// Example that shows simple usage of the INIReader class

#include <iostream>
#include <sstream>
#include "INIReader.h"
#include <cstdlib>
#include <vector>
#include <windows.h>

using namespace std;

typedef BOOL (WINAPI *LPFN_ISWOW64PROCESS) (HANDLE, PBOOL);

LPFN_ISWOW64PROCESS fnIsWow64Process;

BOOL IsWow64()
{
    BOOL bIsWow64 = FALSE;

    //IsWow64Process is not available on all supported versions of Windows.
    //Use GetModuleHandle to get a handle to the DLL that contains the function
    //and GetProcAddress to get a pointer to the function if available.

    fnIsWow64Process = (LPFN_ISWOW64PROCESS) GetProcAddress(
        GetModuleHandle(TEXT("kernel32")),"IsWow64Process");

    if(NULL != fnIsWow64Process)
    {
        if (!fnIsWow64Process(GetCurrentProcess(),&bIsWow64))
        {
            //handle error
        }
    }
    return bIsWow64;
}

int main()
{
    INIReader reader("settings.ini");

    if (reader.ParseError() < 0) {
        cout << "Can't load 'Settings.ini'" << endl;
        return 1;
    }
    vector<string> vect;
    string DriverInfLoc, PortNames, PortNameAndNumber;

    if(IsWow64()) {
        cout << "64 bit" << endl;
        DriverInfLoc = reader.Get("Driver", "DriverInfLoc_x64", "UNKNOWN");
    }
    else {
        cout << "32 bit" << endl;
        DriverInfLoc = reader.Get("Driver", "DriverInfLoc_x86", "UNKNOWN");
    }

    string ServerIP = reader.Get("Server", "serverIP", "UNKNOWN");
    string DriverName = reader.Get("Driver", "DriverName", "UNKNOWN");
    string PrinterName = reader.Get("Queue", "PrinterName", "UNKNOWN");
    string PortName = reader.Get("Queue", "PortName", "UNKNOWN");
    string LPRName = reader.Get("Queue", "LPRName", "UNKNOWN");

    stringstream ss(ServerIP);
    while (ss >> ServerIP)
    {
        vect.push_back(ServerIP);

        if (ss.peek() == ',')
            ss.ignore();
    }


    for (size_t i=0; i< vect.size(); i++) {
        PortNameAndNumber = PortName + "_" + to_string(i + 1);
        string create_port = "Cscript .\\scripts\\Prnport.vbs -a -r " + PortNameAndNumber + " -h " + vect.at(i) + " -o lpr -q " + LPRName;
        const char *cCreate_port = create_port.c_str();
        system(cCreate_port);
        if(i > 0 && i < vect.size()){
            PortNames += "," + PortNameAndNumber;
        }
        else{
            PortNames = PortNameAndNumber;
        }
    }

    string install_driver = "rundll32 printui.dll,PrintUIEntry /ia /f " + DriverInfLoc + " /m " + DriverName ;
    string install_printer = "rundll32 printui.dll,PrintUIEntry /if /b " + PrinterName + " /f " + DriverInfLoc + " /r " + PortNameAndNumber + " /m " + DriverName;
    string add_ports_to_queue = "rundll32 printui.dll,PrintUIEntry /Xs /n " + PrinterName + " PortName " + PortNames + " attributes -EnableBidi";

    const char *cInstall_driver = install_driver.c_str();
    const char *cInstall_printer = install_printer.c_str();
    const char *cAdd_ports_to_queue = add_ports_to_queue.c_str();


    cout << "Installing driver " + DriverName << endl;
	system(cInstall_driver);
	cout << "Installing printer queue " + PrinterName << endl;
	system(cInstall_printer);
    cout << "Configuring printer..."  << endl;
	system(cAdd_ports_to_queue);


    return 0;
}


