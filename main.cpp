#include <iostream>
#include <sstream>
#include "INIReader.h"
#include <cstdlib>
#include <vector>
#include <windows.h>
#include <algorithm>
#include <string>

typedef BOOL (WINAPI *LPFN_ISWOW64PROCESS) (HANDLE, PBOOL);

LPFN_ISWOW64PROCESS fnIsWow64Process;

BOOL IsWow64(){
    BOOL bIsWow64 = FALSE;

    //IsWow64Process is not available on all supported versions of Windows.
    //Use GetModuleHandle to get a handle to the DLL that contains the function
    //and GetProcAddress to get a pointer to the function if available.

    fnIsWow64Process = (LPFN_ISWOW64PROCESS) GetProcAddress(
        GetModuleHandle(TEXT("kernel32")),"IsWow64Process");

    if(NULL != fnIsWow64Process){
        if (!fnIsWow64Process(GetCurrentProcess(),&bIsWow64)){
            //handle error
        }
    }
    return bIsWow64;
}

char* getCmdOption(char ** begin, char ** end, const std::string & option){
    char ** itr = std::find(begin, end, option);
    if (itr != end && ++itr != end)
    {
        return *itr;
    }
    return 0;
}

bool cmdOptionExists(char** begin, char** end, const std::string& option){
    return std::find(begin, end, option) != end;
}

int main(int argc, char* argv[]){
    INIReader reader("settings.ini");
    // Check if ini file found or command line argument where passed. 
    if (reader.ParseError() < 0){
        std::cout << "No ini file found" << std::endl;
        return 1;
    }
    std::vector<string> vect;
    string DriverInfLoc, PortNames, PortNameAndNumber, install_printer, ServerIP, DriverName, PrinterName;
    string PortName, LPRName, IP;

    if(cmdOptionExists(argv, argv+argc, "-HELP") 
        || cmdOptionExists(argv, argv+argc, "-?") 
        || cmdOptionExists(argv, argv+argc, "-help"))
    {
        std::cout << "Usage: addprinter [-HELP] [-IP server_ip][-DRIVER driver][-NAME name][-PORTNAME name][-LPR name][-PRINTER ip]" << std::endl;
        std::cout <<  "Arguments:" << std::endl;
        std::cout <<  "-IP           - IP address of the server" << std::endl;
        std::cout <<  "-LPR          - queue name, applies to TCP LPR ports only" << std::endl;
        std::cout <<  "-PORTNAME     - port name" << std::endl;
        std::cout <<  "-DRIVER       - Driver name to be installed, have to match diver name in .INF file" << std::endl;
        std::cout <<  "-NAME         - Print queue name" << std::endl;
        std::cout <<  "-PRINTER      - ip of a printer to get settings from" << std::endl;
        std::cout <<  "The last example will try to get the device settings at the specified IP address." << std::endl;
        std::cout <<  "If a device is detected, then a TCP port is added with the preferred settings for that device." << std::endl;
        return 1;
    }

    char * option = getCmdOption(argv, argv + argc, "-IP");
    if (option)
    {
        ServerIP = option;
    }
    option = getCmdOption(argv, argv + argc, "-DRIVER");
    if (option)
    {
        DriverName = option;
    }
    option = getCmdOption(argv, argv + argc, "-NAME");
    if (option)
    {
        PrinterName = option;
    }
    option = getCmdOption(argv, argv + argc, "-PORTNAME");
    if (option)
    {
        PortName = option;
    }
    option = getCmdOption(argv, argv + argc, "-LPR");
    if (option)
    {
        LPRName = option;
    }
    option = getCmdOption(argv, argv + argc, "-PRINTER");
    if (option)
    {
        IP = option;
    }

    // Read settings from ini file if parameters are not given.
    if (ServerIP.length() < 1){
        ServerIP = reader.Get("Server", "serverIP", "127.0.0.1");
    } 
    if (DriverName.length() < 1){
        DriverName = reader.Get("Driver", "DriverName", "UNKNOWN");
    }
    if (PrinterName.length() < 1){
        PrinterName = reader.Get("Queue", "PrinterName", "UNKNOWN");
    }
    if (PortName.length() < 1){
        PortName = reader.Get("Queue", "PortName", "UNKNOWN");
    }
    if (LPRName.length() < 1){
        LPRName = reader.Get("Queue", "LPRName", "UNKNOWN");
    }
    if (IP.length() < 1){
        IP = reader.Get("Printer", "IP", "UNKNOWN");
    }

    // Check CPU architecture and load correct drivers.
    if(IsWow64()) {
        DriverInfLoc = reader.Get("Driver", "DriverInfLoc_x64", "UNKNOWN");
    }
    else {
        DriverInfLoc = reader.Get("Driver", "DriverInfLoc_x86", "UNKNOWN");
    }

    std::stringstream ss(ServerIP);
    while (ss >> ServerIP)
    {
        vect.push_back(ServerIP);

        if (ss.peek() == ',')
            ss.ignore();
    }

    for (size_t i=0; i< vect.size(); i++) {
        PortNameAndNumber = PortName + "_" + std::to_string(i + 1);
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
    if (IP != "UNKNOWN" && IP.size() > 0 ) {
        string create_port_1 = "Cscript .\\scripts\\prnport.vbs -a -r " + IP +" -h " + IP +" -o raw -n 9100";
        const char *cCreate_port_1 = create_port_1.c_str();
        system(cCreate_port_1);
        install_printer = "rundll32 printui.dll,PrintUIEntry /if /b " + PrinterName + " /f " + DriverInfLoc + " /r " + IP + " /m " + DriverName;
    }
    else {
        install_printer = "rundll32 printui.dll,PrintUIEntry /if /b " + PrinterName + " /f " + DriverInfLoc + " /r " + PortNameAndNumber + " /m " + DriverName;
    }

    string install_driver = "rundll32 printui.dll,PrintUIEntry /ia /f " + DriverInfLoc + " /m " + DriverName ;
    string add_ports_to_queue = "rundll32 printui.dll,PrintUIEntry /Xs /n " + PrinterName + " PortName " + PortNames + " attributes -EnableBidi";

    const char *cInstall_driver = install_driver.c_str();
    const char *cInstall_printer = install_printer.c_str();
    const char *cAdd_ports_to_queue = add_ports_to_queue.c_str();

	system(cInstall_driver);
	system(cInstall_printer);
	system(cAdd_ports_to_queue);

    return 0;
}


