#include <iostream>
#include <sstream>
#include "INIReader.h"
#include <cstdlib>
#include <vector>
#include <windows.h>
#include <algorithm>
#include <string>
#include <fstream>

typedef BOOL (WINAPI *LPFN_ISWOW64PROCESS) (HANDLE, PBOOL);
LPFN_ISWOW64PROCESS fnIsWow64Process;

BOOL IsWow64(){
    BOOL bIsWow64 = FALSE;
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

void writeToFile(std::string txtToWrite){
    std::ofstream writeToFile;
    writeToFile.open("install.log", std::ios_base::out | std::ios_base::app);
    if(writeToFile.is_open()){
        writeToFile << txtToWrite;
        writeToFile.close();
    }
}

int main(int argc, char* argv[]){
    writeToFile("Geting INI file ready to read");
    INIReader reader("settings.ini");
    // Check if ini file found or command line arguments were passed. 
    if (reader.ParseError() < 0){
        std::cout << "No ini file found" << std::endl;
        writeToFile("\nNo INI file found");
    }
    std::vector<string> vect;
    string driverInfLoc, portNames, portNameAndNumber, install_printer, serverIP, driverName, printerName;
    string portName, LPRName, IP;

    if(cmdOptionExists(argv, argv+argc, "-HELP") 
        || cmdOptionExists(argv, argv+argc, "-?") 
        || cmdOptionExists(argv, argv+argc, "-help"))
    {
        std::cout << "Usage: addprinter [-HELP] [-IP server_ip][-DRIVER driver][-NAME name][-PORTNAME name][-LPR name][-PRINTER ip][-INF path]" << std::endl;
        std::cout <<  "Arguments:" << std::endl;
        std::cout <<  "-IP           - IP address of the server" << std::endl;
        std::cout <<  "-LPR          - queue name, applies to TCP LPR ports only" << std::endl;
        std::cout <<  "-PORTNAME     - port name" << std::endl;
        std::cout <<  "-DRIVER       - Driver name to be installed, have to match diver name in .INF file" << std::endl;
        std::cout <<  "-NAME         - Print queue name" << std::endl;
        std::cout <<  "-PRINTER      - ip of a printer to get settings from" << std::endl;
        std::cout <<  "-INF          - Inf file location e.g. .\\Driver\\x64\\KOAXWJ__.inf" << std::endl;
        return 1;
    }

    writeToFile("\nChecking given arguments");
    // Check if command line arguments are given otherwise read them from INI file.
    char * option = getCmdOption(argv, argv + argc, "-IP");
    if (option){
        serverIP = option;
        writeToFile("\nCommand line argument -IP with value: " + serverIP);
    }
    else {
        serverIP = reader.Get("Server", "serverIP", "127.0.0.1");
        writeToFile("\nRead serverIP from INI file with value: " + serverIP);
    }
    option = getCmdOption(argv, argv + argc, "-DRIVER");
    if (option){
        driverName = option;
        driverName = "\"" + driverName + "\"";
        writeToFile("\nCommand line argument -DRIVER with value: " + driverName);
    }
    else{
        driverName = reader.Get("Driver", "driverName", "KONICA MINOLTA C368SeriesPCL");
        writeToFile("\nRead driverName from INI file with value: " + driverName);
    }
    option = getCmdOption(argv, argv + argc, "-NAME");
    if (option){
        printerName = option;
        writeToFile("\nCommand line argument -NAME with value: " + printerName);
    }
    else{
        printerName = reader.Get("Queue", "printerName", "UNKNOWN");
        writeToFile("\nRead printerName from INI file with value: " + printerName);
    }
    option = getCmdOption(argv, argv + argc, "-PORTNAME");
    if (option){
        portName = option;
        writeToFile("\nCommand line argument -PORTNAME with value: " + portName);
    }
    else{
        portName = reader.Get("Queue", "portName", "UNKNOWN");
        writeToFile("\nRead portName from INI file with value: " + portName);
    }
    option = getCmdOption(argv, argv + argc, "-LPR");
    if (option){
        LPRName = option;
        writeToFile("\nCommand line argument -LPR with value: " + LPRName);
    }
    else{
        LPRName = reader.Get("Queue", "LPRName", "UNKNOWN");
        writeToFile("\nRead LPRName from INI file with value: " + LPRName);
    }
    option = getCmdOption(argv, argv + argc, "-PRINTER");
    if (option){
        IP = option;
        writeToFile("\nCommand line argument -PRINTER with value: " + IP);
    }
    else{
        IP = reader.Get("Printer", "IP", "UNKNOWN");
        writeToFile("\nRead IP from INI file with value: " + IP);
    }
    option = getCmdOption(argv, argv + argc, "-INF");
    if (option){
        driverInfLoc = option;
        driverInfLoc = "\"" + driverInfLoc + "\"";
        writeToFile("\nCommand line argument -INF with value: " + driverInfLoc);
    }else{
        // Check CPU architecture and load correct drivers.
        writeToFile("\nDetecting CPU architecture...");
        if(IsWow64()){
            driverInfLoc = reader.Get("Driver", "DriverInfLoc_x64", ".\\Driver\\x64\\KOAXWJ__.inf");
            writeToFile("\n64 bit windows detected... Using INF file path" + driverInfLoc);
        }
        else{
            driverInfLoc = reader.Get("Driver", "DriverInfLoc_x86", ".\\Driver\\x84\\KOAXWJ__.inf");
            writeToFile("\n32 bit windows detected... Using INF file path" + driverInfLoc);
        }
    }

    std::stringstream ss(serverIP);
    while (ss >> serverIP)
    {
        vect.push_back(serverIP);

        if (ss.peek() == ',')
            ss.ignore();
    }

    for (size_t i=0; i< vect.size(); i++) {
        portNameAndNumber = portName + "_" + std::to_string(i + 1);
        string create_port = "Cscript .\\scripts\\Prnport.vbs -a -r " + portNameAndNumber + " -h " + vect.at(i) + " -o lpr -q " + LPRName;
        writeToFile("\nConstructing printer port to be created: "+ portNameAndNumber);
        const char *cCreate_port = create_port.c_str();
        system(cCreate_port);
        if(i > 0 && i < vect.size()){
            portNames += "," + portNameAndNumber;
        }
        else{
            portNames = portNameAndNumber;
        }
    }
    if (IP != "UNKNOWN" && IP.size() > 0 ) {
        string create_port_1 = "Cscript .\\scripts\\prnport.vbs -a -r " + IP +" -h " + IP +" -o raw -n 9100";
        const char *cCreate_port_1 = create_port_1.c_str();
        system(cCreate_port_1);
        install_printer = "rundll32 printui.dll,PrintUIEntry /if /b " + printerName + " /f " + driverInfLoc + " /r " + IP + " /m " + driverName;
        writeToFile("Trying to get settings from printer " + IP + "\n");
    }
    else {
        install_printer = "rundll32 printui.dll,PrintUIEntry /if /b " + printerName + " /f " + driverInfLoc + " /r " + portNameAndNumber + " /m " + driverName;
    }
    
    string install_driver = "rundll32 printui.dll,PrintUIEntry /ia /f " + driverInfLoc + " /m " + driverName;
    string add_ports_to_queue = "rundll32 printui.dll,PrintUIEntry /Xs /n " + printerName + " portName " + portNames + " attributes -EnableBidi";

    const char *cInstall_driver = install_driver.c_str();
    const char *cInstall_printer = install_printer.c_str();
    const char *cAdd_ports_to_queue = add_ports_to_queue.c_str();

    try{
        writeToFile("\nStaring Driver install: " + install_driver);
        system(cInstall_driver);
        writeToFile("\nDriver install... DONE \nInstalling printer: " + install_printer);
        system(cInstall_printer);
        writeToFile("\nAdding additional ports to queue: " + add_ports_to_queue);
        system(cAdd_ports_to_queue);
    }catch(int e){
        std::cout << "An exception occured." << e << std::endl;
         writeToFile("\nAn exception occured." + e);
    };


    return 0;
}


