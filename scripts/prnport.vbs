'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
' prnport.vbs - Port script for WMI on Windows 
'     used to add, delete and list ports
'     also for getting and setting the port configuration
'
' Usage:
' prnport [-adlgt?] [-r port] [-s server] [-u user name] [-w password]
'                   [-o raw|lpr] [-h host address] [-q queue] [-n number]
'                   [-me | -md ] [-i SNMP index] [-y community] [-2e | -2d]"
'
' Examples
' prnport -a -s server -r IP_1.2.3.4 -e 1.2.3.4 -o raw -n 9100
' prnport -d -s server -r c:\temp\foo.prn
' prnport -l -s server
' prnport -g -s server -r IP_1.2.3.4
' prnport -t -s server -r IP_1.2.3.4 -me -y public -i 1 -n 9100
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
dim   gDebugFlag
const kDebugTrace = 1
const kDebugError = 2

gDebugFlag = false

'
' Operation action values.
'
const kActionAdd          = 0
const kActionDelete       = 1
const kActionList         = 2
const kActionUnknown      = 3
const kActionGet          = 4
const kActionSet          = 5

const kErrorSuccess       = 0
const KErrorFailure       = 1

const kFlagCreateOrUpdate = 0

const kNameSpace          = "root\cimv2"


'
' Constants for the parameter dictionary
'
const kServerName      = 1
const kPortName        = 2
const kDoubleSpool     = 3
const kPortNumber      = 4
const kPortType        = 5
const kHostAddress     = 6
const kSNMPDeviceIndex = 7
const kCommunityName   = 8
const kSNMP            = 9
const kQueueName       = 10
const kUserName        = 11
const kPassword        = 12

'
' Generic strings
'
const L_Empty_Text                 = ""
const L_Space_Text                 = " "
const L_Colon_Text                 = ":"
const L_LPR_Queue                  = "LPR"
const L_Error_Text                 = "Error"
const L_Success_Text               = "Success"
const L_Failed_Text                = "Failed"
const L_Hex_Text                   = "0x"
const L_Printer_Text               = "Printer"
const L_Operation_Text             = "Operation"
const L_Provider_Text              = "Provider"
const L_Description_Text           = "Description"
const L_Debug_Text                 = "Debug:"

'
' General usage messages
'
const L_Help_Help_General01_Text   = "Usage: prnport [-adlgt?] [-r port][-s server][-u user name][-w password]"
const L_Help_Help_General02_Text   = "               [-o raw|lpr][-h host address][-q queue][-n number]"
const L_Help_Help_General03_Text   = "               [-me | -md ][-i SNMP index][-y community][-2e | -2d]"
const L_Help_Help_General04_Text   = "Arguments:"
const L_Help_Help_General05_Text   = "-a     - add a port"
const L_Help_Help_General06_Text   = "-d     - delete the specified port"
const L_Help_Help_General07_Text   = "-g     - get configuration for a TCP port"
const L_Help_Help_General08_Text   = "-h     - IP address of the device"
const L_Help_Help_General09_Text   = "-i     - SNMP index, if SNMP is enabled"
const L_Help_Help_General10_Text   = "-l     - list all TCP ports"
const L_Help_Help_General11_Text   = "-m     - SNMP type. [e] enable, [d] disable"
const L_Help_Help_General12_Text   = "-n     - port number, applies to TCP RAW ports"
const L_Help_Help_General13_Text   = "-o     - port type, raw or lpr"
const L_Help_Help_General14_Text   = "-q     - queue name, applies to TCP LPR ports only"
const L_Help_Help_General15_Text   = "-r     - port name"
const L_Help_Help_General16_Text   = "-s     - server name"
const L_Help_Help_General17_Text   = "-t     - set configuration for a TCP port"
const L_Help_Help_General18_Text   = "-u     - user name"
const L_Help_Help_General19_Text   = "-w     - password"
const L_Help_Help_General20_Text   = "-y     - community name, if SNMP is enabled"
const L_Help_Help_General21_Text   = "-2     - double spool, applies to TCP LPR ports. [e] enable, [d] disable"
const L_Help_Help_General22_Text   = "-?     - display command usage"
const L_Help_Help_General23_Text   = "Examples:"
const L_Help_Help_General24_Text   = "prnport -l -s server"
const L_Help_Help_General25_Text   = "prnport -d -s server -r IP_1.2.3.4"
const L_Help_Help_General26_Text   = "prnport -a -s server -r IP_1.2.3.4 -h 1.2.3.4 -o raw -n 9100"
const L_Help_Help_General27_Text   = "prnport -t -s server -r IP_1.2.3.4 -me -y public -i 1 -n 9100"
const L_Help_Help_General28_Text   = "prnport -g -s server -r IP_1.2.3.4"
const L_Help_Help_General29_Text   = "prnport -a -r IP_1.2.3.4 -h 1.2.3.4"
const L_Help_Help_General30_Text   = "Remark:"
const L_Help_Help_General31_Text   = "The last example will try to get the device settings at the specified IP address."
const L_Help_Help_General32_Text   = "If a device is detected, then a TCP port is added with the preferred settings for that device."

'
' Messages to be displayed if the scripting host is not cscript
'
const L_Help_Help_Host01_Text      = "This script should be executed from the Command Prompt using CScript.exe."
const L_Help_Help_Host02_Text      = "For example: CScript script.vbs arguments"
const L_Help_Help_Host03_Text      = ""
const L_Help_Help_Host04_Text      = "To set CScript as the default application to run .VBS files run the following:"
const L_Help_Help_Host05_Text      = "     CScript //H:CScript //S"
const L_Help_Help_Host06_Text      = "You can then run ""script.vbs arguments"" without preceding the script with CScript."

'
' General error messages
'
const L_Text_Error_General01_Text  = "The scripting host could not be determined."
const L_Text_Error_General02_Text  = "Unable to parse command line."
const L_Text_Error_General03_Text  = "Win32 error code"

'
' Miscellaneous messages
'
const L_Text_Msg_General01_Text    = "Added port"
const L_Text_Msg_General02_Text    = "Unable to delete port"
const L_Text_Msg_General03_Text    = "Unable to get port"
const L_Text_Msg_General04_Text    = "Created/updated port"
const L_Text_Msg_General05_Text    = "Unable to create/update port"
const L_Text_Msg_General06_Text    = "Unable to enumerate ports"
const L_Text_Msg_General07_Text    = "Number of ports enumerated"
const L_Text_Msg_General08_Text    = "Deleted port"
const L_Text_Msg_General09_Text    = "Unable to get SWbemLocator object"
const L_Text_Msg_General10_Text    = "Unable to connect to WMI service"


'
' Port properties
'
const L_Text_Msg_Port01_Text       = "Server name"
const L_Text_Msg_Port02_Text       = "Port name"
const L_Text_Msg_Port03_Text       = "Host address"
const L_Text_Msg_Port04_Text       = "Protocol RAW"
const L_Text_Msg_Port05_Text       = "Protocol LPR"
const L_Text_Msg_Port06_Text       = "Port number"
const L_Text_Msg_Port07_Text       = "Queue"
const L_Text_Msg_Port08_Text       = "Byte Count Enabled"
const L_Text_Msg_Port09_Text       = "Byte Count Disabled"
const L_Text_Msg_Port10_Text       = "SNMP Enabled"
const L_Text_Msg_Port11_Text       = "SNMP Disabled"
const L_Text_Msg_Port12_Text       = "Community"
const L_Text_Msg_Port13_Text       = "Device index"

'
' Debug messages
'
const L_Text_Dbg_Msg01_Text        = "In function DelPort"
const L_Text_Dbg_Msg02_Text        = "In function CreateOrSetPort"
const L_Text_Dbg_Msg03_Text        = "In function ListPorts"
const L_Text_Dbg_Msg04_Text        = "In function GetPort"
const L_Text_Dbg_Msg05_Text        = "In function ParseCommandLine"

main

'
' Main execution starts here
'
sub main

    on error resume next

    dim iAction
    dim iRetval
    dim oParamDict

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(L_Help_Help_Host01_Text & vbCRLF & L_Help_Help_Host02_Text & vbCRLF & _
                          L_Help_Help_Host03_Text & vbCRLF & L_Help_Help_Host04_Text & vbCRLF & _
                          L_Help_Help_Host05_Text & vbCRLF & L_Help_Help_Host06_Text & vbCRLF)

        wscript.quit

    end if

    set oParamDict = CreateObject("Scripting.Dictionary")

    iRetval = ParseCommandLine(iAction, oParamDict)

    if iRetval = 0 then

        select case iAction

            case kActionAdd
                iRetval = CreateOrSetPort(oParamDict)

            case kActionDelete
                iRetval = DelPort(oParamDict)

            case kActionList
                iRetval = ListPorts(oParamDict)

            case kActionGet
                iRetVal = GetPort(oParamDict)

            case kActionSet
                iRetVal = CreateOrSetPort(oParamDict)

            case else
                Usage(true)
                exit sub

        end select

    end if

end sub

'
' Delete a port
'
function DelPort(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text
    DebugPrint kDebugTrace, L_Text_Msg_Port01_Text & L_Space_Text & oParamDict(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Port02_Text & L_Space_Text & oParamDict(kPortName)

    dim oService
    dim oPort
    dim iResult
    dim strServer
    dim strPort
    dim strUser
    dim strPassword

    iResult = kErrorFailure

    strServer   = oParamDict(kServerName)
    strPort     = oParamDict(kPortName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPort = oService.Get("Win32_TCPIPPrinterPort='" & strPort & "'")

    else

        DelPort = kErrorFailure

        exit function

    end if

    '
    ' Check if Get succeeded
    '
    if Err.Number = kErrorSuccess then

        '
        ' Try deleting the instance
        '
        oPort.Delete_

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General08_Text & L_Space_Text & strPort

        else

            wscript.echo L_Text_Msg_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                         & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

            '
            ' Try getting extended error information
            '
            call LastError()

        end if

    else

        wscript.echo L_Text_Msg_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    DelPort = iResult

end function

'
' Add or update a port
'
function CreateOrSetPort(oParamDict)

    on error resume next

    dim oPort
    dim oService
    dim iResult
    dim PortType
    dim strServer
    dim strPort
    dim strUser
    dim strPassword

    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text
    DebugPrint kDebugTrace, L_Text_Msg_Port01_Text & L_Space_Text & oParamDict.Item(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Port02_Text & L_Space_Text & oParamDict.Item(kPortName)
    DebugPrint kDebugTrace, L_Text_Msg_Port06_Text & L_Space_Text & oParamDict.Item(kPortNumber)
    DebugPrint kDebugTrace, L_Text_Msg_Port07_Text & L_Space_Text & oParamDict.Item(kQueueName)
    DebugPrint kDebugTrace, L_Text_Msg_Port13_Text & L_Space_Text & oParamDict.Item(kSNMPDeviceIndex)
    DebugPrint kDebugTrace, L_Text_Msg_Port12_Text & L_Space_Text & oParamDict.Item(kCommunityName)
    DebugPrint kDebugTrace, L_Text_Msg_Port03_Text & L_Space_Text & oParamDict.Item(kHostAddress)

    strServer   = oParamDict(kServerName)
    strPort     = oParamDict(kPortName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    '
    ' If the port exists, then get the settings. Later PutInstance will do an update
    '
    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPort = oService.Get("Win32_TCPIPPrinterPort.Name='" & strPort & "'")

        '
        ' If get was unsuccessful then spawn a new port instance. Later PutInstance will do a create
        '
        if Err.Number <> kErrorSuccess then

            '
            ' Clear the previous error
            '
            Err.Clear

            set oPort = oService.Get("Win32_TCPIPPrinterPort").SpawnInstance_

        end if

    else

        CreateOrSetPort = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General03_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        CreateOrSetPort = kErrorFailure

        exit function

    end if

    oPort.Name          = oParamDict.Item(kPortName)
    oPort.HostAddress   = oParamDict.Item(kHostAddress)
    oPort.PortNumber    = oParamDict.Item(kPortNumber)
    oPort.SNMPEnabled   = oParamDict.Item(kSNMP)
    oPort.SNMPDevIndex  = oParamDict.Item(kSNMPDeviceIndex)
    oPort.SNMPCommunity = oParamDict.Item(kCommunityName)
    oPort.Queue         = oParamDict.Item(kQueueName)
    oPort.ByteCount     = oParamDict.Item(kDoubleSpool)

    PortType     = oParamDict.Item(kPortType)

    '
    ' Update the port object with the settings corresponding
    ' to the port type of the port to be added
    '
    select case lcase(PortType)

            case "raw"

                 oPort.Protocol      = 1

                 if Not IsNull(oPort.Queue) then

                     wscript.echo L_Error_Text & L_Colon_Text & L_Space_Text _
                     & L_Help_Help_General14_Text

                     CreateOrSetPort = kErrorFailure

                     exit function

                 end if

            case "lpr"

                 oPort.Protocol      = 2

                 if IsNull(oPort.Queue) then

                     oPort.Queue = L_LPR_Queue

                 end if

            case else

                 '
                 ' PutInstance will attempt to get the configuration of
                 ' the device based on its IP address. Those settings
                 ' will be used to add a new port
                 '
    end select

    '
    ' Try creating or updating the port
    '
    oPort.Put_(kFlagCreateOrUpdate)

    if Err.Number = kErrorSuccess then

        wscript.echo L_Text_Msg_General04_Text & L_Space_Text & oPort.Name

        iResult = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General05_Text & L_Space_Text & oPort.Name & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                     & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

        iResult = kErrorFailure

    end if

    CreateOrSetPort = iResult

end function

'
' List ports on a machine.
'
function ListPorts(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text

    dim Ports
    dim oPort
    dim oService
    dim iRetval
    dim iTotal
    dim strServer
    dim strUser
    dim strPassword

    iResult = kErrorFailure

    strServer   = oParamDict(kServerName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Ports = oService.InstancesOf("Win32_TCPIPPrinterPort")

    else

        ListPorts = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General06_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        ListPrinters = kErrorFailure

        exit function

    end if

    iTotal = 0

    for each oPort in Ports

        iTotal = iTotal + 1

        wscript.echo L_Empty_Text
        wscript.echo L_Text_Msg_Port01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Port02_Text & L_Space_Text & oPort.Name
        wscript.echo L_Text_Msg_Port03_Text & L_Space_Text & oPort.HostAddress

        if oPort.Protocol = 1 then

            wscript.echo L_Text_Msg_Port04_Text
            wscript.echo L_Text_Msg_Port06_Text & L_Space_Text & oPort.PortNumber

        else

            wscript.echo L_Text_Msg_Port05_Text
            wscript.echo L_Text_Msg_Port07_Text & L_Space_Text & oPort.Queue

            if oPort.ByteCount then

                wscript.echo L_Text_Msg_Port08_Text

            else

                wscript.echo L_Text_Msg_Port09_Text

            end if

        end if

        if oPort.SNMPEnabled then

            wscript.echo L_Text_Msg_Port10_Text
            wscript.echo L_Text_Msg_Port12_Text & L_Space_Text & oPort.SNMPCommunity
            wscript.echo L_Text_Msg_Port13_Text & L_Space_Text & oPort.SNMPDevIndex

        else

            wscript.echo L_Text_Msg_Port11_Text

        end if

        Err.Clear

    next

    wscript.echo L_Empty_Text
    wscript.echo L_Text_Msg_General07_Text & L_Space_Text & iTotal

    ListPorts = kErrorSuccess

end function

'
' Gets the configuration of a port
'
function GetPort(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text
    DebugPrint kDebugTrace, L_Text_Msg_Port01_Text & L_Space_Text & oParamDict(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Port02_Text & L_Space_Text & oParamDict(kPortName)

    dim oService
    dim oPort
    dim iResult
    dim strServer
    dim strPort
    dim strUser
    dim strPassword

    iResult = kErrorFailure

    strServer   = oParamDict(kServerName)
    strPort     = oParamDict(kPortName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPort = oService.Get("Win32_TCPIPPrinterPort.Name='" & strPort & "'")

    else

        GetPort = kErrorFailure

        exit function

    end if

    if Err.Number = kErrorSuccess then

        wscript.echo L_Empty_Text
        wscript.echo L_Text_Msg_Port01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Port02_Text & L_Space_Text & oPort.Name
        wscript.echo L_Text_Msg_Port03_Text & L_Space_Text & oPort.HostAddress

        if oPort.Protocol = 1 then

            wscript.echo L_Text_Msg_Port04_Text
            wscript.echo L_Text_Msg_Port06_Text & L_Space_Text & oPort.PortNumber

        else

            wscript.echo L_Text_Msg_Port05_Text
            wscript.echo L_Text_Msg_Port07_Text & L_Space_Text & oPort.Queue

            if oPort.ByteCount then

                wscript.echo L_Text_Msg_Port08_Text

            else

                wscript.echo L_Text_Msg_Port09_Text

            end if

        end if

        if oPort.SNMPEnabled then

            wscript.echo L_Text_Msg_Port10_Text
            wscript.echo L_Text_Msg_Port12_Text & L_Space_Text & oPort.SNMPCommunity
            wscript.echo L_Text_Msg_Port13_Text & L_Space_Text & oPort.SNMPDevIndex

        else

            wscript.echo L_Text_Msg_Port11_Text

        end if

        iResult = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General03_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    GetPort = iResult

end function

'
' Debug display helper function
'
sub DebugPrint(uFlags, strString)

    if gDebugFlag = true then

        if uFlags = kDebugTrace then

            wscript.echo L_Debug_Text & L_Space_Text & strString

        end if

        if uFlags = kDebugError then

            if Err <> 0 then

                wscript.echo L_Debug_Text & L_Space_Text & strString & L_Space_Text _
                             & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                             & L_Space_Text & Err.Description

            end if

        end if

    end if

end sub

'
' Parse the command line into its components
'
function ParseCommandLine(iAction, oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg05_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown

    set oArgs = Wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-g"
                iAction = kActionGet

            case "-t"
                iAction = kActionSet

            case "-a"
                iAction = kActionAdd

            case "-d"
                iAction = kActionDelete

            case "-l"
                iAction = kActionList

            case "-2e"
                oParamDict.Add kDoubleSpool, true

            case "-2d"
                oParamDict.Add kDoubleSpool, false

            case "-s"
                iIndex = iIndex + 1
                oParamDict.Add kServerName, RemoveBackslashes(oArgs(iIndex))

            case "-u"
                iIndex = iIndex + 1
                oParamDict.Add kUserName, oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                oParamDict.Add kPassword, oArgs(iIndex)

            case "-n"
                iIndex = iIndex + 1
                oParamDict.Add kPortNumber, oArgs(iIndex)

            case "-r"
                iIndex = iIndex + 1
                oParamDict.Add kPortName, oArgs(iIndex)

            case "-o"
                iIndex = iIndex + 1
                oParamDict.Add kPortType, oArgs(iIndex)

            case "-h"
                iIndex = iIndex + 1
                oParamDict.Add kHostAddress, oArgs(iIndex)

            case "-q"
                iIndex = iIndex + 1
                oParamDict.Add kQueueName, oArgs(iIndex)

            case "-i"
                iIndex = iIndex + 1
                oParamDict.Add kSNMPDeviceIndex, oArgs(iIndex)

            case "-y"
                iIndex = iIndex + 1
                oParamDict.Add kCommunityName, oArgs(iIndex)

            case "-me"
                oParamDict.Add kSNMP, true

            case "-md"
                oParamDict.Add kSNMP, false

            case "-?"
                Usage(True)
                exit function

            case else
                Usage(True)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo L_Text_Error_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_text & Err.Description


        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo L_Help_Help_General01_Text
    wscript.echo L_Help_Help_General02_Text
    wscript.echo L_Help_Help_General03_Text
    wscript.echo L_Help_Help_General04_Text
    wscript.echo L_Help_Help_General05_Text
    wscript.echo L_Help_Help_General06_Text
    wscript.echo L_Help_Help_General07_Text
    wscript.echo L_Help_Help_General08_Text
    wscript.echo L_Help_Help_General09_Text
    wscript.echo L_Help_Help_General10_Text
    wscript.echo L_Help_Help_General11_Text
    wscript.echo L_Help_Help_General12_Text
    wscript.echo L_Help_Help_General13_Text
    wscript.echo L_Help_Help_General14_Text
    wscript.echo L_Help_Help_General15_Text
    wscript.echo L_Help_Help_General16_Text
    wscript.echo L_Help_Help_General17_Text
    wscript.echo L_Help_Help_General18_Text
    wscript.echo L_Help_Help_General19_Text
    wscript.echo L_Help_Help_General20_Text
    wscript.echo L_Help_Help_General21_Text
    wscript.echo L_Help_Help_General22_Text
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General23_Text
    wscript.echo L_Help_Help_General24_Text
    wscript.echo L_Help_Help_General25_Text
    wscript.echo L_Help_Help_General26_Text
    wscript.echo L_Help_Help_General27_Text
    wscript.echo L_Help_Help_General28_Text
    wscript.echo L_Help_Help_General29_Text
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General30_Text
    wscript.echo L_Help_Help_General31_Text
    wscript.echo L_Help_Help_General32_Text

    if bExit then

        wscript.quit(1)

    end if

end sub

'
' Determines which program is being used to run this script.
' Returns true if the script host is cscript.exe
'
function IsHostCscript()

    on error resume next

    dim strFullName
    dim strCommand
    dim i, j
    dim bReturn

    bReturn = false

    strFullName = WScript.FullName

    i = InStr(1, strFullName, ".exe", 1)

    if i <> 0 then

        j = InStrRev(strFullName, "\", i, 1)

        if j <> 0 then

            strCommand = Mid(strFullName, j+1, i-j-1)

            if LCase(strCommand) = "cscript" then

                bReturn = true

            end if

        end if

    end if

    if Err <> 0 then

        wscript.echo L_Text_Error_General01_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

    end if

    IsHostCscript = bReturn

end function

'
' Retrieves extended information about the last error that occurred
' during a WBEM operation. The methods that set an SWbemLastError
' object are GetObject, PutInstance, DeleteInstance
'
sub LastError()

    on error resume next

    dim oError

    set oError = CreateObject("WbemScripting.SWbemLastError")

    if Err = kErrorSuccess then

        wscript.echo L_Operation_Text            & L_Space_Text & oError.Operation
        wscript.echo L_Provider_Text             & L_Space_Text & oError.ProviderName
        wscript.echo L_Description_Text          & L_Space_Text & oError.Description
        wscript.echo L_Text_Error_General04_Text & L_Space_Text & oError.StatusCode

    end if

end sub

'
' Connects to the WMI service on a server. oService is returned as a service
' object (SWbemServices)
'
function WmiConnect(strServer, strNameSpace, strUser, strPassword, oService)

   on error resume next

   dim oLocator
   dim bResult

   oService = null

   bResult  = false

   set oLocator = CreateObject("WbemScripting.SWbemLocator")

   if Err = kErrorSuccess then

      set oService = oLocator.ConnectServer(strServer, strNameSpace, strUser, strPassword)

      if Err = kErrorSuccess then

          bResult = true

          oService.Security_.impersonationlevel = 3

          '
          ' Required to perform administrative tasks on the spooler service
          '
          oService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege"

          Err.Clear

      else

          wscript.echo L_Text_Msg_General10_Text & L_Space_Text & L_Error_Text _
                       & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                       & Err.Description

      end if

   else

       wscript.echo L_Text_Msg_General09_Text & L_Space_Text & L_Error_Text _
                    & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                    & Err.Description

   end if

   WmiConnect = bResult

end function

'
' Remove leading "\\" from server name
'
function RemoveBackslashes(strServer)

    dim strRet

    strRet = strServer

    if Left(strServer, 2) = "\\" and Len(strServer) > 2 then

        strRet = Mid(strServer, 3)

    end if

    RemoveBackslashes = strRet

end function

'' SIG '' Begin signature block
'' SIG '' MIIhowYJKoZIhvcNAQcCoIIhlDCCIZACAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' JP2OI/dGh8Dr4i1IHJc1kEljUqgI55azox+O936+ZZqg
'' SIG '' ggriMIIFAzCCA+ugAwIBAgITMwAAAXMwMQcmZbi5swAA
'' SIG '' AAABczANBgkqhkiG9w0BAQsFADCBhDELMAkGA1UEBhMC
'' SIG '' VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
'' SIG '' B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
'' SIG '' b3JhdGlvbjEuMCwGA1UEAxMlTWljcm9zb2Z0IFdpbmRv
'' SIG '' d3MgUHJvZHVjdGlvbiBQQ0EgMjAxMTAeFw0xNzA4MTEy
'' SIG '' MDIzMzVaFw0xODA4MTEyMDIzMzVaMHAxCzAJBgNVBAYT
'' SIG '' AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
'' SIG '' EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
'' SIG '' cG9yYXRpb24xGjAYBgNVBAMTEU1pY3Jvc29mdCBXaW5k
'' SIG '' b3dzMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
'' SIG '' AQEAyGK7bssJSLHOX62dEwXJWctJkTqAJaN7CTcsC8C+
'' SIG '' GxgCarOwpheOfvNiAFdBgxkHkeEOtDkKv2pZcWasQ+Os
'' SIG '' lm0apWYBF6AyUZbdOz8wWLEgIReZ2ryuqKMk+DDsFam1
'' SIG '' q/zPGMfsi23XbNPfpwO08q3kiTcQA648pZ+ZOp3xlGZq
'' SIG '' ucLmCERCN2rOGqye0rzXOOnHi5TLW0FHWVPjeKD9ox0e
'' SIG '' WaW6dT61HT2nT4p/hrzI/81imldOZ/9c1uwqAirVlJ5/
'' SIG '' p1/zJGr6FDnQLF0UxQ2HycAxSaTuathBjTAUStOXyX3V
'' SIG '' XnjZ6sagUOwqVwZYz1ePwtffCXVV8YinyKz7PwIDAQAB
'' SIG '' o4IBfzCCAXswHwYDVR0lBBgwFgYKKwYBBAGCNwoDBgYI
'' SIG '' KwYBBQUHAwMwHQYDVR0OBBYEFDYYodH2Mo0XhhU8MiuE
'' SIG '' l2V+CJRcMFEGA1UdEQRKMEikRjBEMQwwCgYDVQQLEwNB
'' SIG '' T0MxNDAyBgNVBAUTKzIyOTg3OSs3MTk1NTVjYi02ZGUy
'' SIG '' LTQ0NmMtYWNiYS1kOTA4OTRhY2Q4NzIwHwYDVR0jBBgw
'' SIG '' FoAUqSkCOY4WxJd4zZD5nk+a4XxVr1MwVAYDVR0fBE0w
'' SIG '' SzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
'' SIG '' L3BraW9wcy9jcmwvTWljV2luUHJvUENBMjAxMV8yMDEx
'' SIG '' LTEwLTE5LmNybDBhBggrBgEFBQcBAQRVMFMwUQYIKwYB
'' SIG '' BQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9w
'' SIG '' a2lvcHMvY2VydHMvTWljV2luUHJvUENBMjAxMV8yMDEx
'' SIG '' LTEwLTE5LmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3
'' SIG '' DQEBCwUAA4IBAQBdg/ayviP1J3Oypji7pLNSdIetC2sI
'' SIG '' vvvxrnfS0umkKhLV+WBr2P8OG4tZvhDGr4lGPEtkcDA7
'' SIG '' uPE3gzHBG3WvyiyQQb9UMe4IGU+7BrWGHbEG+0K4D3Ha
'' SIG '' xMQ2jAlXj7phOrKHX3Qs/Otuiv8XgiRMw2r/wlNK1xNg
'' SIG '' I/YzxXMUcGXNkGoCaVLxzFCJQnxHUzxUNazRGGnJou+L
'' SIG '' eoru4LrTfDcuwKWP0qYXSERpva/sh0nqvRbjsS39dkeB
'' SIG '' 6XDHYIe4gqBqnK3yVh12X4oFD63z+dvXInTz7gDMWE9V
'' SIG '' FmibuqGjvJlr6aRUhgVhSOlF9GV3bAYI8fIRLH2hroHb
'' SIG '' nI6PMIIF1zCCA7+gAwIBAgIKYQd2VgAAAAAACDANBgkq
'' SIG '' hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
'' SIG '' BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
'' SIG '' HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEy
'' SIG '' MDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
'' SIG '' dGUgQXV0aG9yaXR5IDIwMTAwHhcNMTExMDE5MTg0MTQy
'' SIG '' WhcNMjYxMDE5MTg1MTQyWjCBhDELMAkGA1UEBhMCVVMx
'' SIG '' EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
'' SIG '' ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
'' SIG '' dGlvbjEuMCwGA1UEAxMlTWljcm9zb2Z0IFdpbmRvd3Mg
'' SIG '' UHJvZHVjdGlvbiBQQ0EgMjAxMTCCASIwDQYJKoZIhvcN
'' SIG '' AQEBBQADggEPADCCAQoCggEBAN0Mu6LkLgnj58X3lmm8
'' SIG '' ACG9aTMz760Ey1SA7gaDu8UghNn30ovzOLCrpK0tfGJ5
'' SIG '' Bf/jSj8ENSBw48Tna+CcwDZ16Yox3Y1w5dw3tXRGlihb
'' SIG '' h2AjLL/cR6Vn91EnnnLrB6bJuR47UzV85dPsJ7mHHP65
'' SIG '' ySMJb6hGkcFuljxB08ujP10Cak3saR8lKFw2//1DFQqU
'' SIG '' 4Bm0z9/CEuLCWyfuJ3gwi1sqCWsiiVNgFizAaB1TuuxJ
'' SIG '' 851hjIVoCXNEXX2iVCvdefcVzzVdbBwrXM68nCOLb261
'' SIG '' Jtk2E8NP1ieuuTI7QZIs4cfNd+iqVE73XAsEh2W0Qxio
'' SIG '' suBtGXfsWiT6SAMCAwEAAaOCAUMwggE/MBAGCSsGAQQB
'' SIG '' gjcVAQQDAgEAMB0GA1UdDgQWBBSpKQI5jhbEl3jNkPme
'' SIG '' T5rhfFWvUzAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMA
'' SIG '' QTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAf
'' SIG '' BgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBW
'' SIG '' BgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29D
'' SIG '' ZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEE
'' SIG '' TjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8y
'' SIG '' MDEwLTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEA
'' SIG '' FPx8cVGlecJusu85Prw8Ug9uKz8QE3P+qGjQSKY0TYqW
'' SIG '' BSbuMUaQYXnW/zguRWv0wOUouNodj4rbCdcax0wKNmZq
'' SIG '' jOwb1wSQqBgXpJu54kAyNnbEwVrGv+QEwOoW06zDaO9i
'' SIG '' rN1UbFAwWKbrfP6Up06O9Ox8hnNXwlIhczRa86OKVsgE
'' SIG '' 2gcJ7fiL4870fo6u8PYLigj7P8kdcn9TuOu+Y+DjPTFl
'' SIG '' sIHl8qzNFqSfPaixm8JC0JCEX1Qd/4nquh1HkG+wc05B
'' SIG '' n0CfX+WhKrIRkXOKISjwzt5zOV8+q1xg7N8DEKjTCen0
'' SIG '' 9paFtn9RiGZHGY2isBI9gSpoBXe7kUxie7bBB8e6eoc0
'' SIG '' Aw5LYnqZ6cr8zko3yS2kV3wc/j3cuA9a+tbEswKFAjrq
'' SIG '' s9lu5GkhN96B0fZ1GQVn05NXXikbOcjuLeHN5EVzW9DS
'' SIG '' znqrFhmCRljQXp2Bs2evbDXyvOU/JOI1ogp1BvYYVpnU
'' SIG '' eCzRBRvr0IgBnaoQ8QXfun4sY7cGmyMhxPl4bOJYFwY2
'' SIG '' K5ESA8yk2fItuvmUnUDtGEXxzopcaz6rA9NwGCoKauBf
'' SIG '' R9HVYwoy8q/XNh8qcFrlQlkIcUtXun6DgfAhPPQcwcW5
'' SIG '' kJMOiEWThumxIJm+mMvFlaRdYtagYwggvXUQd30980W5
'' SIG '' n5efy1eAbzOpBM93pGIcWX4xghYZMIIWFQIBATCBnDCB
'' SIG '' hDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEuMCwGA1UEAxMlTWlj
'' SIG '' cm9zb2Z0IFdpbmRvd3MgUHJvZHVjdGlvbiBQQ0EgMjAx
'' SIG '' MQITMwAAAXMwMQcmZbi5swAAAAABczANBglghkgBZQME
'' SIG '' AgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
'' SIG '' AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
'' SIG '' LwYJKoZIhvcNAQkEMSIEIKB3KTvjHOwjJx/hn7dw25GG
'' SIG '' 9ZVY36lHuHhwdldxEuIxMDwGCisGAQQBgjcKAxwxLgws
'' SIG '' QkMyV2E3YWphdlFZK1N5ZWRHVlhiRjBqcDAwZkUyRG50
'' SIG '' VE5yU1BIRWdDdz0wWgYKKwYBBAGCNwIBDDFMMEqgJIAi
'' SIG '' AE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcA
'' SIG '' c6EigCBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2lu
'' SIG '' ZG93czANBgkqhkiG9w0BAQEFAASCAQBAfNFyW32Bwqwm
'' SIG '' Kz1Czt0F+PMB5syBLn8GTOoOR8ZdtIZH0vJ50eSiwnaD
'' SIG '' o79qmXKXrqHYCNbAf8nxa9B0CI8IQr2ULfDVWFSm1yth
'' SIG '' cw7h4D70kwM4Lr055vfVehDpRQ3dsj72VWdnz6ZLrXiB
'' SIG '' JDjiuheN/oAvaeeh8xjXK0qouw6c+ISYJxv7QvT967B+
'' SIG '' 42f/sUqoZ5sRmTrYLTYK16kf/YHgeK0X/37tSGrP2uQp
'' SIG '' mPQSU3MR5b9drqFppcaD6wmdJUn3W7VhM17U8qto+kNl
'' SIG '' ybAmQ8ddDqiQmdEbpiwE7UwhmK9bOZt3NX/zQ3j/ptcB
'' SIG '' sSlCozWwYXkIxC5n3643oYITRTCCE0EGCisGAQQBgjcD
'' SIG '' AwExghMxMIITLQYJKoZIhvcNAQcCoIITHjCCExoCAQMx
'' SIG '' DzANBglghkgBZQMEAgEFADCCATsGCyqGSIb3DQEJEAEE
'' SIG '' oIIBKgSCASYwggEiAgEBBgorBgEEAYRZCgMBMDEwDQYJ
'' SIG '' YIZIAWUDBAIBBQAEINSHgeeGlous7+wN+08SGydqIZ4l
'' SIG '' c0aRFewIFcWPdfN/AgZZzVbkiJcYEjIwMTcwOTI5MDQy
'' SIG '' NzIxLjU5WjAHAgEBgAIB9KCBuKSBtTCBsjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIERTRSBFU046RjZGRi0yREE3LUJCNzUx
'' SIG '' JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
'' SIG '' cnZpY2Wggg7KMIIGcTCCBFmgAwIBAgIKYQmBKgAAAAAA
'' SIG '' AjANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMx
'' SIG '' EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
'' SIG '' ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
'' SIG '' dGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2Vy
'' SIG '' dGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMTAwNzAx
'' SIG '' MjEzNjU1WhcNMjUwNzAxMjE0NjU1WjB8MQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
'' SIG '' ZS1TdGFtcCBQQ0EgMjAxMDCCASIwDQYJKoZIhvcNAQEB
'' SIG '' BQADggEPADCCAQoCggEBAKkdDbx3EYo6IOz8E5f1+n9p
'' SIG '' lGt0VBDVpQoAgoX77XxoSyxfxcPlYcJ2tz5mK1vwFVMn
'' SIG '' BDEfQRsalR3OCROOfGEwWbEwRA/xYIiEVEMM1024OAiz
'' SIG '' Qt2TrNZzMFcmgqNFDdDq9UeBzb8kYDJYYEbyWEeGMoQe
'' SIG '' dGFnkV+BVLHPk0ySwcSmXdFhE24oxhr5hoC732H8RsEn
'' SIG '' HSRnEnIaIYqvS2SJUGKxXf13Hz3wV3WsvYpCTUBR0Q+c
'' SIG '' Bj5nf/VmwAOWRH7v0Ev9buWayrGo8noqCjHw2k4GkbaI
'' SIG '' CDXoeByw6ZnNPOcvRLqn9NxkvaQBwSAJk3jN/LzAyURd
'' SIG '' XhacAQVPIk0CAwEAAaOCAeYwggHiMBAGCSsGAQQBgjcV
'' SIG '' AQQDAgEAMB0GA1UdDgQWBBTVYzpcijGQ80N7fEYbxTNo
'' SIG '' WoVtVTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
'' SIG '' BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNV
'' SIG '' HSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNV
'' SIG '' HR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJB
'' SIG '' dXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBM
'' SIG '' MEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEw
'' SIG '' LTA2LTIzLmNydDCBoAYDVR0gAQH/BIGVMIGSMIGPBgkr
'' SIG '' BgEEAYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93
'' SIG '' d3cubWljcm9zb2Z0LmNvbS9QS0kvZG9jcy9DUFMvZGVm
'' SIG '' YXVsdC5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcA
'' SIG '' YQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUAbQBl
'' SIG '' AG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAAfmiFEN
'' SIG '' 4sbgmD+BcQM9naOhIW+z66bM9TG+zwXiqf76V20ZMLPC
'' SIG '' xWbJat/15/B4vceoniXj+bzta1RXCCtRgkQS+7lTjMz0
'' SIG '' YBKKdsxAQEGb3FwX/1z5Xhc1mCRWS3TvQhDIr79/xn/y
'' SIG '' N31aPxzymXlKkVIArzgPF/UveYFl2am1a+THzvbKegBv
'' SIG '' SzBEJCI8z+0DpZaPWSm8tv0E4XCfMkon/VWvL/625Y4z
'' SIG '' u2JfmttXQOnxzplmkIz/amJ/3cVKC5Em4jnsGUpxY517
'' SIG '' IW3DnKOiPPp/fZZqkHimbdLhnPkd/DjYlPTGpQqWhqS9
'' SIG '' nhquBEKDuLWAmyI4ILUl5WTs9/S/fmNZJQ96LjlXdqJx
'' SIG '' qgaKD4kWumGnEcua2A5HmoDF0M2n0O99g/DhO3EJ3110
'' SIG '' mCIIYdqwUB5vvfHhAN/nMQekkzr3ZUd46PioSKv33nJ+
'' SIG '' YWtvd6mBy6cJrDm77MbL2IK0cs0d9LiFAR6A+xuJKlQ5
'' SIG '' slvayA1VmXqHczsI5pgt6o3gMy4SKfXAL1QnIffIrE7a
'' SIG '' KLixqduWsqdCosnPGUFN4Ib5KpqjEWYw07t0MkvfY3v1
'' SIG '' mYovG8chr1m1rtxEPJdQcdeh0sVV42neV8HR3jDA/czm
'' SIG '' TfsNv11P6Z0eGTgvvM9YBS7vDaBQNdrvCScc1bN+NR4I
'' SIG '' uto229Nfj950iEkSMIIE2TCCA8GgAwIBAgITMwAAAKVI
'' SIG '' F3In+XC+YwAAAAAApTANBgkqhkiG9w0BAQsFADB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0xNjA5MDcx
'' SIG '' NzU2NTBaFw0xODA5MDcxNzU2NTBaMIGyMQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMQwwCgYDVQQLEwNBT0MxJzAlBgNVBAsT
'' SIG '' Hm5DaXBoZXIgRFNFIEVTTjpGNkZGLTJEQTctQkI3NTEl
'' SIG '' MCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
'' SIG '' dmljZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
'' SIG '' ggEBALTaktS9TF7w21cH31lhgTomttMYs2nD/fvMI2/S
'' SIG '' DczTG1YWzFBNvsMS+1zWwTSzwjnJtPsh+jLXYWwKCl5u
'' SIG '' T74Kly/RrQLk4dDvaYGfHzLGTYQm/eJ7qNIiLDzxzCs5
'' SIG '' Mi/+G+yeT2/9i8dU84WdfoJLUhKO7g9jPbY3RPpOW7tf
'' SIG '' 3Y/+5oXXA1IRnsGU/zX7hzL2EyCRp5o3ofJPpDJUS+Ao
'' SIG '' FOXnIYbt1sOamTQhJjB1B5igZehTt9CZWpQSZrvbVlY0
'' SIG '' 1ESRYYnQuNRfOk7lgDrH8lKbOnXX3HaoIFMgxxm3FFV1
'' SIG '' fPvKAiIOoUuB87DmWcUapQuByhcCPmrFC46/wbcCAwEA
'' SIG '' AaOCARswggEXMB0GA1UdDgQWBBTiE6tL8u2xYLh48YWg
'' SIG '' Uf5y57EvpDAfBgNVHSMEGDAWgBTVYzpcijGQ80N7fEYb
'' SIG '' xTNoWoVtVTBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8v
'' SIG '' Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0
'' SIG '' cy9NaWNUaW1TdGFQQ0FfMjAxMC0wNy0wMS5jcmwwWgYI
'' SIG '' KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8v
'' SIG '' d3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Rp
'' SIG '' bVN0YVBDQV8yMDEwLTA3LTAxLmNydDAMBgNVHRMBAf8E
'' SIG '' AjAAMBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3
'' SIG '' DQEBCwUAA4IBAQA2fW15K8YuLyQgyiX1NBWM3PJLRwf0
'' SIG '' oz7uhngI3nuH1gbFWOn7y/MXTh/pMaDG0MJmA5+uzfDs
'' SIG '' nCtZk4JTupERHAqex2IaWqPVFsstpurA8rT/eX77DvAz
'' SIG '' 6k4brlza9FAu6EuoZxkGq8ffwX1hBSIwYc6lmMDAAih9
'' SIG '' aTEpmSBupDbn4pTShGzcDRpJyfXjNlVtbffVWxHUOboA
'' SIG '' 36bRJMtJMbwlgIJgpOsZ5iGCaS9IkrJQxQ8OnTffHrz+
'' SIG '' uHLrk0P+W8YfG16gaF3eDhTkItRqFVbk6OLnrY/KJzhi
'' SIG '' ZtZs2yYSLwmDxa5wQeI5HoQtC7HviXAmUgQy6TwMA55s
'' SIG '' bPhuoYIDdDCCAlwCAQEwgeKhgbikgbUwgbIxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xDDAKBgNVBAsTA0FPQzEnMCUGA1UE
'' SIG '' CxMebkNpcGhlciBEU0UgRVNOOkY2RkYtMkRBNy1CQjc1
'' SIG '' MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
'' SIG '' ZXJ2aWNloiUKAQEwCQYFKw4DAhoFAAMVAJvCNd37siYr
'' SIG '' GhQxZVAvHVOUaYJvoIHBMIG+pIG7MIG4MQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMQwwCgYDVQQLEwNBT0MxJzAlBgNVBAsT
'' SIG '' Hm5DaXBoZXIgTlRTIEVTTjoyNjY1LTRDM0YtQzVERTEr
'' SIG '' MCkGA1UEAxMiTWljcm9zb2Z0IFRpbWUgU291cmNlIE1h
'' SIG '' c3RlciBDbG9jazANBgkqhkiG9w0BAQUFAAIFAN14PUIw
'' SIG '' IhgPMjAxNzA5MjkwMzMyMThaGA8yMDE3MDkzMDAzMzIx
'' SIG '' OFowdDA6BgorBgEEAYRZCgQBMSwwKjAKAgUA3Xg9QgIB
'' SIG '' ADAHAgEAAgI9RDAHAgEAAgIZzDAKAgUA3XmOwgIBADA2
'' SIG '' BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMBoAow
'' SIG '' CAIBAAIDFuNgoQowCAIBAAIDB6EgMA0GCSqGSIb3DQEB
'' SIG '' BQUAA4IBAQAZuYpo/zkY8CSSUVCUG3JGhEu+DFgIvP64
'' SIG '' klRXMac/0FzNGbyVYuD+PgXmhBHym9XfcqqykZZM/Q9b
'' SIG '' 0swP5Inn98hYXld/WZvmPtzdX2Oi+CrVdFr06kPVyNKU
'' SIG '' wTUgmSWkN3cOh4OYkQNcA8viFtHvI6+NKcMVR6OJ95nU
'' SIG '' 0z5r4SkfoZY/1qzwVN1POHt3m6uU62Y2VUtrVyKJjXL0
'' SIG '' VqmyK8me1Ou6iAZxu1nSEqUuenL6p9eh6oDNF6mwGsHZ
'' SIG '' 0v17w2i2oA1JvENV8wl935Ou+G3p1K1kOWxY+DsLBaru
'' SIG '' LX3BHmeDWCLnuXV3jIr/qn8EcnrAuaPszDpL9Z8tZqK1
'' SIG '' MYIC9TCCAvECAQEwgZMwfDELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' UENBIDIwMTACEzMAAAClSBdyJ/lwvmMAAAAAAKUwDQYJ
'' SIG '' YIZIAWUDBAIBBQCgggEyMBoGCSqGSIb3DQEJAzENBgsq
'' SIG '' hkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgfT9m6lMC
'' SIG '' 1YEOOSq5LpnygM6m0GVgX9OLKhvcxz6Hw4gwgeIGCyqG
'' SIG '' SIb3DQEJEAIMMYHSMIHPMIHMMIGxBBSbwjXd+7ImKxoU
'' SIG '' MWVQLx1TlGmCbzCBmDCBgKR+MHwxCzAJBgNVBAYTAlVT
'' SIG '' MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
'' SIG '' ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
'' SIG '' YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
'' SIG '' YW1wIFBDQSAyMDEwAhMzAAAApUgXcif5cL5jAAAAAACl
'' SIG '' MBYEFEwBPvqh0A4zK/e69I7lcHW09v8VMA0GCSqGSIb3
'' SIG '' DQEBCwUABIIBAKI0HxltrtI2Y1rF3WBNSP4lrlh/vgfu
'' SIG '' TWEk3D2mqDWQnafwjjfy+7G8WEC5za0+l60uekbevhuX
'' SIG '' vu/7h0UKRpVXkDmz1/Gg/kiwLIRqseJFgkToa1dQ3UTV
'' SIG '' Nn/B6hVydeGTaHaDtwzEnzO7KV9fAXnDbNyq+VXoGKvO
'' SIG '' vDXT3qEZXCnSRZ9HclCAWf8r5tv9XGzRF8zi9UkbKHWt
'' SIG '' XgkV6ujpNxNLI1PbsNF/t9ZCqiIwwAuaDN57wFbVqpDo
'' SIG '' ECEVkJrCQx/UuRA1EmRjKIKjxduZ44wQpbIgxZtYNGNS
'' SIG '' bsi0dZfLtHAZrwMYCBjzG5Oz9GOzknyOv8FyI9hnXehNVPA=
'' SIG '' End signature block
