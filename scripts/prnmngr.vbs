'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
' prnmngr.vbs - printer script for WMI on Windows 
'     used to add, delete, and list printers and connections
'     also for getting and setting the default printer
'
' Usage:
' prnmngr [-adxgtl?][co] [-s server][-p printer][-m driver model][-r port]
'                       [-u user name][-w password]
'
' Examples:
' prnmngr -a -p "printer" -m "driver" -r "lpt1:"
' prnmngr -d -p "printer" -s server
' prnmngr -ac -p "\\server\printer"
' prnmngr -d -p "\\server\printer"
' prnmngr -x -s server
' prnmngr -l -s server
' prnmngr -g
' prnmngr -t -p "printer"
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
const kDebugTrace = 1
const kDebugError = 2
dim   gDebugFlag

gDebugFlag = false

'
' Operation action values.
'
const kActionUnknown           = 0
const kActionAdd               = 1
const kActionAddConn           = 2
const kActionDel               = 3
const kActionDelAll            = 4
const kActionDelAllCon         = 5
const kActionDelAllLocal       = 6
const kActionList              = 7
const kActionGetDefaultPrinter = 8
const kActionSetDefaultPrinter = 9

const kErrorSuccess            = 0
const KErrorFailure            = 1

const kFlagCreateOnly          = 2

const kNameSpace               = "root\cimv2"

'
' Generic strings
'
const L_Empty_Text                 = ""
const L_Space_Text                 = " "
const L_Error_Text                 = "Error"
const L_Success_Text               = "Success"
const L_Failed_Text                = "Failed"
const L_Hex_Text                   = "0x"
const L_Printer_Text               = "Printer"
const L_Operation_Text             = "Operation"
const L_Provider_Text              = "Provider"
const L_Description_Text           = "Description"
const L_Debug_Text                 = "Debug:"
const L_Connection_Text            = "connection"

'
' General usage messages
'
const L_Help_Help_General01_Text   = "Usage: prnmngr [-adxgtl?][c] [-s server][-p printer][-m driver model]"
const L_Help_Help_General02_Text   = "               [-r port][-u user name][-w password]"
const L_Help_Help_General03_Text   = "Arguments:"
const L_Help_Help_General04_Text   = "-a     - add local printer"
const L_Help_Help_General05_Text   = "-ac    - add printer connection"
const L_Help_Help_General06_Text   = "-d     - delete printer"
const L_Help_Help_General07_Text   = "-g     - get the default printer"
const L_Help_Help_General08_Text   = "-l     - list printers"
const L_Help_Help_General09_Text   = "-m     - driver model"
const L_Help_Help_General10_Text   = "-p     - printer name"
const L_Help_Help_General11_Text   = "-r     - port name"
const L_Help_Help_General12_Text   = "-s     - server name"
const L_Help_Help_General13_Text   = "-t     - set the default printer"
const L_Help_Help_General14_Text   = "-u     - user name"
const L_Help_Help_General15_Text   = "-w     - password"
const L_Help_Help_General16_Text   = "-x     - delete all printers"
const L_Help_Help_General17_Text   = "-xc    - delete all printer connections"
const L_Help_Help_General18_Text   = "-xo    - delete all local printers"
const L_Help_Help_General19_Text   = "-?     - display command usage"
const L_Help_Help_General20_Text   = "Examples:"
const L_Help_Help_General21_Text   = "prnmngr -a -p ""printer"" -m ""driver"" -r ""lpt1:"""
const L_Help_Help_General22_Text   = "prnmngr -d -p ""printer"" -s server"
const L_Help_Help_General23_Text   = "prnmngr -ac -p ""\\server\printer"""
const L_Help_Help_General24_Text   = "prnmngr -d -p ""\\server\printer"""
const L_Help_Help_General25_Text   = "prnmngr -x -s server"
const L_Help_Help_General26_Text   = "prnmngr -xo"
const L_Help_Help_General27_Text   = "prnmngr -l -s server"
const L_Help_Help_General28_Text   = "prnmngr -g"
const L_Help_Help_General29_Text   = "prnmngr -t -p ""\\server\printer"""

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
const L_Text_Msg_General01_Text    = "Added printer"
const L_Text_Msg_General02_Text    = "Unable to add printer"
const L_Text_Msg_General03_Text    = "Added printer connection"
const L_Text_Msg_General04_Text    = "Unable to add printer connection"
const L_Text_Msg_General05_Text    = "Deleted printer"
const L_Text_Msg_General06_Text    = "Unable to delete printer"
const L_Text_Msg_General07_Text    = "Attempting to delete printer"
const L_Text_Msg_General08_Text    = "Unable to delete printers"
const L_Text_Msg_General09_Text    = "Number of local printers and connections enumerated"
const L_Text_Msg_General10_Text    = "Number of local printers and connections deleted"
const L_Text_Msg_General11_Text    = "Unable to enumerate printers"
const L_Text_Msg_General12_Text    = "The default printer is"
const L_Text_Msg_General13_Text    = "Unable to get the default printer"
const L_Text_Msg_General14_Text    = "Unable to set the default printer"
const L_Text_Msg_General15_Text    = "The default printer is now"
const L_Text_Msg_General16_Text    = "Number of printer connections enumerated"
const L_Text_Msg_General17_Text    = "Number of printer connections deleted"
const L_Text_Msg_General18_Text    = "Number of local printers enumerated"
const L_Text_Msg_General19_Text    = "Number of local printers deleted"

'
' Printer properties
'
const L_Text_Msg_Printer01_Text    = "Server name"
const L_Text_Msg_Printer02_Text    = "Printer name"
const L_Text_Msg_Printer03_Text    = "Share name"
const L_Text_Msg_Printer04_Text    = "Driver name"
const L_Text_Msg_Printer05_Text    = "Port name"
const L_Text_Msg_Printer06_Text    = "Comment"
const L_Text_Msg_Printer07_Text    = "Location"
const L_Text_Msg_Printer08_Text    = "Separator file"
const L_Text_Msg_Printer09_Text    = "Print processor"
const L_Text_Msg_Printer10_Text    = "Data type"
const L_Text_Msg_Printer11_Text    = "Parameters"
const L_Text_Msg_Printer12_Text    = "Attributes"
const L_Text_Msg_Printer13_Text    = "Priority"
const L_Text_Msg_Printer14_Text    = "Default priority"
const L_Text_Msg_Printer15_Text    = "Start time"
const L_Text_Msg_Printer16_Text    = "Until time"
const L_Text_Msg_Printer17_Text    = "Job count"
const L_Text_Msg_Printer18_Text    = "Average pages per minute"
const L_Text_Msg_Printer19_Text    = "Printer status"
const L_Text_Msg_Printer20_Text    = "Extended printer status"
const L_Text_Msg_Printer21_Text    = "Detected error state"
const L_Text_Msg_Printer22_Text    = "Extended detected error state"


'
' Printer status
'
const L_Text_Msg_Status01_Text     = "Other"
const L_Text_Msg_Status02_Text     = "Unknown"
const L_Text_Msg_Status03_Text     = "Idle"
const L_Text_Msg_Status04_Text     = "Printing"
const L_Text_Msg_Status05_Text     = "Warmup"
const L_Text_Msg_Status06_Text     = "Stopped printing"
const L_Text_Msg_Status07_Text     = "Offline"
const L_Text_Msg_Status08_Text     = "Paused"
const L_Text_Msg_Status09_Text     = "Error"
const L_Text_Msg_Status10_Text     = "Busy"
const L_Text_Msg_Status11_Text     = "Not available"
const L_Text_Msg_Status12_Text     = "Waiting"
const L_Text_Msg_Status13_Text     = "Processing"
const L_Text_Msg_Status14_Text     = "Initializing"
const L_Text_Msg_Status15_Text     = "Power save"
const L_Text_Msg_Status16_Text     = "Pending deletion"
const L_Text_Msg_Status17_Text     = "I/O active"
const L_Text_Msg_Status18_Text     = "Manual feed"
const L_Text_Msg_Status19_Text     = "No error"
const L_Text_Msg_Status20_Text     = "Low paper"
const L_Text_Msg_Status21_Text     = "No paper"
const L_Text_Msg_Status22_Text     = "Low toner"
const L_Text_Msg_Status23_Text     = "No toner"
const L_Text_Msg_Status24_Text     = "Door open"
const L_Text_Msg_Status25_Text     = "Jammed"
const L_Text_Msg_Status26_Text     = "Service requested"
const L_Text_Msg_Status27_Text     = "Output bin full"
const L_Text_Msg_Status28_Text     = "Paper problem"
const L_Text_Msg_Status29_Text     = "Cannot print page"
const L_Text_Msg_Status30_Text     = "User intervention required"
const L_Text_Msg_Status31_Text     = "Out of memory"
const L_Text_Msg_Status32_Text     = "Server unknown"

'
' Debug messages
'
const L_Text_Dbg_Msg01_Text        = "In function AddPrinter"
const L_Text_Dbg_Msg02_Text        = "In function AddPrinterConnection"
const L_Text_Dbg_Msg03_Text        = "In function DelPrinter"
const L_Text_Dbg_Msg04_Text        = "In function DelAllPrinters"
const L_Text_Dbg_Msg05_Text        = "In function ListPrinters"
const L_Text_Dbg_Msg06_Text        = "In function GetDefaultPrinter"
const L_Text_Dbg_Msg07_Text        = "In function SetDefaultPrinter"
const L_Text_Dbg_Msg08_Text        = "In function ParseCommandLine"

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strPrinter
    dim strDriver
    dim strPort
    dim strUser
    dim strPassword

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(L_Help_Help_Host01_Text & vbCRLF & L_Help_Help_Host02_Text & vbCRLF & _
                          L_Help_Help_Host03_Text & vbCRLF & L_Help_Help_Host04_Text & vbCRLF & _
                          L_Help_Help_Host05_Text & vbCRLF & L_Help_Help_Host06_Text & vbCRLF)

        wscript.quit

    end if

    '
    ' Get command line parameters
    '
    iRetval = ParseCommandLine(iAction, strServer, strPrinter, strDriver, strPort, strUser, strPassword)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionAdd
                 iRetval = AddPrinter(strServer, strPrinter, strDriver, strPort, strUser, strPassword)

            case kActionAddConn
                 iRetval = AddPrinterConnection(strPrinter, strUser, strPassword)

            case kActionDel
                 iRetval = DelPrinter(strServer, strPrinter, strUser, strPassword)

            case kActionDelAll
                 iRetval = DelAllPrinters(kActionDelAll, strServer, strUser, strPassword)

            case kActionDelAllCon
                 iRetval = DelAllPrinters(kActionDelAllCon, strServer, strUser, strPassword)

            case kActionDelAllLocal
                 iRetval = DelAllPrinters(kActionDelAllLocal, strServer, strUser, strPassword)

            case kActionList
                 iRetval = ListPrinters(strServer, strUser, strPassword)

            case kActionGetDefaultPrinter
                 iRetval = GetDefaultPrinter(strUser, strPassword)

            case kActionSetDefaultPrinter
                 iRetval = SetDefaultPrinter(strPrinter, strUser, strPassword)

            case kActionUnknown
                 Usage(true)
                 exit sub

            case else
                 Usage(true)
                 exit sub

        end select

    end if

end sub

'
' Add a printer with minimum settings. Use prncnfg.vbs to
' set the complete configuration of a printer
'
function AddPrinter(strServer, strPrinter, strDriver, strPort, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & strServer
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & strPrinter
    DebugPrint kDebugTrace, L_Text_Msg_Printer04_Text & L_Space_Text & strDriver
    DebugPrint kDebugTrace, L_Text_Msg_Printer05_Text & L_Space_Text & strPort

    dim oPrinter
    dim oService
    dim iRetval

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer").SpawnInstance_

    else

        AddPrinter = kErrorFailure

        exit function

    end if

    oPrinter.DriverName = strDriver
    oPrinter.PortName   = strPort
    oPrinter.DeviceID   = strPrinter

    oPrinter.Put_(kFlagCreateOnly)

    if Err.Number = kErrorSuccess then

        wscript.echo L_Text_Msg_General01_Text & L_Space_Text & strPrinter

        iRetval = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General02_Text & L_Space_Text & strPrinter & L_Space_Text & L_Error_Text _
                     & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

        iRetval = kErrorFailure

    end if

    AddPrinter = iRetval

end function

'
' Add a printer connection
'
function AddPrinterConnection(strPrinter, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text

    dim oPrinter
    dim oService
    dim iRetval
    dim uResult

    '
    ' Initialize return value
    '
    iRetval = kErrorFailure

    '
    ' We connect to the local server
    '
    if WmiConnect("", kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer")

    else

        AddPrinterConnection = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        '
        ' The Err object indicates whether the WMI provider reached the execution
        ' of the function that adds a printer connection. The uResult is the Win32
        ' error code returned by the static method that adds a printer connection
        '
        uResult = oPrinter.AddPrinterConnection(strPrinter)

        if Err.Number = kErrorSuccess then

            if uResult = kErrorSuccess then

                wscript.echo L_Text_Msg_General03_Text & L_Space_Text & strPrinter

                iRetval = kErrorSuccess

            else

                wscript.echo L_Text_Msg_General04_Text & L_Space_Text & L_Text_Error_General03_Text _
                             & L_Space_text & uResult

            end if

        else

            wscript.echo L_Text_Msg_General04_Text & L_Space_Text & strPrinter & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                         & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General04_Text & L_Space_Text & strPrinter & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                     & Err.Description

    end if

    AddPrinterConnection = iRetval

end function

'
' Delete a printer or a printer connection
'
function DelPrinter(strServer, strPrinter, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & strServer
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & strPrinter

    dim oService
    dim oPrinter
    dim iRetval

    iRetval = kErrorFailure

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer.DeviceID='" & strPrinter & "'")

    else

        DelPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        oPrinter.Delete_

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General05_Text & L_Space_Text & strPrinter

            iRetval = kErrorSuccess

        else

            wscript.echo L_Text_Msg_General06_Text & L_Space_Text & strPrinter & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                         & L_Space_Text & Err.Description

            '
            ' Try getting extended error information
            '
            call LastError()

        end if

    else

        wscript.echo L_Text_Msg_General06_Text & L_Space_Text & strPrinter & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                     & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    DelPrinter = iRetval

end function

'
' Delete all local printers and connections on a machine
'
function DelAllPrinters(kAction, strServer, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text

    dim Printers
    dim oPrinter
    dim oService
    dim iResult
    dim iTotal
    dim iTotalDeleted
    dim strPrinterName
    dim bDelete
    dim bConnection
    dim strTemp

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Printers = oService.InstancesOf("Win32_Printer")

    else

        DelAllPrinters = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General11_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        DelAllPrinters = kErrorFailure

        exit function

    end if

    iTotal = 0
    iTotalDeleted = 0

    for each oPrinter in Printers

        strPrinterName = oPrinter.DeviceID

        bConnection = oPrinter.Network

        if kAction = kActionDelAll then

            bDelete = 1

            iTotal = iTotal + 1

        elseif kAction = kActionDelAllCon and bConnection then

            bDelete = 1

            iTotal = iTotal + 1

        elseif kAction = kActionDelAllLocal and not bConnection then

            bDelete = 1

            iTotal = iTotal + 1

        else

            bDelete = 0

        end if

        if bDelete = 1 then

            if bConnection then

                strTemp = L_Space_Text & L_Connection_Text & L_Space_Text

            else

                strTemp = L_Space_Text

            end if

            '
            ' Delete printer instance
            '
            oPrinter.Delete_

            if Err.Number = kErrorSuccess then

                wscript.echo L_Text_Msg_General05_Text & strTemp & oPrinter.DeviceID

                iTotalDeleted = iTotalDeleted + 1

            else

                wscript.echo L_Text_Msg_General06_Text & strTemp & strPrinterName _
                             & L_Space_Text & L_Error_Text & L_Space_Text & L_Hex_Text _
                             & hex(Err.Number) & L_Space_Text & Err.Description

                '
                ' Try getting extended error information
                '
                call LastError()

                '
                ' Continue deleting the rest of the printers despite this error
                '
                Err.Clear

            end if

        end if

    next

    wscript.echo L_Empty_Text

    if kAction = kActionDelAll then

        wscript.echo L_Text_Msg_General09_Text & L_Space_Text & iTotal
        wscript.echo L_Text_Msg_General10_Text & L_Space_Text & iTotalDeleted

    elseif kAction = kActionDelAllCon then

        wscript.echo L_Text_Msg_General16_Text & L_Space_Text & iTotal
        wscript.echo L_Text_Msg_General17_Text & L_Space_Text & iTotalDeleted

    elseif kAction = kActionDelAllLocal then

        wscript.echo L_Text_Msg_General18_Text & L_Space_Text & iTotal
        wscript.echo L_Text_Msg_General19_Text & L_Space_Text & iTotalDeleted

    else

    end if

    DelAllPrinters = kErrorSuccess

end function

'
' List the printers
'
function ListPrinters(strServer, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg05_Text

    dim Printers
    dim oService
    dim oPrinter
    dim iTotal

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Printers = oService.InstancesOf("Win32_Printer")

    else

        ListPrinters = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General11_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        ListPrinters = kErrorFailure

        exit function

    end if

    iTotal = 0

    for each oPrinter in Printers

        iTotal = iTotal + 1

        wscript.echo L_Empty_Text
        wscript.echo L_Text_Msg_Printer01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Printer02_Text & L_Space_Text & oPrinter.DeviceID
        wscript.echo L_Text_Msg_Printer03_Text & L_Space_Text & oPrinter.ShareName
        wscript.echo L_Text_Msg_Printer04_Text & L_Space_Text & oPrinter.DriverName
        wscript.echo L_Text_Msg_Printer05_Text & L_Space_Text & oPrinter.PortName
        wscript.echo L_Text_Msg_Printer06_Text & L_Space_Text & oPrinter.Comment
        wscript.echo L_Text_Msg_Printer07_Text & L_Space_Text & oPrinter.Location
        wscript.echo L_Text_Msg_Printer08_Text & L_Space_Text & oPrinter.SepFile
        wscript.echo L_Text_Msg_Printer09_Text & L_Space_Text & oPrinter.PrintProcessor
        wscript.echo L_Text_Msg_Printer10_Text & L_Space_Text & oPrinter.PrintJobDataType
        wscript.echo L_Text_Msg_Printer11_Text & L_Space_Text & oPrinter.Parameters
        wscript.echo L_Text_Msg_Printer12_Text & L_Space_Text & CSTR(oPrinter.Attributes)
        wscript.echo L_Text_Msg_Printer13_Text & L_Space_Text & CSTR(oPrinter.Priority)
        wscript.echo L_Text_Msg_Printer14_Text & L_Space_Text & CStr(oPrinter.DefaultPriority)

        if CStr(oPrinter.StartTime) <> "" and CStr(oPrinter.UntilTime) <> "" then

            wscript.echo L_Text_Msg_Printer15_Text & L_Space_Text & Mid(Mid(CStr(oPrinter.StartTime), 9, 4), 1, 2) & "h" & Mid(Mid(CStr(oPrinter.StartTime), 9, 4), 3, 2)
            wscript.echo L_Text_Msg_Printer16_Text & L_Space_Text & Mid(Mid(CStr(oPrinter.UntilTime), 9, 4), 1, 2) & "h" & Mid(Mid(CStr(oPrinter.UntilTime), 9, 4), 3, 2)

        end if

        wscript.echo L_Text_Msg_Printer17_Text & L_Space_Text & CStr(oPrinter.Jobs)
        wscript.echo L_Text_Msg_Printer18_Text & L_Space_Text & CStr(oPrinter.AveragePagesPerMinute)
        wscript.echo L_Text_Msg_Printer19_Text & L_Space_Text & PrnStatusToString(oPrinter.PrinterStatus)
        wscript.echo L_Text_Msg_Printer20_Text & L_Space_Text & ExtPrnStatusToString(oPrinter.ExtendedPrinterStatus)
        wscript.echo L_Text_Msg_Printer21_Text & L_Space_Text & DetectedErrorStateToString(oPrinter.DetectedErrorState)
        wscript.echo L_Text_Msg_Printer22_Text & L_Space_Text & ExtDetectedErrorStateToString(oPrinter.ExtendedDetectedErrorState)

        Err.Clear

    next

    wscript.echo L_Empty_Text
    wscript.echo L_Text_Msg_General09_Text & L_Space_Text & iTotal

    ListPrinters = kErrorSuccess

end function

'
' Get the default printer
'
function GetDefaultPrinter(strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg06_Text

    dim oService
    dim oPrinter
    dim iRetval
    dim oEnum

    iRetval = kErrorFailure

    '
    ' We connect to the local server
    '
    if WmiConnect("", kNameSpace, strUser, strPassword, oService) then

        set oEnum    = oService.ExecQuery("select DeviceID from Win32_Printer where default=true")

    else

        SetDefaultPrinter = kErrorFailure

        exit function

    end if

    if Err.Number = kErrorSuccess then

         for each oPrinter in oEnum

            wscript.echo L_Text_Msg_General12_Text & L_Space_Text & oPrinter.DeviceID

         next

         iRetval = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General13_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

    end if

    GetDefaultPrinter = iRetval

end function

'
' Set the default printer
'
function SetDefaultPrinter(strPrinter, strUser, strPassword)

    'on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg07_Text

    dim oService
    dim oPrinter
    dim iRetval
    dim uResult

    iRetval = kErrorFailure

    '
    ' We connect to the local server
    '
    if WmiConnect("", kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer.DeviceID='" & strPrinter & "'")

    else

        SetDefaultPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        '
        ' The Err object indicates whether the WMI provider reached the execution
        ' of the function that sets the default printer. The uResult is the Win32
        ' error code of the spooler function that sets the default printer
        '
        uResult = oPrinter.SetDefaultPrinter

        if Err.Number = kErrorSuccess then

            if uResult = kErrorSuccess then

                wscript.echo L_Text_Msg_General15_Text & L_Space_Text & strPrinter

                iRetval = kErrorSuccess

            else

                wscript.echo L_Text_Msg_General14_Text & L_Space_Text _
                             & L_Text_Error_General03_Text& L_Space_Text & uResult

            end if

        else

            wscript.echo L_Text_Msg_General14_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                         & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General14_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    SetDefaultPrinter = iRetval

end function

'
' Converts the printer status to a string
'
function PrnStatusToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status03_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status04_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status05_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status06_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

    end select

    PrnStatusToString = str

end function

'
' Converts the extended printer status to a string
'
function ExtPrnStatusToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status03_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status04_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status05_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status06_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

        case 8
            str = str + L_Text_Msg_Status08_Text + L_Space_Text

        case 9
            str = str + L_Text_Msg_Status09_Text + L_Space_Text

        case 10
            str = str + L_Text_Msg_Status10_Text + L_Space_Text

        case 11
            str = str + L_Text_Msg_Status11_Text + L_Space_Text

        case 12
            str = str + L_Text_Msg_Status12_Text + L_Space_Text

        case 13
            str = str + L_Text_Msg_Status13_Text + L_Space_Text

        case 14
            str = str + L_Text_Msg_Status14_Text + L_Space_Text

        case 15
            str = str + L_Text_Msg_Status15_Text + L_Space_Text

        case 16
            str = str + L_Text_Msg_Status16_Text + L_Space_Text

        case 17
            str = str + L_Text_Msg_Status17_Text + L_Space_Text

        case 18
            str = str + L_Text_Msg_Status18_Text + L_Space_Text

    end select

    ExtPrnStatusToString = str

end function

'
' Converts the detected error state to a string
'
function DetectedErrorStateToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 0
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status20_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status21_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status22_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status23_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status24_Text + L_Space_Text

        case 8
            str = str + L_Text_Msg_Status25_Text + L_Space_Text

        case 9
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

        case 10
            str = str + L_Text_Msg_Status26_Text + L_Space_Text

        case 11
            str = str + L_Text_Msg_Status27_Text + L_Space_Text

    end select

    DetectedErrorStateToString = str

end function

'
' Converts the extended detected error state to a string
'
function ExtDetectedErrorStateToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 0
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status20_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status21_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status22_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status23_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status24_Text + L_Space_Text

        case 8
            str = str + L_Text_Msg_Status25_Text + L_Space_Text

        case 9
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

        case 10
            str = str + L_Text_Msg_Status26_Text + L_Space_Text

        case 11
            str = str + L_Text_Msg_Status27_Text + L_Space_Text

        case 12
            str = str + L_Text_Msg_Status28_Text + L_Space_Text

        case 13
            str = str + L_Text_Msg_Status29_Text + L_Space_Text

        case 14
            str = str + L_Text_Msg_Status30_Text + L_Space_Text

        case 15
            str = str + L_Text_Msg_Status31_Text + L_Space_Text

        case 16
            str = str + L_Text_Msg_Status32_Text + L_Space_Text

    end select

    ExtDetectedErrorStateToString = str

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
function ParseCommandLine(iAction, strServer, strPrinter, strDriver, strPort, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg08_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex  = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-a"
                iAction = kActionAdd

            case "-ac"
                iAction = kActionAddConn

            case "-d"
                iAction = kActionDel

            case "-x"
                iAction = kActionDelAll

            case "-xc"
                iAction = kActionDelAllCon

            case "-xo"
                iAction = kActionDelAllLocal

            case "-l"
                iAction = kActionList

            case "-g"
                iAction = kActionGetDefaultPrinter

            case "-t"
                iAction = kActionSetDefaultPrinter

            case "-s"
                iIndex = iIndex + 1
                strServer = RemoveBackslashes(oArgs(iIndex))

            case "-p"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

            case "-m"
                iIndex = iIndex + 1
                strDriver = oArgs(iIndex)

            case "-u"
                iIndex = iIndex + 1
                strUser = oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                strPassword = oArgs(iIndex)

            case "-r"
                iIndex = iIndex + 1
                strPort = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
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
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General20_Text
    wscript.echo L_Help_Help_General21_Text
    wscript.echo L_Help_Help_General22_Text
    wscript.echo L_Help_Help_General23_Text
    wscript.echo L_Help_Help_General24_Text
    wscript.echo L_Help_Help_General25_Text
    wscript.echo L_Help_Help_General26_Text
    wscript.echo L_Help_Help_General27_Text
    wscript.echo L_Help_Help_General28_Text
    wscript.echo L_Help_Help_General29_Text

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
        wscript.echo L_Text_Error_General03_Text & L_Space_Text & oError.StatusCode

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

            wscript.echo L_Text_Msg_General11_Text & L_Space_Text & L_Error_Text _
                         & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                         & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General10_Text & L_Space_Text & L_Error_Text _
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
'' SIG '' MIIhpAYJKoZIhvcNAQcCoIIhlTCCIZECAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' JgK02eO+gRAvE2IMUc2ETsWGPR1isJdKW/vFu8QQ+Hag
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
'' SIG '' n5efy1eAbzOpBM93pGIcWX4xghYaMIIWFgIBATCBnDCB
'' SIG '' hDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEuMCwGA1UEAxMlTWlj
'' SIG '' cm9zb2Z0IFdpbmRvd3MgUHJvZHVjdGlvbiBQQ0EgMjAx
'' SIG '' MQITMwAAAXMwMQcmZbi5swAAAAABczANBglghkgBZQME
'' SIG '' AgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
'' SIG '' AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
'' SIG '' LwYJKoZIhvcNAQkEMSIEIJidaMwGxq7TyObIUNZQSw5+
'' SIG '' hdGd7e3flVX5B7IxiE25MDwGCisGAQQBgjcKAxwxLgws
'' SIG '' bE1EZUd0Zm0veitzQjhnVnFsb2JOdzU0NWJxWURZcTVx
'' SIG '' d0hFSmNPS3hMdz0wWgYKKwYBBAGCNwIBDDFMMEqgJIAi
'' SIG '' AE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcA
'' SIG '' c6EigCBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2lu
'' SIG '' ZG93czANBgkqhkiG9w0BAQEFAASCAQBMtxbwmkIv/PE1
'' SIG '' An4oYnwwWDvCEZSbKqsE9oBU1MObjmx4VKtSNuN3jGs3
'' SIG '' Sshu3cRSIpCyoB5ia3j4YoA4LU/r7IyRgFECcOpuhVM7
'' SIG '' Cen9+3fz9sLsMoMxsb5gTxyMMiSCE8Vb5XHby7V/hIPa
'' SIG '' iZ6D3xexvkwsmVjMj0xTHFlAPh2dlvO60VPtdqNTXKkP
'' SIG '' YxLGP5+lj5WL/eIy6aMHjvbof5Dv2rODae1bxThxgto0
'' SIG '' pbWr5eiKrteEwOwUTMDHSQgirJ5bFXmZEaepmw3vLPVP
'' SIG '' bpv4EbeNbYHjkaSBTOrXPovfOPmQDF18TWJMlu4fs10t
'' SIG '' dSNwbgcILf0j9FNhTTsmoYITRjCCE0IGCisGAQQBgjcD
'' SIG '' AwExghMyMIITLgYJKoZIhvcNAQcCoIITHzCCExsCAQMx
'' SIG '' DzANBglghkgBZQMEAgEFADCCATwGCyqGSIb3DQEJEAEE
'' SIG '' oIIBKwSCAScwggEjAgEBBgorBgEEAYRZCgMBMDEwDQYJ
'' SIG '' YIZIAWUDBAIBBQAEIOWmU9iRHkPhY8OXPlGoC837Lw5Y
'' SIG '' TCI6DULeWRNsmWRxAgZZzUEqlAEYEzIwMTcwOTI5MDQy
'' SIG '' NDMxLjI4NVowBwIBAYACAfSggbikgbUwgbIxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xDDAKBgNVBAsTA0FPQzEnMCUGA1UE
'' SIG '' CxMebkNpcGhlciBEU0UgRVNOOjEyRTctMzA2NC02MTEy
'' SIG '' MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
'' SIG '' ZXJ2aWNloIIOyjCCBnEwggRZoAMCAQICCmEJgSoAAAAA
'' SIG '' AAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVT
'' SIG '' MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
'' SIG '' ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
'' SIG '' YXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENl
'' SIG '' cnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcw
'' SIG '' MTIxMzY1NVoXDTI1MDcwMTIxNDY1NVowfDELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
'' SIG '' bWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3DQEB
'' SIG '' AQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/
'' SIG '' aZRrdFQQ1aUKAIKF++18aEssX8XD5WHCdrc+Zitb8BVT
'' SIG '' JwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRDDNdNuDgI
'' SIG '' s0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKE
'' SIG '' HnRhZ5FfgVSxz5NMksHEpl3RYRNuKMYa+YaAu99h/EbB
'' SIG '' Jx0kZxJyGiGKr0tkiVBisV39dx898Fd1rL2KQk1AUdEP
'' SIG '' nAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2
'' SIG '' iAg16HgcsOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlE
'' SIG '' XV4WnAEFTyJNAgMBAAGjggHmMIIB4jAQBgkrBgEEAYI3
'' SIG '' FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8Uz
'' SIG '' aFqFbVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEw
'' SIG '' CwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYD
'' SIG '' VR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
'' SIG '' VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
'' SIG '' ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2Vy
'' SIG '' QXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4w
'' SIG '' TDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3Nv
'' SIG '' ZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAx
'' SIG '' MC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCBkjCBjwYJ
'' SIG '' KwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8v
'' SIG '' d3d3Lm1pY3Jvc29mdC5jb20vUEtJL2RvY3MvQ1BTL2Rl
'' SIG '' ZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBn
'' SIG '' AGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0A
'' SIG '' ZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohR
'' SIG '' DeLG4Jg/gXEDPZ2joSFvs+umzPUxvs8F4qn++ldtGTCz
'' SIG '' wsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM
'' SIG '' 9GASinbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/
'' SIG '' 8jd9Wj8c8pl5SpFSAK84Dxf1L3mBZdmptWvkx872ynoA
'' SIG '' b0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
'' SIG '' M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOd
'' SIG '' eyFtw5yjojz6f32WapB4pm3S4Zz5Hfw42JT0xqUKloak
'' SIG '' vZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45V3ai
'' SIG '' caoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9d
'' SIG '' dJgiCGHasFAeb73x4QDf5zEHpJM692VHeOj4qEir995y
'' SIG '' fmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEegPsbiSpU
'' SIG '' ObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO
'' SIG '' 2ii4sanblrKnQqLJzxlBTeCG+SqaoxFmMNO7dDJL32N7
'' SIG '' 9ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp3lfB0d4wwP3M
'' SIG '' 5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUe
'' SIG '' CLraNtvTX4/edIhJEjCCBNkwggPBoAMCAQICEzMAAACs
'' SIG '' iiG8etKbcvQAAAAAAKwwDQYJKoZIhvcNAQELBQAwfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMTYwOTA3
'' SIG '' MTc1NjU0WhcNMTgwOTA3MTc1NjU0WjCBsjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIERTRSBFU046MTJFNy0zMDY0LTYxMTIx
'' SIG '' JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
'' SIG '' cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
'' SIG '' AoIBAQChxPQ8pY5zMaEXQVyrdwi3qdFDqrvf3HT5ujVW
'' SIG '' aoQJ2Km7+1T3GNnGpbND6PomAcw9NfV+C+RMmCrThpUl
'' SIG '' BVzeuzsNL2Lsj9mMdK83ixebazenSrA0rXLIifWBzKHV
'' SIG '' P6jzsQWo96cHukHZqI8xRp3tYivgapt5LLrn9Rm2Jn+E
'' SIG '' 0h2lKDOw5sIteZiriOMPH2Z7mtgYXmyB8ThgdB46p6fr
'' SIG '' NGpcXr11pa1Vkldl0iY6oBAKQQSxJ5Bn4N7ui5Wj5wDk
'' SIG '' DZzGAg6n1ptMPTPJhL2uosW84YjnSp/2suNap3qOjKEY
'' SIG '' XmpGzvasq5qyqPyvfqfksOfNBaJntfpC8dIDJKrnAgMB
'' SIG '' AAGjggEbMIIBFzAdBgNVHQ4EFgQU2EdwqIU1ixNhr4eQ
'' SIG '' 6EUXJOCYpw0wHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xG
'' SIG '' G8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
'' SIG '' L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
'' SIG '' dHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoG
'' SIG '' CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
'' SIG '' L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNU
'' SIG '' aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/
'' SIG '' BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
'' SIG '' 9w0BAQsFAAOCAQEASOzoXifDGxicXbTOdm43DB9dYxwN
'' SIG '' XaQj0hW0ztBACeWbX6zxee1w7TJeXQx1IfTz1BWJAfVl
'' SIG '' a7HA1oxa0LiF/Uf6rfRvEEmGmHr9wWEbr3xiErllTFE2
'' SIG '' V/mdYBSEwdj74m8QLBeFY37Cjx4TFe+AB/FQly3kvfrJ
'' SIG '' KDaYYTTXTAFYAKpi0AcDTcESXZcshJ/O8UZs9fr0BgrO
'' SIG '' m5h7qeQ1CJNmDMVEqElQt/cFO3dxqrAKv3Fu/nsT5GkA
'' SIG '' Bu+vIibgxX6tNmqccIIXKALgv7zDIaRAAWtXV9Lwnmst
'' SIG '' DUbIp8dH/oSVm3tPnCPlrB3+C28vPnvJdbtJi/yOuETd
'' SIG '' +rLn+qGCA3QwggJcAgEBMIHioYG4pIG1MIGyMQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMQwwCgYDVQQLEwNBT0MxJzAlBgNV
'' SIG '' BAsTHm5DaXBoZXIgRFNFIEVTTjoxMkU3LTMwNjQtNjEx
'' SIG '' MjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' U2VydmljZaIlCgEBMAkGBSsOAwIaBQADFQA5cCWLFMh5
'' SIG '' 3V8MXemLnLOUmfcct6CBwTCBvqSBuzCBuDELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIE5UUyBFU046MjY2NS00QzNGLUM1REUx
'' SIG '' KzApBgNVBAMTIk1pY3Jvc29mdCBUaW1lIFNvdXJjZSBN
'' SIG '' YXN0ZXIgQ2xvY2swDQYJKoZIhvcNAQEFBQACBQDdeD2v
'' SIG '' MCIYDzIwMTcwOTI5MDMzNDA3WhgPMjAxNzA5MzAwMzM0
'' SIG '' MDdaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAN14Pa8C
'' SIG '' AQAwBwIBAAICGSEwBwIBAAICGRIwCgIFAN15jy8CAQAw
'' SIG '' NgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAaAK
'' SIG '' MAgCAQACAxbjYKEKMAgCAQACAwehIDANBgkqhkiG9w0B
'' SIG '' AQUFAAOCAQEAJ2fh5rW+yNAjc91KheNOpMisljLpCr7z
'' SIG '' tcAVzudmeYEIWzWYaiMAASteux0ndPpWw5NHt+QgH/5L
'' SIG '' eWEHXTNJTxp0v6jepOCSxuMDwqrt18Ly/aX+CdfrfWTI
'' SIG '' ShbMlzwW65O2cdat84JwOx8DkPNuPl37dhLwApnROKF2
'' SIG '' 1MG8dy1+WMFsgzMU3pXYRPTuuc2EJy8ZqI+67yGBNzvY
'' SIG '' +Svsb0Q9LnV1uzDGLe+aFl5lEKEU97XfDSba0Tkem0xp
'' SIG '' 4TxorQN2EbMJ+oyImyBoosNVjm3/t1OVhtc052rh6Jsj
'' SIG '' xoweXMwH2+Rm8MVMxH6+rthTwpsHvNVCwfFUU3F2IbTF
'' SIG '' SzGCAvUwggLxAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
'' SIG '' IFBDQSAyMDEwAhMzAAAArIohvHrSm3L0AAAAAACsMA0G
'' SIG '' CWCGSAFlAwQCAQUAoIIBMjAaBgkqhkiG9w0BCQMxDQYL
'' SIG '' KoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEINpbgTgU
'' SIG '' ouQFOuNc8ooVt3GACgliSK91oicXAY/OjQ3LMIHiBgsq
'' SIG '' hkiG9w0BCRACDDGB0jCBzzCBzDCBsQQUOXAlixTIed1f
'' SIG '' DF3pi5yzlJn3HLcwgZgwgYCkfjB8MQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
'' SIG '' dGFtcCBQQ0EgMjAxMAITMwAAAKyKIbx60pty9AAAAAAA
'' SIG '' rDAWBBSYg73TN71BhWa5TJZNN+fdAoiA/DANBgkqhkiG
'' SIG '' 9w0BAQsFAASCAQChv/DIGtqr1wA67AMaoTNptzbqTJHx
'' SIG '' mLVNtXzAP17c1HGl8ORc/gY+LqHIhL/9OY/Lh819U1iU
'' SIG '' guMPNQbpdj1M1HVtJ4dAz4dASDxiOWirisdnVpnuS4kJ
'' SIG '' FTCHYHUzNhN6j8SmsJDDqedsxd8fy/MwnAJ9MaQRjq2N
'' SIG '' RuNM825GeMHKeMbfcsNniTh5LJGhYYsXpvi259ORfP3d
'' SIG '' l2+YO7XPn8JJD7YnGDQIKwVzWcbbDGyxJA9TmZHwLi1Z
'' SIG '' pwvmKeUeXAj0UmPPU8/dBuSG1z2YO63A7MhyOAqaUA6s
'' SIG '' IPeDmWL5EUlM21iqewCAOxGD+8PIjR4/Q7GEAcpdlye6
'' SIG '' Lcv6
'' SIG '' End signature block
