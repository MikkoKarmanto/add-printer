'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
'
' prncnfg.vbs - printer configuration script for WMI on Windows used to get
'     and set printer configuration also used to rename a printer
'
' Usage:
' prncnfg [-gtx?] [-s server] [-p printer] [-u user name] [-w password]
'                 [-z new printer name] [-r port name] [-l location] [-m comment]
'                 [-h share name] [-f sep-file] [-y data-type] [-st start time]
'                 [-ut until time] [-o priority] [-i default priority]
'                 [<+|->rawonly][<+|->keepprintedjobs][<+|->queued][<+|->workoffline]
'                 [<+|->enabledevq][<+|->docompletefirst][<+|->enablebidi]
'
' Examples:
' prncnfg -g -s server -p printer
' prncnfg -x -p printer -w "new Printer"
' prncnfg -t -s server -p Printer -l "Building A/Floor 100/Office 1" -m "Color Printer"
' prncnfg -t -p printer -h "Share" +shared -direct
' prncnfg -t -p printer +rawonly +keepprintedjobs
' prncnfg -t -p printer -st 2300 -ut 0215 -o 10 -i 5
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

const kFlagUpdateOnly = 1

'
' Operation action values.
'
const kActionUnknown   = 0
const kActionSet       = 1
const kActionGet       = 2
const kActionRename    = 3

const kErrorSuccess    = 0
const kErrorFailure    = 1

'
' Constants for the parameter dictionary
'
const kServerName      = 1
const kPrinterName     = 2
const kNewPrinterName  = 3
const kShareName       = 4
const kPortName        = 5
const kDriverName      = 6
const kComment         = 7
const kLocation        = 8
const kSepFile         = 9
const kPrintProc       = 10
const kDataType        = 11
const kParameters      = 12
const kPriority        = 13
const kDefaultPriority = 14
const kStartTime       = 15
const kUntilTime       = 16
const kQueued          = 17
const kDirect          = 18
const kDefault         = 19
const kShared          = 20
const kNetwork         = 21
const kHidden          = 23
const kLocal           = 24
const kEnableDevq      = 25
const kKeepPrintedJobs = 26
const kDoCompleteFirst = 27
const kWorkOffline     = 28
const kEnableBidi      = 29
const kRawOnly         = 30
const kPublished       = 31
const kUserName        = 32
const kPassword        = 33

const kNameSpace       = "root\cimv2"

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

'
' General usage messages
'
const L_Help_Help_General01_Text   = "Usage: prncnfg [-gtx?] [-s server][-p printer][-z new printer name]"
const L_Help_Help_General02_Text   = "               [-u user name][-w password][-r port name][-l location]"
const L_Help_Help_General03_Text   = "               [-m comment][-h share name][-f sep file][-y datatype]"
const L_Help_Help_General04_Text   = "               [-st start time][-ut until time][-i default priority]"
const L_Help_Help_General05_Text   = "               [-o priority][<+|->shared][<+|->direct][<+|->hidden]"
const L_Help_Help_General06_Text   = "               [<+|->published][<+|->rawonly][<+|->queued][<+|->enablebidi]"
const L_Help_Help_General07_Text   = "               [<+|->keepprintedjobs][<+|->workoffline][<+|->enabledevq]"
const L_Help_Help_General08_Text   = "               [<+|->docompletefirst]"
const L_Help_Help_General09_Text   = "Arguments:"
const L_Help_Help_General10_Text   = "-f     - separator file name"
const L_Help_Help_General11_Text   = "-g     - get configuration"
const L_Help_Help_General12_Text   = "-h     - share name"
const L_Help_Help_General13_Text   = "-i     - default priority"
const L_Help_Help_General14_Text   = "-l     - location string"
const L_Help_Help_General15_Text   = "-m     - comment string"
const L_Help_Help_General16_Text   = "-o     - priority"
const L_Help_Help_General17_Text   = "-p     - printer name"
const L_Help_Help_General18_Text   = "-r     - port name"
const L_Help_Help_General19_Text   = "-s     - server name"
const L_Help_Help_General20_Text   = "-st    - start time"
const L_Help_Help_General21_Text   = "-t     - set configuration"
const L_Help_Help_General22_Text   = "-u     - user name"
const L_Help_Help_General23_Text   = "-ut    - until time"
const L_Help_Help_General24_Text   = "-w     - password"
const L_Help_Help_General25_Text   = "-x     - change printer name"
const L_Help_Help_General26_Text   = "-y     - data type string"
const L_Help_Help_General27_Text   = "-z     - new printer name"
const L_Help_Help_General28_Text   = "-?     - display command usage"
const L_Help_Help_General29_Text   = "Examples:"
const L_Help_Help_General30_Text   = "prncnfg -g -s server -p printer"
const L_Help_Help_General31_Text   = "prncnfg -x -s server -p printer -z ""new printer"""
const L_Help_Help_General32_Text   = "prncnfg -t -p printer -l ""Building A/Floor 100/Office 1"" -m ""Color Printer"""
const L_Help_Help_General33_Text   = "prncnfg -t -p printer -h ""Share"" +shared -direct"
const L_Help_Help_General34_Text   = "prncnfg -t -p printer +rawonly +keepprintedjobs"
const L_Help_Help_General35_Text   = "prncnfg -t -p printer -st 2300 -ut 0215 -o 1 -i 5"

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
const L_Text_Msg_General01_Text    = "Renamed printer"
const L_Text_Msg_General02_Text    = "New printer name"
const L_Text_Msg_General03_Text    = "Unable to rename printer"
const L_Text_Msg_General04_Text    = "Unable to get configuration for printer"
const L_Text_Msg_General05_Text    = "Printer always available"
const L_Text_Msg_General06_Text    = "Configured printer"
const L_Text_Msg_General07_Text    = "Unable to configure printer"
const L_Text_Msg_General08_Text    = "Unable to get SWbemLocator object"
const L_Text_Msg_General09_Text    = "Unable to connect to WMI service"
const L_Text_Msg_General10_Text    = "Printer status"
const L_Text_Msg_General11_Text    = "Extended printer status"
const L_Text_Msg_General12_Text    = "Detected error state"
const L_Text_Msg_General13_Text    = "Extended detected error state"

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
const L_Text_Msg_Printer17_Text    = "Status"
const L_Text_Msg_Printer18_Text    = "Job count"
const L_Text_Msg_Printer19_Text    = "Average pages per minute"

'
' Printer attributes
'
const L_Text_Msg_Attrib01_Text     = "direct"
const L_Text_Msg_Attrib02_Text     = "raw_only"
const L_Text_Msg_Attrib03_Text     = "local"
const L_Text_Msg_Attrib04_Text     = "shared"
const L_Text_Msg_Attrib05_Text     = "keep_printed_jobs"
const L_Text_Msg_Attrib06_Text     = "published"
const L_Text_Msg_Attrib07_Text     = "queued"
const L_Text_Msg_Attrib08_Text     = "default"
const L_Text_Msg_Attrib09_Text     = "network"
const L_Text_Msg_Attrib10_Text     = "enable_bidi"
const L_Text_Msg_Attrib11_Text     = "do_complete_first"
const L_Text_Msg_Attrib12_Text     = "work_offline"
const L_Text_Msg_Attrib13_Text     = "hidden"
const L_Text_Msg_Attrib14_Text     = "enable_devq_print"

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
const L_Text_Dbg_Msg01_Text        = "In function RenamePrinter"
const L_Text_Dbg_Msg02_Text        = "New printer name"
const L_Text_Dbg_Msg03_Text        = "In function GetPrinter"
const L_Text_Dbg_Msg04_Text        = "In function SetPrinter"
const L_Text_Dbg_Msg05_Text        = "In function ParseCommandLine"

main

'
' Main execution starts here
'
sub main

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

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionSet
                 iRetval = SetPrinter(oParamDict)

            case kActionGet
                 iRetval = GetPrinter(oParamDict)

            case kActionRename
                 iRetval = RenamePrinter(oParamDict)

            case else
                 Usage(True)
                 exit sub

        end select

    end if

end sub

'
' Rename printer
'
function RenamePrinter(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & oParamDict.Item(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & oParamDict.Item(kPrinterName)
    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text & L_Space_Text & oParamDict.Item(kNewPrinterName)

    dim oPrinter
    dim oService
    dim iRetval
    dim uResult
    dim strServer
    dim strPrinter
    dim strNewName
    dim strUser
    dim strPassword

    iRetval = kErrorFailure

    strServer   = oParamDict.Item(kServerName)
    strPrinter  = oParamDict.Item(kPrinterName)
    strNewName  = oParamDict.Item(kNewPrinterName)
    strUser     = oParamDict.Item(kUserName)
    strPassword = oParamDict.Item(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer.DeviceID='" & strPrinter & "'")

    else

        RenamePrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        uResult = oPrinter.RenamePrinter(strNewName)

        if Err.Number = kErrorSuccess then

            if uResult = kErrorSuccess then

                wscript.echo L_Text_Msg_General01_Text & L_Space_Text & strPrinter
                wscript.echo L_Text_Msg_General02_Text & L_Space_Text & strNewName

                iRetval = kErrorSuccess

            else

                wscript.echo L_Text_Msg_General03_Text & L_Space_Text & strPrinter & L_Space_Text _
                             & L_Text_Error_General03_Text & L_Space_Text & uResult

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

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    RenamePrinter = iRetval

end function

'
' Get printer configuration
'
function GetPrinter(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & oParamDict.Item(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & oParamDict.Item(kPrinterName)

    dim oPrinter
    dim oService
    dim iRetval
    dim uResult
    dim strServer
    dim strPrinter
    dim strAttributes
    dim strStart
    dim strEnd
    dim strUser
    dim strPassword

    iRetval = kErrorFailure

    strServer  = oParamDict.Item(kServerName)
    strPrinter = oParamDict.Item(kPrinterName)
    strUser     = oParamDict.Item(kUserName)
    strPassword = oParamDict.Item(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer='" & strPrinter & "'")

    else

        GetPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        wscript.echo L_Text_Msg_Printer01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Printer02_Text & L_Space_Text & oPrinter.DeviceID
        wscript.echo L_Text_Msg_Printer03_Text & L_Space_Text & oPrinter.ShareName
        wscript.echo L_Text_Msg_Printer04_Text & L_Space_Text & oPrinter.DriverName
        wscript.echo L_Text_Msg_Printer05_Text & L_Space_Text & oPrinter.PortName
        wscript.echo L_Text_Msg_Printer06_Text & L_Space_Text & oPrinter.Comment
        wscript.echo L_Text_Msg_Printer07_Text & L_Space_Text & oPrinter.Location
        wscript.echo L_Text_Msg_Printer08_Text & L_Space_Text & oPrinter.SeparatorFile
        wscript.echo L_Text_Msg_Printer09_Text & L_Space_Text & oPrinter.PrintProcessor
        wscript.echo L_Text_Msg_Printer10_Text & L_Space_Text & oPrinter.PrintJobDatatype
        wscript.echo L_Text_Msg_Printer11_Text & L_Space_Text & oPrinter.Parameters
        wscript.echo L_Text_Msg_Printer13_Text & L_Space_Text & CStr(oPrinter.Priority)
        wscript.echo L_Text_Msg_Printer14_Text & L_Space_Text & CStr(oPrinter.DefaultPriority)

        strStart = Mid(CStr(oPrinter.StartTime), 9, 4)
        strEnd = Mid(CStr(oPrinter.UntilTime), 9, 4)

        if strStart <> "" and strEnd <> "" then

            wscript.echo L_Text_Msg_Printer15_Text & L_Space_Text & Mid(strStart, 1, 2) & "h" & Mid(strStart, 3, 2)
            wscript.echo L_Text_Msg_Printer16_Text & L_Space_Text & Mid(strEnd, 1, 2) & "h" & Mid(strEnd, 3, 2)

        else

            wscript.echo L_Text_Msg_General05_Text

        end if

        strAttributes = L_Text_Msg_Printer12_Text

        if oPrinter.Direct then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib01_Text

        end if

        if oPrinter.RawOnly then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib02_Text

        end if

        if oPrinter.Local then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib03_Text

        end if

        if oPrinter.Shared then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib04_Text

        end if

        if oPrinter.KeepPrintedJobs then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib05_Text

        end if

        if oPrinter.Published then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib06_Text

        end if

        if oPrinter.Queued then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib07_Text

        end if

        if oPrinter.Default then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib08_Text

        end if

        if oPrinter.Network then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib09_Text

        end if

        if oPrinter.EnableBiDi then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib10_Text

        end if

        if oPrinter.DoCompleteFirst then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib11_Text

        end if

        if oPrinter.WorkOffline then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib12_Text

        end if

        if oPrinter.Hidden then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib13_Text

        end if

        if oPrinter.EnableDevQueryPrint then

            strAttributes = strAttributes + L_Space_Text + L_Text_Msg_Attrib14_Text

        end if

        wscript.echo strAttributes
        wscript.echo
        wscript.echo L_Text_Msg_General10_Text & L_Space_Text & PrnStatusToString(oPrinter.PrinterStatus)
        wscript.echo L_Text_Msg_General11_Text & L_Space_Text & ExtPrnStatusToString(oPrinter.ExtendedPrinterStatus)
        wscript.echo L_Text_Msg_General12_Text & L_Space_Text & DetectedErrorStateToString(oPrinter.DetectedErrorState)
        wscript.echo L_Text_Msg_General13_Text & L_Space_Text & ExtDetectedErrorStateToString(oPrinter.ExtendedDetectedErrorState)

        iRetval = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General04_Text & L_Space_Text & oParamDict.Item(kPrinterName) & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                     & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    GetPrinter = iRetval

end function

'
' Configure a printer
'
function SetPrinter(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & oParamDict.Item(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & oParamDict.Item(kPrinterName)

    dim oPrinter
    dim oService
    dim iRetval
    dim uResult
    dim strServer
    dim strPrinter
    dim strUser
    dim strPassword

    iRetval = kErrorFailure

    strServer   = oParamDict.Item(kServerName)
    strPrinter  = oParamDict.Item(kPrinterName)
    strNewName  = oParamDict.Item(kNewPrinterName)
    strUser     = oParamDict.Item(kUserName)
    strPassword = oParamDict.Item(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer='" & strPrinter & "'")

    else

        SetPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        if oParamdict.Exists(kPortName)        then oPrinter.PortName            = oParamDict.Item(kPortName)        end if
        if oParamdict.Exists(kDriverName)      then oPrinter.DriverName          = oParamDict.Item(kDriverName)      end if
        if oParamdict.Exists(kShareName)       then oPrinter.ShareName           = oParamDict.Item(kShareName)       end if
        if oParamdict.Exists(kLocation)        then oPrinter.Location            = oParamDict.Item(kLocation)        end if
        if oParamdict.Exists(kComment)         then oPrinter.Comment             = oParamDict.Item(kComment)         end if
        if oParamdict.Exists(kDataType)        then oPrinter.PrintJobDataType    = oParamDict.Item(kDataType)        end if
        if oParamdict.Exists(kSepFile)         then oPrinter.SeparatorFile       = oParamDict.Item(kSepfile)         end if
        if oParamdict.Exists(kParameters)      then oPrinter.Parameters          = oParamDict.Item(kParameters)      end if
        if oParamdict.Exists(kPriority)        then oPrinter.Priority            = oParamDict.Item(kPriority)        end if
        if oParamdict.Exists(kDefaultPriority) then oPrinter.DefaultPriority     = oParamDict.Item(kDefaultPriority) end if
        if oParamdict.Exists(kPrintProc)       then oPrinter.PrintProc           = oParamDict.Item(kPrintProc)       end if
        if oParamdict.Exists(kStartTime)       then oPrinter.StartTime           = oParamDict.Item(kStartTime)       end if
        if oParamdict.Exists(kUntilTime)       then oPrinter.UntilTime           = oParamDict.Item(kUntilTime)       end if
        if oParamdict.Exists(kQueued)          then oPrinter.Queued              = oParamDict.Item(kQueued)          end if
        if oParamdict.Exists(kDirect)          then oPrinter.Direct              = oParamDict.Item(kDirect)          end if
        if oParamdict.Exists(kShared)          then oPrinter.Shared              = oParamDict.Item(kShared)          end if
        if oParamdict.Exists(kHidden)          then oPrinter.Hidden              = oParamDict.Item(kHidden)          end if
        if oParamdict.Exists(kEnabledevq)      then oPrinter.EnableDevQueryPrint = oParamDict.Item(kEnabledevq)      end if
        if oParamdict.Exists(kKeepPrintedJobs) then oPrinter.KeepPrintedJobs     = oParamDict.Item(kKeepPrintedJobs) end if
        if oParamdict.Exists(kDoCompleteFirst) then oPrinter.DoCompleteFirst     = oParamDict.Item(kDoCompleteFirst) end if
        if oParamdict.Exists(kWorkOffline)     then oPrinter.WorkOffline         = oParamDict.Item(kWorkOffline)     end if
        if oParamdict.Exists(kEnableBidi)      then oPrinter.EnableBidi          = oParamDict.Item(kEnableBidi)      end if
        if oParamdict.Exists(kRawonly)         then oPrinter.RawOnly             = oParamDict.Item(kRawonly)         end if
        if oParamdict.Exists(kPublished)       then oPrinter.Published           = oParamDict.Item(kPublished)       end if

        oPrinter.Put_(kFlagUpdateOnly)

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General06_Text & L_Space_Text & strPrinter

            iRetval = kErrorSuccess

        else

            wscript.echo L_Text_Msg_General07_Text & L_Space_Text & strPrinter & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                         & Err.Description

            '
            ' Try getting extended error information
            '
            call LastError()

        end if

    else

        wscript.echo L_Text_Msg_General04_Text & L_Space_Text & strPrinter & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                     & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    SetPrinter = iRetval

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
function ParseCommandLine(iAction, oParamdict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg05_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-g"
                iAction = kActionGet

            case "-t"
                iAction = kActionSet

            case "-x"
                iAction = kActionRename

            case "-p"
                iIndex = iIndex + 1
                oParamdict.Add kPrinterName, oArgs(iIndex)

            case "-s"
                iIndex = iIndex + 1
                oParamdict.Add kServerName, RemoveBackslashes(oArgs(iIndex))

            case "-r"
                iIndex = iIndex + 1
                oParamdict.Add kPortName, oArgs(iIndex)

            case "-h"
                iIndex = iIndex + 1
                oParamdict.Add kShareName, oArgs(iIndex)

            case "-m"
                iIndex = iIndex + 1
                oParamdict.Add kComment, oArgs(iIndex)

            case "-l"
                iIndex = iIndex + 1
                oParamdict.Add kLocation, oArgs(iIndex)

            case "-y"
                iIndex = iIndex + 1
                oParamdict.Add kDataType, oArgs(iIndex)

            case "-f"
                iIndex = iIndex + 1
                oParamdict.Add kSepFile, oArgs(iIndex)

            case "-z"
                iIndex = iIndex + 1
                oParamdict.Add kNewPrinterName, oArgs(iIndex)

            case "-u"
                iIndex = iIndex + 1
                oParamdict.Add kUserName, oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                oParamdict.Add kPassword, oArgs(iIndex)

            case "-st"
                iIndex = iIndex + 1
                oParamdict.Add kStartTime, "********" & oArgs(iIndex) & "00.000000+000"

            case "-o"
                iIndex = iIndex + 1
                oParamdict.Add kPriority, oArgs(iIndex)

            case "-i"
                iIndex = iIndex + 1
                oParamdict.Add kDefaultPriority, oArgs(iIndex)

            case "-ut"
                iIndex = iIndex + 1
                oParamdict.Add kUntilTime, "********" & oArgs(iIndex) & "00.000000+000"

            case "-queued"
                oParamdict.Add kQueued, false

            case "+queued"
                oParamdict.Add kQueued, true

            case "-direct"
                oParamdict.Add kDirect, false

            case "+direct"
                oParamdict.Add kDirect, true

            case "-shared"
                oParamdict.Add kShared, false

            case "+shared"
                oParamdict.Add kShared, true

            case "-hidden"
                oParamdict.Add kHidden, false

            case "+hidden"
                oParamdict.Add kHidden, true

            case "-enabledevq"
                oParamdict.Add kEnabledevq, false

            case "+enabledevq"
                oParamdict.Add kEnabledevq, true

            case "-keepprintedjobs"
                oParamdict.Add kKeepprintedjobs, false

            case "+keepprintedjobs"
                oParamdict.Add kKeepprintedjobs, true

            case "-docompletefirst"
                oParamdict.Add kDocompletefirst, false

            case "+docompletefirst"
                oParamdict.Add kDocompletefirst, true

            case "-workoffline"
                oParamdict.Add kWorkoffline, false

            case "+workoffline"
                oParamdict.Add kWorkoffline, true

            case "-enablebidi"
                oParamdict.Add kEnablebidi, false

            case "+enablebidi"
                oParamdict.Add kEnablebidi, true

            case "-rawonly"
                oParamdict.Add kRawonly, false

            case "+rawonly"
                oParamdict.Add kRawonly, true

            case "-published"
                oParamdict.Add kPublished, false

            case "+published"
                oParamdict.Add kPublished, true

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

end function

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
    wscript.echo L_Help_Help_General23_Text
    wscript.echo L_Help_Help_General24_Text
    wscript.echo L_Help_Help_General25_Text
    wscript.echo L_Help_Help_General26_Text
    wscript.echo L_Help_Help_General27_Text
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General28_Text
    wscript.echo L_Help_Help_General29_Text
    wscript.echo L_Help_Help_General30_Text
    wscript.echo L_Help_Help_General31_Text
    wscript.echo L_Help_Help_General32_Text
    wscript.echo L_Help_Help_General33_Text
    wscript.echo L_Help_Help_General34_Text
    wscript.echo L_Help_Help_General35_Text

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

            wscript.echo L_Text_Msg_General08_Text & L_Space_Text & L_Error_Text _
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
'' SIG '' MIIhpQYJKoZIhvcNAQcCoIIhljCCIZICAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' ie1efOwNeF5D8l67XxhZUjMjDFmogSjxo1MJEQP87/Wg
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
'' SIG '' n5efy1eAbzOpBM93pGIcWX4xghYbMIIWFwIBATCBnDCB
'' SIG '' hDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEuMCwGA1UEAxMlTWlj
'' SIG '' cm9zb2Z0IFdpbmRvd3MgUHJvZHVjdGlvbiBQQ0EgMjAx
'' SIG '' MQITMwAAAXMwMQcmZbi5swAAAAABczANBglghkgBZQME
'' SIG '' AgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
'' SIG '' AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
'' SIG '' LwYJKoZIhvcNAQkEMSIEIEHXr2fNkPIz71v/XM898g/j
'' SIG '' o6d0nlVsrO6ai8Bf2Tl8MDwGCisGAQQBgjcKAxwxLgws
'' SIG '' RDlHUDNYOHVMTC9BRFFEMUNndGpodDVRdnNUcTRPMlRT
'' SIG '' c0IwbHFvN1ZKMD0wWgYKKwYBBAGCNwIBDDFMMEqgJIAi
'' SIG '' AE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcA
'' SIG '' c6EigCBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2lu
'' SIG '' ZG93czANBgkqhkiG9w0BAQEFAASCAQAolgIIytgnKIiX
'' SIG '' FA1BLzI2Qnu3CgWjKkEbgvK/s/2jp3kXrZqYuvSs4lpi
'' SIG '' O9dFynzr+5ORc677iutcATRaj0sH7Anz6YMGOd2xbERE
'' SIG '' BNhGcpX84Nau0Cx7GZLJTyW7GFf6kS014yaVBDRY3AX9
'' SIG '' Q1CywRT5OSe/NpePwUb4OOfGAqkSh5oUuk8Q1t6kKhU8
'' SIG '' Excrw5eHn1960RIkJOEhc0CcRcOUqENclkgP4BZJq765
'' SIG '' SzcANTccgAwwyILSiy83WbrVacjPUr1uupn+jfTWae1O
'' SIG '' 1WEmstshn8zkJz/Cp23m/2pc2cMKwFdjZpHCTf9ym/6H
'' SIG '' a2GDMqylOm02nE8TcNYWoYITRzCCE0MGCisGAQQBgjcD
'' SIG '' AwExghMzMIITLwYJKoZIhvcNAQcCoIITIDCCExwCAQMx
'' SIG '' DzANBglghkgBZQMEAgEFADCCATwGCyqGSIb3DQEJEAEE
'' SIG '' oIIBKwSCAScwggEjAgEBBgorBgEEAYRZCgMBMDEwDQYJ
'' SIG '' YIZIAWUDBAIBBQAEIGsYORi8iSgfBMH3QN8tIH9kpWfQ
'' SIG '' Pfvw38HLbYpTPb8VAgZZzcgwAtIYEzIwMTcwOTI5MDQy
'' SIG '' NTM0Ljk3N1owBwIBAYACAfSggbikgbUwgbIxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xDDAKBgNVBAsTA0FPQzEnMCUGA1UE
'' SIG '' CxMebkNpcGhlciBEU0UgRVNOOkQyMzYtMzdEQS05NzYx
'' SIG '' MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
'' SIG '' ZXJ2aWNloIIOyzCCBnEwggRZoAMCAQICCmEJgSoAAAAA
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
'' SIG '' CLraNtvTX4/edIhJEjCCBNkwggPBoAMCAQICEzMAAACu
'' SIG '' DtZOlonbAPUAAAAAAK4wDQYJKoZIhvcNAQELBQAwfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMTYwOTA3
'' SIG '' MTc1NjU1WhcNMTgwOTA3MTc1NjU1WjCBsjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIERTRSBFU046RDIzNi0zN0RBLTk3NjEx
'' SIG '' JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
'' SIG '' cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
'' SIG '' AoIBAQDeki/DpJVy9T4NZmTD+uboIg90jE3Bnse2VLjx
'' SIG '' j059H/tGML58y3ue28RnWJIv+lSABp+jPp8XIf2p//DK
'' SIG '' Yb0o/QSOJ8kGUoFYesNTPtqyf/qohLW1rcLijiFoMLAB
'' SIG '' H/GDnDbgRZHxVFxHUG+KNwffdC0BYC3Vfq3+2uOO8czR
'' SIG '' lj10gRHU2BK8moSz53Vo2ZwF3TMZyVgvAvlg5sarNgRw
'' SIG '' AYwbwWW5wEqpeODFX1VA/nAeLkjirCmg875M1XiEyPtr
'' SIG '' XDAFLng5/y5MlAcUMYJ6dHuSBDqLLXipjjYakQopB3H1
'' SIG '' +9s8iyDoBM07JqP9u55VP5a2n/IZFNNwJHeCTSvLAgMB
'' SIG '' AAGjggEbMIIBFzAdBgNVHQ4EFgQUfo/lNDREi/J5QLjG
'' SIG '' oNGcQx4hJbEwHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xG
'' SIG '' G8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
'' SIG '' L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
'' SIG '' dHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoG
'' SIG '' CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
'' SIG '' L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNU
'' SIG '' aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/
'' SIG '' BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
'' SIG '' 9w0BAQsFAAOCAQEAPVlNePD0XDQI0bVBYANTDPmMpk3l
'' SIG '' Ih6gPIilg0hKQpZNMADLbmj+kav0GZcxtWnwrBoR+fpB
'' SIG '' suaowWgwxExCHBo6mix7RLeJvNyNYlCk2JQT/Ga80SRV
'' SIG '' zOAL5Nxls1PqvDbgFghDcRTmpZMvADfqwdu5R6FNyIge
'' SIG '' cYNoyb7A4AqCLfV1Wx3PrPyaXbatskk5mT8NqWLYLshB
'' SIG '' zt2Ca0bhJJZf6qQwg6r2gz1pG15ue6nDq/mjYpTmCDhY
'' SIG '' z46b8rxrIn0sQxnFTmtntvz2Z1jCGs99n1rr2ZFrGXOJ
'' SIG '' S4Bhn1tyKEFwGJjrfQ4Gb2pyA9aKRwUyK9BHLKWC5ZLD
'' SIG '' 0hAaIKGCA3UwggJdAgEBMIHioYG4pIG1MIGyMQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMQwwCgYDVQQLEwNBT0MxJzAlBgNV
'' SIG '' BAsTHm5DaXBoZXIgRFNFIEVTTjpEMjM2LTM3REEtOTc2
'' SIG '' MTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' U2VydmljZaIlCgEBMAkGBSsOAwIaBQADFQDHwb0we6UY
'' SIG '' nmReZ3Q2+rvjmbxo+6CBwTCBvqSBuzCBuDELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIE5UUyBFU046MjY2NS00QzNGLUM1REUx
'' SIG '' KzApBgNVBAMTIk1pY3Jvc29mdCBUaW1lIFNvdXJjZSBN
'' SIG '' YXN0ZXIgQ2xvY2swDQYJKoZIhvcNAQEFBQACBQDdeEav
'' SIG '' MCIYDzIwMTcwOTI5MDQxMjMxWhgPMjAxNzA5MzAwNDEy
'' SIG '' MzFaMHUwOwYKKwYBBAGEWQoEATEtMCswCgIFAN14Rq8C
'' SIG '' AQAwCAIBAAIDAhp4MAcCAQACAhhHMAoCBQDdeZgvAgEA
'' SIG '' MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwGg
'' SIG '' CjAIAgEAAgMW42ChCjAIAgEAAgMehIAwDQYJKoZIhvcN
'' SIG '' AQEFBQADggEBAHJeydCP1OfTwxBq6+snQ7pB7gzCQiT+
'' SIG '' 0pzZtMhFmABJwXiWxphrtOftO1fhPqnS0P7xDx2lXCaa
'' SIG '' zU4e6rxY1qQ65gzy23/UyUbYGbVVFrVXp7j/4ZcAOoQj
'' SIG '' 0LeS/XO1RYCtla7WBLpt67wdj/wPJZA/h6DCLXbKYSzz
'' SIG '' g+Zp4FG1rVZkev6xKbDQDVmcn+1TljvpQcYhETIDw6N1
'' SIG '' mn0rJ1P2tSbZmmq394ZnT/DtIWza17l1RMyNEU+W8Rw0
'' SIG '' gkMw9zTJsXdRxwnMe4XCDr7bxMNtJYl2gJu9MS9W9Xux
'' SIG '' 3AucMEMmePj9hEPHoyiWTrgZlJZLr8eNq72hj3KNavUa
'' SIG '' 1s0xggL1MIIC8QIBATCBkzB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMAITMwAAAK4O1k6WidsA9QAAAAAArjAN
'' SIG '' BglghkgBZQMEAgEFAKCCATIwGgYJKoZIhvcNAQkDMQ0G
'' SIG '' CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCD8hPbo
'' SIG '' RzHfihHiOZd5mUb+2T/QO57D4dgmF6fUzbIATjCB4gYL
'' SIG '' KoZIhvcNAQkQAgwxgdIwgc8wgcwwgbEEFMfBvTB7pRie
'' SIG '' ZF5ndDb6u+OZvGj7MIGYMIGApH4wfDELMAkGA1UEBhMC
'' SIG '' VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
'' SIG '' B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
'' SIG '' b3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
'' SIG '' U3RhbXAgUENBIDIwMTACEzMAAACuDtZOlonbAPUAAAAA
'' SIG '' AK4wFgQUxt/YPeGB4gx4GRcbNHPRcKlwySUwDQYJKoZI
'' SIG '' hvcNAQELBQAEggEAeIDM103oJLP3X7VzrP9s2etaoPQX
'' SIG '' UI9tPADrW20UyHuackQ/C4dAgHVcR6L+VXiSZn2xuID7
'' SIG '' WIgstpwUXMWCd+3DOZSlpnN1r4DK/QdTYE7O5dIq/KEY
'' SIG '' b+pFQLs5QlLqW+jmq9CjRNZwittPORk8LfkltZQ6K6AD
'' SIG '' 1iWmsdeg4bpJuwZjmJan/7bucr0agZj6fak3KqECMcrp
'' SIG '' DHWyLnNUdz0tSzMyQeHgs9YAXpI8kyi5nTgDehp8jOeN
'' SIG '' N6lv5hfvdPYgBx8+k0sVl0nU9RoaPm3XP9z6vaMtCLm3
'' SIG '' 3bcLsoueoULdVDRBYccGDgo5l+Dk2HL8SxBgU6hgEZ4b
'' SIG '' sbO7SA==
'' SIG '' End signature block
