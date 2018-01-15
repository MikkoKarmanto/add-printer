'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
' prndrvr.vbs - driver script for WMI on Windows 
'     used to add, delete, and list drivers.
'
' Usage:
' prndrvr [-adlx?] [-m model][-v version][-e environment][-s server]
'         [-u user name][-w password][-h file path][-i inf file]
'
' Example:
' prndrvr -a -m "driver" -v 3 -e "Windows NT x86"
' prndrvr -d -m "driver" -v 3 -e "Windows x64"
' prndrvr -d -m "driver" -v 3 -e "Windows IA64"
' prndrvr -x -s server
' prndrvr -l -s server
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
const kDebugTrace = 1
const kDebugError = 2
dim gDebugFlag

gDebugFlag = false

'
' Operation action values.
'
const kActionUnknown    = 0
const kActionAdd        = 1
const kActionDel        = 2
const kActionDelAll     = 3
const kActionList       = 4

const kErrorSuccess     = 0
const kErrorFailure     = 1

const kNameSpace        = "root\cimv2"

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
const L_Help_Help_General01_Text   = "Usage: prndrvr [-adlx?] [-m model][-v version][-e environment][-s server]"
const L_Help_Help_General02_Text   = "               [-u user name][-w password][-h path][-i inf file]"
const L_Help_Help_General03_Text   = "Arguments:"
const L_Help_Help_General04_Text   = "-a     - add the specified driver"
const L_Help_Help_General05_Text   = "-d     - delete the specified driver"
const L_Help_Help_General06_Text   = "-e     - environment  ""Windows {NT x86 | X64 | IA64}"""
const L_Help_Help_General07_Text   = "-h     - driver file path"
const L_Help_Help_General08_Text   = "-i     - fully qualified inf file name"
const L_Help_Help_General09_Text   = "-l     - list all drivers"
const L_Help_Help_General10_Text   = "-m     - driver model name"
const L_Help_Help_General11_Text   = "-s     - server name"
const L_Help_Help_General12_Text   = "-u     - user name"
const L_Help_Help_General13_Text   = "-v     - version"
const L_Help_Help_General14_Text   = "-w     - password"
const L_Help_Help_General15_Text   = "-x     - delete all drivers that are not in use"
const L_Help_Help_General16_Text   = "-?     - display command usage"
const L_Help_Help_General17_Text   = "Examples:"
const L_Help_Help_General18_Text   = "prndrvr -a -m ""driver"" -v 3 -e ""Windows NT x86"""
const L_Help_Help_General19_Text   = "prndrvr -d -m ""driver"" -v 3 -e ""Windows x64"""
const L_Help_Help_General20_Text   = "prndrvr -a -m ""driver"" -v 3 -e ""Windows IA64"" -i c:\temp\drv\drv.inf -h c:\temp\drv"
const L_Help_Help_General21_Text   = "prndrvr -l -s server"
const L_Help_Help_General22_Text   = "prndrvr -x -s server"
const L_Help_Help_General23_Text   = "Remarks:"
const L_Help_Help_General24_Text   = "The inf file name must be fully qualified. If the inf name is not specified, the script uses"
const L_Help_Help_General25_Text   = "one of the inbox printer inf files in the inf subdirectory of the Windows directory."
const L_Help_Help_General26_Text   = "If the driver path is not specified, the script searches for driver files in the driver.cab file."
const L_Help_Help_General27_Text   = "The -x option deletes all additional printer drivers (drivers installed for use on clients running"
const L_Help_Help_General28_Text   = "alternate versions of Windows), even if the primary driver is in use. If the fax component is installed,"
const L_Help_Help_General29_Text   = "this option deletes any additional fax drivers. The primary fax driver is also deleted if it is not"
const L_Help_Help_General30_Text   = "in use (i.e. if there is no queue using it). If the primary fax driver is deleted, the only way to"
const L_Help_Help_General31_Text   = "re-enable fax is to reinstall the fax component."

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
const L_Text_Msg_General01_Text    = "Added printer driver"
const L_Text_Msg_General02_Text    = "Unable to add printer driver"
const L_Text_Msg_General03_Text    = "Unable to delete printer driver"
const L_Text_Msg_General04_Text    = "Deleted printer driver"
const L_Text_Msg_General05_Text    = "Unable to enumerate printer drivers"
const L_Text_Msg_General06_Text    = "Number of printer drivers enumerated"
const L_Text_Msg_General07_Text    = "Number of printer drivers deleted"
const L_Text_Msg_General08_Text    = "Attempting to delete printer driver"
const L_Text_Msg_General09_Text    = "Unable to list dependent files"
const L_Text_Msg_General10_Text    = "Unable to get SWbemLocator object"
const L_Text_Msg_General11_Text    = "Unable to connect to WMI service"


'
' Printer driver properties
'
const L_Text_Msg_Driver01_Text     = "Server name"
const L_Text_Msg_Driver02_Text     = "Driver name"
const L_Text_Msg_Driver03_Text     = "Version"
const L_Text_Msg_Driver04_Text     = "Environment"
const L_Text_Msg_Driver05_Text     = "Monitor name"
const L_Text_Msg_Driver06_Text     = "Driver path"
const L_Text_Msg_Driver07_Text     = "Data file"
const L_Text_Msg_Driver08_Text     = "Config file"
const L_Text_Msg_Driver09_Text     = "Help file"
const L_Text_Msg_Driver10_Text     = "Dependent files"

'
' Debug messages
'
const L_Text_Dbg_Msg01_Text        = "In function AddDriver"
const L_Text_Dbg_Msg02_Text        = "In function DelDriver"
const L_Text_Dbg_Msg03_Text        = "In function DelAllDrivers"
const L_Text_Dbg_Msg04_Text        = "In function ListDrivers"
const L_Text_Dbg_Msg05_Text        = "In function ParseCommandLine"

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strModel
    dim strPath
    dim uVersion
    dim strEnvironment
    dim strInfFile
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
    iRetval = ParseCommandLine(iAction, strServer, strModel, strPath, uVersion, _
                               strEnvironment, strInfFile, strUser, strPAssword)

    if iRetval = kErrorSuccess  then

        select case iAction

            case kActionAdd
                iRetval = AddDriver(strServer, strModel, strPath, uVersion, _
                                    strEnvironment, strInfFile, strUser, strPassword)

            case kActionDel
                iRetval = DelDriver(strServer, strModel, uVersion, strEnvironment, strUser, strPassword)

            case kActionDelAll
                iRetval = DelAllDrivers(strServer, strUser, strPassword)

            case kActionList
                iRetval = ListDrivers(strServer, strUser, strPassword)

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
' Add a driver
'
function AddDriver(strServer, strModel, strFilePath, uVersion, strEnvironment, strInfFile, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text

    dim oDriver
    dim oService
    dim iResult
    dim uResult

    '
    ' Initialize return value
    '
    iResult = kErrorFailure

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oDriver = oService.Get("Win32_PrinterDriver")

    else

        AddDriver = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        oDriver.Name              = strModel
        oDriver.SupportedPlatform = strEnvironment
        oDriver.Version           = uVersion
        oDriver.FilePath          = strFilePath
        oDriver.InfName           = strInfFile

        uResult = oDriver.AddPrinterDriver(oDriver)

        if Err.Number = kErrorSuccess then

            if uResult = kErrorSuccess then

                wscript.echo L_Text_Msg_General01_Text & L_Space_Text & oDriver.Name

                iResult = kErrorSuccess

            else

                wscript.echo L_Text_Msg_General02_Text & L_Space_Text & strModel & L_Space_Text _
                             & L_Text_Error_General03_Text & L_Space_Text & uResult

            end if

        else

            wscript.echo L_Text_Msg_General02_Text & L_Space_Text & strModel & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General02_Text & L_Space_Text & strModel & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

    end if

    AddDriver = iResult

end function

'
' Delete a driver
'
function DelDriver(strServer, strModel, uVersion, strEnvironment, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text

    dim oDriver
    dim oService
    dim iResult
    dim strObject

    '
    ' Initialize return value
    '
    iResult = kErrorFailure

    '
    ' Build the key that identifies the driver instance.
    '
    strObject = strModel & "," & CStr(uVersion) & "," & strEnvironment

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oDriver = oService.Get("Win32_PrinterDriver.Name='" & strObject & "'")

    else

        DelDriver = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        '
        ' Delete the printer driver instance
        '
        oDriver.Delete_

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General04_Text & L_Space_Text & oDriver.Name

            iResult = kErrorSuccess

        else

            wscript.echo L_Text_Msg_General03_Text & L_Space_Text & strModel & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                         & L_Space_Text & Err.Description

            call LastError()

        end if

    else

        wscript.echo L_Text_Msg_General03_Text & L_Space_Text & strModel & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                     & L_Space_Text & Err.Description

    end if

    DelDriver = iResult

end function

'
' Delete all drivers
'
function DelAllDrivers(strServer, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text

    dim Drivers
    dim oDriver
    dim oService
    dim iResult
    dim iTotal
    dim iTotalDeleted
    dim vntDependentFiles
    dim strDriverName

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Drivers = oService.InstancesOf("Win32_PrinterDriver")

    else

        DelAllDrivers = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General05_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        DelAllDrivers = kErrorFailure

        exit function

    end if

    iTotal = 0
    iTotalDeleted = 0

    for each oDriver in Drivers

        iTotal = iTotal + 1

        wscript.echo
        wscript.echo L_Text_Msg_General08_Text
        wscript.echo L_Text_Msg_Driver01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Driver02_Text & L_Space_Text & oDriver.Name
        wscript.echo L_Text_Msg_Driver03_Text & L_Space_Text & oDriver.Version
        wscript.echo L_Text_Msg_Driver04_Text & L_Space_Text & oDriver.SupportedPlatform

        strDriverName = oDriver.Name

        '
        ' Example of how to delete an instance of a printer driver
        '
        oDriver.Delete_

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General04_Text & L_Space_Text & oDriver.Name

            iTotalDeleted = iTotalDeleted + 1

        else

            '
            ' We cannot use oDriver.Name to display the driver name, because the SWbemLastError
            ' that the function LastError() looks at would be overwritten. For that reason we
            ' use strDriverName for accessing the driver name
            '
            wscript.echo L_Text_Msg_General03_Text & L_Space_Text & strDriverName & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                         & L_Space_Text & Err.Description

            '
            ' Try getting extended error information
            '
            call LastError()

            Err.Clear

        end if

    next

    wscript.echo L_Empty_Text
    wscript.echo L_Text_Msg_General06_Text & L_Space_Text & iTotal
    wscript.echo L_Text_Msg_General07_Text & L_Space_Text & iTotalDeleted

    DelAllDrivers = kErrorSuccess

end function

'
' List drivers
'
function ListDrivers(strServer, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text

    dim Drivers
    dim oDriver
    dim oService
    dim iResult
    dim iTotal
    dim vntDependentFiles

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Drivers = oService.InstancesOf("Win32_PrinterDriver")

    else

        ListDrivers = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General05_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        ListDrivers = kErrorFailure

        exit function

    end if

    iTotal = 0

    for each oDriver in Drivers

        iTotal = iTotal + 1

        wscript.echo
        wscript.echo L_Text_Msg_Driver01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Driver02_Text & L_Space_Text & oDriver.Name
        wscript.echo L_Text_Msg_Driver03_Text & L_Space_Text & oDriver.Version
        wscript.echo L_Text_Msg_Driver04_Text & L_Space_Text & oDriver.SupportedPlatform
        wscript.echo L_Text_Msg_Driver05_Text & L_Space_Text & oDriver.MonitorName
        wscript.echo L_Text_Msg_Driver06_Text & L_Space_Text & oDriver.DriverPath
        wscript.echo L_Text_Msg_Driver07_Text & L_Space_Text & oDriver.DataFile
        wscript.echo L_Text_Msg_Driver08_Text & L_Space_Text & oDriver.ConfigFile
        wscript.echo L_Text_Msg_Driver09_Text & L_Space_Text & oDriver.HelpFile

        vntDependentFiles = oDriver.DependentFiles

        '
        ' If there are no dependent files, the method will set DependentFiles to
        ' an empty variant, so we check if the variant is an array of variants
        '
        if VarType(vntDependentFiles) = (vbArray + vbVariant) then

            PrintDepFiles oDriver.DependentFiles

        end if

        Err.Clear

    next

    wscript.echo L_Empty_Text
    wscript.echo L_Text_Msg_General06_Text & L_Space_Text & iTotal

    ListDrivers = kErrorSuccess

end function

'
' Prints the contents of an array of variants
'
sub PrintDepFiles(Param)

   on error resume next

   dim iIndex

   iIndex = LBound(Param)

   if Err.Number = 0 then

      wscript.echo L_Text_Msg_Driver10_Text

      for iIndex = LBound(Param) to UBound(Param)

          wscript.echo L_Space_Text & Param(iIndex)

      next

   else

        wscript.echo L_Text_Msg_General09_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

   end if

end sub

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
function ParseCommandLine(iAction, strServer, strModel, strPath, uVersion, _
                          strEnvironment, strInfFile, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg05_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-a"
                iAction = kActionAdd

            case "-d"
                iAction = kActionDel

            case "-l"
                iAction = kActionList

            case "-x"
                iAction = kActionDelAll

            case "-s"
                iIndex = iIndex + 1
                strServer = RemoveBackslashes(oArgs(iIndex))

            case "-m"
                iIndex = iIndex + 1
                strModel = oArgs(iIndex)

            case "-h"
                iIndex = iIndex + 1
                strPath = oArgs(iIndex)

            case "-v"
                iIndex = iIndex + 1
                uVersion = oArgs(iIndex)

            case "-e"
                iIndex = iIndex + 1
                strEnvironment = oArgs(iIndex)

            case "-i"
                iIndex = iIndex + 1
                strInfFile = oArgs(iIndex)

            case "-u"
                iIndex = iIndex + 1
                strUser = oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                strPassword = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err.Number <> 0 then

        wscript.echo L_Text_Error_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_text & Err.Description

        ParseCommandLine = kErrorFailure

    else

        ParseCommandLine = kErrorSuccess

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
    wscript.echo L_Empty_Text
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
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General27_Text
    wscript.echo L_Help_Help_General28_Text
    wscript.echo L_Help_Help_General29_Text
    wscript.echo L_Help_Help_General30_Text
    wscript.echo L_Help_Help_General31_Text

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
'' SIG '' mlUP3KOnaFb8M6G5M6AlnOFVzMN6IRQb2eB8kdGCjCig
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
'' SIG '' LwYJKoZIhvcNAQkEMSIEIHzQ3S/HJfQsF18gElR1iwpO
'' SIG '' GlHP8JWRjg0YO5l6qKz2MDwGCisGAQQBgjcKAxwxLgws
'' SIG '' eHplR2pxSHBHS3N5WjQ4S1RycXpYVjlVb2NNeHVjSFVS
'' SIG '' cjFMbjJEWDBmUT0wWgYKKwYBBAGCNwIBDDFMMEqgJIAi
'' SIG '' AE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcA
'' SIG '' c6EigCBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2lu
'' SIG '' ZG93czANBgkqhkiG9w0BAQEFAASCAQCDMMJQ2Txd6+Qz
'' SIG '' ejo0YYipnVoIZH2U/J8Bo+ctChz+nhHxLryE15PdSClQ
'' SIG '' dAH7zX7Sr3W0nUIYedR2heP9ttfpVcJ2XwVOvHS09/UE
'' SIG '' 0EQjI9Nwrari7l+AXlCk47ozrdTg/iSylgcNOD6K+B7R
'' SIG '' sFvJtjxpMZNMex78pfSgzzo/EFREg7OYQgi/WR5P0drP
'' SIG '' 2H+OBDAH22T6Z1dq/Rr4OLhOVzgzb8AojG9foOTQflwz
'' SIG '' gPxZkc9VSBOUlZU1f6m+rMEZ7pisRldXFPMIontDVwoN
'' SIG '' XhH26yu7z/GpQSErkOSJWe9fNnfiw4Qv4+aHmPHAlLh7
'' SIG '' eF6P7UdBu7hU4DaET8hjoYITRjCCE0IGCisGAQQBgjcD
'' SIG '' AwExghMyMIITLgYJKoZIhvcNAQcCoIITHzCCExsCAQMx
'' SIG '' DzANBglghkgBZQMEAgEFADCCATwGCyqGSIb3DQEJEAEE
'' SIG '' oIIBKwSCAScwggEjAgEBBgorBgEEAYRZCgMBMDEwDQYJ
'' SIG '' YIZIAWUDBAIBBQAEIElbwse6yCI0L4/zl1wcxY2bAYIm
'' SIG '' UYc/ZS/qkrVVdQAxAgZZzUFFmyUYEzIwMTcwOTI5MDQy
'' SIG '' NTE1Ljk0NFowBwIBAYACAfSggbikgbUwgbIxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xDDAKBgNVBAsTA0FPQzEnMCUGA1UE
'' SIG '' CxMebkNpcGhlciBEU0UgRVNOOjEyQjQtMkQ1Ri04N0Q0
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
'' SIG '' CLraNtvTX4/edIhJEjCCBNkwggPBoAMCAQICEzMAAACn
'' SIG '' ZF3FKA8BPUQAAAAAAKcwDQYJKoZIhvcNAQELBQAwfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMTYwOTA3
'' SIG '' MTc1NjUyWhcNMTgwOTA3MTc1NjUyWjCBsjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIERTRSBFU046MTJCNC0yRDVGLTg3RDQx
'' SIG '' JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
'' SIG '' cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
'' SIG '' AoIBAQCm6hpQwF9P0lP5oewEfA3XIFcxIVHYx+4DCi54
'' SIG '' zosuIiNrljYQHwpReoXnXRT+c7LX0nxjVcEKuMaUNOi4
'' SIG '' idkXnAmxGKGRP0NWVVbZW8VP1oN+5OgQiBtMehuQotS1
'' SIG '' AtPR+L+bzv81atUujkyTRnqzfU5S0BgR2MyXzEyU5HKL
'' SIG '' UPzq0lIJEDiDEbiWAzE4XEGNOimHEVgTow6tfa3e+uZu
'' SIG '' ytC4oXNtvrWppWdJawd8eYi0ZjbyMSc4FOI1H8EnXXjI
'' SIG '' 4ioya2eBRlMF1ntDNWpEMO27Ie03SX+yRAmb2lnAKz6S
'' SIG '' +A7AJksXiu8I01TnjAuPio1S9qB5qE9mBjK5iyrxAgMB
'' SIG '' AAGjggEbMIIBFzAdBgNVHQ4EFgQUfA1KC337T/9eECw9
'' SIG '' eW8gTUJQiF0wHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xG
'' SIG '' G8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
'' SIG '' L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
'' SIG '' dHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoG
'' SIG '' CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
'' SIG '' L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNU
'' SIG '' aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/
'' SIG '' BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
'' SIG '' 9w0BAQsFAAOCAQEAmsYkS+eXslIBZIcP4kKTKFZ5dvT5
'' SIG '' ChVnTXoOiYA8PzNC3o5y8lym83izubCF0o8l7duKveEs
'' SIG '' fVZgBBLPJ/NjePhGVFRKFtpr6ly7+4Z6bF5/TXNioDad
'' SIG '' sN6z6c8SYoz68JxsJAxriC6Rl77fTjKMXG4nhCd5m53L
'' SIG '' 3+jsEQsACVx6L2ol9tL2OiqHbUd2zFWvTrbx1Xoas/mc
'' SIG '' HQhIKP1x/14HmyVCsKP4C/1h0l+6dOhh1fPi4ES/KQ3j
'' SIG '' ivBtXYxa4uUODYEH0SuO0nlQma0Boss1Abq+AEKDI3G2
'' SIG '' HWCHqoxb/nvtcIYOCHG1V4UvlBAjbbQfvKNGt+UWS67w
'' SIG '' rra6TqGCA3QwggJcAgEBMIHioYG4pIG1MIGyMQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMQwwCgYDVQQLEwNBT0MxJzAlBgNV
'' SIG '' BAsTHm5DaXBoZXIgRFNFIEVTTjoxMkI0LTJENUYtODdE
'' SIG '' NDElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAg
'' SIG '' U2VydmljZaIlCgEBMAkGBSsOAwIaBQADFQDkgi6dMjZt
'' SIG '' dAAyDx7BIbQN6wx32aCBwTCBvqSBuzCBuDELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIE5UUyBFU046MjY2NS00QzNGLUM1REUx
'' SIG '' KzApBgNVBAMTIk1pY3Jvc29mdCBUaW1lIFNvdXJjZSBN
'' SIG '' YXN0ZXIgQ2xvY2swDQYJKoZIhvcNAQEFBQACBQDdeDnv
'' SIG '' MCIYDzIwMTcwOTI5MDMxODA3WhgPMjAxNzA5MzAwMzE4
'' SIG '' MDdaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAN14Oe8C
'' SIG '' AQAwBwIBAAICIxYwBwIBAAICGN4wCgIFAN15i28CAQAw
'' SIG '' NgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAaAK
'' SIG '' MAgCAQACAxbjYKEKMAgCAQACAwehIDANBgkqhkiG9w0B
'' SIG '' AQUFAAOCAQEAVOffEZTLZZxmBU3Z9TcoL0njq5VJ0N/e
'' SIG '' PbxZ0pRl+B+QgyQRPtNbvkmBT4yGBVXLmgZK+Gyz6QkC
'' SIG '' G/SVZRHh8lgPiVkHluhVI8X/7xcRvw/v69tcZbB+tJM7
'' SIG '' DdfDRXPqiEfB/dgF2RsADLtfgs8UBefslZdqrqqHDJjC
'' SIG '' FXwhEruyBdp0XxDeWHLZvA04Io6c636h54gAxaYll2u3
'' SIG '' rZ7Yicriko6VeLOQzJKK+Bml8zEK2b+wUx7vQZ7yIJo0
'' SIG '' 1deVrBsGTteBYuBVVXuoY8a1SCPV18wq6+5nuycGYa3u
'' SIG '' DbeYOedjCfpe0D5VHDFsb24/XeWmqRbhRESqIBfrXONr
'' SIG '' rzGCAvUwggLxAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
'' SIG '' IFBDQSAyMDEwAhMzAAAAp2RdxSgPAT1EAAAAAACnMA0G
'' SIG '' CWCGSAFlAwQCAQUAoIIBMjAaBgkqhkiG9w0BCQMxDQYL
'' SIG '' KoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEINmjA01J
'' SIG '' GVDWesuGixWwCECT6lbjTmrokr9xb05M0BraMIHiBgsq
'' SIG '' hkiG9w0BCRACDDGB0jCBzzCBzDCBsQQU5IIunTI2bXQA
'' SIG '' Mg8ewSG0DesMd9kwgZgwgYCkfjB8MQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
'' SIG '' dGFtcCBQQ0EgMjAxMAITMwAAAKdkXcUoDwE9RAAAAAAA
'' SIG '' pzAWBBTRevOX/EOiJzQFZy8nR+CsoJpzPjANBgkqhkiG
'' SIG '' 9w0BAQsFAASCAQAwtG5YBloy7iG2sRW+XN7TkMhrDcty
'' SIG '' +SQCW6SD9fK5eXjGJpiOQZ/9mgTAWHIln+63wpKlGEb0
'' SIG '' yCldhlZFGwbK87yVl55NIvq+TlrIuGgQRAXs4A5O8GCT
'' SIG '' I8KlHbZjmwa+tbdPMl7RDOslR1A03Xfpn9ZmYOTJRb1g
'' SIG '' ullYllReS8uo/BCd9rtxnsUvGathY6DKz4L7JClL5hip
'' SIG '' xLjIPxQu33dAMTNXq+CJLJePhYy2O6gnImMoZCRs7hdI
'' SIG '' BL2sQXjQDQARmX3l84yGCk0Qku1jeyjl/vw+3iEk4aTf
'' SIG '' vfERJmdwkQ96yFi61jsz6OMqVMvAf/VFpvOMf9UNSbr/
'' SIG '' g4bR
'' SIG '' End signature block
