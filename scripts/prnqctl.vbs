'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
' prnqctl.vbs - printer control script for WMI on Windows 
'    used to pause, resume and purge a printer
'    also used to print a test page on a printer
'
' Usage:
' prnqctl [-zmex?] [-s server] [-p printer] [-u user name] [-w password]
'
' Examples:
' prnqctl -m -s server -p printer
' prnqctl -x -s server -p printer
' prnqctl -e -b printer
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
const kActionUnknown    = 0
const kActionPause      = 1
const kActionResume     = 2
const kActionPurge      = 3
const kActionTestPage   = 4

const kErrorSuccess     = 0
const KErrorFailure     = 1

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
const L_Help_Help_General01_Text   = "Usage: prnqctl [-zmex?] [-s server][-p printer][-u user name][-w password]"
const L_Help_Help_General02_Text   = "Arguments:"
const L_Help_Help_General03_Text   = "-e     - print test page"
const L_Help_Help_General04_Text   = "-m     - resume the printer"
const L_Help_Help_General05_Text   = "-p     - printer name"
const L_Help_Help_General06_Text   = "-s     - server name"
const L_Help_Help_General07_Text   = "-u     - user name"
const L_Help_Help_General08_Text   = "-w     - password"
const L_Help_Help_General09_Text   = "-x     - purge the printer (cancel all jobs)"
const L_Help_Help_General10_Text   = "-z     - pause the printer"
const L_Help_Help_General11_Text   = "-?     - display command usage"
const L_Help_Help_General12_Text   = "Examples:"
const L_Help_Help_General13_Text   = "prnqctl -e -s server -p printer"
const L_Help_Help_General14_Text   = "prnqctl -m -p printer"
const L_Help_Help_General15_Text   = "prnqctl -x -p printer"

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
const L_Text_Error_General03_Text  = "Unable to get printer instance."
const L_Text_Error_General04_Text  = "Win32 error code"
const L_Text_Error_General05_Text  = "Unable to get SWbemLocator object"
const L_Text_Error_General06_Text  = "Unable to connect to WMI service"


'
' Action strings
'
const L_Text_Action_General01_Text = "Pause"
const L_Text_Action_General02_Text = "Resume"
const L_Text_Action_General03_Text = "Purge"
const L_Text_Action_General04_Text = "Print Test Page"

'
' Debug messages
'
const L_Text_Dbg_Msg01_Text        = "In function ExecPrinter"
const L_Text_Dbg_Msg02_Text        = "Server name"
const L_Text_Dbg_Msg03_Text        = "Printer name"
const L_Text_Dbg_Msg04_Text        = "In function ParseCommandLine"

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strPrinter
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
    iRetval = ParseCommandLine(iAction, strServer, strPrinter, strUser, strPassword)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionPause
                 iRetval = ExecPrinter(strServer, strPrinter, strUser, strPassword, L_Text_Action_General01_Text)

            case kActionResume
                 iRetval = ExecPrinter(strServer, strPrinter, strUser, strPassword, L_Text_Action_General02_Text)

            case kActionPurge
                 iRetval = ExecPrinter(strServer, strPrinter, strUser, strPassword, L_Text_Action_General03_Text)

            case kActionTestPage
                 iRetval = ExecPrinter(strServer, strPrinter, strUser, strPassword, L_Text_Action_General04_Text)

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
' Pause/Resume/Purge printer and print test page
'
function ExecPrinter(strServer, strPrinter, strUser, strPassword, strCommand)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text
    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text & L_Space_Text & strServer
    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text & L_Space_Text & strPrinter

    dim oPrinter
    dim oService
    dim iRetval
    dim uResult

    iRetval = kErrorFailure

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer.DeviceID='" & strPrinter & "'")

    else

        ExecPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if getting a printer instance succeeded
    '
    if Err.Number = kErrorSuccess then

        select case strCommand

            case L_Text_Action_General01_Text
                 uResult = oPrinter.Pause()

            case L_Text_Action_General02_Text
                 uResult = oPrinter.Resume()

            case L_Text_Action_General03_Text
                 uResult = oPrinter.CancelAllJobs()

            case L_Text_Action_General04_Text
                 uResult = oPrinter.PrintTestPage()

            case else
                 Usage(true)

        end select

        '
        ' Err set by WMI
        '
        if Err.Number = kErrorSuccess then

            '
            ' uResult set by printer methods
            '
            if uResult = kErrorSuccess then

                wscript.echo L_Success_Text & L_Space_Text & strCommand & L_Space_Text _
                             & L_Printer_Text & L_Space_Text & strPrinter

                iRetval = kErrorSuccess

            else

                wscript.echo L_Failed_Text & L_Space_Text & strCommand & L_Space_Text _
                             & L_Text_Error_General04_Text & L_Space_Text & uResult

            end if

        else

            wscript.echo L_Failed_Text & L_Space_Text & strCommand & L_Space_Text & L_Error_Text _
                         & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        end if

    else

        wscript.echo L_Text_Error_General03_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    ExecPrinter = iRetval

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
function ParseCommandLine(iAction, strServer, strPrinter, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-z"
                iAction = kActionPause

            case "-m"
                iAction = kActionResume

            case "-x"
                iAction = kActionPurge

            case "-e"
                iAction = kActionTestPage

            case "-p"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

            case "-s"
                iIndex = iIndex + 1
                strServer = RemoveBackslashes(oArgs(iIndex))

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

    if Err.Number = kErrorSuccess then

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
    wscript.echo L_Empty_Text
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
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General12_Text
    wscript.echo L_Help_Help_General13_Text
    wscript.echo L_Help_Help_General14_Text
    wscript.echo L_Help_Help_General15_Text

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

          Err.Clear

      else

          wscript.echo L_Text_Error_General06_Text & L_Space_Text & L_Error_Text _
                       & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                       & Err.Description

      end if

   else

       wscript.echo L_Text_Error_General05_Text & L_Space_Text & L_Error_Text _
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
'' SIG '' MIIeLAYJKoZIhvcNAQcCoIIeHTCCHhkCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' bC2hk3cvC4yTgr8z1VD7j98e4AUJNisLu9VlGTQcB1Gg
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
'' SIG '' n5efy1eAbzOpBM93pGIcWX4xghKiMIISngIBATCBnDCB
'' SIG '' hDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEuMCwGA1UEAxMlTWlj
'' SIG '' cm9zb2Z0IFdpbmRvd3MgUHJvZHVjdGlvbiBQQ0EgMjAx
'' SIG '' MQITMwAAAXMwMQcmZbi5swAAAAABczANBglghkgBZQME
'' SIG '' AgEFAKCCAQQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
'' SIG '' AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
'' SIG '' LwYJKoZIhvcNAQkEMSIEIKTor5MQacZAOIOaZqdAM2TQ
'' SIG '' y2vnErwbew9MUPFEam7PMDwGCisGAQQBgjcKAxwxLgws
'' SIG '' UTd3QVhnMVRkL2NjRUUyVzVOTnJnK0ZGNk5ZR2RHVWxj
'' SIG '' Ti9OTWx2T1hlMD0wWgYKKwYBBAGCNwIBDDFMMEqgJIAi
'' SIG '' AE0AaQBjAHIAbwBzAG8AZgB0ACAAVwBpAG4AZABvAHcA
'' SIG '' c6EigCBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vd2lu
'' SIG '' ZG93czANBgkqhkiG9w0BAQEFAASCAQC0uXEyKHGNoXKO
'' SIG '' Ozyvmo6M2h+VmtNt51NubL8/aQFLzH2ffG+QdhcTqNus
'' SIG '' 3ZQmwdbX/pLrsT5rlzbcNLYR75XARBWj/WX8VAMUvwg7
'' SIG '' njkNMHs6/J+QgJZTteXD2IPX4SLaUEO90THD1KMZrP0B
'' SIG '' dsK04MsFodti/qRRZ4WwUft6qwYPi75yYU1KMOZyGVjl
'' SIG '' wF3O8URQWnMnFXSblHkZ9i6A6cqc67Fi4FfGmhO470jo
'' SIG '' 73SWaqoDbUYazS4Peq2uvSV9doHkLLpkv7f0DhRpmPJq
'' SIG '' 180v0b0H/Zrb4xyxOVAbUmcvIN13DM6Oot9/VRHovvLm
'' SIG '' MTPQKhm6MdPTQ7rbhZ2roYIPzjCCD8oGCisGAQQBgjcD
'' SIG '' AwExgg+6MIIPtgYJKoZIhvcNAQcCoIIPpzCCD6MCAQMx
'' SIG '' DzANBglghkgBZQMEAgEFADCCATwGCyqGSIb3DQEJEAEE
'' SIG '' oIIBKwSCAScwggEjAgEBBgorBgEEAYRZCgMBMDEwDQYJ
'' SIG '' YIZIAWUDBAIBBQAEIM3xPra8KfFn11u7Q3D3vPfgyXbN
'' SIG '' H8nsh2G3BSyY6FkIAgZZzGKV3McYEzIwMTcwOTI5MDQy
'' SIG '' NzM1LjQwMlowBwIBAYACAfSggbikgbUwgbIxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xDDAKBgNVBAsTA0FPQzEnMCUGA1UE
'' SIG '' CxMebkNpcGhlciBEU0UgRVNOOjIxMzctMzdBMC00QUFB
'' SIG '' MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
'' SIG '' ZXJ2aWNloIILUjCCBnEwggRZoAMCAQICCmEJgSoAAAAA
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
'' SIG '' CLraNtvTX4/edIhJEjCCBNkwggPBoAMCAQICEzMAAACv
'' SIG '' NY//0yJFdksAAAAAAK8wDQYJKoZIhvcNAQELBQAwfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMTYwOTA3
'' SIG '' MTc1NjU2WhcNMTgwOTA3MTc1NjU2WjCBsjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQL
'' SIG '' Ex5uQ2lwaGVyIERTRSBFU046MjEzNy0zN0EwLTRBQUEx
'' SIG '' JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
'' SIG '' cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
'' SIG '' AoIBAQCT6ECGDhrLcfADDaaPpOdhEtBh6pBqlVa6t6BN
'' SIG '' Sylk2FGaO11nk62I7kPU3V3U7xJxn20/9eqXKsg6t1UR
'' SIG '' 1KcztXRETZhN4Xqlx92pLMyIHwhvPvHPhOXhbA/nw+45
'' SIG '' gYspMydZ+E/OpZuNBgOu3tRJU2LjWSf16ApMvMbgH1fd
'' SIG '' 0Zv/TgfQbGCYl4GuYQ6QeeQLKMv5vRBrPMFzERBe4uNS
'' SIG '' 9j9heuRFr3VBrUsMtek0GhE3cKA/6KwrtAeTvq6Dydni
'' SIG '' dJFwYbAVekT0hJEN/NskM2YAzCCL7sFoPXEfn30L5dJo
'' SIG '' J4aeabKqsxxKF1T5B2n+d2/HP8ac0vFSUKvHlzijAgMB
'' SIG '' AAGjggEbMIIBFzAdBgNVHQ4EFgQUfR5VxjGK2mALBmaw
'' SIG '' FNlq/HsS9uEwHwYDVR0jBBgwFoAU1WM6XIoxkPNDe3xG
'' SIG '' G8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
'' SIG '' L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
'' SIG '' dHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3JsMFoG
'' SIG '' CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
'' SIG '' L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNU
'' SIG '' aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcnQwDAYDVR0TAQH/
'' SIG '' BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
'' SIG '' 9w0BAQsFAAOCAQEAasgAhnj7DEKcENb3KKwPhGqL/N1b
'' SIG '' 6h9aQESpLhNfixpJdmkSKuaE5vjmutGPPC4p7Us1IOCB
'' SIG '' uvrVgxcVm3VTmjx/Y5DLoAhkL/Nh80rwivBFToiNPGuQ
'' SIG '' jQDiayNJrHlnIOdlM6765CLBQg11ZQ/Unc0e4zDerFqG
'' SIG '' GsKFh0ck7nw84nI4ORhuPHWN5gHZx3dwo9r6WbmJQFmR
'' SIG '' tsXyu4UtndUYf+OTG3J/sQmtEAOuxSUm582jfqK1uPVK
'' SIG '' 1TQUfI/ygsM2CNIwiDZ11wgZa2IXkdMDQrJzjNEVqHUp
'' SIG '' rP+ARQtFsd1y9lk1uVIGzuaIf/8trKzR/OCjz/wcOR/b
'' SIG '' 0BDT2zGCAvUwggLxAgEBMIGTMHwxCzAJBgNVBAYTAlVT
'' SIG '' MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
'' SIG '' ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
'' SIG '' YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
'' SIG '' YW1wIFBDQSAyMDEwAhMzAAAArzWP/9MiRXZLAAAAAACv
'' SIG '' MA0GCWCGSAFlAwQCAQUAoIIBMjAaBgkqhkiG9w0BCQMx
'' SIG '' DQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIA09
'' SIG '' m48CkY2/86giU6pPvGkGKnzTK9g4AzEisBUmEH8sMIHi
'' SIG '' BgsqhkiG9w0BCRACDDGB0jCBzzCBzDCBsQQU2OqsFS1v
'' SIG '' LqUtpv0CA1DigBv+TiMwgZgwgYCkfjB8MQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
'' SIG '' ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAK81j//TIkV2SwAA
'' SIG '' AAAArzAWBBR+bBKIHjGclb/uLLuGOYTxKYUcUzANBgkq
'' SIG '' hkiG9w0BAQsFAASCAQBWLdvHxPPsVNng2qY/NVhMgSO+
'' SIG '' 5h8NEARSssCp9OQxPbOKVjzhFU4VK/n6je+ZOLG9Kysn
'' SIG '' LyOPnLxbI68FpHAHJICZz11cyBl8eUeAd1Pnq1LDCK/6
'' SIG '' Qv73StNao6pd+Xjhncj4TWdypb2r/u/O4YErWTaX00Fn
'' SIG '' tCRO+ek3xdUmF7lVFTXpcJAChBGqgWjlFKTKSXltBIpF
'' SIG '' 3aUYq2PBvB+/66sz53ZOLLOWHBzfXH6U8qSATOF28I8U
'' SIG '' A3Z03IeBmo5lH+60Wu+JtRxUKGWbQxuQTsXd6wPr8YtY
'' SIG '' dg5a/l69pdD1vmOJtummAh1LyngsP50j8mpjXm+HhkTj
'' SIG '' t2NYBtxj
'' SIG '' End signature block
