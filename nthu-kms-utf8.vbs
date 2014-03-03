'===============================================================================
'
'         FILE: nthu-kms-utf8.vbs
'
'        USAGE: ./nthu-kms-utf8.vbs  
'
'  DESCRIPTION: All-in-One KMS Activation Script
'
'      OPTIONS: ---
' REQUIREMENTS: ---
'         BUGS: ---
'        NOTES: ---
'       AUTHOR: SHIE, Li-Yi (lyshie), lyshie@mx.nthu.edu.tw
' ORGANIZATION: 
'      VERSION: 1.0
'      CREATED: 2014-02-26 08:56:10
'     REVISION: ---
'===============================================================================
Option Explicit

' Global Variables
Dim strOSLanguage
Dim strOSCaption
Dim strOSVersion
Dim strOSMajorVersion

' Global Configs
Const strTitle     = "NTHU KMS Activation Script"
Const strKMSServer = "kms.eden.nthu.edu.tw"
Const strKMSPort   = "1688"
Const strValidNet  = "140.114."
Const strURL       = "http://140.114.63.137/cgi-bin/myip.cgi"

' Get my real and external IP address
Function getClientExtIP()
	Dim objHTTP ' object

	Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

	objHTTP.setTimeouts 1000 * 10, 1000 * 10, 1000 * 10, 1000 * 10

	objHTTP.Open "GET", strURL, False

	' lyshie_20140226: **FIXME** to use non-blocking method or timeout
	On Error Resume Next
	objHTTP.Send

	getClientExtIP = objHTTP.responseText
End Function

' Get OS major version
Function getOSMajorVersion(version)
	Dim vs ' array

	vs = Split(version, ".")

	If UBound(vs) >= 0 Then
		getOSMajorVersion = vs(0)
	End IF
End Function

' Get OS information via "WMI Service"
Function getOSInfo()
	Dim strComputer
	Dim objWMIService       ' object
	Dim colOperatingSystems ' collection
	Dim objOperatingSystem  ' object
	Dim colServices         ' collection
	Dim objService          ' object

	strComputer = "."

	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _
	& strComputer _
	& "\root\cimv2")

	' Win32 OS
	On Error Resume Next
	Set colOperatingSystems = objWMIService.ExecQuery _
	("SELECT * FROM Win32_OperatingSystem")

	For Each objOperatingSystem in colOperatingSystems
		strOSCaption  = objOperatingSystem.Caption
		strOSVersion  = objOperatingSystem.Version
		strOSLanguage = objOperatingSystem.OSLanguage
	Next

	strOSMajorVersion = getOSMajorVersion(strOSVersion)
End Function

' Get Office version
Function getWordVersion()
	Dim strVersion
	Dim objWord ' object

	strVersion = "0"

	On Error Resume Next
	Set objWord = CreateObject("Word.Application")
	strVersion = objWord.Version
	objWord.Quit

	getWordVersion = strVersion
End Function

' Get full path name of "ospp.vbs"
Function getOSPP()
	Dim strVersion
	Dim strPath
	Dim strPath_x86
	Dim strOSPP
	Dim objWSHShell ' object
	Dim objFSO      ' object

	strVersion = getWordVersion()

	Set objWSHShell = CreateObject("WScript.Shell")

	' Use registry to query install path
	On Error Resume Next
	strPath = objWSHShell.RegRead("HKLM\Software\Microsoft\Office\" _
	& strVersion _
	& "\Word\InstallRoot\Path")

	' Fallback to default install path
	If strPath = "" Then
		strPath = objWSHShell.ExpandEnvironmentStrings("%ProgramFiles%") _
		& "\Microsoft Office\Office" _
		& CStr(CInt(strVersion)) _
		& "\"

		strPath_x86 = objWSHShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") _
		& "\Microsoft Office\Office" _
		& CStr(CInt(strVersion)) _
		& "\"
	End If

	' Check if 'ospp.vbs' exist
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strOSPP = strPath & "ospp.vbs"

	If Not objFSO.FileExists(strOSPP) Then
		strOSPP = strPath_x86 & "ospp.vbs"

		If Not objFSO.FileExists(strOSPP) Then
			strOSPP = ""
		End If
	End If

	getOSPP = strOSPP
End Function

Function openURL(url)
	Dim objShell ' object

	Set objShell = CreateObject("Shell.Application")

	objShell.Open url
End Function

Function runCommand(cmd)
	Dim objShell ' object
	Dim objExec  ' object
	Dim strResponse

	Set objShell = CreateObject("WScript.Shell")

	Set objExec = objShell.Exec(cmd)

	Do While objExec.Status = 0
		WScript.Sleep 100
	Loop

	strResponse = objExec.StdOut.ReadAll

	If strResponse <> "" Then
		Dim objRE      ' object
		Dim objMatches ' collection
		Dim objMatch   ' object

		Set objRE = New RegExp

		objRE.Global = True
		objRE.IgnoreCase = True
		objRE.Pattern = ".*(success|error).*"

		Set objMatches = objRE.Execute(strResponse)

		For Each objMatch In objMatches
			monoEcho objMatch.Value
		Next

		monoEcho strResponse
	End If
End Function

Function activateOS()
	Dim strCmd

	strCmd = "%ComSpec% /c " _
	& "slmgr.vbs -skms " & strKMSServer & " & " _
	& "slmgr.vbs -ato"

	runCommand(strCmd)
End Function

Function activateOffice(strOSPP)
	Dim strCmd
	Dim boolOSPPService

	strCmd = "%ComSpec% /c "

	' Check if service is running (Office Software Protection Platform)
	On Error Resume Next
	Set colServices = objWMIService.ExecQuery _
	("SELECT * FROM Win32_Service WHERE DisplayName = 'Office Software Protection Platform'")

	boolOSPPService = False

	For Each objService in colServices
		If objService.State = "Running" Then
			boolOSPPService = True
		End If
	Next

	If (Not boolOSPPService) And (strOSMajorVersion <= 5) Then
		strCmd = strCmd _
		& "cscript " & """" & strOSPP & """" & " /osppsvcrestart & "
	End If

	strCmd = strCmd _
	& "cscript " & """" & strOSPP & """" & " /sethst:" & strKMSServer & " & " _
	& "cscript " & """" & strOSPP & """" & " /setprt:" & strKMSPort & " & " _
	& "cscript " & """" & strOSPP & """" & " /act"

	runCommand(strCmd)
End Function

Function yesNo(msg_zh, msg_en)
	Dim intResult

	intResult = MsgBox(msg_zh & vbCrLf & vbCrLf & msg_en, 36, strTitle)

	If intResult = 6 Then
		yesNo = True
	Else
		yesNo = False
	End If
End Function

Function monoEcho(msg)
	MsgBox msg, 64, strTitle
End Function

Function dualEcho(msg_zh, msg_en)
	MsgBox msg_zh & vbCrLf & vbCrLf & msg_en, 64, strTitle
End Function

Function Main()
	Dim clientIP
	Dim ospp

	' Display welcome message
	dualEcho "如果過程中出現錯誤訊息，請記下該代碼，寄至 service@cc.nthu.edu.tw。", _
	"If error code shows up during the processes, " _
	& "please write it down and e-mail to service@cc.nthu.edu.tw"

	clientIP = getClientExtIP()

	If (InStr(clientIP, strValidNet) > 0) Then
		' Windows
		If strOSMajorVersion > 5 Then
			If yesNo("您的作業系統是 " & strOSCaption & "。是否繼續啟用？", _
				"Your OS is " & strOSCaption & ". Continue to activate?") Then
				activateOS
			End If
		Else
			dualEcho "您的作業系統是 " & strOSCaption & "。不需要啟用！", _
			"Your OS is " & strOSCaption & ". Do not need to activate!"
		End If

		' Office
		ospp = getOSPP()

		If ospp <> "" Then
			If yesNO("是否繼續啟用 Microsoft Office？", _
				"Continue to activate Microsoft Office?") Then
				activateOffice(ospp)
			End If
		Else
			dualEcho "您並未安裝較新版本的 Microsoft Office。不需要啟用！", _
			"You did not install a newer version of Microsoft Office. Do not need to activate!"
		End If
	Else
		dualEcho "您的 IP 位址不允許啟用。(" & clientIP & ")", _
		"Your IP address is not allowed to activate. (" & clientIP & ")"

		dualEcho "您應該使用 SSL-VPN 來登入「國立清華大學」校園。", _
		"You should use SSL-VPN to login to NTHU campus."

		openURL "http://net.nthu.edu.tw/2009/sslvpn:info"
	End If
End Function

'------------------------------------------------------------------------------
getOSInfo

If (strOSMajorVersion > 5) And (Not WScript.Arguments.Named.Exists("elevate")) Then
	CreateObject("Shell.Application").ShellExecute WScript.FullName _
	, """" & WScript.ScriptFullName & """" & " /elevate", "", "runas", 1
	WScript.Quit
Else
	dualEcho "以下的動作將以系統管理者的身份執行。", _
	"The following actions will run as administrator!"
	Main
End If 
'------------------------------------------------------------------------------
