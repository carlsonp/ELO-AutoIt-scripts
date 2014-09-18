;Commandline parameters to pass in
;1st - building (make sure it's lowercase)
;e.g. "howe"
;2nd - classroom number
;e.g. "1344"
;3rd - instructor computer, fully qualified domain name
;This can be found by going into the "properties" of the computer
;e.g. "DBL057V1.engineering.iastate.edu"
;4th - VNC launcher file (.vnc file ending)
;e.g. "C:\Users\elocoder\Desktop\1344InstructorPC.vnc"

;https://www.autoitscript.com/autoit3/docs/keywords/include.htm
#include "passwords.au3" ;ELO login passwords

;These includes are for GUI/UI constants
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>

;Constants
Global Const $BUILDING = $CmdLine[1]
Global Const $ROOM = $CmdLine[2]
Global Const $INSTRUCTOR_ADDRESS = $CmdLine[3] ;fully qualified domain name
Global Const $VNC_LAUNCHER = $CmdLine[4]
Global Const $PANOPTO = "C:\Program Files (x86)\Panopto\Focus Recorder\Recorder.exe"
Global Const $CHROME = "C:\Users\elocoder\AppData\Local\Google\Chrome\Application\chrome.exe"
Global Const $FIREFOX = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
Global Const $MYWEBEX_LINK = "https://iowastate.webex.com/mw0401l/mywebex/default.do?siteurl=iowastate&service=10"
Global Const $WEBEX_TRAINING_CENTER_LINK = "https://iowastate.webex.com/mw0401l/mywebex/default.do?siteurl=iowastate&service=7"
Global Const $SESSION_MANAGER_LINK = "https://sws.elo.iastate.edu/session-manager/"
Global Const $ECHO_LINK = "https://" & $BUILDING & $ROOM & "hd.engineering.iastate.edu:8443/advanced"
;Settings
Opt("WinTitleMatchMode", 2) ;matches partial substrings when matching titles
;https://www.autoitscript.com/trac/autoit/ticket/1573
Opt("TCPTimeout", 1000) ;TCP timeout, currently not working in stable version of AutoIt?

;https://www.autoitscript.com/autoit3/docs/guiref/GUIRef.htm
;https://www.autoitscript.com/autoit3/docs/functions/GUICreate.htm
;https://www.autoitscript.com/autoit3/docs/functions/GUICtrlCreateInput.htm
Local $gui = GUICreate("Session Manager Login", 300, 300) ;by default it's centered
GUICtrlCreateLabel("Enter your session manager login", 10, 10)
GUICtrlCreateLabel("ISU NetID:", 10, 30)
Local $username_input = GUICtrlCreateInput("", 80, 30, 120, 20)
GUICtrlCreateLabel("Password:", 10, 60)
Local $password_input = GUICtrlCreateInput("", 80, 60, 120, 20, $ES_PASSWORD)
Local $idLogin = GUICtrlCreateButton("Login", 70, 100, 60)
GUICtrlSetState($idLogin, $GUI_DEFBUTTON)
GUISetState(@SW_SHOW, $gui)

;Loop until the user enters session manager login info or just exits
While True
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			GUIDelete($gui)
			ExitLoop
		Case $idLogin
			$username = GUICtrlRead($username_input)
			$password = GUICtrlRead($password_input)
			GUIDelete($gui)
			ExitLoop
	EndSwitch
WEnd

; Disable user input from the mouse and keyboard.
BlockInput(1)

;https://www.autoitscript.com/autoit3/docs/functions/ProgressOn.htm
ProgressOn("Startup", "Startup", "", (@desktopwidth/2)-100, @desktopheight-175)

ProgressSet(15, "Starting Panopto...")

If WinExists("Panopto Focus") Then
	BlockInput(0)
	$Panopto_MSG = MsgBox($MB_YESNO+$MB_ICONQUESTION, "Panopto Open", "Panopto is already open, would you like to restart Panopto?")
	If $Panopto_MSG == $IDYES Then ;clicked YES
		;close Panopto
		;https://www.autoitscript.com/autoit3/docs/functions/WinClose.htm
		WinClose("Panopto Focus")
		WinWaitClose("Panopto Focus")
	EndIf
	BlockInput(1)
EndIf

If Not WinExists("Panopto Focus") Then
	; Run Panopto application
	; TODO: check return value
	Run($PANOPTO)
	Sleep(2000) ;sleep 2 seconds
	WinActivate("Panopto Focus")
EndIf

If WinExists("Panopto Focus") Then
	WinActivate("Panopto Focus")
	WinMove("Panopto Focus", "", 0, 0, @desktopwidth/2, @desktopheight/2)
	;check to make sure Panopto isn't currently recording
	Local $color = Hex(PixelGetColor(197, 165), 6) ;get green of pause button
	If $color <> "329932" Then
		;Check to see if quality is set to "High"
		Local $color = Hex(PixelGetColor(82, 461), 6) ;get bluish radio item selection
		If $color <> "48BAD7" Then ;<> is not equal to
			MouseClick("left", 82, 461) ;click on "High" quality setting
			MouseClick("left", 232, 460) ;click apply
		EndIf
		;Get color of pixel position to check before mouse click
		Local $color = Hex(PixelGetColor(624, 178), 6)
		If $color == "FFFFFF" Then
			;https://www.autoitscript.com/autoit3/docs/functions/MouseClick.htm
			MouseClick("left", 624, 178, 2)
			Send("^a") ;CTRL + a, select all text
			;Set Panopto file name
			Send("ClassName Number Lecture XX: " & @MON & "-" & @MDAY & "-" & @YEAR)
		Else
			BlockInput(0)
			MsgBox($MB_ICONERROR, "ERROR", "Unable to identify Panopto box location, check the resolution perhaps?")
			BlockInput(1)
		EndIf
	EndIf
EndIf

ProgressSet(30, "Starting Chrome...")

If WinExists("Chrome") Then
	BlockInput(0)
	$Chrome_MSG = MsgBox($MB_YESNO+$MB_ICONQUESTION, "Chrome Open", "Chrome is already open, would you like to restart Chrome?")
	If $Chrome_MSG == $IDYES Then ;clicked YES
		;Chrome can have many processes open, so we have to close them all
		While ProcessExists("chrome.exe")
			Local $PID = ProcessExists("chrome.exe")
			ProcessClose($PID)
			ProcessWaitClose($PID)
		WEnd
	ElseIf $Chrome_MSG == $IDNO Then ;clicked NO
		;open up a new tab for the upcoming content
		WinActivate("Chrome")
		Send('^t') ;CTRL + t, new tab
		Send($MYWEBEX_LINK)
		Send("{ENTER}")
		Sleep(2000)
	EndIf
	BlockInput(1)
EndIf

If Not WinExists("Chrome") Then
	; Run Chrome application
	; TODO: check return value
	;https://www.autoitscript.com/autoit3/docs/functions/ShellExecute.htm
	ShellExecute($CHROME, $MYWEBEX_LINK)
	Sleep(2000) ;sleep 2 seconds
	ProcessWait("chrome.exe")
EndIf

If WinExists("Chrome") Then
	;resize window
	WinMove("Chrome", "", @desktopwidth*0.3, 0, @desktopwidth-(@desktopwidth*0.3), @desktopheight/2)
	WinActivate("Chrome")
EndIf

ProgressSet(40, "Opening links...")

If WinExists("Chrome") Then
	;WebEx is annoying because the login page is not a distinct URL, it uses Javascript
	;Therefore, we have to try to go to a page that requires a login, login, and then go
	;to the training center page.
	WinActivate("Chrome")
	Send($ROOM & $BUILDING)
	Send("{TAB}")
	Send($WEBEX_PASSWORD)
	;pause for a moment, not sure why this is needed, login doesn't work otherwise
	Sleep(200)
	Send("{ENTER}")
	Sleep(2000) ;wait for page to load
	Send("{F6}") ;highlight URL address bar
	Send($WEBEX_TRAINING_CENTER_LINK)
	Send("{ENTER}")

	Send('^t') ;CTRL + t, new tab
	Send($SESSION_MANAGER_LINK)
	Send("{ENTER}")
	Sleep(2000)

	If IsDeclared("username") AND $username <> "" Then
		;Get the URL to see if we're logged in or not
		Send("{F6}")
		Send("^c") ;CTRL + c - copy URL text
		;https://www.autoitscript.com/autoit3/docs/functions/ClipGet.htm
		Local $url = ClipGet()
		If $url == "https://weblogin.iastate.edu/cgi-bin/index.cgi" Then
			;reload the page so we get the cursor in the login box
			Send("{F6}")
			Send($SESSION_MANAGER_LINK)
			Send("{ENTER}")
			Sleep(2000)
			Send($username)
			Send("{TAB}")
			Send($password)
			Send("{ENTER}")
		EndIf
	EndIf
EndIf

ProgressSet(55, "Starting Firefox...")

If WinExists("Firefox") Then
	BlockInput(0)
	$Firefox_MSG = MsgBox($MB_YESNO+$MB_ICONQUESTION, "Firefox Open", "Firefox is already open, would you like to restart Firefox?")
	If $Firefox_MSG == $IDYES Then ;clicked YES
		While ProcessExists("firefox.exe")
			Local $PID = ProcessExists("firefox.exe")
			ProcessClose($PID)
			ProcessWaitClose($PID)
		WEnd
	EndIf
	BlockInput(1)
EndIf

If Not WinExists("Firefox") Then
	; Run Firefox application
	; TODO: check return value
	ShellExecute($FIREFOX, $ECHO_LINK)
	Sleep(2000) ;sleep 2 seconds
	ProcessWait("firefox.exe")
EndIf

If WinExists("Firefox") Then
	WinActivate("Firefox")
	;subtraction is to adjust slightly so it's not below the windows taskbar
	WinMove("Firefox", "", 0, @desktopheight/2, @desktopwidth/2, (@desktopheight/2)-45)
EndIf

ProgressSet(80, "Logging into Echo...")

If Not WinExists("Authentication") Then
	;Firefox was already open, so we open a new tab for Echo
	WinActivate("Firefox")
	Send('^t') ;CTRL + t, new tab
	Send($ECHO_LINK)
	Send("{ENTER}")
	Sleep(1000)
EndIf

If WinExists("Authentication") Then
	WinActivate("Authentication")
	Send($ECHO_BOX_USERNAME)
	Send("{TAB}")
	Send($ECHO_BOX_PASSWORD)
	Send("{ENTER}")
EndIf

ProgressSet(90, "Launching VNC...")

If Not WinExists("TightVNC Viewer") Then
	;http://www.autoitscript.com/forum/topic/131506-port-check-function/
	;https://www.autoitscript.com/autoit3/docs/functions/TCPNameToIP.htm
	TCPStartup()
	Local $TCP_socket = TCPConnect(TCPNameToIP($INSTRUCTOR_ADDRESS), 5900)
	TCPCloseSocket(5900)
	If $TCP_socket > 0 Then
		;VNC test connection passed so let's try to open it
		ShellExecute($VNC_LAUNCHER)
		Sleep(1000) ;sleep 1 second
	Else
		BlockInput(0)
		MsgBox($MB_ICONERROR, "ERROR", "Unable to connect to VNC, check to make sure the instructor computer is turned on and logged in.")
		BlockInput(1)
	EndIf
EndIf

If WinExists("TightVNC Viewer") Then
	WinMove("TightVNC Viewer", "", @desktopwidth/2, @desktopheight/2, @desktopwidth/2, @desktopheight/2)
	WinActivate("TightVNC Viewer")
EndIf

If WinExists("Chrome") Then
	WinActivate("Chrome") ;set focus to Chrome
EndIf

; Enable user input from the mouse and keyboard.
BlockInput(0)
