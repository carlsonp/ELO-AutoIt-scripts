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
Global Const $WEBEX_LINK = "https://iowastate.webex.com/mw0401l/mywebex/default.do?siteurl=iowastate&service=7"
Global Const $SESSION_MANAGER_LINK = "https://sws.elo.iastate.edu/session-manager/"
Global Const $ECHO_LINK = "https://" & $BUILDING & $ROOM & "hd.engineering.iastate.edu:8443/advanced"
;Settings
Opt("WinTitleMatchMode", 2) ;matches partial substrings when matching titles

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

;https://www.autoitscript.com/autoit3/docs/functions/ProgressOn.htm
ProgressOn("Startup", "Startup", "", (@desktopwidth/2)-100, @desktopheight-175)

; Disable user input from the mouse and keyboard.
BlockInput(1)

ProgressSet(15, "Starting Panopto...")

If WinExists("Panopto Focus") Then
	BlockInput(0)
	$Panopto_MSG = MsgBox($MB_YESNO, "Panopto Open", "Panopto is already open, would you like to restart Panopto?")
	If $Panopto_MSG = $IDYES Then ;clicked YES
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
	WinMove("Panopto Focus", "", 0, 0, @desktopwidth/2, @desktopheight/2)
	;Check to see if quality is set to "High"
	Local $color = Hex(PixelGetColor(82, 461), 6) ;get bluish radio item selection
	If $color <> "48BAD7" Then ;<> is not equal to
		MouseClick("left", 82, 461) ;click on "High" quality setting
		MouseClick("left", 232, 460) ;click apply
	EndIf
EndIf

ProgressSet(30, "Starting Chrome...")

If WinExists("Chrome") Then
	BlockInput(0)
	$Chrome_MSG = MsgBox($MB_YESNO, "Chrome Open", "Chrome is already open, would you like to restart Chrome?")
	If $Chrome_MSG = $IDYES Then ;clicked YES
		;Chrome can have many processes open, so we have to close them all
		While ProcessExists("chrome.exe")
			Local $PID = ProcessExists("chrome.exe")
			ProcessClose($PID)
			ProcessWaitClose($PID)
		WEnd
	EndIf
	BlockInput(1)
EndIf

; Run Chrome application
; TODO: check return value
;https://www.autoitscript.com/autoit3/docs/functions/ShellExecute.htm
ShellExecute($CHROME, $WEBEX_LINK)
Sleep(2000) ;sleep 2 seconds
WinMove("Chrome", "", @desktopwidth/2, 0, @desktopwidth/2, @desktopheight/2)
ProcessWait("chrome.exe")

ProgressSet(40, "Opening links...")

;Look for distinctive blue text of "Host Log In" button
;Get color of pixel position to check before mouse click
Local $color = Hex(PixelGetColor(1546, 190), 6)
If $color == "0164CF" Then
	;https://www.autoitscript.com/autoit3/docs/functions/MouseClick.htm
	MouseClick("left", 1546, 190)
	;wait for page to load
	Sleep(2000) ;sleep 2 seconds
	;enter the login credentials
	Send($ROOM & $BUILDING)
	Send("{TAB}")
	Send($WEBEX_PASSWORD)
	Send("{ENTER}")
Else
	BlockInput(0)
	MsgBox(0, "ERROR", "Unable to identify WebEx Host Log In button, check the resolution perhaps?")
	BlockInput(1)
EndIf

Send('^t') ;CTRL + t, new tab
Send($SESSION_MANAGER_LINK)
Send("{ENTER}")

WinActivate("Chrome")
If IsDeclared("username") Then
	Send($username)
	Send("{TAB}")
	Send($password)
	Send("{ENTER}")
EndIf


ProgressSet(55, "Starting Firefox...")

If WinExists("Firefox") Then
	BlockInput(0)
	$Firefox_MSG = MsgBox($MB_YESNO, "Firefox Open", "Firefox is already open, would you like to restart Firefox?")
	If $Firefox_MSG = $IDYES Then ;clicked YES
		While ProcessExists("firefox.exe")
			Local $PID = ProcessExists("firefox.exe")
			ProcessClose($PID)
			ProcessWaitClose($PID)
		WEnd
	EndIf
	BlockInput(1)
EndIf

; Run Firefox application
; TODO: check return value
ShellExecute($FIREFOX, $ECHO_LINK )
Sleep(2000) ;sleep 2 seconds
;subtraction is to adjust slightly so it's not below the windows taskbar
WinMove("Firefox", "", 0, @desktopheight/2, @desktopwidth/2, (@desktopheight/2)-45)
ProcessWait("firefox.exe")

ProgressSet(70, "Setting filename...")

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
	MsgBox(0, "ERROR", "Unable to identify Panopto box location, check the resolution perhaps?")
	BlockInput(1)
EndIf

ProgressSet(80, "Logging into Echo...")

WinActivate("Authentication")
Send($ECHO_BOX_USERNAME)
Send("{TAB}")
Send($ECHO_BOX_PASSWORD)
Send("{ENTER}")

ProgressSet(90, "Launching VNC...")

;http://www.autoitscript.com/forum/topic/131506-port-check-function/
;https://www.autoitscript.com/autoit3/docs/functions/TCPNameToIP.htm
TCPStartup()
Local $TCP_socket = TCPConnect(TCPNameToIP($INSTRUCTOR_ADDRESS), 5900)
TCPCloseSocket(5900)
If $TCP_socket > 0 Then
	;VNC test connection passed so let's try to open it
	ShellExecute($VNC_LAUNCHER)
	Sleep(1000) ;sleep 1 second
	WinMove("TightVNC Viewer", "", @desktopwidth/2, @desktopheight/2, @desktopwidth/2, @desktopheight/2)
Else
	BlockInput(0)
	MsgBox(0, "ERROR", "Unable to connect to VNC, check to make sure the instructor computer is turned on and logged in.")
	BlockInput(1)
EndIf

; Enable user input from the mouse and keyboard.
BlockInput(0)
