;Commandline parameters to pass in
;1st - class room
;e.g. "howe1344"

;https://www.autoitscript.com/autoit3/docs/keywords/include.htm
#include "passwords.au3"

;Constants
Global $PANOPTO = "C:\Program Files (x86)\Panopto\Focus Recorder\Recorder.exe"
Global $CHROME = "C:\Users\elocoder\AppData\Local\Google\Chrome\Application\chrome.exe"
Global $FIREFOX = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
Global $WEBEX_LINK = "https://iowastate.webex.com/mw0401l/mywebex/default.do?siteurl=iowastate&service=7"
Global $SESSION_MANAGER_LINK = "https://sws.elo.iastate.edu/session-manager/"
Global $ECHO_LINK = "https://" & $CmdLine[1] & "hd.engineering.iastate.edu:8443/advanced"
;Settings
Opt("WinTitleMatchMode", 2) ;matches partial substrings when matching titles

;https://www.autoitscript.com/autoit3/docs/functions/ProgressOn.htm
ProgressOn("Startup", "Startup", "", @desktopwidth/2, @desktopheight/2)

; Disable user input from the mouse and keyboard.
BlockInput(1)

ProgressSet(10, "Starting Panopto...")

; Run Panopto application
; TODO: check return value
Run($PANOPTO)
Sleep(2000) ;sleep 2 seconds
WinMove("Panopto Focus", "", 0, 0, @desktopwidth/2, @desktopheight/2)

ProgressSet(20, "Starting Chrome...")

; Run Chrome application
; TODO: check return value
;https://www.autoitscript.com/autoit3/docs/functions/ShellExecute.htm
ShellExecute($CHROME, $WEBEX_LINK)
Sleep(2000) ;sleep 2 seconds
WinMove("Chrome", "", @desktopwidth/2, 0, @desktopwidth/2, @desktopheight/2)
ProcessWait("chrome.exe")

ProgressSet(30, "Opening links...")

; TODO: login (figure out javascript)
Send('^t') ;CTRL + t, new tab
Send($SESSION_MANAGER_LINK)
Send("{ENTER}")


ProgressSet(40, "Starting Firefox...")

; Run Firefox application
; TODO: check return value
ShellExecute($FIREFOX, $ECHO_LINK )
Sleep(2000) ;sleep 2 seconds
;subtraction is to adjust slightly so it's not below the windows taskbar
WinMove("Firefox", "", 0, @desktopheight/2, @desktopwidth/2, (@desktopheight/2)-45)
ProcessWait("firefox.exe")

ProgressSet(50, "Setting filename...")

;Get color of pixel position to check before mouse click
Local $color = Hex(PixelGetColor(624, 178), 6)
If $color == "FFFFFF" Then
	;https://www.autoitscript.com/autoit3/docs/functions/MouseClick.htm
	MouseClick("left", 624, 178, 2)
	Send("^a") ;CTRL + a, select all text
Else
	MsgBox(0, "ERROR", "Unable to identify Panopto box location, check the resolution perhaps?")
EndIf

;Set Panopto file name
Send("ClassName Number Lecture XX: " & @MON & "-" & @MDAY & "-" & @YEAR)

ProgressSet(60, "Logging into Echo...")

WinActivate("Authentication")
Send($ECHO_BOX_USERNAME)
Send("{TAB}")
Send($ECHO_BOX_PASSWORD)
Send("{ENTER}")


; Enable user input from the mouse and keyboard.
BlockInput(0)
