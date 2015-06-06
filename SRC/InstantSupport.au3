#AutoIt3Wrapper_Run_Obfuscator=Y
#Obfuscator_Parameters=/StripOnly

; Disable the scripts ability to pause.
Break(0)

#include "Aut2Exe\Include\GUIConstantsEx.au3"
#include "Aut2Exe\Include\StaticConstants.au3"
#include "Aut2Exe\Include\WindowsConstants.au3"
#include "Aut2Exe\Include\ServiceControl.au3"


; Exit if the script hasn't been compiled
If Not @Compiled Then
	MsgBox(0, "ERROR", 'Script must be compiled before running!', 5)
	Exit
EndIf


; Language strings.
$str_Program_Title = 					"Instant Support"
$str_Button_InstallService = 			"Install As Service"
$str_Button_Exit =						"Exit"
$str_MsgBox_Information = 				"Information"
$str_MsgBox_ExitInstantSupport = 		"Exit Instant Support?"
$str_MsgBox_ServiceInstallation = 		"Service Installation"
$str_MsgBox_Error =						"Error"
$str_MsgBox_RemoveService =				"Remove Service and Uninstall?"
$str_ServiceEnterAnIDNumber = 			"Enter an ID number:"
$str_ServiceInvalidID = 				"Invalid ID entered, service installation canceled."
$str_ServiceProxy =						"Service installation not supported when using HTTP Proxy."
$str_ErrorInstallService =				"Installing service requires administrator privileges."
$str_ErrorStopService =					"Stopping VNC services requires administrator privileges."
$str_ErrorUnknownCommand =				"Unknown command."
$str_ErrorRepeaterConnectionFailed = 	"Connection to the repeater was not possible."
$str_EndSupportSession =				"Are you sure you want to end this support session?"
$str_CloseOtherVNCServers =				"Another VNC server is running which must be stopped. Try to stop other VNC server?"


; Global Vars.
Global $ExtractFiles = True
Global $GenerateID = True
Global $LanMode = False
Global $IDNumber = 123456
Global $WorkingPath = @AppDataDir & "\InstantSupport_Temp_Files"
Global $ProxyEnabled = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable")


; Create unique working path if our default directory already exists. (Possible InstantSupport is already running)
If FileExists( @AppDataDir & "\InstantSupport_Temp_Files" ) Then
	$WorkingPath = @AppDataDir & "\InstantSupport_Temp_Files_" & Random( 100000, 999999,1 )
EndIf


; Command line args.
If $cmdline[0] > 0 Then
	Switch $cmdline[1]

		Case "-installservice"
			If IsAdmin() Then

				InstallService()
				Exit

			Else

				MsgBox( 0, $str_MsgBox_Error, $str_ErrorInstallService, 30 )
				Exit

			EndIf

		Case "-removeservice"
			If IsAdmin() Then

				RemoveService()
				Exit

			Else
				; Elevate to admin to remove.
				ShellExecuteWait( @ScriptFullPath, "-removeservice", @ScriptDir, "runas")
				Exit

			EndIf

		Case "-stopservices"
			If IsAdmin() Then

				StopServices()
				Exit

			Else

				MsgBox( 0, $str_MsgBox_Error, $str_ErrorStopService, 30 )
				Exit

			EndIf

		Case Else
			MsgBox( 0 , $str_MsgBox_Error, $str_ErrorUnknownCommand, 30 )
			Exit

	EndSwitch
EndIf


; Extract files.
If $ExtractFiles Then

	DirCreate( $WorkingPath )
	FileInstall( "InstantSupport_Files\instantsupport.ini", $WorkingPath & "\instantsupport.ini", 1 )
	FileInstall( "InstantSupport_Files\logo.jpg", $WorkingPath & "\logo.jpg", 1 )
	FileInstall( "InstantSupport_Files\SecureVNCPlugin.dsm", $WorkingPath & "\SecureVNCPlugin.dsm", 1 )
	FileInstall( "InstantSupport_Files\ultravnc.ini", $WorkingPath & "\ultravnc.ini", 1 )
	FileInstall( "InstantSupport_Files\winvnc.exe", $WorkingPath & "\InstantSupportVNC.exe", 1 )
	FileInstall( "InstantSupport_Files\unblock.js", $WorkingPath & "\unblock.js", 1 )

	; Unblock InstantSupport.exe to prevent "Windows Security" messages.
	ShellExecuteWait($WorkingPath & "\unblock.js", "", @ScriptDir, "")

	FileCopy( @ScriptDir & "\" & @ScriptName, $WorkingPath & "\InstantSupport.exe", 9 )

EndIf


; Close known VNC servers. --NEEDS WORK--
If ProcessExists( "InstantSupportVNC.exe" ) Or ProcessExists( "WinVNC.exe" ) Then

	If MsgBox( 4, $str_Program_Title, $str_CloseOtherVNCServers ) = 6 Then

		; Stop services if running.
		If _ServiceRunning("", "uvnc_service") or _ServiceRunning("", "winvnc") Then
			If IsAdmin() Then

				ShellExecuteWait($WorkingPath & "\InstantSupport.exe", "-stopservices", @ScriptDir, "")

			Else

				ShellExecuteWait($WorkingPath & "\InstantSupport.exe", "-stopservices", @ScriptDir, "runas")

			EndIf
			Sleep(10000)
		EndIf

		; Kill user mode VNC servers.
		Run( $WorkingPath & "\InstantSupportVNC.exe -kill" )
		Sleep(10000)

		; Kill InstantSupport if running.
		If WinExists( $str_Program_Title ) Then

				WinClose( $str_Program_Title )
				Sleep( 500 )
				Send( "{ENTER}" ) ; Let the other InstantSupport process close normally so it can cleanup.
				Sleep( 500 )
		EndIf

	Else

		InstantSupportExit( True )

	EndIf

EndIf


; Read server settings from instantsupport.ini.
$RepeaterAddress = IniRead( $WorkingPath & "\instantsupport.ini", "Repeater", "Address", "" )
$RepeaterAddressLAN = IniRead( $WorkingPath & "\instantsupport.ini", "Repeater", "AddressLAN", "" )
$RepeaterServerPort = IniRead( $WorkingPath & "\instantsupport.ini", "Repeater", "ServerPort", "" )


; Generate a random ID Number between 200,000 and 999,999 or read the installed ID number.
If $GenerateID Then

	$LowerLimit = 200000
	$UpperLimit = 999999
	$IDNumber = Random( $LowerLimit,$UpperLimit,1 )

Else

	$IDNumber = IniRead( $WorkingPath & "\instantsupport.ini", "InstantSupport", "ID", "" )

EndIf


; Create the GUI.
$InstantSupport = GUICreate( $str_Program_Title, 450, 200, -1, -1, BitOR( $WS_SYSMENU,$WS_CAPTION,$WS_POPUP,$WS_POPUPWINDOW,$WS_BORDER,$WS_CLIPSIBLINGS,$WS_MINIMIZEBOX ) )
GUISetBkColor( 0xFFFFFF )
$Label2 = GUICtrlCreateLabel( $IDNumber, 0, 100, 450, 100, $SS_CENTER )
GUICtrlSetFont( -1, 50, 800, 0, "Arial Black" )
$Pic1 = GUICtrlCreatePic( $WorkingPath & "\logo.jpg", 0, 0, 450, 90, BitOR( $SS_NOTIFY,$WS_GROUP,$WS_CLIPSIBLINGS ) )
GUISetState( @SW_SHOW )
WinSetOnTop( $str_Program_Title, "",1)


; Check to see if the repeater exists unless there is a proxy.
If $ProxyEnabled = False Then

	TCPStartUp()

	; Test WAN address first.
	$socket = TCPConnect( TCPNameToIP( $RepeaterAddress ), $RepeaterServerPort )
	If $socket = -1 Then $LanMode = True

	; Test LAN Address because WAN failed.
	If $LanMode = True Then

		$socket = TCPConnect( TCPNameToIP( $RepeaterAddressLAN ), $RepeaterServerPort )

		If $socket = -1 Then

			; No connections possible, exit.
			WinSetOnTop( $str_Program_Title, "",0 )
			MsgBox( 48, $str_MsgBox_Error, $str_ErrorRepeaterConnectionFailed, 10 )
			InstantSupportExit( True )

		EndIf

	EndIf

	TCPShutdown()

EndIf


; Start the VNC server and make a reverse connection to the repeater.
If $LanMode = True Then
	ShellExecute( $WorkingPath & "\InstantSupportVNC.exe", "-httpproxy -autoreconnect ID:" & $IDNumber & " -connect " & $RepeaterAddressLAN & ":" & $RepeaterServerPort & " -run" )
Else
	ShellExecute( $WorkingPath & "\InstantSupportVNC.exe", "-httpproxy -autoreconnect ID:" & $IDNumber & " -connect " & $RepeaterAddress & ":" & $RepeaterServerPort & " -run" )
EndIf


; Create the tray icon. Default tray menu items (Script Paused/Exit) will not be shown.
Opt( "TrayMenuMode", 1 )
$InstallItem = TrayCreateItem( $str_Button_InstallService )
$ExitItem = TrayCreateItem( $str_Button_Exit )

; Enable the scripts ability to pause. (otherwise tray menu is disabled)
Break(1)

; Main loop.
While 1

	; Close any windows firewall messages that popup. The windows firewall doesn't block outgoing connections anyways. ----Does the OS language change this?----
	If WinExists( "Windows Security Alert" ) Then WinClose( "Windows Security Alert" )

	; If UltraVNC can't connect to the repeater then exit.
	If WinExists( "Initiate Connection" ) Then

		; Stop the VNC server
		ProcessClose( "InstantSupportVNC.exe" )
		ProcessWaitClose( "InstantSupportVNC.exe" )
		WinSetOnTop( $str_Program_Title, "",0 )
		MsgBox( 48, $str_MsgBox_Error, $str_ErrorRepeaterConnectionFailed )
		InstantSupportExit( True )

	EndIf

	; Check for form events.
	$nMsg = GUIGetMsg()
	Switch $nMsg

		Case $GUI_EVENT_CLOSE

			WinSetOnTop( $str_Program_Title, "",0 )
			If MsgBox( 4, $str_Program_Title, $str_EndSupportSession ) = 6 Then

				InstantSupportExit( True )

			EndIf

	EndSwitch

	; Tray events.
	$nMsg = TrayGetMsg()
	Switch $nMsg

		Case $InstallItem

			; Service installation unavailable if using an HTTP Proxy (UltraVNC limitation).
			If $ProxyEnabled = True Then

				WinSetOnTop( $str_Program_Title, "",0 )
				MsgBox( 48, $str_MsgBox_Error, $str_ServiceProxy, 10 )

			Else
				; Choose ID for service installation.
				WinSetOnTop( $str_Program_Title, "",0 )
				$IDNumber = InputBox( $str_MsgBox_ServiceInstallation, $str_ServiceEnterAnIDNumber, $IDNumber ) +0

				If IsNumber($IDNumber) and $IDNumber > 200000 and $IDNumber < 999999 Then

					; Configure ultravnc.ini
					If $LanMode = True Then

						IniWrite( $WorkingPath & "\ultravnc.ini", "admin", "service_commandline", '-autoreconnect ID:' & $IDNumber & ' -connect ' & $RepeaterAddressLAN & ":" & $RepeaterServerPort )

					Else

						IniWrite( $WorkingPath & "\ultravnc.ini", "admin", "service_commandline", '-autoreconnect ID:' & $IDNumber & ' -connect ' & $RepeaterAddress & ":" & $RepeaterServerPort )

					EndIf

					; Configure instantsupport.ini
					IniWrite( $WorkingPath & "\instantsupport.ini", "InstantSupport", "ID", $IDNumber )

					; Kill the VNC server
					$PID = Run( $WorkingPath & "\InstantSupportVNC.exe -kill" )

					; Wait for the winvnc server to close.
					ProcessWaitClose( $PID, 15 )

					; Run installer after the server exits.
					ProcessWaitClose( "InstantSupportVNC.exe" )

					If IsAdmin() Then

						ShellExecute($WorkingPath & "\InstantSupport.exe", "-installservice", @ScriptDir, "")

					Else

						ShellExecute($WorkingPath & "\InstantSupport.exe", "-installservice", @ScriptDir, "runas")

					EndIf


					InstantSupportExit( True )

				Else

					WinSetOnTop( $str_Program_Title, "",0 )
					MsgBox( 0, $str_MsgBox_Information, $str_ServiceInvalidID )

				EndIf

			EndIf


		Case $ExitItem

			WinSetOnTop( $str_Program_Title, "",0 )
			If MsgBox( 4, $str_Program_Title, $str_EndSupportSession ) = 6 Then

				InstantSupportExit( True )

			EndIf

	EndSwitch

WEnd


Func InstallService()

	; Copy files.
	FileCopy( @ScriptDir & "\*.*", @ProgramFilesDir & "\InstantSupport\", 9 )

	; Create uninstall link on All Users desktop.
	FileCreateShortcut( @ProgramFilesDir & '\InstantSupport\InstantSupport.exe', @DesktopCommonDir & '\Uninstall Instant Support.lnk', "", "-removeservice" )

	; Install VNC Service.
	ShellExecute( @ProgramFilesDir & "\InstantSupport\InstantSupportVNC.exe", "-install" )

EndFunc


Func RemoveService()

	WinSetOnTop( $str_Program_Title, "",0 )
	If MsgBox( 4, $str_Program_Title, $str_MsgBox_RemoveService ) = 6 Then

		; Remove Uninstaller
		FileDelete( @DesktopCommonDir & '\Uninstall Instant Support.lnk"' )

		; Remove VNC Service.
		ShellExecute( @ProgramFilesDir & "\InstantSupport\InstantSupportVNC.exe", "-uninstall" )

		; Remove the InstantSupport Service.
		_DeleteSelf( @ProgramFilesDir & "\InstantSupport", 15 )

	EndIf

EndFunc


Func _DeleteSelf( $Path, $iDelay = 5 )

	Local $sCmdFile

	FileDelete( @TempDir & "\scratch.bat" )


	$sCmdFile = 'PING -n ' & $iDelay & ' 127.0.0.1 > nul' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\InstantSupportVNC.exe"' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\InstantSupport.exe"' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\SecureVNCPlugin.dsm"' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\ultravnc.ini"' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\instantsupport.ini"' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\logo.jpg"' & @CRLF _
			& 'DEL /F /Q "' & $Path & '\unblock.js"' & @CRLF _
 			& 'RMDIR "' & $Path & '"' & @CRLF _
			& 'DEL "' & @TempDir & '\scratch.bat"'

	FileWrite( @TempDir & "\scratch.bat", $sCmdFile )

	Run( @TempDir & "\scratch.bat", @TempDir, @SW_HIDE )

EndFunc


Func InstantSupportExit( $DeleteFiles = False )

	; Kill the VNC server
	$PID = Run( $WorkingPath & "\InstantSupportVNC.exe -kill" )

	; Wait for the winvnc server to close.
	ProcessWaitClose( $PID, 15 )

	; Remove temp files.
	If $DeleteFiles = True Then _DeleteSelf( $WorkingPath, 5)

	Exit

EndFunc


Func StopServices()

	_StopService("", "uvnc_service")
	_StopService("", "winvnc")

EndFunc



