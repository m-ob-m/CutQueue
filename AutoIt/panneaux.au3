#include <Constants.au3>
;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Mathieu Grenier
;
; Script Function:
;   Ouvre l'optimisateur de porte pour créer le fichier de panneaux.
;

; CSV à importer selon la valeur en argument.
Local $partListFileName = "71999P.txt"	; Doit être changé par argument de ligne de commande.
$DEBUG = True

If $cmdLine[0] > 0 Then $partListFileName = $cmdLine[1]
Local $batch = StringRegExpReplace($partListFileName, "\A(.*)\.txt\z", "\1")
Debug(@ScriptName & ' started.' & @CRLF)

Local $attemptCounter = 1
Local Const $MAX_ATTEMPTS = 5
Local $mainWindow = 0
While 1
	; Ouvre l'importateur de CSV de CutRite.
	Local $command = '"C:\V90\panel.exe" "' & $partListFileName & '"'
	Local $workingDirectory = "C:\V90\FABRIDOR"
	Local $processID = Run($command, $workingDirectory)
	Debug('Executed command "' & $command & '" in working directory "' & $workingDirectory & '".' & @CRLF)

	; Attendre que l'optimisateur ouvre.
	Local $mainWindowTitle = "Part list - " & $batch
	Local $waitingTime = 5
	$mainWindow = WinWait("[TITLE:" & $mainWindowTitle & "]", "", $waitingTime)
	
	If $mainWindow == 0 Then 
		ProcessClose($processID)
		Debug('Command "' & $command & '" failed to produce the expected window with title "' & $mainWindowTitle & '". Process ' & $processID & ' was terminated.' & @CRLF)
		If $attemptCounter <= $MAX_ATTEMPTS Then
			Debug('Retrying...' & @CRLF)
			$attemptCounter += 1
			ContinueLoop
		Else
			Debug(@ScriptName & ' failed.' & @CRLF)
			Exit 1
		EndIf
	EndIf

	Debug('Found window with title "' & $mainWindowTitle & '" and handle ' & $mainWindow & '.' & @CRLF)
	ExitLoop
WEnd

; Cliquer sur le bouton pour aller vers la vue des panneaux (sur la même fenêtre).
Local $toolBarClassName = "ToolbarWindow32"
Local $toolBarInstance = 1
Local $toolBar = ControlGetHandle($mainWindow, "", "[CLASS:" & $toolBarClassName & "; INSTANCE:" & $toolBarInstance & ";]")
Debug('Found control with class "' & $toolBarClassName & '" and instance number ' & $toolBarInstance & ' having handle ' & $toolBar & ' in window with handle ' & $mainWindow & '.' & @CRLF)

Local $button = "primary"
Local $clicks = 1
Local $coord = [394, 21]
ControlClick($mainWindow, "", $toolBar, $button, $clicks, $coord[0], $coord[1])
Debug('Clicked control with handle ' & $toolBar & ' ' & $clicks & ' time' & (($clicks > 1) ? 's' : '') & ' with ' & $button & ' button at coordinates (' & $coord[0] & ', ' & $coord[1] & ').' & @CRLF)

; Cliquer sur le bouton Oui si nécessaire.
While WinExists($mainWindow) And WinGetTitle($mainWindow) == $mainWindowTitle
	Sleep(500)

	Local $messageBoxWindowClassName = "#32770"
	Local $messageBoxWindowTitle = "Part list"
	Local $messageBoxWindow = WinGetHandle("[CLASS:" & $messageBoxWindowClassName & "; TITLE:" & $messageBoxWindowTitle & ";]")
	If @error == 0 Then
		Debug('Found window with class "' & $messageBoxWindowClassName & '" and title "' & $messageBoxWindowTitle & '" and handle ' & $messageBoxWindow & '.' & @CRLF)

		Local $buttonClass = "Button"
		Local $buttonInstance = 1
		Local $yesButton = ControlGetHandle($messageBoxWindow, "", "[CLASS:" & $buttonClass & "; INSTANCE:" & $buttonInstance & ";]")
		Debug('Found control with class "' & $buttonClass & '" and instance number ' & $buttonInstance & ' having handle ' & $yesButton & ' in window with handle ' & $messageBoxWindow & '.' & @CRLF)

		Local $mouseButton = "primary"
		Local $clicks = 1
		ControlClick($messageBoxWindow, "", $yesButton, $mouseButton, $clicks)
		Debug('Clicked control with handle ' & $yesButton & ' ' & $clicks & ' time' & (($clicks > 1) ? 's' : '') & ' with ' & $mouseButton & ' button.' & @CRLF)
		ExitLoop
	EndIf
WEnd
Debug('Toggled to the "Board list" tab.' & @CRLF)

; Fermer la fenêtre.
While WinExists($mainWindow)
	Sleep(500)
	WinKill($mainWindow)
WEnd
Debug('Closed window with handle ' & $mainWindow & '.' & @CRLF)

; Quitter le script
Debug(@ScriptName & ' completed successfuly.' & @CRLF)
Exit

Func Debug($text)
	If $DEBUG == True Then
		Local $handle = FileOpen(@ScriptDir & "\" & StringRegExpReplace(@ScriptName, "\A(.*)\.[^\.]*\z", "\1") & ".log", $FO_APPEND)
		If $handle <> -1 Then
			FileWriteLine($handle, $text)
			FileClose($handle)
		EndIf
	EndIf
EndFunc