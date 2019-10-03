#include <AutoItConstants.au3>
#include <Constants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <GUIListView.au3>

;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Mathieu Grenier
;
; Script Function:
;   Ouvre l'optimisateur de porte
;

;ConsoleWrite("Starting optimise.exe.")

; CSV à importer selon la valeur en argument
$toOptimise = "71999P.txt"	; Pour fin de tests et debug seulement. Doit être changé par argument
$DEBUG = True

If $cmdLine[0] > 0 Then $toOptimise = $cmdLine[1]
Local $order = StringRegExpReplace($toOptimise, "\A(.*)\.txt\z", "\1")

; Ouvre l'importateur de CSV de CutRite
Local $command = '"C:\V90\Panel.exe" "' & $toOptimise & '"'
Local $workingDirectory = "C:\V90\FABRIDOR"
Run($command, $workingDirectory, @SW_HIDE)
Debug("Executed command " & '"' & $command & '"' & " in working directory " & '"' & $workingDirectory & '"' & ".")

; Attendre que l'optimisateur ouvre
Local $partListWindow = WinWait("[TITLE:Part list - " & $order & ";]")
Debug("Found window with title " & '"' & "Part list" & '"' & " and handle " & '"' & $partListWindow & '"' & "." & @CRLF)

; Clique sur le bouton d'optimisation
ControlClick($partListWindow, "", "[CLASS:ToolbarWindow32; INSTANCE:1;]", "primary", 1, 488, 25)
Debug("Clicked control with class " & '"' & "ToolbarWindow32" & '"' & " and instance number 1 on window with handle " & '"' & $partListWindow & '"' & "." & @CRLF)

; Attendre la fenetre d'optimisation
Local $optimiseMainWindow = WinWait("[TITLE:Optimise - " & $order & ";]")
Debug("Found window with title " & '"' & "Optimise - " & $order & '"' & " and handle " & '"' & $partListWindow & '"' & "." & @CRLF)

; Attendre la fenêtre qui indique que l'optimisation est en cours. Si une fenêtre d'erreur apparaît, le processus a échoué.
Debug("Pre-optimisation process starting." & @CRLF)
While WinExists($optimiseMainWindow) == 1 And WinExists("[TITLE:Optimise; CLASS:#32770;]") == 0
	If WinExists("[TITLE:Error; CLASS:#32770;]") == 1 Then
		Local $optimiseErrorWindow = WinGetHandle("[TITLE:Error; CLASS:#32770;]")
		Local $optimiseErrorWindowGXWNDControl = ControlGetHandle($optimiseErrorWindow, "", "[CLASS:GXWND; INSTANCE:1;]")
		Opt ("WinDetectHiddenText", 1)
		
		; *** ERREUR ***
		ControlClick($optimiseErrorWindow, "", $optimiseErrorWindowGXWNDControl, "primary", 1 , 100, 30)	;Clique pour obtenir le message d'erreur
		Sleep(500)
		Local $errorDescription = ControlGetText($optimiseErrorWindow, "", "[CLASS:Edit; INSTANCE:1;]") & @CRLF
		ControlClick($optimiseErrorWindow, "", $optimiseErrorWindowGXWNDControl, "primary", 1 , 400, 30)	;Clique pour obtenir le message d'erreur
		Sleep(500)
		$errorDescription = $errorDescription & ControlGetText($optimiseErrorWindow, "", "[CLASS:Edit; INSTANCE:1;]")
		
		; Annuler l'optimisation
		Sleep(500)
		ControlClick($optimiseErrorWindow, "", "[CLASS:Button2; INSTANCE:1;]", "primary", 1, 40, 18)
		
		; Fermer la fenetre d'optimisation
		Sleep(500)
		WinKill($optimiseMainWindow)
		
		Local $windowsProcessFailedWindow = Null
		Do
			Debug("Trying to find window with title " & '"' & "BATCH" & '"' & " and class " & '"' & "#32770" & "." & @CRLF)
			$windowsProcessFailedWindow = WinWait("[TITLE:BATCH; CLASS:#32770;]", "", 2)
			If ($windowsProcessFailedWindow <> 0) Then
				Debug("Detected window with title " & '"' & "BATCH" & '"' & " and class " & '"' & "#32770" & '"' & " having handle " & '"' & $windowsProcessFailedWindow & '"' & "." & @CRLF)
				WinKill($windowsProcessFailedWindow)
				Debug("Killed window with title " & '"' & "BATCH" & '"' & " and class " & '"' & "#32770" & '"' & " having handle " & '"' & $windowsProcessFailedWindow & '"' & "." & @CRLF)
			EndIf
		Until $windowsProcessFailedWindow == 0
		
		; Message d'erreur
		ConsoleWrite($errorDescription)
		
		Debug("An error occured during the pre-optimisation process. Exiting..." & @CRLF)
		Exit 1
	EndIf
WEnd
Debug("Pre-optimisation process done." & @CRLF)

; Attendre la fenetre de revue. Durant ce temps, vérifier si la fenêtre d'erreur de RunCalc ouvre
Debug("Optimisation process starting." & @CRLF)
While (WinExists("[TITLE:Review runs]") == 0)
	If WinExists("[TITLE:RUNCALC]") Then
		; Si erreur RUNCALC, fermer tout et envoyer un message d'erreur à PHP
		
		WinKill("[TITLE:RUNCALC]")	; Fermeture de l'erreur de RunCalc
		WinKill($optimiseMainWindow)	; Fermeture de l'optimisateur
		
		; Message d'erreur pour PHP
		Debug('Process "RUNCALC.EXE" crashed unexpectedly. Exiting...' & @CRLF)
		ConsoleWrite('Process "RUNCALC.EXE" crashed unexpectedly.')
		Exit 2
	EndIf
	
	Sleep(500)
WEnd
Debug("Optimisation process done." & @CRLF)

; *****************************
; * CODE DE TRANSFERT VERS CNC*
; *****************************
Debug("Transfer process starting" & @CRLF)
TransfertVersCNC()
Debug("Transfer process done." & @CRLF)
; *****************************

; Message que tout est correct
ConsoleWrite("OK")

; Fermeture du script
Exit 0


; ##########################################################################
; # Name:			TransfertVersCNC
; # Description :	Transfert à la CNC et Détecte Data not correct
; #					pour recommencer le transfer
; ##########################################################################
Func TransfertVersCNC()
	Local $reviewRunsWindow = WinGetHandle("[TITLE:Review runs;]")
	While WinExists("[TITLE:Transfer to machining centre; CLASS:#32770;]") == 0
		WinActivate($reviewRunsWindow)
		Sleep(100)
		ControlClick($reviewRunsWindow, "", "[CLASS:Afx:ToolBar:400000:8:10003:10; INSTANCE:1;]", "primary", 1, 345, 30)
		Sleep(400)
	WEnd

	While WinExists("[TITLE:Transfer to machining centre; CLASS:#32770;]")
		Local $transferWindow = WinGetHandle("[TITLE:Transfer to machining centre; CLASS:#32770;]")
		If ControlGetText($transferWindow, "", "[CLASS:Static; INSTANCE:1]") == "Data not correct" Then	; Détecte la fenetre d'erreur
			While WinExists($transferWindow)	; Fait disparaitre la fenetre d'erreur
				ControlClick($transferWindow, "", "[CLASS:Button2]", "primary", 1, 39, 10)
				Sleep(500)
			Wend
			Sleep(3000)
			TransfertVersCNC()	; Recommence le transfert vers la CNC
		EndIf
	
		Sleep(100)
	WEnd
	
	; Fermeture de la fenetre de revue
	Sleep(500)
	WinKill($reviewRunsWindow)
EndFunc

Func Debug($text)
	If $debug == True Then
		Local $handle = FileOpen(@ScriptDir & "\optimise.log", $FO_APPEND)
		If $handle <> -1 Then
			FileWriteLine($handle, $text)
			FileClose($handle)
		EndIf
	EndIf
EndFunc