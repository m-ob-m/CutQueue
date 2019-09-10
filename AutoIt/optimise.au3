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

If $cmdLine[0] > 0 Then $toOptimise = $cmdLine[1]
Local $order = StringRegExpReplace($toOptimise, "\A(.*)\.txt\z", "\1")

; Ouvre l'importateur de CSV de CutRite
Local $command = '"C:\V90\Panel.exe" "' & $toOptimise & '"'
Local $workingDirectory = "C:\V90\FABRIDOR"
Run($command, $workingDirectory, @SW_HIDE)

; Attendre que l'optimisateur ouvre
Local $partListWindow = WinWait("[TITLE:Part list - " & $order & ";]")
;~ ConsoleWrite("Found window with handle " & '"' & $partListWindow & '"' & ".")

; Clique sur le bouton d'optimisation
ControlClick($partListWindow, "", "[CLASS:ToolbarWindow32; INSTANCE:1;]", "primary", 1 , 488, 25)

; Attendre la fenetre d'optimisation
Local $optimiseMainWindow = WinWait("[TITLE:Optimise - " & $order & ";]")

; Attendre la fenêtre qui indique que l'optimisation est en cours. Si une fenêtre d'erreur apparaît, le processus a échoué.
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
      ControlClick($optimiseErrorWindow, "", "[CLASS:Button2;]", "primary", 1, 40, 18)

      ; Fermer la fenetre d'optimisation
      Sleep(500)
      WinKill($optimiseMainWindow)	; Fermeture de l'optimisateur
      WinKill($optimiseErrorWindow)

      ; Message d'erreur
      ConsoleWrite($errorDescription)

      Exit
   EndIf
WEnd

; Attendre la fenetre de revue. Durant ce temps, vérifier si la fenêtre d'erreur de RunCalc ouvre
While (WinExists("[TITLE:Review runs]") == 0)
   If WinExists("[TITLE:RUNCALC]") Then
	  ; Si erreur RUNCALC, fermer tout et envoyer un message d'erreur à PHP

	  WinKill("[TITLE:RUNCALC]")	; Fermeture de l'erreur de RunCalc
	  WinKill($optimiseMainWindow)	; Fermeture de l'optimisateur

	  ; Message d'erreur pour PHP
	  ConsoleWrite('Process "RUNCALC.EXE" crashed unexpectedly.')
	  Exit
   EndIf

   Sleep(500)
WEnd

; *****************************
; * CODE DE TRANSFERT VERS CNC*
; *****************************
TransfertVersCNC()
; *****************************

; Message que tout est correct
ConsoleWrite("OK")

; Fermeture du script
Exit


; ##########################################################################
; # Name:			TransfertVersCNC
; # Description :	Transfert à la CNC et Détecte Data not correct
; #					pour recommencer le transfer
; ##########################################################################
Func TransfertVersCNC()
   Local $reviewRunsWindow = WinGetHandle("[TITLE:Review runs;]")
   While WinExists("[TITLE:Transfer to machining centre; CLASS:#32770;]") == 0
	   WinActivate($reviewRunsWindow)
	   sleep(100)
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
