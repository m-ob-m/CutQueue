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
$toOptimise = "MERC_1ER_FEV_MDF_34_DOS_ERABLE.txt"	; Pour fin de tests et debug seulement. Doit �tre chang� par argument

if $cmdLine[0] > 0 Then ; Chargement de l'argument en ligne de commande
   $toOptimise = $cmdLine[1]
EndIf

; Il est nécessaire de faire les appels à partir de ce répertoire
FileChangeDir("C:\V90\FABRIDOR")

; Ouvre l'importateur de CSV de CutRite
Run("C:\V90\Panel.exe " & $toOptimise)
;ConsoleWrite("Lauching C:\V90\Panel.exe.")

; Attendre que l'optimisateur ouvre
WinWait("[TITLE:Part list - ]")

; Clique sur le bouton d'optimisation
ControlClick("[TITLE:Part list - ]", "", "ToolbarWindow321", "left",1 ,488,25)

; Attendre la fenetre d'optimisation
WinWait("[TITLE:Optimise - ]")

; Attendre que la fenetre avec le message d'erreur _pnumber.mpr s'ouvre pour cliquer OK
Local $window = WinWait("[TITLE:Error]", "", 1)
If($window <> 0 ) Then
   Opt ("WinDetectHiddenText", 1)
   $texts = StringSplit(WinGetText("[TITLE:Error]"), @LF)

   ; *** ERREUR ***
   ControlClick("[TITLE:Error]", "", "GXWND1", "left",1 ,386,11)	;Clique pour obtenir le message d'erreur
   Sleep(500)

   ; Obtention du message d'erreur
   $erreur = ControlGetText("[TITLE:Error]", "", "Edit1")

   ; Annuler l'optimisation
   Sleep(500)
   ControlClick("[TITLE:Error]", "", "Button2", "left",1 ,40,18)

   ; Fermer la fenetre d'optimisation
   Sleep(500)
   WinKill("[TITLE:Optimise - ]")

   ; Message d'erreur
   ConsoleWrite("<b>ERREUR<b><br><hr><br><span style='color:#CC0000;'>" & $erreur & "<BR>" & $texts[6] & "</span>")

   Exit
EndIf

; Attendre la fenetre de revue. Durant ce temps, vérifier si la fenêtre d'erreur de RunCalc ouvre
While WinExists("[TITLE:Review runs]") = false
   If WinExists("[TITLE:RUNCALC]") Then
	  ; Si erreur RUNCALC, fermer tout et envoyer un message d'erreur à PHP

	  WinKill("[TITLE:RUNCALC]")	; Fermeture de l'erreur de RunCalc
	  WinKill("[TITLE:Optimise - ]")	; Fermeture de l'optimisateur

	  ; Message d'erreur pour PHP
	  ConsoleWrite("<b>ERREUR<b><br><hr><br><span style='color:#CC0000;'>RUNCALC.EXE CRASH</span>")
	  Exit
   EndIf

   Sleep(500)
WEnd

; *****************************
; * CODE DE TRANSFERT VERS CNC*
; *****************************
TransfertVersCNC()
; *****************************

; Fermeture de la fenetre de revue
Sleep(500)
WinKill("[TITLE:Review runs]")

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
   Local $window = WinGetHandle("[TITLE:Review runs]", "")
   Local $control = ControlGetHandle($window, "", "[CLASS:Afx:ToolBar:400000:8:10003:10; INSTANCE:1]")
   While WinExists("[TITLE:Transfer to machining centre]") = false
	   WinActivate($window)
	   sleep(100)
	   ControlClick($window, "", $control, "primary", 1, 345, 30)
	   Sleep(400)
   WEnd

   While WinExists("[TITLE:Transfer to]")
	  if WinExists("[TITLE:Transfer to]", "Data not correct") Then	; Détecte la fenetre d'erreur
		 While WinExists("[TITLE:Transfer to]", "Data not correct")	; Fait disparaitre la fenetre d'erreur
			ControlClick("[TITLE:Transfer to]", "Data not correct", "Button2", "left",1 ,39,10)
			Sleep(500)
		 Wend
		 Sleep(3000)
		 TransfertVersCNC()	; Recommence le transfert vers la CNC
	  EndIf

	  Sleep(100)
   WEnd

EndFunc
