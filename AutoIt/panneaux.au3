#include <Constants.au3>
#include <GUIListView.au3>
;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Mathieu Grenier
;
; Script Function:
;   Ouvre l'optimisateur de porte pour créer le fichier de panneaux
;

; CSV à importer selon la valeur en argument
$toOptimise = "MERC_1ER_FEV_MDF_34_DOS_ERABLE.txt"	; Pour fin de tests et debug seulement. Doit être changé par argument

If $cmdLine[0] > 0 Then ; Chargement de l'argument en ligne de commande
	$toOptimise = $cmdLine[1]
EndIf

; Il est nécessaire de faire les appels à partir de ce répertoire
FileChangeDir("c:\V90\FABRIDOR")

; Ouvre l'importateur de CSV de CutRite
Run("c:\V90\panel.exe " & $toOptimise)

; Attendre que l'optimisateur ouvre
WinWait("[TITLE:Part list - ]")

; Renommer la fenêtre car d'autre porte le même nom
WinSetTitle("[TITLE:Part list - ]", "", "PartMain")

; Cliquer sur le bouton des panneaux
ControlClick("[TITLE:PartMain]", "", "ToolbarWindow321", "left",1 ,394,21)

; Cliquer sur le bouton Oui si nécessaire pour faire fermer PartsMain
While WinExists("[TITLE:PartMain]")
	Sleep(500)
	If WinExists("[TITLE:PartMain]") Then
		ControlClick("[TITLE:Part list]", "", "Button1", "left",1 ,40,18)
	EndIf
WEnd

; Fermer la fenêtre
WinWait("[TITLE:Board list - ]")
While WinExists("[TITLE:Board list - ]")
	Sleep(500)
	WinKill("[TITLE:Board list - ]")
WEnd

; Quitter le script
Exit