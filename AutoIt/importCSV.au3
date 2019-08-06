#include <Constants.au3>
#include <GUIListView.au3>
;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Mathieu Grenier
;
; Script Function:
;   Ouvre l'importateur de piece et importe le fichier selon le nom mis en argument
;

; CSV à importer selon la valeur en argument
$toImport = "MERC_1ER_FEV_MDF_34_DOS_ERABLE.txt"	; Pour fin de tests et debug seulement. Doit être changé par argument

if $cmdLine[0] > 0 Then ; Chargement de l'argument en ligne de commande
   $toImport = $cmdLine[1]
EndIf

; Il est nécessaire de faire les appels à partir de ce répertoire
FileChangeDir("c:\V90\FABRIDOR")

; Ouvre l'importateur de CSV de CutRite
Run("c:\V90\import.exe /PARTS")
;RunAs("microvellum","cuisineideale.cabcor","Cuisine123",$RUN_LOGON_INHERIT ,"c:\V90\import.exe /PARTS", "c:\V90\FABRIDOR")

; Attend que l'importateur soit ouvert
WinWait("[TITLE:Import - parts]")

; Renommer la fenetre car les prochaines fenetre ont le meme nom
WinSetTitle("[TITLE:Import - parts]", "", "CSVImportMain")

; Prend le controle de la liste des CSV et choisi celui à importer
$hList = ControlGetHandle("[TITLE:CSVImportMain]", "", "SysListView321")
For $i = 0 To _GUICtrlListView_GetItemCount($hList) - 1
   if _GUICtrlListView_GetItemText($hList,$i) = $toImport Then
	  _GUICtrlListView_SetItemSelected ($hList,$i)
   EndIf
Next

; Click sur le bouton d'importation
WinActivate("[TITLE:CSVImportMain]")
ControlClick("[TITLE:CSVImportMain]", "", "ToolbarWindow321", "left",1 ,74,25)
Sleep(100)

; Attendre la fenetre de confirmation (si confirmation il y a)
WinWait("[TITLE:Import]", "Replace existing data", 1)

; Si confirmation nécessaire
IF WinExists("[TITLE:Import]", "Replace existing data") Then
   ControlClick("[TITLE:Import]", "Replace existing data", "Button3", "left",1 ,45,13)	;Clique sur le bouton Oui
EndIf

; Attendre que la fenetre confirmation d'ouverture arrive
WinWait("[TITLE:Import - parts]", "Import finished")

;Cliquer sur le bouton Non pour ne pas ouvrir le visualiseur
ControlClick("[TITLE:Import - parts]", "Import finished", "Button2", "left",1 ,40,18)

; Fermeture de la fenetre d'importation
Sleep(500)
WinKill("[TITLE:CSVImportMain]")

; Sortie
Exit

