#include <Constants.au3>
#include <GUIListView.au3>
;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Mathieu Grenier
;
; Script Function:
;   Convertit un fichier WMF en JPG avec ImageMagick
;

$wmf = "$MARDI_31_JAN_MDF_34_BLANC_20001$.wmf "	; Pour fin de tests et debug seulement. Doit être changé par argument
$jpg = "MARDI_31_JAN_MDF_34_BLANC_2_0001.jpg"	; Pour fin de tests et debug seulement. Doit être changé par argument

if $cmdLine[0] > 0 Then ; Chargement des arguments en ligne de commande
   $wmf = $cmdLine[1]
   $jpg = $cmdLine[2]
EndIf

; Il est nécessaire de faire les appels à partir de ce répertoire
FileChangeDir("C:\ImageMagick\")

; Convertir l'image et enlever le gris (modulate = "Brighten")
; Il est important de conserver l'ordre des arguments. Donc resize modulate sharpen
RunWait("C:\ImageMagick\magick.exe C:\V90\FABRIDOR\SYSTEM_DATA\DATA\" & $wmf & " -resize 50% -modulate 140% -sharpen 0x2 C:\V90\FABRIDOR\SYSTEM_DATA\DATA\" & $jpg,"C:\ImageMagick\",@SW_HIDE)

; Quitter le script
Sleep(500)
Exit