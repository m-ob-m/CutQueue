#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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

#include <File.au3>
#include "CutRite.au3"
#include "ToolBox.au3"

AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", $OPT_MATCHEXACT)

; Enables log files.
Local $DEBUG = True

Main()

Func Main()
	Local $partListFileName = ($cmdLine[0] > 0) ? $cmdLine[1] : "71999P.txt" ; CSV à importer selon la valeur en argument.
	Local $temp
	Debug(@ScriptName & ' started with batch "' & _PathSplit($partListFileName, $temp, $temp, $temp, $temp)[3] & '".' & @CRLF)

	Local $attemptCounter = 1
	Local Const $MAX_ATTEMPTS = 5
	Local $processID = Null
	Local $mainWindow = 0
	While Not $mainWindow
		Local $temp = CutRite_StartImportExe()
		$processID = $temp[0]
		$mainWindow = $temp[1]
		If @error <> 0 Then
			KillWindowAndProcess($mainWindow, $processID)
			Debug('The program failed to produce the expected main window. Process ' & $processID & ' was terminated.' & @CRLF)
			If $attemptCounter <= $MAX_ATTEMPTS Then
				Debug('Retrying...' & @CRLF)
				$attemptCounter += 1
				ContinueLoop
			Else
				Debug(@ScriptName & ' failed.' & @CRLF)
				ExitWithCodeAndMessage(1, 'The main window failed to appear ' & $MAX_ATTEMPTS & " " & (($MAX_ATTEMPTS <= 1) ? "time" : "times in a row") & ".")
			EndIf
		EndIf
	WEnd

	Local $errorMessage = CutRite_ImportPartsFromCSV($mainWindow, $partListFileName)
	If @error Then
		Debug(@ScriptName & ' failed.' & @CRLF)
		ExitWithCodeAndMessage(2, $errorMessage)
	EndIf

	; Fermer la fenêtre principale.
	KillWindowAndProcess($mainWindow, $processID)
	Sleep(500)
	If WinExists($mainWindow) Then
		Debug('Could not close window with handle ' & $mainWindow & '.' & @CRLF)
		ExitWithCodeAndMessage(3, 'Could not close window with handle ' & $mainWindow & '.')
	ElseIf ProcessExists($processID) Then
		Debug('Could not terminate process with id ' & $processID & '.' & @CRLF)
		ExitWithCodeAndMessage(4, 'Could not terminate process with id ' & $processID & '.')
	Else
		; Quitter le script
		Debug(@ScriptName & ' completed successfuly.' & @CRLF)
		ExitWithCodeAndMessage(0)
	EndIf
EndFunc