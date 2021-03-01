#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x/NT
; Author:         Mathieu Grenier
;
; Script Function:
;   Ouvre l'optimisateur de porte
;

#include <Constants.au3>
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
	Local $batchName = _PathSplit($partListFileName, $temp, $temp, $temp, $temp)[$PATH_FILENAME]
	Debug(@ScriptName & ' started with batch "' & $batchName & '".' & @CRLF)

	Local $attemptCounter = 1
	Local $MAX_ATTEMPTS = 5
	Local $processID = Null
	Local $partListWindow = 0
	While Not $partListWindow
		Local $temp = CutRite_StartPanelExe($partListFileName)
		$processID = $temp[0]
		$partListWindow = $temp[1]
		If @error <> 0 Then
			ProcessClose($processID)
			Debug('The program failed to produce the expected main window. Process ' & $processID & ' was terminated.' & @CRLF)
			If $attemptCounter <= $MAX_ATTEMPTS Then
				Debug('Retrying...' & @CRLF)
				$attemptCounter += 1
				ContinueLoop
			Else
				Debug(@ScriptName & ' failed.' & @CRLF)
				ExitWithCodeAndMessage(1, 'The main window failed to appear ' & $MAX_ATTEMPTS & " " & (($MAX_ATTEMPTS <= 1) ?  "time" : "times in a row") & ".")
			EndIf
		EndIf
	WEnd

	; Attendre la fenetre de revue. Durant ce temps, vérifier si la fenêtre d'erreur de RunCalc ouvre
	Debug("Optimisation process starting." & @CRLF)
	Local $errorMessage = ""
	Local $reviewRunsWindow = CutRite_Optimize($partListWindow, $batchName)
	If @error <> 0 Then
		$errorMessage = $temp
	Else
		Debug("Optimisation process done." & @CRLF)
	EndIf


	KillWindow($partListWindow)
	KillWindow($reviewRunsWindow)
	KillProcess($processID)

	; Kill process "Report.exe" if it becomes a zombie.
	Local $i = 1
	Local $processName = "Report.exe"
	While($i <= $MAX_ATTEMPTS And ProcessExists($processName))
		$i += 1
		If ProcessClose($processName) == 0 Then
			Debug('Failed to kill zombie process "' & $processName & '" on attempt ' & $i & ' of ' & $MAX_ATTEMPTS & '.' & @CRLF)
			Sleep(500)
		Else
			Debug('Succeeded to kill zombie process "' & $processName & '" on attempt ' & $i & ' of ' & $MAX_ATTEMPTS & '.' & @CRLF)
			ExitLoop
		EndIf
	WEnd

	If $errorMessage <> "" And $errorMessage <> Null Then
		ExitWithCodeAndMessage(5, $errorMessage)
	ElseIf WinExists($partListWindow) Then
		Debug('Could not close window with handle ' & $partListWindow & '.' & @CRLF)
		ExitWithCodeAndMessage(3, 'Could not close window with handle ' & $partListWindow & '.')
	ElseIf WinExists($reviewRunsWindow) Then
		Debug('Could not close window with handle ' & $reviewRunsWindow & '.' & @CRLF)
		ExitWithCodeAndMessage(2, 'Could not close window with handle ' & $reviewRunsWindow & '.')
	ElseIf ProcessExists($processID) Then
		Debug('Could not terminate process with id ' & $processID & '.' & @CRLF)
		ExitWithCodeAndMessage(4, 'Could not terminate process with id ' & $processID & '.')
	Else
		; Quitter le script
		Debug(@ScriptName & ' completed successfuly.' & @CRLF)
		ExitWithCodeAndMessage(0)
	EndIf
EndFunc