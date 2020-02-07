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
#include "toolbox.au3"

AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MustDeclareVars", 1)

Local $DEBUG = True

Main()

Func Main()
	Local $partListFileName = ($cmdLine[0] > 0) ? $cmdLine[1] : "71999P.txt" ; CSV à importer selon la valeur en argument.<
	Local $temp
	Local $batchName = _PathSplit($partListFileName, $temp, $temp, $temp, $temp)[3]
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
				ExitWithCodeAndMessage(1, Null, 'The main window failed to appear ' & $MAX_ATTEMPTS & " " & (($MAX_ATTEMPTS <= 1) ?  "time" : "times in a row") & ".")
			EndIf
		EndIf
	WEnd

	; Attendre la fenetre de revue. Durant ce temps, vérifier si la fenêtre d'erreur de RunCalc ouvre
	Debug("Optimisation process starting." & @CRLF)
	Local $temp = CutRite_Optimize($partListWindow)
	If @error Then
		Local $errorMessage = $temp
		KillWindow($partListWindow)
		Debug(@ScriptName & ' failed: ' & $errorMessage & @CRLF)
		ExitWithCodeAndMessage(3, $errorMessage)
	EndIf
	Local $reviewRunsWindow = $temp
	Debug("Optimisation process done." & @CRLF)

	; Transfer to machining center
	Debug("Transfer to machining center process starting" & @CRLF)
	Local $MAX_ATTEMPTS = 5
	Local $i = 0
	Do
		$i += 1
		Local $errorMessage = CutRite_TransferToMachiningCenter($reviewRunsWindow)
		If @error Then
			Debug('The transfer to machining center process returned error message "' & $errorMessage & '" on attempt ' & $i & ' of ' & $MAX_ATTEMPTS & '.' & @CRLF)
		Else
			Debug('The transfer to machining center process succeeded on attempt ' & $i & ' of ' & $MAX_ATTEMPTS & '.' & @CRLF)
		EndIf
	Until($i < $MAX_ATTEMPTS)
	Debug("Transfer to machining center process done." & @CRLF)

	; Close the "Review runs" window.
	KillWindow($reviewRunsWindow)
	Debug('Closed window with handle ' & $reviewRunsWindow & '.' & @CRLF)

	; Exit script
	Debug(@ScriptName & ' completed successfuly.' & @CRLF)
	ExitWithCodeAndMessage(0, ($errorMessage <> "") ? $errorMessage : "OK")
EndFunc