#include-once
#include <Constants.au3>
#include "toolbox.au3"

AutoItSetOption("MustDeclareVars", 1)

Local Const $CUT_RITE_PANEL_EXE_PATH = "C:\V90\panel.exe"
Local Const $CUT_RITE_WORKING_DIRECTORY = "C:\V90\FABRIDOR"
Local Const $CUT_RITE_ERROR_LIST_CLASS_NAME = "GXWND"
Local Const $CUT_RITE_EDIT_CLASS_NAME = "Edit"
Local Const $CUT_RITE_BUTTON_CLASS_NAME = "Button"
Local Const $CUT_RITE_PART_LIST_WINDOW_TITLE_PART_1 = "Part list - "
Local Const $CUT_RITE_PART_LIST_WINDOW_TOOL_BAR_CLASS_NAME = "ToolbarWindow32"
Local Const $CUT_RITE_MODAL_WINDOW_CLASS_NAME = "#32770"
Local Const $CUT_RITE_GHOST_WINDOW_CLASS_NAME = "Ghost"
Local Const $CUT_RITE_REVIEW_RUNS_WINDOW_TITLE = "Review runs" 
Local Const $CUT_RITE_RUNCALC_WINDOW_TITLE = "RUNCALC" 
Local Const $CUT_RITE_RUNCALC_PROCESS_NAME = "RUNCALC.EXE"
Local Const $CUT_RITE_REVIEW_RUNS_WINDOW_TOOL_BAR_CLASS_NAME = "Afx:ToolBar:400000:8:10003:10"
Local Const $CUT_RITE_TRANSFER_TO_MACHINING_CENTER_WINDOW_TITLE = "Transfer to machining centre"
Local Const $CUT_RITE_TRANSFER_TO_MACHINING_CENTER_ERROR_WINDOW_TITLE = "Transfer to machining centre"
Local Const $CUT_RITE_OVERWRITE_PANELS_WITH_PANELS_FROM_LIBRARY = "Part list"
Local Const $CUT_RITE_OVERWRITE_PANELS_CONFIRMATION_WINDOW = "Error"
Local Const $CUT_RITE_PREOPTIMISE_WINDOW_TITLE_PART_1 = "Optimise - "
Local Const $CUT_RITE_OPTIMISE_WINDOW_TITLE = "Optimise"
Local Const $CUT_RITE_OPTIMISE_ERROR_WINDOW_TITLE = "Error"
Local Const $WINDOWS_BATCH_PROCESS_UNCAUGHT_EXCEPTION_WINDOW_TITLE = "BATCH"

Func CutRite_RetrieveErrorMessageFromGXWNDWindow($errorWindow)
	Local $coord = [100, 30]
	ControlClick($errorWindow, "", "[CLASS:" & $CUT_RITE_ERROR_LIST_CLASS_NAME & "; INSTANCE:1;]", $MOUSE_CLICK_PRIMARY, 1, $coord[0], $coord[1])
	Debug( _
		'Clicked control with class name "' & $CUT_RITE_ERROR_LIST_CLASS_NAME & '" and instance number 1 once with ' & $MOUSE_CLICK_PRIMARY & ' button at relative coordinates (' & $coord[0] & _
		', ' & $coord[1] & ').' & @CRLF _
	)

	Sleep(500)
	Local $errorDescription = ControlGetText($errorWindow, "", "[CLASS:" & $CUT_RITE_EDIT_CLASS_NAME & "; INSTANCE:1;]")
	Debug('Got error message "' & $errorDescription & '" from volatile control with class name "' & $CUT_RITE_EDIT_CLASS_NAME & '" and instance number 1 in window ' & $errorWindow & '.' & @CRLF)

	Local $coord = [400, 30] ; Must redeclare to reassign
	ControlClick($errorWindow, "", "[CLASS:" & $CUT_RITE_ERROR_LIST_CLASS_NAME & "; INSTANCE:1;]", $MOUSE_CLICK_PRIMARY, 1, $coord[0], $coord[1])	;Clique pour obtenir le message d'erreur
	Debug( _
		'Clicked control with class name "' & $CUT_RITE_ERROR_LIST_CLASS_NAME & '" and instance number 1 once with ' & $MOUSE_CLICK_PRIMARY & ' button at relative coordinates (' & $coord[0] & _
		', ' & $coord[1] & ').' & @CRLF _
	)

	Sleep(500)
	Local $temp = ControlGetText($errorWindow, "", "[CLASS:" & $CUT_RITE_EDIT_CLASS_NAME & "; INSTANCE:1;]")
	Debug( _
		'Got error message "' & $temp & '" from volatile control with class name "' & $CUT_RITE_EDIT_CLASS_NAME & '" and instance number 1 in window with handle ' & $errorWindow & '.' & _
		@CRLF _
	)
	$errorDescription &= "	" & $temp

	Local $coord = [40, 18]
	ControlClick($errorWindow, "", "[CLASS:" & $CUT_RITE_BUTTON_CLASS_NAME & "; INSTANCE:2;]", $MOUSE_CLICK_PRIMARY, 1, $coord[0], $coord[1])
	Debug( _
		'Clicked control with class name "' & $CUT_RITE_BUTTON_CLASS_NAME & '" and instance number 2 once with ' & $MOUSE_CLICK_PRIMARY & ' button at relative coordinates (' & $coord[0] & _
		', ' & $coord[1] & ').' & @CRLF _
	)

	Return $errorDescription
EndFunc

Func CutRite_StartPanelExe($partListFileName = Null)
	; Ouvre l'importateur de CSV de CutRite.
	Local $command = '"' & $CUT_RITE_PANEL_EXE_PATH & '"' & (($partListFileName <> Null) ? ' "' & $partListFileName & '"' : '')
	Local $processID = Run($command, $CUT_RITE_WORKING_DIRECTORY)
	Debug('Executed command "' & $command & '" in working directory "' & $CUT_RITE_WORKING_DIRECTORY & '".' & @CRLF)

	; Attendre que l'optimisateur ouvre.
	Local $temp
	Local $mainWindowTitle = $CUT_RITE_PART_LIST_WINDOW_TITLE_PART_1 & _PathSplit($partListFileName, $temp, $temp, $temp, $temp)[3]
	Local $mainWindow = WinWait("[TITLE:" & $mainWindowTitle & "]", "", 5)
	
	Local $temp = [$processID, $mainWindow]
	If $mainWindow == 0 Then 
		Return SetError(1, 0, $temp)
	Else
		Debug('Found window with title "' & $mainWindowTitle & '" and handle ' & $mainWindow & '.' & @CRLF)
		Return $temp
	EndIf
EndFunc

Func CutRite_RefreshPanels($mainWindow)
	Local $mainWindowTitle = WinGetTitle($mainWindow) ; Doit être prélevé avant qu'il change.

	Local $coord = [394, 21]
	ControlClick($mainWindow, "", "[CLASS:" & $CUT_RITE_PART_LIST_WINDOW_TOOL_BAR_CLASS_NAME & "; INSTANCE:1;]", $MOUSE_CLICK_PRIMARY, 1, $coord[0], $coord[1])
	Debug( _
		'Clicked control with class "' & $CUT_RITE_PART_LIST_WINDOW_TOOL_BAR_CLASS_NAME & '" and instance number 1 with ' & $MOUSE_CLICK_PRIMARY & ' button at relative coordinates (' & _
		$coord[0] & ', ' & $coord[1] & ').' & @CRLF _
	)

	; Cliquer sur le bouton Oui si nécessaire.
	While WinExists($mainWindow) And WinGetTitle($mainWindow) == $mainWindowTitle
		Sleep(500)

		Local $confirmationWindow = WinGetHandle("[CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & "; TITLE:" & $CUT_RITE_OVERWRITE_PANELS_WITH_PANELS_FROM_LIBRARY & ";]")
		If Not @error Then
			Debug( _
				'Found window with class "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" and title "' & $CUT_RITE_OVERWRITE_PANELS_WITH_PANELS_FROM_LIBRARY & '" having handle ' & _
				$confirmationWindow & '.' & @CRLF _
			)
			
			ControlClick($confirmationWindow, "", "[CLASS:" & $CUT_RITE_BUTTON_CLASS_NAME & "; INSTANCE:1;]", $MOUSE_CLICK_PRIMARY, 1)
			Debug('Clicked control with class name "' & $CUT_RITE_BUTTON_CLASS_NAME & '" and instance number 1 once with ' & $MOUSE_CLICK_PRIMARY & ' button.' & @CRLF)
		EndIf

		Local $errorWindow = WinGetHandle("[CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & "; TITLE:" & $CUT_RITE_OVERWRITE_PANELS_CONFIRMATION_WINDOW & ";]")
		If Not @error Then
			Debug( _
				'Found window with class name "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" and title "' & $CUT_RITE_OVERWRITE_PANELS_CONFIRMATION_WINDOW & '" having handle ' & $errorWindow & '.' & _
				@CRLF _
			)
			
			Local $errorDescription = CutRite_RetrieveErrorMessageFromGXWNDWindow($errorWindow)
			
			; Fermer la fenetre principale
			KillWindow($mainWindow)
			Debug('Closing window with handle ' & $mainWindow & '.' & @CRLF)

			Return SetError(1, 0, $errorDescription)
		EndIf
	WEnd
	Debug('Toggled to the "Board list" tab.' & @CRLF)
EndFunc

Func CutRite_TransferToMachiningCenter($reviewRunsWindow)
    Local $transferToMachiningCenterWindow = 0
	While Not $transferToMachiningCenterWindow
		WinActivate($reviewRunsWindow)
        Debug('Activated window with handle ' & $reviewRunsWindow & '.' & @CRLF)
		Sleep(100)

        Local $coord = [345, 30]
		ControlClick($reviewRunsWindow, "", "[CLASS:" & $CUT_RITE_REVIEW_RUNS_WINDOW_TOOL_BAR_CLASS_NAME & "; INSTANCE:1;]", $MOUSE_CLICK_PRIMARY, 1, $coord[0], $coord[1])
        Debug( _
			'Clicked control with class name "' & $CUT_RITE_REVIEW_RUNS_WINDOW_TOOL_BAR_CLASS_NAME & '" and instance number 1 once with ' & $MOUSE_CLICK_PRIMARY & _
			' button at relative coordinates (' & $coord[0] & ', ' & $coord[1] & ').' & @CRLF _
		)
		Sleep(400)
		$transferToMachiningCenterWindow = WinGetHandle("[TITLE:" & $CUT_RITE_TRANSFER_TO_MACHINING_CENTER_WINDOW_TITLE & "; CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & ";]")
	WEnd
	Debug( _
		'Found window with class name "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" and title "' & $CUT_RITE_TRANSFER_TO_MACHINING_CENTER_WINDOW_TITLE & '" having handle ' & _
		$transferToMachiningCenterWindow & '.' & @CRLF _
	)

	While WinExists($transferToMachiningCenterWindow)
		Sleep(100)

		Local $errorWindow = WinGetHandle("[TITLE:" & $CUT_RITE_TRANSFER_TO_MACHINING_CENTER_ERROR_WINDOW_TITLE & "; CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & ";]")
		If Not @error Then
			Local $errorList = ControlGetHandle($errorWindow, "", "[CLASS:" & $CUT_RITE_ERROR_LIST_CLASS_NAME & "; INSTANCE:1;]")
			If $errorList Then
				Debug( _
					'Found control with class name "' & $CUT_RITE_ERROR_LIST_CLASS_NAME & '" and instance number ' & $errorListInstance & ' having handle ' & $errorList & _
					' in ambiguous window with class name "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" and title "' & $CUT_RITE_TRANSFER_TO_MACHINING_CENTER_ERROR_WINDOW_TITLE & _
					'" having handle ' & $errorWindow & '.' & @CRLF _
				)
				Return SetErrror(1, 0, CutRite_RetrieveErrorMessageFromGXWNDWindow($errorWindow))
			EndIf
		EndIf
	WEnd
EndFunc

Func CutRite_Optimize($partListWindow)
	Local $coord = [488, 25]
	ControlClick($partListWindow, "", "[CLASS:" & $CUT_RITE_PART_LIST_WINDOW_TOOL_BAR_CLASS_NAME & "; INSTANCE:1;]", $MOUSE_CLICK_PRIMARY, 1, $coord[0], $coord[1])
	Debug( _
		'Clicked control with class name "' & $CUT_RITE_PART_LIST_WINDOW_TOOL_BAR_CLASS_NAME & '" and instance number 1 once with ' & $MOUSE_CLICK_PRIMARY & ' button at relative coordinates (' & _
		$coord[0] & ', ' & $coord[1] & ').' & @CRLF _
	)

	Local $batchName = StringRegExpReplace(WinGetTitle($partListWindow), "\A" & $CUT_RITE_PART_LIST_WINDOW_TITLE_PART_1 & "(.*+)\z", "\1", 1)
	Local $preOptimiseWindowTitle = $CUT_RITE_PREOPTIMISE_WINDOW_TITLE_PART_1 & $batchName
	Local $preOptimiseWindow = WinWait("[TITLE:" & $preOptimiseWindowTitle & ";]", "", 5)
	If Not $preOptimiseWindow Then Return SetError(1, 0, 'Failed to obtain window with title "' & $preOptimiseWindowTitle & '".')
	Debug('Found window with title "' & $preOptimiseWindowTitle & '" having handle ' & $preOptimiseWindow & '.' & @CRLF)

	Local $optimiseWindow = 0
	While WinExists($preOptimiseWindow) And Not $optimiseWindow
		Local $optimiseErrorWindow = WinGetHandle("[TITLE:" & $CUT_RITE_OPTIMISE_ERROR_WINDOW_TITLE & "; CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & ";]")
		If $optimiseErrorWindow Then
			Debug( _
				'Found optimization error window with class name "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" and title "' & $CUT_RITE_OPTIMISE_ERROR_WINDOW_TITLE & '" having handle ' & _
				$optimiseErrorWindow & '.' & @CRLF _
			)

			Local $errorDescription = CutRite_RetrieveErrorMessageFromGXWNDWindow($optimiseErrorWindow)
			
			; Fermer la fenetre d'optimisation
			Sleep(500)
			WinKill($preOptimiseWindow)
			Debug('Closed window with handle ' & $preOptimiseWindow & "." & @CRLF)
			
			Do
				Local $windowsProcessFailedWindow = WinWait("[TITLE:" & $WINDOWS_BATCH_PROCESS_UNCAUGHT_EXCEPTION_WINDOW_TITLE & "; CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & ";]", "", 2)
				If $windowsProcessFailedWindow Then
					Debug( _
						'Found window with class name "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" and title "' & $WINDOWS_BATCH_PROCESS_UNCAUGHT_EXCEPTION_WINDOW_TITLE & '" having handle ' & _
						$windowsProcessFailedWindow & '.' & @CRLF _
					)
					WinKill($windowsProcessFailedWindow)
					Debug('Closed window with handle ' & $windowsProcessFailedWindow & '.' & @CRLF)
				EndIf
			Until Not $windowsProcessFailedWindow
			
			; Message d'erreur
			Return SetError(1, 0, $errorDescription)
		EndIf

		$optimiseWindow = WinGetHandle("[TITLE:" & $CUT_RITE_OPTIMISE_WINDOW_TITLE & "; CLASS:" & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & ";]")
	WEnd
	Debug('Found window with title "' & $CUT_RITE_OPTIMISE_WINDOW_TITLE & '" and class name "' & $CUT_RITE_MODAL_WINDOW_CLASS_NAME & '" having handle ' & $optimiseWindow & '.' & @CRLF)

	Local $reviewRunsWindow = 0
	While Not $reviewRunsWindow
		Local $runCalcWindow = WinGetHandle("[TITLE:" & $CUT_RITE_RUNCALC_WINDOW_TITLE & ";]")
		If $runCalcWindow Then
			Debug( _
				'Found window with title "' & $CUT_RITE_RUNCALC_WINDOW_TITLE & '" having handle ' & $runCalcWindow & ' meaning that the process "' & $CUT_RITE_RUNCALC_PROCESS_NAME & _
				'" failed.' & @CRLF _
			)

			; Si erreur RUNCALC, fermer tout et envoyer un message d'erreur à la console
			KillWindow($runCalcWindow)
			Debug('Closed window with handle ' & $runCalcWindow & '.' & @CRLF)

			$reviewRunsWindow = WinWait("[TITLE:" & $CUT_RITE_REVIEW_RUNS_WINDOW_TITLE & ";]", "", 5)
			If $reviewRunsWindow Then
				Debug('Found window with title "' & $CUT_RITE_REVIEW_RUNS_WINDOW_TITLE & '" having handle ' & $reviewRunsWindow & '.' & @CRLF)

				KillWindow($reviewRunsWindow)
				Debug('Closed window with handle ' & $reviewRunsWindow & '.' & @CRLF)
			EndIf

			Return SetError(2, 0, 'Process "RUNCALC.EXE" crashed unexpectedly.')
		EndIf
		
		Sleep(500)
		$reviewRunsWindow = WinGetHandle("[TITLE:" & $CUT_RITE_REVIEW_RUNS_WINDOW_TITLE & ";]")
	WEnd
	Debug('Found window with title "' & $CUT_RITE_REVIEW_RUNS_WINDOW_TITLE & '" having handle ' & $reviewRunsWindow & '.' & @CRLF)

	Return $reviewRunsWindow
EndFunc