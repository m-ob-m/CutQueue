#include-once
#include <Constants.au3>
#include <File.au3>

AutoItSetOption("MustDeclareVars", 1)

If Not IsDeclared("DEBUG") Then Local $DEBUG = False

Func Debug($text)
	If $DEBUG == True Then
		Local $temp
		Local $handle = FileOpen(@ScriptDir & "\" & _PathSplit(@ScriptFullPath, $temp, $temp, $temp, $temp)[3] & ".log", $FO_APPEND)
		If $handle <> -1 Then
			FileWriteLine($handle, $text)
			FileClose($handle)
		EndIf
	EndIf
EndFunc

Func ExitWithCodeAndMessage($code = 0, $ouputMessage = Null, $errorMessage = Null)
	ConsoleWrite($ouputMessage)
	ConsoleWriteError($errorMessage)
	Exit $code
EndFunc

Func KillWindowAndProcess($windowHandle, $processId)
	KillWindow($windowHandle)
	KillProcess($processId)
EndFunc

Func KillWindow($windowHandle)
	If WinExists($windowHandle) Then
		WinKill($windowHandle)
	EndIf
EndFunc

Func KillProcess($processId)
	If ProcessExists($processId) Then
		ProcessClose($processId)
	EndIf
EndFunc