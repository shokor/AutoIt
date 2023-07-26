#include <Date.au3>

; Scripts to test the function ... Remove ';' in Lines 3-16 for testing.
; Local $LogMsg = ''
; Local $LogFilePath = 'H:\Learning\AutoIt\LogFile.txt' 
; Local $result = False
; 
; $LogMsg = 'Test Message 1' 
; $result = updateLogFile($LogMsg, $LogFilePath, 1)
; 
; $LogMsg = 'Test Message 2' 
; $result = updateLogFile($LogMsg, $LogFilePath, 1)
; 
; If $result = True Then 
;	 ShellExecute($LogFilePath)  ; Open log.txt in the user's default app for txt files.
; EndIf

; Function updateLogFile
; Parameters:	$pLogMsg		: String. Log message to write on a log file.
;                  		$pLogFilePath: String. Absolute file path of a log file (txt file)
;                  		$pOpenMode	: Number. 1 (File open in the append mode) or 2 (Write mode = erase previous contents)
; Returns:			True when updated the file successfully. False when any errors occur.

Func updateLogFile($pLogMsg, $pLogFilePath, $pOpenMode)	
		Local $logFileID = FileOpen($pLogFilePath, $pOpenMode)
		If $logFileID = -1 Then
			MsgBox($MB_SYSTEMMODAL, "Error", "An error occurred when reading the file.")
			Return False
		EndIf
		Local $LogMsg =  _NOW() & ' (' & @UserName & ') ' & $pLogMsg & @CRLF ; Adding DateTime, Username, and line break to the passed log message.
		FileWrite($logFileID,$LogMsg)
		FileClose($logFileID)
		Return True
EndFunc
