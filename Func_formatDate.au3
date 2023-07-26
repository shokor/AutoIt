#include <Date.au3>

; Scripts to test the function ... Remove ';' in Lines 4-20 for testing.
; Local $formattedDate = formatDate('7/1/2023',1)
; MsgBox(false,'Result',$formattedDate)
; 
; Local $formattedDate = formatDate('7/1/2023',2)
; MsgBox(false,'Result',$formattedDate)
; 
; Local $formattedDate = formatDate('7/1/2023',3)
; MsgBox(false,'Result',$formattedDate)
; 
; Local $formattedDate = formatDate('7/1/2023',4)
; MsgBox(false,'Result',$formattedDate)
; 
; Local $formattedDate = formatDate('7/1/2023',5)
; MsgBox(false,'Result',$formattedDate)
; 
; Local $formattedDate = formatDate('2/30/2023',5)	; Sending an invalid date.
; MsgBox(false,'Result',$formattedDate)


; Function formatDate
; Parameters: 	$pDate		:	String: date in MM/DD/YYYY format. Month and Date can be 1 digit number. For example, both 07/01/2023 and 7/1/2023 are acceptable.
;						$pFormat	:	1 yyyymmdd			(zero filled, 20230701)
;										 	2 mm-dd-yyyy		(zero filled, for example 07-01-2023)
;										 	3 mm - dd - yyyy	(zero filled, for example 07 - 01 - 2023)
;										 	4 mmddyyyy			(zero filled, for example 07012023)
;										 	else mm/dd/yyyy	(zero filled, for example 07/01/2023)
Func formatDate($pDate, $pFormat)
	Local $newDate = ''
	$delim1 = StringInStr ($pDate, "/" , 0 , 1)
	$delim2 = StringInStr ($pDate, "/" , 0 , 2)
	$len = StringLen ($pDate)
	$m = StringMid ($pDate, 1 ,$delim1-1)
	$d = StringMid ($pDate, $delim1+1 ,$delim2-$delim1-1)
	$y = StringMid ($pDate, $delim2+1 ,$len-$delim2)

	If (StringLen($m)<2) Then
		$m = '0' & $m
	EndIf
	
	If (StringLen($d)<2) Then
		$d = '0' & $d
	EndIf

   If _DateIsValid($y &'/' & $m &'/' & $d) Then
        Select 
			Case $pFormat = 1
				$newDate = $y & $m & $d
			Case $pFormat = 2
				$newDate = $m &'-' & $d &'-' & $y
			Case $pFormat = 3
				$newDate = $m &' - ' & $d &' - ' & $y
			Case $pFormat = 4
				$newDate = $m & $d & $y
			Case Else
				$newDate = $m &'/' & $d &'/' & $y
		EndSelect		
	Else
        MsgBox($MB_SYSTEMMODAL, "Valid Date", "The specified date is invalid.")
	EndIf

	Return $newDate
EndFunc	

