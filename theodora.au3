#include "..\UDF\Tika\Tika.au3"
#include <Inet.au3>
#include "..\UDF\MIME\_filegetmimetype.au3"
#include "..\UDF\mshytools\basics.au3"
#include "..\UDF\Registry\RegEnumKeyValEx.au3"

Global $LIBRE_LOCATION_EXE = _libre_getloc()

;Theodora.exe doc.docx -tw template -o stdout
;GetText()

Func GetText($sfile)
	Local $text = TikaStreamtext($sfile)
	If @extended = 422 Then $text = _Libre_string($sfile)
	ConsoleWrite($text & @CRLF)
EndFunc   ;==>GetText

Func _libre_getloc()
	Local $reg_libre_loc = _RegEnumKeyEx('HKEY_LOCAL_MACHINE\SOFTWARE\LibreOffice\LibreOffice')
	If IsArray($reg_libre_loc) Then
		Return RegRead($reg_libre_loc[1], 'Path')
	EndIf
EndFunc   ;==>_libre_getloc

Func _Libre_string($sInFile, $debug = False)
	;https://help.libreoffice.org/Common/Starting_the_Software_With_Parameters
	Local $f1, $f2, $f3, $f4
	Local $timeout = TimerInit(), $sOutput = ''
	If (Not FileExists($sInFile)) Then Return SetError(4, 0, 0)
	$aInPath = _PathSplit($sInFile, $f1, $f2, $f3, $f4)
	Local $tmpfile = @TempDir & '\' & $f3 & '.txt'
	Local $str_cmd = '"' & $LIBRE_LOCATION_EXE & '"' & _
			' --headless' & _
			' --convert-to txt' & _
			' --outdir "' & @TempDir & '"' & _ ; remove trailing backslash
			' "' & $sInFile & '"'
	If $debug = True then ConsoleWrite($str_cmd & @CRLF)
	Local $iPID = Run($str_cmd, '', @SW_HIDE, BitOR(4, 2))
	While 1
		$sOutput &= StderrRead($iPID)
		$sOutput &= StdoutRead($iPID)
		Select
			Case @error Or (Not ProcessExists($iPID))
				ExitLoop
			Case TimerDiff($timeout) > 10000
				ProcessClose($iPID)
				Return SetError(3, TimerDiff($timeout), 1) ;timeout
		EndSelect
		Sleep(10)
	WEnd
	If $debug = True then ConsoleWrite($sOutput & @CRLF)
	If FileExists($tmpfile) Then
		Local $hndtxt = FileOpen($tmpfile)
		Local $txt = FileRead($hndtxt)
		FileClose($hndtxt)
		FileDelete($tmpfile)
		Return $txt
	EndIf
	Return -1
EndFunc   ;==>_Libre_string

Func _textWalkEx($sText, $aTemplate); need multiple templates
	Local $int_err = 0
	Local $boo_err = 0
	Local $aOut[UBound($aTemplate) - 1]
	For $i = 0 To UBound($aTemplate) - 1
		Local $len = StringLen($aTemplate[$i][0])
		Local $find = StringInStr($sText, $aTemplate[$i][0], 1)
		If $find Then
			Switch $i
				Case (UBound($aTemplate) - 1) ;											last
					;
				Case Else
					Local $find_nex = StringInStr($sText, $aTemplate[$i + 1][0], 1)
					Local $cut = StringMid($sText, _
							$find + $len, _
							($find_nex - $find) - $len)
					$cut = StringReplace($cut, @LF, ' ')
					$cut = StringReplace($cut, @CR, ' ')
					$cut = StringStripWS($cut, $STR_STRIPLEADING + $STR_STRIPTRAILING)
					If StringLen($cut) < $aTemplate[$i][1] Then
						$aOut[$i] = $cut
					Else
						$boo_err = 1
						$int_err += 1
					EndIf
			EndSwitch
		Else
			$boo_err = 1
			$int_err += 1
		EndIf
	Next
	Return SetError($boo_err, $int_err, $aOut)
EndFunc   ;==>_textWalkEx
