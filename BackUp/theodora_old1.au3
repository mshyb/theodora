#include "..\UDF\Tika\Tika.au3"
#include <Inet.au3>
#include "..\UDF\MIME\_filegetmimetype.au3"
#include "..\UDF\mshytools\basics.au3"
#include "..\UDF\Registry\RegEnumKeyValEx.au3"

Global 	$LIBRE_LOCATION_EXE = _libre_getloc()

Func GetText($sfile)
Local $text = TikaStreamtext($sfile)
	if @extended = 422 then $text = _Libre_string($sfile)
	ConsoleWrite($text & @CRLF)
EndFunc

Func _libre_getloc()
	Local $reg_libre_loc = _RegEnumKeyEx('HKEY_LOCAL_MACHINE\SOFTWARE\LibreOffice\LibreOffice')
	if IsArray($reg_libre_loc) Then
		return RegRead($reg_libre_loc[1], 'Path')
	endif
EndFunc

Func _Libre_string($sInFile)
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
	;ConsoleWrite($str_cmd & @CRLF)
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
	;ConsoleWrite($sOutput & @CRLF)
	If FileExists($tmpfile) Then
		Local $hndtxt = FileOpen($tmpfile)
		Local $txt = FileRead($hndtxt)
		FileClose($hndtxt)
		FileDelete($tmpfile)
		Return $txt
	EndIf
	Return -1

EndFunc   ;==>_LIBRE_TO_MEM