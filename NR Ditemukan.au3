#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <File.au3>
#include <Excel.au3>

$oExcel = _Excel_Open()
$FilePath = _Excel_BookOpen($oExcel, @ScriptDir & "\gabung.xlsx")
$HitungRow = $oExcel.ActiveSheet.UsedRange.Rows.Count

Func tulis()
$file = FileOpen(@ScriptDir & "\sumber.csv")
$fileread = FileRead($file)
$delete = StringReplace($fileread, @CR, '')
$delete = StringReplace($delete, @LF, '')
$caritambah = StringReplace($delete, '"FK', @CRLF&'"FK')
FileWrite(@ScriptDir & "\testing.txt", $caritambah)
EndFunc

Func regex()
   $file = FileOpen(@ScriptDir & "\testing.txt")
   $fileread = FileRead($file)
   For $i = 1 To $HitungRow
   $noseri = _Excel_RangeRead($FilePath, Default, "F" & $i)
   $carinoseri = StringRegExp($fileread, '("FK.*' & $noseri & '.*)', 1)
   If $carinoseri <> 0 Then
	  FileWrite(@ScriptDir & "\separo.txt", $carinoseri[0])
   EndIf
Next
EndFunc

Func regex2()
   $filex = FileOpen(@ScriptDir & "\separo.txt")
   $filereadx = FileRead($filex)
   $caritambah = StringReplace($filereadx, '"FK', @CRLF&'"FK')
   $caritambah = StringReplace($caritambah, '"FAPR', @CRLF&'"FAPR')
   $caritambah = StringReplace($caritambah, '"OF', @CRLF&'"OF')
   FileWrite(@ScriptDir & "\hasil.csv", $caritambah)
EndFunc
regex()
regex2()