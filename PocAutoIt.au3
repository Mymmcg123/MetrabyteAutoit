#include <File.au3>
#include <Excel.au3>
#include <Array.au3>
#include <File.au3>
#include <StringConstants.au3>


Call(excelToNotepad)

Func openKbank()
   Call(removeKbank)
   Local $username = "hostinglotus"
   Local $password = "Fern2523"
   ShellExecute("chrome.exe", "https://online.kasikornbankgroup.com/K-Online/login.jsp?lang=TH&type=sme --new-window --start-fullscreen")
   Sleep(4000)
   Send($username)
   Send("{TAB}")
   Sleep(4000)
   Send($password)
   Send("{ENTER}")
   Sleep(4000)
   Send("{TAB 3}")
   Send("{ENTER}")
   Sleep(4000)
   Send("{TAB 5}")
   Send("{ENTER}")
   Sleep(4000)
   Send("{TAB 23}")
   Send("{ENTER}")
   Send("{DOWN}")
   Send("{ENTER}")
   Sleep(4000)
   Send("{TAB 31}")
   Send("{ENTER}")
EndFunc

Func openScbHostinglotus()
   $checkRun = True
   Call(removeScbHostinglotus)
   Local $username = "nanrat2758"
   Local $password = "Zai@2021"
   ShellExecute("chrome.exe", "https://www.scbbusiness.com/auth/login --new-window --start-fullscreen")
   Sleep(3000)
   Send($username)
   Send("{ENTER}")
   Sleep(2000)
   Send($password)
   Send("{ENTER}")
   Sleep(3000)
   Send("{TAB 4}")
   Send("{ENTER}")
   Sleep(1000)
   Send("{TAB 6}")
   Send("{ENTER}")
   Sleep(2000)
   Send("^a")
   Sleep(500)
   Send("^c")
   Sleep(500)
   Run("notepad.exe")
   Sleep(1000)
   Send("^v")
   Sleep(1000)
   Send("^s")
   Sleep(1000)
   Send(@desktopdir & "\hostinglotus.txt")
   Sleep(500)
   Send("{ENTER}")
EndFunc

Func openScbMetrabyte()
   $checkRun = True
   Call(removeScbMetrabyte)
   Local $username = "nanrat2758"
   Local $password = "Zai@2021"
   ShellExecute("chrome.exe", "https://www.scbbusiness.com/auth/login --new-window --start-fullscreen")
   Sleep(3000)
   Send($username)
   Send("{ENTER}")
   Sleep(2000)
   Send($password)
   Send("{ENTER}")
   Sleep(3000)
   Send("{TAB 2}")
   Send("{ENTER}")
   Sleep(3000)
   Send("{UP}")
   Send("{ENTER}")
   Sleep(3000)
   Send("{TAB 2}")
   Send("{ENTER}")
   Sleep(3000)
   Send("{UP}")
   Send("{ENTER}")
   Sleep(3000)
   Send("{TAB 4}")
   Send("{ENTER}")
   Sleep(1000)
   Send("{TAB 6}")
   Send("{ENTER}")
   Sleep(2000)
   Send("^a")
   Sleep(500)
   Send("^c")
   Sleep(500)
   Run("notepad.exe")
   Sleep(1000)
   Send("^v")
   Sleep(1000)
   Send("^s")
   Sleep(1000)
   Send(@desktopdir & "\metrabyte.txt")
   Sleep(500)
   Send("{ENTER}")
EndFunc

Func excelToNotepad()
   $pathMainStatement = "C:\Users\Mark\Desktop\Metrabyte\MetrabyteAutoit\KbankMetrabyte.txt"
   $linePreviousStatement = FileReadLine($pathMainStatement,1)
   $path = "C:\Users\Mark\Downloads\saving_account.csv"

   $oExcel = _Excel_Open()
   $oWorkbook = _Excel_BookOpen($oExcel,$path)
   $aArrayExcel = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange.Rows("8:30"), 1)
;~    _ArrayDisplay($aArrayExcel, "$aArrayExcel")
   $arrayMetrabyteKbank = StringSplit($linePreviousStatement,@TAB,2)
;~    _ArrayDisplay($arrayMetrabyteKbank, "$arrayMetrabyteKbank")

;~    For $i=0 To 10 Step +1
   If $arrayMetrabyteKbank[0] = $aArrayExcel[22][0] And $arrayMetrabyteKbank[1] = $aArrayExcel[22][1] And $arrayMetrabyteKbank[2] = $aArrayExcel[22][2] And $arrayMetrabyteKbank[3] = $aArrayExcel[22][4] And $arrayMetrabyteKbank[4] = $aArrayExcel[22][5] And $arrayMetrabyteKbank[5] = $aArrayExcel[22][6] Then
	  ConsoleWrite("Match")
   Else
	  ConsoleWrite("Not Match")
   EndIf




;~    if $arrayMetrabyteKbank[1] = $aArrayExcel[4][1] Then
;~ 	  ConsoleWrite("True")
;~    Else
;~ 	  ConsoleWrite("false")
;~    EndIf

;~    $row = 8
;~    While(True)
;~ 	  use excel rangeread line $row and check linePreviousStatement
;~    WEnd



EndFunc

Func removeScbHostinglotus()
   FileRecycle("C:\Users\User\Desktop\hostinglotus.txt")
EndFunc

Func removeScbMetrabyte()
	  FileRecycle("C:\Users\User\Desktop\metrabyte.txt")
EndFunc

Func removeKbank()
   FileRecycle("C:\Users\User\Downloads\saving_account.csv")
EndFunc




