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
   $pathMainStatement = "C:\Users\User\Desktop\KbankMetrabyte.txt"
   $linePreviousStatement = FileReadLine($pathMainStatement,1)
   $path = "C:\Users\User\Downloads\saving_account.csv"
   $pathMetrabyteKbank = "‪C:\Users\User\Desktop\KbankMetrabyte.txt"

   $oExcel = _Excel_Open()
   $oWorkbook = _Excel_BookOpen($oExcel,$path)
   $aArray = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange.Rows("8:23"), 1)
;~    _ArrayDisplay($aArray, "$aArray")
   $arrayMetrabyteKbank = StringSplit($linePreviousStatement,@TAB,2)
;~    _ArrayDisplay($arrayMetrabyteKbank, "$arrayMetrabyteKbank")

   if $arrayMetrabyteKbank[2] = $aArray[0][2] Then
	  ConsoleWrite("True")
   Else
	  ConsoleWrite("false")
   EndIf

;~    Run("notepad.exe")
;~    WinWaitActive("Untitled - Notepad")

;~    $row = 8
;~    While(True)
;~ 	  use excel rangeread line $row and check linePreviousStatement
;~    WEnd

;~    For $i=8 to 23
;~ 	  Send(_Excel_RangeRead($oWorkbook,Default,"A"&$i))
;~ 	  Send("{TAB 2}")
;~ 	  Send(_Excel_RangeRead($oWorkbook,Default,"B"&$i))
;~ 	  Send("{TAB}")
;~ 	  Send(_Excel_RangeRead($oWorkbook,Default,"C"&$i))
;~ 	  Send("{TAB}")
;~ 	  Send(_Excel_RangeRead($oWorkbook,Default,"E"&$i))
;~ 	  Send("{TAB}")
;~ 	  Send(_Excel_RangeRead($oWorkbook,Default,"F"&$i))
;~ 	  Send("{TAB}")
;~ 	  Send(_Excel_RangeRead($oWorkbook,Default,"G"&$i))
;~ 	  Send("{ENTER}")
;~    Next

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




