#include <File.au3>
#include <Excel.au3>
#include <Array.au3>
#include <File.au3>
#include <StringConstants.au3>

Func exitHot()
   Exit
EndFunc   ;==>exitHot
HotKeySet("{Esc}", "exitHot")

;~ Call(openKbank)
;~ Sleep(4000)
Call(excelToNotepadKbank)

Func openKbank()

   Call(removeKbank)

   $path = "C:\Users\User\Downloads\saving_account.csv"
   $excelFile = FileExists($path)

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

   $nTab = 29
   While 1
	  Send("{TAB "&$nTab&"}")
	  Send("{ENTER}")
	  Sleep(5000)
	  If FileExists($path) Then
		 ExitLoop
	  Else
		 $nTab += 1
	  EndIf
   WEnd

   Sleep(4000)
   Send("{TAB 9}")
   Send("{ENTER}")
   Sleep(4000)
   Send("!+{F4}",0)

EndFunc

Func openScbHostinglotus()
   $checkRun = True
   Call(removeScbHostinglotus)
   Local $username = "nanrat2758"
   Local $password = "Zai@2021"
   $chrome = ShellExecute("chrome.exe", "https://www.scbbusiness.com/auth/login --new-window --start-fullscreen")
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
   $data = Run("notepad.exe")
   Sleep(1000)
   Send("^v")
   Sleep(1000)
   Send("^s")
   Sleep(1000)
   Send(@desktopdir & "\hostinglotus.txt")
   Sleep(500)
   Send("{ENTER}")
   Sleep(500)
   ProcessClose($data)
   Sleep(1000)
   Send("!+{F4}",0)
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
   $data = Run("notepad.exe")
   Sleep(1000)
   Send("^v")
   Sleep(1000)
   Send("^s")
   Sleep(1000)
   Send(@desktopdir & "\metrabyte.txt")
   Sleep(500)
   Send("{ENTER}")
   Sleep(1000)
   ProcessClose($data)
   Sleep(1000)
   Send("!+{F4}",0)
EndFunc

Func excelToNotepadKbank()

   Global $checkArrayNull = 0
   Global $checkCollumn = 0
   Global $rows
   $pathMainStatement = "C:\Users\User\Desktop\KbankMetrabyte.txt"
   $linePreviousStatement = FileReadLine($pathMainStatement,1)
   $path = "C:\Users\User\Downloads\saving_account.csv"
   $pathMetrabyteKbank = "‪C:\Users\User\Desktop\KbankMetrabyte.txt"

   $oExcel = _Excel_Open()
   $oWorkbook = _Excel_BookOpen($oExcel,$path)
   Global $NumberOfRows = $oWorkbook.ActiveSheet.UsedRange.Rows.Count
   $aArrayExcel = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange.Rows("8:"&$NumberOfRows), 1)
   $arrayMetrabyteKbank = StringSplit($linePreviousStatement,@TAB,2)

   _ArrayDisplay($aArrayExcel)
   $checkOldFilePattern = StringMid($aArrayExcel[0][0],9,1)
   If $checkOldFilePattern = " " Then
	  ConsoleWrite("Old Pattern")
   Else
	  For $i=0 To $NumberOfRows Step +1
		 if $aArrayExcel[$i][0] = "" Then
			ExitLoop
		 Else
			$date = StringLeft($aArrayExcel[$i][0],8)
			$time = StringMid($aArrayExcel[$i][0],9)
			$year = StringLeft($date,4)
			$day = StringMid($date,5,2)
			$month = StringRight($date,2)
			$hour = StringLeft($time,2)
			$minutes = StringMid($time,3,2)
			$aArrayExcel[$i][0] = $year&"/"&$month&"/"&$day&" "&$hour&":"&$minutes&":"&"00"
		 EndIf
	  Next
   EndIf

   $checkWithdraw= _Excel_RangeRead($oWorkbook,1,"D8:D"&$NumberOfRows)
   For $i=0 To $NumberOfRows-8 Step + 1
	  If $checkWithdraw[$i] = "" Then
		 $rows = $NumberOfRows-9
	  Else
		 $checkCollumn += 1
		 if $checkCollumn > 1 Then
			$rows = $NumberOfRows-10
			ExitLoop
		 EndIf
	  EndIf
   Next

   Global $newDataArray[$rows][7]
   For $i = 0 To $rows-1 Step +1
	  If $arrayMetrabyteKbank[0] = $aArrayExcel[$i][0] And $arrayMetrabyteKbank[1] = $aArrayExcel[$i][2] And $arrayMetrabyteKbank[2] = $aArrayExcel[$i][4] And $arrayMetrabyteKbank[3] = $aArrayExcel[$i][5] And $arrayMetrabyteKbank[4] = $aArrayExcel[$i][6] Then
		 ConsoleWrite("Match")
		 ExitLoop
	  Else
		 If $aArrayExcel[$i][3] = "" Then
			For $j=0 To 6 Step +1
			   $newDataArray[$i][$j] = $aArrayExcel[$i][$j]
			Next
		 Else
			ContinueLoop
		 EndIf
	  EndIf
   Next

   For $i=0 To $rows-1 Step +1
	  If $newDataArray[$i][0] = "" Then
		 $checkArrayNull += 1
	  Else
		 ContinueLoop
	  EndIf
   Next

   While(True)
	  If $checkArrayNull = $rows Then
		 ConsoleWrite("Null")
		 ExitLoop
	  Else
		 $data = Run ( "notepad.exe " & $pathMainStatement, @WindowsDir )
		 Sleep(3000)
		 Send("{ENTER}")
		 Send("{UP}")
		 For $i=0 To $rows-1 Step +1
			If $newDataArray[$i][0] = "" Then
			   ContinueLoop
			EndIf
			For $j=0 To 6 Step +1
			   If $j = 1 or $j = 3  Then
				  ContinueLoop
			   EndIf
			   Send($newDataArray[$i][$j])
			   Send("{TAB}")
			Next
			Send("{BS}")
			Send("{ENTER}")
		 Next
		 Send("{BS}")
		 ExitLoop
	  EndIf
   WEnd

   If $checkArrayNull = $rows Then
	  Sleep(1000)
	  Send("!+{F4}",0)
	  ConsoleWrite("Null")
   Else
	  Sleep(5000)
	  Send("^s")
	  Sleep(5000)
	  ProcessClose($data)
	  Sleep(5000)
	  Send("!+{F4}",0)
   EndIf

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




