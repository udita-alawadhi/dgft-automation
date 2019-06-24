#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#include<excel.au3>

;Local $var = "C:\Users\DELL\Desktop\seleniumtesting.xlsx"
;Local $oExcel = _Excel_Open()
;Local $oWorkbook = _Excel_BookOpen($oExcel, $var)

Sleep(1000)
ControlFocus("Print", "", "ComboBox1")
ControlCommand("Print", "", "ComboBox1", "ShowDropDown")
ControlCommand("Print", "", "ComboBox1", "SelectString", "Microsoft Print to PDF")
Send("{ENTER}")
ControlClick("Print", "", "Button10")
Sleep(2500)
ControlFocus("Save Print Output As", "", "ComboBox1")
;This is for comment.

;Local $read = _Excel_RangeRead($oWorkbook,"Sheet1", "C3") ;This should work in a loop and preferably take value from the python code itself.
;Send("file")
;Send("{ENTER}")
;Sleep(500)
;Send("!+{F4}",0) ;how to send ALT+F4
;ControlSend("Statement of Bank Realisation - Mozilla Firefox", "", "", "!+{F4}")

