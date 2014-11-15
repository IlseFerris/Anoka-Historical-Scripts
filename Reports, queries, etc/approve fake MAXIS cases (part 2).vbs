'VARIABLES TO DECLARE

PND2_row = 7

EMConnect ""

Do

EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then
  MsgBox "Not on PND2"
  StopScript
End if

EMReadScreen case_number, 8, PND2_row, 5
case_number = trim(case_number)
If case_number = "" then stopscript

EMWriteScreen "stat", 20, 13
EMWriteScreen "________", 20, 33
EMWriteScreen case_number, 20, 33
EMWriteScreen "wreg", 20, 71
EMSendKey "<enter>"
EMWaitReady 1, 1

EMWriteScreen "nn", 20, 79
EMSendKey "<enter>"
EMWaitReady 1, 1

EMWriteScreen "Y", 6, 68
EMWriteScreen "30", 8, 50
EMWriteScreen "N", 8, 80
EMWriteScreen "09", 13, 50
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<enter>"
EMWaitReady 1, 1

EMSendKey "<PF3>"
EMWaitReady 1, 1

EMSendKey "<PF3>"
EMWaitReady 1, 1
  
EMWriteScreen "rept", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "PND2", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

PND2_row = PND2_row + 1

Loop until case_number = "" or PND2_row = 19