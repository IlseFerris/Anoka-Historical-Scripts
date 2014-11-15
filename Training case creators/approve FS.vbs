amt_of_times_to_run = 1

EMConnect ""

EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then
  MsgBox "Not on PND2"
  StopScript
End if

MAXIS_row = 7

Do
  EMWriteScreen "e", MAXIS_row, 3
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  EMWriteScreen "fs", 20, 71
  EMSendKey "<enter>"
  EMWaitReady 0, 0  

  EMWriteScreen "fssm", 19, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  EMWriteScreen "app", 19, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  EMReadScreen expedited_check, 9, 16, 36
  If expedited_check = "EXPEDITED" then
    EMWriteScreen "n", 15, 60
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  EMWriteScreen "Y", 16, 51
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen PND2_check, 4, 2, 52
    If PND2_check = "LF) " then
      MsgBox "error"
      stopscript
    End if
  Loop until PND2_check = "PND2"
  
  MAXIS_row = MAXIS_row + 1
  payday_array = ""
Loop until MAXIS_row = amt_of_times_to_run + 7