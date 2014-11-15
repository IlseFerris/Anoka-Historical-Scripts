amt_of_times_to_run = 8
siblings = array("03", "04") 

EMConnect ""

EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then
  MsgBox "Not on PND2"
  StopScript
End if

MAXIS_row = 7

Do

  EMWriteScreen "s", MAXIS_row, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  EMWriteScreen "SIBL", 20, 71
  EMWriteScreen "nn", 20, 79
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  col = 39

  EMWriteScreen "01", 7, 28

  For each kid in siblings
    EMWriteScreen kid, 7, col
    col = col + 4
  Next
  
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen PND2_check, 4, 2, 52
    If PND2_check = "LF) " then
      MsgBox "error"
      stopscript
    End if
  Loop until PND2_check = "PND2"
  
  MAXIS_row = MAXIS_row + 1
Loop until MAXIS_row = amt_of_times_to_run + 7