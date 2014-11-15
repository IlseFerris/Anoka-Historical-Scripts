amt_of_times_to_run = 12
PP_together = "N"
EATS_member_array = array("01", "03") 
EATS_non_member_array = array("24") 
EATS_action = "01" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.



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
  EMWaitReady 0, 0

  EMWriteScreen "EATS", 20, 71
  EMWriteScreen EATS_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  If EATS_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 0, 0
  End if

  EMWriteScreen PP_together, 4, 72
  EMWriteScreen "N", 5, 72

  If PP_together = "N" then
    EMWriteScreen "01", 13, 28
  
    col = 39
    For each EATS_member in EATS_member_array  
      EMWriteScreen EATS_member, 13, col
      col = col + 4
    Next
  
    col = 39
    EMWriteScreen "02", 14, 28
    For each EATS_non_member in EATS_non_member_array  
      EMWriteScreen EATS_non_member, 14, col
      col = col + 4
    Next
  End if

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