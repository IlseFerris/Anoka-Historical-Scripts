amt_of_times_to_run = 6
REST_member = "02"
REST_type = "1" '1 for House, 2 for Land, 3 for Buildings, 4 for Mobile Home, 5 for Life Estate, 6 for Other
REST_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.

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
  
  EMWriteScreen "REST", 20, 71
  EMWriteScreen REST_member, 20, 76
  EMWriteScreen REST_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  
  If REST_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
  End if

  EMWriteScreen REST_type, 6, 39
  EMWriteScreen "OT", 6, 62
  EMWriteScreen "1", 12, 54
  EMWriteScreen "N", 13, 54

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
  payday_array = ""
Loop until MAXIS_row = amt_of_times_to_run + 7