amt_of_times_to_run = 6
TRAC_member = "01"
TRAC_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
TRAC_month_01 = "10/10"
TRAC_month_02 = "11/10"
TRAC_month_03 = "12/10"
TRAC_month_04 = "01/11"


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
  
  EMWriteScreen "TRAC", 20, 71
  EMWriteScreen TRAC_member, 20, 76
  EMWriteScreen TRAC_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  If TRAC_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
  End if

  EMWriteScreen "Y", 6, 76
  EMWriteScreen left(TRAC_month_01, 2), 10, 36
  EMWriteScreen right(TRAC_month_01, 2), 10, 41
  EMWriteScreen left(TRAC_month_02, 2), 11, 36
  EMWriteScreen right(TRAC_month_02, 2), 11, 41
  EMWriteScreen left(TRAC_month_03, 2), 12, 36
  EMWriteScreen right(TRAC_month_03, 2), 12, 41
  EMWriteScreen left(TRAC_month_04, 2), 13, 36
  EMWriteScreen right(TRAC_month_04, 2), 13, 41
  
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