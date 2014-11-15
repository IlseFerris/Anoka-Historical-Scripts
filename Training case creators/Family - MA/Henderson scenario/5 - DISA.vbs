amt_of_times_to_run = 8
DISA_member = "03"
DISA_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
DISA_type = "03"
DISA_start_date_month = "01"
DISA_start_date_day = "01"
DISA_start_date_year = "2012"
cash_status = ""
SNAP_status = ""
HC_status = "03"


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
  
  EMWriteScreen "DISA", 20, 71
  EMWriteScreen DISA_member, 20, 76
  EMWriteScreen DISA_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  If DISA_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
  End if

  EMWriteScreen DISA_start_date_month, 6, 47
  EMWriteScreen DISA_start_date_day, 6, 50
  EMWriteScreen DISA_start_date_year, 6, 53

  If cash_status <> "" then
    EMWriteScreen cash_status, 11, 59
    EMWriteScreen "3", 11, 69
  End if

  If SNAP_status <> "" then
    EMWriteScreen SNAP_status, 12, 59
    EMWriteScreen "3", 12, 69
  End if

  If HC_status <> "" then
    EMWriteScreen HC_status, 13, 59
    EMWriteScreen "3", 13, 69
  End if
  
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