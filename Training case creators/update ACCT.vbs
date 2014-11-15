amt_of_times_to_run = 16
ACCT_member = "24"
ACCT_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
ACCT_type = "SV"
ACCT_number = "99999999-01"
ACCT_location = "Some other bank"
ACCT_balance = "3000"
ACCT_as_of_month = "01"
ACCT_as_of_day = "01"
ACCT_as_of_year = "13"
cash_count_status = "Y"
SNAP_count_status = "Y"
HC_count_status = ""


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
  
  EMWriteScreen "ACCT", 20, 71
  EMWriteScreen ACCT_member, 20, 76
  EMWriteScreen ACCT_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  If ACCT_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 0, 0
  End if

  'Clears out existing info
  EMSendKey string(79, "_")

  EMWriteScreen ACCT_type, 6, 44
  EMWriteScreen ACCT_number, 7, 44
  EMWriteScreen ACCT_location, 8, 44
  EMWriteScreen ACCT_balance, 10, 46
  EMWriteScreen "5", 10, 63
  EMWriteScreen ACCT_as_of_month, 11, 44
  EMWriteScreen ACCT_as_of_day, 11, 47
  EMWriteScreen ACCT_as_of_year, 11, 50

  EMWriteScreen cash_count_status, 14, 50
  EMWriteScreen SNAP_count_status, 14, 57
  EMWriteScreen HC_count_status, 14, 64
  EMWriteScreen "N", 15, 44
  
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
  cases_run = cases_run + 1
  If MAXIS_row = 19 then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    MAXIS_row = 7
  End if

Loop until cases_run = amt_of_times_to_run