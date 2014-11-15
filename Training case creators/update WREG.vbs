amt_of_times_to_run = 16
WREG_action = "01" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Leave blank to ignore this panel.
WREG_member_array = array("01") 
FSET_status = "30"
defer_FSET_indicator = "N" 'Should usually be a "N" or a "_" for exempt people.
ABAWD_status = "09"
GA_basis = "99"

EATS_action = "" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Leave blank to ignore this panel.
PP_together = "N"
EATS_member_array = array("01") 
EATS_non_member_array = array("24") 



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

  If WREG_action <> "" then

    For each WREG_member in WREG_member_array  
      EMWriteScreen "WREG", 20, 71
      EMWriteScreen WREG_member, 20, 76
      EMWriteScreen WREG_action, 20, 79
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    
      If WREG_action <> "nn" then
        EMSendKey "<PF9>"
        EMWaitReady 0, 0
      End if
    
      If WREG_member = "01" then EMWriteScreen "Y", 6, 68
      If WREG_member <> "01" then EMWriteScreen "N", 6, 68
      EMWriteScreen FSET_status, 8, 50
      EMWriteScreen defer_FSET_indicator, 8, 80
      EMWriteScreen ABAWD_status, 13, 50
      EMWriteScreen GA_basis, 15, 50
  
      EMSendKey "<enter>"
      EMWaitReady 0, 0

      EMSendKey "<enter>"
      EMWaitReady 0, 0  
    Next
  End if
'----------------------------------------------------------------------------------------------------
  If EATS_action <> "" then

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

    EMSendKey "<enter>"
    EMWaitReady 0, 0  

  End if
'----------------------------------------------------------------------------------------------------

  
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