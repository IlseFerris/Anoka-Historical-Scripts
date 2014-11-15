BeginDialog case_number_dialog, 0, 0, 161, 42, "Case number"
  EditBox 95, 0, 60, 15, case_number
  ButtonGroup case_number_dialog_ButtonPressed
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

'The following is an experimental way to find a MAXIS screen. 
'It only connects to the MAXIS screen, and could work from a third party script host potentially.
Do
  Dialog case_number_dialog
  If case_number_dialog_ButtonPressed = 0 then stopscript
  EMConnect "A"
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" then
    EMConnect "B"
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then 
      EMConnect "C"
      EMSendKey "<enter>"
      EMWaitReady 1, 1
      EMReadScreen MAXIS_check, 5, 1, 39
      If MAXIS_check <> "MAXIS" then MsgBox "Neither screen appears to be on MAXIS. Get one of your screens to MAXIS before proceeding. You may need to enter a password. If you are in MAXIS, you may have had a configuration error. To fix this, restart BlueZone."
    End If
  End If
Loop until MAXIS_check = "MAXIS"
EMFocus

'This jumps back to SELF
Do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 4, 2, 50
  If SELF_check = "SELF" then exit do
Loop until SELF_check = "SELF"

EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "disq", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

EMReadScreen DISQ_member_check, 34, 24, 2
If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then 
  EMSendKey "<PF3>"
  MsgBox "No IEVS DISQ indicated for this case."
End if