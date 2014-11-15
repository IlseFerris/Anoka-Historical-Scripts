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
  EMReadScreen MMIS_case_note_check, 5, 5, 2
  If MMIS_case_note_check <> "'''''" then
    EMConnect "B"
    EMReadScreen MMIS_case_note_check, 5, 5, 2
    If MMIS_case_note_check <> "'''''" then 
      EMConnect "C"
      EMReadScreen MMIS_case_note_check, 5, 5, 2
      If MMIS_case_note_check <> "'''''" then MsgBox "Neither screen appears to be on MAXIS. Get one of your screens to MAXIS before proceeding. You may need to enter a password. If you are in MAXIS, you may have had a configuration error. To fix this, restart BlueZone."
    End If
  End If
Loop until MMIS_case_note_check = "'''''"
EMFocus

EMSendKey case_number