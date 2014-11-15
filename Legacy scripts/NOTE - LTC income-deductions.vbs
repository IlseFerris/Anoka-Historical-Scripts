EMConnect ""

BeginDialog income_dialog, 0, 0, 381, 257, "Income Dialog"
  Text 5, 5, 25, 10, "Income"
  Text 175, 5, 125, 10, "Amt (leave blank if not provided)"
  EditBox 5, 15, 160, 15, income_01_name
  EditBox 175, 15, 80, 15, income_01_amt
  CheckBox 265, 15, 65, 15, "Proof needed?", income_01_proof_needed
  CheckBox 330, 15, 50, 15, "Excluded?", income_01_excluded
  EditBox 5, 35, 160, 15, income_02_name
  EditBox 175, 35, 80, 15, income_02_amt
  CheckBox 265, 35, 65, 15, "Proof needed?", income_02_proof_needed
  CheckBox 330, 35, 50, 15, "Excluded?", income_02_excluded
  EditBox 5, 55, 160, 15, income_03_name
  EditBox 175, 55, 80, 15, income_03_amt
  CheckBox 265, 55, 65, 15, "Proof needed?", income_03_proof_needed
  CheckBox 330, 55, 50, 15, "Excluded?", income_03_excluded
  EditBox 5, 75, 160, 15, income_04_name
  EditBox 175, 75, 80, 15, income_04_amt
  CheckBox 265, 75, 65, 15, "Proof needed?", income_04_proof_needed
  CheckBox 330, 75, 50, 15, "Excluded?", income_04_excluded
  EditBox 5, 95, 160, 15, income_05_name
  EditBox 175, 95, 80, 15, income_05_amt
  CheckBox 265, 95, 65, 15, "Proof needed?", income_05_proof_needed
  CheckBox 330, 95, 50, 15, "Excluded?", income_05_excluded
  EditBox 5, 115, 160, 15, income_06_name
  EditBox 175, 115, 80, 15, income_06_amt
  CheckBox 265, 115, 65, 15, "Proof needed?", income_06_proof_needed
  CheckBox 330, 115, 50, 15, "Excluded?", income_06_excluded
  EditBox 5, 135, 160, 15, income_07_name
  EditBox 175, 135, 80, 15, income_07_amt
  CheckBox 265, 135, 65, 15, "Proof needed?", income_07_proof_needed
  CheckBox 330, 135, 50, 15, "Excluded?", income_07_excluded
  EditBox 5, 155, 160, 15, income_08_name
  EditBox 175, 155, 80, 15, income_08_amt
  CheckBox 265, 155, 65, 15, "Proof needed?", income_08_proof_needed
  CheckBox 330, 155, 50, 15, "Excluded?", income_08_excluded
  EditBox 5, 175, 160, 15, income_09_name
  EditBox 175, 175, 80, 15, income_09_amt
  CheckBox 265, 175, 65, 15, "Proof needed?", income_09_proof_needed
  CheckBox 330, 175, 50, 15, "Excluded?", income_09_excluded
  EditBox 5, 195, 160, 15, income_10_name
  EditBox 175, 195, 80, 15, income_10_amt
  CheckBox 265, 195, 65, 15, "Proof needed?", income_10_proof_needed
  CheckBox 330, 195, 50, 15, "Excluded?", income_10_excluded
  Text 5, 220, 310, 15, "Note: if you have more than 10 incomes, you will have to manually add them to the case note."
  ButtonGroup ButtonPressed
    OkButton 250, 235, 50, 15
    CancelButton 305, 235, 50, 15
EndDialog

'Dialog income_dialog
'If ButtonPressed = 0 then stopscript


BeginDialog deduction_dialog, 0, 0, 331, 277, "Deduction Dialog"
  Text 5, 5, 35, 10, "Deduction"
  Text 175, 5, 125, 10, "Amt (leave blank if not provided)"
  EditBox 5, 15, 160, 15, deduction_01_name
  EditBox 175, 15, 80, 15, deduction_01_amt
  CheckBox 265, 15, 65, 15, "Proof needed?", deduction_01_proof_needed
  EditBox 5, 35, 160, 15, deduction_02_name
  EditBox 175, 35, 80, 15, deduction_02_amt
  CheckBox 265, 35, 65, 15, "Proof needed?", deduction_02_proof_needed
  EditBox 5, 55, 160, 15, deduction_03_name
  EditBox 175, 55, 80, 15, deduction_03_amt
  CheckBox 265, 55, 65, 15, "Proof needed?", deduction_03_proof_needed
  EditBox 5, 75, 160, 15, deduction_04_name
  EditBox 175, 75, 80, 15, deduction_04_amt
  CheckBox 265, 75, 65, 15, "Proof needed?", deduction_04_proof_needed
  EditBox 5, 95, 160, 15, deduction_05_name
  EditBox 175, 95, 80, 15, deduction_05_amt
  CheckBox 265, 95, 65, 15, "Proof needed?", deduction_05_proof_needed
  EditBox 5, 115, 160, 15, deduction_06_name
  EditBox 175, 115, 80, 15, deduction_06_amt
  CheckBox 265, 115, 65, 15, "Proof needed?", deduction_06_proof_needed
  EditBox 5, 135, 160, 15, deduction_07_name
  EditBox 175, 135, 80, 15, deduction_07_amt
  CheckBox 265, 135, 65, 15, "Proof needed?", deduction_07_proof_needed
  EditBox 5, 155, 160, 15, deduction_08_name
  EditBox 175, 155, 80, 15, deduction_08_amt
  CheckBox 265, 155, 65, 15, "Proof needed?", deduction_08_proof_needed
  EditBox 5, 175, 160, 15, deduction_09_name
  EditBox 175, 175, 80, 15, deduction_09_amt
  CheckBox 265, 175, 65, 15, "Proof needed?", deduction_09_proof_needed
  EditBox 5, 195, 160, 15, deduction_10_name
  EditBox 175, 195, 80, 15, deduction_10_amt
  CheckBox 265, 195, 65, 15, "Proof needed?", deduction_10_proof_needed
  Text 25, 220, 75, 10, "Sign your case note:"
  EditBox 105, 215, 90, 15, worker_sig
  Text 5, 240, 290, 10, "Note: if you have more than 10 deductions, you will have to case note them manually."
  ButtonGroup ButtonPressed
    PushButton 5, 260, 45, 10, "prev. page", prev_page
    OkButton 215, 255, 50, 15
    CancelButton 270, 255, 50, 15
EndDialog


Sub dialog_sub
  Do
    Do
      Dialog income_dialog
      If ButtonPressed = 0 then stopscript
    Loop until ButtonPressed = -1
    Do
      Dialog deduction_dialog
      If ButtonPressed = 0 then stopscript
    Loop until ButtonPressed = -1 or ButtonPressed = 79
    If ButtonPressed = 79 then exit do
  Loop until ButtonPressed = -1
End Sub


'This will force the dialog_sub to restart if the last button pressed isn't "OK"

Do
  dialog_sub
Loop until ButtonPressed = -1

'If ButtonPressed = 0 then stopscript


Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then call dialog_sub
If ButtonPressed = 0 then stopscript
End Sub

Do
  find_case_note
Loop until case_note_ready = "Case Notes (NOTE)" and case_note_mode = "Mode: A" or case_note_mode = "Mode: E"

EMSendKey "<enter>"
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case note. Type your password then try again."
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog deduction_dialog
     If ButtonPressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

EMSendKey "<home>" + "<--LTC income/deductions-->" + "<newline>"
EMSendKey "------------income--------------------------------------------amt------" + "<newline>"
EMSetCursor 17, 3
EMSendKey " >>>>DEDUCTIONS>>>>"
EMSetCursor 6, 3

Do
EMSendKey ".                                                                             "
EMReadScreen space_fill_check, 1, 16, 3
Loop until space_fill_check <> " "








If income_01_excluded = 1 then income_01_amt = income_01_amt & " (excl)"
If income_01_proof_needed = 1 then income_01_amt = "Needs verif"

If income_02_excluded = 1 then income_02_amt = income_02_amt & " (excl)"
If income_02_proof_needed = 1 then income_02_amt = "Needs verif"

If income_03_excluded = 1 then income_03_amt = income_03_amt & " (excl)"
If income_03_proof_needed = 1 then income_03_amt = "Needs verif"

If income_04_excluded = 1 then income_04_amt = income_04_amt & " (excl)"
If income_04_proof_needed = 1 then income_04_amt = "Needs verif"

If income_05_excluded = 1 then income_05_amt = income_05_amt & " (excl)"
If income_05_proof_needed = 1 then income_05_amt = "Needs verif"

If income_06_excluded = 1 then income_06_amt = income_06_amt & " (excl)"
If income_06_proof_needed = 1 then income_06_amt = "Needs verif"

If income_07_excluded = 1 then income_07_amt = income_07_amt & " (excl)"
If income_07_proof_needed = 1 then income_07_amt = "Needs verif"

If income_08_excluded = 1 then income_08_amt = income_08_amt & " (excl)"
If income_08_proof_needed = 1 then income_08_amt = "Needs verif"

If income_09_excluded = 1 then income_09_amt = income_09_amt & " (excl)"
If income_09_proof_needed = 1 then income_09_amt = "Needs verif"

If income_10_excluded = 1 then income_10_amt = income_10_amt & " (excl)"
If income_10_proof_needed = 1 then income_10_amt = "Needs verif"


EMSetCursor 6, 3
If income_01_name <> "" then EMSendKey "    " + income_01_name
EMSetCursor 7, 3
If income_02_name <> "" then EMSendKey "    " + income_02_name
EMSetCursor 8, 3
If income_03_name <> "" then EMSendKey "    " + income_03_name
EMSetCursor 9, 3
If income_04_name <> "" then EMSendKey "    " + income_04_name
EMSetCursor 10, 3
If income_05_name <> "" then EMSendKey "    " + income_05_name
EMSetCursor 11, 3
If income_06_name <> "" then EMSendKey "    " + income_06_name
EMSetCursor 12, 3
If income_07_name <> "" then EMSendKey "    " + income_07_name
EMSetCursor 13, 3
If income_08_name <> "" then EMSendKey "    " + income_08_name
EMSetCursor 14, 3
If income_09_name <> "" then EMSendKey "    " + income_09_name
EMSetCursor 15, 3
If income_10_name <> "" then EMSendKey "    " + income_10_name





EMSetCursor 6, 63
EMSendKey income_01_amt
EMSetCursor 7, 63
EMSendKey income_02_amt
EMSetCursor 8, 63
EMSendKey income_03_amt
EMSetCursor 9, 63
EMSendKey income_04_amt
EMSetCursor 10, 63
EMSendKey income_05_amt
EMSetCursor 11, 63
EMSendKey income_06_amt
EMSetCursor 12, 63
EMSendKey income_07_amt
EMSetCursor 13, 63
EMSendKey income_08_amt
EMSetCursor 14, 63
EMSendKey income_09_amt
EMSetCursor 15, 63
EMSendKey income_10_amt

EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 1


EMSendKey "------------deductions----------------------------------------amt------" + "<newline>"
EMSetCursor 17, 3
EMSendKey "x"
EMSetCursor 6, 3

Do
EMSendKey "                                                                              "
EMReadScreen space_fill_check, 1, 17, 3
Loop until space_fill_check = " "




If deduction_01_proof_needed = 1 then deduction_01_amt = "Needs verif"
If deduction_02_proof_needed = 1 then deduction_02_amt = "Needs verif"
If deduction_03_proof_needed = 1 then deduction_03_amt = "Needs verif"
If deduction_04_proof_needed = 1 then deduction_04_amt = "Needs verif"
If deduction_05_proof_needed = 1 then deduction_05_amt = "Needs verif"
If deduction_06_proof_needed = 1 then deduction_06_amt = "Needs verif"
If deduction_07_proof_needed = 1 then deduction_07_amt = "Needs verif"
If deduction_08_proof_needed = 1 then deduction_08_amt = "Needs verif"
If deduction_09_proof_needed = 1 then deduction_09_amt = "Needs verif"
If deduction_10_proof_needed = 1 then deduction_10_amt = "Needs verif"



EMSetCursor 6, 7
EMSendKey deduction_01_name
EMSetCursor 7, 7
EMSendKey deduction_02_name
EMSetCursor 8, 7
EMSendKey deduction_03_name
EMSetCursor 9, 7
EMSendKey deduction_04_name
EMSetCursor 10, 7
EMSendKey deduction_05_name
EMSetCursor 11, 7
EMSendKey deduction_06_name
EMSetCursor 12, 7
EMSendKey deduction_07_name
EMSetCursor 13, 7
EMSendKey deduction_08_name
EMSetCursor 14, 7
EMSendKey deduction_09_name
EMSetCursor 15, 7
EMSendKey deduction_10_name



EMSetCursor 6, 63
EMSendKey deduction_01_amt
EMSetCursor 7, 63
EMSendKey deduction_02_amt
EMSetCursor 8, 63
EMSendKey deduction_03_amt
EMSetCursor 9, 63
EMSendKey deduction_04_amt
EMSetCursor 10, 63
EMSendKey deduction_05_amt
EMSetCursor 11, 63
EMSendKey deduction_06_amt
EMSetCursor 12, 63
EMSendKey deduction_07_amt
EMSetCursor 13, 63
EMSendKey deduction_08_amt
EMSetCursor 14, 63
EMSendKey deduction_09_amt
EMSetCursor 15, 63
EMSendKey deduction_10_amt




EMSendKey "<newline>" + "---" + "<newline>" + worker_sig

EMSendKey "<PF7>"
EMWaitReady 1, 1
EMWriteScreen " ", 17, 3
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSetCursor 5, 3
EMSendKey "x" + "<enter>"
