EMConnect ""

BeginDialog asset_calc_dialog, 5, 5, 381, 287, "Asset Calculation Dialog"
  EditBox 5, 15, 100, 15, asset_01_name
  EditBox 115, 15, 50, 15, asset_01_date
  EditBox 175, 15, 80, 15, asset_01_amt
  CheckBox 265, 15, 65, 15, "Proof needed?", asset_01_proof_needed
  CheckBox 330, 15, 50, 15, "Excluded?", asset_01_excluded
  EditBox 5, 35, 100, 15, asset_02_name
  EditBox 115, 35, 50, 15, asset_02_date
  EditBox 175, 35, 80, 15, asset_02_amt
  CheckBox 265, 35, 65, 15, "Proof needed?", asset_02_proof_needed
  CheckBox 330, 35, 50, 15, "Excluded?", asset_02_excluded
  EditBox 5, 55, 100, 15, asset_03_name
  EditBox 115, 55, 50, 15, asset_03_date
  EditBox 175, 55, 80, 15, asset_03_amt
  CheckBox 265, 55, 65, 15, "Proof needed?", asset_03_proof_needed
  CheckBox 330, 55, 50, 15, "Excluded?", asset_03_excluded
  EditBox 5, 75, 100, 15, asset_04_name
  EditBox 115, 75, 50, 15, asset_04_date
  EditBox 175, 75, 80, 15, asset_04_amt
  CheckBox 265, 75, 65, 15, "Proof needed?", asset_04_proof_needed
  CheckBox 330, 75, 50, 15, "Excluded?", asset_04_excluded
  EditBox 5, 95, 100, 15, asset_05_name
  EditBox 115, 95, 50, 15, asset_05_date
  EditBox 175, 95, 80, 15, asset_05_amt
  CheckBox 265, 95, 65, 15, "Proof needed?", asset_05_proof_needed
  CheckBox 330, 95, 50, 15, "Excluded?", asset_05_excluded
  EditBox 5, 115, 100, 15, asset_06_name
  EditBox 115, 115, 50, 15, asset_06_date
  EditBox 175, 115, 80, 15, asset_06_amt
  CheckBox 265, 115, 65, 15, "Proof needed?", asset_06_proof_needed
  CheckBox 330, 115, 50, 15, "Excluded?", asset_06_excluded
  EditBox 5, 135, 100, 15, asset_07_name
  EditBox 115, 135, 50, 15, asset_07_date
  EditBox 175, 135, 80, 15, asset_07_amt
  CheckBox 265, 135, 65, 15, "Proof needed?", asset_07_proof_needed
  CheckBox 330, 135, 50, 15, "Excluded?", asset_07_excluded
  EditBox 5, 155, 100, 15, asset_08_name
  EditBox 115, 155, 50, 15, asset_08_date
  EditBox 175, 155, 80, 15, asset_08_amt
  CheckBox 265, 155, 65, 15, "Proof needed?", asset_08_proof_needed
  CheckBox 330, 155, 50, 15, "Excluded?", asset_08_excluded
  EditBox 5, 175, 100, 15, asset_09_name
  EditBox 115, 175, 50, 15, asset_09_date
  EditBox 175, 175, 80, 15, asset_09_amt
  CheckBox 265, 175, 65, 15, "Proof needed?", asset_09_proof_needed
  CheckBox 330, 175, 50, 15, "Excluded?", asset_09_excluded
  EditBox 5, 195, 100, 15, asset_10_name
  EditBox 115, 195, 50, 15, asset_10_date
  EditBox 175, 195, 80, 15, asset_10_amt
  CheckBox 265, 195, 65, 15, "Proof needed?", asset_10_proof_needed
  CheckBox 330, 195, 50, 15, "Excluded?", asset_10_excluded
  EditBox 5, 215, 100, 15, asset_11_name
  EditBox 115, 215, 50, 15, asset_11_date
  EditBox 175, 215, 80, 15, asset_11_amt
  CheckBox 265, 215, 65, 15, "Proof needed?", asset_11_proof_needed
  CheckBox 330, 215, 50, 15, "Excluded?", asset_11_excluded
  ButtonGroup asset_calc_ButtonPressed
    OkButton 240, 265, 50, 15
    CancelButton 295, 265, 50, 15
  Text 5, 240, 290, 20, "Note: if you have more than 11 assets, you will have to calculate manually at this time. The script will calculate and allow you to insert cost of care on the next page."
  Text 175, 5, 125, 10, "Amt (leave blank if not provided)"
  Text 5, 5, 25, 10, "Asset"
  Text 115, 5, 55, 10, "Statement date"
EndDialog

Dialog asset_calc_dialog
If asset_calc_ButtonPressed = 0 then stopscript



BeginDialog cost_of_care_dialog, 0, 0, 236, 87, "Cost of Care"
  EditBox 65, 5, 100, 15, cost_of_care
  CheckBox 10, 50, 185, 10, "Check here if the cost of care is unknown at this time.", cost_of_care_check
  EditBox 85, 65, 105, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 175, 10, 50, 15
    CancelButton 175, 30, 50, 15
  Text 10, 25, 160, 20, "Note: the script will auto subtract the cost of care from the asset total."
  Text 10, 10, 50, 10, "Cost of care:"
  Text 10, 70, 75, 10, "Sign your case note:"
EndDialog

Dialog cost_of_care_dialog
If ButtonPressed = 0 then stopscript

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog cost_of_care_dialog
If buttonpressed = 0 then stopscript
End Sub

Do
  find_case_note
Loop until case_note_ready = "Case Notes (NOTE)" and case_note_mode = "Mode: A" or case_note_mode = "Mode: E"

EMSendKey "<enter>"
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case note. Type your password then try again."
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog cost_of_care_dialog
     IF buttonpressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

EMSendKey "<home>" + "<--LTC asset calculation-->" + "<newline>"
EMSendKey "------------asset-----------------------statement date--------amt------" + "<newline>"
EMSetCursor 17, 3
EMSendKey "x"
EMSetCursor 6, 3

Do
EMSendKey "                                                                              "
EMReadScreen space_fill_check, 1, 17, 3
Loop until space_fill_check = " "







If asset_01_proof_needed = 1 then asset_01_amt = "Needs verif"
If asset_01_excluded = 1 then asset_01_amt = "Excluded"

If asset_02_proof_needed = 1 then asset_02_amt = "Needs verif"
If asset_02_excluded = 1 then asset_02_amt = "Excluded"

If asset_03_proof_needed = 1 then asset_03_amt = "Needs verif"
If asset_03_excluded = 1 then asset_03_amt = "Excluded"

If asset_04_proof_needed = 1 then asset_04_amt = "Needs verif"
If asset_04_excluded = 1 then asset_04_amt = "Excluded"

If asset_05_proof_needed = 1 then asset_05_amt = "Needs verif"
If asset_05_excluded = 1 then asset_05_amt = "Excluded"

If asset_06_proof_needed = 1 then asset_06_amt = "Needs verif"
If asset_06_excluded = 1 then asset_06_amt = "Excluded"

If asset_07_proof_needed = 1 then asset_07_amt = "Needs verif"
If asset_07_excluded = 1 then asset_07_amt = "Excluded"

If asset_08_proof_needed = 1 then asset_08_amt = "Needs verif"
If asset_08_excluded = 1 then asset_08_amt = "Excluded"

If asset_09_proof_needed = 1 then asset_09_amt = "Needs verif"
If asset_09_excluded = 1 then asset_09_amt = "Excluded"

If asset_10_proof_needed = 1 then asset_10_amt = "Needs verif"
If asset_10_excluded = 1 then asset_10_amt = "Excluded"

If asset_11_proof_needed = 1 then asset_11_amt = "Needs verif"
If asset_11_excluded = 1 then asset_11_amt = "Excluded"

EMSetCursor 6, 7
EMSendKey asset_01_name
EMSetCursor 7, 7
EMSendKey asset_02_name
EMSetCursor 8, 7
EMSendKey asset_03_name
EMSetCursor 9, 7
EMSendKey asset_04_name
EMSetCursor 10, 7
EMSendKey asset_05_name
EMSetCursor 11, 7
EMSendKey asset_06_name
EMSetCursor 12, 7
EMSendKey asset_07_name
EMSetCursor 13, 7
EMSendKey asset_08_name
EMSetCursor 14, 7
EMSendKey asset_09_name
EMSetCursor 15, 7
EMSendKey asset_10_name
EMSetCursor 16, 7
EMSendKey asset_11_name

EMSetCursor 6, 45
EMSendKey asset_01_date
EMSetCursor 7, 45
EMSendKey asset_02_date
EMSetCursor 8, 45
EMSendKey asset_03_date
EMSetCursor 9, 45
EMSendKey asset_04_date
EMSetCursor 10, 45
EMSendKey asset_05_date
EMSetCursor 11, 45
EMSendKey asset_06_date
EMSetCursor 12, 45
EMSendKey asset_07_date
EMSetCursor 13, 45
EMSendKey asset_08_date
EMSetCursor 14, 45
EMSendKey asset_09_date
EMSetCursor 15, 45
EMSendKey asset_10_date
EMSetCursor 16, 45
EMSendKey asset_11_date

EMSetCursor 6, 63
EMSendKey asset_01_amt
EMSetCursor 7, 63
EMSendKey asset_02_amt
EMSetCursor 8, 63
EMSendKey asset_03_amt
EMSetCursor 9, 63
EMSendKey asset_04_amt
EMSetCursor 10, 63
EMSendKey asset_05_amt
EMSetCursor 11, 63
EMSendKey asset_06_amt
EMSetCursor 12, 63
EMSendKey asset_07_amt
EMSetCursor 13, 63
EMSendKey asset_08_amt
EMSetCursor 14, 63
EMSendKey asset_09_amt
EMSetCursor 15, 63
EMSendKey asset_10_amt
EMSetCursor 16, 63
EMSendKey asset_11_amt




Dim asset_total

'It sets the variables up to be numeric so that the calculation works
If IsNumeric(asset_01_amt) = "False" then asset_01_amt = 0
If IsNumeric(asset_02_amt) = "False" then asset_02_amt = 0
If IsNumeric(asset_03_amt) = "False" then asset_03_amt = 0
If IsNumeric(asset_04_amt) = "False" then asset_04_amt = 0
If IsNumeric(asset_05_amt) = "False" then asset_05_amt = 0
If IsNumeric(asset_06_amt) = "False" then asset_06_amt = 0
If IsNumeric(asset_07_amt) = "False" then asset_07_amt = 0
If IsNumeric(asset_08_amt) = "False" then asset_08_amt = 0
If IsNumeric(asset_09_amt) = "False" then asset_09_amt = 0
If IsNumeric(asset_10_amt) = "False" then asset_10_amt = 0
If IsNumeric(asset_11_amt) = "False" then asset_11_amt = 0


'Now it calculates the total
asset_total = (0 + asset_01_amt + asset_02_amt + asset_03_amt + asset_04_amt + asset_05_amt + asset_06_amt + asset_07_amt + asset_08_amt + asset_09_amt + asset_10_amt + asset_11_amt)

'Now it gets ready to display the total
EMSetCursor 17, 60
EMSendKey "------------" + "<PF8>"
EMWaitReady 1, 0
EMSetCursor 4, 3
EMSendKey "                                                TOTAL:      " & asset_total
EMSendKey "<newline>"
EMSendKey "                                        COST FOR CARE:      " 

'Now it checks cost for care. If the checkbox was checked, it reads back "unknown". If it isn't checked and the value is numeric, it prints it here.
If IsNumeric(cost_of_care) = False then cost_of_care = "Unknown"
If cost_of_care_check = 1 then cost_of_care = "Unknown"

EMSendKey cost_of_care
EMSendKey "<newline>"
EMSendKey "                                                         ------------" + "<newline>"
EMSendKey "                                 TOTAL COUNTED ASSETS:      "
If cost_of_care <> "Unknown" then EMSendKey (asset_total - cost_of_care)
If cost_of_care = "Unknown" then EMSendKey "Unknown"
EMSendKey "<newline>" + "---" + "<newline>" + worker_sig + "<PF3>"
EMWaitReady 1, 0
EMSetCursor 5, 3
EMSendKey "x" + "<enter>"
