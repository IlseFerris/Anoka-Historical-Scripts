function transmit
  EMSendKey "<enter>"
  EMWaitReady 1, 1
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
End function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 1, 1
End function

function back_to_self
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

EMConnect ""

row = 1
col = 1
EMSearch "Case Nbr:", row, col
If row <> 0 then EMReadScreen case_number, 8, row, col + 10

BeginDialog case_number_dialog, 0, 0, 161, 41, "Case number"
  EditBox 95, 0, 60, 15, case_number
  ButtonGroup ButtonPressed
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

Dialog case_number_dialog
If ButtonPressed = 0 then stopscript
PF3
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  MsgBox "You are not in MAXIS, or you are locked out of your case."
  stopscript
End if


back_to_self

EMWriteScreen "elig", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "07", 20, 43
EMWriteScreen "12", 20, 46
EMWriteScreen "hc", 21, 70
transmit

EMReadScreen person_check, 2, 8, 31
If person_check = "NO" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 29
End if
If person_check <> "NO" then EMWriteScreen "x", 8, 29
transmit

row = 1
col = 1
EMSearch "07/12", row, col
If row = 0 then 
  MsgBox "A 07/12 span could not be found. Try this again. You may need to run the case through background."
  stopscript
End if


EMReadScreen elig_type, 2, 12, col - 2
EMReadScreen budget_type, 1, 13, col + 2
EMWriteScreen "x", 9, col + 2
transmit

EMReadScreen LBUD_check, 4, 3, 45
If LBUD_check = "LBUD" then
  EMReadScreen recipient_amt, 10, 15, 70
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 14, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 15, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 16, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Personal needs ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen home_maintenance_allowance, 10, 17, 32
  If home_maintenance_allowance <> "__________" then deductions = deductions & "Home maintenance allowance ($" & replace(home_maintenance_allowance, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 8, 70
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 9, 70
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 10, 70
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 11, 70
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 12, 70
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 13, 70
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

EMReadScreen SBUD_check, 4, 3, 44
If SBUD_check = "SBUD" then
  EMReadScreen recipient_amt, 10, 16, 71
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 13, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 15, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 16, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 17, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Maintenance needs allowance ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 9, 71
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 10, 71
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 11, 71
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 12, 71
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 13, 71
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 14, 71
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

BeginDialog BBUD_Dialog, 0, 0, 191, 76, "BBUD"
  Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
    PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
    CancelButton 135, 55, 50, 15
EndDialog

EMReadScreen BBUD_check, 4, 3, 47
If BBUD_check = "BBUD" then
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  Dialog BBUD_dialog
  If ButtonPressed = 0 then stopscript
  If ButtonPressed = 4 then
    PF3
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then
      Do
        Dialog BBUD_Dialog
        If buttonpressed = 0 then stopscript
      Loop until MAXIS_check = "MAXIS"
    End if
    back_to_SELF
    EMWriteScreen "stat", 16, 43
    EMWriteScreen "bils", 21, 70
    transmit
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
  End if
End if

BeginDialog COLA_dialog, 5, 5, 376, 121, "COLA"
  Text 5, 10, 35, 10, "Elig type:"
  DropListBox 45, 5, 30, 15, "EX"+chr(9)+"DX", elig_type
  Text 85, 10, 45, 10, "Budget type:"
  DropListBox 135, 5, 30, 15, "L"+chr(9)+"S"+chr(9)+"B", budget_type
  Text 175, 10, 110, 10, "Waiver obilgation/recipient amt:"
  EditBox 285, 5, 85, 15, recipient_amt
  Text 5, 30, 80, 10, "Total countable income:"
  EditBox 90, 25, 280, 15, income
  Text 5, 50, 45, 10, "Deductions:"
  EditBox 50, 45, 320, 15, deductions
  CheckBox 5, 65, 65, 10, "Updated RSPL?", updated_RSPL_check
  CheckBox 85, 65, 110, 10, "Approved new MAXIS results?", approved_check
  CheckBox 210, 65, 70, 10, "Sent DHS-3050?", DHS_3050_check
  CheckBox 290, 65, 85, 10, "Designated provider?", designated_provider_check
  Text 5, 85, 70, 10, "Other (if applicable):"
  EditBox 75, 80, 295, 15, other
  Text 5, 105, 70, 10, "Sign your case note:"
  EditBox 80, 100, 70, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 165, 100, 50, 15
    CancelButton 225, 100, 50, 15
EndDialog




Dialog COLA_dialog
If buttonpressed = 0 then stopscript
PF3 'checking for password prompt
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  Do
    Dialog COLA_dialog
    If buttonpressed = 0 then stopscript
  Loop until MAXIS_check = "MAXIS"
End if
back_to_self

EMWriteScreen "case", 16, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "note", 21, 70
transmit

PF9

EMSendKey "<home>"
EMSendKey "**Approved COLA updates 07/12: " & elig_type & "-" & budget_type & " " & recipient_amt
If budget_type = "L" then EMSendKey " LTC sd**"
If budget_type = "S" then EMSendKey " SISEW waiver obl**"
If budget_type = "B" then EMSendKey " recip amt**"
EMSendKey "<newline>"
EMSendKey "* Income: " & income & "<newline>"
EMSendKey "* Deductions: " & deductions & "<newline>"
EMSendKey "---" & "<newline>"
If updated_RSPL_check = 1 then EMSendKey "* Updated RSPL in MMIS." & "<newline>"
If designated_provider_check = 1 then EMSendKey "* Client has designated provider." & "<newline>"
If approved_check = 1 then EMSendKey "* Approved new MAXIS results." & "<newline>"
If DHS_3050_check = 1 then EMSendKey "* Sent DHS-3050 LTC communication form to facility." & "<newline>"
If other <> "" then EMSendKey "* Other: " & other & "<newline>"
EMSendKey "---" & "<newline>"
EMSendKey worker_sig
