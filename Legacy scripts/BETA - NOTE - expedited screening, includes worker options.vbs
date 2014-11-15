'-----------------------------------------------------------------------------------------------------------------------------------------------------
'        Script name:    CAF1 screening
'        Description:    Allows a worker to enter basic information about a CAF received to determine if it appears expedited.
'       Target users:    OSA and FAS.
'           Division:    Family
'          Author(s):    Ronny Cary
'      Working state:    BETA
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'     Script content:    01. Tries to find a case number, then pops up a dialog message for a worker to enter income, assets, 
'                            rent/mortgage, and utilities from CAF1. It also has a checkbox to start the appointment letter.
'                        02. Script calculates the amounts. If the income + assets are less than the expenses, or the 
'                            income is under $150 AND the assets are under $100, then the client is expedited. The
'                            script shows a messagebox indicating this status.
'                        03. The script checks for IEVS disqualifications for the case.
'                        04. The script navigates to CASE/NOTE and notes the expedited status and the calculation. It then shows a 
'                            messagebox for the worker who wrote the case note, explaining what the next action should be per 
'                            procedure.
'                        05. The script loads up the appointment letter script, if the worker requests it in step 01.
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'       Known issues:    1. Script should check CASE/CURR for current/previous SNAP status. This may affect expedited eligibility. In order to see
'                           this, I will need example cases.
'                        2. Procedures need to be clarified for the messageboxes.
'                        3. Supervisors need to approve the script before it is released.
'   Test breakpoints:    None
'              Notes:    This script should be functionally identical in both Adult and Family. The only difference is that in 
'                        Family, the script needs to load the Family appointment letter, and in Adult it needs to pull up the
'                        Adult appointment letter.
'-----------------------------------------------------------------------------------------------------------------------------------------------------

'SECTION 01
EMConnect ""

BeginDialog exp_screening_dialog, 0, 0, 251, 197, "Expedited Screening Dialog"
  EditBox 55, 5, 80, 15, case_number
  EditBox 100, 25, 65, 15, income
  EditBox 100, 45, 60, 15, assets
  EditBox 115, 65, 50, 15, rent
  EditBox 165, 85, 40, 15, utilities
  CheckBox 15, 105, 55, 10, "Heat (or AC)", heat_AC_check
  CheckBox 75, 105, 45, 10, "Electricity", electric_check
  CheckBox 130, 105, 35, 10, "Phone", phone_check
  CheckBox 5, 125, 180, 10, "APPLed case for another worker? If so, write worker:", APPLed_case
  EditBox 190, 120, 55, 15, worker_name
  CheckBox 25, 140, 65, 10, "Paper Process?", paper_process_check
  EditBox 80, 155, 80, 15, worker_sig
  CheckBox 5, 180, 200, 10, "Check here to start the appointment letter script after this.", appointment_letter_check
  ButtonGroup exp_screening_dialog_ButtonPressed
    OkButton 190, 5, 50, 15
    CancelButton 190, 25, 50, 15
  Text 5, 50, 95, 10, "Cash, checking, or savings: "
  Text 5, 70, 105, 10, "Amounts paid for rent/mortgage:"
  Text 5, 90, 155, 10, "Utilities claimed (fill amount in or check below):"
  Text 5, 160, 70, 10, "Sign your case note:"
  Text 5, 10, 50, 10, "Case number: "
  Text 5, 30, 95, 10, "Income received this month:"
EndDialog

'It will search for a case number.
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row <> 0 then 
  EMReadScreen case_number, 8, row, col + 10
  case_number = replace(case_number, "_", "")
  case_number = replace(case_number, " ", "")
End if

Do
  Dialog exp_screening_dialog
  If exp_screening_dialog_ButtonPressed = 0 then stopscript
'The following is an experimental way to find a MAXIS screen. 
'It only connects to the MAXIS screen, and could work from a third party script host potentially.
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
      If MAXIS_check <> "MAXIS" then 
        MsgBox "Neither screen appears to be on MAXIS. Get one of your screens to MAXIS before proceeding. You may need to enter a password. If you are in MAXIS, you may have had a configuration error. To fix this, restart BlueZone."
        stopscript
      End if
    End If
  End If
  If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) then
    MsgBox "The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
    MAXIS_check = "" 'By clearing this variable we simplify the do loop. The script will not finish until this info is complete.
  End if
  If worker_sig = "" then
    MsgBox "You must sign your case note."
    MAXIS_check = "" 'By clearing this variable we simplify the do loop. The script will not finish until this info is complete.
  End if
Loop until MAXIS_check = "MAXIS"

'SECTION 02
If income = "" then income = 0
If assets = "" then assets = 0
If rent = "" then rent = 0
If utilities = "" then utilities = 0
If phone_check = 1 then utilities = 37                            '$37 is the phone standard for utility calculation as of October 2011.
If electric_check = 1 then utilities = 120                        '$120 is the electric standard for utility calculation as of October 2011.
If electric_check = 1 and phone_check = 1 then utilities = 157    'Phone standard plus electric standard.
If heat_AC_check = 1 then utilities = 402                         '$402 is the maximum utility standard as of October 2011. If a client qualifies for this, they do not get the other two.


If (cint(income) < 150 and cint(assets) < 100) or ((cint(income) + cint(assets)) < (cint(rent) + cint(utilities))) then expedited_status = "client appears expedited"
If (cint(income) + cint(assets) >= cint(rent) + cint(utilities)) and (cint(income) >= 150 or cint(assets) >= 100) then expedited_status = "client does not appear expedited"

'SECTION 03
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
  EMWaitReady 1, 1
  has_DISQ = False
End if

If DISQ_member_check <> "DISQ DOES NOT EXIST FOR ANY MEMBER" then 
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  has_DISQ = True
End if

'SECTION 04
EMWriteScreen "case", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "note", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<PF9>"
EMWaitReady 1, 1
EMSendKey "<home>" 'To get to the top of the case note.
If APPLed_case = 1 and paper_process_check = 1 then EMSendKey "APPLed PP intake for " + worker_name + ", " + expedited_status + "<newline>"
If APPLed_case = 1 and paper_process_check = 0 then EMSendKey "APPLed intake for " + worker_name + ", " + expedited_status + "<newline>"
If APPLed_case = 0 then EMSendKey "Reviewed CAF, " + expedited_status + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey "     CAF 1 income claimed this month: $" & income & "<newline>"
EMSendKey "         CAF 1 liquid assets claimed: $" & assets & "<newline>"
EMSendKey "         CAF 1 rent/mortgage claimed: $" & rent & "<newline>"
EMSendKey "        Utilities (amt/HEST claimed): $" & utilities & "<newline>"
EMSendKey "---" + "<newline>"
If has_DISQ = True then EMSendKey "A DISQ panel exists for someone on this case." + "<newline>"
If has_DISQ = False then EMSendKey "No DISQ panels were found for this case." + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig
If expedited_status = "client appears expedited" then
  MsgBox "This client appears expedited. A same day interview needs to be offered."
End if
If expedited_status = "client does not appear expedited" then
  MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
End if

'SECTION 05
If appointment_letter_check = "1" then run "H:\BlueZone\bzsh.exe Q:\Blue Zone Scripts\Family\MEMO - CAF received.vbs"
stopscript
