'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - expedited screening"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 01
EMConnect ""

BeginDialog exp_screening_dialog, 0, 0, 216, 175, "Expedited Screening Dialog"
  EditBox 55, 5, 80, 15, case_number
  EditBox 100, 25, 50, 15, income
  EditBox 100, 45, 50, 15, assets
  EditBox 115, 65, 50, 15, rent
  CheckBox 15, 95, 55, 10, "Heat (or AC)", heat_AC_check
  CheckBox 75, 95, 45, 10, "Electricity", electric_check
  CheckBox 130, 95, 35, 10, "Phone", phone_check
  DropListBox 70, 115, 120, 15, "intake"+chr(9)+"add-a-program", application_type
  CheckBox 5, 140, 180, 10, "Check here if this appointment is a paper process.", paper_process_check
  EditBox 125, 155, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 10, 50, 15
    CancelButton 160, 30, 50, 15
  Text 5, 10, 50, 10, "Case number: "
  Text 5, 30, 95, 10, "Income received this month:"
  Text 5, 50, 95, 10, "Cash, checking, or savings: "
  Text 5, 70, 105, 10, "Amounts paid for rent/mortgage:"
  GroupBox 5, 85, 165, 25, "Utilities claimed (check below):"
  Text 5, 120, 65, 10, "Application is for:"
  Text 50, 160, 70, 10, "Sign your case note:"
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

paper_process_check = 1 'Paper process intake is the default, so the script needs to default to checked for paper process.

Do
  Dialog exp_screening_dialog
  If ButtonPressed = 0 then stopscript
'The following is an experimental way to find a MAXIS screen. 
'It only connects to the MAXIS screen, and could work from a third party script host potentially.
  EMConnect "A"
  transmit
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" then
    EMConnect "B"
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then 
      EMConnect "C"
      transmit
      EMReadScreen MAXIS_check, 5, 1, 39
      If MAXIS_check <> "MAXIS" then script_end_procedure("Neither screen appears to be on MAXIS. Get one of your screens to MAXIS before proceeding. You may need to enter a password. If you are in MAXIS you may have had a configuration error. To fix this restart BlueZone.")
    End If
  End If
  If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) then
    MsgBox "The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
    MAXIS_check = "" 'By clearing this variable we simplify the do loop. The script will not finish until this info is complete.
  End if
  If worker_signature = "" then
    MsgBox "You must sign your case note."
    MAXIS_check = "" 'By clearing this variable we simplify the do loop. The script will not finish until this info is complete.
  End if
Loop until MAXIS_check = "MAXIS"

'SECTION 02
If income = "" then income = 0
If assets = "" then assets = 0
If rent = "" then rent = 0
If phone_check = 1 then utilities = 40                                               '$40 is the phone standard for utility calculation as of November 2013.
If electric_check = 1 then utilities = 141                                           '$141 is the electric standard for utility calculation as of November 2013.
If electric_check = 1 and phone_check = 1 then utilities = 181                       'Phone standard plus electric standard.
If heat_AC_check = 1 then utilities = 459                                            '$459 is the maximum utility standard as of November 2013. If a client qualifies for this, they do not get the other two.
If phone_check = 0 and electric_check = 0 and heat_AC_check = 0 then utilities = 0   'in case no options are clicked, utilities is set to zero.

income = Abs(income)
assets = Abs(assets)
rent = Abs(rent)
utilities = Abs(utilities)



If (income < 150 and assets < 100) or ((income + assets) < (rent + utilities)) then expedited_status = "client appears expedited"
If (income + assets >= rent + utilities) and (income >= 150 or assets >= 100) then expedited_status = "client does not appear expedited"

'SECTION 03
'This jumps back to SELF
back_to_self

EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
footer_month = datepart("m", date) 'This is so the footer month is correct for a CAF1 case.
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", date) - 2000
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "disq", 21, 70
transmit

EMReadScreen DISQ_member_check, 34, 24, 2
If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then 
  PF3
  has_DISQ = False
End if

If DISQ_member_check <> "DISQ DOES NOT EXIST FOR ANY MEMBER" then 
  PF3
  has_DISQ = True
End if

'SECTION 04
EMWriteScreen "case", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "note", 21, 70
transmit
PF9
EMReadScreen read_only_check, 41, 24, 2
If read_only_check = "YOU HAVE 'READ ONLY' ACCESS FOR THIS CASE" then script_end_procedure("You have read-only access to this case! You may be in inquiry or this may be out of county. Expedited status is indicated as: " & expedited_status & ". Try again or process/track manually.")
EMSendKey "<home>" 'To get to the top of the case note.
If paper_process_check = 1 then EMSendKey "APPLed PP " & application_type & ", " & expedited_status & "<newline>"
If paper_process_check = 0 then EMSendKey "APPLed " & application_type & ", " & expedited_status & "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey "     CAF 1 income claimed this month: $" & income & "<newline>"
EMSendKey "         CAF 1 liquid assets claimed: $" & assets & "<newline>"
EMSendKey "         CAF 1 rent/mortgage claimed: $" & rent & "<newline>"
EMSendKey "        Utilities (amt/HEST claimed): $" & utilities & "<newline>"
EMSendKey "---" + "<newline>"
If has_DISQ = True then EMSendKey "A DISQ panel exists for someone on this case." + "<newline>"
If has_DISQ = False then EMSendKey "No DISQ panels were found for this case." + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_signature
If expedited_status = "client appears expedited" then
  MsgBox "This client appears expedited. A same day interview needs to be offered. If this is a paper process, assign it using the expedited-after hours appointment type."
End if
If expedited_status = "client does not appear expedited" then
  MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
End if

script_end_procedure("")
