'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - overpayment-claim established"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'SECTION 02: DIALOGS

BeginDialog overpayment_dialog, 0, 0, 266, 222, "Overpayment dialog"
  EditBox 60, 5, 70, 15, case_number
  EditBox 120, 25, 140, 15, programs_cited
  EditBox 95, 45, 165, 15, months_of_overpayment
  EditBox 65, 65, 60, 15, discovery_date
  EditBox 200, 65, 60, 15, established_date
  EditBox 60, 85, 200, 15, reason_for_OP
  EditBox 85, 105, 175, 15, supporting_docs
  EditBox 80, 125, 180, 15, responsible_parties
  EditBox 60, 145, 200, 15, total_amt_of_OP
  EditBox 70, 165, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 200, 50, 15
    CancelButton 135, 200, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 115, 10, "Program(s) overpayment cited for:"
  Text 5, 50, 85, 10, "Month(s) of overpayment:"
  Text 5, 70, 55, 10, "Discovery date:"
  Text 135, 70, 60, 10, "Established date:"
  Text 5, 90, 55, 10, "Reason for OP:"
  Text 5, 110, 80, 10, "Supporting docs/verifs:"
  Text 5, 130, 70, 10, "Responsible parties:"
  Text 5, 150, 55, 10, "Total amt of OP:"
  Text 5, 170, 65, 10, "Sign the case note:"
  Text 130, 165, 125, 30, "Remember to ''staple'' the supporting documents to the claim form, and send to your supervisor for approval!"
EndDialog

'SECTION 03: THE SCRIPT

EMConnect ""

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

Do
  Do
    Do
      Dialog overpayment_dialog
      If buttonpressed = 0 then stopscript
      If case_number = "" then MsgBox "You must have a case number to continue!"
    Loop until case_number <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

call write_new_line_in_case_note("**OVERPAYMENT/CLAIM ESTABLISHED**")
call write_editbox_in_case_note("Program(s) overpayment cited for", programs_cited, 6) 
call write_editbox_in_case_note("Month(s) of overpayment", months_of_overpayment, 6) 
call write_editbox_in_case_note("Discovery date", discovery_date, 6) 
call write_editbox_in_case_note("Established date", established_date, 6) 
call write_editbox_in_case_note("Reason for overpayment", reason_for_OP, 6) 
call write_editbox_in_case_note("Supporting documents/verifications", supporting_docs, 6) 
call write_editbox_in_case_note("Responsible parties", responsible_parties, 6) 
call write_editbox_in_case_note("Total overpayment amount", total_amt_of_OP, 6) 
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")
