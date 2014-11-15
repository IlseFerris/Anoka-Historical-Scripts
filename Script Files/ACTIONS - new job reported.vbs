'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - new job reported"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog case_number_dialog, 0, 0, 161, 42, "Case number"
  EditBox 95, 0, 60, 15, case_number
  ButtonGroup ButtonPressed
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

BeginDialog new_job_reported_dialog, 0, 0, 241, 225, "New job reported dialog"
  EditBox 80, 25, 25, 15, HH_memb
  EditBox 45, 45, 195, 15, employer
  EditBox 90, 65, 100, 15, who_reported_job
  ComboBox 100, 85, 90, 15, "phone call"+chr(9)+"office visit"+chr(9)+"mailing"+chr(9)+"fax"+chr(9)+"ES counselor"+chr(9)+"CCA worker", job_report_type
  EditBox 30, 105, 210, 15, notes
  CheckBox 5, 125, 190, 10, "Check here to have the script make a new JOBS panel.", create_JOBS_check
  CheckBox 5, 140, 190, 10, "Check here if you sent a status update to CCA.", CCA_check
  CheckBox 5, 155, 190, 10, "Check here if you sent a status update to ES.", ES_check
  CheckBox 5, 170, 190, 10, "Check here if you sent a Work Number request.", work_number_check
  EditBox 70, 185, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 205, 50, 15
    CancelButton 125, 205, 50, 15
    PushButton 125, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 125, 25, 45, 10, "next panel", next_panel_button
    PushButton 185, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 185, 25, 45, 10, "next memb", next_memb_button
  GroupBox 120, 5, 115, 35, "STAT-based navigation"
  Text 5, 30, 70, 10, "HH member number:"
  Text 5, 50, 40, 10, "Employer:"
  Text 5, 70, 80, 10, "Who reported the job?:"
  Text 5, 90, 90, 10, "How was the job reported?:"
  Text 5, 110, 25, 10, "Notes:"
  Text 5, 190, 60, 10, "Worker signature:"
EndDialog




'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

'Finds a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(replace(case_number, "_", ""))
If isnumeric(case_number) = False then case_number = ""

'Shows the case number dialog
Dialog case_number_dialog
If ButtonPressed = 0 then stopscript

'It sends an enter to force the screen to refresh, in order to check for MAXIS. If MAXIS isn't found the script will stop.
transmit
EMReadScreen MAXIS_check, 5, 1, 39
IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS not found. Are you in MAXIS on the screen you started the script? Check and try again. If it still doesn't work try shutting down BlueZone and starting it up again.")

'Now it enters stat/jobs. It'll check to make sure it gets past the SELF menu and gets onto the JOBS panel.
call navigate_to_screen("stat", "jobs")
EMReadScreen SELF_check, 27, 2, 28
If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("Unable to navigate past the SELF menu. Is your case in background? Wait a few seconds and try again.")
EMReadScreen JOBS_check, 4, 2, 45
If JOBS_check <> "JOBS" then transmit
EMReadScreen PW_check, 4, 21, 21
If PW_check <> "X102" then script_end_procedure("This client is out of county. The script will stop. Check your case number and try again.")

'Declaring some variables to create defaults for the new_job_reported_dialog.
create_JOBS_check = 1
HH_memb = "01"
HH_memb_row = 5 'This helps the navigation buttons work!

'Shows the dialog.
Do
  Do
    Do
      Do
        Do
          Dialog new_job_reported_dialog
          If ButtonPressed = 0 then stopscript
          EMReadScreen STAT_check, 4, 20, 21
          If STAT_check = "STAT" then call stat_navigation
          transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
          EMReadScreen MAXIS_check, 5, 1, 39
          If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again.")
        Loop until ButtonPressed = -1
        If employer = "" then MsgBox "You must type an employer!"
      Loop until employer <> ""
      If who_reported_job = "" then MsgBox "You must type out who reported the job!"
    Loop until who_reported_job <> ""
    If job_report_type = "" then MsgBox "You must select how you heard about the job, or write something in that field yourself."
  Loop until job_report_type <> ""
  If worker_signature = "" then MsgBox "You must sign your case note!"
Loop until worker_signature <> ""

'Creates a new JOBS panel if that was selected.
If create_JOBS_check = 1 then
  EMWriteScreen HH_memb, 20, 76
  EMWriteScreen "nn", 20, 79
  transmit
  EMReadScreen edit_mode_check, 1, 20, 8
  If edit_mode_check = "D" then script_end_procedure("Unable to create a new JOBS panel. Check which member number you provided. Otherwise you may be in inquiry mode. If so shut down inquiry and try again. Or try closing BlueZone.")
  EMWriteScreen "w", 5, 38
  EMWriteScreen "n", 6, 38
  EMWriteScreen employer, 7, 42
  EMReadScreen footer_month, 2, 20, 55
  EMReadScreen footer_year, 2, 20, 58
  EMWriteScreen footer_month, 12, 54
  EMWriteScreen "01", 12, 57
  EMWriteScreen footer_year, 12, 60
  EMWriteScreen "0", 12, 67
  EMWriteScreen "0", 18, 72
  Do
    transmit
    EMReadScreen edit_mode_check, 1, 20, 8
  Loop until edit_mode_check = "D"
End if

'Jumps to case note the info.
call navigate_to_screen("case", "note")
PF9
EMReadScreen edit_mode_check, 1, 20, 9
If edit_mode_check = "D" then script_end_procedure("Unable to create a new case note. Your case may be in inquiry. If so shut down inquiry and try again. Or try closing BlueZone.")

'Now the script will case note what's happened.
EMSendKey ">>>New job for MEMB " & HH_memb & " reported by " + who_reported_job + " via " + job_report_type + "<<<" + "<newline>" 
call write_editbox_in_case_note("Employer", employer, 6)
if CCA_check = 1 then call write_new_line_in_case_note("* Sent status update to CCA.")
if ES_check = 1 then call write_new_line_in_case_note("* Sent status update to ES.")
if work_number_check = 1 then call write_new_line_in_case_note("* Sent Work Number request.")
if notes <> "" then call write_editbox_in_case_note("Notes", notes, 6)
call write_new_line_in_case_note("* Sending employment verification. TIKLed for 10-day return.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Navigating to DAIL/WRIT
call navigate_to_screen("dail", "writ")

'The following will generate a TIKL formatted date for 10 days from now.
call create_MAXIS_friendly_date(date, 10, 5, 18)

'Writing in the rest of the TIKL.
EMSetCursor 9, 3
EMSendKey "Verification of job change should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)." 
transmit
PF3
MsgBox "Success! MAXIS updated for job change, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at " & employer & "."

script_end_procedure("")
