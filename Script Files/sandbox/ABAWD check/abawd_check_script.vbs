'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = ""
start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS--------------------------------------------------------------------------------------------------


BeginDialog ABAWD_dlg, 0, 0, 251, 295, "Dialog"
  Text 15, 10, 195, 20, "Ask client if they meet each exemption and check the boxes they do meet. Then hit OK"
  CheckBox 15, 35, 130, 15, "Work Registration Exempt", wreg
  CheckBox 15, 55, 135, 15, "Under 18", under_18
  CheckBox 15, 75, 135, 15, "Age 50 or older", Over_50
  CheckBox 15, 95, 265, 15, "Responsible for care of a dependent child in the household", dep_child
  CheckBox 15, 115, 150, 15, "Medically certified as pregnant", pregnant
  CheckBox 15, 135, 155, 15, "Employed 20 hours per week", hours_20
  CheckBox 15, 155, 145, 15, "Participating in approved work experience program", work_exp
  CheckBox 15, 175, 160, 15, "Participating in approved E & T program", e_and_t
  CheckBox 15, 195, 220, 15, "Residing in area granted waiver form ABAWD requirements", waiver
  CheckBox 15, 215, 155, 15, "RCA or GA recipient", RCA_GA
  Text 10, 240, 70, 15, "Worker signature"
  EditBox 80, 240, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 265, 50, 15
    CancelButton 135, 265, 50, 15
EndDialog



Do
	dialog ABAWD_dlg
	If buttonpressed = 0 then stopscript
	If worker_signature = "" then msgbox("Please sign your name")
Loop until worker_signature <> ""

case_number = 201294

EMConnect ""

'going to case note
Call navigate_to_screen("case","note")


'writing case note
PF9

EMSetcursor 4, 3

call write_new_line_in_case_note("ABAWD status")

'write reason why
If wreg = 1 then
	call write_new_line_in_case_note("Client states they are work registration exempt")
end if

If under_18 = 1 then
	call write_new_line_in_case_note("Client states they are under 18 years old")
end if

If over_50 = 1 then
	call write_new_line_in_case_note("Client states they are 50 years old or over")
end if

If dep_child = 1 then
	call write_new_line_in_case_note("Client states they are responsible for the care of a dependent child in the household")
end if

If pregnant = 1 then
	call write_new_line_in_case_note("Client states they are medically certified as pregnant")
end if

If hours_20 = 1 then
	call write_new_line_in_case_note("Client states they are employed 20 hours per week")
end if

If work_exp = 1 then
	call write_new_line_in_case_note("Client states they are participatiing in work experience program")
end if

If e_and_t = 1 then
	call write_new_line_in_case_note("Client states they are participating in employment and training program")
end if

If waiver = 1 then
	call write_new_line_in_case_note("Client states they are residing in a waiver area")
end if

If RCA_GA = 1 then
	call write_new_line_in_case_note("Client states they are a RCA or GA recipient")
end if

'write determination
If wreg or under_18 or Over_50 or dep_child or pregnant or hours_20 or work_exp or e_and_t or waiver or RCA_GA = 1 then
	call write_new_line_in_case_note("Client is exempt from ABAWD. This person may not be an ABAWD refer to worker for further verification")
else
	call write_new_line_in_case_note("Client is not exempt from ABAWD. This person is ABAWD and subject to 3 month limit")
end if


call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)