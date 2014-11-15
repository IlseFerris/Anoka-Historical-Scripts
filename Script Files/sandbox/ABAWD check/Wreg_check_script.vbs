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


BeginDialog WREG_dlg, 0, 0, 336, 310, "Dialog"
  Text 15, 10, 195, 20, "Ask client if they meet each exemption and check the boxes they do meet. Then sign and hit OK"
  CheckBox 15, 35, 180, 15, "Temp/Perm Disabled (minimum of 30 days)", disa
  CheckBox 15, 55, 135, 15, "Care of disabled unit member", care_disa
  CheckBox 15, 75, 135, 15, "Under 16", under_16
  CheckBox 15, 95, 265, 15, "Aged 16 or 17 and living with parent or caretaker", living_with_parent
  CheckBox 15, 115, 195, 15, "Responsible for care of a child under age 6", child_care
  CheckBox 15, 135, 295, 20, "Employed or self employed 30 hours per week or equivalent 30 hours X Minimum wage", hours_30
  CheckBox 15, 155, 225, 15, "Receiving/Applied for Unemployment insurance", unemployment
  CheckBox 15, 175, 160, 15, "Enrolled in school or training 1/2 time", school
  CheckBox 15, 195, 270, 20, "Participating regularly in drug addiction/alcohol treatment & rehab program", drug
  CheckBox 15, 220, 155, 15, "Receiving MFIP", MFIP
  CheckBox 15, 240, 200, 15, "Pending/Receiving DWP", DWP
  Text 15, 265, 70, 15, "Worker signature"
  EditBox 80, 265, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 285, 50, 15
    CancelButton 135, 285, 50, 15
EndDialog

Do
	dialog wreg_dlg
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

call write_new_line_in_case_note("WREG status")

'write reason why
If disa = 1 then
	call write_new_line_in_case_note("Client states they are disabled")
end if

If care_disa = 1 then
	call write_new_line_in_case_note("Client states they are responsible for care of a disabled unit member")
end if

If under_16 = 1 then
	call write_new_line_in_case_note("Client states they are under 16")
end if

If living_with_parent = 1 then
	call write_new_line_in_case_note("Client states they are age 16 or 17 and living with a parent or caretaker")
end if

If child_care = 1 then
	call write_new_line_in_case_note("Client states they are responsible for the care of a child less than age 6")
end if

If hours_30 = 1 then
	call write_new_line_in_case_note("Client states they are employed 30 hours per week or equivalent to 30 hours a week at minimum wage")
end if

If unemployment = 1 then
	call write_new_line_in_case_note("Client states they are receiving or applied for unemployment insurance")
end if

If school = 1 then
	call write_new_line_in_case_note("Client states they are enrolled in school/training 1/2 time")
end if

If drug = 1 then
	call write_new_line_in_case_note("Client states they are residing in a waiver area")
end if

If MFIP = 1 then
	call write_new_line_in_case_note("Client states they are a MFIP recipient")
end if

If DWP = 1 then
	call write_new_line_in_case_note("Client states they are a DWP recipient")
end if

'write determination
If disa or care_disa or under_16 or living_with_parent or hours_30 or unemployment or school or drug or MFIP or DWP = 1 then
	call write_new_line_in_case_note("Client is exempt from WREG.")
else
	call write_new_line_in_case_note("Client is not exempt from WREG")
end if

'sign case note
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)