'Stats gathering --------------------------------------------------------------------------------------------
name_of_script = ""
start_time = timer

'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Dialogs ----------------------------------------------------------------------
BeginDialog Case_note, 0, 0, 191, 225, "Case note"
  ButtonGroup ButtonPressed
    OkButton 125, 180, 50, 15
    CancelButton 70, 180, 50, 15
  EditBox 100, 70, 70, 25, worker_signature
  CheckBox 20, 30, 90, 20, "NCP is compliant", Compliant
  Text 15, 75, 70, 15, "Worker Signature"
EndDialog


Do
   	Dialog case_note
	
	If worker_signature = "" then MsgBox ("Please sign case note")
Loop until worker_signature <> ""

case_number = "000000019-01"

EMConnect ""

call navigate_to_Prism_screen("CAAD")

PF5

EMSetCursor 16, 04

call write_new_line_in_PRISM_case_note ("Case Note", case_note, 5)
If compliant = 1 then
call write_new_line_in_PRISM_case_note ("NCP compliant: no enf necessary")
else
call write_new_line_in_PRISM_case_note ("NCP non-compliant: start enforcement")

call write_editbox_in_PRISM_case_note ("Worker", worker_signature,5)
end if


