'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT REVW (sups)"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


'THE SCRIPT

EMConnect ""


BeginDialog x102_dialog, 0, 0, 136, 47, "x102 dialog"
  Text 5, 10, 90, 10, "Enter the x102# (x102***):"
  EditBox 100, 5, 30, 15, x102_number
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 75, 25, 50, 15
EndDialog

dialog x102_dialog
If buttonpressed = 0 then stopscript

transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS is not found on this screen.")

call navigate_to_screen("rept", "revw")

If x102_number <> "" then
  EMWriteScreen x102_number, 21, 10
  transmit
End if

script_end_procedure("")