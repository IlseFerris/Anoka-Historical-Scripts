'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT ACTV - bottom (sups)"
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

'Now it checks to make sure MAXIS is running on this screen. If both are running the script will stop.
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then
  MsgBox "MAXIS is not found. Run this script on the window that has MAXIS running."
  StopScript
End if

do
  PF3
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"
EMWriteScreen "rept", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "actv", 21, 70
transmit

EMWriteScreen x102_number, 21, 17
transmit

do
  PF8
  EMReadScreen test, 21, 24, 2
loop until test = "THIS IS THE LAST PAGE"

script_end_procedure("")