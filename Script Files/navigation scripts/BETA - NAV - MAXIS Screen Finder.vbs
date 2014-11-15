'Informational front-end message, date dependent.
If datediff("d", "12/10/2011", now) < 365 then MsgBox "Let me know what you think! -Ronny"

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BETA - NAV - MAXIS screen finder"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MAXIS_screen_finder_dialog, 0, 0, 261, 265, "MAXIS screen finder"
  EditBox 210, 225, 45, 15, case_number
  ButtonGroup ButtonPressed
    CancelButton 170, 245, 50, 15
    PushButton 75, 25, 45, 10, "STAT/JOBS", STAT_JOBS_button
    PushButton 75, 40, 45, 10, "STAT/UNEA", STAT_UNEA_button
    PushButton 75, 55, 45, 10, "STAT/BUSI", STAT_BUSI_button
    PushButton 75, 90, 45, 10, "CASE/CURR", CASE_CURR_button
    PushButton 75, 105, 45, 10, "ELIG/HC", ELIG_HC_button
    PushButton 75, 120, 45, 10, "ELIG/MFIP", ELIG_MFIP_button
    PushButton 75, 135, 45, 10, "ELIG/DWP", ELIG_DWP_button
    PushButton 75, 175, 45, 10, "STAT/ABPS", STAT_ABPS_button
    PushButton 75, 190, 45, 10, "INFC/CSIA", INFC_CSIA_button
    PushButton 75, 205, 45, 10, "INFC/CSIB", INFC_CSIB_button
    PushButton 75, 220, 45, 10, "INFC/CSIC", INFC_CSIC_button
    PushButton 75, 235, 45, 10, "INFC/CSID", INFC_CSID_button
    PushButton 205, 25, 45, 10, "STAT/MEMB", STAT_MEMB_button
    PushButton 205, 40, 45, 10, "STAT/PARE", STAT_PARE_button
    PushButton 205, 55, 45, 10, "CASE/PERS", CASE_PERS_button
    PushButton 205, 90, 45, 10, "MONY/INQB", MONY_INQB_button
    PushButton 205, 125, 45, 10, "STAT/ADDR", STAT_ADDR_button
    PushButton 205, 140, 45, 10, "STAT/DISA", STAT_DISA_button
    PushButton 205, 155, 45, 10, "STAT/INSA", STAT_INSA_button
    PushButton 205, 170, 45, 10, "STAT/PBEN", STAT_PBEN_button
    PushButton 205, 185, 45, 10, "STAT/SANC", STAT_SANC_button
    PushButton 205, 200, 45, 10, "CASE/NOTE", CASE_NOTE_button
  GroupBox 5, 10, 120, 60, "Income"
  Text 10, 25, 65, 10, "Earned Wages:"
  Text 10, 40, 65, 10, "Unearned Income:"
  Text 10, 55, 65, 10, "Self Employment:"
  GroupBox 5, 75, 120, 75, "PA Programs"
  Text 10, 90, 65, 10, "Current Status:"
  Text 10, 105, 65, 10, "HC:"
  Text 10, 120, 65, 10, "MFIP:"
  Text 10, 135, 65, 10, "DWP:"
  GroupBox 5, 160, 120, 90, "ALF/NCP"
  Text 10, 175, 65, 10, "Absent Parent:"
  Text 10, 190, 65, 10, "CS Interface A:"
  Text 10, 205, 65, 10, "CS Interface B:"
  Text 10, 220, 65, 10, "CS Interface C:"
  Text 10, 235, 65, 10, "CS Interface D:"
  GroupBox 135, 10, 120, 60, "HH members"
  Text 140, 25, 65, 10, "Basic person info:"
  Text 140, 40, 65, 10, "Parent info:"
  Text 140, 55, 65, 10, "Case person info:"
  GroupBox 135, 75, 120, 30, "Money Stuff"
  Text 140, 90, 65, 10, "PA disbursements:"
  GroupBox 135, 115, 120, 100, "Other"
  Text 140, 125, 65, 10, "Address/Phone:"
  Text 140, 140, 65, 10, "Disability status:"
  Text 140, 155, 65, 10, "Insurance:"
  Text 140, 170, 65, 10, "Potential Benefits:"
  Text 140, 185, 65, 10, "Sanction:"
  Text 140, 200, 65, 10, "Case Notes:"
  Text 135, 230, 70, 10, "MAXIS case number:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connect to BlueZone
EMConnect ""

'Finds case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(replace(case_number, "_", "")) 'replaces underscores and spaces in the variable

'Shows dialog
Do
  Dialog MAXIS_screen_finder_dialog
  If buttonpressed = 0 then stopscript
  If isnumeric(case_number) = false then MsgBox "You must enter a valid MAXIS case number! No letters, all numeric."
Loop until isnumeric(case_number) = True

'Figure out if we're in MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check = "IS   " then 'Because of a glitch on MONY/INQB, this will work around rewriting the functions file
  PF3
  EMReadScreen MAXIS_check, 5, 1, 39
End if
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You are not currently in MAXIS. Navigate to MAXIS and try again. Make sure you aren't passworded out!")

'Now it'll navigate to any of the screens chosen
If buttonpressed = STAT_JOBS_button then call navigate_to_screen("STAT", "JOBS")
If buttonpressed = STAT_UNEA_button then call navigate_to_screen("STAT", "UNEA")
If buttonpressed = STAT_BUSI_button then call navigate_to_screen("STAT", "BUSI")
If buttonpressed = CASE_CURR_button then call navigate_to_screen("CASE", "CURR")
If buttonpressed = ELIG_HC_button then call navigate_to_screen("ELIG", "HC__")
If buttonpressed = ELIG_MFIP_button then call navigate_to_screen("ELIG", "MFIP")
If buttonpressed = ELIG_DWP_button then call navigate_to_screen("ELIG", "DWP_")
If buttonpressed = STAT_ABPS_button then call navigate_to_screen("STAT", "ABPS")
If buttonpressed = INFC_CSIA_button then call navigate_to_screen("INFC", "CSIA")
If buttonpressed = INFC_CSIB_button then call navigate_to_screen("INFC", "CSIB")
If buttonpressed = INFC_CSIC_button then call navigate_to_screen("INFC", "CSIC")
If buttonpressed = INFC_CSID_button then call navigate_to_screen("INFC", "CSID")
If buttonpressed = STAT_MEMB_button then call navigate_to_screen("STAT", "MEMB")
If buttonpressed = STAT_PARE_button then call navigate_to_screen("STAT", "PARE")
If buttonpressed = CASE_PERS_button then call navigate_to_screen("CASE", "PERS")
If buttonpressed = MONY_INQB_button then call navigate_to_screen("MONY", "INQB")
If buttonpressed = STAT_ADDR_button then call navigate_to_screen("STAT", "ADDR")
If buttonpressed = STAT_DISA_button then call navigate_to_screen("STAT", "DISA")
If buttonpressed = STAT_INSA_button then call navigate_to_screen("STAT", "INSA")
If buttonpressed = STAT_PBEN_button then call navigate_to_screen("STAT", "PBEN")
If buttonpressed = STAT_SANC_button then call navigate_to_screen("STAT", "SANC")
If buttonpressed = CASE_NOTE_button then call navigate_to_screen("CASE", "NOTE")

script_end_procedure("")