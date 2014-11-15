'VARIABLES TO REMOVE FROM FINAL SCRIPT:
case_number = "189053"
PMI_number = "00001280"

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Open Enrollment"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'The following checks for which screen MMIS is running on.
'<<<<<<<<<<<<<<<<<NOTE: currently set to check for training region. Row should be set to 16 in order to function in test, 15 in production. Sendkey should be "10" and not "11" for production as well.
attn
EMReadScreen MMIS_A_check, 7, 16, 15
IF MMIS_A_check = "RUNNING" then 
  EMSendKey "11"
  transmit
End if
IF MMIS_A_check <> "RUNNING" then 
  attn
  EMConnect "B"
  attn
  EMReadScreen MMIS_B_check, 7, 16, 15
  If MMIS_B_check <> "RUNNING" then script_end_procedure("MMIS does not appear to be running. This script will now stop.")
  If MMIS_B_check = "RUNNING" then 
    EMSendKey "11"
    transmit
  End if
End if

'Shifts user focus to whatever screen ended up getting selected (A or B)
EMFocus

'Sends a PF6 to see if MMIS is running.
PF6
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then 
  objExcel.Workbooks.Close
  objExcel.quit
  script_end_procedure("You appear to be passworded out in MMIS.")
End if

'Gets back to the start screen of MMIS
get_to_MMIS_session_begin

'Gets to the EK01 screen (NOTE: this section may not work for all staff, if they only have EK01.)
EMSetCursor 1, 2
EMSendKey "mw00"
transmit
transmit

'Setting the variables for the next search
row = 1
col = 1

'Looking to confirm we're on C302. If not it'll try to get in the recipient file application setting.
EMSearch "C302", row, col
If row <> 0 then 
  EMSetCursor row, 4
  EMSendKey "x"
  transmit
Else
  row = 1
  col = 1
  EMSearch "EKIQ", row, col
  If row <> 0 then
    EMSetCursor row, 4
    EMSendKey "x"
    transmit
  Else
    row = 1
    col = 1
    EMSearch "RECIPIENT FILE APPLICATION", row, col
    If row = 0 then
      script_end_procedure("MMIS could not be found for this user. Contact the script administrator with your name, x102 number, and a description of the problem.")
    Else
      EMWriteScreen "x", row, col - 3
      transmit
    End if
  End if
End if

'This section starts from C302. Getting to RKEY
EMWriteScreen "x", 8, 3
transmit

'Converting case number into the 8 digit requirements for MMIS.
If len(case_number) < 8 then 'This will generate an 8 digit MAXIS case number.
  Do 
    case_number = "0" & case_number
  Loop until len(case_number) = 8
End if

'Sending the "c" code for RKEY.
EMWriteScreen "c", 2, 19

'Sending the case number for RKEY.
EMWriteScreen case_number, 9, 19

'Getting into the case file path.
transmit

'Navigates to RCIN
EMWriteScreen "RCIN", 1, 8
transmit

'Sets the MMIS_row and MMIS_col variables
MMIS_row = 10
MMIS_col = 1

'Searches for and selects the PMI indicated, and enters the recipient file path.
EMSearch PMI_number, MMIS_row, MMIS_col
If MMIS_row = 0 then
  script_end_procedure("PMI is not found. Doublecheck the PMI you selected and try again.")
Else
  EMWriteScreen "x", MMIS_row, 2
  transmit
End if

'Navigating to RPPH.
EMWriteScreen "RPPH", 1, 8
transmit