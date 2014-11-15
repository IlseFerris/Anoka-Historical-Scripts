BeginDialog Dialog1, 0, 0, 191, 47, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 5, 5, 50, 15, "case number:"
  EditBox 60, 0, 65, 15, case_number
EndDialog

dialog dialog1

EMConnect "A"

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 0

'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog Dialog1
     IF buttonpressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"


'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMSendKey "<attn>"
EMWaitReady 1, 0
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
EMReadScreen MMIS_A_check, 7, 15, 15
If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If training_check = "RUNNING" then EMSendKey "3" + "<enter>"
If production_check = "RUNNING" then EMSendKey "1" + "<enter>"

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"

EMSetCursor 16, 43
EMSendKey "stat"
EMSetCursor 18, 43
EMSendkey "<eraseeof>" + case_number + "<enter>"
EMWaitReady 1, 0
EMReadScreen case_number_stat, 8, 20, 37
EMSendKey "memb" + "<enter>"
EMWaitReady 1, 0

'The following checks for which screen MMIS is running on.
IF MMIS_A_check = "RUNNING" then EMSendKey "<attn>" 
IF MMIS_A_check = "RUNNING" then EMWaitReady 1, 0
IF MMIS_A_check = "RUNNING" then EMSendKey "10" + "<enter>"
IF MMIS_A_check = "RUNNING" then EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMConnect "B"
EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMSendKey "<attn>"
EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMReadScreen MMIS_B_check, 7, 15, 15
If MMIS_A_check <> "RUNNING" and MMIS_B_check <> "RUNNING" then MsgBox "MMIS does not appear to be running. This script will now stop."
If MMIS_A_check <> "RUNNING" and MMIS_B_check <> "RUNNING" then stopscript
IF MMIS_A_check <> "RUNNING" and MMIS_B_check = "RUNNING" then EMSendkey "10" + "<enter>"

EMFocus

  Sub get_to_session_begin 'This sub uses a Do Loop to get to the start screen for MMIS.
    Do 
    EMSendkey "<PF6>"
      EMReadScreen password_prompt2, 38, 2, 23
      IF password_prompt2 = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
    EMWaitReady 1, 0
    EMReadScreen session_start, 18, 1, 7
    Loop until session_start = "SESSION TERMINATED"
  End Sub

get_to_session_begin
EMSetCursor 1, 2
EMSendKey "mw00"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "<enter>"
EMWaitReady 1, 0

'This section may not work for all OSAs, since some only have EK01.
  row = 1
  col = 1
EMSearch "EK01", row, col
If row <> 0 then EMSetCursor row, 4
If row <> 0 then EMSendKey "x"
If row <> 0 then EMSendKey "<enter>"
If row <> 0 then EMWaitReady 1, 0

'This section starts from EK01. OSAs may need to skip the previous section.
EMSetCursor 10, 3
EMSendKey "x"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMFocus

'Now we are in MMIS, and it will get to RELG for the associated SSN
EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_number_stat

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = "_" then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = "_" then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = "_" then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = "_" then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0

