


'-----The following is designed to work from the DWMR scrubber. As such, it should not be ran without the DWMR scrubber starting it.


EMConnect ""

BeginDialog SSN_dialog, 0, 0, 121, 37, "SSN Dialog"
  EditBox 30, 0, 90, 15, SSN
  ButtonGroup SSN_ButtonPressed
    OkButton 10, 20, 50, 15
    CancelButton 65, 20, 50, 15
  Text 5, 5, 20, 10, "SSN:"
EndDialog

Dialog SSN_dialog

'The following checks for which screen MMIS is running on.
EMSendKey "<attn>"
EMWaitReady 1, 0
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
EMReadScreen MMIS_A_check, 7, 15, 15
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

'Now it enters the SSN and goes to RCIP to check the status.
EMWriteScreen "i", 2, 19
EMWriteScreen SSN, 5, 19
EMSendKey "<enter>"
EMWaitReady 1, 0
EMWritescreen "rcip", 1, 8
EMSendKey "<enter>"
EMWaitReady 1, 0
EMReadScreen ID_status, 1, 6, 66
If ID_status = "B" or ID_status = "F" or ID_status = "H" or ID_status = "U" then msgbox "No action required. This number is verified."
If ID_status <> "B" and ID_status <> "F" and ID_status <> "H" and ID_status <> "U" then Msgbox "Script was unable to verify ID. Do this manually. You may need to consult Christa."