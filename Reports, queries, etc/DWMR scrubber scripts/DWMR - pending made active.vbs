


'-----The following is designed to work from the DWMR scrubber. As such, it should not be ran without the DWMR scrubber starting it.


MsgBox "Check this case to see if MAXIS is active. If MAXIS is not active, XFER MCRE and MAXIS to the HC team using Excel."
StopScript


'---At this time, the script just reads a message and then stops. What follows is a way to enter MMIS to check SSN info.

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

'Now it enters the SSN
EMWriteScreen "i", 2, 19
EMWriteScreen SSN, 5, 19
EMSendKey "<enter>"