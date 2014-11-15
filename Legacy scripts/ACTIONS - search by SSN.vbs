EMConnect ""

  Sub MAXIS_on   'This sub checks to see if MAXIS and MMIS are running on the display.
     EMSendKey "<Attn>"
     EMWaitReady 1, 0
     EMReadScreen MAI, 16, 6, 6
     EMReadScreen MAI2, 16, 15, 6
     IF MAI <> "FMPP     RUNNING" Then MsgBox "MAXIS does not appear to be running on this screen. Please have MAXIS and MMIS on the same screen to use this script."
     IF MAI <> "FMPP     RUNNING" Then StopScript
     IF MAI2 <> "MW00 1   RUNNING" Then MsgBox "MMIS does not appear to be running on this screen. Please have MAXIS and MMIS on the same screen to use this script."
     IF MAI2 <> "MW00 1   RUNNING" Then StopScript
     EMSendKey "<attn>"
     EMWaitReady 1, 0
     Call get_to_self
  End Sub

MAXIS_on

  Sub get_to_self 'This sub gets MAXIS to the main menu to enter a command.
     EMReadScreen start_position, 27, 2, 28
     If start_position <> "Select Function Menu (SELF)" Then call not_found
  End Sub

  Sub not_found 'This sub gets MAXIS from any other screen to the SELF menu using PF03. It also checks for the password prompt.
     EMSendKey "<PF3>"
     EMWaitReady 1, 100
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
     Call get_to_self
  End Sub

BeginDialog PERS, 0, 0, 156, 47, "Person search by SSN"
  ButtonGroup ButtonPressed
    OkButton 20, 25, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 60, 10, "SSN (no dashes):"
  EditBox 65, 5, 90, 15, SSN
EndDialog

Dialog PERS

If ButtonPressed = 0 then stopscript
EMSetCursor 16, 43
EMSendKey "PERS"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSetCursor 14, 36
EMSendKey SSN
EMSendKey "<enter>"
EMWaitReady 1, 0
EMReadScreen PERS_worked, 28, 2, 28
If PERS_worked <> "Person Search Display (DSPL)" then MsgBox "That didn't work. Check the error message and try again. If the SSN is right, this case might be MCRE only."
If PERS_worked <> "Person Search Display (DSPL)" then Stopscript
   row = 1
   col = 1
EMSearch " Y ", row, col
EMSetCursor row, 4
EMSendKey "x"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 21, 70
EMSendKey "curr"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMReadScreen inactive, 14, 8, 3
EMSetCursor 21, 14
EMSendKey "<PF1>"
EMWaitReady 1, 0
EMReadScreen worker_name, 19, 19, 10
If inactive = "Case: INACTIVE" then MAXIS_status = "inactive"
If inactive <> "Case: INACTIVE" then MAXIS_status = "active"

EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0

'Now the MMIS part of this script starts.

EMSendKey "<attn>" 
EMWaitReady 1, 0
EMSetCursor 2, 15
EMSendKey "10"
EMSendKey "<enter>"
EMWaitReady 1, 0

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
EMSetCursor 5, 19
EMSendKey SSN
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSetCursor 1, 08
EMSendKey "relg"
EMSendKey "<enter>"
EMWaitReady 1, 0

'Now it reads RELG to determine the current status for this SSN
EMReadScreen MMIS_active_date, 8, 7, 36
IF MMIS_active_date = "99/99/99" then MMIS_active_status = "active" 'Sets active/inactive status based on elig end date.
IF MMIS_active_date <> "99/99/99" then MMIS_active_status = "inactive"
EMReadScreen MMIS_prg, 2, 6, 10 'Reading elig program
EMReadScreen MMIS_elig_type, 2, 6, 33 'Reading elig type
EMReadScreen MMIS_case_number, 8, 6, 73 'Reading MMIS case number

'Now it goes back to RKEY, so it can get case based information.
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSetCursor 5, 19
EMSendKey "<EraseEOF>"
EMSetCursor 9, 19
EMSendKey MMIS_case_number
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSetCursor 1, 8
EMSendKey "rcin"
EMSendKey "<enter>"
EMWaitReady 1, 0

'Now it reads the MMIS worker number, then gets back to MAXIS (starting point).
EMReadScreen MMIS_worker_number, 7, 2, 46
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<attn>"
EMWaitReady 1, 0
EMSetCursor 2, 15
EMSendKey "1"
EMSendKey "<enter>"

'Now it displays the results.
MsgBox "This MAXIS case is " + MAXIS_status + ". The worker is listed as: " + worker_name + ". "& vbNewLine & + _
"MMIS shows case is " + MMIS_active_status + " on " + MMIS_prg + "/" + MMIS_elig_type + ". The case number is " + MMIS_case_number + ". " + _
"The MMIS worker number is " + MMIS_worker_number + "."& vbNewLine & + "Note: if this client is married, spousal information should be checked as well."