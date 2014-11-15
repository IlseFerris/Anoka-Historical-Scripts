EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 0

'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then press ''OK'' to try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

'Now it will look for MMIS. If MMIS is running, it will start checking each case against MMIS.
EMSendKey "<attn>"
EMWaitReady 1, 0

'The following checks for which screen MMIS is running on.
EMReadScreen MMIS_A_check, 7, 16, 15 'This should be row 15 (the middle of the three coordinates!) for production and row 16 for training.
IF MMIS_A_check = "RUNNING" then EMSendKey "11" + "<enter>" 'This should be 10 for production and 11 for training.
IF MMIS_A_check = "RUNNING" then EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMSendKey "<attn>"
IF MMIS_A_check <> "RUNNING" then EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMConnect "B"
EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMSendKey "<attn>"
EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMReadScreen MMIS_B_check, 7, 16, 15 'This should be row 15 for production and row 16 for training.
If MMIS_A_check <> "RUNNING" and MMIS_B_check <> "RUNNING" then MsgBox "MMIS does not appear to be running. This script will now stop."
If MMIS_A_check <> "RUNNING" and MMIS_B_check <> "RUNNING" then stopscript
IF MMIS_A_check <> "RUNNING" and MMIS_B_check = "RUNNING" then EMSendkey "11" + "<enter>" 'This should be 10 for production and 11 for training.

EMFocus 'To bring focus to whatever screen is now running.

'Now we use a Do Loop to get to the start screen for MMIS.
Sub get_to_session_begin
  Do 
  EMSendkey "<PF6>"
  EMReadScreen password_prompt2, 38, 2, 23
  IF password_prompt2 = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
  EMWaitReady 1, 0
  EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
End Sub
get_to_session_begin

'Now we get back into MMIS. We have to skip past the intro screens.
EMSetCursor 1, 2
EMSendKey "mw00"
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<enter>"
EMWaitReady 1, 1

'This section may not work for all OSAs, since some only have EK01. This will find EK01 and enter it.
  row = 1
  col = 1
EMSearch "EK01", row, col
If row <> 0 then EMSetCursor row, 4
If row <> 0 then EMSendKey "x"
If row <> 0 then EMSendKey "<enter>"
If row <> 0 then EMWaitReady 1, 1

'This section starts from EK01. OSAs may need to skip the previous section.
EMSetCursor 12, 3
EMSendKey "x"
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now we're in OKEY and the script will begin making a new NPI number.
EMWriteScreen "a", 2, 65
EMWriteScreen "x", 10, 18
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now we're in the next screen, and the worker has to enter a PMI number.
  BeginDialog PMI_dialog, 0, 0, 216, 42, "PMI dialog"
    Text 5, 5, 145, 10, "Enter your PMI number:"
    EditBox 155, 0, 60, 15, PMI_number
    ButtonGroup PMI_dialog_ButtonPressed
      OkButton 50, 20, 50, 15
      CancelButton 110, 20, 50, 15
  EndDialog
Do 'This Do...Loop gets the PMI number, and will make sure the PMI length is 8 digits long.
  Dialog PMI_dialog
  If PMI_dialog_ButtonPressed = 0 then stopscript
  Do
    If len(PMI_number) < 8 then PMI_number = "0" & PMI_number
    If len(PMI_number) >= 8 then exit Do
  Loop until len(PMI_number) = 8
  If len(PMI_number) <> 8 then MsgBox "This is not eight digits long. The PMI needs to be eight digits long."
Loop until len(PMI_number) = 8
EMWriteScreen PMI_number, 6, 18
EMWriteScreen "x", 18, 28
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now we're on the next screen, and we have to grab the info for MMIS. The script will allow a worker to manually add an address if it is different.
EMReadScreen OPRV_check, 4, 1, 52
If OPRV_check <> "OPRV" then MsgBox "You appear to have left the OPRV screen. Do not navigate away from this screen when entering this info. The script will now stop."
If OPRV_check <> "OPRV" then stopscript
EMReadScreen NPI_ID, 10, 3, 10
EMReadScreen ADDR_line_01, 21, 10, 11
If ADDR_line_01 = "                     " then ADDR_line_01 = ""
EMReadScreen ADDR_line_02, 21, 11, 11
If ADDR_line_02 = "                     " then ADDR_line_02 = ""
EMReadScreen city_line, 12, 12, 11
If city_line = "            " then city_line = ""
EMReadScreen state_line, 2, 12, 39
If state_line = "  " then state_line = ""
EMReadScreen zip_code_line, 5, 12, 54
If zip_code_line = "     " then zip_code_line = ""
EMReadScreen phone_number_line, 12, 13, 11
phone_number_line = replace(phone_number_line, "-", "") 'Because the OKEY screen shows the dashes but RCAD does not take them.
If phone_number_line = "            " then phone_number_line = ""
  BeginDialog ADDR_dialog, 0, 0, 191, 122, "ADDR dialog"
    Text 5, 5, 55, 10, "Address line 1: "
    EditBox 65, 0, 115, 15, ADDR_line_01
    Text 5, 25, 55, 10, "Address line 2: "
    EditBox 65, 20, 115, 15, ADDR_line_02
    Text 5, 45, 20, 10, "City:"
    EditBox 25, 40, 90, 15, city_line
    Text 120, 45, 25, 10, "State:"
    EditBox 150, 40, 30, 15, state_line
    Text 5, 65, 45, 10, "Zip, 5 digits:"
    EditBox 50, 60, 70, 15, zip_code_line
    Text 5, 85, 120, 10, "Phone (no dashes, ie 5555555555):"
    EditBox 125, 80, 55, 15, phone_number_line
    ButtonGroup ADDR_dialog_ButtonPressed
      OkButton 40, 100, 50, 15
      CancelButton 100, 100, 50, 15
  EndDialog
Do
  Dialog ADDR_dialog
  If ADDR_dialog_ButtonPressed = 0 then stopscript
  EMReadScreen OPRV_check, 4, 1, 52
  If OPRV_check <> "OPRV" then MsgBox "You are not in MMIS financial control (the OPRV screen). The script will not continue until you navigate back to that screen!"
  If ADDR_line_01 = "" or city_line = "" or state_line = "" or zip_code_line = "" then msgbox "You can't leave the address lines blank. MMIS needs this info to continue."
Loop until OPRV_check = "OPRV" and (ADDR_line_01 <> "" and city_line <> "" and state_line <> "" and zip_code_line <> "")
EMSendKey "<PF3>"
EMWaitReady 1, 1

'Now we navigate to RKEY.
EMWriteScreen "RKEY", 1, 8
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now we add a new MCRE case.
EMWriteScreen "a", 2, 19
EMWriteScreen "x", 9, 69
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now we update the address of the client.
EMWriteScreen "102", 2, 53
EMWriteScreen ADDR_line_01, 6, 17
EMWriteScreen ADDR_line_02, 7, 17
EMWriteScreen city_line, 8, 17
EMWriteScreen state_line, 9, 17
EMWriteScreen zip_code_line, 9, 30
EMWriteScreen "002", 10, 17
EMWriteScreen phone_number_line, 10, 30
EMSendKey "<enter>"
EMWaitReady 1, 1

'This screen is not used, so we skip past it.
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now we insert the NPI and PMI of the client.
EMWriteScreen NPI_ID, 7, 8
EMWriteScreen PMI_number, 11, 4
EMWriteScreen "01", 11, 13

'This is where the script stops for now. Workers will manually have to create the new case until this script is further developed.

Stopscript