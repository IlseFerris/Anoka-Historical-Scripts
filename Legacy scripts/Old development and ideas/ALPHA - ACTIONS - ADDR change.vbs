EMConnect ""


  row = 1
  col = 1

EMSearch "Case Nbr: ", row, col

EMReadScreen case_number, 8, row, col + 10
If case_number = "AR" then case_number = ""

BeginDialog case_number_dialog, 0, 0, 161, 42, "Case number"
  EditBox 95, 0, 60, 15, case_number
  ButtonGroup ButtonPressed_case_number
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

Dialog case_number_dialog

If ButtonPressed_case_number = 0 then stopscript








'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 0



'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog Dialog1
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"




'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMSendKey "<attn>"
EMWaitReady 1, 0
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
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

'Now it will check STAT/ADDR to show the old address.
EMWaitReady 1, 0
EMSendKey "<home>" + "stat" + "<eraseeof>" + case_number
EMSetCursor 21, 70
EMSendKey "addr" + "<enter>"
EMWaitReady 1, 0

'If there was an error after trying to go to STAT/ADDR, the script will shut down.
    EMReadScreen SELF_error_check, 27, 2, 28
    If SELF_error_check = "Select Function Menu (SELF)" then stopscript

'If the case is already in another county, the script will shut down.
    EMReadScreen county_check, 4, 21, 21
    If county_check <> "X102" then MsgBox "This case is in another county. Forward mail to the appropriate address."
    If county_check <> "X102" then stopscript	

BeginDialog addr_dialog, 0, 0, 221, 157, "Address Dialog"
  ButtonGroup addr_dialog_ButtonPressed
    OkButton 110, 135, 50, 15
    CancelButton 165, 135, 50, 15
  CheckBox 5, 5, 205, 10, "Check here if the address changed in the first of the month. ", first_of_month_check
  Text 5, 20, 155, 15, "Otherwise: type the address as you would in MAXIS, without dashes or spaces (MMDDYY):"
  EditBox 165, 20, 50, 15, address_date
  Text 5, 45, 60, 10, "Street and apt #:"
  EditBox 65, 40, 150, 15, street_listing
  Text 5, 65, 40, 10, "City/St/Zip:"
  EditBox 45, 60, 90, 15, city
  EditBox 145, 60, 25, 15, state
  EditBox 180, 60, 35, 15, zip_code
  Text 5, 85, 125, 10, "County code (leave blank if unknown):"
  Text 5, 105, 65, 10, "Verification status:"
  DropListBox 75, 100, 115, 15, "Needs verification"+chr(9)+"No proof needed"+chr(9)+"Proof received", verification_status
  Text 5, 120, 140, 10, "If proofs have been received, which type:"
  DropListBox 150, 115, 65, 15, "Shelter Form"+chr(9)+"Coltrl Stmt"+chr(9)+"Lease/Rent Doc"+chr(9)+"Mortgage Papers"+chr(9)+"Prop Tax Stmt"+chr(9)+"Ctrct for Deed"+chr(9)+"Utility Stmt"+chr(9)+"Dvr Lic/St ID"+chr(9)+"Other Document"+chr(9)+"No Ver Prvd", proof_type
  CheckBox 165, 85, 50, 10, "Homeless?", homeless_check
  EditBox 135, 80, 20, 15, county
EndDialog


Dialog addr_dialog

Do
If addr_dialog_ButtonPressed = 0 then stopscript
EMReadScreen addr_check, 4, 2, 44
If addr_check <> "ADDR" then MsgBox "Navigate back to ADDR before pressing OK."
If addr_check <> "ADDR" then Dialog addr_dialog
Loop until addr_check = "ADDR"


EMSendKey "<PF9>"
EMWaitReady 1, 0

EMReadScreen footer_month, 2, 20, 55
EMReadScreen footer_year, 2, 20, 58
If first_of_month_check = 1 then EMWriteScreen footer_month, 4, 43
If first_of_month_check = 1 then EMWriteScreen "01", 4, 46
If first_of_month_check = 1 then EMWriteScreen footer_year, 4, 49
EMSetCursor 7, 43
EMSendKey "<eraseeof>"
EMSetCursor 6, 43
EMSendKey "<eraseeof>" + street_listing
EMSetCursor 8, 43
EMSendKey "<eraseeof>" + city
EMSetCursor 8, 66
EMSendKey state
EMSetCursor 9, 43
EMSendKey zip_code
EMSetCursor 9, 66
EMSendKey county
If homeless_check = 1 then EMWriteScreen "y", 10, 43
If homeless_check = 0 then EMWriteScreen "n", 10, 43

stopscript

'Now the script goes into the case note and case notes the action. It will stop if the forwarding address matches MAXIS.
EMSendKey "<pf4>"
EMWaitReady 1, 0
EMSendKey "<PF9>"
EMWaitReady 1, 0
If ButtonPressed2 = 7 then EMSendKey "Returned mail received, new address known" + "<newline>" + "* Forwarding address matches MAXIS, forwarding to new address." + "<newline>" + "---" + "<newline>" + worker_sig
If ButtonPressed2 = 7 then stopscript
EMSendKey "-->Returned mail received<--" + "<newline>"
If forwarding_check = 1 then EMSendKey "* Forwarding address indicated as: " + forwarding_address + "<newline>"
If forwarding_check = 1 then EMSendKey "* Sending verification request to forwarding address. TIKLed for 10-day return." + "<newline>"
If forwarding_check = 0 then EMSendKey "* No forwarding address was indicated." + "<newline>"
If forwarding_check = 0 then EMSendKey "* Sending verification request to last known address. TIKLed for 10-day return." + "<newline>"
EMSendKey "---" + "<newline>" + worker_sig + "<pf3>"

'Now we go back to SELF, in order to TIKL out.
EMWaitReady 1, 0
EMSendKey "<pf3>"
EMWaitReady 1, 0
EMSendKey "<pf3>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendkey "dail"
EMSetCursor 18, 43
EMSendkey case_number
EMSetCursor 21, 70
EMSendkey "writ" + "<enter>"

'The following will generate a TIKL formatted date for 10 days from now.

If DatePart("d", Now + 10) = 1 then TIKL_day = "01"
If DatePart("d", Now + 10) = 2 then TIKL_day = "02"
If DatePart("d", Now + 10) = 3 then TIKL_day = "03"
If DatePart("d", Now + 10) = 4 then TIKL_day = "04"
If DatePart("d", Now + 10) = 5 then TIKL_day = "05"
If DatePart("d", Now + 10) = 6 then TIKL_day = "06"
If DatePart("d", Now + 10) = 7 then TIKL_day = "07"
If DatePart("d", Now + 10) = 8 then TIKL_day = "08"
If DatePart("d", Now + 10) = 9 then TIKL_day = "09"
If DatePart("d", Now + 10) > 9 then TIKL_day = DatePart("d", Now + 10)

If DatePart("m", Now + 10) = 1 then TIKL_month = "01"
If DatePart("m", Now + 10) = 2 then TIKL_month = "02"
If DatePart("m", Now + 10) = 3 then TIKL_month = "03"
If DatePart("m", Now + 10) = 4 then TIKL_month = "04"
If DatePart("m", Now + 10) = 5 then TIKL_month = "05"
If DatePart("m", Now + 10) = 6 then TIKL_month = "06"
If DatePart("m", Now + 10) = 7 then TIKL_month = "07"
If DatePart("m", Now + 10) = 8 then TIKL_month = "08"
If DatePart("m", Now + 10) = 9 then TIKL_month = "09"
If DatePart("m", Now + 10) > 9 then TIKL_month = DatePart("m", Now + 10)

TIKL_year = DatePart("yyyy", Now + 10)

EMWaitReady 1, 0
EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year - 2000
EMSetCursor 9, 3
EMSendKey "Request for address sent 10 days ago. If not responded to, close per returned mail procedure." + "<enter>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
If forwarding_check = 0 then MsgBox "Use the returned mail packet in Compass Forms. Send the completed forms to the most recent address. The script has TIKLed out for you for 10-day return."
If forwarding_check = 1 then MsgBox "Use the returned mail packet in Compass Forms. Send the completed forms to the forwarding address. The script has TIKLed out for you for 10-day return. Do not update STAT/ADDR until you receive a response from the client."

