'FUNCTIONS----------------------------------------------------------------------------------------------------

Function multiple_panel_finder 'Reads out what the current_panel and total_panels are. Needs to declare these variables before running the function
  EMReadScreen panel_amt_check, 8, 2, 72
  current_panel = trim(left(panel_amt_check, instr(panel_amt_check, " Of ")))
  total_panels = trim(right(panel_amt_check, instrrev(panel_amt_check, " Of ")))
End function
dim current_panel
dim total_panels

Function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
End function

Function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
End function

Function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

EMConnect ""

'DEV BITS THAT'LL NEED TO BE REMOVED----------------------------------------------------------------------------------------------------

'CASE NUMBER AND SSN NEEDS TO BE FLESHED OUT FROM DAIL, RIGHT NOW IT'S MANUAL ENTRY<<<<<<<<<<<<<<<<<<<<
case_number = "1071328"
client_SSN = replace("399 20 7774", " ", "")

'ALSO NEEDS TO GET FOOTER MONTH FROM DAIL, RIGHT NOW IT'S MANUAL ENTRY<<<<<<<<<<<<<<<<<<<<<<<<<
footer_month = "02"
footer_year = "13"

'IT'LL NEED TO NAVIGATE TO THE INFC SCREEN FROM THE DAIL, RIGHT NOW IT'LL JUST DO IT MANUALLY (NO DAIL MESSAGE FOR TESTING)
Do
  PF3
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"
EMWriteScreen "infc", 16, 43
transmit
EMWriteScreen "sves", 20, 71
transmit
EMWriteScreen "________", 5, 68
EMWriteScreen client_SSN, 4, 68
EMWriteScreen "tpqy", 20, 70
transmit

'AS WE AREN'T USING AN ACTUAL TPQY RESPONSE, THIS WILL GRAB INFO THAT THE EXISTING SCRIPT SHOULD ALSO BE GRABBING----------------------------------------------------------------------------------------------------

'Transmits past the intro page
transmit

'Grabs RSDI claim number and gross amount, converts RSDI_gross to a number for use in calculation. Also grabs SSN to identify an SSN with a HH memb number.
EMReadScreen RSDI_claim_number, 12, 5, 40
EMReadScreen RSDI_gross, 7, 8, 16
RSDI_gross = abs(trim(RSDI_gross))
EMReadScreen BDXM_SSN, 11, 5, 19

'Removes "00" from the end of an RSDI claim number per policy
If right(RSDI_claim_number, 2) = "00" then RSDI_claim_number = left(RSDI_claim_number, 10)

'THE ADDED PARTS OF THE SCRIPT----------------------------------------------------------------------------------------------------

'Navigates back to SELF
Do
  PF3
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

'Navigates to STAT/MEMB for the footer month of the message
EMWriteScreen "stat", 16, 43
EMWriteScreen "        ", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "memb", 21, 70
transmit

'Looking through HH memb to assign a member number to the SSN. If one is not found the script will close.
Do
  EMReadScreen memb_SSN, 11, 7, 42
  If memb_SSN = BDXM_SSN then exit do
  multiple_panel_finder
  if current_panel = total_panels then
    MsgBox "A HH member could not be matched to this TPQY response."
    StopScript
  Else
    transmit
  End if
Loop until memb_SSN = BDXM_SSN
EMReadScreen memb_number, 2, 4, 33
EMWriteScreen "UNEA", 20, 71
EMWriteScreen memb_number, 20, 76
transmit

'Determines which panel is which using the multiple_panel_finder function and the UNEA claim number.
Do
  EMReadScreen UNEA_claim_number, 15, 6, 37
  UNEA_claim_number = trim(replace(UNEA_claim_number, "_", ""))
  If UNEA_claim_number = RSDI_claim_number then exit do
  multiple_panel_finder
  if current_panel = total_panels then
    MsgBox "A UNEA panel could not be matched to this TPQY response. Check to make sure you don't have a ''00'' suffix on your claim number. If you have a ''00'' suffix, remove it from UNEA, send your case through background and try again."
    StopScript
  Else
    transmit
  End if
Loop until UNEA_claim_number = RSDI_claim_number

'Reads the gross amount, and compares it with the gross amount from the response. If they don't match the script will ask the worker if they'd like it to update.
EMReadScreen UNEA_gross, 8, 18, 68
UNEA_gross = abs(trim(replace(UNEA_gross, "_", "")))
If UNEA_gross = RSDI_gross then
  MsgBox "No change to RSDI amount."
Else
  RSDI_update_box = MsgBox ("Not matched. New amount listed at: " & RSDI_gross & ". Would you like the script to update this panel?", 3) 
  MsgBox RSDI_update_box
  If RSDI_update_box = 2 then Stopscript
  If RSDI_update_box = 6 then 
    PF9
    EMWriteScreen "________", 13, 68
    EMWriteScreen RSDI_gross, 13, 68
    transmit
    transmit
  End if
End if