'-----------------------------------------------------------------------------------------------------------------------------------------------------
'        Script name:    Internal XFER (tentative)
'        Description:    SPEC/XFERs a case to another worker in the county. (tentative)
'       Target users:    FAS and OSAs
'           Division:    Adult
'          Author(s):    Ronny Cary
'      Working state:    Alpha
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'     Script content:    01. A dialog pops up allowing the worker to enter information about the case to be transferred, and select the worker to
'                            transfer the case to.
'                        02. The script determines the short name and x102 numbers for the worker which was selected.
'                        03. The script navigates to MAXIS production (or training).
'                        04. The script navigates to CASE/CURR to determine what programs are open.
'                        05. The script navigates to CASE/NOTE and notes the transfer and information about the transfer.
'                        06. The script navigates to SPEC/XFER and transfers the case to the selected worker.
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'       Known issues:    1. All of the possible workers are not worked in to this script.
'                        2. All of the possible appointment types are not worked in to this script.
'                        3. Should transfer MCRE cases as well, and case note the MCRE case number.
'   Test breakpoints:    None
'              Notes:    None
'-----------------------------------------------------------------------------------------------------------------------------------------------------

'SECTION 01

EMConnect ""

row = 1
col = 1

EMSearch "Case Nbr: ", row, col

EMReadScreen case_number, 8, row, col + 10
If case_number = "AR" or case_number = "________" then case_number = ""

BeginDialog worker_list_dialog, 0, 0, 216, 116, "Worker List"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 5, 25, 105, 10, "Type of XFER in Appointments:"
  DropListBox 115, 20, 95, 15, "EAA - to Gen"+chr(9)+"EAA - to HC (MCRE only paper process)", transfer_type
  Text 5, 45, 65, 10, "Notes on the XFER:"
  EditBox 70, 40, 140, 15, notes
  Text 5, 60, 45, 10, "New worker:"
  DropListBox 55, 60, 70, 15, "Aadland, Darcy"+chr(9)+"Almquist, Netti"+chr(9)+"Cary, Ronald"+chr(9)+"Ferrazzi, Rene"+chr(9)+"Mayer, Randi", worker_ID
  Text 5, 80, 65, 10, "Sign the case note:"
  EditBox 75, 75, 100, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 45, 95, 50, 15
    CancelButton 100, 95, 50, 15
EndDialog





Dialog worker_list_dialog
If Buttonpressed = 0 then stopscript

'SECTION 02

If worker_ID = "Aadland, Darcy" then x102_number = "x102c02"
If worker_ID = "Almquist, Netti" then x102_number = "x102b64"
If worker_ID = "Cary, Ronald" then x102_number =  "x102b82"
If worker_ID = "Ferrazzi, Rene" then x102_number =  "x102880"
If worker_ID = "Mayer, Randi" then x102_number =  "x102268"

short_name_array = Split(worker_ID, ", ")
short_name = short_name_array(1) & " " & left(short_name_array(0), 1) & "."

'SECTION 03

'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMSendKey "<attn>"
EMWaitReady 1, 1
EMReadScreen MAI_check, 3, 1, 33
If MAI_check <> "MAI" then EMWaitReady 1, 5
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If training_check = "RUNNING" then 
  EMSendKey "3" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" then EMWaitReady 1, 5
End if
If production_check = "RUNNING" then 
  EMSendKey "1" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" then EMWaitReady 1, 5
End if

EMSendKey "<enter>" 'It sends this now in order to force a screen refresh, to check for any password prompts
EMWaitReady 1, 1

Do
  EMReadScreen password_prompt, 38, 2, 23
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

'SECTION 04


'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWriteScreen "case", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "curr", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

row = 1
col = 1
EMSearch "Case: ACTIVE", row, col
If row = 0 then active_programs = "MAXIS is inactive"
If row <> 0 then
  col = col + 5 'To get the script to not see "Case: ACTIVE" as the next active thing.
  Do
    col = col + 1 'To get the script to look beyong the current column in determining what's active.
    EMSearch ": ACTIVE", row, col
    If row <> 0 then 
      EMReadScreen found_program, 4, row, col - 4
      If active_programs <> "" then active_programs = active_programs & ", " & trim(found_program)
      If active_programs = "" then active_programs = trim(found_program)
    End if
  Loop until row = 0
End if

'SECTION 05

EMWriteScreen "note", 20, 69
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<PF9>"
EMWaitReady 1, 1

If transfer_type = "EAA - to Gen" then EMSendKey ">>>CASE TRANSFER TO ONGOING WORKER<<<" & "<newline>"
If transfer_type = "EAA - to HC (MCRE only paper process)" then EMSendKey ">>>MCRE-ONLY CASE TRANSFER TO HC WORKER<<<" & "<newline>"
EMSendKey "* Worker receiving case: " & short_name & ", " & x102_number & "<newline>"
EMSendKey "* Notes on transfer: " & notes & "<newline>"
EMSendKey "* Active programs: " & active_programs & "<newline>"
EMSendKey "---" & "<newline>"
EMSendKey worker_sig & ", using automated script."


'SECTION 06

'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWriteScreen "spec", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "xfer", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

EMWriteScreen "x", 7, 16
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<PF9>"
EMWaitReady 1, 1

EMWriteScreen x102_number, 18, 61

EMSendKey "<enter>"
EMWaitReady 1, 1