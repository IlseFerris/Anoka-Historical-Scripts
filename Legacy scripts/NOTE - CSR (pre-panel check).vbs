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

'This Do...loop gets back to SELF
do
EMSendKey "<PF3>"
EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMSendKey "<home>" + "case" + "<eraseeof>" + case_number
EMSetCursor 21, 70
EMSendKey "curr" + "<enter>"
EMWaitReady 1, 0
EMReadScreen inactive_check, 8, 8, 9
If inactive_check = "INACTIVE" then MsgBox "This case is inactive. You should check to see if it can be REINed. Look over STAT and see if it can be REINed. If it gets to REIN status you should be able to use this script to update. If you can't REIN, process manually."
If inactive_check = "INACTIVE" then stopscript


'This Do...loop gets back to SELF
do
EMSendKey "<PF3>"
EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMSetCursor 16, 43
EMSendKey "stat" + "<enter>"
EMWaitReady 1, 0

'----<<This script may need error proofing here: it needs to know when a case is in BG>>----


BeginDialog CSR_dialog, 5, 5, 346, 292, "CSR Dialog"
  DropListBox 40, 5, 75, 15, "paper"+chr(9)+"paperless", CSR_type
  CheckBox 180, 10, 20, 10, "FS", FS_CSR
  CheckBox 210, 10, 25, 10, "HC", HC_CSR
  ButtonGroup ButtonPressed
    PushButton 5, 30, 35, 10, "CSR date:", CSR_date
  EditBox 45, 25, 75, 15, CSR_date
  EditBox 180, 25, 50, 15, recert_date
  ButtonGroup ButtonPressed
    PushButton 5, 50, 35, 10, "HH comp:", HH_comp
  EditBox 45, 45, 295, 15, HH_comp
  ButtonGroup ButtonPressed
    PushButton 5, 70, 25, 10, "JOBS/", JOBS
    PushButton 30, 70, 25, 10, "BUSI/", BUSI
    PushButton 55, 70, 20, 10, "RBIC:", RBIC
  EditBox 80, 65, 260, 15, JOBS
  ButtonGroup ButtonPressed
    PushButton 5, 90, 25, 10, "UNEA/", UNEA
    PushButton 30, 90, 25, 10, "PBEN:", PBEN
  EditBox 60, 85, 280, 15, UNEA
  ButtonGroup ButtonPressed
    PushButton 5, 110, 25, 10, "COEX/", COEX
    PushButton 30, 110, 25, 10, "DCEX:", DCEX
  EditBox 60, 105, 115, 15, COEX
  ButtonGroup ButtonPressed
    PushButton 180, 110, 25, 10, "CASH:", CASH
  EditBox 205, 105, 40, 15, CASH
  ButtonGroup ButtonPressed
    PushButton 5, 130, 25, 10, "ACCT:", ACCT
  EditBox 35, 125, 305, 15, ACCT
  ButtonGroup ButtonPressed
    PushButton 5, 150, 25, 10, "SECU:", SECU
  EditBox 35, 145, 105, 15, SECU
  ButtonGroup ButtonPressed
    PushButton 145, 150, 25, 10, "CARS:", CARS
  EditBox 175, 145, 165, 15, CARS
  ButtonGroup ButtonPressed
    PushButton 5, 170, 25, 10, "REST:", REST
  EditBox 35, 165, 70, 15, REST
  ButtonGroup ButtonPressed
    PushButton 110, 170, 25, 10, "OTHR:", OTHR
  EditBox 140, 165, 95, 15, OTHR
  EditBox 60, 185, 280, 15, other_changes
  DropListBox 50, 210, 90, 15, "complete"+chr(9)+"incomplete", CSR_status
  EditBox 55, 230, 285, 15, verifs_needed
  EditBox 55, 250, 285, 15, actions_taken
  EditBox 50, 270, 50, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 230, 270, 50, 15
    CancelButton 290, 270, 50, 15
  Text 120, 10, 55, 10, "Progs reviewing:"
  Text 5, 235, 50, 10, "Verifs needed:"
  Text 5, 190, 50, 10, "Other changes: "
  Text 5, 275, 40, 10, "Your name:"
  Text 5, 255, 50, 10, "Actions taken:"
  Text 130, 30, 50, 10, "Recert month:"
  Text 5, 210, 40, 10, "CSR status:"
  Text 5, 10, 35, 10, "CSR type: "
EndDialog

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog CSR_dialog
If buttonpressed = 0 then stopscript
End Sub


Do
Dialog CSR_dialog

EMSetCursor 20, 71
If ButtonPressed = 6 then EMSendKey "revw" + "<enter>"
If ButtonPressed = 9 then EMSendKey "memb" + "<enter>"
If ButtonPressed = 11 then EMSendKey "jobs" + "<enter>"
If ButtonPressed = 12 then EMSendKey "busi" + "<enter>"
If ButtonPressed = 13 then EMSendKey "rbic" + "<enter>"
If ButtonPressed = 15 then EMSendKey "unea" + "<enter>"
If ButtonPressed = 16 then EMSendKey "pben" + "<enter>"
If ButtonPressed = 18 then EMSendKey "coex" + "<enter>"
If ButtonPressed = 19 then EMSendKey "dcex" + "<enter>"
If ButtonPressed = 21 then EMSendKey "cash" + "<enter>"
If ButtonPressed = 23 then EMSendKey "acct" + "<enter>"
If ButtonPressed = 25 then EMSendKey "secu" + "<enter>"
If ButtonPressed = 27 then EMSendKey "cars" + "<enter>"
If ButtonPressed = 29 then EMSendKey "rest" + "<enter>"
If ButtonPressed = 31 then EMSendKey "othr" + "<enter>"
If buttonPressed = -1 then call find_case_note

Loop until ButtonPressed = -1 or ButtonPressed = 0

If buttonpressed = 0 then stopscript


Do
  find_case_note
Loop until case_note_ready = "Case Notes (NOTE)" and case_note_mode = "Mode: A" or case_note_mode = "Mode: E"

EMSendKey "<enter>"
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case note. Type your password then try again."
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog CAF_dialog
     IF buttonpressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

IF CSR_type = "paper" then EMSendKey "***CSR for " + recert_date + " received " + CSR_date + ": " + CSR_status + "***"
IF CSR_type = "paperless" then EMSendKey "***Cleared paperless HC CSR for " + recert_date + "***" + "<newline>" + "---" + "<newline>" + worker_sig
IF CSR_type = "paperless" then stopscript
EMSendKey "<NewLine>"
EMSendKey "* CSR for:"
If FS_CSR = 1 then EMSendKey " FS,"
IF HC_CSR = 1 then EMSendKey " HC,"
EMSendKey "<backspace>"
EMSendKey "<newline>"
EMSetCursor 6, 3
EMSendKey "* HH comp: " + HH_comp
EMSetCursor 9, 3
EMSendKey "* JOBS/BUSI/RBIC: " + JOBS
EMSetCursor 13, 3
EMSendKey "* UNEA/PBEN: " + UNEA
EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 0
EMSetCursor 4, 3
EMSendKey "* COEX/DCEX: " + COEX
EMSetCursor 6, 3
EMSendKey "* CASH: " + CASH
EMSetCursor 7, 3
EMSendKey "* ACCT: " + ACCT
EMSetCursor 10, 3
EMSendKey "* SECU: " + SECU
EMSetCursor 12, 3
EMSendKey "* CARS: " + CARS
EMSetCursor 14, 3
EMSendKey "* REST: " + REST
EMSetCursor 15, 3
EMSendKey "* OTHR: " + OTHR

EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 0







'The following is the last page of the case note.
EMSetCursor 4, 3
EMSendKey "* Other changes: " + other_changes
EMSetCursor 7, 3
EMSendKey "* Verifs needed: " + verifs_needed
EMSetCursor 11, 3
EMSendKey "* Actions taken: " + actions_taken
EMSetCursor 16, 3
EMSendKey "---"
EMSetCursor 17, 3
EMSendKey worker_sig
'Now it goes through, deletes the carats and exits the case note.
EMSendKey "<PF7>"
EMWaitReady 1, 0
EMSetCursor 17, 3
EMSendKey " " + "<PF7>"
EMWaitReady 1, 0
EMSetCursor 17, 3
EMSendKey " " + "<PF3>"
EMWaitReady 1, 0
'Now it reenters the case note.
EMSetCursor 5, 3
EMSendKey "x" + "<enter>"


stopscript