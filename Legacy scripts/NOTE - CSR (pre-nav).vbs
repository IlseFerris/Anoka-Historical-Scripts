EMConnect ""

BeginDialog CSR_dialog, 0, 0, 346, 292, "CSR Dialog"
  DropListBox 40, 5, 75, 15, "paper"+chr(9)+"paperless", CSR_type
  CheckBox 180, 10, 20, 10, "FS", FS_CSR
  CheckBox 210, 10, 25, 10, "HC", HC_CSR
  EditBox 40, 25, 80, 15, CSR_date
  EditBox 180, 25, 50, 15, recert_date
  EditBox 45, 45, 295, 15, HH_comp
  EditBox 65, 65, 275, 15, JOBS
  EditBox 50, 85, 290, 15, UNEA
  EditBox 50, 105, 125, 15, COEX
  EditBox 205, 105, 40, 15, CASH
  EditBox 25, 125, 315, 15, ACCT
  EditBox 30, 145, 110, 15, SECU
  EditBox 170, 145, 170, 15, CARS
  EditBox 30, 165, 75, 15, REST
  EditBox 135, 165, 100, 15, OTHR
  EditBox 60, 185, 280, 15, other_changes
  DropListBox 50, 210, 90, 15, "complete"+chr(9)+"incomplete", CSR_status
  EditBox 55, 230, 285, 15, verifs_needed
  EditBox 55, 250, 285, 15, actions_taken
  EditBox 50, 270, 50, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 230, 270, 50, 15
    CancelButton 290, 270, 50, 15
  Text 5, 110, 45, 10, "COEX/DCEX:"
  Text 5, 130, 20, 10, "ACCT:"
  Text 5, 275, 40, 10, "Your name:"
  Text 5, 50, 35, 10, "HH comp:"
  Text 5, 235, 50, 10, "Verifs needed:"
  Text 5, 70, 60, 10, "JOBS/BUSI/RBIC:"
  Text 5, 255, 50, 10, "Actions taken:"
  Text 180, 110, 25, 10, "CASH:"
  Text 5, 170, 25, 10, "REST:"
  Text 110, 170, 25, 10, "OTHR:"
  Text 120, 10, 55, 10, "Progs reviewing:"
  Text 145, 150, 25, 10, "CARS:"
  Text 5, 90, 45, 10, "UNEA/PBEN:"
  Text 5, 150, 25, 10, "SECU:"
  Text 130, 30, 50, 10, "Recert month:"
  Text 5, 190, 50, 10, "Other changes: "
  Text 5, 30, 35, 10, "CSR date:"
  Text 5, 210, 40, 10, "CSR status:"
  Text 5, 10, 35, 10, "CSR type: "
EndDialog

Dialog CSR_dialog
If buttonpressed = 0 then stopscript

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog CSR_dialog
If buttonpressed = 0 then stopscript
End Sub

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
