EMConnect ""

BeginDialog HC_renewal_dialog, 0, 0, 346, 297, "HC renewal dialog"
  EditBox 70, 5, 80, 15, recert_datestamp
  EditBox 210, 5, 50, 15, recert_date
  EditBox 45, 25, 295, 15, HH_comp
  EditBox 65, 45, 275, 15, JOBS
  EditBox 50, 65, 290, 15, UNEA
  EditBox 50, 85, 125, 15, COEX
  EditBox 205, 85, 40, 15, CASH
  EditBox 25, 105, 315, 15, ACCT
  EditBox 30, 125, 110, 15, SECU
  EditBox 170, 125, 170, 15, CARS
  EditBox 30, 145, 75, 15, REST
  EditBox 135, 145, 100, 15, OTHR
  EditBox 260, 145, 80, 15, DISA
  EditBox 60, 165, 280, 15, other_changes
  DropListBox 60, 185, 90, 15, "complete"+chr(9)+"incomplete", Recert_status
  EditBox 55, 205, 285, 15, verifs_needed
  EditBox 55, 225, 285, 15, actions_taken
  CheckBox 5, 245, 100, 10, "MA-EPD? If so, check here", MAEPD_check
  EditBox 60, 270, 90, 15, MAEPD_premium
  CheckBox 155, 275, 65, 10, "Emailed MADE?", MADE
  EditBox 290, 245, 50, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 230, 275, 50, 15
    CancelButton 290, 275, 50, 15
  Text 5, 210, 50, 10, "Verifs needed:"
  Text 5, 50, 60, 10, "JOBS/BUSI/RBIC:"
  Text 5, 230, 50, 10, "Actions taken:"
  Text 180, 90, 25, 10, "CASH:"
  Text 5, 150, 25, 10, "REST:"
  Text 110, 150, 25, 10, "OTHR:"
  Text 145, 130, 25, 10, "CARS:"
  Text 5, 70, 45, 10, "UNEA/PBEN:"
  Text 5, 130, 25, 10, "SECU:"
  Text 160, 10, 50, 10, "Recert month:"
  Text 5, 170, 50, 10, "Other changes: "
  Text 5, 10, 60, 10, "Recert datestamp:"
  Text 5, 190, 50, 10, "Recert status:"
  Text 240, 150, 20, 10, "DISA: "
  Text 5, 110, 20, 10, "ACCT:"
  GroupBox 5, 260, 220, 30, "If MA-EPD..."
  Text 10, 275, 50, 10, "New premium:"
  Text 5, 30, 35, 10, "HH comp:"
  Text 5, 90, 45, 10, "COEX/DCEX:"
  Text 245, 250, 40, 10, "Your name:"
EndDialog

Dialog HC_renewal_dialog

If buttonpressed = 0 then stopscript

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog HC_renewal_dialog
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

EMSendKey "***HC ER for " + recert_date + " received " + recert_datestamp + ": " + recert_status + "***"
EMSendKey "<NewLine>"
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







'The following is the second to the last page of the case note.
EMSetCursor 4, 3
EMSendKey "* DISA: " + DISA
EMSetCursor 6, 3
EMSendKey "* Other changes: " + other_changes
EMSetCursor 8, 3
EMSendKey "* Verifs needed: " + verifs_needed
EMSetCursor 12, 3
EMSendKey "* Actions taken: " + actions_taken
EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 0

'The following is the last page of the case note.
If MAEPD_check = 1 then EMSendKey "* MA-EPD premium: " + MAEPD_premium + ". "
IF MADE = 1 then EMSendKey "Emailed MADE."
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
EMSendKey " " + "<PF7>"
EMWaitReady 1, 0
EMSetCursor 17, 3
EMSendKey " " + "<PF3>"
EMWaitReady 1, 0
'Now it reenters the case note.
EMSetCursor 5, 3
EMSendKey "x" + "<enter>"
