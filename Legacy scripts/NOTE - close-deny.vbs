EMConnect ""

BeginDialog Closed_denied_dialog, 0, 0, 366, 187, "Closed-denied case"
  DropListBox 65, 5, 90, 10, "Closed"+chr(9)+"Denied", closed_or_denied
  CheckBox 165, 10, 180, 10, "If you closed (or denied) the whole case, check here.", whole_case_closed_check
  EditBox 60, 25, 70, 15, date_of_closure
  EditBox 220, 25, 140, 15, progs_closed
  EditBox 75, 45, 285, 15, reason_for_closure
  CheckBox 5, 70, 155, 10, "Are verifs needed? If so, check here and list: ", verifs_needed_check
  EditBox 165, 65, 195, 15, verifs_needed
  EditBox 270, 85, 90, 15, REIN_date
  EditBox 170, 105, 190, 15, open_progs
  EditBox 80, 125, 70, 15, worker_sig
  CheckBox 5, 150, 235, 10, "Check here to TIKL out 10 days to send case to CLS (HC-only denial)", TIKL_check
  ButtonGroup ButtonPressed
    OkButton 255, 165, 50, 15
    CancelButton 310, 165, 50, 15
  Text 5, 50, 65, 10, "Reason for closure:"
  Text 5, 130, 70, 10, "Sign your case note: "
  Text 5, 90, 265, 10, "Last possible REIN date? If not possible, state that client would have to reapply:"
  Text 5, 10, 60, 10, "Closed or denied:"
  Text 5, 30, 55, 10, "Date of closure:"
  Text 140, 30, 80, 10, "What programs closed:"
  Text 5, 110, 165, 10, "Are any programs still open? If so, list them here:"
EndDialog




Dialog Closed_denied_dialog
If buttonpressed = 0 then stopscript

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog Closed_denied_dialog
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

If whole_case_closed_check = 1 then EMSendKey "<home>" + "---" + closed_or_denied + " case for " + date_of_closure + "---" + "<NewLine>"
If whole_case_closed_check = 0 then EMSendKey "<home>" + "---" + closed_or_denied + " " + progs_closed + " for " + date_of_closure + "---" + "<NewLine>"
If whole_case_closed_check = 1 then EMSendKey "* Progs closed: " + progs_closed + "<NewLine>"
EMSendKey "* Reason: " + reason_for_closure + "<newline>"
If verifs_needed_check = 1 then EMSendKey "* Verifs needed: " + verifs_needed + "<newline>"
EMSendKey "* Last possible REIN date: " + REIN_date + "<newline>"
If whole_case_closed_check = 0 then EMSendKey "* Case remains open on: " + open_progs + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig
   row = 1
   col = 1
EMSearch " for ", row, col
EMSetCursor 4, col
If closed_or_denied = "Denied" then EMSendKey "---" + "<eraseeof>"
If TIKL_check = 0 then stopscript
EMReadScreen case_number, 8, 20, 38

'This Do...loop gets back to SELF
do
EMSendKey "<PF3>"
EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

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
EMSendKey "This case was denied 10 days ago. If client has not reapplied, or turned in proofs, and is not open, send to CLS per policy." + "<enter>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
