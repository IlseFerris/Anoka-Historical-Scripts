EMConnect ""

BeginDialog closed_dialog, 0, 0, 366, 127, "Closed case dialog"
  EditBox 65, 5, 55, 15, date_of_closure
  CheckBox 160, 10, 150, 10, "If you closed the whole case, check here.", whole_case_closed_check
  EditBox 55, 25, 75, 15, progs_closed
  EditBox 210, 25, 150, 15, reason_for_closure
  EditBox 110, 45, 250, 15, verifs_needed
  EditBox 270, 65, 90, 15, REIN_date
  EditBox 170, 85, 190, 15, open_progs
  EditBox 80, 105, 70, 15, worker_sig
  ButtonGroup closed_dialog_ButtonPressed
    OkButton 205, 105, 50, 15
    CancelButton 260, 105, 50, 15
  Text 5, 10, 55, 10, "Date of closure:"
  Text 5, 30, 50, 10, "Progs closed:"
  Text 140, 30, 65, 10, "Reason for closure:"
  Text 5, 50, 100, 10, "Verifs needed (if applicable):"
  Text 5, 70, 265, 10, "Last possible REIN date? If not possible, state that client would have to reapply:"
  Text 5, 90, 165, 10, "Are any programs still open? If so, list them here:"
  Text 5, 110, 70, 10, "Sign your case note: "
EndDialog

Do
  all_sections_completed = "" 'Resetting variable
  Dialog closed_dialog
  If closed_dialog_ButtonPressed = 0 then stopscript
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen case_note_ready, 17, 2, 33
  EMReadScreen case_note_mode, 7, 20, 3
  If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
  If date_of_closure = "" or progs_closed = "" or reason_for_closure = "" or REIN_date = "" or worker_sig = "" then all_sections_completed = "False"
  If date_of_closure = "" or progs_closed = "" or reason_for_closure = "" or REIN_date = "" or worker_sig = "" then MsgBox "You must fill in the date of closure, programs closed, the reason for closure, the REIN date, and the signature, in order to proceed with this script."
Loop until case_note_ready = "Case Notes (NOTE)" and (case_note_mode = "Mode: A" or case_note_mode = "Mode: E") and all_sections_completed <> "False"

If whole_case_closed_check = 1 then EMSendKey "<home>" + "---Closed case for " + date_of_closure + "---" + "<NewLine>"
If whole_case_closed_check = 0 then EMSendKey "<home>" + "---Closed " + progs_closed + " for " + date_of_closure + "---" + "<NewLine>"
If whole_case_closed_check = 1 then EMSendKey "* Progs closed: " + progs_closed + "<NewLine>"
EMSendKey "* Reason: " + reason_for_closure + "<newline>"
If verifs_needed <> "" then EMSendKey "* Verifs needed: " + verifs_needed + "<newline>"
EMSendKey "* Last possible REIN date: " + REIN_date + "<newline>"
If open_progs <> "" then EMSendKey "* Case remains open on: " + open_progs + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig
