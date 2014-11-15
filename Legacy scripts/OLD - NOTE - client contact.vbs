EMConnect ""

BeginDialog Contact_dialog, 0, 0, 381, 147, "Client contact"
  DropListBox 50, 0, 90, 10, "Office visit"+chr(9)+"Phone call"+chr(9)+"Voicemail"+chr(9)+"Letter", contact_type
  EditBox 220, 0, 95, 15, cl_name
  EditBox 170, 20, 85, 15, case_number_field
  EditBox 100, 40, 60, 15, phone_number
  CheckBox 165, 45, 160, 10, "Do you want to case note the phone number?", phone_number_case_note
  EditBox 55, 60, 325, 15, issue
  CheckBox 5, 80, 375, 10, "Check here if all you did was leave a generic message. If any other actions were taken, fill out the next section.", left_generic_message
  EditBox 55, 95, 325, 15, actions_taken
  EditBox 80, 115, 70, 15, worker_sig
  CheckBox 5, 135, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL
  ButtonGroup ButtonPressed
    OkButton 325, 5, 50, 15
    CancelButton 325, 25, 50, 15
  Text 150, 5, 70, 10, "Who contacted you?:"
  Text 5, 45, 95, 10, "Phone number (if applicable): "
  Text 5, 100, 50, 10, "Actions taken: "
  Text 5, 5, 45, 10, "Contact type:"
  Text 5, 65, 50, 10, "Issue/subject: "
  Text 5, 120, 70, 10, "Sign your case note: "
  Text 5, 25, 165, 10, "Case number (this does not print in the case note): "
EndDialog


Dialog Contact_dialog
If buttonpressed = 0 then stopscript

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog Contact_dialog
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

EMSendKey "<home>" + Contact_type + " from " + cl_name + "<NewLine>"
If phone_number_case_note = 1 then EMSendKey "* Phone number given: " + phone_number + "<newline>"
EMSendKey "* Issue/subject: " + issue + "<newline>"
If left_generic_message = 1 then EMSendKey "* Actions taken: client did not answer. Left generic message."
If left_generic_message = 0 then EMSendKey "* Actions taken: " + actions_taken
EMSendKey "<newline>" + "---" + "<newline>" + worker_sig
If TIKL = 0 then stopscript

EMSendKey "<PF3>"
EMWaitReady 1, 0
EMReadScreen case_number, 8, 20, 38
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendkey "dail"
EMSetCursor 18, 43
EMSendkey case_number
EMSetCursor 21, 70
EMSendkey "writ" + "<enter>"
