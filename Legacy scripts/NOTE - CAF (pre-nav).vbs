EMConnect ""

BeginDialog CAF_dialog, 0, 0, 346, 432, "CAF Dialog"
  DropListBox 40, 5, 95, 15, "Intake"+chr(9)+"Reapplication"+chr(9)+"Recertification"+chr(9)+"Add program", CAF_type
  CheckBox 180, 10, 20, 10, "FS", FS_app
  CheckBox 210, 10, 25, 10, "HC", HC_app
  CheckBox 240, 10, 30, 10, "Cash", cash_app
  CheckBox 275, 10, 30, 10, "Emer", emer_app
  EditBox 40, 25, 80, 15, CAF_date
  EditBox 220, 25, 50, 15, recert_date
  EditBox 45, 60, 295, 15, HH_comp
  EditBox 25, 80, 140, 15, DISA
  EditBox 235, 80, 105, 15, SCHL
  EditBox 50, 100, 155, 15, AREP
  EditBox 65, 135, 275, 15, JOBS
  EditBox 50, 155, 290, 15, UNEA
  EditBox 80, 175, 260, 15, income_changes
  EditBox 50, 195, 115, 15, SHEL
  EditBox 215, 195, 125, 15, COEX
  EditBox 30, 230, 40, 15, CASH
  EditBox 95, 230, 245, 15, ACCT
  EditBox 30, 250, 110, 15, SECU
  EditBox 170, 250, 170, 15, CARS
  EditBox 30, 270, 75, 15, REST
  EditBox 135, 270, 100, 15, OTHR
  EditBox 260, 270, 80, 15, TRAN
  EditBox 45, 305, 140, 15, MEDI
  EditBox 220, 305, 120, 15, DIET
  EditBox 50, 325, 85, 15, FMED
  EditBox 190, 325, 55, 15, HC_begin
  EditBox 270, 325, 70, 15, ACCI
  CheckBox 5, 355, 50, 10, "Expedited?", expedited
  EditBox 55, 370, 285, 15, verifs_needed
  EditBox 55, 390, 285, 15, actions_taken
  EditBox 50, 410, 50, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 230, 410, 50, 15
    CancelButton 290, 410, 50, 15
  Text 240, 275, 20, 10, "TRAN:"
  Text 5, 10, 35, 10, "CAF type: "
  Text 5, 310, 40, 10, "MEDI/INSA:"
  Text 5, 200, 45, 10, "SHEL/HEST:"
  Text 5, 30, 35, 10, "CAF date:"
  Text 195, 310, 20, 10, "DIET:"
  Text 170, 200, 45, 10, "COEX/DCEX:"
  Text 75, 235, 20, 10, "ACCT:"
  Text 5, 415, 40, 10, "Your name:"
  Text 5, 65, 35, 10, "HH comp:"
  Text 5, 375, 50, 10, "Verifs needed:"
  Text 5, 140, 60, 10, "JOBS/BUSI/RBIC:"
  Text 5, 395, 50, 10, "Actions taken:"
  Text 5, 235, 25, 10, "CASH:"
  Text 170, 85, 65, 10, "SCHL/STIN/STEC:"
  Text 5, 85, 20, 10, "DISA:"
  Text 5, 275, 25, 10, "REST:"
  Text 110, 275, 25, 10, "OTHR:"
  Text 150, 10, 25, 10, "App for:"
  Text 145, 255, 25, 10, "CARS:"
  Text 5, 105, 45, 10, "AREP/SWKR:"
  Text 5, 180, 75, 10, "STWK/Inc. changes?:"
  Text 5, 330, 45, 10, "FMED/BILS:"
  Text 5, 160, 45, 10, "UNEA/PBEN:"
  Text 140, 330, 50, 10, "HC begin date:"
  Text 5, 255, 25, 10, "SECU:"
  Text 250, 330, 20, 10, "ACCI:"
  Text 130, 30, 85, 10, "Recert date (if applicable):"
EndDialog


Dialog CAF_dialog
If buttonpressed = 0 then stopscript

Sub find_case_note
EMReadScreen case_note_ready, 17, 2, 33
EMReadScreen case_note_mode, 7, 20, 3
If case_note_ready <> "Case Notes (NOTE)" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then msgbox "You aren't in a case note on edit mode. You need to be in a case note on edit mode."
If case_note_mode <> "Mode: A" and case_note_mode <> "Mode: E" then Dialog CAF_dialog
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

IF CAF_type = "Intake" then EMSendKey "***Intake***"
IF CAF_type = "Reapplication" then EMSendKey "***Reapplication***"
IF CAF_type = "Recertification" then EMSendKey "***Recertification***"
IF CAF_type = "Add program" then EMSendKey "***Add program***"
EMSendKey "<NewLine>"
EMSendKey "* App for:"
If FS_app = 1 then EMSendKey " FS,"
IF HC_app = 1 then EMSendKey " HC,"
IF cash_app = 1 then EMSendKey " cash,"
If emer_app = 1 then EMSendKey " emer,"
EMSendKey "<backspace>"
If expedited = 1 then EMSendKey ", expedited"
EMSendKey "<newline>"
EMSendKey "* CAF date: " + CAF_date 
If CAF_type = "Recertification" then EMSendKey " (" + recert_date + " recert)"
EMSetCursor 7, 3
EMSendKey "* HH comp: " + HH_comp
EMSetCursor 10, 3
EMSendKey "* DISA: " + DISA
EMSetCursor 12, 3
EMSendKey "* SCHL/STIN/STEC: " + SCHL
EMSetCursor 14, 3
EMSendKey "* AREP/SWKR: " + AREP
EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 0
EMSendKey "* JOBS/BUSI/RBIC: " + JOBS
EMSetCursor 8, 3
EMSendKey "* UNEA/PBEN: " + UNEA
EMSetCursor 11, 3
EMSendKey "* Income changes: " + income_changes
EMSetCursor 14, 3
EMSendKey "* SHEL/HEST: " + SHEL
EMSetCursor 16, 3
EMSendKey "* COEX/DCEX: " + COEX
EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 0
EMSendKey "* CASH: " + CASH
EMSetCursor 5, 3
EMSendKey "* ACCT: " + ACCT
EMSetCursor 8, 3
EMSendKey "* SECU: " + SECU
EMSetCursor 10, 3
EMSendKey "* CARS: " + CARS
EMSetCursor 12, 3
EMSendKey "* REST: " + REST
EMSetCursor 13, 3
EMSendKey "* OTHR: " + OTHR
EMSetCursor 14, 3
EMSendKey "* TRAN: " + TRAN
EMSetCursor 15, 3
EMSendKey "* MEDI/INSA: " + MEDI
EMSetCursor 17, 3
EMSendKey "^" + "<PF8>"
EMWaitReady 1, 0
'The following is the last page of the case note.

EMSendKey "* DIET: " + DIET
EMSetCursor 5, 3
EMSendKey "* FMED: " + FMED
EMSetCursor 6, 3
EMSendKey "* HC begin date: " + HC_begin
EMSetCursor 7, 3
EMSendKey "* ACCI: " + ACCI
EMSetCursor 8, 3
EMSendKey "* Verifs needed: " + verifs_needed
EMSetCursor 12, 3
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
EMSendKey " " + "<PF7>"
EMWaitReady 1, 0
EMSetCursor 17, 3
EMSendKey " " + "<PF3>"
EMWaitReady 1, 0
'Now it reenters the case note.
EMSetCursor 5, 3
EMSendKey "x" + "<enter>"
