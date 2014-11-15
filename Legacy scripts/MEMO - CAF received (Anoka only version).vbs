EMConnect ""

  row = 1
  col = 1

EMSearch "Case Nbr: ", row, col

EMReadScreen case_number, 8, row, col + 10
If case_number = "AR" then case_number = ""

BeginDialog CAF_received_dialog, 0, 0, 146, 181, "CAF received"
  EditBox 70, 0, 50, 15, case_number
  DropListBox 45, 20, 95, 15, "new application"+chr(9)+"recertification", app_type
  EditBox 45, 40, 95, 15, CAF_date
  DropListBox 65, 60, 75, 15, "phone"+chr(9)+"in-person", interview_type
  EditBox 65, 80, 75, 15, interview_date
  EditBox 65, 100, 75, 15, interview_time
  EditBox 75, 120, 65, 15, client_phone
  EditBox 80, 140, 60, 15, worker_sig
  ButtonGroup CAF_received_dialog_ButtonPressed
    OkButton 20, 160, 50, 15
    CancelButton 80, 160, 50, 15
  Text 20, 5, 50, 10, "Case number:"
  Text 10, 25, 30, 10, "App type:"
  Text 10, 45, 35, 10, "CAF date:"
  Text 10, 65, 55, 10, "Interview type:"
  Text 10, 85, 50, 10, "Interview date: "
  Text 10, 105, 50, 10, "Interview time:"
  Text 10, 120, 65, 20, "Client phone (if phone interview):"
  Text 10, 145, 65, 10, "Sign the case note:"
EndDialog

'This Do...loop checks for the password prompt.
Do
  Dialog CAF_received_dialog
  If CAF_received_dialog_ButtonPressed = 0 then stopscript
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  IF MAXIS_check <> "MAXIS" then MsgBox "You need to be in MAXIS for this to work. Please try again."
  If case_number = "" or interview_date = "" or interview_time = "" or worker_sig = "" then MsgBox "You must fill in a case number, interview date/time, and a signature before continuing."
  CAF_date = replace(CAF_date, ".", "/")
  If isdate(CAF_date) = False then Msgbox "You did not enter a valid date (MM/DD/YYYY format). Try again."
  If isdate(CAF_date) = True then CAF_date = cdate(CAF_date)
Loop until MAXIS_check = "MAXIS" and (case_number <> "" and isdate(CAF_date) = True and interview_date <> "" and interview_time <> "" and worker_sig <> "")

If app_type = "recertification" then
  current_month = datepart("m", date)
  current_year = datepart("yyyy", date)
  next_month = current_month + 1
  next_month_year = current_year
  If next_month = 13 then
    next_month = 1
    next_month_year = current_year + 1
  end if
  last_contact_day = cdate(next_month & "/01/" & next_month_year) - 1
End if

If app_type = "new application" then last_contact_day = CAF_date + 31


'This Do...loop gets back to SELF
do
EMSendKey "<PF3>"
EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWaitReady 1, 1


EMSetCursor 16, 43
EMSendKey "spec"
EMSetCursor 18, 43
EMSendkey "<eraseeof>" + case_number
EMSetCursor 21, 70
EMSendkey "memo" + "<enter>"
EMWaitReady 1, 1
'--------------ERROR PROOFING--------------
EMReadScreen still_self, 27, 2, 28 'This checks to make sure we've moved passed SELF.
If still_self = "Select Function Menu (SELF)" then StopScript 
EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
If no_MEMB = "Arrival Date:" then MsgBox "This HH member does not exist."
If no_MEMB = "Arrival Date:" then StopScript
EMReadScreen county, 4, 20, 14 'This will check the county. If this case is not x102, the script will stop.
If county <> "X102" then MsgBox "This case is not in Anoka County. Check your case number and try again."
If county <> "X102" then StopScript
'--------------END ERROR PROOFING--------------
EMSendKey "<PF5>"
EMWaitReady 1, 1
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then MsgBox "You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production."
If memo_display_check = "Memo Display" then stopscript
EMSetCursor 5, 10
EMSendKey "x" + "<enter>"
EMWaitReady 1, 1
EMSetCursor 3, 15
IF app_type = "new application" then EMSendKey "You recently applied for assistance in Anoka County on "
If app_type = "recertification" then EMSendKey "You sent recertification paperwork to Anoka County on "
EMSendKey CAF_date & ". An interview is required to process your application." & "<newline>" & "<newline>"
If interview_type = "phone" then EMSendKey "Your phone interview is scheduled for "
If interview_type = "in-person" then EMSendKey "Your in-office interview is scheduled for "
EMSendKey interview_date & " at " & interview_time & "." & "<newline>" & "<newline>"
If interview_type = "phone" then EMSendKey "We will be calling you at this number: " & client_phone & ". " & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview in the office, please call your worker. "
If interview_type = "in-person" then EMSendKey "Our office is located at:" & "<newline>" & "   2100 3rd Ave, Suite 400" & "<newline>" & "   Anoka, MN 55303" & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number. "
EMSendKey "If we do not hear from you by " & last_contact_day & " we will deny your application."
EMSendKey "<PF4>"
EMWaitReady 1, 1
EMSetCursor 19, 22
EMSendKey "case"
EMSetCursor 19, 70
EMSendKey "note"
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<PF9>"
EMWaitReady 1, 1
If app_type = "new application" then EMSendKey "**New CAF received " & CAF_date & ", appt letter sent in MEMO**" & "<newline>"
If app_type = "recertification" then EMSendKey "**Recert CAF received " & CAF_date & ", appt letter sent in MEMO**" & "<newline>"
EMSendKey "* Appointment is " & interview_date & " at " & interview_time & ". Appointment type is " & interview_type & "." & "<newline>" 
EMSendKey "* Client must complete interview by " & last_contact_day & "." & "<newline>" & "---" & "<newline>" & worker_sig
