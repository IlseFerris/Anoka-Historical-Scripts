'Informational front-end message, date dependent.
If datediff("d", "08/07/2013", now) < 5 then MsgBox "This script has been updated as of 08/07/2013! Here's what's new:" & chr(13) & chr(13) & "Now has an option for ''left client a v/m requesting a call back''." 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - appt letter"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog appointment_letter_dialog, 0, 0, 156, 355, "Appointment letter"
  EditBox 75, 5, 50, 15, case_number
  DropListBox 50, 25, 95, 15, "new application"+chr(9)+"recertification", app_type
  EditBox 50, 45, 95, 15, CAF_date
  CheckBox 10, 65, 130, 10, "Check here if this is a recert and the", no_CAF_check
  DropListBox 70, 90, 75, 15, "phone"+chr(9)+"Anoka Office"+chr(9)+"Blaine Office"+chr(9)+"Col Hts Office"+chr(9)+"Lexington Office", interview_type
  EditBox 70, 110, 75, 15, interview_date
  EditBox 70, 130, 75, 15, interview_time
  EditBox 80, 150, 65, 15, client_phone
  CheckBox 10, 190, 95, 10, "Client appears expedited", expedited_check
  CheckBox 10, 205, 135, 10, "Same day interview offered/declined", same_day_declined_check
  EditBox 10, 240, 135, 15, expedited_explanation
  CheckBox 10, 265, 135, 10, "Check here to have the script update", client_delay_check
  CheckBox 10, 290, 135, 10, "Check here if you left V/M with client", voicemail_check
  EditBox 85, 315, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 335, 50, 15
    CancelButton 85, 335, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 15, 30, 30, 10, "App type:"
  Text 15, 50, 35, 10, "CAF date:"
  Text 30, 75, 105, 10, "CAF hasn't been received yet."
  Text 15, 95, 55, 10, "Interview type:"
  Text 15, 115, 50, 10, "Interview date: "
  Text 15, 135, 50, 10, "Interview time:"
  Text 15, 150, 60, 20, "Client phone (if phone interview):"
  GroupBox 5, 175, 145, 85, "Expedited questions"
  Text 10, 220, 135, 20, "If expedited interview date is not within six days of the application, explain:"
  Text 45, 275, 75, 10, "PND2 for client delay."
  Text 45, 300, 75, 10, "requesting a call back."
  Text 15, 320, 65, 10, "Worker signature:"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Searches for a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If isnumeric(case_number) = False then case_number = ""


'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
Do
  Do
    Do
      Do
        Do
          Do
            Do
              Do
                Dialog appointment_letter_dialog
                If ButtonPressed = 0 then stopscript
                If isnumeric(case_number) = False or len(case_number) > 8 then MsgBox "You must fill in a valid case number. Please try again."
              Loop until isnumeric(case_number) = True and len(case_number) <= 8 
              CAF_date = replace(CAF_date, ".", "/")
              If no_CAF_check = 1 and app_type = "new application" then no_CAF_check = 0 'Shuts down "no_CAF_check" so that it will validate the date entered. New applications can't happen if a CAF wasn't provided.
              If no_CAF_check = 0 and isdate(CAF_date) = False then Msgbox "You did not enter a valid CAF date (MM/DD/YYYY format). Please try again."
            Loop until no_CAF_check = 1 or isdate(CAF_date) = True
            if interview_type = "phone" and client_phone = "" then MsgBox "If this is a phone interview, you must enter a phone number! Pleast try again."
          Loop until interview_type <> "phone" or (interview_type = "phone" and client_phone <> "")
          interview_date = replace(interview_date, ".", "/")
          If isdate(interview_date) = False then MsgBox "You did not enter a valid interview date (MM/DD/YYYY format). Please try again."
        Loop until isdate(interview_date) = True 
        If interview_time = "" then MsgBox "You must type an interview time. Please try again."
      Loop until interview_time <> ""
      If no_CAF_check = 1 then exit do 'If no CAF was turned in, this layer of validation is unnecessary, so the script will skip it.
      If expedited_check = 1 and datediff("d", CAF_date, interview_date) > 6 and expedited_explanation = "" then MsgBox "You have indicated that your case is expedited, but scheduled the interview date outside of the six-day window. An explanation of why is required for QC purposes."
    Loop until expedited_check = 0 or (datediff("d", CAF_date, interview_date) <= 6) or (datediff("d", CAF_date, interview_date) > 6 and expedited_explanation <> "")
    If worker_signature = "" then MsgBox "You must provide a signature for your case note."
  Loop until worker_signature <> ""
  transmit
  EMReadScreen MAXIS_check, 5, 1, 39
  IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You need to be in MAXIS for this to work. Please try again."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'Converting the CAF_date variable to a date, for cases where a CAF was turned in
If no_CAF_check = 0 then CAF_date = cdate(CAF_date)

'Figuring out the last contact day
If app_type = "recertification" then
  next_month = datepart("m", dateadd("m", 1, interview_date))
  next_month_year = datepart("yyyy", dateadd("m", 1, interview_date))
  last_contact_day = dateadd("d", -1, next_month & "/01/" & next_month_year)
End if
If app_type = "new application" then last_contact_day = CAF_date + 31

'Navigating to SPEC/MEMO
call navigate_to_screen("SPEC", "MEMO")

'This checks to make sure we've moved passed SELF.
EMReadScreen SELF_check, 27, 2, 28
If SELF_check = "Select Function Menu (SELF)" then StopScript 

'This will check the county. If this case is not x102, the script will stop.
EMReadScreen county, 4, 20, 14 
If county <> "X102" then 
  MsgBox "This case is not in Anoka County. Check your case number and try again."
  StopScript
End if

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then 
  MsgBox "You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production."
  stopscript
End if
EMWriteScreen "x", 5, 10
transmit

'Writes the MEMO.
EMSetCursor 3, 15
EMSendKey "************************************************************"
IF app_type = "new application" then EMSendKey "You recently applied for assistance in Anoka County on " & CAF_date & ". An interview is required to process your application." & "<newline>" & "<newline>"
If app_type = "recertification" and no_CAF_check = 0 then EMSendKey "You sent recertification paperwork to Anoka County on " & CAF_date & ". An interview is required to process your application." & "<newline>" & "<newline>"
If app_type = "recertification" and no_CAF_check = 1 then EMSendKey "You asked us to set up an interview for your recertification. Remember to send in your forms before the interview." & "<newline>" & "<newline>"
If interview_type = "phone" then EMSendKey "Your phone interview is scheduled for "
If interview_type = "Anoka Office" or interview_type = "Blaine Office" or interview_type = "Col Hts Office" or interview_type = "Lexington Office" then EMSendKey "Your in-office interview is scheduled for "
EMSendKey interview_date & " at " & interview_time & "." & "<newline>" & "<newline>"
If interview_type = "phone" then EMSendKey "We will be calling you at this number: " & client_phone & ". " & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview in the office, please call your worker. "
If interview_type = "Anoka Office" then EMSendKey "Your interview is at the Anoka office, located at:" & "<newline>" & "   2100 3rd Ave, Suite 400" & "<newline>" & "   Anoka, MN 55303" & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number. "
If interview_type = "Blaine Office" then EMSendKey "Your interview is at the Blaine office, located at:" & "<newline>" & "   1201 89th Ave, Suite 400" & "<newline>" & "   Blaine, MN 55434" & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number. "
If interview_type = "Col Hts Office" then EMSendKey "Your interview is at the Columbia Hts office, located at:" & "<newline>" & "   3980 Central Ave NE" & "<newline>" & "   Columbia Heights, MN 55421" & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number. "
If interview_type = "Lexington Office" then EMSendKey "Your interview is at the Lexington office, located at:" & "<newline>" & "   4175 Lovell Road, Suite 128" & "<newline>" & "   Lexington, MN 55014" & "<newline>" & "<newline>" & "If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number. "
EMSendKey "If we do not hear from you by " '& 
EMGetCursor cursor_row, cursor_col
If cursor_row = 17 then PF8
EMSendKey last_contact_day & " we will deny your application." & "<newline>"
EMSendKey "************************************************************"

'Exits the MEMO
PF4

'UPDATES PND2 FOR CLIENT DELAY IF CHECKED
If client_delay_check = 1 then 
  call navigate_to_screen("rept", "pnd2")
  EMGetCursor PND2_row, PND2_col
  for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
    EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
    If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
    EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
    If PND2_HC_status_check = "P" then
      EMWriteScreen "x", PND2_row, 3
      transmit
      person_delay_row = 7
      Do
        EMReadScreen person_delay_check, 1, person_delay_row, 39
        If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
        person_delay_row = person_delay_row + 2
      Loop until person_delay_check = " " or person_delay_row > 20
      PF3
    End if
    EMReadScreen additional_app_check, 14, PND2_row + 1, 17
    If additional_app_check <> "ADDITIONAL APP" then exit for
    PND2_row = PND2_row + 1
  next
  PF3
  EMReadScreen PND2_check, 4, 2, 52
  If PND2_check = "PND2" then
    MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
    PF10
    client_delay_check = 0
  End if
End if

'Navigates to CASE/NOTE
call navigate_to_screen("case", "note")
PF9

'Writes the case note
If app_type = "new application" then EMSendKey "**New CAF received " & CAF_date & ", appt letter sent in MEMO**" & "<newline>"
If same_day_declined_check = 1 then EMSendKey "* Same day interview offered and declined." & "<newline>"
If app_type = "recertification" and no_CAF_check = 0 then EMSendKey "**Recert CAF received " & CAF_date & ", appt letter sent in MEMO**" & "<newline>"
If app_type = "recertification" and no_CAF_check = 1 then EMSendKey "**Client requested recert appointment, letter sent in MEMO**" & "<newline>"
EMSendKey "* Appointment is " & interview_date & " at " & interview_time & "." & "<newline>" 
If expedited_explanation <> "" then call write_editbox_in_case_note("Why interview is more than six days from now", expedited_explanation, 5)
call write_editbox_in_case_note("Appointment type", interview_type, 5)
If client_phone <> "" then call write_editbox_in_case_note("Client phone", client_phone, 5)
call write_new_line_in_case_note("* Client must complete interview by " & last_contact_day & ".")
If client_delay_check = 1 then call write_new_line_in_case_note("* PND2 updated for client delay.")
If voicemail_check = 1 then call write_new_line_in_case_note("* Left client a voicemail requesting a call back.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")
