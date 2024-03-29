'Informational front-end message, date dependent.
If datediff("d", "02/28/2013", now) < 7 then MsgBox "This script has been updated as of 02/28/2013! It now can be used for both SNAP ER NOMIs as well as regular SNAP NOMIs. Follow the prompts and see a sup or PC with questions." 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - NOMI"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog SNAP_ER_NOMI_dialog, 0, 0, 211, 102, "SNAP ER NOMI Dialog"
  Text 5, 5, 50, 10, "Case number:"
  EditBox 60, 0, 65, 15, case_number
  Text 5, 25, 85, 10, "Date of missed interview:"
  EditBox 95, 20, 50, 15, date_of_missed_interview
  Text 5, 45, 85, 10, "Time of missed interview:"
  EditBox 95, 40, 50, 15, time_of_missed_interview
  Text 5, 60, 125, 20, "Recert must be complete by (usually the last day of the current month):"
  EditBox 130, 60, 75, 15, last_day_for_recert
  Text 55, 85, 70, 10, "Sign your case note:"
  EditBox 130, 80, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 5, 50, 15
    CancelButton 155, 25, 50, 15
EndDialog

BeginDialog NOMI_dialog, 0, 0, 151, 125, "NOMI Dialog"
  Text 20, 10, 50, 10, "Case number:"
  EditBox 70, 5, 65, 15, case_number
  Text 5, 30, 85, 10, "Date of missed interview:"
  EditBox 95, 25, 50, 15, date_of_missed_interview
  Text 5, 50, 85, 10, "Time of missed interview:"
  EditBox 95, 45, 50, 15, time_of_missed_interview
  Text 20, 70, 60, 10, "Application date:"
  EditBox 80, 65, 50, 15, application_date
  Text 5, 90, 65, 10, "Worker signature:"
  EditBox 70, 85, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 20, 105, 50, 15
    CancelButton 80, 105, 50, 15
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(replace(case_number, "_", ""))
If isnumeric(case_number) = False then case_number = ""

'Asks if this is a recert. A recert uses a SPEC/MEMO notice, vs. a SPEC/LETR for intakes and add-a-programs.
recert_check = MsgBox("Is this a missed SNAP recertification interview?" & Chr(13) & Chr(13) & "If yes, the SNAP missed recert interview notice will be sent. " & Chr(13) & "If no, the regular NOMI will be sent.", 3)
If recert_check = 2 then stopscript
If recert_check = 6 then

  'Shows dialog, checks for password prompt
  Do
    Do
      Dialog SNAP_ER_NOMI_dialog
      If ButtonPressed = 0 then stopscript
      If case_number = "" then MsgBox "You did not enter a case number. Please try again."
      If date_of_missed_interview = "" then MsgBox "You did not enter a date of missed interview. Please try again."
      If time_of_missed_interview = "" then MsgBox "You did not enter a time of missed interview. Please try again."
      If last_day_for_recert = "" then MsgBox "You did not enter a date the recert must be completed by. Please try again."
      If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
    Loop until case_number <> "" and date_of_missed_interview <> "" and time_of_missed_interview <> "" and last_day_for_recert <> "" and worker_signature <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be outside of MAXIS. You may be locked out of MAXIS, check your screen and try again."
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
    
  'Navigates into SPEC/MEMO
  call navigate_to_screen("SPEC", "MEMO")
  
  'Checks to make sure we're past the SELF menu
  EMReadScreen still_self, 27, 2, 28 
  If still_self = "Select Function Menu (SELF)" then script_end_procedure("Script was not able to get past SELF menu. Is case in background?")
  
  'Checks the county
  EMReadScreen county, 4, 20, 14 
  If county <> "X102" then script_end_procedure("This case is not in Anoka County. Check your case number and try again.")
  
  'Creates a new MEMO. If it's unable the script will stop.
  PF5
  EMReadScreen memo_display_check, 12, 2, 33
  If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
  EMWriteScreen "x", 5, 10
  transmit
  
  'Writes the info into the MEMO.
  EMSetCursor 3, 15
  EMSendKey "************************************************************"
  EMSendKey "You have missed your Food Support interview that was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & "." & "<newline>" & "<newline>"
  EMSendKey "Please contact your worker at the telephone number listed below to reschedule the required Food Support interview." & "<newline>" & "<newline>"
  EMSendKey "The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date." & "<newline>"
  EMSendKey "************************************************************"
  PF4  

  'Navigates to a blank case note
  call navigate_to_screen("case", "note")
  PF9
  
  'Writes the case note
  EMSendKey "**Client missed SNAP recertification interview**" & "<newline>"
  EMSendKey "* Appointment was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & "." & "<newline>" 
  EMSendKey "* A SPEC/MEMO has been sent to the client informing them of missed interview." & "<newline>"
  EMSendKey "---" & "<newline>" 
  EMSendKey worker_signature
  MsgBox "Success! A SPEC/MEMO has been sent with the correct language for a missed SNAP recert. A case note has been made."

Elseif recert_check = 7 then

  'Shows dialog, checks for password prompt
  Do
    Do
      Dialog NOMI_dialog
      If ButtonPressed = 0 then stopscript
      If case_number = "" then MsgBox "You did not enter a case number. Please try again."
      If isdate(date_of_missed_interview) = False then MsgBox "You did not enter a valid date of missed interview. Please try again."
      If time_of_missed_interview = "" then MsgBox "You did not enter a time of missed interview. Please try again."
      If isdate(application_date) = False then MsgBox "You did not enter a valid application date. Please try again."
      If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
    Loop until case_number <> "" and isdate(date_of_missed_interview) = True and time_of_missed_interview <> "" and isdate(application_date) = True and worker_signature <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be outside of MAXIS. You may be locked out of MAXIS, check your screen and try again."
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  
  'Navigates into SPEC/LETR
  call navigate_to_screen("SPEC", "LETR")
  
  'Checks to make sure we're past the SELF menu
  EMReadScreen still_self, 27, 2, 28 
  If still_self = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past the SELF screen. Is your case in background?")
  
  'Checks the county
  EMReadScreen county, 4, 21, 13
  If county <> "X102" then script_end_procedure("This case is not in Anoka County. Check your case number and try again.")
  
  'Opens up the NOMI LETR. If it's unable the script will stop.
  EMWriteScreen "x", 7, 12
  transmit
  EMReadScreen LETR_check, 4, 2, 49
  If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
  
  'Writes the info into the NOMI.
  EMWriteScreen "x", 7, 17
  call create_MAXIS_friendly_date(application_date, 0, 12, 38) 
  call create_MAXIS_friendly_date(date_of_missed_interview, 0, 14, 38) 
  transmit
  PF4
  
  'Navigates to a blank case note
  call navigate_to_screen("case", "note")
  PF9
  
  'Writes the case note
  EMSendKey "**Client missed SNAP interview**" & "<newline>"
  EMSendKey "* Appointment was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & "." & "<newline>" 
  EMSendKey "* A NOMI has been sent via SPEC/LETR informing them of missed interview." & "<newline>"
  EMSendKey "---" & "<newline>" 
  EMSendKey worker_signature
  MsgBox "Success! The NOMI has been sent and a case note has been made."

End if

script_end_procedure("")
