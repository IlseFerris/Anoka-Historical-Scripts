'Removed 07/23/2013 as the procedure changed. Will write a new script if asked. Procedure may change dramatically due to MNsure.


'FUNCTIONS----------------------------------------------------------------------------------------------------
function escape
  EMSendKey "<attn>"
  EMWaitReady -1, 0
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
End function

function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

function PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
End function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
End function

function PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
End function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog MSQ_dialog, 0, 0, 151, 62, "MSQ dialog"
  EditBox 90, 5, 55, 15, PMI_number
  EditBox 85, 25, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 20, 45, 50, 15
    CancelButton 80, 45, 50, 15
  Text 10, 10, 80, 10, "Enter your PMI number:"
  Text 10, 30, 70, 10, "Sign your case note:"
EndDialog

'SHOWING THE DIALOG. IT STOPS IF CANCELLED, AND IT LOOPS UNTIL CONDITIONS ARE MET
Do
  Do
    Dialog MSQ_dialog
    If ButtonPressed = 0 then stopscript
    If len(PMI_number) > 8 or IsNumeric(PMI_number) = False then MsgBox "Your PMI number has to be 8 digits or less, and a number. Try again!"
  Loop until len(PMI_number) <= 8 and IsNumeric(PMI_number) = True
  If worker_signature = "" then MsgBox "You must sign your case note."
Loop until worker_signature <> ""

'IT TAKES THAT PMI AND CONVERTS IT TO AN 8 DIGIT NUMBER
Do
  If len(PMI_number) < 8 then PMI_number = "0" & PMI_number
Loop until len(PMI_number) = 8

'IT CONNECTS TO BLUEZONE SCREEN "B", BECAUSE SCREEN "B" IS TYPICALLY WHERE MMIS RESIDES. IT THEN FINDS MMIS. IF NOT FOUND ON "B", IT WILL TRY SCREEN "A".
EMConnect "B"
escape
EMReadScreen MMIS_check, 7, 15, 15
If MMIS_check <> "RUNNING" then 
  EMConnect "A"
  escape
  EMReadScreen MMIS_check, 7, 15, 15
  If MMIS_check <> "RUNNING" then
    MsgBox "MMIS does not appear to be running. The script will now stop."
    stopscript
  End if
End if
EMWriteScreen "10", 2, 15
transmit

'WE'RE IN MMIS AND THE SCRIPT GETS BACK TO START SCREEN OF MMIS. IF THERE'S A PASSWORD PROMPT IT WILL STOP.
Do
  PF6
  EMReadScreen password_check, 28, 2, 33
  If password_check = "PASSWORD VERIFICATION PROMPT" then stopscript
  EMReadScreen session_terminated_check, 18, 1, 7
Loop until session_terminated_check = "SESSION TERMINATED"

'WE GET BACK IN TO MMIS AND TRANSMIT THROUGH THE FRONT MESSAGES SCREENS
EMWriteScreen "mw00", 1, 2
transmit
transmit

'IT CHECKS TO FIND EK01 (MCRE), AND IT TRANSMITS INTO IT. IF NO EK01 IS FOUND THE SCRIPT WILL STOP.
row = 1
col = 1
EMSearch "EK01", row, col
If row = 0 then
  MsgBox "EK01 (MCRE MMIS) is not found. The script will now stop."
  stopscript
End if
EMWriteScreen "x", row, col - 2
transmit

'IT CHECKS TO FIND RECIPIENT FILE APPLICATION, AND IT TRANSMITS INTO IT. IF NOT FOUND, THE SCRIPT WILL STOP.
row = 1
col = 1
EMSearch "RECIPIENT FILE APPLICATION", row, col
If row = 0 then
  MsgBox "Recipient file application is not found. The script will now stop."
  stopscript
End if
EMWriteScreen "x", row, col - 3
transmit

'WE'RE ON RKEY AND THE SCRIPT TYPES THE PMI NUMBER, AND GOES INTO INQUIRY MODE
EMWriteScreen "i", 2, 19
EMWriteScreen PMI_number, 4, 19
transmit

'IT CHECKS THE SCREEN TO MAKE SURE WE'VE GONE PAST RKEY
EMReadScreen RSUM_check, 4, 1, 51
If RSUM_check <> "RSUM" then stopscript 'Because the error message should be explained in the window, no pop-up is needed.

'IT GOES TO RELG
EMWriteScreen "relg", 1, 8
transmit

'IT GRABS THE CASE NUMBER AND DETERMINES IF THE CASE IS OPEN. IF IT IS NOT OPEN, THE SCRIPT STOPS
EMReadScreen case_number, 8, 6, 73 



'IT DETERMINES IF THE CASE IS OPEN. IF THE CASE IS NOT OPEN, THE SCRIPT STOPS.
EMReadScreen elig_end_date, 8, 7, 36
If elig_end_date <> "99/99/99" then
  MsgBox "This client is not currently open on HC. Process manually. The script will now stop."
  stopscript
End if

'NAVIGATES TO RMSQ.
EMWriteScreen "rmsq", 1, 8
transmit

'CHECKING THE RETURN DATE. IF THE MSQ HAS ALREADY COME BACK, THE SCRIPT WILL STOP.
EMReadScreen return_date, 8, 11, 68
If return_date <> "        " then 
  MsgBox "The MSQ has a return date. It may have already returned. Process manually. The script will now stop."
  StopScript
End if

'COLLECTS DATA ON THE ACCIDENT, RIGHT OFF OF THE RMSQ PANEL, THENNAVIGATES TO THE NEXT NEEDED SCREEN USING A SET-CURSOR AND A PF4.
EMReadScreen MSQ_type_code, 1, 11, 07
EMReadScreen MSQ_origin_code, 1, 11, 11
If MSQ_type_code = "1" then MSQ_type = "automobile accident"
If MSQ_type_code = "2" then MSQ_type = "workers compensation"
If MSQ_type_code = "3" then MSQ_type = "other accident"
If MSQ_type_code = "4" then MSQ_type = "diagnosis indicating accident"
If MSQ_origin_code = "A" then MSQ_origin = "STAT/ACCI interface"
If MSQ_origin_code = "C" then MSQ_origin = "BRS requested MSQ"
If MSQ_origin_code = "D" then MSQ_origin = "Dept of Labor & Industry"
If MSQ_origin_code = "M" then MSQ_origin = "Medicaid claims request"
EMSetCursor 11, 23
PF4

'THE SCRIPT NEEDS TO DETERMINE IF THERE IS A "CUB" QUEUE, A "CPH" QUEUE, OR A "CHF" QUEUE. THE INFORMATION CAN BE FOUND ON DIFFERENT SCREENS.
EMReadScreen CUB_check, 3, 1, 48
If CUB_check = "CUB" then
  EMReadScreen service_date_start, 6, 8, 22   'CHECKS THE SERVICE DATES. 
  service_date_start = left(service_date_start, 2) & "/" & mid(service_date_start, 3, 2) & "/" & right(service_date_start, 2)
  EMReadScreen service_date_end, 6, 8, 29
  service_date_end = left(service_date_end, 2) & "/" & mid(service_date_end, 3, 2) & "/" & right(service_date_end, 2)
  EMWriteScreen "CUB3", 1, 8  
  transmit 'Navigating to the CUB3 screen
  EMReadScreen diagnosis_code, 6, 13, 18
  If diagnoisis_code <> "      " then MSQ_type = MSQ_type & ", (diagnosis code " & trim(diagnosis_code) & ")"
  billing_provider_name = "not found in MMIS" 'BECAUSE THE "CUB" QUEUE SCREENS DO NOT CONTAIN A BILLING PROVIDER NAME
ElseIf CUB_check = " CP" then
  EMReadScreen service_date_start, 6, 9, 57   'CHECKS THE SERVICE DATES. 
  service_date_start = left(service_date_start, 2) & "/" & mid(service_date_start, 3, 2) & "/" & right(service_date_start, 2)
  EMReadScreen service_date_end, 6, 9, 57
  service_date_end = left(service_date_end, 2) & "/" & mid(service_date_end, 3, 2) & "/" & right(service_date_end, 2)
  transmit 'Navigating to the CPH2 screen
  EMReadScreen diagnosis_code, 6, 12, 40
  If diagnoisis_code <> "      " then MSQ_type = MSQ_type & ", (diagnosis code " & trim(diagnosis_code) & ")"
  billing_provider_name = "pharmacy (not found in MMIS)" 'BECAUSE THE "CPH" QUEUE SCREENS DO NOT CONTAIN A BILLING PROVIDER NAME
Else
  EMWriteScreen "CHF2", 1, 8  'getting to the CHF2 screen for the diagnosis code
  transmit 
  EMReadScreen diagnosis_code, 6, 10, 23
  If diagnoisis_code <> "      " then MSQ_type = MSQ_type & ", (diagnosis code " & trim(diagnosis_code) & ")"
  EMWriteScreen "CHF3", 1, 8  'getting to the CHF3 screen
  transmit 
  EMReadScreen CHF3_check, 4, 1, 50   'IF IT IS NOT ON THE "CHF3" SCREEN IT WILL ERROR OUT.
  If CHF3_check <> "CHF3" then 
    MsgBox "There appears to have been an error. Some cases do not go into the correct queue. Try again. If it doesn't work, process manually and notify the script administrator with a description of the error and the client PMI. The script will now stop."
    stopscript
  End if 
  EMReadScreen service_date_start, 6, 6, 6   'CHECKS THE SERVICE DATES. 
  service_date_start = left(service_date_start, 2) & "/" & mid(service_date_start, 3, 2) & "/" & right(service_date_start, 2)
  EMReadScreen service_date_end, 6, 6, 13
  service_date_end = left(service_date_end, 2) & "/" & mid(service_date_end, 3, 2) & "/" & right(service_date_end, 2)
  transmit 'Navigating to the CHF4 screen  
  EMReadScreen billing_provider_name, 30, 14, 25  'GRABBING THE BILLING PROVIDER NAME FROM CHF4.
End if

'IF THE CASE WAS MCRE, IT NAVIGATES TO A BLANK CASE NOTE AND WILL CASE NOTE THE RESULTS, THEN END THE SCRIPT
If IsNumeric(case_number) = False then
  PF6
  PF6
  EMWriteScreen "c", 2, 19
  EMWriteScreen "        ", 4, 19
  EMWriteScreen case_number, 9, 19
  transmit
  PF4
  PF11
  EMReadScreen MMIS_case_note_edit_check, 5, 5, 2
  IF MMIS_case_note_edit_check <> "'''''" then
    MsgBox "Error: MMIS case note edit mode not found. Is case out of county? Check on this and try again. If it doesn't work, this may be a bug."
    StopScript
  End if
  EMSendKey ">>>RECIPIENT HAS NOT RESPONDED TO AN MSQ<<<" & "<newline>"
  EMSendKey "* MSQ type reported as: " & MSQ_type &  "<newline>" 
  EMSendKey "* Origin of info for MSQ is: " & MSQ_origin & "<newline>" 
  EMSendKey "* Service dates: " & service_date_start & " to " & service_date_end & "<newline>"
  EMSendKey "* The billing provider name is: " & billing_provider_name & "<newline>"
  EMSendKey "* Sent MSQ to client. TIKLed for 10-day return." & "<newline>"
  EMSendKey "---" & "<newline>"
  EMSendKey worker_signature & "<newline>"
  EMSendKey "***********************************************************************"
  MsgBox "Success! Now, send the form with the MSQ info on it. The info is in the MMIS case note. " & Chr(13) & Chr(13) & "Don't forget: you need to track for the MSQ's return using a MAXIS TIKL. The script doesn't do this for MCRE cases, you have to do it manually!"
  StopScript 'Because if the case number is not numeric, it's not a MAXIS case, as MCRE case numbers have a letter in them.
End if

'FOR MAXIS CASES, IT CONTINUES WITH A MSGBOX POPPING UP WITH THE RESULTS. IT CONTAINS A CANCEL BUTTON.
MSQ_button_pressed = MsgBox ("Enter the following information into the MSQ that you are sending out." & Chr(13) & "Do not press ENTER or OK until AFTER you fill out the form!!!" & Chr(13) & Chr(13) & "Case number: " & case_number &  Chr(13) & "MSQ type: " & MSQ_type &  Chr(13) & "Origin of info: " & MSQ_origin & Chr(13) & Chr(13) & "Service dates: " & service_date_start & " to " & service_date_end & Chr(13) &  "Billing provider: " & billing_provider_name, 1, "MSQ results")
If MSQ_button_pressed = 2 then stopscript

'GETTING BACK TO MAXIS PRODUCTION, WHICH SHOULD ALWAYS BE ON SCREEN "A". IF MAXIS PRODUCTION IS NOT FOUND THE SCRIPT WILL STOP.
EMConnect "A"
escape
EMReadScreen MAXIS_check, 7, 6, 15
If MAXIS_check <> "RUNNING" then
  MsgBox "MAXIS production does not appear to be running. The script will now stop."
  stopscript
End if
EMWriteScreen "1", 2, 15
transmit

'WE'RE IN MAXIS PRODUCTION AND THE SCRIPT WILL NAVIGATE BACK TO THE SELF SCREEN
Do
  PF3
  EMWaitReady 0, 0
  EMReadScreen SELF_check, 27, 2, 28
  EMReadScreen password_check, 28, 2, 33
  If password_check = "PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of MAXIS. Type your password and press ''OK''"
Loop until SELF_check = "Select Function Menu (SELF)"



'NAVIGATING TO CASE/NOTE FOR THE CASE
EMWriteScreen "case", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "note", 21, 70
transmit

'CREATING A NEW CASE NOTE. USING SENDKEYS INSTEAD OF WRITESCREENS AS SENDKEYS WILL WORK REGARDLESS OF STRING LENGTH. THIS CASE NOTE SHOULD NOT EXCEED ONE PAGE.
PF9
EMSendKey ">>>RECIPIENT HAS NOT RESPONDED TO AN MSQ<<<" & "<newline>"
EMSendKey "* MSQ type reported as: " & MSQ_type &  "<newline>" 
EMSendKey "* Origin of info for MSQ is: " & MSQ_origin & "<newline>" 
EMSendKey "* Service dates: " & service_date_start & " to " & service_date_end & "<newline>"
EMSendKey "* The billing provider name is: " & billing_provider_name & "<newline>"
EMSendKey "* Sent MSQ to client. TIKLed for 10-day return." & "<newline>"
EMSendKey "---" & "<newline>"
EMSendKey worker_signature

'BACK TO SELF SCREEN AS WE NEED TO MAKE A TIKL
Do
  PF3
  EMReadScreen SELF_check, 27, 2, 28
  EMReadScreen password_check, 28, 2, 33
  If password_check = "PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of MMIS. Type your password and press ''OK''"
Loop until SELF_check = "Select Function Menu (SELF)"

'NAVIGATING TO TIKL
EMWriteScreen "dail", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "writ", 21, 70
transmit

'CREATING A TIKL FRIENDLY DATE (MM DD YY, AS THREE SEPARATE VARIABLES)
TIKL_day = DatePart("d", dateadd("d", 10, date))
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day 
TIKL_month = DatePart("m", dateadd("d", 10, date))
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month 
TIKL_year = DatePart("yyyy", dateadd("d", 10, date))
TIKL_year = TIKL_year - 2000

'WRITING THE TIKL. ALWAYS USE A SENDKEY FOR A TIKL AS A WRITESCREEN DOES NOT PRESERVE WORD WRAPPING
EMWriteScreen TIKL_month, 5, 18
EMWriteScreen TIKL_day, 5, 21
EMWriteScreen TIKL_year, 5, 24
EMSetCursor 9, 3
EMSendKey "MSQ requested 10 days ago. If not received, take appropriate action."
transmit
PF3

'END OF SCRIPT MESSAGE
MsgBox "Success! The case with an MSQ has been identified, the info case noted, and a TIKL generated for its return. You should have filled out the form when prompted earlier." & Chr(13) & Chr(13) & "If you did not, navigate back to the case note for this case. The information was stored there."
