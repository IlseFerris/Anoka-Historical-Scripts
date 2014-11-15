'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - INAC scrubber"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog inac_scrubber_dialog, 0, 0, 206, 162, "INAC scrubber"
  Text 5, 5, 200, 10, "This script will transfer the cases in your REPT/INAC to CLS."
  Text 5, 25, 195, 20, "It will check MMIS for each household member, STAT/ABPS for Good Cause status, and CCOL/CLIC for claims."
  Text 5, 55, 195, 20, "Write the information in the boxes below and click ''OK'' to begin. Click ''Cancel'' to exit."
  Text 5, 85, 75, 10, "Sign your case notes:"
  EditBox 80, 80, 80, 15, worker_sig
  Text 5, 105, 90, 10, "Write your worker number:"
  EditBox 100, 100, 60, 15, worker_number
  Text 5, 125, 45, 10, "Footer month:"
  EditBox 55, 120, 35, 15, footer_month
  Text 100, 125, 40, 10, "Footer year:"
  EditBox 145, 120, 35, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 45, 140, 50, 15
    CancelButton 110, 140, 50, 15
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows the dialog
Dialog inac_scrubber_dialog
If buttonpressed = 0 then stopscript

'Converts X102 number to the MAXIS friendly all-caps
worker_number = UCase(worker_number)

'Converts system/footer month and year to a MAXIS-appropriate number, for validation
current_system_month = DatePart("m", Now)
If len(current_system_month) = 1 then current_system_month = "0" & current_system_month
current_system_year = DatePart("yyyy", Now) - 2000
If len(footer_month) <> 2 or isnumeric(footer_month) = False or footer_month > 13 or len(footer_year) <> 2 or isnumeric(footer_year) = False then script_end_procedure("Your footer month and year must be 2 digits and numeric. The script will now stop.")
footer_month_first_day = footer_month & "/01/" & footer_year
date_compare = datediff("d", footer_month_first_day, date)
If date_compare < 0 then script_end_procedure("You appear to have entered a future month and year. The script will now stop.")
If cint(current_system_month) = cint(footer_month) and cint(footer_year) = cint(current_system_year) then script_end_procedure("Do not use this script in the current footer month. These cases need to be in your REPT/INAC for 30 days. The script will now stop.")

'Validates the worker number (should be seven digits and start with "x102" or "X102")
If len(worker_number) <> 7 then script_end_procedure("Your worker number is not 7 digits. Please try again.")
If left(worker_number, 4) <> "x102" and left(worker_number, 4) <> "X102" then script_end_procedure("That worker number is incorrect or does not appear to be in this county. Please try again.")


'SECTION 02

'Connects to MAXIS
EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
transmit
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then script_end_procedure("You appear to be passworded out.")

row = 1
col = 1
EMSearch "MAXIS", row, col
If row <> 1 then script_end_procedure("You need to run this script in the window that has MAXIS on it. Please try again.")

'Gets back to SELF
back_to_self

'Enters REPT/INAC
EMWriteScreen "rept", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "inac", 21, 70
transmit

'Checks to make sure the selected worker number is the default. If not it will navigate to that person.
EMReadScreen worker_number_check, 7, 21, 16
If worker_number_check <> worker_number then
  EMWriteScreen worker_number, 21, 16
  transmit
End if

'Checks to make sure the worker has cases to close. If not the script will end.
EMReadScreen worker_has_cases_to_close_check, 16, 7, 14
If worker_has_cases_to_close_check = "                " then script_end_procedure("This worker does not appear to have any cases to close. If there are cases here email the script administrator a description of the problem and your X102 number.")

'Notifies the worker that we're about to create a Word document.
MsgBox "The script is about to start a Word document. This may take a few moments."

'Before creating the Word document, we create an Excel spreadsheet that runs behind the scenes to collect the case numbers. It's easier to work off of an Excel spreadsheet than an array (for debugging purposes). An array would be faster, however.

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False                                'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = False                          'Set this to false to make alerts go away. This is necessary in production.

'Now it creates a word document with all active claims in it.
'As of 04/01/2013, all get emailed to Steve Gunz. I'm maintaining the alpha split for now in case they hire an additional person.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objselection.typetext "Case numbers with active claims (A-G, email to Christine Herman): "
objselection.TypeParagraph()
objselection.TypeParagraph()

'Assigning values to the Excel spreadsheet.
ObjExcel.Cells(1, 1).Value = "MAXIS number"
ObjExcel.Cells(1, 2).Value = "Name"
ObjExcel.Cells(1, 3).Value = "INAC eff date"
ObjExcel.Cells(1, 4).Value = "Amount due in claims"
ObjExcel.Cells(1, 5).Value = "DAILs?"
ObjExcel.Cells(1, 6).Value = "MMIS?"
ObjExcel.Cells(1, 7).Value = "PMIs"

'Setting the variables for the do...loop.
excel_row = 2 'This sets the variable for the following do...loop.
MAXIS_row = 7 'This sets the variable for the following do...loop.

'This loop grabs the case number, client name, and inac date for each case.
Do
  Do
    EMReadScreen case_number, 8, MAXIS_row, 3          'First it reads the case number, name, and date they closed.
    EMReadScreen client_name, 25, MAXIS_row, 14
    EMReadScreen inac_date, 8, MAXIS_row, 49
    EMReadScreen appl_date, 8, MAXIS_row, 39
    case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each. This is for the Excel spreadsheet, so that we aren't entering blank spaces.
    client_name = Trim(client_name)
    inac_date = Trim(inac_date)
    If appl_date <> inac_date then                     'Because if the two dates equal each other, then this is a denial and not a case closure.
      ObjExcel.Cells(excel_row, 1).Value = case_number   'Then it writes each into the Excel spreadsheet to be used later.
      ObjExcel.Cells(excel_row, 2).Value = client_name
      ObjExcel.Cells(excel_row, 3).Value = inac_date
      excel_row = excel_row + 1
    End if
    MAXIS_row = MAXIS_row + 1
  Loop until MAXIS_row = 19
  MAXIS_row = 7 'Setting the variable for when the do...loop restarts
  PF8
  EMReadScreen last_page_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
Loop until last_page_check = "THIS IS THE LAST PAGE"

'Navigates to CCOL/CLIC
EMWriteScreen "ccol", 20, 22
EMWriteScreen "clic", 20, 70
transmit

'This sets the variable for the following do...loop. It resets as we're going to be looking at each case for claims.
excel_row = 2 

'Grabs any claims due for each case.
Do
  EMWriteScreen "________", 4, 8
  EMWriteScreen ObjExcel.Cells(excel_row, 1).Value, 4, 8
  transmit
  EMReadScreen claims_due, 10, 19, 58
  ObjExcel.Cells(excel_row, 4).Value = claims_due
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""


'This sets the variable for the following do...loop. We'll be splitting the claims from the Excel spreadsheet.
excel_row = 2 

'Setting up the alpha split for the first part of the alphabet
Do
  starting_letter_last_name = left(ObjExcel.Cells(excel_row, 2).Value, 1)
  If starting_letter_last_name = "A" or starting_letter_last_name = "B" or starting_letter_last_name = "C" or starting_letter_last_name = "D" or starting_letter_last_name = "E" or starting_letter_last_name = "F" or starting_letter_last_name = "G" then
    alpha_split = "first third"
  elseif starting_letter_last_name = "H" or starting_letter_last_name = "I" or starting_letter_last_name = "J" or starting_letter_last_name = "K" or starting_letter_last_name = "L" or starting_letter_last_name = "M" or starting_letter_last_name = "N" then 
    alpha_split = "second third"
  else
    alpha_split = "third third"
  End if
  If alpha_split = "first third" and ObjExcel.Cells(excel_row, 4).Value <> 0 then 
    objselection.typetext ObjExcel.Cells(excel_row, 1).Value & ": " & ObjExcel.Cells(excel_row, 2).Value & "; amount due: $" & ObjExcel.Cells(excel_row, 4).Value
    objselection.TypeParagraph()
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'Adding the text to the second part of the document.
objselection.TypeParagraph()
objselection.typetext "Case numbers with active claims (H-N, email to Anna Welch): "
objselection.TypeParagraph()
objselection.TypeParagraph()

'This sets the variable for the following do...loop.
excel_row = 2 

'Setting up the alpha split for the second part of the alphabet
Do
  starting_letter_last_name = left(ObjExcel.Cells(excel_row, 2).Value, 1)
  If starting_letter_last_name = "A" or starting_letter_last_name = "B" or starting_letter_last_name = "C" or starting_letter_last_name = "D" or starting_letter_last_name = "E" or starting_letter_last_name = "F" or starting_letter_last_name = "G" then
    alpha_split = "first third"
  elseif starting_letter_last_name = "H" or starting_letter_last_name = "I" or starting_letter_last_name = "J" or starting_letter_last_name = "K" or starting_letter_last_name = "L" or starting_letter_last_name = "M" or starting_letter_last_name = "N" then 
    alpha_split = "second third"
  else
    alpha_split = "third third"
  End if
  If alpha_split = "second third" and ObjExcel.Cells(excel_row, 4).Value <> 0 then 
    objselection.typetext ObjExcel.Cells(excel_row, 1).Value & ": " & ObjExcel.Cells(excel_row, 2).Value & "; amount due: $" & ObjExcel.Cells(excel_row, 4).Value
    objselection.TypeParagraph()
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'Adding the text to the second part of the document.
objselection.TypeParagraph()
objselection.typetext "Case numbers with active claims (O-Z, email to Steve Gunz): "
objselection.TypeParagraph()
objselection.TypeParagraph()

'This sets the variable for the following do...loop.
excel_row = 2 

'Setting up the alpha split for the third part of the alphabet
Do
  starting_letter_last_name = left(ObjExcel.Cells(excel_row, 2).Value, 1)
  If starting_letter_last_name = "A" or starting_letter_last_name = "B" or starting_letter_last_name = "C" or starting_letter_last_name = "D" or starting_letter_last_name = "E" or starting_letter_last_name = "F" or starting_letter_last_name = "G" then
    alpha_split = "first third"
  elseif starting_letter_last_name = "H" or starting_letter_last_name = "I" or starting_letter_last_name = "J" or starting_letter_last_name = "K" or starting_letter_last_name = "L" or starting_letter_last_name = "M" or starting_letter_last_name = "N" then 
    alpha_split = "second third"
  else
    alpha_split = "third third"
  End if
  If alpha_split = "third third" and ObjExcel.Cells(excel_row, 4).Value <> 0 then 
    objselection.typetext ObjExcel.Cells(excel_row, 1).Value & ": " & ObjExcel.Cells(excel_row, 2).Value & "; amount due: $" & ObjExcel.Cells(excel_row, 4).Value
    objselection.TypeParagraph()
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'Navigating to the DAIL
EMWriteScreen "dail", 20, 20
EMWriteScreen "dail", 20, 69
transmit

'This sets the variable for the following do...loop.
excel_row = 2 

'This checks the DAIL for messages, sends a variable to the Excel spreadsheet. We don't transfer cases with DAIL messages.
Do
  EMWriteScreen "________", 20, 38
  EMWriteScreen ObjExcel.Cells(excel_row, 1).Value, 20, 38
  transmit
  EMReadScreen DAIL_check, 1, 5, 5
  If DAIL_check <> " " then 
    ObjExcel.Cells(excel_row, 5).Value = "Yes"
  Else
    ObjExcel.Cells(excel_row, 5).Value = "No"
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'This sets the variable for the following do...loop.
excel_row = 2 

'Making the header for the next section of the Word document.
objselection.TypeParagraph()
objselection.TypeParagraph()
objselection.typetext "Cases that need to be REINed, STAT/ABPS updated with an ''N'' code for Good Cause Status, and then reapproved for closure:"
objselection.TypeParagraph()

'This do...loop goes into STAT, grabs PMIs for MEMB types 01, 02, 03, 04, and 18, and then navigates to ABPS to get that info.
Do
  back_to_self
  EMWriteScreen "stat", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen ObjExcel.Cells(excel_row, 1).Value, 18, 43
  EMWriteScreen "memb", 21, 70
  transmit
  EMReadScreen SELF_check, 4, 2, 50
  If SELF_check = "SELF" then ObjExcel.Cells(excel_row, 7).Value = "Unable to determine/privileged"
  If SELF_check <> "SELF" then
    excel_col = 7 'This sets the variable for the next do...loop.
    Do
      EMReadScreen PMI_number, 8, 4, 46
      EMReadScreen rel_to_applicant, 2, 10, 42
      If rel_to_applicant = "01" or rel_to_applicant = "02" or rel_to_applicant = "03" or rel_to_applicant = "04" or rel_to_applicant = "18" then ObjExcel.Cells(excel_row, excel_col).Value = PMI_number
      transmit
      excel_col = excel_col + 1
      EMReadScreen no_more_MEMBs_check, 31, 24, 2
    Loop until no_more_MEMBs_check = "ENTER A VALID COMMAND OR PF-KEY"
    EMWriteScreen "abps", 20, 71
    transmit
    EMReadScreen good_cause_check, 1, 5, 47
    If good_cause_check = "P" then
      objselection.typetext ObjExcel.Cells(excel_row, 1).Value & ", " & ObjExcel.Cells(excel_row, 2).Value
      objselection.TypeParagraph()
    End if
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'The following checks for which screen MMIS is running on.
attn
EMReadScreen MMIS_A_check, 7, 15, 15
IF MMIS_A_check = "RUNNING" then 
  EMSendKey "10"
  transmit
End if
IF MMIS_A_check <> "RUNNING" then 
  attn
  EMConnect "B"
  attn
  EMReadScreen MMIS_B_check, 7, 15, 15
  If MMIS_B_check <> "RUNNING" then 
    objExcel.Workbooks.Close
    objExcel.quit
    script_end_procedure("MMIS does not appear to be running. This script will now stop.")
  End if
  If MMIS_B_check = "RUNNING" then 
    EMSendKey "10"
    transmit
  End if
End if

'Shifts user focus to whatever screen ended up getting selected (A or B)
EMFocus

'Sends a PF6 to see if MMIS is running.
PF6
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then 
  objExcel.Workbooks.Close
  objExcel.quit
  script_end_procedure("You appear to be passworded out in MMIS.")
End if

'Gets back to the start screen of MMIS
get_to_MMIS_session_begin

'Gets to the EK01 screen (NOTE: this section may not work for all staff, if they only have EK01.)
EMSetCursor 1, 2
EMSendKey "mw00"
transmit
transmit

'Setting the variables for the next search
row = 1
col = 1

'Looking to confirm we're on EK01. If not it'll try to get in the recipient file application setting.
EMSearch "EK01", row, col
If row <> 0 then 
  EMSetCursor row, 4
  EMSendKey "x"
  transmit
Else
  row = 1
  col = 1
  EMSearch "EKIQ", row, col
  If row <> 0 then
    EMSetCursor row, 4
    EMSendKey "x"
    transmit
  Else
    row = 1
    col = 1
    EMSearch "RECIPIENT FILE APPLICATION", row, col
    If row = 0 then
      objExcel.Workbooks.Close
      objExcel.quit
      script_end_procedure("MMIS could not be found for this user. Contact the script administrator with your x102 number and a description of the problem.")
    Else
      EMWriteScreen "x", row, col - 3
      transmit
    End if
  End if
End if

'This section starts from EK01. Getting to RKEY
EMWriteScreen "x", 10, 3
transmit

'Resetting the variables for the do...loop
excel_row = 2 'Resetting the variable for the next do...loop.
excel_col = 7 'Resetting the variable for the next do...loop.

'Sending the "i" code for MMIS.
EMWriteScreen "i", 2, 19

'This do...loop enters a PMI for each HH member and gets their program status in MMIS.
Do
  Do
    Do
      PMI = ObjExcel.Cells(excel_row, excel_col).Value
      If len(PMI) > 8 then exit do 'Because these cases are privileged and a PMI could not be determined.
      If len(PMI) < 8 then 'This will generate an 8 digit PMI.
        Do 
          PMI = "0" & PMI
        Loop until len(PMI) = 8
      End if
      EMWriteScreen PMI, 4, 19
      transmit
      EMWriteScreen "relg", 1, 8
      transmit
      EMReadScreen MMIS_case_status, 1, 7, 62
      MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
      If len(MAXIS_case_number) < 8 then 'This will generate an 8 digit MAXIS case number.
        Do 
          MAXIS_case_number = "0" & MAXIS_case_number
        Loop until len(MAXIS_case_number) = 8
      End if
      EMReadScreen MMIS_case_number, 8, 6, 73
      If MMIS_case_status = "A" or MMIS_case_status = "P" then
        If isnumeric(MMIS_case_number) = False or MMIS_case_number = MAXIS_case_number then ObjExcel.Cells(excel_row, 6).Value = "Yes"
      End if
      If MMIS_case_status = "C" or MMIS_case_status = "D" then
        EMReadScreen elig_end_date, 8, 7, 36
        If elig_end_date = "99/99/99" then 
          ObjExcel.Cells(excel_row, 6).Value = "Yes"
        Else
          If datediff("m", elig_end_date, now) < 1 and (isnumeric(MMIS_case_number) = False or MMIS_case_number = MAXIS_case_number) then ObjExcel.Cells(excel_row, 6).Value = "Yes"
        End if
      End if
      PF6
      excel_col = excel_col + 1
    Loop until ObjExcel.Cells(excel_row, excel_col).Value = ""
    excel_row = excel_row + 1
    excel_col = 7 'Resetting the variable for the next do...loop.
  Loop until ObjExcel.Cells(excel_row, 1).Value = ""
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'The following checks for which screen MAXIS is running on.
EMConnect "A"
attn
EMReadScreen MAXIS_A_check, 7, 6, 15 
IF MAXIS_A_check = "RUNNING" then 
  EMSendKey "1"
  transmit
End if
IF MAXIS_A_check <> "RUNNING" then 
  attn
  EMConnect "B"
  EMReadScreen MAXIS_B_check, 7, 6, 15
  If MAXIS_B_check <> "RUNNING" then 
    objExcel.Workbooks.Close
    objExcel.quit
    script_end_procedure("MAXIS does not appear to be running. This script will now stop.")
  Else
    EMSendkey "1" 
    transmit
  End if
End if

'Resetting the variable for the next do...loop 
excel_row = 2

'This do...loop updates case notes for all of the cases that don't have DAIL messages or cases still open in MMIS
Do
  back_to_self
  PMI_number = ObjExcel.Cells(excel_row, 7).Value 'It uses this to detemine if the case is privileged or not. Privileged cases would not show a PMI.
  If IsNumeric(PMI_number) = True and ObjExcel.Cells(excel_row, 5).Value = "No" and ObjExcel.Cells(excel_row, 6).Value = "" and datediff("d", ObjExcel.Cells(excel_row, 3).Value, now) > 10 then
    EMWriteScreen "case", 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen ObjExcel.Cells(excel_row, 1).Value, 18, 43
    EMWriteScreen "note", 21, 70
    transmit
    PF9
    EMReadScreen case_note_mode_check, 7, 20, 3 'NOTE: It should not check for this, it should just send the keystrokes on the next line. This is for testing purposes.
    If case_note_mode_check = "Mode: A" then EMSendKey "Sending closed case to CLS via automated script. -" & worker_sig
    If case_note_mode_check <> "Mode: A" then MsgBox "Check, right case?"
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'Resetting the variable for the next do...loop
excel_row = 2 

'This do...loop transfers the cases to X102CLS.
Do 
  back_to_SELF
  PMI_number = ObjExcel.Cells(excel_row, 7).Value 'It uses this to detemine if the case is privileged or not. Privileged cases would not show a PMI.
  If IsNumeric(PMI_number) = True and ObjExcel.Cells(excel_row, 5).Value = "No" and ObjExcel.Cells(excel_row, 6).Value = "" and datediff("d", ObjExcel.Cells(excel_row, 3).Value, now) > 10 then
    EMWriteScreen "spec", 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen ObjExcel.Cells(excel_row, 1).Value, 18, 43
    EMWriteScreen "xfer", 21, 70
    transmit
    EMWriteScreen "x", 7, 16
    transmit
    PF9
    EMWriteScreen "x102CLS", 18, 61
    transmit
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

'Notifies the worker of the success
MsgBox "Success!"  & vbNewLine & vbNewLine &_
"The cases that have HC open in MMIS, or have DAILs generated, are still in your REPT/INAC. Some of these cases may be discrepancies or may be MCRE. Check each one of these manually in MMIS and CCOL/CLIC before sending to CLS. "  & vbNewLine & vbNewLine &_
"A word document has been created. Follow the directions indicated on that document. If you have questions about the procedure, see a program coordinator."

'Closes the Excel workbook and ends the script.
objExcel.Workbooks.Close
objExcel.quit
script_end_procedure("")