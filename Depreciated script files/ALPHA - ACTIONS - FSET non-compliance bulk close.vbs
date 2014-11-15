'Removed 04/17/2014 per Pat's request: replaced with a version which does not close the non-compliant cases (as closure exemptions were
'	considered "too risky" at the time of development). Replaced with BULK - FSET non-compliance TIKLer.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ALPHA - ACTIONS - FSET non-compliance bulk close"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
file_path = "H:\Bulk script projects\FSET no-show lists\No Show List - 04.11.2014.xls" 'Appears to be Excel 97 format
sanction_date = "050114" 'This should be a dialog or should be programmed automatically
number_of_sanctions = "01" 'This should be programmed automatically
excel_row = 2 'Starts with row 2
footer_month_and_year = "0514" 'This should be programmed automatically

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Loading Excel sheet
'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open(file_path) 
objExcel.DisplayAlerts = False 'Set this to false to make alerts go away. This is necessary in production.

'Connecting to BlueZone
EMConnect ""

'Starting the do...loop
Do

	'Navigating back to the SELF menu
	back_to_self
  
	'Entering the correct footer month and transmitting
	EMSetCursor 20, 43
	EMSendKey footer_month_and_year
	transmit
  
	'Pulling the case number and PMI from Excel
	case_number = ObjExcel.Cells(excel_row, 3).Value
	PMI_number = ObjExcel.Cells(excel_row, 2).Value
  
	'If the case_number variable is blank we exit the do...loop
	If case_number = "" then exit do
  
	'Going to MEMB
	call navigate_to_screen("STAT", "MEMB")
  
	'Finding the MEMB number with the corresponding PMI
	Do
		EMReadScreen PMI_on_MEMB, 8, 4, 46
		If trim(PMI_on_MEMB) <> trim(PMI_number) then transmit
	Loop until trim(PMI_on_MEMB) = trim(PMI_number)
  
	'Reading reference number variable for the correct HH member
	EMReadScreen ref_nbr, 2, 4, 33
  
	'Navigating to WREG for the HH member
	EMWriteScreen "WREG", 20, 71
	EMWriteScreen ref_nbr, 20, 76
	transmit
  
	'Gets panel in edit mode
	PF9
  
	'Updates FSET WREG status code as "02" and clears out the defer FSET question
	EMWriteScreen "02", 8, 50
	EMWriteScreen "_", 8, 80
  
	'Updates sanction date
	EMSetCursor 10, 50
	EMSendKey sanction_date

	EMWriteScreen number_of_sanctions, 11, 50	'Writes the number of sanctions in. <<<<<<<<<<<<<<<<<<<For now it just puts an "01" in, this should be programmed dynamically eventually.
	transmit						'Transmitting on the screen

	excel_row = excel_row + 1			'Updating the excel_row variable to do the next case
	MsgBox "check 1"					'<<<<<<<<<STEP THROUGH MSGBOX FOR THE FIRST RUN
Loop until case_number = ""	'End of loop

'Now it goes back to self and starts looking at closed cases from the top of the spreadsheet.
'To do that we must reset variables.

'Resetting excel_row
excel_row = 2 

'Starting do...loop
Do

	'Going back to self
	back_to_self
  
	'Pulling case number from excel. If case_number is blank it will exit the do...loop
	case_number = ObjExcel.Cells(excel_row, 3).Value
	If case_number = "" then exit do
  
	'Navigates to ELIG/FS. It'll read for the SELF menu to see if the case is still in background. If case is in background it will back out and wait a few seconds, then try again.
	Do
		EMWriteScreen "elig", 16, 43
		EMWriteScreen "________", 18, 43
		EMWriteScreen case_number, 18, 43
		EMWriteScreen "fs__", 21, 70
		transmit
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then
			PF3
			Pause 2
		End if
	Loop until SELF_check <> "SELF"
  
	'Checks for STAT edits. If there's a STAT edit it's going to case note and TIKL for the case. If not, it's going to close the case.
	EMReadScreen STAT_edit_check, 4, 24, 2
	If STAT_edit_check = "STAT" then
		'<<<<<<<<<STEP THROUGH MSGBOX FOR THE FIRST RUN
		MsgBox "check 2"
		EMWriteScreen "DAIL", 19, 22
		EMWriteScreen "WRIT", 19, 70
		transmit
		EMSetCursor 9, 3
		EMSendKey "Evaluate case for closure due to FSET noncompliance. Script was not able to autoclose case due to STAT errors. TIKL generated by script. If you have questions, consult a PC."
		transmit
		PF3
	Else
		EMWriteScreen "FSSM", 19, 70
		transmit
		EMWriteScreen "app", 19, 70
		'<<<<<<<<<STEP THROUGH MSGBOX FOR THE FIRST RUN
		MsgBox "check 3"
		transmit
		transmit
		'<<<<<<<<<STEP THROUGH MSGBOX FOR THE FIRST RUN
		MsgBox "check 4"
		ObjExcel.Cells(excel_row, 5).Value = "CLOSED" 'for tracking purposes
		call navigate_to_screen("case", "note") 'Now it case notes
		PF9
		EMSendKey "---SNAP closure for PMI " & ObjExcel.Cells(excel_row, 2).Value & " due to FSET non-compliance---" & "<newline>"
		EMSendKey "* No show date: " & ObjExcel.Cells(excel_row, 1).Value & "<newline>"
		EMSendKey "---" & "<newline>"
		EMSendKey "Case updated automatically via script. "
		'<<<<<<<<<STEP THROUGH MSGBOX FOR THE FIRST RUN
		MsgBox "check 5"
		PF3
	End if
  
	'Jumps to next row
	excel_row = excel_row + 1

Loop until case_number = ""

script_end_procedure("")