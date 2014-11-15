'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - FSET non-compliance TIKLer"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
file_path = "Q:\Blue Zone Scripts\Spreadsheets for script use\FSET non-compliance list\No Show List.xls" 'Appears to be Excel 97 format 
excel_row = 2 'Starts with row 2

'FILESYSTEMOBJECTS FOR SCRIPT----------------------------------------------------------------------------------------------------
Set fso = CreateObject("Scripting.FileSystemObject")

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'Validating user ID
If user_ID_for_validation <> "VKCARY" and _
	user_ID_for_validation <> "SLCARDA" and _
	user_ID_for_validation <> "EABUELOW" and _
	user_ID_for_validation <> "PHBROCKM" _
	then script_end_procedure("User " & user_ID_for_validation & " is not authorized to use this script. To be added to the allowed users' group, email the script administrator, and include the user ID indicated. Thank you!")

'Checks to make sure the file exists. If it doesn't the script will exit.
If fso.FileExists(file_path) = False then script_end_procedure("''No show list'' not found. The list should be saved at " & file_path & ". Check to make sure the file is there and try again.")

'LOADING EXCEL
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open(file_path) 
objExcel.DisplayAlerts = False 'Set this to false to make alerts go away. This is necessary in production.

'Checks the no-show date on the spreadsheet. If it's older than 7 days, it'll assume it hasn't been updated and it'll stop.
no_show_date = cdate(ObjExcel.Cells(2, 1).Value)
If no_show_date < date - 7 = True then script_end_procedure("The no-show list appears to be older than seven days. It may not have been updated to the most recent version. JTC emails a new version of the file weekly on Fridays. If you have not received a new version of the file email JTC and request the most recent ''no show list''.")

'Determining the footer month and year based on the no_show_date variable
footer_month = datepart("m", no_show_date)
If len(footer_month) < 2 then footer_month = "0" & footer_month	'Because MAXIS footer months must be two digits
footer_year = right(datepart("yyyy", no_show_date), 2)		'Only need the last two digits of the year

'Connecting to BlueZone
EMConnect ""

'Starting the do...loop
Do
	'Navigating back to the SELF menu
	back_to_self
  
	'Entering the correct footer month and transmitting
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
	transmit
  
	'Pulling the orientation date, case number, and PMI from Excel
	case_number = ObjExcel.Cells(excel_row, 3).Value
	PMI_number = ObjExcel.Cells(excel_row, 2).Value
	orientation_date = ObjExcel.Cells(excel_row, 1).Value
  
	'If the case_number variable is blank we exit the do...loop
	If case_number = "" then exit do
  
	'Going to STAT/MEMB (to find the HH memb number)
	call navigate_to_screen("STAT", "MEMB")
  
	'Finding the MEMB number with the corresponding PMI
	Do
		EMReadScreen PMI_on_MEMB, 8, 4, 46
		If trim(PMI_on_MEMB) <> trim(PMI_number) then transmit
	Loop until trim(PMI_on_MEMB) = trim(PMI_number)
  
	'Reading reference number variable for the correct HH member
	EMReadScreen ref_nbr, 2, 4, 33
  

	'Going to DAIL/WRIT
	call navigate_to_screen("DAIL", "WRIT")

	'Sends the TIKL
	EMSetCursor 9, 3
	EMSendKey "MEMB " & ref_nbr & " FAILED TO ATTEND FSET ORIENTATION " & orientation_date & ". EVALUATE FOR POSSIBLE SNAP SANCTION AND CLOSURE. CONTACT A PC WITH QUESTIONS. TIKL AUTOGENERATED VIA SCRIPT."
	transmit

	'Exits the TIKL
	PF3

	'Advances excel_row by one
	excel_row = excel_row + 1
Loop until case_number = ""	'End of loop

MsgBox "Success! The script has TIKLed out for the cases indicated on the spreadsheet."
script_end_procedure("")