'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - case XFER"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You appear to be locked out of MAXIS")


'Calls the existing excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\case transfers\03.11.2014 - Ilse to Cassie.xlsx") 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

'Variable setting
excel_row = 2
column_containing_new_worker = 6 '<<<<SHOULD SET IN DIALOG FOR PRODUCTION USE BY OTHERS


Do
	new_worker = ObjExcel.Cells(excel_row, column_containing_new_worker).Value
	case_number = ObjExcel.Cells(excel_row, 1).Value
	If case_number = "" then exit do
	back_to_self
	'Now we navigate to SPEC/XFER
	EMWriteScreen "SPEC", 16, 43
	EMWriteScreen "________", 18, 43
	EMWriteScreen case_number, 18, 43
	EMWriteScreen "XFER", 21, 70
	transmit
	EMWriteScreen "x", 7, 16
	transmit
	PF9
	EMWriteScreen new_worker, 18, 65
	transmit
	excel_row = excel_row + 1
Loop until case_number = ""

script_end_procedure("")

