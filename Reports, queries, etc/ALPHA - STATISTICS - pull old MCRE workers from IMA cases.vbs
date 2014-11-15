''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
excel_row = 5080 'Starts on 2

'Calls the existing excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\ACA\12.19.2013 IMA cases from Infopac.xlsx") 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

'Connects to MMIS
EMConnect ""

Do

'Checks for RKEY
EMReadScreen RKEY_check, 4, 1, 52
If RKEY_check <> "RKEY" then
  MsgBox "RKEY NOT FOUND. STOPPING SCRIPT."
  StopScript
End if

'Taking PMI from the spreadsheet
PMI = ObjExcel.Cells(excel_row, 2).Value 

'If PMI is blank the loop should exit
If PMI = "" then exit do

'Loading the PMI into RKEY, clearing any case numbers, and pressing transmit
EMWriteScreen PMI, 4, 19
EMWriteScreen "        ", 9, 19
transmit

'Navigating to RELG
EMWriteScreen "RELG", 1, 8
transmit

'Now hopefully these cases will have the MCRE number located directly below the current span.
'Loading the MCRE case number into the spreadsheet
EMReadScreen MCRE_case_number, 8, 10, 73
ObjExcel.Cells(excel_row, 8).Value = MCRE_case_number

'Jumping out of this person-based panel and back to RKEY
PF3

'Clearing the PMI, entering the MCRE case number, and pressing transmit
EMWriteScreen "        ", 4, 19
EMWriteScreen MCRE_case_number, 9, 19
transmit

'Navigating to RCIN
EMWriteScreen "RCIN", 1, 8
transmit

'Reading the worker number and loading into the spreadsheet
EMReadScreen MCRE_worker, 7, 2, 46
ObjExcel.Cells(excel_row, 9).Value = MCRE_worker

'Navigating back to RKEY
PF3

'Raising the excel_row variable by 1 to grab the info from the next case
excel_row = excel_row + 1

Loop until PMI = ""
