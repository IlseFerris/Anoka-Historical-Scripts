''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
excel_row = 2 'Starts on 2

'Calls the existing excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\ACA\12.19.2013 IMA cases from Infopac.xlsx") 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

Do

'Taking info from the spreadsheet
ELIG_type = ObjExcel.Cells(excel_row, 5).Value
MCRE_worker = ObjExcel.Cells(excel_row, 9).Value
MAXIS_case_number = ObjExcel.Cells(excel_row, 10).Value 
MAXIS_worker = ObjExcel.Cells(excel_row, 11).Value
MCRE_out_of_county = ObjExcel.Cells(excel_row, 12).Value
affiliated_case_in_county = ObjExcel.Cells(excel_row, 13).Value

If MCRE_worker = "" then stopscript

'Logic
If MCRE_out_of_county = "" then 'First it looks at MCRE being out-of-county. If it's in county it'll just give the worker their last worker.
  receiving_worker = MCRE_worker
ElseIf affiliated_case_in_county = "AFFILIATED CASE FOUND" then 'Second it looks at whether-or-not an affiliated MAXIS case can be found. If so, it sends the case to that worker.
  receiving_worker = MAXIS_worker
ElseIf elig_type = "AX" then 'Third it looks at Adult cases, and sends those to the HCAPP team
  receiving_worker = "HCAPP team"
Else 'All of the rest should go to Jodi
  receiving_worker = "X102B50"
End if

ObjExcel.Cells(excel_row, 14).Value = receiving_worker

excel_row = excel_row + 1

Loop until MCRE_worker = ""
