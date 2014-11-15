''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
excel_row = 1 'Starts on 2
new_receiving_worker = "X102880" 'Starts here, we'll rotate through

'Calls the existing excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\ACA\02.24.2014 Randi's cases to move.xlsx") 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

Do

'Taking info from the spreadsheet
worker_receiving = ObjExcel.Cells(excel_row, 14).Value

If worker_receiving = "" then stopscript

'Logic
If worker_receiving = "X102268" then 
  ObjExcel.Cells(excel_row, 14).Value = new_receiving_worker
  If new_receiving_worker = "X102880" then 
    new_receiving_worker = "X102598"
  ElseIf new_receiving_worker = "X102598" then 
    new_receiving_worker = "X102757"
  ElseIf new_receiving_worker = "X102757" then 
    new_receiving_worker = "X102932"
  ElseIf new_receiving_worker = "X102932" then 
    new_receiving_worker = "X102880"
  End if
End if



excel_row = excel_row + 1



Loop until worker_receiving = ""

