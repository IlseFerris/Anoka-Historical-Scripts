''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
excel_row = 2 'Starts on 2
new_receiving_worker = "X102880" 'Starts here, we'll rotate through

'Calls the existing excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\Randi's cases to XFER to HC team.xlsx") 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

Do

'Taking info from the spreadsheet
old_worker = ObjExcel.Cells(excel_row, 2).Value

If old_worker = "" then exit do

'Logic

ObjExcel.Cells(excel_row, 7).Value = new_receiving_worker
If new_receiving_worker = "X102880" then 
  new_receiving_worker = "X102598"
ElseIf new_receiving_worker = "X102598" then 
  new_receiving_worker = "X102757"
ElseIf new_receiving_worker = "X102757" then 
  new_receiving_worker = "X102932"
ElseIf new_receiving_worker = "X102932" then 
  new_receiving_worker = "X102880"
End if


excel_row = excel_row + 1



Loop until old_worker = ""

'DECLARING VARIABLES
excel_row = 31 'Starts on 2 normally

EMConnect ""

Do

  'Taking info from the spreadsheet
  MCRE_case_number = ObjExcel.Cells(excel_row, 5).Value
  worker_receiving = ObjExcel.Cells(excel_row, 7).Value
  
  If worker_receiving = "" then 
    MsgBox "completed!"
    stopscript
  End if
  
  'First it checks to make sure we're on RKEY. If we aren't on RKEY the script will stop.
  EMReadScreen RKEY_check, 4, 1, 52
  If RKEY_check <> "RKEY" then
    MsgBox "RKEY not found. Script will stop. Excel row = " & excel_row & ". JOT THIS DOWN!"
    stopscript
  End if

  'Getting to RCIN for the case and transferring
  If MCRE_case_number <> "" then
    EMWriteScreen "T", 2, 19
    EMWriteScreen MCRE_case_number, 9, 19
    transmit

    EMWriteScreen "RCIN", 1, 8
    transmit

    EMWriteScreen worker_receiving, 2, 46
    PF9

    EMWriteScreen "Y", 2, 73
    PF3
    PF3
  End if

  ObjExcel.Cells(excel_row, 8).Value = "YES"
  excel_row = excel_row + 1

Loop until worker_receiving = ""

