'This was really more of a temporary script, but I've kept it in this folder just in case it's needed again with all of the IMA hullabaloo.

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
excel_row = 2 'Starts on 2 normally

'Calls the existing excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\ACA\02.24.2014 Randi's cases to move.xlsx") 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

'68

Do

  'Taking info from the spreadsheet
  IMA_case_number = ObjExcel.Cells(excel_row, 1).Value
  worker_receiving = ObjExcel.Cells(excel_row, 14).Value
  
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
  If IMA_Case_number <> "" then
    EMWriteScreen "T", 2, 19
    EMWriteScreen IMA_case_number, 9, 19
    transmit

    EMWriteScreen "RCIN", 1, 8
    transmit

    EMWriteScreen worker_receiving, 2, 46
    PF3
  End if

  excel_row = excel_row + 1

Loop until worker_receiving = ""

