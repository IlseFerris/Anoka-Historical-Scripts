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

'Connects to MAXIS
EMConnect ""

Do

  'Checks for PERS
  EMReadScreen PERS_check, 4, 2, 47
  If PERS_check <> "PERS" then
    MsgBox "PERS NOT FOUND. ROW IS " & excel_row & ". STOPPING SCRIPT."
    StopScript
  End if
  
  'Grabbing case number from spreadsheet. It should only look up cases for the instances where a case number is indicated.
  IMA_case_number = ObjExcel.Cells(excel_row, 1).Value
  If IMA_case_number <> "" then

    'Grabbing worker info from MCRE list. It should only look up cases that aren't in Anoka County.
    MCRE_worker = ObjExcel.Cells(excel_row, 9).Value

    If left(MCRE_worker, 4) <> "X102" then

      'Taking PMI from the spreadsheet
      PMI = ObjExcel.Cells(excel_row, 2).Value 
      
      'If PMI is blank the loop should exit
      If PMI = "" then exit do
      
      'Navigating to PERS
      call navigate_to_screen("pers", "____")
      
      'Loading the PMI into PERS, and pressing transmit
      EMWriteScreen PMI, 15, 36
      transmit
      
      'Navigating to next screen
      EMWriteScreen "X", 8, 5
      transmit
      
      'Checks for the "NO MAXIS CASE EXISTS" error
      EMReadScreen no_MAXIS_case_check, 13, 24, 38
      If no_MAXIS_case_check = "NO MAXIS CASE" then
        ObjExcel.Cells(excel_row, 10).Value = "NONE FOUND FOR PMI"
        PF3
      Else
      
        'Loading the MAXIS case number into the spreadsheet, by searching for the "Y" code for "applicant"
        row = 9 'Setting this high to avoid false results when a "Y" is the middle initial
        col = 1
        EMSearch " Y  ", row, col
        
        'The remainder is in an if statement, because if we can't see the code chances are it's non existant, and we should move on to the next one
        If row <> 0 then 
          EMReadScreen MAXIS_case_number, 8, row, 6
          ObjExcel.Cells(excel_row, 10).Value = MAXIS_case_number
          
          'Reading the worker number and loading into the spreadsheet
          EMReadScreen MAXIS_worker, 7, row, 71
          ObjExcel.Cells(excel_row, 11).Value = MAXIS_worker
        Else
          ObjExcel.Cells(excel_row, 10).Value = "PMI not primary on a known case"
        End if
          
        'Navigating back to PERS
        PF3
        PF3
      End if

    End if

  End if
    
  'Raising the excel_row variable by 1 to grab the info from the next case
  excel_row = excel_row + 1

Loop until MCRE_worker = ""
