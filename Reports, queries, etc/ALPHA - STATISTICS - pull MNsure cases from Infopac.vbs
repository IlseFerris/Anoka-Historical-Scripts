''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
INFOPAC_row = 12 'Sets to 12 because the first page has info start on 12
excel_row = 2 'Starts on 2

'Calls a blank excel sheet
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

'Adds headers to excel
ObjExcel.Cells(1, 1).Value = "IMA CASE #"
ObjExcel.Cells(1, 2).Value = "PMI"
ObjExcel.Cells(1, 3).Value = "RECIP NAME"
ObjExcel.Cells(1, 4).Value = "CASE ADDRESS"
ObjExcel.Cells(1, 5).Value = "ELIG TYPE"

'Connects to INFOPAC
EMConnect ""

Do
  'Finds the row where IMA CASE # starts. They apparently don't all start on the same row.
  row = 1
  col = 1
  EMSearch "IMA CASE #", row, col
  INFOPAC_row = row + 2 

Do

  'Reads the case info. Has to navigate sideways to each page (PF11)
  EMReadScreen IMA_case_number, 8, INFOPAC_row, 4
  EMReadScreen IMA_PMI, 8, INFOPAC_row, 17
  If IMA_PMI = "        " then exit do 'Exits this part of the loop if there is no IMA_PMI

  
  EMReadScreen IMA_recip_name, 28, INFOPAC_row, 28
  PF11
  EMReadScreen IMA_case_address, 50, INFOPAC_row, 22
  PF11
  EMReadScreen IMA_elig_type, 2, INFOPAC_row, 46

  'Writes the info to the excel sheet
  ObjExcel.Cells(excel_row, 1).Value = trim("'" & IMA_case_number)
  ObjExcel.Cells(excel_row, 2).Value = trim("'" & IMA_PMI)
  ObjExcel.Cells(excel_row, 3).Value = trim(IMA_recip_name)
  ObjExcel.Cells(excel_row, 4).Value = trim(IMA_case_address)
  ObjExcel.Cells(excel_row, 5).Value = trim(IMA_elig_type)

  'Navigates back to the beginning
  PF10
  PF10

  'Adds one to the variable so that it pulls the next excel row and the next infopac row
  excel_row = excel_row + 1
  INFOPAC_row = INFOPAC_row + 1

  If INFOPAC_row = 25 then 'There's only 24 rows, if we get to 25 we have to move to the next page
    PF8
    INFOPAC_row = 4 'Because the second part of the page comes in on the fourth row
  End if

Loop until trim(IMA_PMI) = ""

'Jumps to next page (screen in INFOPAC)
PF8

'Resets the INFOPAC row to 

'Checks to see if we're on the last page. If we are on the last page the report will stop
EMReadScreen end_of_report_pages_check, 19, 1, 8

Loop until end_of_report_pages_check = "END OF REPORT PAGES"
