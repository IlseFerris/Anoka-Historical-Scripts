EMConnect ""

start_time = timer


x_number_input = Inputbox ("Type the x102 number you are loading up.")

caps_x_number = UCase (x_number_input)



Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open ("H:\COLA worklist for MAXIS script - " & caps_x_number & ".xlsx")

ObjExcel.Cells(1, 8).Value = "Cases with SS that had STAT errors"

SS_case_to_check = 2 'Setting up the initial variable for the excel spreadsheet to operate. It will pick this cell out.
SS_cases_with_FS_STAT_errors = 2 'Setting up the initial variable for the FS STAT errors to be documented.

'This do...loop gets back to SELF.
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMReadScreen footer_month_check, 5, 20, 43
If footer_month_check <> "01 12" then MsgBox "Wrong footer month"
If footer_month_check <> "01 12" then Stopscript


'The following Do...Loop checks each case for STAT errors. It stores the information in the Excel spreadsheet.

Do until ObjExcel.Cells(SS_case_to_check, 4).Value = "" or ObjExcel.Cells(SS_case_to_check, 4).Value = "        "

  case_number = ObjExcel.Cells(SS_case_to_check, 4).Value

  EMWriteScreen "stat", 16, 43
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMSendKey "<enter>"
  EMWaitReady 1, 0

'The following section checks for Error Prone and Abended cases, so they don't hang the script.
  EMReadScreen error_prone_check, 31, 2, 26
  If error_prone_check = "Error Prone Edit Summary (ERRR)" then EMSendKey "<enter>"
  If error_prone_check = "Error Prone Edit Summary (ERRR)" then EMWaitReady 1, 1
  EMReadScreen abended_check, 31, 8, 27
  If abended_check = "Note: The last STAT session was" then EMSendKey "<enter>"
  If abended_check = "Note: The last STAT session was" then EMWaitReady 1, 1



     row = 1
     col = 1

  EMSearch "FS HAS BEEN INHIBITED", row, col 'This searches for cases that STAT errored out, and did not go through background.
  If row = 0 then EMSearch "BACKGROUND HAS BEEN ABORTED BEFORE HOUSEHOLD COMP FOR FS", row, col
  If row <> "0" then ObjExcel.Cells(SS_cases_with_FS_STAT_errors, 8).Value = case_number 'This writes cases that have STAT errors in the spreadsheet
  If row <> "0" then SS_cases_with_FS_STAT_errors = SS_cases_with_FS_STAT_errors + 1 'This adjusts the row so that the next case with an error is not overwriting the previous one.

     row = 1
     col = 1

  EMSearch "BACKGROUND HAS BEEN ABORTED BEFORE HOUSEHOLD COMP FOR FS", row, col 'This searches for cases that have HH comp changes, and did not go through background.
  If row <> "0" then ObjExcel.Cells(SS_cases_with_FS_STAT_errors, 8).Value = case_number 'This writes cases that have STAT errors in the spreadsheet
  If row <> "0" then SS_cases_with_FS_STAT_errors = SS_cases_with_FS_STAT_errors + 1 'This adjusts the row so that the next case with an error is not overwriting the previous one.

     row = 1
     col = 1

  EMSearch "A Background transaction", row, col 'This searches for cases that are still in background.
  If row <> "0" then ObjExcel.Cells(SS_cases_with_FS_STAT_errors, 8).Value = case_number 'This writes cases that have STAT errors in the spreadsheet
  If row <> "0" then SS_cases_with_FS_STAT_errors = SS_cases_with_FS_STAT_errors + 1 'This adjusts the row so that the next case with an error is not overwriting the previous one.  

'This do...loop gets back to SELF.
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"



SS_case_to_check = SS_case_to_check + 1 'This adjusts the next case number to read from the spreadsheet by one cell.

Loop

MsgBox "You can clear the cases with STAT-errors, and run this script again, as long as you don't save the spreadsheet. Or, running part three will cause the cases to be unincluded in the COLA."


stop_time = timer

MsgBox stop_time - start_time
