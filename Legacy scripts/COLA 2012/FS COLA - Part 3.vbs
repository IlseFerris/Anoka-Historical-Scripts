EMConnect ""

start_time = timer


x_number_input = Inputbox ("Type the x102 number you are loading up.")

caps_x_number = UCase (x_number_input)



Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open ("H:\COLA worklist for MAXIS script - " & caps_x_number & ".xlsx")

If ObjExcel.Cells(1, 8).Value <> "Cases with SS that had STAT errors" then MsgBox "This caseload does not appear to have been run using part 2. Use part 2 before proceeding."
If ObjExcel.Cells(1, 8).Value <> "Cases with SS that had STAT errors" then Stopscript


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


'The following Do...Loop reads cases that had STAT errors, and TIKLs out for them.

Do until ObjExcel.Cells(SS_cases_with_FS_STAT_errors, 8).Value = "" or ObjExcel.Cells(SS_cases_with_FS_STAT_errors, 8).Value = "        "

  case_number = ObjExcel.Cells(SS_cases_with_FS_STAT_errors, 8).Value

  EMWriteScreen "dail", 16, 43
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMWriteScreen "writ", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 0
  
  EMSetCursor 9, 3
  EMSendKey "This case could not have FS COLA approved, due to a stat error. Correct the STAT error and process manually. (TIKL auto-generated with script)" + "<enter>"
  EMWaitReady 1, 0

'This do...loop gets back to SELF.
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

SS_cases_with_FS_STAT_errors = SS_cases_with_FS_STAT_errors + 1 'This adjusts the next case number to read from the spreadsheet by one cell.

Loop

MsgBox "You've TIKLed out for the cases with STAT errors. IMPORTANT: delete the cells in Excel that had STAT errors that you weren't able to correct. Merge the remaining cells up, and then run the next script. Save the excel spreadsheet when you're done merging."


stop_time = timer

MsgBox stop_time - start_time
'objExcel.Quit
