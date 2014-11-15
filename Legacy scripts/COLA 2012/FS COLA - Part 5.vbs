EMConnect ""

start_time = timer


x_number_input = Inputbox ("Type the x102 number you are loading up.")

caps_x_number = UCase (x_number_input)



Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open ("H:\COLA worklist for MAXIS script - " & caps_x_number & ".xlsx")




ineligible_cases = 2 'Setting up the initial variable for the excel spreadsheet to operate. It will pick this cell out for ineligible cases.
RSDI_only_MSA_cases = 2 'Setting up the initial variable for the RSDI only MSA cases to be TIKLed.
SS_cases_not_updated_by_COLA = 2 'Setting up the initial variable for the SS cases that did not receive a COLA update.

'This do...loop gets back to SELF.
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMReadScreen footer_month_check, 5, 20, 43
If footer_month_check <> "01 12" then MsgBox "Wrong footer month"
If footer_month_check <> "01 12" then Stopscript


'The following Do...Loop reads cases that were ineligible, and TIKLs out for them.

Do until ObjExcel.Cells(ineligible_cases, 12).Value = "" or ObjExcel.Cells(ineligible_cases, 12).Value = "        "

  case_number = ObjExcel.Cells(ineligible_cases, 12).Value

  EMWriteScreen "dail", 16, 43
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMWriteScreen "writ", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 0
  
  EMSetCursor 9, 3
  EMSendKey "This case could not have FS COLA approved, because it came up ineligible for MEMB 01. Check the eligibility results and process manually. (TIKL auto-generated with script)" + "<enter>"
  EMWaitReady 1, 0

'This do...loop gets back to SELF.
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

ineligible_cases = ineligible_cases + 1 'This adjusts the next case number to read from the spreadsheet by one cell.

Loop

'The following Do...Loop reads cases that were RSDI-MSA only, and TIKLs out for them.

Do until ObjExcel.Cells(RSDI_only_MSA_cases, 14).Value = "" or ObjExcel.Cells(RSDI_only_MSA_cases, 14).Value = "        "

  case_number = ObjExcel.Cells(RSDI_only_MSA_cases, 14).Value

  EMWriteScreen "dail", 16, 43
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMWriteScreen "writ", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 0
  
  EMSetCursor 9, 3
  EMSendKey "This case could not have FS COLA approved, because it has PA income and no SSI. It may be RSDI-only MSA. Check the eligibility results and process manually. See Erica’s instructions, emailed last week. (TIKL auto-generated with script)" + "<enter>"
  EMWaitReady 1, 0

'This do...loop gets back to SELF.
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

RSDI_only_MSA_cases = RSDI_only_MSA_cases + 1 'This adjusts the next case number to read from the spreadsheet by one cell.

Loop

'The following Do...Loop reads cases that were not updated automatically by the state's COLA, and TIKLs out for them.

Do until ObjExcel.Cells(SS_cases_not_updated_by_COLA, 6).Value = "" or ObjExcel.Cells(SS_cases_not_updated_by_COLA, 6).Value = "        "

  case_number = ObjExcel.Cells(SS_cases_not_updated_by_COLA, 6).Value

  EMWriteScreen "dail", 16, 43
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMWriteScreen "writ", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 0
  
  EMSetCursor 9, 3
  EMSendKey "This case could not have FS COLA approved, because the UNEA panel was not updated by the state. Possibly dual elig for RSDI. Check UNEA and email Christa any claim numbers that don’t show a prospective date of 01/**/12." + "<enter>"
  EMWaitReady 1, 0

'This do...loop gets back to SELF.
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

SS_cases_not_updated_by_COLA = SS_cases_not_updated_by_COLA + 1 'This adjusts the next case number to read from the spreadsheet by one cell.

Loop



stop_time = timer

MsgBox stop_time - start_time

