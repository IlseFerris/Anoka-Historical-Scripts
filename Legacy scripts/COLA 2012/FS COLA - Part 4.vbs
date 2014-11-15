EMConnect ""
start_time = timer

x_number_input = Inputbox ("Type the x102 number you are loading up.")

caps_x_number = UCase (x_number_input)

unapproved_PA_check = "" 'This sets the unapproved_PA_check variable.


Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open ("H:\COLA worklist for MAXIS script - " & caps_x_number & ".xlsx")

ObjExcel.Cells(1, 12).Value = "Ineligible cases with SS"
ObjExcel.Cells(1, 14).Value = "RSDI-only MSA"



SS_case_to_run = 2 'Setting up the initial variable for the excel spreadsheet to operate. It will pick this cell out.
SS_ineligible_case = 2 'Setting up the ineligible SS case variable.
RSDI_only_MSA_case = 2 'Setting up the RSDI-only MSA case variable.

'This do...loop gets back to SELF.
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMReadScreen footer_month_check, 5, 20, 43
If footer_month_check <> "01 12" then MsgBox "Wrong footer month"
If footer_month_check <> "01 12" then Stopscript


'The following Do...Loop reads cases that are ineligible, and adds them to the spreadsheet.

Do until ObjExcel.Cells(SS_case_to_run, 4).Value = "" or ObjExcel.Cells(SS_case_to_run, 4).Value = "        "

  case_number = ObjExcel.Cells(SS_case_to_run, 4).Value

  EMWriteScreen "elig", 16, 43
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMWriteScreen "fs", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 0

  EMReadScreen member_01_elig_check, 8, 7, 57
  If member_01_elig_check <> "ELIGIBLE" then ObjExcel.Cells(SS_ineligible_case, 12).Value = case_number
  If member_01_elig_check <> "ELIGIBLE" then SS_ineligible_case = SS_ineligible_case + 1 'This sets up the next cell to be written.
  If member_01_elig_check = "ELIGIBLE" then EMSendKey "<enter>"
  If member_01_elig_check = "ELIGIBLE" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" then EMSendKey "<enter>" 
  If member_01_elig_check = "ELIGIBLE" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" then EMReadScreen PA_grants_check, 8, 10, 33
  If member_01_elig_check = "ELIGIBLE" then EMReadScreen SSI_check, 8, 12, 33
  If member_01_elig_check = "ELIGIBLE" and (PA_grants_check <> "        " and SSI_check = "        ") then RSDI_only_MSA = "True"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA = "True" then ObjExcel.Cells(RSDI_only_MSA_case, 14).Value = case_number
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA = "True" then RSDI_only_MSA_case = RSDI_only_MSA_case + 1
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "<enter>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "<enter>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMReadScreen new_benefit_amount, 6, 13, 75
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSetCursor 19, 70
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "app" + "<enter>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and PA_grants_check <> "        " and RSDI_only_MSA <> "True" then EMReadScreen unapproved_PA_check, 39, 5, 25 
  If member_01_elig_check = "ELIGIBLE" and PA_grants_check <> "        " and RSDI_only_MSA <> "True" and unapproved_PA_check = "An unapproved PA amount was used in the" then EMWriteScreen "y", 9, 63
  If member_01_elig_check = "ELIGIBLE" and PA_grants_check <> "        " and RSDI_only_MSA <> "True" and unapproved_PA_check = "An unapproved PA amount was used in the" then EMSendKey "<enter>"
  If member_01_elig_check = "ELIGIBLE" and PA_grants_check <> "        " and RSDI_only_MSA <> "True" and unapproved_PA_check = "An unapproved PA amount was used in the" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and PA_grants_check <> "        " and RSDI_only_MSA <> "True" and unapproved_PA_check <> "An unapproved PA amount was used in the" then Msgbox "Check this approval screen for possible data entry issues before proceeding."
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "<enter>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "<PF4>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "<PF9>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "~~~COLA: 2012 FS approved using automated script~~~" + "<newline>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "* STAT/UNEA's PIC has been automatically updated, and new FS results approved." + "<newline>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "* New benefit amount: $" + new_benefit_amount + "<newline>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "---" + "<newline>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "Ronny C.," + "<newline>"
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMSendKey "  processing COLA on behalf of worker, via automated script."
  If member_01_elig_check = "ELIGIBLE" and RSDI_only_MSA <> "True" then EMWaitReady 1, 0

  RSDI_only_MSA = "False" 'Resetting the variable

'This do...loop gets back to SELF.
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

SS_case_to_run = SS_case_to_run + 1 'This adjusts the next case number to read from the spreadsheet by one cell.
unapproved_PA_check = " " 'This resets the unapproved_PA_check variable.

Loop


stop_time = timer

MsgBox stop_time - start_time
'objExcel.Quit
