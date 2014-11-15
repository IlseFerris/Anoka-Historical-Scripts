'-----------------------------------------------------------------------------------------------------------------------------------------------------
'        Script name:    Mass citizenship finder
'        Description:    Finds citizen cases
'       Target users:    Ronny Cary
'           Division:    Adult and Family
'          Author(s):    Ronny Cary
'      Working state:    purpetual development
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'     Script content:    01. 
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'       Known issues:    This script is just for the script developer to use, as it gathers large quantities of data and will lock up a user's computer.
'   Test breakpoints:    None 
'              Notes:    None
'-----------------------------------------------------------------------------------------------------------------------------------------------------

'SECTION 01

worker_number_array = array("x102233", "x10230V", "x102TRP", "x102767", "x102B42", "x102722")

EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 0, 0
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript

row = 1
col = 1
EMSearch "MAXIS", row, col
If row <> 1 then
  MsgBox "You need to run this script in the window that has MAXIS on it. Please try again."
  StopScript
End if

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

ObjExcel.Cells(1, 1).Value = "MAXIS number"
ObjExcel.Cells(1, 2).Value = "Name"
ObjExcel.Cells(1, 3).Value = "x102number"

next_excel_row_start = 2 'This sets the variable for the following.

For each worker_number in worker_number_array

excel_row = next_excel_row_start

'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWriteScreen "rept", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "11", 20, 43 'Forces a footer month/year
EMWriteScreen "12", 20, 46
EMWriteScreen "actv", 21, 70
EMSendKey "<enter>"
EMWaitReady 0, 0
EMReadScreen worker_number_check, 7, 21, 13
If worker_number_check <> worker_number then
  EMWriteScreen worker_number, 21, 13
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End if
  
'SECTION 03
  
MAXIS_row = 7 'This sets the variable for the following do...loop.
Do
  EMReadScreen last_page_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
  Do
    EMReadScreen case_number, 8, MAXIS_row, 12
    EMReadScreen client_name, 21, MAXIS_row, 21
    case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each. This is for the Excel spreadsheet, so that we aren't entering blank spaces.
    client_name = Trim(client_name)
    If case_number <> "" then 
      ObjExcel.Cells(excel_row, 1).Value = case_number   'Then it writes each into the Excel spreadsheet to be used later.
      ObjExcel.Cells(excel_row, 2).Value = client_name
      ObjExcel.Cells(excel_row, 3).Value = worker_number
      excel_row = excel_row + 1
    End if
    MAXIS_row = MAXIS_row + 1
  Loop until MAXIS_row = 19
  MAXIS_row = 7 'Setting the variable for when the do...loop restarts
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
Loop until last_page_check = "THIS IS THE LAST PAGE"

excel_row = next_excel_row_start

Do
  case_number = ObjExcel.Cells(excel_row, 1).Value
  If case_number = "" then exit do
  do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"
  EMWriteScreen "stat", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "memb", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  'It needs to check for error prone cases.
  EMReadScreen MEMB_check, 4, 2, 48
  If MEMB_check <> "MEMB" then 
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  Do
    EMReadScreen panel_info, 8, 2, 72
    memb_number = trim(left(panel_info, 2))
    memb_total = trim(right(panel_info, 2))
    EMReadScreen HH_memb, 2, 4, 33
    EMReadScreen rel_code, 2, 10, 42
    If rel_code <> "24" and rel_code <> "25" and rel_code <> "27" then HH_memb_array = trim(HH_memb_array & " " & HH_memb)
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  Loop until memb_number = memb_total
  HH_memb_array = split(HH_memb_array)

  For each HH_memb in HH_memb_array
    If len(HH_memb) = 1 then HH_memb = "0" & HH_memb

    EMWriteScreen "UNEA", 20, 71
    EMWriteScreen HH_memb, 20, 76    
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    Do
      EMReadScreen panel_info, 8, 2, 72
      UNEA_number = trim(left(panel_info, 2))
      UNEA_total = trim(right(panel_info, 2))
      EMReadScreen SSA_coding, 2, 5, 37
      If SSA_coding = "01" or SSA_coding = "02" or SSA_coding = "03" then
        EMReadScreen end_date_check, 8, 7, 68
        If end_date_check <> "__ __ __" then 
          has_SSA = False
        Else
          has_SSA = True
        End if
      End if
      If UNEA_number <> UNEA_total and has_SSA = False then 
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    Loop until (UNEA_number = UNEA_total) or has_SSA = True

    EMWriteScreen "MEMI", 20, 71
    EMWriteScreen HH_memb, 20, 76    
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    EMReadScreen citizen_check, 1, 10, 49
    If citizen_check = "Y" then
      EMReadScreen SSA_MA_citizenship_ver, 1, 11, 49
      If SSA_MA_citizenship_ver = "_" then
        If has_SSA = True then HH_memb_to_run_array = trim(HH_memb_to_run_array & " " & HH_memb & "E")
        If has_SSA = False then HH_memb_to_run_array = trim(HH_memb_to_run_array & " " & HH_memb & "R")
        EMReadScreen SSN_check, 11, 5, 29
        If SSN_check <> "   -  -    " then
          EMSendKey "<PF9>"
          EMWaitReady 0, 0
          If has_SSA = False then EMWriteScreen "R", 11, 49
          If has_SSA = True then EMWriteScreen "E", 11, 49
        End if
      Else
        HH_memb_to_run_array = trim(HH_memb_to_run_array & " " & HH_memb)
      End if
    End if
    has_SSA = False 'Clearing the variable
  Next

  ObjExcel.Cells(excel_row, 4).Value = HH_memb_to_run_array
  HH_memb_array = "" 'Resetting the variable
  HH_memb_to_run_array = "" 'Resetting the variable
  excel_row = excel_row + 1
Loop until case_number = ""

excel_row = next_excel_row_start

Do
  case_number = ObjExcel.Cells(excel_row, 1).Value
  If case_number = "" then exit do

  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
  
  EMWriteScreen "dail", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "elig", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  row = 6 'Setting to 6 as a test
  col = 1
  
  EMSearch "11 12", row, col
  
  If row <> 0 then
    EMWriteScreen "d", row, 3
    Do
      row = row + 1
      EMReadScreen next_month_check, 5, row, 11
      If next_month_check = "11 12" then EMWriteScreen "d", row, 3
    Loop until next_month_check <> "11 12"
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if
excel_row = excel_row + 1
Loop until case_number = ""

next_excel_row_start = excel_row

EMSendKey "<enter>"
EMWaitReady 0, 0

Next

MsgBox "Success!"



