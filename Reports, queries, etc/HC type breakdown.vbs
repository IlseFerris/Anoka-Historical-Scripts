'SECTION 01

worker_number_array = array("x102b83")

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
ObjExcel.Cells(1, 4).Value = "HC status"
ObjExcel.Cells(1, 5).Value = "FS status"
ObjExcel.Cells(1, 6).Value = "cash status"

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
      EMReadScreen HC_status, 1, MAXIS_row, 64
      EMReadScreen SNAP_status, 1, MAXIS_row, 61
      EMReadScreen cash_status, 9, MAXIS_row, 51
      case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each. This is for the Excel spreadsheet, so that we aren't entering blank spaces.
      client_name = Trim(client_name)
      If case_number <> "" then 
        ObjExcel.Cells(excel_row, 1).Value = case_number   'Then it writes each into the Excel spreadsheet to be used later.
        ObjExcel.Cells(excel_row, 2).Value = client_name
        ObjExcel.Cells(excel_row, 3).Value = worker_number
        ObjExcel.Cells(excel_row, 4).Value = HC_status
        ObjExcel.Cells(excel_row, 5).Value = SNAP_status
        ObjExcel.Cells(excel_row, 6).Value = cash_status
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
    HC_status = ObjExcel.Cells(excel_row, 4).Value
    If HC_status = "A" then
      do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 27, 2, 28
      loop until SELF_check = "Select Function Menu (SELF)"
      EMWriteScreen "elig", 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen "hc__", 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0

      EMReadScreen MA_check, 2, 8, 31
      If MA_check = "MA" then
        EMWriteScreen "x", 8, 29
        EMSendKey "<enter>"
        EMWaitReady 0, 0

        row = 1
        col = 1
        EMSearch "DX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "DX"
        row = 1
        col = 1
        EMSearch "EX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "EX"
      End if
    End if

    excel_row = excel_row + 1
  Loop until case_number = ""
  
  next_excel_row_start = excel_row
  
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
Next

MsgBox "Success!"



