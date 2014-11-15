'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'RIGHT NOW THIS ALSO FILTERS TO JUST THE LEXINGTON ZIP CODES. THOSE CASES IN LEXINGTON WILL DUMP INTO AN EXCEL DOCUMENT.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'VARIABLES TO DECLARE


EMConnect ""

EMReadScreen ACTV_check, 4, 2, 48
If ACTV_check <> "ACTV" then
  MsgBox "Not on ACTV."
  StopScript
End if

row = 7
excel_row = 1

Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.


Do
  EMReadScreen case_number, 8, row, 12
  If case_number <> "        " then case_number_array = case_number_array & trim(case_number) & " "
  row = row + 1
  If row = 19 then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    row = 7
    EMReadScreen last_page_check, 4, 24, 19
  End if
Loop until last_page_check = "PAGE" or case_number = "        "

case_number_array = split(case_number_array)

For each case_number in case_number_array
  If case_number <> "" then
    Do
      EMSendKey "<PF3>"
      EMWaitReady 0, 0
      EMReadScreen SELF_check, 4, 2, 50
    Loop until SELF_check = "SELF"
  
    EMWriteScreen "stat", 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen case_number, 18, 43
    EMWriteScreen "addr", 21, 70
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    EMReadScreen ADDR_check, 4, 2, 44
    If ADDR_check <> "ADDR" then
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if

    EMReadScreen zip_code, 5, 9, 43
    If zip_code = "55014" or zip_code = "55038" then 
      lexington = True
    Else
      lexington = False
    End if

    If lexington = True then
      ObjExcel.Cells(excel_row, 1).Value = case_number
      excel_row = excel_row + 1
    End If

'    EMWriteScreen "busi", 20, 71
'    EMSendKey "<enter>"
'    EMWaitReady 0, 0

'    EMReadScreen BUSI_amt, 1, 2, 78
'    If BUSI_amt <> "0" then 
'      BUSI_total = BUSI_total + 1
'    End If
  End if
Next

Msgbox BUSI_total & " total cases with BUSI panels."