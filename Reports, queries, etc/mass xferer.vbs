function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
End function


function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(x) and STAT_note_check <> "NOTE" then 
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen y, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen x, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen y, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    End if
  End if
End function


'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("h:\adult MA EX or DX (marked).xlsx") 
objExcel.DisplayAlerts = True

'NOW THE SCRIPT GRABS EACH CASE NUMBER OFF OF THE LIST
excel_row = 2 'Setting the variable for the following do...loop.
Do
  case_number = ObjExcel.Cells(excel_row, 1).Value
  if ObjExcel.Cells(excel_row, 8).Value = "" then exit do
  call navigate_to_screen("spec", "xfer")
  EMWriteScreen "x", 7, 16
  transmit
  PF9
  EMWriteScreen "x102886", 18, 61
  transmit
  excel_row = excel_row + 1
loop until ObjExcel.Cells(excel_row, 8).Value = "" 

MsgBox "Success!"