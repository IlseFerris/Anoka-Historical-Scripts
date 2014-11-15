Set ObjExcel = CreateObject("Excel.application")
ObjExcel.Application.Workbooks.Open "h:\Adult Privileged case finding.xlsx"
ObjExcel.Application.Visible = True


'SECTION 06
excel_row = 2 'This sets the variable for the following do...loop.

Do
  If ObjExcel.Cells(excel_row, 4) = "no match" then
   'This Do...loop gets back to SELF
    do
      EMSendKey "<PF3>"
      EMWaitReady 1, 1
      EMReadScreen SELF_check, 27, 2, 28
    loop until SELF_check = "Select Function Menu (SELF)"
    EMWriteScreen "stat", 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen ObjExcel.Cells(excel_row, 1).Value, 18, 43
    EMWriteScreen "memb", 21, 70
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 4, 2, 50
    If SELF_check = "SELF" then 
      EMWaitReady 1, 5
      EMReadScreen SELF_check, 4, 2, 50
      If SELF_check = "SELF" then ObjExcel.Cells(excel_row, 5).Value = "Privileged"
    End if
  End if
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""