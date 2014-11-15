Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open ("Q:\Blue Zone Scripts\x102 list.xlsx")

excel_row = 2

MsgBox objExcel.Cells(excel_row,1).Value


objExcel.quit