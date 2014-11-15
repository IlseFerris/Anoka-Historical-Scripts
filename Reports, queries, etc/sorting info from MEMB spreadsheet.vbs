'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True                                 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open("H:\---report.xlsx") 
objExcel.DisplayAlerts = True                           'Set this to false to make alerts go away. This is necessary in production.

excel_row = 2

Do
  abawd_MEMB_array = ObjExcel.Cells(excel_row, 3).Value
  If abawd_MEMB_array <> "" then
    abawd_MEMB_array = split(abawd_MEMB_array)
    For each MEMB in abawd_MEMB_array
      If MEMB <> "" then
        DISA_MEMBs = ObjExcel.Cells(excel_row, 4).Value
        other_aged_MEMBs = ObjExcel.Cells(excel_row, 5).Value
        If instr(DISA_MEMBs, MEMB) <> 0 or instr(other_aged_MEMBs, MEMB) <> 0 then ObjExcel.Cells(excel_row, 7).Value = ObjExcel.Cells(excel_row, 7).Value & " " & MEMB
      End if
    Next
  area_to_trim = trim(ObjExcel.Cells(excel_row, 7).Value)
  If right(area_to_trim, 1) = "," then ObjExcel.Cells(excel_row, 7).Value = "MEMB " & left(area_to_trim, len(area_to_trim) - 1)
  End if
  excel_row = excel_row + 1
'Loop until excel_row = 25
Loop until excel_row = 8380
