'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True                                'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True                          'Set this to false to make alerts go away. This is necessary in production.

'Assigning values to the Excel spreadsheet.
excel_row = 1

Do
	case_number = cint(rnd * 100)
	ObjExcel.Cells(excel_row, 1).Value = case_number
	case_number_total = case_number_total + case_number
	excel_row = excel_row + 1
Loop until excel_row = 900

MsgBox case_number_total


'Now it creates a word document with all active claims in it.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objselection.typetext case_number_total & " is the ''random'' number VBscript generated."
objselection.TypeParagraph()


