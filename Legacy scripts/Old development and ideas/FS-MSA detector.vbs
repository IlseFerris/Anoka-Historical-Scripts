EMConnect ""

EMReadScreen actv_check, 22, 2, 31
If actv_check <> "Active Caseload (ACTV)" then stopscript
excel_row = 1


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\test.xls"  
Set objWorkbook = objExcel.Workbooks.Add() 

Do
EMReadScreen last_page_check, 21, 24, 02

EMReadScreen first_row_FS_active, 1, 7, 61
EMReadScreen first_row_MSA_active, 4, 7, 51
EMReadScreen first_row_case_number, 8, 7, 12
EMReadScreen first_row_name, 20, 7, 21
If first_row_MSA_active = "MS A" and first_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = first_row_case_number
If first_row_MSA_active = "MS A" and first_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = first_row_name
If first_row_MSA_active = "MS A" and first_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen second_row_FS_active, 1, 8, 61
EMReadScreen second_row_MSA_active, 4, 8, 51
EMReadScreen second_row_case_number, 8, 8, 12
EMReadScreen second_row_name, 20, 8, 21
If second_row_MSA_active = "MS A" and second_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = second_row_case_number
If second_row_MSA_active = "MS A" and second_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = second_row_name
If second_row_MSA_active = "MS A" and second_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen third_row_FS_active, 1, 9, 61
EMReadScreen third_row_MSA_active, 4, 9, 51
EMReadScreen third_row_case_number, 8, 9, 12
EMReadScreen third_row_name, 20, 9, 21
If third_row_MSA_active = "MS A" and third_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = third_row_case_number
If third_row_MSA_active = "MS A" and third_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = third_row_name
If third_row_MSA_active = "MS A" and third_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen fourth_row_FS_active, 1, 10, 61
EMReadScreen fourth_row_MSA_active, 4, 10, 51
EMReadScreen fourth_row_case_number, 8, 10, 12
EMReadScreen fourth_row_name, 20, 10, 21
If fourth_row_MSA_active = "MS A" and fourth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = fourth_row_case_number
If fourth_row_MSA_active = "MS A" and fourth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = fourth_row_name
If fourth_row_MSA_active = "MS A" and fourth_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen fifth_row_FS_active, 1, 11, 61
EMReadScreen fifth_row_MSA_active, 4, 11, 51
EMReadScreen fifth_row_case_number, 8, 11, 12
EMReadScreen fifth_row_name, 20, 11, 21
If fifth_row_MSA_active = "MS A" and fifth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = fifth_row_case_number
If fifth_row_MSA_active = "MS A" and fifth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = fifth_row_name
If fifth_row_MSA_active = "MS A" and fifth_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen sixth_row_FS_active, 1, 12, 61
EMReadScreen sixth_row_MSA_active, 4, 12, 51
EMReadScreen sixth_row_case_number, 8, 12, 12
EMReadScreen sixth_row_name, 20, 12, 21
If sixth_row_MSA_active = "MS A" and sixth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = sixth_row_case_number
If sixth_row_MSA_active = "MS A" and sixth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = sixth_row_name
If sixth_row_MSA_active = "MS A" and sixth_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen seventh_row_FS_active, 1, 13, 61
EMReadScreen seventh_row_MSA_active, 4, 13, 51
EMReadScreen seventh_row_case_number, 8, 13, 12
EMReadScreen seventh_row_name, 20, 13, 21
If seventh_row_MSA_active = "MS A" and seventh_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = seventh_row_case_number
If seventh_row_MSA_active = "MS A" and seventh_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = seventh_row_name
If seventh_row_MSA_active = "MS A" and seventh_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen eighth_row_FS_active, 1, 14, 61
EMReadScreen eighth_row_MSA_active, 4, 14, 51
EMReadScreen eighth_row_case_number, 8, 14, 12
EMReadScreen eighth_row_name, 20, 14, 21
If eighth_row_MSA_active = "MS A" and eighth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = eighth_row_case_number
If eighth_row_MSA_active = "MS A" and eighth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = eighth_row_name
If eighth_row_MSA_active = "MS A" and eighth_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen ninth_row_FS_active, 1, 15, 61
EMReadScreen ninth_row_MSA_active, 4, 15, 51
EMReadScreen ninth_row_case_number, 8, 15, 12
EMReadScreen ninth_row_name, 20, 15, 21
If ninth_row_MSA_active = "MS A" and ninth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = ninth_row_case_number
If ninth_row_MSA_active = "MS A" and ninth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = ninth_row_name
If ninth_row_MSA_active = "MS A" and ninth_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen tenth_row_FS_active, 1, 16, 61
EMReadScreen tenth_row_MSA_active, 4, 16, 51
EMReadScreen tenth_row_case_number, 8, 16, 12
EMReadScreen tenth_row_name, 20, 16, 21
If tenth_row_MSA_active = "MS A" and tenth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = tenth_row_case_number
If tenth_row_MSA_active = "MS A" and tenth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = tenth_row_name
If tenth_row_MSA_active = "MS A" and tenth_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen eleventh_row_FS_active, 1, 17, 61
EMReadScreen eleventh_row_MSA_active, 4, 17, 51
EMReadScreen eleventh_row_case_number, 8, 17, 12
EMReadScreen eleventh_row_name, 20, 17, 21
If eleventh_row_MSA_active = "MS A" and eleventh_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = eleventh_row_case_number
If eleventh_row_MSA_active = "MS A" and eleventh_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = eleventh_row_name
If eleventh_row_MSA_active = "MS A" and eleventh_row_FS_active = "A" then excel_row = excel_row + 1
EMReadScreen twelfth_row_FS_active, 1, 18, 61
EMReadScreen twelfth_row_MSA_active, 4, 18, 51
EMReadScreen twelfth_row_case_number, 8, 18, 12
EMReadScreen twelfth_row_name, 20, 18, 21
If twelfth_row_MSA_active = "MS A" and twelfth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 1).Value = twelfth_row_case_number
If twelfth_row_MSA_active = "MS A" and twelfth_row_FS_active = "A" then ObjExcel.Cells(excel_row, 2).Value = twelfth_row_name
If twelfth_row_MSA_active = "MS A" and twelfth_row_FS_active = "A" then excel_row = excel_row + 1
EMSendKey "<PF8>"
EMWaitReady 1, 1
Loop until last_page_check = "THIS IS THE LAST PAGE"
'objExcel.Quit 