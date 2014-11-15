EMConnect ""

start_time = timer

EMReadScreen actv_check, 22, 2, 31
If actv_check <> "Active Caseload (ACTV)" then stopscript

EMReadScreen worker_number, 7, 21, 13


case_row = 2


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\case printout - " & worker_number & ".xlsx"  
Set objWorkbook = objExcel.Workbooks.Add() 
ObjExcel.Cells(1, 1).Value = "M#"
ObjExcel.Cells(1, 2).Value = "Name"
ObjExcel.Cells(1, 3).Value = "Next review"
ObjExcel.Cells(1, 4).Value = "Cash"
ObjExcel.Cells(1, 5).Value = "FS"
ObjExcel.Cells(1, 6).Value = "HC"


Do
  EMReadScreen last_page_check, 21, 24, 02

  EMReadScreen first_row_cash_active, 1, 7, 54
  EMReadScreen first_row_FS_active, 1, 7, 61
  EMReadScreen first_row_HC_active, 1, 7, 64
  EMReadScreen first_row_case_number, 8, 7, 12
  EMReadScreen first_row_name, 20, 7, 21
  EMReadScreen first_row_recert_month, 2, 7, 42
  ObjExcel.Cells(case_row, 1).Value = first_row_case_number
  ObjExcel.Cells(case_row, 2).Value = first_row_name
  ObjExcel.Cells(case_row, 3).Value = first_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = first_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = first_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = first_row_HC_active
  case_row = case_row + 1

  EMReadScreen second_row_FS_active, 1, 8, 61
  EMReadScreen second_row_cash_active, 1, 8, 54
  EMReadScreen second_row_HC_active, 1, 8, 64
  EMReadScreen second_row_case_number, 8, 8, 12
  EMReadScreen second_row_name, 20, 8, 21
  EMReadScreen second_row_recert_month, 2, 8, 42
  ObjExcel.Cells(case_row, 1).Value = second_row_case_number
  ObjExcel.Cells(case_row, 2).Value = second_row_name
  ObjExcel.Cells(case_row, 3).Value = second_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = second_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = second_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = second_row_HC_active
  case_row = case_row + 1

  EMReadScreen third_row_FS_active, 1, 9, 61
  EMReadScreen third_row_cash_active, 1, 9, 54
  EMReadScreen third_row_HC_active, 1, 9, 64
  EMReadScreen third_row_case_number, 8, 9, 12
  EMReadScreen third_row_name, 20, 9, 21
  EMReadScreen third_row_recert_month, 2, 9, 42
  ObjExcel.Cells(case_row, 1).Value = third_row_case_number
  ObjExcel.Cells(case_row, 2).Value = third_row_name
  ObjExcel.Cells(case_row, 3).Value = third_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = third_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = third_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = third_row_HC_active
  case_row = case_row + 1

  EMReadScreen fourth_row_FS_active, 1, 10, 61
  EMReadScreen fourth_row_cash_active, 1, 10, 54
  EMReadScreen fourth_row_HC_active, 1, 10, 64
  EMReadScreen fourth_row_case_number, 8, 10, 12
  EMReadScreen fourth_row_name, 20, 10, 21
  EMReadScreen fourth_row_recert_month, 2, 10, 42
  ObjExcel.Cells(case_row, 1).Value = fourth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = fourth_row_name
  ObjExcel.Cells(case_row, 3).Value = fourth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = fourth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = fourth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = fourth_row_HC_active
  case_row = case_row + 1

  EMReadScreen fifth_row_FS_active, 1, 11, 61
  EMReadScreen fifth_row_cash_active, 1, 11, 54
  EMReadScreen fifth_row_HC_active, 1, 11, 64
  EMReadScreen fifth_row_case_number, 8, 11, 12
  EMReadScreen fifth_row_name, 20, 11, 21
  EMReadScreen fifth_row_recert_month, 2, 11, 42
  ObjExcel.Cells(case_row, 1).Value = fifth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = fifth_row_name
  ObjExcel.Cells(case_row, 3).Value = fifth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = fifth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = fifth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = fifth_row_HC_active
  case_row = case_row + 1

  EMReadScreen sixth_row_FS_active, 1, 12, 61
  EMReadScreen sixth_row_cash_active, 1, 12, 54
  EMReadScreen sixth_row_HC_active, 1, 12, 64
  EMReadScreen sixth_row_case_number, 8, 12, 12
  EMReadScreen sixth_row_name, 20, 12, 21
  EMReadScreen sixth_row_recert_month, 2, 12, 42
  ObjExcel.Cells(case_row, 1).Value = sixth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = sixth_row_name
  ObjExcel.Cells(case_row, 3).Value = sixth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = sixth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = sixth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = sixth_row_HC_active
  case_row = case_row + 1

  EMReadScreen seventh_row_FS_active, 1, 13, 61
  EMReadScreen seventh_row_cash_active, 1, 13, 54
  EMReadScreen seventh_row_HC_active, 1, 13, 64
  EMReadScreen seventh_row_case_number, 8, 13, 12
  EMReadScreen seventh_row_name, 20, 13, 21
  EMReadScreen seventh_row_recert_month, 2, 13, 42
  ObjExcel.Cells(case_row, 1).Value = seventh_row_case_number
  ObjExcel.Cells(case_row, 2).Value = seventh_row_name
  ObjExcel.Cells(case_row, 3).Value = seventh_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = seventh_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = seventh_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = seventh_row_HC_active
  case_row = case_row + 1

  EMReadScreen eighth_row_FS_active, 1, 14, 61
  EMReadScreen eighth_row_cash_active, 1, 14, 54
  EMReadScreen eighth_row_HC_active, 1, 14, 64
  EMReadScreen eighth_row_case_number, 8, 14, 12
  EMReadScreen eighth_row_name, 20, 14, 21
  EMReadScreen eighth_row_recert_month, 2, 14, 42
  ObjExcel.Cells(case_row, 1).Value = eighth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = eighth_row_name
  ObjExcel.Cells(case_row, 3).Value = eighth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = eighth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = eighth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = eighth_row_HC_active
  case_row = case_row + 1

  EMReadScreen ninth_row_FS_active, 1, 15, 61
  EMReadScreen ninth_row_cash_active, 1, 15, 54
  EMReadScreen ninth_row_HC_active, 1, 15, 64
  EMReadScreen ninth_row_case_number, 8, 15, 12
  EMReadScreen ninth_row_name, 20, 15, 21
  EMReadScreen ninth_row_recert_month, 2, 15, 42
  ObjExcel.Cells(case_row, 1).Value = ninth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = ninth_row_name
  ObjExcel.Cells(case_row, 3).Value = ninth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = ninth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = ninth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = ninth_row_HC_active
  case_row = case_row + 1

  EMReadScreen tenth_row_FS_active, 1, 16, 61
  EMReadScreen tenth_row_cash_active, 1, 16, 54
  EMReadScreen tenth_row_HC_active, 1, 16, 64
  EMReadScreen tenth_row_case_number, 8, 16, 12
  EMReadScreen tenth_row_name, 20, 16, 21
  EMReadScreen tenth_row_recert_month, 2, 16, 42
  ObjExcel.Cells(case_row, 1).Value = tenth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = tenth_row_name
  ObjExcel.Cells(case_row, 3).Value = tenth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = tenth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = tenth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = tenth_row_HC_active
  case_row = case_row + 1

  EMReadScreen eleventh_row_FS_active, 1, 17, 61
  EMReadScreen eleventh_row_cash_active, 1, 17, 54
  EMReadScreen eleventh_row_HC_active, 1, 17, 64
  EMReadScreen eleventh_row_case_number, 8, 17, 12
  EMReadScreen eleventh_row_name, 20, 17, 21
  EMReadScreen eleventh_row_recert_month, 2, 17, 42
  ObjExcel.Cells(case_row, 1).Value = eleventh_row_case_number
  ObjExcel.Cells(case_row, 2).Value = eleventh_row_name
  ObjExcel.Cells(case_row, 3).Value = eleventh_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = eleventh_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = eleventh_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = eleventh_row_HC_active
  case_row = case_row + 1

  EMReadScreen twelfth_row_FS_active, 1, 18, 61
  EMReadScreen twelfth_row_cash_active, 1, 18, 54
  EMReadScreen twelfth_row_HC_active, 1, 18, 64
  EMReadScreen twelfth_row_case_number, 8, 18, 12
  EMReadScreen twelfth_row_name, 20, 18, 21
  EMReadScreen twelfth_row_recert_month, 2, 18, 42
  ObjExcel.Cells(case_row, 1).Value = twelfth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = twelfth_row_name
  ObjExcel.Cells(case_row, 3).Value = twelfth_row_recert_month
  ObjExcel.Cells(case_row, 4).Value = twelfth_row_cash_active
  ObjExcel.Cells(case_row, 5).Value = twelfth_row_FS_active
  ObjExcel.Cells(case_row, 6).Value = twelfth_row_HC_active
  case_row = case_row + 1

EMSendKey "<PF8>"
EMWaitReady 1, 1

Loop until last_page_check = "THIS IS THE LAST PAGE"

objExcel.Activeworkbook.SaveAs strfilename

stop_time = timer

MsgBox stop_time - start_time