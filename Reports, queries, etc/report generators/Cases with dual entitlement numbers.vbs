EMConnect ""

start_time = timer

EMSendKey "<attn>"
EMWaitReady 1, 1


EMReadScreen actv_check, 22, 2, 31
If actv_check <> "Active Caseload (ACTV)" then stopscript
excel_row_variable_col_1 = 2
excel_row_variable_col_2 = 2


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\test.xls"  
Set objWorkbook = objExcel.Workbooks.Add() 
ObjExcel.Cells(1, 1).Value = "M# in ACTV"



Do
  EMReadScreen last_page_check, 21, 24, 02

  EMReadScreen first_row_name, 20, 7, 21
  EMReadScreen first_row_case_number, 8, 7, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = first_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = first_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen second_row_name, 20, 8, 21
  EMReadScreen second_row_case_number, 8, 8, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = second_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = second_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen third_row_name, 20, 9, 21
  EMReadScreen third_row_case_number, 8, 9, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = third_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = third_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen fourth_row_name, 20, 10, 21
  EMReadScreen fourth_row_case_number, 8, 10, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = fourth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = fourth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen fifth_row_name, 20, 11, 21
  EMReadScreen fifth_row_case_number, 8, 11, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = fifth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = fifth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen sixth_row_name, 20, 12, 21
  EMReadScreen sixth_row_case_number, 8, 12, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = sixth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = sixth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen seventh_row_name, 20, 13, 21
  EMReadScreen seventh_row_case_number, 8, 13, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = seventh_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = seventh_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen eighth_row_name, 20, 14, 21
  EMReadScreen eighth_row_case_number, 8, 14, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = eighth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = eighth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen ninth_row_name, 20, 15, 21
  EMReadScreen ninth_row_case_number, 8, 15, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = ninth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = ninth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen tenth_row_name, 20, 16, 21
  EMReadScreen tenth_row_case_number, 8, 16, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = tenth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = tenth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen eleventh_row_name, 20, 17, 21
  EMReadScreen eleventh_row_case_number, 8, 17, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = eleventh_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = eleventh_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

  EMReadScreen twelfth_row_name, 20, 18, 21
  EMReadScreen twelfth_row_case_number, 8, 18, 12
  ObjExcel.Cells(excel_row_variable_col_1, 1).Value = twelfth_row_case_number
  ObjExcel.Cells(excel_row_variable_col_2, 2).Value = twelfth_row_name
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
  excel_row_variable_col_2 = excel_row_variable_col_2 + 1

EMSendKey "<PF8>"
EMWaitReady 1, 1

Loop until last_page_check = "THIS IS THE LAST PAGE"



excel_row_variable_col_1 = 2 'resetting the excel_row_variable_col_1 to read back the case numbers.




Do until ObjExcel.Cells(excel_row_variable_col_1, 1).value = "" or ObjExcel.Cells(excel_row_variable_col_1, 1).value = "        "

  case_number = ObjExcel.Cells(excel_row_variable_col_1, 1).value

'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

  EMWaitReady 1, 1
  EMWriteScreen "stat", 16, 43 
  EMSetCursor 18, 43
  EMSendKey "        "
  EMSetCursor 18, 43
  EMSendKey case_number
  EMSetCursor 21, 70
  EMSendKey "unea" + "<enter>"

  EMWaitReady 1, 1


'The following section checks for Error Prone and Abended cases, so they don't hang the script.
  EMReadScreen error_prone_check, 31, 2, 26
  If error_prone_check = "Error Prone Edit Summary (ERRR)" then EMSendKey "<enter>"
  If error_prone_check = "Error Prone Edit Summary (ERRR)" then EMWaitReady 1, 1
  EMReadScreen abended_check, 31, 8, 27
  If abended_check = "Note: The last STAT session was" then EMSendKey "<enter>"
  If abended_check = "Note: The last STAT session was" then EMWaitReady 1, 1


' This section checks for HH member numbers
  EMReadScreen memb_check_02, 2, 6, 3
  EMReadScreen memb_check_03, 2, 7, 3
  EMReadScreen memb_check_04, 2, 8, 3
  EMReadScreen memb_check_05, 2, 9, 3
  EMReadScreen memb_check_06, 2, 10, 3
  EMReadScreen memb_check_07, 2, 11, 3
  EMReadScreen memb_check_08, 2, 12, 3
  EMReadScreen memb_check_09, 2, 13, 3
  EMReadScreen memb_check_10, 2, 14, 3
  EMReadScreen memb_check_11, 2, 15, 3
  EMReadScreen memb_check_12, 2, 16, 3
  EMReadScreen memb_check_13, 2, 17, 3
  EMReadScreen memb_check_14, 2, 18, 3
  EMReadScreen memb_check_15, 2, 19, 3

'This section figures out how many members there are by using a "limit_reached" variable, which will be turned on when no more members are indicated.
  If memb_check_02 = "  " then HH_size = "01" 
  If memb_check_02 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_03 = "  " then HH_size = "02" 
  If memb_check_03 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_04 = "  " then HH_size = "03" 
  If memb_check_04 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_05 = "  " then HH_size = "04" 
  If memb_check_05 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_06 = "  " then HH_size = "05" 
  If memb_check_06 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_07 = "  " then HH_size = "06" 
  If memb_check_07 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_08 = "  " then HH_size = "07" 
  If memb_check_08 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_09 = "  " then HH_size = "08" 
  If memb_check_09 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_10 = "  " then HH_size = "09" 
  If memb_check_10 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_11 = "  " then HH_size = "10" 
  If memb_check_11 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_12 = "  " then HH_size = "11" 
  If memb_check_12 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_13 = "  " then HH_size = "12" 
  If memb_check_13 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_14 = "  " then HH_size = "13" 
  If memb_check_14 = "  " then limit_reached = "True"
  If limit_reached <> "True" and memb_check_15 = "  " then HH_size = "14" 
  If memb_check_15 = "  " then limit_reached = "True"

  ObjExcel.Cells(excel_row_variable_col_1, 3).value = HH_size


  limit_reached = "False" 'Resetting limit_reached variable.
  HH_size = "error" 'Resetting the variable for HH_size to detect errors.
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1


loop 

stop_time = timer

MsgBox stop_time - start_time