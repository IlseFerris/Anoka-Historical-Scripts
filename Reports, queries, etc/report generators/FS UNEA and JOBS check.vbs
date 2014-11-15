EMConnect ""

start_time = timer

EMReadScreen actv_check, 22, 2, 31
If actv_check <> "Active Caseload (ACTV)" then stopscript
excel_row_variable_col_1 = 2


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\test.xls"  
Set objWorkbook = objExcel.Workbooks.Add() 
ObjExcel.Cells(1, 1).Value = "M# with FS"
ObjExcel.Cells(1, 4).Value = "FS with SS"
ObjExcel.Cells(1, 6).Value = "FS with JOBS"


Do
  EMReadScreen last_page_check, 21, 24, 02

  EMReadScreen first_row_FS_active, 1, 7, 61
  EMReadScreen first_row_case_number, 8, 7, 12
  If first_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = first_row_case_number
  If first_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen second_row_FS_active, 1, 8, 61
  EMReadScreen second_row_case_number, 8, 8, 12
  If second_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = second_row_case_number
  If second_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen third_row_FS_active, 1, 9, 61
  EMReadScreen third_row_case_number, 8, 9, 12
  If third_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = third_row_case_number
  If third_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen fourth_row_FS_active, 1, 10, 61
  EMReadScreen fourth_row_case_number, 8, 10, 12
  If fourth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = fourth_row_case_number
  If fourth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen fifth_row_FS_active, 1, 11, 61
  EMReadScreen fifth_row_case_number, 8, 11, 12
  If fifth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = fifth_row_case_number
  If fifth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen sixth_row_FS_active, 1, 12, 61
  EMReadScreen sixth_row_case_number, 8, 12, 12
  If sixth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = sixth_row_case_number
  If sixth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen seventh_row_FS_active, 1, 13, 61
  EMReadScreen seventh_row_case_number, 8, 13, 12
  If seventh_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = seventh_row_case_number
  If seventh_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen eighth_row_FS_active, 1, 14, 61
  EMReadScreen eighth_row_case_number, 8, 14, 12
  If eighth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = eighth_row_case_number
  If eighth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen ninth_row_FS_active, 1, 15, 61
  EMReadScreen ninth_row_case_number, 8, 15, 12
  If ninth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = ninth_row_case_number
  If ninth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen tenth_row_FS_active, 1, 16, 61
  EMReadScreen tenth_row_case_number, 8, 16, 12
  If tenth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = tenth_row_case_number
  If tenth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen eleventh_row_FS_active, 1, 17, 61
  EMReadScreen eleventh_row_case_number, 8, 17, 12
  If eleventh_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = eleventh_row_case_number
  If eleventh_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

  EMReadScreen twelfth_row_FS_active, 1, 18, 61
  EMReadScreen twelfth_row_case_number, 8, 18, 12
  If twelfth_row_FS_active = "A" then ObjExcel.Cells(excel_row_variable_col_1, 1).Value = twelfth_row_case_number
  If twelfth_row_FS_active = "A" then excel_row_variable_col_1 = excel_row_variable_col_1 + 1

EMSendKey "<PF8>"
EMWaitReady 1, 1

Loop until last_page_check = "THIS IS THE LAST PAGE"

excel_row_variable_col_1 = 2 'resetting the excel_row_variable_col_1 to read back the case numbers.
excel_row_variable_col_4 = 2 'This variable will be used to record cases with UNEA.
excel_row_variable_col_6 = 2 'This variable will be used to record cases with JOBS.


Do

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
  EMSendKey "eats" + "<enter>"

  EMWaitReady 1, 1


'The following section checks for Error Prone and Abended cases, so they don't hang the script.
  EMReadScreen error_prone_check, 31, 2, 26
  If error_prone_check = "Error Prone Edit Summary (ERRR)" then EMSendKey "<enter>"
  If error_prone_check = "Error Prone Edit Summary (ERRR)" then EMWaitReady 1, 1
  EMReadScreen abended_check, 31, 8, 27
  If abended_check = "Note: The last STAT session was" then EMSendKey "<enter>"
  If abended_check = "Note: The last STAT session was" then EMWaitReady 1, 1

'Now it checks EATS to determine the HH size. If there are multiple HH membs in the eating group, it will store all the HH membs.
  EMReadScreen EATS_check, 1, 2, 78
  If EATS_check = 0 then multiple_HH_membs = "False"
  If EATS_check = 1 then EMReadScreen EATS_group_check, 2, 13, 43
  If EATS_group_check = "__" then multiple_HH_membs = "False"
  If EATS_check = 1 and EATS_group_check <> "__" then multiple_HH_membs = "True"
  EMReadScreen all_HH_membs_eat_with_applicant, 1, 4, 72
  If all_HH_membs_eat_with_applicant = "Y" then multiple_HH_membs = "True"

  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_02, 2, 13, 43
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_03, 2, 13, 47
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_04, 2, 13, 51
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_05, 2, 13, 55
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_06, 2, 13, 59
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_07, 2, 13, 63
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_08, 2, 13, 67
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_09, 2, 13, 71
  If multiple_HH_membs = "True" then EMReadScreen UNEA_reference_number_10, 2, 13, 75

  If multiple_HH_membs = "True" and UNEA_reference_number_02 = "__" then read_from_ref_section = "True"

  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_02, 2, 6, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_03, 2, 7, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_04, 2, 8, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_05, 2, 9, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_06, 2, 10, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_07, 2, 11, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_08, 2, 12, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_09, 2, 13, 3
  If read_from_ref_section = "True" then EMReadScreen UNEA_reference_number_10, 2, 14, 3

  If UNEA_reference_number_02 = "  " then UNEA_reference_number_02 = "__" 
  If UNEA_reference_number_03 = "  " then UNEA_reference_number_03 = "__" 
  If UNEA_reference_number_04 = "  " then UNEA_reference_number_04 = "__" 
  If UNEA_reference_number_05 = "  " then UNEA_reference_number_05 = "__" 
  If UNEA_reference_number_06 = "  " then UNEA_reference_number_06 = "__" 
  If UNEA_reference_number_07 = "  " then UNEA_reference_number_07 = "__" 
  If UNEA_reference_number_08 = "  " then UNEA_reference_number_08 = "__" 
  If UNEA_reference_number_09 = "  " then UNEA_reference_number_09 = "__" 
  If UNEA_reference_number_10 = "  " then UNEA_reference_number_10 = "__" 

  JOBS_reference_number_02 = UNEA_reference_number_02
  JOBS_reference_number_03 = UNEA_reference_number_03
  JOBS_reference_number_04 = UNEA_reference_number_04
  JOBS_reference_number_05 = UNEA_reference_number_05
  JOBS_reference_number_06 = UNEA_reference_number_06
  JOBS_reference_number_07 = UNEA_reference_number_07
  JOBS_reference_number_08 = UNEA_reference_number_08
  JOBS_reference_number_09 = UNEA_reference_number_09
  JOBS_reference_number_10 = UNEA_reference_number_10

  EMSetCursor 20, 71
  EMSendKey "unea"
  EMSetCursor 20, 76
  EMSendKey "01" + "<enter>"
  EMWaitReady 1, 0


  If multiple_HH_membs = "True" then next_ref_number = UNEA_reference_number_02 'Setting the variable for the next do...loop.
  If multiple_HH_membs = "False" then next_ref_number = "__"
  all_membs_checked = "False"
'Now it checks UNEA to determine if there is any RSDI/SSI. It starts with memb 01. If there are additional members it adds them in.
  Do
    If next_ref_number = "__" then all_membs_checked = "True"
    Do
      EMReadScreen UNEA_screen, 1, 2, 73
      EMReadScreen UNEA_total, 1, 2, 78
      EMReadScreen income_type, 2, 5, 37
      If income_type = "01" or income_type = "02" or income_type = "03" then case_has_SS = "True"
      EMSendKey "<enter>"
      EMWaitReady 1, 1
    Loop until UNEA_screen = UNEA_total

    If next_ref_number <> "__" or next_ref_number <> "  " then EMWriteScreen next_ref_number, 20, 76  
    If next_ref_number <> "__" or next_ref_number <> "  " then EMSendKey "<enter>"
    If next_ref_number <> "__" or next_ref_number <> "  " then EMWaitReady 1, 1
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_09 then next_ref_number = UNEA_reference_number_10
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_08 then next_ref_number = UNEA_reference_number_09
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_07 then next_ref_number = UNEA_reference_number_08
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_06 then next_ref_number = UNEA_reference_number_07
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_05 then next_ref_number = UNEA_reference_number_06
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_04 then next_ref_number = UNEA_reference_number_05
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_03 then next_ref_number = UNEA_reference_number_04
    If next_ref_number <> "__" and next_ref_number = UNEA_reference_number_02 then next_ref_number = UNEA_reference_number_03
  Loop until all_membs_checked = "True"

  EMWriteScreen "jobs", 20, 71
  EMWriteScreen "01", 20, 76
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  If multiple_HH_membs = "True" then next_ref_number = JOBS_reference_number_02 'Setting the variable for the next do...loop.
  If multiple_HH_membs = "False" then next_ref_number = "__"
  all_membs_checked = "False"

  Do
    If next_ref_number = "__" then all_membs_checked = "True"
    Do
      EMReadScreen JOBS_screen, 1, 2, 73
      EMReadScreen JOBS_total, 1, 2, 78
      If JOBS_total <> "0" then case_has_JOBS = "True"
      EMSendKey "<enter>"
      EMWaitReady 1, 1
    Loop until JOBS_screen = JOBS_total

    If next_ref_number <> "__" or next_ref_number <> "  " then EMWriteScreen next_ref_number, 20, 76  
    If next_ref_number <> "__" or next_ref_number <> "  " then EMSendKey "<enter>"
    If next_ref_number <> "__" or next_ref_number <> "  " then EMWaitReady 1, 1
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_09 then next_ref_number = JOBS_reference_number_10
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_08 then next_ref_number = JOBS_reference_number_09
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_07 then next_ref_number = JOBS_reference_number_08
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_06 then next_ref_number = JOBS_reference_number_07
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_05 then next_ref_number = JOBS_reference_number_06
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_04 then next_ref_number = JOBS_reference_number_05
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_03 then next_ref_number = JOBS_reference_number_04
    If next_ref_number <> "__" and next_ref_number = JOBS_reference_number_02 then next_ref_number = JOBS_reference_number_03
  Loop until all_membs_checked = "True"


  If case_has_SS = "True" then ObjExcel.Cells(excel_row_variable_col_4, 4).Value = case_number
  If case_has_SS = "True" then excel_row_variable_col_4 = excel_row_variable_col_4 + 1


  If case_has_JOBS = "True" then ObjExcel.Cells(excel_row_variable_col_6, 6).Value = case_number
  If case_has_JOBS = "True" then excel_row_variable_col_6 = excel_row_variable_col_6 + 1

  case_has_SS = "False" 'resetting the variable before the next instance of the do...loop runs
  case_has_JOBS = "False" 'resetting the variable before the next instance of the do...loop runs

  excel_row_variable_col_1 = excel_row_variable_col_1 + 1

loop until ObjExcel.Cells(excel_row_variable_col_1, 1).value = ""

stop_time = timer

MsgBox stop_time - start_time