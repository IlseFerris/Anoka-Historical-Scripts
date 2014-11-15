'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

x102_array = array("293", "692", "30V", "B83", "752", "4SS", "C02", "756", "395", "4RS", "631", "104", "GMZ", "B98", "SEC", "SAC", "757", "880", "598", "4BL", "268", "B93", "932", "989", "949", "B64", "BED", "628", "769", "524", "4DK", "750", "4SZ", "950", "742", "4SW", "TRP", "619", "4AS", "894", "987", "4SY", "C08", "624", "200", "616", "4F9", "925", "4BM", "294", "B55", "B48", "A75", "4BV", "B36", "RLM", "B52", "707", "A84", "674", "106", "231", "A18", "733", "962", "213", "A44", "902", "223", "944", "234", "B50", "B97", "618", "225", "C06", "C04", "869", "4SL", "TLP", "C07", "C05", "4ES", "895", "4SX", "978", "222", "107", "767", "722", "247", "4AF", "119", "233", "112", "111", "122", "125", "110", "B20", "872", "117", "643", "967", "630", "518", "118", "4RJ", "601", "116", "928", "120", "113", "114", "797", "121", "123", "124", "126", "282", "127")

EMConnect ""

start_time = timer

call navigate_to_screen("REPT", "ACTV")
excel_row_variable_col_1 = 2


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\test.xlsx"  
Set objWorkbook = objExcel.Workbooks.Add() 
ObjExcel.Cells(1, 1).Value = "x102"
ObjExcel.Cells(1, 2).Value = "M# on SNAP"
ObjExcel.Cells(1, 3).Value = "MEMBs with 09 code on WREG"

For each x102_number in x102_array
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
  EMSendKey "ACTV" & "<enter>"
  EMWaitReady 0, 0
  EMWriteScreen x102_number, 21, 17
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSetCursor 21, 13
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
  EMReadScreen supervisor_name, 20, 22, 16
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  EMReadScreen ACTV_amt_check, 6, 3, 74
  If ACTV_amt_check <> "0 Of 0" then 'skips workers with no active cases
    Do
      MAXIS_row = 7
      Do
        EMReadScreen FS_status_code, 1, MAXIS_row, 61
        EMReadScreen case_number, 8, MAXIS_row, 12
        If FS_status_code = "A" then
          ObjExcel.Cells(excel_row_variable_col_1, 1).Value = x102_number
          ObjExcel.Cells(excel_row_variable_col_1, 2).Value = case_number
          ObjExcel.Cells(excel_row_variable_col_1, 5).Value = supervisor_name
          excel_row_variable_col_1 = excel_row_variable_col_1 + 1
        End if
        MAXIS_row = MAXIS_row + 1
      Loop until case_number = "        "
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
      EMReadScreen last_page_check, 21, 24, 02
    Loop until last_page_check = "THIS IS THE LAST PAGE"
  End if
Next

excel_row_variable_col_1 = 2

Do
  case_number = ObjExcel.Cells(excel_row_variable_col_1, 2).Value 
  back_to_self
  EMWriteScreen "stat", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "wreg", 21, 70
  EMWriteScreen "01", 21, 75
  transmit
  HH_memb_row = 5 'Setting up variable for the following do...loop. 
  Do
    EMReadScreen ABAWD_status, 2, 13, 50
    If ABAWD_status = "09" then
      EMReadScreen HH_ref_nbr, 2, 4, 33
      ObjExcel.Cells(excel_row_variable_col_1, 3).Value = ObjExcel.Cells(excel_row_variable_col_1, 3).Value & HH_ref_nbr & ", "
    End if
    HH_memb_row = HH_memb_row + 1
    EMReadScreen next_HH_memb, 2, HH_memb_row, 3
    If next_HH_memb <> "  " then
      EMWriteScreen next_HH_memb, 20, 76
      transmit
    End if
  Loop until next_HH_memb = "  "
  excel_row_variable_col_1 = excel_row_variable_col_1 + 1
Loop until case_number = ""

stop_time = timer
MsgBox stop_time - start_time