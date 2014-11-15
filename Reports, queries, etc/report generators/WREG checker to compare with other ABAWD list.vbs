'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

x102_array = array("293", "692", "30V", "B83", "752", "4SS", "C02", "756", "395", "4RS", "631", "104", "GMZ", "B98", "SEC", "SAC", "757", "880", "598", "4BL", "268", "B93", "932", "989", "949", "B64", "BED", "628", "769", "524", "4DK", "750", "4SZ", "950", "742", "4SW", "TRP", "619", "4AS", "894", "987", "4SY", "C08", "624", "200", "616", "4F9", "925", "4BM", "294", "B55", "B48", "A75", "4BV", "B36", "RLM", "B52", "707", "A84", "674", "106", "231", "A18", "733", "962", "213", "A44", "902", "223", "944", "234", "B50", "B97", "618", "225", "C06", "C04", "869", "4SL", "TLP", "C07", "C05", "4ES", "895", "4SX", "978", "222", "107", "767", "722", "247", "4AF", "119", "233", "112", "111", "122", "125", "110", "B20", "872", "117", "643", "967", "630", "518", "118", "4RJ", "601", "116", "928", "120", "113", "114", "797", "121", "123", "124", "126", "282", "127")

'x102_array = array("b42", "293")

EMConnect ""

start_time = timer

call navigate_to_screen("REPT", "ACTV")
excel_row_variable_col_1 = 2


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\~~~scanning detection in progress.xlsx"  
Set objWorkbook = objExcel.Workbooks.Open(strFileName) 

Sub blocked_out

objExcel.worksheets.Add().Name = "Persons Affected"

ObjExcel.Cells(1, 1).Value = "x102"
ObjExcel.Cells(1, 2).Value = "First Name"
ObjExcel.Cells(1, 3).Value = "Last Name"
ObjExcel.Cells(1, 4).Value = "PMI"
ObjExcel.Cells(1, 5).Value = "MAXIS case number"
ObjExcel.Cells(1, 6).Value = "Date of Birth"
ObjExcel.Cells(1, 7).Value = "Program ID"
ObjExcel.Cells(1, 8).Value = "ELIG Begin Date"
ObjExcel.Cells(1, 9).Value = "MAXIS WREG ABAWD status"
ObjExcel.Cells(1, 10).Value = "WREG status"
ObjExcel.Cells(1, 11).Value = "Relationship to MEMB 01"
ObjExcel.Cells(1, 12).Value = "Recipient Status"
ObjExcel.Cells(1, 13).Value = "MAXIS Counted Ind"
ObjExcel.Cells(1, 14).Value = "HH member number"

objExcel.worksheets.Add().Name = "Cases checked"

ObjExcel.Cells(1, 1).Value = "x102"
ObjExcel.Cells(1, 2).Value = "M# on SNAP"
ObjExcel.Cells(1, 3).Value = "MEMBs open on FS"

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

excel_case_list_row_variable = 2
excel_person_list_row_variable = 2

Do
  case_number = ObjExcel.Cells(excel_case_list_row_variable, 2).Value 
  If case_number = "" then exit do
  back_to_self
  EMWriteScreen "elig", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "fs", 21, 70
  transmit
  Do 'Only reading the APPROVED version
    EMReadScreen approved_ELIG_check, 8, 3, 3 
    If approved_ELIG_check <> "APPROVED" then
      EMReadScreen version_number, 2, 2, 12
      version_number = cint(version_number)
      version_to_check = version_number - 1
      If len(version_to_check) = 1 then version_to_check = "0" & version_to_check
      EMWriteScreen version_to_check, 19, 78
      transmit
    End if
  Loop until approved_ELIG_check = "APPROVED"
  ELIG_MEMB_row = 7 'Setting up variable for the following do...loop. 
  Do
    EMReadScreen x102_number, 7, 20, 16
    EMReadScreen MEMB_ref_nbr, 2, ELIG_MEMB_row, 10
    EMReadScreen counted_ind, 7, ELIG_MEMB_row, 39
    EMReadScreen recipient_status, 1, ELIG_MEMB_row, 35
    EMReadScreen begin_date, 8, ELIG_MEMB_row, 68
    If MEMB_ref_nbr <> "  " then
      objExcel.worksheets("Persons Affected").Activate
      ObjExcel.Cells(excel_person_list_row_variable, 1).Value = x102_number
      ObjExcel.Cells(excel_person_list_row_variable, 5).Value = case_number
      ObjExcel.Cells(excel_person_list_row_variable, 7).Value = "FS"
      ObjExcel.Cells(excel_person_list_row_variable, 8).Value = begin_date
      ObjExcel.Cells(excel_person_list_row_variable, 12).Value = recipient_status
      ObjExcel.Cells(excel_person_list_row_variable, 13).Value = counted_ind
      ObjExcel.Cells(excel_person_list_row_variable, 14).Value = MEMB_ref_nbr
      excel_person_list_row_variable = excel_person_list_row_variable + 1
      objExcel.worksheets("Cases checked").Activate
    End if
    ELIG_MEMB_row = ELIG_MEMB_row + 1
  Loop until MEMB_ref_nbr = "  "
  excel_case_list_row_variable = excel_case_list_row_variable + 1
Loop until case_number = ""

End sub

MsgBox "found"

excel_person_list_row_variable = 20197
objExcel.worksheets("Persons Affected").Activate

Do
  case_number = ObjExcel.Cells(excel_person_list_row_variable, 5).Value 
  If case_number = "" then exit do
  MEMB_number = ObjExcel.Cells(excel_person_list_row_variable, 14).Value 
  If len(MEMB_number) < 2 then MEMB_number = "0" & MEMB_number
  back_to_self
  EMWriteScreen "stat", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "memb", 21, 70
  EMWriteScreen MEMB_number, 21, 75
  transmit
  EMReadScreen MEMB_check, 4, 2, 48 'Error prone cases require an extra transmit
  If MEMB_check <> "MEMB" then transmit
  EMReadScreen first_name, 12, 6, 63
  EMReadScreen last_name, 25, 6, 30
  EMReadScreen PMI, 8, 4, 46
  EMReadScreen date_of_birth, 10, 8, 42
  EMReadScreen relationship_to_applicant, 2, 10, 42
  EMWriteScreen "wreg", 20, 71
  EMWriteScreen MEMB_number, 20, 76
  transmit
  EMReadScreen WREG_ABAWD_status, 2, 13, 50
  EMReadScreen WREG_status, 2, 8, 50
  ObjExcel.Cells(excel_person_list_row_variable, 2).Value = replace(first_name, "_", "")
  ObjExcel.Cells(excel_person_list_row_variable, 3).Value = replace(last_name, "_", "")
  ObjExcel.Cells(excel_person_list_row_variable, 4).Value = PMI
  ObjExcel.Cells(excel_person_list_row_variable, 6).Value = replace(date_of_birth, " ", "/")
  ObjExcel.Cells(excel_person_list_row_variable, 9).Value = WREG_ABAWD_status
  ObjExcel.Cells(excel_person_list_row_variable, 10).Value = WREG_status
  ObjExcel.Cells(excel_person_list_row_variable, 11).Value = relationship_to_applicant
  excel_person_list_row_variable = excel_person_list_row_variable + 1
Loop until case_number = ""


stop_time = timer
MsgBox stop_time - start_time