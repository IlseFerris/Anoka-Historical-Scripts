function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(x) and STAT_note_check <> "NOTE" then 
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen y, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen x, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen y, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    End if
  End if
End function


Function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
End function

Function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.
excel_row = 2 'For adding rows to the spreadsheet
ObjExcel.Cells(1, 1).Value = "MAXIS #"
ObjExcel.Cells(1, 2).Value = "CLIENT NAME"
ObjExcel.Cells(1, 3).Value = "DOB"
ObjExcel.Cells(1, 4).Value = "SSN"
ObjExcel.Cells(1, 5).Value = "UNEA CLAIM #"
ObjExcel.Cells(1, 6).Value = "CURRENT $"
ObjExcel.Cells(1, 6).Value = "WORKER X102"

x102_number_array = array("4DK", "4AS", "4BL", "4SZ", "BED", "752", "4SW", "247", "692", "C02", "989", "B98", "GMZ", "104", "524", "B83", "932", "395", "631", "B93", "987", "769", "4SX", "293", "628", "233", "756", "30V", "4MG", "b64", "598", "268", "880", "4RS", "4SY", "SEC", "SAC", "894", "4SS", "750", "742", "757", "TRP")
'x102_number_array = array("b83") 'Using smaller array for testing

EMConnect ""

call navigate_to_screen("dail", "dail")

For each x102_number in x102_number_array
  EMWriteScreen x102_number, 21, 10
  transmit
  EMWriteScreen "dail", 20, 70
  transmit
  EMWriteScreen "cola", 20, 70
  transmit
  row = 6
  EMReadScreen message_type_warning, 52, 24, 2
  If message_type_warning <> "NO MESSAGES TYPES: COLA  - USE VIEW/PICK TO RESELECT" then 'I have to contain this in an "IF" scenario because of some workers having no COLA messages.
    Do
      Do
       EMReadScreen railroad_retirement_check, 45, row, 20
        If railroad_retirement_check = "CHECK FOR COLA - UNEA HAS RAILROAD RETIREMENT" then
          case_number_row = row - 1
          Do
            EMReadScreen case_number_check, 9, case_number_row, 63
            If case_number_check = "CASE NBR:" then 
              EMReadScreen case_number, 8, case_number_row, 73
'              ObjExcel.Cells(excel_row, 1).Value = case_number
              excel_row = excel_row + 1
              case_number_array = trim(case_number_array & " " & case_number)
            Else
              case_number_row = case_number_row - 1
            End if
          Loop until case_number_check = "CASE NBR:"
        End if
        row = row + 1
      Loop until row = 19
      PF8
      row = 6
      EMReadScreen page_check, 4, 24, 19
      if page_check = "PAGE" then exit do
    Loop until page_check = "PAGE"
  End If
Next

excel_row = 2 'resetting variable
case_number_array = split(case_number_array)

For each case_number in case_number_array
  call navigate_to_screen("stat", "memb") 'Have to go to MEMB first because of problems with UNEA in inquiry mode.
  call navigate_to_screen("stat", "unea") 
  row = 5
  Do
    EMReadScreen new_MEMB_number, 2, row, 3
    If new_MEMB_number <> "  " then
      MEMB_array = trim(MEMB_array & " " & new_MEMB_number)
      row = row + 1
    End if
  Loop until new_MEMB_number = "  " or row = 20
  MEMB_array = split(MEMB_array)
  For each HH_memb in MEMB_array
    EMWriteScreen "unea", 20, 71
    EMWriteScreen HH_memb, 20, 76
    transmit
    Do
      EMReadScreen UNEA_current_panel, 1, 2, 73
      EMReadScreen UNEA_total_check, 1, 2, 78  
      EMReadScreen UNEA_type_check, 2, 5, 37 'Figuring out if UNEA is type 16
      If UNEA_type_check <> "16" then transmit
    Loop until UNEA_current_panel = UNEA_total_check or UNEA_type_check = "16"
    EMReadScreen UNEA_name, 29, 4, 36
    EMReadScreen RR_claim_number, 15, 6, 37
    EMReadScreen prospective_amt, 8, 18, 68
    EMReadScreen worker_number, 7, 21, 21
    EMWriteScreen "memb", 20, 71
    EMWriteScreen HH_memb, 20, 76
    transmit
    EMReadScreen DOB, 10, 8, 42
    EMReadScreen SSN, 11, 7, 42
    ObjExcel.Cells(excel_row, 1).Value = case_number
    ObjExcel.Cells(excel_row, 2).Value = trim(UNEA_name)
    ObjExcel.Cells(excel_row, 3).Value = replace(DOB, " ", "/")
    ObjExcel.Cells(excel_row, 4).Value = replace(SSN, " ", "-")
    ObjExcel.Cells(excel_row, 5).Value = trim(replace(RR_claim_number, "_", ""))
    ObjExcel.Cells(excel_row, 6).Value = trim(prospective_amt)
    ObjExcel.Cells(excel_row, 7).Value = worker_number
  Next
  excel_row = excel_row + 1
  MEMB_array = ""
Next