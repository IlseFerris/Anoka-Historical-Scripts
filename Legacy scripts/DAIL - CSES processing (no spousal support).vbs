EMConnect ""

Dim line_01_PMI_array
Dim line_02_PMI_array
Dim line_03_PMI_array
Dim line_04_PMI_array
Dim line_05_PMI_array
Dim line_06_PMI_array
Dim line_07_PMI_array
Dim line_08_PMI_array
Dim line_09_PMI_array
Dim line_10_PMI_array
Dim line_11_PMI_array
Dim line_12_PMI_array
Dim line_13_PMI_array

'First it checks to see if PRISM is on the same screen. If not, the script will stop and notify the worker.
EMSendKey "<attn>"
EMWaitReady 1, 50
EMReadScreen PRISM_check, 7, 17, 15
If PRISM_check <> "RUNNING" then MsgBox "You need PRISM running on this BlueZone window to continue. The script will now stop."
If PRISM_check <> "RUNNING" then stopscript
EMSendKey "<attn>"
EMWaitReady 1, 50

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = False 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = False 'Set this to false to make alerts go away. This is necessary in production.

EMSendKey "t" + "<enter>"
EMWaitReady 1, 1
EMReadScreen case_number, 8, 5, 73
EMWriteScreen "cses", 20, 70 'This is set as a TIKL for testing purposes. It should be set as "CSES" for live purposes.
EMSendKey "<enter>"
EMWaitReady 1, 1
EMWriteScreen case_number, 20, 38
EMSendKey "<enter>"
EMWaitReady 1, 1

Sub read_messages

excel_row = 1 'setting this variable for the script

  EMReadScreen line_01_check, 4, 6, 20
  If line_01_check = "    " or line_01_check = "----" then exit sub
  If line_01_check <> "DISB" then MsgBox "This is not a DISB CS message. If you have other CSES messages that are not about CS disbursements, please clear them manually before using this script again. If you have questions, contact the scripts administrator."
  If line_01_check <> "DISB" then objExcel.Workbooks.Close
  If line_01_check <> "DISB" then objExcel.quit
  If line_01_check <> "DISB" then stopscript
  EMWriteScreen "x", 6, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_01_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_01_COEX_amt, 6, row, col + 1
  line_01_COEX_amt = Replace(line_01_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_01_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_01_issue_date, 8, row, col - 8
  EMReadScreen line_01_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_01_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_01_raw_PMI_numbers = line_01_raw_PMI_numbers_initial & line_01_raw_PMI_numbers_overflow
  line_01_PMI_numbers_no_spaces = Replace(line_01_raw_PMI_numbers, " ", "")
  line_01_PMI_array = Split(line_01_PMI_numbers_no_spaces, ",")
  For each x in line_01_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "1"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_01_COEX_amt/line_01_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_01_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_01_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_02_check, 4, 7, 20
  If line_02_check <> "DISB" then exit sub
  EMWriteScreen "x", 7, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_02_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_02_COEX_amt, 6, row, col + 1
  line_02_COEX_amt = Replace(line_02_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_02_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_02_issue_date, 8, row, col - 8
  EMReadScreen line_02_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_02_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_02_raw_PMI_numbers = line_02_raw_PMI_numbers_initial & line_02_raw_PMI_numbers_overflow
  line_02_PMI_numbers_no_spaces = Replace(line_02_raw_PMI_numbers, " ", "")
  line_02_PMI_array = Split(line_02_PMI_numbers_no_spaces, ",")
  For each x in line_02_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_02_COEX_amt/line_02_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_02_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_02_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_03_check, 4, 8, 20
  If line_03_check <> "DISB" then exit sub
  EMWriteScreen "x", 8, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_03_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_03_COEX_amt, 6, row, col + 1
  line_03_COEX_amt = Replace(line_03_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_03_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_03_issue_date, 8, row, col - 8
  EMReadScreen line_03_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_03_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_03_raw_PMI_numbers = line_03_raw_PMI_numbers_initial & line_03_raw_PMI_numbers_overflow
  line_03_PMI_numbers_no_spaces = Replace(line_03_raw_PMI_numbers, " ", "")
  line_03_PMI_array = Split(line_03_PMI_numbers_no_spaces, ",")
  For each x in line_03_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_03_COEX_amt/line_03_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_03_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_03_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_04_check, 4, 9, 20
  If line_04_check <> "DISB" then exit sub
  EMWriteScreen "x", 9, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_04_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_04_COEX_amt, 6, row, col + 1
  line_04_COEX_amt = Replace(line_04_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_04_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_04_issue_date, 8, row, col - 8
  EMReadScreen line_04_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_04_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_04_raw_PMI_numbers = line_04_raw_PMI_numbers_initial & line_04_raw_PMI_numbers_overflow
  line_04_PMI_numbers_no_spaces = Replace(line_04_raw_PMI_numbers, " ", "")
  line_04_PMI_array = Split(line_04_PMI_numbers_no_spaces, ",")
  For each x in line_04_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_04_COEX_amt/line_04_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_04_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_04_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_05_check, 4, 10, 20
  If line_05_check <> "DISB" then exit sub
  EMWriteScreen "x", 10, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_05_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_05_COEX_amt, 6, row, col + 1
  line_05_COEX_amt = Replace(line_05_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_05_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_05_issue_date, 8, row, col - 8
  EMReadScreen line_05_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_05_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_05_raw_PMI_numbers = line_05_raw_PMI_numbers_initial & line_05_raw_PMI_numbers_overflow
  line_05_PMI_numbers_no_spaces = Replace(line_05_raw_PMI_numbers, " ", "")
  line_05_PMI_array = Split(line_05_PMI_numbers_no_spaces, ",")
  For each x in line_05_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_05_COEX_amt/line_05_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_05_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_05_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_06_check, 4, 11, 20
  If line_06_check <> "DISB" then exit sub
  EMWriteScreen "x", 11, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_06_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_06_COEX_amt, 6, row, col + 1
  line_06_COEX_amt = Replace(line_06_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_06_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_06_issue_date, 8, row, col - 8
  EMReadScreen line_06_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_06_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_06_raw_PMI_numbers = line_06_raw_PMI_numbers_initial & line_06_raw_PMI_numbers_overflow
  line_06_PMI_numbers_no_spaces = Replace(line_06_raw_PMI_numbers, " ", "")
  line_06_PMI_array = Split(line_06_PMI_numbers_no_spaces, ",")
  For each x in line_06_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_06_COEX_amt/line_06_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_06_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_06_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_07_check, 4, 12, 20
  If line_07_check <> "DISB" then exit sub
  EMWriteScreen "x", 12, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_07_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_07_COEX_amt, 6, row, col + 1
  line_07_COEX_amt = Replace(line_07_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_07_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_07_issue_date, 8, row, col - 8
  EMReadScreen line_07_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_07_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_07_raw_PMI_numbers = line_07_raw_PMI_numbers_initial & line_07_raw_PMI_numbers_overflow
  line_07_PMI_numbers_no_spaces = Replace(line_07_raw_PMI_numbers, " ", "")
  line_07_PMI_array = Split(line_07_PMI_numbers_no_spaces, ",")
  For each x in line_07_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_07_COEX_amt/line_07_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_07_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_07_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_08_check, 4, 13, 20
  If line_08_check <> "DISB" then exit sub
  EMWriteScreen "x", 13, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_08_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_08_COEX_amt, 6, row, col + 1
  line_08_COEX_amt = Replace(line_08_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_08_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_08_issue_date, 8, row, col - 8
  EMReadScreen line_08_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_08_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_08_raw_PMI_numbers = line_08_raw_PMI_numbers_initial & line_08_raw_PMI_numbers_overflow
  line_08_PMI_numbers_no_spaces = Replace(line_08_raw_PMI_numbers, " ", "")
  line_08_PMI_array = Split(line_08_PMI_numbers_no_spaces, ",")
  For each x in line_08_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_08_COEX_amt/line_08_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_08_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_08_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_09_check, 4, 14, 20
  If line_09_check <> "DISB" then exit sub
  EMWriteScreen "x", 14, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_09_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_09_COEX_amt, 6, row, col + 1
  line_09_COEX_amt = Replace(line_09_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_09_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_09_issue_date, 8, row, col - 8
  EMReadScreen line_09_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_09_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_09_raw_PMI_numbers = line_09_raw_PMI_numbers_initial & line_09_raw_PMI_numbers_overflow
  line_09_PMI_numbers_no_spaces = Replace(line_09_raw_PMI_numbers, " ", "")
  line_09_PMI_array = Split(line_09_PMI_numbers_no_spaces, ",")
  For each x in line_09_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_09_COEX_amt/line_09_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_09_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_09_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_10_check, 4, 15, 20
  If line_10_check <> "DISB" then exit sub
  EMWriteScreen "x", 15, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_10_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_10_COEX_amt, 6, row, col + 1
  line_10_COEX_amt = Replace(line_10_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_10_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_10_issue_date, 8, row, col - 8
  EMReadScreen line_10_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_10_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_10_raw_PMI_numbers = line_10_raw_PMI_numbers_initial & line_10_raw_PMI_numbers_overflow
  line_10_PMI_numbers_no_spaces = Replace(line_10_raw_PMI_numbers, " ", "")
  line_10_PMI_array = Split(line_10_PMI_numbers_no_spaces, ",")
  For each x in line_10_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_10_COEX_amt/line_10_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_10_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_10_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_11_check, 4, 16, 20
  If line_11_check <> "DISB" then exit sub
  EMWriteScreen "x", 16, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_11_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_11_COEX_amt, 6, row, col + 1
  line_11_COEX_amt = Replace(line_11_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_11_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_11_issue_date, 8, row, col - 8
  EMReadScreen line_11_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_11_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_11_raw_PMI_numbers = line_11_raw_PMI_numbers_initial & line_11_raw_PMI_numbers_overflow
  line_11_PMI_numbers_no_spaces = Replace(line_11_raw_PMI_numbers, " ", "")
  line_11_PMI_array = Split(line_11_PMI_numbers_no_spaces, ",")
  For each x in line_11_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_11_COEX_amt/line_11_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_11_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_11_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_12_check, 4, 17, 20
  If line_12_check <> "DISB" then exit sub
  EMWriteScreen "x", 17, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_12_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_12_COEX_amt, 6, row, col + 1
  line_12_COEX_amt = Replace(line_12_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_12_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_12_issue_date, 8, row, col - 8
  EMReadScreen line_12_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_12_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_12_raw_PMI_numbers = line_12_raw_PMI_numbers_initial & line_12_raw_PMI_numbers_overflow
  line_12_PMI_numbers_no_spaces = Replace(line_12_raw_PMI_numbers, " ", "")
  line_12_PMI_array = Split(line_12_PMI_numbers_no_spaces, ",")
  For each x in line_12_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_12_COEX_amt/line_12_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_12_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_12_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMReadScreen line_13_check, 4, 18, 20
  If line_13_check <> "DISB" then exit sub
  EMWriteScreen "x", 18, 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "TYPE", row, col
  EMReadScreen line_13_CS_type, 2, row, col + 5
    row = 1
    col = 1
  EMSearch "$", row, col
  EMReadScreen line_13_COEX_amt, 6, row, col + 1
  line_13_COEX_amt = Replace(line_13_COEX_amt, "F", "")
  EMSearch "CHILD(REN)", row, col
  EMReadScreen line_13_COEX_PMI_total, 1, row, col - 2
  EMSearch " TO PMI(S): ", row, col
  EMReadScreen line_13_issue_date, 8, row, col - 8
  EMReadScreen line_13_raw_PMI_numbers_initial, 40, row, col + 12
  EMReadScreen line_13_raw_PMI_numbers_overflow, 70, row + 1, 5
  line_13_raw_PMI_numbers = line_13_raw_PMI_numbers_initial & line_13_raw_PMI_numbers_overflow
  line_13_PMI_numbers_no_spaces = Replace(line_13_raw_PMI_numbers, " ", "")
  line_13_PMI_array = Split(line_13_PMI_numbers_no_spaces, ",")
  For each x in line_13_PMI_array
    ObjExcel.Cells(excel_row, 1).Value = "2"
    ObjExcel.Cells(excel_row, 2).Value = x
    ObjExcel.Cells(excel_row, 4).Value = line_13_COEX_amt/line_13_COEX_PMI_total
    ObjExcel.Cells(excel_row, 5).Value = line_13_CS_type
    ObjExcel.Cells(excel_row, 6).Value = line_13_issue_date
    excel_row = excel_row + 1
  Next
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
End Sub

read_messages

'Now the script goes into case/curr, and checks to see what programs are currently open.
EMWriteScreen "h", 6, 3
EMSendKey "<enter>"
EMWaitReady 1, 1
  row = 1
  col = 1
EMSearch "Case: INACTIVE", row, col 'First the script looks for the case to be inactive. If it is inactive the script will stop.
If row <> 0 then MsgBox "This case is inactive in MAXIS. The script will now stop. If this case is MCRE only process manually at this time."
If row <> 0 then objExcel.Workbooks.Close
If row <> 0 then objExcel.quit
If row <> 0 then stopscript
  row = 1
  col = 1
EMSearch "MFIP:", row, col 'Now it is looking for MFIP to be active.
If row <> 0 then MFIP_active = "True"
If row = 0 then MFIP_active = "False"
  row = 1
  col = 1
EMSearch "HC:", row, col 'Now it is looking for HC to be active.
If row <> 0 then HC_active = "True"
If row = 0 then HC_active = "False"
  row = 1
  col = 1
EMSearch "FS:", row, col 'Now it is looking for FS to be active.
If row <> 0 then FS_active = "True"
If row = 0 then FS_active = "False"

'Now it gets to STAT/MEMB to associate the HH members with the PMIs
EMWriteScreen "stat", 20, 22
EMWriteScreen "memb", 20, 69
EMSendKey "<enter>"
EMWaitReady 1, 1

EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then MsgBox "This case appears to have been abended. Press ''OK'', then transmit, then try this DAIL message again."
If stat_check <> "STAT" then objExcel.Workbooks.Close
If stat_check <> "STAT" then objExcel.quit
If stat_check <> "STAT" then stopscript

'The following checks for error prone cases.
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then
  EMSendKey "<enter>"
  EMWaitReady 1, 1
End if

'Now we're in STAT/MEMB, and the script will associate a PMI with that HH member.
excel_row = 1 'setting the variable for the following Do...Loop
'The following checks for single-member households. They do not currently work, as the second generation do...loop will not catch the PMI, because the "Enter a valid command" notice doesn't go away.
EMReadScreen second_member_check, 2, 6, 3
If second_member_check = "  " then MsgBox "This is a single-individual household. These are not currently covered by the script. Process manually, and watch your email for a script update which will correct this problem."
If second_member_check = "  " then objExcel.Workbooks.Close
If second_member_check = "  " then objExcel.quit
If second_member_check = "  " then stopscript

Do
  Do
    EMReadScreen all_members_checked, 31, 24, 2
    If all_members_checked = "ENTER A VALID COMMAND OR PF-KEY" then exit do
    EMReadScreen PMI_from_MEMB, 8, 4, 46
    PMI_check = Replace(PMI_from_MEMB, " ", "")
    EMReadScreen HH_memb_number, 2, 4, 33
    EMReadScreen SSN_number, 11, 7, 42
    excel_variable = CStr(ObjExcel.Cells(excel_row, 2).Value)
    If excel_variable = PMI_check then ObjExcel.Cells(excel_row, 3).Value = HH_memb_number
    If excel_variable = PMI_check then ObjExcel.Cells(excel_row, 9).Value = SSN_number
    EMSendKey "<enter>"
    EMWaitReady 1, 1
  Loop until all_members_checked = "ENTER A VALID COMMAND OR PF-KEY"
  If ObjExcel.Cells(excel_row, 3).Value = "" then MsgBox "A HH member could not be determined. A PMI could be missing, or this may be arrears for a child who is no longer in the home. Process manually."
  If ObjExcel.Cells(excel_row, 3).Value = "" then need_to_quit = "True"
  If need_to_quit = "True" then objExcel.Workbooks.Close
  If need_to_quit = "True" then objExcel.quit
  If need_to_quit = "True" then stopscript
  need_to_quit = "False" 'Resetting this variable.
  excel_row = excel_row + 1
  EMWriteScreen "01", 20, 76
  EMSendKey "<enter>"
  EMWaitReady 1, 1
Loop until ObjExcel.Cells(excel_row, 2).Value = ""

'Now it reads the footer month for the case, determines what the retro month would be, and gets to UNEA
EMReadScreen footer_month, 2, 20, 55
EMReadScreen footer_year, 2, 20, 58
retro_month = footer_month - 2
retro_year = footer_year
If retro_month = -1 then retro_year = footer_year - 1
If retro_month = -1 then retro_month = 11
If retro_month = 0 then retro_year = footer_year - 1
If retro_month = 0 then retro_month = 12
If len(footer_month) = 1 then footer_month = "0" & footer_month
If len(retro_month) = 1 then retro_month = "0" & retro_month

EMWriteScreen "unea", 20, 71
EMSendKey "<enter>"
EMWaitReady 1, 1

'Now it gets to the UNEA panel for the first member with CS
excel_row = 1 'setting the variable for the following Do...Loop

'Declaring a sub for MFIP cases.
Sub MFIP_sub
  EMSendKey "<PF9>"
  EMWaitReady 1, 1
'Now it updates the code to be a "6" for verification type
  EMWriteScreen "6", 5, 65
'Now it clears out all of the old data.
  EMSetCursor 13, 25
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 28
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 31
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 39
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 54
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 57
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 60
  EMSendKey "<eraseeof>"
  EMSetCursor 13, 68
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 25
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 28
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 31
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 39
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 54
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 57
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 60
  EMSendKey "<eraseeof>"
  EMSetCursor 14, 68
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 25
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 28
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 31
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 39
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 54
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 57
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 60
  EMSendKey "<eraseeof>"
  EMSetCursor 15, 68
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 25
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 28
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 31
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 39
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 54
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 57
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 60
  EMSendKey "<eraseeof>"
  EMSetCursor 16, 68
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 25
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 28
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 31
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 39
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 54
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 57
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 60
  EMSendKey "<eraseeof>"
  EMSetCursor 17, 68
  EMSendKey "<eraseeof>"
  EMWriteScreen retro_month, 13, 25
  issue_day = day(ObjExcel.Cells(excel_row, 6).Value)
  If len(issue_day) = 1 then issue_day = "0" & issue_day
  EMWriteScreen issue_day, 13, 28
  EMWriteScreen retro_year, 13, 31
  payment_amount = FormatNumber(ObjExcel.Cells(excel_row, 4).Value, 2)
  EMWriteScreen payment_amount, 13, 39
  EMWriteScreen footer_month, 13, 54
  EMWriteScreen issue_day, 13, 57
  EMWriteScreen footer_year, 13, 60
  EMWriteScreen payment_amount, 13, 68
'The following determines if there are multiple amounts that need to be added into the case for MFIP.
  MFIP_memb_excel_row = excel_row 'Setting the variable for the next Do...Loop
  MAXIS_payment_row = 14 'Setting the variable for the MAXIS payment row
  HH_memb_to_check = ObjExcel.Cells(excel_row, 3).Value
  ObjExcel.Cells(excel_row, 8).Value = "checked"
  Do
    MFIP_memb_excel_row = MFIP_memb_excel_row + 1 'This was originally under the following If...then. I moved it 05/11/2012.
    If MAXIS_payment_row >= 18 and ObjExcel.Cells(MFIP_memb_excel_row, 3).Value = HH_memb_to_check then 'I added the HH_memb_to_check section 05/11/2012 in response to the script incorrectly showing over five dates, when there was more than one child on the case.
      MsgBox "There are more than five paydates for this case. At this time, process this manually. If this is a common occurrence, contact the script administrator to have this feature added to the script."
      objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
      objExcel.quit
      stopscript
    End if
    next_issue_day = day(ObjExcel.Cells(MFIP_memb_excel_row, 6).Value)
    next_payment_amount = FormatNumber(ObjExcel.Cells(MFIP_memb_excel_row, 4).Value, 2)
    if len(next_issue_day) = 1 then next_issue_day = "0" & next_issue_day
    If ObjExcel.Cells(MFIP_memb_excel_row, 3).Value = HH_memb_to_check and Cint(income_type_on_UNEA) = ObjExcel.Cells(MFIP_memb_excel_row, 5).Value then 
      EMWriteScreen retro_month, MAXIS_payment_row, 25
      EMWriteScreen next_issue_day, MAXIS_payment_row, 28
      EMWriteScreen retro_year, MAXIS_payment_row, 31
      EMWriteScreen "        ", MAXIS_payment_row, 39
      EMWriteScreen next_payment_amount, MAXIS_payment_row, 39
      ObjExcel.Cells(MFIP_memb_excel_row, 8).Value = "checked"
      EMWriteScreen footer_month, MAXIS_payment_row, 54
      EMWriteScreen next_issue_day, MAXIS_payment_row, 57
      EMWriteScreen retro_year, MAXIS_payment_row, 60
      EMWriteScreen "        ", MAXIS_payment_row, 68
      EMWriteScreen next_payment_amount, MAXIS_payment_row, 68
      MAXIS_payment_row = MAXIS_payment_row + 1
    End If
  Loop until ObjExcel.Cells(MFIP_memb_excel_row, 3).Value = ""
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMSendKey "<enter>"
  EMWaitReady 1, 1
End Sub

'Declaring a sub for FS cases.
Sub FS_sub
'  EMSendKey "<PF9>" 'These two items are needed IF the script is going to edit the UNEA panels. For now they are turned off.
'  EMWaitReady 1, 1

'First it adds the FS amounts together for the month.
  CSES_amt_excel_row = excel_row 'Setting variable for determining the total amount from CSES message
  HH_memb_to_check = ObjExcel.Cells(excel_row, 3).Value
  CSES_amt = ObjExcel.Cells(excel_row, 4).Value
  Do
    CSES_amt_excel_row = CSES_amt_excel_row + 1
    If ObjExcel.Cells(CSES_amt_excel_row, 3).Value = HH_memb_to_check and Cint(income_type_on_UNEA) = ObjExcel.Cells(CSES_amt_excel_row, 5).Value then CSES_amt = CSES_amt + ObjExcel.Cells(CSES_amt_excel_row, 4).Value
    If ObjExcel.Cells(CSES_amt_excel_row, 3).Value = HH_memb_to_check and Cint(income_type_on_UNEA) = ObjExcel.Cells(CSES_amt_excel_row, 5).Value then ObjExcel.Cells(CSES_amt_excel_row, 7).Value = "checked"
  Loop until ObjExcel.Cells(CSES_amt_excel_row, 3).Value = ""

'Now it enters the PIC to determine if the FS amount is appropriate.
  EMWriteScreen "x", 10, 26
  EMSendKey "<enter>"
  EMWaitReady 1, 1

'What follows figures out the lowest_amt and highest_amt of FS on the PIC.
  Dim income_received_01
  Dim income_received_02
  Dim income_received_03
  Dim income_received_04
  Dim income_received_05
  EMReadScreen income_received_01, 8, 9, 25
  EMReadScreen income_received_02, 8, 10, 25
  EMReadScreen income_received_03, 8, 11, 25
  EMReadScreen income_received_04, 8, 12, 25
  EMReadScreen income_received_05, 8, 13, 25
  If income_received_01 = "________" then MsgBox "This case has CS, but does not have a PIC for the client who receives the CS, or the income is listed as anticipated income. You will have to manually update the PIC at this time. After a new range is determined, you can try the script again!"
  If income_received_01 = "________" then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
  If income_received_01 = "________" then objExcel.quit
  If income_received_01 = "________" then stopscript
  If income_received_02 = "________" then income_received_02 = income_received_01
  If income_received_03 = "________" then income_received_03 = income_received_02
  If income_received_04 = "________" then income_received_04 = income_received_03
  If income_received_05 = "________" then income_received_05 = income_received_04
  If abs(income_received_01) <= abs(income_received_02) and abs(income_received_01) <= abs(income_received_03) and abs(income_received_01) <= abs(income_received_04) and abs(income_received_01) <= abs(income_received_05) then lowest_amt = abs(income_received_01)
  If abs(income_received_02) <= abs(income_received_01) and abs(income_received_02) <= abs(income_received_03) and abs(income_received_02) <= abs(income_received_04) and abs(income_received_02) <= abs(income_received_05) then lowest_amt = abs(income_received_02)
  If abs(income_received_03) <= abs(income_received_02) and abs(income_received_03) <= abs(income_received_01) and abs(income_received_03) <= abs(income_received_04) and abs(income_received_03) <= abs(income_received_05) then lowest_amt = abs(income_received_03)
  If abs(income_received_04) <= abs(income_received_02) and abs(income_received_04) <= abs(income_received_03) and abs(income_received_04) <= abs(income_received_01) and abs(income_received_04) <= abs(income_received_05) then lowest_amt = abs(income_received_04)
  If abs(income_received_05) <= abs(income_received_02) and abs(income_received_05) <= abs(income_received_03) and abs(income_received_05) <= abs(income_received_04) and abs(income_received_05) <= abs(income_received_01) then lowest_amt = abs(income_received_05)

  If abs(income_received_01) >= abs(income_received_02) and abs(income_received_01) >= abs(income_received_03) and abs(income_received_01) >= abs(income_received_04) and abs(income_received_01) >= abs(income_received_05) then highest_amt = abs(income_received_01)
  If abs(income_received_02) >= abs(income_received_01) and abs(income_received_02) >= abs(income_received_03) and abs(income_received_02) >= abs(income_received_04) and abs(income_received_02) >= abs(income_received_05) then highest_amt = abs(income_received_02)
  If abs(income_received_03) >= abs(income_received_02) and abs(income_received_03) >= abs(income_received_01) and abs(income_received_03) >= abs(income_received_04) and abs(income_received_03) >= abs(income_received_05) then highest_amt = abs(income_received_03)
  If abs(income_received_04) >= abs(income_received_02) and abs(income_received_04) >= abs(income_received_03) and abs(income_received_04) >= abs(income_received_01) and abs(income_received_04) >= abs(income_received_05) then highest_amt = abs(income_received_04)
  If abs(income_received_05) >= abs(income_received_02) and abs(income_received_05) >= abs(income_received_03) and abs(income_received_05) >= abs(income_received_04) and abs(income_received_05) >= abs(income_received_01) then highest_amt = abs(income_received_05)
  If IsEmpty(highest_amt) = True then highest_amt = abs(income_received_01)
  If lowest_amt = 0 then lowest_amt = income_received_01
  If income_received_01 = "    0.00" or income_received_02 = "    0.00" or income_received_03 = "    0.00" or income_received_04 = "    0.00" or income_received_05 = "    0.00" then lowest_amt = 0
  If CSES_amt >= lowest_amt - (lowest_amt/10) and CSES_amt <= highest_amt + (highest_amt/10) then within_range = "True"
  If CSES_amt < lowest_amt - (lowest_amt/10) or CSES_amt > highest_amt + (highest_amt/10) then within_range = "False"
  If within_range = "False" then MsgBox "The CS received appears to be out of the range for FS. At this time, process this manually."
  If within_range = "False" then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
  If within_range = "False" then objExcel.quit
  If within_range = "False" then stopscript
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMSendKey "<PF10>"
  EMWaitReady 1, 1
End Sub

Dim HC_status

Sub HC_sub
  EMWriteScreen "revw", 20, 71
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen revw_month, 2, 9, 70
  If revw_month <> footer_month then HC_status = "* No HC review due at this time. No changes made for HC."
  If revw_month = footer_month then HC_status = "* A review is due for HC. Awaiting review before processing."
End Sub

'The following is the editing section. If working in inquiry, turn it into a sub by un-commenting the sub sections.

'Sub fake_sub

Do
  EMReadScreen income_end_date_error_check, 50, 24, 2
  If income_end_date_error_check = "RETROSPECTIVE DATE CANNOT BE AFTER INCOME END DATE" then
    MsgBox "You have an income end date on this panel, but the income does not appear to have ended, or it has started up again. Fix this panel, then try the script again."
    objExcel.Workbooks.Close
    objExcel.quit
    stopscript
  End if
  UNEA_number = ObjExcel.Cells(excel_row, 3).Value
  If Len(UNEA_number) = 1 then UNEA_number = "0" & UNEA_number
'  EMWriteScreen "unea", 20, 71
  EMWriteScreen UNEA_number, 20, 76
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen panel_amt_check, 1, 2, 78
  If panel_amt_check <> "1" then 
'    EMWriteScreen "unea", 20, 71
    EMWriteScreen "01", 20, 79
    EMSendKey "<enter>"
    EMWaitReady 1, 1
  End if
  Do
    EMReadScreen income_type_on_UNEA, 2, 5, 37
    If income_type_on_UNEA = "__" then MsgBox "The script cannot find an appropriate CS panel for this case. You may need to add a new panel. Process manually at this time."
    If income_type_on_UNEA = "__" then objExcel.Workbooks.Close
    If income_type_on_UNEA = "__" then objExcel.quit
    If income_type_on_UNEA = "__" then stopscript
    If Cint(income_type_on_UNEA) <> ObjExcel.Cells(excel_row, 5).Value then EMSendKey "<enter>"
    If Cint(income_type_on_UNEA) <> ObjExcel.Cells(excel_row, 5).Value then EMWaitReady 1, 1
    EMReadScreen all_panels_checked, 5, 24, 02
    If all_panels_checked = "ENTER" then MsgBox "The script cannot find an appropriate CS panel for this case. You may need to add a new panel. Process manually at this time."
    If all_panels_checked = "ENTER" then objExcel.Workbooks.Close
    If all_panels_checked = "ENTER" then objExcel.quit
    If all_panels_checked = "ENTER" then stopscript
  Loop until Cint(income_type_on_UNEA) = ObjExcel.Cells(excel_row, 5).Value
  If MFIP_active = "True" and ObjExcel.Cells(excel_row, 8).Value <> "checked" then call MFIP_sub
  If MFIP_active <> "True" and FS_active = "True" and ObjExcel.Cells(excel_row, 7).Value <> "checked" then call FS_sub
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 3).Value = ""

If HC_active = "True" then call HC_sub

'This is a dialog which will ask if the worker wants to case note, if the case was already case noted.
BeginDialog already_case_noted_dialog, 0, 0, 191, 52, "Already case noted?"
  ButtonGroup already_case_noted_dialog_ButtonPressed
    CancelButton 130, 30, 50, 15
    OkButton 130, 10, 50, 15
  Text 10, 10, 105, 35, "You appear to have already case noted this. To case note again, press ''ok''. To exit, press ''cancel''."
EndDialog
already_case_noted_dialog_ButtonPressed = "1" 'setting the variable for the next section.
EMSendKey "<PF4>"
EMWaitReady 1, 1
EMReadScreen CSES_messages_reviewed_check, 28, 5, 25
If CSES_messages_reviewed_check = ":::CSES messages reviewed:::" then dialog already_case_noted_dialog
If already_case_noted_dialog_ButtonPressed = 0 then stopscript
If already_case_noted_dialog_ButtonPressed = 0 then objExcel.Workbooks.Close
If already_case_noted_dialog_ButtonPressed = 0 then objExcel.quit
If already_case_noted_dialog_ButtonPressed = 0 then stopscript
EMSendKey "<PF9>"
EMWaitReady 1, 1
EMReadScreen case_note_mode_check, 7, 20, 3
If case_note_mode_check <> "Mode: A" then MsgBox "You are not in a case note on edit mode. You might be in inquiry. Try the script again in production."
If case_note_mode_check <> "Mode: A" then objExcel.Workbooks.Close
If case_note_mode_check <> "Mode: A" then objExcel.quit
If case_note_mode_check <> "Mode: A" then stopscript
EMSendKey ":::CSES messages reviewed:::" + "<newline>"
If MFIP_active = "True" then EMSendKey "* Updated retro/prospective income amounts." + "<newline>"
If MFIP_active <> "True" and FS_active = "True" then EMSendKey "* FS PIC reviewed, income appears to be in range." + "<newline>"
If MFIP_active = "True" and FS_active = "True" then EMSendKey "* FS PIC ignored, as case also has MFIP." + "<newline>"
If HC_active = "True" then EMSendKey HC_status + "<newline>"
EMSendKey "---" + "<newline>"
BeginDialog worker_sig_dialog, 0, 0, 141, 47, "Worker signature"
  EditBox 15, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 85, 5, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 75, 10, "Sign your case note."
EndDialog
dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then objExcel.Workbooks.Close
If ButtonPressed_worker_sig_dialog = 0 then objExcel.quit
If ButtonPressed_worker_sig_dialog = 0 then stopscript
EMSendKey worker_sig & ", using automated script."

'End sub

If MFIP_active = "True" then 
  MsgBox "MFIP is active, so the script will not check PRISM for this case. It will now stop."
  objExcel.Workbooks.Close
  objExcel.quit
  stopscript
End if

'This jumps to PRISM.
EMSendKey "<attn>"
EMWaitReady 1, 50
EMWriteScreen "12", 2, 15
EMSendKey "<enter>"
EMWaitReady 1, 50

excel_row = 1 'Resetting the variable for the PRISM part of the script.
Do 
If ObjExcel.Cells(excel_row, 9).Value = "" then exit do 'This gets out of the do...loop if there is no SSN indicated.

'The following is a lockout dialog to prevent workers from freezing the PRISM screen.
BeginDialog PRISM_lockout_dialog, 0, 0, 191, 57, "PRISM lockout dialog"
  ButtonGroup PRISM_lockout_dialog_ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 10, 5, 110, 45, "You are locked out of PRISM. Get back to the PRISM main menu before pressing OK. Pressing cancel will cause the script to end."
EndDialog


'Now it returns to the PRISM start screen.
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 0
    EMReadScreen PRISM_check, 5, 1, 36
    EMReadScreen PRISM_person_search_check, 9, 2, 34
    If PRISM_check = "PRISM" and PRISM_person_search_check = "Main Menu" then exit do
    If PRISM_check <> "PRISM" then Dialog PRISM_lockout_dialog
    If PRISM_check <> "PRISM" and PRISM_lockout_dialog_ButtonPressed = 0 then 
      objExcel.Workbooks.Close
      objExcel.quit
      stopscript
    End if
      
  Loop until PRISM_check = "PRISM" and PRISM_person_search_check = "Main Menu"

  Do 'This will check to make sure the excel row isn't duplicating work.
    If ObjExcel.Cells(excel_row, 10).Value = "SSN checked" then excel_row = excel_row + 1
  Loop until ObjExcel.Cells(excel_row, 10).Value = ""

  EMWriteScreen "PESE", 21, 18
  EMSendKey "<enter>"
  EMWaitReady 1, 0

  current_SSN_with_spaces = ObjExcel.Cells(excel_row, 9).Value
  current_SSN = replace(ObjExcel.Cells(excel_row, 9).Value, " ", "")
  EMWriteScreen "            ", 5, 20
  EMWriteScreen "            ", 6, 20
  EMWriteScreen "   ", 7, 20
  EMWriteScreen " ", 9, 13
  EMWriteScreen "          ", 9, 32
  EMWriteScreen "  ", 9, 68
  EMWriteScreen "  ", 9, 76
  EMWriteScreen "          ", 10, 32
  EMWriteScreen "N", 10, 67
  EMWriteScreen "N", 10, 76
  EMWriteScreen "N", 12, 54

  EMSetCursor 10, 13
  EMSendKey current_SSN + "<enter>"
  EMWaitReady 1, 0

  EMWriteScreen "x", 5, 5
  EMSendKey "<enter>"
  EMWaitReady 1, 0

'Now it checks to see if there is more than one case. If there is, the script will have a worker message then stop. If not, the script will select the case.
  EMReadScreen case_amount_check, 1, 7, 17
if case_amount_check <> 1 then
  Do 
    EMReadScreen ind_active_check, 1, 7, 41
    If ind_active_check = "Y" then exit do
    EMReadScreen current_case_check, 1, 7, 12
    If current_case_check = case_amount_check then MsgBox "The script could not determine which child support case is active for this HH member. Check PRISM manually."
    If current_case_check = case_amount_check then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
    If current_case_check = case_amount_check then objExcel.quit
    If current_case_check = case_amount_check then stopscript
    EMSendKey "<PF8>"
    EMWaitReady 1, 0
  Loop until ind_active_check = "Y"
end if

  EMWriteScreen "s", 2, 20
  EMSendKey "<enter>"
  EMWaitReady 1, 0

  EMWriteScreen "CAFS", 21, 17
  EMSendKey "<enter>"
  EMWaitReady 1, 0

'Now we are in CAFS, and the script will read the Obl field to determine if the Obl is CCC, CMS, or CMI.
  EMReadScreen CAFS_check_01, 3, 17, 18
  EMReadScreen CAFS_check_02, 3, 18, 18
  EMReadScreen CAFS_check_03, 3, 19, 18
  EMReadScreen CAFS_check_04, 3, 20, 18
  EMReadScreen CAFS_balance_check_01, 4, 17, 59
  EMReadScreen CAFS_balance_check_02, 4, 18, 59
  EMReadScreen CAFS_balance_check_03, 4, 19, 59
  EMReadScreen CAFS_balance_check_04, 4, 20, 59
  If CAFS_balance_check_01 <> "0.00" and (CAFS_check_01 = "CCC" or CAFS_check_01 = "CMS" or CAFS_check_01 = "CMI") then MsgBox "The Obl type is CCC, CMS, or CMI, and a balance is listed. Process this manually, and check the other children in the household for this as well. Check with a PC if you have any questions. The MAXIS part of the script has already case noted for you."
  If CAFS_balance_check_01 <> "0.00" and (CAFS_check_01 = "CCC" or CAFS_check_01 = "CMS" or CAFS_check_01 = "CMI") then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
  If CAFS_balance_check_01 <> "0.00" and (CAFS_check_01 = "CCC" or CAFS_check_01 = "CMS" or CAFS_check_01 = "CMI") then objExcel.quit
  If CAFS_balance_check_01 <> "0.00" and (CAFS_check_01 = "CCC" or CAFS_check_01 = "CMS" or CAFS_check_01 = "CMI") then stopscript
  If CAFS_balance_check_02 <> "0.00" and (CAFS_check_02 = "CCC" or CAFS_check_02 = "CMS" or CAFS_check_02 = "CMI") then MsgBox "The Obl type is CCC, CMS, or CMI, and a balance is listed. Process this manually, and check the other children in the household for this as well. Check with a PC if you have any questions. The MAXIS part of the script has already case noted for you."
  If CAFS_balance_check_02 <> "0.00" and (CAFS_check_02 = "CCC" or CAFS_check_02 = "CMS" or CAFS_check_02 = "CMI") then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
  If CAFS_balance_check_02 <> "0.00" and (CAFS_check_02 = "CCC" or CAFS_check_02 = "CMS" or CAFS_check_02 = "CMI") then objExcel.quit
  If CAFS_balance_check_02 <> "0.00" and (CAFS_check_02 = "CCC" or CAFS_check_02 = "CMS" or CAFS_check_02 = "CMI") then stopscript
  If CAFS_balance_check_03 <> "0.00" and (CAFS_check_03 = "CCC" or CAFS_check_03 = "CMS" or CAFS_check_03 = "CMI") then MsgBox "The Obl type is CCC, CMS, or CMI, and a balance is listed. Process this manually, and check the other children in the household for this as well. Check with a PC if you have any questions. The MAXIS part of the script has already case noted for you."
  If CAFS_balance_check_03 <> "0.00" and (CAFS_check_03 = "CCC" or CAFS_check_03 = "CMS" or CAFS_check_03 = "CMI") then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
  If CAFS_balance_check_03 <> "0.00" and (CAFS_check_03 = "CCC" or CAFS_check_03 = "CMS" or CAFS_check_03 = "CMI") then objExcel.quit
  If CAFS_balance_check_03 <> "0.00" and (CAFS_check_03 = "CCC" or CAFS_check_03 = "CMS" or CAFS_check_03 = "CMI") then stopscript
  If CAFS_balance_check_04 <> "0.00" and (CAFS_check_04 = "CCC" or CAFS_check_04 = "CMS" or CAFS_check_04 = "CMI") then MsgBox "The Obl type is CCC, CMS, or CMI, and a balance is listed. Process this manually, and check the other children in the household for this as well. Check with a PC if you have any questions. The MAXIS part of the script has already case noted for you."
  If CAFS_balance_check_04 <> "0.00" and (CAFS_check_04 = "CCC" or CAFS_check_04 = "CMS" or CAFS_check_04 = "CMI") then objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
  If CAFS_balance_check_04 <> "0.00" and (CAFS_check_04 = "CCC" or CAFS_check_04 = "CMS" or CAFS_check_04 = "CMI") then objExcel.quit
  If CAFS_balance_check_04 <> "0.00" and (CAFS_check_04 = "CCC" or CAFS_check_04 = "CMS" or CAFS_check_04 = "CMI") then stopscript

'Now it returns to the main menu of PRISM.
  EMSendKey "<PF3>"
  EMWaitReady 1, 0

'Now it marks any SSNs that have already been checked as having been checked. This way it doesn't check them again.
  SSN_check_excel_row = excel_row 'copying the row over so we don't overwrite the overall excel row.
  Do
    If current_SSN_with_spaces = ObjExcel.Cells(SSN_check_excel_row, 9).Value and ObjExcel.Cells(SSN_check_excel_row, 9).Value <> "" then ObjExcel.Cells(SSN_check_excel_row, 10).Value = "SSN checked"
    SSN_check_excel_row = SSN_check_excel_row + 1
  Loop until ObjExcel.Cells(SSN_check_excel_row, 9).Value = ""
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 9).Value = ""

'Now it will navigate back to MAXIS for the ending.
EMSendKey "<attn>"
EMWaitReady 1, 50
EMSendKey "<attn>"
EMWaitReady 1, 50

MsgBox "PRISM checked, no CMI/CMS/CCC obl types indicated on CAFS. The script findings are listed in this case note."

objWorkbook = objExcel.Workbooks.Close '---This is how you close a workbook. Two steps!
objExcel.quit
