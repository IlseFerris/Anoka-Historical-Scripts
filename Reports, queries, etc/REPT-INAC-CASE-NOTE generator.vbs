EMConnect ""



EMReadScreen inac_check, 4, 2, 60
If inac_check <> "INAC" then Msgbox "You are not on REPT/INAC. Please get to REPT/INAC before continuing."
If inac_check <> "INAC" then stopscript


BeginDialog worker_sig_dialog, 0, 0, 191, 57, "Worker signature"
  EditBox 35, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 25, 10, 75, 10, "Sign your case note."
EndDialog

Do
  Dialog worker_sig_dialog
  If ButtonPressed_worker_sig_dialog = 0 then stopscript
  If worker_sig = "" then MsgBox "You must sign your case note. The script will not work until you sign your case note."
Loop until worker_sig <> ""


case_row = 2

EMReadScreen worker_number, 7, 21, 16
EMReadScreen footer_month, 2, 20, 54
EMReadScreen footer_year, 2, 20, 57

Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\closings - " & footer_month & "-" & footer_year & " - " & worker_number & ".xlsx"  
Set objWorkbook = objExcel.Workbooks.Add() 
ObjExcel.Cells(1, 1).Value = "M#"
ObjExcel.Cells(1, 1).HorizontalAlignment = -4108 
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 1).ColumnWidth = 08
ObjExcel.Cells(1, 2).Value = "Name"
ObjExcel.Cells(1, 2).HorizontalAlignment = -4108 
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 2).ColumnWidth = 25
ObjExcel.Cells(1, 3).Value = "INAC date"
ObjExcel.Cells(1, 3).HorizontalAlignment = -4108 
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 3).ColumnWidth = 12
ObjExcel.Cells(1, 4).Value = "Already case noted?"
ObjExcel.Cells(1, 4).HorizontalAlignment = -4108 
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 19
ObjExcel.Cells(1, 5).HorizontalAlignment = -4108 
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 5).ColumnWidth = 4
ObjExcel.Cells(1, 6).Value = "Cash status"
ObjExcel.Cells(1, 6).HorizontalAlignment = -4108 
objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(1, 6).ColumnWidth = 25
ObjExcel.Cells(1, 7).Value = "FS status"
ObjExcel.Cells(1, 7).HorizontalAlignment = -4108 
objExcel.Cells(1, 7).Font.Bold = TRUE
objExcel.Cells(1, 7).ColumnWidth = 25
ObjExcel.Cells(1, 8).Value = "HC status"
ObjExcel.Cells(1, 8).HorizontalAlignment = -4108 
objExcel.Cells(1, 8).Font.Bold = TRUE
objExcel.Cells(1, 8).ColumnWidth = 25



Do until last_page_check = "THIS IS THE LAST PAGE"
  EMReadScreen last_page_check, 21, 24, 02
    if last_page_check = "THIS IS THE LAST PAGE" then exit do

'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"



  EMReadScreen first_row_case_number, 8, 7, 3
  EMReadScreen first_row_name, 24, 7, 14
  EMReadScreen first_row_inactive_date, 8, 7, 49
  ObjExcel.Cells(case_row, 1).Value = first_row_case_number
  ObjExcel.Cells(case_row, 2).Value = first_row_name
  ObjExcel.Cells(case_row, 3).Value = first_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen second_row_case_number, 8, 8, 3
  EMReadScreen second_row_name, 24, 8, 14
  EMReadScreen second_row_inactive_date, 8, 8, 49
  ObjExcel.Cells(case_row, 1).Value = second_row_case_number
  ObjExcel.Cells(case_row, 2).Value = second_row_name
  ObjExcel.Cells(case_row, 3).Value = second_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen third_row_case_number, 8, 9, 3
  EMReadScreen third_row_name, 24, 9, 14
  EMReadScreen third_row_inactive_date, 8, 9, 49
  ObjExcel.Cells(case_row, 1).Value = third_row_case_number
  ObjExcel.Cells(case_row, 2).Value = third_row_name
  ObjExcel.Cells(case_row, 3).Value = third_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen fourth_row_case_number, 8, 10, 3
  EMReadScreen fourth_row_name, 24, 10, 14
  EMReadScreen fourth_row_inactive_date, 8, 10, 49
  ObjExcel.Cells(case_row, 1).Value = fourth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = fourth_row_name
  ObjExcel.Cells(case_row, 3).Value = fourth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen fifth_row_case_number, 8, 11, 3
  EMReadScreen fifth_row_name, 24, 11, 14
  EMReadScreen fifth_row_inactive_date, 8, 11, 49
  ObjExcel.Cells(case_row, 1).Value = fifth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = fifth_row_name
  ObjExcel.Cells(case_row, 3).Value = fifth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen sixth_row_case_number, 8, 12, 3
  EMReadScreen sixth_row_name, 24, 12, 14
  EMReadScreen sixth_row_inactive_date, 8, 12, 49
  ObjExcel.Cells(case_row, 1).Value = sixth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = sixth_row_name
  ObjExcel.Cells(case_row, 3).Value = sixth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen seventh_row_case_number, 8, 13, 3
  EMReadScreen seventh_row_name, 24, 13, 14
  EMReadScreen seventh_row_inactive_date, 8, 13, 49
  ObjExcel.Cells(case_row, 1).Value = seventh_row_case_number
  ObjExcel.Cells(case_row, 2).Value = seventh_row_name
  ObjExcel.Cells(case_row, 3).Value = seventh_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen eighth_row_case_number, 8, 14, 3
  EMReadScreen eighth_row_name, 24, 14, 14
  EMReadScreen eighth_row_inactive_date, 8, 14, 49
  ObjExcel.Cells(case_row, 1).Value = eighth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = eighth_row_name
  ObjExcel.Cells(case_row, 3).Value = eighth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen ninth_row_case_number, 8, 15, 3
  EMReadScreen ninth_row_name, 24, 15, 14
  EMReadScreen ninth_row_inactive_date, 8, 15, 49
  ObjExcel.Cells(case_row, 1).Value = ninth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = ninth_row_name
  ObjExcel.Cells(case_row, 3).Value = ninth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen tenth_row_case_number, 8, 16, 3
  EMReadScreen tenth_row_name, 24, 16, 14
  EMReadScreen tenth_row_inactive_date, 8, 16, 49
  ObjExcel.Cells(case_row, 1).Value = tenth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = tenth_row_name
  ObjExcel.Cells(case_row, 3).Value = tenth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen eleventh_row_case_number, 8, 17, 3
  EMReadScreen eleventh_row_name, 24, 17, 14
  EMReadScreen eleventh_row_inactive_date, 8, 17, 49
  ObjExcel.Cells(case_row, 1).Value = eleventh_row_case_number
  ObjExcel.Cells(case_row, 2).Value = eleventh_row_name
  ObjExcel.Cells(case_row, 3).Value = eleventh_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1


  EMReadScreen twelfth_row_case_number, 8, 18, 3
  EMReadScreen twelfth_row_name, 24, 18, 14
  EMReadScreen twelfth_row_inactive_date, 8, 18, 49
  ObjExcel.Cells(case_row, 1).Value = twelfth_row_case_number
  ObjExcel.Cells(case_row, 2).Value = twelfth_row_name
  ObjExcel.Cells(case_row, 3).Value = twelfth_row_inactive_date
  ObjExcel.Cells(case_row, 3).HorizontalAlignment = -4108 

  case_row = case_row + 1

EMSendKey "<PF8>"
EMWaitReady 1, 1

Loop 

'--------Now the script is checking case note for each case, in order to see if a note was already made.



case_row = 2 'Resetting the case row to investigate.


do until ObjExcel.Cells(case_row, 1).Value = "" or ObjExcel.Cells(case_row, 1).Value = "        "

  case_number = ObjExcel.Cells(case_row, 1).Value 
    If case_number = "" or case_number = "        " then exit do

'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMReadScreen SELF_check, 27, 2, 28
    EMWaitReady 1, 1
  loop until SELF_check = "Select Function Menu (SELF)"

'Now we go into CASE/NOTE for each case.
  EMSendKey "<home>" + "case"
  EMSetCursor 18, 43
  EMSendKey "        "
  EMWriteScreen case_number, 18, 43
  EMSetCursor 21, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 1

  row = 1
  col = 1
  EMSearch "---Closed", row, col
  if row = 0 or row > 10 then ObjExcel.Cells(case_row, 4).Value = "No"
  if row > 0 and row <= 10 then ObjExcel.Cells(case_row, 4).Value = "Yes"
  ObjExcel.Cells(case_row, 4).HorizontalAlignment = -4108 

  case_row = case_row + 1 'setting up the script to check the next row.

loop

'-------Now it checks STAT/REVW for the cases that weren't already case noted.

case_row = 2 'Resetting row for Excel spreadsheet.

do until ObjExcel.Cells(case_row, 1).Value = "" or ObjExcel.Cells(case_row, 1).Value = "        "

  Do
    If ObjExcel.Cells(case_row, 4).Value = "Yes" then case_row = case_row + 1
  Loop until ObjExcel.Cells(case_row, 4).Value = "No" or ObjExcel.Cells(case_row, 4).Value = ""


  case_number = ObjExcel.Cells(case_row, 1).Value 
    If case_number = "" or case_number = "        " then exit do



'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMReadScreen SELF_check, 27, 2, 28
    EMWaitReady 1, 1
  loop until SELF_check = "Select Function Menu (SELF)"

'Now we go into STAT/REVW for each case.
  EMSendKey "<home>" + "stat"
  EMSetCursor 18, 43
  EMSendKey "        "
  EMWriteScreen case_number, 18, 43
  EMSetCursor 21, 70
  EMSendKey "revw" + "<enter>"
  EMWaitReady 1, 1

  EMReadScreen error_prone_check, 4, 2, 52
  If error_prone_check = "ERRR" then EMSendKey "revw" + "<enter>"
  If error_prone_check = "ERRR" then EMWaitReady 1, 1

  EMReadScreen cash_status, 1, 7, 40
  EMReadScreen FS_status, 1, 7, 60
  EMReadScreen HC_status, 1, 7, 73
  EMReadScreen renewal_received_check, 2, 13, 37
  If renewal_received_check = "__" then renewal_received = "No"
  If renewal_received_check <> "__" then renewal_received = "Yes"

  If cash_status <> "_" then ObjExcel.Cells(case_row, 6).Value = "cash closed, no reason determined"
  If cash_status = "T" or cash_status = "N" then ObjExcel.Cells(case_row, 6).Value = "cash closed, no CAF"
  If (cash_status = "T" or cash_status = "N") and renewal_received = "Yes" then ObjExcel.Cells(case_row, 6).Value = "cash closed, incomplete review"


  If FS_status <> "_" then ObjExcel.Cells(case_row, 7).Value = "FS closed, no reason determined"
  If FS_status = "T" or FS_status = "N" then EMWriteScreen "x", 5, 58
  If FS_status = "T" or FS_status = "N" then EMSendKey "<enter>"
  If FS_status = "T" or FS_status = "N" then EMWaitReady 1, 1
  If FS_status = "T" or FS_status = "N" then EMReadScreen SR_footer_month_FS, 2, 9, 26
  If FS_status = "T" or FS_status = "N" then EMReadScreen SR_footer_year_FS, 2, 9, 32
  If (FS_status = "T" or FS_status = "N") and (cint(SR_footer_month_FS) = cint(footer_month) and cint(SR_footer_year_FS) = cint(footer_year)) then ObjExcel.Cells(case_row, 7).Value = "FS closed, no CSR."
  If FS_status = "T" or FS_status = "N" then EMReadScreen ER_footer_month_FS, 2, 9, 64
  If FS_status = "T" or FS_status = "N" then EMReadScreen ER_footer_year_FS, 2, 9, 70
  If (FS_status = "T" or FS_status = "N") and (cint(ER_footer_month_FS) = cint(footer_month) and cint(ER_footer_year_FS) = cint(footer_year)) then ObjExcel.Cells(case_row, 7).Value = "FS closed, no CAF renewal."
    If (ObjExcel.Cells(case_row, 7).Value = "FS closed, no CSR." or ObjExcel.Cells(case_row, 7).Value = "FS closed, no CAF renewal.") and renewal_received = "Yes" then ObjExcel.Cells(case_row, 7).Value = "FS closed, incomplete renewal"
  If FS_status = "T" or FS_status = "N" then EMSendKey "<enter>"
  If FS_status = "T" or FS_status = "N" then EMWaitReady 1, 1


  IF HC_status <> "_" then ObjExcel.Cells(case_row, 8).Value = "HC closed, no reason determined"
  If HC_status = "T" or HC_status = "N" then EMWriteScreen "x", 5, 71
  If HC_status = "T" or HC_status = "N" then EMSendKey "<enter>"
  If HC_status = "T" or HC_status = "N" then EMWaitReady 1, 1
  If HC_status = "T" or HC_status = "N" then EMReadScreen SR_footer_month_HC, 2, 8, 27
    If (HC_status = "T" or HC_status = "N") and SR_footer_month_HC = "__" then EMReadScreen SR_footer_month_HC, 2, 8, 71
  If HC_status = "T" or HC_status = "N" then EMReadScreen SR_footer_year_HC, 2, 8, 33
    If (HC_status = "T" or HC_status = "N") and SR_footer_year_HC = "__" then EMReadScreen SR_footer_year_HC, 2, 8, 77
  If (HC_status = "T" or HC_status = "N") and (cint(SR_footer_month_HC) = cint(footer_month) and cint(SR_footer_year_HC) = cint(footer_year)) then ObjExcel.Cells(case_row, 8).Value = "HC closed, no CSR."
  If HC_status = "T" or HC_status = "N" then EMReadScreen ER_footer_month_HC, 2, 9, 27
  If HC_status = "T" or HC_status = "N" then EMReadScreen ER_footer_year_HC, 2, 9, 33
  If (HC_status = "T" or HC_status = "N") and (cint(ER_footer_month_HC) = cint(footer_month) and cint(ER_footer_year_HC) = cint(footer_year)) then ObjExcel.Cells(case_row, 8).Value = "HC closed, no HC renewal."
    If (ObjExcel.Cells(case_row, 8).Value = "HC closed, no CSR." or ObjExcel.Cells(case_row, 8).Value = "HC closed, no HC renewal.") and renewal_received = "Yes" then ObjExcel.Cells(case_row, 8).Value = "HC closed, incomplete renewal"
  If HC_status = "T" or HC_status = "N" then EMSendKey "<enter>"
  If HC_status = "T" or HC_status = "N" then EMWaitReady 1, 1


  If (ObjExcel.Cells(case_row, 6).Value = "cash closed, no CAF" or ObjExcel.Cells(case_row, 6).Value = "cash closed, incomplete review" or ObjExcel.Cells(case_row, 6).Value = "") and _
     (ObjExcel.Cells(case_row, 7).Value = "FS closed, no CSR." or ObjExcel.Cells(case_row, 7).Value = "FS closed, no CAF renewal." or ObjExcel.Cells(case_row, 7).Value = "FS closed, incomplete renewal" or ObjExcel.Cells(case_row, 7).Value = "") and _
     (ObjExcel.Cells(case_row, 8).Value = "HC closed, no CSR." or ObjExcel.Cells(case_row, 8).Value = "HC closed, no HC renewal." or ObjExcel.Cells(case_row, 8).Value = "HC closed, incomplete renewal" or ObjExcel.Cells(case_row, 8).Value = "") and _
     (ObjExcel.Cells(case_row, 6).Value <> "" or ObjExcel.Cells(case_row, 7).Value <> "" or ObjExcel.Cells(case_row, 8).Value <> "") then _
     needs_case_note = "True"


  If needs_case_note = "True" then EMSendKey "<PF4>"
  If needs_case_note = "True" then EMWaitReady 1, 1
  If needs_case_note = "True" then EMSendKey "<PF9>"
  If needs_case_note = "True" then EMWaitReady 1, 1
  If needs_case_note = "True" then EMSetCursor 4, 3
  If needs_case_note = "True" then EMSendKey "---Closed case for " & footer_month & "/" & footer_year & "---" & "<newline>"
  If needs_case_note = "True" and renewal_received = "Yes" then EMSendKey "* Incomplete renewal. See previous case notes for requested verifications. If they are not received by the end of " & footer_month & "/" & footer_year & ", a new HCAPP or CAF is required." & "<newline>"
  If needs_case_note = "True" and renewal_received = "No" and ObjExcel.Cells(case_row, 6).Value = "cash closed, no CAF" then EMSendKey "* Cash: no CAF renewal. A new CAF is required." & "<newline>"
  If needs_case_note = "True" and renewal_received = "Yes" and ObjExcel.Cells(case_row, 6).Value = "cash closed, incomplete review" then EMSendKey "* Cash is closed for incomplete renewal." & "<newline>"
  If needs_case_note = "True" and renewal_received = "No" and ObjExcel.Cells(case_row, 7).Value = "FS closed, no CSR." then EMSendKey "* FS: no CSR renewal. A new CSR is required. If it is not received by the end of " & footer_month & "/" & footer_year & ", a new CAF is required." & "<newline>"
  If needs_case_note = "True" and renewal_received = "No" and ObjExcel.Cells(case_row, 7).Value = "FS closed, no CAF renewal." then EMSendKey "* FS: no CAF renewal. A new CAF is required." & "<newline>"
  If needs_case_note = "True" and renewal_received = "Yes" and ObjExcel.Cells(case_row, 7).Value = "FS closed, incomplete renewal" then EMSendKey "* FS is closed for incomplete renewal." & "<newline>"
  If needs_case_note = "True" and renewal_received = "No" and ObjExcel.Cells(case_row, 8).Value = "HC closed, no CSR." then EMSendKey "* HC: no CSR renewal. A new CSR is required. If it is not received by the end of " & footer_month & "/" & footer_year & ", a new HCAPP is required. Client may use CAF for HC if applying for other programs as well." & "<newline>"
  If needs_case_note = "True" and renewal_received = "No" and ObjExcel.Cells(case_row, 8).Value = "HC closed, no HC renewal." then EMSendKey "* HC: no renewal. A new HC renewal is required. If it is not received by the end of " & footer_month & "/" & footer_year & ", a new HCAPP is required. Client may use CAF for HC if applying for other programs as well." & "<newline>"
  If needs_case_note = "True" and renewal_received = "Yes" and ObjExcel.Cells(case_row, 8).Value = "HC closed, incomplete renewal" then EMSendKey "* HC is closed for incomplete renewal." & "<newline>"
  If needs_case_note = "True" then EMSendKey "* This case remains in worker number until the last day of " & footer_month & "/" & footer_year & ". After that, the client has to go through intake or HCAPP rotation." & "<newline>"
  If needs_case_note = "True" then EMSendKey "---" & "<newline>"
  If needs_case_note = "True" then EMSendKey worker_sig & ", using automated script."

  needs_case_note = "False" 'Resetting variable.

  case_row = case_row + 1
loop




MsgBox "All of your " & footer_month & "/" & footer_year & " cases that closed for no review have been case noted. A spreadsheet has been generated which displays what closed and why. You can print this and use it to manually track any additional changes, if you'd like."