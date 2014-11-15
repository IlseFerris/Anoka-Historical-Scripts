'VARIABLES TO DECLARE


EMConnect ""

EMReadScreen ACTV_check, 4, 2, 48
If ACTV_check <> "ACTV" then
  MsgBox "Not on ACTV."
  StopScript
End if

EMReadScreen x102_number, 7, 21, 13

row = 7

Do
  EMReadScreen case_number, 8, row, 12
  If case_number <> "        " then
    case_number_array = case_number_array & trim(case_number) & " "
  Else
    EMReadScreen case_count, 4, row - 1, 5
  End if
  row = row + 1
  If row = 19 then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    row = 7
    EMReadScreen last_page_check, 4, 24, 19
  End if
Loop until last_page_check = "PAGE" or case_number = "        "

case_number_array = split(case_number_array)

For each case_number in case_number_array
  If case_number <> "" then
    Do
      EMSendKey "<PF3>"
      EMWaitReady 0, 0
      EMReadScreen SELF_check, 4, 2, 50
    Loop until SELF_check = "SELF"
  
    EMWriteScreen "case", 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen case_number, 18, 43
    EMWriteScreen "note", 21, 70
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    Do
      EMReadScreen case_note_eff_date_last, 8, 18, 6
      If case_note_eff_date_last = "        " then case_note_eff_date_last = "01/01/1900" 'This simplifies the function of the do...loop
      row = 1
      col = 1
      EMSearch "***CSR", row, col
      If col = 25 then
        EMReadScreen case_note_worker_number, 7, row, 16
        EMReadScreen case_note_eff_date, 8, row, 6
        If case_note_worker_number = x102_number and dateadd("m", -12, date) < cdate(case_note_eff_date) then 
          cases_with_scripts_used = cases_with_scripts_used + 1
          exit do
        End if
      Else
        If cdate(case_note_eff_date_last) > dateadd("m", -12, date) then
          EMSendKey "<PF8>"
          EMWaitReady 0, 0
        End if
      End if
    Loop until cdate(case_note_eff_date_last) <= dateadd("m", -12, date)
  End if
Next

MsgBox "Done. Cases with scripts used: " & cases_with_scripts_used & ". Total cases: " & case_count & "."