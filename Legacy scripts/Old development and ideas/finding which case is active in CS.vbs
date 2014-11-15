EMConnect ""


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
