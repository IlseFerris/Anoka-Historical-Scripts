case_number_array = array("179473", "179474", "179475", "178693", "178349", "179105", "179467", "179468", "179469", "179470", "179471", "179472")

EMConnect ""

For each case_number in case_number_array

  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
  
  EMWriteScreen "dail", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "elig", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  row = 6 'Setting to 6 as a test
  col = 1
  
  EMSearch "11 12", row, col
  
  If row <> 0 then
    EMWriteScreen "d", row, 3
    Do
      row = row + 1
      EMReadScreen next_month_check, 5, row, 11
      If next_month_check = "11 12" then EMWriteScreen "d", row, 3
    Loop until next_month_check <> "11 12"
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if
  
Next