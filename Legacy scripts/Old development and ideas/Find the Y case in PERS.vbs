  Do
  row = 1
  col = 1
EMReadScreen PMI_NBR_assigned_check, 7, 24, 2
IF PMI_NBR_assigned_check = "PMI NBR" then row = 0
IF PMI_NBR_assigned_check = "PMI NBR" then exit do
  EMSearch "   Y    ", row, 61
  If row <> 0 then EMReadScreen case_number, 8, row, 6

  If row = 0 then EMSendKey "<PF8>"
  If row = 0 then EMWaitReady 1, 1
  If row = 0 then EMReadScreen last_page_check, 21, 24, 2
  Loop until last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "THIS IS THE ONLY PAGE" or row <> 0