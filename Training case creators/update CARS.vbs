amt_of_times_to_run = 1
CARS_member = "01"
CARS_type = "1" '1 for Car, 2 for truck, 3 for van, 4 for camper, 5 for motorcycle, 6 for trailer, 7 for other
CARS_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
CARS_year = "1999"
CARS_make = "Honda"
CARS_model = "Accord"
CARS_value = "1000"

EMConnect ""

EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then
  MsgBox "Not on PND2"
  StopScript
End if

MAXIS_row = 7

Do
  EMWriteScreen "s", MAXIS_row, 3
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  EMWriteScreen "CARS", 20, 71
  EMWriteScreen CARS_member, 20, 76
  EMWriteScreen CARS_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  If CARS_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 0, 0
  End if

  EMWriteScreen CARS_type, 6, 43
  EMWriteScreen CARS_year, 8, 31
  EMWriteScreen CARS_make, 8, 43
  EMWriteScreen CARS_model, 8, 66
  EMWriteScreen CARS_value, 9, 45
  EMWriteScreen CARS_value, 9, 62
  EMWriteScreen "4", 9, 80
  EMWriteScreen "1", 15, 43
  EMWriteScreen "Y", 15, 76
  EMWriteScreen "N", 16, 43

  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen PND2_check, 4, 2, 52
    If PND2_check = "LF) " then
      MsgBox "error"
      stopscript
    End if
  Loop until PND2_check = "PND2"
  
  MAXIS_row = MAXIS_row + 1
  payday_array = ""
Loop until MAXIS_row = amt_of_times_to_run + 7