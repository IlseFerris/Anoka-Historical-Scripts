amt_of_times_to_run = 10

siblings = array("03", "04", "05") 
SIBL_action = "nn"

CARS_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Leave blank to skip.
CARS_member = "02"
CARS_type = "1" '1 for Car, 2 for truck, 3 for van, 4 for camper, 5 for motorcycle, 6 for trailer, 7 for other
CARS_year = "1999"
CARS_make = "Dodge"
CARS_model = "Caravan"
CARS_value = "850"

DISA_member = "05"
DISA_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
DISA_type = "03"
DISA_start_date_month = "10"
DISA_start_date_day = "01"
DISA_start_date_year = "2013"
cash_status = "03"
SNAP_status = ""
HC_status = ""

UNEA_member = "05"
UNEA_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
UNEA_type = "03"
UNEA_claim_number = "474474474DI"
UNEA_start_date_month = "01"
UNEA_start_date_day = "03"
UNEA_start_date_year = "14"
pay_freq = 1 '1 for monthly, 2 for semi-monthly, 3 for biweekly and 4 for weekly.
pay_amount = 710
update_PIC = False 'Not programmed yet, need to block out a line and enter PIC manually each time.
update_HC_income_estimator = False

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

  If DISA_action <> "" then 'Updating DISA----------------------------------------------------------------------------------------------------
    EMWriteScreen "DISA", 20, 71
    EMWriteScreen DISA_member, 20, 76
    EMWriteScreen DISA_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  
    If DISA_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
  
    EMWriteScreen DISA_start_date_month, 6, 47
    EMWriteScreen DISA_start_date_day, 6, 50
    EMWriteScreen DISA_start_date_year, 6, 53
  
    If cash_status <> "" then
      EMWriteScreen cash_status, 11, 59
      EMWriteScreen "3", 11, 69
    End if
  
    If SNAP_status <> "" then
      EMWriteScreen SNAP_status, 12, 59
      EMWriteScreen "3", 12, 69
    End if
  
    If HC_status <> "" then
      EMWriteScreen HC_status, 13, 59
      EMWriteScreen "3", 13, 69
    End if
    
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  If SIBL_action <> "" then 'Updating SIBL----------------------------------------------------------------------------------------------------
    EMWriteScreen "SIBL", 20, 71
    EMWriteScreen SIBL_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  
    col = 39
  
    EMWriteScreen "01", 7, 28
  
    For each kid in siblings
      EMWriteScreen kid, 7, col
      col = col + 4
    Next
    
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  If CARS_action <> "" then 'Updating CARS----------------------------------------------------------------------------------------------------
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
    EMWriteScreen "5", 10, 60
    EMWriteScreen "4", 9, 80
    EMWriteScreen "1", 15, 43
    EMWriteScreen "Y", 15, 76
    EMWriteScreen "N", 16, 43
  
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  If UNEA_action <> "" then 'Updating UNEA----------------------------------------------------------------------------------------------------
    EMWriteScreen "UNEA", 20, 71
    EMWriteScreen UNEA_member, 20, 76
    EMWriteScreen UNEA_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  
    If UNEA_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
  
    EMReadScreen footer_month_year, 5, 20, 55
    footer_month_year = replace(footer_month_year, " ", "/01/")
    
    'Now it figures out what the paydays would be. It assumes a friday payday.
    date_start_plus_1 = DateAdd("d", 1, footer_month_year) 
    date_start_plus_2 = DateAdd("d", 2, footer_month_year) 
    date_start_plus_3 = DateAdd("d", 3, footer_month_year) 
    date_start_plus_4 = DateAdd("d", 4, footer_month_year) 
    date_start_plus_5 = DateAdd("d", 5, footer_month_year) 
    date_start_plus_6 = DateAdd("d", 6, footer_month_year) 
    If Weekday(footer_month_year, 0) = 6 then first_payday = (footer_month_year)
    If Weekday(date_start_plus_1, 0) = 6 then first_payday = (date_start_plus_1)
    If Weekday(date_start_plus_2, 0) = 6 then first_payday = (date_start_plus_2)
    If Weekday(date_start_plus_3, 0) = 6 then first_payday = (date_start_plus_3)
    If Weekday(date_start_plus_4, 0) = 6 then first_payday = (date_start_plus_4)
    If Weekday(date_start_plus_5, 0) = 6 then first_payday = (date_start_plus_5)
    If Weekday(date_start_plus_6, 0) = 6 then first_payday = (date_start_plus_6)
    If pay_freq = 1 then
      second_payday = ""
      third_payday = ""
      fourth_payday = ""
      fifth_payday = ""
    End if
    If pay_freq = 2 then
      second_payday = dateadd("d", 15, first_payday)
      third_payday = ""
      fourth_payday = ""
      fifth_payday = ""
    End if
    If pay_freq = 3 then
      second_payday = dateadd("d", 14, first_payday)
      third_payday = dateadd("d", 14, second_payday)
      fourth_payday = ""
      fifth_payday = ""
      If datepart("m", third_payday) <> datepart("m", first_payday) then third_payday = ""
    End if
    If pay_freq = 4 then
      second_payday = dateadd("d", 7, first_payday)
      third_payday = dateadd("d", 7, second_payday)
      fourth_payday = dateadd("d", 7, third_payday)
      fifth_payday = dateadd("d", 7, fourth_payday)
      If datepart("m", fifth_payday) <> datepart("m", first_payday) then fifth_payday = ""
    End if
    If first_payday <> "" then payday_array = payday_array & " " & first_payday
    If second_payday <> "" then payday_array = payday_array & " " & second_payday
    If third_payday <> "" then payday_array = payday_array & " " & third_payday
    If fourth_payday <> "" then payday_array = payday_array & " " & fourth_payday
    If fifth_payday <> "" then payday_array = payday_array & " " & fifth_payday
    payday_array = split(trim(payday_array))
    
    'Now it writes the payday info into MAXIS
    row = 13
    For each payday in payday_array
      payday = cdate(payday)
      payday_month = datepart("m", payday)
      If len(payday_month) = 1 then payday_month = "0" & payday_month
      payday_day = datepart("d", payday)
      If len(payday_day) = 1 then payday_day = "0" & payday_day
      payday_year = datepart("yyyy", payday) - 2000
    
      EMWriteScreen payday_month, row, 54
      EMWriteScreen payday_day, row, 57
      EMWriteScreen payday_year, row, 60
      EMWriteScreen "________", row, 68
      EMWriteScreen pay_amount, row, 68
    
      row = row + 1
    Next
     
    'Writes info about the income into MAXIS.
    EMWriteScreen "6", 5, 65
    EMWriteScreen UNEA_type, 5, 37
    EMWriteScreen UNEA_claim_number, 6, 37
    EMWriteScreen UNEA_start_date_month, 7, 37
    EMWriteScreen UNEA_start_date_day, 7, 40
    EMWriteScreen UNEA_start_date_year, 7, 43
    EMWriteScreen pay_freq, 18, 35
  
    If update_HC_income_estimator = True then 
      EMWriteScreen "x", 6, 56
      EMSendKey "<enter>"
      EMWaitReady 0, 0
  
      EMWriteScreen "________", 9, 65
      EMWriteScreen pay_amount, 9, 65
      EMWriteScreen pay_freq, 10, 63
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
  
    If update_PIC = True then 
      EMWriteScreen "x", 10, 26
      EMSendKey "<enter>"
      EMWaitReady 0, 0
  
      EMWriteScreen current_month, 5, 34
      EMWriteScreen current_day, 5, 37
      EMWriteScreen current_year, 5, 40
      EMWriteScreen pay_freq, 5, 64
  
  MsgBox "" '<<<<<<<<<<<<ERROR PROOFING?
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
    
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    payday_array = "" 'Blanking out the array

  End if

    
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
Loop until MAXIS_row = amt_of_times_to_run + 7