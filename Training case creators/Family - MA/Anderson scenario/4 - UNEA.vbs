amt_of_times_to_run = 8 'Max is 12
UNEA_member = "03"
UNEA_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
UNEA_type = "36"
UNEA_claim_number = ""
UNEA_start_date_month = "01"
UNEA_start_date_day = "01"
UNEA_start_date_year = "12"
pay_freq = 2 '1 for monthly, 2 for semi-monthly, 3 for biweekly and 4 for weekly.
pay_amount = 60
update_PIC = False 'Not programmed yet, need to block out a line and enter PIC manually each time.
update_HC_income_estimator = False



current_month = datepart("m", date)
if len(current_month) < 2 then current_month = "0" & current_month
current_day = datepart("d", date)
if len(current_day) < 2 then current_day = "0" & current_day
current_year = datepart("yyyy", date) - 2000

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
  EMWaitReady 1, 1
  
  EMWriteScreen "UNEA", 20, 71
  EMWriteScreen UNEA_member, 20, 76
  EMWriteScreen UNEA_action, 20, 79
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  If UNEA_action <> "nn" then
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
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
    EMWaitReady 1, 1

    EMWriteScreen "________", 9, 65
    EMWriteScreen pay_amount, 9, 65
    EMWriteScreen pay_freq, 10, 63
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<enter>"
    EMWaitReady 1, 1
  End if

  If update_PIC = True then 
    EMWriteScreen "x", 10, 26
    EMSendKey "<enter>"
    EMWaitReady 1, 1

    EMWriteScreen current_month, 5, 34
    EMWriteScreen current_day, 5, 37
    EMWriteScreen current_year, 5, 40
    EMWriteScreen pay_freq, 5, 64

    EMSendKey "<enter>"
    EMWaitReady 1, 1
  End if
  
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen PND2_check, 4, 2, 52
    If PND2_check = "LF) " then
      MsgBox "error"
      stopscript
    End if
  Loop until PND2_check = "PND2"
  
  MAXIS_row = MAXIS_row + 1
  payday_array = ""
Loop until MAXIS_row = amt_of_times_to_run + 7