amt_of_times_to_run = 9

kids_to_add = array("03")

ABPS_action = ""
ABPS_last_name = "Smith"
ABPS_first_name = "Smittie"
ABPS_gender = "M"

PARE_action = ""
both_parents_in_HH = True
stepparent_in_HH = False

EATS_action = "" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Blank out if not used
PP_together = "Y"
EATS_member_array = array("01", "02") 
EATS_non_member_array = array() 

EMPS_action = "" 'at this time it just puts "n" for the other stuff. Blank this out to skip. If updating an existing panel, put "01".
fin_orient_dt_month = "04"
fin_orient_dt_day = "10"
fin_orient_dt_year = "13"
full_time_care_of_child_under_1 = "Y"
EMPS_Exemption_Care_Of_A_Child_Under_One_month = "01"
EMPS_Exemption_Care_Of_A_Child_Under_One_year = "2013"

SHEL_action = "" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Blank it out if not used ("").
SHEL_member = "01"
subsidized_indicator = "n"
shared_indicator = "n"
paid_to_memb = ""
paid_to_name = "lucky landlord"
rent_amt = "400"
rent_proof = "ot"
lot_rent_amt = ""
lot_rent_proof = ""
mortgage_amt = ""
mortgage_proof = ""
insurance_amt = ""
insurance_proof = ""
taxes_amt = ""
taxes_proof = ""
room_amt = ""
room_proof = ""
garage_amt = ""
garage_proof = ""
subsidy_amt = ""
subsidy_proof = ""

HEST_action = "" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Blank it out if not used ("").
HEST_heat_AC_indicator = "Y"
HEST_electric_indicator = ""
HEST_phone_indicator = ""

JOBS_action = "" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Blank it out if not used ("").
JOBS_member = "01"
JOBS_location = "Tudor Time"
JOBS_start_date_month = "12"
JOBS_start_date_day = "10"
JOBS_start_date_year = "11"
JOBS_proof = "1"
pay_freq = 3 '1 for monthly, 2 for semi-monthly, 3 for biweekly and 4 for weekly. Biweekly is buggy on hours.
pay_amount = 500
hours_per_check = 25 'THIS IS REALLY HOURS PER WEEK
update_PIC =  True 'Just does anticipated
update_future_months = True 

ACCT_action = "" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit.
ACCT_member = "01"
ACCT_type = "CK"
ACCT_number = "123456789012"
ACCT_location = "Wells Fargo"
ACCT_balance = "900"
ACCT_as_of_month = "12"
ACCT_as_of_day = "01"
ACCT_as_of_year = "12"
cash_count_status = "Y"
SNAP_count_status = "Y"
HC_count_status = "Y"

WREG_action = "nn" 'If creating a new one, this should be "nn", in lower case, otherwise, it should be the panel to edit. Leave blank to ignore this panel.
WREG_member_array = array("01") 
FSET_status = "30"
defer_FSET_indicator = "N" 'Should usually be a "N" or a "_" for exempt people.
ABAWD_status = "09"
GA_basis = ""


current_month = datepart("m", date)
if len(current_month) < 2 then current_month = "0" & current_month
current_day = datepart("d", date)
if len(current_day) < 2 then current_day = "0" & current_day
current_year = datepart("yyyy", date) - 2000




'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to MAXIS
EMConnect ""

'Checking for PND2. If not on PND2 it'll stop.
EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then
  MsgBox "Not on PND2"
  StopScript
End if

'Setting the MAXIS row to look at.
MAXIS_row = 7

'Grabs the footer month and year because it might need the original if we're updating future months in cases.
EMReadScreen PND2_footer_month, 2, 20, 55
EMReadScreen PND2_footer_year, 2, 20, 58

'OPERATES AS A DO...LOOP TO UPDATE EVERY CASE ON PND2
Do
  'Writing the original footer month in case we updated future months
  EMWriteScreen PND2_footer_month, 20, 55
  EMWriteScreen PND2_footer_year, 20, 58

  'gets into STAT for the case
  EMWriteScreen "s", MAXIS_row, 3
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  If ABPS_action <> "" then 'Updating ABPS----------------------------------------------------------------------------------------------------
    If both_parents_in_HH = False then
      EMWriteScreen "ABPS", 20, 71
      EMWriteScreen ABPS_action, 20, 79
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      
      EMWriteScreen "01", 4, 47
      EMWriteScreen "Y", 4, 73
      EMWriteScreen "N", 5, 47
      EMWriteScreen ABPS_last_name, 10, 30
      EMWriteScreen ABPS_first_name, 10, 63
      EMWriteScreen ABPS_gender, 11, 80
      row = 15
     
      For each kid in kids_to_add
        EMWriteScreen kid, row, 35
        EMWriteScreen "1", row, 53
        EMWriteScreen "1", row, 67
        row = row + 1
      Next
      
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
  End if

  If ACCT_action <> "" then 'Updating ACCT----------------------------------------------------------------------------------------------------
    EMWriteScreen "ACCT", 20, 71
    EMWriteScreen ACCT_member, 20, 76
    EMWriteScreen ACCT_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  
    If ACCT_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
  
    'Clears out existing info
    EMSendKey string(79, "_")
  
    EMWriteScreen ACCT_type, 6, 44
    EMWriteScreen ACCT_number, 7, 44
    EMWriteScreen ACCT_location, 8, 44
    EMWriteScreen ACCT_balance, 10, 46
    EMWriteScreen "5", 10, 63
    EMWriteScreen ACCT_as_of_month, 11, 44
    EMWriteScreen ACCT_as_of_day, 11, 47
    EMWriteScreen ACCT_as_of_year, 11, 50
  
    EMWriteScreen cash_count_status, 14, 50
    EMWriteScreen SNAP_count_status, 14, 57
    EMWriteScreen HC_count_status, 14, 64
    EMWriteScreen "N", 15, 44
    
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  If EATS_action <> "" then 'Updating EATS----------------------------------------------------------------------------------------------------
    EMWriteScreen "EATS", 20, 71
    EMWriteScreen EATS_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  
    If EATS_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
  
    EMWriteScreen PP_together, 4, 72
    EMWriteScreen "N", 5, 72
  
    If PP_together = "N" then
      EMWriteScreen "01", 13, 28
    
      col = 39
      For each EATS_member in EATS_member_array  
        EMWriteScreen EATS_member, 13, col
        col = col + 4
      Next
    
      col = 39
      EMWriteScreen "02", 14, 28
      For each EATS_non_member in EATS_non_member_array  
        EMWriteScreen EATS_non_member, 14, col
        col = col + 4
      Next
    End if
  
    EMSendKey "<enter>"
    EMWaitReady 0, 0  
  End if
 
  If PARE_action <> "" then 'Updating PARE----------------------------------------------------------------------------------------------------
    EMWriteScreen "PARE", 20, 71
    EMWriteScreen PARE_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    
    row = 8
    
    For each kid in kids_to_add
      EMWriteScreen kid, row, 24
      EMWriteScreen "1", row, 53
      EMWriteScreen "OT", row, 71
      row = row + 1
    Next
    
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    
    If both_parents_in_HH = True then
      EMWriteScreen "PARE", 20, 71
      EMWriteScreen "02", 20, 76
      EMWriteScreen PARE_action, 20, 79
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      
      row = 8
      
      For each kid in kids_to_add
        EMWriteScreen kid, row, 24
        EMWriteScreen "1", row, 53
        EMWriteScreen "OT", row, 71
        row = row + 1
      Next
      
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
  
    If stepparent_in_HH = True then
      EMWriteScreen "PARE", 20, 71
      EMWriteScreen "02", 20, 76
      EMWriteScreen PARE_action, 20, 79
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      
      row = 8
      
      For each kid in kids_to_add
        EMWriteScreen kid, row, 24
        EMWriteScreen "2", row, 53
        EMWriteScreen "OT", row, 71
        row = row + 1
      Next
    
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
  End if

  If EMPS_action <> "" then 'Updating EMPS----------------------------------------------------------------------------------------------------
    EMWriteScreen "EMPS", 20, 71
    EMWriteScreen EMPS_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    If EMPS_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if

    EMWriteScreen fin_orient_dt_month, 5, 39
    EMWriteScreen fin_orient_dt_day, 5, 42
    EMWriteScreen fin_orient_dt_year, 5, 45
    EMWriteScreen "n", 8, 76
    EMWriteScreen "n", 9, 76
    EMWriteScreen "n", 10, 76
    EMWriteScreen "no", 11, 76
    EMWriteScreen full_time_care_of_child_under_1, 12, 76
    EMWriteScreen "n", 13, 76

    If full_time_care_of_child_under_1 = "Y" then
      EMWriteScreen "x", 12, 39
      EMSendKey "<enter>"
      EMWaitReady 0, 0

      EMWriteScreen EMPS_Exemption_Care_Of_A_Child_Under_One_month, 7, 22
      EMWriteScreen EMPS_Exemption_Care_Of_A_Child_Under_One_year, 7, 27
      EMSendKey "<enter>"
      EMWaitReady 0, 0

      EMSendKey "<PF3>"
      EMWaitReady 0, 0
    End if

    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  If SHEL_action <> "" then 'Updating SHEL----------------------------------------------------------------------------------------------------
    EMWriteScreen "shel", 20, 71
    EMWriteScreen SHEL_member, 20, 76
    EMWriteScreen SHEL_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    
    If SHEL_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
    
    EMSetCursor 6, 42
    EMSendKey(string(189, "_"))
  
    EMWriteScreen subsidized_indicator, 6, 42
    EMWriteScreen shared_indicator, 6, 60
    EMWriteScreen paid_to_MEMB, 7, 42
    EMWriteScreen paid_to_name, 7, 46
  
    EMWriteScreen rent_amt, 11, 56
    EMWriteScreen rent_proof, 11, 67
    EMWriteScreen lot_rent_amt, 12, 56
    EMWriteScreen lot_rent_proof, 12, 67
    EMWriteScreen mortgage_amt, 13, 56
    EMWriteScreen mortgage_proof, 13, 67
    EMWriteScreen insurance_amt, 14, 56
    EMWriteScreen insurance_proof, 14, 67
    EMWriteScreen taxes_amt, 15, 56
    EMWriteScreen taxes_proof, 15, 67
    EMWriteScreen room_amt, 16, 56
    EMWriteScreen room_proof, 16, 67
    EMWriteScreen garage_amt, 17, 56
    EMWriteScreen garage_proof, 17, 67
    EMWriteScreen subsidy_amt, 18, 56
    EMWriteScreen subsidy_proof, 18, 67

    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if

  If HEST_action <> "" then 'Updating HEST----------------------------------------------------------------------------------------------------
    EMWriteScreen "HEST", 20, 71
    EMWriteScreen HEST_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  
    If HEST_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
  
    EMSetCursor 6, 40
    EMSendKey(string(52, "_"))

    EMWriteScreen "01", 6, 40
    EMWriteScreen "01", 7, 40
    EMWriteScreen "01", 7, 43
    EMWriteScreen "01", 7, 46

    EMWriteScreen HEST_heat_AC_indicator, 13, 60
    If HEST_heat_AC_indicator <> "" then EMWriteScreen "01", 13, 68
    EMWriteScreen HEST_electric_indicator, 14, 60
    If HEST_electric_indicator <> "" then EMWriteScreen "01", 15, 68
    EMWriteScreen HEST_phone_indicator, 15, 60
    If HEST_phone_indicator <> "" then EMWriteScreen "01", 15, 68

    EMSendKey "<enter>"
    EMWaitReady 0, 0

  End if

  If WREG_action <> "" then 'Updating WREG----------------------------------------------------------------------------------------------------
    For each WREG_member in WREG_member_array  
      EMWriteScreen "WREG", 20, 71
      EMWriteScreen WREG_member, 20, 76
      EMWriteScreen WREG_action, 20, 79
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    
      If WREG_action <> "nn" then
        EMSendKey "<PF9>"
        EMWaitReady 0, 0
      End if
    
      If WREG_member = "01" then EMWriteScreen "Y", 6, 68
      If WREG_member <> "01" then EMWriteScreen "N", 6, 68
      EMWriteScreen FSET_status, 8, 50
      EMWriteScreen defer_FSET_indicator, 8, 80
      EMWriteScreen ABAWD_status, 13, 50
      EMWriteScreen GA_basis, 15, 50
  
      EMSendKey "<enter>"
      EMWaitReady 0, 0

      EMSendKey "<enter>"
      EMWaitReady 0, 0  
    Next
  End if

  If JOBS_action <> "" then 'Updating JOBS, does this last for multi-month purposes----------------------------------------------------------------------------------------------------
    EMWriteScreen "JOBS", 20, 71
    EMWriteScreen JOBS_member, 20, 76
    EMWriteScreen JOBS_action, 20, 79
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    
    If JOBS_action <> "nn" then
      EMSendKey "<PF9>"
      EMWaitReady 0, 0
    End if
  
    EMReadScreen footer_month_year, 5, 20, 55
    footer_month_year = replace(footer_month_year, " ", "/01/")
  
    'Clears out existing info
    EMSendKey string(214, "_")
    
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
    row = 12
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
      EMWriteScreen "________", row, 67
      EMWriteScreen pay_amount, row, 67
    
      row = row + 1
    Next
    
    'Writes hours into MAXIS
    monthly_hours = hours_per_check * (ubound(payday_array) + 1)
    EMWriteScreen "___", 18, 72
    EMWriteScreen monthly_hours, 18, 72
    
    'Writes info about the job into MAXIS.
    EMWriteScreen "W", 5, 38
    EMWriteScreen JOBS_proof, 6, 38
    EMWriteScreen JOBS_location, 7, 42
    EMWriteScreen JOBS_start_date_month, 9, 35
    EMWriteScreen JOBS_start_date_day, 9, 38
    EMWriteScreen JOBS_start_date_year, 9, 41
    EMWriteScreen pay_freq, 18, 35

    If update_PIC = True then 
      EMWriteScreen "x", 19, 38
      EMSendKey "<enter>"
      EMWaitReady 0, 0
  
      EMWriteScreen current_month, 5, 34
      EMWriteScreen current_day, 5, 37
      EMWriteScreen current_year, 5, 40
      EMWriteScreen pay_freq, 5, 64
  
      EMWriteScreen hours_per_check, 8, 64
      EMWriteScreen pay_amount/hours_per_check, 9, 66
  
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
    
    If update_future_months = True then 
      Do
        EMSendKey "<enter>"
        EMWaitReady 0, 0
        EMWriteScreen "bgtx", 20, 71
        EMSendKey "<enter>"
        EMWaitReady 0, 0
  
        EMWriteScreen "y", 16, 54
        EMSendKey "<enter>"
        EMWaitReady 0, 0
  
        EMWriteScreen "jobs", 20, 71
        EMWriteScreen JOBS_member, 20, 76
        If JOBS_action = "nn" then
          EMWriteScreen "01", 20, 79
        Else
          EMWriteScreen JOBS_action, 20, 79
        End if
        EMSendKey "<enter>"
        EMWaitReady 0, 0
        EMSendKey "<PF9>"
        EMWaitReady 0, 0
  
        EMReadScreen current_footer_month, 2, 20, 55
        EMReadScreen current_footer_year, 2, 20, 58
  
        JOBS_line_row = 12
        Do
          EMReadScreen JOBS_line_day, 2, JOBS_line_row, 57
          If isnumeric(JOBS_line_day) = False then exit do
          If isdate(current_footer_month & "/" & JOBS_line_day & "/" & current_footer_year) = True Then 
            EMWriteScreen current_footer_month, JOBS_line_row, 54
            EMWriteScreen current_footer_year, JOBS_line_row, 60
          Else
            EMWriteScreen "__", JOBS_line_row, 54
            EMWriteScreen "__", JOBS_line_row, 57
            EMWriteScreen "__", JOBS_line_row, 60
            EMWriteScreen "________", JOBS_line_row, 67
          End if
          JOBS_line_row = JOBS_line_row + 1
        Loop until JOBS_line_row = 17
        first_of_current_month = current_footer_month & "/01/" & current_footer_year
        first_of_next_month = datepart("m", dateadd("m", 1, date)) & "/01/" & datepart("yyyy", dateadd("m", 1, date))
      Loop until cdate(first_of_next_month) = cdate(first_of_current_month)

      EMWriteScreen "x", 19, 54
      EMSendKey "<enter>"
      EMWaitReady 0, 0
  
      EMWriteScreen "________", 11, 63
      EMWriteScreen pay_amount, 11, 63
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
    
    EMSendKey "<enter>"
    EMWaitReady 0, 0

    payday_array = ""
  End if  

  Do 'Exiting the case----------------------------------------------------------------------------------------------------
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