'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script contains functions that the other BlueZone scripts use very commonly. The
'other BlueZone scripts contain a few lines of code that run this script and get the 
'functions. This saves me time in writing and copy/pasting the same functions in
'many different places. Only add functions to this script if they've been tested by
'the workgroups. This document is actively used by live scripts, so it needs to be
'functionally complete at all times.
'
'Here's the code to add (without comments of course):
'
''LOADING ROUTINE FUNCTIONS
'Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
'text_from_the_other_script = fso_command.ReadAll
'fso_command.Close
'Execute text_from_the_other_script
'----------------------------------------------------------------------------------------------------

'FUNCTIONS THAT ARE BEING USED:
'          add_ACCI_to_variable
'          add_ACCT_to_variable
'          add_BUSI_to_variable
'          add_CARS_to_variable
'          add_JOBS_to_variable
'          add_OTHR_to_variable
'          add_RBIC_to_variable
'          add_REST_to_variable
'          add_SECU_to_variable
'          add_UNEA_to_variable
'          attn
'          back_to_SELF
'          create_MAXIS_friendly_date
'          end_excel_and_script
'          find_variable
'          memb_navigation_next
'          memb_navigation_prev
'          navigate_to_screen
'          navigation_buttons
'          new_page_check
'          new_service_heading
'          new_BS_BSI_heading
'          new_CAI_heading
'          panel_navigation_next
'          panel_navigation_prev
'          PF1
'          PF2
'          PF3
'          PF4
'          PF5
'          PF6
'          PF7
'          PF8
'          PF9
'          PF10
'          PF11
'          PF12
'          PF20
'          run_another_script
'          script_end_procedure
'          stat_navigation
'          transmit
'          write_editbox_in_case_note
'          write_new_line_in_case_note
'          write_three_columns_in_case_note

Function add_ACCI_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen ACCI_date, 8, 6, 73
  ACCI_date = replace(ACCI_date, " ", "/")
  If datediff("yyyy", ACCI_date, now) < 5 then
    EMReadScreen ACCI_type, 2, 6, 47
    If ACCI_type = "01" then ACCI_type = "Auto"
    If ACCI_type = "02" then ACCI_type = "Workers Comp"
    If ACCI_type = "03" then ACCI_type = "Homeowners"
    If ACCI_type = "04" then ACCI_type = "No Fault"
    If ACCI_type = "05" then ACCI_type = "Other Tort"
    If ACCI_type = "06" then ACCI_type = "Product Liab"
    If ACCI_type = "07" then ACCI_type = "Med Malprac"
    If ACCI_type = "08" then ACCI_type = "Legal Malprac"
    If ACCI_type = "09" then ACCI_type = "Diving Tort"
    If ACCI_type = "10" then ACCI_type = "Motorcycle"
    If ACCI_type = "11" then ACCI_type = "MTC or Other Bus Tort"
    If ACCI_type = "12" then ACCI_type = "Pedestrian"
    If ACCI_type = "13" then ACCI_type = "Other"
    x = x & ACCI_type & " on " & ACCI_date & ".; "
  End if
End function

Function add_ACCT_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen ACCT_amt, 8, 10, 46
  ACCT_amt = trim(ACCT_amt)
  EMReadScreen counted_check, 1, 14, 64
  If counted_check = "N" then
    ACCT_amt = "excluded"
  Else
    ACCT_amt = "$" & ACCT_amt
  End if
  EMReadScreen ACCT_type, 2, 6, 44
  EMReadScreen ACCT_location, 20, 8, 44
  ACCT_location = replace(ACCT_location, "_", "")
  ACCT_location = split(ACCT_location)
  For each a in ACCT_location
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      If len(a) > 3 then
        new_ACCT_location = new_ACCT_location & b & c & " "
      Else
        new_ACCT_location = new_ACCT_location & a & " "
      End if
    End if
  Next
  EMReadScreen ACCT_ver, 1, 10, 63
  If ACCT_ver = "N" then 
    ACCT_ver = ", no proof provided"
  Else
    ACCT_ver = ""
  End if
  x = x & ACCT_type & " at " & new_ACCT_location & "(" & ACCT_amt & ")" & ACCT_ver & ".; "
  new_ACCT_location = ""
End function

Function add_BUSI_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen BUSI_type, 2, 5, 37
  If BUSI_type = "01" then BUSI_type = "Farming"
  If BUSI_type = "02" then BUSI_type = "Real Estate"
  If BUSI_type = "03" then BUSI_type = "Home Product Sales"
  If BUSI_type = "04" then BUSI_type = "Other Sales"
  If BUSI_type = "05" then BUSI_type = "Personal Services"
  If BUSI_type = "06" then BUSI_type = "Paper Route"
  If BUSI_type = "07" then BUSI_type = "InHome Daycare"
  If BUSI_type = "08" then BUSI_type = "Rental Income"
  If BUSI_type = "09" then BUSI_type = "Other"
  EMWriteScreen "x", 7, 26
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  If cash_check = 1 then
    EMReadScreen BUSI_ver, 1, 9, 73
  ElseIf HC_check = 1 then 
    EMReadScreen BUSI_ver, 1, 12, 73
    If BUSI_ver = "_" then EMReadScreen BUSI_ver, 1, 13, 73
  ElseIf SNAP_check = 1 then
    EMReadScreen BUSI_ver, 1, 11, 73
  End if
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
  If SNAP_check = 1 then
    EMReadScreen BUSI_amt, 8, 11, 68
    BUSI_amt = trim(BUSI_amt)
  ElseIf cash_check = 1 then 
    EMReadScreen BUSI_amt, 8, 9, 54
    BUSI_amt = trim(BUSI_amt)
  ElseIf HC_check = 1 then 
    EMWriteScreen "x", 17, 29
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen BUSI_amt, 8, 15, 54
    If BUSI_amt = "    0.00" then EMReadScreen BUSI_amt, 8, 16, 54
    BUSI_amt = trim(BUSI_amt)
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
  End if
  x = x & trim(BUSI_type) & " BUSI"
  EMReadScreen BUSI_income_end_date, 8, 5, 71
  If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
  If IsDate(BUSI_income_end_date) = True then
    x = x & " (ended " & BUSI_income_end_date & ")"
  Else
    If BUSI_amt <> "" then x = x & ", ($" & BUSI_amt & "/monthly)"
  End if
  If BUSI_ver = "N" or BUSI_ver = "?" then 
    x = x & ", no proof provided.; "
  Else
    x = x & ".; "
  End if
End function

Function add_CARS_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen CARS_year, 4, 8, 31
  EMReadScreen CARS_make, 15, 8, 43
  CARS_make = replace(CARS_make, "_", "")
  EMReadScreen CARS_model, 15, 8, 66
  CARS_model = replace(CARS_model, "_", "")
  CARS_type = CARS_year & " " & CARS_make & " " & CARS_model
  CARS_type = split(CARS_type)
  For each a in CARS_type
    If len(a) > 1 then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      new_CARS_type = new_CARS_type & b & c & " "
    End if
  Next
  EMReadScreen CARS_amt, 8, 9, 45
  CARS_amt = trim(CARS_amt)
  EMReadScreen counted_check, 1, 15, 43
  If counted_check = "8" then
    CARS_amt = "$" & CARS_amt
  Else
    CARS_amt = "excluded"
  End if
  x = x & trim(new_CARS_type) & ", (" & CARS_amt & "); "
  new_CARS_type = ""
End function

Function add_JOBS_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen JOBS_type, 30, 7, 42
  JOBS_type = replace(JOBS_type, "_", ""	)
  JOBS_type = trim(JOBS_type)
  JOBS_type = split(JOBS_type)
  For each a in JOBS_type
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      new_JOBS_type = new_JOBS_type & b & c & " "
    End if
  Next
  If SNAP_check = 1 then
    EMWriteScreen "x", 19, 38
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    SNAP_JOBS_amt = trim(SNAP_JOBS_amt)
    EMReadScreen pay_frequency, 1, 5, 64
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  ElseIf cash_check = 1 then
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    retro_JOBS_amt = trim(retro_JOBS_amt)
  ElseIf HC_check = 1 then 
    EMReadScreen pay_frequency, 1, 18, 35
    EMWriteScreen "x", 19, 54
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End If
  EMReadScreen JOBS_ver, 1, 6, 38
  EMReadScreen JOBS_income_end_date, 8, 9, 49
  If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
  If IsDate(JOBS_income_end_date) = True then
    x = x & new_JOBS_type & "(ended " & JOBS_income_end_date
  Else
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
    x = x & "EI from " & trim(new_JOBS_type)
    If SNAP_check = 1 then
      x = x & ", ($" & SNAP_JOBS_amt & "/" & pay_frequency
    ElseIf cash_check = 1 then
      x = x & ", ($" & retro_JOBS_amt & " budgeted"
    ElseIf HC_check = 1 then 
      x = x & ", ($" & HC_JOBS_amt & "/" & pay_frequency 
    End if
  End if
  If JOBS_ver = "N" or JOBS_ver = "?" then
    x = x & ", no proof provided).; "
  Else
    x = x & ").; "
  End if
End function

Function add_OTHR_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen OTHR_type, 16, 6, 43
  OTHR_type = trim(OTHR_type)
  EMReadScreen OTHR_amt, 10, 8, 40
  OTHR_amt = trim(OTHR_amt)
  EMReadScreen counted_check, 1, 12, 64
  If counted_check = "N" then
    OTHR_amt = "excluded"
  Else
    OTHR_amt = "$" & OTHR_amt
  End if
  x = x & trim(OTHR_type) & ", (" & OTHR_amt & ").; "
  new_OTHR_type = ""
End function

Function add_RBIC_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen RBIC_type, 16, 5, 48
  RBIC_type = trim(RBIC_type)
  EMReadScreen RBIC_amt, 8, 10, 62
  RBIC_amt = trim(RBIC_amt)
  EMReadScreen RBIC_ver, 1, 10, 76
  If RBIC_ver = "N" then RBIC_ver = ", no proof provided"
  EMReadScreen RBIC_end_date, 8, 6, 68
  RBIC_end_date = replace(RBIC_end_date, " ", "/")
  If isdate(RBIC_end_date) = True then
    x = x & trim(RBIC_type) & " RBIC, ended " & RBIC_end_date & RBIC_ver & ".; "
  Else
    x = x & trim(RBIC_type) & " RBIC, ($" & RBIC_amt & RBIC_ver & ").; "
  End if
End function

Function add_REST_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen REST_type, 16, 6, 41
  REST_type = trim(REST_type)
  EMReadScreen REST_amt, 10, 8, 41
  REST_amt = trim(REST_amt)
  EMReadScreen counted_check, 1, 12, 54
  If counted_check = "3" or counted_check = "4" or counted_check = "7" then
    REST_amt = "excluded"
  Else
    REST_amt = "$" & REST_amt
  End if
  x = x & trim(REST_type) & ", (" & REST_amt & ").; "
  new_REST_type = ""
End function


Function add_SECU_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen SECU_amt, 8, 10, 52
  SECU_amt = trim(SECU_amt)
  EMReadScreen counted_check, 1, 15, 64
  If counted_check = "N" then
    SECU_amt = "excluded"
  Else
    SECU_amt = "$" & SECU_amt
  End if
  EMReadScreen SECU_type, 2, 6, 50
  EMReadScreen SECU_location, 20, 8, 50
  SECU_location = replace(SECU_location, "_", "")
  SECU_location = split(SECU_location)
  For each a in SECU_location
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      If len(a) > 3 then
        new_SECU_location = new_SECU_location & b & c & " "
      Else
        new_SECU_location = new_SECU_location & a & " "
      End if
    End if
  Next
  EMReadScreen SECU_ver, 1, 11, 50
  If SECU_ver = "1" then SECU_ver = "agency form provided"
  If SECU_ver = "2" then SECU_ver = "source doc provided"
  If SECU_ver = "3" then SECU_ver = "verified via phone"
  If SECU_ver = "5" then SECU_ver = "other doc verified"
  If SECU_ver = "N" then SECU_ver = "no proof provided"
  x = x & SECU_type & " at " & new_SECU_location & " (" & SECU_amt & "), " & SECU_ver & ".; "
  new_SECU_location = ""
End function

Function add_UNEA_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen UNEA_type, 16, 5, 40
  If UNEA_type = "Unemployment Ins" then UNEA_type = "UC"
  If UNEA_type = "Disbursed Child " then UNEA_type = "CS"
  If UNEA_type = "Disbursed CS Arr" then UNEA_type = "CS arrears"
  UNEA_type = trim(UNEA_type)
  EMReadScreen UNEA_ver, 1, 5, 65
  EMReadScreen UNEA_income_end_date, 8, 7, 68
  If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
  If IsDate(UNEA_income_end_date) = True then
    x = x & UNEA_type & " (ended " & UNEA_income_end_date
  Else
    EMReadScreen UNEA_amt, 8, 18, 68
    UNEA_amt = trim(UNEA_amt)
    If SNAP_check = 1 then
      EMWriteScreen "x", 10, 26
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen SNAP_UNEA_amt, 8, 17, 56
      SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
      EMReadScreen pay_frequency, 1, 5, 64
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    ElseIf cash_check = 1 then
      EMReadScreen retro_UNEA_amt, 8, 18, 39
      retro_UNEA_amt = trim(retro_UNEA_amt)
    ElseIf HC_check = 1 then 
      EMWriteScreen "x", 6, 56
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen HC_UNEA_amt, 8, 9, 65
      HC_UNEA_amt = trim(HC_UNEA_amt)
      EMReadScreen pay_frequency, 1, 10, 63
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      If HC_UNEA_amt = "________" then
        EMReadScreen HC_UNEA_amt, 8, 18, 68
        HC_UNEA_amt = trim(HC_UNEA_amt)
        pay_frequency = "mo budgeted prospectively"
      End if
    End If
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" then pay_frequency = "non-monthly"
    x = x & trim(UNEA_type)
    If SNAP_check = 1 then
      x = x & ", ($" & SNAP_UNEA_amt & "/" & pay_frequency
    ElseIf cash_check = 1 then
      x = x & ", ($" & retro_UNEA_amt & " budgeted"
    ElseIf HC_check = 1 then 
      x = x & ", ($" & HC_UNEA_amt & "/" & pay_frequency
    End if
  End if
  If UNEA_ver = "N" or UNEA_ver = "?" then
    x = x & ", no proof provided).; "
  Else
    x = x & ").; "
  End if
End function

Function attn
  EMSendKey "<attn>"
  EMWaitReady -1, 0
End function

function back_to_SELF
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

Function create_MAXIS_friendly_date(date_variable, variable_length, screen_row, screen_col) 
  var_month = datepart("m", dateadd("d", variable_length, date_variable))
  If len(var_month) = 1 then var_month = "0" & var_month
  EMWriteScreen var_month, screen_row, screen_col
  var_day = datepart("d", dateadd("d", variable_length, date_variable))
  If len(var_day) = 1 then var_day = "0" & var_day
  EMWriteScreen var_day, screen_row, screen_col + 3
  var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
  EMWriteScreen right(var_year, 2), screen_row, screen_col + 6
End function

Function end_excel_and_script
  objExcel.Workbooks.Close
  objExcel.quit
  stopscript
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

Function memb_navigation_next
  HH_memb_row = HH_memb_row + 1
  EMReadScreen next_HH_memb, 2, HH_memb_row, 3
  If isnumeric(next_HH_memb) = False then
    HH_memb_row = HH_memb_row + 1
  Else
    EMWriteScreen next_HH_memb, 20, 76
    EMWriteScreen "01", 20, 79
  End if
End function

Function memb_navigation_prev
  HH_memb_row = HH_memb_row - 1
  EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
  If isnumeric(prev_HH_memb) = False then
    HH_memb_row = HH_memb_row + 1
  Else
    EMWriteScreen prev_HH_memb, 20, 76
    EMWriteScreen "01", 20, 79
  End if
End function

function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(x) and STAT_note_check <> "NOTE" then 
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen y, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen x, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen y, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    End if
  End if
End function

function navigation_buttons 'this works by calling the navigation_buttons function when the buttonpressed isn't -1
  If ButtonPressed = ABPS_button then call navigate_to_screen("stat", "ABPS")
  If ButtonPressed = ACCI_button then call navigate_to_screen("stat", "ACCI")
  If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
  If ButtonPressed = ADDR_button then call navigate_to_screen("stat", "ADDR")
  If ButtonPressed = ALTP_button then call navigate_to_screen("stat", "ALTP")
  If ButtonPressed = AREP_button then call navigate_to_screen("stat", "AREP")
  If ButtonPressed = BILS_button then call navigate_to_screen("stat", "BILS")
  If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
  If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
  If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
  If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
  If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
  If ButtonPressed = DIET_button then call navigate_to_screen("stat", "DIET")
  If ButtonPressed = DISA_button then call navigate_to_screen("stat", "DISA")
  If ButtonPressed = EATS_button then call navigate_to_screen("stat", "EATS")
  If ButtonPressed = ELIG_DWP_button then call navigate_to_screen("elig", "DWP_")
  If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
  If ButtonPressed = ELIG_GA_button then call navigate_to_screen("elig", "GA__")
  If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
  If ButtonPressed = ELIG_MFIP_button then call navigate_to_screen("elig", "MFIP")
  If ButtonPressed = ELIG_MSA_button then call navigate_to_screen("elig", "MSA_")
  If ButtonPressed = ELIG_WB_button then call navigate_to_screen("elig", "WB__")
  If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
  If ButtonPressed = FMED_button then call navigate_to_screen("stat", "FMED")
  If ButtonPressed = HCRE_button then call navigate_to_screen("stat", "HCRE")
  If ButtonPressed = HEST_button then call navigate_to_screen("stat", "HEST")
  If ButtonPressed = IMIG_button then call navigate_to_screen("stat", "IMIG")
  If ButtonPressed = INSA_button then call navigate_to_screen("stat", "INSA")
  If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
  If ButtonPressed = MEDI_button then call navigate_to_screen("stat", "MEDI")
  If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
  If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
  If ButtonPressed = MONT_button then call navigate_to_screen("stat", "MONT")
  If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
  If ButtonPressed = PBEN_button then call navigate_to_screen("stat", "PBEN")
  If ButtonPressed = PDED_button then call navigate_to_screen("stat", "PDED")
  If ButtonPressed = PREG_button then call navigate_to_screen("stat", "PREG")
  If ButtonPressed = PROG_button then call navigate_to_screen("stat", "PROG")
  If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
  If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
  If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
  If ButtonPressed = SCHL_button then call navigate_to_screen("stat", "SCHL")
  If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
  If ButtonPressed = STIN_button then call navigate_to_screen("stat", "STIN")
  If ButtonPressed = STEC_button then call navigate_to_screen("stat", "STEC")
  If ButtonPressed = STWK_button then call navigate_to_screen("stat", "STWK")
  If ButtonPressed = SHEL_button then call navigate_to_screen("stat", "SHEL")
  If ButtonPressed = SWKR_button then call navigate_to_screen("stat", "SWKR")
  If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
  If ButtonPressed = TYPE_button then call navigate_to_screen("stat", "TYPE")
  If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
End function

function new_BS_BSI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then 
    EMSendKey "--------BURIAL SPACE/ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_CAI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then 
    EMSendKey "--------CASH ADVANCE ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_page_check
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 17 then
    EMSendKey ">>>>MORE>>>>"
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    MAXIS_row = 4
  End if
end function

function new_service_heading
  EMGetCursor MAXIS_service_row, MAXIS_service_col
  If MAXIS_service_row = 4 then 
    EMSendKey "--------------SERVICE--------------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_service_row = 5
  end if
End function

Function panel_navigation_next
  EMReadScreen current_panel, 1, 2, 73
  EMReadScreen amount_of_panels, 1, 2, 78
  If current_panel < amount_of_panels then new_panel = current_panel + 1
  If current_panel = amount_of_panels then new_panel = current_panel
  If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
End function

Function panel_navigation_prev
  EMReadScreen current_panel, 1, 2, 73
  EMReadScreen amount_of_panels, 1, 2, 78
  If current_panel = 1 then new_panel = current_panel
  If current_panel > 1 then new_panel = current_panel - 1
  If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
End function

Function PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
End function

Function PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

Function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

Function PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
End function

Function PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
End function

Function PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
End function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

Function PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
End function

Function PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
End function

function PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function run_another_script(script_path)
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  Execute text_from_the_other_script
end function

function stat_navigation
  EMReadScreen STAT_check, 4, 20, 21
  If STAT_check = "STAT" then
    If ButtonPressed = prev_panel_button then 
      EMReadScreen current_panel, 1, 2, 73
      EMReadScreen amount_of_panels, 1, 2, 78
      If current_panel = 1 then new_panel = current_panel
      If current_panel > 1 then new_panel = current_panel - 1
      If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
    End if
    If ButtonPressed = next_panel_button then 
      EMReadScreen current_panel, 1, 2, 73
      EMReadScreen amount_of_panels, 1, 2, 78
      If current_panel < amount_of_panels then new_panel = current_panel + 1
      If current_panel = amount_of_panels then new_panel = current_panel
      If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
    End if
    If ButtonPressed = prev_memb_button then 
      HH_memb_row = HH_memb_row - 1
      EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
      If isnumeric(prev_HH_memb) = False then
        HH_memb_row = HH_memb_row + 1
      Else
        EMWriteScreen prev_HH_memb, 20, 76
        EMWriteScreen "01", 20, 79
      End if
    End if
    If ButtonPressed = next_memb_button then 
      HH_memb_row = HH_memb_row + 1
      EMReadScreen next_HH_memb, 2, HH_memb_row, 3
      If isnumeric(next_HH_memb) = False then
        HH_memb_row = HH_memb_row + 1
      Else
        EMWriteScreen next_HH_memb, 20, 76
        EMWriteScreen "01", 20, 79
      End if
    End if
  End if
End function

function script_end_procedure(closing_message)
  If closing_message <> "" then MsgBox closing_message
  stop_time = timer
  script_run_time = stop_time - start_time
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set ts = fs.OpenTextFile("q:\Blue Zone Scripts\Script Files\STATISTICS - log usage stats.vbs")
  script_to_run = ts.ReadAll
  ts.Close
  Execute script_to_run
  stopscript
end function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

Function write_editbox_in_case_note(x, y, z) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
  variable_array = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in variable_array 
    EMGetCursor row, col 
    If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 17 and col + (len(x)) >= 80) then EMSendKey "<newline>" & space(z)
    If (row = 4 and col = 3) then EMSendKey space(z)
    EMSendKey x & " "
    If right(x, 1) = ";" then 
      EMSendKey "<backspace>" & "<backspace>" 
      EMGetCursor row, col 
      If row = 17 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSendKey space(z)
      Else
        EMSendKey "<newline>" & space(z)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_new_line_in_case_note(x)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_three_columns_in_case_note(col_01_start_point, col_01_variable, col_02_start_point, col_02_variable, col_03_start_point, col_03_variable)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMGetCursor row, col
  EMWriteScreen "                                                                              ", row, 3
  EMSetCursor row, col_01_start_point
  EMSendKey col_01_variable
  EMSetCursor row, col_02_start_point
  EMSendKey col_02_variable
  EMSetCursor row, col_03_start_point
  EMSendKey col_03_variable
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function