Function autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)
  If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
    call navigate_to_screen("stat", "ABPS")
    EMReadScreen ABPS_total_pages, 1, 2, 78
    If ABPS_total_pages <> 0 then 
      Do
        'First it checks the support coop. If it's "N" it'll add a blurb about it to the support_coop variable
        EMReadScreen support_coop_code, 1, 4, 73
        If support_coop_code = "N" then
          EMReadScreen caregiver_ref_nbr, 2, 4, 47
          If instr(support_coop, "Memb " & caregiver_ref_nbr & " not cooperating with child support; ") = 0 then support_coop = support_coop & "Memb " & caregiver_ref_nbr & " not cooperating with child support; "'the if...then statement makes sure the info isn't duplicated. 
        End if
        'Then it gets info on the ABPS themself.
        EMReadScreen ABPS_current, 45, 10, 30
        If ABPS_current = "________________________  First: ____________" then ABPS_current = "Parent unknown"
        ABPS_current = replace(ABPS_current, "  First:", ",")
        ABPS_current = replace(ABPS_current, "_", "")
        ABPS_current = split(ABPS_current)
        For each a in ABPS_current
          b = ucase(left(a, 1))
          c = LCase(right(a, len(a) -1))
          If len(a) > 1 then
            new_ABPS_current = new_ABPS_current & b & c & " "
          Else
            new_ABPS_current = new_ABPS_current & a & " "
          End if
        Next
        ABPS_row = 15 'Setting variable for do...loop
        Do 'Using a do...loop to determine which MEMB numbers are with this parent
          EMReadScreen child_ref_nbr, 2, ABPS_row, 35
          If child_ref_nbr <> "__" then
            amt_of_children_for_ABPS = amt_of_children_for_ABPS + 1
            children_for_ABPS = children_for_ABPS & child_ref_nbr & ", "
          End if
          ABPS_row = ABPS_row + 1
        Loop until ABPS_row > 17
        'Cleaning up the "children_for_ABPS" variable to be more readable
        children_for_ABPS = left(children_for_ABPS, len(children_for_ABPS) - 2) 'cleaning up the end of the variable (removing the comma for single kids)
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it around to change the last comma to an "and"
        children_for_ABPS = replace(children_for_ABPS, ",", "dna ", 1, 1)        'it's backwards, replaces just one comma with an "and"
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it back around 
        if amt_of_children_for_ABPS > 1 then HH_memb_title = " for membs "
        if amt_of_children_for_ABPS <= 1 then HH_memb_title = " for memb "
        variable_written_to = variable_written_to & trim(new_ABPS_current) & HH_memb_title & children_for_ABPS & "; "
        'Resetting variables for the do...loop in case this function runs again
        new_ABPS_current = "" 
        amt_of_children_for_ABPS = 0
        children_for_ABPS = ""
        'Checking to see if it needs to run again, if it does it transmits or else the loop stops
        EMReadScreen ABPS_current_page, 1, 2, 73
        If ABPS_current_page <> ABPS_total_pages then transmit
      Loop until ABPS_current_page = ABPS_total_pages
      'Combining the two variables (support coop and the variable written to)
      variable_written_to = support_coop & variable_written_to
    End if
  Elseif panel_read_from = "ACCI" then '----------------------------------------------------------------------------------------------------ACCI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "ACCI")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCI_total, 1, 2, 78
      If ACCI_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCI_to_variable(variable_written_to)
          EMReadScreen ACCI_panel_current, 1, 2, 73
          If cint(ACCI_panel_current) < cint(ACCI_total) then transmit
        Loop until cint(ACCI_panel_current) = cint(ACCI_total)
      End if
    Next
  Elseif panel_read_from = "ACCT" then '----------------------------------------------------------------------------------------------------ACCT
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "acct")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCT_total, 1, 2, 78
      If ACCT_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCT_to_variable(variable_written_to)
          EMReadScreen ACCT_panel_current, 1, 2, 73
          If cint(ACCT_panel_current) < cint(ACCT_total) then transmit
        Loop until cint(ACCT_panel_current) = cint(ACCT_total)
      End if
    Next
  Elseif panel_read_from = "ADDR" then '----------------------------------------------------------------------------------------------------ADDR
    call navigate_to_screen("stat", "addr")
    EMReadScreen addr_line_01, 22, 6, 43
    EMReadScreen addr_line_02, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 12, 9, 43
    variable_written_to = replace(addr_line_01, "_", "") & "; " & replace(addr_line_02, "_", "") & "; " & replace(city_line, "_", "") & ", " & state_line & " " & replace(zip_line, "__ ", "-")
    variable_written_to = replace(variable_written_to, "; ; ", "; ") 'in case there's only one line on ADDR
  Elseif panel_read_from = "AREP" then '----------------------------------------------------------------------------------------------------AREP
    call navigate_to_screen("stat", "arep")
    EMReadScreen AREP_name, 37, 4, 32
    AREP_name = replace(AREP_name, "_", "")
    AREP_name = split(AREP_name)
    For each word in AREP_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "BILS" then '----------------------------------------------------------------------------------------------------BILS
    call navigate_to_screen("stat", "bils")
    EMReadScreen BILS_amt, 1, 2, 78
    If BILS_amt <> 0 then variable_written_to = "BILS known to MAXIS."
  Elseif panel_read_from = "BUSI" then '----------------------------------------------------------------------------------------------------BUSI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "busi")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen BUSI_total, 1, 2, 78
      If BUSI_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_BUSI_to_variable(variable_written_to)
          EMReadScreen BUSI_panel_current, 1, 2, 73
          If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
        Loop until cint(BUSI_panel_current) = cint(BUSI_total)
      End if
    Next
  Elseif panel_read_from = "CARS" then '----------------------------------------------------------------------------------------------------CARS
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "cars")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen CARS_total, 1, 2, 78
      If CARS_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_CARS_to_variable(variable_written_to)
          EMReadScreen CARS_panel_current, 1, 2, 73
          If cint(CARS_panel_current) < cint(CARS_total) then transmit
        Loop until cint(CARS_panel_current) = cint(CARS_total)
      End if
    Next
  Elseif panel_read_from = "CASH" then '----------------------------------------------------------------------------------------------------CASH
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "cash")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen cash_amt, 8, 8, 39
      cash_amt = trim(cash_amt)
      If cash_amt <> "________" then
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Cash ($" & cash_amt & "); "
      End if
    Next
  Elseif panel_read_from = "COEX" then '----------------------------------------------------------------------------------------------------COEX
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "coex")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen support_amt, 8, 10, 63
      support_amt = trim(support_amt)
      If support_amt <> "________" then
        EMReadScreen support_ver, 1, 10, 36
        If support_ver = "?" or support_ver = "N" then
          support_ver = ", no proof provided"
        Else
          support_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Support ($" & support_amt & "/mo" & support_ver & "); "
      End if
      EMReadScreen alimony_amt, 8, 11, 63
      alimony_amt = trim(alimony_amt)
      If alimony_amt <> "________" then
        EMReadScreen alimony_ver, 1, 11, 36
        If alimony_ver = "?" or alimony_ver = "N" then
          alimony_ver = ", no proof provided"
        Else
          alimony_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Alimony ($" & alimony_amt & "/mo" & alimony_ver & "); "
      End if
      EMReadScreen tax_dep_amt, 8, 12, 63
      tax_dep_amt = trim(tax_dep_amt)
      If tax_dep_amt <> "________" then
        EMReadScreen tax_dep_ver, 1, 12, 36
        If tax_dep_ver = "?" or tax_dep_ver = "N" then
          tax_dep_ver = ", no proof provided"
        Else
          tax_dep_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Tax dep ($" & tax_dep_amt & "/mo" & tax_dep_ver & "); "
      End if
      EMReadScreen other_COEX_amt, 8, 13, 63
      other_COEX_amt = trim(other_COEX_amt)
      If other_COEX_amt <> "________" then
        EMReadScreen other_COEX_ver, 1, 13, 36
        If other_COEX_ver = "?" or other_COEX_ver = "N" then
          other_COEX_ver = ", no proof provided"
        Else
          other_COEX_ver = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Other ($" & other_COEX_amt & "/mo" & other_COEX_ver & "); "
      End if
    Next
  Elseif panel_read_from = "DCEX" then '----------------------------------------------------------------------------------------------------DCEX
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "dcex")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DCEX_row = 11
      Do
      EMReadScreen expense_amt, 8, DCEX_row, 63
      expense_amt = trim(expense_amt)
      If expense_amt <> "________" then
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen child_ref_nbr, 2, DCEX_row, 29
        EMReadScreen expense_ver, 1, DCEX_row, 41
        If expense_ver = "?" or expense_ver = "N" or expense_ver = "_" then
          expense_ver = ", no proof provided"
        Else
          expense_ver = ""
        End if
        variable_written_to = variable_written_to & "Child " & child_ref_nbr & " ($" & expense_amt & "/mo DCEX" & expense_ver & "); "
      End if
      DCEX_row = DCEX_row + 1
      Loop until DCEX_row = 17
    Next
  Elseif panel_read_from = "DIET" then '----------------------------------------------------------------------------------------------------DIET
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "diet")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DIET_row = 8 'Setting this variable for the next do...loop
      EMReadScreen DIET_total, 1, 2, 78
      If DIET_total <> 0 then 
        If HH_member <> "01" then DIET = DIET & "Member " & HH_member & "- "
        Do
          EMReadScreen diet_type, 2, DIET_row, 40
          EMReadScreen diet_proof, 1, DIET_row, 51
          If diet_proof = "_" or diet_proof = "?" or diet_proof = "N" then 
            diet_proof = ", no proof provided"
          Else
            diet_proof = ""
          End if
          If diet_type = "01" then diet_type = "High Protein"
          If diet_type = "02" then diet_type = "Cntrl Protein (40-60 g/day)"
          If diet_type = "03" then diet_type = "Cntrl Protein (<40 g/day)"
          If diet_type = "04" then diet_type = "Lo Cholesterol"
          If diet_type = "05" then diet_type = "High Residue"
          If diet_type = "06" then diet_type = "Preg/Lactation"
          If diet_type = "07" then diet_type = "Gluten Free"
          If diet_type = "08" then diet_type = "Lactose Free"
          If diet_type = "09" then diet_type = "Anti-Dumping"
          If diet_type = "10" then diet_type = "Hypoglycemic"
          If diet_type = "11" then diet_type = "Ketogenic"
          If diet_type <> "__" and diet_type <> "  " then variable_written_to = variable_written_to & diet_type & diet_proof & "; "
          DIET_row = DIET_row + 1
        Loop until DIET_row = 19
      End if
    Next
  Elseif panel_read_from = "DISA" then '----------------------------------------------------------------------------------------------------DISA
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "disa")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen DISA_status, 2, 13, 59
      If DISA_status = "01" or DISA_status = "02" or DISA_status = "03" or DISA_status = "04" then DISA_status = "RSDI/SSI certified"
      If DISA_status = "06" then DISA_status = "SMRT/SSA pends"
      If DISA_status = "08" then DISA_status = "Certified blind"
      If DISA_status = "10" then DISA_status = "Certified disabled"
      If DISA_status = "11" then DISA_status = "Spec cat- disa child"
      If DISA_status = "20" then DISA_status = "TEFRA- disabled"
      If DISA_status = "21" then DISA_status = "TEFRA- blind"
      If DISA_status = "22" then DISA_status = "MA-EPD"
      If DISA_status = "23" then DISA_status = "MA/waiver"
      If DISA_status = "24" then DISA_status = "SSA/SMRT appeal pends"
      If DISA_status = "26" then DISA_status = "SSA/SMRT disa deny"
      If DISA_status = "__" then
        DISA_status = ""
      Else
        EMReadScreen DISA_ver, 1, 13, 69
        If DISA_ver = "?" or DISA_ver = "N" then
          DISA_proof_type = ", no proof provided"
        Else
          DISA_proof_type = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & DISA_status & DISA_proof_type & "; "
      End if
    Next
  Elseif panel_read_from = "EATS" then '----------------------------------------------------------------------------------------------------EATS
    call navigate_to_screen("stat", "eats")
    row = 14
    Do
      EMReadScreen reference_numbers_current_row, 40, row, 39
      reference_numbers = reference_numbers + reference_numbers_current_row  
      row = row + 1
    Loop until row = 18
    reference_numbers = replace(reference_numbers, "  ", " ")
    reference_numbers = split(reference_numbers)
    For each member in reference_numbers
      If member <> "__" and member <> "" then EATS_info = EATS_info & member & ", "
    Next
    EATS_info = trim(EATS_info)
    if right(EATS_info, 1) = "," then EATS_info = left(EATS_info, len(EATS_info) - 1)
    If EATS_info <> "" then variable_written_to = variable_written_to & ", p/p sep from memb(s) " & EATS_info & "."
  Elseif panel_read_from = "FACI" then '----------------------------------------------------------------------------------------------------FACI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "faci")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen FACI_total, 1, 2, 78
      If FACI_total <> 0 then
        row = 14
        Do
          EMReadScreen date_in_check, 4, row, 53
          EMReadScreen date_out_check, 4, row, 77
          If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
          If row > 18 then
            EMReadScreen FACI_page, 1, 2, 73
            If FACI_page = FACI_total then 
              FACI_status = "Not in facility"
            Else
              transmit
              row = 14
            End if
          End if
        Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
        EMReadScreen client_FACI, 30, 6, 43
        client_FACI = replace(client_FACI, "_", "")
        FACI_array = split(client_FACI)
        For each a in FACI_array
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            new_FACI = new_FACI & b & c & " "
          End if
        Next
        client_FACI = new_FACI
        If FACI_status = "Not in facility" then
          client_FACI = ""
        Else
          If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
          variable_written_to = variable_written_to & client_FACI & "; "
        End if
      End if
    Next
  Elseif panel_read_from = "FMED" then '----------------------------------------------------------------------------------------------------FMED
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "fmed")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      fmed_row = 9 'Setting this variable for the next do...loop
      EMReadScreen fmed_total, 1, 2, 78
      If fmed_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          EMReadScreen fmed_type, 2, fmed_row, 25
          EMReadScreen fmed_proof, 2, fmed_row, 32
          EMReadScreen fmed_amt, 8, fmed_row, 70
          If fmed_proof = "__" or fmed_proof = "?_" or fmed_proof = "NO" then 
            fmed_proof = ", no proof provided"
          Else
            fmed_proof = ""
          End if
          If fmed_amt = "________" then
            fmed_amt = ""
          Else
            fmed_amt = " ($" & trim(fmed_amt) & ")"
          End if
          If fmed_type = "01" then fmed_type = "Nursing Home"
          If fmed_type = "02" then fmed_type = "Hosp/Clinic"
          If fmed_type = "03" then fmed_type = "Physicians"
          If fmed_type = "04" then fmed_type = "Prescriptions"
          If fmed_type = "05" then fmed_type = "Ins Premiums"
          If fmed_type = "06" then fmed_type = "Dental"
          If fmed_type = "07" then fmed_type = "Medical Trans/Flat Amt"
          If fmed_type = "08" then fmed_type = "Vision Care"
          If fmed_type = "09" then fmed_type = "Medicare Prem"
          If fmed_type = "10" then fmed_type = "No Spdwn Amt/Waiver Obl"
          If fmed_type = "11" then fmed_type = "Home Care"
          If fmed_type = "12" then fmed_type = "Medical Trans/Mileage Calc"
          If fmed_type = "15" then fmed_type = "Medi Part D premium"
          If fmed_type <> "__" then variable_written_to = variable_written_to & fmed_type & fmed_amt & fmed_proof & "; "
          fmed_row = fmed_row + 1
          If fmed_row = 15 then
            PF20
            fmed_row = 9
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
          End if
        Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"
      End if
    Next
  Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
    call navigate_to_screen("stat", "hcre")
    EMReadScreen variable_written_to, 8, 10, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If variable_written_to = "__/__/__" then EMReadScreen variable_written_to, 8, 11, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then variable_written_to = cdate(variable_written_to) & ""
    If isdate(variable_written_to) = False then variable_written_to = ""
  Elseif panel_read_from = "HCRE-retro" then '----------------------------------------------------------------------------------------------HCRE-retro
    call navigate_to_screen("stat", "hcre")
    EMReadScreen variable_written_to, 5, 10, 64
    If isdate(variable_written_to) = True then
      variable_written_to = replace(variable_written_to, " ", "/01/")
      If DatePart("m", variable_written_to) <> DatePart("m", CAF_datestamp) or DatePart("yyyy", variable_written_to) <> DatePart("yyyy", CAF_datestamp) then
        variable_written_to = variable_written_to
      Else
        variable_written_to = ""
      End if
    End if
  Elseif panel_read_from = "HEST" then '----------------------------------------------------------------------------------------------------HEST
    call navigate_to_screen("stat", "hest")
    EMReadScreen HEST_total, 1, 2, 78
    If HEST_total <> 0 then 
      EMReadScreen heat_air_check, 6, 13, 75
      If heat_air_check <> "      " then variable_written_to = variable_written_to & "Heat/AC.; "
      EMReadScreen electric_check, 6, 14, 75
      If electric_check <> "      " then variable_written_to = variable_written_to & "Electric.; "
      EMReadScreen phone_check, 6, 15, 75
      If phone_check <> "      " then variable_written_to = variable_written_to & "Phone.; "
    End if
  Elseif panel_read_from = "IMIG" then '----------------------------------------------------------------------------------------------------IMIG
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "IMIG")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen IMIG_total, 1, 2, 78
      If IMIG_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen IMIG_type, 30, 6, 48
        variable_written_to = variable_written_to & trim(IMIG_type) & "; "
      End if
    Next
  Elseif panel_read_from = "INSA" then '----------------------------------------------------------------------------------------------------INSA
    call navigate_to_screen("stat", "insa")
    EMReadScreen INSA_amt, 1, 2, 78
    If INSA_amt <> 0 then
      EMReadScreen INSA_name, 38, 10, 38
      INSA_name = replace(INSA_name, "_", "")
      INSA_name = split(INSA_name)
      For each word in INSA_name
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 4 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      Next
      variable_written_to = trim(variable_written_to) & "; "
    End if
  Elseif panel_read_from = "JOBS" then '----------------------------------------------------------------------------------------------------JOBS
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "jobs")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen JOBS_total, 1, 2, 78
      If JOBS_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_JOBS_to_variable(variable_written_to)
          EMReadScreen JOBS_panel_current, 1, 2, 73
          If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
        Loop until cint(JOBS_panel_current) = cint(JOBS_total)
      End if
    Next
  Elseif panel_read_from = "MEDI" then '----------------------------------------------------------------------------------------------------MEDI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "MEDI")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen MEDI_amt, 1, 2, 78
      If MEDI_amt <> "0" then variable_written_to = variable_written_to & "Medicare for member " & HH_member & ".; "
    Next
  Elseif panel_read_from = "MEMB" then '----------------------------------------------------------------------------------------------------MEMB
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "memb")
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen rel_to_applicant, 2, 10, 42
      EMReadScreen client_age, 3, 8, 76
      If client_age = "   " then client_age = 0
      If cint(client_age) >= 21 or rel_to_applicant = "02" then
        number_of_adults = number_of_adults + 1
      Else
        number_of_children = number_of_children + 1
      End if
    Next
    If number_of_adults > 0 then variable_written_to = number_of_adults & "a"
    If number_of_children > 0 then variable_written_to = variable_written_to & ", " & number_of_children & "c"
    If left(variable_written_to, 1) = "," then variable_written_to = right(variable_written_to, len(variable_written_to) - 1)
  Elseif panel_read_from = "MEMI" then '----------------------------------------------------------------------------------------------------MEMI
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "memi")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen citizen, 1, 10, 49
      If citizen = "Y" then citizen = "US citizen"
      If citizen = "N" then citizen = "non-citizen"
      EMReadScreen citizenship_ver, 2, 10, 78
      EMReadScreen SSA_MA_citizenship_ver, 1, 11, 49
      If citizenship_ver = "__" or citizenship_ver = "NO" then cit_proof_indicator = ", no verifs provided"
      If SSA_MA_citizenship_ver = "R" then cit_proof_indicator = ", MEMI infc req'd"
      If (citizenship_ver <> "__" and citizenship_ver <> "NO") or (SSA_MA_citizenship_ver = "A") then cit_proof_indicator = ""
      If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
      variable_written_to = variable_written_to & citizen & cit_proof_indicator & "; "
    Next
  ElseIf panel_read_from = "MONT" then '----------------------------------------------------------------------------------------------------MONT
    call navigate_to_screen("stat", "mont")
    EMReadScreen variable_written_to, 8, 6, 39
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "OTHR" then '----------------------------------------------------------------------------------------------------OTHR
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "othr")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen OTHR_total, 1, 2, 78
      If OTHR_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_OTHR_to_variable(variable_written_to)
          EMReadScreen OTHR_panel_current, 1, 2, 73
          If cint(OTHR_panel_current) < cint(OTHR_total) then transmit
        Loop until cint(OTHR_panel_current) = cint(OTHR_total)
      End if
    Next
  Elseif panel_read_from = "PBEN" then '----------------------------------------------------------------------------------------------------PBEN
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "pben")
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen panel_amt, 1, 2, 78
      If panel_amt <> "0" then
        If HH_member <> "01" then PBEN = PBEN & "Member " & HH_member & "- "
        row = 8
        Do
          EMReadScreen PBEN_type, 12, row, 28
          EMReadScreen PBEN_disp, 1, row, 77
          If PBEN_disp = "A" then PBEN_disp = " appealing"
          If PBEN_disp = "D" then PBEN_disp = " denied"
          If PBEN_disp = "E" then PBEN_disp = " eligible"
          If PBEN_disp = "P" then PBEN_disp = " pends"
          If PBEN_disp = "N" then PBEN_disp = " not applied yet"
          If PBEN_disp = "R" then PBEN_disp = " refused"
          If PBEN_type <> "            " then PBEN = PBEN & trim(PBEN_type) & PBEN_disp & "; "
          row = row + 1
        Loop until row = 14
      End if
    Next
    If PBEN <> "" then variable_written_to = variable_written_to & PBEN
  Elseif panel_read_from = "PREG" then '----------------------------------------------------------------------------------------------------PREG
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "PREG")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen PREG_total, 1, 2, 78
      If PREG_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen PREG_due_date, 8, 10, 53
        If PREG_due_date = "__ __ __" then
          PREG_due_date = "unknown"
        Else
          PREG_due_date = replace(PREG_due_date, " ", "/")
        End if
        variable_written_to = variable_written_to & "Due date is " & PREG_due_date & ".; "
      End if
    Next
  Elseif panel_read_from = "PROG" then '----------------------------------------------------------------------------------------------------PROG
    call navigate_to_screen("stat", "prog") 'THIS WILL DETERMINE THE LAST DATESTAMP ON THE PROG PANEL
    row = 6
    Do
      EMReadScreen appl_prog_date, 8, row, 33
      If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "
      row = row + 1
    Loop until row = 13
    appl_prog_date_array = split(appl_prog_date_array)
    variable_written_to = CDate(appl_prog_date_array(0))
    for i = 0 to ubound(appl_prog_date_array) - 1
      if CDate(appl_prog_date_array(i)) > variable_written_to then 
        variable_written_to = CDate(appl_prog_date_array(i))
      End if
    next
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "RBIC" then '----------------------------------------------------------------------------------------------------RBIC
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "rbic")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen RBIC_total, 1, 2, 78
      If RBIC_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_RBIC_to_variable(variable_written_to)
          EMReadScreen RBIC_panel_current, 1, 2, 73
          If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
        Loop until cint(RBIC_panel_current) = cint(RBIC_total)
      End if
    Next
  Elseif panel_read_from = "REST" then '----------------------------------------------------------------------------------------------------REST
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "rest")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen REST_total, 1, 2, 78
      If REST_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_REST_to_variable(variable_written_to)
          EMReadScreen REST_panel_current, 1, 2, 73
          If cint(REST_panel_current) < cint(REST_total) then transmit
        Loop until cint(REST_panel_current) = cint(REST_total)
      End if
    Next
  Elseif panel_read_from = "REVW" then '----------------------------------------------------------------------------------------------------REVW
    call navigate_to_screen("stat", "revw")
    EMReadScreen variable_written_to, 8, 13, 37
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "SCHL" then '----------------------------------------------------------------------------------------------------SCHL
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "schl")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen school_type, 2, 7, 40
      If school_type = "01" then school_type = "elementary school"
      If school_type = "11" then school_type = "middle school"
      If school_type = "02" then school_type = "high school"
      If school_type = "03" then school_type = "GED"
      If school_type = "07" then school_type = "IEP"
      If school_type = "08" or school_type = "09" or school_type = "10" then school_type = "post-secondary"
      If school_type = "06" or school_type = "__" or school_type = "?_" then
        school_type = ""
      Else
        EMReadScreen SCHL_ver, 2, 6, 63
        If SCHL_ver = "?_" or SCHL_ver = "NO" then
          school_proof_type = ", no proof provided"
        Else
          school_proof_type = ""
        End if
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & school_type & school_proof_type & "; "
      End if
    Next
  Elseif panel_read_from = "SECU" then '----------------------------------------------------------------------------------------------------SECU
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "secu")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SECU_total, 1, 2, 78
      If SECU_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_SECU_to_variable(variable_written_to)
          EMReadScreen SECU_panel_current, 1, 2, 73
          If cint(SECU_panel_current) < cint(SECU_total) then transmit
        Loop until cint(SECU_panel_current) = cint(SECU_total)
      End if
    Next
  Elseif panel_read_from = "SHEL" then '----------------------------------------------------------------------------------------------------SHEL
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "shel")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SHEL_total, 1, 2, 78
      If SHEL_total <> 0 then 
        If HH_member <> "01" then member_number_designation = "Member " & HH_member & "- "
        row = 11
        Do
          EMReadScreen SHEL_amount, 8, row, 56
          If SHEL_amount <> "________" then
            EMReadScreen SHEL_type, 9, row, 24
            EMReadScreen SHEL_proof_check, 2, row, 67
            If SHEL_proof_check = "NO" or SHEL_proof_check = "?_" then 
              SHEL_proof = ", no proof provided"
            Else
              SHEL_proof = ""
            End if
            SHEL_expense = SHEL_expense & "$" & trim(SHEL_amount) & "/mo " & lcase(trim(SHEL_type)) & SHEL_proof & ".; "
          End if
          row = row + 1
        Loop until row = 19
        variable_written_to = variable_written_to & member_number_designation & SHEL_expense
      End if
      SHEL_expense = ""
    Next
  Elseif panel_read_from = "STWK" then '----------------------------------------------------------------------------------------------------STWK
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "STWK")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen STWK_total, 1, 2, 78
      If STWK_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen STWK_verification, 1, 7, 63
        If STWK_verification = "N" then
          STWK_verification = ", no proof provided"
        Else
          STWK_verification = ""
        End if
        EMReadScreen STWK_employer, 30, 6, 46
        STWK_employer = replace(STWK_employer, "_", "")
        STWK_employer = split(STWK_employer)
        For each a in STWK_employer
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            If len(a) > 3 then
              new_STWK_employer = new_STWK_employer & b & c & " "
            Else
              new_STWK_employer = new_STWK_employer & a & " "
            End if
          End if
        Next
        EMReadScreen STWK_income_stop_date, 8, 8, 46
        If STWK_income_stop_date = "__ __ __" then
          STWK_income_stop_date = "at unknown date"
        Else
          STWK_income_stop_date = replace(STWK_income_stop_date, " ", "/")
        End if
      EMReadScreen voluntary_quit, 1, 10, 46
	vol_quit_info = "Vol. Quit " & voluntary_quit
	  IF voluntary_quit = "Y" THEN
		EMReadScreen good_cause, 1, 12, 67
		EMReadScreen fs_pwe, 1, 14, 46
		vol_quit_info = "Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & " FS PWE " & fs_pwe
	  END IF
        variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & ".; " & vol_quit_info
      End if
      new_STWK_employer = "" 'clearing variable to prevent duplicates
    Next
  Elseif panel_read_from = "UNEA" then '----------------------------------------------------------------------------------------------------UNEA
    For each HH_member in HH_member_array
      call navigate_to_screen("stat", "unea")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen UNEA_total, 1, 2, 78
      If UNEA_total <> 0 then 
        If HH_member <> "01" then variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_UNEA_to_variable(variable_written_to)
          EMReadScreen UNEA_panel_current, 1, 2, 73
          If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
        Loop until cint(UNEA_panel_current) = cint(UNEA_total)
      End if
    Next
  End if
  variable_written_to = trim(variable_written_to) '-----------------------------------------------------------------------------------------cleaning up editbox
  if right(variable_written_to, 1) = ";" then variable_written_to= left(variable_written_to, len(variable_written_to) - 1)
  variable_written_to = replace(variable_written_to, "$________/non-monthly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/monthly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/weekly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/biweekly", "amt unknown")
  variable_written_to = replace(variable_written_to, "$________/semimonthly", "amt unknown")
End function
