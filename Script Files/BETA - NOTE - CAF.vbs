'Informational front-end message, date dependent.
If datediff("d", "05/15/2014", now) < 6 then MsgBox "This script has been updated as of 05/15/2014! Here's what's new:" & chr(13) & chr(13) & "Stop work edit box has been expanded and now autofills with stop work information, ABAWD edit box added and fills with ABAWD information, if you have any questions please email Robert or Charles."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BETA - NOTE - CAF"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'New function for stop work
Function autofill_editbox_from_MAXIS_test(HH_member_array, panel_read_from, variable_written_to)
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
	vol_quit_info = ", Vol. Quit " & voluntary_quit
	  IF voluntary_quit = "Y" THEN
		EMReadScreen good_cause, 1, 12, 67
		EMReadScreen fs_pwe, 1, 14, 46
		vol_quit_info = ", Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & ", FS PWE " & fs_pwe
	  END IF
        variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & vol_quit_info & ".; "
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
  Elseif panel_read_from = "WREG" then '---------------------------------------------------------------------------------------------------WREG
    For each HH_member in HH_member_array
	call navigate_to_screen("stat", "wreg")
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	EmWriteScreen "x", 13, 57
	transmit
	 bene_mo_col = (15 + (4*cint(footer_month)))
	  bene_yr_row = 10
 	 month_count = 0
 	   DO
  		  EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
  		    IF is_counted_month = "X" or is_counted_month = "M" THEN abawd_counted_months = abawd_counted_months + 1
   		  bene_mo_col = bene_mo_col - 4
    		    IF bene_mo_col = 15 THEN
        		bene_yr_row = bene_yr_row - 1
   	     		bene_mo_col = 63
   	   	    END IF
    		  month_count = month_count + 1
  	   LOOP until month_count = 36
  	PF3
	EmreadScreen read_abawd_status, 2, 13, 50
	If read_abawd_status = 10 or read_abawd_status = 11 or read_abawd_status = 13 then
	  abawd_status = "Client is ABAWD and has used " & abawd_counted_months & " months"
	else
	  abawd_status = "Client is not ABAWD"
	end if
	variable_written_to = variable_written_to & "Member " & HH_member & "- " & abawd_status & ".; "
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


'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 185, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_check
  CheckBox 50, 60, 30, 10, "HC", HC_check
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_check
  CheckBox 135, 60, 35, 10, "EMER", EMER_check
  DropListBox 70, 80, 75, 15, "Intake"+chr(9)+"Reapplication"+chr(9)+"Recertification"+chr(9)+"Add program", CAF_type
  CheckBox 5, 100, 160, 10, "Disable semicolons?", disable_semicolon_check
  ButtonGroup ButtonPressed
    OkButton 35, 165, 50, 15
    CancelButton 95, 165, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Programs applied for"
  Text 30, 85, 35, 10, "CAF type:"
  Text 15, 110, 160, 50, "(Disabling semicolons will cause your ''income'', ''asset'', and other sections to enter with word wrap, instead of each panel getting it's own line. This can be useful in households with many members, and could help keep case notes from exceeding four pages.)"
EndDialog

BeginDialog CAF_dialog_01, 0, 0, 451, 235, "CAF dialog part 1"
  EditBox 60, 5, 50, 15, CAF_datestamp
  EditBox 60, 25, 50, 15, interview_date
  EditBox 75, 45, 260, 15, HH_comp
  EditBox 35, 65, 200, 15, cit_id
  EditBox 265, 65, 180, 15, IMIG
  EditBox 60, 85, 120, 15, AREP
  EditBox 270, 85, 175, 15, SCHL
  EditBox 60, 105, 210, 15, DISA
  EditBox 310, 105, 135, 15, FACI
  EditBox 35, 135, 410, 15, PREG
  EditBox 35, 155, 410, 15, ABPS
  EditBox 55, 185, 390, 15, verifs_needed
  ButtonGroup ButtonPressed
    PushButton 340, 215, 50, 15, "NEXT", next_to_page_02_button
    CancelButton 395, 215, 50, 15
    PushButton 200, 15, 20, 10, "DWP", ELIG_DWP_button
    PushButton 220, 15, 15, 10, "FS", ELIG_FS_button
    PushButton 235, 15, 15, 10, "GA", ELIG_GA_button
    PushButton 250, 15, 15, 10, "HC", ELIG_HC_button
    PushButton 265, 15, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 285, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 305, 15, 15, 10, "WB", ELIG_WB_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 50, 60, 10, "HH comp/EATS:", EATS_button
    PushButton 240, 70, 20, 10, "IMIG:", IMIG_button
    PushButton 5, 90, 25, 10, "AREP/", AREP_button
    PushButton 30, 90, 25, 10, "ALTP:", ALTP_button
    PushButton 190, 90, 25, 10, "SCHL/", SCHL_button
    PushButton 215, 90, 25, 10, "STIN/", STIN_button
    PushButton 240, 90, 25, 10, "STEC:", STEC_button
    PushButton 5, 110, 25, 10, "DISA/", DISA_button
    PushButton 30, 110, 25, 10, "PDED:", PDED_button
    PushButton 280, 110, 25, 10, "FACI:", FACI_button
    PushButton 5, 140, 25, 10, "PREG:", PREG_button
    PushButton 5, 160, 25, 10, "ABPS:", ABPS_button
    PushButton 150, 215, 25, 10, "ADDR", ADDR_button
    PushButton 175, 215, 25, 10, "MEMB", MEMB_button
    PushButton 200, 215, 25, 10, "MEMI", MEMI_button
    PushButton 225, 215, 25, 10, "PROG", PROG_button
    PushButton 250, 215, 25, 10, "REVW", REVW_button
    PushButton 275, 215, 25, 10, "TYPE", TYPE_button
  GroupBox 195, 5, 130, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 10, 55, 10, "CAF datestamp:"
  Text 5, 30, 55, 10, "Interview date:"
  Text 5, 70, 25, 10, "CIT/ID:"
  Text 5, 190, 50, 10, "Verifs needed:"
  GroupBox 145, 205, 160, 25, "other STAT panels:"
EndDialog

BeginDialog CAF_dialog_02, 0, 0, 451, 315, "CAF dialog part 2"
  EditBox 60, 45, 385, 15, earned_income
  EditBox 70, 65, 375, 15, unearned_income
  EditBox 85, 85, 360, 15, income_changes
  EditBox 65, 105, 380, 15, notes_on_abawd
  EditBox 65, 125, 380, 15, notes_on_income
  EditBox 155, 145, 290, 15, is_any_work_temporary
  EditBox 60, 175, 385, 15, SHEL_HEST
  EditBox 60, 195, 250, 15, COEX_DCEX
  EditBox 65, 225, 380, 15, CASH_ACCTs
  EditBox 155, 245, 290, 15, other_assets
  EditBox 55, 275, 390, 15, verifs_needed
  ButtonGroup ButtonPressed
    PushButton 275, 300, 60, 10, "previous page", previous_to_page_01_button
    PushButton 340, 295, 50, 15, "NEXT", next_to_page_03_button
    CancelButton 395, 295, 50, 15
  GroupBox 5, 5, 130, 25, "ELIG panels:"
  ButtonGroup ButtonPressed
    PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
    PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
    PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
    PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
    PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
  GroupBox 145, 5, 135, 25, "Income panels"
  ButtonGroup ButtonPressed
    PushButton 150, 15, 25, 10, "BUSI", BUSI_button
    PushButton 175, 15, 25, 10, "JOBS", JOBS_button
    PushButton 200, 15, 25, 10, "PBEN", PBEN_button
    PushButton 225, 15, 25, 10, "RBIC", RBIC_button
    PushButton 250, 15, 25, 10, "UNEA", UNEA_button
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  ButtonGroup ButtonPressed
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
  Text 5, 50, 55, 10, "Earned income:"
  Text 5, 70, 65, 10, "Unearned income:"
  ButtonGroup ButtonPressed
    PushButton 5, 90, 75, 10, "STWK/inc. changes:", STWK_button
  Text 5, 110, 50, 10, "ABAWD notes:"
  Text 5, 130, 60, 10, "Notes on income:"
  Text 5, 150, 150, 10, "Is any work temporary? If so, explain details:"
  ButtonGroup ButtonPressed
    PushButton 5, 180, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 180, 25, 10, "HEST:", HEST_button
    PushButton 105, 250, 45, 10, "other assets:", OTHR_button
    PushButton 5, 200, 25, 10, "COEX/", COEX_button
    PushButton 5, 250, 25, 10, "CARS/", CARS_button
    PushButton 5, 230, 25, 10, "CASH/", CASH_button
    PushButton 30, 230, 30, 10, "ACCTs:", ACCT_button
    PushButton 55, 250, 25, 10, "SECU/", SECU_button
    PushButton 80, 250, 25, 10, "TRAN/", TRAN_button
  Text 5, 280, 50, 10, "Verifs needed:"
  ButtonGroup ButtonPressed
    PushButton 30, 200, 25, 10, "DCEX:", DCEX_button
    PushButton 30, 250, 25, 10, "REST/", REST_button
EndDialog


BeginDialog CAF_dialog_03, 0, 0, 451, 340, "CAF dialog part 3"
  EditBox 60, 45, 385, 15, INSA
  EditBox 35, 65, 410, 15, ACCI
  EditBox 35, 85, 175, 15, DIET
  EditBox 245, 85, 200, 15, BILS
  EditBox 35, 105, 290, 15, FMED
  EditBox 390, 105, 55, 15, HC_begin
  EditBox 180, 135, 265, 15, reason_expedited_wasnt_processed
  EditBox 100, 155, 345, 15, FIAT_reasons
  CheckBox 25, 180, 80, 10, "Application signed?", application_signed_check
  CheckBox 110, 180, 50, 10, "Expedited?", expedited_check
  CheckBox 170, 180, 65, 10, "R/R explained?", R_R_check
  CheckBox 240, 180, 80, 10, "Intake packet given?", intake_packet_check
  CheckBox 325, 180, 70, 10, "EBT referral sent?", EBT_referral_check
  CheckBox 25, 195, 95, 10, "Workforce referral made?", WF1_check
  CheckBox 135, 195, 70, 10, "IAAs/OMB given?", IAA_check
  CheckBox 220, 195, 65, 10, "Updated MMIS?", updated_MMIS_check
  CheckBox 295, 195, 105, 10, "Managed care packet sent?", managed_care_packet_check
  CheckBox 25, 210, 115, 10, "Managed care referral made?", managed_care_referral_check
  CheckBox 150, 210, 290, 10, "Check here to have the script update PND2 to show client delay (pending cases only).", client_delay_check
  CheckBox 25, 225, 265, 10, "Check here to have the script create a TIKL to deny at the 30/45 day mark.", TIKL_check
  EditBox 55, 240, 230, 15, other_notes
  ComboBox 330, 240, 115, 15, "incomplete"+chr(9)+"complete", CAF_status
  EditBox 55, 260, 390, 15, verifs_needed
  EditBox 55, 280, 390, 15, actions_taken
  EditBox 395, 300, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 320, 50, 15
    CancelButton 395, 320, 50, 15
    PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
    PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
    PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
    PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
    PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 50, 25, 10, "INSA/", INSA_button
    PushButton 30, 50, 25, 10, "MEDI:", MEDI_button
    PushButton 5, 70, 25, 10, "ACCI:", ACCI_button
    PushButton 5, 90, 25, 10, "DIET:", DIET_button
    PushButton 215, 90, 25, 10, "BILS:", BILS_button
    PushButton 5, 110, 25, 10, "FMED:", FMED_button
    PushButton 330, 110, 55, 10, "HC begin date:", HCRE_button
    PushButton 265, 325, 60, 10, "previous page", previous_to_page_02_button
  GroupBox 5, 5, 130, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 140, 170, 10, "Reason expedited wasn't processed (if applicable):"
  Text 5, 160, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 245, 50, 10, "Other notes:"
  Text 290, 245, 40, 10, "CAF status:"
  Text 5, 265, 50, 10, "Verifs needed:"
  Text 5, 285, 50, 10, "Actions taken:"
  Text 330, 305, 60, 10, "Worker signature:"
EndDialog

BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog

BeginDialog cancel_dialog, 0, 0, 141, 51, "Cancel dialog"
  Text 5, 5, 135, 10, "Are you sure you want to end this script?"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 10, "No, take me back to the script dialog.", no_cancel_button
    PushButton 20, 35, 105, 10, "Yes, close this script.", yes_cancel_button
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_check = 1 'The script should default to having the application signed.


'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You are not in MAXIS or you are locked out of your case.")


'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact a Support Team member.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.


'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
  call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
Else
  call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
End if
If HC_check = 1 and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_check = 1 then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable

'I put these sections in here, just because SHEL should come before HEST, it just looks cleaner.
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST) 
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST) 

'Now it grabs the rest of the info, not dependent on which programs are selected.
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "FMED", FMED)
call autofill_editbox_from_MAXIS(HH_member_array, "IMIG", IMIG)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
call autofill_editbox_from_MAXIS_test(HH_member_array, "STWK", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
call autofill_editbox_from_MAXIS_test(HH_member_array, "WREG", notes_on_abawd)

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
earned_income = trim(earned_income)
if right(earned_income, 1) = ";" then earned_income = left(earned_income, len(earned_income) - 1)
earned_income = replace(earned_income, "$________/non-monthly", "amt unknown")
earned_income = replace(earned_income, "$________/monthly", "amt unknown")
earned_income = replace(earned_income, "$________/weekly", "amt unknown")
earned_income = replace(earned_income, "$________/biweekly", "amt unknown")
earned_income = replace(earned_income, "$________/semimonthly", "amt unknown")
unearned_income = trim(unearned_income)
if right(unearned_income, 1) = ";" then unearned_income = left(unearned_income, len(unearned_income) - 1)
unearned_income = replace(unearned_income, "$________/non-monthly", "amt unknown")
unearned_income = replace(unearned_income, "$________/monthly", "amt unknown")
unearned_income = replace(unearned_income, "$________/weekly", "amt unknown")
unearned_income = replace(unearned_income, "$________/biweekly", "amt unknown")
unearned_income = replace(unearned_income, "$________/semimonthly", "amt unknown")
other_assets = trim(other_assets)
if right(other_assets, 1) = ";" then other_assets = left(other_assets, len(other_assets) - 1)
CASH_ACCTs = trim(CASH_ACCTs)
if right(CASH_ACCTs, 1) = ";" then CASH_ACCTs = left(CASH_ACCTs, len(CASH_ACCTs) - 1)
COEX_DCEX = trim(COEX_DCEX)
if right(COEX_DCEX, 1) = ";" then COEX_DCEX = left(COEX_DCEX, len(COEX_DCEX) - 1)
SHEL_HEST = trim(SHEL_HEST)
if right(SHEL_HEST, 1) = ";" then SHEL_HEST = left(SHEL_HEST, len(SHEL_HEST) - 1)
PREG = trim(PREG)
if right(PREG, 1) = ";" then PREG = left(PREG, len(PREG) - 1)
SCHL = trim(SCHL)
if right(SCHL, 1) = ";" then SCHL = left(SCHL, len(SCHL) - 1)
DISA = trim(DISA)
if right(DISA, 1) = ";" then DISA = left(DISA, len(DISA) - 1)
FACI = trim(FACI)
if right(FACI, 1) = ";" then FACI = left(FACI, len(FACI) - 1)
INSA = trim(INSA)
if right(INSA, 1) = ";" then INSA = left(INSA, len(INSA) - 1)
ACCI = trim(ACCI)
if right(ACCI, 1) = ";" then ACCI = left(ACCI, len(ACCI) - 1)
DIET = trim(DIET)
if right(DIET, 1) = ";" then DIET = left(DIET, len(DIET) - 1)
FMED = trim(FMED)
if right(FMED, 1) = ";" then FMED = left(FMED, len(FMED) - 1)
ABPS = trim(ABPS)
if right(ABPS, 1) = ";" then ABPS = left(ABPS, len(ABPS) - 1)
cit_ID = trim(cit_ID)
if right(cit_ID, 1) = ";" then cit_ID = left(cit_ID, len(cit_ID) - 1)
If cash_check = 1 then programs_applied_for = programs_applied_for & "cash, "
If HC_check = 1 then programs_applied_for = programs_applied_for & "HC, "
If SNAP_check = 1 then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_check = 1 then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)
income_changes = trim(income_changes)
if right(income_changes, 1) = ";" then income_changes= left(income_changes, len(income_changes) - 1)
IMIG = trim(IMIG)
if right(IMIG, 1) = ";" then IMIG = left(IMIG, len(IMIG) - 1)

'The following shuts down the semicolons if selected in the first dialog.
If disable_semicolon_check = 1 then
  earned_income = replace(earned_income, ";", "")
  unearned_income = replace(unearned_income, ";", "")
  CASH_ACCTs = replace(CASH_ACCTs, ";", "")
  other_assets = replace(other_assets, ";", "")
  schl = replace(schl, ";", "")
  disa = replace(disa, ";", "")
  faci = replace(faci, ";", "")
  insa = replace(insa, ";", "")
  acci = replace(acci, ";", "")
  diet = replace(diet, ";", "")
  fmed = replace(fmed, ";", "")
  abps = replace(abps, ";", "")
  preg = replace(preg, ";", "")
  cit_ID = replace(cit_ID, ";", ".") 'I put a period in here because the cit_ID variable does not store a comma or period normally. This should probably be fleshed out at some point.
End if

'SHOULD DEFAULT TO UPDATING PND2 FOR CLIENT DELAY FOR APPLICATION THAT AREN'T RECERTS. SHOULD ALSO DEFAULT TO TIKLING.
If CAF_type <> "Recertification" then
  client_delay_check = 1
  TIKL_check = 1
End if

'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Do
  Do
    Do
      Do
        Do
          Dialog CAF_dialog_01
          If ButtonPressed = 0 then 
            dialog cancel_dialog
            If ButtonPressed = yes_cancel_button then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then call stat_navigation
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed <> next_to_page_02_button then call navigation_buttons
    Loop until ButtonPressed = next_to_page_02_button
    Do
      Do
        Do
          Do
            Dialog CAF_dialog_02
            If ButtonPressed = 0 then 
              dialog cancel_dialog
              If ButtonPressed = yes_cancel_button then stopscript
            End if
          Loop until ButtonPressed <> no_cancel_button
          EMReadScreen STAT_check, 4, 20, 21
          If STAT_check = "STAT" then call stat_navigation
          transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
          EMReadScreen MAXIS_check, 5, 1, 39
          If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
        Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
        If ButtonPressed <> next_to_page_03_button then call navigation_buttons
      Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button
      If ButtonPressed = previous_to_page_01_button then exit do
      Do
        Do
          Do
            Dialog CAF_dialog_03
            If ButtonPressed = 0 then 
              dialog cancel_dialog
              If ButtonPressed = yes_cancel_button then stopscript
            End if
          Loop until ButtonPressed <> no_cancel_button
          EMReadScreen STAT_check, 4, 20, 21
          If STAT_check = "STAT" then call stat_navigation
          transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
          EMReadScreen MAXIS_check, 5, 1, 39
          If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
        Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
        If ButtonPressed <> -1 then call navigation_buttons
        If ButtonPressed = previous_to_page_02_button then exit do
      Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button
    Loop until ButtonPressed = -1
    If ButtonPressed = previous_to_page_01_button then exit do 'In case the script skipped the third page as a result of hitting "previous page" on part 2
    If actions_taken = "" or CAF_datestamp = "" or worker_signature = "" then MsgBox "You need to fill in the datestamp and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
  Loop until actions_taken <> "" and CAF_datestamp <> "" and worker_signature <> "" 
  If ButtonPressed = -1 then dialog case_note_dialog
  If buttonpressed = yes_case_note_button then
    If client_delay_check = 1 and CAF_type <> "Recertification" then 'UPDATES PND2 FOR CLIENT DELAY IF CHECKED
      call navigate_to_screen("rept", "pnd2")
      EMGetCursor PND2_row, PND2_col
      for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
        EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
        If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
        EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
        If PND2_HC_status_check = "P" then
          EMWriteScreen "x", PND2_row, 3
          transmit
          person_delay_row = 7
          Do
            EMReadScreen person_delay_check, 1, person_delay_row, 39
            If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
            person_delay_row = person_delay_row + 2
          Loop until person_delay_check = " " or person_delay_row > 20
          PF3
        End if
        EMReadScreen additional_app_check, 14, PND2_row + 1, 17
        If additional_app_check <> "ADDITIONAL APP" then exit for
        PND2_row = PND2_row + 1
      next
      PF3
      EMReadScreen PND2_check, 4, 2, 52
      If PND2_check = "PND2" then
        MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
        PF10
        client_delay_check = 0
      End if
    End if
    If TIKL_check = 1 and CAF_type <> "Recertification" then
      If cash_check = 1 or EMER_check = 1 or SNAP_check = 1 then
        call navigate_to_screen("dail", "writ")
        call create_MAXIS_friendly_date(CAF_datestamp, 30, 5, 18) 
        EMSetCursor 9, 3
        If cash_check = 1 then EMSendKey "cash/"
        If SNAP_check = 1 then EMSendKey "SNAP/"
        If EMER_check = 1 then EMSendKey "EMER/"
        EMSendKey "<backspace>" & " pending 30 days. Evaluate for possible denial."
        transmit
        PF3
      End if
      If HC_check = 1 then
        call navigate_to_screen("dail", "writ")
        call create_MAXIS_friendly_date(CAF_datestamp, 45, 5, 18) 
        EMSetCursor 9, 3
        EMSendKey "HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out."
        transmit
        PF3
      End if
    End if
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
  End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"


'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = footer_month & "/" & footer_year & " recert"


'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

EMSendKey "<home>" & "***" & CAF_type & CAF_status & "***" & "<newline>"
call write_editbox_in_case_note("CAF datestamp", CAF_datestamp, 6)
if interview_date <> "" then call write_editbox_in_case_note("Interview date", interview_date, 6)
call write_editbox_in_case_note("Programs applied for", programs_applied_for, 6)
If HH_comp <> "" then call write_editbox_in_case_note("HH comp/EATS", HH_comp, 6)
If cit_id <> "" then call write_editbox_in_case_note("Cit/ID", cit_id, 6)
If IMIG <> "" then call write_editbox_in_case_note("IMIG", IMIG, 6)
If AREP <> "" then call write_editbox_in_case_note("AREP", AREP, 6)
If FACI <> "" then call write_editbox_in_case_note("FACI", FACI, 6)
If SCHL <> "" then call write_editbox_in_case_note("SCHL/STIN/STEC", SCHL, 6)
If DISA <> "" then call write_editbox_in_case_note("DISA", DISA, 6)
If PREG <> "" then call write_editbox_in_case_note("PREG", PREG, 6)
If ABPS <> "" then call write_editbox_in_case_note("ABPS", ABPS, 6)
If earned_income <> "" then call write_editbox_in_case_note("Earned income", earned_income, 6)
If unearned_income <> "" then call write_editbox_in_case_note("Unearned income", unearned_income, 6)
If income_changes <> "" then call write_editbox_in_case_note("STWK/inc. changes", income_changes, 6)
IF notes_on_abawd <> "" then call write_editbox_in_case_note("ABAWD Notes", notes_on_abawd, 6)
If notes_on_income <> "" then call write_editbox_in_case_note("Notes on income", notes_on_income, 6)
If is_any_work_temporary <> "" then call write_editbox_in_case_note("Is any work temporary", is_any_work_temporary, 6)
If SHEL_HEST <> "" then call write_editbox_in_case_note("SHEL/HEST", SHEL_HEST, 6)
If COEX_DCEX <> "" then call write_editbox_in_case_note("COEX/DCEX", COEX_DCEX, 6)
If CASH_ACCTs <> "" then call write_editbox_in_case_note("CASH/ACCTs", CASH_ACCTs, 6)
If other_assets <> "" then call write_editbox_in_case_note("Other assets", other_assets, 6)
If INSA <> "" then call write_editbox_in_case_note("INSA", INSA, 6)
If ACCI <> "" then call write_editbox_in_case_note("ACCI", ACCI, 6)
If DIET <> "" then call write_editbox_in_case_note("DIET", DIET, 6)
If BILS <> "" then call write_editbox_in_case_note("BILS", BILS, 6)
If FMED <> "" then call write_editbox_in_case_note("FMED", FMED, 6)
If HC_begin <> "" then call write_editbox_in_case_note("HC begin date", HC_begin, 6)
If application_signed_check = 1 then call write_new_line_in_case_note("* Application was signed.")
If application_signed_check = 0 then call write_new_line_in_case_note("* Application was not signed.")
If expedited_check = 1 then call write_new_line_in_case_note("* Expedited SNAP.")
If reason_expedited_wasnt_processed <> "" then call write_editbox_in_case_note("Reason expedited wasn't processed", reason_expedited_wasnt_processed, 6)
If R_R_check = 1 then call write_new_line_in_case_note("* R/R explained to client.")
If intake_packet_check = 1 then call write_new_line_in_case_note("* Client received intake packet.")
If EBT_referral_check = 1 then call write_new_line_in_case_note("* EBT referral made for client.")
If WF1_check = 1 then call write_new_line_in_case_note("* Workforce referral made.")
If IAA_check = 1 then call write_new_line_in_case_note("* IAAs/OMB given to client.")
If updated_MMIS_check = 1 then call write_new_line_in_case_note("* Updated MMIS.")
If managed_care_packet_check = 1 then call write_new_line_in_case_note("* Client received managed care packet.")
If managed_care_referral_check = 1 then call write_new_line_in_case_note("* Managed care referral made.")
If client_delay_check = 1 then call write_new_line_in_case_note("* PND2 updated to show client delay.")
if FIAT_reasons <> "" then call write_editbox_in_case_note("FIAT reasons", FIAT_reasons, 6)
if other_notes <> "" then call write_editbox_in_case_note("Other notes", other_notes, 6)
if verifs_needed <> "" then call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
call write_editbox_in_case_note("Actions taken", actions_taken, 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")