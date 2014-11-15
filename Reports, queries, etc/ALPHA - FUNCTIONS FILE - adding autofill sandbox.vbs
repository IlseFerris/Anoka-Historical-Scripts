'Variables that need to be considered:
' HH_member (this is the "for each" in the for each...in statements and should be changed to avoid issues in case notes)
' HH_member_array 
' panel that is being read from
' variable that's being written to

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

HH_member_array = array("01", "02")
case_number = "180940"
panel_to_read_from_for_testing = "PBEN"

Function autofill_editbox_from_MAXIS_sandbox(HH_member_array, panel_read_from, variable_written_to)
  If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
    call navigate_to_screen("stat", "ABPS")
    EMReadScreen ABPS_total_pages, 1, 2, 78
    If ABPS_total_pages <> 0 then 
      Do
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
        variable_written_to = variable_written_to & trim(new_ABPS_current) & "; "
        new_ABPS_current = "" 'need to reset the variable
        EMReadScreen ABPS_current_page, 1, 2, 73
        If ABPS_current_page <> ABPS_total_pages then transmit
      Loop until ABPS_current_page = ABPS_total_pages
    End if
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
  Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
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
End function

EMConnect ""

call autofill_editbox_from_MAXIS_sandbox(HH_member_array, panel_to_read_from_for_testing, editbox_to_write_to_for_testing)
MsgBox "* " & panel_to_read_from_for_testing & ": " & editbox_to_write_to_for_testing