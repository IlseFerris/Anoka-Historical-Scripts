'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Functions============================================================
'Performs a MAXIS check-----------------------------------------------
function maxis_check_function
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again."
END function

'Returns the current month to create a starting column for reading the tracking record---------------
function starting_point_month(bene_mo_col)		
  current_month = datepart("m",Date())
  bene_mo_col = (15 + (4*current_month))
END function 

function how_many_abawd_months(abawd_counted_months)
  DO
    call navigate_to_screen("stat", "wreg")
    maxis_check_function
  LOOP until MAXIS_check <> "MAXIS"
    EMSetCursor 13, 57
    EMSendKey "X"
    transmit
  bene_yr_row = 10
  month_count = 0
  abawd_counted_months = 0
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
END function 

function number_of_hh_members(number_of_hh_memb)
  DO
    call navigate_to_screen("stat", "memb")
    maxis_check_function
  LOOP until MAXIS_check <> "MAXIS"
  number_of_hh_memb = 0
  read_row = 3
  DO
    EMReadScreen memb, 2, 5, read_row
    IF memb <> "  " THEN number_of_hh_memb = number_of_hh_memb + 1
    read_row = read_row + 1
  LOOP until memb_yn = "  "
END function

'THE SCRIPT-----------------------------------------------------------------------
EMConnect ""

back_to_SELF

call starting_point_month(bene_mo_col)

case_number = inputbox("Case number...")
call number_of_hh_members(number_of_hh_memb)
'call how_many_abawd_months(abawd_counted_months)
msgbox number_of_hh_memb

'MSGBox("CL has used " & abawd_counted_months & " of ABAWD-counted SNAP months in the past 36 months starting with " & datepart("m",Date()) & "/" & datepart("yyyy",Date()) & ".")


