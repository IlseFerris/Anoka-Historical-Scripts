'Informational front-end message, date dependent.
If datediff("d", "08/05/2013", now) < 4 then MsgBox "This script has been updated as of 08/05/2013! Here's what's new:" & chr(13) & chr(13) & "There's now a ''FIAT reasons'' section at the end of the case note. Make sure you fill in any reasons you FIATed this case. This is required per the PCs."

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - CSR"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

next_month = dateadd("m", + 1, date)

footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog case_number_dialog, 0, 0, 181, 115, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 35, 10, "SNAP", SNAP_check
  CheckBox 55, 60, 35, 10, "WB", WB_check
  CheckBox 95, 60, 30, 10, "HC", HC_check
  CheckBox 10, 80, 100, 10, "Is this a paperless (*) IR?", paperless_check
  ButtonGroup ButtonPressed
    OkButton 35, 95, 50, 15
    CancelButton 95, 95, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Programs recertifying"
EndDialog

BeginDialog CSR_dialog, 0, 0, 451, 275, "CSR dialog"
  EditBox 65, 15, 50, 15, CSR_datestamp
  DropListBox 170, 15, 75, 15, "complete"+chr(9)+"incomplete", CSR_status
  EditBox 40, 35, 280, 15, HH_comp
  EditBox 65, 55, 380, 15, earned_income
  EditBox 70, 75, 375, 15, unearned_income
  EditBox 40, 95, 405, 15, assets
  EditBox 60, 115, 95, 15, SHEL_HEST
  EditBox 220, 115, 95, 15, COEX_DCEX
  EditBox 100, 135, 345, 15, FIAT_reasons
  EditBox 50, 155, 395, 15, other_notes
  EditBox 45, 175, 400, 15, changes
  EditBox 240, 195, 205, 15, verifs_needed
  EditBox 235, 215, 210, 15, actions_taken
  EditBox 380, 235, 65, 15, worker_signature
  CheckBox 105, 240, 175, 10, "Check here to case note grant info from ELIG/FS.", grab_FS_info_check
  CheckBox 105, 255, 210, 15, "Check here if CSR and cash supplement were used as a HRF.", HRF_check
  ButtonGroup ButtonPressed
    OkButton 340, 255, 50, 15
    CancelButton 395, 255, 50, 15
    PushButton 260, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 300, 15, 20, 10, "WB", ELIG_WB_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 120, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 120, 25, 10, "HEST:", HEST_button
    PushButton 165, 120, 25, 10, "COEX/", COEX_button
    PushButton 190, 120, 25, 10, "DCEX:", DCEX_button
    PushButton 10, 210, 25, 10, "BUSI", BUSI_button
    PushButton 35, 210, 25, 10, "JOBS", JOBS_button
    PushButton 10, 220, 25, 10, "RBIC", RBIC_button
    PushButton 35, 220, 25, 10, "UNEA", UNEA_button
    PushButton 75, 210, 25, 10, "ACCT", ACCT_button
    PushButton 100, 210, 25, 10, "CARS", CARS_button
    PushButton 125, 210, 25, 10, "CASH", CASH_button
    PushButton 150, 210, 25, 10, "OTHR", OTHR_button
    PushButton 75, 220, 25, 10, "REST", REST_button
    PushButton 100, 220, 25, 10, "SECU", SECU_button
    PushButton 125, 220, 25, 10, "TRAN", TRAN_button
    PushButton 10, 250, 25, 10, "MEMB", MEMB_button
    PushButton 35, 250, 25, 10, "MEMI", MEMI_button
    PushButton 60, 250, 25, 10, "REVW", REVW_button
  GroupBox 255, 5, 70, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation:"
  Text 5, 20, 55, 10, "CSR datestamp:"
  Text 125, 20, 40, 10, "CSR status:"
  Text 5, 40, 35, 10, "HH comp:"
  Text 5, 60, 55, 10, "Earned income:"
  Text 5, 80, 60, 10, "Unearned income:"
  Text 5, 100, 30, 10, "Assets:"
  Text 5, 140, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 160, 40, 10, "Other notes:"
  Text 5, 180, 35, 10, "Changes?:"
  GroupBox 5, 200, 60, 35, "Income panels"
  GroupBox 70, 200, 110, 35, "Asset panels"
  Text 185, 200, 50, 10, "Verifs needed:"
  Text 185, 220, 50, 10, "Actions taken:"
  GroupBox 5, 240, 85, 25, "other STAT panels:"
  Text 315, 240, 65, 10, "Worker signature:"
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
HH_memb_row = 5
Dim row
Dim col


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to MAXIS
EMConnect ""

'Searching for the case_number variable
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Searching for the footer month
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then call script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'If "paperless" was checked, the script will put a simple case note in and end.
If paperless_check = 1 then
  call navigate_to_screen("case", "note")
  PF9
  EMWriteScreen "***Cleared paperless IR for " & footer_month & "/" & footer_year & "***", 4, 3
  EMWriteScreen "---", 5, 3
  worker_signature = InputBox ("Sign your case note:", "worker signature")
  EMWriteScreen worker_signature, 6, 3
  call script_end_procedure("")
End if

'Navigating to STAT/REVW, checking for error prone cases
call navigate_to_screen("stat", "revw")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then call script_end_procedure("Can't get into STAT. This case may be in background. Wait a few seconds and try again. If this case is not in background email Veronica the case number and footer month.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.


'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabbing SHEL/HEST first, and putting them in this special order that everyone seems to like
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)

'Autofilling the rest
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CSR_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)




'Cleaning up the case note
CSR_month = footer_month & "/" & footer_year
earned_income = trim(earned_income)
if right(earned_income, 1) = ";" then earned_income = left(earned_income, len(earned_income) - 1)
earned_income = replace(earned_income, "$________/non-monthly", "amt unknown")
unearned_income = trim(unearned_income)
if right(unearned_income, 1) = ";" then unearned_income = left(unearned_income, len(unearned_income) - 1)
unearned_income = replace(unearned_income, "$________/non-monthly", "amt unknown")
assets = trim(assets)
if right(assets, 1) = ";" then assets = left(assets, len(assets) - 1)
SHEL_HEST = trim(SHEL_HEST)
if right(SHEL_HEST, 1) = ";" then SHEL_HEST = left(SHEL_HEST, len(SHEL_HEST) - 1)
COEX_DCEX = trim(COEX_DCEX)
if right(COEX_DCEX, 1) = ";" then COEX_DCEX = left(COEX_DCEX, len(COEX_DCEX) - 1)
If cash_check = 1 then programs_recertifying = programs_recertifying & "cash, "
If HC_check = 1 then programs_recertifying = programs_recertifying & "HC, "
If SNAP_check = 1 then programs_recertifying = programs_recertifying & "SNAP, "
programs_recertifying = trim(programs_recertifying)
if right(programs_recertifying, 1) = "," then programs_recertifying = left(programs_recertifying, len(programs_recertifying) - 1)



'Showing the case note dialog
Do
  Do
    Do
      Do
        Do
          Dialog CSR_dialog
          If ButtonPressed = 0 then 
            dialog cancel_dialog
            If ButtonPressed = yes_cancel_button then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then
          If ButtonPressed = prev_panel_button then call panel_navigation_prev
          If ButtonPressed = next_panel_button then call panel_navigation_next
          If ButtonPressed = prev_memb_button then call memb_navigation_prev
          If ButtonPressed = next_memb_button then call memb_navigation_next
        End if
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
      If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
      If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
      If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
      If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
      If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
      If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
      If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
      If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
      If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
      If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
      If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
      If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
      If ButtonPressed = SHEL_button then call navigate_to_screen("stat", "SHEL")
      If ButtonPressed = HEST_button then call navigate_to_screen("stat", "HEST")
      If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
      If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
      If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
      If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
      If ButtonPressed = ELIG_WB_button then call navigate_to_screen("elig", "WB__")
    Loop until ButtonPressed = -1
    If (earned_income = "" and unearned_income = "") or actions_taken = "" or CSR_datestamp = "" or worker_signature = "" then MsgBox "You need to fill in the datestamp, income, and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
  Loop until (earned_income <> "" or unearned_income <> "") and actions_taken <> "" and CSR_datestamp <> "" and worker_signature <> ""
  If ButtonPressed = -1 then dialog case_note_dialog
  If buttonpressed = yes_case_note_button then
    If grab_FS_info_check = 1 then
      call navigate_to_screen("elig", "fs")
      EMReadScreen FSPR_check, 4, 3, 48
      If FSPR_check <> "FSPR" then
        MsgBox "The script couldn't find ELIG/FS. It will now jump to case note."
      Else
        EMWriteScreen "FSSM", 19, 70
        transmit
        EMReadScreen FSSM_line_01, 37, 13, 44
        EMReadScreen FSSM_line_02, 37, 8, 3
        EMReadScreen FSSM_line_03, 37, 10, 3
      End if
    End if
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
  End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'Writing the case note to MAXIS
EMSendKey "<home>" & "***" & CSR_month & " CSR received " & CSR_datestamp & ": " & CSR_status & "***" & "<newline>"
call write_editbox_in_case_note("Programs recertifying", programs_recertifying, 6)
call write_editbox_in_case_note("HH comp", HH_comp, 6)
If earned_income <> "" then call write_editbox_in_case_note("Earned income", earned_income, 6)
If unearned_income <> "" then call write_editbox_in_case_note("Unearned income", unearned_income, 6)
If assets <> "" then call write_editbox_in_case_note("Assets", assets, 6)
If SHEL_HEST <> "" then call write_editbox_in_case_note("SHEL/HEST", SHEL_HEST, 6)
If COEX_DCEX <> "" then call write_editbox_in_case_note("COEX/DCEX", COEX_DCEX, 6)
if FIAT_reasons <> "" then call write_editbox_in_case_note("FIAT reasons", FIAT_reasons, 6)
if other_notes <> "" then call write_editbox_in_case_note("Other notes", other_notes, 6)
If changes <> "" then call write_editbox_in_case_note("Changes", changes, 6)
If HRF_check = 1 then call write_new_line_in_case_note("* CSR and cash supplement used as HRF.")
if verifs_needed <> "" then call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
call write_editbox_in_case_note("Actions taken", actions_taken, 6)
call write_new_line_in_case_note("---")
If FSPR_check = "FSPR" then
  call write_new_line_in_case_note("   " & FSSM_line_01)
  call write_new_line_in_case_note("   " & FSSM_line_02)
  call write_new_line_in_case_note("   " & FSSM_line_03)
  call write_new_line_in_case_note("---")
End if
call write_new_line_in_case_note(worker_signature)

call script_end_procedure("")