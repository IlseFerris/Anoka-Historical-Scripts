'Removed MCRE scripts 01/2014 due to MNsure.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MMIS - MCRE application"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

BeginDialog MCRE_dialog, 5, 5, 291, 392, "MCRE dialog"
  EditBox 60, 5, 70, 15, app_date
  EditBox 200, 5, 40, 15, all_sigs
  EditBox 45, 25, 240, 15, HH_comp
  EditBox 55, 45, 75, 15, MN_residency
  EditBox 195, 45, 80, 15, US_cit_verif
  EditBox 55, 65, 80, 15, current_insa
  EditBox 200, 65, 85, 15, access_to_ESI
  EditBox 70, 85, 55, 15, INSA
  EditBox 235, 85, 50, 15, med_support
  EditBox 90, 105, 40, 15, assets
  EditBox 235, 105, 50, 15, BUSI_assets
  EditBox 60, 125, 225, 15, income
  EditBox 80, 145, 75, 15, RINC
  EditBox 225, 145, 60, 15, premium
  EditBox 155, 165, 105, 15, retro
  EditBox 120, 185, 165, 15, tracking
  EditBox 10, 225, 150, 15, HH_memb_01
  EditBox 170, 225, 45, 15, action_01
  EditBox 235, 225, 45, 15, group_01
  EditBox 10, 245, 150, 15, HH_memb_02
  EditBox 170, 245, 45, 15, action_02
  EditBox 235, 245, 45, 15, group_02
  EditBox 10, 265, 150, 15, HH_memb_03
  EditBox 170, 265, 45, 15, action_03
  EditBox 235, 265, 45, 15, group_03
  EditBox 10, 285, 150, 15, HH_memb_04
  EditBox 170, 285, 45, 15, action_04
  EditBox 235, 285, 45, 15, group_04
  EditBox 10, 305, 150, 15, HH_memb_05
  EditBox 170, 305, 45, 15, action_05
  EditBox 235, 305, 45, 15, group_05
  EditBox 10, 325, 150, 15, HH_memb_06
  EditBox 170, 325, 45, 15, action_06
  EditBox 235, 325, 45, 15, group_06
  EditBox 50, 350, 235, 15, comments
  EditBox 80, 370, 70, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 170, 370, 50, 15
    CancelButton 225, 370, 50, 15
  Text 5, 70, 50, 10, "Current insa?:"
  Text 5, 130, 55, 10, "Income source:"
  Text 140, 10, 60, 10, "All required sigs:"
  Text 160, 150, 65, 10, "Monthly premium:"
  Text 5, 50, 50, 10, "MN residency:"
  Text 5, 170, 145, 10, "Retro MCRE? (If retro elig, send retro letter):"
  Text 5, 90, 60, 10, "INSA last 4 mos?:"
  Text 5, 190, 115, 10, "Tracking? If tracking, write issues:"
  Text 5, 10, 50, 10, "New app date:"
  Text 10, 210, 105, 10, "HH memb (fill in all that apply)"
  Text 170, 210, 45, 10, "action/codes"
  Text 235, 210, 45, 10, "group/status"
  Text 130, 90, 105, 10, "Med Support/# of refrls/names:"
  Text 5, 355, 40, 10, "Comments:"
  Text 135, 50, 65, 10, "US citiz verified?:"
  Text 5, 375, 70, 10, "Sign your case note:"
  Text 5, 110, 80, 10, "Total countable assets:"
  Text 5, 30, 35, 10, "HH comp:"
  Text 140, 110, 95, 10, "BUSI assets under $200k?:"
  Text 5, 150, 75, 10, "Annual income (RINC):"
  Text 145, 70, 55, 10, "Access to ESI?:"
EndDialog

Do
  Dialog MCRE_dialog
  If buttonpressed = 0 then stopscript
  EMReadScreen MMIS_case_note_check, 15, 1, 31
  EMReadScreen MMIS_edit_check, 5, 5, 2
  If MMIS_case_note_check <> "MMIS CASE NOTES" or MMIS_edit_check = "=====" then MsgBox "You are not in MMIS case note edit mode. Please get to MMIS case note edit mode before pressing OK."
Loop until MMIS_case_note_check = "MMIS CASE NOTES" and MMIS_edit_check <> "====="

EMSendKey "<PF11>" 'To check for password lockout.
EMWaitReady 1, 1
Do
   EMReadScreen password_prompt, 38, 2, 23
   IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case note. Type your password then try again."
   IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog MCRE_dialog
   IF buttonpressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

EMSendKey "***************************MCRE APPLICATION***************************" + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* APP DATE: " + app_date + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* ALL REQUIRED SIGNATURES: " + all_sigs + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* HH COMP: " + HH_comp + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* MN RESIDENCY: " + MN_residency + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* US citizenship verified: " + US_cit_verif + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Current insa: " + current_insa + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Access to ESI: " + access_to_ESI + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* INSA last 4 mos: " + INSA + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Med Support/# of refrls/names: " + med_support + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Total countable assets: " + assets + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* BUSI assets under $200k?: " + BUSI_assets + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Income source: " + income + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Annual income (RINC): " + RINC + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Monthly premium: " + premium + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Retro MCRE?: " + retro + "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Tracking/Issues: " + tracking + "<PF11>"
EMWaitReady 1, 1
EMSendKey "-----------------------------MEMB STATUS-----------------------------" + "<PF11>"
EMWaitReady 1, 1
EMSendKey "household member....................action/codes...........group status" + "<PF11>"
EMWaitReady 1, 1
If HH_memb_01 <> "" then EMSendKey "......................................................................."
If HH_memb_01 <> "" then EMWriteScreen HH_memb_01, 5, 8
If HH_memb_01 <> "" then EMWriteScreen action_01, 5, 46
If HH_memb_01 <> "" then EMWriteScreen group_01, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 1, 1
If HH_memb_02 <> "" then EMSendKey "......................................................................."
If HH_memb_02 <> "" then EMWriteScreen HH_memb_02, 5, 8
If HH_memb_02 <> "" then EMWriteScreen action_02, 5, 46
If HH_memb_02 <> "" then EMWriteScreen group_02, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 1, 1
If HH_memb_03 <> "" then EMSendKey "......................................................................."
If HH_memb_03 <> "" then EMWriteScreen HH_memb_03, 5, 8
If HH_memb_03 <> "" then EMWriteScreen action_03, 5, 46
If HH_memb_03 <> "" then EMWriteScreen group_03, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 1, 1
If HH_memb_04 <> "" then EMSendKey "......................................................................."
If HH_memb_04 <> "" then EMWriteScreen HH_memb_04, 5, 8
If HH_memb_04 <> "" then EMWriteScreen action_04, 5, 46
If HH_memb_04 <> "" then EMWriteScreen group_04, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 1, 1
If HH_memb_05 <> "" then EMSendKey "......................................................................."
If HH_memb_05 <> "" then EMWriteScreen HH_memb_05, 5, 8
If HH_memb_05 <> "" then EMWriteScreen action_05, 5, 46
If HH_memb_05 <> "" then EMWriteScreen group_05, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 1, 1
If HH_memb_06 <> "" then EMSendKey "......................................................................."
If HH_memb_06 <> "" then EMWriteScreen HH_memb_06, 5, 8
If HH_memb_06 <> "" then EMWriteScreen action_06, 5, 46
If HH_memb_06 <> "" then EMWriteScreen group_06, 5, 69
EMSendKey "<PF11>"
EMWaitReady 1, 1
EMSendKey "* Comments: " + comments + "<PF11>"
EMWaitReady 1, 1
EMSendKey "---" + "<PF11>"
EMWaitReady 1, 1
EMSendKey worker_sig + "<PF11>"
EMWaitReady 1, 1
EMSendKey "***********************************************************************"

script_end_procedure("")
