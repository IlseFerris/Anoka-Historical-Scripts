'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MMIS - MCRE renewal"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

BeginDialog MCRE_renewal_dialog, 5, 5, 291, 312, "MCRE renewal dialog"
  EditBox 75, 5, 60, 15, app_date
  EditBox 220, 5, 60, 15, redetermination_date
  EditBox 65, 25, 40, 15, all_sigs
  EditBox 150, 25, 135, 15, HH_comp
  EditBox 60, 45, 65, 15, access_to_ESI
  EditBox 235, 45, 50, 15, med_support
  EditBox 90, 65, 40, 15, assets
  EditBox 235, 65, 50, 15, BUSI_assets
  EditBox 65, 85, 220, 15, income
  EditBox 80, 105, 75, 15, RINC
  EditBox 225, 105, 60, 15, premium
  EditBox 10, 145, 150, 15, HH_memb_01
  EditBox 170, 145, 45, 15, action_01
  EditBox 235, 145, 45, 15, group_01
  EditBox 10, 165, 150, 15, HH_memb_02
  EditBox 170, 165, 45, 15, action_02
  EditBox 235, 165, 45, 15, group_02
  EditBox 10, 185, 150, 15, HH_memb_03
  EditBox 170, 185, 45, 15, action_03
  EditBox 235, 185, 45, 15, group_03
  EditBox 10, 205, 150, 15, HH_memb_04
  EditBox 170, 205, 45, 15, action_04
  EditBox 235, 205, 45, 15, group_04
  EditBox 10, 225, 150, 15, HH_memb_05
  EditBox 170, 225, 45, 15, action_05
  EditBox 235, 225, 45, 15, group_05
  EditBox 10, 245, 150, 15, HH_memb_06
  EditBox 170, 245, 45, 15, action_06
  EditBox 235, 245, 45, 15, group_06
  EditBox 50, 270, 235, 15, comments
  EditBox 80, 290, 70, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 170, 290, 50, 15
    CancelButton 225, 290, 50, 15
  Text 170, 130, 45, 10, "action/codes"
  Text 235, 130, 45, 10, "group/status"
  Text 145, 10, 80, 10, "Redetermination date:"
  Text 5, 70, 80, 10, "Total countable assets:"
  Text 110, 30, 35, 10, "HH comp:"
  Text 140, 70, 95, 10, "BUSI assets under $200k?:"
  Text 5, 10, 65, 10, "Renewal app date:"
  Text 5, 110, 75, 10, "Annual income (RINC):"
  Text 5, 275, 40, 10, "Comments:"
  Text 5, 30, 60, 10, "All required sigs:"
  Text 5, 295, 70, 10, "Sign your case note:"
  Text 160, 110, 65, 10, "Monthly premium:"
  Text 130, 50, 105, 10, "Med Support/# of refrls/names:"
  Text 10, 130, 105, 10, "HH memb (fill in all that apply)"
  Text 5, 90, 55, 10, "Income source:"
  Text 5, 50, 55, 10, "Access to ESI?:"
EndDialog

Do
  Dialog MCRE_renewal_dialog
  If buttonpressed = 0 then stopscript
  EMReadScreen MMIS_case_note_check, 15, 1, 31
  EMReadScreen MMIS_edit_check, 5, 5, 2
  If MMIS_case_note_check <> "MMIS CASE NOTES" or MMIS_edit_check = "=====" then MsgBox "You are not in MMIS case note edit mode. Please get to MMIS case note edit mode before pressing OK."
Loop until MMIS_case_note_check = "MMIS CASE NOTES" and MMIS_edit_check <> "====="

EMSendKey "<PF11>" 'To check for password lockout.
EMWaitReady 0, 0
Do
   EMReadScreen password_prompt, 38, 2, 23
   IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case note. Type your password then try again."
   IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog MCRE_renewal_dialog
   IF buttonpressed = 0 then stopscript
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

EMSendKey "*****************************MCRE RENEWAL*****************************" + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* APP DATE: " + app_date + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Redetermination date: " + redetermination_date + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* ALL REQUIRED SIGNATURES: " + all_sigs + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* HH COMP: " + HH_comp + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Access to ESI: " + access_to_ESI + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Med Support/# of refrls/names: " + med_support + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Total countable assets: " + assets + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* BUSI assets under $200k?: " + BUSI_assets + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Income source: " + income + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Annual income (RINC): " + RINC + "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Monthly premium: " + premium + "<PF11>"
EMWaitReady 0, 0
EMSendKey "-----------------------------MEMB STATUS-----------------------------" + "<PF11>"
EMWaitReady 0, 0
EMSendKey "household member....................action/codes...........group status" + "<PF11>"
EMWaitReady 0, 0
If HH_memb_01 <> "" then EMSendKey "......................................................................."
If HH_memb_01 <> "" then EMWriteScreen HH_memb_01, 5, 8
If HH_memb_01 <> "" then EMWriteScreen action_01, 5, 46
If HH_memb_01 <> "" then EMWriteScreen group_01, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 0, 0
If HH_memb_02 <> "" then EMSendKey "......................................................................."
If HH_memb_02 <> "" then EMWriteScreen HH_memb_02, 5, 8
If HH_memb_02 <> "" then EMWriteScreen action_02, 5, 46
If HH_memb_02 <> "" then EMWriteScreen group_02, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 0, 0
If HH_memb_03 <> "" then EMSendKey "......................................................................."
If HH_memb_03 <> "" then EMWriteScreen HH_memb_03, 5, 8
If HH_memb_03 <> "" then EMWriteScreen action_03, 5, 46
If HH_memb_03 <> "" then EMWriteScreen group_03, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 0, 0
If HH_memb_04 <> "" then EMSendKey "......................................................................."
If HH_memb_04 <> "" then EMWriteScreen HH_memb_04, 5, 8
If HH_memb_04 <> "" then EMWriteScreen action_04, 5, 46
If HH_memb_04 <> "" then EMWriteScreen group_04, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 0, 0
If HH_memb_05 <> "" then EMSendKey "......................................................................."
If HH_memb_05 <> "" then EMWriteScreen HH_memb_05, 5, 8
If HH_memb_05 <> "" then EMWriteScreen action_05, 5, 46
If HH_memb_05 <> "" then EMWriteScreen group_05, 5, 69 
EMSendKey "<PF11>"
EMWaitReady 0, 0
If HH_memb_06 <> "" then EMSendKey "......................................................................."
If HH_memb_06 <> "" then EMWriteScreen HH_memb_06, 5, 8
If HH_memb_06 <> "" then EMWriteScreen action_06, 5, 46
If HH_memb_06 <> "" then EMWriteScreen group_06, 5, 69
EMSendKey "<PF11>"
EMWaitReady 0, 0
EMSendKey "* Comments: " + comments + "<PF11>"
EMWaitReady 0, 0
EMSendKey "---" + "<PF11>"
EMWaitReady 0, 0
EMSendKey worker_sig + "<PF11>"
EMWaitReady 0, 0
EMSendKey "***********************************************************************"

script_end_procedure("")
