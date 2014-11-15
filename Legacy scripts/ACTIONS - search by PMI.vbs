EMConnect ""



BeginDialog PMI_number_dialog, 0, 0, 161, 42, "PMI number"
  EditBox 95, 0, 60, 15, PMI_number
  ButtonGroup ButtonPressed_PMI_number
    OkButton 25, 20, 50, 15
    CancelButton 85, 20, 50, 15
  Text 5, 5, 85, 10, "Enter the PMI number:"
EndDialog

Dialog PMI_number_dialog

If ButtonPressed_PMI_number = 0 then stopscript

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 0



'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"



'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMSendKey "<attn>"
EMWaitReady 1, 0
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If training_check = "RUNNING" then EMSendKey "3" + "<enter>"
If production_check = "RUNNING" then EMSendKey "1" + "<enter>"

'This Do...loop gets back to SELF
do
EMSendKey "<PF3>"
EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWaitReady 1, 0

EMSendKey "<home>" + "pers" + "<eraseeof>" + "<enter>"
EMWaitReady 1, 0
EMSetcursor 15, 36
EMSendKey PMI_number + "<enter>"
EMWaitReady 1, 0
EMFocus