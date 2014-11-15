'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - MEDI CEI"
start_time = timer

'LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script




BeginDialog CEI_dialog, 0, 0, 271, 257, "CEI dialog"
  EditBox 60, 0, 70, 15, case_number_01
  EditBox 220, 0, 45, 15, medicare_amount_01
  EditBox 60, 20, 70, 15, case_number_02
  EditBox 220, 20, 45, 15, medicare_amount_02
  EditBox 60, 40, 70, 15, case_number_03
  EditBox 220, 40, 45, 15, medicare_amount_03
  EditBox 60, 60, 70, 15, case_number_04
  EditBox 220, 60, 45, 15, medicare_amount_04
  EditBox 60, 80, 70, 15, case_number_05
  EditBox 220, 80, 45, 15, medicare_amount_05
  EditBox 60, 100, 70, 15, case_number_06
  EditBox 220, 100, 45, 15, medicare_amount_06
  EditBox 60, 120, 70, 15, case_number_07
  EditBox 220, 120, 45, 15, medicare_amount_07
  EditBox 60, 140, 70, 15, case_number_08
  EditBox 220, 140, 45, 15, medicare_amount_08
  EditBox 60, 160, 70, 15, case_number_09
  EditBox 220, 160, 45, 15, medicare_amount_09
  EditBox 60, 180, 70, 15, case_number_10
  EditBox 220, 180, 45, 15, medicare_amount_10
  EditBox 165, 200, 40, 15, reimbursement_month
  EditBox 140, 220, 50, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 165, 240, 50, 15
    CancelButton 220, 240, 50, 15
  Text 155, 105, 60, 10, "Medicare amount:"
  Text 155, 5, 60, 10, "Medicare amount:"
  Text 5, 125, 50, 10, "Case number: "
  Text 155, 45, 60, 10, "Medicare amount:"
  Text 155, 125, 60, 10, "Medicare amount:"
  Text 5, 45, 50, 10, "Case number: "
  Text 5, 145, 50, 10, "Case number: "
  Text 5, 65, 50, 10, "Case number: "
  Text 155, 145, 60, 10, "Medicare amount:"
  Text 5, 25, 50, 10, "Case number: "
  Text 5, 165, 50, 10, "Case number: "
  Text 155, 65, 60, 10, "Medicare amount:"
  Text 155, 165, 60, 10, "Medicare amount:"
  Text 5, 5, 50, 10, "Case number: "
  Text 5, 185, 50, 10, "Case number: "
  Text 5, 85, 50, 10, "Case number: "
  Text 155, 185, 60, 10, "Medicare amount:"
  Text 155, 25, 60, 10, "Medicare amount:"
  Text 55, 205, 110, 10, "Reimbursement month (MM/YY):"
  Text 155, 85, 60, 10, "Medicare amount:"
  Text 70, 225, 65, 10, "Sign the case note:"
  Text 5, 105, 50, 10, "Case number: "
EndDialog

EMConnect ""

Do
  Dialog CEI_dialog
  IF buttonpressed = 0 then stopscript
  'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop. It prioritizes the training region over production.
  EMSendKey "<attn>"
  EMWaitReady 1, 1
  Do
    EMReadScreen MAI_check, 3, 1, 33
    If MAI_check <> "MAI" then EMWaitReady 1, 1
  Loop until MAI_check = "MAI"
  EMReadScreen training_check, 7, 8, 15
  EMReadScreen production_check, 7, 6, 15
  If training_check <> "RUNNING" and production_check <> "RUNNING" then 
    MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
    StopScript
  End if
  If training_check = "RUNNING" then 
    EMSendKey "3" + "<enter>"
  Else
    EMSendKey "1" + "<enter>"
  End if
  EMWaitReady 1, 1
  Do
    EMReadScreen MAI_check, 3, 1, 33
    If MAI_check <> "MAI" then EMWaitReady 1, 1
  Loop until MAI_check <> "MAI"
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  IF MAXIS_check <> "MAXIS" then MsgBox "You appear to be locked out of your case. You might need to type your password."
Loop until MAXIS_check = "MAXIS"



back_to_self





'Now the script will go to case note the contents of each case listed.

If case_number_01 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_01
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_01 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_02 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_02
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_02 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_03 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_03
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_03 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_04 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_04
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_04 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_05 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_05
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_05 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_06 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_06
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_06 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_07 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_07
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_07 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_08 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_08
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_08 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_09 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_09
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_09 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

If case_number_10 <> "" then
  EMWaitReady 1, 1
  EMSetCursor 16, 43
  EMSendKey "case"
  EMSetCursor 18, 43
  EMSendKey "<eraseeof>" + case_number_10
  EMSetCursor 21, 70
  EMSendKey "curr" + "<enter>"
  EMWaitReady 1, 1
  EMReadScreen error_check, 37, 24, 2
  If error_check <> "                                     " then MsgBox "Error! See the bottom of your MAXIS screen."
  If error_check <> "                                     " then stopscript
  EMReadScreen county_check, 4, 21, 14
  If county_check <> "X102" then MsgBox "This case is out of county. The script will now stop."
  If county_check <> "X102" then stopscript
  row = 1
  col = 1
  EMSearch "MA: ACTIVE", row, col
  If row = 0 then
    row = 1
    col = 1
    EMSearch "IMD: ACTIVE", row, col
    If row = 0 then MsgBox "This case is not open on MA or IMD. Check the case status for this client after the script runs. The script will skip this client and move on to the next case."
  End if
  If row <> 0 then
    EMWriteScreen "note", 20, 69
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMSendKey "<PF9>"
    EMWaitReady 1, 1
    'Now it is case noting the contents.
    EMSendKey "Medicare reimbursement for " + reimbursement_month + " sent to fiscal" + "<newline>"
    EMSendKey "* Medicare amount: " + medicare_amount_10 + "<newline>"
    EMSendKey "---" + "<newline>"
    EMSendKey worker_sig
  End if
  'Now it returns to SELF to start again.
  back_to_self
End if

MsgBox "Your cases have been case noted! Don't forget to send the authorization for payment to fiscal, attn: Bonnie Broda."

script_end_procedure("")