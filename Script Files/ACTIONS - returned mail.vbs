'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - returned mail"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'FUNCTIONS

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(x) then
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen y, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 1, 1
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 1, 1
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen x, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen y, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 1, 1
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 1, 1
      End if
    End if
  End if
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 1, 1
end function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 1, 1
end function



EMConnect ""

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If isnumeric(case_number) = False then case_number = ""

BeginDialog returned_mail_dialog, 0, 0, 236, 92, "Returned mail"
  EditBox 140, 0, 65, 15, case_number
  EditBox 100, 50, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 70, 50, 15
    CancelButton 120, 70, 50, 15
  Text 35, 55, 60, 10, "Worker signature:"
  Text 30, 5, 110, 10, "Case number with returned mail:"
  Text 10, 20, 220, 25, "Note: if you have mail with an allowed forwarding address, update MAXIS per policy. Do not use this script with a forwarding address. Ask a supervisor if you have questions about returned mail policy."
EndDialog

Do
  Dialog returned_mail_dialog
  If buttonpressed = 0 then stopscript
Loop until trim(case_number) <> ""
transmit 'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then stopscript


Call navigate_to_screen("case", "note")

'If there was an error after trying to go to CASE/NOTE, the script will shut down.
EMReadScreen SELF_error_check, 27, 2, 28 
If SELF_error_check = "Select Function Menu (SELF)" then stopscript

'Now the script goes into the case note and case notes the action. 

PF9

EMReadScreen mode_check, 7, 20, 3
If mode_check <> "Mode: A" and mode_check <> "Mode: E" then
  MsgBox "Unable to start a case note. Is this inquiry mode? Is this case out of county? Right case number? Check these out and try again!"
  StopScript
End if

EMSendKey "-->Returned mail received<--" + "<newline>"
EMSendKey "* No forwarding address was indicated." + "<newline>"
EMSendKey "* Sending verification request to last known address. TIKLed for 10-day return." + "<newline>"
EMSendKey "---" + "<newline>" + worker_signature

PF3

call navigate_to_screen("dail", "writ")


ten_days_from_today = dateadd("d", date, 10)
TIKL_day = datepart("d", ten_days_from_today)
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_month = datepart("m", ten_days_from_today)
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_year = (datepart("yyyy", ten_days_from_today)) - 2000

EMWriteScreen TIKL_month, 5, 18
EMWriteScreen TIKL_day, 5, 21
EMWriteScreen TIKL_year, 5, 24
EMSetCursor 9, 3
EMSendKey "Request for address sent 10 days ago. If not responded to, close per returned mail procedure." 

transmit
PF3

MsgBox "Use the returned mail packet in Compass Forms. Send the completed forms to the most recent address. The script has case noted that returned mail was received and TIKLed out for 10-day return."
script_end_procedure("")
