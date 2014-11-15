'SECTION 01: FUNCTIONS----------------------------------------------------------------------------------------------------

Function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
End function

Function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

Function PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
End function

Function PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
End function

Function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
End function

Function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" then
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
    EMWaitReady 0, 0
  Else
    Do
      EMSendKey "<PF3>"
      EMWaitReady 0, 0
      EMReadScreen SELF_check, 4, 2, 50
    Loop until SELF_check = "SELF"
    EMWriteScreen x, 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen case_number, 18, 43
    EMWriteScreen y, 21, 70
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMReadScreen abended_check, 7, 9, 27
    If abended_check = "abended" then
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    End if
  End if
  End if
End function

'SECTION 02: DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog new_worker_dialog, 0, 0, 246, 102, "New worker name and x102 dialog"
  EditBox 85, 45, 50, 15, new_worker_name
  EditBox 110, 65, 35, 15, x102_number
  CheckBox 5, 85, 155, 15, "Check here if this worker is an Adult worker.", adult_check
  ButtonGroup ButtonPressed
    OkButton 185, 50, 50, 15
    CancelButton 185, 70, 50, 15
  Text 5, 5, 240, 35, "This script will send a SPEC/MEMO version of the new worker letter for each case on a worker's REPT/ACTV. Do not use this script if the worker does not have all new cases. Also, check the worker's phone number for accuracy prior to using the script."
  Text 15, 50, 65, 10, "New worker name: "
  Text 15, 70, 90, 10, "Worker number (x102###):"
EndDialog


'CONNECTING TO MACHINE.
EMConnect ""

'REQUIRES PASSWORD TO RUN.
password = InputBox("This script is for supervisors only. Enter your password here.")
If password = "" then stopscript
If password <> "countymemo" then
  MsgBox "Incorrect password."
  StopScript
End If

'RUNS THE DIALOG.
Do
  Do
    Dialog new_worker_dialog
    If ButtonPressed = 0 then stopscript
    If len(x102_number) <> 3 then MsgBox "Your x102 number must be exactly 3 digits long."
  Loop until len(x102_number) = 3
  If new_worker_name = "" then MsgBox "You must put in a new worker name."
Loop until new_worker_name <> "" 

'CHECKS TO MAKE SURE MAXIS IS NOT PASSWORDED OUT.
PF3
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then
  MsgBox "MAXIS is not found. Are you passworded out?"
  StopScript
End if

'NAVIGATES TO REPT/ACTV. 
call navigate_to_screen("rept", "actv")
EMReadScreen ACTV_check, 4, 2, 48
If ACTV_check <> "ACTV" then
  MsgBox "Not on REPT/ACTV. Your MAXIS window may have errored/passworded out. Clear any password prompts and try again."
  StopScript
End if

'WRITES THE X102 NUMBER TO BE READ. BACKS UP ONCE TO THE FIRST PAGE, IN CASE THE INITIAL X102 NUMBER WAS THE DEFAULT NUMBER.
EMWriteScreen x102_number, 21, 17
transmit
PF7

'DECLARES THE ROW VARIABLE, SO THAT THE SCRIPT WILL KNOW WHERE TO FIND CASE INFO
row = 7

'READS ALL CASE NUMBERS OFF OF REPT/ACTV
Do
  EMReadScreen case_number, 8, row, 12
  EMReadScreen more_pages_check, 1, 19, 9
  If case_number = "        " then
    row = 7
    PF8
  Else
    case_number = trim(case_number)
    case_number_array = case_number_array & "|" & case_number
    row = row + 1
  End if
Loop until case_number = "        " and more_pages_check <> "+"

'SPLITTING THE CASE NUMBERS INTO AN ARRAY. NOTIFYING THE SUP AS TO HOW MANY CASES WILL GET THE LETTER SENT, AND GIVING THEM ONE LAST CHANCE TO CANCEL THE SCRIPT.
case_number_array = split(case_number_array, "|")
If ubound(case_number_array) < 1 then 
  MsgBox "This worker does not appear to have any cases. The script will now close."
  StopScript
End if
last_chance = MsgBox (ubound(case_number_array) & " cases will have letters sent.", 1, 0)
If last_chance = 2 then stopscript

'NAVIGATING TO SPEC/MEMO FOR EACH CASE, AND WRITING THE MEMO FOR EACH IN THE ARRAY
For each x in case_number_array
  If isnumeric(x) = True then
    case_number = x
    call navigate_to_screen("spec", "memo")
    PF5
    EMWriteScreen "x", 5, 10
    transmit
    EMSendKey "Your case has been reassigned. Your new worker is: " & new_worker_name & "<newline>" & "<newline>"
    EMSendKey "Call your worker:" & "<newline>"
    EMSendKey " > With questions or to report changes." & "<newline>"
    EMSendKey " > To schedule appointments if you need to see your worker." & "<newline>"
    EMSendKey " > If you need to apply for other programs." & "<newline>" & "<newline>"
    EMSendKey "Your case number is " & case_number & ". Have your case number ready when you call. Please read everything we send you, and follow all instructions." & "<newline>" & "<newline>"
    EMSendKey "For EBT cards, call 1-888-997-2227:" & "<newline>"
    EMSendKey " > After 10:00 AM to check for current/new benefit info." & "<newline>"
    EMSendKey " > To report your card lost, stolen, or damaged." & "<newline>"
    PF8
    EMSendKey "   NOTE: There is a $2 replacement fee if your card is lost" & "<newline>"
    EMSendKey "         or stolen, or if you destroyed it when your case" & "<newline>"
    EMSendKey "         was closed." & "<newline>" & "<newline>"
    If adult_check = 0 then EMSendKey "Call 763-422-7320 if you have any questions about Child Support. Call the Employment Service reschedule line at 763-783-4885 if you cannot come to a scheduled Overview or Assessment Workshop."
    PF4
  End if
Next

'SUCCESS BOX
MsgBox "Success!"