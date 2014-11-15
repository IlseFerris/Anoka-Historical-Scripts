'FUNCTIONS

function PF11
  EMSendKey "<PF11>"
  EMWaitReady 1, 1
end function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 1, 1
end function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
end function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 1, 1
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 1, 1
end function

function back_to_self
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

function navigate_to_screen(x, y)
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
    EMWriteScreen y, 21, 70
    EMSendKey "<enter>"
    EMWaitReady 1, 1
  End if
End function

Function write_editbox_in_case_note(x, y)
  z = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in z 'z represents the variable
    EMGetCursor row, col 
    If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then PF8
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 17 and col + (len(x)) >= 80) then EMSendKey "<newline>" & "     "
    If (row = 4 and col = 3) then EMSendKey "     "
    EMSendKey x & " "
  Next
  EMSendKey "<newline>"
End function

Function write_new_line_in_case_note(x)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then PF8
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

Function switch_to_MMIS
  EMSendKey "<attn>"
  Do
    EMWaitReady 1, 1
    EMReadScreen MAI_check, 3, 1, 33
  Loop until MAI_check = "MAI"
  EMReadScreen MMIS_A_check, 7, 15, 15 
  IF MMIS_A_check = "RUNNING" then 
    EMSendKey "10" + "<enter>"
    Do
      EMWaitReady 1, 1
      EMReadScreen MAI_check, 3, 1, 33
    Loop until MAI_check <> "MAI"
  End if
  IF MMIS_A_check <> "RUNNING" then 
    EMSendKey "<attn>"
    EMWaitReady 1, 1
    EMConnect "B"
    EMSendKey "<attn>"
    Do
      EMWaitReady 1, 1
      EMReadScreen MAI_check, 3, 1, 33
    Loop until MAI_check = "MAI"
    EMReadScreen MMIS_B_check, 7, 15, 15
    If MMIS_B_check <> "RUNNING" then 
      MsgBox "MMIS does not appear to be running. This script will now stop."
      stopscript
    End if
    If MMIS_B_check = "RUNNING" then 
      EMSendkey "10" + "<enter>"
      Do
        EMWaitReady 1, 1
        EMReadScreen MAI_check, 3, 1, 33
      Loop until MAI_check <> "MAI"
    End if
  End if
  EMReadScreen MAI_check, 3, 1, 33
  If MAI_check = "MAI" then EMWaitReady 1, 5
End function

Function switch_to_MMIS_training
  EMSendKey "<attn>"
  Do
    EMWaitReady 1, 1
    EMReadScreen MAI_check, 3, 1, 33
  Loop until MAI_check = "MAI"
  EMReadScreen MMIS_A_check, 7, 16, 15 
  IF MMIS_A_check = "RUNNING" then 
    EMSendKey "11" + "<enter>"
    Do
      EMWaitReady 1, 1
      EMReadScreen MAI_check, 3, 1, 33
    Loop until MAI_check <> "MAI"
  End if
  IF MMIS_A_check <> "RUNNING" then 
    EMSendKey "<attn>"
    EMWaitReady 1, 1
    EMConnect "B"
    EMSendKey "<attn>"
    Do
      EMWaitReady 1, 1
      EMReadScreen MAI_check, 3, 1, 33
    Loop until MAI_check = "MAI"
    EMReadScreen MMIS_B_check, 7, 16, 15
    If MMIS_B_check <> "RUNNING" then 
      MsgBox "MMIS does not appear to be running. This script will now stop."
      stopscript
    End if
    If MMIS_B_check = "RUNNING" then 
      EMSendkey "11" + "<enter>"
      Do
        EMWaitReady 1, 1
        EMReadScreen MAI_check, 3, 1, 33
      Loop until MAI_check <> "MAI"
    End if
  End if
  EMReadScreen MAI_check, 3, 1, 33
  If MAI_check = "MAI" then EMWaitReady 1, 5
End function

Function get_to_session_begin 'This function uses a Do Loop to get to the start screen for MMIS.
  Do 
    EMSendkey "<PF6>"
    EMReadScreen password_prompt, 38, 2, 23
    IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
    EMWaitReady 1, 1
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
End function

'THE SCRIPT

EMConnect ""

transmit 'To check for a password prompt

EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  MsgBox "You do not appear to be in MAXIS. The script will now stop."
  Stopscript
End if

call find_variable("Month: ", footer_month, 2)
If footer_month <> "" then call find_variable("Month: " & footer_month & " ", footer_year, 2)
call find_variable("Case Nbr: ", case_number, 8)
case_number = replace(case_number, "_", "")
case_number = trim(case_number)

BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 15, 25, 50, 10, "Footer month:"
  EditBox 65, 20, 25, 15, footer_month
  Text 95, 25, 20, 10, "Year:"
  EditBox 120, 20, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

Dialog case_number_dialog
If buttonpressed = 0 then stopscript

call navigate_to_screen("elig", "hc")

'Script should allow for members other than 01, using the initial dialog.

row = 1
col = 1
EMSearch " 01 ", row, col 'This should be a variable based on the dialog's HH_memb request. Space in front and behind will prevent it from reading the date/time stamp.
If row = 0 then
  MsgBox "That HH member (" & HH_memb & ") is not found. The script will now stop."
  Stopscript
End if
EMWriteScreen "x", row, 29
transmit
EMWriteScreen "x", 18, 3
transmit
EMWriteScreen "X", 6, 3
transmit

BeginDialog spenddown_dialog, 0, 0, 186, 181, "Spenddown dialog"
  Text 5, 10, 30, 10, "SD type:"
  EditBox 40, 5, 15, 15, SD_type
  Text 65, 10, 30, 10, "Method:"
  EditBox 100, 5, 15, 15, SD_method
  Text 125, 10, 35, 10, "Cvrd pop:"
  EditBox 165, 5, 15, 15, cvrd_pop
  Text 15, 25, 30, 10, "MONTH"
  Text 65, 25, 50, 10, "ORIGINAL SD"
  Text 130, 25, 40, 10, "RECIP AMT"
  EditBox 10, 40, 40, 15, month_01
  EditBox 70, 40, 40, 15, original_SD_01
  EditBox 130, 40, 40, 15, recipient_amt_01
  EditBox 10, 60, 40, 15, month_02
  EditBox 70, 60, 40, 15, original_SD_02
  EditBox 130, 60, 40, 15, recipient_amt_02
  EditBox 10, 80, 40, 15, month_03
  EditBox 70, 80, 40, 15, original_SD_03
  EditBox 130, 80, 40, 15, recipient_amt_03
  EditBox 10, 100, 40, 15, month_04
  EditBox 70, 100, 40, 15, original_SD_04
  EditBox 130, 100, 40, 15, recipient_amt_04
  EditBox 10, 120, 40, 15, month_05
  EditBox 70, 120, 40, 15, original_SD_05
  EditBox 130, 120, 40, 15, recipient_amt_05
  EditBox 10, 140, 40, 15, month_06
  EditBox 70, 140, 40, 15, original_SD_06
  EditBox 130, 140, 40, 15, recipient_amt_06
  ButtonGroup ButtonPressed
    OkButton 40, 160, 50, 15
    CancelButton 95, 160, 50, 15
EndDialog

EMReadScreen SD_type, 1, 5, 14
EMReadScreen SD_method, 1, 5, 45
EMReadScreen cvrd_pop, 1, 5, 68
EMReadScreen month_01, 5, 7, 21
If month_01 <> "     " then
  EMReadScreen original_SD_01, 8, 8, 18
  original_SD_01 = trim(original_SD_01)
  EMReadScreen recipient_amt_01, 8, 11, 18
  recipient_amt_01 = trim(recipient_amt_01)
End if
EMReadScreen month_02, 5, 7, 32
If month_02 <> "     " then
  EMReadScreen original_SD_02, 8, 8, 29
  original_SD_02 = trim(original_SD_02)
  EMReadScreen recipient_amt_02, 8, 11, 29
  recipient_amt_02 = trim(recipient_amt_02)
End if
EMReadScreen month_03, 5, 7, 43
If month_03 <> "     " then
  EMReadScreen original_SD_03, 8, 8, 40
  original_SD_03 = trim(original_SD_03)
  EMReadScreen recipient_amt_03, 8, 11, 40
  recipient_amt_03 = trim(recipient_amt_03)
End if
EMReadScreen month_04, 5, 7, 54
If month_04 <> "     " then
  EMReadScreen original_SD_04, 8, 8, 51
  original_SD_04 = trim(original_SD_04)
  EMReadScreen recipient_amt_04, 8, 11, 51
  recipient_amt_04 = trim(recipient_amt_04)
End if
EMReadScreen month_05, 5, 7, 65
If month_05 <> "     " then
  EMReadScreen original_SD_05, 8, 8, 62
  original_SD_05 = trim(original_SD_05)
  EMReadScreen recipient_amt_05, 8, 11, 62
  recipient_amt_05 = trim(recipient_amt_05)
End if
EMReadScreen month_06, 5, 7, 76
If month_06 <> "     " then
  EMReadScreen original_SD_06, 8, 8, 73
  original_SD_06 = trim(original_SD_06)
  EMReadScreen recipient_amt_06, 8, 11, 73
  recipient_amt_06 = trim(recipient_amt_06)
End if

Dialog spenddown_dialog
If buttonpressed = 0 then stopscript

switch_to_MMIS_training
get_to_session_begin

EMSetCursor 1, 2
EMSendKey "mw00"
EMSendKey "<enter>"
EMWaitReady 1, 1
EMSendKey "<enter>"
EMWaitReady 1, 1

'This section may not work for all OSAs, since some only have EK01.
  row = 1
  col = 1
EMSearch "C302", row, col
If row <> 0 then 
  EMSetCursor row, 4
  EMSendKey "x"
  EMSendKey "<enter>"
  EMWaitReady 1, 1
End if
row = 1
col = 1
EMSearch "RECIPIENT FILE APPLICATION", row, col
If row = 0 then
  MsgBox "MMIS could not be found for this user. Contact the script administrator with your name, x102 number, and a description of the problem."
  StopScript
Else
  EMWriteScreen "x", row, col - 3
  EMSendKey "<enter>"
  EMWaitReady 1, 1
End if

EMWriteScreen "c", 2, 19

'MMIS_case_number = case_number
MMIS_case_number = "170655" 'for training/debug purposes only

Do
  If len(MMIS_case_number) < 8 then MMIS_case_number = "0" & MMIS_case_number
Loop until len(MMIS_case_number) = 8
EMWriteScreen MMIS_case_number, 9, 19
transmit
EMWriteScreen "rcin", 1, 8
transmit
EMWriteScreen "x", 11, 2 'It should determine who the other HH member is at this point.
transmit
EMWriteScreen "rspd", 1, 8
transmit
If month_01 <> "     " then
  PF11
  EMWriteScreen SD_type, 5, 9
  EMWriteScreen SD_method, 5, 17
  EMWriteScreen cvrd_pop, 5, 29
  month_01_start = replace(month_01, "/", "/01/")
  month_01_end = dateadd("m", 1, month_01_start)
  month_01_end = dateadd("d", -1, month_01_end)
  month_01_end = left(month_01_end, 6) & right(month_01_end, 2)
  EMWriteScreen month_01_start, 5, 47
  EMWriteScreen month_01_end, 5, 71
  EMWriteScreen cint(original_SD_01), 6, 41
  EMWriteScreen cint(recipient_amt_01), 6, 71
End if
If month_02 <> "     " then
  PF11
  EMWriteScreen SD_type, 5, 9
  EMWriteScreen SD_method, 5, 17
  EMWriteScreen cvrd_pop, 5, 29
  month_02_start = replace(month_02, "/", "/01/")
  month_02_end = dateadd("m", 1, month_02_start)
  month_02_end = dateadd("d", -1, month_02_end)
  month_02_end = left(month_02_end, 6) & right(month_02_end, 2)
  EMWriteScreen month_02_start, 5, 47
  EMWriteScreen month_02_end, 5, 71
  EMWriteScreen cint(original_SD_02), 6, 41
  EMWriteScreen cint(recipient_amt_02), 6, 71
End if
If month_03 <> "     " then
  PF11
  EMWriteScreen SD_type, 5, 9
  EMWriteScreen SD_method, 5, 17
  EMWriteScreen cvrd_pop, 5, 29
  month_03_start = replace(month_03, "/", "/01/")
  month_03_end = dateadd("m", 1, month_03_start)
  month_03_end = dateadd("d", -1, month_03_end)
  month_03_end = left(month_03_end, 6) & right(month_03_end, 2)
  EMWriteScreen month_03_start, 5, 47
  EMWriteScreen month_03_end, 5, 71
  EMWriteScreen cint(original_SD_03), 6, 41
  EMWriteScreen cint(recipient_amt_03), 6, 71
End if
If month_04 <> "     " then
  PF11
  EMWriteScreen SD_type, 5, 9
  EMWriteScreen SD_method, 5, 17
  EMWriteScreen cvrd_pop, 5, 29
  month_04_start = replace(month_04, "/", "/01/")
  month_04_end = dateadd("m", 1, month_04_start)
  month_04_end = dateadd("d", -1, month_04_end)
  month_04_end = left(month_04_end, 6) & right(month_04_end, 2)
  EMWriteScreen month_04_start, 5, 47
  EMWriteScreen month_04_end, 5, 71
  EMWriteScreen cint(original_SD_04), 6, 41
  EMWriteScreen cint(recipient_amt_04), 6, 71
End if
If month_05 <> "     " then
  PF11
  EMWriteScreen SD_type, 5, 9
  EMWriteScreen SD_method, 5, 17
  EMWriteScreen cvrd_pop, 5, 29
  month_05_start = replace(month_05, "/", "/01/")
  month_05_end = dateadd("m", 1, month_05_start)
  month_05_end = dateadd("d", -1, month_05_end)
  month_05_end = left(month_05_end, 6) & right(month_05_end, 2)
  EMWriteScreen month_05_start, 5, 47
  EMWriteScreen month_05_end, 5, 71
  EMWriteScreen cint(original_SD_05), 6, 41
  EMWriteScreen cint(recipient_amt_05), 6, 71
End if
If month_06 <> "     " then
  PF11
  EMWriteScreen SD_type, 5, 9
  EMWriteScreen SD_method, 5, 17
  EMWriteScreen cvrd_pop, 5, 29
  month_06_start = replace(month_06, "/", "/01/")
  month_06_end = dateadd("m", 1, month_06_start)
  month_06_end = dateadd("d", -1, month_06_end)
  month_06_end = left(month_06_end, 6) & right(month_06_end, 2)
  EMWriteScreen month_06_start, 5, 47
  EMWriteScreen month_06_end, 5, 71
  EMWriteScreen cint(original_SD_06), 6, 41
  EMWriteScreen cint(recipient_amt_06), 6, 71
End if