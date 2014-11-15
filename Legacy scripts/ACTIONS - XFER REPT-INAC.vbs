EMConnect "A"

'------The following checks to make sure you are in REPT/INAC.
EMReadScreen INAC_check, 44, 2, 21
If INAC_check <> "Workers Monthly Inactive Cases Report (INAC)" then msgbox "You are not on your REPT/INAC."
If INAC_check <> "Workers Monthly Inactive Cases Report (INAC)" then stopscript

'-------Next it checks the footer month. If it is the current footer month, the script stops

If DatePart("m", Now) = 1 then current_system_month = "01"
If DatePart("m", Now) = 2 then current_system_month = "02"
If DatePart("m", Now) = 3 then current_system_month = "03"
If DatePart("m", Now) = 4 then current_system_month = "04"
If DatePart("m", Now) = 5 then current_system_month = "05"
If DatePart("m", Now) = 6 then current_system_month = "06"
If DatePart("m", Now) = 7 then current_system_month = "07"
If DatePart("m", Now) = 8 then current_system_month = "08"
If DatePart("m", Now) = 9 then current_system_month = "09"
If DatePart("m", Now) = 10 then current_system_month = "10"
If DatePart("m", Now) = 11 then current_system_month = "11"
If DatePart("m", Now) = 12 then current_system_month = "12"




EMReadScreen footer_month, 2, 20, 54 
If current_system_month = footer_month then msgbox "Do not use this script in the current footer month. These cases need to be in your REPT/INAC for 30 days. The script will now stop."
If current_system_month = footer_month then stopscript








'-------------The following declares variables which will cross subs in this script.
Dim case_001_MMIS_end
Dim case_002_MMIS_end
Dim case_003_MMIS_end
Dim case_004_MMIS_end
Dim case_005_MMIS_end
Dim case_006_MMIS_end
Dim case_007_MMIS_end
Dim case_008_MMIS_end
Dim case_009_MMIS_end
Dim case_010_MMIS_end
Dim case_011_MMIS_end
Dim case_012_MMIS_end
Dim case_013_MMIS_end
Dim case_014_MMIS_end
Dim case_015_MMIS_end
Dim case_016_MMIS_end
Dim case_017_MMIS_end
Dim case_018_MMIS_end
Dim case_019_MMIS_end
Dim case_020_MMIS_end
Dim case_021_MMIS_end
Dim case_022_MMIS_end
Dim case_023_MMIS_end
Dim case_024_MMIS_end
Dim case_025_MMIS_end
Dim case_026_MMIS_end
Dim case_027_MMIS_end
Dim case_028_MMIS_end
Dim case_029_MMIS_end
Dim case_030_MMIS_end
Dim case_031_MMIS_end
Dim case_032_MMIS_end
Dim case_033_MMIS_end
Dim case_034_MMIS_end
Dim case_035_MMIS_end
Dim case_036_MMIS_end

'------Now a dialog pops up explaining what's about to happen, and asking for a signature.

BeginDialog start_dialog, 0, 0, 181, 97, "REPT/INAC closer"
  EditBox 45, 55, 80, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 30, 75, 50, 15
    CancelButton 95, 75, 50, 15
  Text 5, 5, 170, 50, "This script will check every case in your REPT/INAC, checking against MMIS for each case and CCOL/CLIC for claims. Be sure that you are in the correct footer month and are at the beginning of your REPT/INAC in production before proceeding. Write your name in the box below and click ''OK'' to begin."
EndDialog

Dialog start_dialog
If buttonpressed = 0 then stopscript

'------------------------Now it sends a PF7 to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<PF7>"
EMWaitReady 1, 0

'This Do...loop checks for the password prompt.
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then Dialog Dialog1
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"


'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMSendKey "<attn>"
EMWaitReady 1, 5
EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
EMReadScreen MMIS_A_check, 7, 15, 15
If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If training_check = "RUNNING" then EMSendKey "3" + "<enter>"
If production_check = "RUNNING" then EMSendKey "1" + "<enter>"



'The following saves the cases from the first three pages of rept/inac as variables. It also gets the inactive eff. dates.
EMWaitReady 1, 10
EMReadScreen case_001, 8, 7, 3
EMReadScreen case_002, 8, 8, 3
EMReadScreen case_003, 8, 9, 3
EMReadScreen case_004, 8, 10, 3
EMReadScreen case_005, 8, 11, 3
EMReadScreen case_006, 8, 12, 3
EMReadScreen case_007, 8, 13, 3
EMReadScreen case_008, 8, 14, 3
EMReadScreen case_009, 8, 15, 3
EMReadScreen case_010, 8, 16, 3
EMReadScreen case_011, 8, 17, 3
EMReadScreen case_012, 8, 18, 3

EMReadScreen case_001_inac_date, 8, 7, 49
EMReadScreen case_002_inac_date, 8, 8, 49
EMReadScreen case_003_inac_date, 8, 9, 49
EMReadScreen case_004_inac_date, 8, 10, 49
EMReadScreen case_005_inac_date, 8, 11, 49
EMReadScreen case_006_inac_date, 8, 12, 49
EMReadScreen case_007_inac_date, 8, 13, 49
EMReadScreen case_008_inac_date, 8, 14, 49
EMReadScreen case_009_inac_date, 8, 15, 49
EMReadScreen case_010_inac_date, 8, 16, 49
EMReadScreen case_011_inac_date, 8, 17, 49
EMReadScreen case_012_inac_date, 8, 18, 49

if isdate(case_001_inac_date) = false then case_001_inac_date = now
if isdate(case_002_inac_date) = false then case_002_inac_date = now
if isdate(case_003_inac_date) = false then case_003_inac_date = now
if isdate(case_004_inac_date) = false then case_004_inac_date = now
if isdate(case_005_inac_date) = false then case_005_inac_date = now
if isdate(case_006_inac_date) = false then case_006_inac_date = now
if isdate(case_007_inac_date) = false then case_007_inac_date = now
if isdate(case_008_inac_date) = false then case_008_inac_date = now
if isdate(case_009_inac_date) = false then case_009_inac_date = now
if isdate(case_010_inac_date) = false then case_010_inac_date = now
if isdate(case_011_inac_date) = false then case_011_inac_date = now
if isdate(case_012_inac_date) = false then case_012_inac_date = now




EMSendKey "<PF8>"
EMWaitReady 1, 10

EMReadScreen second_inac_duplicate_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
EMReadScreen case_013, 8, 7, 3
EMReadScreen case_014, 8, 8, 3
EMReadScreen case_015, 8, 9, 3
EMReadScreen case_016, 8, 10, 3
EMReadScreen case_017, 8, 11, 3
EMReadScreen case_018, 8, 12, 3
EMReadScreen case_019, 8, 13, 3
EMReadScreen case_020, 8, 14, 3
EMReadScreen case_021, 8, 15, 3
EMReadScreen case_022, 8, 16, 3
EMReadScreen case_023, 8, 17, 3
EMReadScreen case_024, 8, 18, 3

EMReadScreen case_013_inac_date, 8, 7, 49
EMReadScreen case_014_inac_date, 8, 8, 49
EMReadScreen case_015_inac_date, 8, 9, 49
EMReadScreen case_016_inac_date, 8, 10, 49
EMReadScreen case_017_inac_date, 8, 11, 49
EMReadScreen case_018_inac_date, 8, 12, 49
EMReadScreen case_019_inac_date, 8, 13, 49
EMReadScreen case_020_inac_date, 8, 14, 49
EMReadScreen case_021_inac_date, 8, 15, 49
EMReadScreen case_022_inac_date, 8, 16, 49
EMReadScreen case_023_inac_date, 8, 17, 49
EMReadScreen case_024_inac_date, 8, 18, 49

if isdate(case_013_inac_date) = false then case_013_inac_date = now
if isdate(case_014_inac_date) = false then case_014_inac_date = now
if isdate(case_015_inac_date) = false then case_015_inac_date = now
if isdate(case_016_inac_date) = false then case_016_inac_date = now
if isdate(case_017_inac_date) = false then case_017_inac_date = now
if isdate(case_018_inac_date) = false then case_018_inac_date = now
if isdate(case_019_inac_date) = false then case_019_inac_date = now
if isdate(case_020_inac_date) = false then case_020_inac_date = now
if isdate(case_021_inac_date) = false then case_021_inac_date = now
if isdate(case_022_inac_date) = false then case_022_inac_date = now
if isdate(case_023_inac_date) = false then case_023_inac_date = now
if isdate(case_024_inac_date) = false then case_024_inac_date = now

EMSendKey "<PF8>"
EMWaitReady 1, 10

EMReadScreen third_inac_duplicate_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
EMReadScreen case_025, 8, 7, 3
EMReadScreen case_026, 8, 8, 3
EMReadScreen case_027, 8, 9, 3
EMReadScreen case_028, 8, 10, 3
EMReadScreen case_029, 8, 11, 3
EMReadScreen case_030, 8, 12, 3
EMReadScreen case_031, 8, 13, 3
EMReadScreen case_032, 8, 14, 3
EMReadScreen case_033, 8, 15, 3
EMReadScreen case_034, 8, 16, 3
EMReadScreen case_035, 8, 17, 3
EMReadScreen case_036, 8, 18, 3

EMReadScreen case_025_inac_date, 8, 7, 49
EMReadScreen case_026_inac_date, 8, 8, 49
EMReadScreen case_027_inac_date, 8, 9, 49
EMReadScreen case_028_inac_date, 8, 10, 49
EMReadScreen case_029_inac_date, 8, 11, 49
EMReadScreen case_030_inac_date, 8, 12, 49
EMReadScreen case_031_inac_date, 8, 13, 49
EMReadScreen case_032_inac_date, 8, 14, 49
EMReadScreen case_033_inac_date, 8, 15, 49
EMReadScreen case_034_inac_date, 8, 16, 49
EMReadScreen case_035_inac_date, 8, 17, 49
EMReadScreen case_036_inac_date, 8, 18, 49

if isdate(case_025_inac_date) = false then case_025_inac_date = now
if isdate(case_026_inac_date) = false then case_026_inac_date = now
if isdate(case_027_inac_date) = false then case_027_inac_date = now
if isdate(case_028_inac_date) = false then case_028_inac_date = now
if isdate(case_029_inac_date) = false then case_029_inac_date = now
if isdate(case_030_inac_date) = false then case_030_inac_date = now
if isdate(case_031_inac_date) = false then case_031_inac_date = now
if isdate(case_032_inac_date) = false then case_032_inac_date = now
if isdate(case_033_inac_date) = false then case_033_inac_date = now
if isdate(case_034_inac_date) = false then case_034_inac_date = now
if isdate(case_035_inac_date) = false then case_035_inac_date = now
if isdate(case_036_inac_date) = false then case_036_inac_date = now

'Now it will look for MMIS. If MMIS is running, it will start checking each case against MMIS.

'The following checks for which screen MMIS is running on.
IF MMIS_A_check = "RUNNING" then EMSendKey "<attn>" 
IF MMIS_A_check = "RUNNING" then EMWaitReady 1, 5
IF MMIS_A_check = "RUNNING" then EMSendKey "10" + "<enter>"
IF MMIS_A_check = "RUNNING" then EMWaitReady 1, 0
IF MMIS_A_check <> "RUNNING" then EMConnect "B"
EMWaitReady 1, 5
IF MMIS_A_check <> "RUNNING" then EMSendKey "<attn>"
EMWaitReady 1, 5
IF MMIS_A_check <> "RUNNING" then EMReadScreen MMIS_B_check, 7, 15, 15
If MMIS_A_check <> "RUNNING" and MMIS_B_check <> "RUNNING" then MsgBox "MMIS does not appear to be running. This script will now stop."
If MMIS_A_check <> "RUNNING" and MMIS_B_check <> "RUNNING" then stopscript
IF MMIS_A_check <> "RUNNING" and MMIS_B_check = "RUNNING" then EMSendkey "10" + "<enter>"

EMFocus

  Sub get_to_session_begin 'This sub uses a Do Loop to get to the start screen for MMIS.
    Do 
    EMSendkey "<PF6>"
      EMReadScreen password_prompt2, 38, 2, 23
      IF password_prompt2 = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
    EMWaitReady 1, 0
    EMReadScreen session_start, 18, 1, 7
    Loop until session_start = "SESSION TERMINATED"
  End Sub

get_to_session_begin
EMSetCursor 1, 2
EMSendKey "mw00"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "<enter>"
EMWaitReady 1, 0

'This section may not work for all OSAs, since some only have EK01.
  row = 1
  col = 1
EMSearch "EK01", row, col
If row <> 0 then EMSetCursor row, 4
If row <> 0 then EMSendKey "x"
If row <> 0 then EMSendKey "<enter>"
If row <> 0 then EMWaitReady 1, 0
If row = 0 then
  Msgbox "EK01 (MCRE) does not appear to be installed for this user. The script will now stop. Contact the script administrator to have this issue resolved."
  StopScript
End if

'This section starts from EK01. OSAs may need to skip the previous section.
EMSetCursor 10, 3
EMSendKey "x"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMFocus

'Now we are in MMIS, and it will start a sub to get info on the first case number.

Sub case_001_MMIS_check
  EMSetCursor 2, 19
  EMSendKey "i"
  EMSetCursor 9, 19
  EMSendKey case_001

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
  EMReadscreen first_MMIS_number_position, 1, 9, 19
  EMSetCursor 9, 19
  If first_MMIS_number_position = " " then EMSendKey "0"
  EMReadscreen second_MMIS_number_position, 1, 9, 20
  EMSetCursor 9, 20
  If second_MMIS_number_position = " " then EMSendKey "0"
  EMReadscreen third_MMIS_number_position, 1, 9, 21
  EMSetCursor 9, 21
  If third_MMIS_number_position = " " then EMSendKey "0"
  EMReadscreen fourth_MMIS_number_position, 1, 9, 22
  EMSetCursor 9, 22
  If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
  EMSendKey "<enter>"
  EMWaitReady 1, 0
  EMSendKey "rcin" + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 11, 2
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "relg" + "<enter>"
  EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. It stores an additional variable indicating that the case should not be XFERed. Then it returns to RKEY.
  EMReadScreen case_001_MMIS_end, 8, 7, 36
  EMSendKey "<PF6>"
  EMWaitReady 1, 0
  EMSendKey "<PF6>"
  EMWaitReady 1, 0

End sub



'---Now the sub for case 002.

Sub case_002_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_002

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_002_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End Sub

'---Now it does it again with case 003.

Sub case_003_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_003

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_003_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End Sub

'---Now it does it again with case 004.

Sub case_004_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_004

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_004_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 005.

Sub case_005_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_005

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_005_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 006.

Sub case_006_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_006

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_006_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 007.

Sub case_007_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_007

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_007_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End Sub

'---Now it does it again with case 008.

Sub case_008_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_008

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_008_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 009.

Sub case_009_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_009

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_009_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 010.

Sub case_010_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_010

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_010_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 011

Sub case_011_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_011

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. It stores an additional variable indicating that the case should not be XFERed. Then it returns to RKEY.

EMReadScreen case_011_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 012.

Sub case_012_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_012

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_012_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 013.

Sub case_013_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_013

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_013_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 014.

Sub case_014_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_014

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_014_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 015.

Sub case_015_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_015

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_015_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 016.

Sub case_016_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_016

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_016_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0


End sub

'---Now it does it again with case 017.

Sub case_017_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_017

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_017_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 018.

Sub case_018_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_018

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_018_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 019.

Sub case_019_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_019

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_019_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 020.

Sub case_020_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_020

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_020_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 021

Sub case_021_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_021

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. It stores an additional variable indicating that the case should not be XFERed. Then it returns to RKEY.

EMReadScreen case_021_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 022.

Sub case_022_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_022

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_022_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 023.

Sub case_023_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_023

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_023_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 024.

Sub case_024_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_024

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_024_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 025.

Sub case_025_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_025

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_025_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 026.

Sub case_026_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_026

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_026_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub


'---Now it does it again with case 027.

Sub case_027_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_027

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_027_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 028.

Sub case_028_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_028

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_028_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 029.

Sub case_029_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_029

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_029_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 030.

Sub case_030_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_030

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_030_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 031

Sub case_031_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_031

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. It stores an additional variable indicating that the case should not be XFERed. Then it returns to RKEY.

EMReadScreen case_031_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 032.

Sub case_032_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_032

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_032_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 033.


Sub case_033_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_033

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_033_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 034.


Sub case_034_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_034

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_034_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 035.

Sub case_035_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_035

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_035_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'---Now it does it again with case 036.

Sub case_036_MMIS_check

EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_036

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now it gets to RELG for this case.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads the case to determine if MMIS is active. Then it returns to RKEY.

EMReadScreen case_036_MMIS_end, 8, 7, 36
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

End sub

'--------The following subs hits the subs for each line, checking each to make sure there is a case number. If duplicate pages were detected it will bypass them.
sub first_page_MMIS_check
  If case_001 <> "        " then call case_001_MMIS_check
  If case_002 <> "        " then call case_002_MMIS_check
  If case_003 <> "        " then call case_003_MMIS_check
  If case_004 <> "        " then call case_004_MMIS_check
  If case_005 <> "        " then call case_005_MMIS_check
  If case_006 <> "        " then call case_006_MMIS_check
  If case_007 <> "        " then call case_007_MMIS_check
  If case_008 <> "        " then call case_008_MMIS_check
  If case_009 <> "        " then call case_009_MMIS_check
  If case_010 <> "        " then call case_010_MMIS_check
  If case_011 <> "        " then call case_011_MMIS_check
  If case_012 <> "        " then call case_012_MMIS_check
End sub

sub second_page_MMIS_check
  If case_013 <> "        " then call case_013_MMIS_check
  If case_014 <> "        " then call case_014_MMIS_check
  If case_015 <> "        " then call case_015_MMIS_check
  If case_016 <> "        " then call case_016_MMIS_check
  If case_017 <> "        " then call case_017_MMIS_check
  If case_018 <> "        " then call case_018_MMIS_check
  If case_019 <> "        " then call case_019_MMIS_check
  If case_020 <> "        " then call case_020_MMIS_check
  If case_021 <> "        " then call case_021_MMIS_check
  If case_022 <> "        " then call case_022_MMIS_check
  If case_023 <> "        " then call case_023_MMIS_check
  If case_024 <> "        " then call case_024_MMIS_check
End sub

sub third_page_MMIS_check
  If case_025 <> "        " then call case_025_MMIS_check
  If case_026 <> "        " then call case_026_MMIS_check
  If case_027 <> "        " then call case_027_MMIS_check
  If case_028 <> "        " then call case_028_MMIS_check
  If case_029 <> "        " then call case_029_MMIS_check
  If case_030 <> "        " then call case_030_MMIS_check
  If case_031 <> "        " then call case_031_MMIS_check
  If case_032 <> "        " then call case_032_MMIS_check
  If case_033 <> "        " then call case_033_MMIS_check
  If case_034 <> "        " then call case_034_MMIS_check
  If case_035 <> "        " then call case_035_MMIS_check
  If case_036 <> "        " then call case_036_MMIS_check
end sub

first_page_MMIS_check
If second_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call second_page_MMIS_check
If third_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call third_page_MMIS_check




'---------------------------------------------------------Now it heads back into MAXIS

EMConnect "A"
EMSendKey "<attn>"
EMWaitReady 1, 5
EMSendKey "<attn>"
EMWaitReady 1, 0

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"

EMSetCursor 16, 43
EMSendKey "ccol"
EMSetCursor 21, 70
EMSendKey "clic" + "<enter>"
EMWaitReady 1, 0

'---------------------Now it checks CCOL/CLIC for claims, and makes a word doc with their contents.

MsgBox "Now we are checking against claims. A word doc will be created, with the names and case numbers of cases with active claims. It may take a few seconds to start Word. Do not use your computer until it loads up. Press OK to continue."

'Now it creates a word document with all active claims in it.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
   Set objSelection = objWord.Selection
   objselection.typetext "Case numbers with active claims: "
   objselection.TypeParagraph()
   objselection.TypeParagraph()

'The following are the subs which will check cases against CCOL/CLIC.

Sub case_001_CCOL_check
  EMSendKey "<eraseeof>" + case_001 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_001_claims, 9, 24, 02
  EMReadScreen case_001_name_and_number, 37, 4, 8
  If case_001_claims <> "NO CLAIMS" then objselection.typetext case_001_name_and_number
  If case_001_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_002_CCOL_check
  EMSendKey "<eraseeof>" + case_002 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_002_claims, 9, 24, 02
  EMReadScreen case_002_name_and_number, 37, 4, 8
  If case_002_claims <> "NO CLAIMS" then objselection.typetext case_002_name_and_number
  If case_002_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_003_CCOL_check
  EMSendKey "<eraseeof>" + case_003 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_003_claims, 9, 24, 02
  EMReadScreen case_003_name_and_number, 37, 4, 8
  If case_003_claims <> "NO CLAIMS" then objselection.typetext case_003_name_and_number
  If case_003_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_004_CCOL_check
  EMSendKey "<eraseeof>" + case_004 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_004_claims, 9, 24, 02
  EMReadScreen case_004_name_and_number, 37, 4, 8
  If case_004_claims <> "NO CLAIMS" then objselection.typetext case_004_name_and_number
  If case_004_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_005_CCOL_check
  EMSendKey "<eraseeof>" + case_005 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_005_claims, 9, 24, 02
  EMReadScreen case_005_name_and_number, 37, 4, 8
  If case_005_claims <> "NO CLAIMS" then objselection.typetext case_005_name_and_number
  If case_005_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_006_CCOL_check
  EMSendKey "<eraseeof>" + case_006 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_006_claims, 9, 24, 02
  EMReadScreen case_006_name_and_number, 37, 4, 8
  If case_006_claims <> "NO CLAIMS" then objselection.typetext case_006_name_and_number
  If case_006_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_007_CCOL_check
  EMSendKey "<eraseeof>" + case_007 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_007_claims, 9, 24, 02
  EMReadScreen case_007_name_and_number, 37, 4, 8
  If case_007_claims <> "NO CLAIMS" then objselection.typetext case_007_name_and_number
  If case_007_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_008_CCOL_check
  EMSendKey "<eraseeof>" + case_008 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_008_claims, 9, 24, 02
  EMReadScreen case_008_name_and_number, 37, 4, 8
  If case_008_claims <> "NO CLAIMS" then objselection.typetext case_008_name_and_number
  If case_008_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_009_CCOL_check
  EMSendKey "<eraseeof>" + case_009 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_009_claims, 9, 24, 02
  EMReadScreen case_009_name_and_number, 37, 4, 8
  If case_009_claims <> "NO CLAIMS" then objselection.typetext case_009_name_and_number
  If case_009_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_010_CCOL_check
  EMSendKey "<eraseeof>" + case_010 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_010_claims, 9, 24, 02
  EMReadScreen case_010_name_and_number, 37, 4, 8
  If case_010_claims <> "NO CLAIMS" then objselection.typetext case_010_name_and_number
  If case_010_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_011_CCOL_check
  EMSendKey "<eraseeof>" + case_011 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_011_claims, 9, 24, 02
  EMReadScreen case_011_name_and_number, 37, 4, 8
  If case_011_claims <> "NO CLAIMS" then objselection.typetext case_011_name_and_number
  If case_011_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_012_CCOL_check
  EMSendKey "<eraseeof>" + case_012 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_012_claims, 9, 24, 02
  EMReadScreen case_012_name_and_number, 37, 4, 8
  If case_012_claims <> "NO CLAIMS" then objselection.typetext case_012_name_and_number
  If case_012_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_013_CCOL_check
  EMSendKey "<eraseeof>" + case_013 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_013_claims, 9, 24, 02
  EMReadScreen case_013_name_and_number, 37, 4, 8
  If case_013_claims <> "NO CLAIMS" then objselection.typetext case_013_name_and_number
  If case_013_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_014_CCOL_check
  EMSendKey "<eraseeof>" + case_014 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_014_claims, 9, 24, 02
  EMReadScreen case_014_name_and_number, 37, 4, 8
  If case_014_claims <> "NO CLAIMS" then objselection.typetext case_014_name_and_number
  If case_014_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_015_CCOL_check
  EMSendKey "<eraseeof>" + case_015 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_015_claims, 9, 24, 02
  EMReadScreen case_015_name_and_number, 37, 4, 8
  If case_015_claims <> "NO CLAIMS" then objselection.typetext case_015_name_and_number
  If case_015_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_016_CCOL_check
  EMSendKey "<eraseeof>" + case_016 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_016_claims, 9, 24, 02
  EMReadScreen case_016_name_and_number, 37, 4, 8
  If case_016_claims <> "NO CLAIMS" then objselection.typetext case_016_name_and_number
  If case_016_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_017_CCOL_check
  EMSendKey "<eraseeof>" + case_017 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_017_claims, 9, 24, 02
  EMReadScreen case_017_name_and_number, 37, 4, 8
  If case_017_claims <> "NO CLAIMS" then objselection.typetext case_017_name_and_number
  If case_017_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_018_CCOL_check
  EMSendKey "<eraseeof>" + case_018 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_018_claims, 9, 24, 02
  EMReadScreen case_018_name_and_number, 37, 4, 8
  If case_018_claims <> "NO CLAIMS" then objselection.typetext case_018_name_and_number
  If case_018_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_019_CCOL_check
  EMSendKey "<eraseeof>" + case_019 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_019_claims, 9, 24, 02
  EMReadScreen case_019_name_and_number, 37, 4, 8
  If case_019_claims <> "NO CLAIMS" then objselection.typetext case_019_name_and_number
  If case_019_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_020_CCOL_check
  EMSendKey "<eraseeof>" + case_020 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_020_claims, 9, 24, 02
  EMReadScreen case_020_name_and_number, 37, 4, 8
  If case_020_claims <> "NO CLAIMS" then objselection.typetext case_020_name_and_number
  If case_020_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_021_CCOL_check
  EMSendKey "<eraseeof>" + case_021 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_021_claims, 9, 24, 02
  EMReadScreen case_021_name_and_number, 37, 4, 8
  If case_021_claims <> "NO CLAIMS" then objselection.typetext case_021_name_and_number
  If case_021_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_022_CCOL_check
  EMSendKey "<eraseeof>" + case_022 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_022_claims, 9, 24, 02
  EMReadScreen case_022_name_and_number, 37, 4, 8
  If case_022_claims <> "NO CLAIMS" then objselection.typetext case_022_name_and_number
  If case_022_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_023_CCOL_check
  EMSendKey "<eraseeof>" + case_023 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_023_claims, 9, 24, 02
  EMReadScreen case_023_name_and_number, 37, 4, 8
  If case_023_claims <> "NO CLAIMS" then objselection.typetext case_023_name_and_number
  If case_023_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_024_CCOL_check
  EMSendKey "<eraseeof>" + case_024 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_024_claims, 9, 24, 02
  EMReadScreen case_024_name_and_number, 37, 4, 8
  If case_024_claims <> "NO CLAIMS" then objselection.typetext case_024_name_and_number
  If case_024_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_025_CCOL_check
  EMSendKey "<eraseeof>" + case_025 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_025_claims, 9, 24, 02
  EMReadScreen case_025_name_and_number, 37, 4, 8
  If case_025_claims <> "NO CLAIMS" then objselection.typetext case_025_name_and_number
  If case_025_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_026_CCOL_check
  EMSendKey "<eraseeof>" + case_026 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_026_claims, 9, 24, 02
  EMReadScreen case_026_name_and_number, 37, 4, 8
  If case_026_claims <> "NO CLAIMS" then objselection.typetext case_026_name_and_number
  If case_026_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_027_CCOL_check
  EMSendKey "<eraseeof>" + case_027 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_027_claims, 9, 24, 02
  EMReadScreen case_027_name_and_number, 37, 4, 8
  If case_027_claims <> "NO CLAIMS" then objselection.typetext case_027_name_and_number
  If case_027_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_028_CCOL_check
  EMSendKey "<eraseeof>" + case_028 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_028_claims, 9, 24, 02
  EMReadScreen case_028_name_and_number, 37, 4, 8
  If case_028_claims <> "NO CLAIMS" then objselection.typetext case_028_name_and_number
  If case_028_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_029_CCOL_check
  EMSendKey "<eraseeof>" + case_029 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_029_claims, 9, 24, 02
  EMReadScreen case_029_name_and_number, 37, 4, 8
  If case_029_claims <> "NO CLAIMS" then objselection.typetext case_029_name_and_number
  If case_029_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_030_CCOL_check
  EMSendKey "<eraseeof>" + case_030 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_030_claims, 9, 24, 02
  EMReadScreen case_030_name_and_number, 37, 4, 8
  If case_030_claims <> "NO CLAIMS" then objselection.typetext case_030_name_and_number
  If case_030_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_031_CCOL_check
  EMSendKey "<eraseeof>" + case_031 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_031_claims, 9, 24, 02
  EMReadScreen case_031_name_and_number, 37, 4, 8
  If case_031_claims <> "NO CLAIMS" then objselection.typetext case_031_name_and_number
  If case_031_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_032_CCOL_check
  EMSendKey "<eraseeof>" + case_032 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_032_claims, 9, 24, 02
  EMReadScreen case_032_name_and_number, 37, 4, 8
  If case_032_claims <> "NO CLAIMS" then objselection.typetext case_032_name_and_number
  If case_032_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_033_CCOL_check
  EMSendKey "<eraseeof>" + case_033 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_033_claims, 9, 24, 02
  EMReadScreen case_033_name_and_number, 37, 4, 8
  If case_033_claims <> "NO CLAIMS" then objselection.typetext case_033_name_and_number
  If case_033_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_034_CCOL_check
  EMSendKey "<eraseeof>" + case_034 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_034_claims, 9, 24, 02
  EMReadScreen case_034_name_and_number, 37, 4, 8
  If case_034_claims <> "NO CLAIMS" then objselection.typetext case_034_name_and_number
  If case_034_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_035_CCOL_check
  EMSendKey "<eraseeof>" + case_035 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_035_claims, 9, 24, 02
  EMReadScreen case_035_name_and_number, 37, 4, 8
  If case_035_claims <> "NO CLAIMS" then objselection.typetext case_035_name_and_number
  If case_035_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub case_036_CCOL_check
  EMSendKey "<eraseeof>" + case_036 + "<enter>"
  EMWaitReady 1, 0
  EMReadScreen case_036_claims, 9, 24, 02
  EMReadScreen case_036_name_and_number, 37, 4, 8
  If case_036_claims <> "NO CLAIMS" then objselection.typetext case_036_name_and_number
  If case_036_claims <> "NO CLAIMS" then objselection.TypeParagraph()
  EMSetCursor 4, 8
End Sub

Sub first_page_CCOL_check
  If case_001 <> "        " then call case_001_CCOL_check
  If case_002 <> "        " then call case_002_CCOL_check
  If case_003 <> "        " then call case_003_CCOL_check
  If case_004 <> "        " then call case_004_CCOL_check
  If case_005 <> "        " then call case_005_CCOL_check
  If case_006 <> "        " then call case_006_CCOL_check
  If case_007 <> "        " then call case_007_CCOL_check
  If case_008 <> "        " then call case_008_CCOL_check
  If case_009 <> "        " then call case_009_CCOL_check
  If case_010 <> "        " then call case_010_CCOL_check
  If case_011 <> "        " then call case_011_CCOL_check
  If case_012 <> "        " then call case_012_CCOL_check
End sub

Sub second_page_CCOL_check
  If case_013 <> "        " then call case_013_CCOL_check
  If case_014 <> "        " then call case_014_CCOL_check
  If case_015 <> "        " then call case_015_CCOL_check
  If case_016 <> "        " then call case_016_CCOL_check
  If case_017 <> "        " then call case_017_CCOL_check
  If case_018 <> "        " then call case_018_CCOL_check
  If case_019 <> "        " then call case_019_CCOL_check
  If case_020 <> "        " then call case_020_CCOL_check
  If case_021 <> "        " then call case_021_CCOL_check
  If case_022 <> "        " then call case_022_CCOL_check
  If case_023 <> "        " then call case_023_CCOL_check
  If case_024 <> "        " then call case_024_CCOL_check
End sub

Sub third_page_CCOL_check
  If case_025 <> "        " then call case_025_CCOL_check
  If case_026 <> "        " then call case_026_CCOL_check
  If case_027 <> "        " then call case_027_CCOL_check
  If case_028 <> "        " then call case_028_CCOL_check
  If case_029 <> "        " then call case_029_CCOL_check
  If case_030 <> "        " then call case_030_CCOL_check
  If case_031 <> "        " then call case_031_CCOL_check
  If case_032 <> "        " then call case_032_CCOL_check
  If case_033 <> "        " then call case_033_CCOL_check
  If case_034 <> "        " then call case_034_CCOL_check
  If case_035 <> "        " then call case_035_CCOL_check
  If case_036 <> "        " then call case_036_CCOL_check
End sub

first_page_CCOL_check
If second_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call second_page_CCOL_check
If third_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call third_page_CCOL_check

'-----------------------Now we will be case noting the closed cases.

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"


'This gets into CASE for the first case number, but does not enter a case note. Subs will be used to enter the case notes.
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 18, 43
EMSendKey case_001 + "<enter>"
EMWaitReady 1, 0

'The following are subs to case note all inactive cases, that aren't active in MMIS.

Sub case_001_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_001
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_002_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_002
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_003_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_003
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_004_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_004
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_005_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_005
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_006_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_006
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_007_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_007
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_008_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_008
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_009_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_009
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_010_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_010
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_011_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_011
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_012_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_012
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_013_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_013
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_014_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_014
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_015_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_015
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_016_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_016
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_017_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_017
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_018_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_018
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_019_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_019
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_020_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_020
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_021_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_021
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_022_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_022
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_023_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_023
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_024_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_024
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_025_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_025
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_026_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_026
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_027_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_027
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_028_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_028
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_029_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_029
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_030_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_030
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_031_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_031
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_032_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_032
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_033_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_033
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_034_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_034
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_035_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_035
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub

Sub case_036_case_note
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_036
  EMSetCursor 20, 70
  EMSendKey "note" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSendKey "Closed case sent to CLS using automated script. -" + worker_sig
  EMSendKey "<PF3>"
  EMWaitReady 1, 0
End sub
Sub first_page_case_note
  IF case_001_MMIS_end <> "99/99/99" and case_001 <> "        " and DateDiff("d", case_001_inac_date, Now) > 10 then call case_001_case_note
  IF case_002_MMIS_end <> "99/99/99" and case_002 <> "        " and DateDiff("d", case_002_inac_date, Now) > 10 then call case_002_case_note
  IF case_003_MMIS_end <> "99/99/99" and case_003 <> "        " and DateDiff("d", case_003_inac_date, Now) > 10 then call case_003_case_note
  IF case_004_MMIS_end <> "99/99/99" and case_004 <> "        " and DateDiff("d", case_004_inac_date, Now) > 10 then call case_004_case_note
  IF case_005_MMIS_end <> "99/99/99" and case_005 <> "        " and DateDiff("d", case_005_inac_date, Now) > 10 then call case_005_case_note
  IF case_006_MMIS_end <> "99/99/99" and case_006 <> "        " and DateDiff("d", case_006_inac_date, Now) > 10 then call case_006_case_note
  IF case_007_MMIS_end <> "99/99/99" and case_007 <> "        " and DateDiff("d", case_007_inac_date, Now) > 10 then call case_007_case_note
  IF case_008_MMIS_end <> "99/99/99" and case_008 <> "        " and DateDiff("d", case_008_inac_date, Now) > 10 then call case_008_case_note
  IF case_009_MMIS_end <> "99/99/99" and case_009 <> "        " and DateDiff("d", case_009_inac_date, Now) > 10 then call case_009_case_note
  IF case_010_MMIS_end <> "99/99/99" and case_010 <> "        " and DateDiff("d", case_010_inac_date, Now) > 10 then call case_010_case_note
  IF case_011_MMIS_end <> "99/99/99" and case_011 <> "        " and DateDiff("d", case_011_inac_date, Now) > 10 then call case_011_case_note
  IF case_012_MMIS_end <> "99/99/99" and case_012 <> "        " and DateDiff("d", case_012_inac_date, Now) > 10 then call case_012_case_note
End sub

Sub second_page_case_note
  IF case_013_MMIS_end <> "99/99/99" and case_013 <> "        " and DateDiff("d", case_013_inac_date, Now) > 10 then call case_013_case_note
  IF case_014_MMIS_end <> "99/99/99" and case_014 <> "        " and DateDiff("d", case_014_inac_date, Now) > 10 then call case_014_case_note
  IF case_015_MMIS_end <> "99/99/99" and case_015 <> "        " and DateDiff("d", case_015_inac_date, Now) > 10 then call case_015_case_note
  IF case_016_MMIS_end <> "99/99/99" and case_016 <> "        " and DateDiff("d", case_016_inac_date, Now) > 10 then call case_016_case_note
  IF case_017_MMIS_end <> "99/99/99" and case_017 <> "        " and DateDiff("d", case_017_inac_date, Now) > 10 then call case_017_case_note
  IF case_018_MMIS_end <> "99/99/99" and case_018 <> "        " and DateDiff("d", case_018_inac_date, Now) > 10 then call case_018_case_note
  IF case_019_MMIS_end <> "99/99/99" and case_019 <> "        " and DateDiff("d", case_019_inac_date, Now) > 10 then call case_019_case_note
  IF case_020_MMIS_end <> "99/99/99" and case_020 <> "        " and DateDiff("d", case_020_inac_date, Now) > 10 then call case_020_case_note
  IF case_021_MMIS_end <> "99/99/99" and case_021 <> "        " and DateDiff("d", case_021_inac_date, Now) > 10 then call case_021_case_note
  IF case_022_MMIS_end <> "99/99/99" and case_022 <> "        " and DateDiff("d", case_022_inac_date, Now) > 10 then call case_022_case_note
  IF case_023_MMIS_end <> "99/99/99" and case_023 <> "        " and DateDiff("d", case_023_inac_date, Now) > 10 then call case_023_case_note
  IF case_024_MMIS_end <> "99/99/99" and case_024 <> "        " and DateDiff("d", case_024_inac_date, Now) > 10 then call case_024_case_note
End sub

Sub third_page_case_note
  IF case_025_MMIS_end <> "99/99/99" and case_025 <> "        " and DateDiff("d", case_025_inac_date, Now) > 10 then call case_025_case_note
  IF case_026_MMIS_end <> "99/99/99" and case_026 <> "        " and DateDiff("d", case_026_inac_date, Now) > 10 then call case_026_case_note
  IF case_027_MMIS_end <> "99/99/99" and case_027 <> "        " and DateDiff("d", case_027_inac_date, Now) > 10 then call case_027_case_note
  IF case_028_MMIS_end <> "99/99/99" and case_028 <> "        " and DateDiff("d", case_028_inac_date, Now) > 10 then call case_028_case_note
  IF case_029_MMIS_end <> "99/99/99" and case_029 <> "        " and DateDiff("d", case_029_inac_date, Now) > 10 then call case_029_case_note
  IF case_030_MMIS_end <> "99/99/99" and case_030 <> "        " and DateDiff("d", case_030_inac_date, Now) > 10 then call case_030_case_note
  IF case_031_MMIS_end <> "99/99/99" and case_031 <> "        " and DateDiff("d", case_031_inac_date, Now) > 10 then call case_031_case_note
  IF case_032_MMIS_end <> "99/99/99" and case_032 <> "        " and DateDiff("d", case_032_inac_date, Now) > 10 then call case_032_case_note
  IF case_033_MMIS_end <> "99/99/99" and case_033 <> "        " and DateDiff("d", case_033_inac_date, Now) > 10 then call case_033_case_note
  IF case_034_MMIS_end <> "99/99/99" and case_034 <> "        " and DateDiff("d", case_034_inac_date, Now) > 10 then call case_034_case_note
  IF case_035_MMIS_end <> "99/99/99" and case_035 <> "        " and DateDiff("d", case_035_inac_date, Now) > 10 then call case_035_case_note
  IF case_036_MMIS_end <> "99/99/99" and case_036 <> "        " and DateDiff("d", case_036_inac_date, Now) > 10 then call case_036_case_note
end sub

first_page_case_note
If second_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call second_page_case_note
If third_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call third_page_case_note






'-------------------Now it will SPEC/XFER the cases that aren't open in MMIS to CLS.

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"

EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "spec"
EMSetCursor 18, 43
EMSendKey "<eraseeof>" + case_001 'This will use case 1 even though we may not be XFERing it, as a case number is required for spec/xfer
EMSetCursor 21, 70
EMSendKey "xfer" + "<enter>"
EMWaitReady 1, 0

'The following are the subs to XFER each case.

Sub case_001_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_001 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_002_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_002 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_003_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_003 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_004_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_004 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_005_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_005 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_006_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_006 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_007_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_007 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_008_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_008 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_009_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_009 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_010_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_010 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_011_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_011 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_012_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_012 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_013_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_013 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_014_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_014 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_015_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_015 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_016_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_016 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_017_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_017 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_018_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_018 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_019_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_019 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_020_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_020 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_021_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_021 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_022_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_022 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_023_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_023 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_024_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_024 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_025_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_025 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_026_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_026 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_027_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_027 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_028_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_028 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_029_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_029 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_030_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_030 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_031_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_031 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_032_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_032 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_033_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_033 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_034_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_034 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_035_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_035 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub

Sub case_036_spec_xfer
  EMSetCursor 20, 38
  EMSendKey "<eraseeof>" + case_036 + "<enter>"
  EMWaitReady 1, 0
  EMSetCursor 7, 16
  EMSendKey "x" + "<enter>"
  EMWaitReady 1, 0
  EMSendKey "<PF9>"
  EMWaitReady 1, 0
  EMSetCursor 18, 61
  EMSendKey "x102cls" + "<enter>"
  EMWaitReady 1 , 0
End Sub


Sub first_page_spec_xfer
  IF case_001_MMIS_end <> "99/99/99" and case_001 <> "        " and DateDiff("d", case_001_inac_date, Now) > 10 then call case_001_spec_xfer
  IF case_002_MMIS_end <> "99/99/99" and case_002 <> "        " and DateDiff("d", case_002_inac_date, Now) > 10 then call case_002_spec_xfer
  IF case_003_MMIS_end <> "99/99/99" and case_003 <> "        " and DateDiff("d", case_003_inac_date, Now) > 10 then call case_003_spec_xfer
  IF case_004_MMIS_end <> "99/99/99" and case_004 <> "        " and DateDiff("d", case_004_inac_date, Now) > 10 then call case_004_spec_xfer
  IF case_005_MMIS_end <> "99/99/99" and case_005 <> "        " and DateDiff("d", case_005_inac_date, Now) > 10 then call case_005_spec_xfer
  IF case_006_MMIS_end <> "99/99/99" and case_006 <> "        " and DateDiff("d", case_006_inac_date, Now) > 10 then call case_006_spec_xfer
  IF case_007_MMIS_end <> "99/99/99" and case_007 <> "        " and DateDiff("d", case_007_inac_date, Now) > 10 then call case_007_spec_xfer
  IF case_008_MMIS_end <> "99/99/99" and case_008 <> "        " and DateDiff("d", case_008_inac_date, Now) > 10 then call case_008_spec_xfer
  IF case_009_MMIS_end <> "99/99/99" and case_009 <> "        " and DateDiff("d", case_009_inac_date, Now) > 10 then call case_009_spec_xfer
  IF case_010_MMIS_end <> "99/99/99" and case_010 <> "        " and DateDiff("d", case_010_inac_date, Now) > 10 then call case_010_spec_xfer
  IF case_011_MMIS_end <> "99/99/99" and case_011 <> "        " and DateDiff("d", case_011_inac_date, Now) > 10 then call case_011_spec_xfer
  IF case_012_MMIS_end <> "99/99/99" and case_012 <> "        " and DateDiff("d", case_012_inac_date, Now) > 10 then call case_012_spec_xfer
End sub

Sub second_page_spec_xfer
  IF case_013_MMIS_end <> "99/99/99" and case_013 <> "        " and DateDiff("d", case_013_inac_date, Now) > 10 then call case_013_spec_xfer
  IF case_014_MMIS_end <> "99/99/99" and case_014 <> "        " and DateDiff("d", case_014_inac_date, Now) > 10 then call case_014_spec_xfer
  IF case_015_MMIS_end <> "99/99/99" and case_015 <> "        " and DateDiff("d", case_015_inac_date, Now) > 10 then call case_015_spec_xfer
  IF case_016_MMIS_end <> "99/99/99" and case_016 <> "        " and DateDiff("d", case_016_inac_date, Now) > 10 then call case_016_spec_xfer
  IF case_017_MMIS_end <> "99/99/99" and case_017 <> "        " and DateDiff("d", case_017_inac_date, Now) > 10 then call case_017_spec_xfer
  IF case_018_MMIS_end <> "99/99/99" and case_018 <> "        " and DateDiff("d", case_018_inac_date, Now) > 10 then call case_018_spec_xfer
  IF case_019_MMIS_end <> "99/99/99" and case_019 <> "        " and DateDiff("d", case_019_inac_date, Now) > 10 then call case_019_spec_xfer
  IF case_020_MMIS_end <> "99/99/99" and case_020 <> "        " and DateDiff("d", case_020_inac_date, Now) > 10 then call case_020_spec_xfer
  IF case_021_MMIS_end <> "99/99/99" and case_021 <> "        " and DateDiff("d", case_021_inac_date, Now) > 10 then call case_021_spec_xfer
  IF case_022_MMIS_end <> "99/99/99" and case_022 <> "        " and DateDiff("d", case_022_inac_date, Now) > 10 then call case_022_spec_xfer
  IF case_023_MMIS_end <> "99/99/99" and case_023 <> "        " and DateDiff("d", case_023_inac_date, Now) > 10 then call case_023_spec_xfer
  IF case_024_MMIS_end <> "99/99/99" and case_024 <> "        " and DateDiff("d", case_024_inac_date, Now) > 10 then call case_024_spec_xfer
End sub

Sub third_page_spec_xfer
  IF case_025_MMIS_end <> "99/99/99" and case_025 <> "        " and DateDiff("d", case_025_inac_date, Now) > 10 then call case_025_spec_xfer
  IF case_026_MMIS_end <> "99/99/99" and case_026 <> "        " and DateDiff("d", case_026_inac_date, Now) > 10 then call case_026_spec_xfer
  IF case_027_MMIS_end <> "99/99/99" and case_027 <> "        " and DateDiff("d", case_027_inac_date, Now) > 10 then call case_027_spec_xfer
  IF case_028_MMIS_end <> "99/99/99" and case_028 <> "        " and DateDiff("d", case_028_inac_date, Now) > 10 then call case_028_spec_xfer
  IF case_029_MMIS_end <> "99/99/99" and case_029 <> "        " and DateDiff("d", case_029_inac_date, Now) > 10 then call case_029_spec_xfer
  IF case_030_MMIS_end <> "99/99/99" and case_030 <> "        " and DateDiff("d", case_030_inac_date, Now) > 10 then call case_030_spec_xfer
  IF case_031_MMIS_end <> "99/99/99" and case_031 <> "        " and DateDiff("d", case_031_inac_date, Now) > 10 then call case_031_spec_xfer
  IF case_032_MMIS_end <> "99/99/99" and case_032 <> "        " and DateDiff("d", case_032_inac_date, Now) > 10 then call case_032_spec_xfer
  IF case_033_MMIS_end <> "99/99/99" and case_033 <> "        " and DateDiff("d", case_033_inac_date, Now) > 10 then call case_033_spec_xfer
  IF case_034_MMIS_end <> "99/99/99" and case_034 <> "        " and DateDiff("d", case_034_inac_date, Now) > 10 then call case_034_spec_xfer
  IF case_035_MMIS_end <> "99/99/99" and case_035 <> "        " and DateDiff("d", case_035_inac_date, Now) > 10 then call case_035_spec_xfer
  IF case_036_MMIS_end <> "99/99/99" and case_036 <> "        " and DateDiff("d", case_036_inac_date, Now) > 10 then call case_036_spec_xfer
end sub


first_page_spec_xfer
If second_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call second_page_spec_xfer
If third_inac_duplicate_check <> "THIS IS THE LAST PAGE" then call third_page_spec_xfer

MsgBox "Success!"  & vbNewLine & vbNewLine &_
"The cases that have HC open in MMIS, or were denied in the last 10 days, are still in your REPT/INAC. Some of these cases may be discrepancies, may be MCRE, or may be active on a spousal case. Check each one of these manually in MMIS and CCOL/CLIC before sending to CLS. "  & vbNewLine & vbNewLine &_
"The script is currently limited to three pages of REPT/INAC that it can process. If you have more than this, run the script again. Copy your claims into an outlook email and send them to the appropriate person(s) in fiscal."
stopscript