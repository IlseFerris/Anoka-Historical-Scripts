'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "assign PMI to MCRE infopac list"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DECLARING VARIABLES
file_path = "H:\MCRE infopac list.xlsx"
excel_row = 2 'Starts with row 2
RCIN_row = 11 'First row on RCIN is 11

'SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Loading Excel sheet
'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Open(file_path) 
objExcel.DisplayAlerts = False 'Set this to false to make alerts go away. This is necessary in production.

'Now it will look for MMIS on both screens, and enter into it.. 
attn
EMReadScreen MMIS_A_check, 7, 15, 15
IF MMIS_A_check = "RUNNING" then
  EMSendKey "10"
  transmit
Else
  attn
  EMConnect "B"
  attn
  EMReadScreen MMIS_B_check, 7, 15, 15
  If MMIS_B_check <> "RUNNING" then 
    MsgBox "MMIS does not appear to be running. This script will now stop."
    stopscript
  Else
    EMSendKey "10"
    transmit
  End if
End if
EMFocus 'Bringing window focus to the second screen if needed.

'Sending MMIS back to the beginning screen and checking for a password prompt
Do 
  PF6
  EMReadScreen password_prompt, 38, 2, 23
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
  EMReadScreen session_start, 18, 1, 7
Loop until session_start = "SESSION TERMINATED"

'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themself into MMIS the first time!)
EMWriteScreen "mw00", 1, 2
transmit
transmit

'The following will select the correct version of MMIS. First it looks for C302, then EK01, then C402.
row = 1
col = 1
EMSearch "C302", row, col
If row <> 0 then 
  If row <> 1 then 'It has to do this in case the worker only has one option (as many LTC and OSA workers don't have the option to decide between MAXIS and MCRE case access). The MMIS screen will show the text, but it's in the first row in these instances.
    EMWriteScreen "x", row, 4
    transmit
  End if
Else 'Some staff may only have EK01 (MMIS MCRE). The script will allow workers to use that if applicable.
  row = 1
  col = 1
  EMSearch "EK01", row, col
  If row <> 0 then 
    If row <> 1 then
      EMWriteScreen "x", row, 4
      transmit
    End if
  Else 'Some OSAs have C402 (limited access). This will search for that.
    row = 1
    col = 1
    EMSearch "C402", row, col
    If row <> 0 then 
      If row <> 1 then
        EMWriteScreen "x", row, 4
        transmit
      End if
    Else 'Some OSAs have EKIQ (limited MCRE access). This will search for that.
      row = 1
      col = 1
      EMSearch "EKIQ", row, col
      If row <> 0 then 
        If row <> 1 then
          EMWriteScreen "x", row, 4
          transmit
        End if
      Else
        MsgBox "C402, C302, EKIQ, or EK01 not found. Your access to MMIS may be limited. Contact Ronny Cary if you have questions about using this script."
        stopscript
      End if
    End if
  End if
End if

'Now it finds the recipient file application feature and selects it.
row = 1
col = 1
EMSearch "RECIPIENT FILE APPLICATION", row, col
EMWriteScreen "x", row, col - 3
transmit

Do

'Pulling case number
case_number = ObjExcel.Cells(excel_row, 5)

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "i", 2, 19
EMWriteScreen case_number, 9, 19
transmit
EMReadscreen RKEY_check, 4, 1, 52
If RKEY_check = "RKEY" then 
  MsgBox "A correct case number was not taken from MAXIS. Check your case number and try again."
  stopscript
End if

'Now it gets to RCIN for this case.
EMWriteScreen "rcin", 1, 8
transmit

'Reads each line to see if the client is active ("A" code in elig col). If so it'll add to the spreadsheet.
Do
  EMReadScreen MEMB_active_check, 1, RCIN_row, 20                                  'Checks the column to see if an "Y" code is indicated in the CV column
  If MEMB_active_check = "Y" then                                                  'If a "Y" code is found it'll add the PMI to the sheet.
    EMReadScreen recipient_ID_for_sheet, 8, RCIN_row, 4                            'Reads the PMI.
    If ObjExcel.Cells(excel_row, 7) = "" then                                      'If blank it'll just write. Otherwise it'll add the PMI to the existing cell.
      ObjExcel.Cells(excel_row, 7) = "'" & recipient_ID_for_sheet
    Else
      ObjExcel.Cells(excel_row, 7) = ObjExcel.Cells(excel_row, 7) & ", " & recipient_ID_for_sheet
    End if
  End if
  RCIN_row = RCIN_row + 1                                                          'Adds one to the variable to check the next row.
Loop until MEMB_active_check = " " or MEMB_active_check = "-"

'Sends a PF3
PF3

'changes to the next excel row
excel_row = excel_row + 1

'resets the RCIN_row
RCIN_row = 11

Loop until excel_row = 3978 'should be 3978 for the first run