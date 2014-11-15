'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "STATISTICS - monthly agency stats"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DEFINING VARIABLES----------------------------------------------------------------------------------------------------
path_for_excel_file = "Q:\Blue Zone Scripts\Spreadsheets for script use\agency stats template.xlsx" 
excel_row = 4 'this is the row the numbers start on the spreadsheet

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Grabbing user ID to validate user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_validation = ucase(objNet.UserName)

'Validating user ID
If user_ID_for_validation <> "VKCARY" and _
   user_ID_for_validation <> "VLANDERS" _
   then script_end_procedure("User " & user_ID_for_validation & " is not authorized to use this script. To be added to the allowed users' group, email Veronica Cary, and include the user ID indicated. Thank you!")

'Connecting to BlueZone
EMConnect ""

'Checking MAXIS
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure "MAXIS not found. Check to make sure you have MAXIS open and you aren't passworded out."

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open(path_for_excel_file) 
objExcel.DisplayAlerts = True



'Getting to REPT/ARST
call navigate_to_screen("rept", "arst")

EMReadScreen accumulations_timestamp, 30, 19, 40
accumulations_timestamp = trim(accumulations_timestamp)

objExcel.worksheets("Monthly Stats - Anoka").Activate
ObjExcel.Cells(2, 10).Value = "(from MAXIS as of " & accumulations_timestamp & ")"


Do
  worker_number = ObjExcel.Cells(excel_row, 1).Value
  If worker_number <> "" then
    EMWriteScreen worker_number, 3, 31
    transmit
    PF8
    EMReadScreen caseload_count, 4, 8, 24
    EMReadScreen MSA_count, 7, 15, 13
    EMReadScreen GA_count, 7, 16, 13
    PF8
    EMReadScreen GRH_count, 7, 8, 13
    EMReadScreen SNAP_count, 7, 14, 13
    PF8
    EMReadScreen HC_count, 7, 8, 13
    PF7
    PF7
    PF7
    ObjExcel.Cells(excel_row, 10).Value = caseload_count
    ObjExcel.Cells(excel_row, 11).Value = SNAP_count
    ObjExcel.Cells(excel_row, 12).Value = GA_count
    ObjExcel.Cells(excel_row, 13).Value = MSA_count
    ObjExcel.Cells(excel_row, 14).Value = GRH_count
    ObjExcel.Cells(excel_row, 15).Value = HC_count
  End if
  excel_row = excel_row + 1
Loop until excel_row = 70

objExcel.worksheets("Monthly Stats - Blaine").Activate
ObjExcel.Cells(2, 10).Value = "(from MAXIS as of " & accumulations_timestamp & ")"
excel_row = 4 'resetting the variable

Do
  worker_number = ObjExcel.Cells(excel_row, 1).Value
  If worker_number <> "" then
    EMWriteScreen worker_number, 3, 31
    transmit
    PF8
    EMReadScreen caseload_count, 4, 8, 24
    EMReadScreen MFIP_count, 7, 11, 13
    EMReadScreen DWP_count, 7, 13, 13
    EMReadScreen WB_count, 7, 14, 13
    PF8
    EMReadScreen SNAP_count, 7, 14, 13
    PF8
    EMReadScreen HC_count, 7, 8, 13
    PF7
    PF7
    PF7
    ObjExcel.Cells(excel_row, 10).Value = caseload_count
    ObjExcel.Cells(excel_row, 11).Value = SNAP_count
    ObjExcel.Cells(excel_row, 12).Value = MFIP_count
    ObjExcel.Cells(excel_row, 13).Value = DWP_count
    ObjExcel.Cells(excel_row, 14).Value = WB_count
    ObjExcel.Cells(excel_row, 15).Value = HC_count
  End if
  excel_row = excel_row + 1
Loop until excel_row = 90

MsgBox "Success! The statistics have loaded."

script_end_procedure("")