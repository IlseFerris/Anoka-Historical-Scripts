'REPT ACTV Excel Generator
'Author - Andy Fink, with help from Ronny C.
'Creates a formatted excel spreadsheet with REPT ACTV data from selected caseworker's caseload

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command 	= run_another_script_fso.OpenTextFile("G:\Scripts\Scripts\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CONNECTS TO MAXIS
EMConnect ""

'DIALOG for worker number
BeginDialog WorkerRequest, 0, 0, 191, 60, "REPT/ACTV"
  EditBox 20, 40, 40, 15, workerNumber
  ButtonGroup ButtonPressed
    OkButton 75, 40, 50, 15
    CancelButton 130, 40, 50, 15
  Text 5, 0, 180, 35, "You are about to create a list of your REPT ACTV in Excel. If you would like to create a list for someone other than yourself please put the last three characters of their X119### below. Otherwise please press OK."
EndDialog

'Run Dialog
Dialog WorkerRequest
If buttonpressed = 0 then stopscript

'CHECKS FOR PASSWORD PROMPT/MAXIS STATUS
transmit
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

'Back to Self screen and navigate to REPT/ACTV
back_to_self
EMWriteScreen "REPT", 16, 43
EMWRiteScreen "        ", 18, 43
EMWriteScreen "ACTV", 21, 70
Transmit

'Checks to see if other worker number is put in and routes script accordingly
If workerNumber <> "" then 
EmWriteScreen workerNumber, 21, 17
Transmit
End If

EmReadScreen workerName, 20, 3, 11
EmReadScreen reptDate, 8, 3, 41
EmReadScreen workerFullNumber, 7, 21, 13


'DEFINES THE EXCEL_ROW VARIABLE FOR WORKING WITH THE SPREADSHEET
excel_row = 2

'Sets counts to zero for worker statistics
cash_active = 0
cash_pending = 0
fs_active = 0
fs_pending = 0
total_cases = 0
cc_active = 0
cc_pending = 0
cc_suspended = 0
hc_active = 0
hc_pending = 0
ea_pending = 0



FS_total = 0
HC_total = 0

'OPENS A NEW EXCEL SPREADSHEET
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
Set objWorkbook = objExcel.Workbooks.Add() 

'Formatting for the excel spreadsheet


ObjExcel.Cells(1, 1).Value = "Maxis #"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 1).ColumnWidth = 9
ObjExcel.Cells(1, 2).Value = "Client Name"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 2).ColumnWidth = 23
ObjExcel.Cells(1, 3).Value = "REVW Date"
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 3).ColumnWidth = 10
ObjExcel.Cells(1, 4).Value = "CA Type"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 7
ObjExcel.Cells(1, 5).Value = "CA"
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 5).ColumnWidth = 3
ObjExcel.Cells(1, 6).Value = "FS"
objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(1, 6).ColumnWidth = 3
ObjExcel.Cells(1, 7).Value = "HC"
objExcel.Cells(1, 7).Font.Bold = TRUE
objExcel.Cells(1, 7).ColumnWidth = 3
ObjExcel.Cells(1, 8).Value = "EA"
objExcel.Cells(1, 8).Font.Bold = TRUE
objExcel.Cells(1, 8).ColumnWidth = 3
ObjExcel.Cells(1, 9).Value = "GR"
objExcel.Cells(1, 9).Font.Bold = TRUE
objExcel.Cells(1, 9).ColumnWidth = 3
ObjExcel.Cells(1, 10).Value = "IVE"
objExcel.Cells(1, 10).Font.Bold = TRUE
objExcel.Cells(1, 10).ColumnWidth = 3
ObjExcel.Cells(1, 11).Value = "FIAT"
objExcel.Cells(1, 11).Font.Bold = TRUE
objExcel.Cells(1, 11).ColumnWidth = 4
ObjExcel.Cells(1, 12).Value = "CC"
objExcel.Cells(1, 12).Font.Bold = TRUE
objExcel.Cells(1, 12).ColumnWidth = 3
objExcel.Cells(1, 14).ColumnWidth = 20
objExcel.Cells(6, 14).Font.Bold = TRUE

'Freezes first 
objExcel.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True


'THIS DO...LOOP DUMPS THE CASE NUMBER AND NAME OF EACH CLIENT INTO A SPREADSHEET
Do until last_page_check = "THIS IS THE LAST PAGE"
  EMReadScreen last_page_check, 21, 24, 02
  EMReadScreen current_page_check, 1, 3, 79

row = 7 'defining the row to look at



	Do
    
      EMReadScreen case_number, 8, row, 12 'grabbing case number
      EMReadScreen client_name, 19, row, 21 'grabbing client name
	EMReadScreen revw_month, 2, row, 42 'grabs review date month
	EMReadScreen revw_day, 2, row, 45 'grabs review date day
	EMReadScreen revw_year, 2, row, 48 'grabs review date year
	EMReadScreen cash_type, 2, row, 51 'grabs cash type
	EMReadScreen cash_status, 1, row, 54 'grabs cash status
	EMReadScreen fs_status, 1, row, 61 'grabs FS status
	EMReadScreen hc_status, 1, row, 64 'grabs HC status
	EMReadScreen ea_status, 1, row, 67 'brabs EA status
	EMReadScreen gr_status, 1, row, 70 'grabs GR status
	EMReadScreen ive_status, 1, row, 74 'grabs IVE status
	EMReadScreen fiat_status, 1, row, 77 'grabs FIAT status
	EMReadScreen cc_status, 1, row, 80 'grabs CC status

	revw_date = revw_month + "/" + revw_day + "/" + revw_year

    	ObjExcel.Cells(excel_row, 1).Value = trim(case_number)
    	ObjExcel.Cells(excel_row, 2).Value = trim(client_name)
	ObjExcel.Cells(excel_row, 3).Value = trim(revw_date)
	ObjExcel.Cells(excel_row, 4).Value = trim(cash_type)
	ObjExcel.Cells(excel_row, 5).Value = trim(cash_status)
	ObjExcel.Cells(excel_row, 6).Value = trim(fs_status)
	ObjExcel.Cells(excel_row, 7).Value = trim(hc_status)
	ObjExcel.Cells(excel_row, 8).Value = trim(ea_status)
	ObjExcel.Cells(excel_row, 9).Value = trim(gr_status)
	ObjExcel.Cells(excel_row, 10).Value = trim(ive_status)
	ObjExcel.Cells(excel_row, 11).Value = trim(fiat_status)
	ObjExcel.Cells(excel_row, 12).Value = trim(cc_status)
    	
	if ObjExcel.Cells(excel_row, 5).Value = "A"  then cash_active = cash_active + 1
      if ObjExcel.Cells(excel_row, 5).Value = "P"  then cash_pending = cash_pending + 1
	if ObjExcel.Cells(excel_row, 6).Value = "A"  then fs_active = fs_active + 1
	if ObjExcel.Cells(excel_row, 6).Value = "P"  then fs_pending = fs_pending + 1
	if ObjExcel.Cells(excel_row, 7).Value = "A"  then hc_active = hc_active + 1
      if ObjExcel.Cells(excel_row, 7).Value = "P"  then hc_pending = hc_pending + 1
	if ObjExcel.Cells(excel_row, 8).Value = "P"  then ea_pending = ea_pending + 1
	if ObjExcel.Cells(excel_row, 12).Value = "A"  then cc_active = cc_active + 1
	if ObjExcel.Cells(excel_row, 12).Value = "P"  then cc_pending = cc_pending + 1
	if ObjExcel.Cells(excel_row, 12).Value = "S"  then cc_suspended = cc_suspended + 1
	

	total_cases = total_cases + 1
	excel_row = excel_row + 1
    	row = row + 1
	
  	Loop until row = 19 or trim(case_number) = ""
	
	
 PF8 'going to the next screen
  if last_page_check = "THIS IS THE LAST PAGE" and current_page_check = "1" then exit do 'allows do...loop to exit if there's only one page
Loop 

if ObjExcel.Cells(excel_row -1, 3).Value = "/  /" then
ObjExcel.Cells(excel_row -1, 3).Value = ""
end if

objExcel.ActiveSheet.PageSetup.LeftHeader = reptDate + "  " + trim(workerName) + "  " + trim(workerFullNumber)
'worker info

ObjExcel.Cells(2, 14).Value = reptDate
ObjExcel.Cells(3, 14).Value = workerName
ObjExcel.Cells(4, 14).Value = workerFullNumber

'case info

ObjExcel.Cells(6, 14).Value = "Total Maxis Cases ="
ObjExcel.Cells(6, 15).Value = total_cases
ObjExcel.Cells(7, 14).Value = "Active cash cases ="
ObjExcel.Cells(7, 15).Value = cash_active
ObjExcel.Cells(8, 14).Value = "Pending cash cases ="
ObjExcel.Cells(8, 15).Value = cash_pending
ObjExcel.Cells(9, 14).Value = "Active FS cases ="
ObjExcel.Cells(9, 15).Value = FS_active
ObjExcel.Cells(10, 14).Value = "Pending FS cases ="
ObjExcel.Cells(10, 15).Value = fs_pending
ObjExcel.Cells(11, 14).Value = "Active CC cases ="
ObjExcel.Cells(11, 15).Value = cc_active
ObjExcel.Cells(12, 14).Value = "Pending CC cases ="
ObjExcel.Cells(12, 15).Value = CC_pending
ObjExcel.Cells(13, 14).Value = "Suspended CC cases ="
ObjExcel.Cells(13, 15).Value = cc_suspended
ObjExcel.Cells(14, 14).Value = "Active HC ="
ObjExcel.Cells(14, 15).Value = hc_active
ObjExcel.Cells(15, 14).Value = "Pending HC ="
ObjExcel.Cells(15, 15).Value = hc_pending
ObjExcel.Cells(16, 14).Value = "Pending EA ="
ObjExcel.Cells(16, 15).Value = ea_pending

objExcel.Sheets(1).Name = workerName