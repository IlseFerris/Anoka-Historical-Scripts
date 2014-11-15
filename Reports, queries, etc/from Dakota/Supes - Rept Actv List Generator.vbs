'REPT ACTV Excel Generator - Updated 5-24-2013
'Author - Andy Fink, with help from Ronny C.
'Creates a formatted excel spreadsheet for supes with REPT ACTV data on a given number of selected workers. 

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command 	= run_another_script_fso.OpenTextFile("G:\Scripts\Scripts\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CONNECTS TO MAXIS
EMConnect ""

'DIALOG for worker number
BeginDialog worker_report, 0, 0, 281, 160, "Multiple Worker Report"
  EditBox 55, 30, 40, 15, worker1
  EditBox 55, 50, 40, 15, worker2
  EditBox 55, 70, 40, 15, worker3
  EditBox 55, 90, 40, 15, worker4
  EditBox 55, 110, 40, 15, worker5
  EditBox 145, 30, 40, 15, worker6
  EditBox 145, 50, 40, 15, worker7
  EditBox 145, 70, 40, 15, worker8
  EditBox 145, 90, 40, 15, worker9
  EditBox 145, 110, 40, 15, worker10
  EditBox 230, 30, 40, 15, worker11
  EditBox 230, 50, 40, 15, worker12
  EditBox 230, 70, 40, 15, worker13
  EditBox 230, 90, 40, 15, worker14
  EditBox 230, 110, 40, 15, worker15
  ButtonGroup ButtonPressed
    OkButton 180, 135, 50, 15
    CancelButton 230, 135, 50, 15
  Text 15, 5, 255, 20, "Please enter the last three characters (X119xxx) of each worker you would like to run a report on and click OK:"
  Text 15, 35, 30, 10, "Worker 1"
  Text 15, 55, 30, 10, "Worker 2"
  Text 15, 75, 30, 10, "Worker 3"
  Text 15, 95, 30, 10, "Worker 4"
  Text 15, 115, 30, 10, "Worker 5"
  Text 110, 35, 30, 10, "Worker 6"
  Text 110, 115, 35, 10, "Worker 10"
  Text 110, 75, 30, 10, "Worker 8"
  Text 110, 95, 30, 10, "Worker 9"
  Text 110, 55, 30, 10, "Worker 7"
  Text 195, 35, 35, 10, "Worker 11"
  Text 195, 55, 35, 10, "Worker 12"
  Text 195, 75, 35, 10, "Worker 13"
  Text 195, 95, 35, 10, "Worker 14"
  Text 195, 115, 35, 10, "Worker 15"
EndDialog



'This function writes worker's stats in selected worksheet*****************************************************************************************
function worker_stats

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

'Starts counts at zero for each worker.
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

'Formats Header of worker's worksheet
objExcel.ActiveSheet.PageSetup.LeftHeader = reptDate + "  " + trim(workerName) + "  " + trim(workerFullNumber)

'Puts worker info into 
ObjExcel.Cells(2, 14).Value = reptDate
ObjExcel.Cells(3, 14).Value = workerName
ObjExcel.Cells(4, 14).Value = workerFullNumber

'case info

ObjExcel.Cells(6, 14).Value = "Total Maxis Cases ="
ObjExcel.Cells(6, 15).Value = total_cases - 1
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

ObjExcel.activeSheet.Name = trim(workerName)

ObjExcel.Sheets("Summary").Activate
ObjExcel.Cells(workersummary_row, 1).Value = workerName
ObjExcel.Cells(workersummary_row, 2).Value = total_cases - 1
ObjExcel.Cells(workersummary_row, 3).Value = cash_active
ObjExcel.Cells(workersummary_row, 4).Value = cash_pending
ObjExcel.Cells(workersummary_row, 5).Value = FS_active
ObjExcel.Cells(workersummary_row, 6).Value = FS_Pending
ObjExcel.Cells(workersummary_row, 7).Value = HC_active
ObjExcel.Cells(workersummary_row, 8).Value = HC_pending
ObjExcel.Cells(workersummary_row, 9).Value = ea_pending
ObjExcel.Cells(workersummary_row, 10).Value = cc_active
ObjExcel.Cells(workersummary_row, 11).Value = CC_pending
ObjExcel.Cells(workersummary_row, 12).Value = cc_suspended

workersummary_row = workersummary_row + 1

End Function
'End of Function worker_stats*************************************************************************************





'Run Dialog
Dialog worker_report
If buttonpressed = 0 then stopscript

'Counts total number of workers to create sheets for
total_worksheets = 0
'Defines first line for worker in summar
workersummary_row = 2

if worker1 <> "" then total_worksheets = total_worksheets + 1 End If
if worker2 <> "" then total_worksheets = total_worksheets + 1 End If
if worker3 <> "" then total_worksheets = total_worksheets + 1 End If
if worker4 <> "" then total_worksheets = total_worksheets + 1 End If
if worker5 <> "" then total_worksheets = total_worksheets + 1 End If
if worker6 <> "" then total_worksheets = total_worksheets + 1 End If
if worker7 <> "" then total_worksheets = total_worksheets + 1 End If
if worker8 <> "" then total_worksheets = total_worksheets + 1 End If
if worker9 <> "" then total_worksheets = total_worksheets + 1 End If
if worker10 <> "" then total_worksheets = total_worksheets + 1 End If
if worker11 <> "" then total_worksheets = total_worksheets + 1 End If
if worker12 <> "" then total_worksheets = total_worksheets + 1 End If
if worker13 <> "" then total_worksheets = total_worksheets + 1 End If
if worker14 <> "" then total_worksheets = total_worksheets + 1 End If
if worker15 <> "" then total_worksheets = total_worksheets + 1 End If



'CHECKS FOR PASSWORD PROMPT/MAXIS STATUS
transmit
Do
     EMReadScreen password_prompt, 38, 2, 23
     IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

'Back to Self screen and navigate to REPT/ACTV
back_to_self

'OPENS A NEW EXCEL SPREADSHEET
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
Set objWorkbook = objExcel.Workbooks.Add() 

'Calculates how many worksheets to add to create enough worksheets for entered workers plus summary
For i = 0 to total_worksheets - 3
objExcel.ActiveWorkbook.Worksheets.Add
Next

'Names and formats summary worksheet
objexcel.worksheets(total_worksheets + 1).Name = "Summary"

ObjExcel.Sheets("Summary").Activate

ObjExcel.Cells(1, 1).Value = "Worker"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 1).ColumnWidth = 23
ObjExcel.Cells(1, 2).Value = "Total Maxis Cases"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 2).ColumnWidth = 18
ObjExcel.Cells(1, 3).Value = "CA ACTV"
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 3).ColumnWidth = 8
ObjExcel.Cells(1, 4).Value = "CA PEND"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 8
ObjExcel.Cells(1, 5).Value = "FS ACTV"
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 5).ColumnWidth = 8
ObjExcel.Cells(1, 6).Value = "FS PEND"
objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(1, 6).ColumnWidth = 8
ObjExcel.Cells(1, 7).Value = "HC ACTV"
objExcel.Cells(1, 7).Font.Bold = TRUE
objExcel.Cells(1, 7).ColumnWidth = 8
ObjExcel.Cells(1, 8).Value = "HC PEND"
objExcel.Cells(1, 8).Font.Bold = TRUE
objExcel.Cells(1, 8).ColumnWidth = 8
ObjExcel.Cells(1, 9).Value = "EA PEND"
objExcel.Cells(1, 9).Font.Bold = TRUE
objExcel.Cells(1, 9).ColumnWidth = 8
ObjExcel.Cells(1, 10).Value = "CC ACTV"
objExcel.Cells(1, 10).Font.Bold = TRUE
objExcel.Cells(1, 10).ColumnWidth = 8
ObjExcel.Cells(1, 11).Value = "CC PEND"
objExcel.Cells(1, 11).Font.Bold = TRUE
objExcel.Cells(1, 11).ColumnWidth = 8
ObjExcel.Cells(1, 12).Value = "CC SUSP"
objExcel.Cells(1, 12).Font.Bold = TRUE
objExcel.Cells(1, 12).ColumnWidth = 8

objExcel.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Formats Header
objExcel.ActiveSheet.PageSetup.LeftHeader = Date





'call worker1
ObjExcel.Sheets("Sheet1").Activate
workernumber = worker1
call worker_stats
if worker2 = "" then stopscript

'worker2
ObjExcel.Sheets("Sheet2").Activate
workernumber = worker2
call worker_stats
if worker3 = "" then stopscript

'worker3
ObjExcel.Sheets("Sheet4").Activate
workernumber = worker3
call worker_stats
if worker4 = "" then stopscript

'worker4
ObjExcel.Sheets("Sheet5").Activate
workernumber = worker4
call worker_stats
if worker5 = "" then stopscript

'worker5
ObjExcel.Sheets("Sheet6").Activate
workernumber = worker5
call worker_stats
if worker6 = "" then stopscript

'worker6
ObjExcel.Sheets("Sheet7").Activate
workernumber = worker6
call worker_stats
if worker7 = "" then stopscript

'worker7
ObjExcel.Sheets("Sheet8").Activate
workernumber = worker7
call worker_stats
if worker8 = "" then stopscript

'worker8
ObjExcel.Sheets("Sheet9").Activate
workernumber = worker8
call worker_stats
if worker9 = "" then stopscript

'worker9
ObjExcel.Sheets("Sheet10").Activate
workernumber = worker9
call worker_stats
if worker10 = "" then stopscript

'worker10
ObjExcel.Sheets("Sheet11").Activate
workernumber = worker10
call worker_stats
if worker11 = "" then stopscript

'worker11
ObjExcel.Sheets("Sheet12").Activate
workernumber = worker11
call worker_stats
if worker12 = "" then stopscript

'worker12
ObjExcel.Sheets("Sheet13").Activate
workernumber = worker12
call worker_stats
if worker13 = "" then stopscript

'worker13
ObjExcel.Sheets("Sheet14").Activate
workernumber = worker13
call worker_stats
if worker14 = "" then stopscript

'worker14
ObjExcel.Sheets("Sheet15").Activate
workernumber = worker14
call worker_stats
if worker15 = "" then stopscript

'worker15
ObjExcel.Sheets("Sheet16").Activate
workernumber = worker15
call worker_stats

stopscript