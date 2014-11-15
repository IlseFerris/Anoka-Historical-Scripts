'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - pull cases into Excel"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog pull_cases_into_excel_dialog, 0, 0, 416, 85, "Pull cases into Excel dialog"
  CheckBox 15, 20, 55, 10, "PREG exists?", preg_check
  CheckBox 15, 35, 90, 10, "All HH membs 19+?", all_HH_membs_19_plus_check
  CheckBox 15, 50, 90, 10, "Number of HH membs?", number_of_HH_membs_check
  CheckBox 15, 65, 90, 10, "ABAWD code", ABAWD_code_check
  DropListBox 180, 15, 95, 10, "REPT/PND2"+chr(9)+"REPT/ACTV", screen_to_use
  EditBox 190, 30, 90, 15, x102_number
  CheckBox 125, 50, 295, 15, "Check here if you're running this for all staff (WARNING: this could take several hours)", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 365, 10, 50, 15
    CancelButton 365, 30, 50, 15
  GroupBox 10, 5, 110, 75, "Additional items to log"
  Text 125, 15, 50, 10, "Screen to use:"
  Text 125, 35, 60, 10, "Worker to check:"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Dialog asks what stats are being pulled
Dialog pull_cases_into_excel_dialog
If buttonpressed = 0 then stopscript

'Adjusting name of script variable for usage stats according to what was done. So, if ACTV was used instead of PND2, it'll indicate that on the script (and thus allow accurate measurement of time savings).
If screen_to_use = "REPT/PND2" then
	name_of_script = "BULK - pull cases into Excel (PND2)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (PND2 all cases)"
ElseIf screen_to_use = "REPT/ACTV" then
	name_of_script = "BULK - pull cases into Excel (ACTV)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (ACTV all cases)"
End if

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You appear to be locked out of MAXIS")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True


'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X102"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"

'If working off of PND2 it sets the 4th  col as APPL DATE, otherwise it'll be NEXT REVW DATE
If screen_to_use = "REPT/PND2" then
	ObjExcel.Cells(1, 4).Value = "APPL DATE"
ElseIf screen_to_use = "REPT/ACTV" then
	ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"	
End if

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 4 because cols 1-3 are already used
If preg_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "PREG EXISTS?"
	preg_col = col_to_use
	col_to_use = col_to_use + 1
End if
If all_HH_membs_19_plus_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "ALL MEMBS 19+?"
	all_HH_membs_19_plus_col = col_to_use
	col_to_use = col_to_use + 1
End if
If number_of_HH_membs_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "NUMBER OF HH MEMBS?"
	number_of_HH_membs_col = col_to_use
	col_to_use = col_to_use + 1
End if
If ABAWD_code_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "ABAWD CODE"
	ABAWD_code_col = col_to_use
	col_to_use = col_to_use + 1
End if


'Setting the variable for what's to come
excel_row = 2

'If all workers are selected, the script will open the worker list stored on the shared drive, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = 1 then
	'Sets variable for worker list
	worker_list_excel_row = 2
	'Loads worker list
	Set objExcelWorkerList = CreateObject("Excel.Application")
	objExcelWorkerList.Visible = True
	Set objWorkbookWorkerList = objExcelWorkerList.Workbooks.Open("Q:\Blue Zone Scripts\Spreadsheets for script use\worker list.xlsx") 
	objExcel.DisplayAlerts = False
	'Adds each X102 number into an array
	Do
		if ObjExcelWorkerList.Cells(worker_list_excel_row, 1).Value = "" then exit do
		x102_array = x102_array & " " & ObjExcelWorkerList.Cells(worker_list_excel_row, 1).Value
		worker_list_excel_row = worker_list_excel_row + 1
	Loop until ObjExcelWorkerList.Cells(worker_list_excel_row, 1).Value = "" 
	objExcelWorkerList.Workbooks.Close
	objExcelWorkerList.Quit
	x102_array = split(x102_array)
Else
	x102_array = split(x102_number)
End if

For each worker in x102_array
'Getting to PND2, if PND2 is the selected option
If screen_to_use = "REPT/PND2" then
	Call navigate_to_screen("rept", "pnd2")
	EMWriteScreen worker, 21, 17
	transmit

	'Grabbing each case number on screen
	Do
		MAXIS_row = 7
		Do
			EMReadScreen case_number, 8, MAXIS_row, 5
			If case_number = "        " then exit do
			EMReadScreen client_name, 22, MAXIS_row, 16
			EMReadScreen APPL_date, 8, MAXIS_row, 38
			ObjExcel.Cells(excel_row, 1).Value = worker
			ObjExcel.Cells(excel_row, 2).Value = case_number
			ObjExcel.Cells(excel_row, 3).Value = client_name
			ObjExcel.Cells(excel_row, 4).Value = replace(APPL_date, " ", "/")
			MAXIS_row = MAXIS_row + 1
			excel_row = excel_row + 1
		Loop until MAXIS_row = 19
		PF8
		EMReadScreen last_page_check, 21, 24, 2
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if

'Getting to ACTV, if ACTV is the selected option
If screen_to_use = "REPT/ACTV" then
	Call navigate_to_screen("rept", "actv")
	EMWriteScreen worker, 21, 17
	transmit

	'Grabbing each case number on screen
	Do
		MAXIS_row = 7
		Do
			EMReadScreen case_number, 8, MAXIS_row, 12
			If case_number = "        " then exit do
			EMReadScreen client_name, 21, MAXIS_row, 21
			EMReadScreen next_REVW_date, 8, MAXIS_row, 42
			ObjExcel.Cells(excel_row, 1).Value = worker
			ObjExcel.Cells(excel_row, 2).Value = case_number
			ObjExcel.Cells(excel_row, 3).Value = client_name
			ObjExcel.Cells(excel_row, 4).Value = replace(next_REVW_date, " ", "/")
			MAXIS_row = MAXIS_row + 1
			excel_row = excel_row + 1
		Loop until MAXIS_row = 19
		PF8
		EMReadScreen last_page_check, 21, 24, 2
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if

next

'Resetting excel_row variable, now we need to start looking people up
excel_row = 2 

Do 
	case_number = ObjExcel.Cells(excel_row, 2).Value
	If case_number = "" then exit do

	'Now pulling PREG info
	If preg_check = 1 then
		call navigate_to_screen("STAT", "PREG")
		EMReadScreen PREG_panel_check, 1, 2, 78
		If PREG_panel_check <> "0" then 
			ObjExcel.Cells(excel_row, preg_col).Value = "Y"
		Else
			ObjExcel.Cells(excel_row, preg_col).Value = "N"
		End if
	End if

	'Now pulling age info
	If all_HH_membs_19_plus_check = 1 then
		call navigate_to_screen("STAT", "MEMB")
		Do
			EMReadScreen MEMB_panel_current, 1, 2, 73
			EMReadScreen MEMB_panel_total, 1, 2, 78
			EMReadScreen MEMB_age, 3, 8, 76
			If MEMB_age = "   " then MEMB_age = "0"
			If cint(MEMB_age) < 19 then has_minor_in_case = True
			transmit
		Loop until MEMB_panel_current = MEMB_panel_total
		If has_minor_in_case <> True then 
			ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "Y"
		Else
			ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "N"
		End if
		has_minor_in_case = "" 'clearing variable
	End if

	'Now pulling number of membs info
	If number_of_HH_membs_check = 1 then
		call navigate_to_screen("STAT", "MEMB")
		EMReadScreen MEMB_panel_total, 1, 2, 78
		ObjExcel.Cells(excel_row, number_of_HH_membs_col).Value = cint(MEMB_panel_total)
	End if

	'Now pulling ABAWD info
	If ABAWD_code_check = 1 then
		call navigate_to_screen("STAT", "WREG")
		EMReadScreen ERRR_check, 4, 2, 52		'Error prone case checking
		If ERRR_check = "ERRR" then transmit	'transmitting if case is error prone
		EMReadScreen WREG_panel_total, 1, 2, 78
		If WREG_panel_total <> "0" then
			WREG_row = 5 'setting variable for do...loop
			WREG_membs_array = "" 'Clearing variable to use in the do...loop
			Do
				EMReadScreen WREG_ref_nbr, 2, WREG_row, 3
				If WREG_ref_nbr = "  " then exit do
				WREG_membs_array = WREG_membs_array & WREG_ref_nbr & ", "
				WREG_row = WREG_row + 1
			Loop until WREG_row = 19
			WREG_membs_array = split(WREG_membs_array, ", ")
			For each WREG_memb in WREG_membs_array
				EMWriteScreen WREG_memb, 20, 76
				transmit
				EMReadScreen ABAWD_status_code, 2, 13, 50
				If WREG_memb <> "" then ABAWD_status = ABAWD_status & WREG_memb & ": " & ABAWD_status_code & ", "
			Next
			ObjExcel.Cells(excel_row, ABAWD_code_col).Value = "'" & left(ABAWD_status, len(ABAWD_status) - 2)
			ABAWD_status = "" 'clearing variable
		End if
	End if

	excel_row = excel_row + 1
Loop until case_number = ""

'Logging usage stats
script_end_procedure("")