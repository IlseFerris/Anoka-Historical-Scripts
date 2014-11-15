'Only removed 03/27/2014 as I'm unlikely to be able to finish before my maternity leave. Will re-evaluate in the future.

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ALPHA - BULK - fake case APPLer"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES TO DECLARE <<<<<<<<<<CREATE DIALOG
Do
	amt_of_cases_to_make = inputbox("Amount of cases to make?", 1)
	If amt_of_cases_to_make = "" then stopscript
	If isnumeric(amt_of_cases_to_make) = False then MsgBox "You must enter a number for the amount of cases to make."
Loop until isnumeric(amt_of_cases_to_make) = True

Do
	application_date = inputbox("Application date:", 1)
	If application_date = "" then stopscript
	If isdate(application_date) = False then MsgBox "You must enter a date here."
Loop until isdate(application_date) = True

'<<<<REPLACE THIS DEFAULT VARIABLE WITH A DIALOG SELECTOR FOR THE DIFFERENT CASES ONCE THE SCRIPT IS TESTED
excel_row = 5

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("H:\Bulk script projects\fake case scenario loading sheet.xlsx") 
objExcel.DisplayAlerts = True



'Loading the variables from the spreadsheet
married_couple = ObjExcel.Cells(excel_row, 1).Value
HH_member_array = split(ObjExcel.Cells(excel_row, 2).Value, ", ")
ADDR_line_01 = ObjExcel.Cells(excel_row, 3).Value
ADDR_line_02 = ObjExcel.Cells(excel_row, 4).Value
city_line = ObjExcel.Cells(excel_row, 5).Value
zip_line = ObjExcel.Cells(excel_row, 6).Value

'Has to determine what SSN number to start with, based on a text file stored on the network. We can't duplicate SSNs so this is vital.
'Variable for the text file
file = "Q:\Blue Zone Scripts\Training case creators\SSN identifier number.txt"

'Opening text file and reading contents into SSN_identifier variable, then closing
Set objFSO = CreateObject("Scripting.FileSystemObject")
set objTS=objFSO.opentextfile(file, 1)
SSN_identifier = objTS.ReadAll
objTS.Close

'Determining the SSN_identified based on the above variable
SSN_identifier = cint(SSN_identifier)

'Connecting to MAXIS
EMConnect ""

'This big for...next statement will APPL the cases.
For i = 1 to amt_of_cases_to_make

	EMReadScreen APPL_check, 4, 2, 45
	If APPL_check <> "APPL" then script_end_procedure("Not on APPL.") 'Must start on the APPL screen, this could change if an additional function were written to get to APPL in the right footer month!

	'The following for...next statement adds each person for the HH member array, splitting the data up and entering it into MAXIS.
	For each HH_member in HH_member_array		
		split_array = split(HH_member, "|")
		member_number = split_array(0)
		first_name = split_array(1)
		last_name = split_array(2)
		year_of_birth = split_array(3)
		gender = split_array(4)
		If member_number = "01" then 			'Enters the member 01 on the APPL panel, then jumps to the next screen.
			call create_MAXIS_friendly_date(application_date, 0, 4, 63)
			EMWriteScreen last_name, 7, 30
			EMWriteScreen first_name, 7, 63
			transmit
			EMReadScreen APPL_check, 4, 2, 45		'Checks to make sure we've moved past the APPL screen. If we haven't, the script will stop.
			If APPL_check = "APPL" then script_end_procedure("Error!")
	      End if
	    
		'Now it enters complete info on the HH members
		If member_number <> "01" then
			EMWriteScreen member_number, 4, 33
			EMWriteScreen last_name, 6, 30
			EMWriteScreen first_name, 6, 63
		End if
		EMWriteScreen "474", 7, 42
		EMWriteScreen "47", 7, 46
		SSN_last_four_digits = SSN_identifier
		Do
			If len(SSN_last_four_digits) < 4 then SSN_last_four_digits = "0" & SSN_last_four_digits 
		Loop until len(SSN_last_four_digits) = 4
		EMWriteScreen SSN_last_four_digits, 7, 49
		EMWriteScreen "P", 7, 68
		EMWriteScreen "01", 8, 42
		EMWriteScreen "01", 8, 45
		EMWriteScreen year_of_birth, 8, 48
		EMWriteScreen "OT", 8, 68
		EMWriteScreen gender, 9, 42
		If member_number = "02" then
			EMWriteScreen "02", 10, 42
		ElseIf member_number = "24" then
			EMWriteScreen "24", 10, 42
		ElseIf member_number <> "01" then
			EMWriteScreen "03", 10, 42
		End if
		EMWriteScreen "DL", 9, 68
		EMWriteScreen "99", 12, 42
		EMWriteScreen "99", 13, 42
		EMWriteScreen "N", 14, 68
		EMWriteScreen "N", 15, 42
		EMWriteScreen "N", 16, 68
		EMWriteScreen "x", 17, 34   'It needs to enter a code for race. It is set to do "unable to determine".
		transmit
		EMWriteScreen "x", 15, 12   
		transmit
		transmit

		'Now it checks to make sure there's no duplicate SSNs. If there is, it goes back and makes another until it's done.
		Do
			EMReadScreen SSN_as_entered, 11, 4, 4
			EMReadScreen first_SSN_listed, 11, 8, 7
			If SSN_as_entered = first_SSN_listed then
				PF3
				EMWriteScreen "474", 7, 42
				EMWriteScreen "47", 7, 46
				SSN_identifier = SSN_identifier + 1
				SSN_last_four_digits = SSN_identifier
					Do
						If len(SSN_last_four_digits) < 4 then SSN_last_four_digits = "0" & SSN_last_four_digits 
					Loop until len(SSN_last_four_digits) = 4
				EMWriteScreen SSN_last_four_digits, 7, 49
				transmit
			End if
		Loop until SSN_as_entered <> first_SSN_listed
    
		'Now it creates the new PMI entry.
		PF8
		PF8
		PF8
		PF8
		PF8
		PF5
		EMWriteScreen "y", 6, 67
		transmit

		'Now it enters MEMI information
		If married_couple = FALSE then marital_status = "N"
		If married_couple = TRUE then marital_status = "M"
		call create_MAXIS_friendly_date(application_date, 0, 6, 35)		'This date format is different from other MAXIS formats (YYYY instead of YY), so there's some workaround code below
		EMReadScreen footer_month, 2, 20, 58					'Reading the footer month
		EMWriteScreen "20" & footer_month, 6, 41					'Adding "20" to the footer month and including it
		EMReadScreen ref_nbr, 2, 4, 33
		If ref_nbr <> "01" and ref_nbr <> "02" then marital_status = "N"
		EMWriteScreen marital_status, 7, 49
		if married_couple = True and ref_nbr = "01" then EMWriteScreen "02", 8, 49
		if married_couple = True and ref_nbr = "02" then EMWriteScreen "01", 8, 49
		age = datepart("yyyy", date) - cint(year_of_birth)
		last_grade_completed = age - 6
		If age > 18 then last_grade_completed = 12
		If age < 6 then last_grade_completed = "00"
		If len(last_grade_completed) = 1 then last_grade_completed = "0" & last_grade_completed
		EMWriteScreen last_grade_completed, 9, 49
		EMWriteScreen "y", 10, 49
		EMWriteScreen "no", 10, 78
		EMWriteScreen "y", 13, 49
		EMWriteScreen "n", 13, 78
		transmit
		SSN_identifier = SSN_identifier + 1
	Next
  
  
	'Now it transmits, to get to the ADDR screen. 
	Transmit
  
	'Now it enters the address.
	EMWriteScreen application_month, 4, 43
	EMWriteScreen application_day, 4, 46
	EMWriteScreen application_year, 4, 49
	EMWriteScreen ADDR_line_01, 6, 43
	EMWriteScreen ADDR_line_02, 7, 43
	EMWriteScreen city_line, 8, 43
	EMWriteScreen "MN", 8, 66
	EMWriteScreen zip_line, 9, 43
	EMWriteScreen "02", 9, 66
	EMWriteScreen "SF", 9, 74
	EMWriteScreen "N", 10, 43
	transmit
	transmit
	PF3
	EMWriteScreen "APPL", 16, 43
	EMWriteScreen "________", 18, 43
	transmit
Next

MsgBox "Done APPLing. SSN identifier ended at: " & SSN_identifier


'Opening up the text file again, writing the new number into the file, then closing
set objTS=objFSO.opentextfile(file, 2)
ObjTS.WriteLine(SSN_identifier)
objTS.Close

'Getting to PND1
PF10
back_to_self
EMWriteScreen "REPT", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "pnd1", 21, 70
transmit

'VARIABLES TO DECLARE FOR PND1-ing

just_memb_01 = ObjExcel.Cells(excel_row, 7).Value
cash_app = ObjExcel.Cells(excel_row, 8).Value
HC_app = ObjExcel.Cells(excel_row, 9).Value
FS_app = ObjExcel.Cells(excel_row, 10).Value
emer_app = ObjExcel.Cells(excel_row, 11).Value
paperless_indicator = ObjExcel.Cells(excel_row, 12).Value

'Now the script will process the PND1 piece for each case
Do
	EMReadScreen PND1_check, 4, 2, 50			'Confirms we aren't on PND1. If we are the script will stop.
	If PND1_check <> "PND1" then script_end_procedure("Not on PND1")
	EMReadScreen case_number, 8, 7, 3			'Grabs top case number. Assumes all PND1 cases are for this same scenario. To change this add case numbers to spreadsheet or array and run from that
	case_number = trim(case_number)			'Removing spaces
	If case_number = "" then exit do			'Because we're probably done if there's no case number
	EMWriteScreen "stat", 20, 13				'Getting to STAT
	EMWriteScreen "________", 20, 33			'Clearing any case number data
	EMWriteScreen case_number, 20, 33			'Entering new case number
	transmit							'Has to go twice to get to new screen
	transmit
	EMWriteScreen "N", 6, 64				'These are always "N"s
	EMWriteScreen "N", 6, 73
	MAXIS_row = 6						'The first row a HH member can be on

	'This will add the data from the spreadsheet to all of the HH members indicated
	Do
		EMWriteScreen cash_app, MAXIS_row, 28
		EMWriteScreen HC_app, MAXIS_row, 37
		EMWriteScreen FS_app, MAXIS_row, 46
		EMWriteScreen emer_app, MAXIS_row, 55
		MAXIS_row = MAXIS_row + 1
		EMReadScreen member_row_check, 2, MAXIS_row, 3
		If just_memb_01 = TRUE and MAXIS_row = 7 then   'Separates the original data to simplify the do...loop. It restores after the loop. This ensures that PROG gets coded correctly.
			actual_cash_app = cash_app
			actual_HC_app = HC_app
			actual_FS_app = FS_app 
			actual_emer_app = emer_app 
			cash_app = "N"
			HC_app = "N"
			FS_app = "N"
			emer_app = "N"
		End if
	Loop until member_row_check = "  "

	If just_memb_01 = True then   'Restoring original values
		cash_app = actual_cash_app
		HC_app = actual_HC_app
		FS_app = actual_FS_app 
		emer_app = actual_emer_app 
	End if

	transmit

	If cash_app = "Y" then
		EMReadScreen appl_month, 2, 6, 33
		EMReadScreen appl_day, 2, 6, 36
		EMReadScreen appl_year, 2, 6, 39
		EMWriteScreen appl_month, 6, 55
		EMWriteScreen appl_day, 6, 58
		EMWriteScreen appl_year, 6, 61
	End if

	If emer_app = "Y" then
		EMReadScreen appl_month, 2, 8, 33
		EMReadScreen appl_day, 2, 8, 36
		EMReadScreen appl_year, 2, 8, 39
		EMWriteScreen appl_month, 8, 55
		EMWriteScreen appl_day, 8, 58
		EMWriteScreen appl_year, 8, 61
		EMWriteScreen "EG", 8, 67
	End if

	If FS_app = "Y" then
		EMReadScreen appl_month, 2, 10, 33
		EMReadScreen appl_day, 2, 10, 36
		EMReadScreen appl_year, 2, 10, 39
		EMWriteScreen appl_month, 10, 55
		EMWriteScreen appl_day, 10, 58
		EMWriteScreen appl_year, 10, 61
	End if

	If HC_app = "Y" then
		EMReadScreen appl_month, 2, 12, 33
		EMReadScreen appl_day, 2, 12, 36
		EMReadScreen appl_year, 2, 12, 39
	End if

	EMWriteScreen "N", 18, 67

	transmit

	If HC_app = "Y" then transmit 'HC cases jump to STAT/HCRE

	application_date = cdate(appl_month & "/" & appl_day & "/" & appl_year)
	six_month_recert_date = dateadd("m", 6, application_date)
	six_month_month = datepart("m", six_month_recert_date)
	If len(six_month_month) = 1 then six_month_month = "0" & six_month_month 
	six_month_year = datepart("yyyy", six_month_recert_date) - 2000
	one_year_recert_date = dateadd("m", 12, application_date)
	one_year_month = datepart("m", one_year_recert_date)
	If len(one_year_month) = 1 then one_year_month = "0" & one_year_month 
	one_year_year = datepart("yyyy", one_year_recert_date) - 2000

	If cash_app = "Y" then
		EMWriteScreen one_year_month, 9, 37
		EMWriteScreen one_year_year, 9, 43
	End if
  
	If FS_app = "Y" then
		EMWriteScreen "N", 15, 75
		EMWriteScreen "x", 5, 58
		transmit
		EMWriteScreen six_month_month, 9, 26
		EMWriteScreen six_month_year, 9, 32
		EMWriteScreen one_year_month, 9, 64
		EMWriteScreen one_year_year, 9, 70
		transmit
		transmit
	End if

	If HC_app = "Y" then
		EMWriteScreen "x", 5, 71
		transmit
		EMWriteScreen six_month_month, 8, 71
		EMWriteScreen six_month_year, 8, 77
		EMWriteScreen one_year_month, 9, 27
		EMWriteScreen one_year_year, 9, 33
		EMWriteScreen paperless_indicator, 9, 71
		transmit
		transmit
	End if

	transmit
	transmit
  
	EMWriteScreen "rept", 16, 43
	EMWriteScreen "________", 18, 43
	EMWriteScreen "pnd1", 21, 70
	transmit
Loop until case_number = ""


'----------------------------------------------------------------------------------------------------PND2 side

Msgbox "wait a few moments to allow the cases to get out of background"

amt_of_times_to_run = amt_of_cases_to_make '<<<<<<<<THIS IS TO MAKE BELOW FUNCTIONS SIMPLER, BUT THIS SHOULD BE UPDATED TO BE DYNAMIC, IE WORK EVERY PND2 CASE OR SOMETHING!

'Loading variables from spreadsheet for PND2 phase
kids_to_add = split(ObjExcel.Cells(excel_row, 13).Value, ", ")
ABPS_action = ObjExcel.Cells(excel_row, 14).Value
ABPS_last_name = ObjExcel.Cells(excel_row, 15).Value
ABPS_first_name = ObjExcel.Cells(excel_row, 16).Value
ABPS_gender = ObjExcel.Cells(excel_row, 17).Value
PARE_action = ObjExcel.Cells(excel_row, 18).Value
both_parents_in_HH = ObjExcel.Cells(excel_row, 19).Value
stepparent_in_HH = ObjExcel.Cells(excel_row, 20).Value
EATS_action = ObjExcel.Cells(excel_row, 21).Value
PP_together = ObjExcel.Cells(excel_row, 22).Value
EATS_member_array = ObjExcel.Cells(excel_row, 23).Value
EATS_non_member_array = ObjExcel.Cells(excel_row, 24).Value
EMPS_action = ObjExcel.Cells(excel_row, 25).Value 		'at this time it just puts "n" for the other stuff. Blank this out to skip. If updating an existing panel, put "01".
fin_orient_dt_month = ObjExcel.Cells(excel_row, 26).Value
fin_orient_dt_day = ObjExcel.Cells(excel_row, 27).Value
fin_orient_dt_year = ObjExcel.Cells(excel_row, 28).Value
full_time_care_of_child_under_1 = ObjExcel.Cells(excel_row, 29).Value
EMPS_Exemption_Care_Of_A_Child_Under_One_month = ObjExcel.Cells(excel_row, 30).Value
EMPS_Exemption_Care_Of_A_Child_Under_One_year = ObjExcel.Cells(excel_row, 31).Value
SHEL_action = ObjExcel.Cells(excel_row, 32).Value
SHEL_member = ObjExcel.Cells(excel_row, 33).Value
subsidized_indicator = ObjExcel.Cells(excel_row, 34).Value
shared_indicator = ObjExcel.Cells(excel_row, 35).Value
paid_to_memb = ObjExcel.Cells(excel_row, 36).Value
paid_to_name = ObjExcel.Cells(excel_row, 37).Value
rent_amt = ObjExcel.Cells(excel_row, 38).Value
rent_proof = ObjExcel.Cells(excel_row, 39).Value
lot_rent_amt = ObjExcel.Cells(excel_row, 40).Value
lot_rent_proof = ObjExcel.Cells(excel_row, 41).Value
mortgage_amt = ObjExcel.Cells(excel_row, 42).Value
mortgage_proof = ObjExcel.Cells(excel_row, 43).Value
insurance_amt = ObjExcel.Cells(excel_row, 44).Value
insurance_proof = ObjExcel.Cells(excel_row, 45).Value
taxes_amt = ObjExcel.Cells(excel_row, 46).Value
taxes_proof = ObjExcel.Cells(excel_row, 47).Value
room_amt = ObjExcel.Cells(excel_row, 48).Value
room_proof = ObjExcel.Cells(excel_row, 49).Value
garage_amt = ObjExcel.Cells(excel_row, 50).Value
garage_proof = ObjExcel.Cells(excel_row, 51).Value
subsidy_amt = ObjExcel.Cells(excel_row, 52).Value
subsidy_proof = ObjExcel.Cells(excel_row, 53).Value
HEST_action = ObjExcel.Cells(excel_row, 54).Value
HEST_heat_AC_indicator = ObjExcel.Cells(excel_row, 55).Value
HEST_electric_indicator = ObjExcel.Cells(excel_row, 56).Value
HEST_phone_indicator = ObjExcel.Cells(excel_row, 57).Value
JOBS_action = ObjExcel.Cells(excel_row, 58).Value
JOBS_member = ObjExcel.Cells(excel_row, 59).Value
JOBS_location = ObjExcel.Cells(excel_row, 60).Value
JOBS_start_date_month = ObjExcel.Cells(excel_row, 61).Value
JOBS_start_date_day = ObjExcel.Cells(excel_row, 62).Value
JOBS_start_date_year = ObjExcel.Cells(excel_row, 63).Value
JOBS_proof = ObjExcel.Cells(excel_row, 64).Value
pay_freq = ObjExcel.Cells(excel_row, 65).Value
pay_amount = ObjExcel.Cells(excel_row, 66).Value
hours_per_check = ObjExcel.Cells(excel_row, 67).Value
update_PIC =  ObjExcel.Cells(excel_row, 68).Value
update_future_months = ObjExcel.Cells(excel_row, 69).Value
ACCT_action = ObjExcel.Cells(excel_row, 70).Value
ACCT_member = ObjExcel.Cells(excel_row, 71).Value
ACCT_type = ObjExcel.Cells(excel_row, 72).Value
ACCT_number = ObjExcel.Cells(excel_row, 73).Value
ACCT_location = ObjExcel.Cells(excel_row, 74).Value
ACCT_balance = ObjExcel.Cells(excel_row, 75).Value
ACCT_as_of_month = ObjExcel.Cells(excel_row, 76).Value
ACCT_as_of_day = ObjExcel.Cells(excel_row, 77).Value
ACCT_as_of_year = ObjExcel.Cells(excel_row, 78).Value
cash_count_status = ObjExcel.Cells(excel_row, 79).Value
SNAP_count_status = ObjExcel.Cells(excel_row, 80).Value
HC_count_status = ObjExcel.Cells(excel_row, 81).Value
WREG_action = ObjExcel.Cells(excel_row, 82).Value
WREG_member_array = split(ObjExcel.Cells(excel_row, 83).Value, ", ") 
FSET_status = ObjExcel.Cells(excel_row, 84).Value
defer_FSET_indicator = ObjExcel.Cells(excel_row, 85).Value
ABAWD_status = ObjExcel.Cells(excel_row, 86).Value
GA_basis = ObjExcel.Cells(excel_row, 87).Value


'Date calculation
current_month = datepart("m", date)
if len(current_month) < 2 then current_month = "0" & current_month
current_day = datepart("d", date)
if len(current_day) < 2 then current_day = "0" & current_day
current_year = datepart("yyyy", date) - 2000

'Getting to PND2
back_to_self
EMWriteScreen "REPT", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "pnd2", 21, 70
transmit

'Checking for PND2. If not on PND2 it'll stop.
EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then script_end_procedure("Not on PND2")

'Setting the MAXIS row to look at.
MAXIS_row = 7

'Grabs the footer month and year because it might need the original if we're updating future months in cases.
EMReadScreen PND2_footer_month, 2, 20, 55
EMReadScreen PND2_footer_year, 2, 20, 58

'OPERATES AS A DO...LOOP TO UPDATE EVERY CASE ON PND2
Do
	'Writing the original footer month in case we updated future months
	EMWriteScreen PND2_footer_month, 20, 55
	EMWriteScreen PND2_footer_year, 20, 58

	'gets into STAT for the case
	EMWriteScreen "s", MAXIS_row, 3
	transmit

	If ABPS_action <> "" then 'Updating ABPS----------------------------------------------------------------------------------------------------
		If both_parents_in_HH = False then
			EMWriteScreen "ABPS", 20, 71
			EMWriteScreen ABPS_action, 20, 79
			transmit
      
			EMWriteScreen "01", 4, 47
			EMWriteScreen "Y", 4, 73
			EMWriteScreen "N", 5, 47
			EMWriteScreen ABPS_last_name, 10, 30
			EMWriteScreen ABPS_first_name, 10, 63
			EMWriteScreen ABPS_gender, 11, 80
			row = 15
     
			For each kid in kids_to_add
				EMWriteScreen kid, row, 35
				EMWriteScreen "1", row, 53
				EMWriteScreen "1", row, 67
				row = row + 1
			Next
      
			transmit
		End if
	End if

	If ACCT_action <> "" then 'Updating ACCT----------------------------------------------------------------------------------------------------
		EMWriteScreen "ACCT", 20, 71
		EMWriteScreen ACCT_member, 20, 76
		EMWriteScreen ACCT_action, 20, 79
		transmit
  
		If ACCT_action <> "nn" then PF9
  
		'Clears out existing info
		EMSendKey string(79, "_")
  
		EMWriteScreen ACCT_type, 6, 44
		EMWriteScreen ACCT_number, 7, 44
		EMWriteScreen ACCT_location, 8, 44
		EMWriteScreen ACCT_balance, 10, 46
		EMWriteScreen "5", 10, 63
		EMWriteScreen ACCT_as_of_month, 11, 44
		EMWriteScreen ACCT_as_of_day, 11, 47
		EMWriteScreen ACCT_as_of_year, 11, 50
  
		EMWriteScreen cash_count_status, 14, 50
		EMWriteScreen SNAP_count_status, 14, 57
		EMWriteScreen HC_count_status, 14, 64
		EMWriteScreen "N", 15, 44
    
		transmit
	End if

	If EATS_action <> "" then 'Updating EATS----------------------------------------------------------------------------------------------------
		EMWriteScreen "EATS", 20, 71
		EMWriteScreen EATS_action, 20, 79
		transmit

		If EATS_action <> "nn" then PF9
  
		EMWriteScreen PP_together, 4, 72
		EMWriteScreen "N", 5, 72
  
		If PP_together = "N" then
			EMWriteScreen "01", 13, 28
			col = 39
			For each EATS_member in EATS_member_array  
				EMWriteScreen EATS_member, 13, col
				col = col + 4
			Next
    
			col = 39
			EMWriteScreen "02", 14, 28
			For each EATS_non_member in EATS_non_member_array  
				EMWriteScreen EATS_non_member, 14, col
				col = col + 4
			Next
		End if
  
		transmit
	End if

	If PARE_action <> "" then 'Updating PARE----------------------------------------------------------------------------------------------------
		EMWriteScreen "PARE", 20, 71
		EMWriteScreen PARE_action, 20, 79
		transmit
    
		row = 8
    
		For each kid in kids_to_add
			EMWriteScreen kid, row, 24
			EMWriteScreen "1", row, 53
			EMWriteScreen "OT", row, 71
			row = row + 1
		Next

		transmit
    
		If both_parents_in_HH = True then
			EMWriteScreen "PARE", 20, 71
			EMWriteScreen "02", 20, 76
			EMWriteScreen PARE_action, 20, 79
			transmit
		      row = 8
			For each kid in kids_to_add
				EMWriteScreen kid, row, 24
				EMWriteScreen "1", row, 53
				EMWriteScreen "OT", row, 71
				row = row + 1
			Next

			transmit
		End if

		If stepparent_in_HH = True then
			EMWriteScreen "PARE", 20, 71
			EMWriteScreen "02", 20, 76
			EMWriteScreen PARE_action, 20, 79
			transmit
			row = 8

			For each kid in kids_to_add
				EMWriteScreen kid, row, 24
				EMWriteScreen "2", row, 53
				EMWriteScreen "OT", row, 71
				row = row + 1
			Next

			transmit
		End if
	End if

	If EMPS_action <> "" then 'Updating EMPS----------------------------------------------------------------------------------------------------
		EMWriteScreen "EMPS", 20, 71
		EMWriteScreen EMPS_action, 20, 79
		transmit
		If EMPS_action <> "nn" then PF9

		EMWriteScreen fin_orient_dt_month, 5, 39
		EMWriteScreen fin_orient_dt_day, 5, 42
		EMWriteScreen fin_orient_dt_year, 5, 45
		EMWriteScreen "n", 8, 76
		EMWriteScreen "n", 9, 76
		EMWriteScreen "n", 10, 76
		EMWriteScreen "no", 11, 76
		EMWriteScreen full_time_care_of_child_under_1, 12, 76
		EMWriteScreen "n", 13, 76

		If full_time_care_of_child_under_1 = "Y" then
			EMWriteScreen "x", 12, 39
			transmit
			EMWriteScreen EMPS_Exemption_Care_Of_A_Child_Under_One_month, 7, 22
			EMWriteScreen EMPS_Exemption_Care_Of_A_Child_Under_One_year, 7, 27
			transmit
			PF3
		End if

		transmit
	End if

	If SHEL_action <> "" then 'Updating SHEL----------------------------------------------------------------------------------------------------
		EMWriteScreen "shel", 20, 71
		EMWriteScreen SHEL_member, 20, 76
		EMWriteScreen SHEL_action, 20, 79
		transmit
		If SHEL_action <> "nn" then PF9
		EMSetCursor 6, 42
		EMSendKey(string(189, "_"))
		EMWriteScreen subsidized_indicator, 6, 42
		EMWriteScreen shared_indicator, 6, 60
		EMWriteScreen paid_to_MEMB, 7, 42
		EMWriteScreen paid_to_name, 7, 46
		EMWriteScreen rent_amt, 11, 56
		EMWriteScreen rent_proof, 11, 67
		EMWriteScreen lot_rent_amt, 12, 56
		EMWriteScreen lot_rent_proof, 12, 67
		EMWriteScreen mortgage_amt, 13, 56
		EMWriteScreen mortgage_proof, 13, 67
		EMWriteScreen insurance_amt, 14, 56
		EMWriteScreen insurance_proof, 14, 67
		EMWriteScreen taxes_amt, 15, 56
		EMWriteScreen taxes_proof, 15, 67
		EMWriteScreen room_amt, 16, 56
		EMWriteScreen room_proof, 16, 67
		EMWriteScreen garage_amt, 17, 56
		EMWriteScreen garage_proof, 17, 67
		EMWriteScreen subsidy_amt, 18, 56
		EMWriteScreen subsidy_proof, 18, 67
		transmit
	End if

	If HEST_action <> "" then 'Updating HEST----------------------------------------------------------------------------------------------------
		EMWriteScreen "HEST", 20, 71
		EMWriteScreen HEST_action, 20, 79
		transmit
		If HEST_action <> "nn" then PF9
		EMSetCursor 6, 40
		EMSendKey(string(52, "_"))
		EMWriteScreen "01", 6, 40
		EMWriteScreen "01", 7, 40
		EMWriteScreen "01", 7, 43
		EMWriteScreen "01", 7, 46
		EMWriteScreen HEST_heat_AC_indicator, 13, 60
		If HEST_heat_AC_indicator <> "" then EMWriteScreen "01", 13, 68
		EMWriteScreen HEST_electric_indicator, 14, 60
		If HEST_electric_indicator <> "" then EMWriteScreen "01", 15, 68
		EMWriteScreen HEST_phone_indicator, 15, 60
		If HEST_phone_indicator <> "" then EMWriteScreen "01", 15, 68
		transmit
	End if

	If WREG_action <> "" then 'Updating WREG----------------------------------------------------------------------------------------------------
		For each WREG_member in WREG_member_array  
			EMWriteScreen "WREG", 20, 71
			EMWriteScreen WREG_member, 20, 76
			EMWriteScreen WREG_action, 20, 79
			transmit
			If WREG_action <> "nn" then PF9
			If WREG_member = "01" then EMWriteScreen "Y", 6, 68
			If WREG_member <> "01" then EMWriteScreen "N", 6, 68
			EMWriteScreen FSET_status, 8, 50
			EMWriteScreen defer_FSET_indicator, 8, 80
			EMWriteScreen ABAWD_status, 13, 50
			EMWriteScreen GA_basis, 15, 50
			transmit
			transmit
		Next
	End if

	If JOBS_action <> "" then 'Updating JOBS, does this last for multi-month purposes----------------------------------------------------------------------------------------------------
		EMWriteScreen "JOBS", 20, 71
		EMWriteScreen JOBS_member, 20, 76
		EMWriteScreen JOBS_action, 20, 79
		transmit
		If JOBS_action <> "nn" then PF9
		EMReadScreen footer_month_year, 5, 20, 55
		footer_month_year = replace(footer_month_year, " ", "/01/")
		'Clears out existing info
		EMSendKey string(214, "_")
		'Now it figures out what the paydays would be. It assumes a friday payday.
		date_start_plus_1 = DateAdd("d", 1, footer_month_year) 
		date_start_plus_2 = DateAdd("d", 2, footer_month_year) 
		date_start_plus_3 = DateAdd("d", 3, footer_month_year) 
		date_start_plus_4 = DateAdd("d", 4, footer_month_year) 
		date_start_plus_5 = DateAdd("d", 5, footer_month_year) 
		date_start_plus_6 = DateAdd("d", 6, footer_month_year) 
		If Weekday(footer_month_year, 0) = 6 then first_payday = (footer_month_year)
		If Weekday(date_start_plus_1, 0) = 6 then first_payday = (date_start_plus_1)
		If Weekday(date_start_plus_2, 0) = 6 then first_payday = (date_start_plus_2)
		If Weekday(date_start_plus_3, 0) = 6 then first_payday = (date_start_plus_3)
		If Weekday(date_start_plus_4, 0) = 6 then first_payday = (date_start_plus_4)
		If Weekday(date_start_plus_5, 0) = 6 then first_payday = (date_start_plus_5)
		If Weekday(date_start_plus_6, 0) = 6 then first_payday = (date_start_plus_6)
		If pay_freq = 1 then
			second_payday = ""
			third_payday = ""
			fourth_payday = ""
			fifth_payday = ""
		End if
		If pay_freq = 2 then
			second_payday = dateadd("d", 15, first_payday)
			third_payday = ""
			fourth_payday = ""
			fifth_payday = ""
		End if
		If pay_freq = 3 then
			second_payday = dateadd("d", 14, first_payday)
			third_payday = dateadd("d", 14, second_payday)
			fourth_payday = ""
			fifth_payday = ""
			If datepart("m", third_payday) <> datepart("m", first_payday) then third_payday = ""
		End if
		If pay_freq = 4 then
			second_payday = dateadd("d", 7, first_payday)
			third_payday = dateadd("d", 7, second_payday)
			fourth_payday = dateadd("d", 7, third_payday)
			fifth_payday = dateadd("d", 7, fourth_payday)
			If datepart("m", fifth_payday) <> datepart("m", first_payday) then fifth_payday = ""
		End if
		If first_payday <> "" then payday_array = payday_array & " " & first_payday
		If second_payday <> "" then payday_array = payday_array & " " & second_payday
		If third_payday <> "" then payday_array = payday_array & " " & third_payday
		If fourth_payday <> "" then payday_array = payday_array & " " & fourth_payday
		If fifth_payday <> "" then payday_array = payday_array & " " & fifth_payday
		payday_array = split(trim(payday_array))    
		'Now it writes the payday info into MAXIS
		row = 12
		For each payday in payday_array
			payday = cdate(payday)
			payday_month = datepart("m", payday)
			If len(payday_month) = 1 then payday_month = "0" & payday_month
			payday_day = datepart("d", payday)
			If len(payday_day) = 1 then payday_day = "0" & payday_day
			payday_year = datepart("yyyy", payday) - 2000
			EMWriteScreen payday_month, row, 54
			EMWriteScreen payday_day, row, 57
			EMWriteScreen payday_year, row, 60
			EMWriteScreen "________", row, 67
			EMWriteScreen pay_amount, row, 67
			row = row + 1
		Next
		'Writes hours into MAXIS
		monthly_hours = hours_per_check * (ubound(payday_array) + 1)
		EMWriteScreen "___", 18, 72
		EMWriteScreen monthly_hours, 18, 72
		'Writes info about the job into MAXIS.
		EMWriteScreen "W", 5, 38
		EMWriteScreen JOBS_proof, 6, 38
		EMWriteScreen JOBS_location, 7, 42
		EMWriteScreen JOBS_start_date_month, 9, 35
		EMWriteScreen JOBS_start_date_day, 9, 38
		EMWriteScreen JOBS_start_date_year, 9, 41
		EMWriteScreen pay_freq, 18, 35
		If update_PIC = True then 
			EMWriteScreen "x", 19, 38
			transmit
			EMWriteScreen current_month, 5, 34
			EMWriteScreen current_day, 5, 37
			EMWriteScreen current_year, 5, 40
			EMWriteScreen pay_freq, 5, 64
			EMWriteScreen hours_per_check, 8, 64
			EMWriteScreen pay_amount/hours_per_check, 9, 66
			transmit
			transmit
		End if
    
		If update_future_months = True then 
			Do
				transmit
				EMWriteScreen "bgtx", 20, 71
				transmit
				EMWriteScreen "y", 16, 54
				transmit
				EMWriteScreen "jobs", 20, 71
				EMWriteScreen JOBS_member, 20, 76
				If JOBS_action = "nn" then
					EMWriteScreen "01", 20, 79
				Else
					EMWriteScreen JOBS_action, 20, 79
				End if
				transmit
				PF9
 
				EMReadScreen current_footer_month, 2, 20, 55
				EMReadScreen current_footer_year, 2, 20, 58

				JOBS_line_row = 12
				Do
					EMReadScreen JOBS_line_day, 2, JOBS_line_row, 57
					If isnumeric(JOBS_line_day) = False then exit do
					If isdate(current_footer_month & "/" & JOBS_line_day & "/" & current_footer_year) = True Then 
						EMWriteScreen current_footer_month, JOBS_line_row, 54
						EMWriteScreen current_footer_year, JOBS_line_row, 60
					Else
						EMWriteScreen "__", JOBS_line_row, 54
						EMWriteScreen "__", JOBS_line_row, 57
						EMWriteScreen "__", JOBS_line_row, 60
						EMWriteScreen "________", JOBS_line_row, 67
					End if
					JOBS_line_row = JOBS_line_row + 1
				Loop until JOBS_line_row = 17
				first_of_current_month = current_footer_month & "/01/" & current_footer_year
				first_of_next_month = datepart("m", dateadd("m", 1, date)) & "/01/" & datepart("yyyy", dateadd("m", 1, date))
			Loop until cdate(first_of_next_month) = cdate(first_of_current_month)

			EMWriteScreen "x", 19, 54
			transmit
 
			EMWriteScreen "________", 11, 63
			EMWriteScreen pay_amount, 11, 63
			transmit
			transmit
		End if
 
		transmit

		payday_array = ""
	End if  

	Do 'Exiting the case----------------------------------------------------------------------------------------------------
		PF3
		EMReadScreen PND2_check, 4, 2, 52
		If PND2_check = "LF) " then script_end_procedure("error")
	Loop until PND2_check = "PND2"
  
	MAXIS_row = MAXIS_row + 1
Loop until MAXIS_row = amt_of_times_to_run + 7