'LOADING ROUTINE FUNCTIONS===========================================================================================
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
'===================================================================================================================

EMConnect ""

'Dialogs===================================================================================================================
	'This dialog is for the WREG exemptions.-----------------------------------------------------------------------
BeginDialog wreg_exemptions, 0, 0, 311, 250, "ABAWD Screening Tool"
  CheckBox 5, 20, 260, 10, "...Permanently or Temporarily disabled or incapacitated (at least 30 days)?", wreg_disa
  CheckBox 5, 35, 270, 10, "...responsible for the care of a disabled household member?", care_of_hh_memb
  CheckBox 5, 50, 275, 10, "...age 60 or older?", age_sixty
  CheckBox 5, 65, 290, 15, "...under the age of 16?", under_sixteen
  CheckBox 5, 85, 275, 10, "...aged 16 or 17 living w/ parent or caregiver?", sixteen_seventeen
  CheckBox 5, 100, 275, 10, "...responsible for the care of a child under 6?", care_child_six
  CheckBox 5, 115, 255, 10, "...employed 30 hours per week or earning at least $935.25/month gross?", employed_thirty
  CheckBox 5, 130, 255, 10, "...receiving or applied for unemployment insurance?", unemployment
  CheckBox 5, 145, 255, 10, "...enrolled in school, training program, or higher education?", enrolled_school
  CheckBox 5, 160, 305, 10, "...participating in a chemical dependency treatment program (not AA or Narc. Anonymous)?", CD_program
  CheckBox 5, 175, 300, 10, "...receiving MFIP?", receiving_MFIP
  CheckBox 5, 190, 305, 10, "...receiving or pending for Diversionary Work Program or Work Benefit?", receiving_DWP_WB
  CheckBox 5, 205, 300, 10, "...applied for SSI (cannot be appealing denial)?", applied_SSI
  ButtonGroup ButtonPressed
    PushButton 205, 235, 50, 15, "NEXT", next_button
    CancelButton 260, 235, 50, 15
  Text 5, 5, 85, 10, "Is the client..."
EndDialog

	'This dialog gets the client's case number.---------------------------------------------------------------------
BeginDialog get_case_number, 0, 0, 181, 100, "ABAWD Screening Tool"
  Text 10, 15, 50, 10, "Case Number: "
  EditBox 90, 10, 50, 15, case_number
  Text 10, 35, 70, 10, "Member Number:"
  EditBox 90, 30, 30, 15, member_number
  Text 10, 55, 75, 10, "Sign your Case Note:"
  EditBox 90, 50, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 45, 85, 50, 15, "Next", next_button
    CancelButton 95, 85, 50, 15
EndDialog


	'This dialog is for the ABAWD exemptions and is used if the CL does not have a WREG exemption.---------------------
BeginDialog abawd_exemptions, 0, 0, 241, 180, "ABAWD Screening Tool"
  CheckBox 5, 20, 230, 15, "...residing in a waivered area", waiver
  CheckBox 5, 35, 185, 15, "...younger than 18 OR 50 or older?", age_exempt
  CheckBox 5, 50, 210, 15, "...medically certified as pregnant?", cert_preg
  CheckBox 5, 65, 210, 15, "...working at least 20 hours per week or 80 hours per month?", working_20
  CheckBox 5, 80, 230, 15, "...receiving RCA or GA?", receiving_cash
  CheckBox 5, 95, 240, 15, "...responsible for the care of a dependent child?", dependent_child
  CheckBox 5, 110, 240, 15, "...a Work Experience participant?", work_exp
  CheckBox 5, 125, 240, 15, "...participating in an approved Employment and Training program?", approved_ET
  ButtonGroup ButtonPressed
    PushButton 45, 160, 50, 15, "Previous", previous_button
    PushButton 100, 160, 50, 15, "Next", next_button
    CancelButton 180, 160, 50, 15
  Text 5, 5, 245, 10, "Is the client..."
EndDialog

	'This dialog allows the screener to ask if the CL has earned an additional 3-month period of ABAWD-counted months---------
BeginDialog earn_additional_months, 0, 0, 366, 95, "ABAWD Screening Tool"
  CheckBox 5, 30, 355, 15, "Has the CL worked at least 80 hours in a month SINCE closing for using their last ABAWD-counted month?", worked_80_since_closing
  CheckBox 5, 50, 355, 15, "Has the CL used a second period of ABAWD-counted months?", has_used_second_period
  ButtonGroup ButtonPressed
    PushButton 165, 75, 50, 15, "Finish", finish_button
    PushButton 110, 75, 50, 15, "Previous", previous_button
    CancelButton 220, 75, 50, 15
  Text 5, 10, 295, 15, "Please navigate to the ABAWD Tracking Record for the appropriate member..."
EndDialog


	'This dialog gets the worker's signature and allows the OSA to enter any comments for the case worker.----------------------
	'The idea being that if the OSA notices irregularities or unusualness (word?) in the ABAWD tracking panel, it---------------
	'can be reported to the worker or the worker can be directed to look deeper into the ABAWD tracking.------------------------
BeginDialog get_worker_comments, 0, 0, 166, 105, "ABAWD Screening Tool"
  EditBox 5, 50, 155, 15, worker_comment
  ButtonGroup ButtonPressed
    PushButton 20, 75, 50, 15, "OK", OK_button
    CancelButton 90, 75, 50, 15
  Text 5, 10, 150, 10, "Case noting CL interaction."
  Text 5, 25, 160, 20, "Any additional comments, please enter here. Press ENTER to complete and Case Note."
EndDialog

	'Cancel discount double-check-----------------------------------------------------------------------------------------------
BeginDialog cancel_dialog, 0, 0, 126, 61, "Cancel dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 100, 15, "yes, cancel this case note", yes
    PushButton 10, 40, 110, 15, "no, do not cancel this case note", no
  Text 10, 5, 110, 10, "Are you sure you want to cancel?"
EndDialog

'FUNCTIONS========================================================================================
function how_many_abawd_months(abawd_counted_months)
  call navigate_to_screen("stat", "wreg")
    EMWriteScreen member_number, 20, 76
    transmit
    EMSetCursor 13, 57
    EMSendKey "X"
    transmit
  current_month = datepart("m",Date())
  bene_mo_col = (15 + (4*current_month))
  bene_yr_row = 10
  month_count = 0
  DO
    EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
      IF is_counted_month = "X" or is_counted_month = "M" THEN abawd_counted_months = abawd_counted_months + 1
    bene_mo_col = bene_mo_col - 4
      IF bene_mo_col = 15 THEN
        bene_yr_row = bene_yr_row - 1
        bene_mo_col = 63
      END IF
    month_count = month_count + 1
  LOOP until month_count = 36
  PF3
END function

Function case_note_and_end
	Dialog get_worker_comments
		IF ButtonPressed = 0 THEN
		      Dialog cancel_script
		      IF ButtonPressed = yes then stopscript
	      END IF
	PF3	
   DO
	call navigate_to_screen("case", "note")
	maxis_check_function
   LOOP until MAXIS_check = "MAXIS"
	PF9
	EMSetCursor 4, 3
      case_note_header = "***Member " & member_number & " has been screened for ABAWD***"
	call write_new_line_in_case_note(case_note_header)
	call write_new_line_in_case_note(abawd_status)
	  IF has_used_second_period = 1 THEN call write_new_line_in_case_note("* CL has used 2nd period of ABAWD eligibility.")
	  IF worked_80_since_closing = 1 AND has_used_second_period <> 1 THEN call write_new_line_in_case_note("* CL has earned additional 3-month period of ABAWD eligibility.")
	  IF wreg_disa = 1 THEN call write_new_line_in_case_note("* Client states they are disabled")
	  IF care_of_hh_memb = 1 THEN call write_new_line_in_case_note("* Client states they are responsible for care of a disabled unit member")
	  IF age_sixty = 1 THEN call write_new_line_in_case_note("* Client is over 60.")
	  IF under_sixteen = 1 THEN call write_new_line_in_case_note("* Client states they are under 16.")
	  IF sixteen_seventeen = 1 THEN call write_new_line_in_case_note("* Client states they are age 16 or 17 and living with a parent or caretaker")
	  IF care_child_six = 1 THEN call write_new_line_in_case_note("* Client states they are responsible for the care of a child less than age 6.")
	  IF employed_thirty = 1 THEN 
		call write_new_line_in_case_note("* Client states they are employed 30 hours per week or equivalent to 30 hours") 
	      call write_new_line_in_case_note("  a week at minimum wage.")
	  End If
	  IF unemployment = 1 THEN call write_new_line_in_case_note("* Client states they are receiving or applied for unemployment insurance.")
	  IF enrolled_school = 1 THEN call write_new_line_in_case_note("* Client states they are enrolled in school/training 1/2 time.")
	  IF CD_program = 1 THEN 
		call write_new_line_in_case_note("* Client states they are enrolled in a sanctioned chemical dependency")
		call write_new_line_in_case_note("  treatment program.")
	  End If
	  IF receiving_MFIP = 1 THEN call write_new_line_in_case_note("* Client states they are a MFIP recipient.")
	  IF receiving_DWP_WB = 1 THEN call write_new_line_in_case_note("* Client states they are a DWP/WB recipient.")
	  IF age_exempt = 1 THEN call write_new_line_in_case_note("* Client states they are under 18 or over 50 years old")
	  IF cert_preg = 1 THEN call write_new_line_in_case_note("* Client states certified as pregnant")
	  IF working_20 = 1 THEN call write_new_line_in_case_note("* Client states they are employed 20 hours per week")
	  IF dependent_child = 1 THEN 
		call write_new_line_in_case_note("* Client states they are responsible for the care of a dependent child in the")
		call write_new_line_in_case_note("  household.")
	  End If 
	  IF work_exp = 1 THEN call write_new_line_in_case_note("* Client states they are participatiing in work experience program")
	  IF approved_ET = 1 THEN call write_new_line_in_case_note("* Client states they are participating in employment and training program")
	  IF waiver = 1 THEN call write_new_line_in_case_note("* Client states they are residing in a waiver area")
	  IF receiving_cash = 1 THEN call write_new_line_in_case_note("* Client states they are a RCA or GA recipient")
		IF worker_comment <> "" THEN 
			worker_comment = "* ADDITIONAL NOTE: " + worker_comment			
			call write_new_line_in_case_note(worker_comment)
		END IF
	call write_new_line_in_case_note("---")
	call write_new_line_in_case_note(worker_signature)
	stopscript
END Function

function maxis_check_function
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again."
END function


'THE SCRIPT=======================================================================================
DO
  dialog get_case_number
    IF ButtonPressed = 0 THEN stopscript
    IF IsNumeric(case_number) = FALSE THEN
	case_number = ""
	MSGBox("Your case number is not a valid case number.")
    END IF
    IF len(case_number) > 8 THEN 
	case_number = ""
	MSGBox("Your case number is not a valid case number.")
    END IF
    IF len(member_number) = 1 THEN member_number = "0" & member_number
    IF len(member_number) > 2 THEN MSGBox("Invalid member number")
    maxis_check_function
    IF worker_signature = "" THEN MSGBox("Please sign your case note.")
LOOP until case_number <> "" and worker_signature <> "" and len(member_number) = 2
   
DO
 call navigate_to_screen("stat", "wreg")
 maxis_check_function
LOOP until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

call how_many_abawd_months(abawd_counted_months)

DO
	Dialog wreg_exemptions
		IF ButtonPressed = 0 THEN
		  Dialog cancel_dialog
		  IF ButtonPressed = yes then stopscript
		End If	
		IF under_sixteen = 1 or wreg_disa = 1 or care_of_hh_memb = 1 or age_sixty = 1 or sixteen_seventeen = 1 or care_child_six = 1 or employed_thirty = 1 or unemployment = 1 or enrolled_school = 1 or CD_program = 1 or receiving_MFIP = 1 or receiving_DWP_WB = 1 or applied_SSI = 1 THEN wreg_exempt = true
		IF under_sixteen = 0 and wreg_disa = 0 and care_of_hh_memb = 0 and age_sixty = 0 and sixteen_seventeen = 0 and care_child_six = 0 and employed_thirty = 0 and unemployment = 0 and enrolled_school = 0 and CD_program = 0 and receiving_MFIP = 0 and receiving_DWP_WB = 0 and applied_SSI = 0 THEN wreg_exempt = false
		IF wreg_exempt = TRUE THEN abawd_status = "* CL is NOT an ABAWD."
	IF wreg_exempt = true THEN call case_note_and_end
  DO
	  Dialog abawd_exemptions
		IF ButtonPressed = 0 THEN
		  Dialog cancel_dialog
		  IF ButtonPressed = yes then stopscript
		End If	
  	      IF waiver = 1 or age_exempt = 1 or cert_preg = 1 or working_20 = 1 or receiving_cash = 1 or dependent_child = 1 or work_exp = 1 or approved_ET = 1 THEN 
		  cl_has_abawd_exemption = true
		  abawd_status = "* CL is NOT an ABAWD."
		End If
		IF ButtonPressed = previous_button then exit do
		IF cl_has_abawd_exemption = true THEN call case_note_and_end
		IF cl_has_abawd_exemption <> true THEN abawd_status = "* CL is ABAWD and has used " & abawd_counted_months & " months of SNAP eligibility."
    DO
		Dialog earn_additional_months
		  IF ButtonPressed = 0 THEN
		    Dialog cancel_dialog
		    IF ButtonPressed = yes then stopscript
		END IF
		IF ButtonPressed = previous_button then exit do
		call case_note_and_end
    LOOP until ButtonPressed = -1
  LOOP until ButtonPressed = -1
LOOP until ButtonPressed = -1




