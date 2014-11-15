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
BeginDialog wreg_exemptions, 0, 0, 311, 225, "ABAWD Screening Tool"
  CheckBox 5, 20, 260, 10, "...Permanently or Temporarily disabled or incapacitated (at least 30 days)?", wreg_disa
  CheckBox 5, 35, 270, 10, "...responsible for the care of a disabled household member?", care_of_hh_memb
  CheckBox 5, 50, 275, 10, "...age 60 or older?", age_sixty
  CheckBox 5, 65, 275, 10, "...aged 16 or 17 living w/ parent or caregiver?", sixteen_seventeen
  CheckBox 5, 80, 275, 10, "...responsible for the care of a child under 6?", care_child_six
  CheckBox 5, 95, 255, 10, "...employed 30 hours per week or earning at least $935.25/month gross?", employed_thirty
  CheckBox 5, 110, 255, 10, "...receiving or applied for unemployment insurance?", unemployment
  CheckBox 5, 125, 255, 10, "...enrolled in school, training program, or higher education?", enrolled_school
  CheckBox 5, 140, 305, 10, "...participating in a chemical dependency treatment program (not AA or Narc. Anonymous)?", CD_program
  CheckBox 5, 155, 300, 10, "...receiving MFIP?", receiving_MFIP
  CheckBox 5, 170, 305, 10, "...receiving or pending for Diversionary Work Program or Work Benefit?", receiving_DWP_WB
  CheckBox 5, 185, 300, 10, "...applied for SSI (cannot be appealing denial)?", applied_SSI
  ButtonGroup ButtonPressed
    PushButton 205, 205, 50, 15, "NEXT", next_button
    CancelButton 260, 205, 50, 15
  Text 5, 5, 85, 10, "Is the client..."
EndDialog

	'This dialog gets the client's case number.---------------------------------------------------------------------
BeginDialog get_case_number, 0, 0, 181, 80, "ABAWD Screening Tool"
  Text 10, 15, 50, 10, "Case Number: "
  EditBox 90, 10, 50, 15, case_number
  Text 10, 35, 75, 10, "Sign your Case Note:"
  EditBox 90, 30, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 40, 60, 50, 15, "Next", next_button
    CancelButton 95, 60, 50, 15
EndDialog

	'This dialog is for the ABAWD exemptions and is used if the CL does not have a WREG exemption.---------------------
BeginDialog abawd_exemptions, 0, 0, 241, 180, "ABAWD Screening Tool"
  CheckBox 5, 20, 230, 15, "...WREG exempt (should autofill from previous screen)?", wreg_exempt
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
	'This dialog allows the OSA to enter the number of ABAWD months the CL has used if the CL is an ABAWD and-----------------
	'had a SNAP case open previously.-----------------------------------------------------------------------------------------
BeginDialog get_abawd_months, 0, 0, 236, 80, "ABAWD Screening Tool"
  ButtonGroup ButtonPressed
    PushButton 65, 60, 50, 15, "Finish", finish_button
    PushButton 10, 60, 50, 15, "Previous", previous_button
    CancelButton 120, 60, 50, 15
  Text 10, 25, 180, 30, "How many ABAWD counted months has the client used (count all months in past 3 years coded X or M)..."
  DropListBox 205, 25, 25, 15, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3", abawd_months
  CheckBox 10, 5, 225, 15, "If checked, CL is not ABAWD. Please hit ENTER.", Check1
EndDialog

	'This dialog gets the worker's signature and allows the OSA to enter any comments for the case worker.----------------------
	'The idea being that if the OSA notices irregularities or unusualness (word?) in the ABAWD tracking panel, it---------------
	'can be reported to the worker or the worker can be directed to look deeper into the ABAWD tracking.------------------------
BeginDialog get_worker_comments, 0, 0, 166, 105, "ABAWD Screening Tool"
  EditBox 5, 50, 155, 15, worker_comment
  ButtonGroup ButtonPressed
    PushButton 20, 75, 50, 15, "Previous", previous_button
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
Function what_is_abawd_status(abawd_status)
IF wreg_exempt = 0 and age_exempt = 0 and cert_preg = 0 and working_20 = 0 and receiving_cash = 0 and dependent_child = 0 and work_exp = 0 and approved_ET = 0 THEN abawd_exempt = true  '1 = true...just used 1 for the sake of fewer key strokes
	IF abawd_exempt = true and abawd_months = 0 THEN abawd_status = "   -CL is ABAWD and has not used any ABAWD-counted months."
	IF abawd_exempt = true and abawd_months = 1 THEN abawd_status = "   -CL is ABAWD and has used 1 ABAWD-counted month."
	IF abawd_exempt = true and abawd_months = 2 THEN abawd_status = "   -CL is ABAWD and has used 2 ABAWD-counted months."
	IF abawd_exempt = true and abawd_months = 3 THEN abawd_status = "   -CL is ABAWD and has used ALL 3 ABAWD-counted months."
IF wreg_exempt = 1 or age_exempt = 1 or cert_preg = 1 or working_20 = 1 or receiving_cash = 1 or dependent_child = 1 or work_exp = 1 or approved_ET = 1 THEN abawd_status = "   -CL is NOT an ABAWD."
END Function

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
    IF worker_signature = "" THEN MSGBox("Please sign your case note.")
LOOP until case_number <> "" and worker_signature <> ""
   
call navigate_to_screen("stat", "wreg")
	EMSetCursor 13, 57
	EMSendKey "X"
	transmit

DO
	Dialog wreg_exemptions
		IF ButtonPressed = 0 THEN
		  Dialog cancel_dialog
		  IF ButtonPressed = yes then stopscript
		End If	
		IF wreg_disa = 1 or care_of_hh_memb = 1 or age_sixty = 1 or sixteen_seventeen = 1 or care_child_six = 1 or employed_thirty = 1 or unemployment = 1 or enrolled_school = 1 or CD_program = 1 or receiving_MFIP = 1 or receiving_DWP_WB = 1 or applied_SSI = 1 THEN wreg_exempt = 1
		IF wreg_disa = 0 and care_of_hh_memb = 0 and age_sixty = 0 and sixteen_seventeen = 0 and care_child_six = 0 and employed_thirty = 0 and unemployment = 0 and enrolled_school = 0 and CD_program = 0 and receiving_MFIP = 0 and receiving_DWP_WB = 0 and applied_SSI = 0 THEN wreg_exempt = 0
  DO
	  Dialog abawd_exemptions
		IF ButtonPressed = 0 THEN
		  Dialog cancel_dialog
		  IF ButtonPressed = yes then stopscript
		End If	
   	      IF wreg_exempt = 1 or age_exempt = 1 or cert_preg = 1 or working_20 = 1 or receiving_cash = 1 or dependent_child = 1 or work_exp = 1 or approved_ET = 1 THEN Check1 = 1
		IF ButtonPressed = previous_button then exit do
    DO
		Dialog get_abawd_months
		  IF ButtonPressed = 0 THEN
		    Dialog cancel_dialog
		    IF ButtonPressed = yes then stopscript
		END IF
		IF ButtonPressed = previous_button then exit do
      DO
	  Dialog get_worker_comments
	    IF ButtonPressed = 0 THEN 
	      Dialog cancel_script
	      IF ButtonPressed = yes then stopscript
	    END IF	
	    IF ButtonPressed = previous_button THEN exit do
      LOOP until ButtonPressed = -1
    LOOP until ButtonPressed = -1
  LOOP until ButtonPressed = -1
LOOP until ButtonPressed = -1

PF3

'CASE/NOTING=================================================================================================
'Treat abawd_yn = 1 to be a true/false statement. abawd_yn = 1  is an attempt to reduce key strokes. ========================================================================================
call what_is_abawd_status(abawd_status)

call navigate_to_screen("case", "note")
	PF9
	call write_new_line_in_case_note("***CL has been screened for ABAWD***")
	call write_editbox_in_case_note("", abawd_status, 4)	
		IF worker_comment <> "" THEN call write_editbox_in_case_note("ADDITIONAL NOTES", worker_comment, 4)
	call write_new_line_in_case_note("---")
	call write_new_line_in_case_note(worker_signature)