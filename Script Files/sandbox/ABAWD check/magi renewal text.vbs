
'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Functions===========================================================
'Performs a MAXIS check-----------------------------------------------


BeginDialog magi_dlg, 0, 0, 196, 100, "Dialog"
  EditBox 60, 15, 125, 15, case_number
  EditBox 60, 35, 25, 15, Approval_month
  EditBox 150, 35, 30, 15, approval_year
  EditBox 85, 55, 95, 15, worker_signature
  Text 0, 15, 55, 15, "Case Number: "
  Text 0, 35, 55, 10, "Approval Month"
  Text 85, 35, 65, 15, "Approval Year:"
  Text 0, 55, 85, 15, "Case worker signature: "
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 120, 75, 50, 15
EndDialog


EMConnect ""

transmit

maxis_check_function

row = 1
col = 1
EMSearch "Case Nbr:", row, col
If row <> 0 then 
  EMReadScreen case_number, 8, row, col + 10
  case_number = replace(case_number, "_", "")
  case_number = trim(case_number)
End if

If isnumeric(case_number) = False then case_number = ""

back_to_SELF

Do
	dialog magi_dlg
	If buttonpressed = 0 then stopscript
	If worker_signature = "" then msgbox("Please sign your name")
Loop until worker_signature <> ""

call navigate_to_screen("stat", "memb")
call HH_member_custom_dialog(HH_member_array)
back_to_SELF

call navigate_to_screen("spec", "wcom")
EMWriteScreen "Y", 3, 74
transmit


FOR each HH_member in HH_member_array
	DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
		EMReadScreen more_pages, 8, 18, 72
		IF more_pages = "MORE:  -" THEN PF7
	LOOP until more_pages <> "MORE:  -"

	read_row = 7
	DO
		EMReadscreen reference_number, 2, read_row, 62 
		EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
		If waiting_check = "Waiting" and reference_number = HH_member THEN 'checking program type and if it's been printed, needs more fool proofing
			EMSetcursor read_row, 13
			EMSendKey "x"
			Transmit
			pf9
		      EMSetCursor 03, 15
      		EMWriteScreen "You will remain eligible for Medical Assistance because of", 3, 15
	      	EMWriteScreen "new rules and guidelines. (Authority: 42 C.F.R. 435.603(a)", 4, 15
	      	EMWriteScreen "(3); Section 1902(e)(14)(A)", 5, 15
		      PF4
			PF3
			exit do
		ELSE
			read_row = read_row + 1
		END IF
		IF read_row = 18 THEN
			PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18??
			read_row = 7
		End if
	LOOP until reference_number = "  "
NEXT

'NEXT STEP: Why is the script not navigating to page 2 in wcom?
'What to do when there are 14 clients or more on a case
'Why no member 20 SPEC/MEMO in Training Case 203942?

back_to_self
call navigate_to_screen("Case", "Note")

pf9

call write_new_line_in_case_note("***Magi renewal***")
FOR EACH HH_member IN HH_member_array
  magi_case_note_line_one = "* Member " & HH_member & " remains eligible for Medical Assistance for an additional year"
  magi_case_note_line_two = "  because of new rules and guidelines."
  call write_new_line_in_case_note(magi_case_note_line_one)
  call write_new_line_in_case_note(magi_case_note_line_two)
NEXT
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)
