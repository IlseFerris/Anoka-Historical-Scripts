'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = ""
'start_time = timer

''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'-----------------------------------------------------------------------------------------------------------


case_number = 201295

call navigate_to_screen("stat", "wreg")
  EMReadScreen WREG_total_pages, 1, 2, 78
    If WREG_total_pages <> 0 then
	EMSetCursor 13, 57
	EMSendKey "X"
	transmit
	  'Do
	    EMReadScreen jan_11, 7, 19, 'what goes here again?
	    msgbox(jan_11)
end If

'look into using/gutting autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)


	
