'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - letterhead generator"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CONNECTS TO BLUEZONE, CHECKS FOR MAXIS
EMConnect ""
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then
  MsgBox "You aren't in MAXIS. The script will now stop."
  StopScript
End if

'SEARCHES FOR A CASE NUMBER.
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
If case_number = "AR" or case_number = "________" then case_number = ""

'DIALOG FOR SCRIPT
BeginDialog letterhead_generator_dialog, 0, 0, 146, 117, "Letterhead Generator Dialog"
  EditBox 70, 5, 60, 15, case_number
  CheckBox 20, 40, 35, 10, "Blank", blank_check
  CheckBox 20, 55, 35, 10, "Client", client_check
  CheckBox 20, 70, 35, 10, "AREP", AREP_check
  CheckBox 20, 85, 35, 10, "SWKR", SWKR_check
  CheckBox 20, 100, 35, 10, "Vendor", vendor_check
  ButtonGroup ButtonPressed
    OkButton 80, 60, 50, 15
    CancelButton 80, 80, 50, 15
  Text 15, 10, 50, 10, "Case number:"
  Text 5, 25, 140, 10, "Who is this going to? Check all that apply:"
EndDialog


'RUNS DIALOG
Dialog letterhead_generator_dialog
If ButtonPressed = 0 then stopscript

'MAKES SURE MAXIS IS NOT PASSWORDED OUT
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then StopScript

'NAVIGATES TO STAT/SUMM FOR ALL CASES, TO PROTECT AGAINST ERROR PRONE
call navigate_to_screen("stat", "summ")

'NAVIGATES TO STAT/MEMB AND STAT/ADDR IF NEEDED, READS CLIENT INFO
If client_check = 1 then
  call navigate_to_screen ("stat", "memb")
  EMReadScreen first_name, 12, 6, 63
  EMReadScreen middle_initial, 1, 6, 79
  EMReadScreen last_name, 25, 6, 30
  call navigate_to_screen ("stat", "addr")
  EMReadScreen first_addr_line, 22, 6, 43
  EMReadScreen second_addr_line, 22, 7, 43
  EMReadScreen city_line, 15, 8, 43
  EMReadScreen state_line, 2, 8, 66
  EMReadScreen zip_line, 12, 9, 43
End if

'NAVIGATES TO STAT/AREP IF NEEDED, AND READS AREP INFO
If AREP_check = 1 then
  call navigate_to_screen ("stat", "arep")
  EMReadScreen arep_line, 37, 4, 32
  EMReadScreen first_arep_addr_line, 22, 5, 32
  EMReadScreen second_arep_addr_line, 22, 6, 32
  EMReadScreen arep_city_line, 15, 7, 32
  EMReadScreen arep_state_line, 2, 7, 55
  EMReadScreen arep_zip_line, 5, 7, 64
End if

'NAVIGATES TO STAT/SWKR IF NEEDED, AND READS SWKR INFO
If SWKR_check = 1 then
  call navigate_to_screen ("stat", "swkr")
  EMReadScreen swkr_line, 35, 6, 32
  EMReadScreen first_swkr_addr_line, 22, 8, 32
  EMReadScreen second_swkr_addr_line, 22, 9, 32
  EMReadScreen swkr_city_line, 15, 10, 32
  EMReadScreen swkr_state_line, 2, 10, 54
  EMReadScreen swkr_zip_line, 5, 10, 63
End if

'NAVIGATES TO STAT/FACI IF NEEDED, AND READS VENDOR INFO, THEN GOES TO MONY/VNDS TO GET INFO
If vendor_check = 1 then
  call navigate_to_screen ("stat", "faci")
  EMReadScreen FACI_total, 1, 2, 78
  If FACI_total <> 0 then
    row = 14
    Do
      EMReadScreen date_in_check, 4, row, 53
      EMReadScreen date_out_check, 4, row, 77
      If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
      If row > 18 then
        EMReadScreen FACI_page, 1, 2, 73
        If FACI_page = FACI_total then 
          FACI_status = "Not in facility"
        Else
          transmit
          row = 14
        End if
      End if
    Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
  End if
  EMReadScreen vendor_number, 8, 5, 43
  If vendor_number = "________" then
    MsgBox "Vendor number not found for this client."
    vendor_check = 0
  Else
    call navigate_to_screen ("mony", "vnds")
    EMWriteScreen vendor_number, 4, 59
    transmit
    EMReadScreen vendor_name_line, 30, 3, 15
    EMReadScreen vendor_c_o_line, 30, 4, 15
    EMReadScreen first_vendor_addr_line, 22, 5, 15
    EMReadScreen second_vendor_addr_line, 22, 6, 15
    EMReadScreen vendor_city_line, 15, 7, 15
    EMReadScreen vendor_state_line, 2, 7, 36
    EMReadScreen vendor_zip_line, 5, 7, 46
  End if
End if

'CONVERTS COLLECTED NAME AND ADDRESS INTO INFO USABLE TO THE SCRIPT
converted_whole_name = (Replace (first_name, "_", "")) & " " & middle_intial & " " & (Replace (last_name, "_", ""))
converted_first_addr_line = (Replace (first_addr_line, "_", ""))
converted_second_addr_line = (Replace (second_addr_line, "_", ""))
converted_city_line = (Replace (city_line, "_", ""))
converted_state_line = (Replace (state_line, "_", ""))
no_underscore_zip_line = (Replace (zip_line, "_", ""))
converted_zip_line = (Replace (no_underscore_zip_line, " ", "-"))

'CONVERTS AREP NAME AND ADDRESS INTO INFO USABLE TO THE SCRIPT
converted_arep_name = (Replace (arep_line, "_", ""))
converted_first_arep_addr_line = (Replace (first_arep_addr_line, "_", ""))
converted_second_arep_addr_line = (Replace (second_arep_addr_line, "_", ""))
converted_arep_city_line = (Replace (arep_city_line, "_", ""))

'CONVERTS SWKR NAME AND ADDRESS INTO INFO USABLE TO THE SCRIPT
converted_swkr_name = (Replace (swkr_line, "_", ""))
converted_first_swkr_addr_line = (Replace (first_swkr_addr_line, "_", ""))
converted_second_swkr_addr_line = (Replace (second_swkr_addr_line, "_", ""))
converted_swkr_city_line = (Replace (swkr_city_line, "_", ""))

'CONVERTS VENDOR NAME AND ADDRESS INTO INFO USABLE TO THE SCRIPT
vendor_name_line = trim(replace(vendor_name_line, "_", ""))
vendor_c_o_line = trim(replace(vendor_c_o_line, "_", ""))
first_vendor_addr_line = trim(replace(first_vendor_addr_line, "_", ""))
second_vendor_addr_line = trim(replace(second_vendor_addr_line, "_", ""))
vendor_city_line = trim(replace(vendor_city_line, "_", ""))
vendor_state_line = trim(replace(vendor_state_line, "_", ""))
vendor_zip_line = trim(replace(vendor_zip_line, "_", ""))

'LOADS LETTERHEAD IF THE CLIENT_CHECK BOX WAS PRESSED
If client_check = 1 then 
  Set objWord = CreateObject("Word.Application")
  objWord.Visible = true
  set objDoc = objWord.Documents.open("L:\Correspondence\Letterhead - Anoka.dotx")
   Set objSelection = objWord.Selection
   objselection.typetext converted_whole_name
   objselection.TypeParagraph()
   objselection.typetext converted_first_addr_line
   objselection.TypeParagraph()
   If converted_second_addr_line <> "" then objselection.typetext converted_second_addr_line
   If converted_second_addr_line <> "" then objselection.TypeParagraph()
   objselection.typetext converted_city_line & ", " & converted_state_line & " " & converted_zip_line
End if

'LOADS LETTERHEAD IF THE AREP_CHECK BOX WAS PRESSED
If AREP_check = 1 then 
  Set objWord = CreateObject("Word.Application")
  objWord.Visible = true
  set objDoc = objWord.Documents.open("L:\Correspondence\Letterhead - Anoka.dotx")
   Set objSelection = objWord.Selection
   objselection.typetext converted_arep_name
   objselection.TypeParagraph()
   objselection.typetext converted_first_arep_addr_line
   objselection.TypeParagraph()
   If converted_second_arep_addr_line <> "" then objselection.typetext converted_second_arep_addr_line
   If converted_second_arep_addr_line <> "" then objselection.TypeParagraph()
   objselection.typetext converted_arep_city_line & ", " & arep_state_line & " " & arep_zip_line
End if

'LOADS LETTERHEAD IF THE SWKR_CHECK BOX WAS PRESSED
If SWKR_check = 1 then 
  Set objWord = CreateObject("Word.Application")
  objWord.Visible = true
  set objDoc = objWord.Documents.open("L:\Correspondence\Letterhead - Anoka.dotx")
  Set objSelection = objWord.Selection
  objselection.typetext converted_swkr_name
  objselection.TypeParagraph()
  objselection.typetext converted_first_swkr_addr_line
  objselection.TypeParagraph()
  If converted_second_swkr_addr_line <> "" then objselection.typetext converted_second_swkr_addr_line
  If converted_second_swkr_addr_line <> "" then objselection.TypeParagraph()
  objselection.typetext converted_swkr_city_line & ", " & swkr_state_line & " " & swkr_zip_line
End if

'LOADS BLANK LETTERHEAD IF THE BLANK_CHECK BOX WAS PRESSED
If BLANK_check = 1 then
  Set objWord = CreateObject("Word.Application")
  objWord.Visible = true
  set objDoc = objWord.Documents.open("L:\Correspondence\Letterhead - Anoka.dotx")
End if

'LOADS LETTERHEAD IF THE VENDOR_CHECK BOX WAS PRESSED
If vendor_check = 1 then
  Set objWord = CreateObject("Word.Application")
  objWord.Visible = true
  set objDoc = objWord.Documents.open("L:\Correspondence\Letterhead - Anoka.dotx")
  Set objSelection = objWord.Selection
  objselection.typetext vendor_name_line
  objselection.TypeParagraph()
  If vendor_c_o_line <> "" then
    objselection.typetext vendor_c_o_line
    objselection.TypeParagraph()
  End if
  objselection.typetext first_vendor_addr_line
  objselection.TypeParagraph()
  If second_vendor_addr_line <> "" then
    objselection.typetext second_vendor_addr_line
    objselection.TypeParagraph()
  End if
  objselection.typetext vendor_city_line & ", " & vendor_state_line & " " & vendor_zip_line
End if

script_end_procedure("")
