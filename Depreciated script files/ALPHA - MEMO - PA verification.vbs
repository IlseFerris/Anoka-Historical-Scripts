'Removed 01/15/2014 as low priority. May revisit in the future.

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog benefit_history_dialog, 0, 0, 171, 165, "Benefit History Dialog"
  EditBox 70, 0, 80, 15, case_number
  EditBox 120, 30, 35, 15, first_month
  EditBox 120, 50, 35, 15, last_month
  CheckBox 15, 95, 25, 10, "GA", GA_check
  CheckBox 50, 95, 30, 10, "MSA", MSA_check
  CheckBox 90, 95, 30, 10, "SNAP", SNAP_check
  CheckBox 130, 95, 30, 10, "MFIP", MFIP_check
  CheckBox 30, 110, 30, 10, "DWP", DWP_check
  CheckBox 70, 110, 30, 10, "GRH", GRH_check
  CheckBox 105, 110, 30, 10, "EGA", EGA_check
  CheckBox 40, 125, 90, 10, "Emergency Assistance", emergency_assistance_check
  ButtonGroup ButtonPressed
    OkButton 30, 145, 50, 15
    CancelButton 90, 145, 50, 15
  Text 20, 5, 50, 10, "Case number:"
  GroupBox 5, 20, 160, 50, "Month range requested:"
  Text 15, 35, 100, 10, "First month needed (MM/YY):"
  Text 15, 55, 100, 10, "Last month needed (MM/YY):"
  GroupBox 5, 80, 160, 60, "Programs proof is requested for:"
EndDialog

'PRELOADING COMMON VARIABLES----------------------------------------------------------------------------------------------------
first_month = datepart("m", (dateadd("m", -2, (date)))) & "/" & (datepart("yyyy", (dateadd("m", -2, (date)))) - 2000)
if len(first_month) = 4 then first_month = "0" & first_month
last_month = datepart("m", date) & "/" & (datepart("yyyy", date) - 2000)
if len(last_month) = 4 then last_month = "0" & last_month

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""


'Finding the case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(replace(case_number, "_", ""))
If isnumeric(case_number) = False then case_number = ""


'Show the dialog
Dialog benefit_history_dialog
If ButtonPressed = 0 then stopscript

'If worker put dashes or periods in for the date field, it'll switch them to slashes. After that it converts the dates to simple arrays for entry into the INQX screen
first_month = replace(first_month, "-", "/")
first_month = replace(first_month, ".", "/")
first_month = split(first_month, "/")
last_month = replace(last_month, "-", "/")
last_month = replace(last_month, ".", "/")
last_month = split(last_month, "/")
If len(first_month(0)) = 1 then first_month(0) = "0" & first_month(0)
If len(first_month(1)) > 2 then first_month(1) = right(first_month(1), 2)
If len(last_month(0)) = 1 then last_month(0) = "0" & last_month(0)
If len(last_month(1)) > 2 then last_month(1) = right(last_month(1), 2)


'It sends an enter to force the screen to refresh, and checks to see if we're in MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS not found. Are you locked out of this case? Navigate to MAXIS and try again.")

'Navigating to STAT/MEMB and grabbing the client/worker information
call navigate_to_screen("stat", "memb")
EMReadScreen first_name, 12, 6, 63
EMReadScreen middle_initial, 1, 6, 79
EMReadScreen last_name, 25, 6, 30
EMSetCursor 21, 21
PF1
EMReadScreen worker, 21, 19, 10
EMReadScreen worker_phone, 12, 19, 45
transmit

'Removing underscores and concatenates first/middle/last names
converted_whole_name = (Replace (first_name, "_", "")) & " " & middle_intial & " " & (Replace (last_name, "_", ""))

'Converting worker information into something Anoka likes (just first name and last initial)
worker = replace(worker, ".", "")
worker = split(worker)
first_name_worker = worker(0)
y = 0
For each x in worker
  x = trim(x)
  If len(x) <= 1 then x = ""
  If x = " " then x = ""
  If len(x) > 1 then x = left(x, 1) 
  worker(y) = x
  y = y + 1
Next
worker(0) = first_name_worker
worker_name = join(worker)
worker_name = trim(worker_name)
worker_name = replace(worker_name, "  ", " ")

'Going to STAT/ADDR to grab the client's address
call navigate_to_screen("stat", "addr")
EMReadScreen first_addr_line, 22, 6, 43
EMReadScreen second_addr_line, 22, 7, 43
EMReadScreen city_line, 15, 8, 43
EMReadScreen state_line, 2, 8, 66
EMReadScreen zip_line, 12, 9, 43

'Removing underscores from address info
converted_first_addr_line = (Replace (first_addr_line, "_", ""))
converted_second_addr_line = (Replace (second_addr_line, "_", ""))
converted_city_line = (Replace (city_line, "_", ""))
converted_state_line = (Replace (state_line, "_", ""))
no_underscore_zip_line = (Replace (zip_line, "_", ""))
converted_zip_line = (Replace (no_underscore_zip_line, " ", "-"))

'Navigates to INQX and enters first/last month
call navigate_to_screen("mony", "inqx")
EMWriteScreen first_month(0), 6, 38
EMWriteScreen first_month(1), 6, 41
EMWriteScreen last_month(0), 6, 53
EMWriteScreen last_month(1), 6, 56

'The following puts "x"s on the boxes when the corresponding check was indicated in the dialog
If GA_check = 1 then EMWriteScreen "x", 11, 5
If MSA_check = 1 then EMWriteScreen "x", 13, 50
If SNAP_check = 1 then EMWriteScreen "x", 9, 5
If EGA_check = 1 then EMWriteScreen "x", 11, 50
If MFIP_check = 1 then EMWriteScreen "x", 10, 5
If DWP_check = 1 then EMWriteScreen "x", 17, 50
If GRH_check = 1 then EMWriteScreen "x", 16, 50
If emergency_assistance_check = 1 then EMWriteScreen "x", 9, 50


'Navigates to next screen
transmit

'Declaring the array we're about to use in the do...loop
dim line_array()

'Setting the variables for the do...loop.
row = 6 
array_size = 0

'This do...loop checks the screen to see if the current row is not blank. If it isn't, it sets the array to be the size of the array_size variable 
'(remember arrays start on 0, so counterintuitive but accurate to start the variable on 0. After that it adds the current line to the array, increases 
'the variables for both array_size and row (the MAXIS row in this case), and then PF8s if we've reached the last possible row (in our case, 18). This 
'repeats until there's no more information to be found in INQD.
Do
  EMReadScreen current_line, 39, row, 7 
  If current_line <> "                                       " then
    redim preserve line_array(array_size)
    line_array(array_size) = current_line
    array_size = array_size + 1
    row = row + 1
    If row = 18 then
      row = 6
      PF8
      EMReadScreen last_page_check, 21, 24, 2 'In case there are exactly 12 records, then this will read if we're on the last page and exit
      If last_page_check = "THIS IS THE LAST PAGE" then exit do
    End if
  End if
Loop until current_line = "                                       "

'Opens up the Anoka Letterhead in Word.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.open("L:\Correspondence\Letterhead - Anoka.dotx")
Set objSelection = objWord.Selection

'Types out the demographic information we collected above (name/address), as well as a blip about the benefits requested.
objselection.typetext converted_whole_name
objselection.TypeParagraph()
objselection.typetext converted_first_addr_line
objselection.TypeParagraph()
If converted_second_addr_line <> "" then objselection.typetext converted_second_addr_line
If converted_second_addr_line <> "" then objselection.TypeParagraph()
objselection.typetext converted_city_line & ", " & converted_state_line & " " & converted_zip_line
objselection.TypeParagraph()
objselection.TypeParagraph()
objselection.typetext "Case number: " & case_number
objselection.TypeParagraph()
objselection.TypeParagraph()
objselection.typetext "As requested, here is a printout of your benefits from the period of " & first_month(0) & "/" & first_month(1) & " to " & last_month(0) & "/" & last_month(1) & ":" 
objselection.TypeParagraph()
objselection.TypeParagraph()

'Switches fonts to Courier New, and types out the benefit information by splitting the information in the array.
objSelection.Font.Name = "Courier New"
objSelection.Font.Size = "12"
objselection.typetext "    DATE           PROGRAM           AMOUNT"
objselection.TypeParagraph()
For each x in line_array
  new_array = split(x, " ")
  new_array_ubound = UBound(new_array)
  If new_array(1) = "FS" then new_array(1) = "SNAP     "
  If new_array(1) = "MS" then new_array(1) = "MSA      "
  If new_array(1) = "GA" then new_array(1) = " GA      "
  If new_array(1) = "EG" then new_array(1) = "EGA      "
  If new_array(1) = "GR" then new_array(1) = "GRH      "
  If new_array(1) = "MF-FS" then new_array(1) = "MFIP-food"
  If new_array(1) = "MF-MF" then new_array(1) = "MFIP-cash"
  If new_array(1) = "DW" then new_array(1) = "DWP      "
  If new_array(1) = "EA" then new_array(1) = "EA       "
  objselection.typetext "  " & new_array(0) & "          " & new_array(1) & "       $" & new_array(new_array_ubound)
  objselection.TypeParagraph()
Next
objselection.TypeParagraph()

'Switches font back to Times New Roman, and adds worker information to the document
objSelection.Font.Name = "Times New Roman"
objSelection.Font.Size = "12"
objselection.typetext "Please let your worker know if you have any other questions. Thank you."
objselection.TypeParagraph()
objselection.TypeParagraph()
objselection.typetext "Worker: " & worker_name
objselection.TypeParagraph()
objselection.typetext "Phone: " & worker_phone