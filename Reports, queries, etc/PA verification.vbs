'SECTION 01

EMConnect ""

  row = 1
  col = 1

EMSearch "Case Nbr: ", row, col

EMReadScreen case_number, 8, row, col + 10
If case_number = "AR" then case_number = ""

BeginDialog benefit_history_dialog, 0, 0, 256, 146, "Benefit History Dialog"
  Text 5, 5, 50, 10, "Case number:"
  EditBox 55, 0, 80, 15, case_number
  GroupBox 5, 20, 180, 50, "Month range requested:"
  Text 15, 35, 100, 10, "First month needed (MM/YY):"
  EditBox 120, 30, 55, 15, first_month
  Text 15, 55, 100, 10, "Last month needed (MM/YY):"
  EditBox 120, 50, 55, 15, last_month
  GroupBox 5, 80, 180, 30, "Programs proof is requested for:"
  CheckBox 15, 95, 25, 10, "GA", GA_check
  CheckBox 50, 95, 30, 10, "MSA", MSA_check
  CheckBox 90, 95, 30, 10, "SNAP", SNAP_check
  Text 5, 120, 70, 10, "Sign your case note:"
  EditBox 80, 115, 115, 15, worker_sig
  CheckBox 5, 135, 220, 10, "Check here to print this on letterhead instead of a SPEC/MEMO.", letterhead_check
  ButtonGroup ButtonPressed
    OkButton 200, 5, 50, 15
    CancelButton 200, 25, 50, 15
EndDialog

Dialog benefit_history_dialog
If ButtonPressed = 0 then stopscript

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



'SECTION 02

EMSendKey "<attn>"
Do
  EMWaitReady 1, 1
  EMReadScreen MAI_check, 3, 1, 33
  If MAI_check = "   " then EMSendKey "<attn>"
Loop until MAI_check = "MAI"

EMReadScreen training_check, 7, 8, 15
EMReadScreen production_check, 7, 6, 15
If training_check = "RUNNING" and production_check = "RUNNING" then MsgBox "You have production and training both running. Close one before proceeding."
If training_check = "RUNNING" and production_check = "RUNNING" then stopscript
If training_check <> "RUNNING" and production_check <> "RUNNING" then MsgBox "You need to run this script on the window that has MAXIS production on it. Please try again."
If training_check <> "RUNNING" and production_check <> "RUNNING" then stopscript
If training_check = "RUNNING" then EMSendKey "3" + "<enter>"
If production_check = "RUNNING" then EMSendKey "1" + "<enter>"
Do
  EMWaitReady 1, 1
  EMReadScreen MAI_check, 3, 1, 33
Loop until MAI_check <> "MAI"

'SECTION 03

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 1

'This Do...loop checks for the password prompt.
Do
  EMReadScreen password_prompt, 38, 2, 23
  IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

'SECTION 04

'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43 
EMWriteScreen "memb", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

EMReadScreen first_name, 12, 6, 63
EMReadScreen middle_initial, 1, 6, 79
EMReadScreen last_name, 25, 6, 30
EMSetCursor 21, 21
EMSendKey "<PF1>"
EMWaitReady 1, 1

EMReadScreen worker, 21, 19, 10
EMReadScreen worker_phone, 12, 19, 45
EMSendKey "<enter>"
EMWaitReady 1, 1

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

EMWriteScreen "addr", 20, 71
EMSendKey "<enter>"
EMWaitReady 1, 1

EMReadScreen first_addr_line, 22, 6, 43
EMReadScreen second_addr_line, 22, 7, 43
EMReadScreen city_line, 15, 8, 43
EMReadScreen state_line, 2, 8, 66
EMReadScreen zip_line, 12, 9, 43

converted_whole_name = (Replace (first_name, "_", "")) & " " & middle_intial & " " & (Replace (last_name, "_", ""))
converted_first_addr_line = (Replace (first_addr_line, "_", ""))
converted_second_addr_line = (Replace (second_addr_line, "_", ""))
converted_city_line = (Replace (city_line, "_", ""))
converted_state_line = (Replace (state_line, "_", ""))
no_underscore_zip_line = (Replace (zip_line, "_", ""))
converted_zip_line = (Replace (no_underscore_zip_line, " ", "-"))

'SECTION 05

'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWriteScreen "mony", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43 
EMWriteScreen "inqx", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

EMWriteScreen first_month(0), 6, 38
EMWriteScreen first_month(1), 6, 41

EMWriteScreen last_month(0), 6, 53
EMWriteScreen last_month(1), 6, 56

If GA_check = 1 then EMWriteScreen "x", 11, 5
If MSA_check = 1 then EMWriteScreen "x", 13, 50
If SNAP_check = 1 then EMWriteScreen "x", 9, 5

EMSendKey "<enter>"
EMWaitReady 1, 1

row = 6 'Setting the variable for the do...loop.
array_size = 0
dim line_array()

Do
  EMReadScreen line_check, 39, row, 7 'set array for +1 variable if the line isn't blank. Keep adding the lines to the array until you reach blank space. Account for multiple pages.
  If line_check <> "                                       " then
    redim preserve line_array(array_size)
    line_array(array_size) = line_check
    array_size = array_size + 1
    row = row + 1
    If row = 18 then
      row = 6
      EMSendKey "<PF8>"
      EMWaitReady 1, 1
    End if
  End if
Loop until line_check = "                                       "

If letterhead_check = 1 then
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
  objselection.TypeParagraph()
  objselection.TypeParagraph()
  objselection.typetext "Case number: " & case_number
  objselection.TypeParagraph()
  objselection.TypeParagraph()
  objselection.typetext "As requested, here is a printout of your benefits from the period of " & first_month(0) & "/" & first_month(1) & " to " & last_month(0) & "/" & last_month(1) & ":" 
  objselection.TypeParagraph()
  objselection.TypeParagraph()
  
  objSelection.Font.Name = "Courier New"
  objSelection.Font.Size = "12"
  objselection.typetext "    DATE           PROGRAM           AMOUNT"
  objselection.TypeParagraph()
  For each x in line_array
    new_array = split(x, " ")
    new_array_ubound = UBound(new_array)
    If new_array(1) = "FS" then new_array(1) = "SNAP"
    If new_array(1) = "MS" then new_array(1) = "MSA "
    If new_array(1) = "GA" then new_array(1) = " GA "
    objselection.typetext "  " & new_array(0) & "          " & new_array(1) & "            $" & new_array(new_array_ubound)
    objselection.TypeParagraph()
  Next
  objselection.TypeParagraph()

  objSelection.Font.Name = "Times New Roman"
  objSelection.Font.Size = "12"
  objselection.typetext "Please let your worker know if you have any other questions. Thank you."
  objselection.TypeParagraph()
  objselection.TypeParagraph()
  objselection.typetext "Worker: " & worker_name
  objselection.TypeParagraph()
  objselection.typetext "Phone: " & worker_phone

End if

If letterhead_check = 0 then

  'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

  EMWriteScreen "spec", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43 
  EMWriteScreen "memo", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  EMSendKey "<PF5>"
  EMWaitReady 1, 1

  EMSendKey "x" & "<enter>"
  EMWaitReady 1, 1

  EMSendKey "As requested, here is a printout of your benefits from the period of " & first_month(0) & "/" & first_month(1) & " to " & last_month(0) & "/" & last_month(1) & ". Please let your worker know if you have any other questions. Thank you." 
  EMSendKey "<newline>"
  EMSendKey "<newline>"
  EMSendKey "Benefits issued as follows: "
  EMSendKey "<newline>"
  EMSendKey "<newline>"  
  EMSendKey "    DATE           PROGRAM           AMOUNT"
  EMSendKey "<newline>"
  For each x in line_array
    new_array = split(x, " ")
    new_array_ubound = UBound(new_array)
    If new_array(1) = "FS" then new_array(1) = "SNAP"
    If new_array(1) = "MS" then new_array(1) = "MSA "
    If new_array(1) = "GA" then new_array(1) = " GA "
    EMSendKey "  " & new_array(0) & "          " & new_array(1) & "            $" & new_array(new_array_ubound)
    EMGetCursor row, col
    If row = 17 then 
      EMSendKey "<PF8>"
      EMWaitReady 1, 1
    End if
    If row <> 17 then EMSendKey "<newline>"
  Next
  EMSendKey "<PF4>"
  EMWaitReady 1, 1
End if

'This Do...loop gets back to SELF
do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

EMWriteScreen "case", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43 
EMWriteScreen "note", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

EMSendKey "<PF9>"
EMWaitReady 1, 1

EMSendKey "**Client requested PA verification**" & "<newline>"
If letterhead_check = 1 then EMSendKey "* Printed on letterhead for client." & "<newline>"
If letterhead_check = 0 then EMSendKey "* Sent using SPEC/MEMO." & "<newline>"
EMSendKey "* Dates requested: " & first_month(0) & "/" & first_month(1) & "-" & last_month(0) & "/" & last_month(1) & "<newline>"
EMSendKey "---" & "<newline>"
EMSendKey worker_sig
