'-----------------------------------------------------------------------------------------------------------------------------------------------------
'        Script name:    Privileged case finder
'        Description:    Finds privileged cases
'       Target users:    Ronny Cary
'           Division:    Adult and Family
'          Author(s):    Ronny Cary
'      Working state:    purpetual development
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'     Script content:    01. 
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'       Known issues:    This script is just for the script developer to use, as it gathers large quantities of data and will lock up a user's computer.
'   Test breakpoints:    None 
'              Notes:    None
'-----------------------------------------------------------------------------------------------------------------------------------------------------

'SECTION 01

family_1_array = ("x102B36|x102200|x102624|x102b48|x102294|x102A71|x1024F9|x102B55|x102RLM|x102925|x102B51|x102616|x102A75|x1024BV") 
family_2_array = ("x102231|x102886|x102674|x102955|x102RLH|x102b35|x102b52") 
family_3_array = ("x102A44|x102223|x102213|x102B50|x102618|x102234|x102B97|x102218|x102902|x102A18|x102733|x102225") 
family_4_array = ("x102978|x102222|x102895|x1024ES|x102A77|x102B09|x102224|x102869|x102C04|x1024MS|x102TLP") 
PSU_array = ("x102601|x102643|x102B20|x102797|x102928|x1024BM|x102C05|x102630|x102872|x102518|x1024SL|x102282|x102C08|x102105|x102106|x102C06|x102C07|x1024BL|x1024SZ|x102BED|x1024SW|x102247|x102GMZ|x102104|x1024SX|x102233|x10230V|x102TRP") 
adult_1_array = ("x102C02|x102B83|x102395|x102631|x1024MG|x1024RS|x1024SS|x102752|x102293|x102756|x102692") 
adult_2_array = ("x102989|x102B98|x102B93|x102SEC|x102SAC|x102932|x102598|x102268|x102880|x102757") 
adult_3_array = ("x1024DK|x1024AS|x102987|x102628|x102B64|x102524|x102769|x102894|x102750|x102742") 

BeginDialog Dialog1, 0, 0, 191, 67, "Dialog"
  Text 15, 15, 55, 10, "Select the unit:"
  DropListBox 80, 15, 95, 10, "Family 1"+chr(9)+"Family 2"+chr(9)+"Family 3"+chr(9)+"Family 4"+chr(9)+"PSU"+chr(9)+"Adult 1"+chr(9)+"Adult 2"+chr(9)+"Adult 3"+chr(9)+"All Family (Ronny only)"+chr(9)+"All Adult (Ronny only)", unit
  ButtonGroup ButtonPressed
    OkButton 40, 45, 50, 15
    CancelButton 100, 45, 50, 15
EndDialog

Dialog Dialog1
If buttonpressed = 0 then stopscript

If unit = "Family 1" then worker_number_array = family_1_array
If unit = "Family 2" then worker_number_array = family_2_array
If unit = "Family 3" then worker_number_array = family_3_array
If unit = "Family 4" then worker_number_array = family_4_array
If unit = "PSU" then worker_number_array = PSU_array
If unit = "Adult 1" then worker_number_array = adult_1_array
If unit = "Adult 2" then worker_number_array = adult_2_array
If unit = "Adult 3" then worker_number_array = adult_3_array
If unit = "All Family (Ronny only)" then worker_number_array = (family_1_array & "|" & family_2_array & "|" & family_3_array & "|" & family_4_array & "|" & PSU_array)
If unit = "All Adult (Ronny only)" then worker_number_array = (adult_1_array & "|" & adult_2_array & "|" & adult_3_array)
'SECTION 02

worker_number_array = split(worker_number_array, "|")

EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 1, 1
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript

row = 1
col = 1
EMSearch "MAXIS", row, col
If row <> 1 then
  MsgBox "You need to run this script in the window that has MAXIS on it. Please try again."
  StopScript
End if

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

ObjExcel.Cells(1, 1).Value = "MAXIS number"
ObjExcel.Cells(1, 2).Value = "Name"
ObjExcel.Cells(1, 3).Value = "x102number"

excel_row = 2 'This sets the variable for the following.

For each worker_number in worker_number_array

  'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"
  
  EMWriteScreen "rept", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen "actv", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen worker_number_check, 7, 21, 13
  If worker_number_check <> worker_number then
    EMWriteScreen worker_number, 21, 13
    EMSendKey "<enter>"
    EMWaitReady 1, 1
  End if
  EMReadScreen worker_has_cases_to_close_check, 16, 7, 21
  If worker_has_cases_to_close_check = "                " then
    MsgBox "This worker does not appear to have any cases."
  End if
  
  'SECTION 03
  
  MAXIS_row = 7 'This sets the variable for the following do...loop.
  Do
    EMReadScreen last_page_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
    Do
      EMReadScreen case_number, 8, MAXIS_row, 12
      EMReadScreen client_name, 21, MAXIS_row, 21
      case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each. This is for the Excel spreadsheet, so that we aren't entering blank spaces.
      client_name = Trim(client_name)
      If case_number <> "" then 
        ObjExcel.Cells(excel_row, 1).Value = case_number   'Then it writes each into the Excel spreadsheet to be used later.
        ObjExcel.Cells(excel_row, 2).Value = client_name
        ObjExcel.Cells(excel_row, 3).Value = worker_number
        excel_row = excel_row + 1
      End if
      MAXIS_row = MAXIS_row + 1
    Loop until MAXIS_row = 19
    MAXIS_row = 7 'Setting the variable for when the do...loop restarts
    EMSendKey "<PF8>"
    EMWaitReady 1, 1
  Loop until last_page_check = "THIS IS THE LAST PAGE"

Next

MsgBox "Complete! Email the spreadsheet to Ronny. Thank you!"
stopscript





