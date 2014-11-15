'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - initialize scripts"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DEFINING VARIABLES----------------------------------------------------------------------------------------------------
path_for_excel_file = "Q:\Blue Zone Scripts\Spreadsheets for script use\worker list.xlsx" 'Path for the excel file the worker list is stored at
excel_row = 2 'this is the row the numbers start on the worker list spreadsheet
ZMD_folder_path = "Q:\Blue Zone Scripts\ZMD files\" 'The file path where the various ZMD file templates are kept
default_ZMD_path = "H:\BlueZone\Config\Anoka Desktop.zmd" 'The location where the "default" ZMD is kept. In Anoka County this is called from the ZenWorks app launcher

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Grabbing user ID to determine user of script. Only some users are allowed to use this script.
Set objNet = CreateObject("WScript.NetWork") 
user_ID_for_scripts = ucase(objNet.UserName)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open(path_for_excel_file) 
objExcel.DisplayAlerts = False

'Finding the user in the Excel file
Do
  If ucase(user_ID_for_scripts) = ObjExcel.Cells(excel_row, 3).Value then
    exit do
  Else
    excel_row = excel_row + 1
  End if
Loop until ObjExcel.Cells(excel_row, 3).Value = ""

'Closing script if the user isn't found
If ObjExcel.Cells(excel_row, 3).Value = "" then
  MsgBox "User not found. Contact Veronica Cary for assistance, you may need to be added to the worker list for scripts."
  objExcel.Workbooks.Close
  objExcel.quit
  Wscript.Quit
End if

'Completing the file path for the ZMD file by combining the ZMD type from the spreadsheet with the ZMD_folder_path variable and the file extension
ZMD_path = ZMD_folder_path & ObjExcel.Cells(excel_row, 5).Value & ".zmd"

'Notifying the user that the script will start
worker_notification = MsgBox("This script will shut down BlueZone and install the update.", 1) 
If worker_notification = 2 then Wscript.Quit

'Killing BlueZone processes
strComputer = "."
strProcessToKill = "BZMD.exe" 
Set objWMIService = GetObject("winmgmts:" _
  & "{impersonationLevel=impersonate}!\\" _ 
  & strComputer & "\root\cimv2") 
Set colProcess = objWMIService.ExecQuery _
  ("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")
count = 0
For Each objProcess in colProcess
  objProcess.Terminate()
  count = count + 1
Next 

'Creating our FileSystemObject to copy the ZMD where it needs to go
dim filesys
set filesys=CreateObject("Scripting.FileSystemObject")


If filesys.FileExists(ZMD_path) Then
  filesys.CopyFile ZMD_path, "H:\BlueZone\Config\Anoka Desktop.zmd", True
  MsgBox "Config complete. User scripts now set to " & ObjExcel.Cells(excel_row, 5).Value & ". You can restart BlueZone now."
Else
  script_end_procedure_wsh("Config error. Source file (" & ZMD_path & ") not found. Your user ID is listed as " & user_ID_for_scripts & ". Email Veronica Cary a screenshot of this error message.")
End if


'Closing the notebook
objExcel.Workbooks.Close
objExcel.quit

'Calling a custom function to end the script and write the usage statistics to the log file.
script_end_procedure_wsh("") 


