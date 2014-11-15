'Removed 05/08/2014, as the process has been simplified (now uses Access table instead of CSV) and added to the FUNCTIONS FILE.

'Getting user name
Set objNet = CreateObject("WScript.NetWork") 
user_ID = objNet.UserName

'Variable for the stats file, it's a CSV
file = "Q:\Blue Zone Scripts\Statistics\usage statistics " & replace(date, "/", ".") & ".csv"

'If the log file doesn't exist, the script has to make one
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(file) = False Then
  Set create_new_FSO = CreateObject("Scripting.FileSystemObject")
  Set create_new = create_new_FSO.CreateTextFile(file, False)
  create_new.WriteLine(", ")
  create_new.close
End If

'Opening text file and reading contents into sText variable, then closing
Set objFSO = CreateObject("Scripting.FileSystemObject")
set objTS=objFSO.opentextfile(file, 1)
sText = objTS.ReadAll
objTS.Close

'creating one variable with the worker data
worker_data = (user_ID & "," & date & "," & time & "," & name_of_script & "," & script_run_time & "," & closing_message)

'Opening up the text file again, writing the worker data into the file, then closing
set objTS=objFSO.opentextfile(file, 2)
ObjTS.WriteLine(worker_data & vbCrLf & replace(sText, ",,,,", ""))
objTS.Close