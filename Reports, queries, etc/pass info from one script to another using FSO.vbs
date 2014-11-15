name_of_script = "test"
Set fs = CreateObject("Scripting.FileSystemObject")
Set ts = fs.OpenTextFile("q:\Blue Zone Scripts\stats file.vbs")
script_to_run = ts.ReadAll
ts.Close
Execute script_to_run