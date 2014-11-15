file = "Q:\Adult Share\Script test\testdoc.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
set objTS=objFSO.opentextfile(file, 1)

sText = objTS.ReadAll
objTS.Close

time_to_compare = CDate(sText) 'Replace this with whatever variable you use as your datestamp.
current_time_with_variable = DateAdd("n", -1, now) 'The middle section represents the amount of minutes back (or forward for positive numbers) you need to compare
If TimeValue(current_time_with_variable) <= TimeValue(time_to_compare) then MsgBox "You just did this! Wait a minute then try again!"
If TimeValue(current_time_with_variable) <= TimeValue(time_to_compare) then stopscript
MsgBox "success"

set objTS=objFSO.opentextfile(file, 2)

ObjTS.WriteLine(now)

objTS.Close

