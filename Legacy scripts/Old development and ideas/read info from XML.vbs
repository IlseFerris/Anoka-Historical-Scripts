Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objDictionary = CreateObject("Scripting.Dictionary") 
Const ForReading = 1 
Set objFile = objFSO.OpenTextFile ("h:\test.xml", ForReading) 
XML_file = objFile.ReadAll
objFile.Close 
 


search_variable = "SSN"
search_string_start = "<ap:" & search_variable & ">"
search_string_end = "</ap:" & search_variable & ">"

If InStr(XML_file, search_string_start) <> 0 then
  start_point = InStr(XML_file, search_string_start) + len(search_string_start)
  end_point = InStr(XML_file, search_string_end)
  search_result = Mid(XML_file, start_point, end_point - start_point)
  MsgBox "String: " & search_result
Else
  MsgBox "string not found"
End if

EMConnect ""

Do
EMSendKey "<PF3>"
EMWaitReady 1, 1
EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

EMWriteScreen "pers", 16, 43
EMSendKey "<enter>"
EMWaitReady 1, 1

If search_variable = "SSN" then
  SSN = split(search_result, "-")
  EMWriteScreen SSN(0), 14, 36
  EMWriteScreen SSN(1), 14, 40
  EMWriteScreen SSN(2), 14, 43
End if


