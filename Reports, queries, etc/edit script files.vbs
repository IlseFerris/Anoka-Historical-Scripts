'Opening spreadsheet
Set ObjExcel = CreateObject("Excel.application")

ObjExcel.Application.Workbooks.Open "q:\Blue Zone Scripts\batch script editing.xlsx"
ObjExcel.Application.Visible = True

excel_row = 1


Do
  file_path = ObjExcel.Cells(excel_row, 1).Value
  name_of_script = ObjExcel.Cells(excel_row, 2).Value
  skip_status = ObjExcel.Cells(excel_row, 3).Value
  mandatory_status = ObjExcel.Cells(excel_row, 4).Value

  If name_of_script = "" then exit do

  If mandatory_status <> "mandatory" and skip_status <> "skip" then

    'Opening script and reading contents into sText variable, then closing
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    set objTS=objFSO.opentextfile(file_path, 1)
    sText = objTS.ReadAll
    objTS.Close

    'Creating one variable with the new text to add, from a text file
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    set objTS=objFSO.opentextfile("Q:\Blue Zone Scripts\text to add to scripts.txt", 1)
    newText = objTS.ReadAll
    objTS.Close

    'Opening up the text file again, writing the worker data into the file, then closing
    set objTS=objFSO.opentextfile(file_path, 2)
    ObjTS.WriteLine(replace(newText, "##newname##", name_of_script) & vbCrLf & sText & vbCrLf & "script_end_procedure("& chr(34) & chr(34) & ")")
    objTS.Close

  End if
  excel_row = excel_row + 1
Loop until name_of_script = ""


MsgBox "Success!"