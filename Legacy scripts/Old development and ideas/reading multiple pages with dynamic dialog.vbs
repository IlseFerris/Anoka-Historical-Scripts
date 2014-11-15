'FUNCTIONS AND SUBS

function PF11
  EMSendKey "<PF11>"
  EMWaitReady 1, 1
end function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 1, 1
end function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
end function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 1, 1
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 1, 1
end function

function back_to_self
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

function navigate_to_screen(x, y)
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" then
  row = 1
  col = 1
  EMSearch "Function: ", row, col
  If row <> 0 then 
    EMReadScreen MAXIS_function, 4, row, col + 10
    row = 1
    col = 1
    EMSearch "Case Nbr: ", row, col
    EMReadScreen current_case_number, 8, row, col + 10
    current_case_number = replace(current_case_number, "_", "")
    current_case_number = trim(current_case_number)
  End if
  If current_case_number = case_number and MAXIS_function = ucase(x) then
    row = 1
    col = 1
    EMSearch "Command: ", row, col
    EMWriteScreen y, row, col + 9
    EMSendKey "<enter>"
    EMWaitReady 1, 1
  Else
    Do
      EMSendKey "<PF3>"
      EMWaitReady 1, 1
      EMReadScreen SELF_check, 4, 2, 50
    Loop until SELF_check = "SELF"
    EMWriteScreen x, 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen case_number, 18, 43
    EMWriteScreen y, 21, 70
    EMSendKey "<enter>"
    EMWaitReady 1, 1
    EMReadScreen abended_check, 7, 9, 27
    If abended_check = "abended" then
      EMSendKey "<enter>"
      EMWaitReady 1, 1
    End if
  End if
  End if
End function

Function write_editbox_in_case_note(x, y) 'x is the header, y is the variable for the edit box which will be put in the case note.
  z = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in z 'z represents the variable
    EMGetCursor row, col 
    If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 1, 1
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 17 and col + (len(x)) >= 80) then EMSendKey "<newline>" & "     "
    If (row = 4 and col = 3) then EMSendKey "     "
    EMSendKey x & " "
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 1, 1
  End if
End function

Function write_new_line_in_case_note(x)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 1, 1
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 1, 1
  End if
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

EMConnect ""

row = 1
col = 1
EMSearch "***", row, col
If row = 0 then
  MsgBox "Case note not found."
  stopscript
End if

EMWriteScreen "x", row, 3
transmit

For i = 0 to 3
row = 1
col = 1
  EMSearch "* Verifs needed: ", row, col
  If row = 0 then PF8
  If row <> 0 then exit for
Next
If row = 0 then
  MsgBox "Verifs needed not found."
  stopscript
End if

EMReadScreen verifs_needed, 60, row, 20
If row = 17 then 
  PF8
  row = 4 
Else
  row = row + 1
End if
Do
  EMReadScreen verifs_needed_next_line, 80, row, 3
  If left(verifs_needed_next_line, 1) <> "-" and left(verifs_needed_next_line, 1) <> "*" then verifs_needed = verifs_needed & " " & verifs_needed_next_line
  If row = 17 then 
    PF8
    row = 4 
  Else
    row = row + 1
  End if
Loop until left(verifs_needed_next_line, 1) = "-" or left(verifs_needed_next_line, 1) = "*" or verifs_needed_next_line = "                                                                                "
verifs_needed_array = split(verifs_needed)
verifs_needed = ""
For each x in verifs_needed_array
  x = replace(x, " ", "")
  If x <> "" then
    If verifs_needed = "" then 
      verifs_needed = x
    Else
      verifs_needed = verifs_needed & " " & x
    End if
  End if
Next
verifs_needed_array = split(verifs_needed, ",")


z = UBound(verifs_needed_array) * 25
if z < 50 then z = 40
BeginDialog Dialog1, 0, 0, 200, z, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  y = 5
  For each x in verifs_needed_array
    CheckBox 10, y, 60, 10, x, x
    y = y + 15
  Next
EndDialog

Dialog