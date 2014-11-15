function PF8
  EMSendKey "<PF8>"
  EMWaitReady 1, 1
end function





statement = "Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department Anoka County Human Services Department "
statement = statement & statement & statement & statement
statement = split(statement, " ")

EMConnect ""

Function write_in_case_note(x, y)
  EMSendKey "* " & x & ": "
  For each x in y 'y represents the variable
    EMGetCursor row, col 
    If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then PF8
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 17 and col + (len(x)) >= 80) then EMSendKey "<newline>" & "     "
    If (row = 4 and col = 3) then EMSendKey "     "
    EMSendKey x & " "
  Next
  EMSendKey "<newline>"
End function

Call write_in_case_note("Statement", statement)