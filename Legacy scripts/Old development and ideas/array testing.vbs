EMConnect ""

row = 7
var = 0


Dim case_number(12)

Do
EMReadScreen case_number_check, 8, row, 5
case_number(var) = case_number_check
row = row + 1
var = var + 1
Loop until case_number_check = "        "


test = Filter(case_number, "        ", False)
'For each x in test
Msgbox Join(test, ", ")
'next