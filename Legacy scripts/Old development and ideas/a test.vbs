date_01 = "06/01/12"
date_02 = "06/01/13"

working_date = date_01
x = DateDiff("m", date_01, date_02)
dim date_total()
redim date_total(x)
For i = 0 to x
  date_total(i) = working_date
  working_date = DateAdd("m", 1, working_date)
Next

