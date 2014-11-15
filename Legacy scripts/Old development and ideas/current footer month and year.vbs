'The following returns the CURRENT footer month and year.

footer_month = datepart("m", now)
If footer_month = "1" then footer_month = "01"
If footer_month = "2" then footer_month = "02"
If footer_month = "3" then footer_month = "03"
If footer_month = "4" then footer_month = "04"
If footer_month = "5" then footer_month = "05"
If footer_month = "6" then footer_month = "06"
If footer_month = "7" then footer_month = "07"
If footer_month = "8" then footer_month = "08"
If footer_month = "9" then footer_month = "09"
footer_year = Right(year(now), 2)
