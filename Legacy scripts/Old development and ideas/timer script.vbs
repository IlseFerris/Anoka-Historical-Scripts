file = "Q:\Adult Share\Script test\testdoc.txt"

Set objExcel = CreateObject("Excel.Application")


time_to_compare = now
current_time_with_variable = DateAdd("n", -1, now) 'The middle section represents the amount of minutes back (or forward for positive numbers) you need to compare
MsgBox current_time_with_variable - time_to_compare
