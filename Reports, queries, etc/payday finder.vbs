EMConnect ""
EMReadScreen PIC_check, 43, 3, 22
EMReadScreen JOBS_check, 17, 2, 33

If PIC_check = "Food Support Prospective Income Calculation" then EMReadScreen wage, 7, 17, 57
If PIC_check = "Food Support Prospective Income Calculation" then EMReadScreen hours_per_pay_date, 6, 16, 51
If PIC_check = "Food Support Prospective Income Calculation" then EMReadScreen pay_frequency_pic_check, 1, 5, 64

If JOBS_check <> "Job Income (JOBS)" and PIC_check <> "Food Support Prospective Income Calculation" then MsgBox "You need to be on JOBS, or on the PIC from JOBS. This script will now stop."
If JOBS_check <> "Job Income (JOBS)" and PIC_check <> "Food Support Prospective Income Calculation" then StopScript



If pay_frequency_pic_check = 1 then pay_frequency = "One Time Per Month"
If pay_frequency_pic_check = 2 then pay_frequency = "Two Times Per Month"
If pay_frequency_pic_check = 3 then pay_frequency = "Every Other Week"
If pay_frequency_pic_check = 4 then pay_frequency = "Every Week"



BeginDialog payday_finder_dialog, 0, 0, 226, 97, "Payday Finder"
  EditBox 60, 10, 50, 15, wage
  DropListBox 65, 30, 100, 15, "One Time Per Month"+chr(9)+"Two Times Per Month"+chr(9)+"Every Other Week"+chr(9)+"Every Week"+chr(9)+"Other", pay_frequency
  DropListBox 115, 55, 85, 15, "Sunday"+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday", day_of_week_paid
  EditBox 105, 75, 75, 15, hours_per_pay_date
  ButtonGroup ButtonPressed
    OkButton 170, 10, 50, 15
    CancelButton 170, 30, 50, 15
  Text 5, 35, 55, 10, "Pay frequency:"
  Text 5, 55, 110, 10, "Day of week paid (if applicable):"
  Text 5, 80, 95, 10, "Average Hours/Pay Date:"
  Text 5, 15, 50, 10, "Gross wage:"
EndDialog

Dialog payday_finder_dialog
If buttonpressed = 0 then stopscript

EMReadScreen PIC_check, 43, 3, 22
If PIC_check = "Food Support Prospective Income Calculation" then EMSendKey "<PF3>"

EMReadScreen edit_mode, 1, 20, 8
If edit_mode <> "A" and edit_mode <> "E" then MsgBox "This is not on edit or add mode. You might be on inquiry mode or display mode. This script will now stop. Try again after moving to edit mode or add mode for JOBS."
If edit_mode <> "A" and edit_mode <> "E" then StopScript

EMReadScreen footer_month, 2, 20, 55
EMReadScreen footer_year, 2, 20, 58

Dim first_payday




If footer_year = 10 then actual_year = 2010
If footer_year = 11 then actual_year = 2011
If footer_year = 12 then actual_year = 2012
If footer_year = 13 then actual_year = 2013
If footer_year = 14 then actual_year = 2014
If footer_year = 15 then actual_year = 2015
If footer_year = 16 then actual_year = 2016
If footer_year = 17 then actual_year = 2017
If footer_year = 18 then actual_year = 2018
If footer_year = 19 then actual_year = 2019
If footer_year = 20 then actual_year = 2020
If footer_year = 21 then actual_year = 2021
If footer_year = 22 then actual_year = 2022
If footer_year = 23 then actual_year = 2023
If footer_year = 24 then actual_year = 2024
If footer_year = 25 then actual_year = 2025
If footer_year = 26 then actual_year = 2026
If footer_year = 27 then actual_year = 2027
If footer_year = 28 then actual_year = 2028
If footer_year = 29 then actual_year = 2029

date_start = footer_month & "/01/" & actual_year


date_start_plus_1 = DateAdd("d", 1, date_start) 
date_start_plus_2 = DateAdd("d", 2, date_start) 
date_start_plus_3 = DateAdd("d", 3, date_start) 
date_start_plus_4 = DateAdd("d", 4, date_start) 
date_start_plus_5 = DateAdd("d", 5, date_start) 
date_start_plus_6 = DateAdd("d", 6, date_start) 

Sub sunday_finder
   If Weekday(date_start, 0) = 1 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 1 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 1 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 1 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 1 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 1 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 1 then first_payday = day(date_start_plus_6)
End Sub

Sub monday_finder
   If Weekday(date_start, 0) = 2 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 2 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 2 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 2 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 2 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 2 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 2 then first_payday = day(date_start_plus_6)
End Sub

Sub tuesday_finder
   If Weekday(date_start, 0) = 3 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 3 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 3 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 3 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 3 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 3 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 3 then first_payday = day(date_start_plus_6)
End Sub

Sub wednesday_finder
   If Weekday(date_start, 0) = 4 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 4 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 4 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 4 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 4 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 4 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 4 then first_payday = day(date_start_plus_6)
End Sub

Sub thursday_finder
   If Weekday(date_start, 0) = 5 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 5 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 5 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 5 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 5 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 5 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 5 then first_payday = day(date_start_plus_6)
End Sub

Sub friday_finder
   If Weekday(date_start, 0) = 6 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 6 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 6 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 6 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 6 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 6 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 6 then first_payday = day(date_start_plus_6)
End Sub

Sub saturday_finder
   If Weekday(date_start, 0) = 7 then first_payday = day(date_start)
   If Weekday(date_start_plus_1, 0) = 7 then first_payday = day(date_start_plus_1)
   If Weekday(date_start_plus_2, 0) = 7 then first_payday = day(date_start_plus_2)
   If Weekday(date_start_plus_3, 0) = 7 then first_payday = day(date_start_plus_3)
   If Weekday(date_start_plus_4, 0) = 7 then first_payday = day(date_start_plus_4)
   If Weekday(date_start_plus_5, 0) = 7 then first_payday = day(date_start_plus_5)
   If Weekday(date_start_plus_6, 0) = 7 then first_payday = day(date_start_plus_6)
End Sub

If day_of_week_paid = "Sunday" then call sunday_finder
If day_of_week_paid = "Monday" then call monday_finder
If day_of_week_paid = "Tuesday" then call tuesday_finder
If day_of_week_paid = "Wednesday" then call wednesday_finder
If day_of_week_paid = "Thursday" then call thursday_finder
If day_of_week_paid = "Friday" then call friday_finder
If day_of_week_paid = "Saturday" then call saturday_finder


If first_payday = 1 then first_payday = "01"
If first_payday = 2 then first_payday = "02"
If first_payday = 3 then first_payday = "03"
If first_payday = 4 then first_payday = "04"
If first_payday = 5 then first_payday = "05"
If first_payday = 6 then first_payday = "06"
If first_payday = 7 then first_payday = "07"
If first_payday = 8 then first_payday = "08"
If first_payday = 9 then first_payday = "09"



'Now it figures out all of the days in that month that are Fridays!

second_payday = first_payday + 7

If second_payday = 7 then second_payday = "07"
If second_payday = 8 then second_payday = "08"
If second_payday = 9 then second_payday = "09"

third_payday = first_payday + 14
fourth_payday = first_payday + 21
fifth_payday = first_payday + 28
If IsDate(footer_month & "/" & fifth_payday & "/" & actual_year) = false then fifth_payday = "false"

'Now it writes the dates and gross income to the screen

'First it clears all of the income from the prospective side.
EMSetCursor 12, 54
EMSendKey "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
EMSetCursor 13, 54
EMSendKey "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
EMSetCursor 14, 54
EMSendKey "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
EMSetCursor 15, 54
EMSendKey "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
EMSetCursor 16, 54
EMSendKey "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
EMSetCursor 18, 72
EMSendKey "<eraseeof>"

Sub first_payday_sub
   EMWriteScreen footer_month, 12, 54
   EMWriteScreen first_payday, 12, 57
   EMWriteScreen footer_year, 12, 60
   EMSetCursor 12, 67
   EMSendKey "<eraseeof>" + wage
End Sub

Sub second_payday_sub
   EMWriteScreen footer_month, 13, 54
   EMWriteScreen second_payday, 13, 57
   EMWriteScreen footer_year, 13, 60
   EMSetCursor 13, 67
   EMSendKey "<eraseeof>" + wage
End Sub

Sub third_payday_sub
   EMWriteScreen footer_month, 14, 54
   EMWriteScreen third_payday, 14, 57
   EMWriteScreen footer_year, 14, 60
   EMSetCursor 14, 67
   EMSendKey "<eraseeof>" + wage
End Sub

Sub fourth_payday_sub
   EMWriteScreen footer_month, 15, 54
   EMWriteScreen fourth_payday, 15, 57
   EMWriteScreen footer_year, 15, 60
   EMSetCursor 15, 67
   EMSendKey "<eraseeof>" + wage
End Sub

Sub fifth_payday_sub
   If fifth_payday <> "false" then EMWriteScreen footer_month, 16, 54
   If fifth_payday <> "false" then EMWriteScreen fifth_payday, 16, 57
   If fifth_payday <> "false" then EMWriteScreen footer_year, 16, 60
   If fifth_payday <> "false" then EMSetCursor 16, 67
   If fifth_payday <> "false" then EMSendKey "<eraseeof>" + wage
End Sub

Sub one_time_per_month_sub
   EMWriteScreen footer_month, 12, 54
   EMWriteScreen "01", 12, 57
   EMWriteScreen footer_year, 12, 60
   EMSetCursor 12, 67
   EMSendKey wage
End Sub

Sub two_times_per_month_sub
   EMWriteScreen footer_month, 12, 54
   EMWriteScreen "01", 12, 57
   EMWriteScreen footer_year, 12, 60
   EMSetCursor 12, 67
   EMSendKey wage
   EMWriteScreen footer_month, 13, 54
   EMWriteScreen "15", 13, 57
   EMWriteScreen footer_year, 13, 60
   EMSetCursor 13, 67
   EMSendKey wage
End Sub

If pay_frequency = "Every Other Week" or pay_frequency = "Every Week" then call first_payday_sub
If pay_frequency = "Every Week" then call second_payday_sub
If pay_frequency = "Every Other Week" or pay_frequency = "Every Week" then call third_payday_sub
If pay_frequency = "Every Week" then call fourth_payday_sub
If pay_frequency = "Every Other Week" or pay_frequency = "Every Week" then call fifth_payday_sub

If pay_frequency = "One Time Per Month" or pay_frequency = "Other" then call one_time_per_month_sub
If pay_frequency = "Two Times Per Month" then call two_times_per_month_sub

'Now it figures the hours worked prospectively, and adds this to the screen.

If pay_frequency = "One Time Per Month" then hours = hours_per_pay_date
If pay_frequency = "Two Times Per Month" then hours = hours_per_pay_date * 2
If pay_frequency = "Every Other Week" then hours = hours_per_pay_date * 2.15
If pay_frequency = "Every Week" then hours = hours_per_pay_date * 4.3

EMWriteScreen fix(hours), 18, 72