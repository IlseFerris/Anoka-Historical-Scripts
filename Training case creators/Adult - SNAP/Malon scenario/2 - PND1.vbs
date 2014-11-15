'VARIABLES TO DECLARE

just_memb_01 = True
cash_app = "N"
HC_app = "N"
FS_app = "Y"
emer_app = "N"
paperless_indicator = "N"
case_number_stop = "" 'To use if you have multiple pnd1 cases, and you need the script to stop moving them at a certain point.


EMConnect ""

Do

EMReadScreen PND1_check, 4, 2, 50
If PND1_check <> "PND1" then
  MsgBox "Not on PND1"
  StopScript
End if

EMReadScreen case_number, 8, 7, 3
case_number = trim(case_number)
If case_number = "" or case_number = case_number_stop then stopscript

EMWriteScreen "stat", 20, 13
EMWriteScreen "________", 20, 33
EMWriteScreen case_number, 20, 33
EMSendKey "<enter>"
EMWaitReady 0, 0
EMSendKey "<enter>"
EMWaitReady 0, 0
EMWriteScreen "N", 6, 64
EMWriteScreen "N", 6, 73

MAXIS_row = 6

Do
  EMWriteScreen cash_app, MAXIS_row, 28
  EMWriteScreen HC_app, MAXIS_row, 37
  EMWriteScreen FS_app, MAXIS_row, 46
  EMWriteScreen emer_app, MAXIS_row, 55
  MAXIS_row = MAXIS_row + 1
  EMReadScreen member_row_check, 2, MAXIS_row, 3
  If just_memb_01 = True and MAXIS_row = 7 then   'Separates the original data to simplify the do...loop. It restores after the loop. This ensures that PROG gets coded correctly.
    actual_cash_app = cash_app
    actual_HC_app = HC_app
    actual_FS_app = FS_app 
    actual_emer_app = emer_app 
    cash_app = "N"
    HC_app = "N"
    FS_app = "N"
    emer_app = "N"
  End if
Loop until member_row_check = "  "

If just_memb_01 = True then   'Restoring original values
  cash_app = actual_cash_app
  HC_app = actual_HC_app
  FS_app = actual_FS_app 
  emer_app = actual_emer_app 
End if

EMSendKey "<enter>"
EMWaitReady 0, 0

If cash_app = "Y" then
  EMReadScreen appl_month, 2, 6, 33
  EMReadScreen appl_day, 2, 6, 36
  EMReadScreen appl_year, 2, 6, 39
  EMWriteScreen appl_month, 6, 55
  EMWriteScreen appl_day, 6, 58
  EMWriteScreen appl_year, 6, 61
End if

If emer_app = "Y" then
  EMReadScreen appl_month, 2, 8, 33
  EMReadScreen appl_day, 2, 8, 36
  EMReadScreen appl_year, 2, 8, 39
  EMWriteScreen appl_month, 8, 55
  EMWriteScreen appl_day, 8, 58
  EMWriteScreen appl_year, 8, 61
  EMWriteScreen "EG", 8, 67
End if

If FS_app = "Y" then
  EMReadScreen appl_month, 2, 10, 33
  EMReadScreen appl_day, 2, 10, 36
  EMReadScreen appl_year, 2, 10, 39
  EMWriteScreen appl_month, 10, 55
  EMWriteScreen appl_day, 10, 58
  EMWriteScreen appl_year, 10, 61
End if

If HC_app = "Y" then
  EMReadScreen appl_month, 2, 12, 33
  EMReadScreen appl_day, 2, 12, 36
  EMReadScreen appl_year, 2, 12, 39
End if

EMWriteScreen "N", 18, 67

EMSendKey "<enter>"
EMWaitReady 0, 0

If HC_app = "Y" then 'HC cases jump to STAT/HCRE
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End if

application_date = cdate(appl_month & "/" & appl_day & "/" & appl_year)
six_month_recert_date = dateadd("m", 6, application_date)
six_month_month = datepart("m", six_month_recert_date)
If len(six_month_month) = 1 then six_month_month = "0" & six_month_month 
six_month_year = datepart("yyyy", six_month_recert_date) - 2000
one_year_recert_date = dateadd("m", 12, application_date)
one_year_month = datepart("m", one_year_recert_date)
If len(one_year_month) = 1 then one_year_month = "0" & one_year_month 
one_year_year = datepart("yyyy", one_year_recert_date) - 2000

If cash_app = "Y" then
  EMWriteScreen one_year_month, 9, 37
  EMWriteScreen one_year_year, 9, 43
End if
  
If FS_app = "Y" then
  EMWriteScreen "N", 15, 75
  EMWriteScreen "x", 5, 58
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMWriteScreen six_month_month, 9, 26
  EMWriteScreen six_month_year, 9, 32
  EMWriteScreen one_year_month, 9, 64
  EMWriteScreen one_year_year, 9, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End if

If HC_app = "Y" then
  EMWriteScreen "x", 5, 71
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMWriteScreen six_month_month, 8, 71
  EMWriteScreen six_month_year, 8, 77
  EMWriteScreen one_year_month, 9, 27
  EMWriteScreen one_year_year, 9, 33
  EMWriteScreen paperless_indicator, 9, 71
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End if

EMSendKey "<enter>"
EMWaitReady 0, 0
EMSendKey "<enter>"
EMWaitReady 0, 0
  
EMWriteScreen "rept", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen "pnd1", 21, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

Loop until case_number = ""