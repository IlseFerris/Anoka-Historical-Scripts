'
' 11:39:09  01/09/2014
' BlueZone v5.1C2 Recorded Visual Basic Script File
'

date_for_PW = datepart("m", date)
If len(date_for_PW) = 1 then date_for_PW = "0" & date_for_PW

EMConnect ""

EMSendKey "cicsdt4"
EMSendKey "<Enter>"
EMWaitReady 0, 0

EMSendKey "pwcst05"
EMSendKey "<NewLine>"
EMSendKey "Train#" & date_for_PW
EMSendKey "<Enter>"
EMWaitReady 0, 0

EMSendKey "QQT4"
EMSendKey "<Enter>"
EMWaitReady 0, 0
