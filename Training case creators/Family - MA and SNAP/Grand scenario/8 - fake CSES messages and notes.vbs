'
' 12:00:42  06/03/2013
' BlueZone v5.1C2 Recorded Visual Basic Script File
'

EMConnect ""

EMSendKey "n"
EMSendKey "<Enter>"
EMWaitReady 10, 1

EMSendKey "<PF9>"
EMWaitReady 10, 1

EMSendKey "<Home>"
EMSendKey "DISB CS (TYPE 36) OF $0.00 FOR 1 CHILD(REN) ISSUED ON"
EMSendKey "<NewLine>"
EMSendKey "04/01/13 TO PMI(S): XXXXXXXX"
EMSendKey "<PF3>"
EMWaitReady 10, 1

EMSendKey "<PF9>"
EMWaitReady 10, 1

EMSendKey "<Home>"
EMSendKey "DISB CS (TYPE 36) OF $80.00 FOR 1 CHILD(REN) ISSUED ON"
EMSendKey "<NewLine>"
EMSendKey "05/01/13 TO PMI(S): XXXXXXXX"
EMSendKey "<PF3>"
EMWaitReady 10, 1

EMSendKey "<PF9>"
EMWaitReady 10, 1

EMSendKey "<Home>"
EMSendKey "DISB CS (TYPE 36) OF $50.00 FOR 1 CHILD(REN) ISSUED ON"
EMSendKey "<NewLine>"
EMSendKey "06/01/13 TO PMI(S): XXXXXXXX"
EMSendKey "<PF3>"
EMWaitReady 10, 1

EMSetCursor 20, 22
EMSendKey "DAIL"
EMSetCursor 20, 70
EMSendKey "WRIT"
EMSendKey "<Enter>"
EMWaitReady 10, 1

EMSendKey "<NewLine>"
EMSendKey "<Home>"
EMSendKey "<NewLine>"
EMSendKey "DISB CS (TYPE 36) OF $50.00 FOR 1 CHILD(REN) ISSUED ON"
EMSendKey "<NewLine>"
EMSendKey "06/01/13 TO PMI(S): XXXXXXXX"
EMSendKey "<PF3>"
EMWaitReady 10, 1

EMWriteScreen "REPT", 20, 22
EMWriteScreen "ACTV", 20, 70
EMSendKey "<Enter>"
EMWaitReady 10, 1