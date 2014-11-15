EMConnect ""

'The following will search for an application case note.

   row = 1
   col = 1

EMSearch "***", row, col
EMSetCursor row, 3
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0

'Now it will search the application case note for the verifs requested section
Do
  row = 1
  col = 1
EMSearch "* Verifs needed: ", row, col
If row = 0 and col = 0 then EMSendKey "<PF8>"
EMWaitReady 1, 0
Loop until row <> 0 and col <> 0

If row = 17 and col = 3 then EMReadScreen verifs_needed, 60, 17, 20
If row = 17 and col = 3 then EMSendKey "<PF8>"
If row = 17 and col = 3 then EMWaitReady 1, 0
If row = 17 and col = 3 then EMReadScreen verifs_needed_second_page, 75, 4, 3

If row <> 17 then EMReadScreen verifs_needed, 160, row, col 

Msgbox verifs_needed & verifs_needed_second_page