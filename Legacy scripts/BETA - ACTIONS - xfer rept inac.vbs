'The following creates a new claims processing document.
MsgBox "This script creates a file on your H drive called ''claims processing''. Open this file after this script runs, and email fiscal with any case numbers indicated. The script has a delayed start while it generates this file."
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
   Set objSelection = objWord.Selection
   objselection.typetext " the proceeding are any case numbers with active claims for this month. Email them to fiscal."
   objselection.TypeParagraph()
   objDoc.SaveAs("H:\claims processing.docx")
   objWord.Quit


EMConnect ""

EMReadScreen INAC_check, 44, 2, 21
If INAC_check <> "Workers Monthly Inactive Cases Report (INAC)" then msgbox "You are not on your REPT/INAC."
If INAC_check <> "Workers Monthly Inactive Cases Report (INAC)" then stopscript
EMReadScreen case_number, 8, 7, 3
EMReadScreen worker_number, 7, 21, 16 'This gets the worker number to ensure that REPT/INAC remains in the right worker's number.

If case_number = "        " then MSgbox "No inactive cases..."
If case_number = "        " then stopscript
EMSendKey "<attn>"
EMWaitReady 1, 0
EMReadScreen MMIS_check, 7, 15, 15 
If MMIS_check <> "RUNNING" then MsgBox "At this time, you will need MAXIS and MMIS running in the same window. When you are done with the script, you can close MMIS again and move it where you want it."
If MMIS_check <> "RUNNING" then stopscript
EMSendKey "10" + "<enter>"      
EMWaitReady 1, 0


  Sub get_to_session_begin 'This sub uses a Do Loop to get to the start screen for MMIS.
    Do 
    EMSendkey "<PF6>"
      EMReadScreen password_prompt2, 38, 2, 23
      IF password_prompt2 = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
    EMWaitReady 1, 0
    EMReadScreen session_start, 18, 1, 7
    Loop until session_start = "SESSION TERMINATED"
  End Sub

get_to_session_begin
EMSetCursor 1, 2
EMSendKey "mw00"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSendKey "<enter>"
EMWaitReady 1, 0

'This section may not work for all OSAs, since some only have EK01.
  row = 1
  col = 1
EMSearch "EK01", row, col
If row <> 0 then EMSetCursor row, 4
If row <> 0 then EMSendKey "x"
If row <> 0 then EMSendKey "<enter>"
If row <> 0 then EMWaitReady 1, 0

'This section starts from EK01. OSAs may need to skip the previous section.
EMSetCursor 10, 3
EMSendKey "x"
EMSendKey "<enter>"
EMWaitReady 1, 0
EMFocus

'Now we are in MMIS, and it will set up an inquiry screen
EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_number

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now we enter to find RELG and determine if the case is still open.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSetCursor 1, 08
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads RELG to determine the current status for this SSN
EMReadScreen MMIS_active_date, 8, 7, 36
IF MMIS_active_date = "99/99/99" then MMIS_active_status = "active" 'Sets active/inactive status based on elig end date.
EMReadScreen MMIS_case_number, 8, 6, 73 'Reading MMIS case number

'If the case is still active, it should not be transferred. The next two lines will close the script and notify the worker if the case is still active.
If MMIS_active_status = "active" then MsgBox "This case is active in MMIS on case number " + MMIS_case_number + ". Either manually transfer this case to the HC team, or manually transfer to CLS (for example, if the case is active at MCRE ops). NOTE: copy and paste your claims report before running this script again."
If MMIS_active_status = "active" then stopscript

'Now it gets out of RELG
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

'Now it goes back into MAXIS to check the case against CCOL.

EMSendKey "<attn>"
EMWaitReady 1, 5
EMSendKey "1" + "<enter>" 
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0

'Now we're getting to CCOL/CLIC for the case
EMSetCursor 16, 43
EMSendKey "ccol"
EMSetCursor 18, 43
EMSendKey case_number
EMSetCursor 21, 70
EMSendKey "clic" + "<enter>"
EMWaitReady 1, 0


Sub create_word_doc
   'Now it gears up a word doc to pick up any claims
   Set objWord = CreateObject("Word.Application")
   'note: if you don’t want word popping up and displaying as this is built– set this next line to False
   objWord.Visible = True
   set objDoc = objWord.Documents.open("H:\claims processing.docx")

   Set objSelection = objWord.Selection
   objselection.Font.Name = "Courier New"
   objselection.Font.Size = "12"
   objselection.typetext case_number
   objselection.TypeParagraph()
   objDoc.SaveAs("H:\claims processing.docx")
end sub

'Now it checks for claims.

EMReadScreen CLIC_message, 55, 24, 2
If CLIC_message <> "NO CLAIMS WERE FOUND FOR THIS CASE, PROGRAM, AND STATUS" then call create_word_doc

'Now it case notes that the case has closed.
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 18, 43
EMSendKey case_number
EmSetCursor 21, 70
EMSendKey "note" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<Pf9>"
EMWaitReady 1, 0
EMSendKey "Case is closed, XFERed to CLS."

'Now it returns to SELF to XFER the case.
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "spec"
EMSetCursor 18, 43
EMSendKey case_number
EmSetCursor 21, 70
EMSendKey "xfer" + "<enter>"
EMWaitReady 1, 0

'Now it XFERs the case and returns to the SELF menu.
EMSetCursor 7, 16
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<Pf9>"
EMWaitReady 1, 0
EMSetCursor 18, 61
EMSendKey "x102cls" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0

'Now it returns to rept/inac
EMSetCursor 16, 43
EMSendKey "rept"
EMSetCursor 18, 43
EMSendKey "<eraseeof>"
EmSetCursor 21, 70
EMSendKey "inac" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 21, 16
EMSendKey worker_number + "<enter>"
EMWaitReady 1, 0




'================================Now the script starts over again as a sub, so that it can run until all the cases are XFERed.



Sub rest_of_inactive


EMReadScreen INAC_check, 44, 2, 21
If INAC_check <> "Workers Monthly Inactive Cases Report (INAC)" then msgbox "You are not on your REPT/INAC."
If INAC_check <> "Workers Monthly Inactive Cases Report (INAC)" then stopscript
EMReadScreen case_number, 8, 7, 3
If case_number = "        " then MSgbox "Out of inactive cases."
If case_number = "        " then stopscript
EMSendKey "<attn>"
EMWaitReady 1, 0
EMSendKey "10" + "<enter>" 
EMWaitReady 1, 0

'Now we are in MMIS, and it will set up an inquiry screen
EMSetCursor 2, 19
EMSendKey "i"
EMSetCursor 9, 19
EMSendKey case_number

'Because a case number is never 8 digits, and MMIS requires it, the following will fill the vacant space with zeroes.
EMReadscreen first_MMIS_number_position, 1, 9, 19
EMSetCursor 9, 19
If first_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen second_MMIS_number_position, 1, 9, 20
EMSetCursor 9, 20
If second_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen third_MMIS_number_position, 1, 9, 21
EMSetCursor 9, 21
If third_MMIS_number_position = " " then EMSendKey "0"
EMReadscreen fourth_MMIS_number_position, 1, 9, 22
EMSetCursor 9, 22
If fourth_MMIS_number_position = " " then EMSendKey "0"

'Now we enter to find RELG and determine if the case is still open.
EMSendKey "<enter>"
EMWaitReady 1, 0
EMSetCursor 1, 08
EMSendKey "rcin" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 11, 2
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "relg" + "<enter>"
EMWaitReady 1, 0

'Now it reads RELG to determine the current status for this SSN
EMReadScreen MMIS_active_date, 8, 7, 36
IF MMIS_active_date = "99/99/99" then MMIS_active_status = "active" 'Sets active/inactive status based on elig end date.
EMReadScreen MMIS_case_number, 8, 6, 73 'Reading MMIS case number

'If the case is still active, it should not be transferred. The next two lines will close the script and notify the worker if the case is still active.
If MMIS_active_status = "active" then MsgBox "This case is active in MMIS on case number " + MMIS_case_number + ". Either manually transfer this case to the HC team, or manually transfer to CLS (for example, if the case is active at MCRE ops). NOTE: copy and paste your claims report before running this script again."
If MMIS_active_status = "active" then stopscript

'Now it gets out of RELG
EMSendKey "<PF6>"
EMWaitReady 1, 0
EMSendKey "<PF6>"
EMWaitReady 1, 0

'Now it goes back into MAXIS to check the case against CCOL.

EMSendKey "<attn>"
EMWaitReady 1, 0
EMSendKey "1" + "<enter>" 
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0

'Now we're getting to CCOL/CLIC for the case
EMSetCursor 16, 43
EMSendKey "ccol"
EMSetCursor 18, 43
EMSendKey case_number
EMSetCursor 21, 70
EMSendKey "clic" + "<enter>"
EMWaitReady 1, 0

'Now it checks for claims.

EMReadScreen CLIC_message, 55, 24, 2
If CLIC_message <> "NO CLAIMS WERE FOUND FOR THIS CASE, PROGRAM, AND STATUS" then call create_word_doc

'Now it case notes that the case has closed.
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 18, 43
EMSendKey case_number
EmSetCursor 21, 70
EMSendKey "note" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<Pf9>"
EMWaitReady 1, 0
EMSendKey "Case is closed, XFERed to CLS."

'Now it returns to SELF to XFER the case.
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSetCursor 16, 43
EMSendKey "spec"
EMSetCursor 18, 43
EMSendKey case_number
EmSetCursor 21, 70
EMSendKey "xfer" + "<enter>"
EMWaitReady 1, 0

'Now it XFERs the case and returns to the SELF menu.
EMSetCursor 7, 16
EMSendKey "x" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<Pf9>"
EMWaitReady 1, 0
EMSetCursor 18, 61
EMSendKey "x102cls" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0
EMSendKey "<Pf3>"
EMWaitReady 1, 0

'Now it returns to rept/inac
EMSetCursor 16, 43
EMSendKey "rept"
EMSetCursor 18, 43
EMSendKey "<eraseeof>"
EmSetCursor 21, 70
EMSendKey "inac" + "<enter>"
EMWaitReady 1, 0
EMSetCursor 21, 16
EMSendKey worker_number + "<enter>"
EMWaitReady 1, 0




End Sub

Do
call rest_of_inactive
Loop until case_number = "        "
