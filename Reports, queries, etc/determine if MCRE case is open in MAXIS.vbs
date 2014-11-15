Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\April MCRE cases.xlsx"  
Set objWorkbook = objExcel.Workbooks.Open(strFileName) 

ObjExcel.Cells(1, 7).Value = "current MCRE worker"
ObjExcel.Cells(1, 8).Value = "current MCRE status"
ObjExcel.Cells(1, 9).Value = "current MCRE status reason"
ObjExcel.Cells(1, 10).Value = "PMI"
ObjExcel.Cells(1, 11).Value = "MAXIS case number"
ObjExcel.Cells(1, 12).Value = "MAXIS worker"
ObjExcel.Cells(1, 13).Value = "MAXIS status"

excel_row = 2 'Setting the variable



EMConnect ""

Sub break_out
EMSendKey "<attn>"
EMWaitReady 1, 2
EMSendKey "10" + "<enter>"
EMWaitReady 1, 2


EMReadScreen RKEY_check, 4, 1, 52
If RKEY_check <> "RKEY" then
  MsgBox "You are not on RKEY. The script will now stop."
  Stopscript
End If

Do
  MCRE_number = ObjExcel.Cells(excel_row, 5).Value
  EMWriteScreen "i", 2, 19
  EMWriteScreen MCRE_number, 9, 19
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen RCAD_check, 4, 1, 50
  If RCAD_check <> "RCAD" then
    MsgBox "You are not on RCAD. The script will now stop."
    Stopscript
  End If
  EMWriteScreen "RCIN", 1, 8
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen current_MCRE_worker, 7, 2, 46
  EMReadScreen current_MCRE_status, 13, 5, 16
  EMReadScreen current_MCRE_status_reason, 14, 5, 44
  EMReadScreen person_01_PMI, 8, 11, 04
  ObjExcel.Cells(excel_row, 7).Value = current_MCRE_worker
  ObjExcel.Cells(excel_row, 8).Value = trim(current_MCRE_status)
  ObjExcel.Cells(excel_row, 9).Value = trim(current_MCRE_status_reason)
  ObjExcel.Cells(excel_row, 10).Value = person_01_PMI
  EMSendKey "<PF6>"
  EMWaitReady 1, 1
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 5).Value = ""
  

EMSendKey "<attn>"
EMWaitReady 1, 2
EMSendKey "<attn>"
EMWaitReady 1, 2


  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then
      MsgBox "MAXIS could not be found. The script will now stop."
      StopScript
    End if
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
  EMWriteScreen "pers", 16, 43
  EMSendKey "<enter>"
  EMWaitReady 1, 1

excel_row = 2 'resetting the variable

Do
  EMWriteScreen "________", 15, 36
  EMWriteScreen ObjExcel.Cells(excel_row, 10).Value, 15, 36
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen MTCH_check, 4, 2, 51
  If MTCH_check <> "MTCH" then stopscript
  EMWriteScreen "x", 8, 5
  EMSendKey "<enter>"
  EMWaitReady 1, 1
    row = 1
    col = 1
  EMSearch "   Y    ", row, col
  If row <> 0 then
    EMReadScreen case_number, 8, row, 6
    EMReadScreen MAXIS_worker_number, 7, row, 71
  End if
  If row = 0 then
    Do
        row = 1
        col = 1
      EMSendKey "<PF8>"
      EMWaitReady 1, 1
      EMSearch "   Y    ", row, col
      EMReadScreen page_check, 11, 24, 2
      If page_check = "THIS IS THE" and row = 0 then 
        case_number = "None found"
        Exit do
      End if
      If row <> 0 then
        EMReadScreen case_number, 8, row, 6
        EMReadScreen MAXIS_worker_number, 7, row, 71
      End if
    Loop until row <> 0
  End if

  ObjExcel.Cells(excel_row, 11).Value = case_number
  ObjExcel.Cells(excel_row, 12).Value = MAXIS_worker_number

  excel_row = excel_row + 1
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen PERS_check, 4, 2, 47
  Loop until PERS_check = "PERS"
Loop until ObjExcel.Cells(excel_row, 10).Value = ""

end sub


excel_row = 2 'resetting the variable

Do
  Do
    EMSendKey "<PF3>"
    EMWaitReady 1, 1
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then
      MsgBox "MAXIS could not be found. The script will now stop."
      StopScript
    End if
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
  Do
    case_number = ObjExcel.Cells(excel_row, 11).Value
    If case_number = "None found" then excel_row = excel_row + 1
  Loop until case_number <> "None found" or ObjExcel.Cells(excel_row, 11).Value = ""

  EMWriteScreen "case", 16, 43
  EMWriteScreen "        ", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "curr", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 1, 1

  EMReadScreen priv_check, 10, 24, 14
  If priv_check = "PRIVILEGED" then ObjExcel.Cells(excel_row, 10).Value = "Privileged"
  If priv_check <> "PRIVILEGED" then EMReadScreen CURR_check, 4, 2, 55
  If priv_check <> "PRIVILEGED" and CURR_check <> "CURR" then MsgBox "Something went wrong, this isn't CURR!"
  If priv_check <> "PRIVILEGED" and CURR_check <> "CURR" then stopscript

    row = 1
    col = 1
  EMSearch "Case: INACTIVE", row, col
  If row <> 0 then MAXIS_status = "Inactive"
  If row = 0 then MAXIS_status = "Active/pending"

  ObjExcel.Cells(excel_row, 13).Value = MAXIS_status



  EMSendKey "<PF3>"
  EMWaitReady 1, 1

  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 11).Value = ""