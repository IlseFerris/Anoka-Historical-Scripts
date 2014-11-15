'SECTION 01

worker_number_array = array("103", "104", "107", "109", "110", "111", "112", "113", "114", "200", "204", "211", "213", "214", "222", "223", "224", "228", "229", "231", "231", "232", "233", "243", "247", "256", "263", "268", "274", "275", "280", "281", "282", "283", "284", "285", "287", "288", "289", "291", "292", "293", "294", "305", "307", "309", "30R", "30X", "30Y", "310", "323", "338", "345", "346", "352", "362", "364", "368", "373", "374", "380", "382", "395", "408", "444", "490", "494", "4AF", "4AL", "4AS", "4B4", "4BL", "4BM", "4BV", "4DK", "4F9", "4G1", "4JE", "4JK", "4JS", "4LD", "4MG", "4MS", "4RH", "4RJ", "4RS", "4SS", "4SW", "4SY", "4SZ", "4TR", "4YK", "4YL", "505", "516", "519", "524", "530", "536", "538", "539", "549", "550", "553", "565", "569", "578", "588", "592", "593", "595", "598", "601", "606", "607", "615", "623", "624", "629", "630", "631", "633", "643", "648", "652", "664", "672", "673", "674", "675", "678", "681", "686", "690", "692", "700", "715", "720", "722", "733", "736", "742", "749", "750", "752", "756", "757", "758", "762", "767", "769", "770", "773", "797", "805", "824", "825", "866", "869", "870", "873", "880", "881", "884", "886", "892", "893", "894", "895", "902", "920", "922", "925", "926", "928", "932", "933", "943", "944", "949", "950", "955", "959", "962", "968", "978", "985", "987", "989", "992", "998", "A01", "A03", "A07", "A12", "A14", "A18", "A19", "A27", "A35", "A40", "A44", "A46", "A51", "A54", "A55", "A62", "A64", "A65", "A74", "A75", "A77", "A84", "B13", "B14", "B15", "B18", "B20", "B24", "B25", "B34", "B35", "B42", "B43", "B46", "B47", "B48", "B50", "B51", "B55", "B58", "B63", "B64", "B68", "B69", "B70", "B72", "B74", "B78", "B83", "B84", "B92", "B93", "B97", "BA1", "BA2", "C02", "C03", "C04", "C05", "C08", "C09", "C10", "C11", "CMB", "ICT", "RLH", "SAC", "SAR", "SAS", "SEC", "TRP", "V57", "106", "218", "225", "234", "30V", "4ES", "4G7", "4SX", "518", "616", "628", "872", "A71", "B36", "B52", "B98", "BED", "C06", "C07", "4SL", "GMZ", "RLM", "TLP", "518")

EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
EMSendKey "<enter>"
EMWaitReady 0, 0
EMReadScreen password_prompt, 38, 2, 23
IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript

row = 1
col = 1
EMSearch "MAXIS", row, col
If row <> 1 then
  MsgBox "You need to run this script in the window that has MAXIS on it. Please try again."
  StopScript
End if

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

ObjExcel.Cells(1, 1).Value = "MAXIS number"
ObjExcel.Cells(1, 2).Value = "Name"
ObjExcel.Cells(1, 3).Value = "x102number"
ObjExcel.Cells(1, 4).Value = "HC status"
ObjExcel.Cells(1, 5).Value = "FS status"
ObjExcel.Cells(1, 6).Value = "cash status"

next_excel_row_start = 2 'This sets the variable for the following.

For each worker_number in worker_number_array

  excel_row = next_excel_row_start
  
  'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"
  
  EMWriteScreen "rept", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen "04", 20, 43 'Forces a footer month/year
  EMWriteScreen "13", 20, 46
  EMWriteScreen "actv", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen worker_number_check, 3, 21, 17
  If worker_number_check <> worker_number then
    EMWriteScreen worker_number, 21, 17
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if
    
  'SECTION 03
    
  MAXIS_row = 7 'This sets the variable for the following do...loop.
  Do
    EMReadScreen last_page_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
    Do
      EMReadScreen case_number, 8, MAXIS_row, 12
      EMReadScreen client_name, 21, MAXIS_row, 21
      EMReadScreen HC_status, 1, MAXIS_row, 64
      EMReadScreen SNAP_status, 1, MAXIS_row, 61
      EMReadScreen cash_status, 9, MAXIS_row, 51
      case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each. This is for the Excel spreadsheet, so that we aren't entering blank spaces.
      client_name = Trim(client_name)
      If case_number <> "" then 
        ObjExcel.Cells(excel_row, 1).Value = case_number   'Then it writes each into the Excel spreadsheet to be used later.
        ObjExcel.Cells(excel_row, 2).Value = client_name
        ObjExcel.Cells(excel_row, 3).Value = worker_number
        ObjExcel.Cells(excel_row, 4).Value = HC_status
        ObjExcel.Cells(excel_row, 5).Value = SNAP_status
        ObjExcel.Cells(excel_row, 6).Value = cash_status
        excel_row = excel_row + 1
      End if
      MAXIS_row = MAXIS_row + 1
    Loop until MAXIS_row = 19
    MAXIS_row = 7 'Setting the variable for when the do...loop restarts
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  Loop until last_page_check = "THIS IS THE LAST PAGE"
  
  excel_row = next_excel_row_start
  

  
  Do
    case_number = ObjExcel.Cells(excel_row, 1).Value
    If case_number = "" then exit do
    HC_status = ObjExcel.Cells(excel_row, 4).Value
    If HC_status = "A" then
      do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 27, 2, 28
      loop until SELF_check = "Select Function Menu (SELF)"
      EMWriteScreen "elig", 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen "hc__", 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0

      EMReadScreen MA_check, 2, 8, 31
      If MA_check = "MA" then
        EMWriteScreen "x", 8, 29
        EMSendKey "<enter>"
        EMWaitReady 0, 0

        row = 1
        col = 1
        EMSearch "14 / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "14"
        row = 1
        col = 1
        EMSearch "AA / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "AA"
        row = 1
        col = 1
        EMSearch "AX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "AX"
        row = 1
        col = 1
        EMSearch "BX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "BX"
        row = 1
        col = 1
        EMSearch "CB / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "CB"
        row = 1
        col = 1
        EMSearch "CK / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "CK"
        row = 1
        col = 1
        EMSearch "CX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "CX"
        row = 1
        col = 1
        EMSearch "DP / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "DP"
        row = 1
        col = 1
        EMSearch "DX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "DX"
        row = 1
        col = 1
        EMSearch "EX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "EX"
        row = 1
        col = 1
        EMSearch "PX / ", row, col
        If row <> 0 then ObjExcel.Cells(excel_row, 7).Value = "PX"
        row = 1
        col = 1
        EMSearch "Excess Inc:                                                                   ", row, col
        If row = 0 then ObjExcel.Cells(excel_row, 8).Value = "spenddown indicated"
      End if
    End if

    excel_row = excel_row + 1
  Loop until case_number = ""
  
  next_excel_row_start = excel_row
  
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
Next

MsgBox "Success!"



