'FUNCTIONS----------------------------------------------------------------------------------------------------
Function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
End function

Function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

x102_array = array("004", "006", "022", "024", "030", "032", "033", "038", "042", "045", "051", "054", "055", "058", "060", "075", "081", "085", "087", "093", "097", "103", "104", "106", "107", "109", "110", "111", "112", "113", "114", "200", "204", "211","213", "214", "218", "222", "223", "224", "225", "228", "229", "231", "231", "232", "233", "234", "243", "247", "256", "263", "268", "274", "275", "280", "281", "282", "283", "284", "285", "287", "288", "289", "291", "292", "293", "294", "305", "307", "309", "30R", "30V", "30X", "30Y", "310", "323", "338", "345", "346", "352", "362", "364", "368", "373", "374", "380", "382", "395", "408", "444", "490", "494", "4AF", "4AL", "4AS", "4B4", "4BL", "4BM", "4BV", "4DK", "4ES", "4F9", "4G1", "4G7", "4JE", "4JK", "4JS", "4LD", "4MG", "4MS", "4RH", "4RJ", "4RS", "4SL", "4SS", "4SW", "4SX", "4SY", "4SZ", "4TR", "4YK", "4YL", "505", "516", "518", "519", "524", "530", "536", "538", "539", "549", "550", "553", "565", "569", "578", "588", "592", "593", "595", "598", "601", "606", "607", "615", "616", "623", "624", "628", "629", "630", "631", "633", "643", "648", "652", "664", "672", "673", "674", "675", "678", "681", "686", "690", "692", "700", "715", "720", "722", "733", "736", "742", "749", "750", "752", "756", "757", "758", "762", "767", "769", "770", "773", "797", "805", "824", "825", "866", "869", "870", "872", "873", "880", "881", "884", "886", "892", "893", "894", "895", "902", "920", "922", "925", "926", "928", "932", "933", "943", "944", "949", "950", "955", "959", "962", "968", "978", "985", "987", "989", "992", "998", "A01", "A03", "A07", "A12", "A14", "A18", "A19", "A27", "A35", "A40", "A44", "A46", "A51", "A54", "A55", "A62", "A64", "A65", "A71", "A74", "A75", "A77", "A84", "B13", "B14", "B15", "B18", "B20", "B24", "B25", "B34", "B35", "B36", "B42", "B43", "B46", "B47", "B48", "B50", "B51", "B52", "B55", "B58", "B63", "B64", "B68", "B69", "B70", "B72", "B74", "B78", "B82", "B83", "B84", "B92", "B93", "B97", "B98", "BA1", "BA2", "BED", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09", "C10", "C11", "CMB", "GMZ", "ICT", "RLH", "RLM", "SAC", "SAR", "SAS", "SEC", "TLP", "TRP", "V57")

EMConnect ""

start_time = timer

EMReadScreen PND2_check, 4, 2, 52
If PND2_check <> "PND2" then 
  MsgBox "Navigate to REPT/PND2 before proceeding."
  stopscript
End if
excel_row_variable_col_1 = 2


Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
strFileName = "h:\test.xlsx"  
Set objWorkbook = objExcel.Workbooks.Add() 
ObjExcel.Cells(1, 1).Value = "x102"
ObjExcel.Cells(1, 2).Value = "M# with FS"
ObjExcel.Cells(1, 3).Value = "FS client delay indicator"

For each x102_number in x102_array
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
  EMSendKey "pnd2" & "<enter>"
  EMWaitReady 0, 0
  EMWriteScreen x102_number, 21, 17
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSetCursor 21, 13
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
  EMReadScreen supervisor_name, 20, 22, 16
  EMSendKey "<enter>"
  EMWaitReady 0, 0

  EMReadScreen PND2_amt_check, 6, 3, 74
  If PND2_amt_check <> "0 Of 0" then 'skips workers with no pending cases
    Do
      MAXIS_row = 7
      Do
        EMReadScreen FS_status_code, 1, MAXIS_row, 62
        EMReadScreen HC_status_code, 1, MAXIS_row, 65
        EMReadScreen case_number, 8, MAXIS_row, 5
        If (FS_status_code <> "_" or HC_status_code <> "_") and FS_status_code <> " " then
          ObjExcel.Cells(excel_row_variable_col_1, 1).Value = x102_number
          ObjExcel.Cells(excel_row_variable_col_1, 2).Value = case_number
          If FS_status_code <> "_" then ObjExcel.Cells(excel_row_variable_col_1, 3).Value = FS_status_code
          If HC_status_code = "P" then
            EMWriteScreen "x", MAXIS_row, 3
            transmit
            EMReadScreen HC_delay_notice_code, 1, 7, 39
            ObjExcel.Cells(excel_row_variable_col_1, 4).Value = HC_delay_notice_code
            PF3
          End if
          ObjExcel.Cells(excel_row_variable_col_1, 5).Value = supervisor_name
          excel_row_variable_col_1 = excel_row_variable_col_1 + 1
        End if
        MAXIS_row = MAXIS_row + 1
      Loop until FS_status_code = " "
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
      EMReadScreen last_page_check, 21, 24, 02
    Loop until last_page_check = "THIS IS THE LAST PAGE"
  End if
Next

stop_time = timer

MsgBox stop_time - start_time