cases_done = 0 'DONT CHANGE THIS, SHOULD ALWAYS BE A 0
row = 7 'DONT CHANGE THIS, SHOULD ALWAYS BE A 7

number_of_cases_to_transfer = inputbox("How many cases are you transferring? Currently a max of 12.")
number_of_cases_to_transfer = cint(number_of_cases_to_transfer)
workers_to_transfer_to = inputbox("Which worker(s) are you XFERing to? Add the last 3 digits of their" & chr(13) & "X102#, separated by a comma.")
workers_to_transfer_to = split(replace(workers_to_transfer_to, " ", ""), ",")

EMConnect ""

EMReadScreen ACTV_check, 4, 2, 48
EMReadScreen PND2_check, 4, 2, 52

'797, 116, 120, 126, 127, 967, 123, 118, 117, 124, 121
'117, 967, 118, 116, 120, 121, 123, 124, 126, 127
'797, 5A7, 5A6, 5A5, 5A4, 5A2, 5A3, 5A1, B83
'797, 5A5, 5A7, 5A6, 5A1, b83, 4es, 395, B93, 122
'PXB, SKM, EMP, SJS, KAS, HLS, 797


If PND2_check = "PND2" then
  case_number_col = 5
  screen_to_return_to = "PND2"
ElseIf ACTV_check = "ACTV" then
  case_number_col = 12
  screen_to_return_to = "ACTV"
Else
  MsgBox "Not on ACTV or PND2"
  StopScript
End if

Do
  For each worker in workers_to_transfer_to
    EMReadScreen case_number, 8, row, case_number_col
    master_array = trim(master_array & " " & trim(case_number) & "|" & worker)
    cases_done = cases_done + 1
    row = row + 1
    If cases_done = number_of_cases_to_transfer then exit for
  Next
Loop until cases_done = number_of_cases_to_transfer

master_array = split(master_array)

for each case_and_worker in master_array
  new_worker = right(case_and_worker, 3)
  case_number = left(case_and_worker, len(case_and_worker) - 4)
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
  'Now we navigate to SPEC/XFER
  EMWriteScreen "SPEC", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen case_number, 18, 43
  EMWriteScreen "XFER", 21, 70
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMWriteScreen "x", 7, 16
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
  EMWriteScreen new_worker, 18, 65
  EMSendKey "<enter>"
  EMWaitReady 0, 0
Next

EMWriteScreen "REPT", 20, 22
EMWriteScreen "________", 20, 38
EMWriteScreen screen_to_return_to, 20, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

MsgBox "transfer complete"

