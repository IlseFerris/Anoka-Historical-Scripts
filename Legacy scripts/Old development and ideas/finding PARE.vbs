EMConnect ""

row = 5
Do 
  EMReadScreen HH_memb_number, 2, row, 3
  If HH_memb_number = "  " then exit do
  HH_memb_array = HH_memb_array & HH_memb_number & "|" 
  row = row + 1
Loop until HH_memb_number = "  "

HH_memb_array = split(HH_memb_array, "|")

For each x in HH_memb_array
  EMWriteScreen "pare", 20, 71
  EMWriteScreen x, 20, 76
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen panel_check, 1, 2, 78
  If panel_check = "1" then parents_array = parents_array & x & "|"
Next

parents_array = split(parents_array, "|")
amt_of_parents = ubound(parents_array) - 1
redim parents_array(amt_of_parents, 1)

For each x in parents_array
  EMWriteScreen "pare", 20, 71
  EMWriteScreen x, 20, 76
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  parents_array(x, 1) = "yes"
Next