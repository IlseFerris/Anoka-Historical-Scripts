function transmit
  EMSendKey "<enter>"
  EMWaitReady 1, 1
End function

EMConnect ""

HH_memb_array = "01 03 04 05"
HH_memb_array = split(HH_memb_array, " ")

For each HH_memb in HH_memb_array
  EMWriteScreen HH_memb, 20, 76
  transmit
  EMReadScreen client_age, 3, 8, 76
  If cint(client_age) >= 21 then number_of_adults = number_of_adults + 1
  If cint(client_age) < 21 then number_of_children = number_of_children + 1
Next

If number_of_adults > 0 then HH_comp = number_of_adults & "a"
If number_of_children > 0 then HH_comp = HH_comp & ", " & number_of_children & "c"
If left(HH_comp, 1) = "," then HH_comp = right(HH_comp, len(HH_comp) - 1)