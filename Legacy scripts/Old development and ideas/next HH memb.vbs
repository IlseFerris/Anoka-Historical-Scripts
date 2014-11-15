EMConnect ""

HH_memb_row = 05

BeginDialog Dialog1, 0, 0, 191, 67, "Dialog"
  ButtonGroup ButtonPressed
    PushButton 5, 5, 80, 10, "Previous panel", Button6
    PushButton 5, 20, 80, 10, "Next Panel", Button8
    PushButton 5, 35, 80, 10, "Previous HH memb", Button5
    PushButton 5, 50, 80, 10, "Next HH memb", Button7
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
EndDialog


Do
Dialog Dialog1


If ButtonPressed = 0 then stopscript

If ButtonPressed = 1 then EMReadScreen stat_check, 4, 20, 21
If stat_check = "STAT" and ButtonPressed = 1 then EMReadScreen current_panel, 1, 2, 73
If stat_check = "STAT" and ButtonPressed = 1 and current_panel = 1 then new_panel = current_panel
If stat_check = "STAT" and ButtonPressed = 1 and current_panel > 1 then new_panel = current_panel - 1
If stat_check = "STAT" and ButtonPressed = 1 and amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
If stat_check = "STAT" and ButtonPressed = 1 and amount_of_panels > 1 then EMSendKey "<enter>"
If stat_check = "STAT" and ButtonPressed = 1 and amount_of_panels > 1 then EMWaitReady 1, 0


If ButtonPressed = 2 then EMReadScreen stat_check, 4, 20, 21
If stat_check = "STAT" and ButtonPressed = 2 then EMReadScreen current_panel, 1, 2, 73
If stat_check = "STAT" and ButtonPressed = 2 then EMReadScreen amount_of_panels, 1, 2, 78
If stat_check = "STAT" and ButtonPressed = 2 and current_panel < amount_of_panels then new_panel = current_panel + 1
If stat_check = "STAT" and ButtonPressed = 2 and current_panel = amount_of_panels then new_panel = current_panel
If stat_check = "STAT" and ButtonPressed = 2 and amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
If stat_check = "STAT" and ButtonPressed = 2 and amount_of_panels > 1 then EMSendKey "<enter>"
If stat_check = "STAT" and ButtonPressed = 2 and amount_of_panels > 1 then EMWaitReady 1, 0



If ButtonPressed = 3 then EMReadScreen stat_check, 4, 20, 21
If stat_check = "STAT" and ButtonPressed = 3 then HH_memb_row = HH_memb_row - 1
If stat_check = "STAT" and ButtonPressed = 3 then EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
   If stat_check = "STAT" and ButtonPressed = 3 and prev_HH_memb = "ef" then HH_memb_row = HH_memb_row + 1
   If stat_check = "STAT" and ButtonPressed = 3 and prev_HH_memb = "ef" then EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
If stat_check = "STAT" and ButtonPressed = 3 then EMWriteScreen prev_HH_memb, 20, 76
If stat_check = "STAT" and ButtonPressed = 3 then EMSendKey "<enter>"
If stat_check = "STAT" and ButtonPressed = 3 then EMWaitReady 1, 0


If ButtonPressed = 4 then EMReadScreen stat_check, 4, 20, 21
If stat_check = "STAT" and ButtonPressed = 4 then HH_memb_row = HH_memb_row + 1
If stat_check = "STAT" and ButtonPressed = 4 then EMReadScreen next_HH_memb, 2, HH_memb_row, 3
   If stat_check = "STAT" and ButtonPressed = 4 and next_HH_memb = "  " then HH_memb_row = HH_memb_row - 1
   If stat_check = "STAT" and ButtonPressed = 4 and next_HH_memb = "  " then EMReadScreen next_HH_memb, 2, HH_memb_row, 3
If stat_check = "STAT" and ButtonPressed = 4 then EMWriteScreen next_HH_memb, 20, 76
If stat_check = "STAT" and ButtonPressed = 4 then EMSendKey "<enter>"
If stat_check = "STAT" and ButtonPressed = 4 then EMWaitReady 1, 0


Loop until ButtonPressed = -1