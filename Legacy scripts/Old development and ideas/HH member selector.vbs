

EMConnect ""

EMReadScreen HH_member_01, 18, 5, 3
EMReadScreen HH_member_02, 18, 6, 3
EMReadScreen HH_member_03, 18, 7, 3
EMReadScreen HH_member_04, 18, 8, 3
EMReadScreen HH_member_05, 18, 9, 3
EMReadScreen HH_member_06, 18, 10, 3
EMReadScreen HH_member_07, 18, 11, 3
EMReadScreen HH_member_08, 18, 12, 3
EMReadScreen HH_member_09, 18, 13, 3
EMReadScreen HH_member_10, 18, 14, 3
EMReadScreen HH_member_11, 18, 15, 3
EMReadScreen HH_member_12, 18, 16, 3
EMReadScreen HH_member_13, 18, 17, 3
EMReadScreen HH_member_14, 18, 18, 3
EMReadScreen HH_member_15, 18, 19, 3

new_variable = 50

If HH_member_03 <> "                  " then new_variable = 65
If HH_member_04 <> "                  " then new_variable = 80
If HH_member_05 <> "                  " then new_variable = 95
If HH_member_06 <> "                  " then new_variable = 110
If HH_member_07 <> "                  " then new_variable = 125
If HH_member_08 <> "                  " then new_variable = 140
If HH_member_09 <> "                  " then new_variable = 155
If HH_member_10 <> "                  " then new_variable = 170
If HH_member_11 <> "                  " then new_variable = 185
If HH_member_12 <> "                  " then new_variable = 200
If HH_member_13 <> "                  " then new_variable = 215
If HH_member_14 <> "                  " then new_variable = 230
If HH_member_15 <> "                  " then new_variable = 245

BeginDialog HH_memb_dialog, 0, 0, 191, new_variable, "HH member dialog"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 10, 5, 105, 10, "Household members to look at:"
  If HH_member_01 <> "                  " then CheckBox 10, 20, 120, 10, HH_member_01, client_01_check
  If HH_member_02 <> "                  " then CheckBox 10, 35, 120, 10, HH_member_02, client_02_check
  If HH_member_03 <> "                  " then CheckBox 10, 50, 120, 10, HH_member_03, client_03_check
  If HH_member_04 <> "                  " then CheckBox 10, 65, 120, 10, HH_member_04, client_04_check
  If HH_member_05 <> "                  " then CheckBox 10, 80, 120, 10, HH_member_05, client_05_check
  If HH_member_06 <> "                  " then CheckBox 10, 95, 120, 10, HH_member_06, client_06_check
  If HH_member_07 <> "                  " then CheckBox 10, 110, 120, 10, HH_member_07, client_07_check
  If HH_member_08 <> "                  " then CheckBox 10, 125, 120, 10, HH_member_08, client_08_check
  If HH_member_09 <> "                  " then CheckBox 10, 140, 120, 10, HH_member_09, client_09_check
  If HH_member_10 <> "                  " then CheckBox 10, 155, 120, 10, HH_member_10, client_10_check
  If HH_member_11 <> "                  " then CheckBox 10, 170, 120, 10, HH_member_11, client_11_check
  If HH_member_12 <> "                  " then CheckBox 10, 185, 120, 10, HH_member_12, client_12_check
  If HH_member_13 <> "                  " then CheckBox 10, 200, 120, 10, HH_member_13, client_13_check
  If HH_member_14 <> "                  " then CheckBox 10, 215, 120, 10, HH_member_14, client_14_check
  If HH_member_15 <> "                  " then CheckBox 10, 230, 120, 10, HH_member_15, client_15_check
EndDialog
Dialog HH_memb_dialog