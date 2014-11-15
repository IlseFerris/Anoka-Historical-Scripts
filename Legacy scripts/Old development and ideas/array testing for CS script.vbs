EMConnect ""



Sub FS_sub
  'What follows figures out the lowest_amt and highest_amt of FS on the PIC.
  Dim income_received(4)
  EMReadScreen income_received(0), 8, 9, 25
  EMReadScreen income_received(1), 8, 10, 25
  EMReadScreen income_received(2), 8, 11, 25
  EMReadScreen income_received(3), 8, 12, 25
  EMReadScreen income_received(4), 8, 13, 25
  If income_received(0) < income_received(1) and income_received(0) < income_received(2) and income_received(0) < income_received(3) and income_received(0) < income_received(4) then lowest_amt = abs(income_received(0))
  If income_received(1) < income_received(0) and income_received(1) < income_received(2) and income_received(1) < income_received(3) and income_received(1) < income_received(4) then lowest_amt = abs(income_received(1))
  If income_received(2) < income_received(1) and income_received(2) < income_received(0) and income_received(2) < income_received(3) and income_received(2) < income_received(4) then lowest_amt = abs(income_received(2))
  If income_received(3) < income_received(1) and income_received(3) < income_received(2) and income_received(3) < income_received(0) and income_received(3) < income_received(4) then lowest_amt = abs(income_received(3))
  If income_received(4) < income_received(1) and income_received(4) < income_received(2) and income_received(4) < income_received(3) and income_received(4) < income_received(0) then lowest_amt = abs(income_received(4))
  If income_received(0) = "________" then income_received(0) = 0
  If income_received(1) = "________" then income_received(1) = 0
  If income_received(2) = "________" then income_received(2) = 0
  If income_received(3) = "________" then income_received(3) = 0
  If income_received(4) = "________" then income_received(4) = 0
  If income_received(0) > income_received(1) and income_received(0) > income_received(2) and income_received(0) > income_received(3) and income_received(0) > income_received(4) then highest_amt = abs(income_received(0))
  If income_received(1) > income_received(0) and income_received(1) > income_received(2) and income_received(1) > income_received(3) and income_received(1) > income_received(4) then highest_amt = abs(income_received(1))
  If income_received(2) > income_received(1) and income_received(2) > income_received(0) and income_received(2) > income_received(3) and income_received(2) > income_received(4) then highest_amt = abs(income_received(2))
  If income_received(3) > income_received(1) and income_received(3) > income_received(2) and income_received(3) > income_received(0) and income_received(3) > income_received(4) then highest_amt = abs(income_received(3))
  If income_received(4) > income_received(1) and income_received(4) > income_received(2) and income_received(4) > income_received(3) and income_received(4) > income_received(0) then highest_amt = abs(income_received(4))
End Sub

FS_sub