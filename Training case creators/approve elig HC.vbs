Function transmit
  EMSendKey "<Enter>"
  EMWaitReady 0, 0
End function

Function new_line
  EMSendKey "<newline>"
  EMWaitReady 0, 0
End function

Function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
End function

Function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
End function

EMConnect ""

Do

EMSendKey "e"
transmit

EMSendKey "hc"
transmit

EMSendKey "x"
transmit

transmit

EMSendKey "app"
transmit

EMSendKey "x"
transmit

EMSendKey "y"
transmit

PF3

new_line

EMGetCursor row, col

Loop until row > 18

If row > 18 then
  PF8
  EMSetCursor 7, 3
End if