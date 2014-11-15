

EMConnect ""

BeginDialog message_dialog, 0, 0, 306, 42, "Message dialog"
  DropListBox 45, 5, 255, 12, "COPAY MONTH IS DUPLICATE, BYPASSED"+chr(9)+"MEDICARE ID# DOES NOT MATCH # ON CLAIM. VRFY ID#."+chr(9)+"PENDING AWAITING PAYMENT RECIPIENT MADE ACTIVE ONGOING.", error_message
  ButtonGroup message_dialog_ButtonPressed
    OkButton 100, 25, 50, 15
    CancelButton 160, 25, 50, 15
  Text 5, 5, 40, 10, "Message:"
EndDialog


Dialog message_dialog

If message_dialog_ButtonPressed = 0 then stopscript

If error_message = "COPAY MONTH IS DUPLICATE, BYPASSED" then MsgBox "This message is informational. No action is needed."
If error_message = "COPAY MONTH IS DUPLICATE, BYPASSED" then stopscript

If error_message = "MEDICARE ID# DOES NOT MATCH # ON CLAIM. VRFY ID#." then run "H:\BlueZone\bzsh.exe Q:\Blue Zone Scripts\Adult\DWMR scrubber scripts\DWMR - MEDICARE ID.vbs"
If error_message = "MEDICARE ID# DOES NOT MATCH # ON CLAIM. VRFY ID#." then stopscript

If error_message = "PENDING AWAITING PAYMENT RECIPIENT MADE ACTIVE ONGOING." then run "H:\BlueZone\bzsh.exe Q:\Blue Zone Scripts\Adult\DWMR scrubber scripts\DWMR - pending made active.vbs"
If error_message = "PENDING AWAITING PAYMENT RECIPIENT MADE ACTIVE ONGOING." then stopscript


msgbox "1"