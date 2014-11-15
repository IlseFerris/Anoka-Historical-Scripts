EMConnect ""

BeginDialog Dialog1, 0, 0, 191, 62, "Dialog1"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog2, 0, 0, 191, 62, "Dialog2"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog3, 0, 0, 191, 62, "Dialog3"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog4, 0, 0, 191, 62, "Dialog4"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog5, 0, 0, 191, 62, "Dialog5"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog6, 0, 0, 191, 62, "Dialog6"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog7, 0, 0, 191, 62, "Dialog7"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog8, 0, 0, 191, 62, "Dialog8"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

BeginDialog Dialog9, 0, 0, 191, 62, "Dialog9"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
    PushButton 5, 5, 45, 20, "prev", Button3
    PushButton 70, 5, 55, 20, "next", Button5
EndDialog

Sub Dialog_1
EMSendKey "<enter>"
Dialog Dialog1
If buttonpressed = 2 then call Dialog_2
If buttonpressed = 1 then call Dialog_1
End Sub

Sub Dialog_2
EMSendKey "<enter>"
Dialog Dialog2
If buttonpressed = 2 then call Dialog_3
If buttonpressed = 1 then call Dialog_1
End Sub

Sub Dialog_3
EMSendKey "<enter>"
Dialog Dialog3
If buttonpressed = 2 then call Dialog_4
If buttonpressed = 1 then call Dialog_2
End Sub

Sub Dialog_4
EMSendKey "<enter>"
Dialog Dialog4
If buttonpressed = 2 then call Dialog_5
If buttonpressed = 1 then call Dialog_3
End Sub

Sub Dialog_5
EMSendKey "<enter>"
Dialog Dialog5
If buttonpressed = 2 then call Dialog_6
If buttonpressed = 1 then call Dialog_4
End Sub

Sub Dialog_6
EMSendKey "<enter>"
Dialog Dialog6
If buttonpressed = 2 then call Dialog_7
If buttonpressed = 1 then call Dialog_5
End Sub

Sub Dialog_7
EMSendKey "<enter>"
Dialog Dialog7
If buttonpressed = 2 then call Dialog_8
If buttonpressed = 1 then call Dialog_6
End Sub

Sub Dialog_8
EMSendKey "<enter>"
Dialog Dialog8
If buttonpressed = 2 then call Dialog_9
If buttonpressed = 1 then call Dialog_7
End Sub

Sub Dialog_9
EMSendKey "<enter>"
Dialog Dialog9
If buttonpressed = 2 then call Dialog_9
If buttonpressed = 1 then call Dialog_8
End Sub

Dialog_1

