EMConnect ""
EMReadscreen test, 10, 10, 22
EMReadscreen test2, 10, 11, 22

BeginDialog Dialog1, 0, 0, 191, 52, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  OptionGroup RadioGroup1
    RadioButton 10, 5, 50, 10, test, Radio1
    RadioButton 10, 20, 60, 10, test2, Radio2
EndDialog

dialog Dialog1