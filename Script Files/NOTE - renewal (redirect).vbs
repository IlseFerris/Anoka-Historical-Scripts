'This grabs the user ID and pumps it into a renewal_user_ID variable
Set objNet_renewal = CreateObject("WScript.NetWork") 
renewal_user_ID = objNet_renewal.UserName

'This grabs the OS, XP uses a different path than 7
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")        
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
  system_type = objOperatingSystem.Caption
Next

'Sets the path based on the system_type variable
If instr(system_type, "Windows 7") <> 0 then
  program_path = "C:\Users\" & renewal_user_ID & "\Documents\"
Else
  program_path = "H:\"
End if

'The dialog
BeginDialog renewal_dialog, 0, 0, 71, 127, "Renewal Dialog"
  ButtonGroup ButtonPressed
    PushButton 5, 5, 60, 10, "CAF", CAF_button
    PushButton 5, 20, 60, 10, "Combined AR", Combined_AR_button
    PushButton 5, 35, 60, 10, "CSR", CSR_button
    PushButton 5, 50, 60, 10, "HC ER", HC_ER_button
    PushButton 5, 65, 60, 10, "HRF (Family)", HRF_family_button
    CancelButton 5, 110, 60, 15
EndDialog

'Shows dialog
Dialog renewal_dialog
If buttonpressed = 0 then stopscript

'Doesn't run as script because that introduced errors in future dialogs for some reason.----------------------------------------------------------------------------------------------------
' So, it runs it as a file right from the script host.
If buttonpressed = CAF_button then
  run program_path & "BlueZone\bzsh.exe Q:\Blue Zone Scripts\Script Files\NOTE - CAF.vbs"
  StopScript
End if

If buttonpressed = CSR_button then
  run program_path & "BlueZone\bzsh.exe Q:\Blue Zone Scripts\Script Files\NOTE - CSR.vbs"
  StopScript
End if

If buttonpressed = Combined_AR_button then
  run program_path & "BlueZone\bzsh.exe Q:\Blue Zone Scripts\Script Files\NOTE - Combined AR.vbs"
  StopScript
End if

If buttonpressed = HC_ER_button then
  run program_path & "BlueZone\bzsh.exe Q:\Blue Zone Scripts\Script Files\NOTE - HC ER.vbs"
  StopScript
End if

If buttonpressed = HRF_family_button then
  run program_path & "BlueZone\bzsh.exe Q:\Blue Zone Scripts\Script Files\NOTE - HRF (Family).vbs"
  StopScript
End if