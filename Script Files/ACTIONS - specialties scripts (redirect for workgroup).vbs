'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Script dialog
BeginDialog specialties_scripts_dialog, 0, 0, 216, 185, "Specialties scripts"
  ButtonGroup ButtonPressed
    PushButton 40, 15, 30, 10, "SAVE", SAVE_button
    PushButton 30, 25, 50, 10, "Sponsor inc.", Sponsor_inc_button
    PushButton 10, 55, 90, 10, "ABAWD Screening Tool", ABAWD_Screening_Tool_button
    PushButton 10, 85, 90, 10, "FSET non-comp TIKLer", FSET_non_comp_TIKLer_button
    PushButton 10, 95, 90, 10, "New worker MEMO", new_worker_MEMO_button
    PushButton 10, 125, 90, 10, "Client Contact (OSA)", client_contact_button
    PushButton 10, 155, 90, 10, "Monthly Agency Stats", monthly_agency_stats_button
    PushButton 145, 15, 30, 10, "1503", LTC_1503_button
    PushButton 115, 25, 90, 10, "Asset Assessment", asset_assessment_button
    PushButton 115, 35, 90, 10, "Asset transfer memo", asset_transfer_memo_button
    PushButton 130, 45, 60, 10, "BILS updater", BILS_updater_button
    PushButton 130, 55, 60, 10, "Burial Assets", burial_assets_button
    PushButton 120, 65, 80, 10, "COLA summary", COLA_summary_button
    PushButton 145, 75, 30, 10, "LTC ER", LTC_ER_button
    PushButton 125, 85, 70, 10, "LTC/GRH list gen", LTC_GRH_list_gen_button
    PushButton 130, 95, 60, 10, "LTC intake", LTC_intake_button
    PushButton 125, 105, 70, 10, "LTC intake approval", LTC_intake_approval_button
    PushButton 115, 115, 90, 10, "LTC verifs needed", LTC_verifs_needed_button
    PushButton 130, 125, 60, 10, "MA approval", MA_approval_button
    PushButton 130, 135, 60, 10, "MA-EPD EI FIAT", MA_EPD_EI_FIAT_button
    PushButton 130, 145, 60, 10, "Spousal Alloc.", spousal_alloc_button
    CancelButton 135, 165, 50, 15
  GroupBox 5, 5, 100, 35, "LEP"
  GroupBox 5, 45, 100, 25, "Mentors"
  GroupBox 5, 75, 100, 35, "Supervisors"
  GroupBox 5, 115, 100, 25, "OSAs"
  GroupBox 5, 145, 100, 25, "Statistics"
  GroupBox 110, 5, 100, 155, "LTC"
EndDialog



'Shows dialog, cancels script if requested
Dialog specialties_scripts_dialog
If buttonpressed = 0 then stopscript

'LEP SCRIPTS----------------------------------------------------------------------------------------------------
If buttonpressed = SAVE_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - SAVE (LEP).vbs")
  StopScript
End if

If buttonpressed = Sponsor_inc_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - sponsor income.vbs")
  StopScript
End if

'MENTOR SCRIPTS----------------------------------------------------------------------------------------------------
If buttonpressed = ABAWD_Screening_Tool_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\BETA - ACTIONS - ABAWD Screening Tool.vbs")
  StopScript
End if

'SUPERVISOR SCRIPTS----------------------------------------------------------------------------------------------------
If buttonpressed = FSET_non_comp_TIKLer_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\BULK - FSET non-compliance TIKLer.vbs")
  StopScript
End if

If buttonpressed = new_worker_MEMO_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\MEMO - new worker MEMO.vbs")
  StopScript
End if

'OSA SCRIPTS----------------------------------------------------------------------------------------------------
If buttonpressed = client_contact_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - client contact (call center).vbs")
  StopScript
End if

'STATISTICS SCRIPTS----------------------------------------------------------------------------------------------------
If buttonpressed = monthly_agency_stats_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\STATISTICS - monthly agency stats.vbs")
  StopScript
End if

'LTC/GRH SCRIPTS----------------------------------------------------------------------------------------------------
If buttonpressed = COLA_summary_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\BETA - NOTE - COLA summary.vbs")
  StopScript
End if

If buttonpressed = LTC_1503_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - 1503.vbs")
  StopScript
End if

If buttonpressed = asset_assessment_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - asset assessment.vbs")
  StopScript
End if

If buttonpressed = burial_assets_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - LTC burial assets.vbs")
  StopScript
End if

If buttonpressed = LTC_ER_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - LTC ER.vbs")
  StopScript
End if

If buttonpressed = LTC_intake_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - LTC intake.vbs")
  StopScript
End if

If buttonpressed = LTC_intake_approval_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - LTC intake approval.vbs")
  StopScript
End if

If buttonpressed = MA_approval_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - MA approval.vbs")
  StopScript
End if

If buttonpressed = LTC_verifs_needed_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\NOTE - verifs needed (LTC).vbs")
  StopScript
End if

If buttonpressed = BILS_updater_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\ACTIONS - BILS updater.vbs")
  StopScript
End if

If buttonpressed = MA_EPD_EI_FIAT_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\ACTIONS - MA-EPD EI FIAT.vbs")
  StopScript
End if

If buttonpressed = LTC_GRH_list_gen_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\ACTIONS - LTC-GRH list gen.vbs")
  StopScript
End if

If buttonpressed = spousal_alloc_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\ACTIONS - Spousal allocation FIATer.vbs")
  StopScript
End if

If buttonpressed = asset_transfer_memo_button then
  call run_another_script("Q:\Blue Zone Scripts\Script Files\MEMO - LTC asset transfer.vbs")
  StopScript
End if

