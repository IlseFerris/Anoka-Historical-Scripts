'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - LTC burial assets"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 01
BeginDialog burial_assets_dialog_01, 0, 0, 286, 125, "Burial assets dialog (01)"
  EditBox 60, 5, 75, 15, case_number
  CheckBox 5, 25, 160, 10, "Applied $1500 of burial services to BFE?", applied_BFE_check
  DropListBox 100, 40, 45, 15, "AFB"+chr(9)+"CSA"+chr(9)+"IBA"+chr(9)+"IFB"+chr(9)+"RBA", type_of_burial_agreement
  EditBox 210, 40, 65, 15, purchase_date
  EditBox 55, 60, 125, 15, issuer_name
  EditBox 215, 60, 55, 15, policy_number
  EditBox 50, 80, 50, 15, face_value
  EditBox 165, 80, 115, 15, funeral_home
  ButtonGroup ButtonPressed
    PushButton 90, 105, 50, 15, "next", next_button
    CancelButton 150, 105, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 45, 90, 10, "Type of burial agreement:"
  Text 155, 45, 55, 10, "Purchase date:"
  Text 5, 65, 50, 10, "Issuer name:"
  Text 185, 65, 30, 10, "Policy #:"
  Text 5, 85, 40, 10, "Face value:"
  Text 110, 85, 50, 10, "Funeral home:"
EndDialog


BeginDialog burial_assets_dialog_02, 0, 0, 305, 380, "Burial Assets Dialog (02)"
  CheckBox 10, 25, 110, 10, "Basic service funeral director", basic_service_funeral_director_check
  Text 50, 5, 30, 10, "SERVICE"
  Text 155, 5, 25, 10, "VALUE"
  Text 240, 5, 30, 10, "STATUS"
  EditBox 140, 20, 55, 15, basic_service_funeral_director_value
  DropListBox 215, 20, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", basic_service_funeral_director_status
  CheckBox 10, 45, 110, 10, "Embalming", embalming_check
  EditBox 140, 40, 55, 15, embalming_value
  DropListBox 215, 40, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", embalming_status
  CheckBox 10, 65, 110, 10, "Other preparation to body", other_preparation_to_body_check
  EditBox 140, 60, 55, 15, other_preparation_to_body_value
  DropListBox 215, 60, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", other_preparation_to_body_status
  CheckBox 10, 85, 110, 10, "Visitation at funeral chapel", visitation_at_funeral_chapel_check
  EditBox 140, 80, 55, 15, visitation_at_funeral_chapel_value
  DropListBox 215, 80, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", visitation_at_funeral_chapel_status
  CheckBox 10, 105, 110, 10, "Visitation at other facility", visitation_at_other_facility_check
  EditBox 140, 100, 55, 15, visitation_at_other_facility_value
  DropListBox 215, 100, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", visitation_at_other_facility_status
  CheckBox 10, 125, 110, 10, "Funeral serv at funeral chapel", funeral_serv_at_funeral_chapel_check
  EditBox 140, 120, 55, 15, funeral_serv_at_funeral_chapel_value
  DropListBox 215, 120, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", funeral_serv_at_funeral_chapel_status
  CheckBox 10, 145, 110, 10, "Funeral serv at other facility", funeral_serv_at_other_facility_check
  EditBox 140, 140, 55, 15, funeral_serv_at_other_facility_value
  DropListBox 215, 140, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", funeral_serv_at_other_facility_status
  CheckBox 10, 165, 110, 10, "Memorial serv at funeral chapel", memorial_serv_at_funeral_chapel_check
  EditBox 140, 160, 55, 15, memorial_serv_at_funeral_chapel_value
  DropListBox 215, 160, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", memorial_serv_at_funeral_chapel_status
  CheckBox 10, 185, 110, 10, "Memorial serv at other facility", memorial_serv_at_other_facility_check
  EditBox 140, 180, 55, 15, memorial_serv_at_other_facility_value
  DropListBox 215, 180, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", memorial_serv_at_other_facility_status
  CheckBox 10, 205, 110, 10, "Graveside service", graveside_service_check
  EditBox 140, 200, 55, 15, graveside_service_value
  DropListBox 215, 200, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", graveside_service_status
  CheckBox 10, 225, 120, 10, "Transfer remains to funeral home", transfer_remains_to_funeral_home_check
  EditBox 140, 220, 55, 15, transfer_remains_to_funeral_home_value
  DropListBox 215, 220, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", transfer_remains_to_funeral_home_status
  CheckBox 10, 245, 110, 10, "Funeral coach", funeral_coach_check
  EditBox 140, 240, 55, 15, funeral_coach_value
  DropListBox 215, 240, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", funeral_coach_status
  CheckBox 10, 265, 110, 10, "Funeral sedan/limousine", funeral_sedan_check
  EditBox 140, 260, 55, 15, funeral_sedan_value
  DropListBox 215, 260, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", funeral_sedan_status
  CheckBox 10, 285, 110, 10, "Service vehicle", service_vehicle_check
  EditBox 140, 280, 55, 15, service_vehicle_value
  DropListBox 215, 280, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", service_vehicle_status
  CheckBox 10, 305, 110, 10, "Forwarding of remains", forwarding_of_remains_check
  EditBox 140, 300, 55, 15, forwarding_of_remains_value
  DropListBox 215, 300, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", forwarding_of_remains_status
  CheckBox 10, 325, 110, 10, "Receiving of remains", receiving_of_remains_check
  EditBox 140, 320, 55, 15, receiving_of_remains_value
  DropListBox 215, 320, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", receiving_of_remains_status
  CheckBox 10, 345, 110, 10, "Direct cremation", direct_cremation_check
  EditBox 140, 340, 55, 15, direct_cremation_value
  DropListBox 215, 340, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", direct_cremation_status
  ButtonGroup ButtonPressed
    PushButton 45, 360, 50, 15, "previous", previous_button
    PushButton 105, 360, 50, 15, "next", next_button
    CancelButton 165, 360, 50, 15
EndDialog

BeginDialog burial_assets_dialog_03, 0, 0, 305, 260, "Burial Assets Dialog (03)"
  Text 30, 5, 100, 10, "BURIAL SPACE/ITEM"
  Text 155, 5, 25, 10, "VALUE"
  Text 240, 5, 30, 10, "STATUS"
  CheckBox 10, 25, 110, 10, "Markers/Headstone", markers_headstone_check
  EditBox 140, 20, 55, 15, markers_headstone_value
  DropListBox 215, 20, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", markers_headstone_status
  CheckBox 10, 45, 110, 10, "Engraving", engraving_check
  EditBox 140, 40, 55, 15, engraving_value
  DropListBox 215, 40, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", engraving_status
  CheckBox 10, 65, 110, 10, "Opening/Closing of space", opening_closing_of_space_check
  EditBox 140, 60, 55, 15, opening_closing_of_space_value
  DropListBox 215, 60, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", opening_closing_of_space_status
  CheckBox 10, 85, 110, 10, "Perpetual Care", perpetual_care_check
  EditBox 140, 80, 55, 15, perpetual_care_value
  DropListBox 215, 80, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", perpetual_care_status
  CheckBox 10, 105, 110, 10, "Casket", casket_check
  EditBox 140, 100, 55, 15, casket_value
  DropListBox 215, 100, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", casket_status
  CheckBox 10, 125, 110, 10, "Vault", vault_check
  EditBox 140, 120, 55, 15, vault_value
  DropListBox 215, 120, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", vault_status
  CheckBox 10, 145, 110, 10, "Cemetery plot", cemetery_plot_check
  EditBox 140, 140, 55, 15, cemetery_plot_value
  DropListBox 215, 140, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", cemetery_plot_status
  CheckBox 10, 165, 110, 10, "Crypt", crypt_check
  EditBox 140, 160, 55, 15, crypt_value
  DropListBox 215, 160, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", crypt_status
  CheckBox 10, 185, 110, 10, "Mausoleum", mausoleum_check
  EditBox 140, 180, 55, 15, mausoleum_value
  DropListBox 215, 180, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", mausoleum_status
  CheckBox 10, 205, 110, 10, "Urns", urns_check
  EditBox 140, 200, 55, 15, urns_value
  DropListBox 215, 200, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", urns_status
  CheckBox 10, 225, 120, 10, "Niches", niches_check
  EditBox 140, 220, 55, 15, niches_value
  DropListBox 215, 220, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", niches_status
  ButtonGroup ButtonPressed
    PushButton 45, 240, 50, 15, "previous", previous_button
    PushButton 105, 240, 50, 15, "next", next_button
    CancelButton 165, 240, 50, 15
EndDialog

BeginDialog burial_assets_dialog_04, 0, 0, 306, 351, "Burial Assets Dialog (04)"
  Text 30, 5, 80, 10, "CASH ADVANCED ITEM"
  Text 155, 5, 25, 10, "VALUE"
  Text 240, 5, 30, 10, "STATUS"
  CheckBox 10, 25, 110, 10, "Certified death certificate", certified_death_certificate_check
  EditBox 140, 20, 55, 15, certified_death_certificate_value
  DropListBox 215, 20, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", certified_death_certificate_status
  CheckBox 10, 45, 110, 10, "Motor escort", motor_escort_check
  EditBox 140, 40, 55, 15, motor_escort_value
  DropListBox 215, 40, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", motor_escort_status
  CheckBox 10, 65, 110, 10, "Clergy honorarium", clergy_honorarium_check
  EditBox 140, 60, 55, 15, clergy_honorarium_value
  DropListBox 215, 60, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", clergy_honorarium_status
  CheckBox 10, 85, 110, 10, "Music honorarium", music_honorarium_check
  EditBox 140, 80, 55, 15, music_honorarium_value
  DropListBox 215, 80, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", music_honorarium_status
  CheckBox 10, 105, 110, 10, "Flowers", flowers_check
  EditBox 140, 100, 55, 15, flowers_value
  DropListBox 215, 100, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", flowers_status
  CheckBox 10, 125, 110, 10, "Obituary notice", obituary_notice_check
  EditBox 140, 120, 55, 15, obituary_notice_value
  DropListBox 215, 120, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", obituary_notice_status
  CheckBox 10, 145, 110, 10, "Crematory charges", crematory_charges_check
  EditBox 140, 140, 55, 15, crematory_charges_value
  DropListBox 215, 140, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", crematory_charges_status
  CheckBox 10, 165, 110, 10, "Acknowledgement card", acknowledgement_card_check
  EditBox 140, 160, 55, 15, acknowledgement_card_value
  DropListBox 215, 160, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", acknowledgement_card_status
  CheckBox 10, 185, 110, 10, "Register book", register_book_check
  EditBox 140, 180, 55, 15, register_book_value
  DropListBox 215, 180, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", register_book_status
  CheckBox 10, 205, 110, 10, "Service folders/prayer cards", service_folders_prayer_cards_check
  EditBox 140, 200, 55, 15, service_folders_prayer_cards_value
  DropListBox 215, 200, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", service_folders_prayer_cards_status
  Text 10, 225, 30, 10, "Other(1):"
  EditBox 45, 220, 85, 15, other_01
  EditBox 140, 220, 55, 15, other_01_value
  DropListBox 215, 220, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", other_01_status
  Text 40, 240, 35, 10, "This is a: "
  DropListBox 80, 240, 45, 15, "service"+chr(9)+"BS-BSI"+chr(9)+"CAI", other_01_type
  Text 10, 260, 30, 10, "Other(2):"
  EditBox 45, 255, 85, 15, other_02
  EditBox 140, 255, 55, 15, other_02_value
  DropListBox 215, 255, 80, 10, "counted"+chr(9)+"excluded"+chr(9)+"unavailable", other_02_status
  Text 40, 275, 35, 10, "This is a: "
  DropListBox 80, 275, 45, 15, "service"+chr(9)+"BS-BSI"+chr(9)+"CAI", other_02_type
  Text 10, 295, 45, 10, "Case action:"
  EditBox 55, 290, 235, 15, case_action
  Text 140, 315, 70, 10, "Sign your case note:"
  EditBox 210, 310, 80, 15, worker_sig
  ButtonGroup ButtonPressed
    PushButton 70, 330, 50, 15, "previous", previous_button
    OkButton 125, 330, 50, 15
    CancelButton 180, 330, 50, 15
EndDialog

BeginDialog cancel_dialog, 0, 0, 126, 61, "Cancel dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 100, 15, "yes, cancel this case note", yes
    PushButton 10, 40, 110, 15, "no, do not cancel this case note", no
  Text 10, 5, 110, 10, "Are you sure you want to cancel?"
EndDialog

'SECTION 02

EMConnect ""

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""


Do
  Do
    Do
      Dialog burial_assets_dialog_01
      If buttonpressed = 0 then 
        Dialog cancel_dialog
        If buttonpressed = yes then stopscript
        If buttonpressed = no then exit do
      End if
      If isnumeric(case_number) = False then MsgBox "You must provide a valid case number."
    Loop until isnumeric(case_number) = True
    Do
      Dialog burial_assets_dialog_02
      If buttonpressed = 0 then 
        Dialog cancel_dialog
        If buttonpressed = yes then stopscript
        If buttonpressed = no then exit do
      End if
      If buttonpressed = previous_button then exit do
      Do
        Dialog burial_assets_dialog_03
        If buttonpressed = previous_button then exit do
        If buttonpressed = 0 then 
          Dialog cancel_dialog
          If buttonpressed = yes then stopscript
          If buttonpressed = no then exit do
        End if
        Do
          Dialog burial_assets_dialog_04
          If buttonpressed = previous_button then exit do
          If buttonpressed = 0 then 
            Dialog cancel_dialog
            If buttonpressed = yes then stopscript
            If buttonpressed = no then exit do
          End if
          transmit
          EMReadScreen MAXIS_check, 5, 1, 39
          If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You don't appear to be in MAXIS. You might be locked out of your case. Please get back into MAXIS production before continuing."
        Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
      Loop until buttonpressed = -1
    Loop until buttonpressed = -1
  Loop until buttonpressed = -1
Loop until buttonpressed = -1


call navigate_to_screen ("case", "note")
PF9


'SECTION 03
'Must convert non-numeric "values" to numeric for calculations to work
If isnumeric(basic_service_funeral_director_value) = False then basic_service_funeral_director_value = 0
If isnumeric(embalming_value) = False then embalming_value = 0
If isnumeric(other_preparation_to_body_value) = False then other_preparation_to_body_value = 0
If isnumeric(visitation_at_funeral_chapel_value) = False then visitation_at_funeral_chapel_value = 0
If isnumeric(visitation_at_other_facility_value) = False then visitation_at_other_facility_value = 0
If isnumeric(funeral_serv_at_funeral_chapel_value) = False then funeral_serv_at_funeral_chapel_value = 0
If isnumeric(funeral_serv_at_other_facility_value) = False then funeral_serv_at_other_facility_value = 0
If isnumeric(memorial_serv_at_funeral_chapel_value) = False then memorial_serv_at_funeral_chapel_value = 0
If isnumeric(memorial_serv_at_other_facility_value) = False then memorial_serv_at_other_facility_value = 0
If isnumeric(graveside_service_value) = False then graveside_service_value = 0
If isnumeric(transfer_remains_to_funeral_home_value) = False then transfer_remains_to_funeral_home_value = 0
If isnumeric(funeral_coach_value) = False then funeral_coach_value = 0
If isnumeric(funeral_sedan_value) = False then funeral_sedan_value = 0
If isnumeric(service_vehicle_value) = False then service_vehicle_value = 0
If isnumeric(forwarding_of_remains_value) = False then forwarding_of_remains_value = 0
If isnumeric(receiving_of_remains_value) = False then receiving_of_remains_value = 0
If isnumeric(direct_cremation_value) = False then direct_cremation_value = 0
If isnumeric(markers_headstone_value) = False then markers_headstone_value = 0
If isnumeric(engraving_value) = False then engraving_value = 0
If isnumeric(opening_closing_of_space_value) = False then opening_closing_of_space_value = 0
If isnumeric(perpetual_care_value) = False then perpetual_care_value = 0
If isnumeric(casket_value) = False then casket_value = 0
If isnumeric(vault_value) = False then vault_value = 0
If isnumeric(cemetery_plot_value) = False then cemetery_plot_value = 0
If isnumeric(crypt_value) = False then crypt_value = 0
If isnumeric(mausoleum_value) = False then mausoleum_value = 0
If isnumeric(urns_value) = False then urns_value = 0
If isnumeric(niches_value) = False then niches_value = 0
If isnumeric(certified_death_certificate_value) = False then certified_death_certificate_value = 0
If isnumeric(motor_escort_value) = False then motor_escort_value = 0
If isnumeric(clergy_honorarium_value) = False then clergy_honorarium_value = 0
If isnumeric(music_honorarium_value) = False then music_honorarium_value = 0
If isnumeric(flowers_value) = False then flowers_value = 0
If isnumeric(obituary_notice_value) = False then obituary_notice_value = 0
If isnumeric(crematory_charges_value) = False then crematory_charges_value = 0
If isnumeric(acknowledgement_card_value) = False then acknowledgement_card_value = 0
If isnumeric(register_book_value) = False then register_book_value = 0
If isnumeric(service_folders_prayer_cards_value) = False then service_folders_prayer_cards_value = 0
If isnumeric(other_01_value) = False then other_01_value = 0
If isnumeric(other_02_value) = False then other_02_value = 0

'This adds all service amounts together.
total_service_amount = cint(basic_service_funeral_director_value) + cint(embalming_value) + cint(other_preparation_to_body_value) + cint(visitation_at_funeral_chapel_value) + cint(visitation_at_other_facility_value) + cint(funeral_serv_at_funeral_chapel_value) + cint(funeral_serv_at_other_facility_value) + cint(memorial_serv_at_funeral_chapel_value) + cint(memorial_serv_at_other_facility_value) + cint(graveside_service_value) + cint(transfer_remains_to_funeral_home_value) + cint(funeral_coach_value) + cint(funeral_sedan_value) + cint(service_vehicle_value) + cint(forwarding_of_remains_value) + cint(receiving_of_remains_value) + cint(direct_cremation_value)
If other_01 <> "" and other_01_type = "service" then total_service_amount = total_service_amount + cint(other_01_value)
If other_02 <> "" and other_02_type = "service" then total_service_amount = total_service_amount + cint(other_02_value)

'This adds all exluded burial space/burial space items (BS/BSI) together.
If markers_headstone_check = 1 and markers_headstone_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(markers_headstone_value)
If engraving_check = 1 and engraving_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(engraving_value)
If opening_closing_of_space_check = 1 and opening_closing_of_space_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(opening_closing_of_space_value)
If perpetual_care_check = 1 and perpetual_care_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(perpetual_care_value)
If casket_check = 1 and casket_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(casket_value)
If vault_check = 1 and vault_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(vault_value)
If cemetery_plot_check = 1 and cemetery_plot_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(cemetery_plot_value)
If crypt_check = 1 and crypt_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(crypt_value)
If mausoleum_check = 1 and mausoleum_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(mausoleum_value)
If urns_check = 1 and urns_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(urns_value)
If niches_check = 1 and niches_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(niches_value)
If other_01 <> "" and other_01_type = "BS-BSI" and other_01_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(other_01_value)
If other_02 <> "" and other_02_type = "BS-BSI" and other_02_status = "excluded" then total_BS_BSI_excluded_amount = total_BS_BSI_excluded_amount + cint(other_02_value)

'This adds all unavailable cash advance items (CAI) together.
If certified_death_certificate_check = 1 and certified_death_certificate_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(certified_death_certificate_value)
If motor_escort_check = 1 and motor_escort_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(motor_escort_value)
If clergy_honorarium_check = 1 and clergy_honorarium_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(clergy_honorarium_value)
If music_honorarium_check = 1 and music_honorarium_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(music_honorarium_value)
If flowers_check = 1 and flowers_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(flowers_value)
If obituary_notice_check = 1 and obituary_notice_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(obituary_notice_value)
If crematory_charges_check = 1 and crematory_charges_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(crematory_charges_value)
If acknowledgement_card_check = 1 and acknowledgement_card_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(acknowledgement_card_value)
If register_book_check = 1 and register_book_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(register_book_value)
If service_folders_prayer_cards_check = 1 and service_folders_prayer_cards_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(service_folders_prayer_cards_value)
If other_01 <> "" and other_01_type = "CAI" and other_01_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(other_01_value)
If other_02 <> "" and other_02_type = "CAI" and other_02_status = "unavailable" then total_unavailable_CAI_amount = total_unavailable_CAI_amount + cint(other_02_value)


'This adds all counted fields together.
If basic_service_funeral_director_check = 1 and basic_service_funeral_director_status = "counted" then total_counted_amount = total_counted_amount + cint(basic_service_funeral_director_value)
If embalming_check = 1 and embalming_status = "counted" then total_counted_amount = total_counted_amount + cint(embalming_value)
If other_preparation_to_body_check = 1 and other_preparation_to_body_status = "counted" then total_counted_amount = total_counted_amount + cint(other_preparation_to_body_value)
If visitation_at_funeral_chapel_check = 1 and visitation_at_funeral_chapel_status = "counted" then total_counted_amount = total_counted_amount + cint(visitation_at_funeral_chapel_value)
If visitation_at_other_facility_check = 1 and visitation_at_other_facility_status = "counted" then total_counted_amount = total_counted_amount + cint(visitation_at_other_facility_value)
If funeral_serv_at_funeral_chapel_check = 1 and funeral_serv_at_funeral_chapel_status = "counted" then total_counted_amount = total_counted_amount + cint(funeral_serv_at_funeral_chapel_value)
If funeral_serv_at_other_facility_check = 1 and funeral_serv_at_other_facility_status = "counted" then total_counted_amount = total_counted_amount + cint(funeral_serv_at_other_facility_value)
If memorial_serv_at_funeral_chapel_check = 1 and memorial_serv_at_funeral_chapel_status = "counted" then total_counted_amount = total_counted_amount + cint(memorial_serv_at_funeral_chapel_value)
If memorial_serv_at_other_facility_check = 1 and memorial_serv_at_other_facility_status = "counted" then total_counted_amount = total_counted_amount + cint(memorial_serv_at_other_facility_value)
If graveside_service_check = 1 and graveside_service_status = "counted" then total_counted_amount = total_counted_amount + cint(graveside_service_value)
If transfer_remains_to_funeral_home_check = 1 and transfer_remains_to_funeral_home_status = "counted" then total_counted_amount = total_counted_amount + cint(transfer_remains_to_funeral_home_value)
If funeral_coach_check = 1 and funeral_coach_status = "counted" then total_counted_amount = total_counted_amount + cint(funeral_coach_value)
If funeral_sedan_check = 1 and funeral_sedan_status = "counted" then total_counted_amount = total_counted_amount + cint(funeral_sedan_value)
If service_vehicle_check = 1 and service_vehicle_status = "counted" then total_counted_amount = total_counted_amount + cint(service_vehicle_value)
If forwarding_of_remains_check = 1 and forwarding_of_remains_status = "counted" then total_counted_amount = total_counted_amount + cint(forwarding_of_remains_value)
If receiving_of_remains_check = 1 and receiving_of_remains_status = "counted" then total_counted_amount = total_counted_amount + cint(receiving_of_remains_value)
If direct_cremation_check = 1 and direct_cremation_status = "counted" then total_counted_amount = total_counted_amount + cint(direct_cremation_value)
If markers_headstone_check = 1 and markers_headstone_status = "counted" then total_counted_amount = total_counted_amount + cint(markers_headstone_value)
If engraving_check = 1 and engraving_status = "counted" then total_counted_amount = total_counted_amount + cint(engraving_value)
If opening_closing_of_space_check = 1 and opening_closing_of_space_status = "counted" then total_counted_amount = total_counted_amount + cint(opening_closing_of_space_value)
If perpetual_care_check = 1 and perpetual_care_status = "counted" then total_counted_amount = total_counted_amount + cint(perpetual_care_value)
If casket_check = 1 and casket_status = "counted" then total_counted_amount = total_counted_amount + cint(casket_value)
If vault_check = 1 and vault_status = "counted" then total_counted_amount = total_counted_amount + cint(vault_value)
If cemetery_plot_check = 1 and cemetery_plot_status = "counted" then total_counted_amount = total_counted_amount + cint(cemetery_plot_value)
If crypt_check = 1 and crypt_status = "counted" then total_counted_amount = total_counted_amount + cint(crypt_value)
If mausoleum_check = 1 and mausoleum_status = "counted" then total_counted_amount = total_counted_amount + cint(mausoleum_value)
If urns_check = 1 and urns_status = "counted" then total_counted_amount = total_counted_amount + cint(urns_value)
If niches_check = 1 and niches_status = "counted" then total_counted_amount = total_counted_amount + cint(niches_value)
If certified_death_certificate_check = 1 and certified_death_certificate_status = "counted" then total_counted_amount = total_counted_amount + cint(certified_death_certificate_value)
If motor_escort_check = 1 and motor_escort_status = "counted" then total_counted_amount = total_counted_amount + cint(motor_escort_value)
If clergy_honorarium_check = 1 and clergy_honorarium_status = "counted" then total_counted_amount = total_counted_amount + cint(clergy_honorarium_value)
If music_honorarium_check = 1 and music_honorarium_status = "counted" then total_counted_amount = total_counted_amount + cint(music_honorarium_value)
If flowers_check = 1 and flowers_status = "counted" then total_counted_amount = total_counted_amount + cint(flowers_value)
If obituary_notice_check = 1 and obituary_notice_status = "counted" then total_counted_amount = total_counted_amount + cint(obituary_notice_value)
If crematory_charges_check = 1 and crematory_charges_status = "counted" then total_counted_amount = total_counted_amount + cint(crematory_charges_value)
If acknowledgement_card_check = 1 and acknowledgement_card_status = "counted" then total_counted_amount = total_counted_amount + cint(acknowledgement_card_value)
If register_book_check = 1 and register_book_status = "counted" then total_counted_amount = total_counted_amount + cint(register_book_value)
If service_folders_prayer_cards_check = 1 and service_folders_prayer_cards_status = "counted" then total_counted_amount = total_counted_amount + cint(service_folders_prayer_cards_value)
If other_01 <> "" and other_01_status = "counted" then total_counted_amount = total_counted_amount + cint(other_01_value)
If other_02 <> "" and other_02_status = "counted" then total_counted_amount = total_counted_amount + cint(other_02_value)

If total_service_amount = "" then total_service_amount = "0"
If total_BS_BSI_excluded_amount = "" then total_BS_BSI_excluded_amount = "0"
If total_unavailable_CAI_amount = "" then total_unavailable_CAI_amount = "0"
If total_counted_amount = "" then total_counted_amount = "0"




'SECTION 04

dim MAXIS_service_row
dim MAXIS_col


'NOTE: "Other" sections need to be included in correct sections. 
EMSendKey "**BURIAL ASSETS**" + "<newline>"
If applied_BFE_check = 1 then EMSendKey "* Applied $1500 of burial services to BFE." & "<newline>"
EMSendKey "* Type: " & type_of_burial_agreement & ". Purchase date: " & purchase_date & "." & "<newline>"
EMSendKey "* Issuer: " & issuer_name & ". Policy #: " & policy_number & "." & "<newline>"
EMSendKey "* Face value: " & face_value & "<newline>"
EMSendKey "* Funeral home: " & funeral_home & "<newline>"
EMSendKey "--------------SERVICE--------------------AMOUNT----------STATUS--------------" & "<newline>"
new_page_check
If basic_service_funeral_director_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "     Basic service funeral director:", 44, "$" & basic_service_funeral_director_value, 59, basic_service_funeral_director_status)
End if
new_page_check
If embalming_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "                          Embalming:", 44, "$" & embalming_value, 59, embalming_status)
End if
new_page_check
If other_preparation_to_body_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "          Other preparation to body:", 44, "$" & other_preparation_to_body_value, 59, other_preparation_to_body_status)
End if
new_page_check
If visitation_at_funeral_chapel_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "       Visitation at funeral chapel:", 44, "$" & visitation_at_funeral_chapel_value, 59, visitation_at_funeral_chapel_status)
End if
new_page_check
If visitation_at_other_facility_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "       Visitation at other facility:", 44, "$" & visitation_at_other_facility_value, 59, visitation_at_other_facility_status)
End if
new_page_check
If funeral_serv_at_funeral_chapel_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "     Funeral serv at funeral chapel:", 44, "$" & funeral_serv_at_funeral_chapel_value, 59, funeral_serv_at_funeral_chapel_status)
End if
new_page_check
If funeral_serv_at_other_facility_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "     Funeral serv at other facility:", 44, "$" & funeral_serv_at_other_facility_value, 59, funeral_serv_at_other_facility_status)
End if
new_page_check
If memorial_serv_at_funeral_chapel_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "    Memorial serv at funeral chapel:", 44, "$" & memorial_serv_at_funeral_chapel_value, 59, memorial_serv_at_funeral_chapel_status)
End if
new_page_check
If memorial_serv_at_other_facility_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "    Memorial serv at other facility:", 44, "$" & memorial_serv_at_other_facility_value, 59, memorial_serv_at_other_facility_status)
End if
new_page_check
If graveside_service_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "                  Graveside service:", 44, "$" & graveside_service_value, 59, graveside_service_status)
End if
new_page_check
If transfer_remains_to_funeral_home_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "   Transfer remains to funeral home:", 44, "$" & transfer_remains_to_funeral_home_value, 59, transfer_remains_to_funeral_home_status)
End if
new_page_check
If funeral_coach_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "                      Funeral coach:", 44, "$" & funeral_coach_value, 59, funeral_coach_status)
End if
new_page_check
If funeral_sedan_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "                      Funeral sedan:", 44, "$" & funeral_sedan_value, 59, funeral_sedan_status)
End if
new_page_check
If service_vehicle_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "                    Service vehicle:", 44, "$" & service_vehicle_value, 59, service_vehicle_status)
End if
new_page_check
If forwarding_of_remains_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "              Forwarding of remains:", 44, "$" & forwarding_of_remains_value, 59, forwarding_of_remains_status)
End if
new_page_check
If receiving_of_remains_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "               Receiving of remains:", 44, "$" & receiving_of_remains_value, 59, receiving_of_remains_status)
End if
new_page_check
If direct_cremation_check = 1 then 
  new_service_heading
  call write_three_columns_in_case_note(3, "                   Direct cremation:", 44, "$" & direct_cremation_value, 59, direct_cremation_status)
End if
new_page_check
If other_01 <> "" and other_01_type = "service" then 
  new_service_heading
  call write_three_columns_in_case_note(38 - len(other_01), other_01 & ":", 44, "$" & other_01_value, 59, other_01_status)
End if
new_page_check
If other_02 <> "" and other_02_type = "service" then 
  new_service_heading
  call write_three_columns_in_case_note(38 - len(other_02), other_02 & ":", 44, "$" & other_02_value, 59, other_02_status)
End if
new_page_check
EMSendKey "--------BURIAL SPACE/ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
new_page_check
If markers_headstone_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                  Markers headstone:", 44, "$" & markers_headstone_value, 59, markers_headstone_status)
End if
new_page_check
If engraving_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                          Engraving:", 44, "$" & engraving_value, 59, engraving_status)
End if
new_page_check
If opening_closing_of_space_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "           Opening closing of space:", 44, "$" & opening_closing_of_space_value, 59, opening_closing_of_space_status)
End if
new_page_check
If perpetual_care_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                     Perpetual care:", 44, "$" & perpetual_care_value, 59, perpetual_care_status)
End if
new_page_check
If casket_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                             Casket:", 44, "$" & casket_value, 59, casket_status)
End if
new_page_check
If vault_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                              Vault:", 44, "$" & vault_value, 59, vault_status)
End if
new_page_check
If cemetery_plot_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                      Cemetery plot:", 44, "$" & cemetery_plot_value, 59, cemetery_plot_status)
End if
new_page_check
If crypt_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                              Crypt:", 44, "$" & crypt_value, 59, crypt_status)
End if
new_page_check
If mausoleum_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                          Mausoleum:", 44, "$" & mausoleum_value, 59, mausoleum_status)
End if
new_page_check
If urns_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                               Urns:", 44, "$" & urns_value, 59, urns_status)
End if
new_page_check
If niches_check = 1 then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(3, "                             Niches:", 44, "$" & niches_value, 59, niches_status)
End if
new_page_check
If other_01 <> "" and other_01_type = "BS-BSI" then 
  new_BS_BSI_heading
  call write_three_columns_in_case_note(38 - len(other_01), other_01 & ":", 44, "$" & other_01_value, 59, other_01_status)
End if
new_page_check
If other_02 <> "" and other_02_type = "BS-BSI" then 
  new_BS_BSI_heading
  other_02_length = len(other_02)
  call write_three_columns_in_case_note(38 - len(other_02), other_02 & ":", 44, "$" & other_02_value, 59, other_02_status)
End if
new_page_check
EMSendKey "--------CASH ADVANCE ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
new_page_check
If certified_death_certificate_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "        Certified death certificate:", 44, "$" & certified_death_certificate_value, 59, certified_death_certificate_status)
End if
new_page_check
If motor_escort_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                       Motor escort:", 44, "$" & motor_escort_value, 59, motor_escort_status)
End if
new_page_check
If clergy_honorarium_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                  Clergy honorarium:", 44, "$" & clergy_honorarium_value, 59, clergy_honorarium_status)
End if
new_page_check
If music_honorarium_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                   Music honorarium:", 44, "$" & music_honorarium_value, 59, music_honorarium_status)
End if
new_page_check
If flowers_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                            Flowers:", 44, "$" & flowers_value, 59, flowers_status)
End if
new_page_check
If obituary_notice_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                    Obituary notice:", 44, "$" & obituary_notice_value, 59, obituary_notice_status)
End if
new_page_check
If crematory_charges_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                  Crematory charges:", 44, "$" & crematory_charges_value, 59, crematory_charges_status)
End if
new_page_check
If acknowledgement_card_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "               Acknowledgement card:", 44, "$" & acknowledgement_card_value, 59, acknowledgement_card_status)
End if
new_page_check
If register_book_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "                      Register book:", 44, "$" & register_book_value, 59, register_book_status)
End if
new_page_check
If service_folders_prayer_cards_check = 1 then 
  new_CAI_heading
  call write_three_columns_in_case_note(3, "       Service folders prayer cards:", 44, "$" & service_folders_prayer_cards_value, 59, service_folders_prayer_cards_status)
End if
new_page_check
If other_01 <> "" and other_01_type = "CAI" then 
  new_CAI_heading
  call write_three_columns_in_case_note(38 - len(other_01), other_01 & ":", 44, "$" & other_01_value, 59, other_01_status)
End if
new_page_check
If other_02 <> "" and other_02_type = "CAI" then 
  new_CAI_heading
  call write_three_columns_in_case_note(38 - len(other_02), other_02 & ":", 44, "$" & other_02_value, 59, other_02_status)
End if
new_page_check
EMSendKey "-----------------------------------------------------------------------------" & "<newline>"
new_page_check
EMSendKey "* Total service amount: $" & total_service_amount & "<newline>"
new_page_check
EMSendKey "* Total BS/BSI excluded amount: $" & total_BS_BSI_excluded_amount & "<newline>"
new_page_check
EMSendKey "* Total unavailable CAI: $" & total_unavailable_CAI_amount & "<newline>"
new_page_check
EMSendKey "* Total counted amount: $" & total_counted_amount & "<newline>"
new_page_check
EMSendKey "* Actions taken: " & case_action & "<newline>"
new_page_check
EMSendKey "---" & "<newline>"
new_page_check
EMSendKey worker_sig & "<newline>"

script_end_procedure("")