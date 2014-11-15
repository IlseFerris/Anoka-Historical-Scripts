'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - TPQY response"
start_time = timer

''LOADING ROUTINE FUNCTIONS
'<<DELETE REDUNDANCIES!
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.



'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog SVES_dialog, 0, 0, 126, 72, "QURY reader dialog"
  Text 5, 5, 100, 10, "Would you like the script to:"
  OptionGroup RadioGroup1
    RadioButton 5, 20, 110, 10, "Read the info and case note it.", Radio2
    RadioButton 5, 35, 75, 10, "Stop here.", Radio1
  ButtonGroup SVES_dialog_ButtonPressed
    OkButton 10, 50, 50, 15
    CancelButton 65, 50, 50, 15
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Reads case number
EMReadScreen case_number, 8, 5, 73

'Navigates to INFC
EMSendKey "i"
transmit

'Navigates to SVES
EMWriteScreen "sves", 20, 71
transmit

'Navigates to TPQY
EMWriteScreen "tpqy", 20, 70
transmit

'Per management decision 12/11/2013, this is where the script stops.  We'll restore the function as soon as clarification comes from DHS.
MsgBox "At this time, this script can no longer case note TPQY information. When we get clarification on new policy, we will update the script and add this feature back. In the meantime, please read the Policy Point issued 12/11/2013 called ''IRA and SSA Updated Safeguarding Requirements''. Thank you!"
StopScript

'Shows the dialog, if cancel or "stop here" is selected the script ends
Dialog SVES_dialog
If SVES_dialog_ButtonPressed = 0 or radiogroup1 = 1 then stopscript

'Reads client name, SSN, and response date
EMReadScreen client_name, 25, 4, 10
EMReadScreen client_SSN, 11, 5, 9
client_SSN = replace(client_SSN, " ", "-") 'This replaces the spaces in the client_SSN variable with dashes.
EMReadScreen response_date, 8, 7, 22
response_date = replace(response_date, " ", "/") 'This converts the response date to a date-based variable, with slashes.
transmit

'Checks to make sure we've moved past BDXP. If not, the script will close. This is errorproofing.
EMReadScreen BDXP_check, 4, 2, 53
If BDXP_check <> "BDXP" then
  MsgBox "Error! Your MAXIS screen could not get to BDXP. Navigate back to the DAIL and try again."
  StopScript
End if

'Reads all of BDXP, and grabs claim numbers.
EMReadScreen line_01_BDXP, 76, 4, 3
EMReadScreen line_02_BDXP, 76, 5, 3
EMReadScreen line_03_BDXP, 76, 6, 3
EMReadScreen line_04_BDXP, 76, 7, 3
EMReadScreen line_05_BDXP, 76, 8, 3
EMReadScreen line_06_BDXP, 76, 9, 3
EMReadScreen line_07_BDXP, 76, 10, 3
EMReadScreen line_08_BDXP, 76, 11, 3
EMReadScreen line_09_BDXP, 76, 12, 3
EMReadScreen line_10_BDXP, 76, 13, 3
EMReadScreen line_11_BDXP, 76, 14, 3
EMReadScreen RSDI_claim_number, 12, 5, 40
If RSDI_claim_number = "            " then RSDI_claim_number = "none shown" 'This replaces the claim number if one isn't shown.
If right(RSDI_claim_number, 2) = "00" then RSDI_claim_number = left(RSDI_claim_number, 10) 'Removes "00" from the end of an RSDI claim number per policy
EMReadScreen RSDI_dual_entl_number, 12, 5, 69
If RSDI_dual_entl_number = "            " then RSDI_dual_entl_number = "none shown" 'This replaces the dual entitlement number if one isn't shown.
EMReadScreen disa_date, 10, 15, 69
If disa_date <> "          " then disa_date = replace(disa_date, " ", "/") 'This converts the disa date to a date-based variable, with slashes.
If disa_date = "          " then disa_date = "none shown"
EMReadScreen most_recent_pay_date, 5, 8, 5
If most_recent_pay_date <> "     " then most_recent_pay_date = replace(most_recent_pay_date, " ", "/03/") 'converts the most recent pay date to a MAXIS friendly date.
If most_recent_pay_date = "     " then most_recent_pay_date = "none shown"
EMReadScreen second_most_recent_pay_date, 5, 9, 5
If second_most_recent_pay_date <> "     " then second_most_recent_pay_date = replace(second_most_recent_pay_date, " ", "/03/") 'converts the second most recent pay date to a MAXIS friendly date.
If second_most_recent_pay_date = "     " then second_most_recent_pay_date = "none shown"
EMReadScreen most_recent_monthly_amount, 7, 8, 16
EMReadScreen second_most_recent_monthly_amount, 7, 9, 16
EMReadScreen net_amount, 7, 8, 32
EMReadScreen RSDI_claim_number, 12, 5, 40                                               '|These four sections are from the add-feature version of this
EMReadScreen RSDI_gross, 7, 8, 16                                                       '|script. They're added in here so that the logic will
If isnumeric(trim(RSDI_gross)) = True then RSDI_gross = abs(trim(RSDI_gross))           '|eventually work smoothly for testing.
EMReadScreen BDXM_SSN, 11, 5, 19                                                        '|
transmit

'Reads Medicare info and all of BDXM.
EMReadScreen medicare_number, 12, 4, 29
If medicare_number = "            " then 
  medicare_number = "none shown"
Else
  EMReadScreen medicare_suffix, 3, 4, 41 
  medicare_suffix = replace(medicare_suffix, "0", "")
  medicare_number = medicare_number & medicare_suffix
End if
EMReadScreen part_a_premium, 6, 6, 64
If part_a_premium = "      " then part_a_premium = "none shown"
EMReadScreen part_a_start_date, 5, 7, 25
If part_a_start_date <> "     " then part_a_start_date = replace(part_a_start_date, " ", "/")
If part_a_start_date = "     " then part_a_start_date = "none shown"
EMReadScreen part_b_premium, 6, 12, 64
If part_b_premium = "      " then part_b_premium = "none shown"
EMReadScreen part_b_start_date, 5, 13, 25
If part_b_start_date <> "     " then part_b_start_date = replace(part_b_start_date, " ", "/")
If part_b_start_date = "     " then part_b_start_date = "none shown"
EMReadScreen line_01_BDXM, 76, 4, 2
EMReadScreen line_02_BDXM, 76, 5, 2
EMReadScreen line_03_BDXM, 76, 6, 2
EMReadScreen line_04_BDXM, 76, 7, 2
EMReadScreen line_05_BDXM, 76, 8, 2
EMReadScreen line_06_BDXM, 76, 9, 2
EMReadScreen line_07_BDXM, 76, 10, 2
EMReadScreen line_08_BDXM, 76, 11, 2
EMReadScreen line_09_BDXM, 76, 12, 2
EMReadScreen line_10_BDXM, 76, 13, 2
EMReadScreen line_11_BDXM, 76, 14, 2
EMReadScreen line_12_BDXM, 76, 15, 2
EMReadScreen line_13_BDXM, 76, 16, 2
EMReadScreen line_14_BDXM, 76, 17, 2
transmit

'Reads the federal living arrangement from SDXE. It also looks for current pay. It uses this later in the script to determine if SSI is current or not.
EMReadScreen fed_liv_arrange, 1, 6, 70
If fed_liv_arrange = " " then fed_liv_arrange = "none shown"
  SDXE_row = 1
  SDXE_col = 1
EMSearch "CURRENT PAY", SDXE_row, SDXE_col
transmit

'Reads all of the info off of SDXP.
EMReadScreen line_01_SDXP, 76, 4, 2
EMReadScreen line_02_SDXP, 76, 5, 2
EMReadScreen line_03_SDXP, 76, 6, 2
EMReadScreen line_04_SDXP, 76, 7, 2
EMReadScreen line_05_SDXP, 76, 8, 2
EMReadScreen line_06_SDXP, 76, 9, 2
EMReadScreen line_07_SDXP, 76, 10, 2
EMReadScreen line_08_SDXP, 76, 11, 2
EMReadScreen line_09_SDXP, 76, 12, 2
EMReadScreen line_10_SDXP, 76, 13, 2
EMReadScreen line_11_SDXP, 76, 14, 2
EMReadScreen line_12_SDXP, 76, 15, 2
EMReadScreen line_13_SDXP, 76, 16, 2
EMReadScreen line_14_SDXP, 76, 17, 2
EMReadScreen most_recent_SSI_pay_date, 5, 8, 3
If most_recent_SSI_pay_date <> "     " then most_recent_SSI_pay_date = replace(most_recent_SSI_pay_date, " ", "/01/") 'converts the most recent pay date to a MAXIS friendly date.
If most_recent_SSI_pay_date = "     " then most_recent_SSI_pay_date = "none shown"
EMReadScreen second_most_recent_SSI_pay_date, 6, 8, 3
If second_most_recent_SSI_pay_date <> "     " then second_most_recent_SSI_pay_date = replace(second_most_recent_SSI_pay_date, " ", "/01/") 'converts the most recent pay date to a MAXIS friendly date.
If second_most_recent_SSI_pay_date = "     " then second_most_recent_SSI_pay_date = "none shown"
EMReadScreen most_recent_monthly_SSI_amount, 6, 8, 13
EMReadScreen second_most_recent_monthly_SSI_amount, 6, 9, 13
transmit

'Navigates back to SELF
do
  PF3
  EMReadScreen SELF_check, 4, 2, 50
loop until SELF_check = "SELF"

'Navigates to a case note.
EMWriteScreen "case", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "note", 21, 70
transmit

'Now it will case note the info.
PF9
EMSendKey ":::SVES RESPONSE: " + response_date + ":::" + "<newline>"
EMSendKey "* Name/SSN: " + client_name
col = 15
EMSearch "  ", 5, col 'This finds the edge of the client's name
EMWriteScreen "/" + client_SSN, 5, col
EMSendKey "<newline>"
EMSendKey "* RSDI claim number"
If RSDI_dual_entl_number <> "none shown" then EMSendKey "s: " + RSDI_claim_number + "/" + RSDI_dual_entl_number + "<newline>"
If RSDI_dual_entl_number = "none shown" then EMSendKey ": " + RSDI_claim_number + ", no dual claim numbers shown." + "<newline>"
EMSendKey "* DISA date: " + disa_date + "<newline>"
If most_recent_pay_date = "none shown" then EMSendKey "* RSDI is not currently in pay." + "<newline>"
If most_recent_pay_date <> "none shown" then EMSendKey "* RSDI current entitlement date/amt: " + most_recent_pay_date + ", $" + most_recent_monthly_amount + " ($" + net_amount + " net)" + "<newline>"
If most_recent_pay_date <> "none shown" then EMSendKey "* RSDI previous entitlement date/amt: " + second_most_recent_pay_date + ", $" + second_most_recent_monthly_amount + "<newline>"
If medicare_number = "none shown" then EMSendKey "* No Medicare info shown." + "<newline>"
If medicare_number <> "none shown" and part_b_premium <> "none shown" then EMSendKey "* Medicare # and premium: " + replace(medicare_number, " ", "") + ", Part A: " + part_a_premium + ", Part B: $" + part_b_premium + "<newline>"
If medicare_number <> "none shown" and part_b_premium = "none shown" then EMSendKey "* Medicare # and premium: " + replace(medicare_number, " ", "") + ", Part A: " + part_a_premium + ", Part B: " + part_b_premium + "<newline>"
If medicare_number <> "none shown" then EMSendKey "* Medicare start dates: Part A: " + part_a_start_date + ", Part B: " + part_b_start_date + "<newline>"
EMSendKey "* Federal living arrangement: " + fed_liv_arrange
If fed_liv_arrange = "A" then EMSendKey " (Own household)"
If fed_liv_arrange = "B" then EMSendKey " (Lives with others)"
EMSendKey "<newline>"
IF SDXE_row = 0 then EMSendKey "* No SSI currently in pay." + "<newline>"
If SDXE_row <> 0 then EMSendKey "* SSI amount: $" + most_recent_monthly_SSI_amount + "<newline>"
If SDXE_row <> 0 then EMSendKey "* Previous SSI amount: $" + second_most_recent_monthly_SSI_amount + "<newline>"
EMSendKey "---" + "<newline>"
    BeginDialog worker_sig_dialog, 0, 0, 141, 47, "Worker signature"
      EditBox 15, 25, 50, 15, worker_sig
      ButtonGroup ButtonPressed_worker_sig_dialog
        OkButton 85, 5, 50, 15
        CancelButton 85, 25, 50, 15
      Text 5, 10, 75, 10, "Sign your case note."
    EndDialog
Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript
EMSendKey worker_sig + "<newline>"
Do
  EMGetCursor case_note_row, case_note_col
  If case_note_row <> 17 then EMSendKey "---" + "<newline>"
Loop until case_note_row = 17
EMWriteScreen ">>>>>>>>PAYMENT HISTORY ON NEXT PAGE>>>>>>>>", 17, 3
PF8

EMSendKey line_01_BDXP + "<newline>"
EMSendKey line_02_BDXP + "<newline>"
EMSendKey line_03_BDXP + "<newline>"
EMSendKey line_04_BDXP + "<newline>"
EMSendKey line_05_BDXP + "<newline>"
EMSendKey line_06_BDXP + "<newline>"
EMSendKey line_07_BDXP + "<newline>"
EMSendKey line_08_BDXP + "<newline>"
EMSendKey line_09_BDXP + "<newline>"
EMSendKey line_10_BDXP + "<newline>"
EMSendKey line_11_BDXP + "<newline>"
Do
  EMGetCursor case_note_row, case_note_col
  If case_note_row <> 17 then EMSendKey "---" + "<newline>"
Loop until case_note_row = 17
EMWriteScreen ">>>>>>>>CONTINUED ON NEXT PAGE>>>>>>>>", 17, 3
PF8

EMSendKey line_01_BDXM + "<newline>"
EMSendKey line_02_BDXM + "<newline>"
EMSendKey line_03_BDXM + "<newline>"
EMSendKey line_04_BDXM + "<newline>"
EMSendKey line_05_BDXM + "<newline>"
EMSendKey line_06_BDXM + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey line_08_BDXM + "<newline>"
EMSendKey line_09_BDXM + "<newline>"
EMSendKey line_10_BDXM + "<newline>"
EMSendKey line_11_BDXM + "<newline>"
EMSendKey line_12_BDXM + "<newline>"
EMSendKey "---" + "<newline>"
EMWriteScreen ">>>>>>>>CONTINUED ON NEXT PAGE>>>>>>>>", 17, 3
PF8

EMSendKey line_01_SDXP + "<newline>"
EMSendKey line_02_SDXP + "<newline>"
EMSendKey line_03_SDXP + "<newline>"
EMSendKey line_04_SDXP + "<newline>"
EMSendKey line_05_SDXP + "<newline>"
EMSendKey line_06_SDXP + "<newline>"
EMSendKey line_07_SDXP + "<newline>"
EMSendKey line_08_SDXP + "<newline>"
EMSendKey line_09_SDXP + "<newline>"
EMSendKey line_10_SDXP + "<newline>"
EMSendKey line_11_SDXP + "<newline>"
EMSendKey line_12_SDXP + "<newline>"
EMSendKey line_13_SDXP + "<newline>"
EMSendKey line_14_SDXP + "<newline>"

script_end_procedure("")