'VARIABLES TO DECLARE

amt_of_cases_to_make = 8
married_couple = False

application_month = "05"
application_day = "01"
application_year = "13"


HH_member_array = array("01|Tina|Smith|1990|M", "03|George|Smith|2010|M")
ADDR_line_01 = "4141 Main St"
ADDR_line_02 = ""
city_line = "Anoka"
zip_line = "55303"

'Has to determine what SSN number to start with, based on a text file stored on the network. We can't duplicate SSNs so this is vital.
'Variable for the text file
file = "Q:\Blue Zone Scripts\Training case creators\SSN identifier number.txt"

'Opening text file and reading contents into SSN_identifier variable, then closing
Set objFSO = CreateObject("Scripting.FileSystemObject")
set objTS=objFSO.opentextfile(file, 1)
SSN_identifier = objTS.ReadAll
objTS.Close

SSN_identifier = cint(SSN_identifier)

EMConnect ""

For i = 1 to amt_of_cases_to_make
  EMReadScreen APPL_check, 4, 2, 45
  If APPL_check <> "APPL" then
    MsgBox "Not on APPL."
    StopScript
  End if
  
  For each HH_member in HH_member_array
    split_array = split(HH_member, "|")
    member_number = split_array(0)
    first_name = split_array(1)
    last_name = split_array(2)
    year_of_birth = split_array(3)
    gender = split_array(4)
    If member_number = "01" then
      'Enters the member 01 on the APPL panel, then jumps to the next screen.
      EMWriteScreen application_month, 4, 63
      EMWriteScreen application_day, 4, 66
      EMWriteScreen application_year, 4, 69
      EMWriteScreen last_name, 7, 30
      EMWriteScreen first_name, 7, 63
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    
      'Checks to make sure we've moved past the APPL screen. If we haven't, the script will stop.
      EMReadScreen APPL_check, 4, 2, 45
      If APPL_check = "APPL" then
        MsgBox "Error!"
        StopScript
      End if
    End if
    
  'Now it enters complete info on the HH members
    If member_number <> "01" then
      EMWriteScreen member_number, 4, 33
      EMWriteScreen last_name, 6, 30
      EMWriteScreen first_name, 6, 63
    End if
    EMWriteScreen "474", 7, 42
    EMWriteScreen "47", 7, 46
    SSN_last_four_digits = SSN_identifier

    Do
      If len(SSN_last_four_digits) < 4 then SSN_last_four_digits = "0" & SSN_last_four_digits 
    Loop until len(SSN_last_four_digits) = 4
    EMWriteScreen SSN_last_four_digits, 7, 49
    EMWriteScreen "P", 7, 68
    EMWriteScreen "01", 8, 42
    EMWriteScreen "01", 8, 45
    EMWriteScreen year_of_birth, 8, 48
    EMWriteScreen "OT", 8, 68
    EMWriteScreen gender, 9, 42
    If member_number = "02" then
      EMWriteScreen "02", 10, 42
    ElseIf member_number = "24" then
      EMWriteScreen "24", 10, 42
    ElseIf member_number <> "01" then
      EMWriteScreen "03", 10, 42
    End if
    EMWriteScreen "DL", 9, 68
    EMWriteScreen "99", 12, 42
    EMWriteScreen "99", 13, 42
    EMWriteScreen "N", 14, 68
    EMWriteScreen "N", 15, 42
    EMWriteScreen "N", 16, 68
    EMWriteScreen "x", 17, 34   'It needs to enter a code for race. It is set to do "unable to determine".
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMWriteScreen "x", 15, 12   
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    
    'Now it checks to make sure there's no duplicate SSNs. If there is, it goes back and makes another until it's done.
    Do
      EMReadScreen SSN_as_entered, 11, 4, 4
      EMReadScreen first_SSN_listed, 11, 8, 7
      If SSN_as_entered = first_SSN_listed then
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMWriteScreen "474", 7, 42
        EMWriteScreen "47", 7, 46
        SSN_identifier = SSN_identifier + 1
        SSN_last_four_digits = SSN_identifier
        Do
          If len(SSN_last_four_digits) < 4 then SSN_last_four_digits = "0" & SSN_last_four_digits 
        Loop until len(SSN_last_four_digits) = 4
        EMWriteScreen SSN_last_four_digits, 7, 49
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    Loop until SSN_as_entered <> first_SSN_listed
    
    'Now it creates the new PMI entry.
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    EMSendKey "<PF5>"
    EMWaitReady 0, 0
    EMWriteScreen "y", 6, 67
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    
    'Now it enters MEMI information
    If married_couple = False then marital_status = "N"
    If married_couple = True then marital_status = "M"
    EMWriteScreen application_month, 6, 35
    EMWriteScreen application_day, 6, 38
    EMWriteScreen "20" & application_year, 6, 41
    EMReadScreen ref_nbr, 2, 4, 33
    If ref_nbr <> "01" and ref_nbr <> "02" then marital_status = "N"
    EMWriteScreen marital_status, 7, 49
    if married_couple = True and ref_nbr = "01" then EMWriteScreen "02", 8, 49
    if married_couple = True and ref_nbr = "02" then EMWriteScreen "01", 8, 49
    age = datepart("yyyy", date) - cint(year_of_birth)
    last_grade_completed = age - 6
    If age > 18 then last_grade_completed = 12
    If age < 6 then last_grade_completed = "00"
    If len(last_grade_completed) = 1 then last_grade_completed = "0" & last_grade_completed
    EMWriteScreen last_grade_completed, 9, 49
    EMWriteScreen "y", 10, 49
    EMWriteScreen "no", 10, 78
    EMWriteScreen "y", 13, 49
    EMWriteScreen "n", 13, 78
    EMSendKey "<enter>"
    EMWaitReady 0, 0
    SSN_identifier = SSN_identifier + 1
  Next
  
  
  'Now it transmits, to get to the ADDR screen. 
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  
  'Now it enters a fake address.
  EMWriteScreen application_month, 4, 43
  EMWriteScreen application_day, 4, 46
  EMWriteScreen application_year, 4, 49
  EMWriteScreen ADDR_line_01, 6, 43
  EMWriteScreen ADDR_line_02, 7, 43
  EMWriteScreen city_line, 8, 43
  EMWriteScreen "MN", 8, 66
  EMWriteScreen zip_line, 9, 43
  EMWriteScreen "02", 9, 66
  EMWriteScreen "SF", 9, 74
  EMWriteScreen "N", 10, 43
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
  EMWriteScreen "APPL", 16, 43
  EMWriteScreen "________", 18, 43
  EMSendKey "<enter>"
  EMWaitReady 0, 0

Next

MsgBox "Done. SSN identifier ended at: " & SSN_identifier

'Opening up the text file again, writing the new number into the file, then closing
set objTS=objFSO.opentextfile(file, 2)
ObjTS.WriteLine(SSN_identifier)
objTS.Close