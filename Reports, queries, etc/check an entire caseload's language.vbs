'VARIABLES TO DECLARE


EMConnect ""

EMReadScreen ACTV_check, 4, 2, 48
If ACTV_check <> "ACTV" then
  MsgBox "Not on ACTV."
  StopScript
End if

row = 7

Do
  EMReadScreen case_number, 8, row, 12
  If case_number <> "        " then case_number_array = case_number_array & trim(case_number) & " "
  row = row + 1
  If row = 19 then
    EMSendKey "<PF8>"
    EMWaitReady 1, 1
    row = 7
    EMReadScreen last_page_check, 4, 24, 19
  End if
Loop until last_page_check = "PAGE" or case_number = "        "

case_number_array = split(case_number_array)

For each case_number in case_number_array
  If case_number <> "" then
    Do
      EMSendKey "<PF3>"
      EMWaitReady 1, 1
      EMReadScreen SELF_check, 4, 2, 50
    Loop until SELF_check = "SELF"
  
    EMWriteScreen "stat", 16, 43
    EMWriteScreen "________", 18, 43
    EMWriteScreen case_number, 18, 43
    EMWriteScreen "memb", 21, 70
    EMSendKey "<enter>"
    EMWaitReady 1, 1

    EMReadScreen MEMB_check, 4, 2, 48
    If MEMB_check <> "MEMB" then
      EMSendKey "<enter>"
      EMWaitReady 1, 1
    End if

    EMReadScreen language_spoken, 2, 12, 42
    language_spoken_array = language_spoken_array & language_spoken & " "
  End if
Next

language_spoken_array = split(trim(language_spoken_array))

Amharic_speakers = filter(language_spoken_array, "09")
Arabic_speakers = filter(language_spoken_array, "10")
ASL_speakers = filter(language_spoken_array, "08")
Burmese_speakers = filter(language_spoken_array, "14")
Cantonese_speakers = filter(language_spoken_array, "15")
English_speakers = filter(language_spoken_array, "99")
French_speakers = filter(language_spoken_array, "16")
Hmong_speakers = filter(language_spoken_array, "02")
Khmer_speakers = filter(language_spoken_array, "04")
Korean_speakers = filter(language_spoken_array, "20")
Karen_speakers = filter(language_spoken_array, "21")
Laotian_speakers = filter(language_spoken_array, "05")
Mandarin_speakers = filter(language_spoken_array, "17")
Oromo_speakers = filter(language_spoken_array, "12")
Russian_speakers = filter(language_spoken_array, "06")
Serbo_Croatian_speakers = filter(language_spoken_array, "11")
Somali_speakers = filter(language_spoken_array, "07")
Spanish_speakers = filter(language_spoken_array, "01")
Swahili_speakers = filter(language_spoken_array, "18")
Tigrinya_speakers = filter(language_spoken_array, "13")
Vietnamese_speakers = filter(language_spoken_array, "03")
Yoruba_speakers = filter(language_spoken_array, "19")
Unknown_speakers = filter(language_spoken_array, "97")
Other_speakers = filter(language_spoken_array, "98")

amt_of_Amharic_speakers = ubound(Amharic_speakers) + 1
amt_of_Arabic_speakers = ubound(Arabic_speakers) + 1
amt_of_ASL_speakers = ubound(ASL_speakers) + 1
amt_of_Burmese_speakers = ubound(Burmese_speakers) + 1
amt_of_Cantonese_speakers = ubound(Cantonese_speakers) + 1
amt_of_English_speakers = ubound(English_speakers) + 1
amt_of_French_speakers = ubound(French_speakers) + 1
amt_of_Hmong_speakers = ubound(Hmong_speakers) + 1
amt_of_Khmer_speakers = ubound(Khmer_speakers) + 1
amt_of_Korean_speakers = ubound(Korean_speakers) + 1
amt_of_Karen_speakers = ubound(Karen_speakers) + 1
amt_of_Laotian_speakers = ubound(Laotian_speakers) + 1
amt_of_Mandarin_speakers = ubound(Mandarin_speakers) + 1
amt_of_Oromo_speakers = ubound(Oromo_speakers) + 1
amt_of_Russian_speakers = ubound(Russian_speakers) + 1
amt_of_Serbo_Croatian_speakers = ubound(Serbo_Croatian_speakers) + 1
amt_of_Somali_speakers = ubound(Somali_speakers) + 1
amt_of_Spanish_speakers = ubound(Spanish_speakers) + 1
amt_of_Swahili_speakers = ubound(Swahili_speakers) + 1
amt_of_Tigrinya_speakers = ubound(Tigrinya_speakers) + 1
amt_of_Vietnamese_speakers = ubound(Vietnamese_speakers) + 1
amt_of_Yoruba_speakers = ubound(Yoruba_speakers) + 1
amt_of_Unknown_speakers = ubound(Unknown_speakers) + 1
amt_of_Other_speakers = ubound(Other_speakers) + 1



Msgbox ""