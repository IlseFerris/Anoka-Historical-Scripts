'COPY A SCREEN FROM INQUIRY INTO TRAINING/PRODUCTION
'ONLY RONNY SHOULD BE USING THIS

EMConnect "B"
EMReadScreen B_line_01, 78, 1, 2
EMReadScreen B_line_02, 78, 2, 2
EMReadScreen B_line_03, 78, 3, 2
EMReadScreen B_line_04, 78, 4, 2
EMReadScreen B_line_05, 78, 5, 2
EMReadScreen B_line_06, 78, 6, 2
EMReadScreen B_line_07, 78, 7, 2
EMReadScreen B_line_08, 78, 8, 2
EMReadScreen B_line_09, 78, 9, 2
EMReadScreen B_line_10, 78, 10, 2
EMReadScreen B_line_11, 78, 11, 2
EMReadScreen B_line_12, 78, 12, 2
EMReadScreen B_line_13, 78, 13, 2
EMReadScreen B_line_14, 78, 14, 2
EMReadScreen B_line_15, 78, 15, 2
EMReadScreen B_line_16, 78, 16, 2
EMReadScreen B_line_17, 78, 17, 2
EMReadScreen B_line_18, 78, 18, 2
EMReadScreen B_line_19, 78, 19, 2
EMReadScreen B_line_20, 78, 20, 2
EMReadScreen B_line_21, 78, 21, 2

EMConnect "A"
EMReadScreen A_line_01, 78, 1, 2
EMReadScreen A_line_02, 78, 2, 2
EMReadScreen A_line_03, 78, 3, 2
EMReadScreen A_line_04, 78, 4, 2
EMReadScreen A_line_05, 78, 5, 2
EMReadScreen A_line_06, 78, 6, 2
EMReadScreen A_line_07, 78, 7, 2
EMReadScreen A_line_08, 78, 8, 2
EMReadScreen A_line_09, 78, 9, 2
EMReadScreen A_line_10, 78, 10, 2
EMReadScreen A_line_11, 78, 11, 2
EMReadScreen A_line_12, 78, 12, 2
EMReadScreen A_line_13, 78, 13, 2
EMReadScreen A_line_14, 78, 14, 2
EMReadScreen A_line_15, 78, 15, 2
EMReadScreen A_line_16, 78, 16, 2
EMReadScreen A_line_17, 78, 17, 2
EMReadScreen A_line_18, 78, 18, 2
EMReadScreen A_line_19, 78, 19, 2
EMReadScreen A_line_20, 78, 20, 2
EMReadScreen A_line_21, 78, 21, 2

new_array = split(B_line_01, "")
MsgBox new_array(0)