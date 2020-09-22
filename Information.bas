Attribute VB_Name = "Information"
'This is a "Trial Use" same program.  The Registry.bas file
'was not written by me.  It came from a Registry Editor
'type program and I can not recall the author.
'The program first checks for a "dll" file
'If the file exist, the trial uses have been used
'If the file does not exist, it checks the runcount
'stored in the registry.
'Even if the fake "dll" is found and deleted, the
'program will create it again.
'The program may not be un-installed and re-used
'because of the runcount in the registry.
'(most people would not know what to look for anyway)

Public Function ENCRYPT(sString As String, lLEn As Long) As String
    Dim I As Long
    Dim NewChar As Long
    I = 1 'can't start a String at 0 :-)
    Do Until I = lLEn + 1
        NewChar = Asc(Mid(sString, I, 1)) + 13
        ENCRYPT = ENCRYPT + Chr(NewChar)
        I = I + 1
    Loop
End Function
Public Function DECRYPT(sString As String, lLEn As Long) As String
    Dim I As Long
    Dim NewChar As Long
    I = 1
    Do Until I = lLEn + 1
        NewChar = Asc(Mid(sString, I, 1)) - 13
        DECRYPT = DECRYPT + Chr(NewChar)
        I = I + 1
    Loop
End Function
