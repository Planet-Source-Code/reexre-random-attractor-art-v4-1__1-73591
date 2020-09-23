Attribute VB_Name = "modPLAYmusic"
'Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public CurNotepos  As Long


Public Sub PlayMusic()
    Dim N          As Long
    Dim D          As Long
    Dim t          As Long
    Dim n2         As Long




    OpenMIDI

    For CurNotepos = 1 To NofPoints Step 1    ' 2000

        'N = 1 + 1000 * (PtX(CurNotepos) - MinX) / (MaxX - MinX)
        N = 12 + 24 * (PtX(CurNotepos) - MinX) / (MaxX - MinX)

        n2 = N Mod 12
        If n2 = 1 Then N = N - 1
        If n2 = 3 Then N = N - 1
        If n2 = 5 Then N = N - 1
        If n2 = 8 Then N = N - 1
        If n2 = 10 Then N = N - 1


        D = 4 * (PtY(CurNotepos) - MinY) / (MaxY - MinY)
        D = 80 * 2 ^ D

        'Beep N, D
        PlayNote 12 + N

        t = GetTickCount
        Do While GetTickCount < D + t + (D * 0.1)
            DoEvents
        Loop
        PlayNoteOff 12 + N

    Next
    CloseMIDI




End Sub

