Attribute VB_Name = "modColors"
Public Const PI As Double = 3.14159265358979
Public Const PI2 As Double = 6.28318530717959


Public Function Atan2(ByVal Dx As Double, ByVal Dy As Double) As Double

    Dim theta          As Double

    If (Abs(Dx) < 0.000001) Then
        If (Abs(Dy) < 0.000001) Then
            theta = 0
        ElseIf (Dy > 0) Then
            theta = 1.5707963267949
            'theta = PI / 2
        Else
            theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        theta = Atn(Dy / Dx)
        If (Dx < 0) Then
            If (Dy >= 0) Then
                theta = theta + PI
            Else
                theta = theta - PI
            End If
        End If
    End If

    Atan2 = theta

    If Atan2 < 0 Then Atan2 = Atan2 + PI2
    If Atan2 > PI2 Then Stop: Atan2 = Atan2 - PI2

End Function
