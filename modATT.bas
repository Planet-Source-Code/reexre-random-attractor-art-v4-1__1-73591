Attribute VB_Name = "modATT"
Option Explicit

Public Mode3D      As Boolean


Private Const PI2 = 6.28318530717959
Public NofPoints   As Long        ' Number of Points

Public PtX()       As Double      ' Points Coordinates
Public PtY()       As Double
Public PtZ()       As Double

Public PtScrX()    As Long        ' Screen Points Coordinates
Public PtScrY()    As Long
Public PtR()       As Single      ' Point Colors
Public PtG()       As Single
Public PtB()       As Single

Public Map()       As Single      'Double      'PreProcessed Output picture
'Public MapIsPixel() As Boolean


Public FastDist()  As Single      'Double      'Light Decay

Public XX          As Double
Public YY          As Double

Public ZZ          As Double


Public XRender     As Long
Public Xpic        As Long
Public Ypic        As Long

Public MinX        As Double      ' Function  Limits
Public MaxX        As Double
Public MinY        As Double
Public MaxY        As Double
Public MinZ        As Double      ' Function  Limits
Public MaxZ        As Double

Public MinR        As Double
Public MinG        As Double
Public MinB        As Double
Public MaxR        As Double
Public MaxG        As Double
Public MaxB        As Double

'#define  Pr  .241
'#define  Pg  .691
'#define  Pb  .068

'Public Const rW As Double = 3.54411764705882
'Public Const gW As Double = 10.1617647058824
'Public Const bW As Double = 1
'Public Const wR As Double = 0.34876989869754
'Public Const wG As Double = 10.1617647058824
'Public Const wB As Double = 9.84081041968162E-02
Public Const wR    As Double = 0.241
Public Const wG    As Double = 0.691
Public Const wB    As Double = 0.068


Public A           As Double

Public V(0 To 13)  As Double
Public VStart()    As Double
Public VEnd()      As Double
Public Const Range As Double = 4


Public XscrStart() As Long
Public XscrEnd()   As Long
Public YscrStart() As Long
Public YscrEnd()   As Long



Public LightR()    As Double
Public LightG()    As Double
Public LightB()    As Double

Public MaxLightR   As Double
Public MaxLightG   As Double
Public MaxLightB   As Double
Public MaxLight    As Double

Public AvR         As Double
Public AvG         As Double
Public AvB         As Double


Public CyR         As Double
Public CyG         As Double
Public Cyb         As Double


Public FN          As String

Public SignX()     As Long
Public SignY()     As Long
Public NS          As Long
Public DoLoop      As Boolean

Public CamDIST     As Double

Public Const MaxDist As Double = 200 + 4    '350 + 4
Public Const MaxDistSq As Double = MaxDist * MaxDist


'Public ATTRACTORmode As Long


Public Sub InitFASTdist()
    Dim X          As Long
    Dim y          As Long
    Dim MinD       As Single

    MinD = 1000000

    '    ReDim FastDist(-frmMAIN.PIC.Width To frmMAIN.PIC.Width, -frmMAIN.PIC.Height To frmMAIN.PIC.Height)
    '    For X = -frmMAIN.PIC.Width To frmMAIN.PIC.Width
    '        For y = -frmMAIN.PIC.Height To frmMAIN.PIC.Height

    ReDim FastDist(-MaxDist * 2 To MaxDist * 2, -MaxDist * 2 To MaxDist * 2)
    For X = -MaxDist * 2 To MaxDist * 2
        For y = -MaxDist * 2 To MaxDist * 2

            FastDist(X, y) = 10000 / (1 + X * X + y * y)
            'If FastDist(X, Y) < MinD Then MinD = FastDist(X, Y)
            'If MinD = 0 Then Stop
        Next
    Next

    'MsgBox MinD



End Sub


Public Function DISTSQLongDouble(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Double
    Dim dX         As Long
    Dim dY         As Long
    dX = X2 - X1
    dY = Y2 - Y1
    DISTSQLongDouble = dX * dX + dY * dY
End Function
Public Function DISTSQDoubleDouble(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
    Dim dX         As Double
    Dim dY         As Double
    dX = X2 - X1
    dY = Y2 - Y1
    DISTSQDoubleDouble = dX * dX + dY * dY
End Function


Public Sub INIT(HowManyPoints)


    ReDim PtX(HowManyPoints)
    ReDim PtY(HowManyPoints)

    If Mode3D Then ReDim PtZ(HowManyPoints) Else: ReDim PtZ(0)

    ReDim PtScrX(HowManyPoints)
    ReDim PtScrY(HowManyPoints)
    ReDim PtR(HowManyPoints)
    ReDim PtG(HowManyPoints)
    ReDim PtB(HowManyPoints)


End Sub
Public Sub RandomizeCURVE(StartPosToo As Boolean)
    Dim i          As Long



    If StartPosToo Then
        PtX(0) = Rnd - 0.5
        PtY(0) = Rnd - 0.5
        PtZ(0) = Rnd - 0.5
    End If

    For i = 0 To 13
        V(i) = Range * (Rnd - 0.5)
    Next

End Sub
Public Function GenerateCURVE(AttMode As Long) As Boolean
    Dim ToDiscard  As Boolean
    Dim i          As Long
    Dim dX         As Double
    Dim dY         As Double
    Dim dd         As Double

    Dim lyapunov   As Double
    Dim d0         As Double
    Dim xE         As Double
    Dim yE         As Double

    Dim xEnew      As Double
    Dim yEnew      As Double

    Dim DDX        As Double
    Dim DDY        As Double

newRND:

    MinX = 1E+29
    MinY = 1E+29
    MaxX = -1E+29
    MaxY = -1E+29

    MinR = 1E+29
    MinG = 1E+29
    MinB = 1E+29

    MaxR = -1E+29
    MaxG = -1E+29
    MaxB = -1E+29

    ToDiscard = False

    Do
        xE = PtX(0) + (Rnd - 0.5) / 1000
        yE = PtY(0) + (Rnd - 0.5) / 1000
        dX = PtX(0) - xE
        dY = PtY(0) - yE
        d0 = Sqr(dX * dX + dY * dY)
    Loop While (d0 <= 0)


    'normal
    For i = 1 To NofPoints
        XX = PtX(i - 1)
        YY = PtY(i - 1)

        'sPECIAL
        '    For I = 2 To NofPoints Step 2
        '        XX = PtX(I - 2)
        '        YY = PtY(I - 2)


        If i Mod 5000 = 0 Then frmMAIN.PB = 100 * (i / NofPoints)
        '        PtX(I) = V(0) + V(1) * XX + V(2) * XX * XX + V(3) + V(4) * XX + V(5) * XX * XX
        '        PtY(I) = V(6) + V(7) * YY + V(8) * YY * YY + V(9) + V(10) * YY + V(11) * YY * YY

        Select Case AttMode
            Case 0
                'Paul Brouke
                PtX(i) = V(0) + V(1) * XX + V(2) * XX * XX + _
                         V(3) * XX * YY + V(4) * YY + V(5) * YY * YY
                PtY(i) = V(6) + V(7) * XX + V(8) * XX * XX + _
                         V(9) * XX * YY + V(10) * YY + V(11) * YY * YY
            Case 1
                'Clifford
                PtX(i) = Sin(V(0) * YY) + V(2) * Cos(V(0) * XX)
                PtY(i) = Sin(V(1) * XX) + V(3) * Cos(V(1) * YY)

            Case 2
                'De Jong
                PtX(i) = Sin(V(0) * YY) - Cos(V(1) * XX)
                PtY(i) = Sin(V(2) * XX) - Cos(V(3) * YY)
            Case 3                'De Jong variation : Johnny Svensson
                PtX(i) = V(3) * Sin(V(0) * XX) - Sin(V(1) * YY)
                PtY(i) = V(2) * Cos(V(0) * XX) + Cos(V(1) * YY)

            Case 4
                'Sprott
                '1720 XNEW = A(1) + X * (A(2) + A(3) * X + A(4) * Y)
                '1730 XNEW = XNEW + Y * (A(5) + A(6) * Y)
                '1830 YNEW = A(7) + X * (A(8) + A(9) * X + A(10) * Y)
                '1930 YNEW = YNEW + Y * (A(11) + A(12) * Y)

                PtX(i) = V(0) + XX * (V(1) + V(2) * XX + V(3) * YY) + YY * (V(4) + V(5) * YY)
                PtY(i) = V(6) + XX * (V(7) + V(8) * XX + V(9) * YY) + YY * (V(10) + V(11) * YY)

            Case 5
                'Philp Ham
                PtX(i) = Tan(XX) * Tan(XX) - Sin(YY) * Sin(YY) + V(0)
                PtY(i) = (V(3) + 3) * Tan(XX) * Sin(YY) + V(1)

            Case 6
                'ABS

                PtX(i) = V(0) + V(1) * XX + V(2) * YY + V(3) * Abs(XX) + V(4) * Abs(YY)
                PtY(i) = V(5) + V(6) * XX + V(7) * YY + V(8) * Abs(XX) + V(9) * Abs(YY)

            Case 7
                'POW

                PtX(i) = V(0) + V(1) * XX + V(2) * YY + V(3) * Abs(XX) + V(4) * Abs(YY) ^ (V(10))
                PtY(i) = V(5) + V(6) * XX + V(7) * YY + V(8) * Abs(XX) + V(9) * Abs(YY) ^ (V(11))

            Case 8
                'SIN
                PtX(i) = V(0) + V(1) * XX + V(2) * YY + V(3) * Sin(V(4) * XX) + V(5) * Sin(V(6) * YY)
                PtY(i) = V(7) + V(8) * XX + V(9) * YY + V(10) * Sin(V(11) * XX) + V(12) * Sin(V(13) * YY)



            Case 9                'AND OR
                PtX(i) = V(0) + V(1) * XX + V(2) * YY + (V(3) * XX) And (V(4) * YY) + (V(5) * XX) Or (V(6) * YY)
                PtY(i) = V(7) + V(8) * XX + V(9) * YY + (V(10) * XX) And (V(11) * YY) + (V(12) * XX) Or (V(13) * YY)


            Case 10



            Case 15

                'Roberto Mior 2
                PtX(i) = Sin(V(0) * XX) + Sin(V(1) * YY * YY)
                PtY(i) = Sin(V(2) * YY) + Sin(V(3) * XX * XX)




        End Select

        'lyapunov*******************************************************************
        xEnew = V(0) + V(1) * xE + V(2) * xE * xE + _
                V(3) * xE * yE + V(4) * yE + V(5) * yE * yE
        yEnew = V(6) + V(7) * xE + V(8) * xE * xE + _
                V(9) * xE * yE + V(10) * yE + V(11) * yE * yE
        '** Calculate the lyapunov exponents **
        If (i > 1000) Then
            dX = PtX(i) - xEnew
            dY = PtY(i) - yEnew
            dd = Sqr(dX * dX + dY * dY)
            lyapunov = lyapunov + Log(Abs(dd / d0))
            xE = PtX(i) + d0 * dX / dd
            yE = PtY(i) + d0 * dY / dd
        End If

        If PtX(i) > MaxX Then MaxX = PtX(i)
        If PtY(i) > MaxY Then MaxY = PtY(i)
        If PtX(i) < MinX Then MinX = PtX(i)
        If PtY(i) < MinY Then MinY = PtY(i)
        '***************************************************************************



        'SpeciaL
        '       PtR(I) = Abs(PtR(I))
        '       PtG(I) = Abs(PtG(I))
        '       PtB(I) = Abs(PtB(I))

        '       PtX(I - 1) = (PtX(I) + PtX(I - 2)) * 0.5
        '       PtY(I - 1) = (PtY(I) + PtY(I - 2)) * 0.5
        '       PtR(I - 1) = (PtR(I) + PtR(I - 2)) * 0.125
        '       PtG(I - 1) = (PtG(I) + PtG(I - 2)) * 0.125
        '       PtB(I - 1) = (PtB(I) + PtB(I - 2)) * 0.125
        '-------------


        '   PtR(I) = V(0) + V(2) * XX + V(4) * XX * XX + _
            V(6) * XX * YY + V(11) * YY + V(1) * YY * YY
        '   PtG(I) = V(5) + V(6) * XX + V(7) * XX * XX + _
            V(8) * XX * YY + V(9) * YY + V(10) * YY * YY
        '   PtB(I) = V(1) + V(3) * XX + V(5) * XX * XX + _
            V(6) * XX * YY + V(7) * YY + V(9) * YY * YY

        PtR(i) = Sin(V(0) * XX) + Sin(V(1) * YY)
        PtG(i) = Sin(V(2) * XX) + Sin(V(3) * YY)
        PtB(i) = Sin(V(4) * XX) + Sin(V(5) * YY)



        If PtR(i) > MaxR Then MaxR = PtR(i)
        If PtG(i) > MaxG Then MaxG = PtG(i)
        If PtB(i) > MaxB Then MaxB = PtB(i)
        If PtR(i) < MinR Then MinR = PtR(i)
        If PtG(i) < MinG Then MinG = PtG(i)
        If PtB(i) < MinB Then MinB = PtB(i)



        If PtX(i) > 1000000 Then ToDiscard = True
        If PtY(i) > 1000000 Then ToDiscard = True
        If PtX(i) < -1000000 Then ToDiscard = True
        If PtY(i) < -1000000 Then ToDiscard = True
        If ToDiscard Then Exit For

        '1e-10
        If (Abs(PtX(i) - XX) < 0.00000001 And Abs(PtY(i) - YY) < 0.00000001) Then ToDiscard = True


    Next
    '   If ToDiscard Then GoTo newRND


    dX = MaxX - MinX
    dY = MaxY - MinY

    MaxY = MaxY + dY * 0.1
    MinY = MinY - dY * 0.1

    MaxX = MaxX + dX * 0.2618
    MinX = MinX - dX * 0.2618


    'If ToDiscard = False Then Stop

    If ToDiscard Then frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "Overflow.." & vbCrLf

    If Not (ToDiscard) Then
        If (Abs(lyapunov) < 10) Then    '10
            'MsgBox ("neutrally stable ")
            frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "neutrally stable " & vbCrLf

            ToDiscard = True
        ElseIf (lyapunov < 0) Then
            'MsgBox ("periodic " & lyapunov)
            frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "periodic " & vbCrLf

            ToDiscard = True
        End If
    End If

    If Not ToDiscard Then

        'DDX = MaxX - MinX
        'DDY = MaxY - MinY
        'If DDX > DDY Then DD = DDX Else: DD = DDY
        'If DD = 0 Then DD = 1
        'For I = 1 To NofPoints
        '    Dx = (PtX(I) - PtX(I - 1))
        '    Dy = (PtY(I) - PtY(I - 1))
        '
        '            '       HSPtoRGB 255 * Atan2(Dx, Dy) / PI2, 255 - 255 * Sqr(Dx * Dx + Dy * Dy) / DD, 255, PtR(I), PtG(I), PtB(I)
        '            HSPtoRGB 255 * Atan2(Dx, Dy) / PI2, 255, 255, PtR(I), PtG(I), PtB(I)
        '
        '            If PtR(I) > MaxR Then MaxR = PtR(I)
        '            If PtG(I) > MaxG Then MaxG = PtG(I)
        '            If PtB(I) > MaxB Then MaxB = PtB(I)
        '            If PtR(I) < MinR Then MinR = PtR(I)
        '            If PtG(I) < MinG Then MinG = PtG(I)
        '            If PtB(I) < MinB Then MinB = PtB(I)

        'Next

        'Make Color Range from 0 to 255
        For i = 1 To NofPoints


            PtR(i) = 255 * (PtR(i) - MinR) / (MaxR - MinR)
            PtG(i) = 255 * (PtG(i) - MinG) / (MaxG - MinG)
            PtB(i) = 255 * (PtB(i) - MinB) / (MaxB - MinB)
        Next

    End If

    GenerateCURVE = Not (ToDiscard)

End Function
Public Function GenerateCURVE3D(AttMode As Long) As Boolean
'Stop

    Dim ToDiscard  As Boolean
    Dim i          As Long
    Dim dX         As Double
    Dim dY         As Double

    Dim dZ         As Double


    MinX = 1E+29
    MinY = 1E+29
    MinZ = 1E+29

    MaxX = -1E+29
    MaxY = -1E+29
    MaxZ = -1E+29

    MinR = 1E+29
    MinG = 1E+29
    MinB = 1E+29

    MaxR = -1E+29
    MaxG = -1E+29
    MaxB = -1E+29



    Scree.Center.X = frmMAIN.PIC.Width \ 2
    Scree.Center.y = frmMAIN.PIC.Height \ 2
    Scree.Size.X = frmMAIN.PIC.Width
    Scree.Size.y = frmMAIN.PIC.Height

    ToDiscard = False


    'normal
    For i = 1 To NofPoints
        XX = PtX(i - 1)
        YY = PtY(i - 1)
        ZZ = PtZ(i - 1)

        'sPECIAL
        '    For I = 2 To NofPoints Step 2
        '        XX = PtX(I - 2)
        '        YY = PtY(I - 2)


        If i Mod 5000 = 0 Then frmMAIN.PB = 100 * (i / NofPoints)
        '        PtX(I) = V(0) + V(1) * XX + V(2) * XX * XX + V(3) + V(4) * XX + V(5) * XX * XX
        '        PtY(I) = V(6) + V(7) * YY + V(8) * YY * YY + V(9) + V(10) * YY + V(11) * YY * YY

        Select Case AttMode
            Case 0

                PtX(i) = Sin(V(0) * ZZ) + V(3) * Sin(V(6) * YY * XX)
                PtY(i) = Sin(V(1) * XX) + V(4) * Sin(V(7) * ZZ * YY)
                'PtZ(i) = Sin(V(2) * YY) + V(5) * Sin(V(8) * XX * ZZ)
                PtZ(i) = Sin(V(2) * YY) + V(5) * Sin(V(8) * XX * YY)

            Case 1

                PtX(i) = V(0) * XX + V(1) * YY + V(2) * XX * YY + V(9) * ZZ
                PtY(i) = V(3) * YY + V(4) * ZZ + V(5) * YY * ZZ + V(10) * XX
                PtZ(i) = V(6) * ZZ + V(7) * XX + V(8) * ZZ * XX + V(11) * YY


        End Select


        If PtX(i) > MaxX Then MaxX = PtX(i)
        If PtY(i) > MaxY Then MaxY = PtY(i)
        If PtX(i) < MinX Then MinX = PtX(i)
        If PtY(i) < MinY Then MinY = PtY(i)
        If PtZ(i) > MaxZ Then MaxZ = PtZ(i)
        If PtZ(i) < MinZ Then MinZ = PtZ(i)

        '***************************************************************************

        PtR(i) = Sin(V(0) * XX) + Sin(V(1) * YY)
        PtG(i) = Sin(V(2) * XX) + Sin(V(3) * YY)
        PtB(i) = Sin(V(4) * XX) + Sin(V(5) * YY)


        If PtR(i) > MaxR Then MaxR = PtR(i)
        If PtG(i) > MaxG Then MaxG = PtG(i)
        If PtB(i) > MaxB Then MaxB = PtB(i)
        If PtR(i) < MinR Then MinR = PtR(i)
        If PtG(i) < MinG Then MinG = PtG(i)
        If PtB(i) < MinB Then MinB = PtB(i)



        If PtX(i) > 1000000 Then ToDiscard = True
        If PtY(i) > 1000000 Then ToDiscard = True
        If PtX(i) < -1000000 Then ToDiscard = True
        If PtY(i) < -1000000 Then ToDiscard = True
        If PtZ(i) > 1000000 Then ToDiscard = True
        If PtZ(i) < -1000000 Then ToDiscard = True

        If ToDiscard Then Exit For

        '1e-10
        If (Abs(PtZ(i) - ZZ) < 0.00000001 And Abs(PtX(i) - XX) < 0.00000001 And Abs(PtY(i) - YY) < 0.00000001) Then ToDiscard = True


    Next
    '   If ToDiscard Then GoTo newRND


    dX = MaxX - MinX
    dY = MaxY - MinY
    dZ = MaxZ - MinZ


    MaxY = MaxY + dY * 0.1
    MinY = MinY - dY * 0.1

    MaxX = MaxX + dX * 0.2618
    MinX = MinX - dX * 0.2618

    MaxZ = MaxZ + dZ * 0.1
    MinZ = MinZ - dZ * 0.1

    'If ToDiscard = False Then Stop

    If ToDiscard Then frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "Overflow.." & vbCrLf


    If Not ToDiscard Then

        For i = 1 To NofPoints

            PtR(i) = 255 * (PtR(i) - MinR) / (MaxR - MinR)
            PtG(i) = 255 * (PtG(i) - MinG) / (MaxG - MinG)
            PtB(i) = 255 * (PtB(i) - MinB) / (MaxB - MinB)
        Next

    End If

    GenerateCURVE3D = Not (ToDiscard)

    If ToDiscard Then Exit Function


    RePosCamera

    UpdateCamera


End Function



Public Sub RenderSimple()
    Dim i          As Long

    If Mode3D Then
        ScreenPOINTS3D

    Else
        ScreenPOINTS
    End If

    frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "Basic Rendering..." & vbCrLf
    DoEvents
    'RENDER
    BitBlt frmMAIN.PIC.hdc, 0, 0, frmMAIN.PIC.Width, frmMAIN.PIC.Height, frmMAIN.PIC.hdc, 0, 0, vbBlackness
    For i = 1 To NofPoints
        SetPixel frmMAIN.PIC.hdc, PtScrX(i), PtScrY(i), RGB(PtR(i), PtG(i), PtB(i))    'vbWhite
        'If I Mod 10000 = 0 Then frmMAIN.PB = 100 * (I / NofPoints)
    Next
    frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "Ready." & vbCrLf
    frmMAIN.PIC.Refresh
    Beep                          '440, 100

End Sub

Public Sub ScreenPOINTS()
    Dim dX         As Double
    Dim dY         As Double
    Dim i          As Long
    If MaxX - MinX = 0 Then Exit Sub
    If MaxY - MinY = 0 Then Exit Sub

    dX = frmMAIN.PIC.Width / (MaxX - MinX)
    dY = frmMAIN.PIC.Height / (MaxY - MinY)

    For i = 1 To NofPoints
        If PtR(i) < 0 Then PtR(i) = 0
        If PtG(i) < 0 Then PtG(i) = 0
        If PtB(i) < 0 Then PtB(i) = 0
        PtScrX(i) = dX * (PtX(i) - MinX)
        PtScrY(i) = dY * (PtY(i) - MinY)
    Next

End Sub
Public Sub ScreenPOINTS3D()


    Dim P          As POINTAPI

    Dim dX         As Double
    Dim dY         As Double

    Dim i          As Long

    Dim X1         As Long
    Dim X2         As Long
    Dim Y1         As Long
    Dim Y2         As Long

    If MaxX - MinX = 0 Then Exit Sub
    If MaxY - MinY = 0 Then Exit Sub


    dX = frmMAIN.PIC.Width / (MaxX - MinX)
    dY = frmMAIN.PIC.Height / (MaxY - MinY)

    X1 = 999999999#
    Y1 = 999999999#
    X2 = -999999999#
    Y2 = -999999999#

    For i = 1 To NofPoints

        P = PointToScreen(Vector(PtX(i), PtY(i), PtZ(i)))

        PtScrX(i) = P.X
        PtScrY(i) = P.y

        If PtR(i) < 0 Then PtR(i) = 0
        If PtG(i) < 0 Then PtG(i) = 0
        If PtB(i) < 0 Then PtB(i) = 0

        If P.X < X1 Then X1 = P.X
        If P.y < Y1 Then Y1 = P.y
        If P.X > X2 Then X2 = P.X
        If P.y > Y2 Then Y2 = P.y

    Next

    ' dX = frmMAIN.PIC.Width / (X2 - X1)
    '   dY = frmMAIN.PIC.Height / (Y2 - Y1)

    '    For I = 1 To NofPoints
    '
    '        PtScrX(I) = dX * (PtX(I) - X1)
    '        PtScrY(I) = dY * (PtY(I) - Y1)
    '
    '    Next


End Sub
Public Sub RenderQUALITY(Preview As Boolean, ContrastV As Double, Ifrom As Long, iTo As Long, Optional mapMAX As Double = 0)
'Before call this, RenderSimple must be called

    Dim FindMaxMap As Boolean

    Dim i          As Long

    Dim XF         As Long
    Dim YF         As Long

    Dim Y1         As Long

    Dim Dis        As Double
    Dim Cr         As Double
    Dim Cg         As Double
    Dim Cb         As Double

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim MapV       As Double

    Dim GLOBALLIGHT As Double

    Dim MaxR2      As Double
    Dim MaxG2      As Double
    Dim MaxB2      As Double


    Dim MapMAX2    As Double
    Dim MapMAX3    As Double
    Dim MapMAX4    As Double


    Xpic = frmMAIN.PIC.Width
    Ypic = frmMAIN.PIC.Height

    If mapMAX = 0 Then FindMaxMap = True
    '
    '    If Contrast Then
    '        GLOBALLIGHT = 0.5 * (4) ^ 2    ' 16     '10 ' 2.5 ' 2    '10'3
    '    Else
    '        GLOBALLIGHT = 2.5 ^ (0.5)     ' for " ^ "
    '        GLOBALLIGHT = 1 ^ (0.25) ' 2.5 ^ (0.5)    ' for " ^ "
    '    End If

    GLOBALLIGHT = (ContrastV * 6) ^ ContrastV


    ReDim LightR(0 To frmMAIN.PIC.Width, 0 To frmMAIN.PIC.Height)
    ReDim LightG(0 To frmMAIN.PIC.Width, 0 To frmMAIN.PIC.Height)
    ReDim LightB(0 To frmMAIN.PIC.Width, 0 To frmMAIN.PIC.Height)


    ReDim Map(0 To Xpic, 0 To Ypic, 0 To 2)
    'ReDim MapIsPixel(0 To Xpic, 0 To Ypic)




ModePERFECTFASTER:
    For i = Ifrom To iTo
        XF = PtScrX(i)
        YF = PtScrY(i)
        Map(XF, YF, 0) = Map(XF, YF, 0) + PtR(i)
        Map(XF, YF, 1) = Map(XF, YF, 1) + PtG(i)
        Map(XF, YF, 2) = Map(XF, YF, 2) + PtB(i)
        'MapIsPixel(Xf, Yf) = True
        '        MAPV = (Map(Xf, Yf, 0) * wR + Map(Xf, Yf, 1) * wG + Map(Xf, Yf, 2) * wB)

        If FindMaxMap Then
            MapV = (Map(XF, YF, 0) + Map(XF, YF, 1) + Map(XF, YF, 2))
            If MapV > mapMAX Then MapMAX4 = MapMAX3: MapMAX3 = MapMAX2: MapMAX2 = mapMAX: mapMAX = MapV
        End If

        If Map(XF, YF, 0) > MaxR2 Then MaxR2 = Map(XF, YF, 0)
        If Map(XF, YF, 1) > MaxG2 Then MaxG2 = Map(XF, YF, 1)
        If Map(XF, YF, 2) > MaxB2 Then MaxB2 = Map(XF, YF, 2)

        If i Mod 1000 = 0 Then frmMAIN.PB = 100 * (i - Ifrom) / (iTo - Ifrom)
    Next


    '**********************************************************************

    For i = 1 To NS
        Map(SignX(i), SignY(i), 0) = 0.02 * MaxR2
        Map(SignX(i), SignY(i), 1) = 0.02 * MaxG2
        Map(SignX(i), SignY(i), 2) = 0.02 * MaxB2
    Next
    '**********************************************************************
    'mapMAX = (mapMAX + MapMAX2 + MapMAX3 + MapMAX4) * 0.25
    If Preview Then
        For XRender = 0 To Xpic
            For Y1 = 0 To Ypic

                If Map(XRender, Y1, 0) Or Map(XRender, Y1, 1) Or Map(XRender, Y1, 2) Then

                    MapV = 16 + MaxDistSq * (Map(XRender, Y1, 0) + Map(XRender, Y1, 1) + Map(XRender, Y1, 2)) / mapMAX
                    MapV = Sqr(MapV)

                    '1000
                    Xfrom = XRender - MapV
                    Xto = XRender + MapV
                    If Xfrom < 0 Then Xfrom = 0
                    If Xto > Xpic Then Xto = Xpic
                    Yfrom = Y1 - MapV
                    Yto = Y1 + MapV
                    If Yfrom < 0 Then Yfrom = 0
                    If Yto > Ypic Then Yto = Ypic

                    For XF = Xfrom To Xto
                        For YF = Yfrom To Yto
                            Dis = FastDist((XF - XRender), (YF - Y1))
                            LightR(XF, YF) = LightR(XF, YF) + Dis * Map(XRender, Y1, 0)
                            LightG(XF, YF) = LightG(XF, YF) + Dis * Map(XRender, Y1, 1)
                            LightB(XF, YF) = LightB(XF, YF) + Dis * Map(XRender, Y1, 2)
                        Next
                    Next
                End If
            Next
            frmMAIN.PB = 100 * (XRender / Xpic)
            DoEvents
        Next
    Else
        For XRender = 0 To Xpic
            For Y1 = 0 To Ypic
                If Map(XRender, Y1, 0) Or Map(XRender, Y1, 1) Or Map(XRender, Y1, 2) Then

                    MapV = 16 + MaxDistSq * (Map(XRender, Y1, 0) + Map(XRender, Y1, 1) + Map(XRender, Y1, 2)) / mapMAX
                    MapV = Sqr(MapV) * 2

                    Xfrom = XRender - MapV
                    Xto = XRender + MapV
                    If Xfrom < 0 Then Xfrom = 0
                    If Xto > Xpic Then Xto = Xpic
                    Yfrom = Y1 - MapV
                    Yto = Y1 + MapV
                    If Yfrom < 0 Then Yfrom = 0
                    If Yto > Ypic Then Yto = Ypic
                    For XF = Xfrom To Xto
                        For YF = Yfrom To Yto
                            Dis = FastDist((XF - XRender), (YF - Y1))
                            LightR(XF, YF) = LightR(XF, YF) + Dis * Map(XRender, Y1, 0)
                            LightG(XF, YF) = LightG(XF, YF) + Dis * Map(XRender, Y1, 1)
                            LightB(XF, YF) = LightB(XF, YF) + Dis * Map(XRender, Y1, 2)
                        Next
                    Next
                End If
            Next
            frmMAIN.PB = 100 * (XRender / Xpic)
            DoEvents
        Next

    End If

cont:

    MaxLightR = 0
    MaxLightG = 0
    MaxLightB = 0
    MaxLight = 0
    '    For XF = 0 To frmMAIN.PIC.Width
    '        For YF = 0 To frmMAIN.PIC.Height
    '
    '            If LightR(XF, YF) * wR + LightG(XF, YF) * wG + LightB(XF, YF) * wB > MaxLight Then
    '                MaxLight = LightR(XF, YF) * wR + LightG(XF, YF) * wG + LightB(XF, YF) * wB
    '            End If
    '        Next
    '    Next

    i = 0
    For XF = 0 To frmMAIN.PIC.Width
        For YF = 0 To frmMAIN.PIC.Height
            If LightR(XF, YF) < 0 Then LightR(XF, YF) = 0
            If LightG(XF, YF) < 0 Then LightG(XF, YF) = 0
            If LightB(XF, YF) < 0 Then LightB(XF, YF) = 0

            'If Contrast Then
            'Else
            'LightR(Xf, Yf) = Sqr(LightR(Xf, Yf))
            'LightG(Xf, Yf) = Sqr(LightG(Xf, Yf))
            'LightB(Xf, Yf) = Sqr(LightB(Xf, Yf))
            LightR(XF, YF) = (LightR(XF, YF)) ^ ContrastV    ' 0.25    '0.5
            LightG(XF, YF) = (LightG(XF, YF)) ^ ContrastV    '0.5
            LightB(XF, YF) = (LightB(XF, YF)) ^ ContrastV    '0.5
            '               LightR(XF, YF) = Log(MaxLight) - Log(MaxLight / (1 + LightR(XF, YF)))
            '               LightG(XF, YF) = Log(MaxLight) - Log(MaxLight / (1 + LightG(XF, YF)))
            '               LightB(XF, YF) = Log(MaxLight) - Log(MaxLight / (1 + LightB(XF, YF)))
            'End If

            If LightR(XF, YF) * wR + LightG(XF, YF) * wG + LightB(XF, YF) * wB > MaxLight Then
                MaxLight = LightR(XF, YF) * wR + LightG(XF, YF) * wG + LightB(XF, YF) * wB
            End If

            ''If LightR(Xf, Yf) > MaxLight Then MaxLight = LightR(Xf, Yf)
            ''If LightG(Xf, Yf) > MaxLight Then MaxLight = LightG(Xf, Yf)
            ''If LightB(Xf, Yf) > MaxLight Then MaxLight = LightB(Xf, Yf)
            'If LightR(Xf, Yf) * wR > MaxLightR Then MaxLightR = LightR(Xf, Yf) * wR
            'If LightG(Xf, Yf) * wG > MaxLightG Then MaxLightG = LightG(Xf, Yf) * wG
            'If LightB(Xf, Yf) * wB > MaxLightB Then MaxLightB = LightB(Xf, Yf) * wB

        Next
    Next

    'MaxLight = Log(MaxLight) * 2

    MaxLightR = GLOBALLIGHT / MaxLight
    MaxLightG = GLOBALLIGHT / MaxLight
    MaxLightB = GLOBALLIGHT / MaxLight
    '    MaxLightR = GLOBALLIGHT / MaxLightR
    '    MaxLightG = GLOBALLIGHT / MaxLightG
    '    MaxLightB = GLOBALLIGHT / MaxLightB

    For XF = 0 To frmMAIN.PIC.Width
        For YF = 0 To frmMAIN.PIC.Height
            Cr = LightR(XF, YF) * MaxLightR
            Cg = LightG(XF, YF) * MaxLightG
            Cb = LightB(XF, YF) * MaxLightB

            Cr = Cr * 255
            Cg = Cg * 255
            Cb = Cb * 255

            If Cr > 255 Then Cr = 255
            If Cg > 255 Then Cg = 255
            If Cb > 255 Then Cb = 255

            SetPixel frmMAIN.PIC.hdc, XF, YF, RGB(Cr, Cg, Cb)
        Next

        frmMAIN.PIC.Refresh
        DoEvents

    Next

    Beep                          '220, 100

End Sub

Public Sub CMDGenerate()

    frmMAIN.txtSTATUS = "Generating Curve..." & vbCrLf

    If Mode3D Then
        Do
            DoEvents
            RandomizeCURVE frmMAIN.chRNDStartPos
        Loop While Not (GenerateCURVE3D(frmMAIN.cmbATTmode.ListIndex))
    Else
        Do
            DoEvents
            RandomizeCURVE frmMAIN.chRNDStartPos
        Loop While Not (GenerateCURVE(frmMAIN.cmbATTmode.ListIndex))

    End If

    frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "SUCCESS!" & vbCrLf

    RenderSimple

End Sub
Public Sub CMDRenderQuality()

'    RenderQUALITY frmMAIN.chPREVmode, frmMAIN.chContrast, 1, NofPoints
    RenderQUALITY frmMAIN.chPREVmode, frmMAIN.sContrast / 100, 1, NofPoints

    FN = Year(Now) & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00") & " " & _
         Format(Hour(Now), "00") & "." & Format(Minute(Now), "00") & "." & Format(Second(Now), "00")
    SaveJPG frmMAIN.PIC.Image, App.Path & "\OUT\" & FN & ".jpg", 99

End Sub

Public Sub LoadSign()
    Dim i          As Long
    Dim sMaxY      As Long
    Dim sMinY      As Long
    sMaxY = -1
    sMinY = 99999

    Open App.Path & "\sign.txt" For Input As 1
    Input #1, NS
    ReDim SignX(NS)
    ReDim SignY(NS)
    For i = 1 To NS
        Input #1, SignX(i)
        SignX(i) = SignX(i) + 8
        Input #1, SignY(i)
        SignY(i) = SignY(i) + frmMAIN.PIC.Height - 52
        If SignY(i) < sMinY Then sMinY = SignY(i)
        If SignY(i) > sMaxY Then sMaxY = SignY(i)
    Next
    Close 1
    'For I = 1 To NS
    'SignX(I) = SignX(I) + (sMaxY - SignY(I)) * 0.5
    'Next
End Sub
Public Sub SaveParam(Optional FileName = "Par.txt")
    Dim i          As Long
    Open App.Path & "\Thumbnails\" & FileName For Output As 1
    Print #1, IIf(Mode3D, "3D", "2D")
    Print #1, "Attractor"
    Print #1, frmMAIN.cmbATTmode.ListIndex
    Print #1, "N of Points"
    Print #1, NofPoints
    Print #1, "StartPoint"
    Print #1, PtX(0)
    Print #1, PtY(0)
    Print #1, PtZ(0)
    Print #1, "Params"
    Print #1, UBound(V)
    For i = 0 To UBound(V)
        Print #1, V(i)
    Next
    Close 1
End Sub
Public Sub LoadParam(Optional FileName = "Par.txt", Optional ToW As Long)
    Dim i          As Long
    Dim N          As Long
    Dim S          As String

    Open App.Path & "\Thumbnails\" & FileName For Input As 1
    Input #1, S
    If S = "3D" Then Mode3D = True Else: Mode3D = False
    frmMAIN.PopulateAttractors
    Input #1, S
    Input #1, S: frmMAIN.cmbATTmode.ListIndex = Val(S)
    Input #1, S
    Input #1, S: NofPoints = Val(S)

    INIT NofPoints

    Input #1, S
    Line Input #1, S: PtX(0) = CDbl(S)
    Line Input #1, S: PtY(0) = CDbl(S)
    Line Input #1, S: PtZ(0) = CDbl(S)
    Input #1, S
    Line Input #1, S: N = S
    For i = 0 To N
        Line Input #1, S: V(i) = CDbl(S)
    Next
    Close 1


    If ToW <> 0 Then
        If ToW = 1 Then
            ReDim VStart(N)
        End If
        If ToW = 2 Then
            ReDim VEnd(N)
        End If

        For i = 0 To N
            If ToW = 1 Then
                VStart(i) = V(i)
            End If

            If ToW = 2 Then
                VEnd(i) = V(i)
            End If

        Next
    End If


End Sub

Public Sub SavePTS(Optional FileN As String = "PTS.txt")
    Dim i          As Long
    Open App.Path & "\" & FileN For Output As 1
    Print #1, NofPoints
    For i = 0 To NofPoints
        Print #1, PtScrX(i)
        Print #1, PtScrY(i)
    Next
    Close 1
End Sub
Public Sub LoadPTS(Optional FileName = "PTS.txt", Optional ToW As Long)
    Dim i          As Long

    Dim S          As String

    Open App.Path & "\" & FileName For Input As 1
    Line Input #1, S: NofPoints = S
    For i = 0 To NofPoints
        Line Input #1, S: PtScrX(i) = S
        Line Input #1, S: PtScrY(i) = S
    Next
    Close 1


    If ToW <> 0 Then
        If ToW = 1 Then
            ReDim XscrStart(NofPoints)
            ReDim YscrStart(NofPoints)
        End If
        If ToW = 2 Then
            ReDim XscrEnd(NofPoints)
            ReDim YscrEnd(NofPoints)
        End If

        For i = 0 To NofPoints
            If ToW = 1 Then
                XscrStart(i) = PtScrX(i)
                YscrStart(i) = PtScrY(i)
            End If

            If ToW = 2 Then
                XscrEnd(i) = PtScrX(i)
                YscrEnd(i) = PtScrY(i)
            End If

        Next
    End If

End Sub


Public Sub RenderALL3D()
    Dim A          As Double
    Dim St         As Double
    Dim j          As Long

    Dim i          As Long
    Dim FN1        As String
    Dim FN2        As String
    Dim I2         As Long
    Dim iTo        As Long

    Dim steps      As Long

    Dim MAX        As Double
    Dim MapV       As Double
    Dim mapMAX     As Double
    Dim MapMAX2    As Double
    Dim MapMAX3    As Double
    Dim MapMAX4    As Double


    Dim XF         As Long
    Dim YF         As Long

    Dim CoolRender3D As Boolean

    steps = InputBox("How many Steps for 360° ?", "Steps", 16)

    j = MsgBox(" 'Cool Render' Quality ? ", vbYesNo, "Quality")
    If j = vbYes Then CoolRender3D = True


    St = PI2 / steps

    GoTo skipL
    'FindAnimation Light
    MAX = 0
    j = 0
    For A = 0 To PI2 - St Step St
        camera.cFrom.X = Cos(A) * CamDIST
        camera.cFrom.Z = Sin(A) * CamDIST
        UpdateCamera
        'clear map
        ReDim Map(0 To frmMAIN.PIC.Width, 0 To frmMAIN.PIC.Height, 0 To 2)
        MapV = 0
        mapMAX = 0
        For i = 1 To NofPoints
            XF = PtScrX(i)
            YF = PtScrY(i)
            Map(XF, YF, 0) = Map(XF, YF, 0) + PtR(i)
            Map(XF, YF, 1) = Map(XF, YF, 1) + PtG(i)
            Map(XF, YF, 2) = Map(XF, YF, 2) + PtB(i)
            'MapIsPixel(Xf, Yf) = True
            '        MAPV = (Map(Xf, Yf, 0) * wR + Map(Xf, Yf, 1) * wG + Map(Xf, Yf, 2) * wB)
            MapV = (Map(XF, YF, 0) + Map(XF, YF, 1) + Map(XF, YF, 2))

            If MapV > mapMAX Then MapMAX4 = MapMAX3: MapMAX3 = MapMAX2: MapMAX2 = mapMAX: mapMAX = MapV

            If i Mod 1000 = 0 Then frmMAIN.PB = 100 * i / NofPoints: DoEvents

        Next
        MAX = MAX + mapMAX
        j = j + 1
    Next
    MAX = MAX / j
    '**********************************************************************+

skipL:



    j = 0
    For A = 0 To PI2 - St Step St
        frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & Int(100 * A / PI2) & "%" & vbCrLf

        camera.cFrom.X = Cos(A) * CamDIST
        camera.cFrom.Z = Sin(A) * CamDIST
        UpdateCamera

        If CoolRender3D Then
            ScreenPOINTS3D
            '            RenderQUALITY frmMAIN.chPREVmode, frmMAIN.chContrast, 1, NofPoints, 0     'max
            RenderQUALITY frmMAIN.chPREVmode, frmMAIN.sContrast / 100, 1, NofPoints, 0    'max

            DoEvents              '----------------
        Else
            '----------------
            RenderSimple
            DoEvents
            '----------------
        End If
        FN1 = App.Path & "\ANIM\" & Format(j, "0000000") & ".jpg"
        SaveJPG frmMAIN.PIC.Image, FN1, 99

        j = j + 1

    Next


    iTo = j - 1
    I2 = j - 1

    For i = 0 To iTo * 4
        I2 = I2 + 1
        FN1 = App.Path & "\ANIM\" & Format(i, "0000000") & ".jpg"
        FN2 = App.Path & "\ANIM\" & Format(I2, "0000000") & ".jpg"
        FileCopy FN1, FN2
    Next

    frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "100%" & vbCrLf

End Sub

Public Sub RePosCamera()
    Scree.Center.X = frmMAIN.PIC.Width \ 2
    Scree.Center.y = frmMAIN.PIC.Height \ 2
    Scree.Size.X = frmMAIN.PIC.Width
    Scree.Size.y = frmMAIN.PIC.Height

    CamDIST = MaxX - MinX
    If MaxY - MinY > CamDIST Then CamDIST = MaxY - MinY
    If MaxZ - MinZ > CamDIST Then CamDIST = MaxZ - MinZ

    CamDIST = CamDIST * 1.5       '2

    camera.cTo.X = (MaxX + MinX) * 0.5
    camera.cTo.y = (MaxY + MinY) * 0.5
    camera.cTo.Z = (MaxZ + MinZ) * 0.5

    camera.cFrom.X = (MaxX + MinX) * 0.5 + CamDIST * Cos(0)
    camera.cFrom.y = (MaxY + MinY) * 0.5 - CamDIST * 1
    camera.cFrom.Z = (MaxZ + MinZ) * 0.5 + CamDIST * Sin(0)

    camera.NearPlane = 0
    camera.FarPlane = 1000000
    camera.Projection = PERSPECTIVE
    camera.cUp.X = 0
    camera.cUp.y = -1
    camera.cUp.Z = 0
    camera.Zoom = 1
    camera.ANGh = 30
    camera.ANGv = 30
End Sub
