VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAIN 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Random Attractor ART  v4.1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar sContrast 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   12
      TabIndex        =   24
      Top             =   4200
      Value           =   50
      Width           =   1575
   End
   Begin VB.CommandButton cmd3Danimation 
      Caption         =   "3D ANIMATION"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   23
      Top             =   9840
      Width           =   1575
   End
   Begin VB.PictureBox PicSave 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2625
      Left            =   4080
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   22
      Top             =   8280
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.CommandButton cmdSAVE 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Save Curve [Remember to save only after ""CoolRender""]"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdLOAD 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Load Curve [Usefull to render with different Size]"
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbortRender 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Abort Render"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "don't want to wait this renderization..."
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdOPT 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Basic Options"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Open Options Tool"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton CMDstopLoop 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop Loop"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      ToolTipText     =   "Automatically continue to Generate Images untill application Exit. (Sometimes stop due to crash. ... bugs to fix)"
      Top             =   10800
      Width           =   1095
   End
   Begin VB.CheckBox chRNDStartPos 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RND start POS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Randomize Start Position too."
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdLittleVAR 
      BackColor       =   &H00FFC0C0&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   15
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton cmdSave2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      ToolTipText     =   "Click here to Set this as the SECOND Curve of Animation"
      Top             =   9000
      Width           =   375
   End
   Begin VB.CommandButton cmdSave1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Click here to Set this as the FIRST curve of Animation"
      Top             =   9000
      Width           =   375
   End
   Begin VB.CommandButton cmdANIMATION 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Make ANIMATION"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Click Here to RENDER Animation frames from 1st Curve to 2nd Curve."
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSTOPmidi 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Stop Music"
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdPLAY 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Play It!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Play this curve as music"
      Top             =   7920
      Width           =   1575
   End
   Begin VB.ComboBox cmbATTmode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtSTATUS 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox chPREVmode 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Preview MODE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Leave it Checked: Not in Preview Mode the Quality is a little bit better."
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton CommandCICLO 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Infinite LOOP"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Automatically continue to Generate Images untill application Exit or StopLoop. (Sometimes stop due to crash. ... bugs to fix)"
      Top             =   10560
      Width           =   1575
   End
   Begin VB.CommandButton CommandRNDCurve 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Generate Curve, if you're not satisfied click again.."
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton CommandV1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Try to Center"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click Here when you see that the curve is not centered."
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdQUAL 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COOL - Render"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "RENDER"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Jellyka - Estrya's Handwriting"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   7665
      Left            =   1920
      ScaleHeight     =   511
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   853
      TabIndex        =   0
      Top             =   360
      Width           =   12795
   End
   Begin VB.Label lContra 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contrast 50%"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ATTRACTOR:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbortRender_Click()
    XRender = Xpic + 1
End Sub

Private Sub cmdANIMATION_Click()

    Dim EN         As Long
    Dim ASTEP

    Dim P1         As Double
    Dim P2         As Double
    Dim i          As Long
    Dim I2         As Long

    Dim FN1        As String
    Dim FN2        As String

    Dim Nfr        As Long

    GoTo Skip

    '    ST = 1
    '    For EN = 2 To NofPoints Step ASTEP
    '        II = II + 1
    '        RenderQUALITY True, ST, EN
    '        FN = "\ANIM\" & Format(II, "0000000")
    '        SaveJPG frmMAIN.PIC.Image, App.Path & FN & ".jpg", 99
    '    Next
    '
    '    EN = NofPoints
    '    For ST = 2 To NofPoints - 2 Step ASTEP
    '        II = II + 1
    '        RenderQUALITY True, ST, EN
    '        FN = "\ANIM\" & Format(II, "0000000")
    '        SaveJPG frmMAIN.PIC.Image, App.Path & FN & ".jpg", 99
    '    Next
    '    EN = NofPoints
    '    For ST = 1 To NofPoints - ASTEP Step ASTEP
    '    EN = ST + ASTEP
    '        II = II + 1
    '        RenderQUALITY True, ST, EN
    '        FN = "\ANIM\" & Format(II, "0000000")
    '        SaveJPG frmMAIN.PIC.Image, App.Path & FN & ".jpg", 99
    '    Next

    GoTo fineok


Skip:

    Nfr = InputBox("Animation Length [Frames]: ", , 120)
    If Nfr = 0 Then Nfr = 10

    LoadParam "par1.txt", 1
    LoadParam "par2.txt", 2

    For I2 = 0 To Nfr
        P1 = 1 - I2 / Nfr
        P2 = 1 - P1

        For i = 0 To 13
            V(i) = VStart(i) * P1 + VEnd(i) * P2
        Next

        GenerateCURVE cmbATTmode.ListIndex
        'RenderSimple
        ScreenPOINTS
        'RenderQUALITY True, frmMAIN.chContrast, 1, NofPoints
        RenderQUALITY True, frmMAIN.sContrast / 100, 1, NofPoints

        FN1 = App.Path & "\ANIM\" & Format(I2, "0000000") & ".jpg"
        SaveJPG frmMAIN.PIC.Image, FN1, 99

    Next

    GoTo fineok

    I2 = Nfr
    For i = Nfr - 1 To 1 Step -1
        I2 = I2 + 1
        FN1 = App.Path & "\ANIM\" & Format(i, "0000000") & ".jpg"
        FN2 = App.Path & "\ANIM\" & Format(I2, "0000000") & ".jpg"
        FileCopy FN1, FN2
    Next


fineok:

End Sub

Private Sub cmdLIttleVar_Click()
    Dim i          As Long

    Me.MousePointer = 11

    For i = 0 To 13
        V(i) = V(i) + 0.3 * (Rnd - 0.5)
    Next

    If GenerateCURVE(cmbATTmode.ListIndex) Then
        RenderSimple
    End If

    Me.MousePointer = 0

End Sub

Private Sub cmdLOAD_Click()
    frmTHUMB.Show

End Sub

Private Sub cmdOPT_Click()
    frmOPT.Show

End Sub

Private Sub cmdPLAY_Click()
    PlayMusic
End Sub

Private Sub cmdSAVE_Click()
    SetStretchBltMode PicSave.hdc, 4
    StretchBlt PicSave.hdc, 0, 0, PicSave.Width, PicSave.Height, PIC.hdc, 0, 0, PIC.Width, PIC.Height, vbSrcCopy
    PicSave.CurrentX = 1
    PicSave.CurrentY = 1
    If Mode3D Then
        PicSave.Print "3D Pts: " & NofPoints
    Else
        PicSave.Print "Pts: " & NofPoints
    End If

    PicSave.Refresh
    SaveJPG PicSave.Image, App.Path & "\Thumbnails\" & FN & ".jpg", 80
    SaveParam FN & ".txt"

    MsgBox FN & " Saved!", vbInformation


End Sub

Private Sub cmdSave1_Click()
    SaveParam "par1.txt"
End Sub

Private Sub cmdSave2_Click()
    SaveParam "par2.txt"
End Sub

Private Sub cmdSTOPmidi_Click()
    CurNotepos = NofPoints

End Sub

Private Sub CMDstopLoop_Click()
    DoLoop = False

End Sub

Private Sub cmd3Danimation_Click()
    RenderALL3D

End Sub

Private Sub CommandCICLO_Click()

    DoLoop = True

    Me.MousePointer = 11
    Do
        CMDGenerate
        CMDRenderQuality
        DoEvents
        DoEvents

    Loop While DoLoop
    Me.MousePointer = 0
End Sub

Private Sub Form_Initialize()
    XPStyle False

End Sub

Private Sub Form_Load()


'Double 8
'Single 4
'Long 4
'Boole 2

'Dim U As Single
'MsgBox Len(U)

    Randomize Timer

    NofPoints = 200000            ' 2000000

    INIT NofPoints

    PIC.Width = 1200
    PIC.Height = PIC.Width * 0.618
    PicSave.Height = PicSave.Width * 0.618


    'PB.Width = PIC.Width
    'PB.Top = PIC.Top + PIC.Height + 10
    InitFASTdist
    LoadSign
    ProcessPrioritySet , , ppbelownormal    'ppidle '

    PopulateAttractors

    If Dir(App.Path & "\OUT", vbDirectory) = "" Then MkDir App.Path & "\OUT"
    If Dir(App.Path & "\ANIM", vbDirectory) = "" Then MkDir App.Path & "\ANIM"

End Sub

Private Sub CommandRNDCurve_Click()

    Me.MousePointer = 11

    CMDGenerate

    Me.MousePointer = 0

    SaveParam

End Sub

Private Sub CommandV1_Click()
'PtX(0) = PtX(0) + (MaxX - MinX) * 0.05
    Me.MousePointer = 11
    PtX(0) = PtX(1000)
    PtY(0) = PtY(1000)
    If Mode3D Then PtZ(0) = PtZ(1000)
    GenerateCURVE cmbATTmode.ListIndex
    RenderSimple

    Me.MousePointer = 0

End Sub


Private Sub cmdQUAL_Click()

    Me.MousePointer = 11

    frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "Cool Rendering" & IIf(frmMAIN.chPREVmode, " (Fast Mode) ", " (Perfect Mode) ") & "..." & vbCrLf
    DoEvents

    CMDRenderQuality

    frmMAIN.txtSTATUS = frmMAIN.txtSTATUS & "Ready." & vbCrLf

    Me.MousePointer = 0

End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        PB.Left = 8
        PB.Width = Me.ScaleWidth - 16
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseMIDI
    End
End Sub

'Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim i          As Long
    Dim C          As Long
    Exit Sub
    If Button = 1 Then
        NS = 0
        PIC.Print "by Roberto Mior"
        PIC.Refresh
        For X = 0 To 500
            For y = 0 To 200
                C = GetPixel(PIC.hdc, X, y)
                If C <> 0 Then
                    NS = NS + 1
                    ReDim Preserve SignX(NS)
                    ReDim Preserve SignY(NS)
                    SignX(NS) = X
                    SignY(NS) = y
                End If
            Next
        Next
        Open App.Path & "\sign.txt" For Output As 1
        Print #1, NS
        For i = 1 To NS
            Print #1, SignX(i)
            Print #1, SignY(i)
        Next
        Close 1
        MsgBox "Saved"
    End If
End Sub

Private Sub sContrast_Change()
    lContra = "Contrast " & sContrast & "%"
End Sub

Private Sub sContrast_Scroll()
    lContra = "Contrast " & sContrast & "%"
End Sub

Private Sub txtSTATUS_Change()
    If Len(txtSTATUS) > 1000 Then txtSTATUS = Right$(txtSTATUS, 100)
    txtSTATUS.SelLength = 1
    txtSTATUS.SelStart = Len(txtSTATUS) - 1
End Sub


Public Sub ReOrder()
    ReDim Xsort(NofPoints) As Double
    ReDim Ysort(NofPoints) As Double
    ReDim Sorted(NofPoints) As Boolean
    Dim SC         As Long
    Dim tx         As Double
    Dim ty         As Double

    Dim i          As Long
    Dim j          As Long
    Dim K          As Long
    Dim D          As Double
    Dim Dmax       As Double
    Dim S          As Long

    Dim Isor       As Long
    Dim Icur       As Long
    Dim Bes        As Long




    Xsort(0) = PtX(0)
    Ysort(0) = PtY(0)
    Isor = 0
AG:

    Dmax = -1
    For i = 1 To NofPoints
        If Not (Sorted(i)) Then
            D = FastDist(PtX(Isor) - PtX(i), PtY(Isor) - PtY(i))
            If D > Dmax Then
                Dmax = D
                Bes = i
            End If
        End If
    Next

    Me.Caption = Isor & "   " & Bes
    Isor = Isor + 1
    Sorted(Bes) = True
    Xsort(Isor) = PtX(Bes)
    Ysort(Isor) = PtY(Bes)


    If Isor <= NofPoints Then GoTo AG

    For i = 1 To NofPoints
        PtX(i) = Xsort(i)
        PtY(i) = Ysort(i)
    Next


End Sub

Public Sub PopulateAttractors()

    If Mode3D Then
        cmbATTmode.Clear
        cmbATTmode.AddItem "3D Mior 1"
        cmbATTmode.AddItem "3D Mior 2"
    Else

        cmbATTmode.Clear
        cmbATTmode.AddItem "Paul Bourke"
        cmbATTmode.AddItem "Clifford Pickover"
        cmbATTmode.AddItem "Peter De Jong"
        cmbATTmode.AddItem "Johnny Svensson"

        cmbATTmode.AddItem "Julien Clinton Sprott"
        cmbATTmode.AddItem "Philp Ham"

        cmbATTmode.AddItem "ABS"
        cmbATTmode.AddItem "POW"
        cmbATTmode.AddItem "AND OR"
        cmbATTmode.AddItem "empty"
        cmbATTmode.AddItem "empty"
        cmbATTmode.AddItem "empty"
        cmbATTmode.AddItem "empty"
        cmbATTmode.AddItem "empty"
        cmbATTmode.AddItem "empty"
        cmbATTmode.AddItem "Roberto Mior 2"


    End If
    cmbATTmode.ListIndex = 0

    frmMAIN.cmd3Danimation.Enabled = Mode3D
End Sub
