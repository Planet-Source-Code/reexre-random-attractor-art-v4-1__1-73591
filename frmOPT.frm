VERSION 5.00
Begin VB.Form frmOPT 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Options"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton o3D 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3D"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton o2D 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2D"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.HScrollBar WW 
      Height          =   375
      LargeChange     =   10
      Left            =   120
      Max             =   350
      Min             =   40
      SmallChange     =   10
      TabIndex        =   5
      Top             =   2400
      Value           =   120
      Width           =   4335
   End
   Begin VB.TextBox TextINFO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmOPT.frx":0000
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reset"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Apply"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.HScrollBar sNOP 
      Height          =   375
      LargeChange     =   10
      Left            =   120
      Max             =   6000
      Min             =   100
      SmallChange     =   10
      TabIndex        =   0
      Top             =   1560
      Value           =   100
      Width           =   4335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"frmOPT.frx":0009
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label labWW 
      BackColor       =   &H00FFC0C0&
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
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label LabelNOP 
      BackColor       =   &H00FFC0C0&
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
      Left            =   285
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
End
Attribute VB_Name = "frmOPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private MP         As Long
Private MPP        As Long

Private Area       As Long

Private VX         As Long
Private VY         As Long

Private NPchanged  As Boolean
Private OldMode3D  As Boolean

Private Type MEMORYSTATUS
    dwLength       As Long
    dwMemoryLoad   As Long
    dwTotalPhys    As Long
    dwAvailPhys    As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private mS         As MEMORYSTATUS

Private Sub cmdApply_Click()
    ApplyChanges
End Sub
Public Sub ApplyChanges()
    NofPoints = CLng(sNOP) * 1000

    frmMAIN.PIC.Cls
    frmMAIN.PIC.Width = VX
    frmMAIN.PIC.Height = VY
    frmMAIN.PIC.Refresh


    If OldMode3D <> Mode3D Then OldMode3D = Mode3D: INIT NofPoints
    If NPchanged Then
        INIT NofPoints
    Else
        If Mode3D Then
            RePosCamera
            UpdateCamera
            ScreenPOINTS3D
        Else
            ScreenPOINTS
        End If
    End If

    InitFASTdist
    LoadSign

    NPchanged = False
    Unload Me

End Sub
Private Sub cmdReset_Click()
    sNOP = NofPoints / 1000
    WW = frmMAIN.PIC.Width * 0.1
    VX = WW * 10
    VY = WW * 10 * 0.618
    NPchanged = False
End Sub



Private Sub Form_Activate()
    Me.Left = frmMAIN.Left + 140 * Screen.TwipsPerPixelX

    Me.Top = frmMAIN.Top + 55 * Screen.TwipsPerPixelY

    If Mode3D Then o3D = True Else: o2D = True

    OldMode3D = Mode3D

End Sub

Private Sub Form_Load()
    sNOP = NofPoints / 1000
    WW = frmMAIN.PIC.Width * 0.1
    VX = WW * 10
    VY = WW * 10 * 0.618
    NPchanged = False
End Sub

Private Sub o2D_Click()
    If o2D Then Mode3D = False
    frmMAIN.PopulateAttractors
    UpDateMem
End Sub

Private Sub o3D_Click()
    If o3D Then Mode3D = True
    frmMAIN.PopulateAttractors
    UpDateMem

End Sub

Private Sub sNOP_Change()
    LabelNOP = "Number of Points: " & CLng(sNOP) * 1000
    UpDateMem
    NPchanged = True

End Sub

Private Sub WW_Change()
    VX = WW * 10
    VY = WW * 10 * 0.618
    labWW = "Picture " & VX & " x " & VY
    UpDateMem
End Sub



Public Sub UpDateMem()
    Dim MDA        As Long
    '    MDA = MaxDist * MaxDist * 4
    MDA = (2 * MaxDist) ^ 2 * 4

    TextINFO = "Approximate Used Memory " & vbCrLf & vbCrLf

    'x,y,xscr,yscr (r,g,b)
    MP = CLng(sNOP) * 1000 * (8 + 8 + IIf(Mode3D, 8, 0) + 4 + 4 + (4 + 4 + 4))

    TextINFO = TextINFO & "for Points: " & (MP \ 1024) \ 1024 & " MBytes" & vbCrLf

    Area = VX * VY

    '    (area*4) * single + area *( SingleR +singleG+SingleB+boolean)
    '    TextINFO = TextINFO & "for Image : " & (((Area * 4) * 4 + Area * (4 + 4 + 4 + 2)) \ 1024) \ 1024 & " Mbytes"
    '   TextINFO = TextINFO & "for Image : " & (((42025 * 4) * 4 + Area * (4 + 4 + 4 + 2)) \ 1024) \ 1024 & " Mbytes"

    MPP = (MDA * 4 + Area * (4 + 4 + 4))
    TextINFO = TextINFO & "for Image : " & MPP \ 1024 \ 1024 & " Mbytes" & vbCrLf & vbCrLf

    GlobalMemoryStatus mS

    TextINFO = TextINFO & "Physical Free Memory: " & mS.dwAvailPhys \ 1024 \ 1024 & " Mbytes" & vbCrLf



End Sub
