VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3375
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pOptions 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   1
      Left            =   6960
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   6510
      Begin VB.Frame Frame3 
         Caption         =   "Skin selector"
         Height          =   1815
         Left            =   225
         TabIndex        =   10
         Top             =   165
         Width           =   3255
         Begin VB.ListBox List1 
            Height          =   1230
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2775
         End
      End
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   6000
      Pattern         =   "*.ini"
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox pOptions 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   0
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   2
      Top             =   600
      Width           =   6510
      Begin VB.Frame Frame4 
         Caption         =   "Rendering option"
         Height          =   735
         Left            =   225
         TabIndex        =   15
         Top             =   1125
         Width           =   2895
         Begin VB.CheckBox cHigh 
            Caption         =   "&High Quality"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Window options"
         Height          =   1695
         Left            =   3345
         TabIndex        =   6
         Top             =   165
         Width           =   2895
         Begin VB.CheckBox cCThrough 
            Caption         =   "&Click through"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox cUnmove 
            Caption         =   "&Unmovable window"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox cOntop 
            Caption         =   "&Always on top"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Transparency"
         Height          =   735
         Left            =   225
         TabIndex        =   3
         Top             =   165
         Width           =   2895
         Begin ComctlLib.Slider sAlpha 
            Height          =   375
            Left            =   30
            TabIndex        =   4
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   327682
            Max             =   255
            SelStart        =   255
            TickStyle       =   3
            Value           =   255
         End
         Begin VB.Label lAlpha 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "255"
            Height          =   195
            Left            =   2520
            TabIndex        =   5
            Top             =   300
            Width           =   270
         End
      End
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3000
      X2              =   3000
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label lOptions 
      Alignment       =   2  'Center
      Caption         =   "Skin"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   195
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1560
      X2              =   1560
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label lOptions 
      Alignment       =   2  'Center
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   1215
   End
End
Attribute VB_Name = "fOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
  Unload Me
End Sub

Private Sub cCThrough_Click()
  CThrough = Abs(cCThrough)
End Sub

Private Sub cHigh_Click()
  rHighQuality = Abs(cHigh)
  LoadSkin
End Sub

Private Sub cOntop_Click()
  OnTop = Abs(cOntop)
End Sub

Private Sub cUnmove_Click()
  Unmovable = Abs(cUnmove)
End Sub

Private Sub Form_Load()
  InitLayeredForm Me
  
  Dim i%
  For i = 0 To pOptions.UBound
    pOptions(i).BorderStyle = 0
    pOptions(i).Move 120, 540
    pOptions(i).Line (0, 0)-(pOptions(i).ScaleWidth - 1, pOptions(i).ScaleHeight - 1), vb3DShadow, B
  Next
  cCThrough = Abs(CThrough)
  cOntop = Abs(OnTop)
  cUnmove = Abs(Unmovable)
  cHigh = Abs(rHighQuality)
  sAlpha.Value = Transparency
  File1.Path = App.Path & "\Skin"
  List1.AddItem "<Default>"
  For i = 0 To File1.ListCount
    If File1.List(i) <> "" Then List1.AddItem Left(File1.List(i), Len(File1.List(i)) - 4)
  Next
  List1.ListIndex = SendMessage(List1.hWnd, LB_FINDSTRING, -1, ByVal SkinName)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  FadeOutForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  LoadSetting
End Sub

Private Sub List1_Click()
  SkinName = List1.Text
  LoadSkin
End Sub

Private Sub lOptions_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lOptions(Index).ForeColor = vbBlue
End Sub

Private Sub lOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lOptions(Index).ForeColor = vbButtonText
  If X > 0 And X < lOptions(Index).Width And Y > 0 And Y < lOptions(Index).Height Then
    Dim i As Byte
    For i = 0 To pOptions.UBound
      pOptions(i).Visible = i = Index
      lOptions(i).FontBold = i = Index
    Next
  End If
End Sub

Private Sub OKButton_Click()
  SaveSetting
  Unload Me
End Sub

Private Sub sAlpha_Change()
  Transparency = sAlpha
  lAlpha = sAlpha
End Sub

Private Sub sAlpha_Scroll()
  sAlpha_Change
End Sub
