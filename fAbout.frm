VERSION 5.00
Begin VB.Form fAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Aero Clock"
   ClientHeight    =   2175
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5295
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pClock 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   180
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   3
      Top             =   180
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "La Volpe for his 32bpp DIB Classes"
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Special Thanks to :"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Aero Clock  v1.2"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2130
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 2007 by Fauzie's Software"
      Height          =   195
      Left            =   2130
      TabIndex        =   2
      Top             =   540
      Width           =   2985
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  InitLayeredForm Me
  
  AttachMessage Me, hwnd, WM_TIMER
  SetTimer hwnd, 1, 1000, 0
  
  UpdateClock
End Sub

Private Sub UpdateClock()
  pClock.Cls
  
  With fClock
    .iBack.Render pClock.hDC, , , pClock.ScaleWidth, pClock.ScaleHeight
'    .iHour.Render pClock.hDC, , , pClock.ScaleWidth, pClock.ScaleHeight, , , , , , , , , ((Int(Timer) Mod 43200) / 360) * 3
'    .iMin.Render pClock.hDC, , , pClock.ScaleWidth, pClock.ScaleHeight, , , , , , , , , Minute(Time) * 6
'    .iSec.Render pClock.hDC, , , pClock.ScaleWidth, pClock.ScaleHeight, , , , , , , , , Second(Time) * 6
  End With
  
  pClock.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  FadeOutForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DetachMessage Me, hwnd, WM_TIMER
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
  'dont remove this
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  'dont remove this
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If iMsg = WM_TIMER Then
    UpdateClock
  End If
End Function
