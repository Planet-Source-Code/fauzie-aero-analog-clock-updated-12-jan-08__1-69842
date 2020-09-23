VERSION 5.00
Begin VB.Form fMenu 
   Caption         =   "Form1"
   ClientHeight    =   435
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   1635
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleWidth      =   1635
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Begin VB.Menu mOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu n1 
         Caption         =   "-"
      End
      Begin VB.Menu mCThrough 
         Caption         =   "&Click through"
      End
      Begin VB.Menu mUnmove 
         Caption         =   "&Unmovable"
      End
      Begin VB.Menu mOnTop 
         Caption         =   "Always on &top"
      End
      Begin VB.Menu n2 
         Caption         =   "-"
      End
      Begin VB.Menu mHigh 
         Caption         =   "&High Quality"
      End
      Begin VB.Menu n3 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "fMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mAbout_Click()
  fAbout.Show vbModal, fClock
End Sub

Private Sub mClose_Click()
  Unload fClock
  Unload Me
End Sub

Private Sub mCThrough_Click()
  mCThrough.Checked = Not mCThrough.Checked
  CThrough = mCThrough.Checked
  MakeWndTransparent fClock.hWnd, CThrough
End Sub

Private Sub mHigh_Click()
  mHigh.Checked = Not mHigh.Checked
  rHighQuality = mHigh.Checked
  LoadSkin
End Sub

Private Sub mOnTop_Click()
  mOnTop.Checked = Not mOnTop.Checked
  OnTop = mOnTop.Checked
  SetTopMost fClock.hWnd, OnTop
End Sub

Private Sub mOptions_Click()
  fOptions.Show vbModal, fClock
End Sub

Private Sub mUnmove_Click()
  mUnmove.Checked = Not mUnmove.Checked
  Unmovable = mUnmove.Checked
End Sub
