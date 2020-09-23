VERSION 5.00
Begin VB.Form fClock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const HTCAPTION = 2

Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

Dim blendFunc32bpp As BLENDFUNCTION
Dim WinSize As Size
Dim SrcPoint As POINTAPI
Dim cTray As New cSysTray

Implements ISubclass

Public iBack As New c32bppDIB
Public iHour As New c32bppDIB, iMin As New c32bppDIB, iSec As New c32bppDIB

Private Sub Form_Load()
  SetIcon hWnd, "MAINICON"
  
  Dim curWinLong As Long
  curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
  SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
  
  cTray.Icon = LoadIcon(App.hInstance, "MAINICON")
  cTray.BeginTrayNotifications hWnd, hWnd, WM_MOUSEMOVE
  
  AttachMessage Me, hWnd, WM_TIMER
  SetTimer hWnd, 1, 1000, 0
  
  With blendFunc32bpp
    .AlphaFormat = AC_SRC_ALPHA
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER
  End With
  UpdateClock
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  If Button = 1 And Not Unmovable Then
    Call ReleaseCapture
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  ElseIf Button = 2 Then
    PopupMenu fMenu.mnu
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    PopupMenu fMenu.mnu
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim i As Integer
  DetachMessage Me, hWnd, WM_TIMER
  UpdateClock
  For i = 255 To 0 Step -2
    blendFunc32bpp.SourceConstantAlpha = i
    Call UpdateLayeredWindow(Me.hWnd, Me.hDC, ByVal 0&, WinSize, iBack.LoadDIBinDC(True), SrcPoint, 0, blendFunc32bpp, ULW_ALPHA)
  Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Cleanup everything
  Call GdiplusShutdown(GDIpToken)
  cTray.RemoveTrayIcon
  SaveSetting
End Sub

Private Function UpdateClock() As Boolean
  Dim Tim As Long
  Tim = Int(Timer)
  iBack.LoadPicture_FromOrignalFormat
  If Not iBack.Alpha Then iBack.MakeTransparent iBack.GetPixel(0, 0)
  iHour.Render 0, , , ScaleWidth, ScaleHeight, , , , , , , , iBack, ((Tim Mod 43200) / 360) * 3
  iMin.Render 0, , , ScaleWidth, ScaleHeight, , , , , , , , iBack, Minute(Time) * 6
  iSec.Render 0, , , ScaleWidth, ScaleHeight, , , , , , , , iBack, Second(Time) * 6

  ' Needed for updateLayeredWindow call
  SrcPoint.x = 0
  SrcPoint.y = 0
  WinSize.cx = Me.ScaleWidth
  WinSize.cy = Me.ScaleHeight
   
  blendFunc32bpp.SourceConstantAlpha = Transparency
   
  Call UpdateLayeredWindow(Me.hWnd, 0, ByVal 0&, WinSize, iBack.LoadDIBinDC(True), SrcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
  'dont remove this
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  'dont remove this
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If iMsg = WM_TIMER Then
    UpdateClock
  End If
End Function
