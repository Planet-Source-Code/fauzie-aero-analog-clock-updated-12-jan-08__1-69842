Attribute VB_Name = "mMain"
Option Explicit

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type Size
    cX As Long
    cY As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const LB_FINDSTRING = &H18F
Public Const WM_TIMER = &H113
Public Const WM_MOUSEMOVE = &H200
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_STYLE As Long = -16
Public Const GWL_EXSTYLE As Long = -20
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hDC As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hDC As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, Image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, BITMAP As Long) As GpStatus

Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
End Type


Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Public SkinName As String, cX As Long, cY As Long, Unmovable As Boolean, OnTop As Boolean, Transparency As Byte, rHighQuality As Boolean, CThrough As Boolean

Public GDIpToken As Long ' Needed to close GDI+

Public Function GetValue(IniFileName As String, Section As String, Key As String, Optional ByVal DefaultValue As String) As String
  On Error GoTo Hell
  Dim Value As String, retval As String, X As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(Section, Key, DefaultValue, retval, Len(retval), IniFileName)
  GetValue = Trim(Left(retval, X))
Exit Function
Hell:
  GetValue = DefaultValue
End Function

Public Function WriteValue(IniFileName As String, Section As String, Key As String, ByVal Value As String) As Boolean
  On Error GoTo Hell
  Dim X As Integer
  X = WritePrivateProfileString(Section, Key, Value, IniFileName)
  If X <> 0 Then WriteValue = True
  Exit Function
Hell:
End Function

Sub Main()
  ' Start up GDI+
  Dim GpInput As GdiplusStartupInput
  GpInput.GdiplusVersion = 1
  If GdiplusStartup(GDIpToken, GpInput) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical
    End
  End If
  
  With fClock
    .iBack.gdiToken = GDIpToken
    .iHour.gdiToken = GDIpToken
    .iMin.gdiToken = GDIpToken
    .iSec.gdiToken = GDIpToken
    
    LoadSetting
    .Show
    .Move cX * Screen.TwipsPerPixelX, cY * Screen.TwipsPerPixelY
  End With
End Sub

Public Sub LoadSetting()
  SkinName = GetValue(App.Path & "\Settings.ini", "Settings", "Skin", "<Default>")
  cX = GetValue(App.Path & "\Settings.ini", "Settings", "X", (Screen.Width - fClock.Width - 300) \ Screen.TwipsPerPixelX)
  cY = GetValue(App.Path & "\Settings.ini", "Settings", "Y", 300 \ Screen.TwipsPerPixelX)
  CThrough = Abs(GetValue(App.Path & "\Settings.ini", "Settings", "ClickThrough", 0))
  Unmovable = Abs(GetValue(App.Path & "\Settings.ini", "Settings", "Unmovable", 0))
  OnTop = Abs(GetValue(App.Path & "\Settings.ini", "Settings", "AlwaysOnTop", 1))
  Transparency = GetValue(App.Path & "\Settings.ini", "Settings", "Alpha", 255)
  rHighQuality = Abs(GetValue(App.Path & "\Settings.ini", "Settings", "HighQuality", 1))
  fMenu.mCThrough.Checked = CThrough
  fMenu.mUnmove.Checked = Unmovable
  fMenu.mOnTop.Checked = OnTop
  fMenu.mHigh.Checked = rHighQuality
  SetTopMost fClock.hWnd, OnTop
  If CThrough Then MakeWndTransparent fClock.hWnd, CThrough
  
  LoadSkin
End Sub

Public Sub LoadSkin()
  Dim sBack$, sHour$, sMin$, sSec$
  With fClock
    .iBack.DestroyDIB
    .iHour.DestroyDIB
    .iMin.DestroyDIB
    .iSec.DestroyDIB
    If LCase(SkinName) = "<default>" Then
      .iBack.LoadPicture_Resource "BACK", "PNG", , , , True
      .iHour.LoadPicture_Resource "HOUR", "PNG", , , , True
      .iMin.LoadPicture_Resource "MINUTE", "PNG", , , , True
      .iSec.LoadPicture_Resource "SECOND", "PNG", , , , True
    Else
      sBack = GetValue(App.Path & "\Skin\" & SkinName & ".ini", "Clock", "Back")
      sHour = GetValue(App.Path & "\Skin\" & SkinName & ".ini", "Clock", "Hour")
      sMin = GetValue(App.Path & "\Skin\" & SkinName & ".ini", "Clock", "Minute")
      sSec = GetValue(App.Path & "\Skin\" & SkinName & ".ini", "Clock", "Second")
      .iBack.LoadPicture_File App.Path & "\Skin\" & sBack, True
      .iHour.LoadPicture_File App.Path & "\Skin\" & sHour, True
      .iMin.LoadPicture_File App.Path & "\Skin\" & sMin, True
      .iSec.LoadPicture_File App.Path & "\Skin\" & sSec, True
      
'      If Not .iBack.Alpha Then .iBack.MakeTransparent .iBack.GetPixel(0, 0)
    End If
    .iBack.HighQualityInterpolation = rHighQuality
    .iHour.HighQualityInterpolation = rHighQuality
    .iMin.HighQualityInterpolation = rHighQuality
    .iSec.HighQualityInterpolation = rHighQuality
    .Width = .iBack.Width * Screen.TwipsPerPixelX
    .Height = .iBack.Height * Screen.TwipsPerPixelY
  End With
End Sub

Public Sub SaveSetting()
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "Skin", SkinName)
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "X", fClock.Left \ Screen.TwipsPerPixelX)
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "Y", fClock.Top \ Screen.TwipsPerPixelY)
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "ClickThrough", Abs(CThrough))
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "Unmovable", Abs(Unmovable))
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "AlwaysOnTop", Abs(OnTop))
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "Alpha", Transparency)
  Call WriteValue(App.Path & "\Settings.ini", "Settings", "HighQuality", Abs(rHighQuality))
End Sub

Public Sub FadeOutForm(theForm As Form)
  Dim i As Integer
  
  For i = 255 To 0 Step -5
    Call SetLayeredWindowAttributes(theForm.hWnd, 0, i, LWA_ALPHA)
  Next
End Sub

'Public Sub FadeInForm(theForm As Form)
'  Dim i As Integer
'  For i = 0 To 255 Step 5
'    Call SetLayeredWindowAttributes(theForm.hWnd, 0, i, LWA_ALPHA)
'  Next
'End Sub

Public Sub InitLayeredForm(theForm As Form)
  Call SetWindowLong(theForm.hWnd, GWL_EXSTYLE, GetWindowLong(theForm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(theForm.hWnd, 0, 255, LWA_ALPHA)
End Sub

Public Sub SetTopMost(hWnd As Long, TopMost As Boolean)
  If TopMost Then
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Else
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub

Public Function MakeWndTransparent(hWnd As Long, Value As Boolean) As Long
  Dim FLong As Long
  FLong = GetWindowLong(hWnd, GWL_EXSTYLE)
  If Value Then
    MakeWndTransparent = SetWindowLong(hWnd, GWL_EXSTYLE, FLong Or WS_EX_TRANSPARENT)
  Else
    MakeWndTransparent = SetWindowLong(hWnd, GWL_EXSTYLE, FLong And Not WS_EX_TRANSPARENT)
  End If
End Function
