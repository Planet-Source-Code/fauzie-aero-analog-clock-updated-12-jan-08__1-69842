VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' *****
' This class came from another project. The remarks may or may not pertain to this project

' almost all of this class from http://vbnet.mvps.org/index.html?code/subclass/shellnotifystructinfo.htm
' added additional routines & modified a couple so they can be used generically
' and also to allow multiple icons per hWnd

' This class will allow you to modify system tray icons for your application

' Any hWnd can have a tray icon. All icons you assign to the tray icons are your
' responsibility for destroying when appropriate. Only top level windows should
' be installing tray icons. This is because should Explorer crash & then
' restart & new taskbar created, the message we look for to restore the icons
' on the new taskbar is only sent to top level windows.

' This class can be setup in your routines at any time and can be deactivated at
' any time, destroyed, or activated any time after the hWnd has begun subclassing.

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32.dll" (ByVal hwnd As Long, ByVal idAni As Long, ByRef lprcFrom As RECT, ByRef lprcTo As RECT) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type




Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, _
   ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long

'Private Type GUID
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(7) As Byte
'End Type
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128      'shell 5+  <> 64 chars max for shell <5
   dwState As Long            'shell 5+  <> not used in this class
   dwStateMask As Long        'shell 5+
   szInfo As String * 256     'shell 5+
   uTimeoutAndVersion As Long 'shell 5+
   szInfoTitle As String * 64 'shell 5+
   dwInfoFlags As Long        'shell 5+
   guidItem(15) As Byte       ' GUID  shell 6+  <> reserved by Windows
End Type

' tray update commands
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
      
' tray update masks
Private Const NIF_MESSAGE = &H1
Private Const NIF_INFO As Long = &H10
Private Const NIF_ICON = &H2
Private Const NIF_STATE As Long = &H8
Private Const NIF_TIP = &H4
Private Const NIS_HIDDEN = &H1

'balloon tip notification messages <> not used here; for ref only
'Private Const NIN_BALLOONSHOW = (WM_USER + 2)
'Private Const NIN_BALLOONHIDE = (WM_USER + 3)
'Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
'Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' tray version settings
Private Const NOTIFYICON_VERSION As Long = &H3
Private Const NIM_SETVERSION As Long = &H4
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Public Enum BalloonIcons
    icNone = &H0        'NIIF_NONE As Long = &H0
    icInfo = &H1        'NIIF_INFO As Long = &H1
    icWarning = &H2     'NIIF_WARNING As Long = &H2
    icError = &H3       'NIIF_ERROR As Long = &H3
    icTrayIcon = &H4    'NIIF_USER As Long = &H4 ???
    icNoSound = &H10    'NIIF_NOSOUND = &H10
    icNoChange = &H40   'class-custom value
End Enum

Private nid As NOTIFYICONDATA   ' cached UDT values

Private rcAni As RECT

Public Sub RemoveTrayIcon()
' remove/inactive the icon from system tray. Also can be called from
' lvSubclasser's ApplyTrayNotify function

' The information in this class is not destroyed; therefore you can easily
' toggle the tray icon by using the lvSubclasser's ApplyTrayNotify function.

    If IsActive Then
        nid.uFlags = 0
        Shell_NotifyIcon NIM_DELETE, nid
        nid.uCallbackMessage = 0 ' this flag determines if tray is active or not
    End If

End Sub


Public Sub InitializeTray(ByVal hIcon As Long, NewTip As String)
' This is a nice to have routine that fills in the icon & tip in one step
' vs calling those idividual properties. It does not active the tray icon.
' You must activate it from your lvSubclasser's ApplyTrayNotify function.

' Parameters
' hIcon is a handle to an existing icon you want to display in the tray
'       -- IMPORTANT: you are responsible for destroying the icon if needed
' Tip is displayed when mouse hovers icon. 128 max unless < Shell32.v5 which is 64

    Tip = NewTip
    Icon = hIcon
End Sub

Public Property Let Icon(ByVal hIcon As Long)
' Used to change the icon displayed on the system tray
' IMPORTANT: You are responsible for destroying icons if needed

    nid.hIcon = hIcon
    If IsActive Then
        nid.uFlags = NIF_ICON
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End Property
Public Property Get Icon() As Long
    Icon = nid.hIcon    ' returns handle to currently assigned icon
End Property

Public Property Let Tip(ByVal NewTip As String)
    ' Used to change the tip displayed when mouse hovers over your icon
    nid.szTip = NewTip & vbNullChar
    If IsActive Then
        nid.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End Property
Public Property Get Tip() As String
    Tip = Left$(nid.szTip, InStr(nid.szTip & Chr$(0), Chr$(0)) - 1)
End Property

' Read-Only properties return balloon title & message.
' Use ShowBalloon to set those values
Public Property Get BalloonTitle() As String
    BalloonTitle = Left$(nid.szInfoTitle, InStr(nid.szInfoTitle & Chr$(0), Chr$(0)) - 1)
End Property
Public Property Get BalloonMessage() As String
    BalloonMessage = Left$(nid.szInfo, InStr(nid.szInfo & Chr$(0), Chr$(0)) - 1)
End Property

Public Property Get isBalloonCapable() As Boolean
    ' property will let you know if you can use balloons on clients pc
    isBalloonCapable = (NOTIFYICONDATA_SIZE > NOTIFYICONDATA_V1_SIZE)
End Property

Public Sub ShowBalloon(Optional ByVal Message As String, _
    Optional ByVal Title As String, _
    Optional ByVal Icon As BalloonIcons = icNoChange)

' Routine shows a balloon tip only after this class has been assigned to your subclasser
' The class does not destroy your tip, title or icon settings; therefore you only need
' to pass the information that you want changed from the last time you called the tip

If NOTIFYICONDATA_SIZE = 0 Then GetTrayVersion

If NOTIFYICONDATA_SIZE > NOTIFYICONDATA_V1_SIZE Then
    ' V1 of Shell32.dll doesn't have balloon capability (pre W2K)
    
    With nid
        ' only make needed changes
        If Len(Message) Then .szInfo = Message & vbNullChar
        If Len(Title) Then .szInfoTitle = Title & vbNullChar
        If (Icon Or icNoChange) <> Icon Then
            If NOTIFYICONDATA_SIZE < NOTIFYICONDATA_V3_SIZE Then
                ' filter out non-XP settings if not right version of tray
                If (Icon Or icTrayIcon) = Icon Then Icon = Icon And Not icTrayIcon
                If (Icon Or icNoSound) = Icon Then Icon = Icon And Not icNoSound
            End If
            .dwInfoFlags = Icon
        End If
        .uFlags = NIF_INFO
    End With
    ' if the class has been added to subclasser, send the change
    If IsActive Then Shell_NotifyIcon NIM_MODIFY, nid
End If
    
End Sub

Public Sub HideBalloon()

' function destroys a balloon if visible

' if the class has been added to subclasser, send the change
If IsActive Then
    
    If NOTIFYICONDATA_SIZE > NOTIFYICONDATA_V1_SIZE Then
        Dim oldNid As NOTIFYICONDATA
        oldNid = nid
        With oldNid
             ' remove all tip info
            .szInfo = vbNullChar
            .szInfoTitle = vbNullChar
            .szTip = vbNullChar
            .uFlags = NIF_INFO Or NIF_TIP
        End With
        ' update
        Shell_NotifyIcon NIM_MODIFY, oldNid
        
        ' replace the original tip info
        nid.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End If
End Sub

Private Sub GetTrayVersion()
' unmodified from source

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         'IsShellVersion = nVerMajor >= version
         Select Case nVerMajor
         Case Is < 5
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE
        Case Is < 6
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE
        Case Else
            NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
        End Select
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Sub

Friend Sub BeginTrayNotifications(hwnd As Long, TrayIconID As Long, uMsg As Long)
' Only called by lvSubclasser

' since this class can be set up on demand, knowing when to activate it isn't
' possible without user dictating that fact.  This routine fills in the
' missing information needed to activate the icon

' determine which version is running on the OS
' calculates the NOTIFYICONDATA_SIZE value
If NOTIFYICONDATA_SIZE = 0 Then GetTrayVersion
    Dim lAction As Long
    If IsActive = True Then lAction = NIM_MODIFY Else lAction = NIM_ADD
    
    With nid
        .cbSize = NOTIFYICONDATA_SIZE
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uTimeoutAndVersion = NOTIFYICON_VERSION
        .uCallbackMessage = uMsg ' custom message
        .uID = TrayIconID
        .hwnd = hwnd
    End With
    Shell_NotifyIcon lAction, nid
            
    If lAction = NIM_ADD Then Shell_NotifyIcon NIM_SETVERSION, nid

' Note: The custom message above isn't really important to you.
' It will never be sent to your ProcessMessage routine.
' The DLL reroutes these messages to a separate event call ProcessTrayIcon.

' The value of the message, should you also be using custom messages is
' WM_USER + &H1962
End Sub

Public Sub Restore(uMsg As Long)
    ' function is called only from lvSubclasser after a new Explorer re-starts after crash
    ' not all icons may be active, reactivate only previously active ones
    If nid.hwnd <> 0 Then
        nid.uCallbackMessage = 0
        BeginTrayNotifications nid.hwnd, nid.uID, uMsg
    End If

End Sub

Private Sub Class_Terminate()
' remove the tray icon if activated
    RemoveTrayIcon
End Sub

Public Property Get IsActive() As Boolean
    ' returns whether or not tray icon is active.
    ' Active means it is sending your hWnd messages
    IsActive = nid.uCallbackMessage <> 0
End Property
