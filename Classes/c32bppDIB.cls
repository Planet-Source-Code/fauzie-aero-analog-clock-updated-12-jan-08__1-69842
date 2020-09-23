VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c32bppDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Credits/Acknowledgements - Thanx goes to:
'   Paul Caton for his class on calling non VB-Friendly DLLs that use _cdecl calling convention
'       Used when calling non VB-friendly zLIB dll versions
'   Alfred Koppold for his PNG, VB-only, decompression routines.
'       Used when zLib & GDI+ not available
'   Carles P.V for his pvResize logic
'       Used when manually scaling images with NearestNeighbor or BiLinear interpolation
'   www.zlib.net for their free zLIB.dll, the standard DLL for compressing/decompressing PNGs
'       Without it, we'd be limited to GDI+ for creating PNGs
'   coders like you that provide constructive criticism to make this class better & more all-inclusive
'       Without your comments, this project probably would have died several versions/updates ago
' For most current updates/enhancements visit the following:
'   Visit http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
' To see a usercontrol applying a version of this class
'   AlphaImage Control. http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=68262&lngWId=1
'   Image List Control. http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=69621&lngWId=1

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                    O V E R V I E W
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' AlphaBlend API :: (msimg32.dll) , GDI+ API (gdiplus.dll)

' About 32bpp pre-multiplied RGB (pARGB) bitmaps, if you are not aware.
'   - These are used specifically for the AlphaBlend API & are GDI+ compatible
'   Advantages:
'       - Images can be per-pixel alpha blended
'       - Opacity can be simultaneously adjusted during rendering
'       - AlphaBlend does both BitBlt & StretchBlt for pARGB images.
'       - Speed: AlphaBlend & GDI+ are pretty quick APIs vs manual blending
'   Disadvantages:
'       - The original RGB values are permanently destroyed during pre-multiplying
'           -- Premultiplied formula: preMultipliedRed=(OriginalRed * Alpha) \ 255
'           -- There is no way to convert pARGB back to non-premultiplied RGB values
'              The formula would be: reconstructedRed=(preMultipliedRed * 255) \ Alpha.
'               but because of integer division when pre-multiplying, the result is very
'               close and if this should be premultiplied again & converted again, the
'               calculated colors can eventually be completely different than the original.
'               Fully opaque pixels & fully transparent pixels are not affected.
'           ** Note: When images are converted to PNG formats, removal of
'              premultiplication is performed to meet PNG specs.
'       - Displaying a pre-multiplied bitmap without AlphaBlend/GDI+ will not result in
'           the image being displayed as expected.
'       - Not ideal for saving due to its size: SizeOf= W x H x 4
'           -- better to save source image instead or compress the DIB bytes using favorite compression utility
'           -- with GDI+ or zLib, image can be converted to PNG for storage
'       - AlphaBlend API is not included/compatible with Win95, NT4 and lower
'       - AlphaBlend on Win9x systems can be buggy, especially when rendering to DIBs vs DDBs
' Note that GDI+ is standard on WinXP+, and can be used on Win98,ME,2K & on NT4 if SP6 is installed
' Download GDI+ from:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdicpp/GDIPlus/GDIPlus.asp
' Download ZLib from: http://www.zlib.net

' ----------------------------------------------
' About Win95, Win98, WinME, NT3.5 & NT4 support
' ----------------------------------------------
' These routines may not honor AlphaBlend if it exists on those systems. Win98's version,
' for example, has several bugs that can crash the application when AlphaBlending to DIBs.
' NT4, NT3.5 & Win95 do not come with AlphaBlend and I do not have WinME to test with.
' Therefore, to support these older systems, the Render routine will alphablend manually
' regardless if the AlhpaBlend API exists on the older system or not. However, this can
' be overridden by you. See isAlphaBlendFriendly routine. Therefore, AlphaBlend is only
' reliable on Win2K and above. XP & above already have GDI+


' Class Purpose:
' ----------------------------------------------
' This class holds the 32bpp image. It also marshals any new image thru
' the battery of parsers to determine best method for converting the image
' to a 32bpp alpha-compatible image. It handles rendering, rotating, scaling,
' mirroring of DIBs using manual processes, AlphaBlend, and/or GDI+.

' The parser order is very important for fastest/best results...
' cPNGparser :: will convert PNG, all bit depths; aborts quickly if not PNG
' cGIFparser :: will convert non-transparent/transparent GIFs; aborts quickly
' cICOpraser :: will convert XP-Alpha, paletted, true color, & Vista PNG icons
'               -- can also convert most non-animated cursors
' cBMPparser :: will convert bitmaps, wmf/emf & jpgs

' As a last resort, when GDI+ exists, anything unable to be processed by the
' parsers (i.e., TIFFs) are sent to GDI+. If GDI+ can process the image, then
' the image will be converted, internally, to PNG to enable additional processing.

' The parsers are efficient. Most image formats have a magic number that give
'   a hint to what type of image the file/stream is. However, checks need to
'   be employed because non-image files could feasibly have those same magic
'   numbers. If the image is determined not to be one the parser is designed
'   to handle, the parser rejects it and the next parser takes over.  The
'   icon parser is slightly different because PNG files can be included into
'   a Vista ico file. When this occurs, the icon parser will pass off the
'   PNG format to the PNG parser automatically.
' And last but not least, the parsers have no advanced knowledge of the image
' format; as far as they are concerned, anything passed is just a byte array

' Class Organization:
' ----------------------------------------------
' Search the class for the words NEW SECTION
' The class routines are organized in the following sections:
'   Class Initialization & Termination Routines
'   Public Properties & Methods (almost 60 and growing)
'       Public Read-Only Properties
'       Public Methods
'       Class to Class Communication Methods
'   Local Support Functions
' ----------------------------------------------

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                       CHANGE HISTORY
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Accompanying FAQ.rtf is updated with every change
' Last changed: 18 Nov 07. See change history within the FAQ file
' 26 Dec 06: First version
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take liberties, changing parameter types, in several APIs throughout

' Used to determine operating system
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Type OSVERSIONINFOEX
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128 ' up to here is OSVERSIONINFO vs EX
   wServicePackMajor As Integer ' 8 bytes larger than OSVERSIONINFO
   wServicePackMinor As Integer
   wSuiteMask As Integer
   wProductType As Byte
   wReserved As Byte
End Type

' APIs used to manage the 32bpp DIB
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef Pointer As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const STRETCH_HALFTONE As Long = &H4&
Private Const OBJ_BITMAP As Long = &H7&
Private Const OBJ_METAFILE As Long = &H9&
Private Const OBJ_ENHMETAFILE As Long = &HD&

' APIs used to create files
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1&

' ////////////////////////////////////////////////////////////////
' Unicode-capable Drag and Drop of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, _
    ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, _
    ByVal paCNT As Long, ByRef paTypes As Integer, _
    ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (lpString As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

' ////////////////////////////////////////////////////////////////
' Unicode-capable Pasting of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Type FORMATETC
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lIndex As Long
    TYMED As Long
End Type
Private Type DROPFILES
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type
Private Type STGMEDIUM
    TYMED As Long
    Data As Long
    pUnkForRelease As IUnknown
End Type
' ////////////////////////////////////////////////////////////////


' used to create the checkerboard pattern on demand
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' used when saving an image or part of the image
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 3) As Long ' reusable UDT for 1 & 2 dim arrays
    ' array items: 0=1D item count, 1=1D LBound, 2=2D item count, 3=2D LBound
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
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
    bmiPalette As Long
End Type

Private Enum eOScapability
    osAlphaBlendUsable = 1  ' then AlphaBlend enabled & used when needed
    osGDIplusUsable = 2     ' then GDI+ enabled & used when needed (set in isGDIplusEnabled)
   'oszLIBusable = 4        ' then zLib enabled & can be used to create/read PNGs (no longer used, tested as needed)
    osWin2KorBetter = 8     ' AlphaBlend capable system else it isn't for these classes. See isAlphaBlendFriendly property for more info
    osWin98MEonly = 16      ' Win98 or WinME. When m_OScap includes osWin98MEonly+osAlphaBlendUsable then user overrode isAlphaBlendFriendly
    osGDIplusNotAvail = 32  ' then NT4 w/less than SP6 or Win95. Otherwise system is GDI+ capable else it isn't
    osIsNT = 64             ' NT-based system. Unicode compatible
End Enum

Public Enum eImageFormat    ' source image format
    imgError = -1  ' no DIB has been initialized
    imgNone = 0    ' no image loaded
    imgBitmap = 1  ' standard bitmap or jpg
    imgIcon = 3    ' standard icon
    imgWMF = 2     ' windows meta file
    imgEMF = 4     ' enhanced WMF
    imgCursor = 5  ' standard cursor
    imgBmpARGB = 6  ' 32bpp bitmap where RGB is not pre-multiplied
    imgBmpPARGB = 7 ' 32bpp bitmap where RGB is pre-multiplied
    imgIconARGB = 8 ' XP-type icon; 32bpp ARGB
    imgGIF = 9      ' gif; if class.Alpha=True, then transparent GIF
    imgPNG = 10     ' PNG image
    imgPNGicon = 11 ' PNG in icon file (Vista)
    imgCursorARGB = 12 ' alpha blended cursors? do they exist yet?
    imgCheckerBoard = 64 ' image is displaying own checkerboard pattern; no true image
End Enum

Public Enum eConstants      ' See SourceIconSizes
    HIGH_COLOR = &HFFFF00
    TRUE_COLOR = &HFF000000
    TRUE_COLOR_ALPHA = &HFFFFFFFF
End Enum

Private m_Tag As Variant            ' user-defined, user usage only. See TAG property
Private m_StretchQuality As Boolean ' if true will use BiLinear or better interpolation
Private m_Handle As Long        ' handle to 32bpp DIB
Private m_Pointer As Long       ' pointer to DIB bits
Private m_Height As Long        ' height of DIB
Private m_Width As Long         ' width of DIB
Private m_hDC As Long           ' DC if self-managing one
Private m_prevObj As Long       ' object deselected from DC when needed
Private m_osCAP As eOScapability ' See eOScapability enumeration above
Private m_Format As eImageFormat ' type of source image
Private m_ManageDC As Boolean   ' does class manage its own DC
Private m_AlphaImage As Boolean ' does the DIB contain alpha/transparency
Private m_GDItoken As Long
Private m_ImageByteCache() As Byte  ' should you want the DIB class to cache original bytes
' ^^ N/A if image is loaded by handle, stdPicture, or resource

Private m_GDIplus As cGDIPlus   ' maintains instance to speed up drawing
Private m_KeepGDIplusActive As Boolean

' NEW SECTION *******************************************************************************
'                   CLASS INITIALIZATION & TERMINATION ROUTINES
' *******************************************************************************************


Private Sub Class_Initialize()

    ' Determine operating system for compatibility of 32bpp images
    ' http://vbnet.mvps.org/code/helpers/iswinversion.htm
    ' http://msdn2.microsoft.com/en-gb/library/ms724834.aspx
    
   Dim osType As OSVERSIONINFOEX
   Const VER_PLATFORM_WIN32_WINDOWS As Long = 1

   ' Retrieve version data for OS.
   osType.dwOSVersionInfoSize = Len(osType)
   If GetVersionEx(osType) = 0 Then
      ' The OSVERSIONINFOEX structure is only supported
      ' in NT4/SP6+ and NT5.x, so we're likely running
      ' on an earlier version of Windows. Revert structure
      ' size to OSVERSIONINFO and try again.
      osType.dwOSVersionInfoSize = Len(osType) - 8
      Call GetVersionEx(osType)
   End If
   
    If osType.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then

        If osType.dwMinorVersion = 0 Then ' Win95; can't use AlphaBlend nor GDI+
            m_osCAP = osGDIplusNotAvail

        Else ' flag as Alphablend disabled, but capable & is Win98/ME
            m_osCAP = osWin98MEonly
        End If

    Else

        If osType.dwMajorVersion > 4 Then ' if Win2K or better
            m_osCAP = osAlphaBlendUsable Or osWin2KorBetter ' flag as AlphaBlend enabled (Win2K or better) and capable

        Else ' WinNT4. If SP6 or better than GDI+ capable else not. Regardless, not AlphaBlend capable
            If osType.wServicePackMajor < 6 Then m_osCAP = osGDIplusNotAvail
        End If
        m_osCAP = m_osCAP Or osIsNT

    End If

    Me.isGDIplusEnabled = True  ' attempt to start GDI+, test system capability, add osGDIplusUsable to m_OScap
    If Me.isGDIplusEnabled Then Me.HighQualityInterpolation = True
    
    m_Format = imgError

End Sub

Private Sub Class_Terminate()
    DestroyDIB ' simply clean up
End Sub

' NEW SECTION *******************************************************************************
'                           PUBLIC PROPERTIES AND METHODS
' *******************************************************************************************


Public Property Let Alpha(isAlpha As Boolean)
    m_AlphaImage = isAlpha  ' determines the flags used for AlphaBlend API
    ' this flag is set by the various image parsers; setting it yourself
    ' can produce less than desirable effects.
    ' Used in Me.Render & Me.TrimImage, cPNGwriter.OptimizeTrueColor & cPNGwriter.PalettizeImage
End Property
Public Property Get Alpha() As Boolean
    Alpha = m_AlphaImage
End Property

Public Property Let HighQualityInterpolation(Value As Boolean)
    ' When possible GDI+ will be used for stretching & rotation.
    ' If GDI+ is used,then high quality equates to BiCubic interpolation
    ' If not used, then BiLinear (manual processing) will be used.
    ' If High Quality is false, then Nearest Neighbor (very fast, less quality) interpolation used
    m_StretchQuality = Value
End Property
Public Property Get HighQualityInterpolation() As Boolean
    HighQualityInterpolation = m_StretchQuality
End Property

Public Property Get ImageType() As eImageFormat
    ImageType = m_Format    ' returns image format of the source image
End Property
Friend Property Let ImageType(iType As eImageFormat)
    m_Format = iType    ' set by the various image parsers. This is not used
    ' anywhere in these classes, you can do with it what you want -- for now.
End Property

Public Property Let ManageOwnDC(bManage As Boolean)
    ' Determines whether or not this class will manage its own DC
    ' If false, then a DC is created each time the image needs to be Rendered
    Dim tDC As Long
    If bManage = False Then     ' removing management of DC
        If Not m_hDC = 0& Then   ' DC does exist, destroy it
            ' first remove the dib, if one exists
            If Not m_Handle = 0& Then SelectObject m_hDC, m_prevObj
            m_prevObj = 0&
        End If
        DeleteDC m_hDC
        m_hDC = 0&
    Else                        ' allowing creation of dc
        If m_hDC = 0& Then        ' create DC only if we have a dib to put in it
            If Not m_Handle = 0& Then
                tDC = GetDC(0&)
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
        End If
    End If
    m_ManageDC = bManage
End Property
Public Property Get ManageOwnDC() As Boolean
    ManageOwnDC = m_ManageDC
End Property

Public Property Get isAlphaBlendFriendly() As Boolean
    isAlphaBlendFriendly = ((m_osCAP And osAlphaBlendUsable) = osAlphaBlendUsable)
    ' WinNT4 & below and Win95 are not shipped with msimg32.dll (AlphaBlend API)
    ' Win98 has bugs & would believe that WinME is buggy too but don't know for sure
    ' Therefore, the Rendering in this class will not use AlphaBlend on these
    ' operating systems even if the DLL exists, but will use GDI+ if available
    ' Can be overridden by setting this property to True
End Property
Public Property Let isAlphaBlendFriendly(Enabled As Boolean)
    ' This has been provided to override safety of using AlphaBlend on Win9x systems.
    ' Caution. Only set this when rendering to a known device dependent bitmap (DDB)
    ' Alphablend can crash when rendering DIB to DIB vs DIB to DDB. Be warned.
    If Enabled = True Then
        ' Overriding in play: allow AlphaBlend if system is Win98 or better
        ' By default this is already set for Win2K or better
        If ((m_osCAP And osWin2KorBetter) = osWin2KorBetter) Then
            m_osCAP = m_osCAP Or osAlphaBlendUsable
        ElseIf ((m_osCAP And osWin98MEonly) = osWin98MEonly) Then
            m_osCAP = m_osCAP Or osAlphaBlendUsable
        End If
    Else
        m_osCAP = m_osCAP And Not osAlphaBlendUsable ' disallow AlphaBlend
    End If
End Property
Public Property Get isGDIplusEnabled() As Boolean
    ' identifies if GDI+ is usable on the system.
    ' Before this property is set, GDI+ is tested to ensure it is usable
    isGDIplusEnabled = ((m_osCAP And osGDIplusUsable) = osGDIplusUsable)
End Property
Public Property Let isGDIplusEnabled(Enabled As Boolean)
    ' Sets the property. If set to False by you, GDI+ will not be used
    ' for any rendering, but still may be used to create PNG files if needed
    
    ' You can reset it to true at any time. If the system won't support
    ' GDI+, then the True setting will simply be ignored -- no harm, no foul
    ' To test success:  c32class.isGDIplusEnabled=True: If c32class.isGDIplusEnabled=True Then ' success
    
    ' Note. My GDI+ dll works on Win95. Therefore, the O/S capabilities check below
    ' has been removed to create a hard test vs simply assuming
    
    If Not Enabled = Me.isGDIplusEnabled Then
        m_osCAP = (m_osCAP And Not osGDIplusUsable)
        If Enabled Then
'            If (m_osCAP And osGDIplusNotAvail) = 0 Then ' else Win95, NT4 SP5 or lower
                Dim cGDIp As cGDIPlus
                If m_GDIplus Is Nothing Then
                    Set cGDIp = New cGDIPlus
                Else
                    Set cGDIp = m_GDIplus
                End If
                If cGDIp.isGDIplusOk() = True Then m_osCAP = m_osCAP Or osGDIplusUsable
'            End If
        End If
    End If
End Property

Public Property Let gdiToken(Token As Long)
    ' Everytime a GDI+ API function is called, the class calls GDI+ apis to
    ' create a GDI+ token first then destroys the token after the function is called.
    
    ' This occurs quite often. However, you can create your own token by calling
    ' GdiplusStartup and then passing the token to each class for the class to use.
    ' You would call GdiplusShutdown during your main form's Terminate event to
    ' release the Token.
    
    ' When Token is zero, the classes will revert to creating a token on demand.
    ' When the Token is not zero, any other DIB class created by this class will
    ' pass the token as needed. The only routine that can create a new instance and
    ' returns that new instance is the CreateDropShadow method. You must set the
    ' token for that class at some point for that dropshadow class to use it.
    
    m_GDItoken = Token
    
End Property
Public Property Get gdiToken() As Long
    ' returns the GDI+ token if one was created
    gdiToken = m_GDItoken
End Property

Public Property Let Tag(vValue As Variant)
    If IsObject(vValue) Then
        Set m_Tag = vValue
    Else
        m_Tag = vValue
    End If
End Property
Public Property Set Tag(vValue As Variant)
    Me.Tag = vValue ' use the object check in the Let property
End Property
Public Property Get Tag() As Variant
    If IsObject(m_Tag) Then
        Set Tag = m_Tag
    Else
        Tag = m_Tag
    End If
End Property


' NEW SECTION *******************************************************************************
'                               PUBLIC READ-ONLY PROPERTIES
' *******************************************************************************************

Public Property Get Width() As Long
    Width = m_Width     ' width of image in pixels
End Property
Public Property Get Height() As Long
    Height = m_Height   ' height of image in pixels
End Property
Public Property Get BitsPointer() As Long
    BitsPointer = m_Pointer ' pointer to the bits of the image
End Property
Public Property Get scanWidth() As Long
    scanWidth = m_Width * 4&    ' number of bytes per scan line
End Property
Public Property Get Handle() As Long
    Handle = m_Handle   ' the picture handle of the image
End Property
Public Property Get isZlibEnabled() As Boolean
    ' Read Only
    ' To create PNG files, GDI+ or zLib is required. This property informs
    ' you if zLIB exists in the system's DLL path
    isZlibEnabled = iparseValidateZLIB(vbNullString, 0, False, False, True)
    
End Property

' This setting will keep the current cGDI+ class active which
' prevents destroying the GDI+ hImage and Token objects.  This
' can speed up rendering. This option SHOULD NOT be set when the
' class is not compiled. Else crashes will happen when
' user hits END or VB toolbar's STOP button while in IDE
Public Property Get KeepGDIplusActive() As Boolean
    KeepGDIplusActive = m_KeepGDIplusActive
End Property
Public Property Let KeepGDIplusActive(keepActive As Boolean)
    m_KeepGDIplusActive = keepActive
    If keepActive = False Then
        Set m_GDIplus = Nothing
    ElseIf Me.isGDIplusEnabled = False Then
        ' can't set to True if GDI+ not installed
        m_KeepGDIplusActive = False
    End If
End Property


' NEW SECTION *******************************************************************************
'                                       PUBLIC METHODS
' *******************************************************************************************

Public Function LoadPicture_File(ByVal FileName As String, _
                                Optional ByVal SaveFormat As Boolean) As Boolean

    ' PURPOSE: Convert passed image file into a 32bpp image
    
    ' Parameters.
    ' FileName :: full path of file. Validation occurs before we continue
    ' SaveFormat :: if true, then the image will be cached as a byte array only
    '   if the image was successfully loaded. Call GetOrginalFormat to retrieve them.
    
    ' Why would you want to save the bytes? If this is being used in a usercontrol,
    ' saving the bytes will almost always be less size than saving the 32bit DIB.
    ' Additionally, these classes have the ability to get different sizes from
    ' the original source (i.e., WMF, icon, cursors) if available, but if the
    ' 32bit DIB is saved, it is a constant size. The potential of different sizes
    ' could allow better resizing of the image vs stretching the DIB.

    On Error Resume Next
    Dim hFile As Long
    
    hFile = iparseGetFileHandle(FileName, True, ((m_osCAP And osIsNT) = osIsNT))
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    
    If GetFileSize(hFile, 0&) > 56 Then
        
        ' no image file/stream can be less than 57 bytes and still be an image
        Dim aDIB() As Byte  ' dummy array
        LoadPicture_File = spt_LoadPictureEx(hFile, FileName, aDIB(), 0&, 0&, SaveFormat)
    
    End If
    CloseHandle hFile
    
End Function

Public Function LoadPicture_Stream(inStream() As Byte, _
                                    Optional ByVal streamStart As Long = 0&, _
                                    Optional ByVal streamLength As Long = 0&, _
                                    Optional ByVal SaveFormat As Boolean) As Boolean
    
    ' PURPOSE: Convert passed array into a 32bpp image
    
    ' Parameters.
    ' inStream:: byte stream containing the image. Validation occurs below
    ' streamStart :: array position of 1st byte of the image file. Validated.
    ' streamLength :: total length of the image file. Validated.
    ' SaveFormat :: if true, then the image will be cached as a byte array only
    '   if the image was successfully loaded. Call GetOrginalFormat to retrieve them.
    
    ' Why would you want to save the bytes? If this is being used in a usercontrol,
    ' saving the bytes will almost always be less size than saving the 32bit DIB.
    ' Additionally, these classes have the ability to get different sizes from
    ' the original source (i.e., WMF, icon, cursors) if available, but if the
    ' 32bit DIB is saved, it is a constant size. The potential of different sizes
    ' could allow better resizing of the image vs stretching the DIB.
    
    If iparseIsArrayEmpty(VarPtrArray(inStream)) = 0& Then Exit Function
    If streamStart < LBound(inStream) Then streamStart = LBound(inStream)
    If streamLength = 0& Then streamLength = UBound(inStream) - streamStart + 1&
    If streamLength < 57 Then Exit Function
    ' no image file/stream can be less than 57 bytes and still be an image
    LoadPicture_Stream = spt_LoadPictureEx(0&, vbNullString, inStream, streamStart, streamLength, SaveFormat)

End Function

Public Function LoadPicture_Resource(ByVal ResIndex As Variant, ByVal ResSection As Variant, _
                            Optional VBglobal As IUnknown, _
                            Optional ByVal streamStart As Long = 0&, _
                            Optional ByVal streamLength As Long = 0&, _
                            Optional ByVal SaveFormat As Boolean) As Boolean

    ' PURPOSE: Convert passed resource into a 32bpp image
    
    ' Parameters.
    ' ResIndex :: the resource file index (i.e., 101)
    ' ResSection :: one of the VB LoadResConstants or String value of
    '       your resource section, i.e., vbResBitmap, vbResIcon, "Custom", etc
    ' VbGlobal :: pass as VB.GLOBAL of the project containing the resource file
    '       - Allows class to be mobile; can exist in DLL or OCX
    '       - if not provided, class will use resource from existing workspace
    '       - For example, if this class was in a compiled OCX, then the only way
    '           to use the host's resource file is passing the host's VB.Global reference
    ' streamStart :: array position of 1st byte of the image file. Validated.
    ' streamLength :: total length of the image file. Validated.
    '   -- See LoadPicture_Stream for the validation
    
    ' Tips:
    ' 1) Store 32bpp bitmaps in the "Custom" resource always. Storing in the
    '       Bitmap resource can change color depth of the image created by VB
    '       depending on your screen settings
    ' 2) Icons, normal bitmaps, & cursors are generally stored in their own sections
    '       However, with icons containing multiple formats, VB will extract the
    '       closest format to 32x32. May want to consider storing these in "Custom"
    ' 3) All other types of images are normally stored in the "Custom" section

    On Error GoTo ExitRoutine
    
    Dim oWorkSpace As VB.Global, tPic As StdPicture
    
    If VBglobal Is Nothing Then
        Set oWorkSpace = VB.Global
    ElseIf TypeOf VBglobal Is VB.Global Then
        Set oWorkSpace = VBglobal
    Else
        Set oWorkSpace = VB.Global
    End If
    
    If VarType(ResSection) = vbString Then
        Dim inStream() As Byte
        ' could be anything, PNG,icon,gif,32bpp bitmap,wmf, etc
        inStream = oWorkSpace.LoadResData(ResIndex, ResSection)
        LoadPicture_Resource = LoadPicture_Stream(inStream, streamStart, streamLength, SaveFormat)
    End If
    LoadPicture_Resource = Not (m_Handle = 0&)
    
ExitRoutine:
    If Err Then Err.Clear
End Function

Public Function LoadPicture_FromOrignalFormat() As Boolean

    ' PURPOSE: Reload the current image from the cached bytes (if any)
    ' If the original bytes were not cached when the image was loaded, then no action
    ' will be taken.  See LoadPicture_File & LoadPicture_Stream
    
    Dim tBytes() As Byte
    tBytes() = m_ImageByteCache() ' copy bytes; original are destroyed when DIB is recreated
    LoadPicture_FromOrignalFormat = Me.LoadPicture_Stream(tBytes, , , True)
    
End Function

Public Function GetPixel(ByVal X As Long, ByVal Y As Long, Optional ByRef AlphaValue As Long, _
                            Optional ByRef asPreMultiplied As Boolean) As Long

    ' Function will return the pixel color value and alpha value from the DIB
    ' Note that the DIB is always referenced top down within this function
    
    ' X is the left coordinate of the pixel to be returned, image always starts at 0,0
    ' Y is the top coordinate of the pixel to be returned, image always starts at 0,0
    ' AlphaValue will contain the alpha value of the pixel.
    ' asPreMultiplied. If false, then premultiplication is removed else it isn't
    
    ' Return value is the RGB color value of the pixel.
    ' If return value is -1 then the X,Y coordinates passed are invalid
    
    ' It is far more efficient to use GetDIBits or overlaying your own array when more than
    ' one pixel is required to be returned; in other words, recommend not using this
    ' function within a loop
    
    AlphaValue = 0&
    If X < 0& Or X > m_Width - 1& Then
        GetPixel = -1&
    ElseIf Y < 0& Or Y > m_Height - 1& Then
        GetPixel = -1&
    Else
        Dim pOffset As Long, pColor As Long
        ' calculate the location of the X,Y coordinate in relation to a bottom-up DIB
        pOffset = iparseSafeOffset(m_Pointer, X * 4& + ((m_Height - Y - 1&) * m_Width * 4&))
        
        ' get the alpha value
        CopyMemory AlphaValue, ByVal iparseSafeOffset(pOffset, 3&), 1&
        
        ' get the pixel color & convert it to RGB
        CopyMemory pColor, ByVal pOffset, 3&
        If asPreMultiplied = True Or (AlphaValue Mod 255) = 0 Then
            GetPixel = ((pColor And &HFF) * &H10000) Or ((pColor \ &H100) And &HFF) * &H100 Or ((pColor \ &H10000) And &HFF)
        Else    ' remove premultiplication
            pOffset = ((255& * (pColor And &HFF)) \ AlphaValue) * &H10000
            pOffset = pOffset Or ((255& * ((pColor \ &H100) And &HFF)) \ AlphaValue) * &H100
            GetPixel = pOffset Or ((255& * ((pColor \ &H10000) And &HFF)) \ AlphaValue)
        End If
    End If

End Function

Public Sub CopyImageTo(cDIBclass As c32bppDIB, Optional ByVal newWidth As Long, _
            Optional ByVal newHeight As Long, Optional ByVal CopyOriginalFormat As Boolean = False)
    
    ' Function replicates the the current image to another DIB class and optionally resizes it
    
    ' NewWidth is optional. if zero, will use the source DIB width. If negative will mirror & resize if needed
    ' NewHeight is optional. if zero, will use the source DIB height. If negative will mirror & resize if needed
    ' If CopyOriginalFormat = True then, and only, if class loaded its image
    '   with the optional SaveFormat=True, then the original image bytes
    '   were cached and will be copied to the target cDIBclass also
    '   See LoadPicture_File & LoadPicture_Stream for more info
    
    Dim dDC As Long, aResized() As Byte
    Dim bUnselect As Boolean, bResetAlphaCap As Boolean
    
    If Not m_Handle = 0& Then                ' do we have an image to copy?
    
        If newWidth = 0& Then newWidth = m_Width
        If newHeight = 0& Then newHeight = m_Height
        
        If cDIBclass Is Nothing Then
            Set cDIBclass = New c32bppDIB  ' was a valid ref passed?
            cDIBclass.gdiToken = m_GDItoken
            cDIBclass.isGDIplusEnabled = Me.isGDIplusEnabled
            cDIBclass.HighQualityInterpolation = Me.HighQualityInterpolation
            cDIBclass.InitializeDIB Abs(newWidth), Abs(newHeight) ' Create new one
        Else
            cDIBclass.gdiToken = m_GDItoken
            If Not (Abs(newWidth) = cDIBclass.Width And Abs(newHeight) = cDIBclass.Height) Then
                cDIBclass.InitializeDIB Abs(newWidth), Abs(newHeight) ' Create new one
            End If
        End If
        cDIBclass.Alpha = m_AlphaImage       ' carry over the alpha flag
        cDIBclass.ImageType = m_Format       ' and image type flag
            
        If newWidth = m_Width And newHeight = m_Height Then
            ' can copy using CopyMemory vs AlphaBlend
            CopyMemory ByVal cDIBclass.BitsPointer, ByVal m_Pointer, newWidth * 4& * newHeight
        Else
            
            If Not (m_osCAP And (osWin98MEonly Or osAlphaBlendUsable)) = 0 Then ' system is Win98/ME with AlphaBlend capability overridden
                ' but we will be resizing DIB to DIB so disallow it for now
                m_osCAP = (m_osCAP And Not osAlphaBlendUsable)
                bResetAlphaCap = True
            End If
                
            bUnselect = (m_prevObj = 0&)
            If Me.isGDIplusEnabled And (m_StretchQuality = True Or Me.isAlphaBlendFriendly = False) Then ' use GDI+ to resize
                Dim cGDIp As cGDIPlus
                If m_GDIplus Is Nothing Then
                    Set cGDIp = New cGDIPlus
                Else
                    Set cGDIp = m_GDIplus
                End If
                dDC = cDIBclass.LoadDIBinDC(True)
                If bUnselect Then Me.LoadDIBinDC True
                cGDIp.RenderGDIplus Me, dDC, 0&, 100&, 0&, 0&, newWidth, newHeight, 0&, 0&, m_Width, m_Height, True, m_GDItoken
                cDIBclass.LoadDIBinDC False
                If bUnselect Then Me.LoadDIBinDC False
        
            ElseIf Me.isAlphaBlendFriendly And m_StretchQuality = False Then ' O/S has no alphablending shortfalls that are known
                dDC = cDIBclass.LoadDIBinDC(True)   ' load target into a DC
                If bUnselect Then Me.LoadDIBinDC True
                Me.Render dDC, 0&, 0&, newWidth, newHeight, 0&, 0&, m_Width, m_Height, , , False, cDIBclass
                cDIBclass.LoadDIBinDC False         ' remove DIB from DC
                If bUnselect Then Me.LoadDIBinDC False
            Else
                ' stretching is involved, resize
                Call spt_pvResize(0&, aResized(), aResized(), cDIBclass) ' routine resizes directly to DIB bytes
            End If
            
            If bResetAlphaCap Then m_osCAP = m_osCAP Or osAlphaBlendUsable
        
        End If
        ' if the original image bytes are to be copied, do them too
        If CopyOriginalFormat = True Then Call cDIBclass.SetOriginalFormat(m_ImageByteCache)
    
    End If
    
End Sub

Public Function GetOrginalFormat(outStream() As Byte) As Boolean

    ' If SaveFormat is true when LoadPicture_Stream or LoadPicture_File was
    ' called, the original bytes were cached when the image was successfully
    ' loaded. Call this to return those original bytes

    ' If there are no original bytes, the function returns False & outStream is uninitialized
    
    outStream() = m_ImageByteCache()
    GetOrginalFormat = Not (iparseIsArrayEmpty(VarPtrArray(m_ImageByteCache)) = 0&)

End Function

Public Function LoadDIBinDC(ByVal bLoad As Boolean) As Long

    ' Purpose: Select/Unselect the DIB into a DC.
    ' Returns the DC handle when image is loaded
    ' Called by image parser if it needs to paint the image into the DIB
       
    If bLoad = True Then
        Dim tDC As Long
        If Not m_Handle = 0& Then    ' do we have an image?
            If m_hDC = 0& Then        ' do we have a DC?
                tDC = GetDC(0&)     ' if not create one
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
            If m_prevObj = 0& Then
                m_prevObj = SelectObject(m_hDC, m_Handle)
            End If
            LoadDIBinDC = m_hDC
        End If
    Else
        If Not m_prevObj = 0& Then
            SelectObject m_hDC, m_prevObj
            If m_ManageDC = False Then
                DeleteObject m_hDC
                m_hDC = 0&
            End If
            m_prevObj = 0&
        End If
    End If
End Function


Public Function InitializeDIB(ByVal Width As Long, ByVal Height As Long) As Boolean

    ' Creates a blank (all black, all transparent) DIB of requested height & width
    
    Dim tBMPI As BITMAPINFO, tDC As Long
    
    DestroyDIB ' clear any pre-existing dib
    
    If Width < 0& Then Exit Function
    If Height = 0& Then
        Exit Function
    ElseIf Height < 0& Then
        Height = Abs(Height) ' no top-down dibs
    End If
    
    On Error Resume Next
    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSize = 40&
        .biSizeImage = .biHeight * .biWidth * 4&
    End With
    If Err Then
        Err.Clear
        ' only possible error would be that Width*Height*4& is absolutely huge
        Exit Function
    End If
    
    tDC = GetDC(0&) ' get screen DC
    m_Handle = CreateDIBSection(tDC, tBMPI, 0&, m_Pointer, 0&, 0&)
    If m_ManageDC = True Then
        ' create a DC if class is managing its own & one isn't created yet
        If m_hDC = 0& Then m_hDC = CreateCompatibleDC(tDC)
    End If
    ' release the screen DC if we captured it
    ReleaseDC 0&, tDC
    
    If Not m_Handle = 0& Then    ' let's hope system resources allowed DIB creation
        m_Width = Width
        m_Height = Height
        ' on my Win95 system, I have found cases where the creation of the DIB
        ' resulted in Windows filling the DIB with garbage, not zeroes as expected
        EraseDIB                ' ensure dib starts out blank
        InitializeDIB = True
    End If

End Function

Public Sub DestroyDIB()
    
    ' PURPOSE: Destroy any existing image
    Set m_GDIplus = Nothing
    If Not m_hDC = 0& Then   ' do we have a DC?
        ' do we have an image; if so get it out of the DC
        If Not m_prevObj = 0& Then SelectObject m_hDC, m_prevObj
        ' destroy our DC, no point in keeping it w/o image
        DeleteObject m_hDC
        m_hDC = 0&
    End If
    ' if we do have an image, destroy it now
    If Not m_Handle = 0& Then
        DeleteObject m_Handle
        Erase m_ImageByteCache
    End If
    ' reset other image attributes
    m_Width = 0&
    m_Height = 0&
    m_Handle = 0&
    m_Pointer = 0&
    m_prevObj = 0&
    m_AlphaImage = False
    m_Format = imgError
End Sub

Public Sub EraseDIB()
    ' Function clears out an existing DIB, making it 100% transparent/black
    If Not m_Handle = 0& Then
        FillMemory ByVal m_Pointer, m_Width * m_Height * 4&, 0
        m_Format = imgNone
        m_AlphaImage = True
    End If
    
End Sub

Public Function Render(ByVal destinationDC As Long, _
                Optional ByVal destX As Long, Optional ByVal destY As Long, _
                Optional ByVal destWidth As Long, Optional ByVal destHeight As Long, _
                Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                Optional ByVal Opacity As Long = 100&, _
                Optional ByVal Blend As Boolean = True, _
                Optional ByVal SetHalfTone As Boolean = True, _
                Optional ByRef destHostDIB As c32bppDIB = Nothing, _
                Optional ByVal Angle As Single = 0!, _
                Optional ByVal CenterOnDestXY As Boolean = False) As Boolean

    ' PURPOSE: Render an existing 32bpp DIB to a target DC
    ' Angle & CenterOnDestXY added to allow this routine to replace the RotateAtCenterPoint & RotateAtTopLeft calls
    
    ' Mirroring: When destWidth & srcWidth are compared, if one is negative and the other positive, mirroring horizontally occurs
    '            When destHeight & srcHeight are compared, if one is negative and the other positive, mirroring vertically occurs
    ' All four of those parameters are optional & any that are not passed will default to the image's width/height as needed
    ' before the mirroring check is applied
    
    ' Parameters. Only destinationDC is required
    ' destinationDC :: target DC to draw to. Ignored if destHostDIB is passed
    ' destX, destY :: the top/left coordinates to draw to, default is 0,0
    ' destWidth, destHeight :: the width and height to draw to, default is the image's width & height
    ' srcX, srcY :: the left & top offset within the DIB
    ' srcWidth, srcHeight :: the amount of DIB to be rendered
    ' Opacity :: how opaque to draw the image, default is 100% opaque
    ' Blend :: no longer used, reserved & left in for backward compatibility
    ' SetHalfTone :: if True, then the destination DC's stretch mode will be modified to
    '       produce better quality results. This option is not available on Win9x systems.
    '       Tip: When AlphaBlending to another DIB set to False
    '            When AlphaBlending to CompatibleBitmap (DDB) or visible DC set to True
    ' destHostDIB :: When rendering from DIB class to DIB class, pass the destination
    '       DIB class to ensure alpha blending occurs correctly on systems that do not
    '       support GDI+ or AlphaBlend APIs. When passed, destinationDC is ignored
    ' Angle :: between -360 and 360. Rotation is clockwise
    ' CenterOnDestXY :: If true then rendering is centered on the destX,destY coordinates
    
    Dim lBlendFunc As Long, tDC As Long, hOldImage As Long
    Dim lStretchMode As Long
    Dim aResizedBytes() As Byte, aMirrorBytes() As Byte
    Dim bStretching As Boolean
    Dim bMirroring As Boolean
    Dim bRotating As Boolean
    Dim bCanUseAlphaBlend As Boolean
    
    Const AC_SRC_OVER = &H0&
    Const AC_SRC_ALPHA = &H1&
    
    ' validate a few things
    If Opacity < 1& Then        ' nothing to render if image is 100% transparent
        Render = Not (m_Handle = 0)
        Exit Function
    ElseIf m_Handle = 0& Then
        Exit Function
    ElseIf destinationDC = 0& Then
        If destHostDIB Is Nothing Then Exit Function
    End If
    
    ' validate optional destination parameters
        If destWidth = 0& Then destWidth = m_Width
        If destHeight = 0& Then destHeight = m_Height
    
    ' validate optional parameters for source image
        If SrcX < 0& Then SrcX = 0&  ' source X,Y cannot be negative
        If SrcY < 0& Then SrcY = 0&  ' but the dest X,Y can be
        If srcWidth = 0& Then
            srcWidth = m_Width
        ElseIf srcWidth < 0& Then        ' locally, only the destWidth can be negative for mirroring, not srcHeight
            destWidth = -destWidth
            srcWidth = -srcWidth
        End If
        If srcHeight = 0& Then
            srcHeight = m_Height
        ElseIf srcHeight < 0& Then      ' locally, only the destHeight can be negative for mirroring, not srcHeight
            destHeight = -destHeight
            srcHeight = -srcHeight
        End If
        If SrcX + srcWidth > m_Width Then srcWidth = m_Width - SrcX
        If SrcY + srcHeight > m_Height Then srcHeight = m_Height - SrcY
        
    
    ' angle is Single. See if passed angle is evenly divisible by 360
    If Int(Angle) = Angle Then
        bRotating = Not ((Angle Mod 360) = 0)
    Else    ' angle has fractional component; therefore can't be Mod 360
        bRotating = True
    End If
    
    ' validate opacity is within range
    Opacity = Abs(Opacity) Mod 100
    If Opacity = 0& Then Opacity = 100&
    
    If CenterOnDestXY = True Then
        If bRotating Then
            destX = destX - Abs(destWidth) \ 2       ' use negative destWidth,destHeight if mirroring - rotation routines expect it
            destY = destY - Abs(destHeight) \ 2
            If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
            Render = spt_RotateImage(destinationDC, Angle, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, Opacity, destHostDIB)
            If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
            Exit Function
        Else
            destX = destX - Abs(destWidth \ 2)  ' not rotating, don't allow negative destWidth,destHeight for calculations
            destY = destY - Abs(destHeight \ 2)
        End If
    ElseIf bRotating Then
        If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
        Render = spt_RotateImage(destinationDC, Angle, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, Opacity, destHostDIB)
        If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
        Exit Function
    End If
    
    ' are we mirroring? rules out AlphaBlend usage if we are
    If destWidth > 0& Then
        bMirroring = (destHeight < 0&)
    Else
        bMirroring = True
    End If
    ' are we resizing? may rule out AlphaBlend usage (stretching DIB to DIB can crash on Win9x)
    If (Abs(destWidth) = srcWidth) Then
        bStretching = Not (Abs(destHeight) = srcHeight)
    Else
        bStretching = True
    End If
    
    ' The following IF tree and above boolean assignments are to determine:
    ' 1. Do we use AlphaBlend
    ' 2. Do we use GDI+
    ' 3. Do we do it manually.
    ' The answer depends on user settings, O/S, graphics manipulation & DLL abilities
    If m_KeepGDIplusActive = False Then
        ' if user provided token, allow GDI+ to override AlphaBlend
        If Me.isGDIplusEnabled = False Or Me.gdiToken = 0& Then
            ' see if alphablend will support the various rendering options
            If Me.isAlphaBlendFriendly Then     ' Win98 or better with AlphaBlend enabled & GDI+ disabled
                If Not bMirroring Then                  ' can't use alphaBlend if mirroring
                    If bStretching Then
                        bCanUseAlphaBlend = Not m_StretchQuality ' can't use alphaBlend if using Bilinear interpolation
                    Else
                        bCanUseAlphaBlend = True
                    End If
                End If
            End If
        End If
    End If
    ' note. If AlphaBlend cannot support the rendering options, we will want to use GDI+
    
    If Me.isGDIplusEnabled = True And bCanUseAlphaBlend = False Then
        ' we will use GDI+ to render when system is not AlphaBlend friendly, user supplied gdi+ token, or AlphaBlend cannot handle rendering options
        Dim cGDIp As cGDIPlus
        If m_GDIplus Is Nothing Then
            Set cGDIp = New cGDIPlus
            If m_KeepGDIplusActive Then Set m_GDIplus = cGDIp
        Else
            Set cGDIp = m_GDIplus
        End If
        If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
        Render = cGDIp.RenderGDIplus(Me, destinationDC, 0&, Opacity, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, m_StretchQuality, m_GDItoken)
        If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
    
    Else
        
        
        If m_hDC = 0& Then  ' do we have a DC to select our image into?
            tDC = GetDC(0&) ' if not create one, if ManageOwnDC=True, we will have one
            m_hDC = CreateCompatibleDC(tDC)
            ReleaseDC 0&, tDC
            hOldImage = SelectObject(m_hDC, m_Handle)
        Else
            ' we have a DC, but is the image selected into it?
            If m_prevObj = 0& Then hOldImage = SelectObject(m_hDC, m_Handle)
        End If
        
        If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
        
        If SetHalfTone Then ' Stretch_Halftone not compatible with win9x
            If ((m_osCAP And osIsNT) = osIsNT) Then lStretchMode = SetStretchBltMode(destinationDC, STRETCH_HALFTONE)
        End If
        
        ' calculate the opacity required & add it to the BlendFunction variable
        lBlendFunc = AC_SRC_OVER Or (((255& * Opacity) \ 100&) * &H10000)
        ' if the image has transparency, then we add the AC_SRC_ALPHA flag too
        If Me.Alpha = True Then lBlendFunc = lBlendFunc Or (AC_SRC_ALPHA * &H1000000)
        Render = Not (AlphaBlend(destinationDC, destX, destY, destWidth, destHeight, m_hDC, SrcX, SrcY, srcWidth, srcHeight, lBlendFunc) = 0&)
        
        If SetHalfTone Then ' Stretch_Halftone not compatible with win9x
            If ((m_osCAP And osIsNT) = osIsNT) Then SetStretchBltMode destinationDC, lStretchMode
        End If
        If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
        
        ' remove the image from the DC if necessary
        If Not hOldImage = 0& Then SelectObject m_hDC, hOldImage
        If Not tDC = 0& Then    ' if we created a DC, let's destroy it now
            DeleteDC m_hDC
            m_hDC = 0&
        End If
        
    End If
    
End Function

Public Function Resize(ByVal newWidth As Long, ByVal newHeight As Long) As Boolean

    ' Function simply resizes the current DIB to passed dimensions of Width & Height
    ' If newWidth is negative, the image will be horizontally mirrored
    ' If newHeight is negative, the image will be vertically mirrored
    
    If m_Handle = 0& Then Exit Function
    
    If newWidth = 0& Then newWidth = m_Width
    If newHeight = 0& Then newHeight = m_Height
    
    If newWidth = m_Width And newHeight = m_Height Then
        Resize = True                   ' no change
        Exit Function
    End If

    Dim tHost As New c32bppDIB
    Dim tDC As Long, sDC As Long
    Dim aDummy() As Byte
    Dim bUnselect As Boolean, bResetAlphaCap As Boolean
    
    With tHost
        .gdiToken = m_GDItoken
        .InitializeDIB Abs(newWidth), Abs(newHeight)   ' create new DIB
        .isGDIplusEnabled = Me.isGDIplusEnabled
        .HighQualityInterpolation = Me.HighQualityInterpolation
        .ImageType = m_Format          ' transfer format & alpha format
        .Alpha = m_AlphaImage
        .SetOriginalFormat m_ImageByteCache
        .Tag = Me.Tag
    End With
    
    bUnselect = (m_prevObj = 0&)
    If m_StretchQuality And Me.isGDIplusEnabled Then ' use GDI+ for resizing
        Dim cGDIp As cGDIPlus
        If m_GDIplus Is Nothing Then
            Set cGDIp = New cGDIPlus
        Else
            Set cGDIp = m_GDIplus
        End If
        tDC = tHost.LoadDIBinDC(True)
        If bUnselect Then Me.LoadDIBinDC True
        cGDIp.RenderGDIplus Me, tDC, 0&, 100&, 0&, 0&, newWidth, newHeight, 0&, 0&, m_Width, m_Height, True, m_GDItoken
        tHost.LoadDIBinDC False
        If bUnselect Then Me.LoadDIBinDC False
    
    ElseIf Me.isAlphaBlendFriendly And m_StretchQuality = False Then
    
        If Not (m_osCAP And (osWin98MEonly Or osAlphaBlendUsable)) = 0& Then ' system is Win98/ME with AlphaBlend capability overridden
            ' but we will be resizing DIB to DIB so disallow it for now
            m_osCAP = (m_osCAP And Not osAlphaBlendUsable)
            bResetAlphaCap = True
        End If
    
        tDC = tHost.LoadDIBinDC(True)               ' select into DC
        If bUnselect Then Me.LoadDIBinDC True
        Me.Render tDC, 0&, 0&, newWidth, newHeight, , , , , , , False, tHost ' stretch AlphaBlend
        tHost.LoadDIBinDC False                     ' remove from DC
        If bUnselect Then Me.LoadDIBinDC False
        If bResetAlphaCap Then m_osCAP = m_osCAP Or osAlphaBlendUsable
        
    Else
        spt_pvResize 0&, aDummy(), aDummy(), tHost   ' use custom resizing routine for nonAlphaBlend-friendly systems
    End If
    
    tHost.CopyImageTo Me, , , True    ' copy it back to us
    
    Set tHost = Nothing                 ' done
    Resize = True

End Function

Public Function MakeTransparent(ByVal TransparentColor As Long, Optional ByVal Revert As Boolean = False) As Boolean

    ' if Revert = False
    '   Function will convert all pixels that are of the TransparentColor to fully transparent.
    '   Additionally, only if the alpha value of the color is fully opaque will the pixel become transparent.
    ' if Revert = True
    '   All fully transparent colors are made fully opaque and changed to the TransparentColor
    
    If m_Handle = 0& Then Exit Function
    
    Dim tSA As SafeArray, dPixels() As Long, bPixels() As Byte
    Dim X As Long, Y As Long, bAlpha As Boolean
    
    ' convert passed color from RGB to BGRA
    TransparentColor = ((TransparentColor And &HFF) * &H10000) Or ((TransparentColor \ &H100) And &HFF) * &H100 _
                    Or ((TransparentColor \ &H10000) And &HFF) Or &HFF000000
    
    iparseOverlayHost_Long dPixels, VarPtr(tSA), 2, m_Height, m_Width, m_Pointer  ' overlay DMA array
    If Revert Then
        ' change all fully transparent pixels to passed color
        For Y = 0& To m_Height - 1&
            For X = 0& To m_Width - 1&
                If dPixels(X, Y) = 0& Then
                    dPixels(X, Y) = TransparentColor
                End If
            Next
        Next
        iparseOverlayHost_Long dPixels, 0, 0, 0, 0, 0  ' remove DMA overlay
        ' image may or may not contain alpha any longer, validate it
        ' The validation routine expects a 2D byte array, not long, so...
        iparseOverlayHost_Byte bPixels, VarPtr(tSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay DMA array
        iparseValidateAlphaChannel bPixels(), False, bAlpha, 0&
        iparseOverlayHost_Byte bPixels, 0, 0, 0, 0, 0  ' remove DMA overlay
        Me.Alpha = bAlpha
        If bAlpha = False Then
            If m_Format = imgBmpPARGB Or m_Format = imgBmpARGB Then m_Format = imgBitmap
        End If
        
    Else
        ' change all fully opaque colors matching TransparentColor to fully transparent
        For Y = 0& To m_Height - 1&
            For X = 0& To m_Width - 1&
                If dPixels(X, Y) = TransparentColor Then
                    dPixels(X, Y) = 0&
                    bAlpha = True
                End If
            Next
        Next
        iparseOverlayHost_Long dPixels, 0, 0, 0, 0, 0  ' remove DMA overlay
        If bAlpha Then
            m_AlphaImage = True
            If m_Format = imgBitmap Then m_Format = imgBmpPARGB
        End If
    End If
    MakeTransparent = True

End Function

Public Function CreateGDIplusToken() As Long
    ' function used to create a GDI+ token for you.
    ' Creating a token does not assign it to the class. Call gdiToken property
    Dim cGDI As cGDIPlus
    Dim gToken As Long
    If m_GDIplus Is Nothing Then
        Set cGDI = New cGDIPlus
    Else
        Set cGDI = m_GDIplus
    End If
    If cGDI.isGDIplusOk(gToken, True) = True Then
        CreateGDIplusToken = gToken
    End If
End Function

Public Function DestroyGDIplusToken(ByVal Token As Long) As Boolean
    ' function will release GDI+ for you. If the passed Token is used
    ' by other classes, you must set the other class' gdiToken property to zero
    ' before calling this routine.
    If Not Token = 0 Then
        Dim cGDI As cGDIPlus
        If m_GDIplus Is Nothing Then
            Set cGDI = New cGDIPlus
        Else
            Set cGDI = m_GDIplus
        End If
        DestroyGDIplusToken = cGDI.InitializeGDIplus(Token, True)
    End If
End Function

' NEW SECTION *******************************************************************************
'                   CLASS TO CLASS COMMUNICATION METHODS
' *******************************************************************************************

Friend Sub SetOriginalFormat(inStream() As Byte)

    ' Purpose: Pass the original image file/bytes to this DIB from another DIB
    ' This is only called by the CopyImageTo function. Note it is Friend vs Public
    
    m_ImageByteCache() = inStream()

End Sub


' NEW SECTION *******************************************************************************
'                       LOCAL SUPPORT FUNCTIONS FOR THIS CLASS ONLY
' *******************************************************************************************


Private Function spt_RotateImage(ByVal hDC As Long, ByVal Angle As Single, _
                            ByVal TopX As Long, ByVal TopY As Long, _
                            ByVal destWidth As Long, ByVal destHeight As Long, _
                            ByVal SrcX As Long, ByVal SrcY As Long, _
                            ByVal srcWidth As Long, ByVal srcHeight As Long, _
                            ByVal Opacity As Long, _
                            ByRef destHostDIB As c32bppDIB) As Boolean

    ' Internal function will rotate an image by passed Angle and render to the passed hDC.
    ' This function simultaneously rotates, scales and then blends.
    
    ' Note: Me.HighQualityInterpolation property setting is used to determine quality of rotation/scaling
    ' Called only by the Render method
    
    
    ' first see if we can do this via GDI+
    If Me.isGDIplusEnabled Then
        Dim cGDIp As cGDIPlus
        If m_GDIplus Is Nothing Then
            Set cGDIp = New cGDIPlus
        Else
            Set cGDIp = m_GDIplus
        End If
        If cGDIp.RenderGDIplus(Me, hDC, Angle, Opacity, TopX, TopY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, m_StretchQuality, m_GDItoken) = True Then
            spt_RotateImage = True
            Exit Function
        End If
    End If
        
    Dim cosTx As Double, sinTx As Double
    Dim cosTy As Double, sinTy As Double
    Dim scalerX As Double, scalerY As Double
    
    Dim maxX As Long, maxY As Long, maxSize As Long
    Dim ctrX As Long, ctrY As Long
    Dim xOffset As Double, yOffset As Double
    Dim targetX As Double, targetY As Double
    
    Dim dSA As SafeArray, sSA As SafeArray
    Dim dBytes() As Byte, sBytes() As Byte
    Dim lRow As Long, lCol As Long
    
    ' following variables are used for the BiLinear interpolation only
    Dim tgtY As Long, tgtX As Long
    Dim srcPixel As Long, dstPixel As Long, srcRow As Long
    Dim edgeOffsetX As Long, edgeOffsetY As Long
    Dim fY As Double, fX As Double, iX As Long, iY As Long
    Dim R As Double, G As Double, B As Double, a As Double
    
    ' determine the scale to use based off the passed
    ' source and destination widths,heights
    scalerX = destWidth / srcWidth      ' scale x coordinates
    scalerY = destHeight / srcHeight    ' scale y coordinates
    
    ' convert angle to radians & calculate scaled COS/SIN of the angle
    ' Multiplying by Negative so we rotate clockwise
    sinTx = -((Angle Mod 360) * (4& * Atn(1))) / 180  ' convert Degree to Radian
    cosTy = Cos(sinTx) / scalerY      ' get cosine of angle (Y coordinates)
    sinTy = Sin(sinTx) / scalerY      ' get sine of angle (Y coordinates)
    
    cosTx = Cos(sinTx) / scalerX      ' get cosine of angle (X coordinates)
    sinTx = Sin(sinTx) / scalerX      ' get sine of angle  (X coordinates)
    ' determine maximum size image we will need to cover any angle
    maxSize = Sqr(destWidth * destWidth + destHeight * destHeight)
    
    On Error GoTo eh
    ' create a temporary array to hold the rotated image
    ReDim dBytes(0 To maxSize * 4& - 1, 0 To maxSize - 1)
    
    ' overlay the temp DIB and our host DIB
    If iparseIsArrayEmpty(VarPtrArray(sBytes)) = 0& Then
        iparseOverlayHost_Byte sBytes, VarPtr(sSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay DMA array
    End If
    
    ' bottom up dib, vertical offset is from bottom, not top
    SrcY = UBound(sBytes, 2) - srcHeight - SrcY + 1&
    ' set up offsets for calculating rotated/scaled points
    maxX = srcWidth + SrcX      ' the right edge of the source image
    maxY = srcHeight + SrcY     ' the bottom edge of the source image
    ' determine where the center of the source selected bounds falls within the maxSize bounds
    ctrX = srcWidth \ 2 + SrcX  ' the center of the source image
    ctrY = srcHeight \ 2 + SrcY
    ' calculate offsets to "center" destination image in its maxSize window
    xOffset = ((maxSize - srcWidth) \ 2) - SrcX
    yOffset = ((maxSize - srcHeight) \ 2) - SrcY
    
    ' here's where we apply all of the above offsets
    ' This is a bit complicated because we allow any angle rotation,
    ' and also allowing portions of the image or entire image to be rotated
    ' and also allows scaling up or down
    If m_StretchQuality = False Then
    
        For lRow = -yOffset To maxSize - yOffset - 1&
            
            ' Calculate the point in the source image needed for the rotated point in destination image
            ' This only needs to be done once per image scan line & contains many math executions
            targetX = (-xOffset - ctrX) * cosTx + (lRow - ctrY) * sinTx + ctrX
            targetY = (lRow - ctrY) * cosTy - (-xOffset - ctrX) * sinTy + ctrY
        
            For lCol = -xOffset To maxSize - xOffset - 1&
            
                ' validate rotated point is within bounds of the image/portion
                If targetY >= SrcY Then                    ' is Y within area of source?
                    If targetY < maxY Then
                        If targetX >= SrcX Then            ' is X within area of source?
                            If targetX < maxX Then
                                ' validation complete, copy pixel to destination
                                CopyMemory dBytes((xOffset + lCol) * 4&, yOffset + lRow), sBytes(Int(targetX) * 4&, Int(targetY)), 4&
                            End If
                        End If
                    End If
                End If
        
                ' Excellent optimization I found (wish I remember where so I can give credit)
                ' But the logic is simple: once the initial X,Y coordinates for the
                ' current source row is found, the next point is always a constant value from
                ' the last point. In this case, increments of cosT & sinT.
                ' Therefore we don't need to recalculate targetX,targetY for every point
                ' since we did it once for current row. Thus we have 2 simple additions per pixel
                ' vs 4 multiplications & 12 additions per pixel
                targetX = targetX + cosTx
                targetY = targetY - sinTy
            Next
        Next
        
    Else
        ' BiLinear interpolation with rotation. This can produce better quality
        ' results but takes significantly (x4) longer. Recommend using this option
        ' when you need a static rotated image, but when rotating via a scrollbar
        ' or some other method where scrolling is expected to be repeated often,
        ' then use the non-BiLinear method. COMPILED IS MUCH FASTER !!!
        
        ' Up to 4 source pixels (16 bytes) are blended for each destination pixel (4 bytes)
        srcWidth = maxX - 1& ' reuse variable & subtract now vs subtraction for every pixel in the image
        For lRow = -yOffset To maxSize - yOffset - 1&
            
            ' Calculate the rotated point in relation to host image
            ' These calcs only needs to be done once per image scan line
            targetX = (-xOffset - ctrX) * cosTx + (lRow - ctrY) * sinTx + ctrX
            targetY = (lRow - ctrY) * cosTy - (-xOffset - ctrX) * sinTy + ctrY
        
            For lCol = -xOffset To maxSize - xOffset - 1&
                If targetY >= SrcY Then                    ' is Y within area of source?
                    If targetY < maxY Then
                        If targetX >= SrcX Then            ' is X within area of source?
                            If targetX < maxX Then
                            
                                tgtY = Int(targetY)     ' whole number of the double
                                If tgtY = 0& Then       ' for top down images check for last row vs 0
                                    ' last row of source image, will use only this row
                                    edgeOffsetY = 0&
                                    fY = 0#
                                Else
                                    ' will use this row & next row for blending
                                    edgeOffsetY = 1&
                                    fY = 1# - (targetY - tgtY) ' for top down images, use: fY = targetY-tgtY
                                    ' ^ Y coordinate fraction; pct of next vertical pixel that is used
                                End If
                                               
                                R = 0#: G = 0#: B = 0#: a = 0#
                                
                                tgtX = Int(targetX)  ' coordinate rounded down to whole number
                                If tgtX = srcWidth Then
                                    ' at far edge of source image, will use only this pixel for blending
                                    edgeOffsetX = 0&
                                    fY = 0#: fX = 0#
                                Else
                                    ' will use this pixel and next pixel for blending
                                    edgeOffsetX = 1&
                                    fX = targetX - tgtX
                                    ' ^ X coordinate fraction, pct of next horizontal pixel that is used
                                End If
                                
                                For iY = 0& To edgeOffsetY
                                    scalerY = Abs(iY - fY)   ' percentage of current row's pixel to blend
                                    
                                    If Not scalerY = 1& Then  ' else zero
                                        srcRow = tgtY - iY    ' for top down images Add iY vs subtracting
                                        For iX = 0& To edgeOffsetX
                                            scalerX = Abs(fX - iX)   ' percentage of current column's pixel to blend
                                            
                                            If Not scalerX = 1& Then  ' else zero
                                                scalerX = (1# - scalerX) * (1# - scalerY) ' combine percentages
                                                srcPixel = (tgtX + iX) * 4&
                                                B = B + sBytes(srcPixel, srcRow) * scalerX
                                                G = G + sBytes(srcPixel + 1&, srcRow) * scalerX
                                                R = R + sBytes(srcPixel + 2&, srcRow) * scalerX
                                                a = a + sBytes(srcPixel + 3&, srcRow) * scalerX
                                            End If
                                            
                                        Next
                                    End If
                    
                                Next
                                dstPixel = (xOffset + lCol) * 4&
                                iY = lRow + yOffset
                                dBytes(dstPixel, iY) = Int(B)
                                dBytes(dstPixel + 1&, iY) = Int(G)
                                dBytes(dstPixel + 2&, iY) = Int(R)
                                dBytes(dstPixel + 3&, iY) = Int(a)
                            End If
                        End If
                    End If
                End If
            targetX = targetX + cosTx
            targetY = targetY - sinTy
            Next
        Next
    End If
    
    If sSA.pvData = 0& Then ' image was also mirrored
        Erase sBytes()
    Else                    ' remove overlay of non-mirrored image
        iparseOverlayHost_Byte sBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    End If
    
    xOffset = (TopX + destWidth \ 2) - (maxSize \ 2)
    yOffset = (TopY + destHeight \ 2) - (maxSize \ 2)
    
eh:
    If Err Then
        Stop
        Err.Clear ' troubleshooting only, should be removed before compiling to final app
        Resume
    End If
End Function

Private Function spt_LoadPictureEx(ByVal FileHandle As Long, FileName As String, aStream() As Byte, _
                            streamOffset As Long, streamLength As Long, _
                            SaveFormat As Boolean, Optional bitDepth As Long = 32) As Boolean
      
    ' PURPOSE: Marshal passed file/array to image classes for conversion to 32bpp image
    ' For parameter information, see LoadPicture_File & LoadPicture_Stream
    
    Me.DestroyDIB
    
    ' various image parsers, in order of precedence
    ' All 4 recognize transparency
    Dim cPNG As cPNGparser  ' very fast to abort if not a PNG file
    Dim cBMP As cBMPparser  ' catchall. Handles bitmaps, wmf, emf & jpgs
    Dim cGDI As cGDIPlus
    
    Dim bReturn As Boolean  ' function return value
    Dim rtnRead As Long
    
    If bitDepth < 0& Then
        bitDepth = 32
    ElseIf bitDepth > 32 Then
        bitDepth = 32
    End If
    
    Set cPNG = New cPNGparser   ' see if image is a PNG; aborts quickly if not
    If FileHandle = 0 Then
        bReturn = cPNG.LoadStream(aStream(), Me, streamOffset, streamLength, m_GDItoken)
    Else     ' note: processing from file is slightly faster than via array
        bReturn = cPNG.LoadFile(FileHandle, FileName, Me, m_GDItoken)
        'If bReturn = True Then Close #fileNum         ' close the file
    End If
    If Err Then MsgBox Err.Description
    Set cPNG = Nothing
    If Not bReturn Then
        If Not FileHandle = 0& Then
            streamOffset = 0&
            streamLength = GetFileSize(FileHandle, 0&)
            'streamLength = LOF(fileNum) ' cache length of file
            ReDim aStream(streamOffset To streamLength - 1&)
            'Get #fileNum, , aStream()   ' populate our stream with the file contents
            'Close #fileNum
            SetFilePointer FileHandle, 0&, 0&, 0&
            ReadFile FileHandle, aStream(streamOffset), streamLength, rtnRead, ByVal 0&
        End If
        Set cBMP = New cBMPparser
        bReturn = cBMP.LoadStream(aStream(), Me, streamOffset, streamLength)
        Set cBMP = Nothing
    End If
    If m_Handle = 0& Then ' hmmm, not an image file processed here, we can try
                          ' one more thing... Send it GDI+ if possible
        If Me.isGDIplusEnabled Then
            If m_GDIplus Is Nothing Then
                Set cGDI = New cGDIPlus
            Else
                Set cGDI = m_GDIplus
            End If
            If cGDI.GDIplusLoadPNG(FileName, aStream(), Me, m_GDItoken) = True Then
                ' it worked, whatever file type it was. Convert it to PNG
                If SaveFormat = True Then
                    Call cGDI.SaveToPNG(vbNullString, aStream, Me, m_GDItoken)
                    FileHandle = 0& ' prevent next IF from trying to re-load it from file if applicable
                End If
                Me.ImageType = imgPNG
            End If
        End If
    End If
    If Not m_Handle = 0 Then
        If SaveFormat = True Then ' we will cache the original bytes
            If iparseIsArrayEmpty(VarPtrArray(aStream)) = 0& And Not FileHandle = 0& Then
                ' we loaded the image from the file and not a stream (PNG), need to get stream
                'fileNum = FreeFile()
                'Open FileName For Binary Access Read As #fileNum
                If streamLength = 0& Then streamLength = GetFileSize(FileHandle, 0&)
                ReDim m_ImageByteCache(0 To streamLength - 1)
                SetFilePointer FileHandle, 0&, 0&, 0&
                ReadFile FileHandle, m_ImageByteCache(0), streamLength, rtnRead, ByVal 0&
                'Get #fileNum, 1, m_ImageByteCache
                'Close #fileNum
            Else
                m_ImageByteCache() = aStream()
            End If
        End If
        spt_LoadPictureEx = True
    End If

End Function

Private Function spt_pvResize(ByVal destDC As Long, _
                        rSizedBytes() As Byte, rMirror() As Byte, _
                        Optional tHost As c32bppDIB, _
                        Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                        Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                        Optional ByVal destX As Long, Optional ByVal destY As Long, _
                        Optional ByRef destWidth As Long, Optional ByRef destHeight As Long) As Boolean
                            
    ' Function resizes an alpha image, maintaining premultiplied pixels & alpha values
    ' Code originally by Carles P.V. but significantly modified for this project.
    
    ' Parameters:
    ' destDC :: DC being rendered to, may be null
    ' rSizedbytes() : array to hold resized alpha section; not used if tHost is not Nothing
    ' tHost : when resizing to another DIB class, the destination DIB class
    ' srcX,Y : the coordinates of the source image to start resizing from
    ' srcWidth,srcHeight : the width/height of the source image to resize from
    ' destX,Y : the coordinates of the destination image to resize to
    ' destWidth,destHeight : the width/height of the destination image to resize to

    If srcWidth = 0& Then srcWidth = m_Width
    If srcHeight = 0& Then srcHeight = m_Height

    Dim aNewBits() As Byte, dSA As SafeArray   ' new size, overlay of DIB pointer
    Dim aOldBits() As Byte, tSA As SafeArray   ' old size, overlay of DIB pointer
    
    Dim xLUdbl() As Double                     ' look up table (LUT)
    Dim xRatio As Double, yRatio As Double     ' scaled ratios
    Dim srcPixel As Long, dstPixel As Long     ' source/destination pixel locations
    Dim lCol As Long, lRow As Long             ' loop variables
    
    Dim newWidth As Long, newHeight As Long
    
    ' following used with BiLinear scaling
    Dim fX As Double, fY As Double
    Dim tgtX As Long, tgtY As Long
    Dim edgeOffsetX As Long, edgeOffsetY As Long
    Dim iX As Long, iY As Long
    Dim R As Double, G As Double, B As Double, a As Double
    Dim scalerX As Double, scalerY As Double
    
    ' fill in optional parameters
    If Not tHost Is Nothing Then
        newWidth = tHost.Width
        newHeight = tHost.Height
        ' Scaling ratio (ratio of actual image to scaled image)
        xRatio = srcWidth / newWidth
        yRatio = srcHeight / newHeight
        If newWidth > tHost.Width Then newWidth = tHost.Width
        If newHeight > tHost.Height Then newHeight = tHost.Height
    Else
        newWidth = Abs(destWidth)
        newHeight = Abs(destHeight)
        
        ' Scaling ratio (ratio of actual image to scaled image)
        xRatio = srcWidth / newWidth
        yRatio = srcHeight / newHeight

        ' safety checks, recalculation of bounding destination size
        ' if not done, we could very easily access unallocated memory.
        If destX < 0 Then   ' negative DC offset
            newWidth = newWidth + destX    ' reduce width to process
            destX = -destX                 ' used to offset LUT; adjust so not processing bytes not used
        Else
            destX = 0&                        ' fits within destination bitmap; no offsetting needed
        End If
        
        ' now to check the vertical
        If destY < 0& Then  ' negative DC offset
            newHeight = newHeight + destY
            destY = 0&
        Else
            destY = 0&
        End If
        
    End If
    If newHeight < 1& Or newWidth < 1& Then Exit Function
    
    ' set DMA overlays on destination & source
    If tHost Is Nothing Then
        ReDim rSizedBytes(0& To newWidth * 4& - 1&, 0& To newHeight - 1&)
        dSA.pvData = VarPtr(rSizedBytes(0&, 0&))
    Else
        dSA.pvData = tHost.BitsPointer ' called by CopyImageTo & Resize routines
    End If
    iparseOverlayHost_Byte aNewBits, VarPtr(dSA), 2, newHeight, newWidth * 4&, dSA.pvData ' overlay DMA array
    
    If iparseIsArrayEmpty(VarPtrArray(rMirror)) = 0& Then
        iparseOverlayHost_Byte aOldBits, VarPtr(tSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay DMA array
    Else ' source is the mirrored DIB; clipped as needed
        iparseOverlayHost_Byte aOldBits, VarPtr(tSA), 2, UBound(rMirror, 2) + 1&, UBound(rMirror, 1) + 1&, VarPtr(rMirror(0&, 0&))    ' overlay DMA array
    End If
    
    On Error GoTo eh
    
    SrcY = UBound(aOldBits, 2) - srcHeight - SrcY + 1&     ' adjust Y position in source for bottom up DIBs
    
    If (m_StretchQuality = False) Then ' nearest neighbor algorithm (very fast but lowest quality)
        
        ' Scaling LUT, cache actual X position of DIB in relation to scaled X
        ' Cache one scan line of X coords so we don't have to calculate for every pixel
        ReDim xLUdbl(0 To newWidth - 1&)
        For lCol = 0 To newWidth - 1&
            ' offset destX used for negative coordinates, X is location in source to start blending at
            xLUdbl(lCol) = Int(((lCol + destX) * xRatio) + SrcX) * 4&
        Next
        
        For lRow = newHeight - 1& To 0& Step -1&
            '^ current scanline for the scaled image
            ' offset destY is used for negative coordinates
            srcPixel = Int((lRow + destY) * yRatio) + SrcY  ' recalcualted once per scanline
            ' current scanline for the scaled image
            dstPixel = 0&
            For lCol = 0& To newWidth - 1&
                ' copy into resized array the nearest raw/actual pixel
                CopyMemory aNewBits(dstPixel, lRow), aOldBits(Int(xLUdbl(lCol)), srcPixel), 4&
                dstPixel = dstPixel + 4&
            Next lCol
        Next lRow
        
    Else
        
        ' BiLinear interoplation, up to 4 source pixels (16 bytes) are blended for each destination pixel (4 bytes)
 
        ReDim xLUd(0 To newWidth - 1&)   ' work with doubles, we need the decimal portions
        ' Cache one scan line of X coords so we don't have to calculate for every pixel
        For lCol = 0& To newWidth - 1&
            ' offset destX used for negative coordinates, X is location in source to start blending at
            xLUd(lCol) = (((lCol + destX) * xRatio) + SrcX)
        Next
        srcWidth = srcWidth - 1& ' subtract now vs subtracting in loop below
        For lRow = newHeight - 1& To 0& Step -1&
            
            fY = (lRow + destY) * yRatio + SrcY ' get the scaled source row
            tgtY = Int(fY)              ' get whole number of double
            If tgtY = 0& Then         ' for top down images, test for last row vs 0
                ' last row of source image, will use only this row
                edgeOffsetY = 0&
                fY = 0#
            Else
                ' will use this row & next row for blending
                edgeOffsetY = 1&
                fY = Abs(1# - (fY - tgtY)) ' for top down images, use fY = fY-tgtY
                ' ^ Y coordinate fraction; pct of next vertical pixel that is used
            End If
            
            For lCol = 0& To newWidth - 1&
                
                R = 0#: G = 0#: B = 0#: a = 0#
                
                tgtX = Int(xLUd(lCol))  ' coordinate rounded down to whole number
                If tgtX = srcWidth Then
                    ' at far edge of source image, will use only this pixel for blending
                    edgeOffsetX = 0&
                    fX = 0#
                Else
                    ' will use this pixel and next pixel for blending
                    edgeOffsetX = 1&
                    fX = xLUd(lCol) - tgtX
                    ' ^ X coordinate fraction, pct of next horizontal pixel that is used
                End If
                
                For iY = 0& To edgeOffsetY
    
                    scalerY = Abs(iY - fY)   ' percentage of current row's pixel to blend
                    If Not scalerY = 1# Then  ' else result will be zero
                    
                        For iX = 0& To edgeOffsetX
                            scalerX = Abs(fX - iX)   ' percentage of current column's pixel to blend
                            
                            If Not scalerX = 1# Then  ' else result will be zero
                                scalerX = (1# - scalerX) * (1# - scalerY) ' combine percentages
    
                                ' Build the blended RGB values. For top down images Add iY vs subtracting
                                srcPixel = (tgtX + iX) * 4&
                                B = B + aOldBits(srcPixel, tgtY - iY) * scalerX
                                G = G + aOldBits(srcPixel + 1&, tgtY - iY) * scalerX
                                R = R + aOldBits(srcPixel + 2&, tgtY - iY) * scalerX
                                a = a + aOldBits(srcPixel + 3&, tgtY - iY) * scalerX
    
                            End If
                        Next
                    End If
                Next
                iX = lCol * 4&
                ' update destination with adjusted pixel
                aNewBits(iX, lRow) = Int(B)
                aNewBits(iX + 1&, lRow) = Int(G)
                aNewBits(iX + 2&, lRow) = Int(R)
                aNewBits(iX + 3&, lRow) = Int(a)
            Next
        Next
    End If
    
    iparseOverlayHost_Byte aOldBits, 0, 0, 0, 0, 0  ' remove DMA overlay
    iparseOverlayHost_Byte aNewBits, 0, 0, 0, 0, 0  ' remove DMA overlay
    
    ' the passed destWidth,destHeight params are used when rendering; we are just sizing now
    destWidth = newWidth ' the parameter is ByRef, update it now
    destHeight = newHeight ' the parameter is ByRef, update it now
    Erase rMirror()
    
    spt_pvResize = True
eh:
    If Err Then
        Err.Clear   ' troubleshooting only, should be removed before compiling to final app
        Stop
        Resume
    End If
End Function
