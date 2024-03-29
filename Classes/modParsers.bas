Attribute VB_Name = "modParsers"
Option Explicit

' Module contains functions that are required by two or more classes.

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs or different versions, of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.

' Though many of these routines are made Public so that the classes can use them,
' you should not call these routines from your own project. Those that you may
' wish to call anyway, ensure you pass valid, expected parameters. Within the
' classes, parameters are validated and these routines may not have additional
' validation checks which could result in crashes or memory leaks if used incorrectly.

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' 32 bytes as used. Can be used for 1D and/or 2D arrays
End Type
Private Type PictDesc
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

' used to see if DLL exported function exists
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' GDI32 APIs
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' User32 APIs
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' Kernel32/User32 APIs for Unicode Filename Support
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const FILE_ATTRIBUTE_NORMAL = &H80&

Public Function iparseIsArrayEmpty(FarPointer As Long) As Long
  ' test to see if an array has been initialized
  CopyMemory iparseIsArrayEmpty, ByVal FarPointer, 4&
End Function

Public Function iparseByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    iparseByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

Public Function iparseArrayToPicture(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty when calling class' LoadStream was called
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), iparseArrayToPicture)
            End If
        End If
    End If

End Function

Public Function iparseHandleToStdPicture(ByVal hImage As Long, ByVal imgType As Long) As IPicture

    ' function creates a stdPicture object from a image handle (bitmap or icon)
    
    Dim lpPictDesc As PictDesc, aGUID(0 To 3) As Long
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hHandle = hImage
        .hPal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, iparseHandleToStdPicture)
    
End Function

Public Function iparseReverseLong(ByVal inLong As Long) As Long

    ' fast function to reverse a long value from big endian to little endian
    ' PNG files contain reversed longs
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim b4 As Long
    Dim lHighBit As Long
    
    lHighBit = inLong And &H80000000
    If lHighBit Then
      inLong = inLong And Not &H80000000
    End If
    
    b1 = inLong And &HFF
    b2 = (inLong And &HFF00&) \ &H100&
    b3 = (inLong And &HFF0000) \ &H10000
    If lHighBit Then
      b4 = inLong \ &H1000000 Or &H80&
    Else
      b4 = inLong \ &H1000000
    End If
    
    If b1 And &H80& Then
      iparseReverseLong = ((b1 And &H7F&) * &H1000000 Or &H80000000) Or _
          b2 * &H10000 Or b3 * &H100& Or b4
    Else
      iparseReverseLong = b1 * &H1000000 Or _
          b2 * &H10000 Or b3 * &H100& Or b4
    End If

End Function

Public Function iparseValidateDLL(ByVal DllName As String, ByVal dllProc As String) As Boolean
    
    ' PURPOSE: Test a DLL for a specific function.
    
    Dim LB As Long, pa As Long
    
    'attempt to open the DLL to be checked
    LB = LoadLibrary(DllName)
    If LB Then
        'if so, retrieve the address of one of the function calls
        pa = GetProcAddress(LB, dllProc)
        ' free references
        FreeLibrary LB
    End If
    iparseValidateDLL = (Not (LB = 0 Or pa = 0))
    
End Function

Public Function iparseValidateZLIB(ByRef DllName As String, ByRef Version As Long, _
                                ByRef isCDECL As Boolean, ByRef hasCompression2 As Boolean, _
                                Optional ByVal bTestOnly As Boolean) As Boolean
    
    ' PURPOSE: Test ZLib availability and calling convention.
    
    ' About zLIB.  There are several versions ranging from v1.2.3 (latest) to v1.0.? (earliest).
    ' Zlib is used to compress/decompress PNG files, among other things.
    
    ' However, zLIB is written with C calling convention (cdecl) which is unusable with VB.  There
    ' are other modified versions out there that were converted to stdcall calling convention which
    ' is what VB expects. But, we don't know the calling convention of the zLIB in advance, do we?
    
    ' Allowing VB to call cdecl directly results in crashes or invalid function returns. The class
    ' cCDECL is one created by Paul Caton that uses assembly to "wrap" the cdecl call into a stdcall.
    ' But we still need to know the calling convention so we know to use cCDECL or simple API calls.
    
    Dim LB As Long, pa As Long
    Dim asmVal As Integer
    
    DllName = "zlib1.dll"       ' test for newer version first
    For Version = 2& To 1& Step -1&
        LB = LoadLibrary(DllName) 'attempt to open the DLL to be checked
        If LB Then
            hasCompression2 = Not (GetProcAddress(LB, "compress2") = 0)
            pa = GetProcAddress(LB, "crc32") ' retrieve the address of the "crc32" exported function
            If Not pa = 0& Then
                
                If bTestOnly Then Exit For
                Do
                    ' Note: this method will not work for every DLL, nor every function within a DLL.
                    ' I have analyzed 5 versions of zlib (some cdecl, some stdcall) using disassemblers
                    ' and am confident this will work for the crc32 function in all versions from v1.2.3 down.
                    
                    ' Looking for an exit code:
                    CopyMemory asmVal, ByVal pa, 1&
                    Select Case asmVal
                        Case &HC3               ' exit code, no stack clean up
                            CopyMemory asmVal, ByVal iparseSafeOffset(pa, -1&), 1&
                            If Not asmVal = &H33 Then   ' else 0x33C3 is an XOR function, not exit code
                                isCDECL = True      ' DLL uses cdecl calling convention, we use cCDECL
                                Exit For
                            End If
                        Case &HC2
                            CopyMemory asmVal, ByVal iparseSafeOffset(pa, 1&), 2&
                            If asmVal = &HC Then ' exit code with clean up of 12 bytes (the 3 crc32 parameters)
                                isCDECL = False  ' DLL uses stdcall calling convention, we use APIs
                                Exit For
                            Else
                                asmVal = 0
                            End If
                    End Select
                    pa = iparseSafeOffset(pa, 1&)
                Loop
            End If
            ' unmap library
            FreeLibrary LB
            LB = 0&
            hasCompression2 = False
        End If
        DllName = "zlib.dll"    ' test for older version next, if necessary
    Next Version
    
    If Not LB = 0& Then FreeLibrary LB
    iparseValidateZLIB = (Not (Version = 0&))
    
End Function


Public Sub iparseValidateAlphaChannel(inStream() As Byte, bPreMultiply As Boolean, bIsAlpha As Boolean, imgType As Long)

    ' Purpose: Modify 32bpp DIB's alpha bytes depending on whether or not they are used
    
    ' Parameters
    ' inStream(). 2D array overlaying the DIB to be checked
    ' bPreMultiply. If true, image will be premultiplied if not already
    ' bIsAlpha. Returns whether or not the image contains transparency
    ' imgType. If passed as -1 then image is known to be not alpha, but will have its alpha values set to 255
    '          When routine returns, imgType is either imgBmpARGB, imgBmpPARGB or imgBitmap

    Dim X As Long, Y As Long
    Dim lPARGB As Long, zeroCount As Long, opaqueCount As Long
    Dim bPARGB As Boolean, bAlpha As Boolean

    ' see if the 32bpp is premultiplied or not and if it is alpha or not
    If Not imgType = -1 Then
        For Y = 0 To UBound(inStream, 2)
            For X = 3 To UBound(inStream, 1) Step 4
                Select Case inStream(X, Y)
                Case 0
                    If lPARGB = 0 Then
                        ' zero alpha, if any of the RGB bytes are non-zero, then this is not pre-multiplied
                        If Not inStream(X - 1, Y) = 0 Then
                            lPARGB = 1 ' not premultiplied
                        ElseIf Not inStream(X - 2, Y) = 0 Then
                            lPARGB = 1
                        ElseIf Not inStream(X - 3, Y) = 0 Then
                            lPARGB = 1
                        End If
                        ' but don't exit loop until we know if any alphas are non-zero
                    End If
                    zeroCount = zeroCount + 1 ' helps in decision factor at end of loop
                Case 255
                    ' no way to indicate if premultiplied or not, unless...
                    If lPARGB = 1 Then
                        lPARGB = 2    ' not pre-multiplied because of the zero check above
                        Exit For
                    End If
                    opaqueCount = opaqueCount + 1
                Case Else
                    ' if any Exit For's below get triggered, not pre-multiplied
                    If lPARGB = 1 Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(X - 3, Y) > inStream(X, Y) Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(X - 2, Y) > inStream(X, Y) Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(X - 1, Y) > inStream(X, Y) Then
                        lPARGB = 2: Exit For
                    End If
                End Select
            Next
            If lPARGB = 2 Then Exit For
        Next
        
        ' if we got all the way thru the image without hitting Exit:For then
        ' the image is not alpha unless the bAlpha flag was set in the loop
        
        If zeroCount = (X \ 4) * (UBound(inStream, 2) + 1) Then ' every alpha value was zero
            bPARGB = False: bAlpha = False ' assume RGB, else 100% transparent ARGB
            ' also if lPARGB=0, then image is completely black
        ElseIf opaqueCount = (X \ 4) * (UBound(inStream, 2) + 1) Then ' every alpha is 255
            bPARGB = False: bAlpha = False
        Else
            Select Case lPARGB
                Case 2: bPARGB = False: bAlpha = True ' 100% positive ARGB
                Case 1: bPARGB = False: bAlpha = True ' now 100% positive ARGB
                Case 0: bPARGB = True: bAlpha = True
            End Select
        End If
    End If
    
    ' see if caller wants the non-premultiplied alpha channel premultiplied
    If bAlpha = True Then
        If bPARGB Then ' else force premultiplied
            imgType = imgBmpPARGB
        Else
            imgType = imgBmpARGB
            If bPreMultiply = True Then
                For Y = 0 To UBound(inStream, 2)
                    For X = 3 To UBound(inStream, 1) Step 4
                        If inStream(X, Y) = 0 Then
                            CopyMemory inStream(X - 3, Y), 0&, 4&
                        ElseIf Not inStream(X, Y) = 255 Then
                            For lPARGB = X - 3 To X - 1
                                inStream(lPARGB, Y) = ((0& + inStream(lPARGB, Y)) * inStream(X, Y)) \ &HFF
                            Next
                        End If
                    Next
                Next
                bAlpha = True
            End If
        End If
    Else
        imgType = imgBitmap
        If bPreMultiply = True Then
            For Y = 0 To UBound(inStream, 2)
                For X = 3 To UBound(inStream, 1) Step 4
                    inStream(X, Y) = 255
                Next
            Next
        End If
    End If
    bIsAlpha = bAlpha

End Sub

Public Function iparseSafeOffset(ByVal Ptr As Long, Offset As Long) As Long

    ' ref http://support.microsoft.com/kb/q189323/ ' unsigned math
    ' Purpose: Provide a valid pointer offset
    
    ' If a pointer +/- the offset wraps around the high bit of a long, the
    ' pointer needs to change from positive to negative or vice versa.
    
    ' A return of zero indicates the offset exceeds the min/max unsigned long bounds
    
    Const MAXINT_4NEG As Long = -2147483648#
    Const MAXINT_4 As Long = 2147483647
    
    If Offset = 0 Then
        iparseSafeOffset = Ptr
    Else
    
        If Offset < 0 Then ' subtracting from pointer
            If Ptr < MAXINT_4NEG - Offset Then
                ' wraps around high bit (backwards) & changes to Positive from Negative
                iparseSafeOffset = MAXINT_4 - ((MAXINT_4NEG - Ptr) - Offset - 1)
            ElseIf Ptr > 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr > -Offset Then iparseSafeOffset = Ptr + Offset
            Else
                iparseSafeOffset = Ptr + Offset
            End If
        Else    ' Adding to pointer
            If Ptr > MAXINT_4 - Offset Then
                ' wraps around high bit (forward) & changes to Negative from Positive
                iparseSafeOffset = MAXINT_4NEG + (Offset - (MAXINT_4 - Ptr) - 1)
            ElseIf Ptr < 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr < -Offset Then iparseSafeOffset = Ptr + Offset
            Else
                iparseSafeOffset = Ptr + Offset
            End If
        End If
    End If

End Function

Public Function iparseGetFileHandle(ByVal FileName As String, bOpen As Boolean, Optional ByVal useUnicode As Boolean = False) As Long

    ' Function uses APIs to read/create files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    Const FILE_SHARE_WRITE As Long = &H2
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    
    Dim Flags As Long, Access As Long
    Dim Disposition As Long, Share As Long
    
    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If bOpen Then
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    Else
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        If useUnicode Then
            Flags = GetFileAttributesW(StrPtr(FileName))
        Else
            Flags = GetFileAttributes(FileName)
        End If
        If Flags < 0& Then Flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    End If
    
    If useUnicode Then
        iparseGetFileHandle = CreateFileW(StrPtr(FileName), Access, Share, ByVal 0&, Disposition, Flags, 0&)
    Else
        iparseGetFileHandle = CreateFile(FileName, Access, Share, ByVal 0&, Disposition, Flags, 0&)
    End If

End Function

Public Function iparseDeleteFile(FileName As String, Optional ByVal useUnicode As Boolean = False) As Boolean

    ' Function uses APIs to delete files :: unicode supported

    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If useUnicode Then
        If Not (SetFileAttributesW(StrPtr(FileName), FILE_ATTRIBUTE_NORMAL) = 0&) Then
            iparseDeleteFile = Not (DeleteFileW(StrPtr(FileName)) = 0&)
        End If
    Else
        If Not (SetFileAttributes(FileName, FILE_ATTRIBUTE_NORMAL) = 0&) Then
            iparseDeleteFile = Not (DeleteFile(FileName) = 0&)
        End If
    End If

End Function

Public Function iparseFileExists(FileName As String, Optional ByVal useUnicode As Boolean) As Boolean
    ' test to see if a file exists
    Const INVALID_HANDLE_VALUE = -1&
    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If useUnicode Then
        iparseFileExists = Not (GetFileAttributesW(StrPtr(FileName)) = INVALID_HANDLE_VALUE)
    Else
        iparseFileExists = Not (GetFileAttributes(FileName) = INVALID_HANDLE_VALUE)
    End If
End Function

Public Sub iparseOverlayHost_Byte(aOverlay() As Byte, ptrSafeArray As Long, nrDims As Long, ElemCount_Dim1 As Long, ElemCount_Dim2 As Long, ByVal memPtr As Long)

    ' Routine overlays a BYTE array on top of some memory address. Passing incorrect values will crash the app and maybe the system
    ' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10
    
    ' aOverlay() is an uninitialized, dynamic Byte array. If initialized, call Erase on the array before passing it
    ' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
    ' nrDims must be 1 or 2. Not used if memPtr=0
    ' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
    ' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
    ' memPtr is the memory address to overlay the array onto
    
    If memPtr = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), memPtr, 4&      ' remove overlay
    Else
        Dim tSA As SafeArray
        With tSA
            .cbElements = 1     '1=byte
            .pvData = memPtr    'memory address
            .cDims = nrDims     'nr of dimensions
            If nrDims = 2 Then
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (1st dimension)
                .rgSABound(1).cElements = ElemCount_Dim2  'number array items (2nd dimension)
            Else
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (only one dimension)
            End If
            ' Note: the .LBound members of .rgSABound are always zero. Set them on routine's return if needed. Remember right to left order
        End With
        CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
    End If

End Sub

Public Sub iparseOverlayHost_Long(aOverlay() As Long, ptrSafeArray As Long, nrDims As Long, ElemCount_Dim1 As Long, ElemCount_Dim2 As Long, ByVal memPtr As Long)

    ' Routine overlays a LONG array on top of some memory address. Passing incorrect values will crash the app and maybe the system
    ' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10
    
    ' aOverlay() is an uninitialized, dynamic Long array. If initialized, call Erase on the array before passing it
    ' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
    ' nrDims must be 1 or 2. Not used if memPtr=0
    ' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
    ' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
    ' memPtr is the memory address to overlay the array onto
    
    If memPtr = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), memPtr, 4&      ' remove overlay
    Else
        Dim tSA As SafeArray
        With tSA
            .cbElements = 4     '4=long
            .pvData = memPtr    'memory address
            .cDims = nrDims     'nr of dimensions
            If nrDims = 2 Then
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (1st dimension)
                .rgSABound(1).cElements = ElemCount_Dim2  'number array items (2nd dimension)
            Else
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (only one dimension)
            End If
            ' Note: the .LBound members of .rgSABound are always zero. Set them on routine's return if needed. Remember right to left order
        End With
        CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
    End If

End Sub



