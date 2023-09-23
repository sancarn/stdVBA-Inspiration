VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Import/Read Images from Executable Files ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' The binary file contents are never cached.
' The class terminates once user selects one of the resources within the binary file
' Though animated cursors/icons are valid resource types, I have scanned 1000s of dll/exe/ocx files
'    and have never found one. Therefore, those resource types are excluded from this class.
' The following resource types are processed: icon, cursor, bitmap, AVI

Private Declare Function LoadLibraryExW Lib "kernel32" (ByVal lpLibFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function LoadLibraryExA Lib "kernel32" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FindResourceA Lib "kernel32" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Private Declare Function FindResourceW Lib "kernel32" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function LoadImageA Lib "user32.dll" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImageW Lib "user32.dll" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal uFlags As Long) As Long

Private Enum ExtractionModeEnum
    exm_WantImage = 0
    exm_WantAttr = 1
    exm_WantData = 2
End Enum

Private m_File As String            ' Binary's file name
Private m_ResName As String         ' resource ID for current image
Private m_GroupCount As Long        ' number of groups: ICO, CUR, BMP, AVI
Private m_Groups() As Long          ' 2D array identifying which group & nr of images in group
Private m_Index As Long             ' current group & image index
Private m_BMPI As BITMAPINFOHEADER  ' description of currently selected item
Private m_Token As cGDIpToken

Private Const MAXGROUPS As Long = 4&
Private Const RT_CURSOR As Long = 1&
Private Const RT_BITMAP As Long = 2&
Private Const RT_ICON As Long = 3&
Private Const RT_AVI As Long = 4&       ' not an actual resource type ID
Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2&

Public Function IsBinaryResource(FileName As String, cImageData As cGDIpMultiImage) As Long

    ' called to validate file is an executable
    ' Note that binaries are never accepted/loaded by array
    
    ' value of 1 indicates binary with images
    ' value of 0 indicates not a binary
    ' any other value indicates binary but no images or failed to load
    
    If m_Token.Token = 0& Then Exit Function
    
    If pvValidatePEheader(FileName) Then            ' call quick parser
        IsBinaryResource = True
        If pvScanBinary(FileName) = True Then       ' find which resource types we want & count
            cImageData.InitializeContent ObjPtr(Me), lvicPicTypeFromBinaries, m_Groups(1, 0)
            m_File = FileName
            cImageData.SetIndex Empty, 0&, &HF0000000 ' load first resource found
            If cImageData.Handle Then
                IsBinaryResource = 1&               ' return success or reset passed class
            Else
                Set cImageData = New cGDIpMultiImage
            End If
        End If
    End If
End Function

Public Function ExtractData(outArray() As Byte, gHandle As Long) As Boolean

    ' called to return the content of a specific resource in original format
    
    If m_ResName <> vbNullString Then
        If Me.ImageFormat() = lvicPicTypeBitmap Then
            ExtractData = (modCommon.SaveAsBMP(outArray(), gHandle, saveTo_Array) <> 0&)
        Else
            ExtractData = (pvExtractResource(outArray(), 0&, exm_WantData) <> 0&)
        End If
    End If

End Function

Public Function SetIndex(ByVal actualIndex As Long, ByVal Group As Long, outData() As Byte, _
                            newIndex As Long, ImgCount As Long) As Long

    ' sets the group & index of the binary resource & loads 1st image of that group
    ' newIndex, outData(), and ImgCount are returned as needed
    
    ' Group passed as 1-based index
    If m_GroupCount = 0& Then Exit Function
    
    ' returns a GDI+ image handle of the requested binary resource
    If Group < 0& Then
        Group = 1&
    ElseIf Group > MAXGROUPS Then
        Exit Function
    End If
    If actualIndex < 0& Then actualIndex = 0&
    Group = Group - 1&
    
    If m_Groups(0, Group) = 0& Then Exit Function  ' no images for this group
    If actualIndex >= m_Groups(1, Group) Then Exit Function ' out of bounds
    
    m_Index = actualIndex Or Group * &H10000        ' cache current group & index
    newIndex = actualIndex + 1&
    ImgCount = m_Groups(1, Group)
                                                    ' return image handle and array (if png format)
    SetIndex = pvExtractResource(outData(), 0&, exm_WantImage)

End Function

Public Property Get GroupCount() As Long
    ' simply returns number of "groups" found in the binary resource
    GroupCount = m_GroupCount
End Property

Public Property Get GroupFormat(GroupIndex As Long) As ImageFormatEnum
    ' returns the type of image format for the specified group
    If m_GroupCount = 0& Then Exit Property
    If GroupIndex < 1& Then GroupIndex = (m_Index \ &H10000) + 1&
    If GroupIndex > 0& And GroupIndex <= MAXGROUPS Then
        If m_GroupCount Then
            Select Case m_Groups(0, GroupIndex - 1&)
            Case RT_ICON:   GroupFormat = lvicPicTypeIcon
            Case RT_BITMAP: GroupFormat = lvicPicTypeBitmap
            Case RT_CURSOR: GroupFormat = lvicPicTypeCursor
            Case RT_AVI:    GroupFormat = lvicPicTypeAVI
            Case Else:      GroupFormat = lvicPicTypeNone
            End Select
        End If
    End If
End Property

Public Property Get ImageFormat() As ImageFormatEnum
    ' returns the current image's format
    If m_GroupCount Then ImageFormat = Me.GroupFormat(0&)
End Property

Public Property Get GroupImageCount(GroupIndex As Long)
    ' returns the number of images found in the requested group
    If m_GroupCount = 0& Then Exit Property
    If GroupIndex < 1& Then GroupIndex = (m_Index \ &H10000) + 1&
    If GroupIndex > 0& And GroupIndex <= MAXGROUPS Then
        GroupImageCount = m_Groups(1, GroupIndex - 1&)
    End If
End Property

Public Sub GetDimensions(ByVal Index As Long, Width As Single, Height As Single)
    ' returns the bounds of the requested image
    If Index = (m_Index And &HFFFF&) Then
        Width = m_BMPI.biWidth
        Height = m_BMPI.biHeight
    Else
        If m_GroupCount = 0& Then Exit Sub
        If Index < 0& Then Index = 0&
        If Index < m_Groups(1, (m_Index \ &H10000)) Then
            Dim tBMP As BITMAPINFOHEADER, bDummy() As Byte
            tBMP = m_BMPI
            Call pvExtractResource(bDummy(), Index, exm_WantAttr)
            Width = m_BMPI.biWidth
            Height = m_BMPI.biHeight
            m_BMPI = tBMP
        End If
    End If
End Sub

Public Property Get BitDepth(Index As Long) As Long
    ' returns the color depth of the requested image
    If Index = (m_Index And &HFFFF&) Then
        BitDepth = m_BMPI.biBitCount
    Else
        If m_GroupCount = 0& Then Exit Property
        If Index < 0& Then Index = 0&
        If Index < m_Groups(1, (m_Index \ &H10000)) Then
            Dim tBMP As BITMAPINFOHEADER, bDummy() As Byte
            tBMP = m_BMPI
            Call pvExtractResource(bDummy(), Index, exm_WantAttr)
            BitDepth = m_BMPI.biBitCount
            m_BMPI = tBMP
        End If
    End If
End Property

Private Function pvScanBinary(FileName As String) As Boolean

    ' function scans a binary to extract number of groups & images per group

    Dim hMod As Long, lType As Long
    Dim lResult As Long, X As Long
    Dim bAVI() As Byte, sID As String
    
    m_GroupCount = 0&
    ReDim m_Groups(0 To 1, 0 To MAXGROUPS - 1&)
    m_Index = &HFFFF&
    
    If g_UnicodeSystem Then                 ' first load the target file
        hMod = LoadLibraryExW(StrPtr(FileName), 0&, LOAD_LIBRARY_AS_DATAFILE)
    Else
        hMod = LoadLibraryExA(FileName, 0&, LOAD_LIBRARY_AS_DATAFILE)
    End If
    If hMod Then
        sID = "AVI" & vbNullChar
        For X = 1& To MAXGROUPS
            Select Case X
            Case 1: lType = RT_BITMAP
            Case 2: lType = RT_ICON
            Case 3: lType = RT_CURSOR
            Case 4 ' AVI doesn't have a RT_xxxx constant. It must be known by its string name
                If g_UnicodeSystem Then
                    lType = StrPtr(sID)
                Else
                    bAVI() = StrConv(sID, vbFromUnicode): lType = VarPtr(bAVI(0))
                End If
            End Select
            lResult = &H80000000            ' flag indicating we want count, nothing more
            If g_UnicodeSystem Then         ' call that function
                EnumResourceNamesW hMod, lType, AddressOf modCommon.EnumResNamesProcW, lResult
            Else
                EnumResourceNamesA hMod, lType, AddressOf modCommon.EnumResNamesProcA, lResult
            End If
            If (lResult And &HFFFF&) Then               ' found some?
            
                If (lType And &HFFFFFFF8) = 0& Then     ' annotate the type found
                    m_Groups(0, m_GroupCount) = lType
                Else
                    m_Groups(0, m_GroupCount) = RT_AVI
                End If                                  ' update the image count in that group
                m_Groups(1, m_GroupCount) = (lResult And &H7FFF&)
                m_GroupCount = m_GroupCount + 1&        ' keep track of groups found
            End If
        Next
        FreeLibrary hMod                                ' done; unload library
    End If
    pvScanBinary = (m_GroupCount > 0&)                  ' return success/failure
    
End Function

Private Function pvValidatePEheader(FileName As String) As Boolean

    ' routine parses a file to determine if its a 32bit executable
    ' Executables can contain embedded images (resource file)
    ' We will not process 64 bit executables
    
    ' Note that we do not rely on file extensions here

    Dim hHandle As Long
    hHandle = modCommon.GetFileHandle(FileName, False)
    If hHandle = INVALID_HANDLE_VALUE Or hHandle = 0& Then Exit Function
    
    Dim lResult As Long, lRead As Long
    
    On Error GoTo ExitRoutine
    SetFilePointer hHandle, 0&, 0&, 0&
    ReadFile hHandle, lResult, 2&, lRead, ByVal 0&          ' get 1st 2 bytes of 64 byte MS_DOS_HEADER
    If lResult = &H5A4D& Then                               ' "MZ"? then possibly an executable
        SetFilePointer hHandle, 60&, 0&, 0&                 ' get the last 4 bytes. offset to PE header
        ReadFile hHandle, lResult, 4&, lRead, ByVal 0&      ' validate not out of bounds
        If lResult > 60& And lResult < GetFileSize(hHandle, 0&) Then
            SetFilePointer hHandle, lResult, 0&, 0&         ' move to that offset & read 1st 4 bytes
            ReadFile hHandle, lResult, 4&, lRead, ByVal 0&
            If lResult = &H4550 Then                        ' matches expected magic number: PE00
                ReadFile hHandle, lResult, 2&, lRead, ByVal 0&
                pvValidatePEheader = ((lResult And &HFFFF&) = &H14C&) ' then 32bit resource
'                 Note to self: If issues arise using this method, more info can be retrieved about the
'                 executable. Unrem following & test lResult against values provided at this link:
'                 http://msdn.microsoft.com/en-us/library/ms680313%28v=vs.85%29.aspx
'                 SetFilePointer hHandle, 16&, 0&, 1&         ' move forward 16 bytes & get next 2 bytes
'                 ReadFile hHandle, lResult, 2&, lRead, ByVal 0&
            End If
        End If
    End If
ExitRoutine:
CloseHandle hHandle

End Function

Private Function pvExtractResource(outData() As Byte, resIndex As Long, Mode As ExtractionModeEnum) As Long

    ' function has 3 purposes. The Mode parameter:
    ' 1) return an image handle to be displayed in the property page
    ' 2) return the image as a byte array, complete format, for saving/caching
    ' 3) return image attributes of only size & color depth

    Dim hMod As Long, hRes As Long, hGlobal As Long
    Dim sID As String, lType As Long
    Dim lResult As Long, lCount As Long
    Dim bAVI() As Byte, sAVI As String
    Dim cImgData As cGDIpMultiImage

    If g_UnicodeSystem Then                 ' first load the target file
        hMod = LoadLibraryExW(StrPtr(m_File), 0&, LOAD_LIBRARY_AS_DATAFILE)
    Else
        hMod = LoadLibraryExA(m_File, 0&, LOAD_LIBRARY_AS_DATAFILE)
    End If
    If hMod Then                            ' Identify which ordinal we want to find
        If Mode = exm_WantAttr Then
            resIndex = (m_Index And &HFFFF0000) Or resIndex
        Else
            resIndex = m_Index
        End If                              ' also identify which group we are looking in
        If m_Groups(0, (resIndex \ &H10000)) = RT_AVI Then
            sAVI = "AVI" & vbNullChar
            If g_UnicodeSystem Then
                lType = StrPtr(sAVI)
            Else
                bAVI() = StrConv(sAVI, vbFromUnicode)
                lType = VarPtr(bAVI(0))
            End If
        Else
            lType = m_Groups(0, (resIndex \ &H10000))
        End If
        If Mode = exm_WantData Then         ' we may already have the resource ID to search for
            sID = m_ResName
        Else                                ' else we need to find it
            lResult = (resIndex And &HFFFF&)
            If g_UnicodeSystem Then     ' call that function
                EnumResourceNamesW hMod, lType, AddressOf modCommon.EnumResNamesProcW, lResult
            Else
                EnumResourceNamesA hMod, lType, AddressOf modCommon.EnumResNamesProcA, lResult
            End If
    
            If lResult = 0& Then   ' it's a pointer to a string
                lType = 0&
            Else
                CopyMemory ByVal VarPtr(sID), lResult, 4&: lResult = 0&
            End If
        End If
        If sID <> vbNullString Then                     ' get to the target source bits
            If g_UnicodeSystem Then
                hRes = FindResourceW(hMod, StrPtr(sID), lType)
            Else
                hRes = FindResourceA(hMod, sID, lType)
            End If
            If hRes Then hGlobal = LoadResource(hMod, hRes)
            If hGlobal Then lCount = SizeofResource(hMod, hRes)
            If lCount > 44& Then lResult = LockResource(hGlobal)
            ' LockResource does not actually lock memory; it is just used to obtain a pointer to the memory containing the resource data
            ' therefore, one does not unlock the resource, nor use any Globalxxx functions on it. System takes care of releasing it
            If lResult = 0& Then
                lType = 0&: sID = vbNullString          ' failure for some reason
            Else
                Select Case lType
                    Case RT_BITMAP                      ' call function that will correctly handle alpha bitmaps
                        lResult = pvBitmapFromResource(lResult, lCount, Mode)
                        If Mode = exm_WantImage Then
                            If lResult = 1& Then        ' not a 32bpp alpha bitmap, but a bitmap
                                FreeResource hRes: hRes = 0&
                                If g_UnicodeSystem Then ' create bitmap using APIs
                                    lResult = LoadImageW(hMod, StrPtr(sID), 0&, 0&, 0&, 0&)
                                Else
                                    lResult = LoadImageA(hMod, sID, 0&, 0&, 0&, 0&)
                                End If
                                If lResult Then         ' clone the bitmap & return the handle
                                    Call GdipCreateBitmapFromHBITMAP(lResult, 0&, pvExtractResource)
                                    DeleteObject lResult ' destroy the API bitmap
                                End If
                            Else
                                pvExtractResource = lResult
                            End If
                        Else
                            pvExtractResource = lResult
                        End If
                    Case RT_ICON    ' call function to extract or return icon image/data (including PNGs)
                        pvExtractResource = pvIconFromResource(lResult, lCount, outData(), Mode)
                    Case RT_CURSOR  ' call function to extract or return cursor image/data (including PNGs)
                        pvExtractResource = pvCursorFromResource(lResult, lCount, outData(), Mode)
                    Case Else       ' call function to extract or return AVI
                        pvExtractResource = pvAVIfromResource(lResult, lCount, outData(), Mode)
                End Select
            End If
            If hRes Then FreeResource hRes          ' clean up
        End If
    End If
    If hMod Then FreeLibrary hMod                   ' clean up
    If Mode = exm_WantImage Then m_ResName = sID    ' save for eventual data extraction/caching
    
End Function

Private Function pvBitmapFromResource(resPtr As Long, resSize As Long, Mode As ExtractionModeEnum) As Long

    ' function returns one of three things
    ' 1. Simply fills in the m_BMPI structure and returns nothing else
    ' 2. Returns a GDI+ 32bpp image handle (applies only when alpha channel is used)
    ' 3. Reutrns value of 1 to indicate resource API to be used to create the bitmap handle if needed
    
    Dim tSA As SafeArray, bmpData() As Long, hHandle As Long
    Dim tSize As RECTI
    Dim tBD As BitmapData
    
    CopyMemory m_BMPI, ByVal resPtr, 40&
    If Mode = exm_WantAttr Then Exit Function
    
    On Error GoTo ExitRoutine
    If (m_BMPI.biWidth < 1& Or m_BMPI.biHeight = 0&) Then Exit Function
    If m_BMPI.biBitCount = 32& Then
        m_BMPI.biSizeImage = (m_BMPI.biWidth * Abs(m_BMPI.biHeight) * 4&)
        If resSize >= m_BMPI.biSizeImage Then
            With tSA
                .cbElements = 4
                .cDims = 2
                .pvData = resPtr + resSize - m_BMPI.biSizeImage
                .rgSABound(0).cElements = Abs(m_BMPI.biHeight)
                .rgSABound(1).cElements = m_BMPI.biWidth
            End With
            CopyMemory ByVal VarPtrArray(bmpData), VarPtr(tSA), 4&
            tBD.PixelFormat = modCommon.ValidateAlphaChannel(bmpData())
            CopyMemory ByVal VarPtrArray(bmpData), 0&, 4&
            If tBD.PixelFormat <> lvicColor32bpp Then
                tBD.stride = m_BMPI.biWidth * 4&
                tSize.nHeight = Abs(m_BMPI.biHeight): tSize.nWidth = m_BMPI.biWidth
                If GdipCreateBitmapFromScan0(tSize.nWidth, tSize.nHeight, 0&, tBD.PixelFormat, ByVal 0&, hHandle) = 0& Then
                    If m_BMPI.biHeight < 0& Then            ' top down image (do these even exist in res files?)
                        tBD.Scan0Ptr = tSA.pvData
                    Else
                        tBD.stride = -tBD.stride
                        tBD.Scan0Ptr = tSA.pvData + m_BMPI.biSizeImage + tBD.stride
                    End If
                    If GdipBitmapLockBits(hHandle, tSize, ImageLockModeUserInputBuf Or ImageLockModeWrite, tBD.PixelFormat, tBD) = 0& Then
                        If GdipBitmapUnlockBits(hHandle, tBD) = 0& Then
                            pvBitmapFromResource = hHandle
                            Exit Function
                        End If
                    End If
                    GdipDisposeImage hHandle
                End If
            End If
        End If
    End If
    pvBitmapFromResource = 1&
    
ExitRoutine:
End Function

Private Function pvIconFromResource(resPtr As Long, resSize As Long, outData() As Byte, Mode As ExtractionModeEnum) As Long

    ' function returns one of three things
    ' 1. Simply fills in the m_BMPI structure and returns nothing else
    ' 2. A GDI+ image handle of the selected icon (PNGs are not created here but data returned to calling class)
    ' 3. An array containing the selected image in icon format
    
    Dim lResult As Long, Cx As Long, Cy As Long
    Dim bIsPng As Boolean, tSource As Variant
    
    On Error GoTo ExitRoutine
    
    CopyMemory m_BMPI, ByVal resPtr, 40&
    If m_BMPI.biSize = png_Signature1 Then
        If m_BMPI.biWidth <> png_Signature2 Then Exit Function
        Cx = modCommon.ReverseLong(m_BMPI.biCompression)
        If Cx < 1& Or Cx > 256& Then Exit Function
        Cy = modCommon.ReverseLong(m_BMPI.biSizeImage)
        If Cy < 1& Or Cy > 256& Then Exit Function
        ' 24th byte from resource start will be the PNG bit depth
        ' 25th byte from resource start will be the color type
        CopyMemory m_BMPI.biBitCount, ByVal resPtr + 24, 2&
        Select Case (m_BMPI.biBitCount And &H7F00) \ &H100
            Case 2: If (m_BMPI.biBitCount And &HFF) = 8 Then m_BMPI.biBitCount = 24
            Case 6: If (m_BMPI.biBitCount And &HFF) = 8 Then m_BMPI.biBitCount = 32
            Case Else ' use reported depth
                m_BMPI.biBitCount = (m_BMPI.biBitCount And &HFF)
        End Select
        bIsPng = True
    End If
    
    If Not Mode = exm_WantAttr Then ' want entire resource
        If bIsPng And Mode = exm_WantImage Then
            ReDim outData(0 To resSize - 1&)
            CopyMemory outData(0), ByVal resPtr, resSize
        Else
            ReDim outData(0 To resSize + 21&)
            CopyMemory outData(22), ByVal resPtr, resSize
            outData(2) = 1
            outData(4) = 1
            outData(18) = 22
            CopyMemory outData(14), resSize, 4&             ' size of resource
            If bIsPng Then
                outData(10) = 1                             ' planes
                outData(12) = (m_BMPI.biBitCount And &HFF)
            Else
                CopyMemory outData(10), m_BMPI.biPlanes, 4& ' copy planes & bitcount
            End If
        End If
    End If
    If Not bIsPng Then
        If m_BMPI.biWidth < 1& Or m_BMPI.biWidth > 256& Then Exit Function
        Cx = m_BMPI.biWidth
        If m_BMPI.biHeight = 0& Or Abs(m_BMPI.biHeight) > 512& Then Exit Function
        Cy = m_BMPI.biHeight \ 2
        If Not Mode = exm_WantAttr Then
            If m_BMPI.biBitCount < 8 Then
                If m_BMPI.biBitCount = 1 Then outData(8) = 2 Else outData(8) = 16
            End If
        End If
    End If
    If Not Mode = exm_WantAttr Then
        If Cx < 256& Then outData(6) = Cx
        If Cy < 256& Then outData(7) = Cy
        If bIsPng Then
            pvIconFromResource = lvicPicTypePNG
        Else
            If Mode = exm_WantImage Then
                Dim cICO As New cFunctionsICO
                Dim cImageData As New cGDIpMultiImage
                If cICO.IsIconResource(cImageData, outData, True) = 0& Then Exit Function
                cImageData.CacheSourceInfo tSource, pvIconFromResource, 0&, False, True, True
                Erase outData()
                Set cImageData = Nothing
                Set cICO = Nothing
            Else
                pvIconFromResource = Mode
            End If
        End If
    Else
        m_BMPI.biHeight = Cy
    End If
ExitRoutine:
End Function

Private Function pvCursorFromResource(resPtr As Long, resSize As Long, outData() As Byte, Mode As ExtractionModeEnum) As Long

    ' function returns one of three things
    ' 1. Simply fills in the m_BMPI structure and returns nothing else
    ' 2. A GDI+ image handle of the selected cursor (PNGs are not created here but data returned to calling class)
    ' 3. An array containing the selected image in cursor format
    
    Dim lResult As Long, Cx As Long, Cy As Long
    Dim bIsPng As Boolean, tSource As Variant
    
    On Error GoTo ExitRoutine
    
    CopyMemory m_BMPI, ByVal resPtr, 40&
    If m_BMPI.biSize = png_Signature1 Then
        If m_BMPI.biWidth <> png_Signature2 Then Exit Function
        Cx = modCommon.ReverseLong(m_BMPI.biCompression)
        If Cx < 1& Or Cx > 256& Then Exit Function
        Cy = modCommon.ReverseLong(m_BMPI.biSizeImage)
        If Cy < 1& Or Cy > 256& Then Exit Function
        ' 24th byte from resource start will be the PNG bit depth
        ' 25th byte from resource start will be the color type
        CopyMemory m_BMPI.biBitCount, ByVal resPtr + 24, 2&
        Select Case (m_BMPI.biBitCount And &H7F00) \ &H100
            Case 2: If (m_BMPI.biBitCount And &HFF) = 8 Then m_BMPI.biBitCount = 24
            Case 6: If (m_BMPI.biBitCount And &HFF) = 8 Then m_BMPI.biBitCount = 32
            Case Else ' use reported depth
                m_BMPI.biBitCount = (m_BMPI.biBitCount And &HFF)
        End Select
        bIsPng = True
    End If
    
    If Not Mode = exm_WantAttr Then ' want entire resource
        If bIsPng And Mode = exm_WantImage Then
            ReDim outData(0 To resSize - 5&)
            CopyMemory outData(0), ByVal resPtr, resSize - 4&
        Else
            ReDim outData(0 To resSize + 17&)       ' exclude 4 byte hotspot header
            CopyMemory outData(22), ByVal resPtr + 4&, resSize - 4&
            outData(2) = 1
            outData(4) = 1
            CopyMemory outData(10), ByVal resPtr, 4&                ' copy hotspot info
            CopyMemory outData(14), CLng(resSize - 4&), 4&          ' size of resource
            outData(18) = 22
        End If
    End If
    If Not bIsPng Then
        If m_BMPI.biWidth < 1& Or m_BMPI.biWidth > 256& Then Exit Function
        Cx = m_BMPI.biWidth
        If m_BMPI.biHeight = 0& Or Abs(m_BMPI.biHeight) > 512& Then Exit Function
        Cy = m_BMPI.biHeight \ 2
        If Not Mode = exm_WantAttr Then
            If m_BMPI.biBitCount < 8 Then
                If m_BMPI.biBitCount = 1 Then outData(8) = 2 Else outData(8) = 16
            End If
        End If
    End If
    If Not Mode = exm_WantAttr Then
        If Cx < 256& Then outData(6) = Cx
        If Cy < 256& Then outData(7) = Cy
        If bIsPng = True Then
            pvCursorFromResource = lvicPicTypePNG
        Else
            If Mode = exm_WantImage Then
                If m_BMPI.biSize = png_Signature1 Then
                    pvCursorFromResource = lvicPicTypePNG
                Else
                    Dim cICO As New cFunctionsICO
                    Dim cImageData As New cGDIpMultiImage
                    If cICO.IsIconResource(cImageData, outData, True) = 0& Then Exit Function
                    cImageData.CacheSourceInfo tSource, pvCursorFromResource, 0&, False, True, True
                    Erase outData()
                    Set cImageData = Nothing
                    Set cICO = Nothing
                End If
            Else
                pvCursorFromResource = Mode
            End If
        End If
    Else
        m_BMPI.biHeight = Cy
    End If

ExitRoutine:
End Function

Private Function pvAVIfromResource(resPtr As Long, resSize As Long, outData() As Byte, Mode As ExtractionModeEnum) As Long

    ' function returns one of three things
    ' 1. Simply fills in the m_BMPI structure and returns nothing else
    ' 2. A GDI+ image handle of the 1st frame of the selected AVI
    ' 3. An array containing the selected AVI
    
    Dim tSA As SafeArray, bData() As Byte, bOK As Boolean
    Dim cImageData As New cGDIpMultiImage
    Dim cAVI As New cFunctionsAVI
    Dim fSize As RECTF, tSource As Variant
    
    On Error GoTo ExitRoutine
    
    If Mode = exm_WantData Then
        ReDim outData(0 To resSize - 1&)
        CopyMemory outData(0), ByVal resPtr, resSize
        pvAVIfromResource = Mode
    Else
        With tSA
            .cbElements = 1
            .cDims = 1
            .pvData = resPtr
            .rgSABound(0).cElements = resSize
        End With
        CopyMemory ByVal VarPtrArray(bData()), VarPtr(tSA), 4&
        bOK = cAVI.IsAVIResource(bData(), cImageData)
        CopyMemory ByVal VarPtrArray(bData()), 0&, 4&
        tSA.pvData = 0&
        If bOK Then
            GdipGetImageBounds cImageData.Handle, fSize, UnitPixel
            m_BMPI.biWidth = fSize.nWidth
            m_BMPI.biHeight = fSize.nHeight
            GdipGetImagePixelFormat cImageData.Handle, m_BMPI.biSize
            m_BMPI.biBitCount = (m_BMPI.biSize And &HFF00&) \ &H100&
            If Mode = exm_WantImage Then
                cImageData.CacheSourceInfo tSource, pvAVIfromResource, 0&, False, True, True
            End If
        End If
        Set cImageData = Nothing
        Set cAVI = Nothing
    End If
ExitRoutine:
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(bData()), 0&, 4&
End Function

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
End Sub

Private Sub Class_Terminate()
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub
