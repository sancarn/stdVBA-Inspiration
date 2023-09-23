VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write ICO/CUR image format ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' This class contains icon functions & basically does 100% parsing of an icon resource & its images
' See comments in GDIpImage class. Icons are handled differently than other image formats
' Class also parses animated cursors & is written to handle PNGs in icon and cursor files
' and has functions to save to icon/cursor format

' Icons/cursors can come in 4 flavors: normal icon/cursor, animated cursor & Win7 animated cursor, DLL
' Animated cursors use a RIFF format and Win7 complicates it be adding groups of cursors within the format

' So to enable this routine to handle all icons/cursors similarly, I use a header format that looks like:
'   4 byte image/group type & count
'       &HF.......
'           low bit set = cursor else icon
'           bit 2 set = anmimated cursor (RIFF format)
'           bit 3 set = grouped (RIFF format)
'       &H.FFF.... = Nr images. If RIFF then nr frames per group. 4095 max icons per resource
'       &H....FFFF = Nr groups. If not RIFF then will be zero. Max 32767 groups
'   4 byte current Index (byte offset of 4)
'   16 bytes of BitmapInfoHeader that describes the current image (byte offset of 8)
'   4 byte offset, per image, into the resource where icon header starts. (byte offset of 24)

' About binaries
' - The source file is never cached. These could be system DLLs, executables, & OCXs.
' - Whenever image from binary is saved, only the current image is saved, never the entire binary
' - The image itself is only processed/extracted on demand
' XP Shell32.dll for example as over 1400 icons & Vista has over 2000 & Win7 over 3000

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type BITMAPINFO     ' 1bpp header
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1 To 2) As Long
End Type

Private Type ICONDIRENTRY   ' ICONS                     CURSORS
    bWidth As Byte          ' width 256=0               width 256=0
    bHeight As Byte         ' height 256=0              height 256=0
    bColorCount As Byte     ' > 8bpp = 0                > 8bpp = 0
    bReserved As Byte
    wPlanes As Integer      ' must be 1                 X hot spot
    wBitCount As Integer    ' bit depth                 Y hot spot
    dwBytesInRes As Long    ' nr bytes in diretory      nr bytes in diretory
    dwImageOffset As Long   ' offset to 1st byte        offset to 1st byte
End Type
Private Type ICONDIR
    idReserved As Integer
    idType As Integer       ' value is 1                value is 2
    idCount As Integer      ' nr idEntries              nr idEntries
    idEntries() As ICONDIRENTRY
End Type
Private m_MultiDestData() As Byte   ' used when creating icons from multiple images
Private m_Header() As Byte

Public Function IconSizes(sourceData As Variant, ByVal IconIndex As Long, ByVal GroupIndex As Long, _
                            Width As Single, Height As Single) As Boolean

    ' function returns the size of the requested icon/cursor at IconIndex
    
    Dim icde As ICONDIRENTRY
    Dim lOffset As Long, lValue As Long
    Dim inArray() As Byte, imgArray() As Byte
    
    If IconIndex < 0& Then Exit Function
    
    CopyMemory lValue, m_Header(4), 4&
    If lValue = IconIndex Then
        CopyMemory lValue, m_Header(12), 4& ' width
        Width = lValue
        CopyMemory lValue, m_Header(16), 4& ' height
        Height = lValue
    Else
        modCommon.MoveArrayToVariant sourceData, inArray(), False
        lOffset = pvLocateDirEntry(inArray(), IconIndex, GroupIndex, 0&)
        If lOffset Then
            CopyMemory icde, inArray(lOffset), 16&                      ' cache icon directory entry
            Width = icde.bWidth: If Width = 0& Then Width = 256&        ' set width/height for this icon/cursor
            Height = icde.bHeight: If Height = 0& Then Height = 256&
        End If
        modCommon.MoveArrayToVariant sourceData, inArray(), True
    End If
    IconSizes = True
    
End Function

Public Property Get ImageCount() As Long
    On Error Resume Next
    ImageCount = m_Header(2) Or (m_Header(3) And &HF) * &H100&
End Property

Public Property Get GroupCount() As Long
    On Error Resume Next
    GroupCount = m_Header(0) Or m_Header(1) * &H100&
End Property

Public Property Get ImageFormat() As ImageFormatEnum
    On Error GoTo ExitRoutine
    If (m_Header(3) And &H10) Then ' cursor of some sort
        If (m_Header(3) And &H20) Then ' animated vs not
            ImageFormat = lvicPicTypeAnimatedCursor
        Else
            ImageFormat = lvicPicTypeCursor
        End If
    Else
        ImageFormat = lvicPicTypeIcon
    End If
ExitRoutine:
End Property


Public Function IconBitDepth(sourceData As Variant, ByVal IconIndex As Long, ByVal GroupIndex As Long) As Long

    ' function returns the bit depth of the requested icon/cursor at IconIndex

    Dim icde As ICONDIRENTRY
    Dim lOffset As Long, lValue As Long
    Dim inArray() As Byte, imgData() As Byte
    
    If IconIndex < 0& Then Exit Function
    
    CopyMemory lValue, m_Header(4), 4&
    If lValue = IconIndex Then
        IconBitDepth = m_Header(22)
    Else
        modCommon.MoveArrayToVariant sourceData, inArray(), False
        lOffset = pvLocateDirEntry(inArray(), IconIndex, GroupIndex, lValue)
        If lOffset Then
            CopyMemory icde, inArray(lOffset), 16&                          ' cache icon directory entry
            If Not (icde.wBitCount = 0 Or ((m_Header(3) And &H10)) = &H10) Then
                IconBitDepth = icde.wBitCount                               ' right or wrong, report what was provided for icons
            Else                                                            ' for missing bitcount or cursors, get from headers
                IconBitDepth = pvGetBitDepthFromBMPheader(inArray(), icde.dwImageOffset + lValue)
            End If
        End If
        modCommon.MoveArrayToVariant sourceData, inArray(), True
    End If
    
End Function

Public Function IconToBits(sourceData As Variant, ByVal IconIndex As Long, ByVal GroupIndex As Long, _
                            outArray() As Byte, Width As Single, Height As Single, newIndex As Long, newCount As Long) As Long
    
    ' function fills in the last 4 parameters
    ' IconIndex: zero-bound index of the icon/cursor to retrieve
    '   If less than zero, best match for system size & highest color depth will be retrieved
    '   If less than zero, newIndex will contain the index of the icon processed else left unchanged
    ' outArray(): If PNG, outArray() is entire PNG image format
    ' Width & Height: size of icon at IconIndex

    ' Return value is one of the following
    '   lvicPicTypeNone = failed to process inArray()
    '   lvicPicTypePNG = outArray() filled, complete PNG format for use with GdipLoadImageFromStream
    '   other = GDI+ image handle
    
    Dim BMI As BITMAPINFOHEADER, icde As ICONDIRENTRY, tSA As SafeArray
    Dim tSApal As SafeArray, aPal() As Long
    Dim X As Long, Y As Long, lOffset As Long, lHdrOffset As Long
    Dim bitAnd As Long, bAlpha As Boolean
    Dim srcPtr As Long, dstPtr As Long
    Dim scanWidth As Long, dstScanWidth As Long
    Dim palIndex As Long, lMaskPtr As Long, mskScanWidth As Long
    Dim lShift As Long, rShift As Long, lMaskShift As Long
    Dim inArray() As Byte, pot16() As Long
    Dim imgData() As Byte, tImage As GDIpImage, IStream As IUnknown
    Dim tSize As RECTI, tBMP As BitmapData, cPal As ColorPalette
    Const BI_BITFIELDS As Long = 3&

    On Error GoTo ExitRoutine
    modCommon.MoveArrayToVariant sourceData, inArray(), False
    
    X = m_Header(2) Or (m_Header(3) And &HF) * &H100&                         ' icon count
    If IconIndex < 0& Then                                                  ' caller wants best system match for non-grouped resource
        IconIndex = pvGetBestSystemMatch(inArray(), X)
        newIndex = IconIndex + 1&
    ElseIf GroupIndex < 0& Then                                             ' caller wants best system match for grouped resource
        GroupIndex = pvGetBestSystemMatch(inArray(), X) + 1&
        newIndex = GroupIndex
    ElseIf IconIndex >= X Then                                              ' ensure index (0-bound) does not exceed count
        GoTo ExitRoutine
    ElseIf GroupIndex > (m_Header(0) Or m_Header(1) * &H100&) Then            ' ensure group (1-bound) does not exceed count
        GoTo ExitRoutine
    End If
    
    lOffset = pvLocateDirEntry(inArray(), IconIndex, GroupIndex, lHdrOffset) ' location of the group's icon's directory entry
    If lOffset = 0& Then GoTo ExitRoutine
    
    CopyMemory icde, inArray(lOffset), 16&                              ' cache icon directory entry
    CopyMemory BMI, inArray(icde.dwImageOffset + lHdrOffset), 24&       ' get bitmap info header
    CopyMemory m_Header(8), BMI, 16&                                    ' store image description
    CopyMemory m_Header(4), IconIndex, 4&                               ' cache current index
    
    Width = icde.bWidth: If Width = 0& Then Width = 256&                    ' set width/height for this icon/cursor
    Height = icde.bHeight: If Height = 0& Then Height = 256&
    
    If BMI.biSize = png_Signature1 Then
        If BMI.biWidth = png_Signature2 Then                                ' png image
            ReDim outArray(0 To icde.dwBytesInRes - 1&)                     ' resize array & fill it in
            CopyMemory outArray(0), inArray(icde.dwImageOffset + lHdrOffset), icde.dwBytesInRes
            BMI.biBitCount = pvGetBitDepthFromBMPheader(inArray(), icde.dwImageOffset + lHdrOffset)
            BMI.biWidth = Width: BMI.biHeight = Height
            CopyMemory m_Header(8), BMI, 16&                                ' store image description
            IconToBits = lvicPicTypePNG                                     ' done
            GoTo ExitRoutine
        End If
    End If
    
    scanWidth = modCommon.ByteAlignOnWord(BMI.biBitCount, Width)                      ' cache nr bytes per scan line
    mskScanWidth = modCommon.ByteAlignOnWord(1, Width)
    If BMI.biBitCount <= 8 Then
        If BMI.biClrUsed = 0& Then BMI.biClrUsed = 2 ^ BMI.biBitCount
    End If
    lOffset = (scanWidth + mskScanWidth) * Height              ' determine start of pixel data
    lOffset = icde.dwImageOffset + icde.dwBytesInRes - lOffset + lHdrOffset '   calculated from end of icon resource
        
    ' determine transparency
    If BMI.biBitCount = 32 Then
        bAlpha = True
    Else
        lHdrOffset = lOffset + scanWidth * Height                           ' where mask starts in source data
        For Y = 0& To Height - 1&
            lShift = 128
            srcPtr = Y * mskScanWidth + lHdrOffset                          ' where current mask row starts
            For X = 0& To Width - 1&                                        ' make pixel opaque if needed
                If (inArray(srcPtr) And lShift) Then
                    bAlpha = True
                    Y = Height
                    Exit For
                End If
                If lShift = 1& Then lShift = 128 Else lShift = lShift \ 2&
            Next
        Next
    End If
    
    Select Case BMI.biBitCount
        Case 1: X = lvicColor1bpp: cPal.Count = BMI.biClrUsed
            lShift = &H2: bitAnd = &H1
        Case 4: X = lvicColor4bpp: cPal.Count = BMI.biClrUsed
            lShift = &H10: bitAnd = &HF
        Case 8: X = lvicColor8bpp: cPal.Count = BMI.biClrUsed
            lShift = &H1: bitAnd = &HFF
        Case 16, 24: X = lvicColor24bpp
        Case 32: X = lvicColor32bpp
        Case Else
            GoTo ExitRoutine
    End Select
    If cPal.Count Then  ' paletted. Overlay Long array on entries to increase access speed vs CopyMemory
        With tSApal
            .pvData = VarPtr(inArray(lOffset - cPal.Count * 4&))    ' where palette starts
            .cbElements = 4                                         ' Long vs. byte array
            .cDims = 1                                              ' 1D array
            .rgSABound(0).cElements = cPal.Count                    ' number of Longs in array
        End With
    End If
    If bAlpha Then X = lvicColor32bppAlpha: cPal.Count = 0&         ' flag used at end of routine
    
    If GdipCreateBitmapFromScan0(Width, Height, 0&, X, ByVal 0&, IconToBits) Then GoTo ExitRoutine
    tSize.nHeight = Height: tSize.nWidth = Width
    If GdipBitmapLockBits(IconToBits, tSize, ImageLockModeWrite, X, tBMP) Then
        GdipDisposeImage IconToBits
        IconToBits = 0&
        GoTo ExitRoutine
    End If
    With tSA
        .cbElements = 1
        .cDims = 1
        .pvData = tBMP.Scan0Ptr
        .rgSABound(0).cElements = tBMP.Height * tBMP.stride
    End With
    Erase outArray()
    CopyMemory ByVal VarPtrArray(outArray), VarPtr(tSA), 4&         ' overlay arrays on image bits & palette
    If tSApal.pvData Then CopyMemory ByVal VarPtrArray(aPal), VarPtr(tSApal), 4&
    
    With BMI
        dstScanWidth = tBMP.stride                                          ' cache nr bytes in 32bpp scan line
        Select Case .biBitCount
        
        Case 32                                                             ' no mask processing here
            For Y = 0& To Height - 1&
                srcPtr = Y * scanWidth + lOffset                            ' where current source row starts
                dstPtr = dstScanWidth * (Height - Y - 1&)                   ' where current destinatino row starts
                If bAlpha Then
                    CopyMemory outArray(dstPtr), inArray(srcPtr), dstScanWidth
                Else
                    For dstPtr = dstPtr To dstPtr + dstScanWidth - 1& Step 4&
                        CopyMemory outArray(dstPtr), inArray(srcPtr), 3&
                        srcPtr = srcPtr + 4&                                ' move source pointer along
                    Next
                End If
            Next
            
        Case 8, 4, 1                                                        ' paletted icons
            lHdrOffset = lOffset + scanWidth * Height
            For Y = 0& To Height - 1&
                If lShift = &H2 Then rShift = &H80 Else rShift = lShift
                srcPtr = Y * scanWidth + lOffset                            ' where current source row starts
                dstPtr = dstScanWidth * (Height - Y - 1&)                   ' where current destinatino row starts
                lMaskShift = &H80
                lMaskPtr = lHdrOffset + Y * mskScanWidth                    ' where AND mask starts
                If bAlpha Then
                    For X = 0& To Width - 1&                                ' get palette index & color location from packed byte
                        palIndex = ((inArray(srcPtr) \ rShift) And bitAnd)
                        If (inArray(lMaskPtr) And lMaskShift) = 0 Then      ' pixel is fully opaque
                            CopyMemory outArray(dstPtr), aPal(palIndex) Or &HFF000000, 4&
                        ElseIf aPal(palIndex) = 0& Then                     ' pixel is transparent
                            CopyMemory outArray(dstPtr), aPal(palIndex), 3&
                        Else                                                ' pixel is XOR'd with bkg
                            CopyMemory outArray(dstPtr), &HFF000000 Or (vbWhite Xor aPal(palIndex)), 4&
                        End If
                        dstPtr = dstPtr + 4&                                ' move target pointer along
                        If rShift = 1& Then
                            srcPtr = srcPtr + 1&
                            If lShift = &H2 Then rShift = &H80 Else rShift = lShift
                        Else
                            rShift = rShift \ lShift
                        End If
                        If lMaskShift = 1& Then                             ' move AND mask pointer along
                            lMaskShift = &H80: lMaskPtr = lMaskPtr + 1&
                        Else
                            lMaskShift = lMaskShift \ 2&
                        End If
                    Next
                Else
                    CopyMemory outArray(dstPtr), inArray(srcPtr), dstScanWidth
                End If
            Next
            
        Case 24                                                             ' 24bpp icons
            lHdrOffset = lOffset + scanWidth * Height
            For Y = 0& To Height - 1&
                srcPtr = Y * scanWidth + lOffset                            ' where current source row starts
                dstPtr = dstScanWidth * (Height - Y - 1&)                   ' where current destinatino row starts
                lMaskShift = &H80
                lMaskPtr = lHdrOffset + Y * mskScanWidth
                If bAlpha Then
                    For dstPtr = dstPtr To dstPtr + dstScanWidth - 1& Step 4&
                        If (inArray(lMaskPtr) And lMaskShift) = 0 Then      ' see similar comments in previous section
                            CopyMemory outArray(dstPtr), inArray(srcPtr), 4&
                            outArray(dstPtr + 3&) = 255
                        Else
                            CopyMemory palIndex, inArray(srcPtr), 3&
                            If palIndex = 0& Then
                                CopyMemory outArray(dstPtr), palIndex, 3&
                            Else
                                CopyMemory outArray(dstPtr), ((vbWhite Xor palIndex) Or &HFF000000), 4&
                            End If
                        End If
                        srcPtr = srcPtr + 3&                                ' move source pointer along
                        If lMaskShift = 1& Then
                            lMaskShift = &H80: lMaskPtr = lMaskPtr + 1&
                        Else
                            lMaskShift = lMaskShift \ 2&
                        End If
                    Next
                Else
                    CopyMemory outArray(dstPtr), inArray(srcPtr), dstScanWidth
                End If
            Next
        Case 16                                                             ' 16bpp icons, assume 555 format
        
            lHdrOffset = lOffset + scanWidth * Height
            ReDim pot16(0 To 31)
            For X = 0& To 31&
                pot16(X) = X * 8& + (X Mod 8&)
            Next
            For Y = 0& To Height - 1&
                srcPtr = Y * scanWidth + lOffset                            ' where current source row starts
                dstPtr = dstScanWidth * (Height - Y - 1&)                   ' where current destinatino row starts
                lMaskShift = &H80
                lMaskPtr = lHdrOffset + Y * mskScanWidth
                For X = 0& To Width * 2& - 1& Step 2&
                    bitAnd = (inArray(srcPtr + X + 1) And &H7F) * &H100& Or inArray(srcPtr + X)
                    ' shift the 5,5,5 and ref a 5 bit/32 entry palette
                    palIndex = pot16((bitAnd And &H1F)) Or _
                        pot16(((bitAnd \ &H20) And &H1F)) * &H100& Or _
                        pot16((bitAnd \ &H400) And &H1F) * &H10000
                    If bAlpha Then
                        If (inArray(lMaskPtr) And lMaskShift) = 0 Then      ' see similar comments in previous section
                            CopyMemory outArray(dstPtr), palIndex Or &HFF000000, 4&
                        Else
                            If palIndex = 0& Then
                                CopyMemory outArray(dstPtr), palIndex, 3&
                            Else
                                outArray(dstPtr + 3&) = 255
                                CopyMemory outArray(dstPtr), (vbWhite Xor palIndex) Or &HFF000000, 4&
                            End If
                        End If
                        dstPtr = dstPtr + 4&
                        If lMaskShift = 1& Then
                            lMaskShift = &H80: lMaskPtr = lMaskPtr + 1&
                        Else
                            lMaskShift = lMaskShift \ 2&
                        End If
                    Else
                        CopyMemory outArray(dstPtr), palIndex, 3&
                        dstPtr = dstPtr + 3&
                    End If
                Next
            Next
            Erase pot16()
            
        Case Else
            GoTo ExitRoutine
        End Select
        
    End With
    
ExitRoutine:
    If tSApal.pvData Then CopyMemory ByVal VarPtrArray(aPal), 0&, 4&
   If tSA.pvData Then
        GdipBitmapUnlockBits IconToBits, tBMP
        CopyMemory ByVal VarPtrArray(outArray), 0&, 4&
        If Err Then
            GdipDisposeImage IconToBits
            IconToBits = 0&
        ElseIf cPal.Count Then  ' then no alpha & paletted; ensure palette is fully opaque
            palIndex = lOffset - BMI.biClrUsed * 4&
            For X = 1& To cPal.Count
                cPal.Entries(X) = inArray(palIndex) Or inArray(palIndex + 1&) * &H100& Or inArray(palIndex + 2&) * &H10000 Or &HFF000000
                palIndex = palIndex + 4&
            Next
            GdipSetImagePalette IconToBits, cPal
        End If
    End If
    modCommon.MoveArrayToVariant sourceData, inArray(), True
    
End Function

Public Function IsIconResourceFile(cImageData As cGDIpMultiImage, FileHandle As Long, CacheData As Boolean, Optional FileName As String) As Boolean

    ' Routine quickly scans a file to determine if it contains icon resources
    ' This initial check is only a simple check, other routines will process data completely

    Dim icd As ICONDIR, icde() As ICONDIRENTRY, BMI As BITMAPINFOHEADER
    Dim lNeeded As Long, lMax As Long, lPtr As Long, lRead As Long
    
    Const aniRIFF As Long = 1179011410                  ' magic number for animated cursors
    Const aniACON As Long = 1313817409 ' flag indicating animated icon
    
'    If Left$(FileName, 4) = "*DLL" Then
'        IsIconResourceFile = pvScanBinary(cImageData, Mid$(FileName, 5), CacheData)
'    Else
        lMax = GetFileSize(FileHandle, 0&)
        If lMax < 57& Then Exit Function                     ' ensure enough bytes for the ICONDIRECTORY
        
        SetFilePointer FileHandle, 0&, 0&, 0&: ReadFile FileHandle, lNeeded, 4&, lRead, ByVal 0&
        If lNeeded = aniRIFF Then
            ReadFile FileHandle, lNeeded, 4&, lRead, ByVal 0&
            ReadFile FileHandle, lNeeded, 4&, lRead, ByVal 0&
            IsIconResourceFile = (lNeeded = aniACON)
            Exit Function
        Else
            SetFilePointer FileHandle, 0&, 0&, 0&: ReadFile FileHandle, icd, 6&, lRead, ByVal 0& ' cache the ICONDIRECTORY
            If icd.idCount < 1 Then                         ' no icons or not an icon file
                lRead = 0&
            ElseIf icd.idType < 1 Or icd.idType > 2 Then    ' per MSDN, must be 1 or 2 (1=icon,2=cursor)
                lRead = 0&
            End If
        End If
        
        If lRead Then
            lPtr = 6&                                           ' move array pointer
            lNeeded = 16&                                       ' length of directory entry
            If lPtr + lNeeded * icd.idCount <= lMax Then
                ReDim icde(1 To icd.idCount)                    ' size our entries
                lNeeded = icd.idCount * lNeeded                 ' & cache them
                SetFilePointer FileHandle, lPtr, 0&, 0&: ReadFile FileHandle, icde(1), lNeeded, lRead, ByVal 0&
                lNeeded = lNeeded + 6&                          ' move array pointer
                For lPtr = 1 To icd.idCount
                    ' tally total bytes used by directories & ensure no directory overflows into another
                    lNeeded = lNeeded + icde(lPtr).dwBytesInRes
                Next
                If lNeeded <= lMax Then IsIconResourceFile = True
            End If
        End If
'    End If
    
End Function

Public Function IsIconResource(cImageData As cGDIpMultiImage, inArray() As Byte, CacheData As Boolean) As Long

    ' Icons have no image format magic number. So we must parse the file
    ' If any major icon format inconsistencies exist, assume file is not an icon file
    ' This routine will spawn another if format is an animated cursor
    
    ' Maybe this validation routine is too restrictive, but better to abort malformatted icons,
    '   than to let it thru & crash from invalid values/offsets.
    ' To include multiple icon/cursor resources when 1 or more corrupt images may exist, the routine
    '   will simply ignore corrupted ones
    
    Dim icd As ICONDIR, icde As ICONDIRENTRY, ICDirXRef() As Long, BMI As BITMAPINFOHEADER
    Dim lSize As Long, lMax As Long, lPtr As Long, icType As Long, lNeeded As Long
    Dim X As Long, Y As Long, tArray() As Byte, tSource As Variant, lResOffset As Long
    
    ' animated cursor related variables:
    Dim seqPtr As Long, durPtr As Long, nrJiffies As Long, durSeq() As Long
    Const aniRIFF As Long = 1179011410 ' magic number for animated cursors
    
    lMax = UBound(inArray) + 1&
    If lMax < 57& Then Exit Function                    ' ensure enough bytes for at last 1 icon
    
    CopyMemory X, inArray(lPtr), 4&
    If X = aniRIFF Then
        ' pass off to helper function to parse & fill in the icDir & IcDirE() structures
        If pvConvertAniCursor(inArray(), ICDirXRef(), lMax, seqPtr, durPtr, nrJiffies) = False Then Exit Function
        icType = lvicPicTypeAnimatedCursor
    Else
        CopyMemory icd, inArray(lPtr), 6&               ' cache the ICONDIRECTORY
        If icd.idCount < 1 Then Exit Function           ' no icons or not an icon file
        If icd.idType < 1 Or icd.idType > 2 Then Exit Function ' per MSDN, must be 1 or 2 (1=icon,2=cursor)
        If icd.idCount > &HFFF Then Exit Function       ' not designed to handle 4000+ icons
        If icd.idCount * 16& + 6& > lMax Then Exit Function ' not enough bytes in source
        
        If icd.idType = 1 Then icType = lvicPicTypeIcon Else icType = lvicPicTypeCursor
        ReDim ICDirXRef(0 To icd.idCount)
        ICDirXRef(0) = 1& Or icd.idCount * &H10000 Or (icd.idType - 1) * &H10000000
        lPtr = 6&
        For X = 1& To icd.idCount
            ICDirXRef(X) = lPtr + (X - 1&) * 16&        ' icon dir entry position
        Next
    End If
    
    lSize = UBound(ICDirXRef) * 4&                      ' include 4 bytes per entry + 4 byte index + 16 byte BitmapInfoHeader
    For Y = 1& To (ICDirXRef(0) And &HFFFF&)            ' loop thru groups
        icd.idCount = (ICDirXRef(0) And &HFFF0000) \ &H10000
        For X = 0& To icd.idCount - 1                   ' loop thru group icons
            lPtr = (Y - 1&) * icd.idCount + X + 1&      ' locate header offset
            
            ' get the icon directory entry for further validations
            ' and cache the array location to icd.idReserved should icon directory need to be modified
            If icType = lvicPicTypeAnimatedCursor Then  ' RIFF format
                If (ICDirXRef(0) And &H40000000) Then   ' grouped RIFF (Win7+)
                    lResOffset = ICDirXRef(lPtr) + (Y - 1&) * 16&
                Else
                    lResOffset = ICDirXRef(lPtr)
                End If                                  ' calculate offset in relation to raw data
                CopyMemory icde, inArray(lResOffset), 16&
                icde.dwImageOffset = icde.dwImageOffset + ICDirXRef(lPtr) - 6&
            Else
                lResOffset = ICDirXRef(lPtr)
                CopyMemory icde, inArray(lResOffset), 16&
            End If
            
            If icde.dwBytesInRes < 1& Then
                ICDirXRef(lPtr) = -1&                  ' don't use this one
            ElseIf icde.dwImageOffset + icde.dwBytesInRes > lMax Then
                ICDirXRef(lPtr) = -1&                  ' don't use this one
            Else
                ' validate actual resource size within reported resource size
                CopyMemory BMI, inArray(icde.dwImageOffset), 24&  ' only need 24 bytes of bitmapinfoheader for this test
                Select Case BMI.biBitCount                        ' paletted icons; calculate maximum pixel info size
                    Case 1, 2, 4, 8
                        If pvValidateDimensions(BMI.biWidth, BMI.biCompression) = False Then
                            ICDirXRef(lPtr) = 0&      ' don't use this one
                        ElseIf pvValidateDimensions(BMI.biHeight \ 2&, BMI.biCompression) = False Then
                            ICDirXRef(lPtr) = 0&      ' don't use this one
                        End If
                        If ICDirXRef(lPtr) Then
                            BMI.biSizeImage = (modCommon.ByteAlignOnWord(BMI.biBitCount, BMI.biWidth) + modCommon.ByteAlignOnWord(1, BMI.biWidth)) * (BMI.biHeight \ 2)
                            lNeeded = (2 ^ BMI.biBitCount) * 4& + BMI.biSize
                            If BMI.biSizeImage + lNeeded > icde.dwBytesInRes Then ICDirXRef(lPtr) = 0& ' don't use this one
                        End If
                    Case 16, 24, 32                               ' non-paletted icons; calculate maximum pixel info size
                        If pvValidateDimensions(BMI.biWidth, BMI.biCompression) = False Then
                            ICDirXRef(lPtr) = 0&      ' don't use this one
                        ElseIf pvValidateDimensions(BMI.biHeight \ 2&, BMI.biCompression) = False Then
                            ICDirXRef(lPtr) = 0&      ' don't use this one
                        End If
                        If ICDirXRef(lPtr) Then
                            BMI.biSizeImage = (modCommon.ByteAlignOnWord(BMI.biBitCount, BMI.biWidth) + modCommon.ByteAlignOnWord(1, BMI.biWidth)) * (BMI.biHeight \ 2)
                            If BMI.biSizeImage + BMI.biSize > icde.dwBytesInRes Then ICDirXRef(lPtr) = 0& ' don't use this one
                        End If
                    Case Else                           ' if PNG, assume formatted correctly
                        If Not (BMI.biSize = png_Signature1 And BMI.biWidth = png_Signature2) Then
                            ICDirXRef(lPtr) = 0& ' don't use this one
                        Else
                            CopyMemory BMI.biWidth, inArray(icde.dwImageOffset + 16&), 4&
                            CopyMemory BMI.biHeight, inArray(icde.dwImageOffset + 20&), 4&
                            BMI.biWidth = modCommon.ReverseLong(BMI.biWidth)
                            BMI.biHeight = modCommon.ReverseLong(BMI.biHeight)
                            If pvValidateDimensions(BMI.biWidth, 0&) = False Then
                                ICDirXRef(lPtr) = 0&      ' don't use this one
                            ElseIf pvValidateDimensions(BMI.biHeight, 0&) = False Then
                                ICDirXRef(lPtr) = 0&      ' don't use this one
                            Else
                                BMI.biHeight = BMI.biHeight * 2&
                            End If
                        End If
                End Select
            End If
            If ICDirXRef(lPtr) = 0& Then
                icd.idCount = icd.idCount - 1&              ' reduce group icon count
            Else
                ' validate icon directory entry dimensions with actual bitmap dimensions
                If icde.bWidth <> BMI.biWidth Then
                    If Not (icde.bWidth = 0 And BMI.biWidth = 256&) Then
                        If BMI.biWidth = 256& Then inArray(lResOffset) = 0 Else inArray(lResOffset) = BMI.biWidth
                    End If
                End If
                If icde.bHeight <> BMI.biHeight \ 2 Then
                    If Not (icde.bHeight = 0 And BMI.biHeight = 512&) Then
                        If BMI.biHeight = 512& Then inArray(lResOffset + 1&) = 0 Else inArray(lResOffset + 1&) = BMI.biHeight \ 2
                    End If
                End If
                'ICDirXRef(lPtr) = ICDirXRef(lPtr) + lSize   ' adjust in relation to header vs source
            End If
        Next
        If icd.idCount = 0& Then Exit Function              ' couldn't use any icons within a group; assume more corrupted
    Next

    
    ReDim m_Header(0 To lSize + 23&)
    CopyMemory m_Header(24), ICDirXRef(1), lSize
    CopyMemory m_Header(0), ICDirXRef(0), 4&
    m_Header(7) = &H80 'force current index to be a negative value; i.e., no image currently selected
    
    icd.idCount = (ICDirXRef(0) And &HFFF0000) \ &H10000
    cImageData.InitializeContent ObjPtr(Me), icType, (icd.idCount)
    If icType = lvicPicTypeAnimatedCursor Then              ' supply duration & sequence from RIFF if provided
        If seqPtr Then
            'CopyMemory lMax, inArray(seqPtr + lSize), 4&    ' get sequence byte length
            CopyMemory lMax, inArray(seqPtr), 4&     ' get sequence byte length
            If lMax < 4& Then
                seqPtr = 0&
            Else
                lMax = lMax \ 4&
                ReDim ICDirXRef(1 To lMax)                  ' size temporary array to retrieve values
                ReDim durSeq(0 To 1, 1 To lMax)
                'CopyMemory ICDirXRef(1), inArray(seqPtr + lSize + 4&), lMax * 4&
                CopyMemory ICDirXRef(1), inArray(seqPtr + 4&), lMax * 4&
                For lPtr = 1& To lMax                       ' transfer sequence to our array
                    durSeq(0, lPtr) = ICDirXRef(lPtr)
                Next
            End If
        End If
        If durPtr Then                              ' get duration byte length
            'CopyMemory lMax, inArray(durPtr + lSize), 4&
            CopyMemory lMax, inArray(durPtr), 4&
            If lMax < 4& Then
                durPtr = 0&
            Else
                lMax = lMax \ 4&
                If seqPtr = 0& Then                 ' size array & ensure duration/sequence same size
                    ReDim durSeq(0 To 1, 1 To lMax)
                Else
                    If lMax > UBound(durSeq, 2) Then lMax = UBound(durSeq, 2)
                End If
                ReDim ICDirXRef(1 To lMax)          ' transfer sequence to our array
                'CopyMemory ICDirXRef(1), inArray(durPtr + lSize + 4&), lMax * 4&
                CopyMemory ICDirXRef(1), inArray(durPtr + 4&), lMax * 4&
                For lPtr = 1 To lMax                ' convert duration from jiffies to milliseconds
                    durSeq(1, lPtr) = ((ICDirXRef(lPtr) And &HFFFF&) * 1000&) \ 60&
                Next
            End If
        End If
    End If
    Erase ICDirXRef()
    
    ' for all icons/cursors/animated cursors, a sequence/duration array is used for comformity
    ' Setup default values if not already filled in from RIFF info above
    If seqPtr = 0& Then
        X = 0&                                              ' default sequence values
        If durPtr = 0& Then ReDim durSeq(0 To 1, 1 To icd.idCount)
        For lPtr = 1& To UBound(durSeq, 2)
            durSeq(0, lPtr) = X
            X = X + 1&
            If X = icd.idCount Then X = 0&                  ' ensure don't exceed icon count
        Next
        seqPtr = -1&                                        ' flag for next loop
    End If
    If durPtr = 0& And nrJiffies > 0& Then                  ' default duration values
        If seqPtr = 0& Then ReDim durSeq(0 To 1, 1 To icd.idCount)
        For lPtr = 1& To icd.idCount                        ' though not duration RIFF block exists
            durSeq(1, lPtr) = nrJiffies                     ' the RIFF did provide a default per-frame jiffies value
        Next
    End If
    cImageData.SetSequenceDuration durSeq()            ' send off & done
    
    modCommon.MoveArrayToVariant tSource, inArray, True
    cImageData.SetIndex tSource, 0&, &HF0000000
    modCommon.MoveArrayToVariant tSource, inArray, False
    If cImageData.Handle = 0& Then
        Set cImageData = New cGDIpMultiImage
    Else
        cImageData.CacheSourceInfo VarPtrArray(inArray), 0&, 0&, CacheData, False
        IsIconResource = icType
    End If
        
    
End Function

Private Function pvValidateDimensions(BmpSize As Long, Compression As Long) As Boolean

    ' helper function when parsing icons
    
    Dim bOK As Boolean
    If Compression = 0& Then                        ' compression must be zero else don't include icon
        bOK = Not (BmpSize < 1& Or BmpSize > 256&)  ' size must be between 1 nad 256 else don't include icon
    End If
    pvValidateDimensions = bOK

End Function

Private Function pvGetBestSystemMatch(iconBits() As Byte, Count As Long) As Long

' Use similar logic that Windows uses. http://msdn.microsoft.com/en-us/library/ms997538.aspx
'   1. The image closest in size to the requested size is chosen.
'   2. If two or more images of that size are present, the one that matches the color depth of the display is chosen.
'   3. If none exactly match the color depth of the display, Windows chooses the image with the greatest color depth without exceeding the color depth of the display.
'   4. If all the size-matched images exceed the color depth of the display, the one with the lowest color depth is chosen.
'   5. Windows treats all color depths of 8 or more bpp as equal. For example, it is pointless to have a 16x16 256 color image and a 16x16 16bpp image in the same resource—Windows will simply choose the first one it encounters.
'   6. When the display is in 8bpp mode, Windows will prefer a 16 color icon over a 256 color icon, and will display all icons using the system default palette.

'   Exception to rules above. Since this project is for displaying icons, not creating icons, the following rules apply
'   1. The image closest in size to the requested size is chosen.
'   2. If two or more images of that size are present, the one with the image with the greatest color depth

' FYI: Size matching is done by true distance from corner to corner.

    Dim icoCX As Long, icoCY As Long, I As Long, lBest As Long
    Dim dTarget As Currency, dSource As Currency, dSizes() As Currency
    Dim Index() As Long, icde() As ICONDIRENTRY, isCursor As Boolean
    Dim lHdrSize As Long
    
    Const SM_CXICON As Long = 11&
    Const SM_CYICON As Long = 12&
    
    If Count = 1& Then Exit Function
    ReDim icde(1 To Count)
    For I = 1 To Count
        icoCX = pvLocateDirEntry(iconBits(), I - 1&, 1&, lHdrSize)
        If icoCX Then CopyMemory icde(I), iconBits(icoCX), 16&
    Next
    
    icoCX = GetSystemMetrics(SM_CXICON)
    icoCY = GetSystemMetrics(SM_CYICON)
    dTarget = Sqr(icoCX * icoCX + icoCY * icoCY)
    
    ReDim Index(1 To Count)
    ReDim dSizes(1 To Count)
    
    ' Priority #1 - best size
    For I = 1 To Count
        If icde(I).dwBytesInRes Then
            icoCX = icde(I).bWidth: If icoCX = 0& Then icoCX = 256&
            icoCY = icde(I).bHeight: If icoCY = 0& Then icoCY = 256&
            dSource = (dTarget - Sqr(icoCX * icoCX + icoCY * icoCY))
            For icoCX = 1 To I - 1
                If dSource < dSizes(icoCX) Then
                    CopyMemory dSizes(icoCX + 1), dSizes(icoCX), (I - icoCX) * 8&
                    CopyMemory Index(icoCX + 1), Index(icoCX), (I - icoCX) * 4&
                    Exit For
                End If
            Next
            dSizes(icoCX) = dSource
            Index(icoCX) = I
        End If
    Next
    
    ' sizes are in ascending order from largest to smallest in relation to system icon metrics
    ' larger icons will have a negative values while smaller will have postive values
    For I = 1 To UBound(icde)
        If dSizes(I) >= 0& Then Exit For
    Next
    
    icoCY = Count
    If I > icoCY Then                                           ' all icons are larger than system size
        For I = icoCY - 1 To 1 Step -1&                         ' use smallest size available
            If dSizes(I) < dSizes(icoCX) Then Exit For
        Next
        icoCX = I + 1
    Else                                                        ' icons match or exceed desired size; use smallest one
        icoCX = I                                               ' use smallest size available
        For I = icoCX + 1 To icoCY
            If dSizes(I) > dSizes(icoCX) Then Exit For
        Next
        icoCY = I - 1
    End If
    Erase dSizes()
    
    ' Priority #2 - highest color depth
    isCursor = (m_Header(3) And &H10)
    lBest = icoCX
    If (icde(Index(lBest)).wBitCount = 0 Or isCursor = True) Then icde(Index(lBest)).wBitCount = pvGetBitDepthFromBMPheader(iconBits(), icde(Index(lBest)).dwImageOffset + lHdrSize)
    For I = icoCX + 1 To icoCY
        If (icde(Index(I)).wBitCount = 0 Or isCursor = True) Then icde(Index(I)).wBitCount = pvGetBitDepthFromBMPheader(iconBits(), icde(Index(I)).dwImageOffset + lHdrSize)
        If icde(Index(I)).wBitCount > icde(Index(lBest)).wBitCount Then lBest = I
    Next
    pvGetBestSystemMatch = Index(lBest) - 1&

End Function

Private Function pvConvertAniCursor(theStream() As Byte, ICDirXRef() As Long, streamLength As Long, _
                                    sequencePtr As Long, durationPtr As Long, jiffies As Long) As Boolean

    ' function parses out the animated cursor

    Const aniACON As Long = 1313817409 ' flag indicating animated icon
    Const aniANIH As Long = 1751740001 ' ani cursor header data (36 bytes)
    Const aniICON As Long = 1852793705 ' individual icon entries
    Const aniFRAM As Long = 1835102822 ' flag indicating Frame, no additional data
    Const aniIART As Long = 1414676809 ' flag indicating Author/Artist, 4 byte length + Data
    Const aniINAM As Long = 1296125513 ' flag indicating Name/Description, 4 byte length + Data
    Const aniLIST As Long = 1414744396 ' flag indicating List data, 4 byte length + Data
    Const aniRATE As Long = 1702125938 ' flag indicating frame delay/rate, 4 bytes per frame
    Const aniSEQ As Long = 544302451   ' flag indicating frame sequence, 4 bytes per frame
    ' note regarding data length fields. The 4 byte length is not included in the data length

    Dim icPtr As Long, icValue As Long, icOffset As Long
    Dim icIndex As Long, nrIcons As Long, nrGroups As Long
    
    On Error GoTo ExitRoutine
    icPtr = 4&
    Do
        icPtr = icPtr + 4&
        If icPtr >= streamLength Then
            Exit Do
        End If
        CopyMemory icValue, theStream(icPtr), 4&
        Select Case icValue
        Case aniANIH            ' animated cursor header/UDT: 36 bytes unless modified in the future
            CopyMemory icValue, theStream(icPtr + 4&), 4&       ' data length; should be 36 for animated cursors
            If icPtr + icValue + 4& > streamLength Then Exit Function
            CopyMemory nrIcons, theStream(icPtr + 12&), 4&      ' nr icons per group of RIFF ICON blocks
            If nrIcons < 1& Or nrIcons > &HFFF& Then Exit Function
            CopyMemory jiffies, theStream(icPtr + 16&), 4&      ' default jiffies if no RATE block
            If jiffies < 0& Then
                jiffies = 0&
            Else
                jiffies = ((jiffies And &HFFFF&) * 1000&) / 60&
            End If
            icPtr = icPtr + icValue + 4&                        ' move pointer to next data field
        Case aniACON, aniFRAM                                   ' no data length field
            ' nothing to do, header only
        Case aniLIST                                            ' can occur more than once
            icPtr = icPtr + 4&
            If nrIcons = 0& Then
                CopyMemory icValue, theStream(icPtr), 4&        ' data length
                icPtr = icPtr + icValue
            End If
        Case aniICON                                            ' listing of icons
            If icIndex = nrIcons Then
                CopyMemory icValue, theStream(icPtr + 4&), 4&
                icPtr = icPtr + icValue + 4&
            Else
                CopyMemory icValue, theStream(icPtr + 4&), 4&   ' size of data block
                If icValue + 4& + icPtr > streamLength Then Exit Function ' corrupted
                
                CopyMemory nrGroups, theStream(icPtr + 12&), 2& ' nr of icons in this group
                If icIndex = 0& Then
                    If nrGroups < 1& Or nrGroups > &HFFF& Then Exit Function ' corrupted or not prepared to handle 4000+ cursors
                    ReDim ICDirXRef(0 To nrIcons * nrGroups)    ' size our header Xref table
                    ICDirXRef(0) = nrGroups Or nrIcons * &H10000 Or &H30000000 ' flag=cursor & RIFF
                    If nrGroups > 1& Then ICDirXRef(0) = ICDirXRef(0) Or &H40000000 ' flag=grouped RIFF
                End If                                          ' ensure nr icons in block <> bigger than other blocks
                If nrGroups > (ICDirXRef(0) And &HFFFF&) Then nrGroups = (ICDirXRef(0) And &HFFFF&)
                icPtr = icPtr + 8&                              ' set pointer to 1st icon directory in block
                For nrGroups = 0& To nrGroups - 1&              ' update XRef table with that pointer
                    icOffset = nrIcons * nrGroups + icIndex + 1&
                    ICDirXRef(icOffset) = icPtr + 6&
                Next
                icIndex = icIndex + 1&                          ' nr ICON blocks read
                icPtr = icPtr + icValue - 4&                    ' update pointer
            End If
        Case aniRATE
            CopyMemory icValue, theStream(icPtr + 4&), 4&       ' each has a data length field
            durationPtr = icPtr + 4&
            icPtr = durationPtr + icValue
            
        Case aniSEQ                                             ' valid entries not used herein
            CopyMemory icValue, theStream(icPtr + 4&), 4&       ' each has a data length field
            sequencePtr = icPtr + 4&
            icPtr = sequencePtr + icValue
        Case Else
            CopyMemory icValue, theStream(icPtr + 4&), 4&       ' each has a data length field
            icPtr = icPtr + icValue + 4&
        End Select
    Loop
    
    If icIndex Then pvConvertAniCursor = True
    
ExitRoutine:
    If Err Then Err.Clear
End Function

Public Function HICONtoArray(cImageData As cGDIpMultiImage, ByVal hIcon As Long, CacheData As Boolean) As Long

    ' Function takes an HICON handle and converts it to 1,4,8,24,32 bit icon file format
    ' If return value is False, outArray() contents are undefined
    ' Note: Bit reduction is in play. Example: If original source for HICON was 24 bit
    '   and it can be reduced/saved as 8 bit or lower without color loss, it will.
    ' Note: The end result's quality should be identical to HICON
    ' Note: Icon handles retrieved from stdPictures are cursors, not icons per GetIconInfo API
    ' Note: The array containing the icon/cursor will include the custom header (described at top of class)

    Dim tDC As Long, maskScan As Long, lResult As Long
    Dim ICI As ICONINFO, BHI As BITMAPINFO
    Dim icd As ICONDIR, icde As ICONDIRENTRY
    Dim cPal As ColorPalette, outArray() As Byte, lDur() As Long
    Dim tSource As Variant
    
    If hIcon = 0& Then Exit Function
    If GetIconInfo(hIcon, ICI) = 0& Then Exit Function
    
    ' A properly formatted icon file will contain this information:
    ' :: 6 byte ICONDIRECTORY structure
    ' :: 16 byte ICONDIRECTORYENTRY structure
    ' If stored in PNG format then
    '    :: The entire PNG
    ' Else
    '    :: 40 byte BITMAPINFOHEADER structure
    '    If paletted then: Palette entries, each in BGRA format
    '    :: Color data packed & word-aligned per Bitcount of 1,4,8,16,24,32 bits per pixel
    '    :: 1-bit word-aligned Mask data, even if mask not used (i.e., 32bpp)
    '    Size of any single icon's file can be calculated as:
    '    FileSize = 62 + NrPaletteEntries*4 + (modCommon.ByteAlignOnWord(BitCount,Width) + modCommon.ByteAlignOnWord(1,Width))*Height)
    ' Icon sizes are limited to maximum dimensions of 256x256
    
    On Error GoTo Catch_Exception
    tDC = GetDC(0&)
    BHI.bmiHeader.biSize = 40&
    If ICI.hbmColor = 0& Then                               ' black and white icon (rare, but so easy)
        If GetDIBits(tDC, ICI.hbmMask, 0, 0&, ByVal 0&, BHI, 0&) Then
            With BHI.bmiHeader
                .biClrUsed = 2&                             ' should be filled in, but ensure it is so
                .biClrImportant = .biClrUsed                ' should be filled in, but ensure it is so
                .biCompression = 0&
                .biSizeImage = modCommon.ByteAlignOnWord(1, .biWidth)
                BHI.bmiColors(2) = vbWhite                  ' set 2nd palette entry to white
                ' size array to the entire icon file format, includes Icon Directory structure, bitmap header, palette, & mask
                ReDim outArray(0 To .biSizeImage * .biHeight + 69&)
                ' 69 = (62 for headers + 8 color table) - 1
                ' this next call gets the entire icon data; just need to fill in the directory a bit further down this routine
                If GetDIBits(tDC, ICI.hbmMask, 0, .biHeight, outArray(70), BHI, 0&) = 0& Then
                    .biBitCount = 0
                Else
                    .biClrUsed = 2&                         ' fill in; last GetDIBits call erased it
                    .biClrImportant = .biClrUsed            ' fill in; last GetDIBits call erased it
                    .biHeight = .biHeight \ 2&              ' set to real height, not height*2 as is now
                End If
            End With
            CopyMemory outArray(62), BHI.bmiColors(1), 8&
        End If
        DeleteObject ICI.hbmMask: ICI.hbmMask = 0&          ' destroy; no longer needed
    
    Else                                                    ' color icon vs black & white
        If GetDIBits(tDC, ICI.hbmColor, 0, 0&, ByVal 0&, BHI, 0&) Then
            With BHI.bmiHeader
                .biBitCount = 32
                .biCompression = 0&
                .biClrImportant = 0&
                .biClrUsed = 0&
                .biSizeImage = 0&
                ReDim outArray(0 To .biWidth * .biHeight * 4& - 1&)
                If GetDIBits(tDC, ICI.hbmColor, 0, .biHeight, outArray(0), BHI, 0&) = 0 Then
                    .biBitCount = 0
                Else
                    maskScan = modCommon.ByteAlignOnWord(1, .biWidth) ' scan width for the mask portion of this icon
                    lResult = pvReduceBitmap32bpp(outArray(), .biWidth, .biHeight, , 22& + .biSize, maskScan * .biHeight, .biClrImportant)
                    Select Case lResult                     ' finish the bitmapinfo header
                        Case 0: .biBitCount = 0&
                        Case lvicColor32bppAlpha
                            ' nothing to do
                        Case lvicColor24bpp
                            .biBitCount = 24
                        Case Else
                            .biBitCount = ((lResult \ &H100&) And &HFF&)
                            .biClrUsed = 2 ^ .biBitCount
                    End Select
                    .biSizeImage = modCommon.ByteAlignOnWord(.biBitCount, .biWidth) * .biHeight
                    .biClrImportant = .biClrUsed
                End If
            End With
        End If
    End If
        
    If BHI.bmiHeader.biBitCount Then                                    ' tidy up
        ReDim m_Header(0 To 27)
        With BHI.bmiHeader
            icd.idCount = 1                                             ' fill in the directory
            If ICI.fIcon = 0 Then icd.idType = 2 Else icd.idType = 1
            If .biWidth < 256& Then icde.bWidth = .biWidth              ' fill in the directory entry
            If .biHeight < 256& Then icde.bHeight = .biHeight
            If .biClrUsed < 256& Then icde.bColorCount = .biClrUsed
            If icd.idType = 1 Then                                      ' icons
                icde.wBitCount = .biBitCount
                icde.wPlanes = (BHI.bmiHeader.biPlanes And &HFFFF&)
            Else                                                        ' cursors
                icde.wPlanes = (ICI.xHotspot And &HFFFF&)
                icde.wBitCount = (ICI.yHotspot And &HFFFF&)
                m_Header(3) = &H10                                      ' custom header flag indicating cursor
            End If
            icde.dwImageOffset = 22&
            icde.dwBytesInRes = UBound(outArray) - 21&
            
            .biHeight = .biHeight + .biHeight                           ' icon's store height*2 in bitmap header
        End With
        CopyMemory outArray(0), icd, 6&                                 ' copy the icon directory & bitmap header
        CopyMemory outArray(6), icde, 16&
        CopyMemory outArray(icde.dwImageOffset), BHI, BHI.bmiHeader.biSize
        
        ' done with the icon directory, now to the mask portion
        If ICI.hbmMask Then
            BHI.bmiColors(1) = vbBlack: BHI.bmiColors(2) = vbWhite      ' set up black/white palette
            With BHI.bmiHeader                                          ' set up bitmapinfo header
                .biBitCount = 1
                .biClrUsed = 2&
                .biClrImportant = .biClrUsed
                .biHeight = .biHeight \ 2&
                .biSizeImage = 0&
                icde.dwImageOffset = icde.dwBytesInRes - maskScan * .biHeight + icde.dwImageOffset  ' location where mask will be written
            End With
            GetDIBits tDC, ICI.hbmMask, 0&, BHI.bmiHeader.biHeight, outArray(icde.dwImageOffset), BHI, 0&  ' get the mask
        End If
        m_Header(0) = 1: m_Header(2) = 1
        m_Header(24) = 6: m_Header(8) = 255
        
        If ICI.fIcon = 1 Then lResult = lvicPicTypeIcon Else lResult = lvicPicTypeCursor
        cImageData.InitializeContent ObjPtr(Me), lResult, 1&
        ReDim lDur(0 To 1, 1 To 1)
        cImageData.SetSequenceDuration lDur()
        modCommon.MoveArrayToVariant tSource, outArray(), True
        cImageData.SetIndex tSource, 0&, &HF0000000
        modCommon.MoveArrayToVariant tSource, outArray(), False
        If cImageData.Handle Then
            cImageData.CacheSourceInfo VarPtrArray(outArray), 0&, lResult, CacheData, False
            HICONtoArray = lResult
        Else
            Set cImageData = New cGDIpMultiImage
        End If
    End If
    
Catch_Exception:                                                        ' clean up
    ReleaseDC 0&, tDC
    If ICI.hbmColor Then DeleteObject ICI.hbmColor
    If ICI.hbmMask Then DeleteObject ICI.hbmMask
    
End Function

Private Function pvGetBitDepthFromBMPheader(theStream() As Byte, Offset As Long) As Byte

    ' Bitmap's header structure will have bitcount on 14th byte from offset

    ' PNG's IHDR structure
    '    Width As Long              << cannot be negative
    '    Height As Long             << cannot be negative
    '    BitDepth As Byte           << must be 1,2,4,8,16
    '    ColorType As Byte          << must be 0,2,3,4,6
    '    Compression As Byte        << must be zero
    '    Filter As Byte             << must be zero
    '    Interlacing As Byte        << must be zero or one
    
    ' bit depth calculations
    '   Color    Allowed    Interpretation
    '   Type    Bit Depths
    '
    '   0       1,2,4,8,16  Each pixel is a grayscale sample.
    '
    '   2       8,16        Each pixel is an R,G,B triple.
    '
    '   3       1,2,4,8     Each pixel is a palette index; a PLTE chunk must appear.
    '
    '   4       8,16        Each pixel is a grayscale sample, followed by an alpha sample.
    '
    '   6       8,16        Each pixel is an R,G,B triple, followed by an alpha sample.
    
    Dim lValue As Long
    ' get bit depth from bitmap header unless png
    CopyMemory lValue, theStream(Offset), 4&
    If lValue <> png_Signature1 Then
        pvGetBitDepthFromBMPheader = theStream(Offset + 14&)
    Else
        ' Offset + 24& will be the PNG bit depth
        ' Offset + 25& will be the color type
        lValue = theStream(Offset + 24&)
        Select Case theStream(Offset + 25&)
            Case 2: If lValue = 8 Then lValue = 24
            Case 6: If lValue = 8 Then lValue = 32
            Case Else ' use reported depth
        End Select
        pvGetBitDepthFromBMPheader = lValue
    End If
End Function

Private Function pvLocateDirEntry(inArray() As Byte, IconIndex As Long, GroupIndex As Long, resourceOffset As Long) As Long

    ' helper function returns the address of the icon directory entry for passed IconIndex
    ' Of the 4 types of icons/cursors (see top comments regarding custom header)....
    ' 1) standard icon, the custom header contains the offset directly into the directory entry
    '       the dwImageOffset member is offset by 6 bytes (size of icon directory)
    ' 2) non-grouped animated cursors, the header contains offset into the directory entry
    '       the dwImageOffset member is offset by the location of the entry within the RIFF format
    ' 3) grouped animated cursors, the header contains offset of directory entry for each group
    '       The correct entry is then adjusted at 16 bytes * GroupIndex
    '       the dwImageOffset member is offset by the location of the entry within the RIFF format
    ' 4) binaries: function not used

    Dim nrGroups As Long, nrIcons As Long, lOffset As Long
    
    nrGroups = (m_Header(0) + m_Header(1) * &H100&)
    nrIcons = (m_Header(2) + (m_Header(3) And &HF) * &H100&)
    
    If IconIndex < 0& Or IconIndex >= nrIcons Then Exit Function
    If GroupIndex < 1& Then
        GroupIndex = 1&
    Else
        If GroupIndex > nrGroups Then Exit Function
    End If
        
    ' location for the 1st icondirectory entry is cached in inArray:
    resourceOffset = 0&
    CopyMemory lOffset, m_Header(((GroupIndex - 1&) * nrIcons + IconIndex) * 4& + 24&), 4&
    If lOffset Then
        '   (GroupIndex-1)*IconIndex+4 then offset 16 for the icon needed
        If (m_Header(3) And &H20) Then ' For RIFFs we use group index
            resourceOffset = lOffset - 6&
            If (m_Header(3) And &H40) Then lOffset = lOffset + (GroupIndex - 1&) * 16& ' grouped RIFF
        Else
            ' For standard icons, use icon directory index
            resourceOffset = 0&
        End If
    End If
    pvLocateDirEntry = lOffset
    
End Function

Private Function pvReduceBitmap32bpp(Bits() As Byte, Width As Long, Height As Long, _
                                    Optional ImageOffset As Long = 0&, _
                                    Optional HeaderSize As Long = 0&, _
                                    Optional FooterSize As Long = 0&, _
                                    Optional PaletteIndexesUsed As Long) As ImageColorFormatEnum

    ' purpose is to reduce 32 bpp images to a lower bit depth if possible, without color loss
    ' Bits() will always be resized on return
    ' This routine is used on source icon/cursor image bits. At this point in code, the image
    '   is not a GDI+ object. Therefore; can't use cColorReduction class
    
    Dim outArray() As Byte, lPixels() As Long
    Dim bNewColor As Boolean, Palette(1 To 256) As Long
    Dim lPrevPal As Long, X As Long, Y As Long
    Dim palIndex As Long, palPtr As Long
    Dim tSA As SafeArray, lColorType As Long
    Dim lScanWidth As Long, bitCount As Long
    Dim rShift As Byte, lShift As Byte, cOffset As Long
    
    With tSA
        .cbElements = 4
        .cDims = 1
        .pvData = VarPtr(Bits(ImageOffset))
        .rgSABound(0).cElements = Width * Height
    End With
    CopyMemory ByVal VarPtrArray(lPixels), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine
    
    PaletteIndexesUsed = 0&
    lPrevPal = lPixels(PaletteIndexesUsed) Xor 1&   ' forces mismatch in loop start
    lColorType = lvicColor32bppAlpha               ' default type
    For X = 0& To Width * Height - 1&               ' process each color
        If lPixels(X) <> lPrevPal Then
            If (lPixels(X) And &HFF000000) Then        ' uses alpha channel; 32bpp
                lColorType = lvicColor32bppAlpha
                PaletteIndexesUsed = 0&
                Exit For
                
            ElseIf lColorType = lvicColor32bppAlpha Then ' continue processing else identified as potential 24bpp icon
                palIndex = modCommon.FindColor(Palette(), lPixels(X), PaletteIndexesUsed, bNewColor) ' have we seen this color?
                If bNewColor Then                       ' if not, add to our palette
                    If PaletteIndexesUsed = 256& Then      ' max'd out on palette entries; treat as 24bpp
                        lColorType = lvicColor24bpp ' but don't exit loop cause we don't know now
                        PaletteIndexesUsed = 0&
                    Else                                ' prepare to add to our palette if new
                        PaletteIndexesUsed = PaletteIndexesUsed + 1&
                        If palIndex < PaletteIndexesUsed Then  ' keep our palette in ascending order for binary search
                            CopyMemory Palette(palIndex + 1&), Palette(palIndex), (PaletteIndexesUsed - palIndex) * 4&
                        End If
                        Palette(palIndex) = lPixels(X) ' add color now
                    End If
                End If
            End If
            lPrevPal = lPixels(X) ' track for faster looping
        End If
    Next
    If PaletteIndexesUsed Then
        Select Case PaletteIndexesUsed
            Case Is < 3&: lColorType = lvicColor1bpp: rShift = &H2
            Case Is < 17&: lColorType = lvicColor4bpp: rShift = &H10
            Case Else: lColorType = lvicColor8bpp: rShift = &H1
        End Select
        bitCount = ((lColorType \ &H100&) And &HFF&)
        lScanWidth = modCommon.ByteAlignOnWord(bitCount, Width) ' scan width of destination's color data
        ' size array to the entire icon file format, includes Icon Directory structure, bitmap header, palette, & mask
        If bitCount <= 8& Then cOffset = 2 ^ bitCount * 4&
        cOffset = cOffset + HeaderSize
        ReDim outArray(0 To lScanWidth * Height - 1& + cOffset + FooterSize)
        
        X = 0&
        lPrevPal = lPixels(X) Xor 1&                ' forces mismatch when loop starts
        For Y = X To Height - 1&                    ' start packing the palette indexes into bytes
            palPtr = Y * lScanWidth + cOffset       ' position where that byte will start for current row
            If rShift = &H2 Then lShift = &H80 Else lShift = rShift
            For X = X To X + Width - 1&   ' process each row of the source bitmap
                ' locate the color in our palette & subtract one (palette is 1-based, indexes are 0-based)
                If lPrevPal <> lPixels(X) Then
                    palIndex = modCommon.FindColor(Palette(), lPixels(X), PaletteIndexesUsed, bNewColor) - 1&
                    lPrevPal = lPixels(X) ' track for faster looping
                End If
                outArray(palPtr) = outArray(palPtr) Or palIndex * lShift ' pack the index
                If lShift = 1& Then               ' done with this byte
                    palPtr = palPtr + 1&            ' move destination to next byte
                    If rShift = &H2 Then lShift = &H80 Else lShift = rShift
                Else
                    lShift = lShift \ rShift ' adjust position where next index will be written
                End If
            Next
        Next
        CopyMemory outArray(HeaderSize), Palette(1), PaletteIndexesUsed * 4&
            
    ElseIf lColorType = lvicColor32bppAlpha Then
        ReDim outArray(0 To Width * Height * 4& + HeaderSize + FooterSize - 1&)
        CopyMemory outArray(HeaderSize), lPixels(0), Width * Height * 4&
        
    Else ' lvicColor24bpp
        lScanWidth = modCommon.ByteAlignOnWord(24&, Width) ' scan width of destination's color data
        ReDim outArray(0 To lScanWidth * Height - 1& + HeaderSize + FooterSize)
        X = 0&
        For Y = 0& To Height - 1&
            palPtr = Y * lScanWidth + HeaderSize
            For X = X To X + Width - 1&
                CopyMemory outArray(palPtr), lPixels(X), 3&
                palPtr = palPtr + 3&
            Next
        Next
    End If
    
ExitRoutine:
    CopyMemory ByVal VarPtrArray(lPixels), 0&, 4&
    If lColorType Then
        Erase Bits()
        CopyMemory ByVal VarPtrArray(Bits), ByVal VarPtrArray(outArray), 4&
        CopyMemory ByVal VarPtrArray(outArray), 0&, 4&
        pvReduceBitmap32bpp = lColorType
    End If
    
End Function

Private Function pvConvert24bppIcon(SourceHandle As Long, SS As SAVESTRUCT, tData() As Byte, as24bpp As Boolean) As Long

    ' VB stdPicture cannot support 32bpp icon/cursors.
    ' This function converts any bit depth to 24bpp

    Dim tBMP As BitmapData, tSA As SafeArray, srcData() As Byte
    Dim mShift As Byte, mPtr As Long, mOffset As Long
    Dim sizeI As RECTI, mScanWidth As Long, dScanWidth As Long
    Dim X As Long, Y As Long, dPtr As Long, bPP As Long
    Dim dOffset As Long, srcScanWidth As Long
    
    mScanWidth = modCommon.ByteAlignOnWord(1, SS.Width)
    If as24bpp Then bPP = lvicColor24bpp Else bPP = lvicColor32bpp
    dScanWidth = modCommon.ByteAlignOnWord((bPP And &HFF00&) \ &H100&, SS.Width)
    ReDim tData(0 To (mScanWidth + dScanWidth) * SS.Height + 61&)
    tBMP.Scan0Ptr = VarPtr(tData(62& + dScanWidth * SS.Height - dScanWidth))
    tBMP.stride = -dScanWidth
    sizeI.nHeight = SS.Height: sizeI.nWidth = SS.Width
    
    ' get the color bits first
    If GdipBitmapLockBits(SourceHandle, sizeI, ImageLockModeRead Or ImageLockModeUserInputBuf, bPP, tBMP) Then Exit Function
    GdipBitmapUnlockBits SourceHandle, tBMP: tBMP.Scan0Ptr = 0&
    If (SS.reserved1 And 1&) = 0& Then    ' source is not alpha; mask will be all black; done
        pvConvert24bppIcon = bPP
    Else
        If GdipBitmapLockBits(SourceHandle, sizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
        With tSA
            .cbElements = 1: .cDims = 1
            .pvData = tBMP.Scan0Ptr
            If tBMP.stride < 0& Then tSA.pvData = tSA.pvData + (tBMP.Height - 1&) * tBMP.stride
            srcScanWidth = Abs(tBMP.stride)
            .rgSABound(0).cElements = srcScanWidth * tBMP.Height
        End With
        CopyMemory ByVal VarPtrArray(srcData), VarPtr(tSA), 4&
        On Error GoTo ExitRoutine
        
        mOffset = UBound(tData) - mScanWidth + 1&
        dOffset = 62& + dScanWidth * SS.Height - dScanWidth
        With SS
            For Y = 0& To .Height - 1&
                mShift = &H80
                mPtr = mOffset - Y * mScanWidth
                dPtr = dOffset - Y * dScanWidth
                X = Y * srcScanWidth
                If bPP = lvicColor24bpp Then
                    For X = X + 3& To X + srcScanWidth - 1& Step 4&
                        If srcData(X) <= .AlphaTolerancePct Then
                            tData(mPtr) = tData(mPtr) Or mShift
                            tData(dPtr) = 0: tData(dPtr + 1) = 0: tData(dPtr + 2) = 0
                        End If
                        If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
                        dPtr = dPtr + 3&
                    Next
                Else
                    For X = X + 3& To X + srcScanWidth - 1& Step 4&
                        If srcData(X) <= .AlphaTolerancePct Then
                            tData(mPtr) = tData(mPtr) Or mShift
                            tData(dPtr) = 0: tData(dPtr + 1) = 0: tData(dPtr + 2) = 0: tData(dPtr + 3) = 0
                        Else
                            tData(dPtr + 3) = 255
                        End If
                        If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
                        dPtr = dPtr + 4&
                    Next
                End If
            Next
        End With
        pvConvert24bppIcon = lvicColor24bpp
    End If
    
ExitRoutine:
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
    If tBMP.Scan0Ptr Then GdipBitmapUnlockBits SourceHandle, tBMP
End Function

Public Function SaveAsICO(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                            SS As SAVESTRUCT, MIS As MULTIIMAGESAVESTRUCT, FrameNumber As Long) As Long
                            
    ' saves image as a ICO, CUR, HICON, HCURSOR
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    ' returnMedium.
    '   If saveTo_Array then returnObject is the 0-Bound array & function return value is size of array in bytes
    '   If saveTo_GDIhandle then returnObject is HBITMAP and return value is non-zero if successful
    '   If saveTo_File then returnObject is passed HFILE, return value is non-zero if successful. File is not closed
    '   If saveTo_stdPictureture then returnObject is passed stdPicture, return value is non-zero if successful
    '   If saveTo_Clipboard then return value is non-zero if successful
    '   If saveTo_DataObject then returnObject is passed DataObject, return value is non-zero if successful
    '   If saveTo_GDIpHandle then
    '       if creating own GDIpImage class then returnObject is the class & return value is saveTo_GDIpHandle
    '       else returnObject is handle's IStream source & function return value is the GDI+ handle
    
    If SourceHandle = 0& Then Exit Function
    
    Dim tData() As Byte, srcData() As Byte, tSA As SafeArray
    Dim tBMPi As BITMAPINFOHEADER, tBMP As BitmapData
    Dim tgtHandle As Long, lType As Long, lRtnMedium As Long
    Dim lResult As Long, srcDepth As Long, tSize As RECTI
    Dim IIStream As IUnknown, tmpPic As StdPicture, tmpDO As DataObject
    Dim cPal As cColorReduction, tGDIpImage As GDIpImage, tObject As Object
    Dim icoDIR As ICONDIRENTRY, lbOffset As Long
    
    Dim mShift As Byte, mOffset As Long, mPtr As Long, mScanWdith As Long
    Dim rShift As Byte, lShift As Byte, dPtr As Long, dScanWidth As Long
    Dim X As Long, Y As Long
    Const ICRESVER As Long = &H30000
    
    Select Case (SS.reserved1 And &HF0&) \ &H10                 ' icon or cursor?
        Case lvicPicTypeIcon, lvicSaveAs_HICON
            lType = lvicPicTypeIcon
        Case Else
            lType = lvicPicTypeCursor
    End Select
    
    If MIS.Images > 0& Then
        lRtnMedium = saveTo_Array
        lbOffset = LBound(MIS.Image)
        If MIS.Image(FrameNumber + lbOffset).IconEmbeddedAsPNG Then
            If modCommon.SaveAsPNG(tData(), SourceHandle, saveTo_Array, SS) = 0& Then Exit Function
            lResult = 1&
        End If
    Else
        lRtnMedium = returnMedium
    End If
    
    If lResult = 0& Then
        srcDepth = modCommon.ColorDepthToColorType(SS.ColorDepth, SourceHandle) ' what bit depth currently?
        Set cPal = New cColorReduction
        
        tgtHandle = SourceHandle
        ' VB stdPictures don't support 32bpp icons; handle them separately
        If lRtnMedium = saveTo_stdPicture Then
            If SS.ColorDepth <= lvicConvert_TrueColor24bpp Then        ' try colorless reduction
                SS.reserved1 = SS.reserved1 Or &H10000000
                lResult = cPal.PalettizeToArray(SourceHandle, tData(), orient_WantMask Or orient_BottomUp Or orient_WantPaletteInArray, 62&, 0&, alpha_Simple, SS)
                If lResult Then
                    srcDepth = lResult                              ' ok, either 1,4,8 bit icon now; else
                Else                                                ' couldn't reduce with default reduction
                    If (srcDepth And &HFF00&) \ &H100& > 24& Then   ' force to 24bpp if source is 32bpp
                        SS.ColorDepth = lvicConvert_TrueColor24bpp
                    ElseIf SS.ColorDepth = lvicDefaultReduction Then ' and set to no reduction if attempting default reduction
                        SS.ColorDepth = lvicNoColorReduction
                    End If
                End If
                SS.reserved1 = SS.reserved1 Xor &H10000000          ' remove flag
            ElseIf SS.ColorDepth > lvicConvert_TrueColor24bpp Then ' force to 24bpp
                SS.ColorDepth = lvicConvert_TrueColor24bpp
            End If
        ElseIf SS.ColorDepth = lvicConvert_TrueColor32bpp_RGB Then
            lResult = pvConvert24bppIcon(SourceHandle, SS, tData(), False)
            If lResult Then srcDepth = lvicColor32bpp Else Exit Function    ' abort if failure
        End If
        If lResult = 0& Then
            If SS.ColorDepth = lvicConvert_TrueColor24bpp Then
                ' convert to 24bpp; color reduction routines designed for palettes
                lResult = pvConvert24bppIcon(SourceHandle, SS, tData(), True)
                If lResult Then srcDepth = lResult Else Exit Function       ' abort if failure
            ElseIf SS.ColorDepth > lvicNoColorReduction Then
                If SS.ColorDepth = lvicDefaultReduction Then        ' colorless reudction?
                    SS.reserved1 = SS.reserved1 Or &H10000000: lResult = srcDepth
                ElseIf SS.ColorDepth < lvicConvert_TrueColor24bpp Then ' palettizing?
                    lResult = srcDepth
                End If                                                      ' else 32bpp
                If lResult Then
                    If lRtnMedium = saveTo_Clipboard Or lRtnMedium = saveTo_DataObject Then SS.reserved1 = SS.reserved1 Or &H20000000
                    lResult = cPal.PalettizeToArray(SourceHandle, tData(), orient_WantMask Or orient_BottomUp Or orient_WantPaletteInArray, 62&, 0&, alpha_Simple, SS)
                    If lResult Then
                        srcDepth = lResult
                    ElseIf SS.ColorDepth > lvicNoColorReduction Then
                        If lRtnMedium = saveTo_stdPicture Then Exit Function ' palettizing failed; abort
                    End If
                End If
            End If
        End If
        Set cPal = Nothing
        
        If lResult = 0& And SS.ExtractCurrentFrameOnly = False And SS.reserved2 <> 0& And _
            (SS.reserved1 And &HFF00&) \ &H100& = lType And MIS.Images = 0 And _
            lRtnMedium <> saveTo_stdPicture And lRtnMedium <> saveTo_GDIhandle Then
            CopyMemory tObject, SS.reserved2, 4&
            Set tGDIpImage = tObject
            CopyMemory tObject, 0&, 4&
            'If tGDIpImage.ImageCount < 2& Then tGDIpImage.ExtractImageData tData() Else Set tGDIpImage = Nothing
            tGDIpImage.ExtractImageData tData()
        End If
        If tGDIpImage Is Nothing Then
        
            With tBMPi                                                  ' fill in bitmap info
                .biBitCount = (srcDepth And &HFF00&) \ &H100&
                If .biBitCount <= 8 Then
                    .biClrImportant = 2 ^ .biBitCount                   ' palette size
                    .biClrUsed = .biClrImportant
                End If
                .biHeight = SS.Height * 2&                              ' always reported double actual size
                .biWidth = SS.Width
                .biPlanes = 1
                .biSize = 40
                dScanWidth = modCommon.ByteAlignOnWord(.biBitCount, .biWidth)     ' scanwidth
                .biSizeImage = dScanWidth * SS.Height
            End With
        
            If lResult = 0& Then                                        ' else tData() already contains color & mask bits
                mScanWdith = modCommon.ByteAlignOnWord(1, SS.Width)               ' we are using source image; gather data
                X = 62& + tBMPi.biClrUsed * 4&                          ' offset to where color data starts
                mOffset = dScanWidth * SS.Height + X                    ' offset to where mask data starts
                ReDim tData(0 To mOffset + mScanWdith * SS.Height - 1&) ' size target array
                
                tBMP.Scan0Ptr = VarPtr(tData(mOffset - dScanWidth))     ' where to copy color data to
                tBMP.stride = -dScanWidth                               ' want data copied bottom up from array
                tSize.nHeight = SS.Height: tSize.nWidth = SS.Width      ' copy all data
                If GdipBitmapLockBits(SourceHandle, tSize, ImageLockModeRead Or ImageLockModeUserInputBuf, srcDepth, tBMP) Then Exit Function
                GdipBitmapUnlockBits SourceHandle, tBMP                 ' done; now open it for reading
                If GdipBitmapLockBits(SourceHandle, tSize, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
                
                With tSA
                    .cbElements = 1: .cDims = 1
                    .pvData = tBMP.Scan0Ptr
                    If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
                    dScanWidth = Abs(tBMP.stride)
                    .rgSABound(0).cElements = dScanWidth * SS.Height
                End With
                CopyMemory ByVal VarPtrArray(srcData), VarPtr(tSA), 4&  ' overlay source
                
                mOffset = UBound(tData) - mScanWdith + 1&
                With SS
                    For Y = 0& To .Height - 1&                            ' loop thru, creating mask
                        mPtr = mOffset - Y * mScanWdith
                        X = Y * dScanWidth: mShift = &H80
                        For X = X + 3& To X + dScanWidth - 1& Step 4&       ' create the mask
                            If srcData(X) <= .AlphaTolerancePct Then tData(mPtr) = tData(mPtr) Or mShift
                            If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
                        Next
                    Next
                End With
                GdipBitmapUnlockBits SourceHandle, tBMP                 ' release source & overlay
                CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
                If tBMPi.biBitCount <= 8& Then                          ' get palette if source is 8bpp or less
                    Set cPal = New cColorReduction
                    cPal.CopyRemotePalette VarPtr(tData(62)), SourceHandle, False, 4&
                    Set cPal = Nothing
                    For mPtr = 65& To tBMPi.biClrImportant * 4& - 1& Step 4&
                        If tData(mPtr) <= SS.AlphaTolerancePct Then
                            tData(mPtr - 1&) = 0: tData(mPtr - 2&) = 0: tData(mPtr - 3&) = 0
                        End If
                        tData(mPtr) = 0                                 ' remove alpha from palette
                    Next
                    tBMPi.biSizeImage = tBMPi.biSizeImage + tBMPi.biClrUsed * 4&
                End If
            End If
                
            CopyMemory tData(22), tBMPi, tBMPi.biSize       ' fill in the icon/cursor directory & diretory entry
            tData(4) = 1                                    ' nr icons/cursors
            If SS.Width < 256& Then tData(6) = SS.Width     ' width
            If SS.Height < 256& Then tData(7) = SS.Height   ' height
            If tBMPi.biBitCount = 4& Then                   ' color count
                tData(8) = 16
            ElseIf tBMPi.biBitCount = 1& Then
                tData(8) = 2
            End If                                          ' tData(9) reserved
            CopyMemory tData(14), CLng(UBound(tData) - 21&), 4& ' bytes in resource
            tData(18) = 22                                  ' offset to image
            If lType = lvicPicTypeIcon Then
                tData(2) = 1                                ' icon
                tData(10) = 1                               ' planes
                tData(12) = tBMPi.biBitCount                ' bit count
            Else
                tData(2) = 2                                ' cursor
                With SS                                     ' validate hotspot
                    If .CursorHotSpotX < 0 Then
                        .CursorHotSpotX = 0
                    ElseIf .CursorHotSpotX > .Width - 1 Then
                        .CursorHotSpotX = .Width - 1
                    End If
                    If .CursorHotSpotY < 0 Then
                        .CursorHotSpotY = 0
                    ElseIf .CursorHotSpotY > .Height - 1 Then
                        .CursorHotSpotY = .Height - 1
                    End If
                    CopyMemory tData(10), .CursorHotSpotX, 2&   ' populate hotspot
                    CopyMemory tData(12), .CursorHotSpotY, 2&
                End With
            End If
        End If
        
        Select Case lRtnMedium
        Case saveTo_GDIhandle
            If lType = lvicPicTypeIcon Then ' HICON
                lResult = CreateIconFromResourceEx(tData(22), UBound(tData) - 21&, True, ICRESVER, SS.Width, SS.Height, 0&)
            Else            ' HCURSOR
                CopyMemory tData(18), tData(10), 4& ' place hotspot in front of data, increase resource size
                lResult = CreateIconFromResourceEx(tData(18), UBound(tData) - 17&, False, ICRESVER, SS.Width, SS.Height, 0&)
            End If
            SaveAsICO = lResult: returnObject = lResult
        Case saveTo_Array
            SaveAsICO = UBound(tData) + 1
            If MIS.Images = 0& Then modCommon.MoveArrayToVariant returnObject, tData(), True
        Case saveTo_File
            WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
            SaveAsICO = (lResult > UBound(tData))
        Case saveTo_stdPicture
            If lType = lvicPicTypeCursor Then tData(2) = 1 ' say this is icon vs. cursor
            ' this patch appears to work when cursors are loaded into stdPicture objects; else cursor comes out B&W
            Set tmpPic = modCommon.ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
            If Not tmpPic Is Nothing Then
                Set returnObject = tmpPic
                SaveAsICO = (tmpPic.Handle <> 0&)
            End If
        Case Else
            If lRtnMedium = saveTo_GDIplus Then
                Set tGDIpImage = modCommon.LoadImage(tData(), True, , True)
                If tGDIpImage.Handle Then
                    SaveAsICO = saveTo_GDIplus
                    Set returnObject = tGDIpImage
                End If
            Else
                If tGDIpImage Is Nothing Then Set tGDIpImage = modCommon.LoadImage(tData(), True, , True)
                If tGDIpImage.Handle Then
                    If GdipCreateHBITMAPFromBitmap(tGDIpImage.Handle, lResult, SS.RSS.FillColorARGB Or &HFF000000) = 0& Then
                        Set tGDIpImage = Nothing
                        SaveAsICO = True
                        Set tmpPic = modCommon.HandleToStdPicture(lResult, vbPicTypeBitmap)
                        If tmpPic Is Nothing Then
                            DeleteObject lResult: SaveAsICO = False
                        ElseIf lRtnMedium = saveTo_Clipboard Then
                            Clipboard.SetData tmpPic
                            If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat tData(), g_ClipboardFormat
                        ElseIf lRtnMedium = saveTo_DataObject Then
                            Set tmpDO = returnObject
                            tmpDO.SetData tmpPic, vbCFBitmap
                            If g_ClipboardFormat Then tmpDO.SetData tData(), g_ClipboardFormat
                            Set tmpDO = Nothing
                        End If
                        Set tmpPic = Nothing
                    End If
                End If
            End If
        End Select
    End If
    
    If MIS.Images > 0& Then
        If MIS.Image(FrameNumber + lbOffset).IconEmbeddedAsPNG Then
            X = UBound(tData) + 1&
        Else
            X = UBound(tData) - 21&
        End If
        If FrameNumber = 0& Then
            ReDim m_MultiDestData(0 To MIS.Images * 16& + 5& + X)
            CopyMemory m_MultiDestData(0), CLng(MIS.Images * 16& + 6&), 4&
            CopyMemory m_MultiDestData(4), MIS.Images, 2&
        Else
            ReDim Preserve m_MultiDestData(0 To UBound(m_MultiDestData) + X)
        End If
        CopyMemory icoDIR.dwImageOffset, m_MultiDestData(0), 4&
        If MIS.Image(FrameNumber + lbOffset).IconEmbeddedAsPNG Then
            If SS.Height < 256 Then icoDIR.bHeight = SS.Height
            If SS.Width < 256 Then icoDIR.bWidth = SS.Width
            icoDIR.dwBytesInRes = X
            If lType = lvicPicTypeIcon Then
                icoDIR.wBitCount = pvGetBitDepthFromBMPheader(tData(), 0&)
                Select Case icoDIR.wBitCount
                    Case 1: icoDIR.bColorCount = 2
                    Case 4: icoDIR.bColorCount = 16
                End Select
            Else
                icoDIR.wPlanes = SS.CursorHotSpotX
                icoDIR.wBitCount = SS.CursorHotSpotY
            End If
            CopyMemory m_MultiDestData(FrameNumber * 16& + 6&), icoDIR, 16&
            CopyMemory m_MultiDestData(icoDIR.dwImageOffset), tData(0), X
        Else
            CopyMemory m_MultiDestData(icoDIR.dwImageOffset), tData(22), X
            CopyMemory m_MultiDestData(FrameNumber * 16& + 6&), tData(6), 16&
            CopyMemory m_MultiDestData(FrameNumber * 16& + 18&), icoDIR.dwImageOffset, 4&
        End If
        
        If FrameNumber < MIS.Images - 1& Then
            CopyMemory m_MultiDestData(0), CLng(icoDIR.dwImageOffset + X), 4&
            SaveAsICO = 1&
        Else
            ' last frame processed
            m_MultiDestData(0) = 0: m_MultiDestData(1) = 0: m_MultiDestData(3) = 0
            If lType = lvicPicTypeIcon Then m_MultiDestData(2) = 1 Else m_MultiDestData(2) = 2
            CopyMemory m_MultiDestData(4), MIS.Images, 2&
            Select Case returnMedium
                Case saveTo_GDIplus
                    Set tGDIpImage = modCommon.LoadImage(m_MultiDestData(), True, , True)
                    If tGDIpImage.Handle <> 0& Then
                        SaveAsICO = returnMedium
                        Set returnObject = tGDIpImage
                    End If
                Case saveTo_Array
                    SaveAsICO = UBound(m_MultiDestData) + 1&
                    modCommon.MoveArrayToVariant returnObject, m_MultiDestData(), True
                Case saveTo_File
                    WriteFile CLng(returnObject), m_MultiDestData(0), UBound(m_MultiDestData) + 1&, lResult, ByVal 0&
                    SaveAsICO = (lResult > UBound(m_MultiDestData))
                Case saveTo_stdPicture, saveTo_Clipboard, saveTo_DataObject
                    Set tmpPic = modCommon.ArrayToPicture(VarPtr(m_MultiDestData(0)), UBound(m_MultiDestData) + 1&)
                    If Not tmpPic Is Nothing Then
                        If returnMedium = saveTo_Clipboard Then
                            Clipboard.SetData tmpPic
                            If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat m_MultiDestData(), g_ClipboardFormat
                        ElseIf returnMedium = saveTo_DataObject Then
                            Set tmpDO = returnObject
                            tmpDO.SetData tmpPic, vbCFBitmap
                            If g_ClipboardFormat Then tmpDO.SetData m_MultiDestData(), g_ClipboardFormat
                            Set tmpDO = Nothing
                        Else
                            Set returnObject = tmpPic
                        End If
                        SaveAsICO = tmpPic.Handle
                        Set tmpPic = Nothing
                    End If
            End Select
            Erase m_MultiDestData()
        End If
    End If
    Erase tData()
End Function
