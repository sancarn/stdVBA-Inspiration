VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  -----======== PURPOSE: Read/Write PCX image format & Convert to Bitmap ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' PCX Format specifications retrieved from following
' http://netghost.narod.ru/gff/vendspec/pcx/pcx.txt

' Class currently supports the following formats
'   Reading/Writing: 1, 16 & 256 color paletted images, 24 & 32 bpp true color images

Option Explicit

Public Function IsPCXResource(inArray() As Byte, ByVal FileHandle As Long) As Long

    ' Function attempts to rule out whether stream contains a PCX file
    ' Return value = 1 if PCX, -1 if corrupted PCX, 0 if not PCX
    
    Dim Cx As Integer, Cy As Integer, scanWidth As Integer
    Dim tData(0 To 68) As Byte, lSize As Long
    
    If FileHandle Then
        lSize = GetFileSize(FileHandle, 0&)
        If lSize < 129& Then Exit Function
        SetFilePointer FileHandle, 0&, 0&, 0&
        ReadFile FileHandle, tData(0), 69&, lSize, ByVal 0&
    Else
        If UBound(inArray) < 129& Then Exit Function
        CopyMemory tData(0), inArray(0), 69&
    End If
    If tData(0) = 10 And tData(2) < 2 Then                          ' magic number & encoding flag
        Select Case tData(1)                                        ' version info
            Case 0, 2, 3, 4, 5
                CopyMemory scanWidth, tData(4), 2&
                CopyMemory Cx, tData(8), 2&                         ' validate size & scanwidth
                Cx = Cx - scanWidth + 1
                CopyMemory scanWidth, tData(6), 2&
                CopyMemory Cy, tData(10), 2&
                Cy = Cy - scanWidth + 1
                CopyMemory scanWidth, tData(66), 2&
                ' for supported planes & color depths, ensure the bytes per line provided
                ' by the pcx file meets/exceeds minimal bytes required for bitcount
                
                ' To support some poorly formatted PCX files, the following logic will be used
                ' 1. If 4 planes: If 8 or 32 bits per pixel then 32bpp else always 4 bpp
                ' 2. If 3 planes: If 8 or 24 bits per pixel then 24bpp
                ' 3. If 1 plane: 1,4,8 bits per pixel then 1,4,8 bpp
                
                If (Cx < 1 Or Cy < 1 Or scanWidth < 1) = False Then
                    Select Case tData(65)                           ' color planes
                    Case 1
                        If tData(3) = 8 Then                        ' 256 color
                            If UBound(inArray) > 848 Then IsPCXResource = Abs(scanWidth >= Cx)
                            ' 848 = 128 + 720 (128=pcx header); (720=768 palette bytes less 16 bytes)
                        ElseIf tData(3) = 1 Then                    ' B&W
                            IsPCXResource = Abs(scanWidth >= Cx \ 8&)
                        ElseIf tData(3) = 4 Then                    ' 16 color (non-standard pcx)
                            IsPCXResource = Abs(scanWidth >= Cx \ 2&)
                        End If
                    Case 3  ' 24 bit
                        If tData(3) = 8 Or tData(3) = 24 Then       ' true color (no alpha)
                            IsPCXResource = Abs(scanWidth >= Cx)
                        End If
                    Case 4
                        If tData(3) = 8 Or tData(3) = 32 Then       ' true color with alpha (non-standard)
                            IsPCXResource = Abs(scanWidth >= Cx)
                        Else                                        ' treat as 4bpp
                            IsPCXResource = Abs(scanWidth * 4& >= Cx \ 2&)
                        End If
                    Case Else
                    End Select
                End If
            Case Else
        End Select
    End If

End Function

Public Function LoadPCXResource(inStream() As Byte) As Long

    ' NOTE: Must call IsPCXResource first & function must return 1

    ' Function processes an array of bytes suspected of containing a PCX file.
    ' Return value is a sourceless GDI+ handle
    
    Dim m_Palette As ColorPalette
    Dim m_BMP As BitmapData
    Dim Offset As Long
    
    If UBound(inStream) > 128 Then
        Select Case inStream(1)
        Case 0, 2, 3, 4, 5                                      ' pcx versions supported
            Select Case inStream(65)                            ' number of color planes
            Case 1
                If inStream(3) = 8 Then
                    If pvCreatePalette(inStream(), m_Palette, 8&, 16&) Then m_BMP.PixelFormat = lvicColor8bpp
                ElseIf inStream(3) = 1 Then                     ' black & white
                    If pvCreatePalette(inStream(), m_Palette, 1&, 16&) Then m_BMP.PixelFormat = lvicColor1bpp
                ElseIf inStream(3) = 4 Then                     ' 4 bpp 16 colors
                    If pvCreatePalette(inStream(), m_Palette, 4&, 16&) Then m_BMP.PixelFormat = lvicColor4bpp
                End If
            Case 3                                              ' (3 planes @ 8 bits) = 24 bpp
                If inStream(3) = 8 Or inStream(3) = 24 Then m_BMP.PixelFormat = lvicColor24bpp
            Case 4
                If inStream(3) = 8 Or inStream(3) = 32 Then ' (4 planes @ 8 bits) = 32 bpp
                    m_BMP.PixelFormat = lvicColor32bppAlpha
                Else                                        ' (4 planes @ 1 bits) = 4 bpp
                    If pvCreatePalette(inStream(), m_Palette, 4&, 16&) Then m_BMP.PixelFormat = lvicColor4bpp
                End If
            Case Else
            End Select
        Case Else
        End Select
    End If
        
    CopyMemory Offset, inStream(4), 2&
    CopyMemory m_BMP.Width, inStream(8), 2&                     ' get width/height from stream
    m_BMP.Width = m_BMP.Width - Offset + 1&
    CopyMemory Offset, inStream(6), 2&
    CopyMemory m_BMP.Height, inStream(10), 2&
    m_BMP.Height = m_BMP.Height - Offset + 1&
    
    Select Case m_BMP.PixelFormat
        Case lvicColor1bpp, lvicColor8bpp
            LoadPCXResource = pvProcessColorMap(inStream(), (inStream(66) Or inStream(67) * &H100&), m_BMP, m_Palette, (inStream(2) = 1))
        Case lvicColor4bpp
            If inStream(65) = 1 Then
                LoadPCXResource = pvProcessColorMap(inStream(), (inStream(66) Or inStream(67) * &H100&), m_BMP, m_Palette, (inStream(2) = 1))
            Else
                LoadPCXResource = pvProcess4bpp_4planes(inStream(), (inStream(66) Or inStream(67) * &H100&), m_BMP, m_Palette, (inStream(2) = 1))
            End If
        Case lvicColor24bpp, lvicColor32bppAlpha
            LoadPCXResource = pvProcessTrueColor(inStream(), (inStream(66) Or inStream(67) * &H100&), m_BMP, (inStream(2) = 1))
    End Select
    
End Function

Public Function SaveAsPCX(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                           SaveOptions As SAVESTRUCT) As Long

    ' saves image as a PCX
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    ' PCX will save any 1 bit image as black and white; even if the image is not black & white
    ' There are 2 workarounds
    '   1) Save PCX at 24 or 32 bpp to force 1bpp non-b&w image with color
    '   2) Rewrite the 1bpp b&w image as 4 bpp.
    ' I use option #2. 4bpp image is far smaller than 24bpp image
    
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
    
    
    Dim bProcessed As Boolean, lResult As Long
    Dim srcDepth As ColorDepthEnum, cPal As ColorPalette
    Dim palData() As Byte, tData() As Byte
    Dim srcData() As Byte, srcSA As SafeArray
    Dim tmpPic As StdPicture, tmpDO As DataObject
    Dim cColorReduction As cColorReduction, tSizeF As RECTF
    Dim tSizeI As RECTI, tBMP As BitmapData, Compression As Boolean
    Dim tGDIpImage As GDIpImage, tObject As Object
    
    If SourceHandle = 0& Then Exit Function
    If IsMissing(SaveOptions) Then Exit Function ' this function is never called by other than modCommon.SaveImage
    
    If SaveOptions.CompressionJPGQuality <> lvicFormatUncompressed Then Compression = True
    srcDepth = modCommon.ColorDepthToColorType(lvicDefaultReduction, SourceHandle)
    Set cColorReduction = New cColorReduction
    
    If (SaveOptions.reserved1 And 1&) = 0& Then      ' 1bpp workaround; if b&w will remain b&w
        If modCommon.ColorDepthToColorType(lvicNoColorReduction, SourceHandle) = lvicColor1bpp Then
            GdipGetImagePaletteSize SourceHandle, lResult
            If lResult = 16& Then         ' 2 colors + 2 flags
                GdipGetImagePalette SourceHandle, cPal, lResult
                If Not (cPal.Entries(1) = &HFF000000 And cPal.Entries(2) = &HFFFFFFFF) Then
                    SaveOptions.ColorDepth = lvicConvert_16Colors    ' going to force 4bpp
                    SaveOptions.PaletteType = lvicPaletteAdaptive
                Else
                    SaveOptions.ColorDepth = lvicNoColorReduction    ' no point; already 1bpp
                End If
            End If
            lResult = 0&
        End If
    End If
    If SaveOptions.ColorDepth > lvicNoColorReduction Then
        If returnMedium = saveTo_Clipboard Or returnMedium = saveTo_DataObject Then SaveOptions.reserved1 = SaveOptions.reserved1 Or &H20000000
        If SaveOptions.ColorDepth = lvicDefaultReduction Then
            SaveOptions.reserved1 = SaveOptions.reserved1 Or &H10000000
            lResult = cColorReduction.PalettizeToArray(SourceHandle, palData(), orient_TopDown Or orient_4bppIndexesMin, 0&, 0&, alpha_None, SaveOptions)
        ElseIf SaveOptions.ColorDepth < lvicConvert_TrueColor24bpp Then
            lResult = cColorReduction.PalettizeToArray(SourceHandle, palData(), orient_TopDown Or orient_4bppIndexesMin, 0&, 0&, alpha_None, SaveOptions)
            If lResult = 0& Then Exit Function          ' palettizing failed
        End If
        If lResult Then srcDepth = lResult: bProcessed = True
    End If
    
    If (bProcessed = False And (SaveOptions.reserved1 And &HFF00&) \ &H100& = lvicPicTypePCX And SaveOptions.reserved2 <> 0&) Then
        ' if we get here, then image has not be bit reduced above, is already in PCX format & a handle
        ' to the original data is available. No need to convert PCX to PCX; just extract the data
        CopyMemory tObject, SaveOptions.reserved2, 4&
        Set tGDIpImage = tObject
        CopyMemory tObject, 0&, 4&
        If tGDIpImage.ExtractImageData(tData()) = False Then Exit Function
        
    Else    ' convert to PCX
    
        tSizeI.nHeight = SaveOptions.Height: tSizeI.nWidth = SaveOptions.Width
        If (srcDepth And &HFF00&) \ &H100& < 24& Then
            If bProcessed Then
                srcSA.pvData = VarPtr(palData(0))
            ElseIf GdipBitmapLockBits(SourceHandle, tSizeI, ImageLockModeRead, srcDepth, tBMP) Then
                Exit Function
            Else
                srcSA.pvData = tBMP.Scan0Ptr
                If tBMP.stride < 0& Then srcSA.pvData = srcSA.pvData + (tBMP.Height - 1&) * tBMP.stride
            End If
            With srcSA
                .cbElements = 1
                .cDims = 2
                .rgSABound(0).cElements = tSizeI.nHeight
                .rgSABound(1).cElements = modCommon.ByteAlignOnWord((srcDepth And &HFF00&) \ &H100&, tSizeI.nWidth)
            End With
            CopyMemory ByVal VarPtrArray(srcData), VarPtr(srcSA), 4&
            lResult = pvCompressColorMap(tData(), srcData(), Compression, tSizeI.nWidth, tSizeI.nHeight, (srcDepth And &HFF00&) \ &H100&, (tBMP.stride < 0&))
            CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
            If tBMP.Scan0Ptr Then GdipBitmapUnlockBits SourceHandle, tBMP
            If lResult Then
                If bProcessed Then
                    cColorReduction.CopyLocalPalette VarPtr(tData(lResult)), True, 3&
                Else
                    cColorReduction.CopyRemotePalette VarPtr(tData(lResult)), SourceHandle, True, 3&
                End If
            End If
        Else
            If bProcessed Then ' palData is 24/32 bpp format
                srcSA.pvData = VarPtr(palData(0))
                lResult = 24&
            ElseIf GdipBitmapLockBits(SourceHandle, tSizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then
                Exit Function
            Else
                lResult = 32&
                srcSA.pvData = tBMP.Scan0Ptr
                If tBMP.stride < 0& Then srcSA.pvData = srcSA.pvData + (tBMP.Height - 1&) * tBMP.stride
            End If
            With srcSA
                .cbElements = 1
                .cDims = 2
                .rgSABound(0).cElements = tSizeI.nHeight
                .rgSABound(1).cElements = modCommon.ByteAlignOnWord(lResult, tSizeI.nWidth)
            End With
            CopyMemory ByVal VarPtrArray(srcData), VarPtr(srcSA), 4&
            pvCompressTrueColor tData(), srcData(), Compression, tSizeI.nWidth, tSizeI.nHeight, lResult, (srcDepth And &HFF00&) \ &H100, (tBMP.stride < 0&)
            CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
            If tBMP.Scan0Ptr Then GdipBitmapUnlockBits SourceHandle, tBMP
        End If
        Erase palData()
    
        ' fill in pcx header
        If Compression Then tData(2) = 1                        ' compression flag
        tData(0) = 10: tData(1) = 5                             ' magic number, version
        tData(4) = 1: tData(6) = 1                              ' left/top offset
        CopyMemory tData(8), SaveOptions.Width, 2&                       ' width
        CopyMemory tData(10), SaveOptions.Height, 2&                     ' height
        tData(12) = 96: tData(14) = 96                          ' DPI resolution (horizontal & vertical)
    End If
    Set cColorReduction = Nothing
    
    If returnMedium = saveTo_Array Then
        SaveAsPCX = UBound(tData) + 1&
        modCommon.MoveArrayToVariant returnObject, tData(), True
    ElseIf returnMedium = saveTo_File Then
        WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
        SaveAsPCX = (lResult > UBound(tData))
    ElseIf returnMedium = saveTo_GDIplus Then
        Set returnObject = modCommon.LoadImage(tData(), True, , True)
        If returnObject.Handle Then SaveAsPCX = saveTo_GDIplus
    Else
        lResult = 0&
        If tGDIpImage Is Nothing Then Set tGDIpImage = modCommon.LoadImage(tData(), , , True)
        If tGDIpImage Then GdipCreateHBITMAPFromBitmap tGDIpImage.Handle, lResult, SaveOptions.RSS.FillColorARGB Or &HFF000000
        Set tGDIpImage = Nothing
        If lResult Then
            SaveAsPCX = True
            Set tmpPic = modCommon.HandleToStdPicture(lResult, vbPicTypeBitmap)
            If tmpPic Is Nothing Then
                DeleteObject lResult: SaveAsPCX = False
            ElseIf returnMedium = saveTo_stdPicture Then
                Set returnObject = tmpPic
                SaveAsPCX = (tmpPic.Handle <> 0&)
            ElseIf returnMedium = saveTo_Clipboard Then
                Clipboard.SetData tmpPic
                If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat tData(), g_ClipboardFormat
            ElseIf returnMedium = saveTo_DataObject Then
                Set tmpDO = returnObject
                tmpDO.SetData tmpPic, vbCFBitmap
                If g_ClipboardFormat Then tmpDO.SetData tData(), g_ClipboardFormat
                Set tmpDO = Nothing
            End If
            Set tmpPic = Nothing
        End If
    End If
    Erase tData()

End Function

Private Function pvCreatePalette(inStream() As Byte, inPal As ColorPalette, BitDepth As Long, cMapLoc As Long) As Boolean

    ' helper function locates image palette and converts it to ARGB format

    Dim Index As Long, nrBytes As Long, xStop As Long
    Dim lPtrEOF As Long, lptrHDR As Long
    
    Select Case BitDepth                                ' when not found, try these ideas
    Case 8&
        If inStream(UBound(inStream) - 768&) = 12& Then     ' per specs, palette should be at very end of file
            lPtrEOF = UBound(inStream) - 767&               ' found, use it
        ElseIf inStream(UBound(inStream) - 720&) = 12& Then ' see if all but 16 colors are at end of file
            lptrHDR = cMapLoc                               ' yep; assume split between header & end of file
            lPtrEOF = UBound(inStream) - 719&
        End If
        inPal.Count = 256&
    Case 4&                                                 ' must be in the header
        lptrHDR = cMapLoc
        inPal.Count = 16&
    Case 1&
        inPal.Entries(1) = &HFF000000: inPal.Entries(2) = &HFFFFFFFF
        inPal.Count = 2&
        pvCreatePalette = True
    End Select
    
    If (lptrHDR Or lPtrEOF) Then                            ' have a pointer or two?
        nrBytes = inPal.Count * 3&
        If lptrHDR Then                                     ' start with the header palette
            If nrBytes > 48& Then xStop = 48& Else xStop = nrBytes  ' calc nr of bytes to extract
            For lptrHDR = lptrHDR To lptrHDR + xStop - 1 Step 3&
                Index = Index + 1&
                inPal.Entries(Index) = inStream(lptrHDR) * &H10000 Or inStream(lptrHDR + 1&) * &H100& Or inStream(lptrHDR + 2&) Or &HFF000000
            Next
            nrBytes = nrBytes - xStop                       ' adjust nr bytes needed
        End If
        
        If lPtrEOF Then                                     ' get remaining palette entries from end of file
            For lPtrEOF = lPtrEOF To lPtrEOF + nrBytes - 1& Step 3&
                Index = Index + 1&
                inPal.Entries(Index) = inStream(lPtrEOF) * &H10000 Or inStream(lPtrEOF + 1&) * &H100& Or inStream(lPtrEOF + 2&) Or &HFF000000
            Next
        End If
        
        pvCreatePalette = True
        
    End If
    

End Function

Private Function pvCompressColorMap(arrRLE() As Byte, rawStream() As Byte, Compressed As Boolean, _
                                    ByVal imgWidth As Long, ByVal imgHeight As Long, bitCount As Long, bInvertedSrc As Boolean) As Long

    ' Function compresses an 1 or 8 bpp paletted image to pcx standards
    ' rawStream passed as 2D palette indexes relative to passed Palette()
    ' Note: bits extracted from GDI+ are top down, not bottom up
    ' Note: Palette() is RGB format @ 3 bytes per color

    If bitCount = 4& Then                                    ' handled separately as 1 bpp @ 4 planes
        pvCompressColorMap = pvCompressColorMap_4Planes(arrRLE(), rawStream(), Compressed, imgWidth, imgHeight, bitCount, bInvertedSrc)
        Exit Function
    End If

    Dim X As Long, Y As Long, lStep As Long, lPrev As Byte
    Dim ePtr As Long, bEncoded() As Byte, BytesPerLine As Long
    Dim palSize As Long, lCount As Long
    
    If bitCount = 8& Then
        BytesPerLine = imgWidth + (imgWidth And 1&)
        palSize = 769& ' < 768 for palette & 1 for palette intro
    Else
        BytesPerLine = imgWidth \ 8 - ((imgWidth Mod 8) > 0)
        BytesPerLine = BytesPerLine + (BytesPerLine And 1&)
    End If
    
    If bInvertedSrc Then
        X = imgHeight - 1&: Y = 0&: lStep = -1&
    Else
        Y = imgHeight - 1&: X = 0&: lStep = 1&
    End If
    
    If Compressed = False Then
        ePtr = 128&
        ' size array exactly as needed: 128 header, image data, palette flag, palette
        ReDim arrRLE(0 To 127& + BytesPerLine * imgHeight + palSize)
        For Y = X To Y Step lStep
            CopyMemory arrRLE(ePtr + Y * BytesPerLine), rawStream(X, Y), BytesPerLine
        Next
        
    Else
    
        ' add buffer for an encoded line
        ReDim bEncoded(0 To BytesPerLine * 2)
        ReDim arrRLE(0 To 127& + palSize)                        ' size for header & palette
        
        For Y = X To Y Step lStep
            ePtr = 0&: lCount = 1&
            lPrev = rawStream(ePtr, Y)                          ' set start color index
            For X = 1& To BytesPerLine - 1&
                If rawStream(X, Y) = lPrev Then                 ' increment count as needed
                    If lCount > 62& Then                        ' and add compressed data to buffer
                        bEncoded(ePtr) = &HFF: bEncoded(ePtr + 1&) = lPrev
                        lCount = lCount - 63&: ePtr = ePtr + 2&
                    End If
                    lCount = lCount + 1&
                Else
                    If lCount = 1& Then                         ' add compressed color to buffer
                        If (lPrev And &HC0) = &HC0 Then
                            bEncoded(ePtr) = &HC1: bEncoded(ePtr + 1&) = lPrev
                            ePtr = ePtr + 2&
                        Else
                            bEncoded(ePtr) = lPrev: ePtr = ePtr + 1&
                        End If
                    Else
                        bEncoded(ePtr) = lCount Or &HC0: bEncoded(ePtr + 1&) = lPrev
                        ePtr = ePtr + 2&
                    End If
                    lCount = 1&                                 ' reset count
                    lPrev = rawStream(X, Y)                     ' set new start color index
                End If
            Next
            
            If lCount = 1& Then
                If (lPrev And &HC0) = &HC0 Then
                    bEncoded(ePtr) = &HC1: bEncoded(ePtr + 1&) = lPrev
                    ePtr = ePtr + 2&
                Else
                    bEncoded(ePtr) = lPrev: ePtr = ePtr + 1&
                End If
            Else
                bEncoded(ePtr + 1&) = lPrev
                If lCount > 63& Then                           ' handle improbable scenario
                    bEncoded(ePtr) = &HFF
                    bEncoded(ePtr + 2&) = &HC0 Or (lCount - 63&)
                    bEncoded(ePtr + 3&) = lPrev
                    ePtr = ePtr + 4&
                Else
                    bEncoded(ePtr) = lCount Or &HC0
                    ePtr = ePtr + 2&
                End If
            End If
            
            lCount = UBound(arrRLE)                        ' append encoded data to out array
            ReDim Preserve arrRLE(0 To lCount + ePtr)
            CopyMemory arrRLE(lCount - palSize + 1&), bEncoded(0), ePtr
            
        Next
        Erase bEncoded()
        
    End If
    
    arrRLE(3) = bitCount                                        ' bits per pixel
    arrRLE(65) = 1                                              ' color planes
    CopyMemory arrRLE(66), BytesPerLine, 2&                     ' bytes per plane
    If bitCount = 8& Then
        pvCompressColorMap = UBound(arrRLE) - 767&              ' set location for palette insert
        arrRLE(pvCompressColorMap - 1&) = &HC                   ' set palette flag
    Else
        pvCompressColorMap = 16&
    End If
    
End Function


Private Function pvCompressColorMap_4Planes(arrRLE() As Byte, rawStream() As Byte, Compressed As Boolean, _
                                    ByVal imgWidth As Long, ByVal imgHeight As Long, bitCount As Long, bInvertedSrc As Boolean) As Long
                                    
    ' Function compresses a 4bpp paletted image to pcx standards (1 bpp @ 4 planes)
    ' rawStream passed as 2D palette indexes relative to passed Palette()
    ' Note: bits extracted from GDI+ are top down, not bottom up
    ' Note: Palette() is RGB format @ 3 bytes per color
    
    Dim X As Long, Y As Long, p As Long
    Dim bEncoded() As Byte, bScan() As Byte
    Dim srcPtr As Long, dstPtr As Long
    Dim lScanWidth As Long, BytesPerLine As Long
    Dim lShift As Byte, rShift As Byte
    
    BytesPerLine = imgWidth \ 8 - ((imgWidth Mod 8) > 0) ' scanwidth per plane
    lScanWidth = BytesPerLine * 4&
    ReDim bScan(0 To lScanWidth - 1&)                   ' buffer for 1 uncompressed scanline
    
    If Compressed = False Then                          ' size destination array
        ReDim arrRLE(0 To 127& + lScanWidth * imgHeight) ' exact size here
    Else
        ReDim arrRLE(0 To 127)                          ' header size for now
        ReDim bEncoded(0 To lScanWidth * 2&)            ' buffer of 1 compressed scanline
    End If
    
    If bInvertedSrc Then
        p = imgHeight - 1&: Y = 0&: X = -1&
    Else
        Y = imgHeight - 1&: p = 0&: X = 1&
    End If
    
    'For Y = Y To imgHeight - 1&
    For Y = p To Y Step X
        ' build complete scanline before compressing
        ' RawStream() has 4bpp format same as Windows Bitmap specs
        ' PCX specs take each nibble & splits it into 4 different arrays
        ' Example: for 1 byte (2 palette indexes)...
        '   Source: 7654 3210
        '   Destination per plane (1 thru 4): 40 51 62 73
        rShift = 128: lShift = rShift: srcPtr = 0&: dstPtr = srcPtr
        For X = dstPtr To imgWidth - 1&
            For p = lScanWidth - BytesPerLine To 0& Step -BytesPerLine
                ' strip off highest bits & append to planes 4,3,2,1 in that order
                bScan(dstPtr + p) = bScan(dstPtr + p) Or ((rawStream(srcPtr, Y) And rShift) \ rShift) * lShift
                rShift = rShift \ 2
            Next
            If rShift = 0 Then              ' 8 bits stripped?
                srcPtr = srcPtr + 1&        ' move source to next byte
                rShift = 128                ' reset
            End If
            If lShift = 1 Then              ' compressed byte done?
                lShift = 128                ' reset
                dstPtr = dstPtr + 1&        ' move dest to next byte
            Else
                lShift = lShift \ 2         ' continue packing dest bytes
            End If
        Next
        If Compressed = False Then
            X = 128& + Y * lScanWidth
            CopyMemory arrRLE(X), bScan(0), lScanWidth
        Else
            X = pvCompressMultiPlane(bScan(), bEncoded())
            p = UBound(arrRLE)              ' append encoded data to out array
            ReDim Preserve arrRLE(0 To p + X)
            CopyMemory arrRLE(p + 1&), bEncoded(0), X
        End If
        FillMemory bScan(0), lScanWidth, 0
    Next
    
    arrRLE(65) = 4                                      ' number of color planes
    arrRLE(3) = 1                                       ' bits per pixel
    CopyMemory arrRLE(66), BytesPerLine, 2&             ' bytes per scanwidth
    pvCompressColorMap_4Planes = 16&
    
End Function

Private Sub pvCompressTrueColor(arrRLE() As Byte, rawStream() As Byte, Compressed As Boolean, _
                                    ByVal imgWidth As Long, ByVal imgHeight As Long, srcBitDepth As Long, _
                                    destBitDepth As Long, bInvertedSrc As Boolean)

    
    ' function compresses a true color image (24 or 32 bpp) into standard pcx format
    ' Note: 32bpp may not be recognized by other applications
        
    Dim ePtr As Long, X As Long, Y As Long, p As Long
    Dim bEncoded() As Byte, bScan() As Byte
    Dim lScanWidth As Long, BytesPerLine As Long
    Dim srcBpp As Long, dstBpp As Long
    
    srcBpp = srcBitDepth \ 8                                ' bytes per pixel: source
    dstBpp = destBitDepth \ 8                               ' bytes per pixel: destination
    BytesPerLine = imgWidth + (imgWidth And 1&)
    lScanWidth = imgWidth * srcBpp                          ' scanwidth: source
    ReDim bScan(0 To BytesPerLine * dstBpp - 1&)            ' holds 1 uncompressed scanline
    
    If bInvertedSrc Then
        ePtr = imgHeight - 1&: Y = 0&: X = -1&
    Else
        Y = imgHeight - 1&: ePtr = 0&: X = 1&
    End If
    
    If Compressed = False Then
    
        ' size array exactly as needed: 128 header, image data
        ReDim arrRLE(0 To 127& + BytesPerLine * imgHeight * dstBpp)
        For Y = ePtr To Y Step X
            X = 0&: ePtr = Y * (BytesPerLine * dstBpp) + 128&
            For p = 2& To lScanWidth - 1& Step srcBpp      ' loop thru scanline
                bScan(X) = rawStream(p, Y)                 ' add R,G,B,Alpha channels to PCX color planes
                bScan(X + BytesPerLine) = rawStream(p - 1&, Y)
                bScan(X + BytesPerLine + BytesPerLine) = rawStream(p - 2&, Y)
                If dstBpp = 4& Then bScan(X + BytesPerLine * 3&) = rawStream(p + 1&, Y)
                X = X + 1&                                      ' move PCX pointer along
            Next                                                ' transfer uncompressed planes
            CopyMemory arrRLE(ePtr), bScan(0), BytesPerLine * dstBpp
        Next
        
    Else
    
        ReDim bEncoded(0 To BytesPerLine * dstBpp * 2&)         ' holds 1 compressed scanline
        ReDim arrRLE(0 To 127)                                  ' destination array
        
        For Y = ePtr To Y Step X
            ePtr = 0&
            For X = 2& To lScanWidth - 1& Step srcBpp           ' loop thru scanline
                bScan(ePtr) = rawStream(X, Y)                   ' add R,G,B,Alpha channels to PCX color planes
                bScan(ePtr + BytesPerLine) = rawStream(X - 1&, Y)
                bScan(ePtr + BytesPerLine + BytesPerLine) = rawStream(X - 2&, Y)
                If dstBpp = 4& Then bScan(ePtr + BytesPerLine * 3&) = rawStream(X + 1&, Y)
                ePtr = ePtr + 1&                                ' move PCX pointer along
            Next
            p = pvCompressMultiPlane(bScan(), bEncoded())
            ' transfer data to return array
            X = UBound(arrRLE)
            ReDim Preserve arrRLE(0 To X + p)
            CopyMemory arrRLE(X + 1&), bEncoded(0), p
        Next
        
    End If
    
    arrRLE(3) = 8                                       ' bits per pixel
    arrRLE(65) = dstBpp                                 ' number of color planes
    CopyMemory arrRLE(66), BytesPerLine, 2&             ' bytes per scanwidth
    
End Sub

Private Function pvPalettizeImage(Palette() As Byte, _
                                bmpInfo As BitmapData, outStream() As Byte, _
                                HeaderBytes As Long, FooterBytes As Long) As Long

    ' Function determines if image can be palettized
    ' This does not futher optimize the palette other than reducing palette entry count

    ' Non-alpha only images passed here; validated by the calling routine.
    ' If alpha palette entries are allowed, will modify this routine at that point

    Dim X As Long, Y As Long, lScanWidth As Long
    Dim Color As Long, prevColor As Long
    Dim tPalette() As Long
    Dim newColor As Boolean
    Dim lCol As Long, Index As Long
    Dim lShift As Byte, rShift As Byte, bitCount As Byte
    Dim palCount As Long, srcSA As SafeArray, inStream() As Long

    ' count unique colors (maximum of 256 if we are to palettize)
    ' If any alpha values exist, we abort. Otherwise we continue until we know for sure it is not alpha

    With srcSA
        .cbElements = 4
        .cDims = 2
        .pvData = bmpInfo.Scan0Ptr
        If bmpInfo.stride < 0& Then .pvData = .pvData + (bmpInfo.Height - 1&) * bmpInfo.stride
        .rgSABound(0).cElements = bmpInfo.Height
        .rgSABound(1).cElements = bmpInfo.Width
    End With
    CopyMemory ByVal VarPtrArray(inStream), VarPtr(srcSA), 4&
    On Error GoTo ExitRoutine
    
    ReDim tPalette(1 To 256)    ' palette

    ' force impossible first match for loop below
    prevColor = inStream(0, 0) Xor 1&
    
    For Y = 0& To bmpInfo.Height - 1&
        For X = 0& To bmpInfo.Width - 1&

            Color = inStream(X, Y)              ' get next pixel
            If Not Color = prevColor Then       ' is this same as previous color?
                ' use binary search routine to locate it if it already exists
                Index = modCommon.FindColor(tPalette, Color, palCount, newColor)
                If newColor = True Then         ' color not in palette yet
                    palCount = palCount + 1&    ' increment palette count
                    If palCount = 257& Then     ' maxed palette out?
                        Y = bmpInfo.Height       ' aborts outer loop
                        Exit For
                    Else
                        If Index < palCount Then ' keep palette entries sorted for binary search routine
                            CopyMemory tPalette(Index + 1&), tPalette(Index), (palCount - Index) * 4&
                        End If
                        tPalette(Index) = Color ' add new color to the palette
                    End If
                End If
                prevColor = Color               ' cache current color for faster looping
            End If
        Next
    Next

    If palCount < 257& Then ' we can palettize

        ' build the image indexes array
        Select Case palCount
        Case Is < 3
            If (tPalette(1) = &HFF000000 And tPalette(2) = &HFFFFFFFF) Then
                bitCount = 1: rShift = 2
            Else
                bitCount = 4: rShift = 16
            End If
        Case Is < 17
                bitCount = 4: rShift = 16
        Case Else
                bitCount = 8: rShift = 8
        End Select
        lScanWidth = modCommon.ByteAlignOnWord(bitCount, bmpInfo.Width)
        ReDim outStream(0 To lScanWidth - 1&, 0 To bmpInfo.Height - 1&)

        prevColor = inStream(0, 0) Xor 1&
        For Y = 0 To bmpInfo.Height - 1&
            If bitCount = 8 Then
                lShift = 1
            Else
                If bitCount = 4 Then lShift = 16 Else lShift = 128
            End If
            lCol = 0&
            For X = lCol To bmpInfo.Width - 1&
                Color = inStream(X, Y)
                If prevColor <> Color Then
                    Index = modCommon.FindColor(tPalette(), Color, palCount, False) - 1&
                    prevColor = Color
                End If
                outStream(lCol, Y) = outStream(lCol, Y) Or Index * lShift
                If lShift = 1 Then
                    If bitCount = 4 Then
                        lShift = 16
                    ElseIf bitCount = 1 Then
                        lShift = 128
                    End If
                    lCol = lCol + 1&
                Else
                    lShift = lShift \ rShift
                End If
            Next
        Next

        ' transfer palette to passed byte array (3 pixels per entry, not 4)
        ReDim Palette(0 To palCount * 3& - 1&)
        For X = 0 To palCount - 1
            lCol = X * 3&: Color = tPalette(X + 1)
            Palette(lCol + 2) = Color And &HFF
            Palette(lCol + 1) = (Color And &HFF00&) \ &H100&
            Palette(lCol) = (Color And &HFF0000) \ &H10000
        Next

        pvPalettizeImage = bitCount

    End If
ExitRoutine:
    CopyMemory ByVal VarPtrArray(inStream), 0&, 4&
End Function

Private Function pvProcessColorMap(inStream() As Byte, BytesPerLine As Long, bmpi As BitmapData, cPal As ColorPalette, bRLE As Boolean) As Long

    ' Function converts 1, 4 or 8 bpp paletted PCX (compressed/uncompressed) to GDI+ image
    ' Note: The byte order written to PCX is same as written to BMP. No additional bit shifting
    '   or masking is needed. Uncompress & copy from PCX to BMP as needed

    Dim ptr As Long, lCount As Long, lColor As Byte
    Dim X As Long, Y As Long
    Dim lInc As Long, scanWidth As Long
    Dim bBits() As Byte, bBuffer() As Byte
    Dim maxStream As Long
    Dim tSA As SafeArray, tRect As RECTI
    
    ' create destination GDI+ image
    If GdipCreateBitmapFromScan0(bmpi.Width, bmpi.Height, 0&, bmpi.PixelFormat, ByVal 0&, pvProcessColorMap) Then Exit Function
    tRect.nHeight = bmpi.Height: tRect.nWidth = bmpi.Width
    ' get pointer to its bits & abort if failure
    If GdipBitmapLockBits(pvProcessColorMap, tRect, ImageLockModeWrite, bmpi.PixelFormat, bmpi) Then
        GdipDisposeImage pvProcessColorMap
        pvProcessColorMap = 0&
        Exit Function
    End If
    With tSA                                                    ' overlay array on GDI+ bits
        .cbElements = 1
        .cDims = 2
        .pvData = bmpi.Scan0Ptr
        .rgSABound(0).cElements = bmpi.Height
        .rgSABound(1).cElements = bmpi.stride
    End With
    CopyMemory ByVal VarPtrArray(bBits), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine
    
    ptr = 128&                                                  ' all pcx images start at byte 128
    ' ensure don't overflow dib bpp
    If bmpi.stride > BytesPerLine Then scanWidth = BytesPerLine Else scanWidth = bmpi.stride
            
    If bRLE Then                                                ' compressed pcx
    
        ReDim bBuffer(0 To BytesPerLine + 63&)                  ' large enough for 1 pcx scanline
        maxStream = UBound(inStream)                            ' test for EOF
        For Y = Y To bmpi.Height - 1&
            For X = X To BytesPerLine - 1&                      ' process pcx scanline
                If ptr > maxStream Then                         ' eof check
                    Y = bmpi.Height - 1&: Exit For
                End If
                If (inStream(ptr) And &HC0) = &HC0 Then         ' RLE?
                    If ptr + 1& > maxStream Then                ' eof check
                        Y = bmpi.Height - 1&: Exit For
                    End If
                    lCount = (inStream(ptr) And &H3F)           ' number of RLE bytes
                    FillMemory bBuffer(X), lCount, inStream(ptr + 1&) ' transfer to buffer
                    X = X + lCount - 1&                         ' update bytes read
                    ptr = ptr + 2&                              ' move source pointer along
                Else
                    bBuffer(X) = inStream(ptr)                  ' single RLE
                    ptr = ptr + 1&                              ' move source pointer along
                End If
            Next
            CopyMemory bBits(0, Y), bBuffer(0), scanWidth       ' update dib from buffer
            X = X - BytesPerLine                                ' cross-scanline RLE check...
            If X > 0& Then CopyMemory bBuffer(0), bBuffer(BytesPerLine), X
        
        Next
        
    Else                                                        ' uncompressed pcx
    
        ' validate file has enough bytes for dib
        If BytesPerLine * bmpi.Height + 128& > UBound(inStream) + 1& Then Exit Function
        For Y = Y To bmpi.Height - 1&
            CopyMemory bBits(X, Y), inStream(ptr), scanWidth    ' transfer from source to dib bits
            ptr = ptr + BytesPerLine                            ' move source pointer to next scanline
        Next
    End If
    
ExitRoutine:
    CopyMemory ByVal VarPtrArray(bBits), 0&, 4&                 ' remove overlay
    GdipBitmapUnlockBits pvProcessColorMap, bmpi                ' unlock bits
    If Err Then
        GdipDisposeImage pvProcessColorMap                      ' destroy; failure
        pvProcessColorMap = 0&
    Else
        GdipSetImagePalette pvProcessColorMap, cPal             ' update palette
    End If

End Function

Private Function pvProcess4bpp_4planes(inStream() As Byte, ByVal BytesPerLine As Long, bmpi As BitmapData, cPal As ColorPalette, bRLE As Boolean) As Long

    ' Function converts 4 bpp (1 bit per 4 planes) PCX (compressed/uncompressed) to GDI+ image

    Dim ptr As Long, lCount As Long
    Dim X As Long, Y As Long
    Dim dibPtr As Long, lInc As Long
    Dim bBits() As Byte, bBuffer() As Byte
    
    Dim lScanWidth As Long, maxStream As Long
    Dim rShift As Integer, lShift As Integer
    Dim tSA As SafeArray, tRect As RECTI
    
    ' create destination GDI+ image
    If GdipCreateBitmapFromScan0(bmpi.Width, bmpi.Height, 0&, bmpi.PixelFormat, ByVal 0&, pvProcess4bpp_4planes) Then Exit Function
    tRect.nHeight = bmpi.Height: tRect.nWidth = bmpi.Width
    ' get pointer to its bits & abort if failure
    If GdipBitmapLockBits(pvProcess4bpp_4planes, tRect, ImageLockModeWrite, bmpi.PixelFormat, bmpi) Then
        GdipDisposeImage pvProcess4bpp_4planes
        pvProcess4bpp_4planes = 0&
        Exit Function
    End If
    With tSA                                                ' overlay array on GDI+ bits
        .cbElements = 1
        .cDims = 2
        .pvData = bmpi.Scan0Ptr
        .rgSABound(0).cElements = bmpi.Height
        .rgSABound(1).cElements = bmpi.stride
    End With
    CopyMemory ByVal VarPtrArray(bBits), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine
    
    ptr = 128&                                              ' all pcx images start at byte 128
    lScanWidth = BytesPerLine
    BytesPerLine = BytesPerLine * 4&                        ' number bytes per scanline pcx
    
    If bRLE Then                                            ' compressed pcx
        
        ReDim bBuffer(0 To BytesPerLine + 63&)              ' large enough for 1 pcx scan line
        maxStream = UBound(inStream)                        ' test for EOF
        For Y = Y To bmpi.Height - 1
            
            For X = X To BytesPerLine - 1&                  ' process pcx scanline
                If ptr > maxStream Then                     ' eof check
                    Y = bmpi.Height - 1&: Exit For
                End If
                If (inStream(ptr) And &HC0) = &HC0 Then     ' RLE?
                    If ptr + 1& > maxStream Then            ' eof check
                        Y = bmpi.Height - 1&: Exit For
                    End If
                    lCount = (inStream(ptr) And &H3F)       ' number RLE bytes
                    FillMemory bBuffer(X), lCount, inStream(ptr + 1&) ' transfer to buffer
                    X = X + lCount - 1&                     ' update bytes read
                    ptr = ptr + 2&                          ' move source pointer along
                Else
                    bBuffer(X) = inStream(ptr)              ' single RLE
                    ptr = ptr + 1&                          ' move source pointer along
                End If
            Next
            
            lShift = 128: rShift = lShift                   ' set left & right bit shift values
            lInc = lScanWidth * 2&                          ' set place in buffer to begin reading
            dibPtr = 0&                                     ' set start of dest scanline
            For lCount = dibPtr To bmpi.Width - 1&          ' ok, process
                ' 4 planes, each plane uses 1 bit to be combined with other planes to form 4 bit index btwn 0 & 15
                ' Each byte in each plane contains info for up to 8 pixels
                ' Working from left to right, bitwise...
                '   4th channel bit is bit 8/4 of dest palette index
                '   3rd channel bit is bit 7/3 of dest palette index
                '   2nd channel bit is bit 6/2 of dest palette index
                '   1st channel bit is bit 5/1 of dest palette index
                ' We get that index and either shift it 16 or 0 bits to pack 1 byte, 4bpp dest index
                bBits(dibPtr, Y) = bBits(dibPtr, Y) Or _
                    ((bBuffer(lInc + lScanWidth) And rShift) \ rShift) * lShift Or _
                    ((bBuffer(lInc) And rShift) \ rShift) * lShift \ 2 Or _
                    ((bBuffer(lInc - lScanWidth) And rShift) \ rShift) * lShift \ 4 Or _
                    ((bBuffer(lInc - lScanWidth - lScanWidth) And rShift) \ rShift) * lShift \ 8
                    
                If (lCount And 1&) Then                     ' done with dest byte?
                    lShift = 128                            ' set higher left bit shift
                    dibPtr = dibPtr + 1&                    ' increment dest pixel
                Else
                    lShift = 8                              ' set lower left bit shift
                End If
                If rShift = 1 Then                          ' done with the source byte?
                    rShift = 128                            ' reset right bit shift
                    lInc = lInc + 1&                        ' increment source pixel
                Else
                    rShift = rShift \ 2                     ' decrement right bit shift
                End If
            Next
            X = X - BytesPerLine                            ' cross-scanline RLE check...
            If X > 0& Then CopyMemory bBuffer(0), bBuffer(BytesPerLine), X
        Next
        Erase bBuffer()
    
    Else                                                    ' uncompressed pcx
    
        ' validate file has enough bytes for dib
        If BytesPerLine * bmpi.Height + ptr > UBound(inStream) + 1& Then Exit Function
        For Y = Y To bmpi.Height - 1&
            lShift = 128: rShift = lShift                   ' set left & right bit shift values
            X = Y * BytesPerLine + ptr + lScanWidth * 2&  ' set place in buffer to begin reading
            dibPtr = 0&                                     ' set start of dest scanline
            For lCount = dibPtr To bmpi.Width - 1&                ' ok, process
                bBits(dibPtr, Y) = bBits(dibPtr, Y) Or _
                    ((inStream(X + lScanWidth) And rShift) \ rShift) * lShift Or _
                    ((inStream(X) And rShift) \ rShift) * lShift \ 2 Or _
                    ((inStream(X - lScanWidth) And rShift) \ rShift) * lShift \ 4 Or _
                    ((inStream(X - lScanWidth - lScanWidth) And rShift) \ rShift) * lShift \ 8
                If (lCount And 1&) Then                     ' done with dest byte?
                    lShift = 128                            ' set higher left bit shift
                    dibPtr = dibPtr + 1&                    ' increment dest pixel
                Else
                    lShift = 8                              ' set lower left bit shift
                End If
                If rShift = 1 Then                          ' done with the source byte?
                    rShift = 128                            ' reset right bit shift
                    X = X + 1&                              ' increment source pixel
                Else
                    rShift = rShift \ 2                     ' decrement right bit shift
                End If
            Next
        Next
    End If
    
ExitRoutine:
    CopyMemory ByVal VarPtrArray(bBits), 0&, 4&             ' remove overlay
    GdipBitmapUnlockBits pvProcess4bpp_4planes, bmpi        ' unlock bits
    If Err Then
        GdipDisposeImage pvProcess4bpp_4planes              ' destroy; failure
        pvProcess4bpp_4planes = 0&
    Else
        GdipSetImagePalette pvProcess4bpp_4planes, cPal     ' update palette
    End If

End Function

Private Function pvProcessTrueColor(inStream() As Byte, ByVal BytesPerLine As Long, bmpi As BitmapData, bRLE As Boolean) As Long

    ' Function converts 24/32 bpp PCX (compressed/uncompressed) to GDI+ image

    Dim ptr As Long, lCount As Long, lColor As Byte
    Dim X As Long, Y As Long, Channel As Long
    Dim dibPtr As Long, lInc As Long, bPP As Long
    Dim bBits() As Byte, bBuffer() As Byte, bDIB() As Long
    
    Dim lScanWidth As Long, maxStream As Long
    Dim tSA As SafeArray, vSA As SafeArray, tRect As RECTI
    
    tRect.nHeight = bmpi.Height: tRect.nWidth = bmpi.Width
    If bmpi.PixelFormat = lvicColor32bppAlpha Then
        bPP = 4&
        bmpi.stride = bmpi.Width * 4&
        ReDim bBits(0 To bmpi.stride - 1&, 0 To bmpi.Height - 1&)
        ' with 32bpp images, we don't know yet if it is ARGB, pARGB, or RGB. So use local array
    Else
        bPP = 3&
        If GdipCreateBitmapFromScan0(bmpi.Width, bmpi.Height, 0&, bmpi.PixelFormat, ByVal 0&, pvProcessTrueColor) Then Exit Function
        If GdipBitmapLockBits(pvProcessTrueColor, tRect, ImageLockModeWrite, bmpi.PixelFormat, bmpi) Then
            GdipDisposeImage pvProcessTrueColor
            pvProcessTrueColor = 0&
            Exit Function
        End If
        With tSA
            .cbElements = 1
            .cDims = 2
            .pvData = bmpi.Scan0Ptr
            .rgSABound(0).cElements = bmpi.Height
            .rgSABound(1).cElements = bmpi.stride
        End With
        CopyMemory ByVal VarPtrArray(bBits), VarPtr(tSA), 4&
    End If
    On Error GoTo ExitRoutine
    
    ptr = 128&                                                  ' all pcx images start at byte 128
    BytesPerLine = BytesPerLine * bPP                       ' number bytes per scanline pcx
    If bmpi.stride > BytesPerLine Then lScanWidth = BytesPerLine Else lScanWidth = bmpi.stride
        
    If bRLE Then                                            ' compressed pcx
        
        ReDim bBuffer(0 To BytesPerLine + 63&)              ' large enough for 1 pcx scan line
        maxStream = UBound(inStream)                        ' test for EOF
        For Y = Y To bmpi.Height - 1&
            For X = X To BytesPerLine - 1&                  ' process pcx scanline
                If ptr > maxStream Then                     ' eof check
                    Y = bmpi.Height - 1&: Exit For
                End If
                If (inStream(ptr) And &HC0) = &HC0 Then     ' RLE?
                    If ptr + 1& > maxStream Then            ' eof check
                        Y = bmpi.Height - 1&: Exit For
                    End If
                    lCount = (inStream(ptr) And &H3F)       ' number RLE bytes
                    FillMemory bBuffer(X), lCount, inStream(ptr + 1&) ' transfer to buffer
                    X = X + lCount - 1&                     ' update bytes read
                    ptr = ptr + 2&                          ' move source pointer along
                Else
                    bBuffer(X) = inStream(ptr)              ' single RLE
                    ptr = ptr + 1&                          ' move source pointer along
                End If
            Next
            lInc = 0&                                       ' transfer buffer to dib bits
            For Channel = lInc To bPP - 1&                  ' process Reds, Greens, Blues then Alphas
                ' pcx is RGB but dib bits are BGR, ensure dibptr is correct per channel
                If (Channel And 1&) Then dibPtr = Channel Else dibPtr = Channel Xor 2&
                For dibPtr = dibPtr To lScanWidth - 1& Step bPP
                    bBits(dibPtr, Y) = bBuffer(lInc)
                    lInc = lInc + 1&                        ' move buffer pointer along
                Next                                        ' move buffer pointer to next pcx scanline
                lInc = lInc + ((BytesPerLine \ bPP) * (Channel + 1&) - lInc)
            Next
            X = X - BytesPerLine                            ' cross-scanline RLE check...
            If X > 0& Then
                CopyMemory bBuffer(0), bBuffer(BytesPerLine), X
            End If
        Next
        Erase bBuffer()
    
    Else                                                    ' uncompressed pcx
        
        ' validate file has enough bytes for dib
        If BytesPerLine * bmpi.Height + ptr > UBound(inStream) + 1& Then Exit Function
        For Y = Y To bmpi.Height - 1&
            lInc = ptr                                      ' start position for this scanline
            For Channel = 0& To bPP - 1&                    ' process Reds, Greens, Blues then Alphas
                ' pcx is RGB but dib bits are BGR, ensure dibptr is correct per channel
                If (Channel And 1&) Then dibPtr = Channel Else dibPtr = Channel Xor 2&
                For dibPtr = dibPtr To lScanWidth - 1& Step bPP
                    bBits(dibPtr, Y) = inStream(ptr)
                    ptr = ptr + 1&                          ' move source pointer along
                Next                                        ' move source pointer to next pcx scanline
                ptr = ptr + ((BytesPerLine \ bPP) * (Channel + 1&) - (ptr - lInc))
            Next
        Next
    
    End If
    
ExitRoutine:
    If bmpi.PixelFormat = lvicColor24bpp Then
        CopyMemory ByVal VarPtrArray(bBits), 0&, 4&         ' remove overlay
        GdipBitmapUnlockBits pvProcessTrueColor, bmpi       ' unlock bits
        If Err Then
            GdipDisposeImage pvProcessTrueColor             ' failure; abort
            pvProcessTrueColor = 0&
        End If
    ElseIf Err Then
        pvProcessTrueColor = 0&
        If tSA.pvData Then CopyMemory ByVal VarPtrArray(bBits), 0&, 4&
    Else
        With vSA
            .cbElements = 4: .cDims = 2
            .cLocks = 0: .fFeatures = 0&
            .pvData = VarPtr(bBits(0, 0))
            .rgSABound(0).cElements = bmpi.Height
            .rgSABound(1).cElements = bmpi.Width
        End With
        CopyMemory ByVal VarPtrArray(bDIB), VarPtr(vSA), 4&
        bmpi.PixelFormat = modCommon.ValidateAlphaChannel(bDIB()) ' is data ARGB, pARGB or RGB?
        CopyMemory ByVal VarPtrArray(bDIB), 0&, 4&
        
        bmpi.Scan0Ptr = VarPtr(bBits(0, 0))                 ' transfer bits to newly created GDI+ image
        If GdipCreateBitmapFromScan0(bmpi.Width, bmpi.Height, 0&, bmpi.PixelFormat, ByVal 0&, pvProcessTrueColor) = 0& Then
            If GdipBitmapLockBits(pvProcessTrueColor, tRect, ImageLockModeWrite Or ImageLockModeUserInputBuf, bmpi.PixelFormat, bmpi) Then
                GdipDisposeImage pvProcessTrueColor         ' oops; failure
                pvProcessTrueColor = 0&
            Else
                GdipBitmapUnlockBits pvProcessTrueColor, bmpi   ' unlock bits
            End If
        End If
        CopyMemory ByVal VarPtrArray(bBits), 0&, 4&
    End If
    
End Function

Private Function pvCompressMultiPlane(bScan() As Byte, bEncoded() As Byte) As Long

    Dim lPrev As Byte
    Dim lCount As Long
    Dim ePtr As Long
    Dim X As Long
    
    lPrev = bScan(0)                                    ' now compress the PCX scanline
    lCount = 1&: ePtr = 0&
    For X = 1& To UBound(bScan)                         ' loop thru PCX scanline
        If bScan(X) = lPrev Then
            If lCount > 62& Then
                bEncoded(ePtr) = &HFF: bEncoded(ePtr + 1&) = lPrev
                lCount = lCount - 63&: ePtr = ePtr + 2&
            End If
            lCount = lCount + 1&
        Else
            If lCount = 1& Then
                If (lPrev And &HC0) = &HC0 Then
                    bEncoded(ePtr) = &HC1: bEncoded(ePtr + 1&) = lPrev
                    ePtr = ePtr + 2&
                Else
                    bEncoded(ePtr) = lPrev: ePtr = ePtr + 1&
                End If
            Else
                bEncoded(ePtr) = lCount Or &HC0: bEncoded(ePtr + 1&) = lPrev
                ePtr = ePtr + 2&
            End If
            lCount = 1&
            lPrev = bScan(X)
        End If
    Next
    ' clean up scan line
    If lCount = 1& Then
        If (lPrev And &HC0) = &HC0 Then
            bEncoded(ePtr) = &HC1: bEncoded(ePtr + 1&) = lPrev
            ePtr = ePtr + 2&
        Else
            bEncoded(ePtr) = lPrev: ePtr = ePtr + 1&
        End If
    Else
        bEncoded(ePtr + 1&) = lPrev
        If lCount > 63& Then
            bEncoded(ePtr) = &HFF
            bEncoded(ePtr + 2&) = &HC0 Or (lCount - 63&)
            bEncoded(ePtr + 3&) = lPrev
            ePtr = ePtr + 4&
        Else
            bEncoded(ePtr) = lCount Or &HC0
            ePtr = ePtr + 2&
        End If
    End If
    
    pvCompressMultiPlane = ePtr

End Function
