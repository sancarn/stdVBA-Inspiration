VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsTGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write TGA image format & Convert to Bitmap ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' Targa (TGA) Format specifications retrieved from wotsit.org
' additional info from: http://netghost.narod.ru/gff/graphics/summary/tga.htm

    ' TGA is loosely formatted with hardly any restrictions to force conformity
    ' to its structure. Basically, documentation recommends a user identify TGA by
    ' its extension or v2 footer. But v1 of the structure had no footer and those
    ' files are widely available today, still.
    
    ' FYI. The only real difference btwn a normal bitmap and a true color uncompressed tga is
    ' that bitmap scan lines are word aligned & tga are byte aligned and also the headers:
    ' a bitmap has a 40 byte header & tga has a 44 byte header/footer.
    ' But if tga isn't true color, then that is where the similarities pretty much end, which is
    ' probably why so many routines out on the net only support true color tgas: image type of 2.
    

' TGA header follows (18 byte structure)
' Offset + 0 : 1 byte     ID Length (max of 255 characters that follow the header; can contain anything)
' Offset + 1 : 1 byte     Color Map Type (1=palette included, 0=no palette, other values are reserved)
' Offset + 2 : 1 byte     Image Type (0,1,2,3,9,10,11,32,33). 0=no img, 1&9=paletted, 2&10=true color, 3&11=Grayscale, 32&33 huffman compressed, 9&10&11 are compressed)
' Offset + 3 : 2 bytes    Color Map Index (offset into the data where 1st palette entry is found)
' Offset + 5 : 2 bytes    Number of Map Entries (number of palette entries)
' Offset + 7 : 1 byte     Entry bit depth (bit depth of each palette entry: 8,15,16,24,32)
' Offset + 8 : 2 bytes    X origin where image should be displayed on screen (rarely used)
' Offset + 10: 2 bytes    Y origin where image should be displayed on screen (rarely used)
' Offset + 12: 2 bytes    Image Width (unsigned)
' Offset + 14: 2 bytes    Image Height (unsigned)
' Offset + 16: 1 byte     Pixel bit depth (any value is valid except zero; this class does 8,15,16,24,32)
' Offset + 17: 1 byte     Image Descriptor (image orientation & alpha usage)
' Total 18 bytes

' See pvCreateTGAextensionBlock for description of the Extension Block for V2 TGA format (495 blocks)
' The other optional block for TGAs, v2, is the developer block which is not used (variable size & unlimited)

Private pal32LUT(0 To 31) As Byte  ' 5-bit, 32 entry, palette lookup table

Public Function LoadTGAResource(inStream() As Byte) As Long

    ' NOTE: Must call IsTGAResource first & function must return True

    ' Function processes an array of bytes suspected of containing a TGA file.
    ' Return value is a sourceless GDI+ handle


    ' Parameters:
    ' insSream() :: a byte array containing a TGA
    ' cHost :: an initialized c32bppDIB
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    
    If UBound(inStream) < 57 Then Exit Function
    
    Select Case inStream(2)  ' check image type
    
    Case 1, 9  ' uncompressed/compressed color map (paletted)
        LoadTGAResource = pvProcessColorMap(inStream(), UBound(inStream) + 1&)
    
    Case 2, 3, 10, 11 ' uncompressed/compressed true color / grayscale (easy)
        LoadTGAResource = pvProcessTrueColor(inStream(), UBound(inStream) + 1&)
    
    Case 0, 32, 33 ' won't get here because IsTGAResource already ruled it out; just FYI
        ' 0 unsupported. No image contained; probably not a tga
        ' 32 & 33 unsupported. Uses Huffman Delta compression & couldn't find
        '         documentation to support those types and no code exists within to process them
    End Select

End Function

'Public Function SaveToStream(outStream() As Byte, Handle As Long, ImageType As Long, _
                    Optional useCompression As Boolean = False, _
                    Optional useBitReduction As Boolean = False) As Boolean

Public Function SaveAsTGA(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                          SS As SAVESTRUCT) As Long

    ' saves image as a TGA
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
    
    Dim imgDataLen As Long                          ' size of final image compress/uncompressed
    Dim palType As Long                             ' pallete info if image is palettized
    Dim bPaletted As Boolean, bGrayScale As Boolean ' paletted image types
    Dim aHeaderFooter() As Byte                     ' TGA header or Footer
    Dim aAppBlock() As Byte                         ' optional TGA application block data
    Dim appSize As Long, appOffset As Long          ' application block properties
    Dim theStream() As Byte, aCompressed() As Byte  ' source data and compressed data
    Dim outStream() As Byte
    Dim bAlpha As Boolean, bOK As Boolean
    Dim lResult As Long, srcDepth As Long
    Dim srcSA As SafeArray, dSA As SafeArray
    Dim sizeI As RECTI, tBMP As BitmapData
    Dim cPal As cColorReduction, Compression As Boolean
    Dim tmpPic As StdPicture, tmpDO As DataObject
    Dim tGDIpImage As GDIpImage, tObject As Object

    On Error GoTo ExitRoutine ' begin error trapping

    If SS.CompressionJPGQuality <> lvicFormatUncompressed Then Compression = True
    Set cPal = New cColorReduction
    srcDepth = modCommon.ColorDepthToColorType(SS.ColorDepth, SourceHandle)
    If SS.ColorDepth <> lvicConvert_TrueColor32bpp_pARGB Then
        If returnMedium = saveTo_Clipboard Or returnMedium = saveTo_DataObject Then SS.reserved1 = SS.reserved1 Or &H20000000
        If SS.ColorDepth = lvicDefaultReduction Then
            If SS.ColorDepth = lvicDefaultReduction Then SS.reserved1 = SS.reserved1 Or &H10000000
            If (srcDepth And &HFF00&) \ &H100& > 8& Then
                lResult = cPal.PalettizeToArray(SourceHandle, theStream(), orient_8bppIndexes Or orient_TopDown Or orient_SortGrayscale, 0&, 0&, alpha_Complex, SS)
                If lResult Then srcDepth = lResult
            End If
        ElseIf SS.ColorDepth < lvicConvert_TrueColor24bpp And SS.ColorDepth > lvicDefaultReduction Then
            lResult = cPal.PalettizeToArray(SourceHandle, theStream(), orient_8bppIndexes Or orient_TopDown Or orient_SortGrayscale, 0&, 0&, alpha_Complex, SS)
            If lResult Then srcDepth = lResult Else Exit Function                   ' failure
        End If
    End If
    
    If lResult = 0& Then
        If SS.reserved2 <> 0& And (SS.reserved1 And &HFF00&) \ &H100& = lvicPicTypeTGA Then
            CopyMemory tObject, SS.reserved2, 4&
            Set tGDIpImage = tObject
            CopyMemory tObject, 0&, 4&
            bOK = tGDIpImage.ExtractImageData(outStream)
            If (outStream(2) And 8) \ 8 = Abs(CInt(Compression)) Then
                lResult = srcDepth
            Else
                Set tGDIpImage = Nothing
            End If
        End If
        If lResult = 0& And (srcDepth And &HFF00&) \ &H100& <= 8& Then          ' need to ensure at least 8bpp indexes
            SS.ColorDepth = lvicConvert_256Colors
            SS.PaletteType = lvicPaletteAdaptive
            SS.reserved1 = SS.reserved1 Xor &H10000000
            lResult = cPal.PalettizeToArray(SourceHandle, theStream(), orient_8bppIndexes Or orient_TopDown Or orient_SortGrayscale, 0&, 0&, alpha_Complex, SS)
            If lResult Then srcDepth = lResult Else Exit Function               ' failure
        End If
    End If
    
    If tGDIpImage Is Nothing Then
        ReDim aHeaderFooter(0 To 17)                    ' create tga header
        aHeaderFooter(16) = (srcDepth And &HFF00&) \ &H100&
        If aHeaderFooter(16) = 8 Then
            If (SS.reserved1 And 1&) Then
                aHeaderFooter(17) = 8
                aHeaderFooter(7) = 32
                aHeaderFooter(2) = 1
                aHeaderFooter(1) = 1
                CopyMemory aHeaderFooter(5), cPal.Count, 2&
                palType = 4&
            Else
                If cPal.IsGrayscale() = True And cPal.Count = 256& Then
                    aHeaderFooter(2) = 3        ' grayscale; no palette
                    palType = 0&
                Else
                    aHeaderFooter(7) = 24       ' bits per palette entry
                    aHeaderFooter(2) = 1: aHeaderFooter(1) = 1
                    CopyMemory aHeaderFooter(5), cPal.Count, 2&
                    palType = 3&
                End If
                aHeaderFooter(16) = 8           ' index bit depth
            End If
        Else
            aHeaderFooter(2) = 2                ' true color
            If lResult = 0& Then
                If aHeaderFooter(16) < 8 Then cPal.ImportRemotePalette (SourceHandle)
                If lResult = 0& Then
                    sizeI.nWidth = SS.Width: sizeI.nHeight = SS.Height
                    If GdipBitmapLockBits(SourceHandle, sizeI, ImageLockModeRead, srcDepth, tBMP) Then Exit Function
                    With dSA
                        .cbElements = 1
                        .cDims = 1
                        .pvData = tBMP.Scan0Ptr
                        If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
                        .rgSABound(0).cElements = SS.Height * Abs(tBMP.stride)
                    End With
                    CopyMemory ByVal VarPtrArray(theStream), VarPtr(dSA), 4&
                End If
            End If
        End If
    
        ' Compression can't be used on images less than 3 pixels in width
        ' Not a tga requirement but my own due to the way I coded the compression routine
        If SS.Width < 3& Then Compression = False
        CopyMemory aHeaderFooter(12), SS.Width, 2&
        CopyMemory aHeaderFooter(14), SS.Height, 2&
        If tBMP.stride > -1& Then aHeaderFooter(17) = aHeaderFooter(17) Or 32 ' top down image
        ' ^^ else leave this 0; means bottom up image, left to right which is how DIB is stored
        ' other header items not applicable or will be filled in as routine continues
    
        If Compression Then
            ' pass off to another function to compress the image
            imgDataLen = pvCompressImage(aCompressed(), theStream(), 0&, aHeaderFooter(16) \ 8, (SS.Width), (SS.Height))
            If imgDataLen > UBound(theStream) Then
                ' if compressed image larger than original, use original. Theoretically possible
                imgDataLen = UBound(theStream) + 1&
            Else
                ' transfer compressed data to theStream, but don't overwrite the
                ' original image if theStream is overlayed on it
                If Not dSA.pvData = 0& Then ' theStream is overlayed on original source
                    ' memory hack, we will move the aCompressed pointer to theStream pointer
                    ' thereby making theStream and aCompressed the same, then remove the aCompressed pointer.
                    CopyMemory ByVal VarPtrArray(theStream), ByVal VarPtrArray(aCompressed), 4&
                    CopyMemory ByVal VarPtrArray(aCompressed), 0&, 4&  ' remove aCompressed pointer
                    dSA.pvData = 0& ' prevent last line of code in routine from removing the safe array
                Else
                    CopyMemory theStream(0), aCompressed(0), imgDataLen
                End If
                aHeaderFooter(2) = aHeaderFooter(2) Or 8 ' compression flag
            End If
            Erase aCompressed
    
        Else
            ' no bit reduction, no removal of premultiplied pixels
            imgDataLen = UBound(theStream) + 1
    
        End If
    
        ' create the return array, adding header, palette, image data, application block & footer
    
        ' tga footer consists of 26 bytes
        ' 4 bytes for extension area offset. See pvCreateTGAextensionBlock
        ' 4 bytes for developer area offset. Not used in this project
        ' 18 bytes for TRUEVISION-XFILE. << include period
        ' 1 nullchar terminating byte
    
        ' calc size of array needed
        If aHeaderFooter(16) = 32 Then  ' create extension block for other TGA readers' use
            appSize = pvCreateTGAextensionBlock(aCompressed(), srcDepth = lvicColor32bppAlpha, srcDepth = lvicColor32bppAlphaMultiplied)
            lResult = appSize
        Else
            lResult = 0&
        End If
        If palType Then lResult = lResult + cPal.Count * palType        ' include palette entries
        lResult = imgDataLen + lResult + 44&                            ' include header/footer & image data
        ReDim outStream(0 To lResult - 1&)                              ' size array & copy header
        CopyMemory outStream(0), aHeaderFooter(0), 18&                  ' then create footer
        aHeaderFooter() = StrConv(String$(8, 0) & "TRUEVISION-XFILE." & vbNullChar, vbFromUnicode)
        lResult = 18&                                                   ' offset for writing & write palette if needed
        If palType Then lResult = lResult + cPal.CopyLocalPalette(VarPtr(outStream(18)), False, palType) * palType
        CopyMemory outStream(lResult), theStream(0), imgDataLen         ' write image data
        If dSA.pvData = 0& Then Erase theStream()                       ' free up memory
        If appSize Then                                                 ' write extension block if needed
            lResult = UBound(outStream) - appSize - 25&
            CopyMemory outStream(lResult), aCompressed(0), appSize
            Erase aCompressed()
        End If                                                          ' write footer
        CopyMemory outStream(UBound(outStream) - 25&), aHeaderFooter(0), 26&
        Erase aHeaderFooter()
                                                                        ' release overlays
        If dSA.pvData Then CopyMemory ByVal VarPtrArray(theStream), 0&, 4&: dSA.pvData = 0&
        If tBMP.Scan0Ptr Then GdipBitmapUnlockBits SourceHandle, tBMP: tBMP.Scan0Ptr = 0&
    
    End If
    
    If returnMedium = saveTo_Array Then                             ' set return object
        modCommon.MoveArrayToVariant returnObject, outStream(), True
        SaveAsTGA = True
    ElseIf returnMedium = saveTo_File Then
        WriteFile CLng(returnObject), outStream(0), UBound(outStream) + 1&, lResult, ByVal 0&
        SaveAsTGA = (lResult > UBound(outStream))
    ElseIf returnMedium = saveTo_GDIplus Then
        Set returnObject = modCommon.LoadImage(outStream(), True, , True)
        If returnObject.Handle Then SaveAsTGA = saveTo_GDIplus
    Else
        lResult = 0&
        If tGDIpImage Is Nothing Then Set tGDIpImage = modCommon.LoadImage(outStream(), , , True)
        If tGDIpImage.Handle Then GdipCreateHBITMAPFromBitmap tGDIpImage.Handle, lResult, SS.RSS.FillColorARGB Or &HFF000000
        Set tGDIpImage = Nothing
        If lResult Then
            SaveAsTGA = True
            Set tmpPic = modCommon.HandleToStdPicture(lResult, vbPicTypeBitmap)
            If tmpPic Is Nothing Then
                DeleteObject lResult: SaveAsTGA = False
            ElseIf returnMedium = saveTo_stdPicture Then
                Set returnObject = tmpPic
                SaveAsTGA = (tmpPic.Handle <> 0&)
            ElseIf returnMedium = saveTo_Clipboard Then
                Clipboard.SetData tmpPic
                If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat outStream(), g_ClipboardFormat
            ElseIf returnMedium = saveTo_DataObject Then
                Set tmpDO = returnObject
                tmpDO.SetData tmpPic, vbCFBitmap
                If g_ClipboardFormat Then tmpDO.SetData outStream(), g_ClipboardFormat
                Set tmpDO = Nothing
            End If
            Set tmpPic = Nothing
        End If
    End If
    Erase outStream()

ExitRoutine:
If dSA.pvData Then CopyMemory ByVal VarPtrArray(theStream), 0&, 4&
If tBMP.Scan0Ptr Then GdipBitmapUnlockBits SourceHandle, tBMP
If Err Then
    SaveAsTGA = False
    Err.Clear
End If
End Function

Public Function IsTGAResource(inStream() As Byte, Optional FileHandle As Long = 0&) As Long

    ' Function attempts to rule out whether stream contains a TGA file
    ' Return value = 1 if TGA, -1 if corrupted TGA, 0 if not TGA
    
    ' There are some limitations when the image is paletted:
    ' -- will only support palette indexes of 8 bits, 1 byte
    '    This is not a TGA restriction; rather my own. Until I find or am provided
    '    with sample images other than 1 byte indexes; image is denied

    Dim tData(0 To 23) As Byte
    Dim expectedSize As Long
    Dim Cx As Long, Cy As Long
    Dim tVal As Long, pixBpp As Long, lSize As Long
    
    If FileHandle Then
        lSize = GetFileSize(FileHandle, 0&)
        If lSize < 57& Then Exit Function
        SetFilePointer FileHandle, lSize - 23&, 0&, 0&
        ReadFile FileHandle, tData(0), 23&, tVal, ByVal 0&
    Else
        lSize = UBound(inStream) + 1&
        If lSize < 57& Then Exit Function
        CopyMemory tData(0), inStream(lSize - 22), 23&
    End If
    ' check for v2 TGA footer first
    If InStr(1, StrConv(tData(), vbUnicode), "TRUEVISION-XFILE.", vbTextCompare) > 0 Then IsTGAResource = -1  ' default return value here
    ' if the above is true,
    '   :: we will still do some validation before we return.
    '   :: Parser must return True even if image is invalid; this prevents other parsers
    '           from attempting to process the data, when we know 100% it is a tga
    
    If lSize > 18& Then ' possible to have a one pixel B&W image = 19 bytes
        If FileHandle Then
            ReDim inStream(0 To 17)
            SetFilePointer FileHandle, 0&, 0&, 0&
            ReadFile FileHandle, inStream(0), 18&, tVal, ByVal 0&
            tVal = 0&
        End If
        ' This will attempt to validate both version 1 & 2 tgas before we actually parse image data
    
        ' v1 has no magic number we can use. We will try to exclude & process if we cannot
    
        ' we will look for unusual values or known values that cannot exist. If any are found,
        ' our routines won't process the file, so other parsers can have a chance if applicable
        
        ' Text 1: Validate image type
        Select Case inStream(2)
            Case 1, 2, 3, 9, 10, 11 ' supported
            Case Else ' 0=no image, 32 & 33 may be valid, but no documentation on those types available to me
                Exit Function
        End Select
    
        ' Test 2: Image size
        CopyMemory Cx, inStream(12), 2&
        CopyMemory Cy, inStream(14), 2&
        ' width/height are unsigned; any negative values would indicate absolutely huge image which
        ' we are not prepared to process or simply that this is not a tga file
        If Cx < 1& Or Cy < 1& Then Exit Function ' invalid tga or not a tga
        ' so loosely formatted, this still may not be a TGA format; try a simple overflow test
        On Error Resume Next
        Cx = (Cx * Cy * 4&) \ (Cy * 4&)
        If Err Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
        
        ' Test 3. Ignore unsupported pixel bit depths
        Select Case inStream(16)
            Case 8: pixBpp = 1 ' supported
            Case 15, 16 ' supported if not paletted
                If (inStream(2) And Not 8) = 1 Then Exit Function ' paletted
                pixBpp = 2
            Case 24, 32 ' supported if not paletted
                If (inStream(2) And Not 8) = 1 Then Exit Function ' paletted
                pixBpp = (inStream(16) \ 8)
            Case Else: Exit Function ' per documentation; other bit depths may be valid, but unsupported here
        End Select
        
        ' Test 4: paletted image validations
        If (inStream(2) And Not 8) = 1 Then
            If inStream(1) <> 1 Then Exit Function       ' gotta have a palette if paletted image
        End If
        
        ' Test 5. Validate palette attributes. These can exist in image even if image is not paletted
        If inStream(1) = 1 Then
            CopyMemory expectedSize, inStream(5), 2&   ' are nr entries valid?
            If expectedSize < 1& Then Exit Function
            CopyMemory tVal, inStream(3), 2&           ' is palette offset valid?
            If tVal < 0& Then Exit Function
            expectedSize = expectedSize + tVal                  ' how many indexes are expected
            
            Select Case inStream(7)                    ' ignore unsupported palette bit depths
                Case 8, 24, 32 ' supported
                    expectedSize = expectedSize * (inStream(7) \ 8)
                Case 15, 16 ' supported
                    expectedSize = expectedSize * 2&
                Case Else: Exit Function ' per documentation; nothing else is valid
            End Select
        End If
        
        ' Test 6: Minimal file size expected
        If (inStream(2) And 8) = 0& Then ' uncompressed, else compressed
            expectedSize = expectedSize + (pixBpp * Cx * Cy)
        End If
        If IsTGAResource = -1& Then expectedSize = expectedSize + 26& ' has footer
        If expectedSize + inStream(0) + 18& <= lSize Then IsTGAResource = 1&
        ' file may be much larger, but should be at least the size of expectedSize
        
    End If

End Function

Private Function pvProcessTrueColor(inStream() As Byte, streamLength As Long) As Long

    ' Processes compressed/uncompressed, TrueColor & GrayScale TGA files
    ' Note: the isTGA function already validated key portions of the header; so this routine does no
    ' further validation, other than to ensure compressed images uncompress correctly

    Dim aUncompressed() As Byte
    Dim X As Long, Y As Long, lColorType As Long
    Dim dBytes() As Byte, bBits() As Long
    Dim lScanWidth As Long, dstScanWidth As Long, rowOffset As Long
    Dim bAlpha As Boolean, bFailure As Boolean
    Dim bPP As Long, dstBpp As Long, int16 As Integer
    Dim Width As Long, Height As Long
    
    Dim dstStepX As Long, dstStartX As Long, dstX As Long
    Dim dstStepY As Long, dstStartY As Long, dstStopY As Long
    Dim tBMP As BitmapData, tSize As RECTI
    Dim cPal As ColorPalette, tSA As SafeArray, vSA As SafeArray
    
    dstBpp = inStream(16)
    Select Case dstBpp  ' pixel bit/byte depth
        Case 15: bPP = 2
            pvCreate16bitLUT
            dstBpp = 3: lColorType = lvicColor24bpp
        Case 16: bPP = 2
            pvCreate16bitLUT
            dstBpp = 3: lColorType = lvicColor24bpp
            ' 16bpp can have alpha too; either fully transparent or fully opaque; 1bit value
            ' See if flag is set & verify usage from version 2 tga data if available
            If (inStream(17) And &HF) > 0& Then
                bAlpha = pvAllowAlphaChannel(inStream(), 0, streamLength)
                If bAlpha Then dstBpp = 4: lColorType = lvicColor32bppAlpha
            End If
        Case 24: bPP = 3
            lColorType = lvicColor24bpp: dstBpp = bPP
        Case 32: bPP = 4: dstBpp = bPP
            lColorType = lvicColor32bppAlpha
        Case 8 ' grayscale
            bPP = 1: dstBpp = bPP
            cPal.Count = 256: lColorType = lvicColor8bpp
            For X = 0& To cPal.Count - 1&: cPal.Entries(X + 1&) = &HFF000000 Or X Or X * &H100& Or X * &H10000: Next
    End Select
    
    If inStream(1) = 1 Then    ' color map exists before true color data (unusual but not prohibited)
        CopyMemory int16, inStream(3), 2&           ' number of palette entries to skip
        If int16 < 0& Then Exit Function            ' invalid entry, invlaid TGA
        CopyMemory rowOffset, inStream(5), 2&       ' number of color map entries
        Select Case inStream(7) ' bit depth per entry
        Case 8, 24, 32
            rowOffset = (rowOffset + int16) * (inStream(7) \ 8)
        Case 15, 16
            rowOffset = (rowOffset + int16) * 2
        End Select
        
    End If
    
    CopyMemory Width, inStream(12), 2&             ' get image width
    CopyMemory Height, inStream(14), 2&             ' and image height
    
    lScanWidth = bPP * Width                               ' source/raw scan width
    rowOffset = 18& + inStream(0) + rowOffset ' identify where first image data byte exists
    
    If (inStream(17) And 32) Then   ' top down image
        dstStepY = 1&: dstStartY = 0&: dstStopY = Height - dstStepY
    Else                                            ' bottom up image (like our dib)
        dstStepY = -1&: dstStartY = Height + dstStepY: dstStopY = 0&
    End If
    If (inStream(17) And 16) Then   ' right to left
        dstStepX = -dstBpp: dstStartX = Width * dstBpp - 1&
    Else
        dstStepX = dstBpp: dstStartX = 0&              ' left to right image
    End If
    
    lScanWidth = Width * dstBpp
    dstScanWidth = modCommon.ByteAlignOnWord(bPP * 8, Width)
    
    On Error GoTo ExitRoutine
    If inStream(16) = 32 Then
        If (inStream(2) And 8) = 0 Then
            ' 32bpp uncompressed but may be ARGB, pARGB, RGB. Test it
            With vSA
                .cbElements = 4
                .cDims = 2
                .pvData = VarPtr(inStream(rowOffset))
                .rgSABound(0).cElements = Height
                .rgSABound(1).cElements = Width
            End With
            CopyMemory ByVal VarPtrArray(bBits), VarPtr(vSA), 4&
            lColorType = modCommon.ValidateAlphaChannel(bBits, 0&)
            CopyMemory ByVal VarPtrArray(bBits), 0&, 4&
            tBMP.PixelFormat = lColorType
        Else    ' 32bpp compressed; have to validate channel after decompressing
            ReDim dBytes(0 To dstScanWidth - 1&, 0 To Height - 1&)
        End If
    Else
        tBMP.PixelFormat = lColorType
    End If
    
    If tBMP.PixelFormat Then
        If GdipCreateBitmapFromScan0(Width, Height, 0&, lColorType, ByVal 0&, pvProcessTrueColor) Then Exit Function
        tSize.nHeight = Height: tSize.nWidth = Width
        If GdipBitmapLockBits(pvProcessTrueColor, tSize, ImageLockModeWrite, lColorType, tBMP) Then
            GdipDisposeImage pvProcessTrueColor
            Exit Function
        End If
        With tSA
            .cbElements = 1
            .cDims = 2
            .pvData = tBMP.Scan0Ptr
            .rgSABound(0).cElements = Height
            .rgSABound(1).cElements = tBMP.stride
        End With
        CopyMemory ByVal VarPtrArray(dBytes), VarPtr(tSA), 4&
    End If
    
    If (inStream(2) And 8) = 8 Then        ' compressed
        ReDim aUncompressed(0 To lScanWidth - 1) ' 8,24,32 bpp array
        For Y = dstStartY To dstStopY Step dstStepY
            
            If pvUncompressTrueColor(inStream(), rowOffset, aUncompressed(), bPP, dstBpp, Width) = False Then
                bFailure = True
                Exit For
            End If
            
            If dstStartX = 0 Then                ' left to right source
                CopyMemory dBytes(dstStartX, Y), aUncompressed(0), lScanWidth
            Else                    ' right to left and/or 24bpp source
                dstX = dstStartX
                If dstBpp = 1 Then
                    For X = 0 To lScanWidth Step dstBpp
                        dBytes(dstX, Y) = aUncompressed(X)
                        dstX = dstX + dstStepX
                    Next
                Else
                    For X = 0 To lScanWidth Step dstBpp
                        CopyMemory dBytes(dstX, Y), aUncompressed(X), dstBpp
                        dstX = dstX + dstStepX
                    Next
                End If
            End If
        Next
        Erase aUncompressed()
        If Not bFailure Then
            If inStream(16) = 32 And tBMP.Scan0Ptr = 0& Then
                With vSA
                    .cbElements = 4: .cDims = 2
                    .pvData = VarPtr(dBytes(0, 0))
                    .rgSABound(0).cElements = Height
                    .rgSABound(1).cElements = Width
                End With
                CopyMemory ByVal VarPtrArray(bBits), VarPtr(vSA), 4&
                lColorType = modCommon.ValidateAlphaChannel(bBits(), 0&)
                CopyMemory ByVal VarPtrArray(bBits), 0&, 4&
                If GdipCreateBitmapFromScan0(Width, Height, 0&, lColorType, ByVal 0&, pvProcessTrueColor) Then Exit Function
                tSize.nHeight = Height: tSize.nWidth = Width
                tBMP.Scan0Ptr = VarPtr(dBytes(0, 0))
                tBMP.stride = dstScanWidth
                If GdipBitmapLockBits(pvProcessTrueColor, tSize, ImageLockModeWrite Or ImageLockModeUserInputBuf, lColorType, tBMP) Then
                    GdipDisposeImage pvProcessTrueColor
                    Exit Function
                End If
            End If
        End If
        
    Else ' uncompressed/raw
     
        Select Case bPP
        Case 3, 4 ' 32-24 bits per pixel / 4-3 bytes per pixel
            ' simple & straightforward copy
            If dstStartX = 0 And bPP = 4& Then   ' faster updating when left to right & 32bpp
                If dstStartY = 0& Then  ' even faster if image is bottom up
                    CopyMemory dBytes(0, 0), inStream(rowOffset), lScanWidth * Height
                Else
                    For Y = dstStartY To dstStopY Step dstStepY
                        CopyMemory dBytes(dstX, Y), inStream(rowOffset), lScanWidth
                        rowOffset = rowOffset + lScanWidth
                    Next
                End If
            Else            ' either top-down or right-left source at 1,24,32 bpp
                For Y = dstStartY To dstStopY Step dstStepY
                    dstX = dstStartX
                    For X = 0 To lScanWidth - 1 Step dstBpp
                        CopyMemory dBytes(dstX, Y), inStream(rowOffset + X), dstBpp
                        dstX = dstX + dstStepX
                    Next
                    rowOffset = rowOffset + lScanWidth
                Next
            End If
            
        Case 2 ' 15 or 16 bits per pixel / 2 bytes per pixel
            
            ' some bit shifting required
            lScanWidth = bPP * Width
            For Y = dstStartY To dstStopY Step dstStepY
                dstX = dstStartX
                For X = 0 To lScanWidth - 1 Step bPP
                    ' use a 5,5,5 pattern. Build 15 bits into a 16 bit/2 byte Integer
                    int16 = (inStream(rowOffset + X + 1) And &H7F) * &H100& Or inStream(rowOffset + X)
                    ' shift the 5,5,5 and ref a 5 bit/32 entry palette
                    dBytes(dstX, Y) = pal32LUT((int16 And &H1F))
                    dBytes(dstX + 1, Y) = pal32LUT(((int16 \ &H20) And &H1F))
                    dBytes(dstX + 2, Y) = pal32LUT((int16 \ &H400) And &H1F)
                    If bAlpha Then
                        If (inStream(rowOffset + X + 1) And &H80) = 0 Then dBytes(dstX + 3, Y) = &HFF
                    End If
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + lScanWidth
            Next
        
        Case 1    ' 8 bits per pixel / 1 byte per pixel (grayscale)
            For Y = dstStartY To dstStopY Step dstStepY
                dstX = dstStartX
                For X = 0 To lScanWidth - 1
                    dBytes(dstX, Y) = inStream(rowOffset + X)
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + lScanWidth
            Next
            
        End Select
    
    End If
    
ExitRoutine:
    If tBMP.Scan0Ptr Then GdipBitmapUnlockBits pvProcessTrueColor, tBMP
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(dBytes), 0&, 4&
    If Err Then bFailure = True
    If bFailure Then
        GdipDisposeImage pvProcessTrueColor
        pvProcessTrueColor = 0&
    Else
        If lColorType = lvicColor8bpp Then GdipSetImagePalette pvProcessTrueColor, cPal
    End If
    
    
End Function

Private Function pvProcessColorMap(inStream() As Byte, streamLength As Long) As Long

    ' Processes compressed/uncompressed, paletted TGA images
    ' Note: the isTGA function already validated key portions of the header; so this routine does no
    ' further validation, other than to ensure compressed images uncompress correctly

    Dim bFailure As Boolean
    Dim pal() As Byte, aUncompressed() As Byte
    Dim X As Long, Y As Long, lColorType As Long
    Dim lScanWidth As Long, rowOffset As Long
    Dim bAlpha As Boolean, bPalAlpha As Boolean, lType As Long
    Dim bPP As Long, palBpp As Long
    Dim mapSize As Long, int16 As Integer
    Dim dBytes() As Byte, dSA As SafeArray
    Dim Width As Long, Height As Long
    Dim tBMP As BitmapData, tSize As RECTI, cPal As ColorPalette, tSA As SafeArray
    
    Dim dstStepX As Long, dstStartX As Long, dstX As Long
    Dim dstStepY As Long, dstStartY As Long, dstStopY As Long
    
    'Select Case inStream(16 + Offset) ' source pixel bit/byte depth
    '    Case 8: bPP = 1
    'End Select
    ' currently, only 8bit, 1byte palette indexes are supported
    bPP = 1
    
    CopyMemory mapSize, inStream(5), 2&       ' number of palette entries
    Select Case inStream(7)                   ' bit depth per entry
    Case 8: palBpp = 1
    Case 15: palBpp = 2
        pvCreate16bitLUT
    Case 16: palBpp = 2
        pvCreate16bitLUT
        ' It is possible for the palette to have alpha and the pixels to also have alpha
        ' If pixels have alpha, then their alpha overrides any palette alpha
        bPalAlpha = pvAllowAlphaChannel(inStream(), 0, streamLength)
    Case 24: palBpp = 3
    Case 32: palBpp = 4
        bPalAlpha = pvAllowAlphaChannel(inStream(), 0, streamLength)
    End Select
    
    CopyMemory int16, inStream(3), 2& ' number of palette entries to skip
    rowOffset = (mapSize + int16) * palBpp + 18& + inStream(0) ' this is where the palette indexes begin
    
    CopyMemory Width, inStream(12), 2&             ' get image width
    CopyMemory Height, inStream(14), 2&             ' and image height
    
    If mapSize < 257 Then
        'ReDim pal(0 To 1023)    ' oversize palette to prevent any errors if palette index exceeds palette
    Else
        ReDim pal(0 To mapSize * 4 - 1) ' use actual size of palette (note: can't get here 'cause of 8bpp index restriction above
    End If
    
    ' determine where in the source palette to start reading entries from
    X = 1&: Y = rowOffset - (mapSize * palBpp)
    cPal.Count = mapSize
    For Y = Y To rowOffset - 1 Step palBpp
        Select Case palBpp
        Case 1:
            cPal.Entries(X) = inStream(Y) Or inStream(Y) * &H100& Or inStream(Y) * &H10000 Or &HFF000000
        Case 2
            int16 = (inStream(Y + 1) And &H7F) * &H100& Or inStream(Y)
            ' shift the 5,5,5 and ref a 5 bit/32 entry palette
            cPal.Entries(X) = pal32LUT(int16 And &H1F) Or _
                            pal32LUT((int16 \ &H20) And &H1F) * &H100& Or _
                            pal32LUT((int16 \ &H400) And &H1F) * &H10000 Or &HFF000000
            If bPalAlpha Then               ' validate alpha for each color
                If (inStream(Y + 1) And &H80) Then
                    cPal.Entries(X) = cPal.Entries(X) Xor &HFF000000
                    bAlpha = True
                End If
            End If
        Case 3, 4
            cPal.Entries(X) = inStream(Y) Or inStream(Y + 1) * &H100& Or inStream(Y + 2) * &H10000 Or &HFF000000
            If bPalAlpha Then               ' validate alpha for each color
                Select Case inStream(Y + 3)
                    Case 255 ' do nothing
                    Case 0
                        cPal.Entries(X) = cPal.Entries(X) Xor &HFF000000
                        bAlpha = True
                    Case Is > 127
                        cPal.Entries(X) = (cPal.Entries(X) Xor &HFF000000) Or (inStream(Y + 3) - 128) * &H1000000 Or &H80000000
                        bAlpha = True
                    Case Else
                        cPal.Entries(X) = (cPal.Entries(X) Xor &HFF000000) Or inStream(Y + 3) * &H1000000
                        bAlpha = True
                End Select
            End If
        End Select
        X = X + 1&
    Next
    If bAlpha = False Then bPalAlpha = False
    
    If (inStream(17) And 32) Then   ' top down image
        dstStepY = 1&: dstStartY = 0&: dstStopY = Height - 1&
    Else                                            ' bottom up (like our dib)
        dstStepY = -1&: dstStartY = Height - 1&: dstStopY = 0&
    End If
    If (inStream(17) And 16) Then   ' right to left
        If bPalAlpha Then
            dstStepX = -4&: dstStartX = Width * 4& - 5&
        Else
            dstStepX = -1&: dstStartX = Width - 1&
        End If
    ElseIf bPalAlpha Then
        dstStepX = 4&: dstStartX = 0&               ' left to right
    Else
        dstStepX = 1&: dstStartX = 0&               ' left to right
    End If
    
'    If bPalAlpha = True Then
'        lColorType = lvicColor32bppAlpha
'        tSA.rgSABound(1).cElements = modCommon.ByteAlignOnWord(32, Width)
    If mapSize > 256& Then
        tSA.rgSABound(1).cElements = modCommon.ByteAlignOnWord(24, Width)
        lColorType = lvicColor24bpp
    Else
        tSA.rgSABound(1).cElements = modCommon.ByteAlignOnWord(8, Width)
        lColorType = lvicColor8bpp
    End If
    If GdipCreateBitmapFromScan0(Width, Height, 0&, lColorType, ByVal 0&, pvProcessColorMap) Then Exit Function
    tSize.nHeight = Height: tSize.nWidth = Width
    If GdipBitmapLockBits(pvProcessColorMap, tSize, ImageLockModeWrite, lColorType, tBMP) Then
        GdipDisposeImage pvProcessColorMap
        Exit Function
    End If
    With tSA
        .cbElements = 1
        .cDims = 2
        .pvData = tBMP.Scan0Ptr
        .rgSABound(0).cElements = Height
    End With
    CopyMemory ByVal VarPtrArray(dBytes), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine
    
    If (inStream(2) And 8) = 8 Then        ' compressed
    
        ReDim aUncompressed(0 To Width - 1&) ' 32bpp format
        For Y = dstStartY To dstStopY Step dstStepY
            
            If pvUncompressColorMap(inStream(), rowOffset, aUncompressed(), Width) = False Then
                bFailure = True
                Exit For
            End If
            
            If dstStartX = 0& And lColorType = lvicColor8bpp Then
                ' left to right image, 8bpp
                CopyMemory dBytes(dstX, Y), aUncompressed(0), Width
                
            ElseIf lColorType = lvicColor8bpp Then          ' 8bpp right to left
                dstX = dstStartX
                For X = 0& To Width - 1&
                    dBytes(dstX, Y) = aUncompressed(X)
                    dstX = dstX + dstStepX
                Next
            Else                                            ' 32bpp in either direction
                dstX = dstStartX
                For X = 0& To Width - 1&
                    CopyMemory dBytes(dstX, Y), cPal.Entries(aUncompressed(X) + 1&), 4&
                    dstX = dstX + dstStepX
                Next
            End If
        Next
        Erase aUncompressed()
    
    Else
    
        ' 8 bits per pixel / 1 byte per pixel
        ' copy 1 byte at a time, skip alpha (overwrite it to 255 at end)
        Select Case lColorType
        Case lvicColor32bppAlpha    ' 8 bpp with alpha
            For Y = dstStartY To dstStopY Step dstStepY
                If lScanWidth + rowOffset > streamLength Then
                    bFailure = True
                    Exit For
                End If
                dstX = dstStartX
                For X = 0& To Width - 1&
                    CopyMemory dBytes(dstX, Y), cPal.Entries(inStream(rowOffset + X + 1&)), 4&
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + Width
            Next
        Case lvicColor24bpp
            ' n/a as palette indexes can only be 8 byte (restricting indexes 0-255)
        Case Else           ' 8 bpp
            For Y = dstStartY To dstStopY Step dstStepY
                If lScanWidth + rowOffset > streamLength Then
                    bFailure = True
                    Exit For
                End If
                
                dstX = dstStartX
                For X = 0 To Width - 1
                    dBytes(dstX, Y) = inStream(rowOffset + X)
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + Width
            Next
        End Select
    End If
        
ExitRoutine:
    GdipBitmapUnlockBits pvProcessColorMap, tBMP
    CopyMemory ByVal VarPtrArray(dBytes), 0&, 4&
    If Err Then bFailure = True
    If bFailure Then
        GdipDisposeImage pvProcessColorMap
        pvProcessColorMap = 0&
    Else
        If lColorType = lvicColor8bpp Then
            If bPalAlpha Then cPal.Flags = 1&
            GdipSetImagePalette pvProcessColorMap, cPal
        End If
    End If
End Function

Private Function pvUncompressTrueColor(inStream() As Byte, Offset As Long, bScan() As Byte, bPP As Long, _
                                    dstBpp As Long, Width As Long) As Boolean

    ' RLE uncompression of TGA images: TrueColor only

    Dim pixCount As Long, pixLoc As Long
    Dim int16 As Integer, bReturn As Boolean
    Dim tileStop As Long, tileCx As Long
    Dim X As Long, lScanWidth As Long, dstScanWidth As Long
    
    bReturn = True
    lScanWidth = UBound(inStream)
    dstScanWidth = Width * dstBpp
    
    Do Until pixCount = Width
    
        X = (inStream(Offset) And Not 128)
        If dstBpp * X + pixLoc > dstScanWidth Then
            bReturn = False ' corrupt data or not tga; will exceed scanwidth
            Exit Do
        End If
    
        If (inStream(Offset) And 128) = 0 Then  ' this set of pixels not RLE compressed
            ' when not compressed, the control byte says how many non-compressed pixels to process
            
            If Offset + X * bPP > lScanWidth Then
                bReturn = False ' corrupt data or not tga; will exceed scanwidth
                Exit Do
            End If
            
            Offset = Offset + 1                 ' move to next byte
            Select Case bPP
            Case 4                  ' 32bpp; straight copy
                X = inStream(Offset - 1) + 1
                CopyMemory bScan(pixLoc), inStream(Offset), X * bPP
                pixLoc = pixLoc + X * dstBpp
                Offset = Offset + X * bPP
            Case 3                  ' 24 bpp
                For X = 0 To inStream(Offset - 1)
                    CopyMemory bScan(pixLoc), inStream(Offset), bPP ' copy 3 bytes
                    pixLoc = pixLoc + dstBpp
                    Offset = Offset + bPP       ' move 3 bytes
                Next                            ' already positioned on next RLE control byte
            Case 2                  ' 16 bpp
                For X = 0 To (inStream(Offset - 1)) ' create 15bit integer & shift into palette
                    int16 = (inStream(Offset + 1) And &H7F) * &H100 Or inStream(Offset)
                    bScan(pixLoc) = pal32LUT((int16 And &H1F))
                    bScan(pixLoc + 1) = pal32LUT(((int16 \ &H20) And &H1F))
                    bScan(pixLoc + 2) = pal32LUT((int16 \ &H400) And &H1F)
                    If dstBpp = 4 Then
                        If (inStream(Offset + 1) And &H80) = 0 Then bScan(pixLoc + 3) = &HFF
                    End If
                    pixLoc = pixLoc + dstBpp
                    Offset = Offset + bPP       ' move 2 bytes
                Next                            ' already positioned on next RLE control byte
            Case 1                  ' 8 bpp/grayscale
                For X = 0 To (inStream(Offset - 1))
                    bScan(pixLoc) = inStream(Offset)    ' add grayscale values
                    pixLoc = pixLoc + 1
                    Offset = Offset + bPP       ' move to next byte
                Next                            ' already positioned on next RLE control byte
            End Select
            
            pixCount = pixCount + X
            
        Else                                    ' RLE compressed. TGA does RLE runs on pixels; not bits/bytes
            
            If Offset + bPP > lScanWidth Then
                bReturn = False ' corrupt data or not tga; will exceed scanwidth
                Exit Do
            End If
            
            Select Case bPP
            Case 3, 4                           ' 24/32bpp
                CopyMemory bScan(pixLoc), inStream(Offset + 1), bPP ' get color from next bytes
            Case 2
                int16 = (inStream(Offset + 2) And &H7F) * &H100 Or inStream(Offset + 1)
                ' shift the 5,5,5 and ref a 5 bit/32 entry palette
                bScan(pixLoc) = pal32LUT((int16 And &H1F))
                bScan(pixLoc + 1) = pal32LUT(((int16 \ &H20) And &H1F))
                bScan(pixLoc + 2) = pal32LUT((int16 \ &H400) And &H1F)
                If dstBpp = 4 Then
                    If (inStream(Offset + 1) And &H80) = 0 Then bScan(pixLoc + 3) = &HFF
                End If
            Case 1
                bScan(pixLoc) = inStream(Offset + 1)
            End Select
            
            ' here we use tiling techniques to quickly fill a scan line with repeated pixels
            tileCx = dstBpp                              ' how many pixels copied so far
            tileStop = ((inStream(Offset) And Not 128) + 1) * tileCx ' how many pixels will be duplicated
            If tileStop + pixLoc > dstScanWidth Then
                bReturn = False             ' corrupt data; copying will exceed scanwidth
                Exit Do
            End If
            
            Do Until tileCx + tileCx > tileStop     ' copy another group of pixels
                CopyMemory bScan(pixLoc + tileCx), bScan(pixLoc), tileCx
                tileCx = tileCx * 2                 ' increment group by x2
            Loop
            tileCx = tileStop - tileCx              ' see if any partial tiling is needed
            If tileCx Then                          ' and copy required remaining pixels
                CopyMemory bScan(pixLoc + tileStop - tileCx), bScan(pixLoc), tileCx
            End If
            pixLoc = pixLoc + tileStop      ' move bScan() pointer to next byte to be filled
            Offset = Offset + bPP + 1       ' move to next RLE control byte
            pixCount = pixCount + tileStop \ dstBpp ' number of pixels processed during decompression
        End If
        
    Loop
    pvUncompressTrueColor = bReturn

End Function

Private Function pvUncompressColorMap(inStream() As Byte, Offset As Long, bScan() As Byte, Width As Long) As Boolean

    ' RLE uncompression of TGA images: Paletted TGA only
    
    Dim pixCount As Long
    Dim bReturn As Boolean, lSourceLen As Long
    Dim X As Long, lLen As Long
    
    bReturn = True
    lSourceLen = UBound(inStream)
    Do Until pixCount = Width
        
        lLen = (inStream(Offset) And Not 128) + 1&
        If lLen + pixCount > Width Then
            bReturn = False ' corrupt data or not tga; will exceed scanwidth
            Exit Do
        End If
    
        If (inStream(Offset) And 128) = 0 Then
            ' not RLE encoded
            If Offset + 1& + lLen > lSourceLen Then
                bReturn = False
                Exit Do
            End If
            CopyMemory bScan(pixCount), inStream(Offset + 1), lLen
            Offset = Offset + lLen + 1&
        Else
            If Offset + 1& > lSourceLen Then
                bReturn = False
                Exit Do
            End If
            FillMemory bScan(pixCount), lLen, inStream(Offset + 1)
            Offset = Offset + 2
        End If
        pixCount = pixCount + lLen
    
    Loop
    pvUncompressColorMap = bReturn
    
End Function

Private Sub pvCreate16bitLUT()
    ' Initialize a 32 entry palette
    Dim X As Byte
    If pal32LUT(31) = 0 Then ' not yet created
        For X = 0 To 31
            pal32LUT(X) = X * 8 + (X Mod 8)
        Next
    End If
    
End Sub

Private Sub pvFormatTrueColor(inStream() As Long, outStream() As Byte, bmpInfo As BitmapData, _
                           HeaderBytes As Long, FooterBytes As Long)

    ' Function converts to 24bpp
    
    Dim X As Long, Y As Long
    Dim tOffset As Long
    Dim Color As Byte

    ' convert to 24bpp
    ReDim outStream(0& To bmpInfo.Width * bmpInfo.Height * 3& + HeaderBytes + FooterBytes - 1&)
    tOffset = HeaderBytes
    For Y = 0& To bmpInfo.Height - 1&
        For X = 0& To bmpInfo.Width - 1&
            CopyMemory outStream(tOffset), inStream(X, Y), 3&
            tOffset = tOffset + 3&
        Next
    Next

End Sub

Private Function pvPalettizeImage(Palette() As Byte, inStream() As Long, IsGrayscale As Boolean, _
                                bmpInfo As BitmapData, outStream() As Byte, _
                                HeaderBytes As Long, FooterBytes As Long) As Boolean

    ' Function determines if image can be palettized
    ' This does not futher optimize the palette other than reducing palette entry count

    ' Non-alpha only images passed here; validated by the calling routine.
    ' If alpha palette entries are allowed, will modify this routine at that point

    Dim X As Long, Y As Long
    Dim Color As Long, prevColor As Long
    Dim tPalette() As Long
    Dim newColor As Boolean
    Dim lCol As Long, Index As Long
    Dim palCount As Long

    ' count unique colors (maximum of 256 if we are to palettize)
    ' If any alpha values exist, we abort. Otherwise we continue until we know for sure it is not alpha

    ReDim tPalette(1 To 256)    ' palette
    IsGrayscale = True          ' default until proven otherwise

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
                        If IsGrayscale Then     ' validate grayscale potential
                            If (Color And &HFF&) = ((Color And &HFF00&) \ &H100&) Then
                                If Not (Color And &HFF&) = ((Color And &HFF0000) \ &H10000) Then IsGrayscale = False
                            Else
                                IsGrayscale = False
                            End If
                        End If
                    End If
                End If
                prevColor = Color               ' cache current color for faster looping
            End If
        Next
    Next

    If palCount < 257& Then ' we can palettize

        ' build the image indexes array
        ReDim outStream(0 To bmpInfo.Width * bmpInfo.Height + HeaderBytes + FooterBytes - 1)
        lCol = HeaderBytes      ' where to begin writing palette indexes

        If IsGrayscale Then     ' no palette required, the grayscale values are the indexes
            For Y = 0 To bmpInfo.Height - 1&
                For X = 0 To bmpInfo.Width - 1&
                    outStream(lCol) = (inStream(X, Y) And &HFF&)
                    lCol = lCol + 1&
                Next
            Next
            Erase Palette()

        Else

            ' force impossible first match for loop below
            prevColor = inStream(0, 0) Xor 1&
            For Y = 0& To bmpInfo.Height - 1&
                For X = 0& To bmpInfo.Width - 1&
                    Color = inStream(X, Y)              ' get next color
                    If Not Color = prevColor Then       ' use binary search to find it if needed
                        ' when found, convert 1 base to 0 base indexing
                        Index = modCommon.FindColor(tPalette(), Color, palCount, False) - 1&
                        prevColor = Color               ' cache for faster looping
                    End If
                    outStream(lCol) = Index             ' write the index
                    lCol = lCol + 1                     ' increment next index location
                Next
            Next
            ' transfer palette to passed byte array (3 pixels per entry, not 4)
            ReDim Palette(0 To palCount * 3 - 1)
            For X = 0 To palCount - 1
                CopyMemory Palette(X * 3), tPalette(X + 1), 3&
            Next
        End If

        pvPalettizeImage = True

    End If

End Function

Private Function pvCompressImage(arrRLE() As Byte, rawStream() As Byte, Offset As Long, bPP As Long, _
                                imgWidth As Long, imgHeight As Long) As Long

    ' FYI: RLE = Run Lengh Encoding
    ' Simple RLE rules:
    '   Runs will not cross scanlines; restricted to one scanline only
    '   Control Byte high bit, if set, means compressed run, else raw/uncompressed run
    '   -- Max run length is 128
    '   -- (Byte And Not 128)+1 = length of run
    '   -- One control byte precedes every RLE run
    '   If compressed, one pixel is added to RLE, these are the bytes to be repeated
    '   If uncompressed, next n pixels are added to RLE, uncompressed bytes

    ' Function returns length of compressed image which may be smaller than the array size
    ' If failure; function returns value > raw stream size

    Dim X As Long, Y As Long
    Dim Mode As Long, bPurge As Boolean
    Dim prevPixel As Long, curPixel As Long
    Dim locRLE As Long, locSource As Long
    Dim curCount As Long, maxBytes As Long
    Dim scanWidth As Long, scanWidthWord As Long
    ' Mode: 0=undefined; 1=compressing, 2=not compressing; 4=toggle compression; 8=reset to undefined

    maxBytes = UBound(rawStream)
    ReDim arrRLE(0 To maxBytes)                ' initial size of RLE encoded bytes
    scanWidth = imgWidth * bPP                          ' raw scan line width
    scanWidthWord = modCommon.ByteAlignOnWord(bPP * 8&, imgWidth)

On Error GoTo ExitRoutine
    For Y = 0& To imgHeight - 1&
        locSource = Y * scanWidthWord                       ' source start position
        CopyMemory prevPixel, rawStream(locSource), bPP     ' get 1st pixel for comparison
        curCount = 0&: Mode = 0&
        For X = locSource + bPP To locSource + scanWidth - 1& Step bPP
            CopyMemory curPixel, rawStream(X), bPP          ' get next pixel to compare
            If curCount = 127& Then                         ' max run reached?
                bPurge = True: Mode = Mode Or 8&
            ElseIf curPixel = prevPixel Then                ' compression capable
                If Mode = 1& Then                           ' already compressing?
                    curCount = curCount + 1&                '   increment count
                ElseIf Mode = 2& Then                       ' wasn't compressing; but now will be
                    Mode = Mode Or 4&: bPurge = True        ' flag to toggle mode & write current run
                    curCount = curCount - 1&
                Else                                        ' undefined; now is defined: compressing
                    Mode = 1&: curCount = curCount + Mode
                End If
            ElseIf Mode = 1& Then                           ' was compressing; not now
                Mode = Mode Or 8&: bPurge = True            ' write current run & reset mode to undefined
            Else                                            ' undefined or defined as non-compression
                Mode = 2&: curCount = curCount + 1&         ' update count & ensure mode is set
                prevPixel = curPixel                        ' set new comparison pixel
            End If
            
            If bPurge Then                                  ' writing the run
                arrRLE(locRLE) = curCount
                If (Mode And 1&) Then                       ' compression
                    arrRLE(locRLE) = curCount Or 128: curCount = bPP
                Else                                        ' non-compressed
                    arrRLE(locRLE) = curCount: curCount = bPP * curCount + bPP
                End If
                If locRLE + curCount + 1& > maxBytes Then   ' ensure we have enough bytes in buffer
                    locRLE = maxBytes + 1&: Y = imgHeight   ' else abort
                    Exit For
                End If                                      ' copy compressed/uncompressed bytes
                CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), curCount
                locRLE = locRLE + curCount + 1&             ' update next writable byte in compressed buffer
                curCount = 0&: bPurge = False               ' reset
                locSource = X                               ' update/track
                prevPixel = curPixel                        ' set new comparison pixel
                If (Mode And 4&) Then                       ' toggling mode
                    Mode = 3& - (Mode And &H3&)             ' if 1 now 2; else 1
                    If Mode = 1& Then curCount = curCount + Mode
                ElseIf (Mode And 8&) Then                   ' resetting to undefined
                    Mode = 0&
                End If
            End If
        Next
        If locRLE + 1& > maxBytes Then Exit For
                                                            ' finish off the current scanline
        If (Mode And 1&) Then                               ' same rules as above
            arrRLE(locRLE) = curCount Or 128: curCount = bPP
        Else
            arrRLE(locRLE) = curCount: curCount = bPP * curCount + bPP
        End If
        If locRLE + curCount + 1& > maxBytes Then
            locRLE = maxBytes + 1&: Exit For
        End If
        CopyMemory arrRLE(locRLE + 1), rawStream(locSource), curCount
        locRLE = locRLE + curCount + 1&
        
    Next

    pvCompressImage = locRLE  ' return RLE length

ExitRoutine:
    If Err Then pvCompressImage = maxBytes + 1&
End Function

Private Function pvCreateTGAextensionBlock(outStream() As Byte, isAlpha As Boolean, preMultipliedAlhpa As Boolean) As Long

    ' used only if saving TGA in pre-multiplied alpha format
    
    ' Offset +   0:   2 bytes     length of block. Always 495 for V2 tga extension
    ' Offset +   2:  41 bytes     author name; must be null terminated ascii text
    ' Offset +  43: 324 bytes     author comments; 4 lines of 81 ascii characters; 81st is null char
    ' Offset + 367:  12 bytes     date field in 2 byte entries: m d y h n s
    ' Offset + 379:  41 bytes     job name/id; 40 ascii characters, 41st is null char
    ' Offset + 420:   6 bytes     job time in format: 2 bytes for hours (0-65535), 2 for mins (0-59), 2 for secs (0-59)
    ' Offset + 426:  41 bytes     software creation id; 40 ascii characters; 41st is null char
    ' Offset + 467:   3 bytes     software version; similar to version, major, minor
    ' Offset + 470:   4 bytes     key color (default transparent bkg color if transparency not supported. vb:BGRA)
    ' Offset + 474:   4 bytes     pixel aspect ratio: 2 bytes for numerator & 2 for denominator. Ability to indicate non-square pixels
    ' Offset + 478:   4 bytes     gamma ratio: 2 bytes for numerator & 2 for denominator. Should be in range of 1.0 to 10.0
    ' Offset + 482:   4 bytes     gamma correction offset to a gamma correction table. Size is 256 * 4 bytes
    ' Offset + 486:   4 bytes     postage stamp offset to where mini image pixel data located; mini-image must not be compressed
    ' Offset + 490:   4 bytes     scanline offset to where scan line table begins; table = 4 bytes * image height
    ' Offset + 494:   1 byte      attribute type: 0=no alpha; 1=undefined alpha field & ignore
    '                                             2=undefined alpha field & retain
    '                                             3=useful alpha field data; 4=premultiplied alpha; others are reserved/not used
    
    Dim bBlockData() As Byte, iVal As Integer
    Const v2BlockLength As Long = 495     ' Version 2 Extension block is 495 bytes
    
    ReDim outStream(0 To v2BlockLength - 1)
    CopyMemory outStream(0), v2BlockLength, 2&
    ' add software ID info
    bBlockData = StrConv("LaVolpe Alpha Image Control", vbFromUnicode)
    CopyMemory outStream(426), bBlockData(0), UBound(bBlockData) + 1
    outStream(467) = 3: outStream(469) = 1
    
    If preMultipliedAlhpa = True Then   ' add info about alpha blending used
        outStream(494) = 4
    ElseIf isAlpha Then ' else zero for no alpha
        outStream(494) = 3
    End If
    
    CopyMemory outStream(470), -1&, 4&  ' add a default bkg color if other apps don't support alpha (White, fully opaque)
    
    iVal = Month(Date): CopyMemory outStream(367), iVal, 2& ' add current system date/time
    iVal = Day(Date): CopyMemory outStream(367 + 2), iVal, 2&
    iVal = Year(Date): CopyMemory outStream(367 + 4), iVal, 2&
    iVal = Hour(Now()): CopyMemory outStream(367 + 6), iVal, 2&
    iVal = Minute(Now()): CopyMemory outStream(367 + 8), iVal, 2&
    iVal = Second(Now()): CopyMemory outStream(367 + 10), iVal, 2&
    
    pvCreateTGAextensionBlock = v2BlockLength
    
End Function

Private Function pvAllowAlphaChannel(theStream() As Byte, Offset As Long, length As Long) As Boolean

    ' Helper Function. Called when 16bpp images/palettes contain an alpha bit and for 32bpp palettes.
    
    ' This function will check if version 2 data exists and if so whether or not v2 says
    ' to ignore the alpha channel. Only if v2 exists and ignores the alpha will this
    ' function return False. 32bpp true color images are handled via modParsers.iparseValidateAlphaChannel
    
    Dim tData(0 To 29) As Byte
    Dim tVal As Long, appBlockLoc As Long
    Dim bReturn As Boolean
    
    Const v2BlockLength As Long = 495 ' v2 application/extension block length
    
    bReturn = True      ' default - don't override & allow any alpha
    If length > 43& Then ' check for v2 TGA footer first
        CopyMemory tData(0), theStream(Offset + length - 29&), 30&
        tVal = InStr(1, StrConv(tData(), vbUnicode), "TRUEVISION-XFILE.", vbTextCompare)
        If tVal > 8& Then
            ' ok, does it have an application/extension block? 8 bytes before TRUEVISION
            CopyMemory appBlockLoc, tData(tVal - 9&), 4&
            ' ensure it is not corrupt; must be a valid pointer & must be within the stream/file
            If appBlockLoc > 0& And appBlockLoc + Offset + v2BlockLength < length Then
                ' ok, it appears valid, double check
                CopyMemory tVal, theStream(Offset + appBlockLoc), 2&
                If tVal = v2BlockLength Then ' we have v2 app data; expected
                    ' fine, let's get the overriding value
                    bReturn = (theStream(Offset + appBlockLoc + v2BlockLength - 1&) > 2)
                    ' If the byte = 0, 1 or 2; then alpha is not to be used
                End If
            End If
        End If
    Else
        
    End If
    pvAllowAlphaChannel = bReturn

End Function


