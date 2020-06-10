VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTGAparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write TGA image format & Convert to Bitmap ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' TGA Format specifications retrieved from wotsit.org
' additional info from: http://netghost.narod.ru/gff/graphics/summary/tga.htm

    ' TGA is loosely formatted with hardly any restrictions to force conformity
    ' to its structure. Basically, documentation recommends a user identify TGA by
    ' its extension or v2 footer. But v1 of the structure had no footer and those
    ' files are widely available today, still.
    
    ' FYI. The only real difference btwn a normal bitmap and a true color uncompressed tga is
    ' that bitmap scan lines are word aligned & tga are byte aligned and also the the headers:
    ' a bitmap has a 40 byte header & tga has a 44 byte header/footer.
    ' But if tga not true color, then that is where the similarities pretty much end, which is
    ' probably why so many routines out on the net only support true color tgas: image type of 2.
    

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND ' reusable UDT for 1 & 2 dim arrays
End Type

' TGA header follows (18 byte structure)
' Offset + 0 : 1 byte     ID Length (max of 255 characters that follow the header; can contain anything)
' Offset + 1 : 1 byte     Color Map Type (1=palette included, 0=no palette, other values are reserved)
' Offset + 2 : 1 byte     Image Type (0,1,2,3,9,10,11,32,33). 0=no img, 1&9=paletted, 2&10=true color, 3&11=Grayscale, 32&33 huffman compressed, 9&10&11 are compressed)
' Offset + 3 : 2 bytes    Color Map Index (offset into the palette where 1st palette entry is found)
' Offset + 5 : 2 bytes    Number of Map Entries (number of palette entries)
' Offset + 7 : 1 byte     Entry bit depth (bit depth of each palette entry: 8,15,16,24,32)
' Offset + 8 : 2 bytes    X origin where image should be displayed on screen (rarely used)
' Offset + 10: 2 bytes    Y origin where image should be displayed on screen (rarely used)
' Offset + 12: 2 bytes    Image Width (unsigned)
' Offset + 14: 2 bytes    Image Height (unsigned)
' Offset + 16: 1 byte     Pixel bit depth (any value is valid except zero; this class does 8,15,16,24,32)
' Offset + 17: 1 byte     Image Descriptor (image orientation & alpha usage)
' Total 18 bytes

' See CreateTGAextensionBlock for description of the Extension Block for V2 TGA format (495 blocks)
' The other optional block for TGAs, v2, is the developer block which is not used (variable size & unlimited)

Private pal32LUT(0 To 31) As Byte  ' 5-bit, 32 entry, palette lookup table

Public Function LoadStream(inStream() As Byte, cHost As c32bppDIB, _
                    Optional ByVal streamOffset As Long, Optional ByVal streamLength As Long) As Boolean

    ' Function processes an array of bytes suspected of containing a TGA file.
    ' Note: If desired, process filenames by simply opening the file and caching the bytes in an array
    '   and then pass the array to this function, which is what c32bppDIB.LoadPicture_File does.

    ' Parameters:
    ' insSream() :: a byte array containing a TGA
    ' cHost :: an initialized c32bppDIB
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    
    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim bOk As Boolean
    If isTGA(inStream, streamOffset, streamLength, bOk) = False Then Exit Function
    
    ' definitely a tga (v2 footer was included); else potentially a tga (no footer to guarantee)
    
    If bOk Then ' passed key validation checks; assume it is a tga & process it
    
        Select Case inStream(2 + streamOffset) ' check image type
        
        Case 1, 9  ' uncompressed/compressed color map (paletted)
            ProcessColorMap inStream(), streamOffset, streamLength, cHost
        
        Case 2, 3, 10, 11 ' uncompressed/compressed true color / grayscale (easy)
            ProcessTrueColor inStream(), streamOffset, streamLength, cHost
        
        Case 0, 32, 33 ' won't get here because isTGA already ruled it out; just FYI
            ' 0 unsupported. No image contained; probably not a tga
            ' 32 & 33 unsupported. Uses Huffman Delta compression & couldn't find
            '         documentation to support those types and no code exists within to process them
        End Select
        
        LoadStream = (cHost.Handle <> 0&)
        
    End If
    
End Function

Friend Function SaveAsTGA(ByVal FileName As String, outStream() As Byte, cHost As c32bppDIB, _
                    Optional ByVal allowPreMultipliedAlpha As Boolean = True, _
                    Optional ByVal useCompression As Boolean = False, _
                    Optional ByVal useBitReduction As Boolean = False) As Boolean

    ' Function converts & saves a 32bpp image to TGA format, either as a file or array/stream
    
    ' Parameters:
    ' FileName. Provide the file name to save to, if not provided then outStream is populated
    ' outStream. Will be populated only if FileName is vbNullString
    ' cHost. The c32bppDIB class containing the image to convert to TGA
    ' allowPreMultipliedAlpha. Note: Valid apps that use TGA should be able to read premultiplied pixels
    '   If true, premultiplied pixels are written
    '   If false, pixels are converted to non-premultiplied
    ' useCompression. Compression slightly slows down saving routine & slows down reading; use if size is important
    '   If true, image will first be compressed using tga supported compression algorithms
    '   If false, image will not be compressed
    ' useBitReduction. If true, image will be converted to 256 color, 8 bpp paletted image if possible
    '   and grayscale images will be reduced to 8 bpp format.
    
    ' Bit Reduction Example: 256x256 non-alpha image: non-alpha being the key words
    '   :: 32bpp format, no bit reduction: 256 * 256 * 4 = 262,144 bytes
    '   :: 24bpp format equates to 256 * 256 * 3 = 196,608 bytes (tga is not word aligned)
    ' -- by default. All non-alpha images will be reduced to 24bpp regardless of options. Palettizing is optional
    '   :: if 256 colors or less, 256 * 256 * 1 + (paletteSize*3 for palette) = 66,304 bytes max
    '   :: if grayscale, 256 * 256 * 1 = 65,536 bytes (tga uses no palette for grayscale)
    ' Add 44 bytes for the header+footer and that is the minimum file size required

    If cHost Is Nothing Then Exit Function
    If cHost.Handle = 0& Then Exit Function
        
    ' we pass off the work to another function which will return a sized array containing the TGA
    On Error GoTo ExitRoutine
    If FileName = vbNullString Then                     ' return array to user
        SaveAsTGA = SaveToStream(outStream(), cHost, allowPreMultipliedAlpha, useCompression, useBitReduction)
    Else
        Dim hFile As Long, lRead As Long                ' save to a file
        hFile = iparseGetFileHandle(FileName, False)
        If (hFile = INVALID_HANDLE_VALUE) Then Exit Function
        If SaveToStream(outStream(), cHost, allowPreMultipliedAlpha, useCompression, useBitReduction) = True Then
            WriteFile hFile, outStream(0), UBound(outStream) + 1, lRead, ByVal 0&
            SaveAsTGA = (lRead = UBound(outStream) + 1)
        End If
    End If

ExitRoutine:
If Not hFile = 0& Then CloseHandle hFile
If Err Then Err.Clear

End Function

Private Function SaveToStream(outStream() As Byte, cHost As c32bppDIB, _
                    Optional allowPreMultipliedAlpha As Boolean = True, _
                    Optional useCompression As Boolean = False, _
                    Optional useBitReduction As Boolean = False) As Boolean

    ' Function converts a 32bpp image to TGA format and returns the array containing its format
    
    ' This is the workhorse of the SaveAsTGA function. See SaveAsTGA for parameter descriptions
    
    Dim imgDataLen As Long                          ' size of final image compress/uncompressed
    Dim palSize As Long, aPalette() As Byte         ' pallete info if image is palettized
    Dim bPaletted As Boolean, bGrayScale As Boolean ' paletted image types
    Dim aHeaderFooter() As Byte                     ' TGA header or Footer
    Dim aAppBlock() As Byte                         ' optional TGA application block data
    Dim appSize As Long, appOffset As Long          ' application block properties
    Dim theStream() As Byte, aCompressed() As Byte  ' source data and compressed data
    Dim dSA As SafeArray                            ' DMA overlay if needed
    Dim bAlpha As AlphaTypeEnum
        
    ' Compression can't be used on images less than 3 pixels in width
    ' Not a tga requirement but my own due to the way I coded the compression routine
    If cHost.Width < 3& Then useCompression = False
        
    On Error GoTo ExitRoutine ' begin error trapping
        
    ReDim aHeaderFooter(0 To 17)                    ' create tga header
    iparseOverlayHost_2DbyHost theStream(), VarPtr(dSA), cHost
    iparseValidateAlphaChannel theStream(), False, bAlpha, 0&
    iparseOverlayHost_2DbyHost theStream(), 0&, Nothing
    
    If useBitReduction = True And bAlpha = AlphaNone Then
        ' see if we can palletize or reduce to 8 bpp grayscale
        
        bPaletted = PalettizeImage(aPalette(), bGrayScale, cHost, theStream(), 0, 0)
        If bPaletted Then               ' can palettize & palette reutrned & grayscale flag
            If bGrayScale Then
                aHeaderFooter(2) = 3    ' b&w flag
                bPaletted = False       ' grayscale does not use palettes
            Else
                palSize = UBound(aPalette) + 1 ' add palette info to the header
                aHeaderFooter(2) = 1    ' paletted image type
                aHeaderFooter(1) = 1    ' palette exists
                CopyMemory aHeaderFooter(5), palSize \ 3, 2& ' how many palette entries
                aHeaderFooter(7) = 24   ' bits per entry
            End If
            aHeaderFooter(16) = 8       ' bit depth
        End If
        
    End If
    
    If bPaletted = False And bGrayScale = False Then ' didn't palettize; theStream not populated
    
        If allowPreMultipliedAlpha = True And bAlpha > AlphaNone Then ' just using raw data
        
            iparseOverlayHost_Byte theStream(), VarPtr(dSA), 1, cHost.scanWidth * cHost.Height, 0&, cHost.BitsPointer
            ' ^^ we can actually copy the dib data into theStream, but why?
            ' We can use it directly & don't need to take time & memory to make a copy
            aHeaderFooter(16) = 32      ' bit depth
            
        Else ' reduce to 24bpp or remove premultiplication
        
            FormatTrueColor theStream(), cHost, 0, 0, bAlpha
            If bAlpha Then         ' no 24bpp reduction performed
                aHeaderFooter(16) = 32  ' bit depth
            Else
                aHeaderFooter(16) = 24  ' 24bpp reduction applied
            End If
        End If
        aHeaderFooter(2) = 2            ' true color
        
    End If
    CopyMemory aHeaderFooter(12), cHost.Width, 2&
    CopyMemory aHeaderFooter(14), cHost.Height, 2&
  ' aHeaderFooter(17)=0
  ' ^^ leave this 0 on purpose; means bottom up image, left to right which is how DIB is stored
  ' other header items not applicable or will be filled in as routine continues
        
    If useCompression Then
        ' pass off to another function to compress the image
        imgDataLen = CompressImage(aCompressed(), theStream(), 0&, aHeaderFooter(16) \ 8, cHost.Width, cHost.Height)
        If imgDataLen > UBound(theStream) Then
            ' if compressed image larger than original, use original. Theoretically possible
            imgDataLen = UBound(theStream) + 1&
        Else
            ' transfer compressed data to theStream, but don't overwrite the
            ' original DIB if theStream is overlayed on it
            If Not dSA.pvData = 0& Then ' theStream is overlayed on original DIB
                ' memory hack, we will move the aCompressed pointer to theStream pointer
                ' thereby making theStream and aCompressed the same, then remove the aCompressed pointer. Scary.
                ' Other option is to remove DMA overlay on theStream, then copy aCompressed to theStream
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
    ' 4 bytes for extension area offset. See CreateTGAextensionBlock
    ' 4 bytes for developer area offset. Not used in this project
    ' 17 bytes for TRUEVISION-XFILE. << include period
    ' 1 nullchar terminating byte
    
    If bAlpha Then ' palette won't apply for alpha (not this version anyway)
        aHeaderFooter(17) = aHeaderFooter(17) Or 8  ' uses alpha channel
        ' when premultiplied alpha is used we will add the optional application block.
        ' These classes do not need, nor rely on, that optional data. These classes can correctly
        ' determine if alpha values are used and whether pixels are premultiplied or not. However,
        ' other apps may rely on the application block to properly display tga files. Therefore,
        ' we add the optional block -- to help ensure tga is compatible with other apps.
        If allowPreMultipliedAlpha Then
            appSize = CreateTGAextensionBlock(aAppBlock(), (bAlpha > AlphaNone), allowPreMultipliedAlpha)
            ReDim outStream(0 To imgDataLen + 43& + appSize) ' 43=header+footer-1
            appOffset = UBound(outStream) - appSize - 25&    ' get position where app block to be written
        Else
            ReDim outStream(0 To imgDataLen + 43&)  ' 43=header+footer-1
        End If
        CopyMemory outStream(0), aHeaderFooter(0), 18&      ' write the header
        ' create the footer & modify it
        aHeaderFooter() = StrConv(String$(8, 0) & "TRUEVISION-XFILE." & vbNullChar, vbFromUnicode)
        If Not appOffset = 0& Then
            CopyMemory aHeaderFooter(0), appOffset, 4&          ' write offset where app block can be found
            CopyMemory outStream(appOffset), aAppBlock(0), appSize ' write the app block
            Erase aAppBlock
        End If
        
    Else
    
        If bPaletted Then                                   ' include palette size in outStream
            ReDim outStream(0 To imgDataLen + 43& + palSize) ' 43=header+footer-1
            CopyMemory outStream(18), aPalette(0), palSize  ' copy palette
        Else
            ReDim outStream(0 To imgDataLen + 43&)          ' size without palette; 43=header+footer-1
        End If
        CopyMemory outStream(0), aHeaderFooter(0), 18&      ' write the header & create footer
        aHeaderFooter() = StrConv(String$(8, 0) & "TRUEVISION-XFILE." & vbNullChar, vbFromUnicode)
    End If
    ' header written in above lines
    CopyMemory outStream(18& + palSize), theStream(0), imgDataLen        ' write image data
    CopyMemory outStream(UBound(outStream) - 25&), aHeaderFooter(0), 26& ' write the footer
    
    SaveToStream = True
    
ExitRoutine:
If Err Then Err.Clear      ' I cannot foresee any errors other than out of memory
If Not dSA.pvData = 0& Then iparseOverlayHost_Byte theStream(), 0&, 0&, 0&, 0&, 0&

End Function

Private Function isTGA(inStream() As Byte, Offset As Long, Length As Long, isValid As Boolean) As Boolean

    ' Function attempts to rule out whether stream contains a TGA file
    ' There are some limitations when the image is paletted:
    ' -- will only support palette indexes of 8 bits, 1 byte
    '    This is not a TGA restriction; rather my own. Until I find or am provided
    '    with sample images other than 1 byte indexes; image is denied

    Dim tData(0 To 23) As Byte
    Dim expectedSize As Long
    Dim cX As Long, cY As Long
    Dim tVal As Long, pixBpp As Long
    
    isValid = False
    If Length > 23 Then ' check for v2 TGA footer first
        CopyMemory tData(0), inStream(Offset + Length - 22), 23&
        If InStr(1, StrConv(tData(), vbUnicode), "TRUEVISION-XFILE.", vbTextCompare) > 0 Then isTGA = True
        ' if the above is true,
        '   :: we will still do some validation before we return.
        '   :: Parser must return True even if image is invalid; this prevents other parsers
        '           from attempting to process the data, when we know 100% it is a tga
    End If
    
    If Length > 18 Then ' possible to have a one pixel B&W image = 19 bytes
        
        ' This will attempt to validate both version 1 & 2 tgas before we actually parse image data
    
        ' v1 has no magic number we can use. We will try to exclude & process if we cannot
    
        ' we will look for unusual values or known values that cannot exist. If any are found,
        ' our routines won't process the file, so other parsers can have a chance if applicable
        
        ' Text 1: Validate image type
        Select Case inStream(Offset + 2)
            Case 1, 2, 3, 9, 10, 11 ' supported
            Case Else ' 0=no image, 32 & 33 may be valid, but no documentation on those types available to me
                Exit Function
        End Select
    
        ' Test 2: Image size
        CopyMemory cX, inStream(Offset + 12), 2&
        CopyMemory cY, inStream(Offset + 14), 2&
        ' width/height are unsigned; any negative values would indicate absolutely huge image which
        ' we are not prepared to process or simply that this is not a tga file
        If cX < 1& Or cY < 1& Then Exit Function ' invalid tga or not a tga
        
        ' Test 3. Ignore unsupported pixel bit depths
        Select Case inStream(Offset + 16)
            Case 8: pixBpp = 1 ' supported
            Case 15, 16 ' supported if not paletted
                If (inStream(Offset + 2) And Not 8) = 1 Then Exit Function ' paletted
                pixBpp = 2
            Case 24, 32 ' supported if not paletted
                If (inStream(Offset + 2) And Not 8) = 1 Then Exit Function ' paletted
                pixBpp = (inStream(Offset + 16) \ 8)
            Case Else: Exit Function ' per documentation; other bit depths may be valid, but unsupported here
        End Select
        
        ' Test 4: paletted image validations
        If (inStream(Offset + 2) And Not 8) = 1 Then
            If Not inStream(Offset + 1) = 1 Then Exit Function  ' gotta have a palette if paletted image
        End If
        
        ' Test 5. Validate palette attributes. These can exist in image even if image is not paletted
        If inStream(Offset + 1) = 1 Then
            CopyMemory expectedSize, inStream(Offset + 5), 2&   ' are nr entries valid?
            If expectedSize < 1& Then Exit Function
            CopyMemory tVal, inStream(Offset + 3), 2&           ' is palette offset valid?
            If tVal < 0 Then Exit Function
            expectedSize = expectedSize + tVal                  ' how many indexes are expected
            
            Select Case inStream(Offset + 7)                    ' ignore unsupported palette bit depths
                Case 8, 24, 32 ' supported
                    expectedSize = expectedSize * (inStream(Offset + 7) \ 8)
                Case 15, 16 ' supported
                    expectedSize = expectedSize * 2&
                Case Else: Exit Function ' per documentation; nothing else is valid
            End Select
        End If
        
        ' Test 6: Minimal file size expected
        If (inStream(Offset + 2) And 8) = 0& Then ' uncompressed, else compressed
            expectedSize = expectedSize + (pixBpp * cX * cY)
        End If
        If isTGA = True Then expectedSize = expectedSize + 26 ' has footer
        If expectedSize + inStream(Offset) + 18& <= Length Then
            ' file may be much larger, but should be at least the size of expectedSize
            isValid = True
            isTGA = True
        End If
        
    End If

End Function

Private Sub ProcessTrueColor(inStream() As Byte, Offset As Long, streamLength As Long, cHost As c32bppDIB)

    ' Processes compressed/uncompressed, TrueColor & GrayScale TGA files
    ' Note: the isTGA function already validated key portions of the header; so this routine does no
    ' further validation, other than to ensure compressed images uncompress correctly

    Dim aUncompressed() As Byte
    Dim X As Long, Y As Long
    Dim dBytes() As Byte, dSA As SafeArray
    Dim lScanWidth As Long, rowOffset As Long
    Dim bAlpha As AlphaTypeEnum, lType As Long
    Dim bPP As Long, int16 As Integer
    Dim maxWidth As Long
    
    Dim dstStepX As Long, dstStartX As Long, dstX As Long
    Dim dstStepY As Long, dstStartY As Long, dstStopY As Long

    Select Case inStream(16 + Offset) ' pixel bit/byte depth
        Case 15: bPP = 2
            Create16bitLUT
        Case 16: bPP = 2
            Create16bitLUT
            ' 16bpp can have alpha too; either fully transparent or fully opaque; 1bit value
            ' See if flag is set & verify usage from version 2 tga data if available
            If (inStream(17 + Offset) And 31) > 0& Then
                bAlpha = AllowAlphaChannel(inStream(), Offset, streamLength)
            End If
        Case 8, 24, 32: bPP = inStream(16 + Offset) \ 8
    End Select
    
    If inStream(Offset + 1) = 1 Then   ' color map exists before true color data (unusual but not prohibited)
        CopyMemory int16, inStream(Offset + 3), 2&          ' number of palette entries to skip
        If int16 < 0& Then Exit Sub                         ' invalid entry, invlaid TGA
        CopyMemory rowOffset, inStream(Offset + 5), 2&      ' number of color map entries
        Select Case inStream(Offset + 7) ' bit depth per entry
        Case 8, 24, 32
            rowOffset = (rowOffset + int16) * (inStream(Offset + 7) \ 8)
        Case 15, 16
            rowOffset = (rowOffset + int16) * 2
        End Select
        
    End If
    
    CopyMemory X, inStream(12 + Offset), 2&            ' get image width
    CopyMemory Y, inStream(14 + Offset), 2&            ' and image height
    
    lScanWidth = bPP * X                               ' source/raw scan width
    rowOffset = 18& + inStream(Offset) + rowOffset ' identify where first image data byte exists
    
    cHost.InitializeDIB X, Y                          ' create blank dib to write to
    
    If (inStream(Offset + 17) \ 32 And 1) = 1 Then  ' top down image
        dstStepY = -1&: dstStartY = Y + dstStepY: dstStopY = 0&
    Else                                            ' bottom up image (like our dib)
        dstStepY = 1&: dstStartY = 0&: dstStopY = Y - dstStepY
    End If
    If (inStream(Offset + 17) \ 64 And 1) = 1 Then  ' right to left
        dstStepX = -4&: dstStartX = X * 4& - 5&
    Else
        dstStepX = 4&: dstStartX = 0&               ' left to right image
    End If
    
    iparseOverlayHost_Byte dBytes(), VarPtr(dSA), 2, Y, cHost.scanWidth, cHost.BitsPointer
    
    If (inStream(Offset + 2) And 8) = 8 Then        ' compressed
        ReDim aUncompressed(0 To X * 4& - 1)       ' 32bpp array
        For Y = dstStartY To dstStopY Step dstStepY
            
            If UncompressTrueColor(inStream(), rowOffset, aUncompressed(), bPP, cHost.Width, bAlpha) = False Then
                cHost.DestroyDIB ' corrupt data or not a tga file
                iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0
                Exit Sub
            End If
            
            If dstStepX = bPP Then  ' left to right source & 32bpp
                CopyMemory dBytes(dstX, Y), aUncompressed(0), lScanWidth
            Else                    ' right to left and/or 24bpp source
                dstX = dstStartX
                For X = 0 To UBound(aUncompressed) Step 4
                    CopyMemory dBytes(dstX, Y), aUncompressed(X), 4&
                    dstX = dstX + dstStepX
                Next
            End If
        Next
        Erase aUncompressed()
        
    Else
    
        maxWidth = Offset + streamLength ' prevent writing past allocated dib memory
        ' If isTGA function is working correctly, any uncompressed image that would not
        ' contain enough bytes to fill up the image would be rejected; therefore we do
        ' not slow down the loops checking for potential scanwidth violations.
        
        Select Case bPP
        Case 3, 4 ' 32-24 bits per pixel / 4-3 bytes per pixel
            ' simple & straightforward copy
            If dstStepX = bPP Then      ' faster updating when left to right & 32bpp
                If dstStartY = 0& Then  ' even faster if image is bottom up
                    CopyMemory dBytes(0, 0), inStream(rowOffset), lScanWidth * (dstStopY + 1)
                Else
                    For Y = dstStartY To dstStopY Step dstStepY
                        CopyMemory dBytes(dstX, Y), inStream(rowOffset), lScanWidth
                        rowOffset = rowOffset + lScanWidth
                    Next
                End If
            Else            ' either top-down or right-left source or 24bpp
                For Y = dstStartY To dstStopY Step dstStepY
                    dstX = dstStartX
                    For X = 0 To lScanWidth - 1 Step bPP
                        CopyMemory dBytes(dstX, Y), inStream(rowOffset + X), bPP
                        dstX = dstX + dstStepX
                    Next
                    rowOffset = rowOffset + lScanWidth
                Next
            End If
            
        Case 2 ' 15 or 16 bits per pixel / 2 bytes per pixel
            
            ' some bit shifting required
            For Y = dstStartY To dstStopY Step dstStepY
                dstX = dstStartX
                For X = 0 To lScanWidth - 1 Step bPP
                    ' use a 5,5,5 pattern. Build 15 bits into a 16 bit/2 byte Integer
                    int16 = (inStream(rowOffset + X + 1) And &H7F) * &H100 Or inStream(rowOffset + X)
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
            ' copy 1 byte at a time, skip alpha (overwrite it to 255 at end)
            For Y = dstStartY To dstStopY Step dstStepY
                dstX = dstStartX
                For X = 0 To lScanWidth - 1
                    dBytes(dstX, Y) = inStream(rowOffset + X)
                    dBytes(dstX + 1, Y) = dBytes(dstX, Y)
                    dBytes(dstX + 2, Y) = dBytes(dstX, Y)
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + lScanWidth
            Next
        End Select
    
    End If
    
    iparseValidateAlphaChannel dBytes, True, bAlpha, lType
    iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0

    cHost.Alpha = bAlpha
    cHost.ImageType = imgTGA

End Sub

Private Sub ProcessColorMap(inStream() As Byte, Offset As Long, streamLength As Long, cHost As c32bppDIB)

    ' Processes compressed/uncompressed, paletted TGA images
    ' Note: the isTGA function already validated key portions of the header; so this routine does no
    ' further validation, other than to ensure compressed images uncompress correctly

    Dim pal() As Byte, aUncompressed() As Byte
    Dim X As Long, Y As Long
    Dim lScanWidth As Long, rowOffset As Long
    Dim bAlpha As AlphaTypeEnum, bPalAlpha As Boolean, lType As Long
    Dim bPP As Long, palBpp As Long
    Dim mapSize As Long, int16 As Integer
    Dim dBytes() As Byte, dSA As SafeArray
    
    Dim dstStepX As Long, dstStartX As Long, dstX As Long
    Dim dstStepY As Long, dstStartY As Long, dstStopY As Long
    Dim maxWidth As Long
    
    'Select Case inStream(16 + Offset) ' source pixel bit/byte depth
    '    Case 8: bPP = 1
    'End Select
    ' currently, only 8bit, 1byte palette indexes are supported
    bPP = 1
    
    CopyMemory mapSize, inStream(Offset + 5), 2&       ' number of palette entries
    Select Case inStream(Offset + 7)                   ' bit depth per entry
    Case 8: palBpp = 1
    Case 15: palBpp = 2
        Create16bitLUT
    Case 16: palBpp = 2
        Create16bitLUT
        ' It is possible for the palette to have alpha and the pixels to also have alpha
        ' If pixels have alpha, then their alpha overrides any palette alpha
        bPalAlpha = AllowAlphaChannel(inStream(), Offset, streamLength)
    Case 24: palBpp = 3
    Case 32: palBpp = 4
        bPalAlpha = AllowAlphaChannel(inStream(), Offset, streamLength)
    End Select
    
    CopyMemory int16, inStream(Offset + 3), 2& ' number of palette entries to skip
    rowOffset = (mapSize + int16) * palBpp + 18& + inStream(Offset) ' this is where the palette indexes begin
    
    CopyMemory X, inStream(12 + Offset), 2&            ' get image width
    CopyMemory Y, inStream(14 + Offset), 2&            ' and image height
    
    If mapSize < 256 Then
        ReDim pal(0 To 1023)    ' oversize palette to prevent any errors if palette index exceeds palette
    Else
        ReDim pal(0 To mapSize * 4 - 1)           ' use actual size of palette
    End If
    cHost.InitializeDIB X, Y                      ' create blank dib to write to
    
    ' determine where in the source palette to start reading entries from
    X = 0&: Y = rowOffset - (mapSize * palBpp)
    For Y = Y To rowOffset - 1 Step palBpp
        Select Case palBpp
        Case 1:
            pal(X) = inStream(Y)
            pal(X + 1) = inStream(Y)
            pal(X + 2) = inStream(Y)
        Case 2
            int16 = (inStream(Y + 1) And &H7F) * &H100 Or inStream(Y)
            ' shift the 5,5,5 and ref a 5 bit/32 entry palette
            pal(X) = pal32LUT(int16 And &H1F)
            pal(X + 1) = pal32LUT((int16 \ &H20) And &H1F)
            pal(X + 2) = pal32LUT((int16 \ &H400) And &H1F)
            If bPalAlpha Then
                If (inStream(Y + 1) And &H80) = 0& Then pal(X + 3) = &HFF
            End If
        Case 3, 4
            CopyMemory pal(X), inStream(Y), 3
            If bPalAlpha Then pal(X + 3) = inStream(Y + palBpp - 1) ' 32bpp palettes
        End Select
        X = X + 4
    Next
    
    If (inStream(Offset + 17) \ 32 And 1) = 1 Then  ' top down image
        dstStepY = -1&: dstStartY = cHost.Height - 1&: dstStopY = 0&
    Else                                            ' bottom up (like our dib)
        dstStepY = 1&: dstStartY = 0&: dstStopY = cHost.Height - 1&
    End If
    If (inStream(Offset + 17) \ 64 And 1) = 1 Then  ' right to left
        dstStepX = -4&: dstStartX = cHost.Width * 4& - 5&
    Else
        dstStepX = 4&: dstStartX = 0&               ' left to right
    End If
    
    iparseOverlayHost_Byte dBytes(), VarPtr(dSA), 2, cHost.Height, cHost.scanWidth, cHost.BitsPointer
    
    If (inStream(Offset + 2) And 8) = 8 Then        ' compressed
    
        lScanWidth = cHost.scanWidth                ' dib scan width
        ReDim aUncompressed(0 To lScanWidth - 1) ' 32bpp format
        For Y = dstStartY To dstStopY Step dstStepY
            
            If UncompressColorMap(inStream(), rowOffset, aUncompressed(), bPP, cHost.Width, pal()) = False Then
                cHost.DestroyDIB    ' corrupt data or not a tga
                iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0
                Exit Sub
            End If
            If dstStepX = 4& Then  ' left to right image, copy entire scan line in one call
                CopyMemory dBytes(dstX, Y), aUncompressed(0), lScanWidth
            Else                    ' right to left
                dstX = dstStartX
                For X = 0 To UBound(aUncompressed) Step 4
                    CopyMemory dBytes(dstX, Y), aUncompressed(X), 4
                    dstX = dstX + dstStepX
                Next
            End If
        Next
        Erase aUncompressed()
    
    Else
    
        maxWidth = Offset + streamLength ' prevent writing past allocated dib memory
        lScanWidth = bPP * cHost.Width  ' source/raw scan width
    
        ' 8 bits per pixel / 1 byte per pixel
        ' copy 1 byte at a time, skip alpha (overwrite it to 255 at end)
        For Y = dstStartY To dstStopY Step dstStepY
            If lScanWidth + rowOffset > maxWidth Then
                ' safety check that should never fire if isTGA function is working correctly
                cHost.DestroyDIB ' corrupt data or not a tga file
                iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0
                Exit Sub
            End If
            dstX = dstStartX
            For X = 0 To lScanWidth - 1
                CopyMemory dBytes(dstX, Y), pal(inStream(rowOffset + X) * 4&), 4&
                dstX = dstX + dstStepX
            Next
            rowOffset = rowOffset + lScanWidth
        Next
        
    End If
    
    iparseValidateAlphaChannel dBytes, True, bAlpha, lType
    iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0

    cHost.Alpha = bAlpha
    cHost.ImageType = imgTGA

End Sub

Private Function UncompressTrueColor(inStream() As Byte, Offset As Long, bScan() As Byte, bPP As Long, _
                                    Width As Long, AlphaType As AlphaTypeEnum) As Boolean

    ' RLE uncompression of TGA images: TrueColor only

    Dim pixCount As Long, pixLoc As Long
    Dim int16 As Integer, bReturn As Boolean
    Dim tileStop As Long, tileCx As Long
    Dim X As Long, lScanWidth As Long
    
    bReturn = True
    lScanWidth = Width * 4&
    Do Until pixCount = Width
        If (inStream(Offset) And 128) = 0 Then  ' this set of pixels not RLE compressed
            ' when not compressed, the control byte says how many non-compressed pixels to process
            
            If bPP * inStream(Offset) + pixLoc > lScanWidth Then
                bReturn = False ' corrupt data or not tga; will exceed scanwidth
                Exit Do
            End If
            Offset = Offset + 1                 ' move to next byte
            
            Select Case bPP
            Case 4                  ' 32bpp; straight copy
                X = inStream(Offset - 1) + 1
                CopyMemory bScan(pixLoc), inStream(Offset), X * bPP
                pixLoc = pixLoc + X * bPP
                Offset = Offset + X * bPP
            Case 3                  ' 24 bpp
                For X = 0 To inStream(Offset - 1)
                    CopyMemory bScan(pixLoc), inStream(Offset), bPP ' copy 3 bytes
                    pixLoc = pixLoc + 4
                    Offset = Offset + bPP       ' move 3 bytes
                Next                            ' already positioned on next RLE control byte
            Case 2                  ' 16 bpp
                For X = 0 To (inStream(Offset - 1)) ' create 15bit integer & shift into palette
                    int16 = (inStream(Offset + 1) And &H7F) * &H100 Or inStream(Offset)
                    bScan(pixLoc) = pal32LUT((int16 And &H1F))
                    bScan(pixLoc + 1) = pal32LUT(((int16 \ &H20) And &H1F))
                    bScan(pixLoc + 2) = pal32LUT((int16 \ &H400) And &H1F)
                    If AlphaType > AlphaNone Then
                        If (inStream(Offset + 1) And &H80) = 0 Then bScan(pixLoc + 3) = &HFF
                    End If
                    pixLoc = pixLoc + 4
                    Offset = Offset + bPP       ' move 2 bytes
                Next                            ' already positioned on next RLE control byte
            Case 1                  ' 8 bpp/grayscale
                For X = 0 To (inStream(Offset - 1))
                    bScan(pixLoc) = inStream(Offset)    ' add grayscale values
                    bScan(pixLoc + 1) = bScan(pixLoc)
                    bScan(pixLoc + 2) = bScan(pixLoc)
                    pixLoc = pixLoc + 4
                    Offset = Offset + bPP       ' move to next byte
                Next                            ' already positioned on next RLE control byte
            End Select
            
            pixCount = pixCount + X
            
        Else                                    ' RLE compressed. TGA does RLE runs on pixels; not bits/bytes
            
            Select Case bPP
            Case 3, 4                           ' 24/32bpp
                CopyMemory bScan(pixLoc), inStream(Offset + 1), bPP ' get color from next bytes
            Case 2
                int16 = (inStream(Offset + 2) And &H7F) * &H100 Or inStream(Offset + 1)
                ' shift the 5,5,5 and ref a 5 bit/32 entry palette
                bScan(pixLoc) = pal32LUT((int16 And &H1F))
                bScan(pixLoc + 1) = pal32LUT(((int16 \ &H20) And &H1F))
                bScan(pixLoc + 2) = pal32LUT((int16 \ &H400) And &H1F)
                If AlphaType > AlphaNone Then
                    If (inStream(Offset + 1) And &H80) = 0 Then bScan(pixLoc + 3) = &HFF
                End If
            Case 1
                bScan(pixLoc) = inStream(Offset + 1)
                bScan(pixLoc + 1) = bScan(pixLoc)
                bScan(pixLoc + 2) = bScan(pixLoc)
            End Select
            
            ' here we use tiling techniques to quickly fill a scan line with repeated pixels
            tileCx = 4                              ' how many pixels copied so far
            tileStop = ((inStream(Offset) And Not 128) + 1) * tileCx ' how many pixels will be duplicated
            If tileStop + pixLoc > lScanWidth Then
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
            pixCount = pixCount + tileStop \ 4 ' number of pixels processed during decompression
        End If
        
    Loop
    UncompressTrueColor = bReturn

End Function

Private Function UncompressColorMap(inStream() As Byte, Offset As Long, bScan() As Byte, bPP As Long, _
                                    Width As Long, Palette() As Byte) As Boolean

    ' RLE uncompression of TGA images: Paletted TGA only
    
    Dim pixCount As Long, pixLoc As Long
    Dim int16 As Integer, bReturn As Boolean
    Dim tileStop As Long, tileCx As Long
    Dim X As Long, lScanWidth As Long
    Dim maxSourceCount As Long
    
    bReturn = True
    lScanWidth = Width * 4&
    maxSourceCount = UBound(bScan)
    Do Until pixCount = Width
        If (inStream(Offset) And 128) = 0 Then
            ' not RLE encoded
            If bPP * inStream(Offset) + pixLoc > lScanWidth Then
                bReturn = False ' corrupt data or not tga; will exceed scanwidth
                Exit Do
            ElseIf maxSourceCount < pixLoc + inStream(Offset) * 4& Then
                bReturn = False ' corrupt data or not tga; will exceed destination size
                Exit Do
            End If
            
            Offset = Offset + 1
            For X = 0 To inStream(Offset - 1)
                CopyMemory bScan(pixLoc), Palette(inStream(Offset) * 4), 4
                pixLoc = pixLoc + 4
                Offset = Offset + bPP
            Next
            pixCount = pixCount + X
            
        Else
            
            CopyMemory bScan(pixLoc), Palette(inStream(Offset + 1) * 4), 4
            
            tileCx = 4&                              ' how many pixels copied so far
            tileStop = ((inStream(Offset) And Not 128) + 1) * tileCx ' how many bytes will be duplicated
            If tileStop + pixLoc > lScanWidth Then
                bReturn = False     ' corrupt data or not a tga; will exceed scanwidth
                Exit Do
            ElseIf maxSourceCount < tileStop + pixLoc Then
                bReturn = False     ' corrupt data or not tga; will exceed destination size
                Exit Do
            End If
            
            Do Until tileCx + tileCx > tileStop     ' copy another group of pixels
                CopyMemory bScan(pixLoc + tileCx), bScan(pixLoc), tileCx
                tileCx = tileCx + tileCx            ' increment group by x2
            Loop
            tileCx = tileStop - tileCx              ' see if any partial tiling is needed
            If tileCx Then                          ' and copy required remaining pixels
                CopyMemory bScan(pixLoc + tileStop - tileCx), bScan(pixLoc), tileCx
            End If
            pixLoc = pixLoc + tileStop      ' move bScan() pointer to next byte to be filled
            Offset = Offset + bPP + 1       ' move to next RLE control byte
            pixCount = pixCount + tileStop \ 4 ' number of pixels processed during decompression
        
        End If
    
    Loop
    UncompressColorMap = bReturn
    
End Function

Private Sub Create16bitLUT()
    ' Initialize a 32 entry palette
    Dim X As Byte
    If pal32LUT(31) = 0 Then ' not yet created
        For X = 0 To 31
            pal32LUT(X) = X * 8 + (X Mod 8)
        Next
    End If
    
End Sub

Private Sub FormatTrueColor(outStream() As Byte, cHost As c32bppDIB, _
                           HeaderBytes As Long, FooterBytes As Long, AlphaType As AlphaTypeEnum)

    ' Function converts to 24bpp or removes premultiplied pixels
    
    Dim X As Long, Y As Long
    Dim tOffset As Long
    Dim scanWidth As Long, Color As Byte

    Dim dBytes() As Byte, dSA As SafeArray
    
    iparseOverlayHost_Byte dBytes(), VarPtr(dSA), 2&, cHost.Height, cHost.scanWidth, cHost.BitsPointer
    
    If AlphaType Then
        scanWidth = cHost.Width * 4&
        ReDim outStream(0& To scanWidth * cHost.Height + HeaderBytes + FooterBytes - 1&)
        tOffset = HeaderBytes
        For Y = 0& To cHost.Height - 1&
            For X = 0& To scanWidth - 1& Step 4&
                ' remove pre-multiplication
                Select Case dBytes(X + 3&, Y)
                Case 255
                    CopyMemory outStream(tOffset), dBytes(X, Y), 4&
                Case 0 ' do nothing
                Case Else
                    Color = dBytes(X + 3&, Y)
                    outStream(tOffset) = (255& * dBytes(X, Y) \ Color)
                    outStream(tOffset + 1&) = (255& * dBytes(X + 1&, Y) \ Color)
                    outStream(tOffset + 2&) = (255& * dBytes(X + 2, Y) \ Color)
                    outStream(tOffset + 3&) = Color
                End Select
                tOffset = tOffset + 4&
            Next
        Next
    Else  ' convert to 24bpp
        scanWidth = cHost.Width * 4&
        ReDim outStream(0& To cHost.Width * cHost.Height * 3& + HeaderBytes + FooterBytes - 1&)
        tOffset = HeaderBytes
        For Y = 0& To cHost.Height - 1&
            For X = 0& To scanWidth - 1& Step 4&
                CopyMemory outStream(tOffset), dBytes(X, Y), 3&
                tOffset = tOffset + 3&
            Next
        Next
    End If
    
    iparseOverlayHost_Byte dBytes(), 0&, 0&, 0&, 0&, 0&
    
End Sub

Private Function PalettizeImage(Palette() As Byte, isGrayScale As Boolean, cHost As c32bppDIB, outStream() As Byte, _
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
    Dim scanWidth As Long, palCount As Long
    
    Dim dBytes() As Byte, dSA As SafeArray
    
    iparseOverlayHost_Byte dBytes(), VarPtr(dSA), 2, cHost.Height, cHost.scanWidth, cHost.BitsPointer
    
    ' count unique colors (maximum of 256 if we are to palettize)
    ' If any alpha values exist, we abort. Otherwise we continue until we know for sure it is not alpha
    
    ReDim tPalette(1 To 256)    ' palette
    isGrayScale = True          ' default until proven otherwise
    scanWidth = cHost.scanWidth - 1& ' cache vs recalculating each scanline
    
    ' force impossible first match for loop below
    prevColor = ((dBytes(0, 0) + 1) Mod 255) Or _
                ((dBytes(1, 0) + 1) Mod 255) * &H100 Or _
                ((dBytes(2, 0) + 1) Mod 255) * &H10000
    
    For Y = 0& To cHost.Height - 1&
        For X = 0& To scanWidth Step 4&
        
            CopyMemory Color, dBytes(X, Y), 4&  ' get next pixel
            If Not Color = prevColor Then       ' is this same as previous color?
                ' use binary search routine to locate it if it already exists
                Index = iparseFindColor(tPalette, Color, palCount, newColor)
                If newColor = True Then         ' color not in palette yet
                    palCount = palCount + 1&    ' increment palette count
                    If palCount = 257& Then     ' maxed palette out?
                        Y = cHost.Height        ' aborts outer loop
                        Exit For
                    Else
                        If Index < palCount Then ' keep palette entries sorted for binary search routine
                            CopyMemory tPalette(Index + 1&), tPalette(Index), (palCount - Index) * 4&
                        End If
                        tPalette(Index) = Color ' add new color to the palette
                        If isGrayScale Then     ' validate grayscale potential
                            If dBytes(X, Y) = dBytes(X + 1&, Y) Then
                                If Not dBytes(X, Y) = dBytes(X + 2&, Y) Then isGrayScale = False
                            Else
                                isGrayScale = False
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
        ReDim outStream(0 To cHost.Width * cHost.Height + HeaderBytes + FooterBytes - 1)
        lCol = HeaderBytes      ' where to begin writing palette indexes
        
        If isGrayScale Then     ' no palette required, the grayscale values are the indexes
            For Y = 0 To cHost.Height - 1&
                For X = 0 To cHost.scanWidth - 1& Step 4&
                    outStream(lCol) = dBytes(X, Y)
                    lCol = lCol + 1&
                Next
            Next
            Erase Palette()
        
        Else
            
            ' force impossible first match for loop below
            prevColor = ((dBytes(0, 0) + 1) Mod 255) Or _
                        ((dBytes(1, 0) + 1) Mod 255) * &H100 Or _
                        ((dBytes(2, 0) + 1) Mod 255) * &H10000
            
            For Y = 0& To cHost.Height - 1&
                For X = 0& To cHost.scanWidth - 1& Step 4
                    CopyMemory Color, dBytes(X, Y), 4&  ' get next color
                    If Not Color = prevColor Then       ' use binary search to find it if needed
                        ' when found, convert 1 base to 0 base indexing
                        Index = iparseFindColor(tPalette(), Color, palCount, False) - 1&
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
        
        PalettizeImage = True
        
    End If
    iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0

End Function

Private Function CompressImage(arrRLE() As Byte, rawStream() As Byte, Offset As Long, bPP As Long, _
                                imgWidth As Long, imgHeight As Long) As Long

    ' TGA compression routine. Feel free to recommend optimizations
    ' FYI: RLE = Run Lengh Encoding
    ' Simple RLE rules:
    '   Runs will not cross scanlines; restricted to one scanline only
    '   Control Byte high bit, if set, means compressed run, else raw/uncompressed run
    '   -- Max run length is 128
    '   -- (Byte And Not 128)+1 = length of run
    '   -- One control byte precedes every RLE run
    '   If compressed, one pixel is added to RLE, these are the bytes to be repeated
    '   If uncompressed, next n pixels are added to RLE, uncompressed bytes
    
    ' Parameters:
    '   arrRLE() an array to hold RLE scanline on return; return size is not important
    '   rawStream() an array of raw pixel data, byte aligned, on bPP byte boundaries
    '   Offset is where in raw data the first pixel appears; normally LBound(rawStream)
    '   bPP is how many bytes per pixel raw data is formatted to: 1, 2, 3 or 4 only
    '   imgWidth is the width of the image
    '   imgHeight is the height of the image
    
    ' Function returns length of compressed image which may be smaller than the array size
    ' If function returns zero, the returned array should be considered null
    
    ' Testing UBound() below slows down routines just a hair but better safe than crashing

    Dim X As Long, Y As Long
    Dim Mode As Long, ModeChange As Long
    Dim prevPixel As Long, curPixel As Long
    Dim locRLE As Long, locSource As Long
    Dim curCount As Long, maxBytes As Long
    Dim scanWidth As Long, lastScanByte As Long
    
    Const MaxRunLength As Long = 128&
    Const modeUndefined = 0&
    Const modeCompress = 1&
    Const modeRaw = 2&
    
    ReDim arrRLE(0 To UBound(rawStream))                ' initial size of RLE encoded bytes
    scanWidth = imgWidth * bPP                          ' raw scan line width
    
    For Y = 0 To imgHeight - 1
        Mode = modeUndefined                            ' reset each scan line
        ModeChange = modeUndefined                      ' reset each scan line
        locSource = Offset + scanWidth * Y              ' 1st pixel on current scan line
        lastScanByte = locSource + scanWidth            ' 1st pixel on next scan line
        maxBytes = MaxRunLength * bPP + locSource       ' which byte constitues max RLE position
        '^^ note: we calc this vs using a counter to see if we are exceeding MaxRunLength,
        '   we only have to recalc when mode changes vs updating counter every pixel
        CopyMemory prevPixel, rawStream(locSource), bPP ' get initial pixel for comparison
        
        For X = locSource + bPP To lastScanByte - 1& Step bPP
        
            CopyMemory curPixel, rawStream(X), bPP      ' get next pixel to compare
            
            If curPixel = prevPixel Then
                If Mode = modeRaw Then                  ' changing from uncompressed to compressed
                    ModeChange = modeCompress
                ElseIf Mode = modeUndefined Then
                    Mode = modeCompress                 ' starting new compressed run
                End If
            
            Else
                If Mode = modeCompress Then             ' changing from compressed to uncompressed
                    ModeChange = modeRaw
                ElseIf Mode = modeUndefined Then
                    Mode = modeRaw                      ' starting new uncompressed run
                End If
                prevPixel = curPixel                    ' set new base pixel to compare against
            End If
            
            ' check for change in compression mode
            If ModeChange Then
                
                If ModeChange = modeCompress Then       ' was uncompressed; now can compress
                    curCount = (X - locSource - bPP) \ bPP ' nr of uncompressed pixels; don't include last two pixels
                    If curCount Then                    ' can be zero if 1st two pixels on new run are same
                        arrRLE(locRLE) = curCount - 1&  ' nr of compressed pixels
                        curCount = curCount * bPP
                        ' ensure do not write to unitialized memory
                        If locRLE + 1 + curCount >= UBound(arrRLE) Then ReDim Preserve arrRLE(0 To locRLE + curCount + scanWidth)
                        CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), curCount
                        locSource = X - bPP             ' where new RLE run starts
                        locRLE = locRLE + 1& + curCount ' where next RLE run is written
                    End If
                
                Else                                    ' was compressing; can't continue
                    ' ensure do not write to unitialized memory
                    If locRLE + bPP + 1 > UBound(arrRLE) Then ReDim Preserve arrRLE(0 To locRLE + scanWidth)
                    curCount = (X - locSource) \ bPP    ' nr of compressed pixels
                    arrRLE(locRLE) = MaxRunLength Or (curCount - 1&)
                    CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), bPP
                    locSource = X                       ' where new RLE run starts
                    locRLE = locRLE + 1& + bPP          ' where next RLE run is written
                End If
                
                maxBytes = MaxRunLength * bPP + locSource   ' where max RLE length would end
                Mode = ModeChange                       ' current compression mode
                ModeChange = modeUndefined              ' reset change flag
                
            ElseIf X = maxBytes Then                    ' check for max RLE; TGA only supports 128 maximum
            
                curCount = MaxRunLength * bPP           ' RLE byte length
                ' ensure do not write to unitialized memory
                If locRLE + curCount + 1 > UBound(arrRLE) Then ReDim Preserve arrRLE(0 To locRLE + curCount + scanWidth)
                If Mode = modeCompress Then             ' was compressing; add to RLE
                    arrRLE(locRLE) = 255                ' (128-1) OR 128 = 255
                    CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), bPP
                    locRLE = locRLE + 1& + bPP          ' location for next RLE data
                    
                Else                                    ' on an uncompressed RLE run
                
                    arrRLE(locRLE) = MaxRunLength - 1&
                    CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), curCount
                    locRLE = locRLE + curCount + 1&     ' location for next RLE data
                End If
                
                If X + bPP = lastScanByte Then
                    ' we are on the last pixel; set its compression mode to uncompressed so end of loop checks adds it
                    Mode = modeRaw
                Else
                    maxBytes = curCount + X             ' calc next max RLE location
                    Mode = modeUndefined                ' reset compression mode
                End If
                locSource = X                           ' where next RLE run starts
            End If
        Next
        
        If Not Mode = modeUndefined Then                ' finish current RLE run. Should always be needed
            curCount = (X - locSource) \ bPP            ' number of pixels in this scan line run
            
            If Mode = modeCompress Then                 ' currently compressing
                ' ensure do not write to unitialized memory
                If locRLE + bPP + 1 > UBound(arrRLE) Then ReDim Preserve arrRLE(0 To locRLE + bPP + 1)
                arrRLE(locRLE) = (curCount - 1&) Or MaxRunLength ' add nr of compressed pixels
                CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), bPP
                locRLE = locRLE + 1& + bPP              ' increment nr of RLE bytes written
                
            Else                                        ' currently uncompressed RLE
                arrRLE(locRLE) = curCount - 1&          ' add nr of uncompressed pixels
                curCount = curCount * bPP
                ' ensure do not write to unitialized memory
                If locRLE + curCount + 1 > UBound(arrRLE) Then ReDim Preserve arrRLE(0 To locRLE + curCount + 1)
                CopyMemory arrRLE(locRLE + 1&), rawStream(locSource), curCount
                locRLE = locRLE + 1& + curCount     ' increment nr of RLE bytes written
                
            End If
        End If
        
    Next
    
    CompressImage = locRLE  ' return RLE length
    
End Function

Private Function CreateTGAextensionBlock(outStream() As Byte, isAlpha As AlphaTypeEnum, preMultipliedAlhpa As Boolean) As Long

    
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
    bBlockData = StrConv("LaVolpe c32bppDIB Suite", vbFromUnicode)
    CopyMemory outStream(426), bBlockData(0), UBound(bBlockData) + 1
    outStream(467) = 3: outStream(469) = 1
    
    If preMultipliedAlhpa = True Then   ' add info about alpha blending used
        outStream(494) = 4
    ElseIf isAlpha > AlphaNone Then ' else zero for no alpha
        outStream(494) = 3
    End If
    
    CopyMemory outStream(470), -1&, 4&  ' add a default bkg color if other apps don't support alpha (White, fully opaque)
    
    iVal = Month(Date): CopyMemory outStream(367), iVal, 2& ' add current system date/time
    iVal = Day(Date): CopyMemory outStream(367 + 2), iVal, 2&
    iVal = Year(Date): CopyMemory outStream(367 + 4), iVal, 2&
    iVal = Hour(Now()): CopyMemory outStream(367 + 6), iVal, 2&
    iVal = Minute(Now()): CopyMemory outStream(367 + 8), iVal, 2&
    iVal = Second(Now()): CopyMemory outStream(367 + 10), iVal, 2&
    
    CreateTGAextensionBlock = v2BlockLength
    
End Function

Private Function AllowAlphaChannel(theStream() As Byte, Offset As Long, Length As Long) As Boolean

    ' Helper Function. Called when 16bpp images/palettes contain an alpha bit and for 32bpp palettes.
    
    ' This function will check if version 2 data exists and if so whether or not v2 says
    ' to ignore the alpha channel. Only if v2 exists and ignores the alpha will this
    ' function return False. 32bpp true color images are handled via modParsers.iparseValidateAlphaChannel
    
    Dim tData(0 To 29) As Byte
    Dim tVal As Long, appBlockLoc As Long
    Dim bReturn As Boolean
    
    Const v2BlockLength As Long = 495 ' v2 application/extension block length
    
    bReturn = True      ' default - don't override & allow any alpha
    If Length > 43& Then ' check for v2 TGA footer first
        CopyMemory tData(0), theStream(Offset + Length - 29&), 30&
        tVal = InStr(1, StrConv(tData(), vbUnicode), "TRUEVISION-XFILE.", vbTextCompare)
        If tVal > 8& Then
            ' ok, does it have an application/extension block? 8 bytes before TRUEVISION
            CopyMemory appBlockLoc, tData(tVal - 9&), 4&
            ' ensure it is not corrupt; must be a valid pointer & must be within the stream/file
            If appBlockLoc > 0& And appBlockLoc + Offset + v2BlockLength < Length Then
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
    AllowAlphaChannel = bReturn

End Function
