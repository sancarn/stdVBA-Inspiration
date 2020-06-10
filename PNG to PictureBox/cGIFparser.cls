VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGIFparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Convert GIF to Bitmap and Bitmap to GIF ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.

' used to extract data from a converted GIF
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function GetNearestPaletteIndex Lib "gdi32.dll" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePalette Lib "gdi32.dll" (ByRef lpLogPalette As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

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
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type tNode                              'Octree node struct
    bIsLeaf As Boolean                          'Leaf flag
    bAddedReduce As Boolean                     'Linked list flag
    vRed As Long                                'Red Value
    vGreen As Long                              'Green Value
    vBlue As Long                               'Blue Value
    cClrs As Long                               'Pixel count
    iChildren(0 To 1, 0 To 1, 0 To 1) As Long   'Child pointers
    iNext As Long                               'Next reducable node
End Type
Private OctreeNodes() As tNode, OctNodeCount As Long
Private OctDepth As Byte, OctTopGarbage As Long
Private OctColor As Long, OctReduction() As Long
Private c_Normalized32bpp() As Long

' variables used during GIF compression
Private c_aMask() As Long           ' Power of 2 look up table
Private c_aBuff() As Byte           ' general use byte array
Private c_BitBucketCount As Long    ' bit bucket bit counter
Private c_BitBucketBuff As Long     ' bit bucket contents
Private c_lCurrentBits As Long      ' current lzw compression size (variable)
Private c_lSubBlockSize As Long     ' flag to track last byte in a data subBlock
Private c_Palette() As Long         ' GIF palette working array
Private c_hPalette As Long          ' windows palette

Private c_GIFdata() As Byte     ' source bytes (mapped array, never initialized)
Private c_GIFframe() As Byte    ' 1st GIF frame from source bytes

Public Function LoadStream(inStream() As Byte, cHost As c32bppDIB, _
                    Optional ByVal streamOffset As Long, Optional ByVal streamLength As Long) As Boolean

    ' Parameters:
    ' insSream() :: a byte array containing a GIF
    ' cHost :: an initialized c32bppDIB
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    
    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim tTSA As SafeArray
    ' overlay our module level array onto the passed array
    With tTSA
        .cbElements = 1         ' byte array
        .cDims = 1              ' 1 dimensional
        .pvData = VarPtr(inStream(streamOffset))
        .rgSABound(0).cElements = streamLength
    End With
    CopyMemory ByVal VarPtrArray(c_GIFdata), VarPtr(tTSA), 4& ' apply overlay
    
    ' call routine to parse the GIF & convert it to 32bpp
    LoadStream = ParseGIF(cHost)
    CopyMemory ByVal VarPtrArray(c_GIFdata), 0&, 4&    ' remove overlay
    
End Function

Private Function ParseGIF(cHost As c32bppDIB) As Boolean

    On Error Resume Next
    ' a modified routine from some of my other GIF postings
    ' This version is scaled back and only extracts first frame
    
    ' This routine has one limitation. Some rare GIFs do not follow the
    ' standards and when those are encountered, the routine will return
    ' True to prevent GIF from being sent to other parsers.  This is
    ' important because the BMP, the last parser, sends the stream to an API
    ' to convert an unknown image to a stdPicture. If the GIF stream
    ' isn't formatted within standards that API hangs the application.
    
    Dim gLong As Long
    Dim aPointer As Long
    Dim gHeaderLen As Long
    Dim g87aStart As Long, g87aStop As Long
    Dim g89aStart As Long, g89aStop As Long
    
    ' transparency flags and variables use to tweak GIF
    Dim transUsed As Byte, TransIndex As Long
    Dim aLocalTbl As Long, gColorsUsed As Long
    Dim uniquePalette(0 To 767) As Byte
    Dim p As Long
    
    On Error GoTo ExitReadRoutine
    
    ' read signature
    ReDim c_GIFframe(0 To 5)
    CopyMemory c_GIFframe(0), c_GIFdata(0), 6&
    Select Case LCase(StrConv(c_GIFframe, vbUnicode))
        Case "gif89a", "gif87a"
        Case Else
            Exit Function
    End Select
        
    ' skip to the global color table information
    If (c_GIFdata(10) And 128) = 128 Then ' color table used? If so, skip it
        gColorsUsed = 2& ^ (1& + (c_GIFdata(10) And &H7)) ' count colors
        gHeaderLen = gColorsUsed * 3& + 13&
    Else 'no global color table; probably uses local color tables
        gHeaderLen = 13&
    End If
    aPointer = gHeaderLen
    
    Do
        Select Case c_GIFdata(aPointer)    ' read a single byte
        Case 0  ' block terminators
            aPointer = aPointer + 1&
            
        Case 33 'Extension Introducer
            aPointer = aPointer + 1&
            
            Select Case c_GIFdata(aPointer) ' read the extension type
            
            Case 255    ' application extension
                ' Get the length of extension: will always be 11
                aPointer = aPointer + c_GIFdata(aPointer + 1&) + 2&
                Call SkipGifBlock(aPointer)
                
            Case 249    ' Graphic Control Label
                        ' (description of frame & is an optional block) 8 bytes
                transUsed = (c_GIFdata(aPointer + 2&) And 1&)
                If transUsed = 1& Then ' has transparency?
                    TransIndex = c_GIFdata(aPointer + 5&) ' cache transparency index
                End If
                g89aStart = aPointer - 1&    ' location where 89a block starts
                aPointer = aPointer + 7&     ' move to end of block
                
            Case Else   ' Comment block, plain text extension, or Unknown extension
                aPointer = aPointer + 1&
                Call SkipGifBlock(aPointer)
            End Select
                
        Case 44 ' Image Descriptor (image dimensions & color table)
                ' mark position where image description starts
            g87aStart = aPointer
            aPointer = aPointer + 9& ' image data starts 10 bytes after header
            ' next byte indicates if local color table used
            If (c_GIFdata(aPointer) And 128) = 128 Then   ' local color table used?
                gColorsUsed = 2& ^ (1& + (c_GIFdata(aPointer) And &H7)) ' count colors
                aPointer = aPointer + gColorsUsed * 3&
                aLocalTbl = 1&  ' flag indicating colors from local table vs global table
            End If
            aPointer = aPointer + 2& ' move to position of first data block
            Call SkipGifBlock(aPointer)
                
            g87aStop = aPointer - 1&    ' this is where the data ends
            If g87aStop - g87aStart < 3& Then Exit Function ' invalid frame
            Exit Do
            
        Case Else
            ' shouldn't happen; abort with what we have
            Exit Function
        End Select
    Loop
    
    If Not (g87aStart = 0& Or gColorsUsed = 0&) Then ' we have a valid gif frame
    
        ' rebuild the GIF file to include only the 1st frame read
        If g89aStart > 0 Then   ' gif is 89a format
            ' resize array, copy header info & gif89a info
            ReDim c_GIFframe(0 To gHeaderLen + (g87aStop - g87aStart + 1&) + 8&)
            CopyMemory c_GIFframe(0), c_GIFdata(0), gHeaderLen
            CopyMemory c_GIFframe(gHeaderLen), c_GIFdata(g89aStart), 8&
            aPointer = gHeaderLen + 8&  ' adjust pointer for gif87a info
        Else
            ' resize array and copy header info only
            ReDim c_GIFframe(0 To gHeaderLen + (g87aStop - g87aStart + 1&))
            CopyMemory c_GIFframe(0), c_GIFdata(0), gHeaderLen
            aPointer = gHeaderLen
        End If
        ' now copy the gif 87a info
        CopyMemory c_GIFframe(aPointer), c_GIFdata(g87aStart), g87aStop - g87aStart + 1&
        c_GIFframe(UBound(c_GIFframe)) = 59 ' trailer/end of file
    
        ' fix up the left/top & width/height of overall frame
        CopyMemory c_GIFframe(aPointer + 1), 0&, 4& ' make frame left/top zero
        CopyMemory c_GIFframe(6), c_GIFframe(aPointer + 5), 4& ' make window & frame size same
    
        If transUsed = 1& Then
            ' Fix up the color table/indexes for images with transparency
            ' Why? Instead of parsing/decompressing the GIF manually, we will allow an API to do it for us.
            ' But that API can re-index the GIF which means we may lose track of the transparency
            ' color/index.  This happens whenever any color in the GIF's palette is duplicated.
            ' To prevent this from occuring, we simply replace the GIF's palette with another
            ' palette of non-duplicated entries.
            ' BTW: This way of creating GIF is still much faster than parsing the GIF by hand
            If aLocalTbl = 1& Then ' local color table else global
                ' local color table starts 10 bytes after the gif87a block
                aPointer = gHeaderLen + 10&  ' location of table within single frame array
                aLocalTbl = g87aStart + 10&  ' location of table within souce array
                ' offset single frame array when gif89a structure is used
                If Not g89aStart = 0& Then aPointer = aPointer + 8&
            Else
                aPointer = 13&        ' global table location
                aLocalTbl = 13&       ' same in both arrays
            End If
            For p = 1& To gColorsUsed - 1&  ' create non-duplicating color palette
                gLong = p * 3&
                uniquePalette(gLong) = p
                uniquePalette(gLong + 1) = p
                uniquePalette(gLong + 2) = p
            Next
            ' replace the old palette with the new one
            CopyMemory c_GIFframe(aPointer), uniquePalette(0), gColorsUsed * 3&
            Erase uniquePalette()
        Else
            TransIndex = -1&
        End If
        
        ' all done parsing the GIF file, send it to routine to convert it to a 32bpp
        ParseGIF = ConvertGIFto32bpp(TransIndex, aLocalTbl, cHost)
        
    End If
    
ExitReadRoutine:
If Err Then
    Err.Clear           ' this is a GIF format, but the format is invalid
    cHost.DestroyDIB    ' something is wrong; don't allow it to continue
    ParseGIF = True     ' to other parsers
End If
End Function

Private Sub SkipGifBlock(ByRef Ptr As Long)
    ' Routine skips a block of data within the GIF file
    Dim curByte As Byte
    curByte = c_GIFdata(Ptr)
    Do While Not curByte = 0
        Ptr = Ptr + 1& + curByte
        curByte = c_GIFdata(Ptr)
    Loop
    Ptr = Ptr + 1&
End Sub

Private Function ConvertGIFto32bpp(TransIndex As Long, tblOffset As Long, cHost As c32bppDIB) As Boolean

    ' Function converts GIF to a standard picture and then premultiplies RGB values based on the
    ' GIFs transparent index, if applicable.
    ' Note: The c_GIFframe array was already processed/filled in the ParseGIF function
    
    Dim tPic As StdPicture, tBMP As BITMAP
    
    ' used for parsing a transparent gif
    Dim X As Long, Y As Long, m As Long, dX As Long, Index As Long
    Dim gSA As SafeArray, dSA As SafeArray
    Dim Pow2(0 To 8) As Long, dibBytes() As Byte
    Dim maskShift As Long, maskAND As Long
    Dim hostDC As Long
    
    ' first: have API create a stdPicture for us
    Set tPic = iparseArrayToPicture(c_GIFframe, 0&, 1& + UBound(c_GIFframe))
    Erase c_GIFframe
    If Not tPic Is Nothing Then
        
        ' a VB stdPicture is a DIB, therefore it has a handle to the DIB bits; get it
        GetGDIObject tPic.Handle, Len(tBMP), tBMP
        If Not tBMP.bmBits = 0& Then
        
            ' have host create application's 32bpp DIB
            cHost.InitializeDIB tBMP.bmWidth, tBMP.bmHeight
            
            ' we only need to parse the palette & indexes if transparency is used
            If TransIndex = -1& Then                ' opaque GIF
                ' render GIF to our DIB DC, then ensure all alpha bytes are 255
                hostDC = cHost.LoadDIBinDC(True)
                tPic.Render hostDC + 0&, 0&, 0&, tBMP.bmWidth + 0&, tBMP.bmHeight + 0&, _
                    0&, tPic.Height, tPic.Width, -tPic.Height, ByVal 0&
                cHost.LoadDIBinDC False
                With dSA
                    .cbElements = 1
                    .cDims = 2
                    .pvData = cHost.BitsPointer
                    .rgSABound(0).cElements = cHost.Height
                    .rgSABound(1).cElements = cHost.scanWidth
                End With
                CopyMemory ByVal VarPtrArray(dibBytes), VarPtr(dSA), 4&
                iparseValidateAlphaChannel dibBytes(), True, 0&, -1&
                CopyMemory ByVal VarPtrArray(dibBytes), 0&, 4&
                cHost.Alpha = AlphaNone
                
            Else
                ' next: getting ready to parse the paletted stdPic
                Pow2(0) = 1&
                For X = 1& To tBMP.bmBitsPixel   ' power of 2 array
                    Pow2(X) = Pow2(X - 1&) * 2&
                Next
                maskAND = Pow2(tBMP.bmBitsPixel) - 1& ' AND mask for stdPic indexes
                ' we need to overlay arrays onto the GIF and the host's DIB pointers
                With gSA
                    .cbElements = 1
                    .cDims = 2
                    .pvData = tBMP.bmBits
                    .rgSABound(0).cElements = tBMP.bmHeight
                    .rgSABound(1).cElements = iparseByteAlignOnWord(tBMP.bmBitsPixel, tBMP.bmWidth)
                End With
                With dSA
                    .cbElements = 1
                    .cDims = 2
                    .pvData = cHost.BitsPointer
                    .rgSABound(0).cElements = cHost.Height
                    .rgSABound(1).cElements = cHost.scanWidth
                End With
                CopyMemory ByVal VarPtrArray(c_GIFframe), VarPtr(gSA), 4&
                CopyMemory ByVal VarPtrArray(dibBytes), VarPtr(dSA), 4&
                
                ' last: start parsing stdPic's paletted DIB
                For Y = 0& To tBMP.bmHeight - 1&
                    dX = 0&: m = 0&   ' reset dX=host DIB's X & M=stdPic DIB's X
                    maskShift = 8& - tBMP.bmBitsPixel    ' 1st bit to process
                    
                    ' note: do not loop thru using gif ScanWidth. If the GIF
                    ' width is not DWORD ligned , you will overflow the target
                    ' DIB width and eventually write to uninitialized memory
                    For X = 1& To tBMP.bmWidth&
                        ' get the palette index by shifting bits
                        Index = ((c_GIFframe(m, Y) \ Pow2(maskShift)) And maskAND)
                        
                        If Not Index = TransIndex Then  ' 100% opaque else 100% transparent
                            Index = Index * 3& + tblOffset
                            dibBytes(dX, Y) = c_GIFdata(Index + 2&)     ' make BGR vs RGB
                            dibBytes(dX + 1, Y) = c_GIFdata(Index + 1&)
                            dibBytes(dX + 2, Y) = c_GIFdata(Index)
                            dibBytes(dX + 3, Y) = 255
                        End If
                        
                        ' adjust for parsing/shifting the next index
                        If maskShift = 0& Then
                            maskShift = 8& - tBMP.bmBitsPixel ' start new byte
                            m = m + 1&                        ' next GIF byte
                        Else
                            maskShift = maskShift - tBMP.bmBitsPixel ' adjust
                        End If
                        dX = dX + 4&                          ' next Host pixel
                    Next
                Next
                ' done, remove overlays
                CopyMemory ByVal VarPtrArray(c_GIFframe), 0&, 4&
                CopyMemory ByVal VarPtrArray(dibBytes), 0&, 4&
                cHost.Alpha = AlphaSimple
            End If
            cHost.ImageType = imgGIF
            ConvertGIFto32bpp = True
        End If
    End If

End Function

   
Public Function SaveAsGIF(cHost As c32bppDIB, ByVal FileName As String, outStream() As Byte, AdaptivePalette As Boolean, OpaquenessTolerance As Byte) As Boolean

'--- Return as GIF; need to LZW compress bitmap bytes into GIF format
    
' HEAVILY MODIFIED AUTHOR'S ORIGINAL CODE (by Carles P.V.)
' - Reorganized the original routine; removed GOTOs and flow is smoother in hash routine
' - Pulled 11 global declarations into this routine.
' - The following original routines were combined into this one:
'   pvClearBlock, pvClearTable, pvCharInit & InitMasks
' - Overall changes might be a hair slower for small gifs but faster for bigger gifs
'   example using Win98's Clouds.BMP (8bit, 640x480); in IDE: 160 ms faster on my pc
' - Modified to work with 32bpp images and compression from 8 bits (256 colors) to 2 bits (4 colors)

  Const MAX_BITS                    As Long = 12  ' Per GIF docs, 12 is the Max
  Const MAX_BITSHIFT                As Long = 2 ^ MAX_BITS
  Const MAX_CODE                    As Long = 2 ^ MAX_BITS ' Should NEVER generate this code
  Const TABLE_SIZE                  As Long = 5003         ' 80% occupancy (hash)
  
  Dim LIdx     As Long      ' hash table index
  Dim lFCode   As Long      ' pixel pattern (hashed)
  Dim lC       As Long      ' most recent pixel
  Dim lEnt     As Long      ' previous known pattern
  Dim lDisp    As Long
  Dim m_lShift As Long
  
  Dim m_lCodeCount As Long ' count of codes used
  Dim m_lMaxCode As Long   ' Maximum code, given m_lBits
  '-- Block compression parameters.
  Dim m_lInitBits      As Long  ' baseline LZW compression size
  Dim m_lClearCode     As Long  ' clear code table & start again
  Dim m_lHashTable(0 To TABLE_SIZE - 1) As Long
  Dim m_lCodeTable(0 To TABLE_SIZE - 1) As Long

  ' Added by LaVolpe to read from the custom DIB strip
  Dim bExistingPattern As Boolean             ' indicates hash match found
  Dim dibRow As Long, dibCol As Long, aPtr As Long ' DIB looping parameters
  
  Dim AlphaType As AlphaTypeEnum ' alpha state: none, simple, complex
  Dim ourPalette() As Byte, palCount As Long ' palette data
  Dim FileNum As Long
  
  If Not FileName = vbNullString Then
    FileNum = iparseGetFileHandle(FileName, False)
    If (FileNum = INVALID_HANDLE_VALUE) Then Exit Function
  End If
  
    Normalize32bppImage cHost, AlphaType, OpaquenessTolerance
    
    On Error GoTo ExitRoutine
    ' have routines create a palette based on passed option
    ReDim c_Palette(1 To 256)
    ' the local palette is always first. If the image has less than 255/256 colors, then
    ' we will use those colors vs trying to adapt a palette to the image.
    ' The reason we say 255 vs 256 is that we reserve one for transparency if needed
    If CreateLocalPalette(cHost, AlphaType, palCount) = False Then
        If AdaptivePalette = True Then ' image has more than 255/256 colors
            ' if user wanted an Adaptive paletten perform that, else fall back to a websafe palette
            If CreateOptimalPalette(AlphaType, palCount) = False Then CreateWebSafePalette palCount, AlphaType
        Else ' fall back to websafe palette
            CreateWebSafePalette palCount, AlphaType
            AdaptivePalette = True
        End If
    Else ' we can use the colors from the image itself
        AdaptivePalette = False
    End If
    
    If AdaptivePalette Then
        ' here we will create a windows palette and let windows determine best color match
        ' for each color in the DIB by using the GetNearestPaletteIndex API.
    
        ' create a byte version of the LOGPALETTTE structure (4bytes+PalCount*4)
        ReDim ourPalette(1 To palCount * 4 + 4)
        CopyMemory ourPalette(5), c_Palette(1), palCount * 4& ' palette entries (RGB+flag)
        CopyMemory ourPalette(1), &H300, 2&                   ' version
        If AlphaType = AlphaNone Then                         ' nr entries
            CopyMemory ourPalette(3), palCount, 2&
        Else
            CopyMemory ourPalette(3), palCount - 1&, 2&
        End If
        c_hPalette = CreatePalette(ourPalette(1))
        Erase ourPalette()
        If c_hPalette = 0& Then
            Erase c_Palette()
            Erase c_Normalized32bpp()
            If FileName Then CloseHandle FileNum
            Exit Function   ' failed to create a palette, abort
        End If
    End If
    
    ' calc the compression bit depth
    Select Case palCount
        Case Is > 128: m_lInitBits = 8  ' 256 colors
        Case Is > 64: m_lInitBits = 7   ' 128 colors
        Case Is > 32: m_lInitBits = 6   ' 64 colors
        Case Is > 16: m_lInitBits = 5   ' 32 colors
        Case Is > 8: m_lInitBits = 4    ' 16 colors
        Case Is > 4: m_lInitBits = 3    ' 8 colors
        Case Else: m_lInitBits = 2      ' 1-4 colors
    End Select
    
    'Initialize Masks -- Init LUT for fast 2 ^ x - 1 (was InitMasks routine)
    ReDim c_aMask(0 To 16)
    c_aMask(1) = 1&
    For LIdx = 2& To 16&
        c_aMask(LIdx) = 2& * (c_aMask(LIdx - 1&) + 1&) - 1&
    Next LIdx
    
    ' oversize/guesstimate compressed GIF data & include GIF block information
    ReDim c_GIFframe(0 To 800 + (cHost.Height * cHost.Width))
    '^ 800= 13 hdr + 10 img descrip + 8 img ctrl + 768 palette + 1 EOF flag

    ' start building the GIF frame by hand, a byte at a time
    CopyMemory c_GIFframe(0), &H38464947, 4  ' add the 6byte GIF89a signature
    CopyMemory c_GIFframe(4), &H6139&, 2
    CopyMemory c_GIFframe(6), cHost.Width, &H2   ' now the width
    CopyMemory c_GIFframe(8), cHost.Height, &H2  ' & height
    ' add the color table flag and table bit depth, less 1
    c_GIFframe(10) = 128 Or (m_lInitBits - 1)
    ' can skip next 2 bytes: bkg window palette index & pixel aspect ratio
    aPtr = 13 ' keep track of next array element
    For LIdx = 1 To palCount ' copy RGB palette
        CopyMemory c_GIFframe(aPtr), c_Palette(LIdx), 3&
        aPtr = aPtr + 3
    Next
    If c_hPalette Then Erase c_Palette ' no longer needed if hPalette created
    If LIdx < c_aMask(m_lInitBits) + 2 Then
        ' move pointer to position after last palette entry (not all entries have been used)
        aPtr = aPtr + ((c_aMask(m_lInitBits) + 1) - palCount) * 3&
    End If
    ' add the image control block if needed
    If AlphaType > AlphaNone Then
        CopyMemory c_GIFframe(aPtr), &H4F921, 3& ' Introducer(33);BlockID(249);BlockLen(4)
        c_GIFframe(aPtr + 3) = 1                 ' transparency flag
        ' skip next 2 bytes which is the Delay time; not needed for single frame
        c_GIFframe(aPtr + 6) = palCount - 1 ' palette functions herein always use last index
        ' skip next byte which is the block terminator - zero
        aPtr = aPtr + 8 ' next position in array
    End If
    ' add the image description block
    c_GIFframe(aPtr) = 44 ' BlockID
    ' Left and Top coords are unnecessary for single frame; leave zeros
    CopyMemory c_GIFframe(aPtr + 5), c_GIFframe(6), &H4 ' width & height
    ' the packed byte would be next; but we'll leave it as zero
    ' since frame is not interlaced and not using a local color table.
    c_GIFframe(aPtr + 10) = m_lInitBits  ' LZW compression size; 1st byte of compressed image
    aPtr = aPtr + 11                     '10=img description block size + LZW compression byte
    ' next comes compressing DIB into LZW sub blocks, then finishing off stream
    
    '-- Reset output buffer values
    c_BitBucketBuff = 0
    c_BitBucketCount = 0
    ReDim c_aBuff(0 To 254) ' (was pvCharInit)

    '-- Set up the necessary startup values
    m_lInitBits = m_lInitBits + 1& ' per lzw compression specs, increment one
    c_lCurrentBits = m_lInitBits
    m_lMaxCode = c_aMask(c_lCurrentBits) ' max pattern items
    m_lClearCode = c_aMask(m_lInitBits - 1) + 1 ' flag indicating reset decoding tables
    m_lCodeCount = m_lClearCode + 2
    
    '-- Set hash code range bound for shifting
    lFCode = TABLE_SIZE
    Do While lFCode < 65536
        m_lShift = m_lShift + 1
        lFCode = lFCode + lFCode
    Loop
    m_lShift = 1 + c_aMask(8 - m_lShift)
    'Added by LaVolpe -- quick erase, setting all table entries to -1 (was pvClearTable)
    FillMemory m_lHashTable(0), TABLE_SIZE * 4, 255  ' clear hash table
    
    '-- Start...
    
    c_lSubBlockSize = 1 ' position for 1st byte in data sub block
    Call pvOutputCode(m_lClearCode, aPtr) ' all images begin with a clear table flag
    
    ' start LZW patterns
    lEnt = GetPaletteIndex(0, cHost.Height - 1, palCount)
    dibCol = 1& ' start looping on 2nd pixel (zero bound array)
    
    'Added by LaVolpe -- looping is my enhancement; removed various GOSUBs
    ' hash algorithm reorganized, comments added, modified very little
    For dibRow = cHost.Height - 1 To 0 Step -1
    
        ' process each pixel in line of image
        For dibCol = dibCol To cHost.Width - 1
        
            lC = GetPaletteIndex(dibCol, dibRow, palCount) ' get next palette index
            
            lFCode = lC * MAX_BITSHIFT + lEnt   ' add to existing pattern
            LIdx = (lC * m_lShift) Xor lEnt     ' XOR hashing
    
            If LIdx >= TABLE_SIZE Then LIdx = 0 ' added by LaVolpe (sanity check)
            
            If (m_lHashTable(LIdx) = lFCode) Then   ' found existing pattern
                lEnt = m_lCodeTable(LIdx)
            Else
                If (m_lHashTable(LIdx) > -1) Then ' else Empty slot
                    
                    lDisp = TABLE_SIZE - LIdx     ' Secondary hash (after G. Knott)
                    If (LIdx = 0) Then lDisp = 1
                    
                    Do  ' Hash Probing
                        LIdx = LIdx - lDisp
                        If (LIdx < 0) Then LIdx = LIdx + TABLE_SIZE
        
                        If (m_lHashTable(LIdx) = lFCode) Then
                            lEnt = m_lCodeTable(LIdx)
                            bExistingPattern = True
                            Exit Do
                        End If
        
                    Loop While (m_lHashTable(LIdx) > 0) ' Continue probing
                End If
                
                If bExistingPattern = True Then
                    bExistingPattern = False    ' reset flag
                Else
                    Call pvOutputCode(lEnt, aPtr)    ' write previous pattern
                    lEnt = lC  ' set current palette index as previous pattern
                    m_lCodeTable(LIdx) = m_lCodeCount ' store code & hash index
                    m_lHashTable(LIdx) = lFCode
                    
                    ' check for LZW compression increments
                    If m_lCodeCount > m_lMaxCode Then
                        ' ran out of codes for current compression size
                        If (m_lCodeCount = MAX_CODE) Then
                            ' add clear code to output stream (was pvClearBlock)
                            Call pvOutputCode(m_lClearCode, aPtr)
                            FillMemory m_lHashTable(0), TABLE_SIZE * 4, 255  ' clear hash table
                            m_lCodeCount = m_lClearCode + 1 ' reset code counter to 1 less 'cause it is incremented right away
                            c_lCurrentBits = m_lInitBits    ' reset baseline LZW compression size
                        Else
                            c_lCurrentBits = c_lCurrentBits + 1 ' increment compression size
                        End If
                        m_lMaxCode = c_aMask(c_lCurrentBits) ' new max count for current compression size
                    End If
                    m_lCodeCount = m_lCodeCount + 1  ' increment the number of patterns
                End If
            End If
        Next dibCol
        dibCol = 0  ' reset to 1st pixel of line
    Next dibRow

    '--  Put out the final code & image data termination code
    Call pvOutputCode(lEnt, aPtr)
    ' finish off the stream
    Call pvOutputCode(m_lClearCode + 1, aPtr) ' add end of image data flag
    Call pvOutputCode(-1, aPtr)  ' flush remaining bytes in output bitbucket
    ' the next byte would be a sub block terminator which is zero; ignore & go on
    c_GIFframe(aPtr + 1) = 59   ' add end of file flag
    aPtr = aPtr + 2
        
    If FileNum Then
        SetFilePointer FileNum, 0&, 0&, 0&
        WriteFile FileNum, c_GIFframe(0), aPtr, lEnt, ByVal 0&
        SaveAsGIF = (aPtr = lEnt)
    Else
        ReDim outStream(0 To aPtr - 1&)
        CopyMemory outStream(0), c_GIFframe(0), aPtr
        SaveAsGIF = True
    End If

ExitRoutine:
    If Err Then
'        Stop        ' for testing if error occurs
        Err.Clear
'        Resume      ' for testing if error occurs
    End If
    If FileNum Then CloseHandle FileNum
    Erase c_Normalized32bpp()
    Erase c_GIFframe()
    Erase c_aMask()
    Erase c_aBuff()
    Erase c_Palette()
    If c_hPalette Then DeleteObject c_hPalette
End Function

Private Function CreateLocalPalette(cHost As c32bppDIB, AlphaType As AlphaTypeEnum, palCount As Long) As Boolean
    
    ' Function determines if image can be palettized (256 colors or less)
    ' If the image has alphablended pixels (i.e., anything other than 0 or 255 for the alpha channel),
    '   then the pixels will be treated as non-alpha blended and the result will be poor.
    
    Dim X As Long, Y As Long
    Dim prevColor As Long, newColor As Boolean
    Dim maxPalEntry As Long, Index As Long
    
    If AlphaType = AlphaComplex Then maxPalEntry = 255 Else maxPalEntry = 256
    
    ' count unique colors (maximum of 256 if we are to palettize)
    prevColor = -1&
    palCount = 0&
    For Y = 0& To UBound(c_Normalized32bpp, 2)
        For X = 0& To UBound(c_Normalized32bpp, 1)
            If Not c_Normalized32bpp(X, Y) = prevColor Then      ' is this same as previous color?
                If c_Normalized32bpp(X, Y) > -1& Then
                    prevColor = c_Normalized32bpp(X, Y) ' cache current color for faster looping
                    
                    ' use binary search routine to locate it if it already exists
                    Index = iparseFindColor(c_Palette(), c_Normalized32bpp(X, Y), palCount, newColor)
                    If newColor = True Then         ' color not in palette yet
                        palCount = palCount + 1&    ' increment palette count
                        If palCount > maxPalEntry Then Exit Function ' maxed palette out?
                        If Index < palCount Then ' keep palette entries sorted for binary search routine
                            CopyMemory c_Palette(Index + 1&), c_Palette(Index), (palCount - Index) * 4&
                        End If
                        c_Palette(Index) = c_Normalized32bpp(X, Y) ' add new color to the palette
                    End If
                End If
            End If
        Next
    Next
    If palCount <= maxPalEntry Then
        If maxPalEntry = 255 Then
            palCount = palCount + 1
            c_Palette(palCount) = &H7FFFFFFF
        End If
        CreateLocalPalette = True ' we can palettize
    End If
End Function

Private Function GetPaletteIndex(Column As Long, Row As Long, palCount As Long) As Long

    If c_Normalized32bpp(Column, Row) = -1& Then ' special code for full transparency
        GetPaletteIndex = palCount - 1           ' see Normalize32bppImage for more
    ElseIf c_hPalette Then
        ' using a windows palette so we can use this API to find nearest match
        GetPaletteIndex = GetNearestPaletteIndex(c_hPalette, c_Normalized32bpp(Column, Row))
    Else ' using palette created from the image's actual colors (256 or less)
        GetPaletteIndex = iparseFindColor(c_Palette(), c_Normalized32bpp(Column, Row), palCount, False) - 1
    End If

End Function


Private Sub pvOutputCode(ByVal lCode As Long, ByRef arrayPtr As Long)
' Helper function for the SaveAsGIF routine
' MODIFIED/REORGANIZED AUTHOR'S ORIGINAL CODE (by Carles P.V.)
' - Also combined original pvCharOut and pvFlushChar routines herein

    If lCode < 0 Then
    
        If Not c_BitBucketCount = 0 Then ' add last bits of the output buffer to array
            c_aBuff(c_lSubBlockSize) = c_BitBucketBuff
            c_lSubBlockSize = c_lSubBlockSize + 1
        End If
        If Not c_lSubBlockSize = 1 Then ' still have bytes to write
            c_aBuff(0) = c_lSubBlockSize
            ' ensure array large enough for this block + 5 extra bytes needed to finish off stream
            If UBound(c_GIFframe) < arrayPtr + c_lSubBlockSize + 6 Then
                ReDim Preserve c_GIFframe(0 To arrayPtr + c_lSubBlockSize + 6)
            End If
            CopyMemory c_GIFframe(arrayPtr), c_aBuff(0), c_lSubBlockSize + 1
            arrayPtr = arrayPtr + c_lSubBlockSize + 1
        End If
    
    Else
        ' add latest code to the bitbucket & track total bit count
        c_BitBucketBuff = c_BitBucketBuff Or (lCode * (c_aMask(c_BitBucketCount) + 1))
        c_BitBucketCount = c_BitBucketCount + c_lCurrentBits
    
        Do Until c_BitBucketCount < 8
            ' remove 8 bits at a time and place in output buffer (byte array)
            c_aBuff(c_lSubBlockSize) = (c_BitBucketBuff And &HFF&)
            c_BitBucketBuff = c_BitBucketBuff \ &H100&
            c_BitBucketCount = c_BitBucketCount - 8
            If (c_lSubBlockSize = 254) Then
                ' max allowable subblock data size is 255 byte blocks; write it
                c_aBuff(0) = c_lSubBlockSize
                If UBound(c_GIFframe) < arrayPtr + 260 Then
                    ' sanity check; unless the image is not compressible we
                    ' shouldn't trigger this Redim. ^^ 260=255+5 end of file/image bytes
                    ReDim Preserve c_GIFframe(0 To arrayPtr + 512)
                End If
                CopyMemory c_GIFframe(arrayPtr), c_aBuff(0), c_lSubBlockSize + 1
                arrayPtr = arrayPtr + c_lSubBlockSize + 1
                c_lSubBlockSize = 0
            End If
            ' keep track of next byte position to write to
            c_lSubBlockSize = c_lSubBlockSize + 1
        Loop
    End If

End Sub

' THE PALETTE ROUTINES BELOW AND THIS ROUTINE ARE FROM THE FOLLOWING SOURCE.
' I have modified each slightly, so get the original routines from that source...
' source: http://www.vbaccelerator.com/home/VB/Code/vbMedia/Image_Processing/Colour_Depth_Reduction/VB6_Colour_Depth_Sample.asp

Private Function CreateOptimalPalette(AlphaType As AlphaTypeEnum, palCount As Long, _
                                Optional ByVal nLevels As Long = 6) As Boolean

' slightly modified by me

Dim X As Long, Y As Long
Dim pPal As Boolean
    
    If AlphaType = AlphaComplex Then palCount = 255 Else palCount = 256

   On Error GoTo eh
   
   If nLevels = 0 Then nLevels = 1
   OctDepth = nLevels
   ReDim OctReduction(1 To nLevels)
   
   ReDim OctreeNodes(1 To 50) 'Allocates initial storage
   OctTopGarbage = 0
   OctNodeCount = 1
   OctColor = 0
   
   For Y = 0 To UBound(c_Normalized32bpp, 2)
      For X = 0 To UBound(c_Normalized32bpp, 1)
         'Adds the current pixel to the color octree
         If c_Normalized32bpp(X, Y) > -1& Then
            AddClr 1, 1, 0, 255, 0, 255, 0, 255, _
                c_Normalized32bpp(X, Y) And &HFF, (c_Normalized32bpp(X, Y) \ &H100&) And &HFF, (c_Normalized32bpp(X, Y) \ &H10000) And &HFF
         
            Do While OctColor > palCount
               'Combine the levels to get down to desired palette size
               If CombineNodes = False Then Exit Do
            Loop
            If OctColor > palCount Then Exit Function
         End If
      Next X
   Next Y
   
   palCount = 0&
   'Go through octree and extract colors
   For Y = 1 To UBound(OctreeNodes)
      If OctreeNodes(Y).bIsLeaf Then
         With OctreeNodes(Y)
            'Add .vRed / .cClrs, .vGreen / .cClrs, .vBlue / .cClrs
            palCount = palCount + 1
            c_Palette(palCount) = (.vRed \ .cClrs) Or _
                                (.vGreen \ .cClrs) * &H100& Or _
                                (.vBlue \ .cClrs) * &H10000
         End With
      End If
   Next Y
   If AlphaType > AlphaSimple Then palCount = palCount + 1
    CreateOptimalPalette = True
    
eh:
    If Err Then Err.Clear
    Erase OctreeNodes()
    Erase OctReduction()
    
End Function

Private Sub AddClr(ByVal iBranch As Long, ByVal nLevel As Long, _
      ByVal vMinR As Byte, ByVal vMaxR As Byte, _
      ByVal vMinG As Byte, ByVal vMaxG As Byte, _
      ByVal vMinB As Byte, ByVal vMaxB As Byte, _
      ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
        
    ' helper function for CreateOptimalPalette
    Dim iR As Byte, iG As Byte, iB As Byte
    Dim vMid As Long, iIndex As Long
    
   
   'Find mid values for colors and decide which path to take
   'Also update max and min values for later call to self
   vMid = vMinR / 2 + vMaxR / 2
   If r > vMid Then
      iR = 1
      vMinR = vMid
   Else
      iR = 0
      vMaxR = vMid
   End If
   
   vMid = vMinG / 2 + vMaxG / 2
   If g > vMid Then
      iG = 1
      vMinG = vMid
   Else
      iG = 0
      vMaxG = vMid
   End If
   
   vMid = vMinB / 2 + vMaxB / 2
   If b > vMid Then
      iB = 1
      vMinB = vMid
   Else
      iB = 0
      vMaxB = vMid
   End If
   
   If OctreeNodes(iBranch).iChildren(iR, iG, iB) = 0 Then  'If no child here then...
      iIndex = GetFreeNode 'Get a new node index
      OctreeNodes(iBranch).iChildren(iR, iG, iB) = iIndex
      OctreeNodes(iBranch).cClrs = OctreeNodes(iBranch).cClrs + 1
      With OctreeNodes(iIndex)   'Clear/set data
         .bIsLeaf = (nLevel = OctDepth): .vGreen = 0
         .iNext = 0: .cClrs = 0: .vBlue = 0: .vRed = 0
      End With
   Else  'Has a child here
      iIndex = OctreeNodes(iBranch).iChildren(iR, iG, iB)
   End If
   
   
   If OctreeNodes(iIndex).bIsLeaf Then  'If it is a leaf
      With OctreeNodes(iIndex)
         If .cClrs = 0 Then OctColor = OctColor + 1
         .cClrs = .cClrs + 1
         .vRed = .vRed + r
         .vGreen = .vGreen + g
         .vBlue = .vBlue + b
      End With
   Else  'Otherwise
      With OctreeNodes(iIndex)
         If .bAddedReduce = False Then    'If 2 or more colors, add to  reducable OctreeNodes list
            .iNext = OctReduction(nLevel)
            OctReduction(nLevel) = iIndex
            .bAddedReduce = True
            'Debug.Print "Reducable Added: " & iIndex, "Level = " & nLevel, "Top was = " & .iNext
         End If
      End With
      
      'Search a level deeper
      AddClr iIndex, nLevel + 1, vMinR, vMaxR, vMinG, vMaxG, _
         vMinB, vMaxB, r, g, b
   End If
End Sub

Private Function CombineNodes() As Boolean
Dim i As Long, iIndex As Long
Dim iR As Byte, iG As Byte, iB As Byte
Dim nR As Long, nG As Long, Nb As Long, nPixel As Long
   
    ' helper function for CreateOptimalPalette
   
   For i = OctDepth To 1 Step -1   'Find deepest reducable level
      If OctReduction(i) <> 0 Then Exit For
   Next i
   
   If i = 0 Then Exit Function
   iIndex = OctReduction(i)
   OctReduction(i) = OctreeNodes(iIndex).iNext
   
   For i = 0 To 7
      'iR = IIf((i And 1) = 1, 1, 0)
      If (i And 1) = 1 Then
         iR = 1
      Else
         iR = 0
      End If
      'iG = IIf((i And 2) = 2, 1, 0)
      If (i And 2) = 2 Then
         iG = 1
      Else
         iG = 0
      End If
      'iB = IIf((i And 4) = 4, 1, 0)
      If (i And 4) = 4 Then
         iB = 1
      Else
         iB = 0
      End If
      
      If OctreeNodes(iIndex).iChildren(iR, iG, iB) <> 0 Then 'If there is a child
         With OctreeNodes(OctreeNodes(iIndex).iChildren(iR, iG, iB))
            'Add red, green, blue, and pixel count to running total
            nR = nR + .vRed
            nG = nG + .vGreen
            Nb = Nb + .vBlue
            nPixel = nPixel + .cClrs
            FreeNode OctreeNodes(iIndex).iChildren(iR, iG, iB) 'Free the node
            OctColor = OctColor - 1
         End With
         OctreeNodes(iIndex).iChildren(iR, iG, iB) = 0 'Clear the link
      End If
   Next i
   
   OctColor = OctColor + 1
   
   'Set the new node data
   With OctreeNodes(iIndex)
      .cClrs = nPixel
      .bIsLeaf = True
      .vRed = nR
      .vBlue = Nb
      .vGreen = nG
   End With
   CombineNodes = True
   
End Function

Private Sub FreeNode(ByVal iNode As Long)
    ' helper function for CreateOptimalPalette
   OctreeNodes(iNode).iNext = OctTopGarbage
   OctTopGarbage = iNode
   OctreeNodes(iNode).bIsLeaf = False 'Necessary for final loop through
   OctreeNodes(iNode).bAddedReduce = False
   OctNodeCount = OctNodeCount - 1
End Sub

Private Function GetFreeNode() As Long
    ' helper function for CreateOptimalPalette
    Dim i As Long, iR As Byte, iG As Byte, iB As Byte
   OctNodeCount = OctNodeCount + 1
   If OctTopGarbage = 0 Then
      If OctNodeCount > UBound(OctreeNodes) Then
         i = OctNodeCount * 1.1
         ReDim Preserve OctreeNodes(1 To i)
      End If
      GetFreeNode = OctNodeCount
   Else
      GetFreeNode = OctTopGarbage
      OctTopGarbage = OctreeNodes(OctTopGarbage).iNext
      For i = 0 To 7
         'iR = IIf((i And 1) = 1, 1, 0)
         If (i And 1) = 1 Then
            iR = 1
         Else
            iR = 0
         End If
         'iG = IIf((i And 2) = 2, 1, 0)
         If (i And 2) = 2 Then
            iG = 1
         Else
            iG = 0
         End If
         'iB = IIf((i And 4) = 4, 1, 0)
         If (i And 4) = 4 Then
            iB = 1
         Else
            iB = 0
         End If
         OctreeNodes(GetFreeNode).iChildren(iR, iG, iB) = 0
      Next i
   End If
End Function

Private Sub CreateWebSafePalette(palCount As Long, AlphaType As AlphaTypeEnum)

    ' slightly modified to return colors to a c_palette() array
    Dim X As Long
    Dim r As Integer, g As Integer, b As Integer
    For r = 0 To 255 Step 51
        For g = 0 To 255 Step 51
            For b = 0 To 255 Step 51
                X = X + 1&
                c_Palette(X) = RGB(r, g, b)
            Next
        Next
    Next
    
    ' now fill most of remaing entries with gray shades
    For r = 8 To 255 Step 8
        X = X + 1&
        c_Palette(X) = RGB(r, r, r)
    Next
    
    If AlphaType = AlphaNone Then palCount = X + 1& Else palCount = X + 2&
    
End Sub

Private Sub Normalize32bppImage(cHost As c32bppDIB, AlphaType As AlphaTypeEnum, opaqueTolerance As Byte)
    ' transform 32bpp (including premultiplied pixels) to RGB format
    ' GIFs obviously do not support per-pixel opaqueness
    ' so what we do here is to make it more bearable
    Dim X As Long, Y As Long, scanWidth As Long
    Dim Alpha As Long, dstX As Long
    Dim tSA As SafeArray
    
    On Error GoTo ExitRoutine
    iparseOverlayHost_2DbyHost c_GIFframe(), VarPtr(tSA), cHost
    iparseValidateAlphaChannel c_GIFframe, False, AlphaType, 0& ' used in other palette routines if necessary
    scanWidth = cHost.scanWidth - 1&
    
    ReDim c_Normalized32bpp(0 To cHost.Width - 1, 0 To cHost.Height - 1)
    For Y = 0 To cHost.Height - 1
        dstX = 0&
        For X = 0 To scanWidth Step 4
            Select Case c_GIFframe(X + 3&, Y)
            Case 0: ' full transparent
                c_Normalized32bpp(dstX, Y) = -1& ' special flag for full transparency
            Case 255
                c_Normalized32bpp(dstX, Y) = c_GIFframe(X + 2, Y) Or (c_GIFframe(X + 1, Y) * &H100&) Or (c_GIFframe(X, Y) * &H10000)
            Case Is < opaqueTolerance
                ' per user request, make any pixel < opaqueTolerance fully transparent
                c_Normalized32bpp(dstX, Y) = -1& ' special flag for full transparency
            Case Else
                ' per user request, make any pixel >=opaqueTolerance fully opqaue
                Alpha = c_GIFframe(X + 3&, Y)
                c_Normalized32bpp(dstX, Y) = ((c_GIFframe(X + 2&, Y) * 255&) \ Alpha) Or _
                                             ((c_GIFframe(X + 1&, Y) * 255&) \ Alpha) * &H100& Or _
                                             ((c_GIFframe(X, Y) * 255&) \ Alpha) * &H10000
            End Select
            dstX = dstX + 1&
        Next
    Next
ExitRoutine:
    iparseOverlayHost_2DbyHost c_GIFframe(), 0&, Nothing
End Sub

