VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write PNG/APNG image formats ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' APNG are rare today, but may become popular down the road
'  and was designed to compete with the MNG format
' References:
'   APNG - http://animatedpng.com/ also http://animatedpng.com/index.php/documentation/
'   MNG - http://www.libpng.org/pub/mng/mngdocs.html

' APNGs are to PNG what animated GIFs are to GIF. In fact, APNG pretty much
' replicated the logic of animated GIFs and applied it to the PNG format. Yes, APNG
' is a normal PNG file with additional 'chunks' added that provide frame
' descriptions, and additional frame images. PNG viewers that are not APNG-aware
' will display a single frame without even knowing they are doing it. APNG-aware
' viewers can display the animated PNG or a single frame.

' APNG is not officially recognized by PNG. This is probably because APNG competes
' with MNG (PNG's official animated PNG format). But MNG is a different file format
' than PNG and requires separate parsing routines. Whereas APNG uses PNG format,
' so a PNG viewer basically only needs to handle some additional 'chunks' to
' make itself APNG-aware.

' MNGs are similar to APNGs but are so much more
' Because APNG & GIF are nearly identical in the way they work, this class and the
'   cFunctionsGIF class are very similar, but not identical, in their Rendering function

' APNG restrictions when compared to GIF
' - Color tables
'   :: Each GIF frame can have its own color table, 256 colors or less
'   :: If color table used, every APNG frame must use the same color table
' - Frame sizes
'   :: Each GIF frame can be a different size & must fit within the overall GIF 'window/canvas'
'   :: APNG 1st image must be sized to overall 'window/canvas'. Additional frames can be individually sized
' - Bit depth
'   :: Each GIF frame can be a different bit depth
'   :: Every APNG frame must be the same bit depth & PNG color type

' Parsed PNG chunk names & their numerical equivalent (those used in this class)
' Per png specs; using alphabet chars is a no-no should system not support those characters
' http://www.libpng.org/pub/png/spec/1.1/PNG-Chunks.html
Private Const chnk_IHDR As Long = &H52444849 'Image header
Private Const chnk_IDAT As Long = &H54414449 'Image data
Private Const chnk_IEND As Long = &H444E4549 'End of Image
' //// APNG chunk names
Private Const chnk_acTL = &H4C546361    ' Animation Ctrl Chunk: must appear before first IDAT
Private Const chnk_fcTL = &H4C546366    ' Frame Ctrl Chunk: if before IDAT then IDAT is first frame; preceedes fdAT for additional frames
Private Const chnk_fdAT = &H54416466    ' Frame Data Chunk. Same as IDAT except additonal 4 byte sequence identifier prepended

Private Type APNGfDatStruct
    Pointer As Long                 ' pointer in source data where frame image data begins
    length As Long                  ' length of the image data block
    CRC32 As Long                   ' the CRC32 value of the image data
End Type
Private Type APNGFrameStruct
    DatCount As Long                ' how many blocks for entire frame; number of fDat()
    fDat() As APNGfDatStruct        ' collection of data block info
    Delay As Long                   ' frame delay
    Size As RECTI                   ' frame position
    Disposal As Integer             ' packed: 1st byte = disposal code; 2nd byte = transparency usage
End Type
Private Type APNGDataStruct
    FrameCount As Long              ' number of frames; number of Frames()
    Frames() As APNGFrameStruct     ' collection of frame structures
    Loops As Long                   ' number of animation loops
    Offset As Long                  ' offset where 1st data block will begin in the stream
    Size As POINTAPI                ' size of the overall animation
    Frame As Long                   ' current frame
    Code2Size As Long               ' if PNG Code 2 used, what height buffer do we need?
End Type

Private CRC32LUT() As Long          ' used for creating/parsing CRC32 in network byte order
Private m_APNGdat As APNGDataStruct ' our animated png
Private m_BlankCanvas() As Byte     ' array to serve as both erasing frames & caching Disposal code 3 data
Private m_IStream As IUnknown       ' the png's data stream
Private m_Token As cGDIpToken       ' GDI+ token

Public Function IsPNGResourceFile(ByVal FileHandle As Long) As Long

    ' returns one of 3 values
    ' 0 = not a png
    ' 1 = single frame png
    ' other = multi-frame png

    If m_Token.Token = 0& Then Exit Function
    
    Dim lValue As Long, lRead As Long
    SetFilePointer FileHandle, 0&, 0&, 0&
    ReadFile FileHandle, lValue, 4&, lRead, ByVal 0&
    If lValue = png_Signature1 Then
        ReadFile FileHandle, lValue, 4&, lRead, ByVal 0&
        If lValue = png_Signature2 Then
            ' minimal parsing is performed in next call; looking for APNG
            IsPNGResourceFile = pvParsePng(m_BlankCanvas(), FileHandle)
        End If
    End If

End Function

Public Function IsPNGResource(cImageData As cGDIpMultiImage, inStream() As Byte, CacheData As Boolean) As Long

    ' returns one of 3 values
    ' 0 = not a png
    ' 1 = single frame png
    ' other = multi-frame png

    If m_Token.Token = 0& Then Exit Function
    
    Dim lValue As Long, SeqDelay() As Long
    Dim hHandle As Long, I As Long, J As Long
    Const CRCpolynomial As Long = &HEDB88320
    ' &HEDB88320 is the official polynomial used by CRC32 in PKZip & zLIB.
    
    CopyMemory lValue, inStream(0), 4&
    If lValue = png_Signature1 Then
        CopyMemory lValue, inStream(4), 4&
        If lValue = png_Signature2 Then
            
            ReDim CRC32LUT(0 To 255)                            ' create a CRC lookup table (LUT)
            For I = 0& To 255&
                lValue = I
                For J = 8& To 1& Step -1&
                    If (lValue And 1&) Then
                        lValue = (((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF) Xor CRCpolynomial
                    Else
                        lValue = ((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    End If
                Next J
                CRC32LUT(I) = lValue
            Next I
            IsPNGResource = pvParsePng(inStream())
            Erase CRC32LUT()
            
            If IsPNGResource > 1& Then
                If GdipCreateBitmapFromScan0(m_APNGdat.Size.X, m_APNGdat.Size.Y, 0&, lvicColor32bppAlpha, ByVal 0&, hHandle) = 0& Then
                    ' Create a general use array used during frame navigation
                    If m_APNGdat.Code2Size Then
                        ReDim m_BlankCanvas(0 To m_APNGdat.Size.X * 4& * m_APNGdat.Code2Size - 1&)
                    Else
                        Erase m_BlankCanvas
                    End If
                    ' tidy up some PNG info to aid the rendering routine to process frames faster/efficiently
                    ' Don't let first frame have disposal code of 3. It's pointless
                    If (m_APNGdat.Frames(0).Disposal And &HFF&) = 2 Then m_APNGdat.Frames(0).Disposal = (m_APNGdat.Frames(0).Disposal And &HFF00&) Or 1
                    ' Prepare to create first frame & then create it
                    m_APNGdat.Frame = m_APNGdat.FrameCount - 1&
                    ' Don't let last frame have any disposal code either. It's pointless
                    m_APNGdat.Frames(m_APNGdat.Frame).Disposal = (m_APNGdat.Frames(m_APNGdat.Frame).Disposal And &HFF00&)
                    ' Ensure we can load the first frame else assume corrupted
                    If Me.MoveToFrame(0&, inStream(), hHandle) Then
                        cImageData.InitializeContent ObjPtr(Me), lvicPicTypePNG, m_APNGdat.FrameCount
                        ReDim SeqDelay(0 To 1, 1 To m_APNGdat.FrameCount)
                        For I = 1 To m_APNGdat.FrameCount
                            SeqDelay(0, I) = I - 1&
                            SeqDelay(1, I) = m_APNGdat.Frames(I - 1&).Delay
                        Next
                        cImageData.SetSequenceDuration SeqDelay()            ' send off & done
                        cImageData.CacheSourceInfo VarPtrArray(inStream()), hHandle, lvicPicTypePNG, CacheData, False
                        IsPNGResource = True
                    Else
                        Set m_IStream = Nothing
                        Erase m_BlankCanvas()
                    End If
                End If
            End If
        End If
    End If

End Function

Private Function pvParsePng(Stream() As Byte, Optional FileHandle As Long, Optional ParsableFrame As Long) As Long

    ' Return values
    '   0 = failure
    '   1 = single frame PNG
    '  other value = multi-frame PNG

    ' GDI+ cannot process frames of an APNG file; therefore, if we want it to process them,
    '   we need to trick GDI+ into thinking it is multiple PNG files. We can
    '   process these ourselves using zLIB and/or manual decompression of PNG
    '   data, but it is soooo much slower than GDI+.
    ' The routine scans the source file, tracks where fdAT and IDAT chunks exist, their sizes
    
    Dim ptrArray As Long        ' used to ensure parsing doesn't go past end of a corrupted array
    Dim readRtn As Long         ' bytes read when ReadFile is used
    Dim ChunkName As Long       ' name of the chunk (netwk byte order)
    Dim ChunkLen As Long        ' length of the chunk
    Dim crc32Value As Long      ' crc value of chunk (netwk byte order)
    Dim streamLength As Long    ' size of array to prevent reading past end of array
    Dim maxDataSize As Long     ' sum of the largest frame's data chunks
    Dim curDataSize As Long     ' sum of current frame's data chunks
    Dim lBounds As Long
   
    pvParsePng = 1&             ' default return value
    
    If FileHandle Then          ' when passed, we are just testing for APNG vs PNG
        ReDim Stream(0 To 4096)
        ReadFile FileHandle, Stream(0), 4096, readRtn, ByVal 0&
        streamLength = readRtn - 4&
    Else
        ptrArray = 8&
        streamLength = UBound(Stream) + 1&
    End If
    If streamLength < 57& Then Exit Function
    
    With m_APNGdat                ' reset should class have been called more than once
        .Frame = 0&
        .FrameCount = 0&
        .Loops = 0&
        .Offset = 0&
        .Size.X = 0&: .Size.Y = 0&
    End With
    
    Do ' scan & minimally process PNG file for eventual hand off to GDI+
        ' Chunks consist of 4 bytes for the length of the chunk
        '                 + 4 bytes for chunk name
        '                 + n bytes for the chunk
        '                 + 4 bytes for a CRC value (crc includes: data chunks + name chunk)
        CopyMemory ChunkLen, Stream(ptrArray), 4& ' length of the current chunk
        ChunkLen = modCommon.ReverseLong(ChunkLen) ' longs are big endian, need little endian for Windows
        If ChunkLen < 0& Then Exit Function ' if corrupted data, abort
        
        If ptrArray + ChunkLen + 12& > streamLength Then
            If FileHandle Then
                readRtn = 0&            ' haven't found IDAT yet; cache some more bytes
                ReadFile FileHandle, Stream(4), 4092, readRtn, ByVal 0&
                If readRtn = 0& Then Exit Function
                streamLength = readRtn - 4&
                ptrArray = 0&
            Else
                Exit Function
            End If
        End If
        CopyMemory ChunkName, Stream(ptrArray + 4&), 4& ' extract the chunk name
        CopyMemory crc32Value, Stream(ptrArray + ChunkLen + 8&), 4&
        
        Select Case ChunkName
        ' when scanning, we are looking for the IDAT & fdAT chunks within the file/array
        ' These offsets will be cached so they can be inserted into an array to be processed by GDI+
        ' A lot of work to make GDI+ process the APNG, but worth it. GDI+ is much faster than manual processing
        
        Case chnk_IDAT ' APNG terms: Default image to be displayed if parser does not support APNG
                       '             Also the 1st frame in animated PNG only if a fcTL and acTL preceeded this
            If m_APNGdat.FrameCount = 0& Then       ' we do not have an APNG file
                If ParsableFrame = 0& Then
                    Exit Function                   ' exit here, not animated png
                Else                                ' set/called by SaveAsAPNG
                    With m_APNGdat.Frames(0)
                        ReDim Preserve .fDat(0 To .DatCount)
                        .fDat(.DatCount).Pointer = ptrArray
                        .fDat(.DatCount).length = ChunkLen
                        .DatCount = .DatCount + 1&
                    End With
                End If
            ElseIf FileHandle Then                  ' checking for APNG only
                Erase m_APNGdat.Frames()
                pvParsePng = m_APNGdat.FrameCount
                Exit Function
            ElseIf (m_APNGdat.Frame And &HFFFF&) = 1& Then ' will be the first frame
                With m_APNGdat.Frames(0)
                    ReDim Preserve .fDat(0 To .DatCount)
                    .fDat(.DatCount).CRC32 = crc32Value
                    .fDat(.DatCount).length = ChunkLen
                    .fDat(.DatCount).Pointer = ptrArray + 8&
                    .DatCount = .DatCount + 1&
                    curDataSize = curDataSize + ChunkLen + 12&
                End With
            End If          ' else skipped; is default PNG for non-APNG parsers, but not for APNG parsers
            If m_APNGdat.Offset = 0& Then m_APNGdat.Offset = ptrArray  ' location where IDATs start
            
        Case chnk_fdAT ' APNG terms: Same as IDAT except it has an additional 4 byte prefix indicating sequence number
            If zChunk_fdAT(Stream(), ChunkLen, ptrArray + 4&, crc32Value) Then Exit Function
            With m_APNGdat.Frames(m_APNGdat.Frame \ &H10000 - 1&)
                ReDim Preserve .fDat(0 To .DatCount)
                .fDat(.DatCount).length = ChunkLen - 4&
                .fDat(.DatCount).Pointer = ptrArray + 12&
                .DatCount = .DatCount + 1&
                curDataSize = curDataSize + ChunkLen + 12&
            End With
            
        Case chnk_acTL ' APNG terms: animated control chunk -- if present, then an APNG format & must exist before IDAT
            ' process this chunk. Contains frame count
            ' we won't be caching this for GDI+
            If zChunk_acTL(Stream(), ChunkLen, ptrArray + 4&, crc32Value) Then Exit Function
            If FileHandle Then
                pvParsePng = m_APNGdat.FrameCount
                Erase m_APNGdat.Frames()
                Exit Function
            End If
            
        Case chnk_fcTL ' APNG terms: frame control chunk -- one per frame & preceedes series of fdAT chunks, must not preceed acTL
            ' notes: if IDAT is 1st frame, no fdAT will exist btwn this fcTL and next fcTL
            '        if IDAT is not 1st frame, this fcTL occurred after IDAT
            '        fdAT chunks can be separated by other ancillary chunks
            ' we won't be caching this for GDI+
            If zChunk_fcTL(Stream(), ChunkLen, ptrArray + 4&, crc32Value) Then Exit Function
            If curDataSize > maxDataSize Then maxDataSize = curDataSize
            curDataSize = 0&
                
        Case chnk_IHDR ' APNG terms: The overall description of the animated PNG
            ' process this to get overall width/height and other pertinent info
            If zChunk_IHDR(Stream(), ChunkLen, ptrArray + 4&, 0&, crc32Value) Then Exit Function
            
        Case chnk_IEND ' APNG/PNG: the end of the stream. If additonal bytes follow, they are not used
            If ParsableFrame Then
                pvParsePng = ptrArray
                Exit Function
            End If
            Exit Do
            
        Case Else ' other PNG chunks
            ' if they appear before last IDAT they'll be included in the IStream we'll create
            ' otherwise they will be truncated. Any chunks appearing after last IDAT and before IEND
            ' are ancillary and not needed for proper rendering of PNG (generally, they are text-type chunks)
            ' fdAT & all but 1 max fcTL chunks will appear between last IDAT and IEND, so they will be stripped out too
            If ParsableFrame > 1& And m_APNGdat.Offset > 0& Then Exit Function
        
        End Select
        ptrArray = ptrArray + ChunkLen + 12& ' move to next position in the array
    Loop
    
    ' finish building frame info
    If curDataSize > maxDataSize Then maxDataSize = curDataSize
    ReDim m_BlankCanvas(0 To maxDataSize + m_APNGdat.Offset + 11&)        ' size to account for largest frame
    CopyMemory m_BlankCanvas(0), Stream(0), m_APNGdat.Offset              ' copy the PNG header info
    Set m_IStream = modCommon.IStreamFromArray(VarPtr(m_BlankCanvas(0)), UBound(m_BlankCanvas) + 1&)
    Erase m_BlankCanvas()
    If Not m_IStream Is Nothing Then
        ' The fdAT chunk is identical to an IDAT chunk except that it contains an 4-byte prefix
        '   which is the APNG sequence number. But the CRC32 value is for fdAT not IDAT & includes that 4-byte prefix.
        ' Well when we convert fdAT to IDAT in pvRenderFrame, we need to ensure a proper CRC32 value.
        ' This loop creates the CRC32 values once vs each time the frame is displayed
        If m_APNGdat.Frames(0).fDat(0).CRC32 = 0& Then ptrArray = 0& Else ptrArray = 1&
        ' ^^ determine if we skip 1st frame which we will if 1st frame is the IDAT vs fdAT type
        For ptrArray = ptrArray To m_APNGdat.FrameCount - 1&
            If (m_APNGdat.Frames(ptrArray).Disposal And &HFF&) = 2& Then
                With m_APNGdat.Frames(ptrArray).Size
                    lBounds = .nHeight
                    If .nTop < 0& Then lBounds = lBounds + .nTop
                    If lBounds > m_APNGdat.Size.Y Then lBounds = m_APNGdat.Size.Y
                    If lBounds > m_APNGdat.Code2Size Then m_APNGdat.Code2Size = lBounds
                End With
            End If
            For curDataSize = 0& To m_APNGdat.Frames(ptrArray).DatCount - 1&
                With m_APNGdat.Frames(ptrArray).fDat(curDataSize)
                    CopyMemory ChunkName, Stream(.Pointer - 4&), 4&                 ' backup these 4 bytes
                    CopyMemory Stream(.Pointer - 4&), chnk_IDAT, 4&                 ' make this fdAT an IDAT
                    .CRC32 = pvCRCArray(Stream(), .Pointer - 4&, .length + 4&, 0&)  ' CRC the chunk name + frame chunks
                    CopyMemory Stream(.Pointer - 4&), ChunkName, 4&                 ' replace original 4 bytes
                End With
            Next
        Next
        pvParsePng = m_APNGdat.FrameCount                                             ' return frame count
    End If
    
ExitRoutine:
End Function

Private Function pvCRCArray(theArray() As Byte, StartPos As Long, CRCLen As Long, curCRCValue As Long) As Long

    ' Standard CRC algorithm, converting result to newtwork byte order
    ' Returns -1 failure or 0 success, or the CRC value depending on parameters...
    
    ' This has been specifically modified to replicate zLIB's CRC32 function
    ' Params:
    '   theArray(): any initialized byte array
    '   StartPos: the 1st byte position to begin CRC
    '   CRCLen: how many bytes to CRC
    '   curCRCValue: If provided, compare this to result for validation
    '                If not provided, return the CRC value
    
    Dim I As Long, J As Long, crc32val As Long, iLookup As Long
    
    crc32val = &HFFFFFFFF 'start with -1&
    For I = StartPos To StartPos + CRCLen - 1&
        iLookup = (crc32val And &HFF) Xor theArray(I)
        crc32val = (((crc32val And &HFFFFFF00) \ &H100&) And &HFFFFFF) Xor CRC32LUT(iLookup)
    Next I
    pvCRCArray = modCommon.ReverseLong(Not (crc32val)) ' change result to newtwork byte order
    If curCRCValue Then pvCRCArray = (pvCRCArray = curCRCValue) 'return either 0 or -1

End Function

Private Function zChunk_acTL(Stream() As Byte, bufLen As Long, StreamOffset As Long, crcValue As Long) As Long
'    0   num_frames     (unsigned int)    Number of frames
'    4   num_plays      (unsigned int)    Number of times to loop this APNG.  0 indicates infinite looping.
    On Error Resume Next
    Dim lRtn As Long, lValue As Long
    
    If m_APNGdat.FrameCount Then
        lRtn = 1& ' this chunk already processed; there cannot be more than one of these
    Else
        lRtn = Not (pvCRCArray(Stream(), StreamOffset, bufLen + 4&, crcValue))
        If lRtn = 0& Then
            CopyMemory m_APNGdat.FrameCount, Stream(StreamOffset + 4&), 4&
            m_APNGdat.FrameCount = modCommon.ReverseLong(m_APNGdat.FrameCount)
            If m_APNGdat.FrameCount < 2& Then ' must =>1 per specifications else invalid aPNG
                lRtn = 1&
            ElseIf m_APNGdat.FrameCount > &HFFF& Then ' way too big; gotta be something wrong
                lRtn = 1&
            Else
                ReDim Preserve m_APNGdat.Frames(0 To m_APNGdat.FrameCount - 1&)
                CopyMemory m_APNGdat.Loops, Stream(StreamOffset + 8&), 4&
                m_APNGdat.Loops = modCommon.ReverseLong(m_APNGdat.Loops)
                If m_APNGdat.Loops < 0& Then m_APNGdat.Loops = 0& ' zero=indefinite looping else exact loop count
            End If
        End If
    End If
    zChunk_acTL = lRtn
End Function

Private Function zChunk_fcTL(Stream() As Byte, bufLen As Long, StreamOffset As Long, crcValue As Long) As Long
'    0    sequence_number       (unsigned int)   Sequence number of the animation chunk, starting from 0
'    4    width                 (unsigned int)   Width of the following frame
'    8    height                (unsigned int)   Height of the following frame
'   12    x_offset              (unsigned int)   X position at which to render the following frame
'   16    y_offset              (unsigned int)   Y position at which to render the following frame
'   20    delay_num             (unsigned short) Frame delay fraction numerator
'   22    delay_den             (unsigned short) Frame delay fraction denominator
'   24    dispose_op            (byte)           Type of frame area disposal to be done after rendering this frame
'   25    blend_op              (byte)           Type of frame area rendering for this frame
    On Error Resume Next
    Dim lRtn As Long, lValue As Long, Sequence As Long
    
    ' per APNG specs. the acTL chunk (tells us how many frames) is not guaranteed to precede
    ' the first fcTL chunk. This is handled here
    
    If m_APNGdat.FrameCount = 0& Then ' acTL chunk not processed yet
        ReDim m_APNGdat.Frames(0 To 0)
    End If
    lRtn = Not (pvCRCArray(Stream(), StreamOffset, bufLen + 4&, crcValue))
    If lRtn = 0& Then
        CopyMemory Sequence, Stream(StreamOffset + 4&), 4&
        Sequence = modCommon.ReverseLong(Sequence)
        If Sequence = (m_APNGdat.Frame And &HFFFF&) Then
            m_APNGdat.Frame = (m_APNGdat.Frame \ &H10000 + 1&) * &H10000 Or Sequence + 1& ' set next expected sequence & next frame
            With m_APNGdat.Frames(m_APNGdat.Frame \ &H10000 - 1&)
                CopyMemory .Size.nWidth, Stream(StreamOffset + 8&), 4&
                .Size.nWidth = modCommon.ReverseLong(.Size.nWidth)
                CopyMemory .Size.nLeft, Stream(StreamOffset + 16&), 4&
                .Size.nLeft = modCommon.ReverseLong(.Size.nLeft)
                If .Size.nLeft + .Size.nWidth > m_APNGdat.Size.X Or .Size.nWidth < 1& Or .Size.nLeft < 0& Then
                    lRtn = 1&
                Else
                    CopyMemory .Size.nHeight, Stream(StreamOffset + 12&), 4&
                    .Size.nHeight = modCommon.ReverseLong(.Size.nHeight)
                    CopyMemory .Size.nTop, Stream(StreamOffset + 20&), 4&
                    .Size.nTop = modCommon.ReverseLong(.Size.nTop)
                    If .Size.nTop + .Size.nHeight > m_APNGdat.Size.Y Or .Size.nHeight < 1& Or .Size.nTop < 0& Then
                        lRtn = 1&
                    Else
                        CopyMemory .Delay, Stream(StreamOffset + 24&), 4&
                        .Delay = modCommon.ReverseLong(.Delay)
                        If (.Delay And &HFFFF&) = 0& Then
                            .Delay = (.Delay \ &H10000) / 100& * 1000&
                        Else
                            .Delay = (.Delay \ &H10000) / (.Delay And &HFFFF&) * 1000&
                        End If
                        .Disposal = Stream(StreamOffset + 29&) * &H100 Or Stream(StreamOffset + 28&)
                        ' disposal codes
                        '0 = APNG_DISPOSE_OP_NONE: no disposal is done on this frame before rendering the next; the contents of the output buffer are left as is.
                        '1 = APNG_DISPOSE_OP_BACKGROUND: the Frame 's region of the output buffer is to be cleared to fully transparent black before rendering the next frame.
                        '2 = APNG_DISPOSE_OP_PREVIOUS: the Frame 's region of the output buffer is to be reverted to the previous contents before rendering the next frame.
                        ' ROP codes
                        '0   APNG_BLEND_OP_SOURCE ... completely overwrites previous frame
                        '1   APNG_BLEND_OP_OVER ... blended onto previous frame
                    End If
                End If
            End With
        Else
            lRtn = 1&   ' per APNG specs, if out-of-order sequence, treat as error
        End If
    End If
    zChunk_fcTL = lRtn

End Function

Private Function zChunk_fdAT(Stream() As Byte, bufLen As Long, StreamOffset As Long, crcValue As Long) As Long

'    0    sequence_number       (unsigned int)   Sequence number of the animation chunk, starting from 0
'    4    frame_data            X bytes          Frame data for this frame
    
    Dim lRtn As Long, lValue As Long, Sequence As Long
    On Error Resume Next
    
    If m_APNGdat.FrameCount = 0& Then
        lRtn = 1&                   ' no acTL chunk processed; this chunk is invalid
    Else
        lRtn = Not (pvCRCArray(Stream(), StreamOffset, bufLen + 4&, crcValue))
        If lRtn = 0& Then
            CopyMemory Sequence, Stream(StreamOffset + 4&), 4&
            Sequence = modCommon.ReverseLong(Sequence)
            If Sequence = (m_APNGdat.Frame And &HFFFF&) Then
                m_APNGdat.Frame = (m_APNGdat.Frame And &HFFFF0000) Or Sequence + 1& ' set next expected sequence
            Else
                lRtn = 1&   ' per APNG specs, if out-of-order sequence, treat as error
            End If
        End If
    End If
    zChunk_fdAT = lRtn

End Function


Private Function zChunk_IHDR(Stream() As Byte, bufLen As Long, StreamOffset As Long, cmprSize As Long, crcValue As Long) As Long
                
    ' IHDR structure
    '    Width As Long              << cannot be negative
    '    Height As Long             << cannot be negative
    '    BitDepth As Byte           << must be 1,2,4,8,16
    '    ColorType As Byte          << must be 0,2,3,4,6
    '    Compression As Byte        << must be zero
    '    Filter As Byte             << must be zero
    '    Interlacing As Byte        << must be zero or one
    
    On Error Resume Next
    Dim lRtn As Long, lValue As Long
    
    lRtn = Not (pvCRCArray(Stream(), StreamOffset, bufLen + 4, crcValue))
    If lRtn = 0& Then
        
        CopyMemory m_APNGdat.Size.X, Stream(StreamOffset + 4&), 4&
        m_APNGdat.Size.X = modCommon.ReverseLong(m_APNGdat.Size.X)
        CopyMemory m_APNGdat.Size.Y, Stream(StreamOffset + 8&), 4&
        m_APNGdat.Size.Y = modCommon.ReverseLong(m_APNGdat.Size.Y)
        
        If m_APNGdat.Size.X < 1& Or m_APNGdat.Size.Y < 1& Then
            lRtn = 1& 'Corrupted Image Header. Cannot continue.
        Else
            If Not Stream(StreamOffset + 14&) = 0 Then
                lRtn = 1& ' Invalid Compression Flag in Header. Cannot continue.
            Else
                If Not Stream(StreamOffset + 15&) = 0 Then
                    lRtn = 1& 'Invalid Filter Flag in Header. Cannot continue.
                Else
                    Select Case Stream(StreamOffset + 12&)
                    Case 1&, 2&, 4&, 8&, 16&
                        ' it is a valid bit depth
                        Select Case Stream(StreamOffset + 13&)
                        Case 0&, 2&, 3&, 4&, 6&
                            ' it is a valid color type
                            If Stream(StreamOffset + 16&) > 1& Then
                                lRtn = 1& 'Invalid Interlacing Flag in Header. Cannot continue.
                            End If
                        Case Else
                            lRtn = 1& 'Invalid Color Type Flag in Header. Cannot continue.
                        End Select
                    Case Else
                        lRtn = 1& 'Invalid Bit Depth Flag in Header. Cannot continue.
                    End Select
                    
                End If  ' Filter flag
            End If  ' Compression flag
        End If  ' Dimensions
        
        If lRtn = 0& Then
            ' check for png sizes that would cause overflow errors in other calculations...
            ' This has 2 basic checks
            ' check DWord width alignment * height first are within bounds
            lValue = 32& * m_APNGdat.Size.X * m_APNGdat.Size.Y ' max number of bytes needed for DIB
            If Err Then
                Err.Clear
                lRtn = 1&
            End If
            
        End If
    End If

    zChunk_IHDR = lRtn

End Function

Private Function pvGetStreamAddress(LockHandle As Long) As Long

    ' function returns a handle to a Locked stream object
    '   and also releases that object
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If LockHandle Then
        GlobalUnlock LockHandle
    ElseIf Not m_IStream Is Nothing Then
        If GetHGlobalFromStream(ObjPtr(m_IStream), LockHandle) = 0 Then
            pvGetStreamAddress = GlobalLock(LockHandle)
        End If
    End If

End Function

Private Sub pvRenderFrame(FrameNr As Long, sourceData() As Byte, Handle As Long)

    ' where all the magic happens ;)
    ' FYI: FrameNr in this routine is one-bound, not zero-bound

    Dim gHandle As Long, lPtr As Long, hLock As Long, d As Long
    Dim bmpDat As BitmapData, bmpDatSrc As BitmapData, bmpErase As BitmapData
    Dim boundsSrc As RECTI, boundsDst As RECTI, boundsErase As RECTI
    Dim tSA As SafeArray, aFrame() As Byte, hGraphics As Long
    
    On Error GoTo ExitRoutine
    ' the Handle passed herein was created from the stream maintained in this class
    ' We will update the stream vs. re-creating it; faster overall
    With tSA                ' use a DMA overlay on the stream handle
        .cbElements = 1
        .cDims = 1
        .pvData = pvGetStreamAddress(hLock)
        If .pvData = 0& Then Exit Sub
        .rgSABound(0).cElements = GlobalSize(.pvData)
    End With                ' set the DMA handle
    CopyMemory ByVal VarPtrArray(aFrame), VarPtr(tSA), 4&
    
    With m_APNGdat.Frames(FrameNr)                                    ' update the stream
        lPtr = m_APNGdat.Offset
        CopyMemory aFrame(16), modCommon.ReverseLong(.Size.nWidth), 4&      ' make png bounds same size as frame
        CopyMemory aFrame(20), modCommon.ReverseLong(.Size.nHeight), 4&
        For d = 0& To .DatCount - 1&
            CopyMemory aFrame(lPtr), modCommon.ReverseLong(.fDat(d).length), 4& ' insert chunk length
            CopyMemory aFrame(lPtr + 4), chnk_IDAT, 4&                  ' and chunk name as IDAT
            CopyMemory aFrame(lPtr + 8&), sourceData(.fDat(d).Pointer), .fDat(d).length
            lPtr = lPtr + .fDat(d).length + 12&                         ' ^ copy chunk
            CopyMemory aFrame(lPtr - 4&), .fDat(d).CRC32, 4&            ' append CRC32 value
        Next
    End With
    CopyMemory aFrame(lPtr), 0&, 4&                                     ' append IEND
    CopyMemory aFrame(lPtr + 4&), chnk_IEND, 4&                         ' and done with this frame
    CopyMemory aFrame(lPtr + 8&), &H826042AE, 4&                        ' CRC value for IEND
        
    CopyMemory ByVal VarPtrArray(aFrame), 0&, 4&: tSA.pvData = 0&       ' done; remove overlay &
    pvGetStreamAddress hLock                                            ' unlock pointer
    If GdipLoadImageFromStream(ObjPtr(m_IStream), gHandle) Then Exit Sub    ' abort on error
    
    ' post-frame actions per previous frame's disposal code
    If FrameNr = 0& Then    ' first frame, completely erase the image
        boundsErase.nHeight = m_APNGdat.Size.Y: boundsErase.nWidth = m_APNGdat.Size.X
        If GdipGetImageGraphicsContext(Handle, hGraphics) = 0& Then GdipGraphicsClear hGraphics, 0&
    
    Else
        Select Case (m_APNGdat.Frames(m_APNGdat.Frame).Disposal And &HFF&)
        Case 0  ' do nothing, drawing remains unchanged
        Case 1  ' blank out previous frame's area. Note: Same as GIF's Code2
            boundsErase = m_APNGdat.Frames(m_APNGdat.Frame).Size
            If GdipGetImageGraphicsContext(Handle, hGraphics) = 0& Then
                GdipSetClipRectI hGraphics, boundsErase.nLeft, boundsErase.nTop, boundsErase.nWidth, boundsErase.nHeight, 0&
                GdipGraphicsClear hGraphics, 0&
            End If
        Case 2      ' copy previously cached scanlines.  Note: Same as GIF's Code3
            boundsErase = m_APNGdat.Frames(m_APNGdat.Frame).Size
            boundsErase.nLeft = 0&: boundsErase.nWidth = m_APNGdat.Size.X
            If boundsErase.nHeight > m_APNGdat.Code2Size Then boundsErase.nHeight = m_APNGdat.Code2Size
            bmpErase.Scan0Ptr = VarPtr(m_BlankCanvas(0))
            bmpErase.stride = m_APNGdat.Size.X * 4& ' set scanwidth & update
            ' ^^ Tip: Unless retrieving just 1 row starting at left column, entire row will always be set/returned
            If GdipBitmapLockBits(Handle, boundsErase, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor32bppAlpha, bmpErase) = 0& Then
                GdipBitmapUnlockBits Handle, bmpErase
            End If                                  ' ensure our blank canvas is blank
        End Select
    End If
    
    ' set the area of the next frame
    boundsDst = m_APNGdat.Frames(FrameNr).Size                ' good, let's set some bounds
    boundsSrc = boundsDst
    
    If (m_APNGdat.Frames(FrameNr).Disposal And &HFF&) = 2 Then
        ' with dispoal code of 3, need to cache current canvas & replace when next frame is rendered
        If hGraphics Then GdipDeleteGraphics hGraphics: hGraphics = 0&
        boundsErase = boundsDst
        boundsErase.nLeft = 0&: boundsErase.nWidth = m_APNGdat.Size.X
        If boundsErase.nHeight > m_APNGdat.Code2Size Then boundsErase.nHeight = m_APNGdat.Code2Size
        If GdipBitmapLockBits(Handle, boundsErase, ImageLockModeRead, lvicColor32bppAlpha, bmpDatSrc) = 0& Then
            CopyMemory m_BlankCanvas(0), ByVal bmpDatSrc.Scan0Ptr, bmpDatSrc.Height * bmpDatSrc.stride
            ' ^^ Tip: Unless retrieving just 1 row starting at left column, entire row will always be set/returned
            GdipBitmapUnlockBits Handle, bmpDatSrc
        End If
    End If
        
    ' some efficiency tweaks to prevent rendering PNG if possible
    If FrameNr = 0& Then                                        ' never render; just slide new image to our PNG handle
        lPtr = 0&
    ElseIf (m_APNGdat.Frames(FrameNr).Disposal \ &H100) = 0 Then   ' no rendering needed if no blending used
        lPtr = 0&
    ElseIf (m_APNGdat.Frames(m_APNGdat.Frame).Disposal And &HFF&) = 1 Then ' if the previous frame's action was to erase, then
        If boundsDst.nTop >= boundsErase.nTop Then              ' if this frame is <= previous frame's dimensions
            If boundsDst.nLeft >= boundsErase.nLeft Then        ' then we don't need to render
                If boundsErase.nWidth + boundsErase.nLeft >= boundsDst.nLeft + boundsDst.nWidth Then
                    If boundsErase.nHeight + boundsErase.nTop >= boundsDst.nTop + boundsDst.nHeight Then lPtr = 0&
                End If
            End If
        End If
    End If
    If lPtr = 0& Then                                           ' update PNG with current frame
        boundsSrc.nTop = 0&: boundsSrc.nLeft = 0&
        If hGraphics Then GdipDeleteGraphics hGraphics: hGraphics = 0&
        If GdipBitmapLockBits(gHandle, boundsSrc, ImageLockModeRead, lvicColor32bppAlpha, bmpDatSrc) = 0& Then
            bmpErase = bmpDatSrc
            If GdipBitmapLockBits(Handle, boundsDst, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor32bppAlpha, bmpErase) = 0& Then
                GdipBitmapUnlockBits Handle, bmpErase            ' update the PNG now
            End If
            GdipBitmapUnlockBits gHandle, bmpDatSrc
        End If
        
    Else                                                        ' rendering to DC & will update PNG from rendered section
        If hGraphics = 0& Then Call GdipGetImageGraphicsContext(Handle, hGraphics) Else GdipResetClip hGraphics
        If hGraphics Then
            With boundsDst                                      ' render source first, then new frame
                GdipDrawImageRectRectI hGraphics, gHandle, .nLeft, .nTop, .nWidth, .nHeight, 0&, 0&, .nWidth, .nHeight, UnitPixel, 0&, 0&, 0&
            End With
        End If
    End If

ExitRoutine:                                                        ' and clean up
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(aFrame), 0&, 4&
    If hGraphics Then GdipDeleteGraphics hGraphics              ' clean up
    If gHandle Then GdipDisposeImage gHandle
    m_APNGdat.Frame = FrameNr                                       ' track current frame
End Sub

Public Property Get LoopCount() As Long
    LoopCount = m_APNGdat.Loops
End Property

Public Function MoveToFrame(ByVal FrameNr As Long, ByRef sourceInfo As Variant, ByVal Handle As Long) As Boolean

    ' called to move a frame forward or backward

    Dim bSkippedFrames As Boolean
    Dim sourceData() As Byte
    
    If FrameNr = m_APNGdat.Frame Then
        Exit Function
    ElseIf FrameNr >= m_APNGdat.FrameCount Then
        FrameNr = 0&
    ElseIf FrameNr < 0& Then
        FrameNr = m_APNGdat.FrameCount - 1&
        bSkippedFrames = True
    ElseIf FrameNr < m_APNGdat.Frame Then
        bSkippedFrames = (FrameNr > 0&)
    ElseIf FrameNr > m_APNGdat.Frame + 2& Then
        bSkippedFrames = True
    End If
    
    modCommon.MoveArrayToVariant sourceInfo, sourceData(), False
    If bSkippedFrames Then
        For FrameNr = 0& To FrameNr - 1&
            pvRenderFrame FrameNr, sourceData(), Handle
        Next
    End If
    pvRenderFrame FrameNr, sourceData(), Handle
    modCommon.MoveArrayToVariant sourceInfo, sourceData(), True
    MoveToFrame = True

End Function

Public Function ConvertGIF2APNG(returnObject As Variant, sourceData() As Byte) As Boolean

    ' function takes a parsed animated GIF and converts it to an animated PNG

    Dim cGIF As cFunctionsGIF, MIS As MULTIIMAGESAVESTRUCT
    Dim tData() As Byte, f As Long
    Dim tSource As Variant, tReturn As Variant
    
    Set cGIF = New cFunctionsGIF
    If cGIF.IsGIFResource(g_NewImageData, sourceData(), False, True) = False Then Exit Function
    
    modCommon.MoveArrayToVariant tSource, sourceData(), True
    ' call conversion for first frame which also fills in the MIS structure
    If cGIF.ConvertFrameToPNG(tSource, tReturn, 0&, MIS) Then
        If SaveAsAPNG(tReturn, MIS, 1&) Then    ' convert frame to APNG format
            For f = 1& To MIS.Images - 1&       ' continue with subsequent frames
                If cGIF.ConvertFrameToPNG(tSource, tReturn, f, MIS) Then
                    If SaveAsAPNG(tReturn, MIS, f + 1&) = 0& Then Exit For
                End If
            Next
            If f = MIS.Images Then      ' all frames processed
                modCommon.MoveArrayToVariant tReturn, tData(), False        ' get APNG as byte array
                modCommon.MoveArrayToVariant returnObject, tData(), True    ' place it in return object
                ConvertGIF2APNG = True
            End If
        End If
    End If
    modCommon.MoveArrayToVariant tSource, sourceData(), False               ' return orginal source
    Set cGIF = Nothing

End Function

Public Function SaveAsAPNG(returnObject As Variant, MIS As MULTIIMAGESAVESTRUCT, FrameNumber As Long) As Long

    ' function converts a MULTIIMAGESAVESTRUCT into an animated PNG as an array of bytes
    ' FrameNumber passed to this routine is 1-bound only
    
    ' APNG has some specific instructions.
    '   - All frames must use the same compression/filter methods, bit depth, color type and palette (if used)
    '   - We do not create the PNG ourselves, we use GDI+. We have no way of directing how GDI+ creates a PNG 'frame'
    '       -- Filter method, 6 are available per PNG specifications.
    '           :: I disagree that it must be same for each frame as the filter method is encoded per scan line
    '               of image and per specs, each scan line can use a different filter
    '       -- Color Type & Bit Depth. Since each frame's type/depth cannot be individually set, it must be same for all
    '       -- Palette. Since any PNG image uses only 1 palette, every frame must use same palette, if palette is used
    '           :: Also, there is no way to guarantee if same palette is used, that GDI+ won't optimize/re-order it on a case-by-case basis
    '   - With above said, how do we guarantee we can meet these restrictions without writing the PNG manually?
    '       - Palette. Not an issue if frames are 24/32bpp
    '       - Color Type & Bit Depth. Not an issue if all frames forced to 32bpp or 24bpp
    '       - These routines create only 32bpp APNGs & rely on GDI+ compression to minimize file size
    '       - The two functions that create the APNG frame source canvas are:
    '           :: modCommon.pvValidateMultiImgStruct
    '           :: cFunctionsGIF.ConvertFrameToPNG

    Dim tData() As Byte
    Dim lResult As Long
    Dim lValue As Long, f As Long, lbOffset As Long
    
    If Not (Not CRC32LUT()) = 0& Then
        Const CRCpolynomial As Long = &HEDB88320
        ReDim CRC32LUT(0 To 255)                            ' create a CRC lookup table (LUT)
        For f = 0& To 255&
            lValue = f
            For lResult = 8& To 1& Step -1&
                If (lValue And 1&) Then
                    lValue = (((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF) Xor CRCpolynomial
                Else
                    lValue = ((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                End If
            Next lResult
            CRC32LUT(f) = lValue
        Next f
    End If
    Debug.Assert App.hInstance  ' hack to prevent IDE calculation errors associated with NOT NOT call
    
    lbOffset = LBound(MIS.Image) - 1&
    modCommon.MoveArrayToVariant returnObject, tData(), False   ' get source PNG data
    
    If FrameNumber = 1& Then    ' handle 1st frame differently
        
        ReDim m_APNGdat.Frames(0 To 0)                          ' initialize structure (module-level)
        lResult = pvParsePng(tData(), , FrameNumber)            ' find IDAT & IEND locations & validate PNG
        If lResult Then
            lResult = lResult - m_APNGdat.Frames(0).fDat(0).Pointer ' size of PNG data we need to copy
            If (MIS.Image(lbOffset + 1&).GIFFrameInfo.DisposalCode And lvicAPNG_DefaultImage_NotAnimated) = 0& Then
                ' this frame is part of the animation
                lValue = 58& ' 26 byte fcTL chunk + 8 byte acTL chunk + 24 bytes for chunk headers
            Else ' this frame is not part of the animation
                lValue = 20& ' 8 byte acTL chunk + 12 byte header
            End If
            ReDim m_BlankCanvas(0 To UBound(tData) + lValue)
            ' transfer up to the 1st IDAT chunk
            CopyMemory m_BlankCanvas(0), tData(0), m_APNGdat.Frames(0).fDat(0).Pointer
            ' move pointer  20 or 58 bytes further & transfer rest of png data
            CopyMemory m_BlankCanvas(m_APNGdat.Frames(0).fDat(0).Pointer + lValue), tData(m_APNGdat.Frames(0).fDat(0).Pointer), lResult
            
            If lValue = 58& Then    ' if frame is part of animation, include the fcTL chunk
                Call pvWrite_fcTL(lbOffset + 1&, 0, m_APNGdat.Frames(0).fDat(0).Pointer + 20&, MIS)
                m_APNGdat.Code2Size = 1&
            End If
            lValue = m_APNGdat.Frames(0).fDat(0).Pointer ' where we will add the acTL chunk to
            ' build the acTL chunk
            CopyMemory m_BlankCanvas(lValue), modCommon.ReverseLong(8&), 4&        ' chunk length
            CopyMemory m_BlankCanvas(lValue + 4&), chnk_acTL, 4&                     ' chunk name
            CopyMemory m_BlankCanvas(lValue + 8&), modCommon.ReverseLong(MIS.Images), 4& ' number of frames
            CopyMemory m_BlankCanvas(lValue + 12&), modCommon.ReverseLong(MIS.GIFOverview.LoopCount), 4& ' loop count
            CopyMemory m_BlankCanvas(lValue + 16&), pvCRCArray(m_BlankCanvas(), lValue + 4&, 12&, 0&), 4& ' chunk CRC32 value
            
            SaveAsAPNG = True
        End If
        
    Else ' frames #2-#n
    
        m_APNGdat.Frames(0).DatCount = 0&                   ' reset for next frame
        If pvParsePng(tData(), , FrameNumber) = 0& Then Exit Function   ' find IDAT locations & validate PNG
        lValue = UBound(m_BlankCanvas) - 11&: lResult = 49& ' 49=12+38-1 space for IEND & fcTL chunks
        With m_APNGdat.Frames(0)
            For f = 0& To .DatCount - 1&                    ' tally IDAT sizes
                lResult = lResult + .fDat(f).length + 16&   ' add 12 byte header + 4 byte fDAT sequence number
            Next
        End With
        ReDim Preserve m_BlankCanvas(0 To lValue + lResult) ' resize overall array
        Call pvWrite_fcTL(lbOffset + FrameNumber, m_APNGdat.Code2Size, lValue, MIS) ' write the fcTL chunk
        lValue = lValue + 38&                               ' move pointer along
        m_APNGdat.Code2Size = m_APNGdat.Code2Size + 1&      ' increment chunk sequence
        With m_APNGdat.Frames(0)
            For f = 0 To .DatCount - 1&
                CopyMemory m_BlankCanvas(lValue), modCommon.ReverseLong(.fDat(f).length + 4&), 4&   ' chunk length+sequence
                CopyMemory m_BlankCanvas(lValue + 4), chnk_fdAT, 4&                                 ' chunk name
                CopyMemory m_BlankCanvas(lValue + 8), modCommon.ReverseLong(m_APNGdat.Code2Size), 4& ' sequence number
                CopyMemory m_BlankCanvas(lValue + 12), tData(.fDat(f).Pointer + 8&), .fDat(f).length  ' copy image data
                Erase tData()                                                                       ' write the crc value
                CopyMemory m_BlankCanvas(lValue + .fDat(f).length + 12&), pvCRCArray(m_BlankCanvas(), lValue + 4&, .fDat(f).length + 8&, 0&), 4& ' chunk CRC32 value
                lValue = lValue + .fDat(f).length + 16&                                             ' move pointer along
                m_APNGdat.Code2Size = m_APNGdat.Code2Size + 1&                                      ' increment sequence
            Next
        End With
        Set MIS.Image(lbOffset + FrameNumber).Picture = Nothing
        SaveAsAPNG = True
    End If
    
    If FrameNumber = MIS.Images Then
        Erase CRC32LUT()
        CopyMemory m_BlankCanvas(lValue), 0&, 4&                                    ' append IEND
        CopyMemory m_BlankCanvas(lValue + 4&), chnk_IEND, 4&
        CopyMemory m_BlankCanvas(lValue + 8&), &H826042AE, 4&                       ' CRC value for IEND
        modCommon.MoveArrayToVariant returnObject, m_BlankCanvas(), True            ' move result to return object
    End If

End Function

Private Sub pvWrite_fcTL(FrameNumber As Long, SeqNumber As Long, Offset As Long, MIS As MULTIIMAGESAVESTRUCT)

    ' helper function for creating animated PNG

    ' build the fcTL chunk, 26 bytes + 12 byte header/footer
    CopyMemory m_BlankCanvas(Offset), modCommon.ReverseLong(26&), 4&                ' chunk length
    CopyMemory m_BlankCanvas(Offset + 4&), chnk_fcTL, 4&                            ' chunk name
    CopyMemory m_BlankCanvas(Offset + 8&), modCommon.ReverseLong(SeqNumber), 4&     ' sequence number
    CopyMemory m_BlankCanvas(Offset + 12&), modCommon.ReverseLong(m_APNGdat.Size.X), 4& ' frame width
    CopyMemory m_BlankCanvas(Offset + 16&), modCommon.ReverseLong(m_APNGdat.Size.Y), 4& ' frame height
    
    With MIS.Image(FrameNumber).GIFFrameInfo                                        ' frame offsets & delay
        CopyMemory m_BlankCanvas(Offset + 20&), modCommon.ReverseLong(.XOffset), 4& ' X offset
        CopyMemory m_BlankCanvas(Offset + 24&), modCommon.ReverseLong(.YOffset), 4& ' Y offset
        CopyMemory m_BlankCanvas(Offset + 28&), modCommon.ReverseLong(.DelayTime * &H10000 Or &H3E8&), 4&
        Select Case (.DisposalCode And &H3&)                                        ' frame disposal
            Case lvicGIF_RestorePrevious: m_BlankCanvas(Offset + 32&) = 2
            Case lvicGIF_Erase: m_BlankCanvas(Offset + 32&) = 1
            Case Else: m_BlankCanvas(Offset + 32&) = 0
        End Select                                                                  ' frame blendop code
        If (.DisposalCode And lvicAPNG_BlendOpOver) Then m_BlankCanvas(Offset + 33&) = 1 Else m_BlankCanvas(Offset + 33&) = 0
    End With
    CopyMemory m_BlankCanvas(Offset + 34&), pvCRCArray(m_BlankCanvas(), Offset + 4&, 30&, 0&), 4& ' chunk CRC32 value

End Sub

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
End Sub

Private Sub Class_Terminate()
    ' clean up
    Set m_IStream = Nothing
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub



