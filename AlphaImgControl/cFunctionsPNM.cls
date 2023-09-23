VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsPNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write PNM & PAM image format & Convert to Bitmap ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' PNM/PAM Format specifications retrieved from following
' PNM is a generalized extension for: PBM, PGM, PPM. PNM is not an official file extension
'   PNM = Portable aNy Map
'   PBM = Portable Bit Map (black & white only)
'       http://ftp.parisc-linux.org/cgi-bin/man/man2html?pbm+5
'   PGM = Portable Gray Map (grayscale only)
'       http://ftp.parisc-linux.org/cgi-bin/man/man2html?5+pgm
'   PPM = Portable Pixel Map (24 bit RGB only)
'       http://ftp.parisc-linux.org/cgi-bin/man/man2html?5+ppm
'   PAM = Portable Arbritray Map (a newer format of PBM/PGM/PPM)
'       http://ftp.parisc-linux.org/cgi-bin/man/man2html?pam+5

' NOTES
' Though it is possible to have PNM files contain multiple images, this class
'   is not designed to handle them & only 1st image will be parsed.
'   Should I find some good examples of multi-image PNM, then I'll include the parsing
' Also note that it is possible for PGM & PPM files to have R,G,B components greater than
'   255, maxing out at 65535. These images are processed as if they were dual planed.
'   Only 1 plane will be processed
' A good page of various images in PBM, PGM, PPM (both binary and ASCII)
'   http://people.sc.fsu.edu/~jburkardt/data/data.html

' PNM binary mode can be slightly smaller than standard bitmaps
'                       Windows Bitmap      PNM (pbm, pgm, ppm)
'   Header Size         40 + color table    9-15 for pgm,ppm & 7-11 for pbm (assume no comments in file)
'   Pixel data          word aligned        byte aligned which may result in 3 bytes less per scan line
'   1 bpp format (2 colors)  supported      only supported if image is black & white else RGB format
'   4 bpp format (16 colors) supported      not supported; RGB format must be used
'   8 bpp format (256 color) supported      only supported if image is grayscale else RGB format

' PNM ASCII can be extremely large files. With RGB for example, a single pixel in binary mode
'   would use 3 bytes, but in ASCII mode can use 12+ bytes (4x larger). Minimally, they will
'   almost always be at least twice the size of binary mode, but expect 3-5x larger
' Even though the binary mode files can be smaller, they take a bit longer to process. This is
'   because they are stored on disk/memory as RGB where bitmaps are BGR and conversion is needed

' PAM files are always read/written as binary

Private m_LUT() As Byte
Private m_Token As cGDIpToken

Public Function IsPNMResource(theStream() As Byte, FileHandle As Long) As Long

    Dim lRead As Long, lValue As Long
    Dim w As Long, h As Long, pMode As Long
    
    ' All PNM files have a magic number of: P1, P2, P3, P4, P5, or P6
    ' P1 thru P3 are ASCII (non-compressed) format
    ' P4 thur P6 are binary (compressed) format
    ' P7 is the PAM format
    If FileHandle Then
        SetFilePointer FileHandle, 0&, 0&, 0&
        ReadFile FileHandle, lValue, 2&, lRead, ByVal 0&
    Else
        lValue = theStream(0) Or theStream(1) * &H100&
    End If
    If (lValue And &HFF&) = 80 Then     ' capital P
        Select Case lValue \ &H100&
        Case 49 To 54                   ' numeric 1 thru 6
            lValue = pvParseHeader(theStream(), w, h, lRead)
            If lValue Then
                On Error Resume Next    ' a quick overflow test should w or h be very large
                w = (w * h * 4&) \ (h * 4&)
                If Err Then
                    Err.Clear
                Else
                    If lRead > 255& Then lRead = 2& Else lRead = 1& ' binary size of 8 or 16 bit image
                    ' see if size of data minimally supports image type
                    Select Case theStream(1)
                    Case 52 ' binary b&w
                        lRead = ((UBound(theStream) - lValue + 1&) >= (w * h) \ 8)
                    Case 53 ' binary grayscale
                        lRead = ((UBound(theStream) - lValue + 1&) >= w * h * lRead)
                    Case 54 ' binary rgb
                        lRead = ((UBound(theStream) - lValue + 1&) >= w * h * 3 * lRead)
                        
                     ' ascii; guesstimate minimal size
                    Case 51 ' RGB
                        lRead = (UBound(theStream) - lValue + 2&) \ 2 >= (w * h * 3&)
                    Case Else ' grayscale & B&W (minimum of 1 byte per pixel)
                        lRead = (UBound(theStream) - lValue + 2&) \ 2 >= w * h
                    End Select
                    If lRead = True Then IsPNMResource = lvicPicTypePNM
                End If
            End If
        Case 55                         ' numeric 7 (PAM format is P7)
            lRead = pvParseHeaderPAM(theStream(), w, h, 0&, lValue, pMode)
            If lRead Then
                Select Case (pMode And &H7F)
                Case 52     ' b&w
                    lRead = ((UBound(theStream) - lValue + 1&) >= (w * h) \ 8)
                Case 53, 54 ' grayscale/rgb (either with/without alpha)
                    lRead = ((UBound(theStream) - lValue + 1&) >= w * h * lValue)
                End Select
                If lRead = True Then IsPNMResource = lvicPicTypePAM
            End If
        Case Else
        End Select
    End If

End Function

Public Function LoadPNMResource(inStream() As Byte) As Long

    ' function converts passed PNM array into a GDI+ bitmap

    Dim I As Long, MaxVal As Long, aPos As Long
    Dim tData() As Byte, pPAMmode As Long
    Dim bmpi As BitmapData, lDepth As Long
    Dim tRect As RECTI, pal As ColorPalette

    ' get the position of the 1st pixel within the file
    If inStream(1) = 55 Then    ' PAM vs. PNM
        aPos = pvParseHeaderPAM(inStream(), bmpi.Width, bmpi.Height, MaxVal, lDepth, pPAMmode)
    Else
        aPos = pvParseHeader(inStream(), bmpi.Width, bmpi.Height, MaxVal)
    End If
    If aPos = 0& Or aPos >= UBound(inStream) Then Exit Function
    
    ' At this point the pixel data will start at aPos
    ' Only numeric characters and white space are allowed til end of pixel data
    ' If any non-numeric, non-white space exists, routines will assume corrupt or non-valid PNM
    
    ' when MaxVal>255 we ignore one plane of the 2 plane/16 bit R,G,B components
    If MaxVal > 0& Then                     ' create lookup table for relative luminance
        ReDim m_LUT(0 To (MaxVal And &HFF))
        For I = 1& To UBound(m_LUT) - 1&
            m_LUT(I) = (I * 255&) \ (MaxVal And &HFF&)
        Next
        m_LUT(I) = 255
    End If
    
    If pPAMmode Then                                ' PAM format
        Select Case pPAMmode
        Case 54     ' simple RGB type
            bmpi.PixelFormat = lvicColor24bpp
        Case 53     ' grayscale
            bmpi.PixelFormat = lvicColor8bpp
        Case 52     ' black and white
            bmpi.PixelFormat = lvicColor1bpp
        Case Else   ' rgb alpha or grayscale alpha (processed as 32bpp ARGB)
            bmpi.PixelFormat = lvicColor32bppAlpha
        End Select
    Else
        Select Case inStream(1)                     ' type of PNM
        Case 49, 52 ' PBM (black & white)
            bmpi.PixelFormat = lvicColor1bpp
        Case 50, 53 ' PGM (grayscale)
            bmpi.PixelFormat = lvicColor8bpp
        Case Else   ' PPM (rgb)
            bmpi.PixelFormat = lvicColor24bpp
        End Select
        pPAMmode = 256&
    End If
    ' determine scan width & size output array
    bmpi.stride = modCommon.ByteAlignOnWord((bmpi.PixelFormat And &HFF00&) \ &H100&, bmpi.Width)
    ReDim tData(0 To bmpi.stride * bmpi.Height - 1&)
    
    ' call method that does the pixel parsing
    Select Case bmpi.PixelFormat
        Case lvicColor32bppAlpha
            bmpi.Scan0Ptr = pvParseTrueColorAlpha(inStream(), tData(), bmpi, MaxVal, pPAMmode, aPos)
        Case lvicColor24bpp
            bmpi.Scan0Ptr = pvParseTrueColor(inStream(), tData(), bmpi.Width, bmpi.Height, MaxVal, aPos)
        Case lvicColor8bpp
            bmpi.Scan0Ptr = pvParseGrayscale(inStream(), tData(), bmpi.Width, bmpi.Height, MaxVal, aPos)
        Case lvicColor1bpp
            bmpi.Scan0Ptr = pvParseBlackWhite(inStream(), tData(), bmpi.Width, bmpi.Height, aPos, (pPAMmode <> 0))
    End Select
    Erase m_LUT()
    
    If bmpi.Scan0Ptr Then                           ' prepare to create GDI+ bitmap
        bmpi.Scan0Ptr = VarPtr(tData(0))
        tRect.nWidth = bmpi.Width
        tRect.nHeight = bmpi.Height
        
        If GdipCreateBitmapFromScan0(bmpi.Width, bmpi.Height, 0&, bmpi.PixelFormat, ByVal 0&, LoadPNMResource) = 0& Then
            If GdipBitmapLockBits(LoadPNMResource, tRect, ImageLockModeWrite Or ImageLockModeUserInputBuf, bmpi.PixelFormat, bmpi) Then
                GdipDisposeImage LoadPNMResource         ' oops; failure
                LoadPNMResource = 0&
            Else
                GdipBitmapUnlockBits LoadPNMResource, bmpi   ' unlock bits
            End If
            
            If LoadPNMResource Then                         ' handle paletted formats
                If bmpi.PixelFormat = lvicColor8bpp Then
                    pal.Count = 256&
                    For I = 0& To 255&
                        pal.Entries(I + 1&) = I Or I * &H100& Or I * &H10000 Or &HFF000000
                    Next
                    GdipSetImagePalette LoadPNMResource, pal
                ElseIf bmpi.PixelFormat = lvicColor1bpp Then
                    pal.Count = 2&
                    pal.Entries(1) = &HFF000000
                    pal.Entries(2) = &HFFFFFFFF
                    GdipSetImagePalette LoadPNMResource, pal
                End If
            End If
        End If
    End If
    
End Function

Private Function pvParseBlackWhite(inStream() As Byte, outStream() As Byte, Width As Long, Height As Long, _
                                    StreamOffset As Long, invertPalette As Boolean) As Boolean

    ' Parse Black & White (PBM) files
    ' Note that Black is 1, not 0. This is due to 1 was meant to be: Ink On vs. full intensity

    Dim totalProcessed As Long
    Dim L As Long, c As Long
    Dim nrRows As Long
    Dim scanWidth As Long
    Dim bLShift As Long
    
    On Error GoTo ExitRoutine
    scanWidth = modCommon.ByteAlignOnWord(1, Width)
    bLShift = 128
    If inStream(1) < 52 Then                ' ASCII vs. binary
        For StreamOffset = StreamOffset To UBound(inStream)
            Select Case inStream(StreamOffset)
            Case 48 To 57 ' numeric values
                If c = 0& Then c = StreamOffset ' index where pixel can be found
            Case 32, 9 To 13    ' all white space; any pixel info is complete
                If c Then
                    ' convert ascii 0 to bitmap white of 1
                    If inStream(c) = 48 Then outStream(L) = outStream(L) Or bLShift
                    If bLShift = 1& Then        ' adjust shifting
                        bLShift = 128&
                        L = L + 1&              ' move dest pointer along
                    Else
                        bLShift = bLShift \ 2&
                    End If                      ' keep track of rows completed
                    totalProcessed = totalProcessed + 1&
                    If totalProcessed = Width Then
                        nrRows = nrRows + 1&
                        If nrRows = Height Then Exit For
                        bLShift = 128&          ' reset shifting value
                        L = nrRows * scanWidth  ' ensure pointing at new scanline
                        totalProcessed = 0&     ' reset
                    End If
                    c = 0&                      ' reset
                End If
            Case Else
                Exit For    ' unexpected, bug out with what we've got
            End Select
        Next
        ' assign final pixel if no white-space followed it
        If totalProcessed < Width Then
            If c Then
                If inStream(c) = 48 Then outStream(L) = outStream(L) Or bLShift
            End If
        End If
    Else                                        ' binary mode
        ' validate bytes to read
        If (Width = scanWidth * 8&) Then        ' faster if naturally word-aligned
            If StreamOffset + (Width * Height) \ 8 <= UBound(inStream) + 1& Then
                c = StreamOffset + (Width * Height) \ 8 - 1&
                If ((Width * Height) Mod 8) Then c = c + 1
            Else
                c = UBound(inStream)
            End If
            If invertPalette Then               ' PNM vs PAM
                For StreamOffset = StreamOffset To c
                    outStream(L) = inStream(StreamOffset) Xor &HFF ' invert colors
                    L = L + 1&
                Next
            Else
                CopyMemory outStream(L), inStream(StreamOffset), scanWidth * Height
            End If
        Else                                    ' a bit slower; need to provide word-alignment
            If invertPalette Then bLShift = &HFF Else bLShift = 0&
            For StreamOffset = StreamOffset To UBound(inStream)
                outStream(L) = inStream(StreamOffset) Xor bLShift ' flip black to white & vice versa if needed
                totalProcessed = totalProcessed + 8&
                If totalProcessed >= Width Then         ' keep track of rows completed
                    nrRows = nrRows + 1&
                    If nrRows = Height Then Exit For
                    L = nrRows * scanWidth - 1&         ' ensure pointing at new scanline
                    totalProcessed = 0&                 ' reset
                End If
                L = L + 1&                              ' move dest pointer along
            Next
            ' handle last pixel if needed
            If totalProcessed < Width Then outStream(L) = inStream(StreamOffset) Xor bLShift
        End If
    End If
ExitRoutine:
    pvParseBlackWhite = True

End Function

Private Function pvParseGrayscale(inStream() As Byte, outStream() As Byte, Width As Long, Height As Long, maxValue As Long, StreamOffset As Long) As Boolean

    ' Parse GrayScale (PGM) files
    ' Note that grayscale values are relative to the MaxValue set in the file

    Dim totalProcessed As Long, lDenom As Long
    Dim L As Long, I As Long, c As Long
    Dim nrRows As Long, scanWidth As Long
    Dim int16 As Integer
    
    ' per specs, the grayscale should be modified by a 2.2 gamma value
    ' however, most viewers I've tested treat the grayscale without gamma correction
'    Const m_fGamma As Double = 2.2
'    Dim gLUT(0 To 255) As Byte
'    For I = 0 To 255
'        c = (255# * ((I / 255#) ^ (1# / m_fGamma))) ' + 0.5
'        If (c > 255) Then c = 255
'        gLUT(I) = c
'    Next I
'    I = 0&: c = I
   
    On Error GoTo ExitRoutine
    scanWidth = modCommon.ByteAlignOnWord(8, Width)
    If inStream(1) < 52 Then                            ' ASCII vs. binary
        c = &H80000000
        For StreamOffset = StreamOffset To UBound(inStream)
            Select Case inStream(StreamOffset)
            Case 48 To 57 ' numeric values
                c = (c And &HFFFFFF) * 10& + (inStream(StreamOffset) And &HF&)
            Case 32, 9 To 13    ' all white space
                If c > -1& Then
                    If c <= maxValue Then outStream(L) = m_LUT(c And &HFF&) ' apply it to our grayscale palette
                    c = &H80000000
                    totalProcessed = totalProcessed + 1&        ' keep track of completed rows
                    If totalProcessed = Width Then
                        nrRows = nrRows + 1&
                        If nrRows = Height Then Exit For
                        L = scanWidth * nrRows - 1&             ' ensure pointing to new scanline
                        totalProcessed = 0&
                    End If
                    L = L + 1&                                  ' move dest pointer along
                End If
            Case Else
                Exit For    ' unexpected, bug out with what we've got
            End Select
        Next
        ' assign final pixel if no white-space followed it
        If c > -1& Then
            If c <= maxValue Then outStream(L) = m_LUT(c And &HFF)
        End If
        
    Else                                                ' Binary mode
        If (maxValue = 255&) And (Width = scanWidth) Then
            ' validate number of bytes to read
            If StreamOffset + (Width * Height) <= UBound(inStream) + 1& Then
                c = StreamOffset + (Width * Height) - 1&
            Else
                c = UBound(inStream)
            End If                                          ' when naturally word-aligned much faster
            CopyMemory outStream(0), inStream(StreamOffset), c - StreamOffset + 1&
            
        Else
            If maxValue < 256& Then c = 1& Else c = 2&
            For StreamOffset = StreamOffset To UBound(inStream) Step c ' convert to relational luminance
                outStream(L) = m_LUT(inStream(StreamOffset))
                totalProcessed = totalProcessed + 1&    ' keep track of rows completed
                If totalProcessed = Width Then
                    nrRows = nrRows + 1&
                    L = scanWidth * nrRows - 1&         ' ensure pointing to new scanline
                    totalProcessed = 0&
                End If
                L = L + 1&                              ' move dest pointer along
            Next
        End If
    End If
    
ExitRoutine:
    pvParseGrayscale = True

End Function

Private Function pvParseTrueColorAlpha(inStream() As Byte, outStream() As Byte, bmpd As BitmapData, maxValue As Long, pMode As Long, StreamOffset As Long) As Boolean

    ' PAM rgb_alpha routine
    ' PAM bits are R,G,B,A order and we need B,G,R,A order

    Dim totalProcessed As Long
    Dim L As Long, b As Long, g As Long, a As Long
    Dim nrRows As Long, scanWidth As Long
    
    scanWidth = bmpd.Width * 4&
    If (pMode And &H7F) = 54 Then ' RGB alpha else grayscale alpha
        If maxValue = 255& Then
            For StreamOffset = StreamOffset To StreamOffset + scanWidth * bmpd.Height - 1& Step 4&
                outStream(L) = inStream(StreamOffset + 2&)
                outStream(L + 1&) = inStream(StreamOffset + 1&)
                outStream(L + 2&) = inStream(StreamOffset)
                outStream(L + 3&) = inStream(StreamOffset + 3&)
                L = L + 4&
            Next
        Else
            If maxValue < 256& Then
                g = 1&: b = 2&: a = 3&
            Else
                g = 2&: b = 4&: a = 6&
            End If
            For StreamOffset = StreamOffset To StreamOffset + bmpd.Width * bmpd.Height * b * 2& - 1& Step b * 2&
                outStream(L + 2&) = m_LUT(inStream(StreamOffset))
                outStream(L + 1&) = m_LUT(inStream(StreamOffset + g))
                outStream(L) = m_LUT(inStream(StreamOffset + b))
                outStream(L + 3&) = m_LUT(inStream(StreamOffset + a))
                totalProcessed = totalProcessed + 1&    ' keep track of completed rows
                If totalProcessed = bmpd.Width Then
                    nrRows = nrRows + 1&
                    L = scanWidth * nrRows - 4&         ' ensure pointing to new scanline
                    totalProcessed = 0&
                End If
                L = L + 4&                              ' move dest pointer along
            Next
        End If
    Else                            ' grayscale with alpha
        If maxValue = 255& Then
            For StreamOffset = StreamOffset To StreamOffset + bmpd.Width * bmpd.Height * 2& - 1& Step 2&
                outStream(L) = m_LUT(inStream(StreamOffset))
                outStream(L + 1&) = outStream(L)
                outStream(L + 2&) = outStream(L)
                outStream(L + 3&) = m_LUT(inStream(StreamOffset + 1&))
                L = L + 4&
            Next
        Else
            If maxValue < 256& Then
                g = 1&: b = 2&: a = 3&
            Else
                g = 2&: b = 4&: a = 6&
            End If
            For StreamOffset = StreamOffset To StreamOffset + bmpd.Width * bmpd.Height * b& - 1& Step b
                outStream(L) = m_LUT(inStream(StreamOffset))
                outStream(L + 1&) = outStream(L)
                outStream(L + 2&) = outStream(L)
                outStream(L + 3&) = m_LUT(inStream(StreamOffset + g))
                totalProcessed = totalProcessed + 1&    ' keep track of completed rows
                If totalProcessed = bmpd.Width Then
                    nrRows = nrRows + 1&
                    L = scanWidth * nrRows - 4&         ' ensure pointing to new scanline
                    totalProcessed = 0&
                End If
                L = L + 4&                              ' move dest pointer along
            Next
            
        End If
    End If
    pvParseTrueColorAlpha = True
    
End Function


Private Function pvParseTrueColor(inStream() As Byte, outStream() As Byte, Width As Long, Height As Long, maxValue As Long, StreamOffset As Long) As Boolean

    ' Parse Color (PPM) files
    ' Note that R,G,B componnent values are relative to the MaxValue set in the file

    Dim totalProcessed As Long, totalPixels As Long
    Dim L As Long, I As Long, c As Long
    Dim nrRows As Long, scanWidth As Long, lRGBoffset As Long
    
    On Error GoTo ExitRoutine
    scanWidth = modCommon.ByteAlignOnWord(24, Width)
    totalPixels = Width * 3&
    If inStream(1) < 52 Then                                ' ASCII vs. binary
        ' values in text file are R,G,B. However, we need to build bitmap as B,G,R
        lRGBoffset = 2&                                     ' set offset for Red channel
        c = &H80000000
        For StreamOffset = StreamOffset To UBound(inStream)
            Select Case inStream(StreamOffset)
            Case 48 To 57 ' numeric values
                c = (c And &HFFFFFF) * 10& + (inStream(StreamOffset) And &HF&)
            Case 32, 9 To 13 ' all white space
                If c > -1& Then
                    If c <= maxValue Then outStream(L + lRGBoffset) = m_LUT(c And &HFF)
                    c = &H80000000                          ' reset
                    totalProcessed = totalProcessed + 1&    ' keep track of completed rows
                    If totalPixels = totalProcessed Then
                        nrRows = nrRows + 1&
                        If nrRows = Height Then Exit For
                        L = scanWidth * nrRows - 1&         ' ensure pointing to new scanline
                        totalProcessed = 0&
                        lRGBoffset = -2&
                    End If
                    L = L + 1&                              ' move dest pointer along
                    If lRGBoffset = -2& Then lRGBoffset = -lRGBoffset Else lRGBoffset = lRGBoffset - 2&
                End If
            Case Else
                Exit For    ' unexpected, bug out with what we've got
            End Select
        Next
        ' assign final pixel if no white-space followed it
        If c > -1& Then
            If c <= maxValue Then outStream(L + lRGBoffset) = m_LUT(c And &HFF)
        End If
    Else                                                ' Binary mode
        If (maxValue = 255&) And (scanWidth = Width * 3&) Then
            ' validate number of bytes to read
            If StreamOffset + (Width * Height * 3&) <= UBound(inStream) + 1& Then
                c = StreamOffset + (Width * Height * 3&) - 1&
            Else
                c = UBound(inStream) - 2&
            End If                                          ' a bit faster when word-aligned
            For StreamOffset = StreamOffset To c Step 3
                outStream(L + 2&) = inStream(StreamOffset)
                outStream(L + 1&) = inStream(StreamOffset + 1&)
                outStream(L) = inStream(StreamOffset + 2&)
                L = L + 3&
            Next
        Else
            If maxValue < 256& Then c = 1&: I = 2& Else c = 2&: I = 4&
            For StreamOffset = StreamOffset To UBound(inStream) Step 3& * c
                outStream(L + 2&) = m_LUT(inStream(StreamOffset))
                outStream(L + 1&) = m_LUT(inStream(StreamOffset + c))
                outStream(L) = m_LUT(inStream(StreamOffset + I))
                totalProcessed = totalProcessed + 1&    ' keep track of completed rows
                If totalProcessed = Width Then
                    nrRows = nrRows + 1&
                    L = scanWidth * nrRows - 3&         ' ensure pointing to new scanline
                    totalProcessed = 0&
                End If
                L = L + 3&                              ' move dest pointer along
            Next
        End If
    End If
    
ExitRoutine:
    pvParseTrueColor = True

End Function

Private Function pvParseHeader(inStream() As Byte, Width As Long, Height As Long, MaxVal As Long) As Long

    ' routine parses out the dimensions, MaxValue & start of pixel data from PBM, PGM, PPM files

    ' the header consists of magic number, width, height & max value (except for PBM).
    ' Each of those values are "white space" delimited
    ' Comments can appear anywhere, but must begin with hash symbol & end in line feed

    Dim b As Long, bComment As Boolean
    Dim lResult As Long
    Dim bEOT As Boolean, bBOT As Boolean        ' EOT=end of tokenizing, BOT=begin of tokenizing
    
    On Error GoTo ExitRoutine
    Width = 0&: Height = Width: MaxVal = Width
    lResult = &H80000000
    For b = 2& To UBound(inStream)
        Select Case inStream(b)
        Case 48 To 57                           ' numerical data
            If bComment = False Then
                If bBOT Then
                    pvParseHeader = b
                    Exit For
                Else
                    lResult = (lResult And &H7FFFFFFF) * 10& + (inStream(b) And &HF)
                End If
            End If
        Case 10, 13                             ' carriage return
            bComment = False                    ' comments must end with a LF
            bEOT = (lResult > -1&)              ' any non-comment value is complete now
        Case 32, 9, 11, 12 ' white space
            bEOT = (lResult > -1&)              ' any non-comment value is complete now
        Case 35 ' # comment
            bComment = True                     ' comment & anything can follow til next LF
        Case Else
            If bComment = False Then Exit For
        End Select
        
        If bEOT Then                            ' done parsing a wanted value
            If lResult > -1& Then               ' any zero value fails parsing
                Select Case 0&
                Case Width
                    Width = lResult
                Case Height
                    Height = lResult            ' B&W don't have max value line
                    If inStream(1) = 49 Or inStream(1) = 52 Then bBOT = True
                Case Else
                    If lResult < &H10000 Then   ' per specs 65535 is max
                        MaxVal = lResult
                        bBOT = True
                    Else
                        Exit For
                    End If
                End Select
                If bBOT Then
                    If inStream(1) > 51 Then
                        pvParseHeader = b + 1&  ' binary mode: next byte must start binary data
                        Exit For
                    End If
                End If
                bEOT = False
                lResult = &H80000000
            Else
                Exit For
            End If
        End If
    Next
ExitRoutine:
            
End Function

Private Function pvParseHeaderPAM(inStream() As Byte, Width As Long, Height As Long, _
                                  MaxVal As Long, Depth As Long, PNMmode As Long) As Long

    ' PAM headers are standard and look like this:
    
    ' P7
    ' WIDTH xxx
    ' HEIGHT xxx
    ' DEPTH xxx
    ' MAXVAL xxx
    ' TRUPLTYPE xxx
    ' ENDHDR
    
    Dim sLines() As String, lPos As Long
    Dim bLines() As Byte, X As Long, lDelim As Long
    
    Const ENDHDRsuffix As Long = &H52444844   ' the last 4 chars of ENDHDR
    
    For X = 3& To UBound(inStream) - 6&
        If inStream(X) = 78 Then ' looking for letter N of the flag: ENDHDR
            If inStream(X - 1) = 69 Then                                ' previous letter E ?
                If inStream(X + 5) = 10 Or inStream(X + 5) = 13 Then    ' ends in LF or CR ?
                    CopyMemory lDelim, inStream(X + 1&), 4&             ' see if rest is good: DHDR
                    If lDelim = ENDHDRsuffix Then
                        lPos = X + 6&
                        If inStream(lPos - 1&) = 13 Then lPos = lPos + 1&
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    If lPos = 0& Then Exit Function
    
    ReDim bLines(0 To X - 6&)           ' lets convert byte>string following P7 to just before ENDHDR
    CopyMemory bLines(0), inStream(3), UBound(bLines) + 1&
    sLines() = Split(StrConv(bLines(), vbUnicode), vbLf)
    Erase bLines()                      ' loop thru what we converted & parse
    
    For X = 0& To UBound(sLines)        ' ensure parsed tokens are handled in this class
        lDelim = InStr(sLines(X), " ")
        If lDelim Then
            Select Case Left$(sLines(X), lDelim - 1&)
            Case "WIDTH"
                Width = Val(Mid$(sLines(X), lDelim + 1&))
                If Width < 1& Then Exit For
            Case "HEIGHT"
                Height = Val(Mid$(sLines(X), lDelim + 1&))
                If Height < 1& Then Exit For
            Case "DEPTH"
                Depth = Val(Mid$(sLines(X), lDelim + 1&))
                If Depth < 1& Or Depth > 8 Then Exit For
            Case "MAXVAL"
                MaxVal = Val(Mid$(sLines(X), lDelim + 1&))
                If MaxVal < 1& Or MaxVal > 65535 Then Exit For
            Case "TUPLTYPE"
                If pvParseHeaderPAM = 0& Then
                    If (Width = 0& Or Height = 0& Or MaxVal = 0& Or Depth = 0&) Then Exit For
                    ' just in case lines are delimited with vbCrLf (unexepected) vs vbLf (expected)
                    Do Until Right$(sLines(X), 1) <> vbCr
                        sLines(X) = Left$(sLines(X), Len(sLines(X)) - 1&)
                    Loop
                    Select Case Mid$(sLines(X), lDelim + 1&)
                        Case "RGB": PNMmode = 54&                       ' process like ppm binary
                            If Not (Depth = 3& Or Depth = 6&) Then Exit For
                        Case "RGB_ALPHA": PNMmode = 54& Or 128&         ' process alpha channel
                            If Not (Depth = 4& Or Depth = 8&) Then Exit For
                        Case "GRAYSCALE": PNMmode = 53&                 ' process like pgm binary
                            If Depth > 2 Then Exit For
                        Case "GRAYSCALE_ALPHA": PNMmode = 53& Or 128&   ' process alpha channel
                            If Not (Depth = 2& Or Depth = 4&) Then Exit For
                        Case "BLACKANDWHITE": PNMmode = 52&             ' prcoess like pbm binary
                            If Depth > 1& Then Exit For
                        Case Else
                            Exit For
                    End Select
                    pvParseHeaderPAM = lPos
                Else
                    pvParseHeaderPAM = 0&       ' not prepared to handle multiple TRUPLTYPE entries
                    Exit For
                End If
            Case Else
                Exit For
            End Select
        End If
    Next

End Function

Public Function SaveAsPNM(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                           SS As SAVESTRUCT, SaveAsPam As Boolean) As Long

    ' saves image as a PBM, PGM, PPM, or PAM
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    ' Grayscale (PGM only created if image is truly grayscale at any bit depth other than 1)

    ' True compression value will save PNM in binary mode; else ASCII mode
    ' for PAM, image bits are always binary vs. ASCII

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
    
    Dim dstDepth As Long, aPal() As Byte
    Dim hHandle As Long, imgDataLen As Long
    Dim palType As Long, imgType As Long
    Dim theStream() As Byte, outStream() As Byte
    Dim bOK As Boolean, bAlpha As Long
    Dim lResult As Long, srcDepth As Long
    Dim dSA As SafeArray
    Dim sizeI As RECTI, tBMP As BitmapData
    Dim cPal As cColorReduction, Compression As Boolean
    Dim tmpPic As StdPicture, tmpDO As DataObject
    Dim tGDIpImage As GDIpImage, tObject As Object
    Dim X As Long, Y As Long

    If SaveAsPam Then
        bAlpha = alpha_None
        Compression = True
        imgType = lvicPicTypePAM
    Else
        imgType = lvicPicTypePNM
        If SS.CompressionJPGQuality <> lvicFormatUncompressed Then Compression = True
    End If
    Set cPal = New cColorReduction
    srcDepth = modCommon.ColorDepthToColorType(SS.ColorDepth, SourceHandle)
    If returnMedium = saveTo_Clipboard Or returnMedium = saveTo_DataObject Then SS.reserved1 = SS.reserved1 Or &H20000000
    
    If SS.ColorDepth = lvicDefaultReduction Then
        SS.reserved1 = SS.reserved1 Or &H10000000
        If (srcDepth And &HFF00&) \ &H100& > 8& Then
            lResult = cPal.PalettizeToArray(SourceHandle, theStream(), orient_8bppIndexes Or orient_TopDown Or orient_SortGrayscale Or orient_PNMformat, 0&, 0&, bAlpha, SS)
            If lResult Then srcDepth = lResult
        End If
    ElseIf SS.ColorDepth < lvicConvert_TrueColor24bpp And SS.ColorDepth > lvicDefaultReduction Then
        lResult = cPal.PalettizeToArray(SourceHandle, theStream(), orient_8bppIndexes Or orient_TopDown Or orient_SortGrayscale Or orient_PNMformat, 0&, 0&, bAlpha, SS)
        If lResult Then srcDepth = lResult Else Exit Function                   ' failure
    End If
    
    If lResult = 0& Then
        If SS.reserved2 <> 0& And (SS.reserved1 And &HFF00&) \ &H100& = imgType Then
            CopyMemory tObject, SS.reserved2, 4&
            Set tGDIpImage = tObject
            CopyMemory tObject, 0&, 4&
            bOK = tGDIpImage.ExtractImageData(outStream)
            If (outStream(1) > 51) = Compression Then
                lResult = srcDepth
            Else
                Set tGDIpImage = Nothing
            End If
        End If
        If lResult = 0& And (srcDepth And &HFF00&) \ &H100& <= 8& Then          ' need to ensure at least 8bpp indexes
            SS.ColorDepth = lvicConvert_256Colors
            SS.PaletteType = lvicPaletteAdaptive
            SS.reserved1 = SS.reserved1 Xor &H10000000
            lResult = cPal.PalettizeToArray(SourceHandle, theStream(), orient_8bppIndexes Or orient_TopDown Or orient_SortGrayscale Or orient_PNMformat, 0&, 0&, bAlpha, SS)
            If lResult Then srcDepth = lResult
        End If
    End If
    
    If tGDIpImage Is Nothing Then
        If lResult = 0& Then ' using source
            If cPal.ImportRemotePalette(SourceHandle) = False Then Set cPal = New cColorReduction
            ' should only get here if 1) b&w, 2) 8bpp grayscale 3) 16/24bpp color no alpha
            Select Case (srcDepth And &HFF00&) \ &H100
            Case 1
                If cPal.IsBlackWhite Then
                    dstDepth = 1&: tBMP.PixelFormat = lvicColor1bpp
                ElseIf cPal.IsGrayscale(True) Then
                    dstDepth = 8: tBMP.PixelFormat = lvicColor1bpp
                Else
                    dstDepth = 24&: tBMP.PixelFormat = lvicColor24bpp
                End If
            Case 8
                If cPal.IsGrayscale(True) Then
                    dstDepth = 8: tBMP.PixelFormat = lvicColor8bpp
                Else
                    dstDepth = 24&: tBMP.PixelFormat = lvicColor24bpp
                End If
            Case 32
                If srcDepth = lvicColor32bpp Or imgType = lvicPicTypePNM Then
                    dstDepth = 24&: tBMP.PixelFormat = lvicColor24bpp
                Else
                    dstDepth = 32&: tBMP.PixelFormat = srcDepth
                End If
            Case Else
                dstDepth = 24&: tBMP.PixelFormat = lvicColor24bpp
            End Select
        Else
            Select Case (srcDepth And &HFF00&) \ &H100
            Case 1
                dstDepth = 1&: tBMP.PixelFormat = lvicColor1bpp
            Case 8 ' can only be grayscale else returned as 24bpp format
                dstDepth = 8&: tBMP.PixelFormat = lvicColor8bpp
            Case 32
                If srcDepth = lvicColor32bpp Or imgType = lvicPicTypePNM Then
                    dstDepth = 24&: tBMP.PixelFormat = lvicColor24bpp
                Else
                    dstDepth = 32&: tBMP.PixelFormat = srcDepth
                End If
            Case Else
                dstDepth = 24&: tBMP.PixelFormat = lvicColor24bpp
            End Select
        End If
        If dstDepth = 8& Then ' convert palette to red component only for our algo
            ReDim aPal(0 To cPal.Count * 4& - 1&)
            cPal.CopyLocalPalette VarPtr(aPal(0)), True, 4&
            X = 0&: For Y = 0 To UBound(aPal) Step 4&
                aPal(X) = aPal(Y)
                X = X + 1
            Next
        End If
        Set cPal = Nothing
        
        If lResult = 0& Then
            sizeI.nWidth = SS.Width: sizeI.nHeight = SS.Height
            If GdipBitmapLockBits(SourceHandle, sizeI, ImageLockModeRead, tBMP.PixelFormat, tBMP) Then Exit Function
            With dSA
                .cbElements = 1
                .cDims = 1
                .pvData = tBMP.Scan0Ptr
                If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
                .rgSABound(0).cElements = Abs(tBMP.stride) * tBMP.Height
            End With
            CopyMemory ByVal VarPtrArray(theStream), VarPtr(dSA), 4&
        ElseIf tGDIpImage Is Nothing Then   ' inStream has our data
            tBMP.Height = SS.Height: tBMP.Width = SS.Width
            tBMP.stride = modCommon.ByteAlignOnWord((tBMP.PixelFormat And &HFF00&) \ &H100&, SS.Width)
        End If
        If pvBuildPNM(outStream(), theStream(), tBMP, dstDepth, 255& + (dstDepth = 1&) * 255&, Compression, aPal(), imgType) = False Then GoTo ExitRoutine
        Erase aPal()
        If dSA.pvData Then
            GdipBitmapUnlockBits SourceHandle, tBMP
            CopyMemory ByVal VarPtrArray(theStream), 0&, 4&
            dSA.pvData = 0&
        End If
    Else
        Set cPal = Nothing
    End If
    Erase theStream()
    Erase aPal()
    
    If returnMedium = saveTo_Array Then                             ' set return object
        modCommon.MoveArrayToVariant returnObject, outStream(), True
        SaveAsPNM = True
    ElseIf returnMedium = saveTo_File Then
        WriteFile CLng(returnObject), outStream(0), UBound(outStream) + 1&, lResult, ByVal 0&
        SaveAsPNM = (lResult > UBound(outStream))
    ElseIf returnMedium = saveTo_GDIplus Then
        Set returnObject = modCommon.LoadImage(outStream(), True, , True)
        If returnObject.Handle Then SaveAsPNM = saveTo_GDIplus
        SaveAsPNM = saveTo_GDIplus
    Else
        lResult = 0&
        If tGDIpImage Is Nothing Then Set tGDIpImage = modCommon.LoadImage(outStream(), , , True)
        If tGDIpImage.Handle Then GdipCreateHBITMAPFromBitmap tGDIpImage.Handle, lResult, SS.RSS.FillColorARGB Or &HFF000000
        Set tGDIpImage = Nothing
        If lResult Then
            SaveAsPNM = True
            Set tmpPic = modCommon.HandleToStdPicture(lResult, vbPicTypeBitmap)
            If tmpPic Is Nothing Then
                DeleteObject lResult: SaveAsPNM = False
            ElseIf returnMedium = saveTo_stdPicture Then
                Set returnObject = tmpPic
                SaveAsPNM = (tmpPic.Handle <> 0&)
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
    If dSA.pvData Then
        GdipBitmapUnlockBits SourceHandle, tBMP
        CopyMemory ByVal VarPtrArray(theStream), 0&, 4&
    End If
End Function

Private Function pvBuildPNM(outStream() As Byte, inStream() As Byte, bmpInfo As BitmapData, _
                                dstDepth As Long, maxValue As Long, bCompressed As Boolean, _
                                colorPal() As Byte, imgType As ImageFormatEnum) As Boolean

    Dim X As Long, Y As Long, L As Long
    Dim imgDataLen As Long, pOffset As Long
    Dim aHeaderRmks() As Byte, sValue As String
    Dim lShift As Long, dPos As Long
    Dim aTens() As Byte, aCNotes() As Byte

    On Error GoTo ExitRoutine
    ' PNM FORMAT LOOKS LIKE THIS....
    '       P# + vbLF
    '       Width + space + Height + vbLF
    '       MaxValue + vbLF  (unless black & white where MaxValue is always assumed to be 1)
    ' PAM FORMAT LOOKS LIKE THIS...
    '       P7 + vbLF
    '       WIDTH xxx + vbLF
    '       HEIGHT xxx + vbLF
    '       DEPTH xxx + vbLF
    '       MAXVAL xxx + vbLF
    '       TUPLETYPE xxx + vbLF (types: BLACKANDWHITE,GRAYSCALE,RGB,RGB_ALPHA,GRAYSCALE_ALPHA
    '       ENDHDER + vbLF
    If imgType = lvicPicTypePNM Then
        aHeaderRmks = StrConv("# " & CStr(Year(Date)) & ": Created by LaVolpe AlphaImage Control" & vbLf, vbFromUnicode)
        imgDataLen = Len(CStr(bmpInfo.Width)) + Len(CStr(bmpInfo.Height)) + UBound(aHeaderRmks) + 6&
    End If
    If maxValue = 0& Then
        If bCompressed Then
            X = (bmpInfo.Width \ 8): If (bmpInfo.Width Mod 8) Then X = X + 1&
            imgDataLen = imgDataLen + X * bmpInfo.Height
            L = 52& ' b&w compressed
        Else
            imgDataLen = imgDataLen + (bmpInfo.Width * bmpInfo.Height * 2&)
            L = 49& ' b&w uncompressed
        End If
    Else
        If imgType = lvicPicTypePNM Then imgDataLen = imgDataLen + Len(CStr(maxValue)) + 1&
        If bCompressed Then
            imgDataLen = imgDataLen + bmpInfo.Width * (dstDepth \ 8) * bmpInfo.Height
            If dstDepth = 8& Then L = 53& Else L = 54& ' color/grayscale compressed
        
        '/// counting space requirements before sizing array; won't need to re-size/re-dim constantly
        ElseIf dstDepth = 24& Then
            For Y = 0& To bmpInfo.Height - 1&
                X = Y * Abs(bmpInfo.stride)
                For X = X To X + bmpInfo.Width * 3& - 1&
                    If inStream(X) > 99 Then
                        imgDataLen = imgDataLen + 4&
                    ElseIf inStream(X) < 10 Then
                        imgDataLen = imgDataLen + 2&
                    Else
                        imgDataLen = imgDataLen + 3&
                    End If
                Next
            Next
            L = 51& ' color uncompressed
            
        ElseIf dstDepth = 8& Then
            If bmpInfo.PixelFormat = lvicColor1bpp Then
                ' 2 color image, but not black & white, it is grayscale
                If Len(CStr(colorPal(0))) = Len(CStr(colorPal(1))) Then
                    imgDataLen = imgDataLen + (Len(CStr(colorPal(0))) + 1&) * bmpInfo.Width * bmpInfo.Height
                Else
                    For Y = 0& To bmpInfo.Height - 1&
                        X = Y * Abs(bmpInfo.stride)
                        lShift = 128&
                        For dPos = 1& To bmpInfo.Width
                            Select Case (colorPal((inStream(X) And lShift) \ lShift) And &HFF)
                            Case Is > 99
                                imgDataLen = imgDataLen + 4&
                            Case Is < 10
                                imgDataLen = imgDataLen + 2&
                            Case Else
                                imgDataLen = imgDataLen + 3&
                            End Select
                            If lShift = 1& Then
                                X = X + 1&
                                lShift = 128&
                            Else
                                lShift = lShift \ 2&
                            End If
                        Next
                    Next
                End If
            Else
                For Y = 0& To bmpInfo.Height - 1&
                    X = Y * Abs(bmpInfo.stride)
                    For X = X To X + bmpInfo.Width - 1&
                        If colorPal(inStream(X)) > 99 Then
                            imgDataLen = imgDataLen + 4&
                        ElseIf colorPal(inStream(X)) < 10 Then
                            imgDataLen = imgDataLen + 2&
                        Else
                            imgDataLen = imgDataLen + 3&
                        End If
                    Next
                Next
            End If
            L = 50& ' uncompressed grayscale
        End If
    End If
    ' output the pnm type & remarks
    If imgType = lvicPicTypePAM Then    ' build the PAM header
        sValue = "WIDTH " & CStr(bmpInfo.Width) & vbLf & "HEIGHT " & CStr(bmpInfo.Height) & vbLf
        If dstDepth < 24& Then
            sValue = sValue & "DEPTH 1" & vbLf
        ElseIf dstDepth = 32& Then
            sValue = sValue & "DEPTH 4" & vbLf
        Else
            sValue = sValue & "DEPTH 3" & vbLf
        End If
        If maxValue = 0& Then
            sValue = sValue & "MAXVAL 1" & vbLf & "TUPLTYPE "
        Else
            sValue = sValue & "MAXVAL " & CStr(maxValue) & vbLf & "TUPLTYPE "
        End If
        Select Case dstDepth
            Case 1: sValue = sValue & "BLACKANDWHITE"
            Case 8: sValue = sValue & "GRAYSCALE"
            Case 24: sValue = sValue & "RGB"
            Case 32: sValue = sValue & "RGB_ALPHA"
        End Select
        sValue = sValue & vbLf & "ENDHDR" & vbLf
        imgDataLen = imgDataLen + Len(sValue) + 3&
    End If
    ReDim outStream(0 To imgDataLen - 1&)
    outStream(0) = 80: outStream(2) = 10
    If imgType = lvicPicTypePAM Then
        outStream(1) = 55 ' 7
        aHeaderRmks() = StrConv(sValue, vbFromUnicode)
        CopyMemory outStream(3), aHeaderRmks(0), Len(sValue)
        L = UBound(aHeaderRmks) + 4&
        Erase aHeaderRmks()
    Else
        outStream(1) = L
        CopyMemory outStream(3), aHeaderRmks(0), UBound(aHeaderRmks) + 1&
        L = UBound(aHeaderRmks) + 3&
        Erase aHeaderRmks()
        
        ' output the size & max value
        sValue = CStr(bmpInfo.Width)
        For X = 1 To Len(sValue)
            L = L + 1&
            outStream(L) = Asc(Mid$(sValue, X, 1))
        Next
        L = L + 1&: outStream(L) = 32
        sValue = CStr(bmpInfo.Height)
        For X = 1& To Len(sValue)
            L = L + 1&
            outStream(L) = Asc(Mid$(sValue, X, 1))
        Next
        L = L + 1&: outStream(L) = 10
        If maxValue Then ' set max value, else b&w
            sValue = CStr(maxValue)
            For X = 1& To Len(sValue)
                L = L + 1&
                outStream(L) = Asc(Mid$(sValue, X, 1))
            Next
            L = L + 1&: outStream(L) = 10
        End If
        L = L + 1&
    End If
    sValue = vbNullString
    
    ' now write the format to the out array
    If bCompressed Then
        For Y = 0& To bmpInfo.Height - 1&
            X = Y * Abs(bmpInfo.stride)
            If dstDepth = 32 Then       ' true color needs BGR converted to RGB + Alpha
                For X = X To X + bmpInfo.Width * 4& - 1& Step 4&
                    outStream(L) = inStream(X + 2&)
                    outStream(L + 1&) = inStream(X + 1&)
                    outStream(L + 2&) = inStream(X)
                    outStream(L + 3&) = inStream(X + 3&)
                    L = L + 4&
                Next
            ElseIf dstDepth = 24& Then  ' true color needs BGR converted to RGB
                For X = X To X + bmpInfo.Width * 3& - 1& Step 3&
                    outStream(L) = inStream(X + 2&)
                    outStream(L + 1&) = inStream(X + 1&)
                    outStream(L + 2&) = inStream(X)
                    L = L + 3&
                Next
            ElseIf bmpInfo.PixelFormat = lvicColor1bpp Then ' b&w need to have bits inverted
                X = Y * Abs(bmpInfo.stride)
                If dstDepth = 1& Then
                    For X = X To X + bmpInfo.Width \ 8& - 1&
                        outStream(L) = inStream(X) Xor &HFF
                        L = L + 1&
                    Next
                    If (bmpInfo.Width Mod 8&) Then
                        outStream(L) = inStream(X) Xor &HFF
                        L = L + 1&
                    End If
                Else
                    lShift = 128&
                    For L = L To L + bmpInfo.Width - 1&
                        If (inStream(X) And lShift) Then outStream(L) = colorPal(0) Else outStream(L) = colorPal(1)
                        If lShift = 1& Then
                            X = X + 1&
                            lShift = 128&
                        Else
                            lShift = lShift \ 2&
                        End If
                    Next
                End If
            ElseIf UBound(colorPal) = 1023& Then ' 256 color, 8 bpp is straight copy
                CopyMemory outStream(L), inStream(X), bmpInfo.Width
                L = L + bmpInfo.Width
            Else
                For X = X To X + bmpInfo.Width - 1&
                    outStream(L) = colorPal(inStream(X))
                    L = L + 1&
                Next
            End If
        Next
        
    ' uncompressed
    Else
        If dstDepth > 1& Then   ' build lookup table for 2 & 3 digit color components
            aTens() = StrConv("10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73 74 75 76 77 78 79 80 81 82 83 84 85 86 87 88 89 90 91 92 93 94 95 96 97 98 99 ", vbFromUnicode)
            aCNotes() = StrConv("100 101 102 103 104 105 106 107 108 109 110 111 112 113 114 115 116 117 118 119 120 121 122 123 124 125 126 127 128 129 130 131 132 133 134 135 136 137 138 139 140 141 142 143 144 145 146 147 148 149 150 151 152 153 154 155 156 157 158 159 160 161 162 163 164 165 166 167 168 169 170 171 172 173 174 175 176 177 178 179 180 181 182 183 184 185 186 187 188 189 190 191 192 193 194 195 196 197 198 199 200 201 202 203 204 205 206 207 208 209 210 211 212 213 214 215 216 217 218 219 220 221 222 223 224 225 226 227 228 229 230 231 232 233 234 235 236 237 238 239 240 241 242 243 244 245 246 247 248 249 250 251 252 253 254 255 ", vbFromUnicode)
        End If
        imgDataLen = L
        For Y = 0& To bmpInfo.Height - 1&
            X = Y * Abs(bmpInfo.stride)
            If bmpInfo.PixelFormat = lvicColor24bpp Then
                For X = X To X + bmpInfo.Width * 3& - 1& Step 3&
                    For dPos = X + 2& To X Step -1&
                        If inStream(dPos) > 99 Then
                            CopyMemory outStream(L), aCNotes((inStream(dPos) - 100&) * 4&), 4&: L = L + 4&
                        ElseIf inStream(dPos) < 10 Then
                            outStream(L) = inStream(dPos) + 48
                            outStream(L + 1) = 32: L = L + 2&
                        Else
                            CopyMemory outStream(L), aTens((inStream(dPos) - 10&) * 3&), 3&: L = L + 3&
                        End If
                    Next
                Next
                
            ElseIf bmpInfo.PixelFormat = lvicColor8bpp Then
                For X = X To X + bmpInfo.Width - 1&
                    dPos = colorPal(inStream(X))
                    If dPos > 99 Then
                        CopyMemory outStream(L), aCNotes((dPos - 100&) * 4&), 4&: L = L + 4&
                    ElseIf dPos < 10 Then
                        outStream(L) = dPos + 48
                        outStream(L + 1) = 32: L = L + 2&
                    Else
                        CopyMemory outStream(L), aTens((dPos - 10&) * 3&), 3&: L = L + 3&
                    End If
                Next
            
            ElseIf bmpInfo.PixelFormat = lvicColor1bpp Then
                ' black & white needs to be parsed & colors inverted
                lShift = 128&                   ' width is byte aligned & rows start with fresh byte
                If dstDepth = 1& Then
                    For dPos = 1& To bmpInfo.Width
                        If (inStream(X) And lShift) Then outStream(L) = 48 Else outStream(L) = 49
                        outStream(L + 1) = 32: L = L + 2&
                        If lShift = 1& Then
                            lShift = 128&
                            X = X + 1&
                        Else
                            lShift = lShift \ 2&
                        End If
                    Next
                Else
                    For pOffset = 1& To bmpInfo.Width
                        dPos = colorPal((inStream(X) And lShift) \ lShift)
                        If dPos > 99 Then
                            CopyMemory outStream(L), aCNotes((dPos - 100&) * 4&), 4&: L = L + 4&
                        ElseIf dPos < 10 Then
                            outStream(L) = dPos + 48
                            outStream(L + 1) = 32: L = L + 2&
                        Else
                            CopyMemory outStream(L), aTens((dPos - 10&) * 3&), 3&: L = L + 3&
                        End If
                        If lShift = 1& Then
                            lShift = 128&
                            X = X + 1&
                        Else
                            lShift = lShift \ 2&
                        End If
                    Next
                End If
            End If
            ' clean up scan line. Want to keep "lines" in ASCII file to 70 characters or less
            ' to do so, we simply change a "white space" of chr(32) to chr(10)
            Do
                Select Case L - imgDataLen
                Case 70&
                    outStream(L - 1&) = 10
                    imgDataLen = L
                Case Is > 70&
                    For imgDataLen = L - ((L - imgDataLen) - 66&) To L - 1&
                        If outStream(imgDataLen) = 32 Then
                            outStream(imgDataLen) = 10
                            Exit For
                        End If
                    Next
                    imgDataLen = imgDataLen + 1&
                Case Else
                    Exit Do
                End Select
            Loop
        Next
        outStream(UBound(outStream)) = 10
        Erase aTens()
        Erase aCNotes()
    End If
    
    pvBuildPNM = True
ExitRoutine:
End Function

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
End Sub

Private Sub Class_Terminate()
    ' clean up
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub

