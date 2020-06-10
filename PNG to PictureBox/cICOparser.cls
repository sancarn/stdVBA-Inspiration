VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cICOparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read Icon image format & Convert to Bitmap ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take some liberties in several API declarations throughout

' Used for creating array overlays at other memory addresses
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' used to create images as needed
Private Declare Function DrawIcon Lib "user32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, ByRef lpInitBits As Any, ByRef lpInitInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long         ' +4 from .biSize
    biHeight As Long        ' +8
    biPlanes As Integer     ' +12
    biBitCount As Integer   ' +14
    biCompression As Long   ' +16
    biSizeImage As Long     ' +20
    biXPelsPerMeter As Long ' +24
    biYPelsPerMeter As Long ' +28
    biClrUsed As Long       ' +32
    biClrImportant As Long  ' 40th byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette(0 To 255) As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray          ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND ' reusable UDT for 1 & 2 dim arrays
End Type

Private Type ICONDIRENTRY
    bWidth As Byte '// Width, in pixels, of the image
    bHeight As Byte '// Height, in pixels, of the image
    bColorCount As Byte '// Number of colors in image (0 if >=8bpp)
    bReserved As Byte '// Reserved ( must be 0)
    wPlanes As Integer '// Color Planes
    wBitCount As Integer '// Bits per pixel
    dwBytesInRes As Long '// How many bytes in this resource?
    dwImageOffset As Long '// Where in the file is this image?
End Type

Private Type ICONDIR
    idReserved As Integer '// Reserved (must be 0)
    idType As Integer '// Resource Type (1 for icons)
    idCount As Integer '// How many images?
    idEntries() As ICONDIRENTRY '// An entry for each image (idCount of 'em)
End Type

Private Const png_Signature1 As Long = 1196314761   ' 1st 8 bytes of a PNG file start with these 8 bytes
Private Const png_Signature2 As Long = 169478669

Private m_icDirE() As ICONDIRENTRY  ' collection of icon directory entries
Private m_icDir As ICONDIR          ' icon directory
Private m_Bits() As Byte            ' icon bits

Public Property Get Height(Index As Long) As Long
    Height = m_icDirE(Index).bHeight    ' height of icon
    If Height = 0& Then Height = 256&   ' 256x256 icons are identified as 0 in the icon structure
End Property
Public Property Get Width(Index As Long) As Long
    Width = m_icDirE(Index).bHeight     ' width of icon
    If Width = 0& Then Width = 256&     ' 256x256 icons are identified as 0 in the icon structure
End Property

Public Property Get IsIconPNG(Index As Long) As Boolean
    IsIconPNG = m_icDirE(Index).wPlanes = 255   ' custom flag to distinguish PNG from icon
End Property
Public Property Get bitDepth(Index As Long) As Long
    bitDepth = m_icDirE(Index).wBitCount    ' bit count/depth of icon
End Property

Public Property Get IconCount() As Long
    IconCount = m_icDir.idCount
End Property

Public Property Get ColorCount(Index As Long) As Long
    ' for paletted non-PNG images, number of colors that exist
    ' This should be straight forward and is generally supplied in the icon entry's .bColorCount
    ' member. But maybe .bColorCount may not be telling us the truth or it may be missing.
    
    ' To get the proper number supplied with the icon/bitmap, we will add the total bytes
    ' used for the image & mask bytes, then add that to the bytes used for the header.
    ' The difference/4 will always be correct.
    Dim imageBits As Long, headerBits As Long
    If m_icDirE(Index).wBitCount < 9 Then
        imageBits = ColorByteCount(Index) + MaskByteCount(Index)
        headerBits = m_Bits(m_icDirE(Index).dwImageOffset)
        ColorCount = (m_icDirE(Index).dwBytesInRes - (imageBits + headerBits)) \ 4&
    End If
End Property

Public Property Get ColorByteOffset(Index As Long) As Long
    ' Return the position in the source stream where the 1st byte of the color image
    ' can be found; not called for PNGs
    Dim Offset As Long
    CopyMemory Offset, m_Bits(m_icDirE(Index).dwImageOffset), 4& ' header bytes
    Offset = m_icDirE(Index).dwImageOffset + Offset ' shift offset to where icon structure begins
    ' when image is paletted, the palette is included too
    If m_icDirE(Index).wBitCount < 16 Then          ' get number of colors used in image
        Offset = Offset + (2& ^ m_icDirE(Index).wBitCount) * 4& ' add that to the offset
    End If
    ColorByteOffset = Offset
End Property

Private Property Get MaskByteOffset(Index As Long) As Long
    ' Return the position in the source stream where the 1st byte of the mask image
    ' can be found; not called for PNGs. Here we work from the end of the icon structure
    Dim Offset As Long
    
    ' Note: 32bpp icons have masks too
    Offset = m_icDirE(Index).dwImageOffset + m_icDirE(Index).dwBytesInRes
    Offset = Offset - MaskByteCount(Index)
    MaskByteOffset = Offset
End Property

Private Property Get ColorByteCount(Index As Long) As Long
    ' Return the number of image bytes used for the color image; not PNGs
    ColorByteCount = iparseByteAlignOnWord(m_icDirE(Index).wBitCount, m_icDirE(Index).bWidth) * m_icDirE(Index).bHeight
End Property
Private Property Get MaskByteCount(Index As Long) As Long
    ' Return the number of image bytes used for the mask image; not PNGs
    MaskByteCount = iparseByteAlignOnWord(1, m_icDirE(Index).bWidth) * m_icDirE(Index).bHeight
End Property

Public Function LoadStream(inStream() As Byte, _
                ByVal desiredWidth As Long, ByVal desiredHeight As Long, _
                cHost As c32bppDIB, streamOffset As Long, streamLength As Long, _
                icoBitDepth As Long, Optional GlobalToken As Long) As Boolean

    ' Purpose: Parse byte stream to determine if it is an icon file.
    '   If it is an icon file, then select the best match for the passed
    '   size and create our application's main image from the icon
    ' Note: GIF, JPG, BMP, PNG & other formats have a magic number that
    '   indicates what type of file it is. Icons/cursors do not; so we parse & error check
    
    ' Parameters:
    ' inStream() :: an array of the icon file; can consist of more than one icon
    ' desiredWidth :: width of icon to use, if available, else used for closest match
    ' desiredHeight :: height of icon to use, if available, else used for closest match
    ' cHost :: the application's image class
    
    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim icEntry As Long, icValue As Long
    Dim icPtr As Long, icBytesNeed As Long
    Dim bIconFile As Boolean
    
    Dim tDC As Long, hDib As Long, dDC As Long, hObj As Long
    Dim tSA As SafeArray
    Dim tBMPI As BITMAPINFO
    Dim cPNG As cPNGparser
    
    With tSA                    ' overlay the passed stream with our module-level array
        .cbElements = 1         ' as byte
        .cDims = 1              ' as 1 dimensional
        .pvData = VarPtr(inStream(LBound(inStream)))
        If streamLength = 0 Then streamLength = UBound(inStream) + 1
        .rgSABound(0).cElements = streamLength
    End With
    CopyMemory ByVal VarPtrArray(m_Bits), VarPtr(tSA), 4& ' establish overlay
    
    icBytesNeed = 6&                        ' length of the ICONDIRECTORY
    If icPtr + icBytesNeed <= streamLength Then    ' ensure enough bytes exist
        bIconFile = True                    ' good, let's continue
        ' cache the ICONDIRECTORY
        CopyMemory m_icDir.idReserved, m_Bits(icPtr), icBytesNeed
        If m_icDir.idCount < 1 Then         ' no icons or not an icon file
            bIconFile = False
        ElseIf Not m_icDir.idReserved = 0 Then  ' per MSDN, must be zero
            bIconFile = False
        ElseIf m_icDir.idType < 1 Or m_icDir.idType > 2 Then
            bIconFile = False               ' per MSDN, must be 1 or 2 (1=icon,2=cursor)
        Else
            icPtr = icPtr + icBytesNeed     ' move array pointer
            icBytesNeed = 16&               ' length of directory entry
            If icPtr + icBytesNeed * m_icDir.idCount > streamLength Then
                bIconFile = False           ' not enough bytes for expected entries
            Else
                ReDim m_icDirE(1 To m_icDir.idCount)        ' size our entries
                icBytesNeed = m_icDir.idCount * icBytesNeed ' & cache them
                CopyMemory m_icDirE(1).bWidth, m_Bits(icPtr), icBytesNeed
                icBytesNeed = icBytesNeed + 6&       ' move array pointer
                For icEntry = 1 To m_icDir.idCount
                    ' each entry indicates how many bytes are used for it.
                    ' total the bytes and ensure enough bytes exist
                    icBytesNeed = icBytesNeed + m_icDirE(icEntry).dwBytesInRes
                Next
                If icBytesNeed > streamLength Then bIconFile = False ' not enough bytes
            End If
        End If
    End If
    
    If bIconFile Then
    
        ' Through experience, I have found the bitcount of the icons
        ' contained within the IconDirectoryEntry structures may be
        ' wrong or may not be filled in. Here, we will erase them & fill
        ' them in from the bitmap info headers that exist in the array.
        icBytesNeed = 0&
        For icEntry = 1 To m_icDir.idCount
            m_icDirE(icEntry).wPlanes = 1 ' not required, but used as a flag internally as indicating valid or invalid image
            ' get bitcount from the bitmap header
            CopyMemory icBytesNeed, m_Bits(m_icDirE(icEntry).dwImageOffset + 14), 2&
            
            If icBytesNeed = 0 Then ' if it is zero (shouldn't be); use the bitcount from the icon entry structure
                ' ensure the icon entry bitcount is not zero...
                If m_icDirE(icEntry).wBitCount = 0 Then
                    bIconFile = False
                Else
                    CopyMemory m_Bits(m_icDirE(icEntry).dwImageOffset + 14), m_icDirE(icEntry).wBitCount, 2&
                End If
                
            ElseIf icBytesNeed = 21060& Then
                ' flag for PNG, double check & abort if not -- then it is not an icon file
                bIconFile = ParsePNGheader(icEntry)
                
            Else    ' use the bitcount from the bitmap header
                m_icDirE(icEntry).wBitCount = icBytesNeed
            End If
        Next
    
        If bIconFile = True Then
            LoadStream = True

            If Not cHost Is Nothing Then
            
                ' appears we have a valid icon file. Find closest match for requested size
                If desiredWidth < 1 Then desiredWidth = 32& ' default if none provided
                If desiredHeight < 1 Then desiredHeight = 32&
                icEntry = GetBestMatch(desiredWidth, desiredHeight, icoBitDepth)
                
                If Not icEntry = 0 Then ' else something is wrong with the icon structure(s) in this file
                
                    If IsIconPNG(icEntry) Then ' png flag
                        
                        ' we need to pass this off to a PNG class for parsing/processing
                        Set cPNG = New cPNGparser
                        LoadStream = cPNG.LoadStream(inStream, cHost, m_icDirE(icEntry).dwImageOffset, m_icDirE(icEntry).dwBytesInRes, GlobalToken)
                        Set cPNG = Nothing
                        If Not cHost.Handle = 0& Then cHost.ImageType = imgPNGicon
                        
                    Else
                        ' create the main application's image, blank.
                        cHost.InitializeDIB Width(icEntry), Height(icEntry)
                        
                        ' copy the bitmap information header and fix it. Per MSDN, not all
                        ' members of the header are required to be filled in. We need them.
                        CopyMemory tBMPI.bmiHeader, m_Bits(m_icDirE(icEntry).dwImageOffset), 40&
                        With tBMPI.bmiHeader
                            .biClrUsed = ColorCount(icEntry)    ' fix when bitcount <= 8bpp
                            .biHeight = Height(icEntry)         ' height is doubled; fix it
                            .biSizeImage = 0                    ' erase; don't need this
                            .biXPelsPerMeter = 0                ' erase; don't need this
                            .biYPelsPerMeter = 0                ' erase; don't need this
                        End With
                        ' copy the fixed header back into the array
                        CopyMemory m_Bits(m_icDirE(icEntry).dwImageOffset), tBMPI.bmiHeader, 40&
                        
                        ' the next part of the routine is to create a compatible bitmap using
                        ' maximum screen colors on the system.  We will use the API to create it
                        ' for us from the bitmap header we just tweaked above. Otherwise we would
                        ' have to parse the bits ourselves, bloating code to handle 7 possible bit
                        ' depths in combination with several compression algorithms & various RGB masks.
                        tDC = GetDC(0&)
                        hDib = CreateDIBitmap(tDC, tBMPI.bmiHeader, 4, m_Bits(ColorByteOffset(icEntry)), m_Bits(m_icDirE(icEntry).dwImageOffset), 0&)
                        
                        If hDib = 0& Then
                            ReleaseDC 0&, tDC
                            cHost.DestroyDIB
                            ' major problem here; the icon contained in the stream appears to be faulty
                            ' we can't use it. Abort.
                        Else
                            ' here we are defining our application's image.
                            With tBMPI.bmiHeader
                                .biSize = 40&
                                .biBitCount = 32            ' 32bpp
                                .biHeight = cHost.Height    ' same width & height
                                .biWidth = cHost.Width      ' of the source image
                                .biPlanes = 1
                                .biSizeImage = .biHeight * .biWidth * 4&
                            End With
                            ' transfer the image bits from the bitmap created from the array to
                            ' our application's image
                            GetDIBits tDC, hDib, 0&, cHost.Height, ByVal cHost.BitsPointer, tBMPI, 0&
                            ReleaseDC 0&, tDC               ' release dc; don't need it any longer
                            DeleteObject hDib               ' kill the source bitmap; not needed
                            ApplyAlphaMask icEntry, cHost   ' add the alpha channel to app's image
                        End If
                    End If
                End If
            End If
        End If
    End If
    CopyMemory ByVal VarPtrArray(m_Bits), 0&, 4&    ' remove overlay
    
End Function

Public Function ConvertstdPicTo32bpp(Handle As Long, cHost As c32bppDIB) As Boolean

    ' Purpose: Convert a single icon from a stdPicture or handle to a 32bpp bitmap

    If Handle = 0& Then Exit Function
    
    Dim tSA As SafeArray
    Dim icoInfo As ICONINFO, tBMPI As BITMAPINFO
    Dim tBMPc As BITMAPINFO, tBMPm As BITMAPINFO
    Dim tDC As Long, hostDC As Long
    Dim X As Long, Y As Long, bAlpha As AlphaTypeEnum
    
    ' see if we can get the icon information
    If GetIconInfo(Handle, icoInfo) = 0& Then Exit Function
    
    m_icDir.idCount = 1
    m_icDir.idType = icoInfo.fIcon ' 0=icon, 1=cursor
    ReDim m_icDirE(1 To 1)  ' we will have 1 entry
    
    tDC = GetDC(0&)
    
    If Not icoInfo.hbmColor = 0& Then    ' do we have a color image? no for B&W
        tBMPc.bmiHeader.biSize = 40    ' let's fill in the BitmapInfo header
        If GetDIBits(tDC, icoInfo.hbmColor, 0&, 0&, ByVal 0&, tBMPc, 0&) = 0& Then
            m_icDir.idCount = 0 ' oops; something critical happened
        Else
            With tBMPI.bmiHeader    ' now fill in our destination description
                .biBitCount = 32
                .biHeight = tBMPc.bmiHeader.biHeight
                .biWidth = tBMPc.bmiHeader.biWidth
                .biPlanes = 1
                .biSize = 40&
                cHost.InitializeDIB .biWidth, .biHeight ' setup destination DIB
            End With
            ' use API again, to pass the bits from the color icon image to our DIB
            GetDIBits tDC, icoInfo.hbmColor, 0&, tBMPc.bmiHeader.biHeight, ByVal cHost.BitsPointer, tBMPI, 0&
        
            ' we will ensure the passed icon is not already a 32bpp ARGB image
            ' stdPictures won't be this way, but a call to LoadIconFromFile API can load XP icons
            With tSA
                .cbElements = 1
                .cDims = 2
                .pvData = cHost.BitsPointer
                .rgSABound(0).cElements = cHost.Height
                .rgSABound(1).cElements = cHost.scanWidth
            End With
            CopyMemory ByVal VarPtrArray(m_Bits), VarPtr(tSA), 4&
            m_icDirE(1).wBitCount = 1
            For Y = 0 To cHost.Height - 1
                For X = 3 To cHost.scanWidth - 1 Step 4
                    If Not m_Bits(X, Y) = 0 Then
                        m_icDirE(1).wBitCount = 32  ' looking for any non-zero alpha byte
                        Exit For
                    End If
                Next
                If m_icDirE(1).wBitCount = 32 Then Exit For
            Next
            If m_icDirE(1).wBitCount = 32 Then
                ' premultiply DIB as needed & set host imagetype, alpha properties
                iparseValidateAlphaChannel m_Bits, True, bAlpha, 0&
                If m_icDir.idType = 1 Then cHost.ImageType = imgIconARGB Else cHost.ImageType = imgCursorARGB
                cHost.Alpha = bAlpha
            End If
            CopyMemory ByVal VarPtrArray(m_Bits), 0&, 4&
        End If
    Else
        m_icDirE(1).wBitCount = 1   ' b&w icon/cursor
    End If
    
    ' pretty much same thing for the B&W, 1bpp mask
    ' Valid icons always have a mask, so no need to check .hbmMask=0 since this
    ' icon exists in a stdPicture and that, in itself, validated the icon for us
    If m_icDirE(1).wBitCount = 1 Then   ' else already processed as 32bpp icon/cursor
        If m_icDir.idCount = 1 Then
            tBMPm.bmiHeader.biSize = 40&
            If GetDIBits(tDC, icoInfo.hbmMask, 0, 0, ByVal 0&, tBMPm, 0) = 0 Then
                m_icDir.idCount = 0 ' oops; something critical happened
            Else
                With tBMPI.bmiHeader
                    If icoInfo.hbmColor = 0& Then
                        ' we have a b&w icon
                        .biBitCount = 32
                        .biHeight = tBMPm.bmiHeader.biHeight \ 2
                        .biWidth = tBMPm.bmiHeader.biWidth
                        .biPlanes = 1
                        .biSize = 40&
                        ' render the icon onto our dib
                        ' Note: in IDE, icon/cursor will be b&w, but when compiled
                        ' if the cursor had colors, the colors will be shown
                        cHost.InitializeDIB .biWidth, .biHeight  ' setup destination DIB
                        hostDC = cHost.LoadDIBinDC(True)
                        DrawIcon hostDC, 0, 0, Handle
                        cHost.LoadDIBinDC False
                        
                    End If
                    ' size our local array to hold the mask bits; these will be used
                    ' to tweak the 32bpp DIB's alpha channel in ApplyAlphaMask
                    ReDim m_Bits(0 To iparseByteAlignOnWord(1, .biWidth) * .biHeight - 1&)
                End With
                
                ' prepare bitmap info for our 1bpp mask array
                tBMPI.bmiPalette(1) = vbWhite
                With tBMPI.bmiHeader
                    .biBitCount = 1
                    .biClrUsed = 2
                End With
                ' use API again to pass the 1bpp image to our array
                GetDIBits tDC, icoInfo.hbmMask, 0&, tBMPI.bmiHeader.biHeight, m_Bits(0), tBMPI, 0&
                
                ' fill in the icon entry structure
                With m_icDirE(1)
                    .bHeight = tBMPI.bmiHeader.biHeight
                    .bWidth = tBMPm.bmiHeader.biWidth
                    .dwBytesInRes = UBound(m_Bits) + 1& ' we only have a mask in our array
                    .wBitCount = 1  ' the bitmap retrieved from the icon/cursor can be 32bpp
                    .wPlanes = 1    ' so we force the ApplyAlphaMask to use the 1bpp parsing routine
                End With
            
            End If
        End If
    End If
    ReleaseDC 0&, tDC
    
    ' clean up; GetIconInfo creates up to 2 bitmaps we must destroy
    If Not icoInfo.hbmColor = 0& Then DeleteObject icoInfo.hbmColor
    If Not icoInfo.hbmMask = 0& Then DeleteObject icoInfo.hbmMask
    
    If m_icDir.idCount = 1 Then           ' no errors encountered
        If m_icDirE(1).wBitCount = 1 Then ' now apply the mask
            ApplyAlphaMask 1&, cHost
            Erase m_Bits()
        End If
        ConvertstdPicTo32bpp = True
    End If
    
End Function

Private Sub ApplyAlphaMask(Index As Long, cHost As c32bppDIB)

    ' Purpose: Either blend or simulate transparency for icons
    ' The primary DIB for this application is 32bpp. Icons may or
    '   may not be 32bpp. When 32bpp, the icon RGB values are not
    '   pre-multiplied; so we need to pre-multiply them.  When
    '   the icon is not 32bpp, then it may have transparency,
    '   and we will modify our 32bpp image to identify which
    '   pixels are transparent and which are not.

    Dim dX As Long, X As Long, Y As Long, m As Long
    Dim aDIB() As Byte
    Dim Pow2(0 To 7) As Long
    Dim maskShift As Long, maskPtr As Long
    Dim maskScanWidth As Long, maskOffset As Long
    Dim bAlpha As AlphaTypeEnum
    
    Dim tSA As SafeArray
    With tSA                ' overlay the 32bpp dib
        .cbElements = 1     ' as bytes
        .cDims = 2          ' as 2D array
        .pvData = cHost.BitsPointer
        .rgSABound(0).cElements = cHost.Height
        .rgSABound(1).cElements = cHost.scanWidth
    End With
    CopyMemory ByVal VarPtrArray(aDIB), VarPtr(tSA), 4& ' establish overlay
    
    ' separate routines for 32bpp images and non-32bpp images.
    ' 32bpp images have the alpha mask combined with the RGB values. The
    ' transparency mask also exists but won't be used for 32bpp images.
    
    If m_icDirE(Index).wBitCount = 32 Then      ' alphablended icon
                                                ' get location of 1st color byte
        maskPtr = ColorByteOffset(Index) + 3&   ' then move to the alpha byte
        For Y = 0& To cHost.Height - 1&         ' loop thru scan lines
            For X = 0& To cHost.scanWidth - 1& Step 4&
                Select Case m_Bits(maskPtr)
                Case 0          ' 100% transparent
                    CopyMemory aDIB(X, Y), 0&, 4&
                Case 255        ' 100% opaque
                    aDIB(X + 3, Y) = 255
                Case Else       ' blend; calculation from MSDN
                    For dX = X To X + 2&
                        aDIB(dX, Y) = ((0& + m_Bits(maskPtr)) * aDIB(dX, Y)) \ &HFF
                    Next
                    aDIB(dX, Y) = m_Bits(maskPtr) ' keep the alpha byte value
                End Select
                maskPtr = maskPtr + 4&  ' move mask pointer to next alpha byte
            Next
        Next
        If m_icDir.idType = 1 Then cHost.ImageType = imgIconARGB Else cHost.ImageType = imgCursorARGB
        
    Else    ' 1,2,4,8,16,24 bpp images - not alpha blended, no alpha-channel
        
        Pow2(0) = 1&     ' build a power of two lookup table to parse the 1bpp mask
        For X = 1& To UBound(Pow2)
            Pow2(X) = Pow2(X - 1&) * 2&
        Next
        maskOffset = MaskByteOffset(Index)  ' location where mask starts
        maskScanWidth = iparseByteAlignOnWord(1, cHost.Width) ' how many mask bytes per scan line
        For Y = 0& To cHost.Height - 1&     ' loop thru the scan lines
            maskPtr = Y * maskScanWidth + maskOffset  ' adjust mask pointer per scan line
            maskShift = 7&                            ' bit position of mask
            dX = 3&
            ' note: do not loop thru using maskScanWidth. If the icon is a custom
            ' icon that has no DWORD aligned width, you will overflow the target
            ' DIB width and eventually write to uninitialized memory
            For X = 1& To cHost.Width
                If (m_Bits(maskPtr) And Pow2(maskShift)) = 0 Then ' is pixel transparent?
                    aDIB(dX, Y) = 255        ' nope, make it 100% opaque
                Else                        ' else make it 100% transparent
                    CopyMemory aDIB(dX - 3&, Y), 0&, 4&
                    bAlpha = AlphaComplex
                End If
                If maskShift = 0& Then   ' when we get to zero, the mask byte is read
                    maskShift = 7&       ' reset for next mask byte
                    maskPtr = maskPtr + 1& ' move to next mask byte
                Else
                    maskShift = maskShift - 1& ' adjust mask shifter
                End If
                dX = dX + 4&             ' move the 32bpp pointer along
            Next
        Next
        If m_icDir.idType = 1 Then cHost.ImageType = imgIcon Else cHost.ImageType = imgCursor
    End If
    iparseValidateAlphaChannel aDIB(), False, bAlpha, 0&
    CopyMemory ByVal VarPtrArray(aDIB), 0&, 4&  ' remove overlay
    cHost.Alpha = bAlpha
    
End Sub

Private Function GetBestMatch(cX As Long, cY As Long, icoBitDepth As Long) As Long

    ' Purpose: Find the nearest match to the passed Size.
    
    ' Note that this routine is weighted for monitors set at 32bit.
    ' If this is not acceptable, then modify algorithm slightly
    '   from adding weight of:  Abs(32 - bitDepth(icEntry))
    '   to adding weight of: Abs([ScreenColorDepth] - bitDepth(icEntry))
    
    ' additionally, the weighting is customized to favor larger icons over smaller ones
    ' when stretching would be needed. The thought is that stretching down almost always
    ' produces better quality graphics than stretching up.

    Dim Weights() As Long
    Dim icEntry As Long, bestMatch As Long
    Dim lWeight As Long
    
    If m_icDir.idCount > 1 Then ' more than one icon?
    
        ReDim Weights(-1 To m_icDir.idCount)
        ' set least desirable weight: some large number
        Weights(0) = 10000&
        
        For icEntry = 1 To m_icDir.idCount
            ' simple weight; use the difference between desired size & icon size
            If Not m_icDirE(icEntry).wBitCount = 0 Then     ' if a image within icon file is faulty, we ignore it
                
                lWeight = Width(icEntry) - cX ' & penalize if stretching larger is needed
                If cX > Width(icEntry) Then lWeight = lWeight * 2&
                Weights(icEntry) = lWeight
                
                lWeight = Height(icEntry) - cY ' & penalize if stretching larger is needed
                If cY > Height(icEntry) Then lWeight = lWeight * 2&
                Weights(icEntry) = Weights(icEntry) + lWeight
                
                ' add the weight for bit depth
                Weights(icEntry) = Weights(icEntry) + Abs(icoBitDepth - bitDepth(icEntry))
                
                If m_icDirE(icEntry).wBitCount > 32 Then Weights(icEntry) = -10000&  ' if future icons are something like 48bpp
                
                ' compare; one with lowest value wins
                If Weights(icEntry) = 0 Then
                    bestMatch = icEntry
                    Exit For
                ElseIf Weights(icEntry) < Weights(0) Then
                    If Weights(icEntry) > 0 Then        ' basically rejects icons that need to be stretched up
                        Weights(0) = Weights(icEntry)
                        bestMatch = icEntry
                    End If
                End If
            End If
        Next
        If bestMatch = 0 Then ' every image is too small and must be stretched. We will get the highest negative value now
            For icEntry = icEntry - 1& To 1& Step -1&
                Weights(icEntry) = Abs(Weights(icEntry)) + Abs(32& - bitDepth(icEntry))
                If Weights(icEntry) < Weights(0) Then
                    Weights(0) = Weights(icEntry)
                    bestMatch = icEntry
                End If
                If bestMatch = 0& Then bestMatch = 1&
            Next
        End If
        
    Else ' only one icon/PNG
        If m_icDirE(1).wBitCount = 0 Then bestMatch = 0& Else bestMatch = 1&
    
    End If
    GetBestMatch = bestMatch
    
End Function

Private Function ParsePNGheader(Index As Long) As Boolean

    ' PNG's IHDR structure
    '    Width As Long              << cannot be negative
    '    Height As Long             << cannot be negative
    '    BitDepth As Byte           << must be 1,2,4,8,16
    '    ColorType As Byte          << must be 0,2,3,4,6
    '    Compression As Byte        << must be zero
    '    Filter As Byte             << must be zero
    '    Interlacing As Byte        << must be zero or one
    
    Dim lValue As Long, Offset As Long
    Const chnk_IHDR As Long = &H52444849 'Image header PNG flag
    
    On Error GoTo ExitRoutine:
    ' get the image width; the value will be a reversed long
    With m_icDirE(Index)
        
        .wPlanes = 255 ' flag for png
        
        ' verify this is a png signture
        CopyMemory lValue, m_Bits(m_icDirE(Index).dwImageOffset), 4&
        If lValue = png_Signature1 Then ' probably a png (Vista Icon)
            ' the 1st 4 bytes were verified, verify next 4 bytes
            CopyMemory lValue, m_Bits(m_icDirE(Index).dwImageOffset + 4), 4&
            If lValue = png_Signature2 Then  ' definitely a png (Vista Icon)
        
            ' If this is a valid PNG, the next 4 bytes would be 13 (size of header)
            ' and the following 4 bytes would be the header name (chnk_IHDR)
            CopyMemory lValue, m_Bits(.dwImageOffset + 12&), 4&
            
                If lValue = chnk_IHDR Then
                
                    ' get PNG's width
                    CopyMemory lValue, m_Bits(.dwImageOffset + 16&), 4&
                    lValue = iparseReverseLong(lValue)
                    Select Case lValue
                        Case 256: .bWidth = 0&
                        Case 1 To 255: .bWidth = lValue
                        Case Else: .wBitCount = 0& ' prevent processing PNG as an option
                    End Select
                    
                    ' do the same for the height
                    CopyMemory lValue, m_Bits(.dwImageOffset + 20&), 4&
                    lValue = iparseReverseLong(lValue)
                    Select Case lValue
                        Case 256: .bHeight = 0&
                        Case 1 To 255: .bHeight = lValue
                        Case Else: .wBitCount = 0& ' prevent processing PNG as an option
                    End Select
                
                    If .wBitCount = 0 Then
                        .wBitCount = m_Bits(.dwImageOffset + 24&)
                        If .wBitCount = 16 Then
                            .wBitCount = 32 ' for our purposes a 48bpp image is a 32bpp image
                            
                        ElseIf Not .wBitCount = 0 Then
                            Select Case m_Bits(.dwImageOffset + 25&)
                            Case 4, 6: .wBitCount = 32  ' alpha png
                            Case 2: .wBitCount = 24     ' true color
                            Case Else                   ' no change in interpretation
                            End Select
                        End If
                    End If
                    
                    ' the remaining bytes of the IHDR are not needed for the icon class
                    ParsePNGheader = (.wBitCount > 0)
            
                End If
            End If
        End If
    End With

ExitRoutine:
If Err Then
    Err.Clear
    m_icDirE(Index).wBitCount = 0
End If
End Function

