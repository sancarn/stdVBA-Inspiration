VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Process/Render Images using GDI+ ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._


' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take some liberties in several API declarations throughout

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

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long

' following are used for saving dib to PNG (testing phase only)
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, gdiBitmapData As Any, BITMAP As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum
Private Type EncoderParameter
    GUID(0 To 3)   As Long
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type
'-- Encoder Parameters structure
Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type
Private Type ImageCodecInfo
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, ByRef clrMatrix As Any, ByRef grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imgAttr As Long) As Long
Private Const ColorAdjustTypeBitmap As Long = 1

Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const PixelFormat32bppPARGB As Long = &HE200B
Private Const InterpolationModeNearestNeighbor As Long = &H5&
Private Const InterpolationModeHighQualityBicubic As Long = &H7&
Private Const InterpolationModeHighQualityBilinear As Long = &H6&

' Following are used only if PNG file is being processed by GDI+
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, hImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As Long) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, PixelFormat As Long) As Long
Private Const UnitPixel As Long = &H2&
Private Type RECTF
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type

' used for workaround of VB not exposing IStream interface
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Private m_hImage As Long    ' cached when RenderGDIplus is called
Private m_Token As Long     ' cached whenever GDI+ function is called
Private m_Attr As Long      ' current image attributes: mirroring, etc
' m_Token is destroyed when class terminates

Public Function isGDIplusOk(Optional gdiToken As Long, Optional KeepToken As Boolean = False) As Boolean

    ' Function starts GDI+ and returns true if no errors occurred
    
    ' does the system even have GDI+ on it?
    If m_Token = 0& Then
        If iparseValidateDLL("gdiplus.dll", "GdiplusStartup") Then
            ' it does, so attempt to start GDI+
            isGDIplusOk = InitializeGDIplus(gdiToken, False)
            If Not KeepToken Then InitializeGDIplus gdiToken, True    ' shut it down
        End If
    Else
        isGDIplusOk = True
    End If
End Function

Public Function SaveToPNG(FileName As String, outStream() As Byte, cHost As c32bppDIB, Optional ByVal GlobalToken As Long) As Boolean

    ' Function uses GDI+ to create a PNG file or stream
    ' Parameters:
    ' FileName. If PNG is to be saved to file, provide the file name, otherwise PNG will be saved to array
    ' outStream(). If FileName is vbNullString, then PNG is saved to this array, zero-bound
    ' cHost. The c32bppDIB class containing the image to convert to PNG

    If cHost.Handle = 0& Then Exit Function
    ' does the system have GDI+ on it?
    If GlobalToken = 0& Then
        If m_Token = GlobalToken Then
            If iparseValidateDLL("gdiplus.dll", "GdiplusStartup") = False Then Exit Function
            If InitializeGDIplus(m_Token, False) = False Then Exit Function
        End If
    End If
    
    Dim uEncCLSID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    ' Note: GdipCreateBitmapFromGdiDib does not handle 32bpp DIBs correctly
    If m_hImage = 0& Then
        Call GdipCreateBitmapFromScan0(cHost.Width, cHost.Height, cHost.scanWidth, PixelFormat32bppPARGB, ByVal cHost.BitsPointer, m_hImage)
    End If
    
    If Not m_hImage = 0& Then
        On Error Resume Next
        ' retrieve information GDI+ will use for conversion
        If Not pvGetEncoderClsID("image/png", uEncCLSID) = -1 Then
            ' dib is bottom up, scan0 does top down, so flip it
            Call ModifyImageAttributes(6)
            
            If FileName = vbNullString Then
                ' Saving to stream/array. Create a null stream (IUnknown object)
                Erase outStream
                Set IIStream = CreateStream(outStream)
                ' have GDI+ save the 32bpp image to the IUnknown in a PNG format
                If GdipSaveImageToStream(m_hImage, IIStream, uEncCLSID(0&), ByVal 0&) = 0& Then
                    ' now we need to get that array to pass back to client
                    ArrayFromStream IIStream, outStream()
                    SaveToPNG = True
                End If
            Else    ' saving to file
                ' Note: If you are calling this from outside the c32bppDIB class, the file
                ' must not exist; otherwise, the function fails.
                SaveToPNG = (GdipSaveImageToFile(m_hImage, StrPtr(FileName), uEncCLSID(0&), ByVal 0&) = 0&)
            End If
        End If
    End If
    
End Function

Public Function SaveToJPG(FileName As String, outStream() As Byte, cHost As c32bppDIB, Optional ByVal Quality As Long = 80, Optional ByVal TransFillColor As Long = vbBlack, Optional ByVal GlobalToken As Long) As Boolean

    ' Function uses GDI+ to create a JPG file or stream
    ' Parameters:
    ' FileName. If JPG is to be saved to file, provide the file name, otherwise JPG will be saved to array
    ' outStream(). If FileName is vbNullString, then JPG is saved to this array, zero-bound
    ' cHost. The c32bppDIB class containing the image to convert to JPG

    If cHost.Handle = 0& Then Exit Function
    ' does the system have GDI+ on it?
    If GlobalToken = 0 Then
        If m_Token = GlobalToken Then
            If iparseValidateDLL("gdiplus.dll", "GdiplusStartup") = False Then Exit Function
            If InitializeGDIplus(m_Token, False) = False Then Exit Function
        End If
    End If
    
    Dim uEncCLSID(0 To 3) As Long
    Dim aEncParams() As Byte
    Dim IIStream As IUnknown
    Dim uEncParams As EncoderParameters
    Dim tSA As SafeArray, dibBytes() As Byte
    Dim hasAlpha As AlphaTypeEnum, oldAttr As Long, oldImage As Long
    
    Dim X As Long, Y As Long, newDib() As Byte
    Dim dAlpha As Long, pAlpha As Byte
    
    Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    
    ' Note: GdipCreateBitmapFromGdiDib does not handle 32bpp DIBs correctly
    iparseOverlayHost_2DbyHost dibBytes(), VarPtr(tSA), cHost
    iparseValidateAlphaChannel dibBytes(), False, hasAlpha, 0&
    iparseOverlayHost_2DbyHost dibBytes(), 0&, cHost
    
    If m_hImage = 0& Or hasAlpha > AlphaNone Then
        If hasAlpha Then
            oldImage = m_hImage
            oldAttr = m_Attr
            m_Attr = 0&
            iparseBlendImageToColor cHost, TransFillColor, dibBytes()
            Call GdipCreateBitmapFromScan0(cHost.Width, cHost.Height, cHost.scanWidth, PixelFormat32bppPARGB, dibBytes(0, 0), m_hImage)
        Else
            Call GdipCreateBitmapFromScan0(cHost.Width, cHost.Height, cHost.scanWidth, PixelFormat32bppPARGB, ByVal cHost.BitsPointer, m_hImage)
        End If
    End If
    
    If Not m_hImage = 0& Then
        On Error Resume Next
        ' retrieve information GDI+ will use for conversion

        If Not pvGetEncoderClsID("image/jpeg", uEncCLSID) = -1 Then
            '-- Set encoder params. (Quality)
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                Call CLSIDFromString(StrPtr(EncoderQuality), .GUID(0))
                If Quality < 30 Then
                    Quality = 30
                ElseIf Quality > 100 Then
                    Quality = 100
                End If
                .Value = VarPtr(Quality)
            End With
            Call CopyMemory(aEncParams(1), uEncParams, Len(uEncParams))
            Call ModifyImageAttributes(6)
            
            If FileName = vbNullString Then
                Erase outStream
                Set IIStream = CreateStream(outStream)
                If GdipSaveImageToStream(m_hImage, IIStream, uEncCLSID(0&), aEncParams(1)) = 0& Then
                    ' now we need to get that array to pass back to client
                    ArrayFromStream IIStream, outStream()
                    SaveToJPG = True
                End If
            Else ' save the JPG file first
                ' Note: If you are calling this from outside the c32bppDIB class, the file
                ' must not exist; otherwise, the function fails.
                SaveToJPG = (GdipSaveImageToFile(m_hImage, StrPtr(FileName), uEncCLSID(0), aEncParams(1)) = 0&)
            End If
        End If
    End If
    If hasAlpha Then
        GdipDisposeImage m_hImage
        m_hImage = oldImage
        m_Attr = oldAttr
        Erase dibBytes()
    End If

End Function

Private Function pvPtrToStrW(ByVal lpsz As Long) As String
  ' supporting routine for SaveToPNG; converts String Pointer to String
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0&) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2&)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
    
End Function

Private Function pvGetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
  ' supporting routine for SaveToPNG
  Dim Num      As Long
  Dim Size     As Long
  Dim LIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0&) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For LIdx = 1& To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(LIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            CopyMemory ClassID(0), ICI(LIdx).ClassID(0), 16& ' Save the Class ID
            pvGetEncoderClsID = LIdx      ' Return the index number for success
            Exit For
        End If
    Next LIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Friend Function RenderGDIplus(cHost As c32bppDIB, ByVal hDC As Long, _
                            ByVal Angle As Single, ByVal Alpha As Long, _
                            ByVal destX As Long, ByVal destY As Long, _
                            ByVal destWidth As Long, ByVal destHeight As Long, _
                            ByVal SrcX As Long, ByVal SrcY As Long, _
                            ByVal srcWidth As Long, ByVal srcHeight As Long, _
                            ByVal highQuality As Boolean, _
                            ByVal grayScale As eGrayScaleFormulas, _
                            ByVal GlobalToken As Long, _
                            Optional ByVal LightnessAdj As Single = 0!) As Boolean

    ' Function renders a 32bpp to passed DC.
    ' GDI+ can literally do most anything with an image; just gotta know how to set it up
    
    ' Parameters
    ' c32bppDIB. Class containing image to render
    ' hDC. The destination DC to render to
    ' Angle. A value between -360 and 360 used for rotation. 0 is no rotation
    ' Alpha. A value between 0 and 100 used for global tranparency. 100 is fully opaque
    ' destX,Y. The top,left corner of the DC to render the image to
    ' destWidth,Height. The target size of the rendered image
    ' srcX,Y. The top,left corner of the image to be rendered
    ' srcWidth,Height. The size of the source to be rendered
    ' highQuality. If true, then BiCubic interpolation will be used, else NearestNeighbor will be used
    ' grayScale. One of the eGrayScaleFormulas
    ' GlobalToken. When provided it is a valid GDI token
    ' LigthnessAdj. Percentage (-100 to 100) of more/less lightness for the image
    
    If Alpha = 0& Then
        RenderGDIplus = True    ' full transparent, nothing to render
        Exit Function
    End If
    
    If GlobalToken = 0& Then
        If m_Token = GlobalToken Then
            If InitializeGDIplus(m_Token, False) = False Then Exit Function
        End If
    End If

    Dim hGraphics As Long, hImgAttr As Long
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim mirrorROP As Long
    
    ' have GDI+ create a DIB from our host pointer, DIB will be mirrored vertically (upside down)
    If m_hImage = 0& Then
        Call GdipCreateBitmapFromScan0(cHost.Width, cHost.Height, cHost.scanWidth, PixelFormat32bppPARGB, ByVal cHost.BitsPointer, m_hImage)
    End If
    
    If Not m_hImage = 0& Then
        
        If GdipCreateFromHDC(hDC, hGraphics) = 0& Then   ' wrap GDI+ around our target DC
            
            If Not hGraphics = 0& Then
                
                ' Interpolation quality?
                If highQuality = True Then
                    Call GdipSetInterpolationMode(hGraphics, InterpolationModeHighQualityBicubic)
                Else ' Note: There is a 3rd quality which falls between these: InterpolationModeHighQualityBilinear
                    Call GdipSetInterpolationMode(hGraphics, InterpolationModeNearestNeighbor)
                End If
                
                ' calculate flags/offsets if we are mirroring and/or rotating
                If destHeight < 0& Then
                    destHeight = -destHeight               ' no flipping needed; bottom up dibs are flipped vertically naturally
                Else
                    mirrorROP = 6&                         ' flip vertically (mirror vertically)
                End If
                If destWidth < 0& Then
                    mirrorROP = mirrorROP Xor 4&           ' flip horizontally (mirror horizontally)
                    destWidth = -destWidth
                End If
                
                Call ModifyImageAttributes(mirrorROP)     ' apply attributes to the image as needed
                
                If Not ((grayScale = gsclNone) And (LightnessAdj = 0!)) Then
                    ' grayscaling is in play
                    If GdipCreateImageAttributes(hImgAttr) = 0 Then
                        If Not grayScale = gsclNone Then
                            Call iparseGrayScaleRatios(grayScale, clrMatrix(0, 0), clrMatrix(0, 1), clrMatrix(0, 2))
                            clrMatrix(1, 0) = clrMatrix(0, 0)
                            clrMatrix(2, 0) = clrMatrix(0, 0)
                            clrMatrix(1, 1) = clrMatrix(0, 1)
                            clrMatrix(2, 1) = clrMatrix(0, 1)
                            clrMatrix(1, 2) = clrMatrix(0, 2)
                            clrMatrix(2, 2) = clrMatrix(0, 2)
                        Else
                            clrMatrix(0, 0) = 1
                            clrMatrix(1, 1) = 1
                            clrMatrix(2, 2) = 1
                        End If
                        clrMatrix(3, 3) = 1 ' global alpha value
                        clrMatrix(4, 4) = 1 ' required; cannot be anything else
                        If Not LightnessAdj = 0! Then
                            clrMatrix(0, 4) = LightnessAdj / 100    ' red added/subtracted brightness
                            clrMatrix(1, 4) = clrMatrix(0, 4)       ' same for blue
                            clrMatrix(2, 4) = clrMatrix(0, 4)       ' same for green
                        End If
                        If Not GdipSetImageAttributesColorMatrix(hImgAttr, ColorAdjustTypeBitmap, -1&, clrMatrix(0, 0), clrMatrix(0, 0), 0&) = 0& Then
                            GdipDisposeImageAttributes hImgAttr
                            hImgAttr = 0&
                        End If
                    End If
                End If
                
                If Angle = 0& And Alpha = 100& Then ' no blending and no rotation being used

                    RenderGDIplus = (GdipDrawImageRectRectI(hGraphics, m_hImage, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&) = 0&)
                
                Else ' we are blending and/or rotating
                    
                    If hImgAttr = 0& Then ' else grayscaling also & hImagAttr already created
                        If GdipCreateImageAttributes(hImgAttr) = 0& Then ' create image attributes for blending/rotating
                            clrMatrix(0, 0) = 1
                            clrMatrix(1, 1) = 1
                            clrMatrix(2, 2) = 1
                            clrMatrix(4, 4) = 1 ' required; cannot be anything else
                        End If
                        If Not LightnessAdj = 0! Then
                            clrMatrix(0, 4) = LightnessAdj / 100!   ' red added/subtracted brightness
                            clrMatrix(1, 4) = clrMatrix(0, 4)       ' same for blue
                            clrMatrix(2, 4) = clrMatrix(0, 4)       ' same for green
                        End If
                    End If
                    ' Global Blending?
                    clrMatrix(3, 3) = CSng(Alpha / 100&) ' value between 0 & 1
                    
                    If GdipSetImageAttributesColorMatrix(hImgAttr, ColorAdjustTypeBitmap, -1&, clrMatrix(0, 0), clrMatrix(0, 0), 0&) = 0& Then
                        If Angle = 0& Then   ' not rotating
                            RenderGDIplus = (GdipDrawImageRectRectI(hGraphics, m_hImage, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&) = 0&)
                            
                        Else ' rotating & maybe blending too... different coordinates system used when rotating
                        
                            If GdipRotateWorldTransform(hGraphics, Angle + 180, 0&) = 0& Then
                                GdipTranslateWorldTransform hGraphics, destX + (destWidth \ 2), destY + (destHeight \ 2), 1&
                            End If
                            RenderGDIplus = (GdipDrawImageRectRectI(hGraphics, m_hImage, destWidth \ 2, destHeight \ 2, -destWidth, -destHeight, SrcX, SrcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&) = 0&)
                        End If
                    End If
                End If
                    
                If Not hImgAttr = 0& Then GdipDisposeImageAttributes hImgAttr ' clean up
                GdipDeleteGraphics hGraphics ' clean up
                
            End If
        End If
        
    End If

End Function

Friend Function GDIplusLoadPNG(FileName As String, pngStream() As Byte, cHost As c32bppDIB, Optional ByVal GlobalToken As Long) As Boolean
'Exit Function      ' un-rem to test/force PNG loading without GDI+

    ' Purpose: Use GDI+ to load a PNG either by fileName or by array/stream
    ' FileName :: if vbNullString, then the pngStream() array will contain
    '             the PNG else FileName is full path & name of the PNG file
    ' Note: FileName and/or pngStream() have been validated before this routine is called
    
    ' does the system have GDI+ on it?
    If GlobalToken = 0& Then
        If m_Token = GlobalToken Then
            If iparseValidateDLL("gdiplus.dll", "GdiplusStartup") = False Then Exit Function
            If InitializeGDIplus(m_Token, False) = False Then Exit Function
        End If
    End If
     
    Dim hImage As Long, hGraphics As Long
    Dim tDC As Long, lRtn As Long
    Dim rDimensions As RECTF, pStream As IUnknown
    
    On Error GoTo ExitRoutine
    
    If FileName = vbNullString Then ' we need an array vs file name
        ' hack of my own. Create an IUnknown Stream that has the same properties
        ' and minimum methods needed as the IStream interface which VB does not
        ' expose. Once the stream is created, we have GDI+ load from it
        Set pStream = CreateStream(pngStream())
        If Not pStream Is Nothing Then Call GdipLoadImageFromStream(pStream, hImage)

    Else    ' we use the passed file name; have GDI+ load the file
        Call GdipLoadImageFromFile(StrPtr(FileName), hImage)
    End If
    
    If Not hImage = 0& Then
        ' get size of PNG
        lRtn = GdipGetImageBounds(hImage, rDimensions, UnitPixel)
        If lRtn = 0& Then

            ' build 32bpp
            cHost.InitializeDIB CLng(rDimensions.nWidth), CLng(rDimensions.nHeight)

            ' wrap a GDI+ DC around our DIB's DC
            tDC = cHost.LoadDIBinDC(True)
            lRtn = GdipCreateFromHDC(tDC, hGraphics)
            If lRtn = 0& Then
                ' now draw the PNG into our 32bpp. GDI+ is nice enough to pre-multiply
                ' the RGB values for us during the rendering
                With rDimensions
                    GdipDrawImageRectRectI hGraphics, hImage, 0&, 0&, .nWidth, .nHeight, .nLeft, .nTop, .nWidth, .nHeight, UnitPixel, 0&, 0&, 0&
                End With
                GdipDeleteGraphics hGraphics    ' remove the GDI+ DC wrapper
                hGraphics = 0&
            End If
            cHost.LoadDIBinDC False ' unselect our DIB
        End If

        If lRtn = 0& Then                    ' return results
            Dim b() As Byte, oSA As SafeArray, bAlpha As AlphaTypeEnum
            GDIplusLoadPNG = True
            'Call GdipGetImagePixelFormat(hImage, lRtn)
            'cHost.Alpha = (lRtn = PixelFormat32bppARGB Or lRtn = PixelFormat32bppPARGB)
            '^^ PixelFormat32bppARGB will return True for 32bpp PNG & entire alpha channel=255
            ' so we will look ourselves
            GdipDisposeImage hImage             ' destroy the GDI+ image
            iparseOverlayHost_2DbyHost b(), VarPtr(oSA), cHost
            iparseValidateAlphaChannel b(), False, bAlpha, 0
            iparseOverlayHost_2DbyHost b(), 0&, Nothing
            cHost.Alpha = bAlpha
            cHost.ImageType = imgPNG
        Else
            GdipDisposeImage hImage             ' destroy the GDI+ image
            cHost.DestroyDIB
        End If
        hImage = 0&
    End If
        
ExitRoutine:
    If Not hGraphics = 0& Then GdipDeleteGraphics hGraphics
    If Not hImage = 0& Then GdipDisposeImage hImage
End Function

Private Function ArrayFromStream(Stream As IUnknown, arrayBytes() As Byte) As Boolean

    ' Purpose: Return the array contained in an IUnknown interface
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If Not Stream Is Nothing Then
    
        If GetHGlobalFromStream(ByVal ObjPtr(Stream), o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    ArrayFromStream = True
                End If
            End If
        End If

    End If
    
End Function

Private Function CreateStream(byteContent() As Byte, Optional byteOffset As Long = 0&) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_lngLowerBound As Long
    Dim o_lngByteCount  As Long
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If iparseIsArrayEmpty(VarPtrArray(byteContent)) = 0& Then ' create a growing stream as needed
         Call CreateStreamOnHGlobal(0, 1, CreateStream)
    Else                                        ' create a fixed stream
         o_lngByteCount = UBound(byteContent) - byteOffset + 1
         o_hMem = GlobalAlloc(&H2&, o_lngByteCount)
         If o_hMem <> 0 Then
             o_lpMem = GlobalLock(o_hMem)
             If o_lpMem <> 0 Then
                 CopyMemory ByVal o_lpMem, byteContent(byteOffset), o_lngByteCount
                 Call GlobalUnlock(o_hMem)
                 Call CreateStreamOnHGlobal(o_hMem, 1, CreateStream)
             End If
         End If
     End If
    
HandleError:
End Function

Private Sub ModifyImageAttributes(newAttributes As Long)

    ' Function is used to apply and/or remove attributes from the m_hImage object
    
    If newAttributes = m_Attr Then Exit Sub         ' nothing to do
    If m_hImage = 0& Then Exit Sub                  ' no image
    
    If (m_Attr And newAttributes) = 0& Then         ' current attributes don't contain any of the new attributes
        If m_Attr Then GdipImageRotateFlip m_hImage, m_Attr ' remove those
        GdipImageRotateFlip m_hImage, newAttributes ' apply new attriubtes
    
    Else                                            ' current attributes have the new attributes and more
        m_Attr = (m_Attr And Not newAttributes)     ' remove the additional attributes
        GdipImageRotateFlip m_hImage, m_Attr        ' and apply to the image
    End If
    
    m_Attr = newAttributes
    
End Sub

Friend Function InitializeGDIplus(gToken As Long, ShutDown As Boolean) As Boolean
    
    ' function starts/stops GDI+
    On Error Resume Next
    If ShutDown Then
        If m_hImage Then GdipDisposeImage m_hImage ' clean up
        If Not gToken = 0& Then GdiplusShutdown gToken
        m_Attr = 0&
    Else
    
        Dim gdiSI As GdiplusStartupInput
        gdiSI.GdiplusVersion = 1
        If GdiplusStartup(gToken, gdiSI) = 0& Then
            InitializeGDIplus = Not (gToken = 0&)
        Else
            gToken = 0&
        End If
     
    End If
    If Err Then Err.Clear
    
End Function

Private Sub Class_Terminate()
    ' shut down GDI+ and destroy cached token & hImage if necessary
    InitializeGDIplus m_Token, True
End Sub


