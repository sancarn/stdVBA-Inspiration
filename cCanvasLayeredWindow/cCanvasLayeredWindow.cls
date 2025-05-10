VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCanvasLayeredWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Type POINTAPI
    X                        As Long
    Y                        As Long
End Type

Private Type Size
    cx                       As Long
    cy                       As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                   As Long
    biWidth                  As Long
    biHeight                 As Long
    biPlanes                 As Integer
    biBitCount               As Integer
    biCompression            As Long
    biSizeImage              As Long
    biXPelsPerMeter          As Long
    biYPelsPerMeter          As Long
    biClrUsed                As Long
    biClrImportant           As Long
End Type

Private Type BITMAPINFO
    bmiHeader                As BITMAPINFOHEADER
    bmiColors(3)             As Byte
End Type

Private Type BLENDFUNCTION
    BlendOp                  As Byte
    BlendFlags               As Byte
    SourceConstantAlpha      As Byte
    AlphaFormat              As Byte
End Type

Private Const ULW_ALPHA         As Long = &H2
Private Const BI_RGB            As Long = 0&
Private Const DIB_RGB_COLORS    As Long = 0&
Private Const AC_SRC_ALPHA      As Long = &H1

Private Const GWL_EXSTYLE       As Long = -20
Private Const WS_EX_TOPMOST     As Long = &H8&
Private Const WS_EX_LAYERED     As Long = &H80000

Private c_lhDC          As Long
Private c_lDIB          As Long
Private m_Pointer       As Long
Private m_OldBmp        As Long
Private tSIZE           As Size

Private tBLENDFUNCTION  As BLENDFUNCTION


' for save as jpg
Private Type GUID
    Data1          As Long
    Data2          As Integer
    Data3          As Integer
    Data4(0 To 7)  As Byte
End Type
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID           As GUID
    NumberOfValues As Long
    Type           As Long
    Value          As Long
End Type
Private Type EncoderParameters
    Count          As Long
    Parameter      As EncoderParameter
End Type
    
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long


' for save as png
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, RECT As RECT, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    scan0 As Long
    Reserved As Long
End Type

Private Type BitmapFileHeader
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
    
Public Sub saveimage(ByVal filename As String, saveformat As SavePictureFormats, Optional jpgqualtiy As Long = 100)

'Dim lBufferIn() As Byte
'Dim m_GDIplus As New cGDIplusBase
'Dim hImage As Long, hGraphics As Long
'hImage = m_GDIplus.CreateImageFromScan0(0, 0, 321, 571, Color32bppAlpha)
'hGraphics = m_GDIplus.CreateGraphicsFromHIMAGE(hImage)
'MsgBox m_GDIplus.SaveAsPNG(c_lDIB, lBufferIn())
'MsgBox UBound(lBufferIn)
'm_GDIplus.DisposeGraphics hGraphics
'm_GDIplus.DisposeImage hImage
'
'm_GDIplus.CreateGraphicsFromHIMAGE
Dim myc As New PicSave
    myc.SavePicture c_lDIB, filename, saveformat, jpgqualtiy
    Set myc = Nothing
End Sub

    
''    '���� ����� ����
''    Public Sub SaveJPG(ByVal FileName As String, Optional ByVal quality As Long)
''        Dim tSI        As GdiplusStartupInput
''        Dim lRes       As Long
''        Dim lGDIP      As Long
''        Dim lBitmap    As Long
''
''        ' Initialize GDI+
''        tSI.GdiplusVersion = 1
''        lRes = GdiplusStartup(lGDIP, tSI)
''
''        If lRes = 0 Then
''            ' Create the GDI+ bitmap from the image handle
''            lRes = GdipCreateBitmapFromHBITMAP(c_lDIB, 0, lBitmap)
''            If lRes = 0 Then
''                Dim tJpgEncoder As GUID
''                Dim tParams As EncoderParameters
''                ' Initialize the encoder GUID
''                CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
''                ' Initialize the encoder parameters
''                tParams.Count = 1
''                With tParams.Parameter    ' Quality
''                    ' Set the Quality GUID
''                    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
''                    .NumberOfValues = 1
''                    .Type = 4
''                    .Value = VarPtr(quality)
''                End With
''
''                ' Save the image
''                lRes = GdipSaveImageToFile( _
''                lBitmap, _
''                StrPtr(FileName), _
''                tJpgEncoder, _
''                tParams)
''                ' Destroy the bitmap
''                GdipDisposeImage lBitmap
''            End If
''            ' Shutdown GDI+
''            GdiplusShutdown lGDIP
''        End If
''        If lRes Then
''            Err.Raise 5, , "Cannot save the image. GDI+ Error:" & lRes
''        End If
''    End Sub
''
''    '���� ����� ����
''    Public Sub Savepng(ByVal FileName As String)      '= 80
''                Dim lBufferIn() As Byte
''                Dim m_GDIplus As New cGDIplusBase
''                Dim hImage As Long, hGraphics As Long
''                'hImage = m_GDIplus.CreateImageFromScan0(0, 0, 321, 571, Color32bppAlpha)
''                hGraphics = m_GDIplus.CreateGraphicsFromHDC(c_lhDC)
''                MsgBox m_GDIplus.SaveAsPNG(hGraphics, lBufferIn())
''                MsgBox UBound(lBufferIn)
''                m_GDIplus.DisposeGraphics hGraphics
''                m_GDIplus.DisposeImage hImage
''
'''                hImage = m_GDIplus.CreateImageFromScan0(0, 0, 100, 100, Color32bppAlpha)
'''
'''
'''                'Const UnitPixels As Long = 2
'''
'''                'hGraphics = m_GDIplus.CreateGraphicsFromHIMAGE(hImage)
'''                hGraphics = m_GDIplus.CreateGraphicsFromHDC(hdc)
'''
'''                '                Dim lWidth As Single
'''                'Dim lHeight As Single
'''                'GdipGetImageDimension hImage, lWidth, lHeight
'''
'''
'''                Dim lBufferIn() As Byte
'''                m_GDIplus.SaveAsPNG hGraphics, lBufferIn()
'''
'''                    Open App.Path & "\test.png" For Binary As #1
'''    Put #1, 1, lBufferIn()
'''    Close #1
'''    m_GDIplus.DisposeGraphics hGraphics
'''    Erase lBufferIn()
'''
'''   '             m_GDIplus.DisposeImage hImage
'''
'''
'''        Exit Sub
''
''''        Dim tSI        As GdiplusStartupInput
''''        Dim lRes       As Long
''''        Dim lGDIP      As Long
''''        Dim lBitmap    As Long
''''
''''
''''        Dim lrect As Rect
''''        Dim lWidth As Single
''''        Dim lHeight As Single
''''        Const PixelFormat32bppARGB = &H26200A
''''        Dim bmpData As BitmapData
''''        Dim lBufferIn() As Byte
''''        Dim lX As Long, ly As Long
''''
''''
''''
''''        ' Initialize GDI+
''''        tSI.GdiplusVersion = 1
''''        lRes = GdiplusStartup(lGDIP, tSI)
''''
''''        If lRes = 0 Then
''''            ' Create the GDI+ bitmap from the image handle
''''            lRes = GdipCreateBitmapFromHBITMAP(c_lDIB, 0, lBitmap)
''''            If lRes = 0 Then
''''
''''                'GdipImageRotateFlip lBitmap, 6 ' Rotate180FlipX
''''                GdipGetImageDimension lBitmap, lWidth, lHeight
''''                lrect.Right = lWidth
''''                lrect.Bottom = lHeight
''''
''''                If GdipBitmapLockBits(lBitmap, lrect, &H1, PixelFormat32bppARGB, bmpData) = 0 Then
''''                    ReDim lBufferIn(1 To 4, 1 To bmpData.Width, 1 To bmpData.Height) As Byte
''''                    RtlMoveMemory lBufferIn(1, 1, 1), ByVal bmpData.scan0, bmpData.stride * bmpData.Height
''''                    Call GdipBitmapUnlockBits(lBitmap, bmpData)
''''                    ' Inverse les valeurs pour passer de BGRA � RGBA
''''                    'Dim lBufferOut() As Byte
''''                    'ReDim lBufferOut(1 To 4, 1 To bmpData.Width, 1 To bmpData.Height) As Byte
''''
''''                    '                For lX = 1 To bmpData.Width
''''                    '                    For ly = 1 To bmpData.Height
''''                    '                        lBufferOut(1, lX, ly) = lBufferIn(3, lX, ly)
''''                    '                        lBufferOut(2, lX, ly) = lBufferIn(2, lX, ly)
''''                    '                        lBufferOut(3, lX, ly) = lBufferIn(1, lX, ly)
''''                    '                        lBufferOut(4, lX, ly) = lBufferIn(4, lX, ly)
''''                    '                        ' Si une couleur de transparence est sp�cifi�e
''''                    '                       ' If pTransparentColor <> -1 Then
''''                    '                            'If RGB(lBufferIn(3, lX, ly), lBufferIn(2, lX, ly), lBufferIn(1, lX, ly)) = pTransparentColor Then
''''                    '                                'lBufferOut(4, lX, ly) = 0
''''                    '                            'Else
''''                    '                                lBufferOut(4, lX, ly) = 255
''''                    '                            'End If
''''                    '                       ' End If
''''                    '                    Next
''''                    '                Next
''''                End If
''''
''''
''''
''''
''''
''''            End If
''''            '        ' Shutdown GDI+
''''            GdiplusShutdown lGDIP
''''        End If
''''        '    If lRes Then
''''        '
''''        '        Err.Raise 5, , "Cannot save the image. GDI+ Error:" & lRes
''''        '    End If
''    End Sub
''
''
    
    
    Public Sub DestroyCanvas()
        If c_lDIB Then
            Call DeleteObject(SelectObject(c_lhDC, m_OldBmp)): c_lDIB = 0
            Call DeleteDC(c_lhDC): c_lhDC = 0
        End If
    End Sub
    
    Public Function CreateCanvas(ByVal Width As Long, ByVal Height As Long) As Boolean
        Dim tBITMAPINFO     As BITMAPINFO
        
        If c_lDIB <> 0 Then DestroyCanvas
        
        tSIZE.cx = Width
        tSIZE.cy = Height
        
        With tBITMAPINFO.bmiHeader
            .biSize = Len(tBITMAPINFO.bmiHeader)
            .biBitCount = 32
            .biHeight = tSIZE.cy
            .biWidth = tSIZE.cx
            .biPlanes = 1
            .biSizeImage = .biWidth * .biHeight * 4
        End With
        
        c_lhDC = CreateCompatibleDC(0)
        c_lDIB = CreateDIBSection(c_lhDC, tBITMAPINFO, DIB_RGB_COLORS, m_Pointer, 0&, 0&)
        m_OldBmp = SelectObject(c_lhDC, c_lDIB)
        
        CreateCanvas = c_lDIB <> 0
    End Function
    
    Public Sub Clear()
        If m_Pointer Then FillMemory ByVal m_Pointer, tSIZE.cx * tSIZE.cy * 4&, 0
    End Sub
    
    Public Property Get hdc() As Long
        hdc = c_lhDC
    End Property
    
    Public Sub UpdateLayered(hwnd As Long, Optional ByVal Alpha As Byte = 255)
        Dim tPT As POINTAPI
        Dim ExStyle As Long
        
        ExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        If (ExStyle And WS_EX_LAYERED) <> WS_EX_LAYERED Then
            Call SetWindowLong(hwnd, GWL_EXSTYLE, ExStyle Or WS_EX_LAYERED)
        End If
        
        With tBLENDFUNCTION
            .AlphaFormat = AC_SRC_ALPHA
            .SourceConstantAlpha = Alpha
        End With
        
        Call UpdateLayeredWindow(hwnd, 0&, ByVal 0&, tSIZE, c_lhDC, tPT, 0&, tBLENDFUNCTION, ULW_ALPHA)
    End Sub
    
    Private Sub Class_Terminate()
        DestroyCanvas
    End Sub
