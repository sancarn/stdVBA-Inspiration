VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Win32/Non-GDI+ DC/Bitmap handling ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' Class is used when a control needs to have a permanent DC (FastRedraw=True for example)
' There can be up to 2 bitmaps created for the control, a 32bpp and 24bpp, again depends
'   on the control's requirements.

Private m_hDC As Long
Private Type CANVASSTRUCT
    hBmp As Long
    Cx As Long
    Cy As Long
    hPtr As Long
    bPP As Long
End Type

Private m_FastRedraw As CANVASSTRUCT
Private m_PrePost As CANVASSTRUCT

Public Sub UpdateDC(PrePostBmp As Boolean)

    ' basically a SelectObject call.
    ' Owner is responsible for calling this in pairs: one to select bmp into dc, the other to seclect out of dc
    If PrePostBmp Then
        m_PrePost.hBmp = SelectObject(m_hDC, m_PrePost.hBmp)
    Else
        m_FastRedraw.hBmp = SelectObject(m_hDC, m_FastRedraw.hBmp)
    End If
    
End Sub

Public Sub TransferFastRedrawToPrePost(X As Long, Y As Long, Cx As Long, Cy As Long, srcX As Long, srcY As Long, Transparency As Long)

    ' Transfers the content of m_FastRedraw to m_PrePost
    ' m_PrePost must already be selected into the DC
    Dim tDC As Long
    If m_hDC Then
        tDC = CreateCompatibleDC(m_hDC)
        If tDC Then
            m_FastRedraw.hBmp = SelectObject(tDC, m_FastRedraw.hBmp)
            AlphaBlend m_hDC, X, Y, Cx, Cy, tDC, srcX, srcY, Cx, Cy, &H1000000 Or ((((100& - Transparency) * 255&) \ 100&) * &H10000)
            m_FastRedraw.hBmp = SelectObject(tDC, m_FastRedraw.hBmp)
            DeleteDC tDC
        End If
    End If
    
End Sub

Public Function ResizeBitmap(ByVal Width As Long, ByVal Height As Long, BitDepth As Long, PrePostBmp As Boolean) As Boolean
    
    ' passing 0 width & height destroys the bitmap
    
    If m_hDC = 0& Then Exit Function
    
    Dim BHI As BITMAPINFOHEADER, hCanvas As CANVASSTRUCT
    
    If PrePostBmp Then hCanvas = m_PrePost Else hCanvas = m_FastRedraw
    
    If (Width = hCanvas.Cx And Height = hCanvas.Cy And BitDepth = hCanvas.bPP) Then
        ResizeBitmap = True
    Else
        If hCanvas.hBmp Then
            With hCanvas
                DeleteObject .hBmp
                .hBmp = 0&: .Cx = 0&: .Cy = 0&: .hPtr = 0&: .bPP = 0&
            End With
        End If
        If (Width = 0& Or Height = 0&) Then
            ResizeBitmap = True
        Else
        
            With BHI
                .biBitCount = BitDepth
                .biHeight = Height
                .biWidth = Width
                .biPlanes = 1
                .biSize = 40
            End With
            hCanvas.hBmp = CreateDIBSection(m_hDC, BHI, 0&, hCanvas.hPtr, 0&, 0&)
            If hCanvas.hBmp Then
                hCanvas.Cx = Width: hCanvas.Cy = Abs(Height): hCanvas.bPP = BitDepth
                ResizeBitmap = True
            End If
        End If
    End If
    
    If PrePostBmp Then
        m_PrePost = hCanvas
    Else
        m_FastRedraw = hCanvas
    End If
    If hCanvas.bPP = 32& Then FillMemory ByVal hCanvas.hPtr, Width * hCanvas.Cy * 4&, 0
    
End Function

Public Sub EraseBitmap(PrePostBmp As Boolean)
    
    ' ensures DIB is zeroed out
    
    Dim hCanvas As CANVASSTRUCT
    If PrePostBmp Then hCanvas = m_PrePost Else hCanvas = m_FastRedraw
    With hCanvas
        If .bPP = 24& Then
            FillMemory ByVal .hPtr, modCommon.ByteAlignOnWord(.bPP, .Cx) * .Cy, 0
        ElseIf .bPP = 32& Then
            FillMemory ByVal .hPtr, .Cx * .Cy * 4&, 0
        End If
    End With
End Sub

Public Sub FillBitmap(PrePostBmp As Boolean, ByVal Color As Long)

    ' Fill's bitmap with solid color

    Dim hGraphics As Long, hBrush As Long
    Dim hCanvas As CANVASSTRUCT, fRect As RECTI
    
    If PrePostBmp Then hCanvas = m_PrePost Else hCanvas = m_FastRedraw
    With hCanvas
        If .hBmp Then
            If .bPP = 24& Then
                If (Color And &H80000000) Then
                    hBrush = GetSysColorBrush(Color And &HFF&)
                Else
                    hBrush = CreateSolidBrush(Color)
                End If
                SetRect fRect, 0&, 0&, .Cx, .Cy
                FillRect m_hDC, fRect, hBrush
                If (Color And &H80000000) = 0& Then DeleteObject hBrush
            
            ElseIf g_TokenClass.Token Then
                If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
                Color = (Color And &HFF00&) Or ((Color And &HFF&) * &H10000) Or ((Color And &HFF0000) \ &H10000)
                If GdipCreateFromHDC(m_hDC, hGraphics) = 0& Then
                    GdipGraphicsClear hGraphics, &HFF000000 Or Color
                    GdipDeleteGraphics hGraphics
                End If
            End If
        End If
    End With

End Sub

Public Property Get Height(PrePostBmp As Boolean) As Long
    If PrePostBmp Then Height = m_PrePost.Cy Else Height = m_FastRedraw.Cy
End Property

Public Property Get Width(PrePostBmp As Boolean) As Long
    If PrePostBmp Then Width = m_PrePost.Cx Else Width = m_FastRedraw.Cx
End Property

Public Property Get scanWidth(PrePostBmp As Boolean)
    If PrePostBmp Then
        scanWidth = modCommon.ByteAlignOnWord(m_PrePost.bPP, m_PrePost.Cx)
    Else
        scanWidth = modCommon.ByteAlignOnWord(m_FastRedraw.bPP, m_FastRedraw.Cx)
    End If
End Property

Public Property Get Pointer(PrePostBmp As Boolean, Invert As Boolean) As Long
    
    ' returns the DIB pointer. Inverted pointers are used with GdipCreateFromScan0 when
    ' image is bottom up (as is in DIBs). The pointer needs to be set to the first column
    ' of the last scan line else image is loaded upside down by GDI+
    
    If PrePostBmp Then
        If Invert Then
            Pointer = m_PrePost.hPtr + (m_PrePost.Cy - 1&) * Me.scanWidth(PrePostBmp)
        Else
            Pointer = m_PrePost.hPtr
        End If
    ElseIf Invert Then
        Pointer = m_FastRedraw.hPtr + (m_FastRedraw.Cy - 1&) * Me.scanWidth(PrePostBmp)
    Else
        Pointer = m_FastRedraw.hPtr
    End If
End Property

Public Property Get hBitmap(PrePostBmp As Boolean) As Long
    If PrePostBmp Then hBitmap = m_PrePost.hBmp Else hBitmap = m_FastRedraw.hBmp
End Property

Public Property Get BitDepth(PrePostBmp As Boolean) As Long
    If PrePostBmp Then BitDepth = m_PrePost.bPP Else BitDepth = m_FastRedraw.bPP
End Property

Public Sub FreeHandle(PrePostBmp As Boolean)
    ' in some cases a GDI+ bitmap will be created from the hBitmap member of this class
    ' In those cases, can't allow this class to destroy the handle upon termination
    ' Note: before GDI+ creates the handle, ensure bitmap is not selected into a DC
    If PrePostBmp Then m_PrePost.hBmp = 0& Else m_FastRedraw.hBmp = 0&
End Sub

Public Property Get DC() As Long
    DC = m_hDC
End Property

Private Sub Class_Initialize()
    Dim tDC As Long
    tDC = GetDC(GetDesktopWindow())
    m_hDC = CreateCompatibleDC(tDC)
    ReleaseDC GetDesktopWindow(), tDC
End Sub

Private Sub DestroyGDIOjects()
    If m_hDC Then DeleteDC m_hDC: m_hDC = 0&
    If m_PrePost.hBmp Then
        With m_PrePost
            DeleteObject .hBmp
            .hBmp = 0&: .Cx = 0&: .Cy = 0&: .hPtr = 0&
        End With
    End If
    If m_FastRedraw.hBmp Then
        DeleteObject m_FastRedraw.hBmp
        m_FastRedraw = m_PrePost
    End If
End Sub

Private Sub Class_Terminate()
    DestroyGDIOjects
End Sub
