Attribute VB_Name = "MdlImgListAddImage"
Option Explicit
'-----------------------------------------
'Autor:     Leandro Ascierto
'Web:       www.leandroascierto.com.ar
'Date:      30 Oct 2009
'Creditos:  LaVolpe, Cobein

'Revición: 18/01/2010
'   Se implemento La lectura desde Recursos
'   se implemento la lectura desde Stream
'-------------------------------------------
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDrawImage Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipGetImageBounds Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipDrawImageRect Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long

Private Type RECTF
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type IconHeader
    ihReserved          As Integer
    ihType              As Integer
    ihCount             As Integer
End Type

Private Type IconEntry
    ieWidth             As Byte
    ieHeight            As Byte
    ieColorCount        As Byte
    ieReserved          As Byte
    iePlanes            As Integer
    ieBitCount          As Integer
    ieBytesInRes        As Long
    ieImageOffset       As Long
End Type

Private Const IconVersion           As Long = &H30000
Private Const LR_LOADFROMFILE       As Long = &H10
Private Const IMAGE_ICON            As Long = 1
Private Const PixelFormat32bppARGB  As Long = &H26200A
Private Const UnitPixel             As Long = &H2&

'------------------------------------------------------------------------
'Inserta una imagen (Ico,Png,jpg,bmp, etc.) al imagelist desde archivo
'------------------------------------------------------------------------
Public Function AddImageFromFile(ByVal FileName As String, ImgList) As Boolean

    Dim hIcon As Long
    Dim GDIsi As GdiplusStartupInput
    Dim gToken As Long
    Dim hGraphics As Long
    Dim hBitmap As Long
    Dim ResizeBmp As Long
    Dim ResizeGra As Long
    Dim R As RECTF
    Dim lWidth As Long
    Dim lHeight As Long
    Dim FileType As String
    
    On Local Error GoTo AddImageFromFile_Error
    
    lWidth = ImgList.ImageWidth
    lHeight = ImgList.ImageHeight
    
    FileType = UCase(Right(FileName, 3))
    
    If FileType = "ICO" Or FileType = "CUR" Then
        hIcon = LoadImage(App.hInstance, FileName, IMAGE_ICON, lWidth, lHeight, LR_LOADFROMFILE)
        If hIcon <> 0 Then
            AddImageFromFile = ImgLstAddAlphaIcon(hIcon, ImgList)
            Exit Function
        End If
    End If
    
    GDIsi.GdiplusVersion = 1&
    
    If GdiplusStartup(gToken, GDIsi) = 0 Then

        If GdipLoadImageFromFile(StrPtr(FileName), hBitmap) = 0 Then
        
            Call GdipGetImageBounds(hBitmap, R, UnitPixel)
            
            If lWidth <> R.Width Or lHeight <> R.Height Then
        
                If GdipCreateBitmapFromScan0(lWidth, lHeight, 0&, PixelFormat32bppARGB, ByVal 0&, ResizeBmp) = 0 Then
              
                    If GdipGetImageGraphicsContext(ResizeBmp, ResizeGra) = 0 Then
                       
                        If GdipDrawImageRect(ResizeGra, hBitmap, 0, 0, lWidth, lHeight) = 0 Then
                             
                            If GdipCreateHICONFromBitmap(ResizeBmp, hIcon) = 0 Then
                                AddImageFromFile = ImgLstAddAlphaIcon(hIcon, ImgList)
                            End If
                        
                        End If
                        
                        Call GdipDeleteGraphics(ResizeGra)
                        
                    End If
                    
                    Call GdipDisposeImage(ResizeBmp)
    
                End If
                
            Else
            
                If GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0 Then
                    AddImageFromFile = ImgLstAddAlphaIcon(hIcon, ImgList)
                End If

            End If
            
            GdipDisposeImage hBitmap
            
        End If
 
        GdiplusShutdown gToken
        
    End If
    
AddImageFromFile_Error:
    
End Function

'------------------------------------------------------------------------
'Inserta una imagen (Ico,Png,jpg,bmp, etc.) al imagelist desde Recurso
'------------------------------------------------------------------------
Public Function AddImageFromRes(ByVal ResIndex As Variant, ByVal ResSection As Variant, ImgList, Optional VBglobal As IUnknown) As Boolean
   
    Dim bvData()    As Byte
    Dim oVBglobal   As VB.Global
    Dim hIcon As Long
    
    On Local Error GoTo AddImageFromRes_Error

    If VBglobal Is Nothing Then
        Set oVBglobal = VB.Global
    ElseIf TypeOf VBglobal Is VB.Global Then
        Set oVBglobal = VBglobal
    ElseIf VBglobal Is Nothing Then
        Set oVBglobal = VB.Global
    End If
   
   
    bvData = oVBglobal.LoadResData(ResIndex, ResSection)

    If bvData(2) = vbResIcon Or bvData(2) = vbResCursor Then
        hIcon = LoadIconFromStream(bvData, ImgList.ImageWidth, ImgList.ImageHeight)
        Debug.Print ResIndex
    Else
        hIcon = LoadImageFromStream(bvData, ImgList.ImageWidth, ImgList.ImageHeight)
    End If
    
    If hIcon <> 0 Then
        AddImageFromRes = ImgLstAddAlphaIcon(hIcon, ImgList)
    End If


AddImageFromRes_Error:

End Function

'--------------------------------------------------------------------------------
'Lee una imagen (Png, jpg, bmp, etc.) desde un array de bits y devuelve un icono
'--------------------------------------------------------------------------------
Public Function LoadImageFromStream(ByRef bvData() As Byte, ByVal lWidth As Long, ByVal lHeight As Long) As Long
    
    Dim IStream     As IUnknown
    Dim GDIsi As GdiplusStartupInput
    Dim TR          As RECTF
    Dim hIcon       As Long
    Dim ResizeBmp As Long
    Dim ResizeGra As Long
    Dim hBitmap As Long
    Dim gToken As Long
    
    On Local Error GoTo LoadImageFromStream_Error
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function

    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
   
    If Not IStream Is Nothing Then
        GDIsi.GdiplusVersion = 1&
        If GdiplusStartup(gToken, GDIsi) = 0 Then
            If GdipLoadImageFromStream(IStream, hBitmap) = 0 Then
            
                Call GdipGetImageBounds(hBitmap, TR, UnitPixel)
                
                If lWidth <> TR.Width Or lHeight <> TR.Height Then
            
                    If GdipCreateBitmapFromScan0(lWidth, lHeight, 0&, PixelFormat32bppARGB, ByVal 0&, ResizeBmp) = 0 Then
                    
                        If GdipGetImageGraphicsContext(ResizeBmp, ResizeGra) = 0 Then
                           
                            If GdipDrawImageRect(ResizeGra, hBitmap, 0, 0, lWidth, lHeight) = 0 Then
                                 
                                If GdipCreateHICONFromBitmap(ResizeBmp, hIcon) = 0 Then
                    
                                    LoadImageFromStream = hIcon
                    
                                End If
                                
                             End If
                            
                            Call GdipDeleteGraphics(ResizeGra)
                            
                        End If
                        
                        Call GdipDisposeImage(ResizeBmp)

                    End If
                Else
                
                   If GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0 Then
                   
                        LoadImageFromStream = hIcon
                        
                   End If

                End If
    
                
            End If
            
            GdiplusShutdown gToken: gToken = 0
        End If
    End If

    Set IStream = Nothing
    
LoadImageFromStream_Error:

End Function
'------------------------------------------------------------------
'Lee una imagen ICO, CUR desde un array de bits y devuelve un icono
'------------------------------------------------------------------
Public Function LoadIconFromStream(ByRef bytIcoData() As Byte, ByVal lWidth As Long, ByVal lHeight As Long) As Long

    Dim tIconHeader     As IconHeader
    Dim tIconEntry()    As IconEntry
    Dim MaxBitCount     As Long
    Dim MaxSize         As Long
    Dim Aproximate      As Long
    Dim IconID          As Long
    Dim hIcon           As Long
    Dim i               As Long
  
    On Local Error GoTo LoadIconFromStream_Error
    
    If Not IsArrayDim(VarPtrArray(bytIcoData)) Then Exit Function

    Call CopyMemory(tIconHeader, bytIcoData(0), Len(tIconHeader))

    If tIconHeader.ihCount >= 1 Then
    
        ReDim tIconEntry(tIconHeader.ihCount - 1)
        
        Call CopyMemory(tIconEntry(0), bytIcoData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
        
        IconID = -1
           
        For i = 0 To tIconHeader.ihCount - 1
            If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
        Next

       
        For i = 0 To tIconHeader.ihCount - 1
            If MaxBitCount = tIconEntry(i).ieBitCount Then
                MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                If MaxSize > Aproximate And MaxSize <= (lWidth + lHeight) Then
                    Aproximate = MaxSize
                    IconID = i
                End If
            End If
        Next
                   
        If IconID = -1 Then Exit Function
       
        With tIconEntry(IconID)
            hIcon = CreateIconFromResourceEx(bytIcoData(.ieImageOffset), .ieBytesInRes, 1, IconVersion, lWidth, lHeight, &H0)
            If hIcon <> 0 Then
                LoadIconFromStream = hIcon
            End If
        End With
       
    End If

LoadIconFromStream_Error:

End Function


Private Function ImgLstAddAlphaIcon(ByVal hIcon As Long, ImgList) As Boolean
On Local Error GoTo ImgLstAddAlphaIcon_Error

    ImgList.ListImages.Add , , ImgList.Parent.Icon
    ImageList_ReplaceIcon ImgList.hImageList, ImgList.ListImages.Count - 1, hIcon
    DestroyIcon hIcon
    ImgLstAddAlphaIcon = True
    
ImgLstAddAlphaIcon_Error:
End Function


Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

