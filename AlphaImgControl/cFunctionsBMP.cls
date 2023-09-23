VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write BMP image format ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' GDI+ doesn't honor alpha channel for 32bpp bitmaps
' This class will process bitmaps & determine if the alpha channel is used and if so,
'   will manually create the bitmap else allow GDI+ to load it.


Public Function IsBitmapResourceFile(FileHandle As Long) As Long

    Dim lValue As Long, Offset As Long, lBpp As Long, Width As Long, Height As Long
    Dim aDIB() As Long, tSA As SafeArray, lRead As Long
    
    ' manually parse the bitmap header vs using VB's LoadPicture.
    ' Why? because VB's LoadPicture will convert the image into a screen
    ' compatible bitmap; where if screen resolution was less than true color,
    ' a 32bpp image would end up being 24bp or less vs 32bpp
    
    ' only 32bpp bitmaps are completely scanned here, other bitmap formats are quickly passed
    
    SetFilePointer FileHandle, 0&, 0&, 0&: ReadFile FileHandle, lValue, 2&, lRead, ByVal 0&
    If lValue = &H4D42 Then                         ' is it a bmp magic number
        SetFilePointer FileHandle, 28&, 0&, 0&: ReadFile FileHandle, lBpp, 2&, lRead, ByVal 0&
        SetFilePointer FileHandle, 18, 0&, 0&: ReadFile FileHandle, Width, 4&, lRead, ByVal 0&
        SetFilePointer FileHandle, 22&, 0&, 0&: ReadFile FileHandle, Height, 4&, lRead, ByVal 0&
        
        ' validate size
        ' width must be at least 1 pixel & height must be a least 1 pixel
        If Width < 1 Or Height = 0& Then Exit Function ' -Height indicates top down DIB
        
        On Error Resume Next
        SetFilePointer FileHandle, 10&, 0&, 0&: ReadFile FileHandle, Offset, 4&, lRead, ByVal 0&
        ' start of image
        ' validate enough bytes exist for the image
        lValue = (GetFileSize(FileHandle, 0&)) - (modCommon.ByteAlignOnWord(lBpp, Width) * Abs(Height) + Offset)
        If Err Then     ' should some overflow occur
            Err.Clear
            Exit Function
        End If
        If lValue >= 0& Then
            If lBpp = 32 Then
                lValue = 0&             ' not coded to handle compression<>0
                SetFilePointer FileHandle, 30&, 0&, 0&: ReadFile FileHandle, lValue, 4&, lRead, ByVal 0&
                If lValue = 0& Then
                    ReDim aDIB(0 To Width - 1, 0 To Abs(Height) - 1&)
                    lValue = Width * Abs(Height) * 4&
                    SetFilePointer FileHandle, Offset, 0&, 0&: ReadFile FileHandle, aDIB(0, 0), lValue, lRead, ByVal 0&
                    If lValue = lRead Then IsBitmapResourceFile = modCommon.ValidateAlphaChannel(aDIB())
                    Erase aDIB()
                    Exit Function
                End If
            End If
        End If
        IsBitmapResourceFile = lvicColor24bpp
    End If
            
End Function

Public Function IsBitmapResource(theStream() As Byte, Width As Long, Height As Long, ColorFormat As Long) As Long
    
    Dim lValue As Long, Offset As Long, lBpp As Long
    Dim aDIB() As Long, tSA As SafeArray, tBMP As BitmapData, tSize As RECTI
    
    ' manually parse the bitmap header vs using VB's LoadPicture.
    ' Why? because VB's LoadPicture will convert the image into a screen
    ' compatible bitmap; where if screen resolution was less than true color,
    ' a 32bpp image would end up being 24bpp or less vs 32bpp
    
    ' Functions return value is one of 3 possibilities
    ' 0 = not a bitmap or possibly corrupted
    ' 1 = bitmap, non-32bpp for our purposes
    ' other = a pointer into the passed stream where 1st scanline can be found
    '   If image is bottom up (usually the case), Width is returned negative
    
    If UBound(theStream) < 57 Then Exit Function
    CopyMemory lValue, theStream(0&), 2&            ' get 1st 2 bytes of the stream
    If lValue = &H4D42 Then                         ' is it a bmp magic number
        CopyMemory lBpp, theStream(28), 2&          ' bit count
        CopyMemory Width, theStream(18), 4&         ' width
        CopyMemory Height, theStream(22), 4&        ' height
        
        ' validate size
        ' width must be at least 1 pixel & height must be a least 1 pixel
        If Width < 1 Or Height = 0& Then Exit Function ' -Height indicates top down DIB
        
        On Error Resume Next
        CopyMemory Offset, theStream(10), 4&        ' start of image
        ' validate enough bytes exist for the image
        lValue = (1& + UBound(theStream)) - (modCommon.ByteAlignOnWord(lBpp, Width) * Abs(Height) + Offset)
        If Err Then                                 ' should some overflow occur
            Err.Clear
            Exit Function
        End If
        If lValue >= 0& Then                        ' is array big enough?
            If lBpp = 32 Then                       ' else we will allow VB to convert it for us
                                                    ' because it doesn't contain transparency anyway
                CopyMemory lValue, theStream(30&), 4&  ' compression
                If lValue = 0& Then                 ' manually handle no-compression bitmaps
                    With tSA
                        .cbElements = 4
                        .cDims = 2
                        .pvData = VarPtr(theStream(Offset))
                        .rgSABound(0).cElements = Abs(Height)
                        .rgSABound(1).cElements = Width
                    End With
                    
                    ' see if 32bpp uses alpha channel
                    CopyMemory ByVal VarPtrArray(aDIB), VarPtr(tSA), 4&
                    ColorFormat = modCommon.ValidateAlphaChannel(aDIB())
                    CopyMemory ByVal VarPtrArray(aDIB), 0&, 4& ' remove overlay
                    
                    If ColorFormat <> lvicColor32bpp Then
                        If Height < 0& Then         ' top down dib; rare
                            tBMP.Scan0Ptr = VarPtr(theStream(Offset))
                            tBMP.stride = Width * 4&
                        Else                        ' invert scanline offset & width
                            tBMP.Scan0Ptr = VarPtr(theStream(Offset + Width * 4& * (Height - 1&)))
                            tBMP.stride = Width * -4&
                        End If
                        tSize.nHeight = Abs(Height): tSize.nWidth = Width
                        If GdipCreateBitmapFromScan0(tSize.nWidth, tSize.nHeight, 0&, ColorFormat, ByVal 0&, IsBitmapResource) = 0& Then
                            If GdipBitmapLockBits(IsBitmapResource, tSize, ImageLockModeWrite Or ImageLockModeUserInputBuf, ColorFormat, tBMP) Then
                                GdipDisposeImage IsBitmapResource
                                IsBitmapResource = 1&
                            Else
                                GdipBitmapUnlockBits IsBitmapResource, tBMP
                            End If
                        End If
                        Exit Function
                    End If
                End If
            End If
        End If
        IsBitmapResource = 1&
    End If


End Function

Public Function SaveAsBMP(returnObject As Variant, SourceHandle As Long, _
                        ByVal returnMedium As SaveAsMedium, Optional SaveOptions As Variant) As Long

    ' saves image as a BMP or HBITMAP
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
    
    Dim bOK As Boolean, SS As SAVESTRUCT
    Dim uEncCLSID(0 To 3) As Long, tData() As Byte
    Dim tBMPi As BITMAPINFOHEADER, tSA As SafeArray
    Dim tBMP As BitmapData, tSize As RECTI, tSizeF As RECTF
    Dim tgtHandle As Long, srcDepth As Long, lResult As Long
    Dim X As Long, Y As Long, dPtr As Long, lShift As Byte
    Dim cPal As cColorReduction, tGDIpImage As GDIpImage, tObject As Object
    Dim tPic As StdPicture, tDO As DataObject
    
    tgtHandle = SourceHandle                                    ' get requested bitdepth
    If IsMissing(SaveOptions) Then                           ' will be missing when called from GDIpImage
        SS.RSS.FillColorARGB = modCommon.Color_RGBtoARGB(vbWindowBackground, 255&)
        GdipGetImageBounds SourceHandle, tSizeF, UnitPixel          ' get image size
        SS.Height = tSizeF.nHeight: SS.Width = tSizeF.nWidth
    Else
        SS = SaveOptions
    End If
    srcDepth = modCommon.ColorDepthToColorType(SS.ColorDepth, SourceHandle)
    If SS.ColorDepth > lvicNoColorReduction Then
        Set cPal = New cColorReduction
        If returnMedium = saveTo_Clipboard Or returnMedium = saveTo_DataObject Then SS.reserved1 = SS.reserved1 Or &H20000000
        If SS.ColorDepth = lvicDefaultReduction Then
            If SS.ColorDepth = lvicDefaultReduction Then SS.reserved1 = SS.reserved1 Or &H10000000
            lResult = cPal.PalettizeToArray(SourceHandle, tData(), orient_BottomUp Or orient_WantPaletteInArray, 54&, 0&, alpha_None, SS)
            If lResult Then srcDepth = lResult
        ElseIf SS.ColorDepth < lvicConvert_TrueColor24bpp Then
            lResult = cPal.PalettizeToArray(SourceHandle, tData(), orient_BottomUp Or orient_WantPaletteInArray, 54&, 0&, alpha_None, SS)
            If lResult Then srcDepth = lResult Else Exit Function   ' palettizer failed
        End If
        Set cPal = Nothing
    End If
    
    If lResult = 0& And SS.reserved2 <> 0& And (SS.reserved1 And &HFF00&) \ &H100& = lvicPicTypeBitmap Then
        CopyMemory tObject, SS.reserved2, 4&
        Set tGDIpImage = tObject
        CopyMemory tObject, 0&, 4&
        bOK = tGDIpImage.ExtractImageData(tData)
    End If
    
    If tGDIpImage Is Nothing Then
        With tBMPi
            .biBitCount = (srcDepth And &HFF00&) \ &H100
            .biHeight = SS.Height
            .biWidth = SS.Width
            .biPlanes = 1
            .biSize = 40
            .biSizeImage = modCommon.ByteAlignOnWord(.biBitCount, .biWidth)
            If .biBitCount <= 8& Then
                .biClrUsed = 2 ^ .biBitCount
                .biClrImportant = .biClrUsed
            End If
        End With
        
        If lResult = 0& Then                                        ' else tData() already filled
            
            Select Case tBMPi.biBitCount
            Case 32, 24:    ReDim tData(0 To tBMPi.biSizeImage * SS.Height + 53&)
            Case 8:         ReDim tData(0 To tBMPi.biSizeImage * SS.Height + 1077&)
            Case 4:         ReDim tData(0 To tBMPi.biSizeImage * SS.Height + 117&)
            Case 1:         ReDim tData(0 To tBMPi.biSizeImage * SS.Height + 61&)
            Case Else:      Exit Function
            End Select
            ' ^^ array sizes include 14 byte file header + 40 byte bitmapinfo header + palette if any
                    
            tBMP.PixelFormat = srcDepth
            tBMP.stride = -tBMPi.biSizeImage
            tBMP.Scan0Ptr = VarPtr(tData(UBound(tData) - tBMPi.biSizeImage + 1&))
            tSize.nWidth = SS.Width: tSize.nHeight = SS.Height
            If GdipBitmapLockBits(SourceHandle, tSize, ImageLockModeRead Or ImageLockModeUserInputBuf, srcDepth, tBMP) Then Exit Function
            GdipBitmapUnlockBits SourceHandle, tBMP
            
            If tBMPi.biBitCount <= 8& Then              ' transfer the palette
                Set cPal = New cColorReduction
                lResult = cPal.CopyRemotePalette(VarPtr(tData(54)), SourceHandle, False, 4&)
                For lResult = 57& To lResult * 4& + 53& Step 4&
                    tData(lResult) = 0                  ' remove alpha from palette
                Next
            End If
        
        End If
        
        tBMPi.biSizeImage = tBMPi.biSizeImage * SS.Height   ' fixup this member
        CopyMemory tData(14), tBMPi, tBMPi.biSize           ' copy the bitmap info header
                                                            ' fill in the file header info
        CopyMemory tData(0), &H4D42, 2&
        CopyMemory tData(2), CLng(UBound(tData) + 1&), 4&
        CopyMemory tData(10), CLng(54 + tBMPi.biClrUsed * 4&), 4&
        bOK = True
    End If
    
    If bOK Then
        If returnMedium = saveTo_Array Then
            SaveAsBMP = (UBound(tData) + 1&)
            modCommon.MoveArrayToVariant returnObject, tData(), True
        ElseIf returnMedium = saveTo_File Then
            WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
            SaveAsBMP = (lResult > UBound(tData))
        Else
            If returnMedium = saveTo_GDIplus Then Set tGDIpImage = modCommon.LoadImage(tData, True, , True)
            If tGDIpImage Is Nothing Then Set tGDIpImage = modCommon.LoadImage(tData, True, , True)
            If tGDIpImage.Handle Then
                If returnMedium = saveTo_GDIplus Then
                    SaveAsBMP = saveTo_GDIplus
                    Set returnObject = tGDIpImage
                Else
                    If GdipCreateHBITMAPFromBitmap(tGDIpImage.Handle, lResult, SS.RSS.FillColorARGB Or &HFF000000) Then Exit Function
                    Set tGDIpImage = Nothing
                    If returnMedium = saveTo_GDIhandle Then
                        returnObject = lResult
                        SaveAsBMP = lResult
                    Else
                        Set tPic = modCommon.HandleToStdPicture(lResult, vbPicTypeBitmap)
                        If tPic Is Nothing Then
                            DeleteObject lResult
                        ElseIf returnMedium = saveTo_stdPicture Then
                            Set returnObject = tPic
                            SaveAsBMP = (tPic.Handle <> 0&)
                        ElseIf returnMedium = saveTo_Clipboard Then
                            Clipboard.SetData tPic
                            If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat tData(), g_ClipboardFormat
                            SaveAsBMP = True
                        ElseIf returnMedium = saveTo_DataObject Then
                            Set tDO = returnObject
                            tDO.SetData tPic, vbCFBitmap
                            If g_ClipboardFormat Then tDO.SetData tData(), g_ClipboardFormat
                            SaveAsBMP = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function


