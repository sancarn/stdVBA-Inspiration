VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write GIF image format ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' GDI+ has some issues with animated GIFs, in fact several, so
'   instead of living with the shortfalls, this class handles them
' Logic employed....
' 1) Parse the GIF, extracting some frame information and store in custom structures below
' 2) Cache the original GIF bytes (entire GIF). This is used as a template to create frames on the fly
' 3) Render each frame as needed. GIF disposal codes dictate how frames are rendered, erased, combined.

Private Type FRAMESTRUCT
    DataStart As Long           ' where frame data actually begins in resource
    DataLen As Long             ' size of the frame's data
    ImgDescrip As Long          ' where the image descriptor block starts
    ColorTable As Long          ' where the frame's color table starts
    ColorCount As Long          ' size of frame's color table
    Bounds As RECTI             ' size of frame, including any offsets
    Disposal As Byte            ' frame's disposal code
    Transparency As Byte        ' 1 if frame contains transparency else zero
End Type
Private Type GIFStruct
    Size As POINTAPI            ' overall width,height
    LoopCount As Long           ' nr loops suggested
    FrameCount As Long          ' nr frames
    Frame As Long               ' current frame
    Code3Size As Long           ' size of m_BlankCanvas if any frames use GIF disposal code #3
    GlobalOffset As Long        ' offset to where data in the m_IStream begins GIF data
    Frames() As FRAMESTRUCT  ' one per frame
End Type
Private m_IStream As IUnknown
Private m_BlankCanvas() As Byte
Private m_GIFdata As GIFStruct
Private m_Token As cGDIpToken

Public Function IsGIFResource(cImageData As cGDIpMultiImage, inStream() As Byte, CacheData As Boolean, Optional ParseOnly As Boolean) As Boolean

    ' function validates passed array is gif format & if so, parses the gif & returns its handle
    ' ParseOnly flag is only set when converting from animated GIF to animated PNG & called from cFunctionsPNG

    If m_Token.Token = 0& Then Exit Function
    
    Dim hHandle As Long
    Dim SeqDelay() As Long, X As Long, dispCode3 As Long
    
    ReDim m_BlankCanvas(0 To 5)
    CopyMemory m_BlankCanvas(0), inStream(0), 6&
    Select Case LCase$(StrConv(m_BlankCanvas(), vbUnicode))
    
        Case "gif89a", "gif87a"                         ' ok, got a GIF
            Call pvProcessGIFfile(inStream())           ' parse it now
            If m_GIFdata.FrameCount Then                ' success?
                If ParseOnly Then                       ' set only when converting GIF to APNG
                    IsGIFResource = (m_GIFdata.FrameCount > 1&)
                    Exit Function
                End If
                ' create a blank GDI+ bitmap
                If GdipCreateBitmapFromScan0(m_GIFdata.Size.X, m_GIFdata.Size.Y, 0&, lvicColor32bppAlpha, ByVal 0&, hHandle) = 0& Then
                    ' tidy up some GIF info to aid the rendering routine to process frames faster/efficiently
                    ' Don't let last frame have any disposal code. It's pointless
                    m_GIFdata.Frames(m_GIFdata.FrameCount - 1).Disposal = 0
                    ' Create a general use array used during frame navigation
                    If m_GIFdata.Code3Size Then
                        ReDim m_BlankCanvas(0 To m_GIFdata.Size.X * 4& * m_GIFdata.Code3Size - 1&)
                    Else
                        Erase m_BlankCanvas
                    End If
                    ' Prepare to create first frame & then create it
                    m_GIFdata.Frame = m_GIFdata.FrameCount - 1&
                    ' Ensure we can load the first frame else assume corrupted
                    If Me.MoveToFrame(0&, inStream(), hHandle) Then
                        cImageData.InitializeContent ObjPtr(Me), lvicPicTypeGIF, m_GIFdata.FrameCount
                        ReDim SeqDelay(0 To 1, 1 To m_GIFdata.FrameCount)
                        For X = 1& To m_GIFdata.FrameCount
                            SeqDelay(0, X) = X - 1&
                            With m_GIFdata.Frames(X - 1)
                                 If .ImgDescrip > .DataStart Then SeqDelay(1, X) = (((inStream(.DataStart + 4&) Or inStream(.DataStart + 5&) * &H100&)) And &HFFFF&) * 10&
                            End With
                        Next
                        cImageData.SetSequenceDuration SeqDelay()            ' send off & done
                        cImageData.CacheSourceInfo VarPtrArray(inStream()), hHandle, lvicPicTypeGIF, CacheData, False
                        IsGIFResource = True
                    Else
                        Set m_IStream = Nothing
                        Erase m_BlankCanvas()
                    End If
                End If
            End If
        Case Else
    End Select
    
End Function

Public Function IsGIFResourceFile(FileHandle As Long) As Boolean

    ' function simply checks if the passed file is GIF
    
    If m_Token.Token = 0& Then Exit Function
    
    Dim lRead As Long
    ReDim m_BlankCanvas(0 To 5)
    SetFilePointer FileHandle, 0&, 0&, 0&: ReadFile FileHandle, m_BlankCanvas(0), 6&, lRead, ByVal 0&
    If lRead = 6& Then
        Select Case LCase$(StrConv(m_BlankCanvas(), vbUnicode))
            Case "gif89a", "gif87a"
                IsGIFResourceFile = True
            Case Else
        End Select
    End If
    
End Function

Public Function SaveAsGIF(returnObject As Variant, ByVal SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                            SS As SAVESTRUCT, EncoderClassPtr As Long, MIS As MULTIIMAGESAVESTRUCT, FrameNumber As Long) As Long
                            
    ' saves image as a GIF
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    ' returnMedium.
    '   If saveTo_Array then returnObject is the 0-Bound array & function return value is size of array in bytes
    '   If saveTo_File then returnObject is passed HFILE, return value is non-zero if successful. File is not closed
    '   If saveTo_stdPictureture then returnObject is passed stdPicture, return value is non-zero if successful
    '   If saveTo_Clipboard then return value is non-zero if successful
    '   If saveTo_DataObject then returnObject is passed DataObject, return value is non-zero if successful
    '   If saveTo_GDIpHandle then
    '       if creating own GDIpImage class then returnObject is the class & return value is saveTo_GDIpHandle
    '       else returnObject is handle's IStream source & function return value is the GDI+ handle
    
    ' Note: GDI+ always requires 8 bpp indexes; regardless of the color depth
    ' Passing a format other than 8bbp causes GDI+ to convert to 32bpp then back down to 8bpp using halftone palette
    ' see following link for more: http://support.microsoft.com/kb/318343
    
    If SourceHandle = 0& Then Exit Function
    
    Dim tData() As Byte, bOk As Boolean
    Dim tgtHandle As Long, lResult As Long, lRtnMedium As Long
    Dim IIStream As IUnknown, tmpPic As StdPicture, tmpDO As DataObject
    Dim cPal As cColorReduction, tGDIpImage As GDIpImage, tObject As Object
    
    ' GIFs are always sent to color reduction routines. The reason is that we can only use
    '   8bpp indexes and only simple transparency. Flags used in the reduction calls will
    '   prevent color loss if possible
    
    Set cPal = New cColorReduction
    If MIS.Images > 1& Then lRtnMedium = saveTo_Array Else lRtnMedium = returnMedium
        
    If SS.ColorDepth = lvicDefaultReduction Then
        SS.reserved1 = SS.reserved1 Or &H10000000
        tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_Simple, SS, orient_8bppIndexes)
        If tgtHandle Then
            If (modCommon.ColorDepthToColorType(lvicNoColorReduction, tgtHandle) And &HFF00&) \ &H100& <> 8& Then
                GdipDisposeImage tgtHandle: tgtHandle = 0&
            End If
        End If
        SS.reserved1 = SS.reserved1 Xor &H10000000
    ElseIf SS.ColorDepth > lvicNoColorReduction Then
        tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_Simple, SS, orient_8bppIndexes)
    End If
    If tgtHandle = 0& Then                  ' not processed above
        If lRtnMedium <> saveTo_stdPicture And SS.ExtractCurrentFrameOnly = False And MIS.Images = 0& And _
            SS.reserved2 <> 0& And (SS.reserved1 And &HFF00&) \ &H100& = lvicPicTypeGIF Then
            CopyMemory tObject, SS.reserved2, 4&
            Set tGDIpImage = tObject
            CopyMemory tObject, 0&, 4&
            tgtHandle = SourceHandle
        Else
            SS.ColorDepth = lvicConvert_256Colors
            If SS.Palette_Handle = 0& Then SS.PaletteType = lvicPaletteAdaptive
            tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_Simple, SS, orient_8bppIndexes)
            If tgtHandle = 0& Then Exit Function            ' can't do it
        End If
    End If
    Set cPal = Nothing
    
    ' do we need a GDI+ object?
    If tGDIpImage Is Nothing Then
        Set IIStream = modCommon.IStreamFromArray(0&, 0&)
        If Not IIStream Is Nothing Then
            bOk = (GdipSaveImageToStream(tgtHandle, IIStream, ByVal EncoderClassPtr, ByVal 0&) = 0&)
        End If
    Else
        bOk = tGDIpImage.ExtractImageData(tData)
    End If
    
    If bOk Then
        If lRtnMedium = saveTo_GDIplus Then
            If tGDIpImage Is Nothing Then
                If GdipLoadImageFromStream(ObjPtr(IIStream), SaveAsGIF) = 0& Then returnObject = IIStream
            Else
                Set tGDIpImage = modCommon.LoadImage(tData(), , , True)
                If tGDIpImage.Handle Then
                    SaveAsGIF = lRtnMedium
                    Set returnObject = tGDIpImage
                End If
            End If
        Else
            If tGDIpImage Is Nothing Then
                bOk = modCommon.IStreamToArray(ObjPtr(IIStream), tData)
            Else
                bOk = True
            End If
            If bOk Then
                If lRtnMedium = saveTo_Array Then
                    SaveAsGIF = UBound(tData) + 1&
                    If MIS.Images = 0& Then modCommon.MoveArrayToVariant returnObject, tData(), True
                ElseIf lRtnMedium = saveTo_File Then
                    WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
                    SaveAsGIF = (lResult > UBound(tData))
                ElseIf lRtnMedium = saveTo_stdPicture Then
                    Set tmpPic = modCommon.ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
                    If Not tmpPic Is Nothing Then
                        Set returnObject = tmpPic
                        SaveAsGIF = tmpPic.Handle
                    End If
                Else
                    GdipCreateHBITMAPFromBitmap tgtHandle, lResult, SS.RSS.FillColorARGB Or &HFF000000
                    If lResult Then
                        Set tmpPic = modCommon.HandleToStdPicture(lResult, vbPicTypeBitmap)
                        If tmpPic Is Nothing Then
                            DeleteObject lResult
                        Else
                            SaveAsGIF = bOk
                            If lRtnMedium = saveTo_Clipboard Then
                                Clipboard.SetData tmpPic
                                If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat tData(), g_ClipboardFormat
                            Else
                                Set tmpDO = returnObject
                                tmpDO.SetData tmpPic, vbCFBitmap
                                If g_ClipboardFormat Then tmpDO.SetData tData(), g_ClipboardFormat
                                Set tmpDO = Nothing
                            End If
                            Set tmpPic = Nothing
                        End If
                    End If
                    Erase tData()
                End If
            End If
        End If
    End If
    
    If tgtHandle <> SourceHandle Then
        If tgtHandle Then GdipDisposeImage tgtHandle
    End If
    
    If MIS.Images > 1& And SaveAsGIF <> 0& Then
        Set MIS.Image(LBound(MIS.Image) + FrameNumber).Picture = Nothing
        
        ' need to append GIF to our array
        If FrameNumber = 0& Then                    ' 1st frame
            ReDim m_BlankCanvas(0 To UBound(tData) + 4& * MIS.Images)
        Else                                        ' 2nd & later frames
            ReDim Preserve m_BlankCanvas(0 To UBound(m_BlankCanvas) + UBound(tData) + 1&)
        End If
        ' store size of GIF data & then append gif data to end of array
        CopyMemory m_BlankCanvas(FrameNumber * 4&), CLng(UBound(tData) + 1), 4&
        CopyMemory m_BlankCanvas(UBound(m_BlankCanvas) - UBound(tData)), tData(0), UBound(tData) + 1&
        
        ' if on last frame, need to combine individual frames to animated gif format
        If MIS.Images = FrameNumber + 1& Then
            SaveAsGIF = 0&
            If pvCreateAnimatedGIF(MIS, tData()) Then    ' send to function that combines frames
                Select Case returnMedium
                    Case saveTo_GDIplus
                        Set tGDIpImage = modCommon.LoadImage(tData(), True, , True)
                        If tGDIpImage.Handle <> 0& Then
                            SaveAsGIF = returnMedium
                            Set returnObject = tGDIpImage
                        End If
                    Case saveTo_Array
                        SaveAsGIF = UBound(tData) + 1&
                        modCommon.MoveArrayToVariant returnObject, tData(), True
                    Case saveTo_File
                        WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
                        SaveAsGIF = (lResult > UBound(tData))
                    Case saveTo_stdPicture, saveTo_Clipboard, saveTo_DataObject
                        Set tmpPic = modCommon.ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
                        If Not tmpPic Is Nothing Then
                            If returnMedium = saveTo_Clipboard Then
                                Clipboard.SetData tmpPic
                                If g_ClipboardFormat Then modCommon.SetClipboardCustomFormat tData(), g_ClipboardFormat
                            ElseIf returnMedium = saveTo_DataObject Then
                                Set tmpDO = returnObject
                                tmpDO.SetData tmpPic, vbCFBitmap
                                If g_ClipboardFormat Then tmpDO.SetData tData(), g_ClipboardFormat
                                Set tmpDO = Nothing
                            Else
                                Set returnObject = tmpPic
                            End If
                            SaveAsGIF = tmpPic.Handle
                            Set tmpPic = Nothing
                        End If
                End Select
            End If
            Erase m_BlankCanvas()
        End If
    End If
    
End Function

Public Function ConvertFrameToPNG(sourceInfo As Variant, returnObject As Variant, FrameNumber As Long, MIS As MULTIIMAGESAVESTRUCT) As Boolean

    ' called only when converting animated GIFs to animated PNG
    ' frame numbers are passed to this routine as zero bound.
    
    Dim gHandle As Long, lPtr As Long, hLock As Long
    Dim tSA As SafeArray, aFrame() As Byte, sourceData() As Byte
    Dim srcBmpData As BitmapData, dstBmpData As BitmapData
    Dim srcRect As RECTI, dstRect As RECTI, SS As SAVESTRUCT
        
    modCommon.MoveArrayToVariant sourceInfo, sourceData(), False
    On Error GoTo ExitRoutine
    
'/////////// Get desired frame \\\\\\\\\\\\
    ' We will update our template stream vs. re-creating it; faster overall
    With tSA                ' use a DMA overlay on the stream handle
        .cbElements = 1
        .cDims = 1
        .pvData = pvGetStreamAddress(hLock)
        If .pvData = 0& Then GoTo ExitRoutine
        .rgSABound(0).cElements = m_GIFdata.GlobalOffset + m_GIFdata.Frames(FrameNumber).DataLen + 1&
    End With                ' set the DMA handle
    CopyMemory ByVal VarPtrArray(aFrame), VarPtr(tSA), 4&
    
    With m_GIFdata.Frames(FrameNumber)                              ' update the stream
        lPtr = m_GIFdata.GlobalOffset
        CopyMemory aFrame(lPtr), sourceData(.DataStart), .DataLen
        
        aFrame(lPtr + .DataLen) = 59
        CopyMemory aFrame(6), .Bounds.nWidth, 2&   ' make gif width same as frame
        CopyMemory aFrame(8), .Bounds.nHeight, 2&   ' make gif height same as frame
        CopyMemory aFrame(.ImgDescrip - .DataStart + lPtr + 1), 0&, 2& ' make gif left zero
        CopyMemory aFrame(.ImgDescrip - .DataStart + lPtr + 3), 0&, 2& ' make gif top zero & remove disposal code
        If .ImgDescrip > .DataStart Then aFrame(lPtr + 3) = (aFrame(lPtr + 3) And 227)
    
    End With
    
    CopyMemory ByVal VarPtrArray(aFrame), 0&, 4&                ' done; remove overlay & unlock pointer
    pvGetStreamAddress hLock
    tSA.pvData = 0&
    
    ' if all went well, we can create the GIF from the stream we just updated
    If GdipLoadImageFromStream(ObjPtr(m_IStream), gHandle) Then GoTo ExitRoutine

    If FrameNumber = 0& Then                                    ' first time thru: fill in structure
        MIS.GIFOverview.LoopCount = m_GIFdata.LoopCount
        MIS.GIFOverview.WindowHeight = m_GIFdata.Size.Y
        MIS.GIFOverview.WindowWidth = m_GIFdata.Size.X
        ReDim MIS.Image(0 To m_GIFdata.FrameCount - 1&)
        MIS.Images = m_GIFdata.FrameCount
    End If

    With m_GIFdata.Frames(FrameNumber)                          ' update frame info
        MIS.Image(FrameNumber).GIFFrameInfo.DisposalCode = .Disposal
        If FrameNumber Then
            lPtr = 1& ' flag indicating to include blend op flag
            If m_GIFdata.Frames(FrameNumber - 1&).Disposal = lvicGIF_Erase Then
                srcRect = m_GIFdata.Frames(FrameNumber).Bounds
                With m_GIFdata.Frames(FrameNumber - 1&).Bounds
                    If srcRect.nLeft + srcRect.nWidth <= .nLeft + .nWidth Then
                        If srcRect.nTop + srcRect.nHeight <= .nTop + .nHeight Then
                            ' no reason to blend (extra overhead) if frame is being placed over erased area
                            If srcRect.nLeft >= .nLeft And srcRect.nTop >= .nTop Then lPtr = 0&
                        End If
                    End If
                End With
            End If
            If lPtr Then ' blend only if frame contains transparency
                If modCommon.HasTransparency(gHandle) Then _
                    MIS.Image(FrameNumber).GIFFrameInfo.DisposalCode = MIS.Image(FrameNumber).GIFFrameInfo.DisposalCode Or lvicAPNG_BlendOpOver
            End If
        End If
        If .ImgDescrip > .DataStart Then
            MIS.Image(FrameNumber).GIFFrameInfo.DelayTime = (((sourceData(.DataStart + 4&) Or sourceData(.DataStart + 5&) * &H100&)) And &HFFFF&) * 10&
        End If
        
        ' default transfer boundaries from source to destination
        dstRect.nHeight = .Bounds.nHeight: dstRect.nWidth = .Bounds.nWidth
        srcRect = dstRect
        
        If FrameNumber = 0& Then
            SS.Width = m_GIFdata.Size.X                                     ' dimensions of canvas size
            SS.Height = m_GIFdata.Size.Y
            MIS.Image(FrameNumber).GIFFrameInfo.XOffset = 0&
            MIS.Image(FrameNumber).GIFFrameInfo.YOffset = 0&
            ' by specs, 1st frame of APNG must be same size as overall canvas.
            If Not (m_GIFdata.Size.X = .Bounds.nWidth And m_GIFdata.Size.Y = .Bounds.nHeight) Then
                ' transfer bits from current frame to new image. Define areas for copying
                dstRect.nLeft = .Bounds.nLeft: dstRect.nTop = .Bounds.nTop
            End If
        Else
            SS.Width = .Bounds.nWidth
            SS.Height = .Bounds.nHeight
            MIS.Image(FrameNumber).GIFFrameInfo.XOffset = .Bounds.nLeft
            MIS.Image(FrameNumber).GIFFrameInfo.YOffset = .Bounds.nTop
        End If
        If Not MIS.Image(0).Picture Is Nothing Then
            If (MIS.Image(0).Picture.Width = .Bounds.nWidth And MIS.Image(0).Picture.Height = .Bounds.nHeight) Then
                ' re-use existing image vs. creating new one if images are same size
                Set MIS.Image(FrameNumber).Picture = MIS.Image(0).Picture
            End If
        End If
        If MIS.Image(FrameNumber).Picture Is Nothing Then
            SS.ColorDepth = lvicConvert_TrueColor32bpp_ARGB             ' 32bpp depth
            ' create a blank image of the overall canvas size
            Set MIS.Image(FrameNumber).Picture = modCommon.LoadBlankImage(SS, True)  ' create canvas sized image (bitmap)
            If MIS.Image(FrameNumber).Picture.Handle = 0& Then GoTo ExitRoutine
        End If
        
    End With
    
    ' transfer GIF bytes to our bitmap bytes
    If GdipBitmapLockBits(gHandle, srcRect, ImageLockModeRead, lvicColor32bppAlpha, srcBmpData) = 0& Then
        dstBmpData = srcBmpData
        If GdipBitmapLockBits(MIS.Image(FrameNumber).Picture.Handle, dstRect, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor32bppAlpha, dstBmpData) = 0& Then
            GdipBitmapUnlockBits MIS.Image(FrameNumber).Picture.Handle, dstBmpData
        End If
        GdipBitmapUnlockBits gHandle, srcBmpData
    End If
    GdipDisposeImage gHandle: gHandle = 0&
    ' convert bitmap to png
    modCommon.SaveAsPNG aFrame(), MIS.Image(FrameNumber).Picture.Handle, saveTo_Array
    ' return the results
    modCommon.MoveArrayToVariant returnObject, aFrame(), True
    ConvertFrameToPNG = True

ExitRoutine:    ' clean up as needed
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(aFrame), 0&, 4&
    If gHandle Then GdipDisposeImage gHandle
    modCommon.MoveArrayToVariant sourceInfo, sourceData(), True
End Function

Private Function pvCreateAnimatedGIF(MIS As MULTIIMAGESAVESTRUCT, outArray() As Byte) As Boolean

    ' FUNCTION COMBINES SEVERAL SINGLE-FRAME GIFS INTO AN ANIMATED GIF
    
    ' GDI+ v1.0 cannot create animated GIFs, but we can...
    ' The trick we are using is rather simple
    '   Animated GIF frames contain a minimum of: color table, 2 control blocks & compressed data
    '   We used GDI+ to create the compressed data per frame
    '   GDI+ also creates 1 of the control blocks and if transparency exists in image, both blocks
    '   We simply rearrange the exsisting individual GIF bytes into animated GIF format, modifying very little

    ' Limitations
    '   - A true animated GIF editor would optimize the color tables of each frame. Ideally, one table would be
    '       created that contains colors of a majority of frames so that color table is shared & written once
    '   - Each frame has its own color table. This will produce larger files vs. a true animated gif editor
    '   To help reduce overall file size, the following two options are used
    '   1) Inform users via the SavePictureGDIp.RTF file that supplying their own palette in the SaveStructure
    '       will produce optimal file sizes
    '   2) Loop thru each frame, identify which table is used most often and set that one as the global table

    Dim f As Long, gTblIdx As Long
    Dim fPtr As Long, aPtr As Long
    Dim X As Long, Y As Long
    Dim lbOffset As Long, lSize As Long
    Dim fs() As FRAMESTRUCT
    
    ' VALIDATE WINDOW SIZE VS. FRAME SIZE
    lbOffset = LBound(MIS.Image)
    ReDim fs(0 To MIS.Images - 1&)
    fPtr = MIS.Images * 4&
    For f = gTblIdx To MIS.Images - 1&
        
        fs(f).Bounds.nWidth = m_BlankCanvas(fPtr + 6&) Or m_BlankCanvas(fPtr + 7&) * &H100&
        If MIS.Image(f + lbOffset).GIFFrameInfo.XOffset > 0& Then fs(f).Bounds.nWidth = fs(f).Bounds.nWidth + MIS.Image(f + lbOffset).GIFFrameInfo.XOffset
        If fs(f).Bounds.nWidth > X Then X = fs(f).Bounds.nWidth
        
        fs(f).Bounds.nHeight = m_BlankCanvas(fPtr + 8&) Or m_BlankCanvas(fPtr + 9&) * &H100&
        If MIS.Image(f + lbOffset).GIFFrameInfo.YOffset > 0& Then fs(f).Bounds.nHeight = fs(f).Bounds.nHeight + MIS.Image(f + lbOffset).GIFFrameInfo.YOffset
        If fs(f).Bounds.nHeight > Y Then Y = fs(f).Bounds.nHeight
        
        CopyMemory fs(f).DataLen, m_BlankCanvas(f * 4&), 4& ' get size of gif frame in bytes
        fPtr = fPtr + fs(f).DataLen                     ' move pointer along
    Next
    If MIS.GIFOverview.WindowWidth < X Then MIS.GIFOverview.WindowWidth = X
    If MIS.GIFOverview.WindowHeight < Y Then MIS.GIFOverview.WindowHeight = Y
    
    ' CALCULATE OUTPUT ARRAY SIZE AND POSITION FRAMES
    lSize = 32&                                     ' GIF primary header size + NETSCAPE2.0 app block
    fPtr = MIS.Images * 4&
    For f = gTblIdx To MIS.Images - 1&
        lSize = lSize + fs(f).DataLen - 14&         ' add to overall size & subtract header
        ' determine color table size to offset start of gif control block
        fs(f).ColorCount = (2& ^ ((m_BlankCanvas(fPtr + 10&) And &H7&) + 1&)) * 3&
        fs(f).ColorTable = fPtr + 13&
        ' if no control block, add 8 bytes for one to be inserted later
        fs(f).DataStart = fs(f).ColorCount + fPtr + 13&
        If m_BlankCanvas(fs(f).DataStart) = 44 Then lSize = lSize + 8&
        If MIS.Image(f + lbOffset).GIFFrameInfo.XOffset < 0& Then        ' auto-center frame if requested
            MIS.Image(f + lbOffset).GIFFrameInfo.XOffset = (MIS.GIFOverview.WindowWidth - fs(f).Bounds.nWidth) \ 2&
        End If
        If MIS.Image(f + lbOffset).GIFFrameInfo.YOffset < 0& Then
            MIS.Image(f + lbOffset).GIFFrameInfo.YOffset = (MIS.GIFOverview.WindowHeight - fs(f).Bounds.nHeight) \ 2&
        End If
        fPtr = fPtr + fs(f).DataLen                 ' move pointer along
    Next
    
    ' SIMPLE COLOR TABLE OPTIMIZATION
    ' We will run thru each color table to see if one or more is identical.
    ' If this occurs, then we will make the most used one the global table
    For f = gTblIdx To MIS.Images - 1&
        fs(f).ImgDescrip = 1&
        For X = gTblIdx To f - 1&
            If fs(f).ColorCount >= fs(X).ColorCount Then
                For Y = gTblIdx To fs(X).ColorCount - 1&
                    If m_BlankCanvas(fs(f).ColorTable + Y) <> m_BlankCanvas(fs(X).ColorTable + Y) Then Exit For
                Next
                If Y = fs(X).ColorCount Then                    ' table(y) replicated
                    fs(X).ImgDescrip = fs(X).ImgDescrip + 1&    ' increment replication count
                    fs(f).ColorTable = fs(X).ColorTable         ' re-direct color table offset
                    Exit For
                End If
            End If
        Next
    Next                                                        ' decide which table used most often
    For f = 1& To MIS.Images - 1&                               ' the one that uses most bytes is tie breaker
        If fs(f).ImgDescrip > 1& Then
            If fs(f).ImgDescrip >= fs(gTblIdx).ImgDescrip Then  ' keep track of highest usage
                If fs(f).ImgDescrip * fs(f).ColorCount > fs(gTblIdx).ImgDescrip * fs(gTblIdx).ColorCount Then
                    gTblIdx = f                                 ' cache new global table index
                End If
            End If
        End If
    Next

    ' SIZE OUTPUT ARRAY & COPY GIF HEADER + GLOBAL COLOR TABLE
    fPtr = MIS.Images * 4&
    lSize = lSize - ((fs(gTblIdx).ImgDescrip - 1&) * fs(gTblIdx).ColorCount)
    ReDim outArray(0 To lSize)
    CopyMemory outArray(0), m_BlankCanvas(fPtr), 13&
    CopyMemory outArray(13), m_BlankCanvas(fs(gTblIdx).ColorTable), fs(gTblIdx).ColorCount
    If gTblIdx > 0& Then    ' use that table's bit depth & sorted property
        outArray(10) = (outArray(10) And Not &H7&) Or _
            (m_BlankCanvas(fs(gTblIdx).ColorTable - 3&) And &H7&) Or _
            ((m_BlankCanvas(fs(gTblIdx).ColorTable - 3&) And &H10&) \ 2&)
    End If
    
    ' UPDATE OVERALL WINDOW SIZE
    CopyMemory outArray(6), MIS.GIFOverview.WindowWidth, 2&
    CopyMemory outArray(8), MIS.GIFOverview.WindowHeight, 2&
    aPtr = 13& + fs(gTblIdx).ColorCount             ' move pointer along
    
    ' write the loop count block
    outArray(aPtr) = 33                         ' application introducer
    outArray(aPtr + 1&) = 255                   ' application type
    outArray(aPtr + 2&) = 11                    ' block size
    ' here we write NETSCAPE2.0 followed by a byte value of 3 (block size)
    CopyMemory outArray(aPtr + 3&), 1398031694, 4&   ' NETS
    CopyMemory outArray(aPtr + 7&), 1162887491, 4&   ' CAPE
    CopyMemory outArray(aPtr + 11&), 53489202, 4&    ' 2.0 3
    outArray(aPtr + 15&) = 1                    ' write constant of 1 & then the loop count
    If MIS.GIFOverview.LoopCount < 0& Then MIS.GIFOverview.LoopCount = 0&
    CopyMemory outArray(aPtr + 16&), MIS.GIFOverview.LoopCount, 2& ' outArray(aPtr+18)=0
    aPtr = aPtr + 19&                           ' move pointer along
    
    ' WRITE EACH FRAME INTO THE OUTPUT ARRAY
    For f = 0& To MIS.Images - 1&
        X = fs(f).DataStart
        ' find start of block 44 and insert color table
        If m_BlankCanvas(X) = 33 Then           ' block 249 is in use, preceeded by 33 introducer
            CopyMemory outArray(aPtr), m_BlankCanvas(X), 8&
            X = X + 8&
        Else                                    ' create a blank block 249 entry
            outArray(aPtr) = 33
            outArray(aPtr + 1) = 249
            outArray(aPtr + 2) = 4
        End If
        If MIS.Image(f + lbOffset).GIFFrameInfo.DelayTime < 0 Then ' write the frame delay time
            outArray(aPtr + 4) = 0
        Else
            CopyMemory outArray(aPtr + 4), CLng(MIS.Image(f + lbOffset).GIFFrameInfo.DelayTime \ 10&), 2&
        End If
        Select Case (MIS.Image(f + lbOffset).GIFFrameInfo.DisposalCode And &H3&) ' write the frame disposal code
            Case 1&, 2&, 3&
                outArray(aPtr + 3&) = outArray(aPtr + 3&) Or ((MIS.Image(f + lbOffset).GIFFrameInfo.DisposalCode And &H3&) * 4&)
            Case Else
        End Select
        aPtr = aPtr + 8&                        ' move pointer along
        
        ' Block 44 will be provided for every frame, simply copy it; then modify X,Y offsets
        CopyMemory outArray(aPtr), m_BlankCanvas(X), 10&
        CopyMemory outArray(aPtr + 1&), MIS.Image(f + lbOffset).GIFFrameInfo.XOffset, 2&
        CopyMemory outArray(aPtr + 3&), MIS.Image(f + lbOffset).GIFFrameInfo.YOffset, 2&
        
        If fs(f).ColorTable = fs(gTblIdx).ColorTable Then
            aPtr = aPtr + 10&
        Else                                    ' write the local color table & packed byte
            outArray(aPtr + 9&) = (m_BlankCanvas(fPtr + 10&) And &H7&) Or 128
            If f = 0& Then                      ' use that table's sorted property
                outArray(aPtr + 9&) = outArray(aPtr + 9&) Or ((m_BlankCanvas(fPtr + 10&) And &H8) * 2&)
            End If
            CopyMemory outArray(aPtr + 10&), m_BlankCanvas(fs(f).ColorTable), fs(f).ColorCount
            aPtr = aPtr + fs(f).ColorCount + 10&
        End If
        
        ' CALCULATE SIZE OF COMPRESSED DATA WE NEED TO COPY
        lSize = fs(f).DataLen - fs(f).ColorCount - 24&   ' << 24=13 byte header + 10 byte descriptor + 1 byte termination code
        ' if a block 249 was provided, we remove that 8 byte block from what we'll be copying
        If m_BlankCanvas(fs(f).DataStart) = 33 Then lSize = lSize - 8&
        ' now copy the data
        CopyMemory outArray(aPtr), m_BlankCanvas(fPtr + fs(f).DataLen - lSize - 1&), lSize
        aPtr = aPtr + lSize                         ' move pointers along
        fPtr = fPtr + fs(f).DataLen
    Next
    outArray(aPtr) = 59                             ' add EOF marker
'    If UBound(outArray) <> aPtr Then Stop
    pvCreateAnimatedGIF = True

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
    ' FYI: FrameNr in this routine is zero-bound, not 1-bound

    Dim gHandle As Long, lPtr As Long, hLock As Long, hGraphics As Long
    Dim bmpDatSrc As BitmapData, bmpErase As BitmapData
    Dim boundsSrc As RECTI, boundsDst As RECTI, boundsErase As RECTI
    Dim tSA As SafeArray, aFrame() As Byte
    
    On Error GoTo ExitRoutine
    
'/////////// Get next frame \\\\\\\\\\\\
    ' We will update our template stream vs. re-creating it; faster overall
    With tSA                ' use a DMA overlay on the stream handle
        .cbElements = 1
        .cDims = 1
        .pvData = pvGetStreamAddress(hLock)
        If .pvData = 0& Then Exit Sub
        .rgSABound(0).cElements = m_GIFdata.GlobalOffset + m_GIFdata.Frames(FrameNr).DataLen + 1&
    End With                ' set the DMA handle
    CopyMemory ByVal VarPtrArray(aFrame), VarPtr(tSA), 4&
    
    With m_GIFdata.Frames(FrameNr)                              ' update the stream
        lPtr = m_GIFdata.GlobalOffset
        CopyMemory aFrame(lPtr), sourceData(.DataStart), .DataLen
        
        aFrame(lPtr + .DataLen) = 59
        CopyMemory aFrame(6), .Bounds.nWidth, 2&   ' make gif width same as frame
        CopyMemory aFrame(8), .Bounds.nHeight, 2&   ' make gif height same as frame
        CopyMemory aFrame(.ImgDescrip - .DataStart + lPtr + 1), 0&, 2& ' make gif left zero
        CopyMemory aFrame(.ImgDescrip - .DataStart + lPtr + 3), 0&, 2& ' make gif top zero & remove disposal code
        If .ImgDescrip > .DataStart Then aFrame(lPtr + 3) = (aFrame(lPtr + 3) And 227)
    End With
    CopyMemory ByVal VarPtrArray(aFrame), 0&, 4&                ' done; remove overlay & unlock pointer
    pvGetStreamAddress hLock
    tSA.pvData = 0&
    
    ' if all went well, we can create the GIF from the stream we just updated
    If GdipLoadImageFromStream(ObjPtr(m_IStream), gHandle) Then Exit Sub       ' error
    
'//////////  post-frame actions per previous frame's disposal code \\\\\\\\\\\\\
    If FrameNr = 0& Then    ' first frame, completely erase the image
        boundsErase.nHeight = m_GIFdata.Size.Y: boundsErase.nWidth = m_GIFdata.Size.X
        If GdipGetImageGraphicsContext(Handle, hGraphics) = 0& Then
            GdipGraphicsClear hGraphics, 0&
        End If
    Else
        Select Case m_GIFdata.Frames(m_GIFdata.Frame).Disposal
        Case 0, 1   ' do nothing, drawing remains unchanged
        Case 2      ' blank out previous frame's area
            boundsErase = m_GIFdata.Frames(m_GIFdata.Frame).Bounds
            If GdipGetImageGraphicsContext(Handle, hGraphics) = 0& Then
                GdipSetClipRectI hGraphics, boundsErase.nLeft, boundsErase.nTop, boundsErase.nWidth, boundsErase.nHeight, 0&
                GdipGraphicsClear hGraphics, 0&
            End If
        Case 3      ' reverting to previously cached scanlines
            boundsErase = m_GIFdata.Frames(m_GIFdata.Frame).Bounds
            boundsErase.nLeft = 0&: boundsErase.nWidth = m_GIFdata.Size.X
            If boundsErase.nHeight > m_GIFdata.Code3Size Then boundsErase.nHeight = m_GIFdata.Code3Size
            bmpErase.Scan0Ptr = VarPtr(m_BlankCanvas(0))
            bmpErase.stride = m_GIFdata.Size.X * 4& ' set scanwidth & update
            ' ^^ Tip: Unless retrieving just 1 row starting at left column, entire row will always be set/returned
            If GdipBitmapLockBits(Handle, boundsErase, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor32bppAlpha, bmpErase) = 0& Then
                GdipBitmapUnlockBits Handle, bmpErase
            End If                                  ' ensure our blank canvas is blank
        End Select
    End If
    
    ' set the area of the next frame
    boundsDst = m_GIFdata.Frames(FrameNr).Bounds
    boundsSrc = boundsDst

        
'////////// pre-frame actions per current frame's disposal code \\\\\\\\\\\\
    If m_GIFdata.Frames(FrameNr).Disposal = 3 Then
        ' with dispoal code of 3, need to cache current canvas & replace when next frame is rendered
        If hGraphics Then GdipDeleteGraphics hGraphics: hGraphics = 0&
        boundsErase = boundsDst
        boundsErase.nLeft = 0&: boundsErase.nWidth = m_GIFdata.Size.X
        If boundsErase.nHeight > m_GIFdata.Code3Size Then boundsErase.nHeight = m_GIFdata.Code3Size
        If GdipBitmapLockBits(Handle, boundsErase, ImageLockModeRead, lvicColor32bppAlpha, bmpDatSrc) = 0& Then
            CopyMemory m_BlankCanvas(0), ByVal bmpDatSrc.Scan0Ptr, bmpDatSrc.Height * bmpDatSrc.stride
            ' ^^ Tip: Unless retrieving just 1 row starting at left column, entire row will always be set/returned
            GdipBitmapUnlockBits Handle, bmpDatSrc
        End If
    End If
    
    ' some efficiency tweaks to prevent rendering GIF if possible
    If FrameNr = 0& Then                                        ' never render; just slide new image to our GIF handle
        lPtr = 0&
    ElseIf m_GIFdata.Frames(FrameNr).Transparency = 0 Then      ' no rendering needed if no transparency used
        lPtr = 0&
    ElseIf m_GIFdata.Frames(m_GIFdata.Frame).Disposal = 2 Then  ' if the previous frame's action was to erase, then
        If boundsDst.nTop >= boundsErase.nTop Then              ' if this frame is <= previous frame's dimensions
            If boundsDst.nLeft >= boundsErase.nLeft Then        ' then we don't need to render
                If boundsErase.nWidth + boundsErase.nLeft >= boundsDst.nLeft + boundsDst.nWidth Then
                    If boundsErase.nHeight + boundsErase.nTop >= boundsDst.nTop + boundsDst.nHeight Then lPtr = 0&
                End If
            End If
        End If
    End If
        
'////////// rendering actions \\\\\\\\\\\\\
    If lPtr = 0& Then
        boundsSrc.nTop = 0&: boundsSrc.nLeft = 0&
        If hGraphics Then GdipDeleteGraphics hGraphics: hGraphics = 0&
        If GdipBitmapLockBits(gHandle, boundsSrc, ImageLockModeRead, lvicColor32bppAlpha, bmpDatSrc) = 0& Then
            bmpErase = bmpDatSrc
            If GdipBitmapLockBits(Handle, boundsDst, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor32bppAlpha, bmpErase) = 0& Then
                GdipBitmapUnlockBits Handle, bmpErase            ' update the GIF now
            End If
            GdipBitmapUnlockBits gHandle, bmpDatSrc
        End If
        
    Else
        If hGraphics = 0& Then Call GdipGetImageGraphicsContext(Handle, hGraphics) Else GdipResetClip hGraphics
        If hGraphics Then
            With boundsDst                                      ' render source first, then new frame
                GdipDrawImageRectRectI hGraphics, gHandle, .nLeft, .nTop, .nWidth, .nHeight, 0&, 0&, .nWidth, .nHeight, UnitPixel, 0&, 0&, 0&
            End With
        End If
    End If

ExitRoutine:
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(aFrame), 0&, 4&
    If hGraphics Then GdipDeleteGraphics hGraphics              ' clean up
    If gHandle Then GdipDisposeImage gHandle
    m_GIFdata.Frame = FrameNr                                   ' track current frame

End Sub

Public Property Get LoopCount() As Long
    LoopCount = m_GIFdata.LoopCount
End Property

Public Function MoveToFrame(ByVal FrameNr As Long, ByRef sourceInfo As Variant, ByVal Handle As Long) As Boolean

    Dim bSkippedFrames As Boolean
    Dim sourceData() As Byte
    
    If FrameNr = m_GIFdata.Frame Then
        Exit Function
    ElseIf FrameNr >= m_GIFdata.FrameCount Then
        FrameNr = 0&
    ElseIf FrameNr < 0& Then
        FrameNr = m_GIFdata.FrameCount - 1&
        bSkippedFrames = True
    ElseIf FrameNr < m_GIFdata.Frame Then
        bSkippedFrames = (FrameNr > 0&)
    ElseIf FrameNr > m_GIFdata.Frame + 1& Then
        bSkippedFrames = True
    End If
    
    ' Because each frame is not rendered ahead of time and cached, when moving more than 1 frame forward
    ' or moving backwards, we must render 2 or more frames in order, to produce the final result
    
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

Private Sub pvProcessGIFfile(inStream() As Byte)

    ' GIF parser. Parses key data, does not decompress GIF data
    Dim X As Long
    Dim lGlobalCount As Long
    Dim lPtr As Long, lMax As Long
    Dim lFrameSize As Long, lSize As Long
    Dim gifFrame As FRAMESTRUCT, newFrame As FRAMESTRUCT
    Dim bCode3 As Boolean, lBounds As Long
    
    With m_GIFdata
        Erase .Frames
        .FrameCount = 0&
        .LoopCount = 0&
        .Size.X = 0&
        .Size.Y = 0&
        .Frame = 0&
        .Code3Size = 0&
    End With
    
    lMax = UBound(inStream) + 1&                                ' prevent reading past array bounds
    
    m_GIFdata.Size.X = inStream(6) Or inStream(7) * &H100&      ' get overall dimensions
    m_GIFdata.Size.Y = inStream(8) Or inStream(9) * &H100&
    If (inStream(10) And 128&) Then                             ' global color table used?
        With gifFrame                                           ' if so, cache location & size
            lGlobalCount = (2& ^ ((inStream(10) And &H7&) + 1&)) * 3&
            .ColorTable = 13&
            .ColorCount = lGlobalCount
            lPtr = .ColorTable + lGlobalCount                   ' set pointer to where next GIF block starts
        End With
    Else
        lPtr = 13&                                              ' set pointer to where next GIF block starts
    End If
    
    Do While lPtr < lMax                                        ' process the other GIF blocks
    
        Select Case inStream(lPtr)
        Case 0                                                  ' shouldn't get here; but just in case
            lPtr = lPtr + 1&
        Case 33                                                 ' application extensions (optional)
            Select Case inStream(lPtr + 1&)
            Case 255                                            ' looking for loop count if it  exists
                If lPtr + inStream(lPtr + 2&) + 3& > lMax Then Exit Do ' corrupt gif
                ReDim m_BlankCanvas(0 To inStream(lPtr + 2&) - 1&)
                CopyMemory m_BlankCanvas(0), inStream(lPtr + 3&), inStream(lPtr + 2&)
                If LCase(StrConv(m_BlankCanvas(), vbUnicode)) = "netscape2.0" Then
                    If inStream(lPtr + inStream(lPtr + 2&) + 3&) = 3& Then
                        If lPtr + 17& > lMax Then Exit Do        ' corrupt gif
                        m_GIFdata.LoopCount = inStream(lPtr + 16&) Or inStream(lPtr + 17&) * &H100&
                    End If
                End If
                Erase m_BlankCanvas()
            Case 249                                            ' image control block
                If lPtr + 3& > lMax Then Exit Do                ' corrupt gif
                With gifFrame
                    .DataStart = lPtr                           ' cache where this starts & extract key info
                    .Disposal = (inStream(lPtr + 3&) And &H1C&) \ &H4
                    If .Disposal = 3 Then
                        ' Don't let first frame have disposal code of 3. It's pointless
                        If m_GIFdata.FrameCount = 0& Then .Disposal = 2 Else bCode3 = True
                    End If
                    .Transparency = (inStream(lPtr + 3&) And 1)
                End With
            Case Else
            End Select                                          ' don't care about other types of app blocks (i.e. text,comments,etc)
            lPtr = lPtr + inStream(lPtr + 2&) + 3&              ' go onto the next block
            If pvReadPastBlock(lPtr, inStream(), lMax) = False Then Exit Do
        
        Case 44                                                 ' image descriptor (required)
            With gifFrame
                If lPtr + 11& > lMax Then Exit Do               ' corrupt GIF
                If .DataStart = 0& Then .DataStart = lPtr
                .ImgDescrip = lPtr                              ' cache where this block starts
                CopyMemory .Bounds.nLeft, inStream(lPtr + 1&), 2&
                CopyMemory .Bounds.nTop, inStream(lPtr + 3&), 2&
                CopyMemory .Bounds.nWidth, inStream(lPtr + 5&), 2&
                CopyMemory .Bounds.nHeight, inStream(lPtr + 7&), 2&
'Debug.Print m_GIFdata.FrameCount; .Bounds.nLeft; .Bounds.nTop
                If bCode3 Then
                    ' tracking largest array needed to support disposal code 3 via CopyMemory logic
                    ' ensure a mal-formatted GIF frame size won't cause crash
                    lBounds = .Bounds.nHeight
                    If .Bounds.nTop < 0& Then lBounds = lBounds + .Bounds.nTop
                    If lBounds > m_GIFdata.Size.Y Then lBounds = m_GIFdata.Size.Y
                    If lBounds > m_GIFdata.Code3Size Then m_GIFdata.Code3Size = lBounds
                    bCode3 = False
                End If
            End With
            If (inStream(lPtr + 9&) And 128&) Then              ' local color table used?
                gifFrame.ColorTable = lPtr + 10&                ' if so, cache location & size
                gifFrame.ColorCount = 2& ^ ((inStream(lPtr + 9&) And &H7&) + 1&) * 3&
                lPtr = lPtr + gifFrame.ColorCount + 11&         ' adjust pointer to start of compressed data
                If lPtr > lMax Then Exit Do                     ' corrupt gif
            ElseIf lGlobalCount Then
                gifFrame.ColorTable = 13&                       ' no local table, uses global table
                gifFrame.ColorCount = lGlobalCount
                lPtr = lPtr + 11&                               ' adjust pointer to start of compressed data
            Else
                Exit Do                                         ' no local, no global, abort.
            End If                                              ' FYI: Per GIF specs, the decoder can supply any palette desired
            With m_GIFdata
                ReDim Preserve .Frames(0 To .FrameCount)        ' append this frame to the frame collection
                .Frames(.FrameCount) = gifFrame
            End With
            If pvReadPastBlock(lPtr, inStream(), lMax) = False Then Exit Do ' skip to next GIF block
            lSize = lPtr - gifFrame.DataStart                   ' calculate size of the GIF data for this frame
            m_GIFdata.Frames(m_GIFdata.FrameCount).DataLen = lSize
            m_GIFdata.FrameCount = m_GIFdata.FrameCount + 1&
            If lSize > lFrameSize Then lFrameSize = lSize       ' track largest frame, data-size
            gifFrame = newFrame
        Case Else
            Exit Do                                             ' 59=GIF trailer. We will abort if any other unknown blocks appear
        End Select
    Loop
    
    If m_GIFdata.FrameCount Then                                ' some finalizing now
        lFrameSize = lFrameSize + lGlobalCount + 15&            ' size of a IStream we will create
        ReDim m_BlankCanvas(0 To lFrameSize - 1&)
        CopyMemory m_BlankCanvas(0), inStream(0), 13&           ' copy the GIF header info
        If lGlobalCount Then                                    ' and global color table if it exists
            CopyMemory m_BlankCanvas(13), inStream(13), lGlobalCount
            m_GIFdata.GlobalOffset = 13& + lGlobalCount
        Else
            m_GIFdata.GlobalOffset = 13&                        ' position where in stream, each frame data will start
        End If                                                  ' create the stream
        Set m_IStream = modCommon.IStreamFromArray(VarPtr(m_BlankCanvas(0)), lFrameSize)
        If m_IStream Is Nothing Then                            ' abort if failure
            Erase m_GIFdata.Frames()
            Erase m_BlankCanvas()
            m_GIFdata.FrameCount = 0&
        End If
    End If
    
End Sub

Private Function pvReadPastBlock(lPtr As Long, bData() As Byte, DataLen As Long) As Boolean
    
    ' helper function for GIF parser. Skips a GIF block & stops on next block
    
    Do Until bData(lPtr) = 0
        lPtr = lPtr + bData(lPtr) + 1&
        If lPtr >= DataLen Then Exit Function
    Loop
    lPtr = lPtr + 1&
    pvReadPastBlock = True
End Function

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
    Set m_IStream = Nothing
End Sub

Private Sub Class_Terminate()
    ' clean up
    Set m_IStream = Nothing
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub
