VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColorReduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Color quantization & forced palette conversion routines ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' More complex than I initially imagined; but mostly due to image format constraints.
' Some formats can use transparency in palettes while others cannot, while some use masks instead.
' Regarding palettes only (256 colors or less)...
' -- No transparency
'       JPG,BMP,PCX will always have any transparency filled before palette reduction occurs
' -- Simple transparency is where a pixel is either 100% transparent or 100% opaque, nothing else:
'       GIF,PNG,TGA,TIF support simple transparency. GDI+ bitmaps do, but standard GDI bitmaps do not
' -- Complex transparency is where a pixel can have any level of transparency/opaqueness
'       PNG,TGA. ICO/CUR/BMP/PCX/TIF support this but not in paletted form
' -- Palette that never contains transparency, but simple transparency represented by a separate mask
'       ICO,CUR
' -- WMF/EMF; any palette reduction may change color values but result will be in WMF/EMF format

' Default/colorless bit reduction can be applied to all image formats. The logic is simple and
'   straightforward. Any bit reduction for colorless ensures no color loss and no alpha loss
' -- If 32bpp, see if image can be dropped to 24bpp, 8bpp, 4pp or 1bpp (2 colors, not necessarily B&W)
'       If alpha exists, BMP,PCX cannot be reduced using colorless reduction
' -- If 24bpp or lower, see if image can be dropped to 8bpp, 4bpp, 1bpp(2 colors) respectively
' -- If 32/24bpp format, colorless reduction is the only type of reduction applied in this routine
'       unless reduction to a palette is requested.

' Reduction to palette can produce color loss and/or alpha loss. The AlphaTolerancePct value is always applied
' to palette reduction. That value converts complex transparency to simple transparency
'   -- GIF,TGA. Regardless of eventual bit depth, palette indexes must be returned as 8bpp. Format/GDI+ requirements
'   -- PCX,BMP. Fill color is always applied if transparency exists. Palette will not contain transparency
'   -- PCX,TIF. If reduction is 1bpp and not black&white then paletted indexes must be returned as 4bpp
'   -- PNG,TGA,WMF/EMF. Indexes can contain alpha values. WMF/EMF always result in metafile format
'       :: TGA has a special case where if palette is grayscale, then no palette is used
'   -- ICO,CUR. Indexes will not contain transparency, a mask defines simple transparency
'   -- JPG. Result is always 24bpp.
' The palette type and color reduction type determine the eventual bit depth of the paletted image
'-----------------------------------------------------------------------------------------

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(1 To 256) As Long
End Type

Private Declare Function CreatePalette Lib "gdi32.dll" (ByRef lpLogPalette As LOGPALETTE) As Long
Private Declare Function GetNearestPaletteIndex Lib "gdi32.dll" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetPaletteEntries Lib "gdi32.dll" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, ByRef lpPaletteEntries As Any) As Long

Private Type OCTTREENODE
    Pixels As Long              '// Number of pixels represented by this leaf
    Reds As Long                '// Sum of red components
    Greens As Long              '// Sum of green components
    Blues As Long               '// Sum of blue components
    NextNode As Long            '// Pointer to next reducible node
    Child(0 To 7) As Long       '// Pointers to child nodes
    isLeaf As Boolean           '// TRUE if node has no children
End Type
Private Type OCTTREE
    Nodes As Long               '// number of active nodes in tree
    RecycleIndex As Long        '// index to next recyclable node
    MaxDepth As Long            '// tree depth: max is 7, lower values equal greater speed, less adaptive
    Node() As OCTTREENODE       '// the active nodes
    ReducibleNodes() As Long    '// lowest level node refs that can be combined with higher level nodes
End Type

Private pTree As OCTTREE
Private pPalette As ColorPalette

Public Property Get Count() As Long
    Count = pPalette.Count  ' return the number of palette entries
End Property

Public Function IsGrayscale(Optional ProperOnly As Boolean) As Boolean
    ' TGA has a special case for grayscale images. Function test for grayscale on sorted palette
    Dim p As Long, g As Long, gl As Long
    gl = -1& ' set last grayscale value processed
    For p = 1& To pPalette.Count
        If (pPalette.Entries(p) And &HFF) <> ((pPalette.Entries(p) And &HFF00&) \ &H100&) Then Exit For
        If (pPalette.Entries(p) And &HFF) <> ((pPalette.Entries(p) And &HFF0000) \ &H10000) Then Exit For
        If (pPalette.Entries(p) And &HFF000000) <> &HFF000000 Then Exit For
        If ProperOnly Then
            ' testing for incrementing grayscale values
            g = (pPalette.Entries(p) And &HFF&)
            If g > gl Then gl = g Else Exit For
        End If
    Next
    IsGrayscale = (p > pPalette.Count)

End Function

Public Function IsBlackWhite() As Boolean
    If pPalette.Count = 2& Then
        IsBlackWhite = (pPalette.Entries(1) = &HFF000000 And pPalette.Entries(2) = &HFFFFFFFF)
    End If
End Function

Public Function ImportRemotePalette(ImageHandle As Long) As Boolean
    ' extracts a bitmap's palette to our class palette
    If ImageHandle Then
        Dim pSize As Long
        If GdipGetImagePaletteSize(ImageHandle, pSize) = 0& Then
            ImportRemotePalette = (GdipGetImagePalette(ImageHandle, pPalette, pSize) = 0&)
        End If
    End If

End Function

Public Function CopyLocalPalette(DestPointer As Long, RGBorder As Boolean, RGBLength As Long) As Long
    
    ' applies class palette to passed pointer in RGB, ARGB order
    Dim p As Long
    If pPalette.Count Then
        If RGBorder Then
            For p = 1 To pPalette.Count
                CopyMemory ByVal DestPointer, CLng((pPalette.Entries(p) And &HFF0000) \ &H10000 Or _
                                            (pPalette.Entries(p) And &HFF00&) Or _
                                            (pPalette.Entries(p) And &HFF&) * &H10000), RGBLength
                DestPointer = DestPointer + RGBLength
            Next
        ElseIf RGBLength = 4& Then
            CopyMemory ByVal DestPointer, pPalette.Entries(1), pPalette.Count * RGBLength
        Else
            For p = 1 To pPalette.Count
                CopyMemory ByVal DestPointer, pPalette.Entries(p), RGBLength
                DestPointer = DestPointer + RGBLength
            Next
        End If
        CopyLocalPalette = pPalette.Count
    End If
End Function

Public Function CopyRemotePalette(ByVal DestPointer As Long, ImageHandle As Long, RGBorder As Boolean, RGBLength As Long) As Long
    
    ' applies a bitmap's palette to passed pointer in RGB, ARGB order
   
    Dim cPal As ColorPalette, p As Long
    
    If ImageHandle = 0& Then Exit Function
    GdipGetImagePaletteSize ImageHandle, p
    If p = 0& Then Exit Function
    GdipGetImagePalette ImageHandle, cPal, p
    
    If cPal.Count Then
        If RGBorder Then
            For p = 1 To cPal.Count
                CopyMemory ByVal DestPointer, CLng((cPal.Entries(p) And &HFF0000) \ &H10000 Or _
                                            (cPal.Entries(p) And &HFF00&) Or _
                                            (cPal.Entries(p) And &HFF&) * &H10000), RGBLength
                DestPointer = DestPointer + RGBLength
            Next
        ElseIf RGBLength = 4& Then
            CopyMemory ByVal DestPointer, cPal.Entries(1), cPal.Count * RGBLength
        Else
            For p = 1 To cPal.Count
                CopyMemory ByVal DestPointer, cPal.Entries(p), RGBLength
                DestPointer = DestPointer + RGBLength
            Next
        End If
        CopyRemotePalette = cPal.Count
    End If

End Function

Public Function PalettizeToHandle(Handle As Long, _
                AlphaLimit As AlphaReductionLimit, _
                SS As SAVESTRUCT, Optional PalIndexOptions As RawDataOrientation) As Long

    ' color reduces passed image to a GDI+ handle using various reduction algos
    Dim bDummy() As Byte, lOrientation As RawDataOrientation
    
    pPalette.Count = 0&: pPalette.Flags = 0&
    If Handle = 0& Then Exit Function
    
    lOrientation = orient_GDIpHandle Or (PalIndexOptions And Not (orient_BottomUp Or orient_TopDown Or orient_WantMask))
    If SS.reserved1 And &H10000000 Then                 ' no forced color reduction, default colorless reduction
        PalettizeToHandle = pvColorlosslessReduce(Handle, bDummy(), lOrientation, 0&, 0&, AlphaLimit, SS, 0&)
    ElseIf SS.PaletteType = lvicPaletteAdaptive Then
        PalettizeToHandle = pvCreateAdaptivePalette(Handle, bDummy(), lOrientation, 0&, 0&, AlphaLimit, SS)
    Else
        PalettizeToHandle = pvCreateStandardPalette(Handle, bDummy(), lOrientation, 0&, 0&, AlphaLimit, SS)
    End If
    If PalettizeToHandle = 0& Then pPalette.Count = 0&

End Function

Public Function PalettizeToArray(Handle As Long, outData() As Byte, _
                DataOrientation As RawDataOrientation, _
                HeaderSize As Long, FooterSize As Long, _
                AlphaLimit As AlphaReductionLimit, _
                SS As SAVESTRUCT) As Long

        
    ' color reduces passed image to an array using various reduction algos
    pPalette.Count = 0&: pPalette.Flags = 0&
    If Handle = 0& Then Exit Function

    If SS.reserved1 And &H10000000 Then                 ' no forced color reduction, default colorless reduction
        PalettizeToArray = pvColorlosslessReduce(Handle, outData(), DataOrientation, HeaderSize, FooterSize, AlphaLimit, SS, 0&)
    ElseIf SS.PaletteType = lvicPaletteAdaptive Then
        PalettizeToArray = pvCreateAdaptivePalette(Handle, outData(), DataOrientation, HeaderSize, FooterSize, AlphaLimit, SS)
    Else
        PalettizeToArray = pvCreateStandardPalette(Handle, outData(), DataOrientation, HeaderSize, FooterSize, AlphaLimit, SS)
    End If
    If PalettizeToArray = 0& Then pPalette.Count = 0&
    
End Function

Public Function ApplyAlphaTolerance(Handle As Long, ByVal TolerancePct As Long, ByVal Width As Long, ByVal Height As Long, bReplaceImage As Boolean) As Long

    ' reduces complex transparency to simple transparency, no color reduction performed
    ' passed image is always 32bpp
    
    Dim sizeI As RECTI
    Dim tBMPsrc As BitmapData, tBMPdst As BitmapData
    Dim srcSA As SafeArray, dstSA As SafeArray
    Dim srcData() As Long, dstData() As Long
    Dim X As Long, wrkPic As Long, Alpha As Long
    
    sizeI.nHeight = Height: sizeI.nWidth = Width
    If GdipCreateBitmapFromScan0(Width, Height, 0&, lvicColor32bppAlpha, ByVal 0&, wrkPic) Then Exit Function
    If GdipBitmapLockBits(wrkPic, sizeI, ImageLockModeWrite, lvicColor32bppAlpha, tBMPdst) Then
        GdipDisposeImage wrkPic
        Exit Function
    End If
    With dstSA
        .cbElements = 4
        .cDims = 1
        .pvData = tBMPdst.Scan0Ptr
        .rgSABound(0).cElements = sizeI.nHeight * sizeI.nWidth
    End With
    If GdipBitmapLockBits(Handle, sizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMPsrc) Then
        GdipDisposeImage wrkPic
        Exit Function
    End If
    srcSA = dstSA
    srcSA.pvData = tBMPsrc.Scan0Ptr
    CopyMemory ByVal VarPtrArray(dstData), VarPtr(dstSA), 4&
    CopyMemory ByVal VarPtrArray(srcData), VarPtr(srcSA), 4&
    On Error GoTo ExitRoutine
    
    For X = 0& To sizeI.nWidth * sizeI.nHeight - 1&
        If (srcData(X) And &HFF000000) = &HFF000000 Then
            dstData(X) = srcData(X)
        Else
            Alpha = (srcData(X) And &H7F000000) \ &H1000000
            If srcData(X) < 0& Then Alpha = Alpha Or &H80
            If Alpha > TolerancePct Then            ' keep alpha value
                dstData(X) = srcData(X)
            Else                                    ' remove alpha value
                dstData(X) = srcData(X) And &HFFFFFF
            End If
        End If
    Next
    
ExitRoutine:                                        ' clean up
    CopyMemory ByVal VarPtrArray(dstData), 0&, 4&
    CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
    GdipBitmapUnlockBits Handle, tBMPsrc
    GdipBitmapUnlockBits wrkPic, tBMPdst
    If Err Then
        GdipDisposeImage wrkPic
    ElseIf bReplaceImage Then
        GdipDisposeImage Handle
        Handle = wrkPic
    End If
    ApplyAlphaTolerance = wrkPic
End Function

Public Function pvCreateAdaptivePalette(Handle As Long, outData() As Byte, _
                                            DataOrientation As RawDataOrientation, _
                                            ByVal HeaderSize As Long, FooterSize As Long, _
                                            AlphaLimit As AlphaReductionLimit, _
                                            SS As SAVESTRUCT) As Long

    ' Oct-Tree color quantization. The logic I think was well decscribed on this link...
    ' ACKNOWLEDGEMENTS: http://www.microsoft.com/msj/archive/S3F1.aspx
    ' And converted from this source: http://www.microsoft.com/msj/archive/s3f1a.htm

    ' Works well for 4bpp & 8bpp; not so good for black and white. This is a limitation of the logic.
    ' When trying to reduce to 2 colors, the tree-trimming logic combines all child nodes of a specific
    ' node to average the colors. Well, if the tree is down to just 1 parent with several children (colors),
    ' all of them are combined and you are now left with just 1 color. A tweak of the logic has been employed
    ' to prevent this scenario; but its end result may not always be ideal; especially if the last two
    ' remaining colors are similar in lightness/darkness.
    
    If SS.ColorDepth = lvicConvert_BlackWhite Then  ' handle B&W differently
        pvCreateAdaptivePalette = pvCreateBlackAndWhite(Handle, outData(), DataOrientation, HeaderSize, FooterSize, AlphaLimit, SS)
        Exit Function
    End If
    
    ' first see if this can be done as-is, without quantization
    pPalette.Count = 0&: pPalette.Flags = 0&
    pvCreateAdaptivePalette = pvColorlosslessReduce(Handle, outData(), DataOrientation, HeaderSize, FooterSize, _
                                AlphaLimit, SS, modCommon.ColorDepthToColorType(SS.ColorDepth, Handle))
    If pvCreateAdaptivePalette Then Exit Function

    Dim sizeI As RECTI, tBMP As BitmapData, dBMP As BitmapData
    Dim tSA As SafeArray, inStream() As Long
    Dim dSA As SafeArray, outStream() As Byte
    Dim I As Long, J As Long, MaxColorCount As Long
    Dim hPal As Long
    Dim nShift As Long, lShift As Long
    Dim nLeafCount As Long, nodeLevel As Long
    Dim lPrevColor As Long, lLeafIndex As Long
    Dim r As Integer, g As Integer, b As Integer
    Dim cIndex As Long, cOffset As Long, bNew As Boolean
    Dim pMask() As Byte, mPtr As Long, mShift As Byte
    Dim mScanWidth As Long, mOffset As Long
    Dim bWantMask As Boolean, lIcoTransIndex As Long

    Select Case modCommon.ColorDepthToColorType(SS.ColorDepth, Handle) ' set max depth allowed
        Case lvicColor1bpp: MaxColorCount = 2&
        Case lvicColor4bpp: MaxColorCount = 16&
        Case Else: MaxColorCount = 256&
    End Select
    pPalette.Count = 0&: pPalette.Flags = 0&
    bWantMask = (DataOrientation And orient_WantMask)
    If (SS.reserved1 And 1&) Then                           ' image has transparency
        If bWantMask = False Then                           ' need palette entry for transparency?
            If MaxColorCount > 2& Then MaxColorCount = MaxColorCount - 1&
            pPalette.Flags = 1&                             ' flag indicating reduced palette; we add it back later
        End If
    End If
    
    sizeI.nHeight = SS.Height: sizeI.nWidth = SS.Width      ' open image for reading
    If GdipBitmapLockBits(Handle, sizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
    
    ' the mask is very important. We use it to determine which pixels fail alpha-tolerance tests and
    ' to indicate whether pixel will be searched for in the palette or simply assigned a static palette index
    ' If being called for icons/cursors, the mask will be returned as part of the outData array
    mScanWidth = modCommon.ByteAlignOnWord(1, SS.Width)
    ReDim pMask(0 To mScanWidth * SS.Height - 1&)
    
    With tSA
        .cbElements = 4
        .cDims = 2
        .pvData = tBMP.Scan0Ptr
        If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
        .rgSABound(0).cElements = SS.Height
        .rgSABound(1).cElements = SS.Width
    End With
    CopyMemory ByVal VarPtrArray(inStream), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine ' begin error trapping
        
    With pTree
        .Nodes = 1&                             ' how many active nodes in the tree
        .MaxDepth = 7&                          ' max of 7 bits per R,G,B channel to be used
        .RecycleIndex = 0&                      ' nodes removed that can be re-added later
        ReDim .Node(1 To MaxColorCount)         ' arbritrary start count
        ReDim .ReducibleNodes(0 To .MaxDepth)   ' indexes to nodes that can be reduced (linked list)
    End With
    
    lPrevColor = lPrevColor Xor 1&              ' start processing image, top to bottom
    
    For I = 0& To sizeI.nHeight - 1&            ' mask position will always be bottom-up
        mPtr = (SS.Height - I - 1&) * mScanWidth: mShift = &H80
        With SS
            For J = 0& To sizeI.nWidth - 1&
                If lPrevColor <> inStream(J, I) Then    ' same color as before?
                    lPrevColor = inStream(J, I)         ' if not, cache and process
                    If (lPrevColor And &HFF000000) = &HFF000000 Then
                        lLeafIndex = 1&                 ' no transparency, use as-is
                    Else                                ' does pixel pass alpha tolerance test?
                        r = (lPrevColor And &H7F000000) \ &H1000000
                        If lPrevColor < 0& Then r = r Or &H80
                        If r > .AlphaTolerancePct Then lLeafIndex = 1& Else lLeafIndex = 0&
                    End If                              ' if test passed, add color else don't
                End If
                If lLeafIndex Then                      ' color to be added; extract R,G,B channels
                    r = lPrevColor And &HFF&
                    g = (lPrevColor And &HFF00&) \ &H100&
                    b = (lPrevColor And &HFF0000) \ &H10000
                    nodeLevel = -1&: nShift = &H100&    ' reset each pass
                    Do
                        If pTree.Node(lLeafIndex).isLeaf Then ' tally pixel count & channel count
                            With pTree.Node(lLeafIndex)
                                .Pixels = .Pixels + 1&
                                .Reds = .Reds + r
                                .Greens = .Greens + g
                                .Blues = .Blues + b
                            End With
                            Exit Do
                        Else
                            nodeLevel = nodeLevel + 1&  ' find child index for this color part
                            nShift = nShift \ 2&
                            cOffset = (b And nShift) \ nShift Or ((g And nShift) \ nShift) * 2& Or ((r And nShift) \ nShift) * 4&
                            cIndex = pTree.Node(lLeafIndex).Child(cOffset)
                            If cIndex = 0& Then         ' if no child, need to create one
                                cIndex = pvCreateNode(nodeLevel, nLeafCount)
                                pTree.Node(lLeafIndex).Child(cOffset) = cIndex
                            End If
                            lLeafIndex = cIndex         ' keep track of current node index
                        End If
                    Loop
                    If nLeafCount > MaxColorCount Then  ' if too many colors; reduce to lower count
                        If pvTrimTree(nLeafCount, MaxColorCount) = False Then
                            MaxColorCount = -MaxColorCount: Exit For    ' flag indicating reduction to costly
                        End If                          ' see pvTrimTree for a bit more
                    End If
                Else
                    pMask(mPtr) = pMask(mPtr) Or mShift ' keep track of alpha pixels
                End If                                  ' and update mask shift state & pointer movement
                If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
            Next
        End With
        If MaxColorCount < 0& Then                      ' crappy, I know. But need to finish of this scanline
            For J = J To SS.Width - 1&                  ' the mask is required for the 2nd half of the routine
                If (inStream(J, I) And &HFF000000) <> &HFF000000 Then pMask(mPtr) = pMask(mPtr) Or mShift
                If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
            Next
            Exit For
        End If
    Next
    If MaxColorCount < 0& Then                          ' yep, still crappy. Now we need to finish off the image
        MaxColorCount = -MaxColorCount                  ' this error checking only happens when reducing to
        For I = I + 1& To SS.Height - 1&                ' very low bit depths
            mPtr = (SS.Height - I - 1&) * mScanWidth: mShift = &H80
            For J = 0& To SS.Width - 1&
                If (inStream(J, I) And &HFF000000) <> &HFF000000 Then pMask(mPtr) = pMask(mPtr) Or mShift
                If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
            Next
        Next
    End If
    
    pvGetTreeColors 1&, pPalette.Count                  ' retrieve palette colors
    Erase pTree.ReducibleNodes()                        ' release some memory
    Erase pTree.Node
    
    If bWantMask Then                                   ' patch to ensure a opaque black index exists
        lIcoTransIndex = pvFixupIconPalette(MaxColorCount) - 1&
        If lIcoTransIndex = -1& Then                    ' doesn't exist, try again
            hPal = pvConvertPaletteToGDI(MaxColorCount, 0&, True) ' create palette reserving spot
            pPalette.Count = pPalette.Count - 1&        ' reduce count & call again
            lIcoTransIndex = pvFixupIconPalette(MaxColorCount) - 1&
            If lIcoTransIndex = -1& Then                ' should not fail here, but if it does...
                DeleteObject hPal
                hPal = 0&
            End If
        Else                                            ' black exists; create palette
            hPal = pvConvertPaletteToGDI(lvicPaletteAdaptive, 0&)
        End If
    Else
        hPal = pvConvertPaletteToGDI(lvicPaletteAdaptive, 0&) ' convert to GDI palette
    End If
    If hPal = 0& Then GoTo ExitRoutine                  ' failure; abort
    
    If (DataOrientation And orient_BlackIs1Not0) Then
        For I = 1 To pPalette.Count
            If pPalette.Entries(I) = &HFF000000 Then pPalette.Entries(I) = &HFF000001
        Next
    End If
    If pPalette.Flags = 1& Then                         ' add back palette entry we took away at top of routine
        pPalette.Count = pPalette.Count + 1&            ' make its entry the backcolor at full transparency
        pPalette.Entries(pPalette.Count) = (SS.RSS.FillColorARGB And &HFFFFFF)
    End If
    
    If MaxColorCount = 2& Then
        With pPalette                                   ' ensure darker color first; this will be changed to Black & White later
            If (.Entries(1) And &HFF) + ((.Entries(1) And &HFF00&) \ &H100&) + ((.Entries(1) And &HFF0000) \ &H10000) \ 3 < _
                (.Entries(2) And &HFF) + ((.Entries(2) And &HFF00&) \ &H100&) + ((.Entries(2) And &HFF0000) \ &H10000) \ 3 Then
                I = .Entries(1): .Entries(1) = .Entries(2): .Entries(2) = I
            End If
        End With
    End If
    If (DataOrientation And orient_8bppIndexes) Then
        pvCreateAdaptivePalette = lvicColor8bpp: nShift = &H1
    Else
        Select Case pPalette.Count                      ' determine destination bitdepth
            Case Is < 3&
                If (DataOrientation And orient_4bppIndexesMin) Then
                    nShift = &H10: pvCreateAdaptivePalette = lvicColor4bpp
                Else
                    nShift = &H2: pvCreateAdaptivePalette = lvicColor1bpp
                End If
            Case Is < 17&: nShift = &H10: pvCreateAdaptivePalette = lvicColor4bpp
            Case Else: nShift = &H1: pvCreateAdaptivePalette = lvicColor8bpp
        End Select
    End If
    ' added checks below for orient_PNMformat. PNM files cannot use paletted images unless grayscale or b&w
    If (DataOrientation And orient_SortGrayscale) Then
        If pPalette.Flags = 0& And pPalette.Count > 2& Then
            If Me.IsGrayscale() And pPalette.Count = 256& Then
                For I = 0& To pPalette.Count - 1&
                    pPalette.Entries(I + 1&) = I Or I * &H100& Or I * &H10000 Or &HFF000000
                Next
            End If
        End If
    End If
    If (DataOrientation And orient_PNMformat) Then
        If Me.IsGrayscale() Then
            If pPalette.Count = 2& Then
                If Not (pPalette.Entries(1) = &HFF000000 And pPalette.Entries(2) = &HFFFFFFFF) Then pvCreateAdaptivePalette = lvicColor24bpp
            End If
        Else
            pvCreateAdaptivePalette = lvicColor24bpp
        End If
    End If
                                                        ' destination scanwidth
    cOffset = modCommon.ByteAlignOnWord((pvCreateAdaptivePalette And &HFF00&) \ &H100&, SS.Width)
    If (DataOrientation And &H3) = 0& Then              ' create return image & ensure we can open it
        dBMP.PixelFormat = pvCreateAdaptivePalette
        If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, dBMP.PixelFormat, ByVal 0&, pvCreateAdaptivePalette) Then GoTo ExitRoutine
        If GdipBitmapLockBits(pvCreateAdaptivePalette, sizeI, ImageLockModeWrite, dBMP.PixelFormat, dBMP) Then
            GdipDisposeImage pvCreateAdaptivePalette    ' failure; abort
            pvCreateAdaptivePalette = 0&
            GoTo ExitRoutine
        End If
        dSA.pvData = dBMP.Scan0Ptr                      ' we'll be writing to here
        dSA.rgSABound(0).cElements = cOffset * SS.Height
    Else
        If (DataOrientation And orient_WantPaletteInArray) Then
            I = 2 ^ ((pvCreateAdaptivePalette And &HFF00&) \ &H100) * 4&
        Else
            I = 0&
        End If
        If bWantMask Then
            dSA.rgSABound(0).cElements = (cOffset + mScanWidth) * SS.Height
        Else
            dSA.rgSABound(0).cElements = cOffset * SS.Height
        End If
        ReDim outData(0 To HeaderSize + FooterSize + dSA.rgSABound(0).cElements + I - 1&)
        dSA.pvData = VarPtr(outData(HeaderSize + I))    ' we'll be writing to here
    End If
    dSA.cbElements = 1
    dSA.cDims = 1
    CopyMemory ByVal VarPtrArray(outStream), VarPtr(dSA), 4& ' setup overlay
    
    lPrevColor = inStream(0, 0) Xor 1&                  ' now walk thru the image 1 more time
    For I = 0& To sizeI.nHeight - 1&                    ' and set the pixel indexes
        If (DataOrientation And orient_BottomUp) Then
            cIndex = (SS.Height - I - 1&) * cOffset
        Else
            cIndex = I * cOffset + HeaderSize
        End If
        If (DataOrientation And orient_PNMformat) Then
            For J = 0& To sizeI.nWidth - 1&
                If lPrevColor <> inStream(J, I) Then
                    lPrevColor = inStream(J, I)
                    lLeafIndex = GetNearestPaletteIndex(hPal, (lPrevColor And &HFFFFFF)) + 1&
                End If                                      ' update the packed index
                outStream(cIndex) = (pPalette.Entries(lLeafIndex) And &HFF&)
                outStream(cIndex + 1) = (pPalette.Entries(lLeafIndex) And &HFF00&) \ &H100&
                outStream(cIndex + 2) = (pPalette.Entries(lLeafIndex) And &HFF0000) \ &H10000
                cIndex = cIndex + 3&
            Next
        Else
            mPtr = (SS.Height - I - 1&) * mScanWidth: mShift = &H80
            If nShift = &H2 Then lShift = mShift Else lShift = nShift
            For J = 0& To sizeI.nWidth - 1&
                If lPrevColor <> inStream(J, I) Then
                    lPrevColor = inStream(J, I)
                    If (pMask(mPtr) And mShift) Then        ' transparent pixel
                        If bWantMask Then lLeafIndex = lIcoTransIndex Else lLeafIndex = pPalette.Count - 1&
                    Else                                    ' else for opaque, we look for it
                        lLeafIndex = GetNearestPaletteIndex(hPal, (lPrevColor And &HFFFFFF))
                    End If
                End If                                      ' update the packed index
                outStream(cIndex) = outStream(cIndex) Or lLeafIndex * lShift
                If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
                If lShift = 1 Then                          ' move destination pointer along & update shift value
                    If nShift = &H2 Then lShift = &H80 Else lShift = nShift
                    cIndex = cIndex + 1&
                Else
                    lShift = lShift \ nShift
                End If
            Next
        End If
    Next
    
    GdipBitmapUnlockBits Handle, tBMP                   ' done; clean up
    CopyMemory ByVal VarPtrArray(inStream), 0&, 4&: tSA.pvData = 0&
    If dBMP.Scan0Ptr Then
        GdipBitmapUnlockBits pvCreateAdaptivePalette, dBMP
        dBMP.Scan0Ptr = 0&
    End If
    CopyMemory ByVal VarPtrArray(outStream), 0&, 4&: dSA.pvData = 0&
    
    If (DataOrientation And (&H3 Or orient_PNMformat)) = 0& Then  ' write the palette to return image
        GdipSetImagePalette pvCreateAdaptivePalette, pPalette
    ElseIf (DataOrientation And orient_WantPaletteInArray) Then  ' write the palette to return array
        If bWantMask Then
            I = 2 ^ ((pvCreateAdaptivePalette And &HFF00&) \ &H100) * 4&
            CopyMemory outData(HeaderSize + cOffset * SS.Height + I), pMask(0), mScanWidth * SS.Height
        End If
        I = pPalette.Count * 4&
        CopyMemory outData(HeaderSize), pPalette.Entries(1), I
        For I = HeaderSize + 3& To HeaderSize + I - 1& Step 4&
            outData(I) = 0
        Next
    End If
    
ExitRoutine:
    If tSA.pvData Then
        GdipBitmapUnlockBits Handle, tBMP
        CopyMemory ByVal VarPtrArray(inStream), 0&, 4&
    End If
    If dBMP.Scan0Ptr Then
        GdipBitmapUnlockBits pvCreateAdaptivePalette, dBMP
        If Err Then
            GdipDisposeImage pvCreateAdaptivePalette
            pvCreateAdaptivePalette = 0&
        End If
    End If
    If dSA.pvData Then CopyMemory ByVal VarPtrArray(outStream), 0&, 4&
    If hPal Then DeleteObject hPal
End Function

Private Function pvCreateNode(nLevel As Long, pLeafCount As Long) As Long

    ' helper function for oct-tree quanitzation
    On Error GoTo ExitRoutine ' begin error trapping

    Dim Index As Long
    pTree.Nodes = pTree.Nodes + 1&
    If pTree.RecycleIndex Then
        Index = pTree.RecycleIndex
        With pTree.Node(Index)
            pTree.RecycleIndex = .NextNode
            .Blues = 0: .Reds = 0: .Greens = 0
            .isLeaf = False: .Pixels = 0
            Erase .Child()
            .NextNode = 0&
        End With
    Else
        Index = UBound(pTree.Node)
        If pTree.Nodes > Index Then
            ReDim Preserve pTree.Node(1 To Index + 50)
        End If
        Index = pTree.Nodes
    End If
    If (nLevel = pTree.MaxDepth) Then
        pTree.Node(Index).isLeaf = True
        pLeafCount = pLeafCount + 1&
    Else
        pTree.Node(Index).NextNode = pTree.ReducibleNodes(nLevel)
        pTree.ReducibleNodes(nLevel) = Index
    End If
    pvCreateNode = Index

ExitRoutine:
End Function

Private Function pvTrimTree(pLeafCount As Long, ColorCount As Long) As Boolean

    ' helper function for oct-tree quanitzation
    Dim I As Long, pNode As Long, FinalCut As Boolean
    Do While ColorCount < pLeafCount

        For I = pTree.MaxDepth To 0 Step -1
            If pTree.ReducibleNodes(I) Then Exit For
        Next
        If I < 0& Then              ' if this happens, we are done. Down to just 1 leaf
            FinalCut = True         ' which means we'll reduce to just 1 color
            pNode = 1&: I = 0&      ' abort with the minimal required colors
        Else
            pNode = pTree.ReducibleNodes(I)
            pTree.Node(pNode).isLeaf = True
        End If
        pTree.ReducibleNodes(I) = pTree.Node(pNode).NextNode
        
        With pTree.Node(pNode)
            For I = 0 To 7
                If .Child(I) Then
                    .Reds = .Reds + pTree.Node(.Child(I)).Reds
                    .Greens = .Greens + pTree.Node(.Child(I)).Greens
                    .Blues = .Blues + pTree.Node(.Child(I)).Blues
                    .Pixels = .Pixels + pTree.Node(.Child(I)).Pixels

                    If pTree.RecycleIndex Then pTree.Node(.Child(I)).NextNode = pTree.RecycleIndex
                    pTree.RecycleIndex = .Child(I)
                    .Child(I) = 0
                    pTree.Nodes = pTree.Nodes - 1&
                    pLeafCount = pLeafCount - 1&
                    If FinalCut Then
                        If pLeafCount = ColorCount Then Exit Do
                    End If
                End If
            Next
        End With
        pLeafCount = pLeafCount + 1
    Loop
    
    pvTrimTree = Not FinalCut

End Function

Private Sub pvGetTreeColors(Index As Long, pIndex As Long)

    ' helper function for oct-tree quanitzation
    ' recursive routine to get the palette colors from the leafs
    Dim c As Long
    If pTree.Node(Index).isLeaf Then
        With pTree.Node(Index)
            pIndex = pIndex + 1&
            .Reds = .Reds \ .Pixels: If .Reds > 255& Then .Reds = 255&
            .Greens = .Greens \ .Pixels: If .Greens > 255& Then .Greens = 255&
            .Blues = .Blues \ .Pixels: If .Blues > 255& Then .Blues = 255&
            pPalette.Entries(pIndex) = .Reds Or .Greens * &H100& Or .Blues * &H10000 Or &HFF000000
        End With
    Else
        For c = 0& To 7&
            If pTree.Node(Index).Child(c) Then
                pvGetTreeColors pTree.Node(Index).Child(c), pIndex
            End If
        Next
    End If

End Sub

Private Function pvColorlosslessReduce(Handle As Long, outData() As Byte, _
                                            DataOrientation As RawDataOrientation, _
                                            HeaderSize As Long, FooterSize As Long, _
                                            AlphaLimit As AlphaReductionLimit, _
                                            SS As SAVESTRUCT, StaticDepth As Long) As Long
                                            
    ' Routine will attempt to reduce bit depth to less than CurrentBitDepth
    ' If successful, return value is non-zero and is dependent on DataOrientation parameter
    ' If DataOrientation is non-zero, then requestor wants pixel information in array vs GDI+ handle
    '   -- called by other routines when saving image to array, file, stdPicture
    '   -- array data is a bottom-up/top-down based on DataOrientation
    '   -- if AlphaLimit is alpha_Mask then mask is appended to the data
    ' If DataOrientation is zero, the requestor wants a GDI+ image
    
    Dim bOK As Boolean, bNewEntry As Boolean
    Dim tSize As RECTI, sizeF As RECTF
    Dim tgtColorDepth As Long, CurrentBitDepth As Long
    Dim X As Long, Y As Long, prevColor As Long
    Dim lShift As Byte, rShift As Long, pIndex As Long
    ' source image info
    Dim tBMP As BitmapData, tSA As SafeArray, srcData() As Long
    ' reduced image info
    Dim dBMP As BitmapData, dSA As SafeArray, dstData() As Byte
    Dim dPtr As Long, dstScanWidth As Long
    Dim mskScanWidth As Long, mskOffset As Long
    Dim mShift As Byte, mPtr As Long
    Dim bWantMask As Boolean, icoTransIndex As Long
    
    If StaticDepth Then     ' called from pvCreateAdaptivePalette
        Select Case (StaticDepth And &HFF00&) \ &H100&
            Case 8&: tgtColorDepth = 256&: CurrentBitDepth = lvicColor24bpp
            Case 4&: tgtColorDepth = 16&: CurrentBitDepth = lvicColor8bpp
            Case 1&: tgtColorDepth = 2&: CurrentBitDepth = lvicColor4bpp
            Case Else: Exit Function
        End Select
    Else
        CurrentBitDepth = modCommon.ColorDepthToColorType(lvicNoColorReduction, Handle)
        Select Case (CurrentBitDepth And &HFF00&) \ &H100&
            Case Is > 8&: tgtColorDepth = 256&  ' if can't palettize to at least 8bpp, may abort if can't get to 24bpp
            Case 8&: tgtColorDepth = 16&        ' if can't palettize to at least 4bpp, abort
            Case 4&: tgtColorDepth = 2&         ' if can't palettize to at least 1bpp, abort
            Case Else: Exit Function            ' abort, can't get lower than 1bpp
        End Select
    End If
    
    tSize.nHeight = SS.Height: tSize.nWidth = SS.Width
    If GdipBitmapLockBits(Handle, tSize, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
    With tSA
        .cbElements = 4
        .cDims = 1
        .pvData = tBMP.Scan0Ptr
        If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
        .rgSABound(0).cElements = tSize.nWidth * tSize.nHeight
    End With
    CopyMemory ByVal VarPtrArray(srcData), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine
    
    prevColor = srcData(0) Xor 1&
    For X = 0& To tSA.rgSABound(0).cElements - 1&
        If prevColor <> srcData(X) Then
            prevColor = srcData(X)
            pIndex = modCommon.FindColor(pPalette.Entries(), prevColor, pPalette.Count, bNewEntry)
            If (prevColor And &HFF000000) <> &HFF000000 Then
                pPalette.Flags = 1&
                If AlphaLimit = alpha_None Then             ' no alpha allowed
                    pPalette.Count = tgtColorDepth + 1&
                    Exit For
                ElseIf AlphaLimit = alpha_Simple Then  ' only 100% alpha allowed (GIF,ICO,CUR)
                    If (prevColor And &HFF000000) <> 0& Then
                        pPalette.Count = tgtColorDepth + 1&
                        Exit For
                    End If
                    bNewEntry = False
                End If                                      ' else complex alpha allowed (TGA, PNG)
            End If
            If bNewEntry Then
                pPalette.Count = pPalette.Count + 1&
                If pPalette.Count > tgtColorDepth Then
                    Exit For
                End If
                If pIndex < pPalette.Count Then
                    CopyMemory pPalette.Entries(pIndex + 1&), pPalette.Entries(pIndex), (pPalette.Count - pIndex) * 4&
                End If
                pPalette.Entries(pIndex) = prevColor
            End If
        End If
    Next
    
    bWantMask = (DataOrientation And orient_WantMask)
    If bWantMask = True And pPalette.Flags = 1& Then
        icoTransIndex = pvFixupIconPalette(tgtColorDepth) - 1&
        If icoTransIndex = -1& Then pPalette.Count = 257&
    End If
    ' if simple transparency, then we need a pPalette index for the transparent color for GIFs
    If AlphaLimit = alpha_Simple And pPalette.Flags = 1& And bWantMask = False Then tgtColorDepth = tgtColorDepth - 1&
    If pPalette.Count > tgtColorDepth Then                   ' reduce to 24bpp or 32bpp with Alpha Tolerance?
        If ((CurrentBitDepth And &HFF00&) \ &H100& = 32&) Then     ' only if source is 32bpp and only if ...
            If pPalette.Flags = 0& Or (DataOrientation And orient_WantMask) Then ' no alpha exists or reducing icon/cursor
                bOK = True
                For X = X + 1& To tSA.rgSABound(0).cElements - 1&
                    If (srcData(X) And &HFF000000) <> &HFF000000 Then   ' has alpha
                        If AlphaLimit = alpha_None Then                 ' looking for 24bpp no mask; abort
                            bOK = False
                            Exit For
                        ElseIf (srcData(X) And &HFF000000) <> 0& Then   ' looking for simple not complex transparency
                            bOK = False
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        pPalette.Count = 0&
    Else
        bOK = True
    End If
    If Not bOK Then
        CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
        GdipBitmapUnlockBits Handle, tBMP
        Exit Function
    End If
        
    bNewEntry = False
    prevColor = srcData(0) Xor 1&
        
    ' if we got here, bit reduction is in play
    If (DataOrientation And orient_PNMformat) Then
        If pPalette.Count > 0& And pPalette.Flags = 0& Then ' can palettize to lower bit depth
            ' PNM formats cannot use paletted format unless grayscale or b&w
            If Me.IsGrayscale(True) = False Then pPalette.Count = 0&
        Else
            pPalette.Count = 0&
        End If
    End If
    If pPalette.Count Then
        If (DataOrientation And orient_8bppIndexes) Then
            tgtColorDepth = 256&: rShift = &H1: dBMP.PixelFormat = lvicColor8bpp
            If (DataOrientation And orient_PNMformat) Then
                If Me.IsBlackWhite Then
                    tgtColorDepth = 2&: rShift = &H2: dBMP.PixelFormat = lvicColor1bpp
                End If
            End If
        Else
            If StaticDepth = 0& Then
                If ((pPalette.Flags And 1&) And AlphaLimit = alpha_Simple) Then X = pPalette.Count + 1& Else X = pPalette.Count
            Else
                X = tgtColorDepth
            End If
            Select Case X
            Case Is < 3&
                dBMP.PixelFormat = lvicColor1bpp: tgtColorDepth = 2&: rShift = &H2
                If pvFixupBWpalette() = False Then
                    If (DataOrientation And orient_4bppIndexesMin) Then
                        dBMP.PixelFormat = lvicColor4bpp: tgtColorDepth = 16&: rShift = &H10
                    End If
                End If
            Case Is < 17&
                dBMP.PixelFormat = lvicColor4bpp: tgtColorDepth = 16&: rShift = &H10
            Case Else
                dBMP.PixelFormat = lvicColor8bpp: tgtColorDepth = 256&: rShift = &H1
            End Select
        End If
        ' note: unlike other routines herein, don't need to sort grayscale if requested
        ' the sorted palette already has them in ascending order if palette is grayscale
        
        dstScanWidth = modCommon.ByteAlignOnWord((dBMP.PixelFormat And &HFF00&) \ &H100, tSize.nWidth)
        If (DataOrientation And &H3) = 0& Then
            If GdipCreateBitmapFromScan0(tSize.nWidth, tSize.nHeight, 0&, dBMP.PixelFormat, ByVal 0&, pvColorlosslessReduce) Then GoTo ExitRoutine
            If GdipBitmapLockBits(pvColorlosslessReduce, tSize, ImageLockModeWrite, dBMP.PixelFormat, dBMP) Then
                GdipDisposeImage pvColorlosslessReduce
                pvColorlosslessReduce = 0&
                GoTo ExitRoutine
            End If
            dSA.rgSABound(0).cElements = tSize.nHeight * dstScanWidth
            dSA.pvData = dBMP.Scan0Ptr
        Else
            If (DataOrientation And orient_WantPaletteInArray) And tgtColorDepth <> lvicColor24bpp Then
                X = tgtColorDepth * 4&
            Else
                X = 0&
            End If
            If (DataOrientation And orient_WantMask) Then
                mskScanWidth = modCommon.ByteAlignOnWord(1, tSize.nWidth)
                dSA.rgSABound(0).cElements = tSize.nHeight * (dstScanWidth + mskScanWidth)
                mskOffset = dSA.rgSABound(0).cElements - mskScanWidth
            Else
                dSA.rgSABound(0).cElements = tSize.nHeight * dstScanWidth
            End If
            ReDim outData(0 To HeaderSize + FooterSize + dSA.rgSABound(0).cElements + X - 1&)
            dSA.pvData = VarPtr(outData(HeaderSize + X))
        End If
        dSA.cbElements = 1
        dSA.cDims = 1
        CopyMemory ByVal VarPtrArray(dstData), VarPtr(dSA), 4&
            
        For Y = 0& To tSize.nHeight - 1&
            If mskScanWidth Then mShift = &H80: mPtr = mskOffset - Y * mskScanWidth
            If (DataOrientation And orient_BottomUp) Then
                dPtr = dstScanWidth * (tSize.nHeight - Y - 1&)
            Else
                dPtr = Y * dstScanWidth
            End If
            X = Y * tSize.nWidth
            If rShift = &H2 Then lShift = &H80 Else lShift = rShift
            For X = X To X + tSize.nWidth - 1&
                If prevColor <> srcData(X) Then
                    prevColor = srcData(X)
                    bNewEntry = ((prevColor And &HFF000000) <> &HFF000000)
                    If bNewEntry Then
                        If AlphaLimit = alpha_Complex Then  ' TGA,PNG
                            pIndex = modCommon.FindColor(pPalette.Entries(), prevColor, pPalette.Count, False) - 1&
                        Else    ' all others (ico/cur doesn't matter which index, GIF will use last index)
                            If bWantMask Then pIndex = icoTransIndex Else pIndex = pPalette.Count
                        End If
                    Else
                        pIndex = modCommon.FindColor(pPalette.Entries(), prevColor, pPalette.Count, False) - 1&
                    End If
                End If
                If mskOffset Then
                    If bNewEntry Then dstData(mPtr) = dstData(mPtr) Or mShift
                    If mShift = 1 Then mPtr = mPtr + 1&: mShift = &H80 Else mShift = mShift \ 2
                End If
                dstData(dPtr) = dstData(dPtr) Or pIndex * lShift
                If lShift = 1 Then
                    If rShift = &H2 Then lShift = &H80 Else lShift = rShift
                    dPtr = dPtr + 1&
                Else
                    lShift = lShift \ rShift
                End If
            Next
        Next
        If (DataOrientation And orient_BlackIs1Not0) Then
            For X = 1& To pPalette.Count
                If pPalette.Entries(X) = &HFF000000 Then pPalette.Entries(X) = &HFF000001
            Next
        End If
        If pPalette.Flags = 1 And AlphaLimit = alpha_Simple Then
            If bWantMask = False Then
                pPalette.Count = pPalette.Count + 1&
                pPalette.Entries(pPalette.Count) = (SS.RSS.FillColorARGB And &HFFFFFF)
            End If
        End If
        If pvColorlosslessReduce = 0& Then
            pvColorlosslessReduce = dBMP.PixelFormat
            If (DataOrientation And orient_WantPaletteInArray) Then
                X = pPalette.Count * 4&
                CopyMemory outData(HeaderSize), pPalette.Entries(1), X
                For X = HeaderSize + 3& To HeaderSize + X - 1& Step 4&
                    outData(X) = 0
                Next
            End If
        End If
    
    Else            ' reduction to 24bpp
    
        CopyMemory ByVal VarPtrArray(srcData), 0&, 4&: tSA.pvData = 0&  ' release overlay and unlock source
        GdipBitmapUnlockBits Handle, tBMP: tBMP.Scan0Ptr = 0&           ' get new pointer for 24bpp vs 32bpp data
        If (DataOrientation And &H3) = 0& Then
            ' create new 24bpp GDI+ image
            If GdipCreateBitmapFromScan0(tSize.nWidth, tSize.nHeight, 0&, lvicColor24bpp, ByVal 0&, pvColorlosslessReduce) Then GoTo ExitRoutine
            GdipBitmapLockBits Handle, tSize, ImageLockModeRead, lvicColor24bpp, tBMP
            dBMP = tBMP                                                     ' make destination & source pointers same
            GdipBitmapLockBits pvColorlosslessReduce, tSize, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor24bpp, dBMP
        Else
            dstScanWidth = modCommon.ByteAlignOnWord(24, tSize.nWidth)
            If bWantMask Then
                mskScanWidth = modCommon.ByteAlignOnWord(1, tSize.nWidth)
                ReDim outData(0 To HeaderSize + FooterSize + (dstScanWidth + mskScanWidth) * tSize.nHeight - 1&)
                mskOffset = UBound(outData) + 1 - mskScanWidth - FooterSize
            Else
                ReDim outData(0 To HeaderSize + FooterSize + dstScanWidth * tSize.nHeight - 1&)
            End If
            If (DataOrientation And orient_BottomUp) Then
                tBMP.Scan0Ptr = VarPtr(outData(dstScanWidth * tSize.nHeight - dstScanWidth + HeaderSize))
                tBMP.stride = -dstScanWidth
            Else
                tBMP.Scan0Ptr = VarPtr(outData(HeaderSize))
                tBMP.stride = dstScanWidth
            End If
            GdipBitmapLockBits Handle, tSize, ImageLockModeRead Or ImageLockModeUserInputBuf, lvicColor24bpp, tBMP
            GdipBitmapUnlockBits Handle, tBMP: tBMP.Scan0Ptr = 0&           ' get new pointer for 24bpp vs 32bpp data
            pvColorlosslessReduce = lvicColor24bpp
        End If
        If bWantMask = True Then
            GdipBitmapLockBits Handle, tSize, ImageLockModeRead, lvicColor32bppAlpha, tBMP
            tSA.pvData = tBMP.Scan0Ptr
            If tBMP.stride < 0& Then tSA.pvData = tSA.pvData + (tBMP.Height - 1&) * tBMP.stride
            CopyMemory ByVal VarPtrArray(srcData), VarPtr(tSA), 4&
            X = 0&
            For Y = 0& To tSize.nHeight - 1&
                mPtr = mskOffset - Y * mskScanWidth: mShift = &H80
                dPtr = HeaderSize + dstScanWidth * (tSize.nHeight - Y - 1&)
                For X = X To X + tSize.nWidth - 1&
                    If prevColor <> srcData(X) Then
                        prevColor = srcData(X)
                        bNewEntry = ((prevColor And &HFF000000) <> &HFF000000)
                    End If
                    If bNewEntry Then
                        outData(mPtr) = outData(mPtr) Or mShift
                        outData(dPtr) = 0: outData(dPtr + 1) = 0: outData(dPtr + 2) = 0
                    End If
                    If mShift = 1 Then mShift = &H80: mPtr = mPtr + 1& Else mShift = mShift \ 2
                    dPtr = dPtr + 3
                Next
            Next
        End If
    End If

ExitRoutine:
    If dSA.pvData Then CopyMemory ByVal VarPtrArray(dstData), 0&, 4&
    If dBMP.Scan0Ptr Then
        GdipBitmapUnlockBits pvColorlosslessReduce, dBMP
        GdipSetImagePalette pvColorlosslessReduce, pPalette
    End If
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
    If tBMP.Scan0Ptr Then GdipBitmapUnlockBits Handle, tBMP
End Function

Public Sub ApplyPaletteToHandle(PaletteCount As Long, ByVal hUserPalette As Long, hBitmap As Long)

    ' called to apply a Windows GDI palette to a GDI+ bitmap
    Dim p As Long, b As Long, g As Long, r As Long
    
    pPalette.Flags = 0
    pPalette.Count = PaletteCount
    If hUserPalette = 0& Then
        Select Case PaletteCount
        Case 2&
            pPalette.Entries(1) = &HFF000000: pPalette.Entries(2) = &HFFFFFFFF
        Case 16&
            p = 9&
            For b = 0& To 255& Step 255&: For g = 0& To 255& Step 255&: For r = 0& To 255& Step 255&
                pPalette.Entries(p) = r Or g * &H100& Or b * &H10000 Or &HFF000000
                p = p + 1&
            Next: Next: Next: p = 1&
            For b = 0& To 128& Step 128&: For g = 0& To 128& Step 128&: For r = 0& To 128& Step 128&
                pPalette.Entries(p) = r Or g * &H100& Or b * &H10000 Or &HFF000000
                p = p + 1&
            Next: Next: Next
        Case Else
            hUserPalette = CreateHalftonePalette(0&)
            If hUserPalette = 0& Then      ' failure? Use websafe instead
                For r = 0 To 255& Step 51&: For g = 0 To 255& Step 51&: For b = 0 To 255& Step 51&
                    p = p + 1&
                    pPalette.Entries(p) = r Or g * &H100& Or b * &H10000 Or &HFF000000
                Next: Next: Next
            End If
        End Select
    End If
    If hUserPalette Then
        p = GetPaletteEntries(hUserPalette, 0&, 0&, ByVal 0&)
        If p > pPalette.Count Then p = pPalette.Count Else pPalette.Count = p
        If p Then
            GetPaletteEntries hUserPalette, 0&, p, ByVal VarPtr(pPalette.Entries(1))
            For p = 1 To p: pPalette.Entries(p) = pPalette.Entries(p) Or &HFF000000: Next
        End If
    End If
    If pPalette.Count Then GdipSetImagePalette hBitmap, pPalette

End Sub

Private Function pvConvertPaletteToGDI(PaletteType As PaletteTypeEnum, hUserPalette As Long, Optional TransIndexNeeded As Boolean = False) As Long

    ' routine creates a GDI palette and also a GDI+ palette structure based on passed parameters
    ' When an index is required for transparent color, it will always be placed at end of palette

    Dim LOGpal As LOGPALETTE, p As Long
    Dim r As Long, g As Long, b As Long
    LOGpal.palVersion = &H300
    
    Select Case PaletteType
        Case lvicPaletteDefault
            Select Case pPalette.Count
            Case 2&                         ' B&W palette
                pPalette.Entries(1) = &HFF000000: pPalette.Entries(2) = &HFFFFFFFF
            Case 16&                        ' 16 color halftones
                If TransIndexNeeded Then p = 8& Else p = 9&
                For b = 0& To 255& Step 255&: For g = 0& To 255& Step 255&: For r = 0& To 255& Step 255&
                    pPalette.Entries(p) = r Or g * &H100& Or b * &H10000 Or &HFF000000
                    p = p + 1&
                Next: Next: Next: p = 1&
                For b = 0& To 128& Step 128&: For g = 0& To 128& Step 128&: For r = 0& To 128& Step 128&
                    pPalette.Entries(p) = r Or g * &H100& Or b * &H10000 Or &HFF000000
                    p = p + 1&
                Next: Next: Next
                ' when a transparency index is required, we only use 15 colors & make one transparent
                If TransIndexNeeded = False Then pPalette.Entries(p) = 192& Or 192& * &H100& Or 192& * &H10000 Or &HFF000000
            Case Else                       ' 216 color halftones
                pvConvertPaletteToGDI = CreateHalftonePalette(0&)
                If pvConvertPaletteToGDI = 0& Then      ' failure? Use websafe instead
                    For r = 0 To 255& Step 51&: For g = 0 To 255& Step 51&: For b = 0 To 255& Step 51&
                        p = p + 1&
                        pPalette.Entries(p) = r Or g * &H100& Or b * &H10000 Or &HFF000000
                    Next: Next: Next
                    ' when a transparency index is required, we'll add a transparent color
                    If TransIndexNeeded Then pPalette.Count = p + 1& Else pPalette.Count = p
                Else                        ' 256 color halftones
                    pPalette.Count = GetPaletteEntries(pvConvertPaletteToGDI, 0&, 0&, ByVal 0&)
                    GetPaletteEntries pvConvertPaletteToGDI, 0&, pPalette.Count, ByVal VarPtr(pPalette.Entries(1))
                    If TransIndexNeeded Then
                        ' when a transparency index is required, we only use 255 colors & make one transparent
                        p = GetNearestPaletteIndex(pvConvertPaletteToGDI, RGB(128, 128, 128)) + 1&
                        If p < pPalette.Count Then
                            CopyMemory pPalette.Entries(p), pPalette.Entries(p + 1&), (pPalette.Count - p) * 4&
                        End If
                        DeleteObject pvConvertPaletteToGDI: pvConvertPaletteToGDI = 0&
                    End If
                    For p = 1 To pPalette.Count: pPalette.Entries(p) = pPalette.Entries(p) Or &HFF000000: Next
                    If pvConvertPaletteToGDI Then Exit Function
                End If
            End Select
        
        Case lvicPaletteUserDefined     ' user-defined palette passed; scrutinize a bit
        
            If hUserPalette = 0& Then Exit Function
            pPalette.Count = GetPaletteEntries(hUserPalette, 0&, 0&, ByVal 0&)
            If pPalette.Count = 0& Then Exit Function
            
            GetPaletteEntries hUserPalette, 0&, pPalette.Count, ByVal VarPtr(pPalette.Entries(1))
            For p = 1& To pPalette.Count: pPalette.Entries(p) = pPalette.Entries(p) Or &HFF000000: Next
            If TransIndexNeeded Then
                ' need a transparent index; use one of the colors if necessary else append one
                If pPalette.Count < 3& Then
                    pPalette.Count = 3&: p = 0&     ' 1bpp will be 4bpp; append transparency
                ElseIf pPalette.Count < 16& Then
                    pPalette.Count = p: p = 0&      ' room exists; append transparency
                ElseIf pPalette.Count < 256& Then
                    pPalette.Count = p: p = 0&      ' room exists; append transparency
                End If
                If p Then                           ' no room exists; use an existing color
                    p = GetNearestPaletteIndex(hUserPalette, RGB(128, 128, 128)) + 1&
                    If p < pPalette.Count Then
                        CopyMemory pPalette.Entries(p), pPalette.Entries(p + 1&), (pPalette.Count - p) * 4&
                    End If
                End If
                pvConvertPaletteToGDI = 0&
            Else
                pvConvertPaletteToGDI = hUserPalette
            End If
        Case Else
            ' adaptive, palette already filled
    End Select
    If pvConvertPaletteToGDI = 0& Then
        If TransIndexNeeded = True Then
            LOGpal.palNumEntries = pPalette.Count - 1&
            pPalette.Entries(pPalette.Count) = 0&
        Else
            LOGpal.palNumEntries = pPalette.Count
        End If
        CopyMemory LOGpal.palPalEntry(1), pPalette.Entries(1), LOGpal.palNumEntries * 4&
        pvConvertPaletteToGDI = CreatePalette(LOGpal)
    End If

End Function

Private Function pvFixupIconPalette(MaxColorCount As Long) As Long

    ' for icons/cursors to draw properly, the transparency index must point to the color black.
    ' routine looks for opauqe black: &HFF000000. If not found, it adds it if it can be found
    Dim p As Long, pIndex As Long
    
    If pPalette.Count > MaxColorCount Then Exit Function
    For pIndex = 1& To pPalette.Count
        If (pPalette.Entries(pIndex) And &HFFFFFF) = 0& Then
            pPalette.Entries(pIndex) = &HFF000000
            Exit For ' looking for opaque black
        End If
    Next
    If pIndex > pPalette.Count Then             ' not found in the palette
        If pPalette.Count < MaxColorCount Then  ' but have room for entry
            pPalette.Count = pIndex
            pPalette.Entries(pIndex) = &HFF000000
            pvFixupIconPalette = pIndex
        End If
    Else                                        ' exists
        pvFixupIconPalette = pIndex
    End If
    
End Function

Private Function pvFixupBWpalette() As Boolean

    ' we ensure B&W pallete contains Black as first entry & White as second entry
    ' If not B&W then function returns false
    If pPalette.Count = 2 Then
        If pPalette.Entries(1) = &HFFFFFFFF Then
            If pPalette.Entries(2) = &HFF000000 Then
                pPalette.Entries(2) = pPalette.Entries(1)
                pPalette.Entries(1) = &HFF000000
                pvFixupBWpalette = True
            End If
        ElseIf pPalette.Entries(1) = &HFF000000 Then
            pvFixupBWpalette = (pPalette.Entries(2) = &HFFFFFFFF)
        End If
    End If

End Function

Private Function pvCreateStandardPalette(Handle As Long, outData() As Byte, _
                                            DataOrientation As RawDataOrientation, _
                                            HeaderSize As Long, FooterSize As Long, _
                                            AlphaLimit As AlphaReductionLimit, _
                                            SS As SAVESTRUCT) As Long

    ' routine converts passed image to a 2, 16, 256 color paletted image

    Dim X As Long, Y As Long, hPal As Long, lMaskIndex As Long
    Dim pIndex As Long, lPrevColor As Long, lResult As Long
    Dim tgtColorDepth As Long, dstScanWidth As Long
    Dim lShift As Long, rShift As Long, bAlpha As Boolean
    Dim mskOffset As Long, mShift As Long, mScanWidth As Long, mPtr As Long
    Dim dPtr As Long, srcSA As SafeArray, srcData() As Long
    Dim tBMP As BitmapData, tSizeI As RECTI, bWantMask As Boolean
    Dim dBMP As BitmapData, dSA As SafeArray, outStream() As Byte
    
    If SS.ColorDepth = lvicConvert_BlackWhite Then  ' handle B&W differently
        pvCreateStandardPalette = pvCreateBlackAndWhite(Handle, outData(), DataOrientation, HeaderSize, FooterSize, AlphaLimit, SS)
        Exit Function
    End If
    
    bWantMask = (DataOrientation And orient_WantMask)   ' icon/cursor?
    If SS.PaletteType = lvicPaletteUserDefined Then     ' user passed palette
        ' validate palette & revert to standard palette if needed
        hPal = pvConvertPaletteToGDI(lvicPaletteUserDefined, SS.Palette_Handle, ((SS.reserved1 And 1&) And bWantMask = False))
        If hPal = 0& Then   ' either Palette_Handle is zero or passed invalid handle
            pPalette.Count = 256&: tgtColorDepth = 8&: rShift = &H1: dBMP.PixelFormat = lvicColor8bpp
            SS.PaletteType = lvicPaletteDefault
            hPal = pvConvertPaletteToGDI(lvicPaletteDefault, 0&, ((SS.reserved1 And 1&) And bWantMask = False))
        Else                ' determine bitdepth from color count
            Select Case pPalette.Count
            Case Is < 3&: tgtColorDepth = 1&: rShift = &H2: dBMP.PixelFormat = lvicColor1bpp
            Case Is < 17&: tgtColorDepth = 4&: rShift = &H10: dBMP.PixelFormat = lvicColor4bpp
            Case Else: tgtColorDepth = 8&: rShift = &H1: dBMP.PixelFormat = lvicColor8bpp
            End Select
        End If
    Else                                                ' using default halftone palette
        If SS.ColorDepth = lvicConvert_16Colors Then
            pPalette.Count = 16&: tgtColorDepth = 4&: rShift = &H10: dBMP.PixelFormat = lvicColor4bpp
        Else
            pPalette.Count = 256&: tgtColorDepth = 8&: rShift = &H1: dBMP.PixelFormat = lvicColor8bpp
        End If
        hPal = pvConvertPaletteToGDI(lvicPaletteDefault, 0&, ((SS.reserved1 And 1&) And bWantMask = False))
    End If
    
    If hPal Then                                        ' got a palette; do some more tweaks
        If (DataOrientation And orient_BlackIs1Not0) Then
            ' WMFs cannot be allowed to contain a palette that only consists of black & white, regardless
            ' how many colors are in the palette. Transparency is produced and is not compatible with stdPictures
            For X = 1& To pPalette.Count    ' change black to near black
                If pPalette.Entries(X) = &HFF000000 Then pPalette.Entries(X) = &HFF000001
            Next
        End If
        If (SS.reserved1 And 1&) And bWantMask = True Then  ' for icons, ensure we have an index for black
            lMaskIndex = pvFixupIconPalette(2 ^ tgtColorDepth) - 1&
            If lMaskIndex = -1& Then                        ' tweak; no black exists
                If hPal <> SS.Palette_Handle Then DeleteObject hPal
                hPal = pvConvertPaletteToGDI(SS.PaletteType, SS.Palette_Handle, True)
                pPalette.Entries(pPalette.Count) = &HFF000000
                lMaskIndex = pvFixupIconPalette(2 ^ tgtColorDepth) - 1&
                If lMaskIndex = 0& Then
                    If hPal <> SS.Palette_Handle Then DeleteObject hPal
                    hPal = 0&
                End If
            End If
        End If
    End If
    If hPal = 0& Then Exit Function
    
    
    If (DataOrientation And orient_8bppIndexes) Then ' override above if formatting for GIF
        rShift = &H1: dBMP.PixelFormat = lvicColor8bpp: tgtColorDepth = 8&
    ElseIf (DataOrientation And orient_4bppIndexesMin) Then ' override above if this is requested
        If pPalette.Count < 3& Then rShift = &H10: dBMP.PixelFormat = lvicColor4bpp: tgtColorDepth = 4&
    End If
    tBMP.Width = SS.Width: tBMP.Height = SS.Height
    tSizeI.nWidth = tBMP.Width: tSizeI.nHeight = tBMP.Height
    If GdipBitmapLockBits(Handle, tSizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then
        If SS.Palette_Handle <> hPal Then DeleteObject hPal
        Exit Function
    End If
    With srcSA
        .cbElements = 4
        .cDims = 2
        .pvData = tBMP.Scan0Ptr
        If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
        .rgSABound(0).cElements = tBMP.Height
        .rgSABound(1).cElements = tBMP.Width
    End With
    CopyMemory ByVal VarPtrArray(srcData), VarPtr(srcSA), 4&
    On Error GoTo ExitRoutine
    
    ' set up return object/array
    If (DataOrientation And orient_PNMformat) Then
        If Me.IsGrayscale() = False Then
            tgtColorDepth = 24&
            dBMP.PixelFormat = lvicColor24bpp
        End If
    End If
    
    dstScanWidth = modCommon.ByteAlignOnWord(tgtColorDepth, tSizeI.nWidth)
    If (DataOrientation And &H3) = 0& Then
        If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, dBMP.PixelFormat, ByVal 0&, pvCreateStandardPalette) Then GoTo ExitRoutine
        If GdipBitmapLockBits(pvCreateStandardPalette, tSizeI, ImageLockModeWrite, dBMP.PixelFormat, dBMP) Then
            GdipDisposeImage pvCreateStandardPalette
            GoTo ExitRoutine
        End If
        dSA.pvData = dBMP.Scan0Ptr
        dSA.rgSABound(0).cElements = dstScanWidth * SS.Height
    Else ' array
        If (DataOrientation And orient_WantPaletteInArray) Then X = (2 ^ tgtColorDepth) * 4& Else X = 0&
        If bWantMask Then
            mScanWidth = modCommon.ByteAlignOnWord(1, SS.Width)
            dSA.rgSABound(0).cElements = (dstScanWidth + mScanWidth) * SS.Height
            mskOffset = dSA.rgSABound(0).cElements - mScanWidth
        Else
            dSA.rgSABound(0).cElements = dstScanWidth * SS.Height
        End If
        ReDim outData(0 To dSA.rgSABound(0).cElements + HeaderSize + FooterSize + X - 1&)
        dSA.pvData = VarPtr(outData(HeaderSize + X))
        pvCreateStandardPalette = dBMP.PixelFormat
    End If
    dSA.cbElements = 1
    dSA.cDims = 1
    CopyMemory ByVal VarPtrArray(outStream), VarPtr(dSA), 4&
            
    If (SS.reserved1 And 1&) Then
        If bWantMask = False Then
            pPalette.Flags = 1&         ' indicates transparency exists in palette
            pPalette.Entries(pPalette.Count) = (SS.RSS.FillColorARGB And &HFFFFFF)
        End If
    ElseIf (DataOrientation And orient_SortGrayscale) Then
        If pPalette.Count = 256& And hPal = SS.Palette_Handle Then
            If Me.IsGrayscale() Then    ' ensure grayscale is ascending order
                For X = 0& To 255&
                    pPalette.Entries(X + 1&) = X Or X * &H100& Or X * &H10000 Or &HFF000000
                Next
                If SS.Palette_Handle <> hPal Then DeleteObject hPal
                hPal = pvConvertPaletteToGDI(lvicPaletteAdaptive, 0&, False)
            End If
        End If
    End If
    
    lPrevColor = srcData(0, 0) Xor 1            ' process the pixel data
    With SS
        For Y = 0& To tSizeI.nHeight - 1
            If mskOffset Then mPtr = mskOffset - Y * mScanWidth: mShift = &H80
            If (DataOrientation And orient_BottomUp) Then
                dPtr = dstScanWidth * (SS.Height - Y - 1&)
            Else
                dPtr = Y * dstScanWidth
            End If
            If rShift = &H2 Then lShift = &H80 Else lShift = rShift
            For X = 0& To tSizeI.nWidth - 1
                If lPrevColor <> srcData(X, Y) Then
                    lPrevColor = srcData(X, Y)
                    bAlpha = ((lPrevColor And &HFF000000) <> &HFF000000)
                    If bAlpha Then
                        lResult = (lPrevColor And &H7F000000) \ &H1000000
                        If lPrevColor < 0& Then lResult = lResult Or &H80
                        If lResult > .AlphaTolerancePct Then
                            pIndex = GetNearestPaletteIndex(hPal, (lPrevColor And &HFFFFFF))
                        Else
                            If bWantMask Then pIndex = lMaskIndex Else pIndex = pPalette.Count - 1&
                        End If
                    Else
                        pIndex = GetNearestPaletteIndex(hPal, (lPrevColor And &HFFFFFF))
                    End If
                End If
                If tgtColorDepth = 24& Then
                    CopyMemory outStream(dPtr), pPalette.Entries(pIndex + 1&), 3&
                    dPtr = dPtr + 3&
                Else
                    outStream(dPtr) = outStream(dPtr) Or pIndex * lShift
                    If lShift = &H1 Then
                        If rShift = &H2 Then lShift = &H80 Else lShift = rShift
                        dPtr = dPtr + 1&
                    Else
                        lShift = lShift \ rShift
                    End If
                End If
                If bWantMask Then
                    If bAlpha Then outStream(mPtr) = outStream(mPtr) Or mShift
                    If mShift = 1 Then mPtr = mPtr + 1&: mShift = &H80 Else mShift = mShift \ 2
                End If
            Next
        Next
    End With
    
ExitRoutine:                                ' clean up & apply palette
    CopyMemory ByVal VarPtrArray(srcData), 0&, 4&
    GdipBitmapUnlockBits Handle, tBMP
    If dBMP.Scan0Ptr Then
        GdipBitmapUnlockBits pvCreateStandardPalette, dBMP
        If Err Then
            GdipDisposeImage pvCreateStandardPalette
            pvCreateStandardPalette = 0&
        Else
            GdipSetImagePalette pvCreateStandardPalette, pPalette
        End If
    ElseIf Err Then
        pvCreateStandardPalette = 0&
    ElseIf (DataOrientation And orient_WantPaletteInArray) Then
        Y = pPalette.Count * 4&
        CopyMemory outData(HeaderSize), pPalette.Entries(1), Y
        For X = HeaderSize + 3& To HeaderSize + Y - 1& Step 4&
            outData(X) = 0
        Next
    End If
    If dSA.pvData Then CopyMemory ByVal VarPtrArray(outStream), 0&, 4&
    If SS.Palette_Handle <> hPal Then DeleteObject hPal
    
End Function

Private Function pvCreateBlackAndWhite(Handle As Long, outData() As Byte, _
                                            DataOrientation As RawDataOrientation, _
                                            ByVal HeaderSize As Long, FooterSize As Long, _
                                            AlphaLimit As AlphaReductionLimit, _
                                            SS As SAVESTRUCT) As Long

    ' function creates a b&w image in either 1bpp or 4bpp depending on transparency
    ' 4bpp? No such thing as 1bpp b&w image that also contains transparency, except for icon/cur
    ' Note: Before coming to this routine, we are playing with a copy & not the original image
     
    ' This is my own custom method and may not be to your liking. If not, you can always supply your own
    ' The logic follows
    '   - Render image using lvicCCIR709 (done before this routine called)
    '   - Run thru image once, adding all colors to a palette (max of 256 colors)
    '   - Run thru image again, locate pixel in palette & if > 1/2 way point then white else black
    ' This logic prevents images having vast majority intensity from becoming all white or all black.
    ' For example, create a 2 color Red & Blue image. Run it thru most converters & image will be all black
    
    
    Dim tSizeF As RECTF, sizeI As RECTI
    Dim I As Long, J As Long, bAlpha As Boolean
    Dim lResult As Long, lColorType As Long
    Dim tBMP As BitmapData, dBMP As BitmapData
    Dim scanWidth As Long, mScanWidth As Long
    Dim tSA As SafeArray, inStream() As Long
    Dim dSA As SafeArray, outStream() As Byte
    Dim sSA As SafeArray, bwStream() As Byte, srcPtr As Long
    Dim mShift As Byte, mPtr As Long, mskOffset As Long
    Dim lShift As Byte, rShift As Byte, dPtr As Long
    Dim lPrevColor As Long, lMedian As Long
    Dim bWantMask As Boolean, bWantNewPic As Boolean
    
    If GdipGetImageGraphicsContext(Handle, lResult) Then Exit Function
    
    ' ensure we can now read what we drew
    sizeI.nWidth = SS.Width: sizeI.nHeight = SS.Height
    If GdipBitmapLockBits(Handle, sizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
    With tSA
        .cbElements = 4
        .cDims = 2
        .pvData = tBMP.Scan0Ptr
        If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
        .rgSABound(0).cElements = SS.Height
        .rgSABound(1).cElements = SS.Width
    End With
    CopyMemory ByVal VarPtrArray(inStream), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine                               ' begin error trapping
    
    ' calculate destination bitdepth and shifting values
    bWantMask = (DataOrientation And orient_WantMask)
    If (SS.reserved1 And 1&) = 1& And AlphaLimit <> alpha_None Then pPalette.Flags = 1&
    If (DataOrientation And orient_8bppIndexes) Then
        If (DataOrientation And orient_PNMformat) Then
            lColorType = lvicColor1bpp: rShift = &H2
        Else
            lColorType = lvicColor8bpp: rShift = &H1
        End If
    ElseIf (DataOrientation And orient_4bppIndexesMin) Then
        lColorType = lvicColor4bpp: rShift = &H10
    ElseIf pPalette.Flags = 1& And (DataOrientation And orient_WantPaletteInArray) = 0& Then
        lColorType = lvicColor4bpp: rShift = &H10
    Else
        lColorType = lvicColor1bpp: rShift = &H2
    End If
    
    ' loop thru the image, counting colors
    lPrevColor = inStream(I, J) Xor 1&
    With SS
        For J = 0& To SS.Height - 1&
            For I = 0& To SS.Width - 1&
                If inStream(I, J) <> lPrevColor Then
                    lPrevColor = inStream(I, J)
                    If (lPrevColor And &HFF000000) = &HFF000000 Then
                        bAlpha = False
                    Else
                        lResult = (lPrevColor And &H7F000000) \ &H1000000
                        If lPrevColor < 0& Then lResult = lResult Or &H80
                        bAlpha = (lResult <= .AlphaTolerancePct)
                    End If
                    If bAlpha = False Then
                        lResult = modCommon.FindColor(pPalette.Entries(), lPrevColor Or &HFF000000, pPalette.Count, bAlpha)

                        If bAlpha Then
                            pPalette.Count = pPalette.Count + 1&
                            If lResult < pPalette.Count Then CopyMemory pPalette.Entries(lResult + 1&), pPalette.Entries(lResult), (pPalette.Count - lResult) * 4&
                            pPalette.Entries(lResult) = lPrevColor Or &HFF000000
                            If pPalette.Count = 256& Then       ' stop counting; palette maxed out?
                                J = SS.Height: Exit For
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End With
    
    If pPalette.Count = 0& Then Exit Function   ' Alpha ratio is too low; all pixels excluded
    ' handle special cases
    If pPalette.Count = 1& Then  ' one color only
        If (pPalette.Entries(1) And &HFF&) < 128& Then lMedian = 2&
    Else
        lMedian = pPalette.Count \ 2&
    End If
    
    scanWidth = modCommon.ByteAlignOnWord((lColorType And &HFF00&) \ &H100&, SS.Width)    ' destination scanwidth
    mScanWidth = modCommon.ByteAlignOnWord(1, SS.Width)                                   ' mask scanwidth
    If (SS.reserved1 And &H20000000) Then       ' calling routine wants source image B&W for other reasons
        If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, lvicColor1bpp, ByVal 0&, pvCreateBlackAndWhite) Then GoTo ExitRoutine
        If GdipBitmapLockBits(pvCreateBlackAndWhite, sizeI, ImageLockModeWrite, lvicColor1bpp, dBMP) Then
            GdipDisposeImage pvCreateBlackAndWhite          ' error; abort
            GoTo ExitRoutine
        End If
        bWantNewPic = True
        With sSA
            .cbElements = 1: .cDims = 1
            .pvData = dBMP.Scan0Ptr                          ' write b&w data here
            .rgSABound(0).cElements = dBMP.Height * mScanWidth
        End With
        CopyMemory ByVal VarPtrArray(bwStream), VarPtr(sSA), 4&
    End If
    
    If (DataOrientation And orient_WantPaletteInArray) Then
        I = 2 ^ ((lColorType And &HFF00&) \ &H100&) * 4&
    Else
        I = 0&
    End If
    If bWantMask Then                                       ' mask wanted; adjust array to include palette too
        dSA.rgSABound(0).cElements = SS.Height * (scanWidth + mScanWidth)
        mskOffset = dSA.rgSABound(0).cElements - mScanWidth ' offset to where mask will be written
    Else
        dSA.rgSABound(0).cElements = SS.Height * scanWidth
    End If
    ReDim outData(0 To dSA.rgSABound(0).cElements + HeaderSize + FooterSize + I - 1&)
    dSA.pvData = VarPtr(outData(HeaderSize + I))        ' write b&w data here
    dSA.cbElements = 1
    dSA.cDims = 1
    CopyMemory ByVal VarPtrArray(outStream), VarPtr(dSA), 4&
    
    lPrevColor = inStream(0, 0) Xor 1&                      ' loop thru the image one more time
    With SS
        For I = 0& To sizeI.nHeight - 1&
            mPtr = mskOffset - I * mScanWidth: mShift = &H80    ' mask pointer & shifting value
            srcPtr = I * mScanWidth                             ' source pointer if re-writing image
            If (DataOrientation And orient_BottomUp) Then       ' destination pointer & shifting value
                dPtr = (.Height - I - 1&) * scanWidth
            Else
                dPtr = I * scanWidth
            End If
            If rShift = &H2 Then lShift = &H80 Else lShift = rShift
            For J = 0& To sizeI.nWidth - 1&
                If lPrevColor <> inStream(J, I) Then
                    lPrevColor = inStream(J, I)
                    If (lPrevColor And &HFF000000) = &HFF000000 Then    ' has transparency; how to handle it
                        bAlpha = False                                  ' flag to include this color
                    Else
                        lResult = (lPrevColor And &H7F000000) \ &H1000000
                        If lPrevColor < 0& Then lResult = lResult Or &H80
                        bAlpha = (lResult <= .AlphaTolerancePct)    ' flag to allow/exclude color
                    End If
                    If bAlpha = False Then
                        lResult = modCommon.FindColor(pPalette.Entries(), lPrevColor Or &HFF000000, pPalette.Count, bAlpha)
                    End If
                End If
                If bAlpha Then                              ' transparent pixel
                    If bWantMask Then                       ' add to mask
                        outStream(mPtr) = outStream(mPtr) Or mShift
                    Else                                    ' make it ref palette index #3
                        outStream(dPtr) = outStream(dPtr) Or lShift * 2
                    End If
                ElseIf lResult > lMedian Then               ' else add if white vs. black
                    outStream(dPtr) = outStream(dPtr) Or lShift ' and update source if rewriting it
                    If bWantNewPic Then bwStream(srcPtr) = bwStream(srcPtr) Or mShift
                End If
                If mShift = 1 Then                          ' update mask/source shift & pointers
                    mShift = &H80: mPtr = mPtr + 1&: srcPtr = srcPtr + 1&
                Else
                    mShift = mShift \ 2                     ' update destination shift
                End If
                If lShift = 1 Then                          ' done with target byte?
                    If rShift = &H2 Then lShift = &H80 Else lShift = rShift
                    dPtr = dPtr + 1&                        ' update shift & move pointer along
                Else
                    lShift = lShift \ rShift                ' update shift
                End If
            Next
        Next
    End With
    pPalette.Count = 2&                                 ' default palette
    If (DataOrientation And orient_BlackIs1Not0) Then pPalette.Entries(1) = &HFF000001 Else pPalette.Entries(1) = &HFF000000
    pPalette.Entries(2) = &HFFFFFFFF
    
    GdipBitmapUnlockBits Handle, tBMP: tBMP.Scan0Ptr = 0&                       ' release source
    CopyMemory ByVal VarPtrArray(inStream), 0&, 4&: tSA.pvData = 0&             ' remove source overlay
    CopyMemory ByVal VarPtrArray(outStream), 0&, 4&: dSA.pvData = 0&            ' remove destination overlay
    If bWantNewPic Then
        GdipBitmapUnlockBits pvCreateBlackAndWhite, dBMP: dBMP.Scan0Ptr = 0&    ' release new source
        CopyMemory ByVal VarPtrArray(bwStream), 0&, 4&: sSA.pvData = 0&         ' and its overlay
        GdipDisposeImage Handle                                                 ' destroy old source
        Handle = pvCreateBlackAndWhite                                          ' set to new source
        GdipSetImagePalette Handle, pPalette                                    ' set new source palette
        pvCreateBlackAndWhite = 0&                                              ' default return value
    End If
    
    If pPalette.Flags = 1& And bWantMask = False Then                           ' source has transparency; up to 4bpp if needed
        pPalette.Count = 3&                             ' add another palette entry & make it transparent
        pPalette.Entries(3) = (SS.RSS.FillColorARGB And &HFFFFFF)
    End If
    
    If (DataOrientation And &H3) = 0& Then                  ' return a GDI+ bitmap
        If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, lColorType, ByVal 0&, pvCreateBlackAndWhite) Then Exit Function
        tBMP.Scan0Ptr = VarPtr(outData(HeaderSize))
        tBMP.stride = scanWidth
        If GdipBitmapLockBits(pvCreateBlackAndWhite, sizeI, ImageLockModeWrite Or ImageLockModeUserInputBuf, lColorType, tBMP) Then
            GdipDisposeImage pvCreateBlackAndWhite
            pvCreateBlackAndWhite = 0&
        Else
            GdipBitmapUnlockBits pvCreateBlackAndWhite, tBMP
            GdipSetImagePalette pvCreateBlackAndWhite, pPalette
        End If
        tBMP.Scan0Ptr = 0&
    Else
        If (DataOrientation And orient_WantPaletteInArray) Then ' make 2nd palette entry White
            outData(HeaderSize + 4) = 255: outData(HeaderSize + 5) = 255: outData(HeaderSize + 6) = 255
        End If
        pvCreateBlackAndWhite = lColorType                  ' return value
    End If
    
ExitRoutine:                                                ' clean up if error encountered
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(inStream), 0&, 4&
    If dSA.pvData Then CopyMemory ByVal VarPtrArray(outStream), 0&, 4&
    If sSA.pvData Then CopyMemory ByVal VarPtrArray(bwStream), 0&, 4&
    If tBMP.Scan0Ptr Then GdipBitmapUnlockBits Handle, tBMP
    If dBMP.Scan0Ptr Then GdipBitmapUnlockBits pvCreateBlackAndWhite, tBMP
    If Err Then
        If pvCreateBlackAndWhite Then GdipDisposeImage pvCreateBlackAndWhite
        pvCreateBlackAndWhite = 0&
    End If
End Function
