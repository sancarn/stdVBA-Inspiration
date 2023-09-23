VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpMultiImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Manage multi-image/frame formats ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' Image Formats supported:
'   Animated GIFs. Single GIF formats are handled by the GDIpImage class
'   Multi-page TIFFs. Single TIF formats are handled by GDIpImage class
'   Icons & Cursors. Regardless how many exist in source, class handles all icons/cursors
'   Animated cursors. Regardless how many exist in source, class handles all animated cursors
'   Animated PNGs. Single PNG formats are handled by the GDIpImage class
'   Segmented bitmaps. These are like a film strip, one bitmap, but frames built-in like tiles/grid
'   Binaries (exe,dll,ocx). Class handles all cursor,icon,bitmap,AVI images in the executable

' This class may contain a separate GDI+ handle than contained in its parent GDIpImage class
'   All Icons/Cursors/binaries. GDI+ handle always maintained here, GDIpImage class never maintains one
'       Source data for the handle maintained here too; may be IStream or bits array & a subset of original source
'       -- for 32bpp alpha icons/cursors, no separate sub-source data is maintained
'       -- icons/cursors never cloned
'   TIFF. GDI+ handle maintained here only if the frame is cloned.
'       -- See GDIpImage remarks at top of GDIpImage class for cloning comments
'       -- Source data maintained here only if handle maintained here
'   GIF/PNG. GDI+ handle maintained here for multi-frame formats only
'       -- individual frame data maintained by m_GIFfunction/m_PNGfunction as apprropriate
'   Segmented bitmaps. GDI+ handle never maintained here; GDIpImage maintains it
'       -- Source data never maintained here
'       -- Segmented images never cloned
'       -- Cannot segment icons/cursors. GIF/TIFF cannot be segmented if they contain multiple images

' Animation tracking is done by using a sequence/duration array.
'   - Incrementing image index automatically wraps to beginning if increment surpasses sequence count
'   - Decrementing image index automaitcally wraps to end if decrement results in zero
'   - The sequence count can exceed actual frame count. A cycle occurs when sequence is fully run
'       :: Example: animated cursors may have 5 frames, but sequence may look like: 1,2,3,4,5,4,3,2
'   - Any multi-frame image can be animated, including icons, cursors, GIFs, PNGs, TIFFs, segemented bitmaps


Public Event FrameChanged(Index As Long, Width As Single, Height As Single, UpdateDimensions As Boolean)

Private Type PropertyItem
    ID As Long
    length As Long
    Type As Long
    Value As Long
End Type

Private m_Mode As Long                      ' type of multi-image format. See InitializeContent
Private m_SeqDur() As Long                  ' sequence/duration array
Private m_Index As Long                     ' current index in sequence
Private m_Handle As Long                    ' GDI+ handle, as needed
Private m_Count As Long                     ' sequence count
Private m_CountActual As Long               ' actual frame count; can be more/less than sequence count
Private m_GUID() As Long                    ' GUID used for navigating frames in GIF/TIFF
Private m_Source As Variant                 ' source data if GDI+ handle exists, as needed
Private m_CachedData As Variant             '
Private m_IconFunction As cFunctionsICO     ' used for parsing icons/cursors
Private m_GIFfunction As cFunctionsGIF      ' used for parsing gif frames
Private m_PNGfunction As cFunctionsPNG      ' used for parsing animated PNG
Private m_AVIfunction As cFunctionsAVI      ' used for extracting AVI frames
Private m_DLLfunction As cFunctionsDLL      ' used for parsing/extracting DLL/EXE resources

Public Property Get Index() As Long
    Index = (m_Index And &HFFFF&)           ' return index of current image
End Property                                ' index is based from sequence

Public Property Get FrameFromIndex() As Long
    FrameFromIndex = m_SeqDur(0, Me.Index)  ' return real index for segmented sequence
End Property

Public Property Get ImageFormat(GroupIndex As Long) As ImageFormatEnum
    Select Case m_Mode
    Case lvicPicTypeFromBinaries
        If GroupIndex = -1& Then
            ImageFormat = m_DLLfunction.ImageFormat()
        Else
            ImageFormat = m_DLLfunction.GroupFormat(GroupIndex)
        End If
    Case lvicPicTypeIcon
        ImageFormat = m_IconFunction.ImageFormat()
    Case Else
        ImageFormat = m_Mode
    End Select
End Property

Public Function ExtractImageData(imgSource As Variant, outArray() As Byte) As Boolean

    If m_Mode = lvicPicTypeFromBinaries Then                ' have class return data in proper format
        ExtractImageData = m_DLLfunction.ExtractData(outArray(), m_Handle)
    
    ElseIf VarType(imgSource) = (vbArray Or vbByte) Then    ' return byte array contained in imgSource
        Dim bDummy() As Byte
        modCommon.MoveArrayToVariant imgSource, bDummy(), False
        outArray() = bDummy()
        modCommon.MoveArrayToVariant imgSource, bDummy(), True
        On Error Resume Next
        ExtractImageData = UBound(outArray())
    
    ElseIf VarType(imgSource) = vbString Then                ' file, should be locked & should still be readable
        Dim hFile As Long, lRead As Long
        hFile = modCommon.GetFileHandle(CStr(imgSource), False)
        If hFile <> INVALID_HANDLE_VALUE Then
            ReDim outArray(0 To GetFileSize(hFile, 0&) - 1&) ' read data from that file
            ReadFile hFile, outArray(0), UBound(outArray) + 1&, lRead, ByVal 0&
            CloseHandle hFile
            ExtractImageData = (lRead > UBound(outArray))
        End If
    
    ElseIf VarType(imgSource) = vbDataObject Then
        If TypeOf imgSource Is IUnknown Then             ' data contained in a stream, return its bytes
            ExtractImageData = modCommon.IStreamToArray(ObjPtr(imgSource), outArray())
        End If
    End If

End Function

Public Sub SetIndex(sourceData As Variant, gHandle As Long, ByVal newIndex As Long, Optional GroupIndex As Long = 0&)
    
    ' routine retrieves image from a multi-image format
    ' Icons/Cursors are created on demand
    ' TIFF/GIF frames are navigated via GDI+ functions
    ' Passing &HF0000000 for the newIndex has special meaning....
    '   Icons/Cursors: Find first frame that meets system requirements
    '   Animated Cursors: Find first image in group that meets system requirements (can contain multiple groups of cursors)
    '   Binaries (dll/exe/ocx). Move to first image in first group
    '   Segmented/GIF/TIFF: Move to first image
    
    Dim lResult As Long, tData() As Byte, bDummy() As Byte, IStream As IUnknown
    Dim actualIndex As Long, bOK As Boolean, bBinary As Boolean
    Dim rSize As RECTF, hHandle As Long
    
    If newIndex < 1& Then
        If newIndex = &HF0000000 Then                           ' special value used for 1st time load
            Select Case m_Mode
                Case lvicPicTypeIcon
                    actualIndex = -1&                           ' flag for IconToBits call
                    newIndex = m_Index Xor 1&                   ' ensure these 2 don't match
                Case lvicPicTypeFromBinaries
                    actualIndex = -1&
                    newIndex = m_Index Xor 1&
                Case Else
                    m_Index = 1&: newIndex = m_Index            ' always use first image
            End Select
        Else                                                    ' roll around to last in the sequence
            If m_Count = 1& Then
                newIndex = m_Count
            Else
                newIndex = m_Count + (newIndex Mod m_Count)
                If newIndex = 0& Then newIndex = m_Count
            End If
            If m_Mode = lvicPicTypeFromBinaries Then
                actualIndex = newIndex - 1&
            Else
                actualIndex = m_SeqDur(0, newIndex)
            End If
        End If
    Else
        If newIndex > m_Count Then                              ' roll around to 1st in the sequence
            If m_Count = 1& Then
                newIndex = m_Count
            Else
                newIndex = newIndex Mod m_Count
                If newIndex = 0& Then newIndex = m_Count
            End If
        End If
        If m_Mode = lvicPicTypeFromBinaries Then
            actualIndex = newIndex - 1&
        Else
            actualIndex = m_SeqDur(0, newIndex)
        End If
    End If
    
    If GroupIndex = 0& Then                                     ' initially, this byte may not be set
        GroupIndex = (m_Index \ &H10000)
        If GroupIndex = 0& Then GroupIndex = 1&
    End If
    If (newIndex Or GroupIndex * &H10000) = m_Index Then Exit Sub
    
    Select Case m_Mode
    Case lvicPicTypeIcon, lvicPicTypeFromBinaries ' icon/cursor/animated cursor/binary
        Me.KillClone
        If m_Mode = lvicPicTypeIcon Then
            lResult = m_IconFunction.IconToBits(sourceData, actualIndex, GroupIndex, tData(), rSize.nWidth, rSize.nHeight, newIndex, m_Count)
        Else
            lResult = m_DLLfunction.SetIndex(actualIndex, GroupIndex, tData(), newIndex, m_Count)
        End If
        Select Case lResult
        Case lvicPicTypePNG                                     ' png image in array
            Set IStream = modCommon.IStreamFromArray(VarPtr(tData(0)), UBound(tData) + 1&)
            Erase tData()
            If GdipLoadImageFromStream(ObjPtr(IStream), hHandle) = 0& Then
                Set m_Source = IStream
                GdipGetImageBounds hHandle, rSize, UnitPixel
                bOK = True
            End If
        Case lvicPicTypeNone                                    ' something was wrong with image
            If newIndex < 0& Then newIndex = 1&
            If GroupIndex < 0& Then GroupIndex = 1&
        Case Else                                               ' GDI+ image
            hHandle = lResult
            If m_Mode = lvicPicTypeFromBinaries Then GdipGetImageBounds hHandle, rSize, UnitPixel
            bOK = True
        End Select
        If GroupIndex = -1& Then
            GroupIndex = newIndex
            newIndex = 1&
        End If
    Case lvicPicTypeTIFF                                        ' multi-page TIFF
        Me.KillClone
        On Error GoTo ExitRoutine
        If GdipImageSelectActiveFrame(gHandle, m_GUID(0), actualIndex) = 0& Then
            GdipGetImageBounds gHandle, rSize, UnitPixel
            bOK = True
        End If
    Case lvicPicTypeGIF                                         ' animated GIF
        On Error GoTo ExitRoutine
        If m_GIFfunction.MoveToFrame(actualIndex, sourceData, gHandle) Then
            GdipGetImageBounds gHandle, rSize, UnitPixel
            bOK = True
        End If
    Case lvicPicTypePNG                                         ' animated PNG
        On Error GoTo ExitRoutine
        If m_PNGfunction.MoveToFrame(actualIndex, sourceData, gHandle) Then
            GdipGetImageBounds gHandle, rSize, UnitPixel
            bOK = True
        End If
    Case lvicPicTypeBitmap                                      ' segmented bitmap
        bOK = True                                              ' nothing to do here, routine changes the index only
    Case lvicPicTypeAVI
        Me.KillClone
        On Error GoTo ExitRoutine
        hHandle = m_AVIfunction.MoveToFrame(actualIndex)
        If hHandle Then
            GdipGetImageBounds hHandle, rSize, UnitPixel
            bOK = True
        End If
    Case Else
        Exit Sub
    End Select
    
ExitRoutine:
    If bOK Then m_Handle = hHandle
    m_Index = newIndex Or GroupIndex * &H10000
    RaiseEvent FrameChanged(newIndex, rSize.nWidth, rSize.nHeight, bOK)
End Sub

Public Property Get Frames(Optional GroupIndex As Long) As Long
    If m_Mode = lvicPicTypeFromBinaries Then
        Frames = m_DLLfunction.GroupImageCount(GroupIndex)
    Else
        Frames = m_Count
    End If
End Property

Public Property Get FramesActual() As Long
    FramesActual = m_CountActual
End Property

Public Property Get Handle() As Long
    Handle = m_Handle
End Property

Public Property Get Interval(Index As Long) As Long
    
    If m_Mode = lvicPicTypeFromBinaries Then
        Interval = 0&
    ElseIf Not (Index < 1& Or m_Index > m_Count) Then
        Interval = m_SeqDur(1, Index)
    Else
        Interval = m_SeqDur(1, Me.Index)
    End If
End Property

Public Property Get IsBinary() As Boolean
    IsBinary = (m_Mode = lvicPicTypeFromBinaries)
End Property

Public Function SetSequenceDuration(newSequence() As Long) As Boolean
   
   ' The passed sequence must be a 2D Long array in the following format
   ' newSequence(0 To 1, 1 to nrItems in Sequence)
   ' newSequence(0, nr) are the frame references
   '    -- cannot exceed number of actual frames
   '    -- references are 0-bound
   ' newSequence(1, nr) are the frame durations
   '    -- must be between 0 & 60000 to support VB timer intervals
   ' Any deviations will be modified
   
    On Error Resume Next
    If UBound(newSequence, 2) < 1 Then Exit Function
    If UBound(newSequence, 1) <> 1 Then Exit Function
    If LBound(newSequence, 1) <> 0 Then Exit Function
    If Err Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    Dim X As Long
    
    m_SeqDur() = newSequence()
    m_Count = UBound(m_SeqDur, 2)
    For X = 1& To m_Count
        If m_SeqDur(0, X) < 0& Then m_SeqDur(0, X) = 0&
        If m_SeqDur(0, X) >= m_CountActual Then m_SeqDur(0, X) = m_CountActual - 1&
        If (m_SeqDur(1, X) And &HFFFF&) > 60000 Then m_SeqDur(1, X) = 60000
    Next

End Function

Public Sub GetSequenceDuration(outArray() As Long)
    outArray = m_SeqDur
End Sub

Public Function GetAniGIFloops(Handle As Long) As Long
    
    If m_Mode = lvicPicTypeGIF Then
        GetAniGIFloops = m_GIFfunction.LoopCount
    ElseIf m_Mode = lvicPicTypePNG Then
        GetAniGIFloops = m_PNGfunction.LoopCount
    End If

End Function

Public Function GetWidth(sourceData As Variant, gHandle As Long, Index As Long) As Single
    If (Index < 1& Or Index > m_Count) Then Index = Me.Index
    pvGetFrameDimensions sourceData, gHandle, Index, GetWidth, 0!
End Function

Public Function GetHeight(sourceData As Variant, gHandle As Long, Index As Long) As Single
    If (Index < 1& Or Index > m_Count) Then Index = Me.Index
    pvGetFrameDimensions sourceData, gHandle, Index, 0!, GetHeight
End Function

Public Function GetGroupCount(GroupIndex As Long) As Long

    If m_Mode = lvicPicTypeIcon Then
        GetGroupCount = m_IconFunction.GroupCount
    ElseIf m_Mode = lvicPicTypeFromBinaries Then
        GetGroupCount = m_DLLfunction.GroupCount
    Else
        GetGroupCount = 1
    End If

End Function

Public Function GetBitDepth(sourceData As Variant, ByVal gHandle As Long, ByVal Index As Long) As Long
    If (Index < 1& Or Index > m_Count) Then Index = Me.Index
    Select Case m_Mode
        Case lvicPicTypeFromBinaries
            GetBitDepth = m_DLLfunction.BitDepth(Index - 1&)
        Case lvicPicTypeIcon
            GetBitDepth = m_IconFunction.IconBitDepth(sourceData, m_SeqDur(0, Index), (m_Index \ &H10000))
        Case Else
            If gHandle = 0& Then gHandle = m_Handle
            If gHandle Then
                GdipGetImagePixelFormat gHandle, Index
                GetBitDepth = (Index And &HFF00&) \ &H100&
            End If
    End Select
End Function

Public Sub InitializeContent(Handle As Long, ImageType As Long, ActualCount As Long)

    Const FrameDimensionPage As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"

    ' this is called from various processing routines to setup the the class

    Dim X As Long, tObj As Object
    On Error GoTo ExitRoutine
    
    Me.KillClone
    m_CachedData = Empty
    
    Select Case ImageType
    
    Case lvicPicTypeIcon, lvicPicTypeCursor, lvicPicTypeAnimatedCursor
        m_Mode = lvicPicTypeIcon
        Set m_IconFunction = Nothing
        CopyMemory tObj, Handle, 4&
        Set m_IconFunction = tObj
        CopyMemory tObj, 0&, 4&
        ' SetSequenceDuration routine must be called to populate sequence & count
    
    Case lvicPicTypeFromBinaries
        m_Mode = ImageType
        m_Count = ActualCount
        Set m_DLLfunction = Nothing
        CopyMemory tObj, Handle, 4&
        Set m_DLLfunction = tObj
        CopyMemory tObj, 0&, 4&
        m_Index = 1&
    
    Case lvicPicTypeGIF
        Set m_GIFfunction = Nothing
        CopyMemory tObj, Handle, 4&
        Set m_GIFfunction = tObj
        CopyMemory tObj, 0&, 4&
        m_Mode = ImageType
    
    Case lvicPicTypeTIFF
        m_Mode = ImageType
        ReDim m_GUID(0 To 3)
        ' get count & set cycle=count
        CLSIDFromString StrPtr(FrameDimensionPage), m_GUID(0)
        Call GdipImageGetFrameCount(Handle, m_GUID(0), ActualCount)
        If ActualCount > 1& Then                        ' count < 2 & this class will be aborted used
            m_Count = ActualCount
            ReDim m_SeqDur(0 To 1, 1 To m_Count)
            For X = 1& To m_Count
                m_SeqDur(0, X) = X - 1&
            Next
            m_Index = 1&
        End If
    
    Case lvicPicTypeAVI
        Set m_AVIfunction = Nothing
        CopyMemory tObj, Handle, 4&
        Set m_AVIfunction = tObj
        CopyMemory tObj, 0&, 4&
        m_Mode = lvicPicTypeAVI
        m_Index = 1&
        
    Case lvicPicTypePNG
        Set m_PNGfunction = Nothing
        CopyMemory tObj, Handle, 4&
        Set m_PNGfunction = tObj
        CopyMemory tObj, 0&, 4&
        m_Mode = lvicPicTypePNG
        
    Case Else
        m_Mode = lvicPicTypeBitmap                          ' segmented bitmap
        ' SetSequenceDuration routine must be called to populate sequence & count
    End Select
    m_CountActual = ActualCount
   
ExitRoutine:
    If Err Then
        Err.Clear
        m_Count = 0&
    End If
End Sub

Public Sub KillClone()
    If m_Handle Then
        GdipDisposeImage m_Handle
        m_Handle = 0&
    End If
    m_Source = Empty
End Sub

Public Function CreateClone(gHandle As Long) As Long
    Call Me.KillClone
    m_Handle = modCommon.CreateSourcelessHandle(gHandle)
    CreateClone = m_Handle
End Function

Public Function CacheSourceInfo(sourceData As Variant, Handle As Long, ImageFormat As ImageFormatEnum, CacheData As Boolean, bRetrieve As Boolean, Optional bAssume As Boolean)
    
    ' this routine is used by the modCommon.LoadImage function to cache processed information
    ' this routine is then used by the GDIpImage class to retrieve that information
    
    Dim bDummy() As Byte
    If bRetrieve Then
        CacheData = CBool(m_Index And &H80000000)
        ImageFormat = m_Mode
        If bAssume Then
            Handle = m_Handle: m_Handle = 0&
        Else
            Select Case ImageFormat
            Case lvicPicTypeAnimatedCursor, lvicPicTypeIcon, lvicPicTypeCursor, lvicPicTypeFromBinaries, lvicPicTypeAVI
                ' nothing to do regarding the handle; handle is maintained here, not the GDIpImage class
                m_Index = m_Index And &H7FFFFFFF
            Case Else
                ' handle is maintained by the GDIpImage class, not here
                Handle = m_Handle: m_Handle = 0&
                m_Index = 1&
            End Select
        End If
        If IsObject(m_CachedData) Then
            Set sourceData = m_CachedData
        ElseIf VarType(m_CachedData) = (vbArray Or vbByte) Then
            modCommon.MoveArrayToVariant m_CachedData, bDummy, False
            modCommon.MoveArrayToVariant sourceData, bDummy, True
        Else
            sourceData = m_CachedData
        End If
        m_CachedData = Empty
    Else
        If m_Mode = 0& Then     ' otherwise, handle was created in InitializeContent
            Me.KillClone        ' and not destroyed here as this class maintains it
            m_Handle = Handle   ' also the image format exists already in m_Mode
            m_Mode = ImageFormat
        Else
            Select Case m_Mode
            Case lvicPicTypeGIF, lvicPicTypePNG, lvicPicTypeTIFF, lvicPicTypeAVI
                m_Handle = Handle
            End Select
        End If
        
        If IsObject(sourceData) Then
            Set m_CachedData = sourceData
        ElseIf VarType(sourceData) = vbLong Then
            CopyMemory ByVal VarPtrArray(bDummy), ByVal CLng(sourceData), 4&
            CopyMemory ByVal CLng(sourceData), 0&, 4&
            modCommon.MoveArrayToVariant m_CachedData, bDummy, True
        Else
            m_CachedData = sourceData
        End If                  ' m_index may already contain a valid value
        m_Index = Abs(CacheData) * &H80000000 Or m_Index
    End If
End Function

Private Sub pvGetGIFDelays(Handle As Long)
    
'    Dim bPropData() As Byte, lDuration() As Long
'    Dim I As Long
'    Dim lPropSize As Long
'    Dim tPropItem As PropertyItem
'
'    Const PropertyTagTypeLong = 4&
'    Const PropertyTagTypeShort = 3&
'    Const PropertyTagFrameDelay As Long = &H5100&
'
'    On Error GoTo ErrorHandler
'    If GdipGetPropertyItemSize(Handle, PropertyTagFrameDelay, lPropSize) = 0& Then
'        ' get the frame delays
'        ReDim bPropData(0 To lPropSize - 1)
'        If GdipGetPropertyItem(Handle, PropertyTagFrameDelay, lPropSize, bPropData(0)) = 0& Then
'            CopyMemory tPropItem, bPropData(0), 16&
'            If tPropItem.Type = PropertyTagTypeLong Then
'                If tPropItem.Length = m_Count * PropertyTagTypeLong Then
'                    ReDim lDuration(1 To m_Count)
'                    CopyMemory lDuration(1), ByVal tPropItem.Value, tPropItem.Length
'                    For I = 1 To m_Count
'                        m_SeqDur(1, I) = (lDuration(I) And &HFFF&) * 10&
'                        If m_SeqDur(1, I) > 60000 Then m_SeqDur(1, I) = 60000
'                    Next
'                    Erase lDuration()
'                End If
'            End If
'        End If
'    End If
ErrorHandler:
End Sub

Private Sub pvGetFrameDimensions(sourceData As Variant, gHandle As Long, Index As Long, Width As Single, Height As Single)
    On Error Resume Next
    Dim imgSize As RECTF
    Select Case m_Mode
        Case lvicPicTypeTIFF ' tiffs
            Call GdipImageSelectActiveFrame(gHandle, m_GUID(0), m_SeqDur(0, Index))
            GdipGetImageBounds gHandle, imgSize, UnitPixel
            Call GdipImageSelectActiveFrame(gHandle, m_GUID(0), m_SeqDur(0, Me.Index))
            Width = imgSize.nWidth: Height = imgSize.nHeight
        Case lvicPicTypeGIF, lvicPicTypePNG, lvicPicTypeAVI
            If gHandle Then
                GdipGetImageBounds gHandle, imgSize, UnitPixel
            ElseIf m_Handle Then
                GdipGetImageBounds m_Handle, imgSize, UnitPixel
            End If
            Width = imgSize.nWidth: Height = imgSize.nHeight
        Case lvicPicTypeFromBinaries
            Call m_DLLfunction.GetDimensions(Index - 1&, Width, Height)
        Case Else ' icons & cursors
            Call m_IconFunction.IconSizes(sourceData, m_SeqDur(0, Index), (m_Index \ &H10000), Width, Height)
    End Select

End Sub

Private Sub Class_Terminate()
    Call Me.KillClone
End Sub
