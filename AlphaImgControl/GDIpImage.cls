VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIpImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'  -----======== PURPOSE: Manage/Render GDI+ Images ========-----
' Class Instancing = MultiUse (can be created by users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' This class is the return value of the usercontrol's .Picture property

' GDI+ is a really nice tool and has many functions to load image formats that VB cannot.
' It also has many functions to render images quickly while applying niffty effects
'   that negate long-winded manual VB6 code: rotation, bicubic sizing, color modifications, etc
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' However, GDI+ has several limitations and some bugs that must be overcome...
' 1. Non-Indexed Icons 24bpp+ seem to be problems. Cursors won't load
' 2. Fails to load icons if a PNG is embedded (Vista+)
' 3. Bitmaps that use alpha-channels, ignores the alpha-channel
' 4. Metafiles cannot be rendered using image attributes (i.e., grayscaling); cannot be mirrored
' 5. When using image attributes, GDI+ flipping & cloning functions on multi-image formats (gif/tif),
'       can result in no longer being able to select any frame/page other than the first.
'    Even without attributes applied, manual mirroring and 180 degree rotations can break this navigation also.
' 6. Has some issues with GIF frames and some rare cases where JPG files load but report 0x0 dimensions
' 7. Additional image formats are processed but not recognized by GDI+: TGA, PCX, PGM, PPM, PBM, PAM, MP3, WMA, AVI, animated cursors/PNG
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' To overcome the above problems the following rules are applied to numbered items above
' 1. All icons/cursors are converted internally to 32bpp bitmap formats. Original format is cached
'    Note: When first loaded, icon/cursors will be scanned & one that best fits system icon size & highest bit depth is used
' 2. Since icons are custom processed, PNGs are not an issue and handled explicitly internally
' 3. Bitmap is tested to see if alpha-channel is used & if so, loaded differently to honor that channel
' 4. Metafiles are converted internally to 32bpp Bitmap
' 5. A copy of the current frame is made & a new GDI+ handle is created from that copy.
'       That copy is rendered instead, thus preventing the breaking of frame navigation
'       This "cloning" is done on the fly and removed when no longer needed. Adds some code overhead
' 6. GIFs are processed manually. Original GIF data maintained. JPGs may be processed manually & original data maintained
' 7. For the unsupported image formats, manual parsing of format is required & additional notes follow
'       Animated cursors/PNGs. Original format is cached
'       TGA/PCX/PBM/PPM/PGM/PAM. Original format is cached
'       MP3/WMA. If images can be extracted from the file, the original format is lost and also...
'           If only 1 image is detected, original format is maintained (not the file, but the image format itself)
'           If 2 or more images detected, original formats are lost & images are converted to multi-page TIFF
'       Binaries(DLL/OCX/EXE). Image offsets are cached, image extracted on demand; binary file never cached
'       AVI. The original format is cached only when this class' ExtractImageData is called. Otherwise, AVIs are played from file
' About memory usage...
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' GDI+ requires any source used for an image to remain available & unchanged. This means when you load
'       an image from a file, the file cannot be modified, moved or deleted. If you load an image from
'       a stream or byte array, that stream/array must not be destroyed.
' This class handles those situations differently....
' - by File. Class does not cache file data unless LoadPictureGDIplus's "Prevent_FileLock_AsyncDownload" parameter is true
'       Otherwise, the file will be locked by GDI+ and must not be altered until this class is destroyed
'       Image formats GDI+ has issues with (noted above) are excluded and file is read into an array & processed manually
'       Exceptions:
'           AVIs are always loaded via file name, data never cached unless ExtractImageData called
'           Binaries are never cached or loadable from array. Individual images are extracted/processed on demand
' - by URL. Data is read into an array. Array is kept until class destroyed
' - by array/stream. Class caches array or creates IStream object. Array/IStream kept until class destroyed
' - by stdPicture object. Depends on stdPicture object content:
'       :: GIF/JPG. If original format can be extracted, then extracted into IStream else VB provides bitmap only
'       :: WMF/EMF. Converted to 24/32bpp bitmap as needed. Metafile format is cached
'       :: Icon/Cursor. Original format cached to array
'       :: Bitmap. Regardless of bit depth, alpha-channel assumed to not be used. No additional info cached
' - by image handle (icon/cursor & bitmap only). Stand-alone images always created; passed handle not managed/destroyed
' - by Clipboard. Supports bitmap/emf/wmf only & handled same as stdPicture object. Files handled as mentioned above
' - by DataObject. Handled same as Files or stdPicture object, depending on DataObject contents.
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' Since this requirement by GDI+ to keep original source data around results in extra memory usage,
' you have an optional parameter in the LoadPictureGDIplus function to create a stand-alone Image. After any image
' is loaded, the class has a property (KeepOriginalFormat) you can set to convert the current image to stand-alone.
' About stand-alone images...
' - Uses less memory consumption & may be desirable if many classes are creating images
' - All stand-alone non-Bitmap images are identified as PNG but are bitmap internally; except bitmap sources
'   For non-Bitmap source images, returned bytes will be PNG by default else original bitmap depth if Bitmap
'   :: Why default to PNG? Many formats support transparency levels. All viewers honor PNG's alpha channels.
' - Multi-image formats: GIF/TIF/PNG/AVI/ico/cur/segmented image. Only the current frame will be used
' - Original image & source data is permanently lost; thus reduces memory consumption, especially for large images
' - Depending on size of image, a speed hit can be noticed while loading an image then converting to stand-alone
' - Note: Though images loaded by file name add no known extra memory usage, the file does remain locked by GDI+
'       It may be much faster to load very large images by filename and allow GDI+ to lock the file
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

' About sharing classes
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' If you have multiple controls wanting to display the same image, you don't want to load that image multiple times
'   which causes unnecessary memory consumption and processing time. Here are the limitations...
' - Multi-image formats (gif/tif/cur/ico/png/segmented image). The following properties/methods effect all shared instances
'   -- ImageIndex
'   -- ImageGroup
'   -- Setting KeepOriginalFormat=False
'   -- SegmentImage
' - Even though more than 1 control may be sharing an image, each control can render it uniquely
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

' About animation
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' - The class' ImageIndex property returns a 1-bound "frame" index of any loaded resource containing multiple images.
'   When incremented and Index exceeds number of contained frames, the Index automatically wraps back to beginning. The
'   reverse is true when decrementing Indexes. This default circular wrap allows any routine to animate without knowing
'   or testing within a loop the start/end of the animation sequence. A simple loop might look like this...
'
'       # thisClass.ImageIndex = 1
'       # Do
'       #   ' clear drawing area
'       #   Call thisClass.Render(hDC, X, Y, ...)
'       #   thisClass.ImageIndex = thisClass.ImageIndex + 1
'       #   ' pause thisClass.FrameDuration() milliseconds
'       #   ' FrameDuration() can be zero & you should supply a default, like maybe 100 ms
'       # Loop
'       # thisClass.ImageIndex = 1
'       # Call thisClass.Render(hDC, X, Y, ...)
'
' - The ImageCount returned by this class may be smaller or larger than the actual number of images/frames
'   contained in the loaded image.
'   :: When icons/cursors are processed, any corrupt images within the resource are ignored.
'       Therefore, the ImageCount can be smaller than expected
'   :: Many animated cursors contain a sequence section that identifies which image is displayed when.
'       If that sequence is smaller (unused cursors) than the number of contained cursors, ImageCount can be smaller
'       than expected. If that sequence repeats cursors, ImageCount can be larger than expected.
'       Example: Windows' stopwtch.ani contains just 8 cursors, however, the sequence consists of 45 frames
' - This class has the ability to use a single image (bitmap, gif, png, etc) and segment it or divide it into multiple images.
'   When that option is invoked, the ImageCount will change from 1 to number of segments, the Width & Height will also change.
'   The actual image is not modified, only the way it is viewed/reported.
'   -- cannot segment an image that has multiple images (animated gif/png, ico, cur, avi, animated cursor, multi-page TIFF)
'      if you need to segment an image when you are otherwise prevented, extract that frame only,
'       using SavePictureGDIplus, into a new image class and segment that copy
' ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Event ImageChanged(FrameIndex As Long)
Public Event AsyncDownloadDone(Success As Boolean, ErrorCode As AsyncDownloadStatusEnum)

Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixOffsetMode As Long) As Long
' v1.1 function
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByRef Source As RECTF, ByVal xForm As Long, ByVal Effect As Long, ByVal imageAttributes As Long, ByVal srcUnit As Long) As Long
' support functions for GdipDrawImageFX to enable rendering at other than 0,0 & allow partial renders
Private Declare Function GdipCreateMatrix2 Lib "GdiPlus.dll" (ByVal mM11 As Single, ByVal mM12 As Single, ByVal mM21 As Single, ByVal mM22 As Single, ByVal mDx As Single, ByVal mDy As Single, ByRef mMatrix As Long) As Long
Private Declare Function GdipDeleteMatrix Lib "GdiPlus.dll" (ByVal mMatrix As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long

Private Enum ClassFlagsEnum
    flag_CanClone = 1                       ' multi-frame TIFF only. GIFs are now handled manually
    flag_Cloned = 2                         ' current handle is a clone; always kept in m_MultiImageAttr class
    flag_Segmented = 4                      ' image treated as several tiled images
    flag_NoEvent = 8                        ' prevent raising a ImageChanged event
    flag_CacheAsyncData = 32                ' used when loading image asyncrhonously
End Enum
Private Type TileStruct                     ' used for segmented images (See SegmentImage)
    Width As Long                           ' Tile width
    Height As Long                          ' Tile height
    Rows As Long                            ' Nr of horizontal tiles
    Columns As Long                         ' Nr of vertical tiles
End Type
Private Type POINTF                         ' used in RenderSkewed method
    X As Single
    Y As Single
End Type

Private m_Token As cGDIpToken
Private WithEvents m_AsyncClient As cAsyncClient
Attribute m_AsyncClient.VB_VarHelpID = -1
Private WithEvents m_MultiImageAttr As cGDIpMultiImage
Attribute m_MultiImageAttr.VB_VarHelpID = -1
Private m_Handle As Long                    ' GDI+ handle created from original source data
Private m_Format As ImageFormatEnum         ' format of image, not necessarily same as original image format
Private m_Source As Variant                 ' image source data, not ncessarily same as was provided
Private m_Size As RECTF                     ' image dimensions
Private m_Flags As ClassFlagsEnum           ' see ClassFlagsEnum
Private m_Tiles As TileStruct               ' used for segmented images
Private m_Group As Long                     ' a Win7 animated cursor group

Public Property Get Handle() As Long
Attribute Handle.VB_UserMemId = 0
    ' Returns current image handle. May be actual or clone
    ' must be default property. DO NOT move this property within this code
    ' To verify/set to default
    '   1. IDE Menu: Tools | Procedure Attributes
    '   2. Find Handle in the "Name" combobox
    '   3. Click the "Advanced" button
    '   4. Select "(Default)" in the "Procedure ID" combobox
    '   5. Click the "Apply" button
    If m_MultiImageAttr Is Nothing Then
        Handle = m_Handle
    Else
        Handle = m_MultiImageAttr.Handle
        If Handle = 0& Then Handle = m_Handle
    End If
End Property

Public Property Get AnimationLoops() As Long
    If Not m_MultiImageAttr Is Nothing Then AnimationLoops = m_MultiImageAttr.GetAniGIFloops(m_Handle)
End Property

Public Property Get BitDepth(Optional Index As Long) As Long
    ' never cached. Get via APIs or parsing as needed
    If Me.Handle Then
        If m_MultiImageAttr Is Nothing Then
            Dim Depth As Long
            GdipGetImagePixelFormat m_Handle, Depth
            BitDepth = (Depth And &HFF00&) \ &H100&
        Else
            BitDepth = m_MultiImageAttr.GetBitDepth(m_Source, m_Handle, Index)
        End If
    End If
End Property

Public Function ExtractImageData(sourceData() As Byte, Optional ImageFormatReturned As ImageFormatEnum) As Boolean

    ' Routine returns a byte array of the entire format, suitable to save to file or send to another
    ' control or routine. If the return value is False, the return parameter values should be ignored.
    ' Sourceless images (KeepOriginalFormat=False), will return PNG format. See notes at top of this class
    
    Dim hFile As Long, lResult As Long
    
    If Me.Handle = 0& Then Exit Function
    ImageFormatReturned = Me.ImageFormat
    
    If Not m_MultiImageAttr Is Nothing Then
        ExtractImageData = m_MultiImageAttr.ExtractImageData(m_Source, sourceData())
    
    ElseIf IsEmpty(m_Source) Then                               ' sourceless/stand-alone image, return PNG format
        If ImageFormatReturned = lvicPicTypeBitmap Then     ' except when 32bpp bitmap is used
            ExtractImageData = (modCommon.SaveAsBMP(sourceData(), m_Handle, saveTo_Array) > 0&)
        Else
            ExtractImageData = (modCommon.SaveAsPNG(sourceData(), m_Handle, saveTo_Array) > 0&)
        End If
    
    Else
        If VarType(m_Source) = vbString Then                ' file, should be locked & should still be readable
            hFile = modCommon.GetFileHandle(CStr(m_Source), False)    ' read data from that file
            If hFile <> INVALID_HANDLE_VALUE Then
                ReDim sourceData(0 To GetFileSize(hFile, 0&) - 1&)
                ReadFile hFile, sourceData(0), UBound(sourceData) + 1&, lResult, ByVal 0&
                CloseHandle hFile
                ExtractImageData = (lResult > UBound(sourceData))
            End If
        
        ElseIf VarType(m_Source) = (vbArray Or vbByte) Then ' byte array, Icons/Cursors done differently
            If ImageFormatReturned = lvicPicTypeBitmap Then
                If (m_Source(0) Or m_Source(1) * &H100&) = &H4D42 Then
                    sourceData() = m_Source
                    ExtractImageData = True
                Else
                    ExtractImageData = (modCommon.SaveAsBMP(sourceData(), Me.Handle, saveTo_Array) > 0&)
                End If
            Else
                sourceData() = m_Source                     ' non-Icon/Cursor, simply return the array
                ExtractImageData = True
            End If
        
        ElseIf VarType(m_Source) = vbDataObject Then        ' data contained in a stream, return its bytes
            If TypeOf m_Source Is IUnknown Then
                ExtractImageData = modCommon.IStreamToArray(ObjPtr(m_Source), sourceData())
            End If
        End If
    End If
    
End Function

Public Property Get FrameDuration(Optional Index As Long)
    ' Returns the duration of frame at Index. May be zero & you should use a default in its place
    If Not m_MultiImageAttr Is Nothing Then FrameDuration = m_MultiImageAttr.Interval(Index)
End Property

Public Property Get Height(Optional Index As Long) As Long
    ' Returns height of current image
    If (m_Flags And flag_Segmented) Then                        ' return tile size
        Height = m_Tiles.Height
    Else
        Height = m_Size.nHeight
        If Index > 0& Then                                     ' return current image size
            If Not m_MultiImageAttr Is Nothing Then
                If m_MultiImageAttr.Index <> Index Then        ' return size of another frame
                    Height = m_MultiImageAttr.GetHeight(m_Source, m_Handle, Index)
                End If
            End If
        End If
    End If
End Property

Public Property Get UsesTransparency() As Boolean
    ' returns whether or not the raw image contains transparency
    UsesTransparency = modCommon.HasTransparency(Me.Handle)
End Property

Public Property Get ImageCount() As Long
    ' Returns number of images contained in sequence and/or parsed from the source
    If m_MultiImageAttr Is Nothing Then
        ImageCount = 1&
    Else
        ImageCount = m_MultiImageAttr.Frames
    End If
End Property

Public Property Get ImageFormat() As ImageFormatEnum
    ' returns image format currently used by the class
    ' This format is not necessarily the same that was originally loaded. See KeepOriginalFormat
    If m_MultiImageAttr Is Nothing Then
        ImageFormat = m_Format
    ElseIf (m_Flags And flag_Segmented) Then
        ImageFormat = m_Format
    Else
        ImageFormat = m_MultiImageAttr.ImageFormat(Me.ImageGroup)
    End If
End Property

Public Property Get ImageGroup() As Long
    ' Win7 can have multiple cursors per animated frame/group. Return current group.
    ' Will return 1 for all other image formats
    If m_Group = 0& Then m_Group = 1&
    ImageGroup = m_Group
End Property

Public Property Let ImageGroup(newIndex As Long)
    If Not m_MultiImageAttr Is Nothing Then
        If Me.Segmented = False Then
            If newIndex > 0& And newIndex <= Me.ImageGroups Then
                If newIndex <> m_Group Then
                    m_Group = newIndex
                    m_MultiImageAttr.SetIndex m_Source, 0&, 1&, m_Group
                End If
            End If
        End If
    End If
End Property

Public Property Get ImageGroupFormat(Optional ByVal GroupIndex As Long) As ImageFormatEnum

    If m_MultiImageAttr Is Nothing Then
        ImageGroupFormat = Me.ImageFormat
    Else
        If GroupIndex < 1& Then GroupIndex = Me.ImageGroup
        ImageGroupFormat = m_MultiImageAttr.ImageFormat(GroupIndex)
    End If
End Property

Public Property Get ImageGroupCount(Optional ByVal GroupIndex As Long) As Long

    Dim imgData() As Byte
    
    If m_MultiImageAttr Is Nothing Then
        ImageGroupCount = 1&
    Else
        If GroupIndex < 1& Then GroupIndex = Me.ImageGroup
        ImageGroupCount = m_MultiImageAttr.GetGroupCount(GroupIndex)
    End If

End Property

Public Property Get ImageGroups() As Long
    ' Win7 can have multiple cursors per animated frame/group. Return number of groups.
    If m_MultiImageAttr Is Nothing Then
        ImageGroups = 1&
    Else
        ImageGroups = m_MultiImageAttr.GetGroupCount(m_Group)
    End If
End Property

Public Property Get ImageIndex() As Long
    ' Returns the index of the current frame
    If m_MultiImageAttr Is Nothing Then
        ImageIndex = 1&
    Else
        ImageIndex = m_MultiImageAttr.Index
    End If
End Property
Public Property Let ImageIndex(newIndex As Long)
    ' sets the current frame
    If Not m_MultiImageAttr Is Nothing Then
        m_MultiImageAttr.SetIndex m_Source, m_Handle, newIndex
    End If
End Property

Public Property Get IsExecutableResource() As Boolean
    ' returns whether or not this class is currently accessing an executable (dll, ocx, exe)
    If Not m_MultiImageAttr Is Nothing Then IsExecutableResource = m_MultiImageAttr.IsBinary
    
End Property

Public Property Get KeepOriginalFormat() As Boolean
    ' If we have a handle, then
    '   handle came from this class &
    '       m_Source will only be Empty if we have no original bytes
    '       m_Source contains an array, IStream object, or filename & original bytes exist in that
    '   or handle came from m_MultiImageAttr and that class always has a source
    If Me.Handle Then KeepOriginalFormat = (IsEmpty(m_Source) = False)
End Property
Public Property Let KeepOriginalFormat(KeepFormat As Boolean)
    ' cannot change from False to True. Once original bytes are destroyed, no way to get them back as-is
    ' When set to false, this property creates a sourceless bitmap internally (transparency included as needed),
    '   and the reported format will always be PNG externally, including extracting any data as PNG
    If KeepFormat = False Then
        If Me.KeepOriginalFormat = True Then
            ' convert to sourceless image
            Dim lResult As Long
            Select Case Me.ImageFormat
                Case lvicPicTypeMetafile, lvicPicTypeEMetafile, lvicPicTypeTGA, lvicPicTypePCX, lvicPicTypePNM
                    ' these formats are already sourceless by default; nothing to do
                Case Else
                    lResult = modCommon.CreateSourcelessHandle(Me.Handle)
                    If lResult Then
                        If m_Handle Then GdipDisposeImage m_Handle
                        m_Handle = lResult
                    Else
                        Exit Property
                    End If
            End Select
            Set m_MultiImageAttr = Nothing
            m_Source = Empty
            m_Flags = 0&
        End If
        If Me.ImageFormat <> lvicPicTypeBitmap Then m_Format = lvicPicTypePNG
    End If
End Property

Public Property Get Left() As Long
    ' should return zero for all but segmented images.
    ' Segemented images will return the Left edge of the current "Tile"
    If (m_Flags And flag_Segmented) Then
        Left = (m_MultiImageAttr.FrameFromIndex Mod m_Tiles.Columns) * m_Tiles.Width
    Else
        Left = m_Size.nLeft
    End If
End Property

Public Property Get Top() As Long
    ' should return zero for all but segmented images.
    ' Segemented images will return the Top edge of the current "Tile"
    If (m_Flags And flag_Segmented) Then
        Top = (m_MultiImageAttr.FrameFromIndex \ m_Tiles.Columns) * m_Tiles.Height
    Else
        Top = m_Size.nTop
    End If
End Property

Public Function RegionFromImage(Optional ByVal OpaqueLevel As Long = 128&, _
                        Optional ByVal XOffset As Long = 0&, Optional ByVal YOffset As Long = 0&) As Long

    ' returns a Windows region handle, or zero if no image assigned or result would be a rectangular region same size as image
    ' the OpaqueLevel must be a value between 0 and 255
    '   if the pixel opaqueness is <= OpaqueLevel then that pixel not included in the region
    
    ' You are responsible for destroying the returned region
    
    ' Notes: The region returned can be used as the HitTest property for the AlphaImgCtl.
    '   Doing so, would exclude any mouse actions that were not contained with the passed region, thus
    '   creating a hit test area of only the pixels contained in the image: a image-shaped region
    ' 1. This routine uses this class' image & dimensions, as non-rotated, non-mirrored
    ' 2. If you are planning to use this region as a HitTest property of the AlphaImgCtl, here are some suggestions:
    '   - Set the AlphaImgCtl's AutoSize property to lvicSingleAngle
    '   - If mirroring, resizing or rotating, create a snapshot of the image and create region from that snapshot..
    '       #   Dim tImg As GDIpImage, hRgn As Long
    '       #   Set tImg = AlphaImgCtl1.SaveImageAsDrawnToGDIpImage()
    '       #   hRgn = tImg.RegionFromImage()
    '       #   If hRgn Then AlphaImgCtl1.HitTest = hRgn
    '       #   Set tImg = Nothing
    '   - Else if using same size image, very similar code
    '       #   Dim hRgn As Long
    '       #   hRgn = AlphaImgCtl1.Picture.RegionFromImage()
    '       #   If hRgn Then AlphaImgCtl1.HitTest = hRgn
    '   - If the target control is not AutoSize'd then you may have to offset the returned region to ensure it
    '       overlays the rendered image in the control
    ' 3. The region is not DPI aware. If you wish to make it DPI aware, you should magnify the returned region
    
    RegionFromImage = modCommon.CreateShapedRegion(Me.Handle, Me.Width, Me.Height, OpaqueLevel, XOffset, YOffset)

End Function

Public Function Render(ByVal hDC As Long, Optional ByVal destX As Long = 0&, Optional ByVal destY As Long = 0&, _
                        Optional ByVal destWidth As Long = 0&, Optional ByVal destHeight As Long = 0&, _
                        Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                        Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&, _
                        Optional ByVal Angle As Single = 0!, _
                        Optional ByVal GDIplusAttributesHandle As Long = 0&, _
                        Optional ByVal GDIplusHGraphicsHandle As Long = 0&, _
                        Optional ByVal GDIplusEffectsHandle As Long = 0&, _
                        Optional ByVal Interpolation As RenderInterpolation = lvicAutoInterpolate) As Boolean

    ' Customizable rendering routine, depending your level of experience with GDI+
    ' By providing parameters for GDIplusHGraphicsHandle, GDIplusAttributesHandle,
    '   and GDIplusEffectsHandle all GDI+ options are made available to you
    ' Additionally, since the Handle is exposed, you can use your own custom rendering routines
    ' But read comments at top of class regarding GIF/TIFF navigation bugs to ensure your don't mess them up
    
    ' Parmaters
    '   :: hDC is destination device context handle
    '       - Ignored when GDIplusHGraphicsHandle is non-zero else must be provided
    '   :: destX,destY are coords to begin rendering to destination
    '   :: destWidth,destHeight are the rendered dimensions
    '       - See Mirroring below
    '   :: srcX,srcY are coords to begin rendering from source image
    '       - Should not be negative (metafiles may be an exception)
    '   :: srcWidth,srcHeight are dimensions of source to be rendered
    '       - See Mirroring below
    '   :: Angle is angle of rotation
    '       - Ignored if GDIplusHGraphicsHandle is non-zero
    '   :: GDIplusAttributesHandle contains a GDI+ Image Attributes handle to adjust rendered pixel values
    '       - GDIplusAttributesHandle is not destroyed nor modified by this routine
    '   :: GDIplusHGraphicsHandle if non-zero should already be transformed, translated, smoothed, etc
    '       - Angle parameter is ignored when GDIplusHGraphicsHandle is non-zero
    '       - hDC parameter is ignored when GDIplusHGraphicsHandle is non-zero
    '       - GDIplusHGraphicsHandle is not destroyed nor modified by this routine
    '   :: GDIplusEffectsHandle if non-zero must be a valid GDI+ v1.1 image effects handle
    '       - can use the GDIpEffects class and pass its EffectsHandle(x) as the GDIplusEffectsHandle parameter
    '       - hDC parameter is ignored when GDIplusHGraphicsHandle is non-zero
    '       - GDIplusEffectsHandle is not destroyed nor modified by this routine
    '       - You should validate loaded version of GDI+ supports Effects (GDIplusTokenVersion = 1.1)
    '   :: Interpolation
    '       - When set to lvicAutoInterpolate, high quality bicubic scaling done only when scaling or rotation is performed
    ' Mirroring....
    ' When the sign of destWidth and srcWidth parameters are different, horizontal mirroring will occur
    ' When the sign of destHeight and srcHeight parameters are different, vertical mirroring will occur
    
    Dim tHandle As Long, Index As Long, srcRect As RECTF
    Dim hGraphics As Long, bClone As Boolean, hMatrix As Long
    Dim scalerX As Single, scalerY As Single, rot180MirrAdj As Single
    Dim actualAngle As Single, svGraphics As Long
    
    tHandle = Me.Handle
    If tHandle = 0& Then
        Render = True
        Exit Function
    End If
    
    If GDIplusHGraphicsHandle = 0& Then                             ' no GDI+ graphics handle passed
        If hDC = 0& Then Exit Function
        If GdipCreateFromHDC(hDC, hGraphics) Then Exit Function
    Else                                                            ' use handle passed
        hGraphics = GDIplusHGraphicsHandle
        GdipSaveGraphics hGraphics, svGraphics
    End If
    
    If Interpolation < lvicAutoInterpolate Or Interpolation > lvicHighQualityBicubic Then Interpolation = lvicAutoInterpolate
    If srcX = 0& Then srcX = m_Size.nLeft                           ' set default source X,Y values
    If srcY = 0& Then srcY = m_Size.nTop
    
    If (m_Flags And flag_Segmented) Then                            ' get coords relative to segmented tile
        Index = m_MultiImageAttr.FrameFromIndex()
        srcY = srcY + (Index \ m_Tiles.Columns) * m_Tiles.Height
        srcX = srcX + (Index Mod m_Tiles.Columns) * m_Tiles.Width
        If srcWidth = 0& Then srcWidth = m_Tiles.Width
        If srcHeight = 0& Then srcHeight = m_Tiles.Height
        If destWidth = 0& Then destWidth = m_Tiles.Width
        If destHeight = 0& Then destHeight = m_Tiles.Height
    Else
        If srcWidth = 0& Then srcWidth = m_Size.nWidth              ' get coords relative to entire image
        If srcHeight = 0& Then srcHeight = m_Size.nHeight
        If destWidth = 0& Then destWidth = m_Size.nWidth
        If destHeight = 0& Then destHeight = m_Size.nHeight
    End If
    
    If Angle < 0! Then
        actualAngle = 360! + (Int(Angle) Mod 360!) + (Angle - Int(Angle))
    Else
        actualAngle = (Int(Angle) Mod 360!) + (Angle - Int(Angle))
    End If
    
    ' handle mirroring options
    ' We need to use the srcWidth/srcHeight parameters for mirroring, so if
    '   destWidth or destHeight are negative, flip sign for both the dest_ & src_ parameters
    
    If destWidth < 0& Then destWidth = -destWidth: srcWidth = -srcWidth
    If destHeight < 0& Then destHeight = -destHeight: srcHeight = -srcHeight
    
    If actualAngle = 180! Then
        srcWidth = -srcWidth
        srcHeight = -srcHeight
        actualAngle = 0!
    ElseIf srcWidth < 0& And srcHeight < 0& And actualAngle > 0! Then
        If actualAngle < 180! Then actualAngle = actualAngle + 180! Else actualAngle = actualAngle - 180!
        srcWidth = -srcWidth
        srcHeight = -srcHeight
    End If
    
    If GDIplusEffectsHandle Then                                            ' calculate matrix info
        scalerX = destWidth / srcWidth                                      ' scaling ratios; includes mirroring
        scalerY = destHeight / srcHeight
    
        srcRect.nTop = srcY: srcRect.nLeft = srcX
        srcRect.nHeight = srcHeight: srcRect.nWidth = srcWidth
        
        If srcWidth < 0& Then                                          ' adjust if mirroring horizontally
            scalerX = -scalerX
            destX = destX + srcX * scalerX
            srcRect.nLeft = m_Size.nWidth - srcX - 1&
        Else
            destX = destX - srcX * scalerX
        End If
        If srcHeight < 0& Then                                         ' adjust if mirroring vertically
            scalerY = -scalerY
            destY = destY + srcY * scalerY
            srcRect.nTop = m_Size.nHeight - srcY - 1&
        Else
            destY = destY - srcY * scalerY
        End If
        ' note to self: by-hand xform proves using above calculations work
        ' no adjustment to DestX,DestY if no mirroring, rect W/H must be positive values
        ' summary of calculation, per source pixel
        ' For srcY = srcRect.nTop To srcRect.nTop+srcRect.nHeight - 1
        '   For srcX = srcRect.nLeft To srcRect.nLeft+srcRect.nWidth - 1
        '       pixel from srcX,srcY transforms to: DestX + (srcX * ScalerX), DestY + (srcY * ScalerY)
        ' Next: Next
    Else
        If srcWidth < 0& Then srcX = m_Size.nWidth - srcX - 1&            ' offset the srcX and/or srcY parameters
        If srcHeight < 0& Then srcY = m_Size.nHeight - srcY - 1&
    End If
    
    ' to handle multi-image formats that can break when rendered with attributes and/or mirroring, we create clones as needed
    If (m_Flags And flag_CanClone) Then                                     ' clonable depending on rendering attributes
        If GDIplusAttributesHandle Then
            bClone = True                                                   ' clone when rendering with attributes
        Else
            bClone = (srcWidth < 0& Or srcHeight < 0& Or actualAngle = 180!)  ' and/or when mirroring
        End If
        If bClone Then
            If (m_Flags And flag_Cloned) = 0& Then                          ' not cloned, needs to be
                tHandle = m_MultiImageAttr.CreateClone(m_Handle)
                If tHandle = 0& Then                                        ' if clone failure, render without mirroring/attributes
                    GDIplusAttributesHandle = 0&                            ' don't render with these
                    GDIplusEffectsHandle = 0&
                    srcHeight = Abs(srcHeight)                              ' ensure no mirroring
                    srcWidth = Abs(srcWidth)
                    If actualAngle = 180! Then actualAngle = 0!
                    tHandle = m_Handle                                      ' get our handle again
                Else
                    m_Flags = m_Flags Xor flag_Cloned                       ' update flag
                End If
            End If
        ElseIf (m_Flags And flag_Cloned) Then                               ' already cloned
            m_MultiImageAttr.KillClone                                      ' release memory used by clone
            m_Flags = m_Flags Xor flag_Cloned                               ' update flag
            tHandle = m_Handle
        End If
    End If
    
    If actualAngle = 0! Then
    
        If Interpolation = lvicAutoInterpolate Then                         ' set high precision drawing when scaling
            If ((Abs(srcWidth) = destWidth) And (Abs(srcHeight) = destHeight)) = True Then
                Interpolation = lvicNearestNeighbor
            Else
                Interpolation = lvicHighQualityBicubic
            End If
        End If
        Call GdipSetInterpolationMode(hGraphics, Interpolation + 2&)
        Call GdipSetPixelOffsetMode(hGraphics, SmoothingModeAntiAlias)
        
        If GDIplusEffectsHandle Then                                        ' prepare to draw with v1.1 effects
                                                                            ' create a scaling matrix for the destination area
            Call GdipCreateMatrix2(scalerX, 0&, 0&, scalerY, destX, destY, hMatrix)
            Render = (GdipDrawImageFX(hGraphics, tHandle, srcRect, hMatrix, GDIplusEffectsHandle, GDIplusAttributesHandle, UnitPixel) = 0&)
        Else
            
            Render = (GdipDrawImageRectRectI(hGraphics, tHandle, destX, destY, destWidth, destHeight, _
                                            srcX, srcY, srcWidth, srcHeight, UnitPixel, GDIplusAttributesHandle, 0&, 0&) = 0&)
        End If
    
    Else ' we are rotating
        
        If srcWidth < 0& Then
            ' 180 degree rotation bug noted below
            ' when rotating 180 degrees with just horizontal mirroring, rendering X,Y coords get whacked
            ' Note: Error does not surface if image is flipped with GdipImageRotateFlip
            ' So we will change rotation just a bit to allow it to render correctly
            If actualAngle > 179.96! And actualAngle < 180.04! Then
                If actualAngle < 180! Then rot180MirrAdj = -0.05! Else rot180MirrAdj = 0.05!
            End If
        End If
        
        If GdipRotateWorldTransform(hGraphics, actualAngle + rot180MirrAdj, 0&) = 0& Then ' transform 0,0 to center of rendered image
            GdipTranslateWorldTransform hGraphics, destX + (destWidth / 2), destY + (destHeight / 2), 1&
            
            If Interpolation = lvicAutoInterpolate Then Interpolation = lvicHighQualityBicubic
            Call GdipSetInterpolationMode(hGraphics, Interpolation + 2&)
            Call GdipSetPixelOffsetMode(hGraphics, SmoothingModeAntiAlias)
            If Not (actualAngle = 0! Or actualAngle = 180!) Then Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
        
            If GDIplusEffectsHandle Then                                        ' prepare todraw with v1.1 effects
                Call GdipCreateMatrix2(scalerX, 0&, 0&, scalerY, -destWidth \ 2, -destHeight \ 2, hMatrix)
                Render = (GdipDrawImageFX(hGraphics, tHandle, srcRect, hMatrix, GDIplusEffectsHandle, GDIplusAttributesHandle, UnitPixel) = 0&)
            Else
                Render = (GdipDrawImageRectRectI(hGraphics, tHandle, -destWidth \ 2, -destHeight \ 2, destWidth, destHeight, _
                                                srcX, srcY, srcWidth, srcHeight, UnitPixel, GDIplusAttributesHandle, 0&, 0&) = 0&)
            End If
        
        End If
    End If
    
    If GDIplusHGraphicsHandle = 0& Then                         ' clean up
        GdipDeleteGraphics hGraphics
    Else
        GdipRestoreGraphics hGraphics, svGraphics
    End If
    If hMatrix Then GdipDeleteMatrix hMatrix
    
End Function

Public Function RenderSkewed(ByVal hDC As Long, _
                ByVal TopLeftX As Single, ByVal TopLeftY As Single, _
                ByVal TopRightX As Single, ByVal TopRightY As Single, _
                ByVal BottomLeftX As Single, ByVal BottomLeftY As Single, _
                Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&, _
                Optional ByVal GDIplusAttributesHandle As Long = 0&, _
                Optional ByVal GDIplusHGraphicsHandle As Long = 0&, _
                Optional ByVal Interpolation As RenderInterpolation = lvicAutoInterpolate) As Boolean

    ' Similar to this class' Render function except it renders parallelograms only.
    ' A parallelogram is simply a rectangle/square skewed/stretched or not
    ' Routine does not support rotation nor v1.1 GDI+ effects
    '   -- if rotation and/or effects desired. Render to separate image with desired effects,
    '       then call this routine to skew that rendering

    ' Customizable rendering routine, depending on your level of experience with GDI+
    ' Additionally, since the Handle is exposed, you can use your own custom rendering routines
    ' But read comments at top of class regarding GIF/TIFF navigation bugs to ensure your don't mess them up
    
    ' Parmaters
    '   :: hDC is destination device context handle
    '       - Ignored when GDIplusHGraphicsHandle is non-zero else must be provided
    '   :: TopLeftX,TopLeftY are destination coords to begin rendering top left of source
    '   :: TopRightX,TopRightY are destination coords to begin rendering top right of source
    '   :: BottomLeftX,BottomLeftY are destination coords to begin rendering bottom left of source
    '       - The bottom right coords are never needed and calculated by GDI+
    
    '   :: srcX,srcY are coords to begin rendering from source image
    '       - Should not be negative (metafiles may be an exception)
    '   :: srcWidth,srcHeight are dimensions of source to be rendered
    '       - See Mirroring below
    '   :: GDIplusAttributesHandle contains a GDI+ Image Attributes handle to adjust rendered pixel values
    '       - GDIplusAttributesHandle is not destroyed nor modified by this routine
    '   :: GDIplusHGraphicsHandle if non-zero should already be transformed, translated, smoothed, etc
    '       - hDC parameter is ignored when GDIplusHGraphicsHandle is non-zero
    '       - GDIplusHGraphicsHandle is not destroyed nor modified by this routine
    '   :: Interpolation
    '       - When set to lvicAutoInterpolate, high quality bicubic always used
    ' Mirroring....
    ' When the sign of srcWidth parameter is negative, horizontal mirroring will occur
    ' When the sign of srcHeight parameter is negative, vertical mirroring will occur
    
    Dim tHandle As Long, Index As Long, svGraphics As Long
    Dim hGraphics As Long, bClone As Boolean
    Dim thePoints(0 To 2) As POINTF
    
    tHandle = Me.Handle
    If tHandle = 0& Then
        RenderSkewed = True
        Exit Function
    End If
    
    If GDIplusHGraphicsHandle = 0& Then                             ' no GDI+ graphics handle passed
        If hDC = 0& Then Exit Function
        If GdipCreateFromHDC(hDC, hGraphics) Then Exit Function
    Else                                                            ' use handle passed
        hGraphics = GDIplusHGraphicsHandle
        GdipSaveGraphics hGraphics, svGraphics
    End If
    
    If Interpolation < lvicAutoInterpolate Or Interpolation > lvicHighQualityBicubic Then Interpolation = lvicHighQualityBicubic
    
    If srcX = 0& Then srcX = m_Size.nLeft                           ' set default source X,Y values
    If srcY = 0& Then srcY = m_Size.nTop
    
    If (m_Flags And flag_Segmented) Then                            ' get coords relative to segmented tile
        Index = m_MultiImageAttr.FrameFromIndex()
        srcY = srcY + (Index \ m_Tiles.Columns) * m_Tiles.Height
        srcX = srcX + (Index Mod m_Tiles.Columns) * m_Tiles.Width
        If srcWidth = 0& Then srcWidth = m_Tiles.Width
        If srcHeight = 0& Then srcHeight = m_Tiles.Height
    Else
        If srcWidth = 0& Then srcWidth = m_Size.nWidth              ' get coords relative to entire image
        If srcHeight = 0& Then srcHeight = m_Size.nHeight
    End If
    If srcWidth < 0& Then srcX = m_Size.nWidth - srcX - 1&            ' offset the srcX and/or srcY parameters
    If srcHeight < 0& Then srcY = m_Size.nHeight - srcY - 1&
    If srcWidth < 0& And srcHeight < 0& Then
        srcX = srcX + 1: srcY = srcY + 1
    End If
    
    ' to handle multi-image formats that can break when rendered with attributes and/or mirroring, we create clones as needed
    If (m_Flags And flag_CanClone) Then                                     ' clonable depending on rendering attributes
        If GDIplusAttributesHandle Then
            bClone = True                                                   ' clone when rendering with attributes
        Else
            bClone = (srcWidth < 0& Or srcHeight < 0&)                      ' and/or when mirroring
        End If
        If bClone Then
            If (m_Flags And flag_Cloned) = 0& Then                          ' not cloned, needs to be
                tHandle = m_MultiImageAttr.CreateClone(m_Handle)
                If tHandle = 0& Then                                        ' if clone failure, render without mirroring/attributes
                    GDIplusAttributesHandle = 0&                            ' don't render with these
                    srcHeight = Abs(srcHeight)                              ' ensure no mirroring
                    srcWidth = Abs(srcWidth)
                    tHandle = m_Handle                                      ' get our handle again
                Else
                    m_Flags = m_Flags Xor flag_Cloned                       ' update flag
                End If
            End If
        ElseIf (m_Flags And flag_Cloned) Then                               ' already cloned
            m_MultiImageAttr.KillClone                                      ' release memory used by clone
            m_Flags = m_Flags Xor flag_Cloned                               ' update flag
            tHandle = m_Handle
        End If
    End If
    
    thePoints(0).X = TopLeftX: thePoints(0).Y = TopLeftY
    thePoints(1).X = TopRightX: thePoints(1).Y = TopRightY
    thePoints(2).X = BottomLeftX: thePoints(2).Y = BottomLeftY
    
    If Interpolation = lvicAutoInterpolate Then Interpolation = lvicHighQualityBicubic
    Call GdipSetInterpolationMode(hGraphics, Interpolation + 2&)
    Call GdipSetPixelOffsetMode(hGraphics, SmoothingModeAntiAlias)
    
    RenderSkewed = (GdipDrawImagePointsRect(hGraphics, tHandle, ByVal VarPtr(thePoints(0)), 3&, _
                    srcX, srcY, srcWidth, srcHeight, UnitPixel, GDIplusAttributesHandle, 0&, 0&) = 0&)
    
    If GDIplusHGraphicsHandle = 0& Then
        GdipDeleteGraphics hGraphics                                        ' clean up
    Else
        GdipRestoreGraphics hGraphics, svGraphics
    End If
    
End Function

Public Property Get Segmented() As Boolean
    ' returns if user created a segmented image from source
    Segmented = CBool(m_Flags And flag_Segmented)
End Property

Public Sub SegmentImage(ByVal Columns As Long, ByVal Rows As Long, Optional ByVal TileCount As Long, _
                        Optional ByVal FrameDuration As Long = 0&, Optional ByVal Sequence As Variant)

    ' Call to split an image into tiles. The tiles must all be same size, not all tiles must be used
    ' Tiles also do not have to be square
    ' When only some tiles are used, the used tiles must be consecutive from top left to bottom right
    ' Exception: If supplying the Sequence, then the sequence identifies which tiles are used/not used
    
    ' The image to be segmented cannot be a multi-image object. If this is needed, the desired frame
    '   of that multi-image format should be saved to another class & it can be segmented: SavePictureGDIplus
    ' So if this class' ImageCount property is > 1, this class will not segement the image
    
    ' A segmented image can be created from one of these formats
    '   an image consisting of a single row & multiple columns (horizontal strip)
    '   an image consisting of a single column, muliple rows (vertical strip)
    '   an image consisting of multiple rows & multiple columns (grid)
    ' If not all tiles will be used, provide the optional TileCount parameter. See Sequence below
    ' Optional FrameDuration parameter caches number of milliseconds between frames
    ' Optional Sequence parameter identifies which tiles are used and in what order
    '   Two formats are allowed in thie parameter. Sequence is 1-bound, 1st tile is 1
    '   1) comma-delimited list of tiles: 1,3,5,2,4,3,2
    '   2) a byte/integer/long array of tiles
    '   If the sequence is not provided, the sequence will be 1 to nr of tiles or TileCount(if applicable)
    
    ' To remove segmentation, call routine again and pass 0 for both Columns and Rows
    
    If Me.Handle = 0& Then Exit Sub
    
    If (m_Flags And flag_Segmented) = 0& Then
        If Me.ImageCount > 1& Then Exit Sub                 ' will not segment animated-potential images
    End If
    If (Columns < 1& Or Rows < 1&) Then                     ' removing segmentation?
        If (Columns = 0& And Rows = 0&) Then                ' only if both columns/rows are zero
            If (m_Flags And flag_Segmented) Then            ' else passing negative values & not accepted
                m_Flags = m_Flags Xor flag_Segmented
                Set m_MultiImageAttr = Nothing
                RaiseEvent ImageChanged(1&)
            End If
        End If
        Exit Sub
    End If
    
    Dim sSeq() As String, newSEQ() As Long
    Dim lCount As Long, bGotSequence As Boolean
    
    ' handle optional Sequence parameter
    If IsMissing(Sequence) = False Then
        On Error Resume Next
        If VarType(Sequence) = vbString Then                ' expecting comma-delimited string
            sSeq() = Split(CStr(Sequence), ",")
            bGotSequence = (UBound(sSeq) > LBound(sSeq))
            If bGotSequence Then TileCount = UBound(sSeq) + 1&
        ElseIf (VarType(Sequence) And vbArray) Then
            bGotSequence = (UBound(Sequence) > LBound(Sequence))
            If bGotSequence Then
                Select Case (VarType(Sequence) And Not vbArray)
                Case vbByte, vbInteger, vbLong
                    TileCount = (UBound(Sequence) - LBound(Sequence)) + 1&
                Case Else
                    bGotSequence = False                    ' not acceptable array type
                End Select
            End If
            If Err Then Err.Clear
            On Error GoTo 0
        End If
    End If
    
    If Not bGotSequence Then                                ' no valid sequence parameter passed
        lCount = Rows * Columns
        If TileCount < 1& Then
            TileCount = lCount
        ElseIf TileCount > lCount Then                      ' invalid TileCount provided; fix it up
            TileCount = lCount
        End If
    End If
        
    With m_Tiles
        .Columns = Columns
        .Rows = Rows
        .Width = m_Size.nWidth \ .Columns
        .Height = m_Size.nHeight \ .Rows
    End With
    
    Set m_MultiImageAttr = New cGDIpMultiImage
    m_MultiImageAttr.InitializeContent 0&, lvicPicTypeBitmap, TileCount
    
    ReDim newSEQ(0 To 1, 1 To TileCount)                    ' set default segment indexes sequentially
    If bGotSequence Then                                    ' user-defined sequence
        On Error Resume Next
        If VarType(Sequence) = vbString Then                ' comma-delimited string
            For lCount = 1& To TileCount
                newSEQ(0, lCount) = Val(sSeq(lCount)) - 1&  ' ensure values in range of 1-TileCount
                If Err Then                                 ' internally, range is 0 to TileCount-1
                    Err.Clear
                    newSEQ(0, lCount) = 0&
                ElseIf newSEQ(0, lCount) > TileCount - 1& Then
                    newSEQ(0, lCount) = TileCount - 1&
                ElseIf newSEQ(0, lCount) < 0& Then
                    newSEQ(0, lCount) = 0&
                End If
                newSEQ(1, lCount) = FrameDuration
            Next
        Else
            Columns = 1&                                    ' sequence in numeric array type
            For lCount = LBound(Sequence) To UBound(Sequence)
                newSEQ(0, Columns) = CLng(Sequence(lCount)) - 1&
                If Err Then
                    Err.Clear
                    newSEQ(0, Columns) = 0&
                ElseIf newSEQ(0, Columns) > TileCount - 1& Then
                    newSEQ(0, Columns) = TileCount - 1&
                ElseIf newSEQ(0, Columns) < 0& Then
                    newSEQ(0, Columns) = 0&
                End If
                newSEQ(1, Columns) = FrameDuration
                Columns = Columns + 1&
            Next
        End If
        On Error GoTo 0
    Else
        For lCount = 1& To TileCount
            newSEQ(0, lCount) = lCount - 1&
            newSEQ(1, lCount) = FrameDuration
        Next
    End If
    m_MultiImageAttr.SetSequenceDuration newSEQ()
    
    m_Flags = m_Flags Or flag_NoEvent Or flag_Segmented
    m_MultiImageAttr.SetIndex m_Source, m_Handle, &HF0000000
    m_Flags = (m_Flags And Not flag_NoEvent)
    RaiseEvent ImageChanged(1&)

End Sub

Public Property Get Width(Optional Index As Long) As Long
    ' Returns width of current image
    If (m_Flags And flag_Segmented) Then
        Width = m_Tiles.Width
    Else
        Width = m_Size.nWidth
        If Index > 0& Then
            If Not m_MultiImageAttr Is Nothing Then
                If m_MultiImageAttr.Index <> Index Then
                    Width = m_MultiImageAttr.GetWidth(m_Source, m_Handle, Index)
                End If
            End If
        End If
    End If
End Property

Public Property Get AsyncDownloadStatus() As AsyncDownloadStatusEnum
    ' Returns status of an async download, if applicable
    If m_AsyncClient Is Nothing Then
        ' async download is either complete or doesn't apply
        If Me.Handle Then
            AsyncDownloadStatus = lvicAsync_Ok      ' we have an image assigned
        Else
            AsyncDownloadStatus = lvicAsync_Invalid ' no image is assigned (download failed if started)
        End If
    Else
        AsyncDownloadStatus = m_AsyncClient.State   ' currently downloading or queued (some other value)
    End If
    
End Property

Public Property Get AsyncDownloadURL() As String
    ' returns the URL used for an async download attempt
    ' Note: This property is only valid until the download succeeds or fails
    If Not m_AsyncClient Is Nothing Then AsyncDownloadURL = m_AsyncClient.URL
End Property

Private Sub pvClear()   ' clean-up routine

    On Error Resume Next

    If Not m_AsyncClient Is Nothing Then
        m_AsyncClient.Abort
        Set m_AsyncClient = Nothing
    End If
    If m_Handle Then
        GdipDisposeImage m_Handle
        m_Handle = 0&
    End If
    Set m_MultiImageAttr = Nothing
    m_Flags = 0&
    m_Source = Empty
    m_Size.nHeight = 0!
    m_Size.nWidth = 0!
    m_Format = lvicPicTypeNone
End Sub

Private Sub pvGetImageData()

    If g_NewImageData Is Nothing Then Exit Sub              ' no image was loaded by LoadPictureGDIplus
    
    Dim bCacheFormat As Boolean

    Set m_MultiImageAttr = g_NewImageData                   ' reference public class
                                                            ' extract image information from that class
    m_MultiImageAttr.CacheSourceInfo m_Source, m_Handle, m_Format, bCacheFormat, True
    Set g_NewImageData = Nothing                            ' release public class
    
    If m_Format > lvicPicTypeNone Then
        Select Case m_Format                                    ' special handling...
            Case lvicPicTypeTIFF                                ' test for multiple images
                m_MultiImageAttr.InitializeContent m_Handle, m_Format, 0&
                If m_MultiImageAttr.Frames > 1& Then            ' has more than one image
                    m_Flags = m_Flags Or flag_CanClone          ' flag for rendering purposes
                Else
                    Set m_MultiImageAttr = Nothing              ' no longer needed
                End If                                          ' get image dimensions
                Call GdipGetImageBounds(Me.Handle, m_Size, UnitPixel)
                
            Case lvicPicTypeCursor, lvicPicTypeIcon, lvicPicTypeAnimatedCursor, lvicPicTypeFromBinaries
                Call GdipGetImageBounds(Me.Handle, m_Size, UnitPixel)
                
            Case lvicPicTypeGIF, lvicPicTypePNG
                If m_MultiImageAttr.Frames < 2& Then Set m_MultiImageAttr = Nothing
                Call GdipGetImageBounds(Me.Handle, m_Size, UnitPixel)
            
            Case lvicPicTypeAVI
                Call GdipGetImageBounds(Me.Handle, m_Size, UnitPixel)
            
            Case Else
                Set m_MultiImageAttr = Nothing                  ' all others. Release, no-longer needed
                If m_Format = lvicPicTypeAsyncDL Then
                    Set m_AsyncClient = m_Source
                    m_Source = Empty
                    m_Format = lvicPicTypeNone
                    If bCacheFormat Then m_Flags = m_Flags Or flag_CacheAsyncData
                    ' in order to allow the LoadPictureGDIplus function to return with this class,
                    '   a delay in starting async download is performed. Otherwise, if error occurs
                    '   during activation, the user will not receive the RaisedEvent because at this
                    '   point, LoadPictureGDIplus has not exited yet
                    SetTimer g_AsyncController.DelayWindow, ObjPtr(m_AsyncClient), 50, AddressOf modCommon.CommonTimerProc
                    Exit Sub
                Else
                    Call GdipGetImageBounds(Me.Handle, m_Size, UnitPixel) ' get image dimensions
                End If
        End Select
        Me.KeepOriginalFormat = bCacheFormat                    ' create stand-alone image if needed
    Else
        Set m_MultiImageAttr = Nothing
    End If

End Sub

Private Sub Class_Initialize()
    
    ' class retrieves processed image from modCommon's modCommon.LoadImage routine
    
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
    If m_Token.Token = 0& Then Set g_NewImageData = Nothing
    
    Call pvGetImageData
    
End Sub

Private Sub m_AsyncClient_DownloadComplete(DataPointer As Long, ErrorCode As Long, Rcvd As Boolean)
    Rcvd = True
    If DataPointer = 0& Then
        RaiseEvent AsyncDownloadDone(False, (ErrorCode))
    ElseIf DataPointer = -1& Then
        Call pvGetImageData
        RaiseEvent AsyncDownloadDone(True, 0&)
    Else
        Dim bDummy() As Byte
        CopyMemory ByVal VarPtrArray(bDummy), ByVal DataPointer, 4&
        CopyMemory ByVal DataPointer, 0&, 4&
        If modCommon.LoadImage(bDummy(), (m_Flags And flag_CacheAsyncData), , , Me) Is Nothing Then
            RaiseEvent AsyncDownloadDone(False, (lvicAsync_NoImage))
        Else
            Call pvGetImageData
            RaiseEvent AsyncDownloadDone(True, 0&)
        End If
    End If
    Set m_AsyncClient = Nothing
    m_Flags = m_Flags And Not flag_CacheAsyncData
    RaiseEvent ImageChanged(0&)
End Sub

Private Sub m_MultiImageAttr_FrameChanged(Index As Long, Width As Single, Height As Single, UpdateDimensions As Boolean)

    ' Whenever the Index in the multi-image class changes, this event is fired
    
    If UpdateDimensions Then                                ' if not segmented, cache width/height of image
        If (m_Flags And flag_Segmented) = 0& Then
            m_Size.nWidth = Width: m_Size.nHeight = Height
        End If
    End If
    
    If (m_Flags And flag_NoEvent) Then                      ' this class changed index for its own reason
        m_Flags = m_Flags Xor flag_NoEvent                  ' don't fire event back to any client
    Else
        m_Flags = (m_Flags And Not flag_Cloned)
        RaiseEvent ImageChanged(Index)                      ' inform any client that image has changed
    End If
End Sub

Public Function AlphaMask(maskArray() As Byte, ByVal SetMask As Boolean) As Boolean

    ' Routine returns or sets an Alpha mask
    ' Masks can be used to change how an image is blended.
    
    ' The class' image, if applying a mask, must be 32 bits per pixel
    
    ' If returning a mask...
    '   pass SetMask as False
    '   maskArray() will be zero LBound and a 1D array.
    '       it will contain 1 byte for each pixel in the passed picture and is not word-aligned
    '       Example: If image was 20x10, array would contain 200 elements
    ' If setting a mask
    '   pass SetMask as True
    '   maskArray() must contain at least 1 byte per pixel of the passed picture
    '       maskArray can be any dimensions and any LBound
    
    ' Real world example of masking.
    ' Let's say you wanted to draw a rectangular image so that it fades from the center out toward the edges.
    ' Well, you'd create a routine that would calculate alpha values for each pixel in the image. That
    '   routine would calculate the alpha values for the center mass as fully opaque or values of 255.
    '   The routine then would work its way around the center mass until it finishes at the outer edges,
    '   reducing the alpha values for the remaining pixels. What you would end up with is an array of
    '   bytes (1 per pixel) that contains values ranging from 0 to 255. You would then pass that byte
    '   array to this function and it would apply the array as an Alpha mask to the passed picture.
    '   Redraw your picture and presto!
    ' Another example could be using the mask as a simple shadow effect
    
    ' Retrieving a mask most likely would only be useful to cache the original mask and/or to
    '   modify it and eventually passing it back to this function to apply it
    
    ' Example usage for simple shadow:
    '  #    Dim tImg As GDIpImage, tMask() As Byte
    '  #    Dim ss As SAVESTRUCT
    '  #
    '  #    If AlphaImgCtl1.AlphaMask(tMask(), False) Then ' get mask of source image
    '  #
    '  #        ss.Width = AlphaImgCtl1.Picture.Width    ' create new blank image
    '  #        ss.Height = AlphaImgCtl1.Picture.Height  ' make it same size as original & all gray
    '  #        ss.ColorDepth = lvicConvert_TrueColor32bpp_ARGB
    '  #        ss.RSS.FillColorARGB = ConvertRGBtoARGB(RGB(96,96,96))
    '  #        ss.RSS.FillColorUsed = True
    '  #        Set tImg = New GDIpImage
    '  #        If SavePictureGDIplus(Nothing, tImg, , ss) Then ' image created
    '  #            If tImg.AlphaMask(tMask(), True) Then ' now apply the mask to the new image
    '  #                AlphaImgCtl2.TransparencyPct = 60
    '  #                Set AlphaImgCtl2.Picture = tImg ' set image to control
    '  #            End If                              ' the control would be behind source control
    '  #        End If                                  ' and shifted a bit to imitate a shadow
    '  #    End If
    
    If SetMask Then
        If Me.BitDepth() < 32& Then Exit Function
    Else
        If Me.Handle = 0& Then Exit Function
    End If

    Dim tData() As Byte, tMask() As Byte
    Dim tBMP As BitmapData, tSize As RECTI
    Dim tSA As SafeArray, sSA As SafeArray
    Dim X As Long, Y As Long, z As Long, d As Long
    
    tSize.nHeight = m_Size.nHeight                      ' needed for GdipBitmapLockBits
    tSize.nWidth = m_Size.nWidth                        ' disallow null images
    If tSize.nHeight < 1& Or tSize.nWidth < 1& Then Exit Function
    
    If SetMask Then
    
        ' see if passed array has enough bytes for the passed picture's mask
        X = modCommon.NormalizeArray(maskArray(), tMask(), True, z)
        If (X < 1&) Or (X < tSize.nWidth * tSize.nHeight) Then Exit Function
        ' attempt to lock the image so we can write the mask
        If GdipBitmapLockBits(Me.Handle, tSize, ImageLockModeRead Or ImageLockModeWrite, lvicColor32bppAlpha, tBMP) Then Exit Function
        With sSA                                        ' overlay our array on passed mask array
            .cbElements = 1                             ' this allows us to use a 1D, 0-LBound array
            .cDims = 1                                  ' regardless of number of dimensions and bounds
            .pvData = z                                 ' of the passed array. And all without copying anything
            .rgSABound(0).cElements = X
        End With
        CopyMemory ByVal VarPtrArray(tMask), VarPtr(sSA), 4&
        With tSA                                        ' overlay an array on the picture data
            .cbElements = 1                             ' handle upside down images
            .cDims = 1
            .pvData = tBMP.Scan0Ptr
            If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
            .rgSABound(0).cElements = Abs(tBMP.stride) * tBMP.Height
        End With
        CopyMemory ByVal VarPtrArray(tData), VarPtr(tSA), 4&
        
        On Error GoTo ExitRoutineSet
        If tBMP.stride < 0& Then                        ' if upside down then
            X = UBound(tData): Y = 3&: z = -4&          ' use these loop parameters
        Else
            Y = UBound(tData): X = 3&: z = 4&           ' if right-side up then
        End If                                          ' use these loop parameters
        For X = X To Y Step z
            tData(X) = tMask(d)                         ' apply the mask
            d = d + 1&
        Next
        
    Else                                ' retrieving a mask vs. setting one
        
        ' attempt to lock the image so we can read the alpha data
        If GdipBitmapLockBits(Me.Handle, tSize, ImageLockModeRead, lvicColor32bppAlphaMultiplied, tBMP) Then Exit Function
        With tSA                                        ' overlay an array on the picture data
            .cbElements = 1                             ' handle upside down images
            .cDims = 1
            .pvData = tBMP.Scan0Ptr
            If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
            .rgSABound(0).cElements = Abs(tBMP.stride) * tBMP.Height
        End With
        CopyMemory ByVal VarPtrArray(tData), VarPtr(tSA), 4&
        
        On Error GoTo ExitRoutineGet
        ReDim maskArray(0 To tBMP.Width * tBMP.Height - 1&)
        
        If tBMP.stride < 0& Then                        ' if upside down then
            X = UBound(tData): Y = 3&: z = -4&          ' use these loop parameters
        Else
            Y = UBound(tData): X = 3&: z = 4&           ' if right-side up then
        End If                                          ' use these loop parameters
        For X = X To Y Step z
            maskArray(d) = tData(X)                     ' get the mask
            d = d + 1&
        Next
    
    End If
    AlphaMask = True
    
ExitRoutineSet:                                         ' ensure overlay removed
    If sSA.pvData Then CopyMemory ByVal VarPtrArray(tMask), 0&, 4&
ExitRoutineGet:                                         ' ensure overlay removed
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(tData), 0&, 4&
    GdipBitmapUnlockBits Me.Handle, tBMP           ' ensure image is unlocked

End Function

Private Sub Class_Terminate()
    Call pvClear
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub

