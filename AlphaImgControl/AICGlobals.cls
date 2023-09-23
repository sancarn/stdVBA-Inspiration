VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AICGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'  -----======== PURPOSE: Provide global declarations & functions to users ========-----
' Class Instancing = GlobalMultiUse
' Note: You do not create/instantiate an instance of this class. The methods/properties in this
' class are automatically available to you whenever an AlphaImgCtl is placed onto your form
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' Place enumerations, constants, structures & functions/routines in this class that
'   should be made available to the the user and/or to any of projects classes.
' Class used to prevent 'Circular References' errors resulting from WithEvents usage
'   and provide global objects to the user that can be used without creating class instances

Public Type RENDERSTYLESTRUCT2          ' used for SavePictureGDIplus, PaintPictureGDIplus, MergePictureGDIplus
    Angle As Single                     ' Value between -360 and 360 degrees
    Effects As GDIpEffects              ' class used to create image attributes and/or v1.1 Effects
    EffectType As EffectsEnum           ' which created effect of the Effects classs to be applied to image
    Mirrored As MirroredEnum            ' MirroredEnum value
    Quality As RenderInterpolation      ' optional quality setting when rendering
    FillColorARGB As Long               ' ARGB fill color (use ConvertRGBtoARGB). Applied only if FillColorUsed=True
    FillBrushGDIplus_Handle As Long     ' GDI+ brush handle (You are responsible for destroying it). Overrides FillColorARGB
    FillColorUsed As Boolean            ' Determines if FillColorARGB will be applied
End Type

Public Type SAVESTRUCT                  ' used for SavePictureGDIplus, PaintPictureGDIplus, MergePictureGDIplus
    AlphaTolerancePct As Long           ' Determines percentage of alpha to be rendered transparent
                                        ' ^^ If transparency > percentage, pixel becomes 100% opaque else 100% transparent
                                        ' ^^ applies when palettizing images or using color reduction
                                        ' ^^ value of zero means image is not processed to change transparency levels
    ColorDepth As ColorDepthEnum        ' Color reduction/depth as appropriate (See SavePictureGDIplus for more details)
    PaletteType As PaletteTypeEnum      ' Used if ColorDepth is 2,16,256 colors
    Palette_Handle As Long              ' Used if PaletteType is lvicPaletteUserDefined (You are responsible for destroying it)
    RSS As RENDERSTYLESTRUCT2           ' optional rendering style
    CompressionJPGQuality As CompressionQualityEnum ' options for saving to JPG,TIF,PCX,TGA (ignored for all others)
    CursorHotSpotX As Integer           ' if saving as cursor, optional X,Y hotspot
    CursorHotSpotY As Integer           ' ^^ will not be allowed to exist outside image dimensions
    Width As Integer                    ' Destination width, default is image width
    Height As Integer                   ' Destination height, default is image height
    RotatingCanGrowImage As Boolean     ' If true, bounds are expanded to render image else image is contained within bounds
    ExtractCurrentFrameOnly As Boolean  ' Forces multi-frame images to only save the current frame (extraction)
    reserved1 As Long                   ' internal use or possibly for expansion
    reserved2 As Long                   ' internal use or possibly for expansion
End Type

Public Type TIFFPAGESTRUCT              ' Used for creating multiple page TIFFs (1 structure per page)
    Picture As GDIpImage                ' Instance of a GDIpImage
    SS As SAVESTRUCT                    ' optional rendering style
End Type
Public Type MULTIPAGETIFFSTRUCT         ' Used for creating multiple page TIFFs
    Pages As Long                       ' number of pages, number of TIFFPAGESTRUCT structures
    TPS() As TIFFPAGESTRUCT             ' picture source & optional rendering styles for the page
End Type

' //// new MULTIIMAGESAVESTRUCT and its embedded members \\\\ now allows creation of animated PNG (APNG)
Public Type GIFFRAMESTRUCT              ' Used for creating animated GIFs & PNGs (1 structure per frame)
    DelayTime As Long                   ' delay in ms
    DisposalCode As GIFDisposalCodes    ' what happens to a frame after it is displayed
    XOffset As Integer                  ' optional offset: use -1 to auto-center on WindowWidth
    YOffset As Integer                  ' optional offset: use -1 to auto-center on WindowHeight
End Type
Public Type GIFOVERVIEWSTRUCT           ' Used for creating animated GIFs/PNGs (describes overall GIF/PNG)
    LoopCount As Long                   ' Number of loops to complete before animation stops: 0=infinite
    WindowWidth As Integer              ' Overall width of GIF/PNG. Leave zero for auto-calculation
    WindowHeight As Integer             ' Overall height of GIF/PNG. Leave zero for auto-calculation
End Type
Public Type IMAGEINFOSTRUCT             ' Used to describe individual images in a multi-image format (ico,cur,gif,tiff,png)
    Picture As GDIpImage                ' Image to be saved
    FrameNumber As Long                 ' If image contains multiple images, the frame/image index. Zero=current index
    GroupNumber As Long                 ' Animated cursors within Win7+ may be grouped. Else leave zero
    SS As SAVESTRUCT                    ' Rendering options for this image, including non-default image width/height settings
    GIFFrameInfo As GIFFRAMESTRUCT      ' If saving to animated GIF, frame description
    IconEmbeddedAsPNG As Boolean        ' If saving to icon/cursor, save this image in PNG format (Vista+)
End Type                                ' Note: if saving to ico/cur, SS contains palette & cursor hotspot info

Public Type MULTIIMAGESAVESTRUCT        ' Used for saving 1 or more images to icon, cursor, GIF, PNG, TIFF
    Images As Long                      ' Number of Image() members
    Image() As IMAGEINFOSTRUCT          ' One for each image to be added to icon, cursor, GIF, PNG, TIFF resource
    GIFOverview As GIFOVERVIEWSTRUCT    ' If saving to animated GIF/PNG, the overall GIF/PNG description
End Type
' ///////////////////////////////// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Type MERGEINFOSTRUCT             ' Used for MergePictureGDIplus
    Picture As GDIpImage                ' The picture to use
    RSS As RENDERSTYLESTRUCT2           ' Optional rendering styles
    Left As Integer                     ' Left rendering coordinate
    Top As Integer                      ' Top rendering coordinate
    Width As Integer                    ' Destination width, default is image width
    Height As Integer                   ' Destination height, default is image height
End Type
Public Type PICTUREMERGESTRUCT          ' Used for MergePictureGDIplus
    CanvasWidth As Long                 ' The overall new picture width
    CanvasHeight As Long                ' The overall new picture height
    Pictures As Long                    ' Number of pictures to merge
    MIS() As MERGEINFOSTRUCT      ' 1 structure fore each picutre
End Type

Public Type REDEYECORRECTIONAREA        ' a simple GDI RECT structure
    Left As Long                        ' used with EffectsEnum.lvicRedEyeCorrectionFX
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type ASSOCIATEDICON
    FileName As String                  ' can be a numeric value indicating a system imagelist icon index
    DesiredSize As SHILIconSizeEnum     ' if IconType includes lvicAssocIconIndex then the size indicates which system imagelist to use
    IconType As AssocIconTypeEnum       ' determines what type of icon will be retrieved
    IndexReturned As Long               ' on return, system image list icon index used to create icon
End Type                                ' (IndexReturned And &HFFFFFF)=icon index, (IndexReturned \ &H10000000)= icon size

Public Enum SaveAsFormatEnum            ' used for SavePictureGDIplus (See function for more details)
    lvicSaveAsCurrentFormat = 0         ' save in original format
    lvicSaveAsBitmap = 1                ' to bitmap
    lvicSaveAsMetafile = 2              ' to WMF
    lvicSaveAsIcon = 3                  ' to Icon
    lvicSaveAsEMetafile = 4             ' to EMF
    lvicSaveAsCursor = 5                ' to Cursor
    lvicSaveAsJPEG = 6                  ' to JPEG. SavePictureGDIplus Compression parameter used. Default=80%
    lvicSaveAsPNG = 7                   ' to PNG
    lvicSaveAsTGA = 8                   ' to TGA. SavePictureGDIplus Compression parameter used
    lvicSaveAsGIF = 9                   ' to GIF
    lvicSaveAsTIFF = 10                 ' to TIFF. Can save multiple images as multipage TIFF. Compression parameter used
    lvicSaveAsPCX = 11                  ' to PCX.  SavePictureGDIplus Compression parameter used
    lvicSaveAsMetafile_NonPlaceable = 12 ' to non-placeable WMF
    lvicSaveAs_HBITMAP = 13             ' to Bitmap handle. You are responsible for destroying handle
    lvicSaveAs_HICON = 14               ' to Icon handle. You are responsible for destroying handle
    lvicSaveAs_HCURSOR = 15             ' to Cursor handle. You are responsible for destroying handle
    lvicSaveAsPNM = 16                  ' to PNM (depends on color depth): PBM (1bpp), PGM (grayscale) or PPM (true color)
    lvicSaveAsPAM = 17                  ' to PAM which is newer version of PNM that supports transparency
End Enum

Public Enum PaletteTypeEnum             ' (See SavePictureGDIplus for more details)
    lvicPaletteDefault = 0              ' 8bpp=Halftone; 4bpp=VGA; 1bpp=B&W
    lvicPaletteAdaptive = 1             ' creates palette from actual image colors. Color count defined by ColorDepthEnum
    lvicPaletteUserDefined = 2          ' you provide a GDI palette handle & you destroy it too. Defines color count
End Enum

Public Enum ColorDepthEnum              ' (See SavePictureGDIplus for more details)
    lvicNoColorReduction = 0            ' may promote to lvicConvert_TrueColorAlpha if optional rendering style creates transparency
    lvicDefaultReduction = 1            ' will reduce bit depth if no colors need to be modified
    lvicConvert_BlackWhite = 2          ' Image is grayscaled and converted to B&W when default palette type is used
    lvicConvert_16Colors = 3            ' forces to 16 color image. Use lvicPaletteAdaptive or lvicPaletteUserDefined
    lvicConvert_256Colors = 4           ' forces to 256 color image. Use any PaletteTypeEnum
    lvicConvert_TrueColor24bpp = 5      ' forces to 24bpp. Complete loss of transparency
    lvicConvert_TrueColor32bpp_RGB = 6  ' forces to 32bpp RGB (no transparency supported)
    lvicConvert_TrueColor32bpp_ARGB = 7 ' forces to 32bpp ARGB (alpha channel used)
    lvicConvert_TrueColor32bpp_pARGB = 8 ' forces to 32bpp pARGB. Bitmap & TGA supported only. Else defaults to 32bpp ARGB
    lvicApplyAlphaTolerance = 16        ' applies EFFECTSSTYLESTRUCT.AlphaTolerancePct before any effects occur
    ' note about lvicApplyAlphaTolerance. It is OR'd with one of the other ColorDepthEnum values
    ' ... If no transparency exists, then not applied
    ' ... Else EFFECTSSTYLESTRUCT.AlphaTolerancePct is applied to the source image. The result is used from then on
    
End Enum

Public Enum GrayScaleRatioEnum
    lvicNoGrayScale = 0&                ' no grayscaling
    lvicNTSCPAL = 1&                    ' R=R*.299, G=G*.587, B=B*.114
    lvicCCIR709 = 2&                    ' R=R*.213, G=G*.715, B=B*.072
    lvicSimpleAverage = 3&              ' R,G,and B = (R+G+B)/3
    lvicRedMask = 4&                    ' uses only the Red sample value: RGB = Red / 3
    lvicGreenMask = 5&                  ' uses only the Green sample value: RGB = Green / 3
    lvicBlueMask = 6&                   ' uses only the Blue sample value: RGB = Blue / 3
    lvicRedGreenMask = 7&               ' uses Red & Green sample value: RGB = (Red+Green) / 2
    lvicBlueGreenMask = 8&              ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
    lvicSepia = 9&                      ' mixture of RGB to create a brown-scale effect
    lvicBlackWhite = 10&
End Enum

Public Enum ImageColorFormatEnum
    lvicColor1bpp = &H30101      'Specifies that the format is 1 bit per pixel, indexed.
    lvicColor4bpp = &H30402      'Specifies that the format is 4 bits per pixel, indexed.
    lvicColor8bpp = &H30803      'Specifies that the format is 8 bits per pixel, indexed.
    lvicColor16bppGrayscale = &H101004 'Specifies that the format is 16 bits per pixel, grayscale.
    lvicColor16bpp555 = &H21005  'Specifies that the format is 16 bits per pixel; 5 bits each are used for the red, green, and blue components. The remaining bit is not used.
    lvicColor16bpp565 = &H21006  'Specifies that the format is 16 bits per pixel; 5 bits are used for the red component, 6 bits are used for the green component, and 5 bits are used for the blue component.
    lvicColor16bpp555Alpha = &H61007  'Specifies that the format is 16 bits per pixel; 1 bit is used for the alpha component, and 5 bits each are used for the red, green, and blue components.
    lvicColor24bpp = &H21808     'Specifies that the format is 24 bits per pixel; 8 bits each are used for the red, green, and blue components.
    lvicColor32bpp = &H22009     'Specifies that the format is 32 bits per pixel; 8 bits each are used for the red, green, and blue components. The remaining 8 bits are not used.
    lvicColor32bppAlpha = &H26200A 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components.
    lvicColor32bppAlphaMultiplied = &HE200B     'Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.
    lvicColor48bpp = &H10300C    'Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.
    lvicColor64bppAlpha = &H34400D 'Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components.
    lvicColor64bppAlphaMultiplied = &H1C400E 'Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.
    ' ^^ fyi: pixel format constants are encoded like so
    '   1st byte: sequential indexing of 1-14
    '   2nd byte: bit count, 1,4,8,16,24,32,48,64
    '   3rd byte: attributes containing 1 or more of these flags
    '       1=paletted, 2=RGB?, 4=AlphaUsed, 8=AlphaPremultiplied, 16=2-byte colors, 32=??
End Enum

Public Enum ImageFormatEnum
    lvicPicTypeUnknown = -1&
    lvicPicTypeNone = 0             ' 1st 5 are same as VB's PictureTypeConstants
    lvicPicTypeBitmap = 1
    lvicPicTypeMetafile = 2         ' WMF
    lvicPicTypeIcon = 3             ' if ImageCount>1 then has multiple icons in resource
    lvicPicTypeEMetafile = 4        ' EMF
    lvicPicTypeCursor = 5           ' if ImageCount>1 then has multiple cursors in resource
    lvicPicTypeJPEG = 6
    lvicPicTypePNG = 7              ' if ImageCount>1 then animated PNG (APNG)
    lvicPicTypeTGA = 8
    lvicPicTypeGIF = 9              ' if ImageCount>1 then animated GIF
    lvicPicTypeTIFF = 10            ' if ImageCount>1 then has multiple pages in resource
    lvicPicTypePCX = 11
    lvicPicTypeAnimatedCursor = 12
    lvicPicTypeAVI = 13             ' if ImageCount>1 then can be animated
    lvicPicTypePNM = 14             ' PBM - 1bpp, PGM 8bpp, or PPM 24bpp (bit depth identifies which)
    lvicPicTypePAM = 15             ' PAM format is a newer version of PNM
End Enum

Public Enum AnimationActionEnum
    lvicAniCmdStop = 0              ' stops any current animation, resets frame to 1st index
    lvicAniCmdStart = 1             ' starts animation and raises the AnimationInitialize event
    lvicAniCmdPause = 2             ' pauses animation on current frame
    lvicAniCmdResume = 3            ' resumes paused animation or if not paused, same as lvicAniCmdStart
    lvicAniCmdGetState = 4          ' returns animation state, either: lvicAniCmdStop, lvicAniCmdStart, lvicAniCmdPause
    lvicAniCmdGetMinDuration = 5    ' returns minimal duration for any frame
    lvicAniCmdGetFrameIndex = 6     ' returns current frame index
    lvicAniCmdGetLoopCount = 7      ' returns loop count set during AnimationInitialize event
    lvicAniCmdSetMinDuration = 8    ' sets/changes the minimal duration for any frame
    lvicAniCmdGetMaxDuration = 9    ' returns maximum duration for any frame
    lvicAniCmdSetMaxDuration = 10   ' sets/changes the maximum duration for any frame
    lvicAniCmdSetLoopCount = 11     ' sets/changes the animation loops
End Enum

Public Enum MirroredEnum
    lvicMirrorNone = 0              ' no mirroring
    lvicMirrorHorizontal = 1        ' mirror on horizontal axis
    lvicMirrorVertical = 2          ' mirror on vertical axis
    lvicMirrorBoth = 3              ' mirror both horizontally & vertically
End Enum

Public Enum EffectsEnum             ' used for the GDIpEffects class. v1.1 of GDI+
    lvicNoEffects = 0
    lvicBlurFX = 1
    lvicSharpenFX = 2
    lvicTintFX = 3
    lvicBrightnessContrastFX = 4
    lvicHueSaturationLightnessFX = 5
    lvicColorBalanceFX = 6
    lvicToneLevelsFX = 7
    lvicColorCurveFX = 8
    lvicRedEyeCorrectionFX = 9
End Enum
Public Enum CurveAdjustmentsEnum    ' used when creating a ColorCurve effect. v1.1 of GDI+
    lvicAdjustExposure = 0
    lvicAdjustDensity = 1
    lvicAdjustContrast = 2
    lvicAdjustHighlight = 3
    lvicAdjustShadow = 4
    lvicAdjustMidtone = 5
    lvicAdjustWhiteSaturation = 6
    lvicAdjustBlackSaturation = 7
End Enum
Public Enum CurveChannel            ' used when creating a ColorCurve effect. v1.1 of GDI+
    lvicCurveChannelAll = 0
    lvicCurveChannelRed = 1
    lvicCurveChannelGreen = 2
    lvicCurveChannelBlue = 3
End Enum

Public Enum TileOrderEnum           ' used by TilePictureGDIplus
    lvicTile_NoFlip = 0             ' each tile is rendered normally
    lvicTile_FlipX = 1              ' every other tile is horizontally mirrored
    lvicTile_FlipY = 2              ' every other tile is vertically mirrored
    lvicTile_FlipXY = 3             ' every other tile is mirroed both horizontally/vertically
End Enum

Public Enum AsyncDownloadStatusEnum
    lvicAsync_WaitingOnReply = 0    ' sent request, but no answer yet (may time out, may not)
    lvicAsync_Queued = 5000         ' download is queued, max downloads are taking place
    lvicAsync_Downloading = 5001    ' download is in progress, not finished nor aborted
    lvicAsync_Invalid = 5002        ' no async download taking place, no picture assigned
    lvicAsync_NoImage = 5003        ' downloaded information did not result in image creation
    lvicAsync_Processing = 5004     ' download was successful & image is being processed
    lvicAsync_UserAborted = 5005    ' user aborted download
    lvicAsync_Ok = 200              ' http status code of 200 (OK) & image was created OK
    ' else error code per http standards: http://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html
End Enum

Public Enum CompressionQualityEnum  ' ignored for any formats other than JPG,PCX,TGA,TIF. See SAVESTRUCT
    ' values +1 to +100 are user-defined for JPG
    lvicDefaultCompressionQuality = 0
    lvicFormatUncompressed = -1
    lvicFormatCompressed = -2       ' default for TIF, PCX, TGA, PNM
    lvicJPGQuality10pct = -3
    lvicJPGQuality20pct = -4
    lvicJPGQuality30pct = -5
    lvicJPGQuality40pct = -6
    lvicJPGQuality50pct = -7
    lvicJPGQuality60pct = -8
    lvicJPGQuality70pct = -9
    lvicJPGQuality80pct = -10       ' default for JPG
    lvicJPGQuality90pct = -11
    lvicJPGQuality100pct = -12
    ' if wanting a JPG quality of any other value; supply it as a value btwn 1 & 100 in SAVESTRUCT.CompressionJPGQuality
End Enum

Public Enum RenderInterpolation     ' optional parameter to GDIpImage class' Render method
    lvicAutoInterpolate = 0         ' defaults to lvicHighQualityBicubic if scaling or rotating
    lvicBilinear = 1                ' better quality scaling than lvicNearestNeighbor
    ' ^^ This mode is not suitable for shrinking an image below 50 percent of its original size
    lvicBicubic = 2                 ' better quality scaling than lvicBilinear
    ' ^^ This mode is not suitable for shrinking an image below 25 percent of its original size
    lvicNearestNeighbor = 3         ' same scaling that VB uses (very fast, questionable quality)
    lvicHiqhQualityBilinear = 4
    lvicHighQualityBicubic = 5
    ' ^^ last two: GDI+ does some pre-filtering to ensure highest quality rendering
    '   I believe the filtering is error diffusion. This can be seen with a B&W image that is scaled
    '   to much smaller size. The image is no longer B&W but grayscale
    ' http://msdn.microsoft.com/en-us/library/ms534141(v=VS.85).aspx
    ' last but not least: The values above are actual GDI+ constant values minus 2. So lvicBilinear is really 3.
    '   The GDIpImage.Render class adds 2. The reduction in values is to simply start enumeration at 1
End Enum
    
Public Enum GIFDisposalCodes
    lvicGIF_NoAction = 0     ' no action required by decoder
    lvicGIF_NoDiscard = 1    ' decoder must leave current frame as is
    lvicGIF_Erase = 2        ' decoder erases frame's dimensions with DC background
    lvicGIF_RestorePrevious = 3 ' decoder replaces current frame's dimensions with display before frame rendered
    ' following are for animated PNG only...
    '^^ note for APNG: lvicGIF_NoAction & lvicGIF_NoDiscard have same effect
    lvicAPNG_BlendOpOver = 16 ' renders/blends current frame over existing composite
    ' rule of thumb for using lvicAPNG_BlendOpOver: if frame contains transparency use lvicAPNG_BlendOpOver
    lvicAPNG_DefaultImage_NotAnimated = 32 ' 1st image of APNG is not part of the animation else it is
    ' ^^ the 1st image can be part of the animation or not. Keep in mind that not all PNG viewers out there
    ' support APNG. And if not supported, the viewer will display only the 1st frame
    ' ref: https://wiki.mozilla.org/APNG_Specification#.60fcTL.60:_The_Frame_Control_Chunk
End Enum

Public Enum SHILIconSizeEnum        ' used with the ASSOCIATEDICON structure
' http://msdn.microsoft.com/en-us/library/windows/desktop/bb762185%28v=VS.85%29.aspx
    lvicSHIL_Large_32 = 0           ' 32x32
    lvicSHIL_Small_16 = 1           ' 16x16
    lvicSHIL_ExtraLarge_48 = 2      ' 48x48 (XP & above) else defaults to lvicSHIL_Large_32
    lvicSHIL_Jumbo_256 = 4          ' 256x256 (Vista & above) else defaults to lvicSHIL_Large_32
End Enum
Public Enum AssocIconTypeEnum       ' used with the ASSOCIATEDICON structure
    lvicAssocGenericIcon = 0        ' The file/folder file type icon will be retrieved; no file access performed
    lvicAssocActualIcon = 1         ' The file/folder will be accessed to retrieve icon; N/A for drives
    lvicAssocIconOpened = 8         ' Some icons have a closed & open/selected version
    lvicAssocIconPIDL = 16          ' PIDL determines which icon to use
    lvicAssocIconIndex = 32         ' Icon index is known in advance
End Enum

Private m_Token As cGDIpToken       ' cached public GDI+ token class; prevents it getting destroyed before us

Public Function GDIplusTokenExists() As Boolean
    
    ' Routine simply returns whether or not a GDI+ token is created.
    ' You can avoid creating and destroying your own if this function returns True
    ' However, the token is destroyed when the last usercontrol is destroyed/unloaded
    
    ' If this function ever returns false, GDI+ does not exist on the system or cannot be initialized
    
    GDIplusTokenExists = Not (m_Token.Token = 0&)
    
End Function

Public Function GDIplusTokenVersion() As Single
    
    ' Routine simply returns the version and only 3 possible values are returned
    ' 0 = GDI+ not loaded & GDIplusTokenExists would return False also
    ' 1 = GDI+ v1.0 (file version < 6)
    ' 1.1 = GDI+ v1.1 (file version > 6) and this version supports Effects
    
    GDIplusTokenVersion = m_Token.Version
    
End Function

Public Function CF_AlphaImageCtl() As Integer

    ' Passes the clipboard format registered for this control
    ' When this control places items on the clipboard or into a drag/drop data object, it
    '   also copies a custom format that contains the image data. This custom format is
    '   recognized by the control and allows the format to be extracted from the
    '   clipboard or data object. The data object & clipboard only support VB image formats
    '   of bitmap, wmf, emf. The custom clipboard format supports all image formats.
    
    ' you can test whether it exists or not in your applications....
    ' Example: If Clipboard.GetFormat(CF_AlphaImageCtl) Then
    ' Example from OLE_DragDrop event: If Data.GetFormat(CF_AlphaImageCtl) Then
    
    CF_AlphaImageCtl = g_ClipboardFormat

End Function

Public Property Get AsyncDownloadsEnabled() As Boolean

    ' The control may have the ability to download URLs asyncrhonously; thus not locking your GUI
    ' If this property returns false, then syncrhonous downloads will always be performed.
    
    ' Notes about that async downloads
    '   - The routine is not robustly designed. It will request the destination for the
    '       image data via the URL you pass. If the destination returns anything other than an OK status,
    '       the download will fail
    '   - If the passed URL requires a log-in, the user will be supplied with a log-in window
    '   - Should the simple async routines not provide exactly what you need, you should
    '       supply your own download methods and simply pass the image data your routines
    '       gather to the LoadPictureGDIplus function
    
    If g_AsyncController Is Nothing Then Set g_AsyncController = New cAsyncController
    AsyncDownloadsEnabled = g_AsyncController.AsyncModeAvailable

End Property

Public Property Get AVIEnabled() As Boolean

    ' since we are using standard AVI DLLs, should not have issues initializing the
    ' required dll classes. This property simply confirms whether the classes loaded or not
    
    AVIEnabled = m_Token.AVIClassesInitiated

End Property

Public Function AsyncGetDownloadStates(Optional Downloading As Long, Optional WaitingReplies As Long, Optional Queued As Long) As Long

    ' function will return count of any async downloads are pending or queued
    ' Passing optional parameters will tell you how many async downloads are active, waiting on reply and/or queued
    
    Downloading = 0&: WaitingReplies = 0&: Queued = 0&
    If Not g_AsyncController Is Nothing Then
        g_AsyncController.GetStatus Downloading, WaitingReplies, Queued
        AsyncGetDownloadStates = Downloading + WaitingReplies + Queued
    End If

End Function

Public Sub AsyncAbortDownloads(Optional ByVal ActiveOnly As Boolean = True, Optional ByVal QueuedOnly As Boolean)

    ' function simply aborts any current async downloads
    ' ActiveOnly will abort only those that are currently in progress
    ' QueuedOnly will abort only those that are awaiting a slot to become active (10 max downloads at 1 time)
    ' Passing true for both parameters aborts all downloads (default behavior)

    If Not g_AsyncController Is Nothing Then Call g_AsyncController.Abort(ActiveOnly, QueuedOnly)

End Sub

Public Function GetScaledImageSize(ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
                               ByVal destWidth As Long, ByVal destHeight As Long, _
                               ByRef ScaledWidth As Long, ByRef ScaledHeight As Long, _
                               Optional ByVal Angle As Single = 0!, _
                               Optional ByVal CanScaleUp As Boolean = True, _
                               Optional ByVal CanClip As Boolean = True) As Boolean

    ' Function returns scaled (maintaining scale ratio) for passed destination width/height
    ' The CanScaleUp when set to false will never return scaled sizes > than 1:1
    ' The CanClip parameter, if False, will ensure rotated image does not exceed destination width/height
    ' The scaled width & height is returned in the ScaledWidth & ScaledHeight parameters
    ' If function returns false, return parameters are undefined
    
    GetScaledImageSize = modCommon.GetScaledImageSizes(ImageWidth, ImageHeight, destWidth, destHeight, ScaledWidth, ScaledHeight, Angle, CanScaleUp, CanClip)
    
End Function

Public Function GetScaledContainerSize(ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
                            ByRef ContainerWidth As Long, ByRef ContainerHeight As Long, _
                            Optional ByVal Angle As Single = 0!) As Boolean

    ' Function returns the size of a container required to render the passed dimensions at passed Angle
    ' The scaled width & height are returned in the ContainerWidth & ContainerHeight parameters
    ' If function returns false, return parameters are undefined
    
    GetScaledContainerSize = modCommon.GetScaledCanvasSize(ImageWidth, ImageHeight, ContainerWidth, ContainerHeight, Angle)

End Function

Public Function ConvertRGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 255&) As Long

    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to ARGB format
    ' VB system colors are acceptable, i.e., vbButtonFace
    
    ' What is ARGB? VB stores colors as RGB. In a 4 byte value, two bytes used for R,G,B: BBGGRR
    ' GDI+ requires it in reverse and includes the alpha channel: AARRGGBB

    ConvertRGBtoARGB = modCommon.Color_RGBtoARGB(RGBColor, Opacity)
    
End Function

Public Function ConvertARGBtoRGB(ByVal ARGBcolor As Long, Optional ByRef Opacity As Long) As Long

    ' This routine is the opposite of ConvertRGBtoARGB
    ' Returned color is always RGB format, Opacity parameter will contain ARGBcolor opacity (0-255)

    ' What is ARGB? VB stores colors as RGB. In a 4 byte value, two bytes used for R,G,B: BBGGRR
    ' GDI+ requires it in reverse and includes the alpha channel: AARRGGBB

    ConvertARGBtoRGB = modCommon.Color_ARGBtoRGB(ARGBcolor, Opacity)

End Function

Public Function GetPixelGDIplus(Picture As GDIpImage, ByVal X As Long, ByVal Y As Long, Optional ByRef Opacity As Long) As Long

    ' This routine will return the pixel, within an image, in RGB color format
    ' If passed X,Y are out of the picture's bounds, the result is -1&
    ' If color needed for a GDI+ function, you may want to pass the result & the Opacity parameter to ConvertRGBtoARGB

    ' Note: The pixel color you may have set using SetPixelGDIplus is not guaranteed to be the same
    '   color returned by GetPixelGDIplus. This is due to the image's color format and rounding issues.
    '   Example: Setting a pixel to White in a pre-multiplied image at 0% opacity, will return Black vs White

    GetPixelGDIplus = -1&
    If m_Token.Token = 0& Then Exit Function
    If Picture Is Nothing Then Exit Function
    If Picture.Handle = 0& Then Exit Function
    
    If GdipBitmapGetPixel(Picture.Handle, X, Y, X) = 0& Then GetPixelGDIplus = modCommon.Color_ARGBtoRGB(X, Opacity)

End Function

Public Function SetPixelGDIplus(Picture As GDIpImage, ByVal X As Long, ByVal Y As Long, ByVal RGBColor As Long, Optional ByVal Opacity As Long = 255&) As Boolean

    ' This routine will set the pixel, within an image, in ARGB color format
    ' If passed X,Y are out of the picture's bounds, the result is false
    ' VB system colors are acceptable, i.e., vbButtonFace
    ' Changing a pixel in a GDIpImage assigned to a control, you must refresh that control to see the results immediately

    If m_Token.Token = 0& Then Exit Function
    If Picture Is Nothing Then Exit Function
    If Picture.Handle = 0& Then Exit Function
    
    If Opacity < 0& Then
        Opacity = 0&
    ElseIf Opacity > 255& Then
        Opacity = 255&
    End If
    
    SetPixelGDIplus = (GdipBitmapSetPixel(Picture.Handle, X, Y, modCommon.Color_RGBtoARGB((RGBColor And &HFFFFFF), Opacity)) = 0&)

End Function

Public Function LoadPictureGDIplus(SourceImageData As Variant, _
                                    Optional ByVal KeepOriginalFormat As Boolean = True, _
                                    Optional ByVal Prevent_FileLock_AsyncDownload As Boolean = True) As GDIpImage
                                    

'///// SEE THE DOCUMENT PROVIDED WITH THIS CONTROL FOR DETAILED INFO: LoadPictureGDIp.RTF
                                    
    ' Note: To create a blank GDI+ bitmap from no source, see SavePictureGDIplus

    ' Image formats supported
    '------------------------------------------------------------------------------------
    ' format    paletted    true-color  32bpp-Alpha   Notes
    '------------------------------------------------------------------------------------
    ' AVI       supported   supported   supported     (12)
    ' BMP       supported   supported   supported     (1,3,5)
    ' CUR/ANI   supported   supported   supported     (1,4,5,6)
    ' GIF       supported   n/a         n/a           (4,5,6)
    ' JPG       n/a         supported   n/a           (5,8)
    ' ICO       supported   supported   supported     (1,4,5)
    ' MP3       n/a         n/a         n/a           (4,7)
    ' PAM       supported   supported   supported     (11)
    ' PCX       supported   supported   supported     (2)
    ' PNG       supported   supported   supported     (1,6)
    ' PNM       supported   supported   n/a           (10)
    ' TGA       supported   supported   supported     (1,3)
    ' TIF       supported   supported   supported     (1,4)
    ' WMA       n/a         n/a         n/a           (4,7)
    ' WMF/EMF   n/a         n/a         n/a           (5,9)
    '------------------------------------------------------------------------------------
    ' (1) 16 bpp images are supported. Converted to 24 bpp or 32 bpp depending on transparency
    ' (2) 16 bpp images are not supported
    ' (3) Premultiplied pixel format supported
    ' (4) If contains multiple images, all are accessible
    ' (5) Can be loaded from a VB stdPicture or VB .Picture property
    ' (6) Animated format supported
    ' (7) If multiple images exist they are converted to multipage TIFF else original format used
    ' (8) JPEG2000 not supported
    ' (9) Unlike VB, non-placeable WMF are also supported. Always defaulted to scaled 256x256
    '(10) PNM is "portable any map" and is not a file type but does include these: pbm, pgm, ppm
    '(11) Only grayscale images can be palettized. If not black and white nor grayscale, then RGB or ARGB
    '(12) Required codec must be installed. See LoadPictureGDIp.rtf for more info
    '------------------------------------------------------------------------------------
    
    ' SourceImageData must be one of the following:
    ' To erase image & release memory, use 0&, 0, Nothing, vbNullString, "", Empty
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(0&)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus("")
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Nothing)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Empty)
    
    ' To create an image, must be a valid, complete image format in one of the following
    '   ASSOCIATEDICON structure
    '       .FileName member in one of these formats
    '           full path/filename
    '           full path of a folder/drive and should be terminated with a backslash
    '           file extension (i.e., .txt, .exe)
    '           Numeric value of the icon index within a Windows system imagelist
    '           PIDL. http://ccrp.mvps.org/index.html?support/faqs/faqbrowse.htm
    '       .DesiredSize is one of these values
    '           lvicSHIL_Small_16 will attempt to return a 16x16 icon
    '           lvicSHIL_Large_32 (Default) will attempt to return a 32x32 icon
    '           lvicSHIL_ExtraLarge_48 will attempt to return a 48x48 icon (XP+)
    '           lvicSHIL_Jumbo_256 will attempt to return a 256x256 icon (Vista+)
    '           any other values will default to lvicSHIL_Large_32
    '           Note: that the actual size returned is not guaranteed to be what was desired. See IndexReturned
    '           Note: If .IconType includes lvicAssocIconIndex, then this size indicates which system imagelist to use
    '       .IconType is one or more of the following OR'd together as needed
    '           lvicAssocActualIcon: .FileName relates to an actual file or folder
    '           lvicAssocGenericIcon (Default): .FileName relates to a file type
    '           lvicAssocIconOpened: some icons have a closed & opened/selected version
    '           lvicAssocIconPIDL: .FileName is a system item list ID (pIDL)
    '           lvicAssocIconIndex: .FileName is an icon index from an actual system image list
    '               ^^ .DesiredSize identifies which system image list
    '       .IndexReturned indicates which system image list and icon index were used for the returned icon
    '           to get the icon index: (.IndexReturned And &HFFFFFF)
    '           to get size of system image list as a SHILIconSizeEnum value: .IndexReturned \ &H10000000
    '   File (full path & file name), unicode supported, UNC supported
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(App.Path & "\Globe.png")
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus("\\AppsServer\Images\Logo.gif")
    '       ** DO NOT pass Prevent_FileLock_AsyncDownload as False for removable drives, mapped or UNC paths
    '           -- Above warning provided should the drive become unavailable & GDI+ can no longer read the file
    '   URL (full URL must be provided), unicode supported. PreventFileLock will always be treated as True
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus("http://www.vbforums.com/image.php?u=92357&dateline=1263165890")
    '       Success completely dependent on target server, permissions, and internet connectivity
    '       Method is synchronous unless Prevent_FileLock_AsyncDownload is False
    '           Should Async download fail, will fall back to synchronous download
    '       Async method only supported for HTTP & HTTPS protocols
    '   Byte Array (any dimensions, any LBound)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(LoadResData(101, "Custom"))
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(myByteArray())
    '   stdPicture or VB picture objects
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Me.Icon)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Picture1.Picture)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(ImageList1.ListImages(1).Picture)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(LoadResPicture(101, vbResBitmap))
    '       Exception: If you are loading a 32bpp bitmap that uses the alpha channel into
    '           a stdPicture object. Send the stdPicture by handle instead. All stdPicture
    '           objects are processed by ignoring the alpha channel
    '           Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Image1.Picture.Handle)
    '   Clipboard
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Clipboard)
    '       The clipboard can contain a file, image object, url or file path/name
    '       If multiple objects are on the clipboard; they are processed in this order
    '           1. This control's custom clipboard format
    '           2. Files, only first one is used
    '           3. Picture formats of metafiles, bitmaps, dibs
    '           4. Text (pasted url, file path/name, base64) -- unicode supported
    '   Data object, result of a drag/drop event
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Data)
    '       The DataObject can contain a file or image object
    '       When files exist in data object, only first one is used
    '   Bitmap or icon handle, you are responsible for destroying the handle as appropriate
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(hIcon)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(hBitmap)
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Me.Icon.Handle)
    '       Example: hIcon = ImageList1.ListImages(1).ExtractIcon
    '                AlphaImgCtl1.Picture = LoadPictureGDIplus(hIcon)
    '   Another GDIpImage class. A separate copy of the class and image will be created
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(AlphaImgCtl2.Picture)
    '   Screen object. Creates a screen capture in a 24bpp bitmap
    '       Example: AlphaImgCtl1.Picture = LoadPictureGDIplus(Screen)
    
    ' KeepOriginalFormat if True will attempt to cache the original source data so it can be
    '   extracted and used by the user. Some formats will never be cached and those are:
    '       -- bitmap/icon loaded by handle and also MP3/WMA
    '   If parameter is False, the image will be converted internally to bitmap and if extracted,
    '       will be returned in original format, but not guaranteed to be same color depth/compression
    
    ' Prevent_FileLock_AsyncDownload:
    '   when source is a file
    '       - if true will load the entire file to memory and process the data.
    '           Otherwise, the file is simply used by GDI+ and will remain locked/write-protected
    '           until the image is destroyed. Some files will always be treated as Prevent_FileLock_AsyncDownload=True
    '           and those are: icon, cursors, animated cursors/PNG, TGA, PCX, MP3, WMA, PNM, PAM,
    '               and bitmaps that use the alpha channel
    '   when source is a url
    '       - if true then synchronous downloads will be performed else asyncrhonous
    '           If async download is not supported (see AsyncDownloadsEnabled property) then
    '           routines will fall back to synchronous downloads
    

    If m_Token.Token = 0& Then Exit Function

    Dim tPic As GDIpImage
    Set tPic = modCommon.LoadImage(SourceImageData, KeepOriginalFormat, Prevent_FileLock_AsyncDownload)
    If tPic Is Nothing Then
        Err.Raise 481                       ' return invalid picture
    Else
        Set LoadPictureGDIplus = tPic       ' else return the image class
    End If
        
End Function

Public Function SavePictureGDIplus(Picture As GDIpImage, Destination As Variant, _
                                    Optional ByVal ImageFormat As SaveAsFormatEnum = lvicSaveAsCurrentFormat, _
                                    Optional SaveOptions As Variant) As Boolean
                                    
'///// SEE THE DOCUMENT PROVIDED WITH THIS CONTROL FOR DETAILED INFO: SavePictureGDIp.RTF
    
    ' To create a blank GDI+ bitmap from no source, the following must be passed
    ' Picture :: pass as Nothing
    ' Destination :: pass an instantiated GDIpImage class
    ' SaveOptions :: pass a SAVESTRUCT with the following members filled
    '   .Width & .Height will be the blank bitmap's dimensions
    '   .ColorDepth as requested.
    '       Any invalid values or a value of lvicDefaultReduction/lvicNoColorReduction will default to lvicConvert_TrueColor32bpp_ARGB
    '   .Palette_Handle as non-zero if .ColorDepth is 2,16 or 256 colors
    '       if .Palette_Handle is zero and .ColorDepth is 2,16 or 256 colors then halftone palette will be applied
    '   For 16bpp or higher color depths...
    '       If you wish to have the blank image filled with color then supply one of these members
    '           .RSS.FillBrushGDIplus_Handle or
    '           .RSS.FillColorUsed = True and also .RSS.FillColorARGB
    ' All other parameters and structure members are ignored. ImageFormat will always be treated as Bitmap
    
    ' For all other scenarios...
    ' Parameters:
    '   Picture :: must be a GDIpImage object, either created by you, or a AlphaImgCtrl.Picture object
    '           See remarks below regarding multi-page TIFFs; in that case parameter can be Nothing
    
    '   Destination :: must be one of the following
    '       - FileName must contain full path, not write-protected. Unicode file names supported
    '             :Example. SavePictureGDIplus AlphaImgCtrl2.Picture, "C:\Screenshot.jpg", lvicSaveAsJPEG, 50
    '       - Byte array. Destination resized as needed. Will be zero bound, 1D array
    '             :Example. SavePictureGDIplus AlphaImgCtrl2.Picture, byteArray(), lvicSaveAsPNG
    '       - GDIpImage object. Must pass initialized class, i.e., Dim cImage As New GDIpImage
    '             :Example. SavePictureGDIplus AlphaImgCtrl2.Picture, cImage, lvicSaveAsCurrentFormat
    '       - stdPicture object. Must pass initialized object, i.e., Dim myPic As New stdPicture
    '           If any transparency used, will be filled with SaveOptions.RSS.FillColorARGB
    '               If RenderingStyle not passed, vbWindowBackground used
    '           On return, to apply to a Picturebox, image control: Set Picture1.Picture = myPic
    '             :Example. SavePictureGDIplus AlphaImgCtrl2.Picture, myPic, lvicSaveAsCurrentFormat
    '       - Data object, passed during OLEStartDrag or OLESetData events
    '           Data object will contain stdPicture object and special format recognized by this control
    '           That special format ensures another control can process the data without color loss or format loss
    '           If transparency exists, stdPicture effected same as the stdPicture destination comments above
    '           Ensure you clear the Data object first if needed: Data.Clear
    '             :Example. SavePictureGDIplus AlphaImgCtrl2.Picture, Data, lvicSaveAsCurrentFormat
    '       - Clipboard. Result will be pasted to the clipboard
    '           Same comments as the Data object above
    '           Ensure you clear the Clipboard first if needed: Clipboard.Clear
    '             :Example. SavePictureGDIplus AlphaImgCtrl2.Picture, Clipboard, lvicSaveAsCurrentFormat
    '       - Icon or Bitmap Handle (SaveAs_HICON & SaveAs_HBITMAP)
    '           Destination will contain the handle. You are responsible for destroying the handle
    
    '   ImageFormat :: The requested format.
    '       - If SaveAsCurrentFormat then
    '           To retrieve the original source data if it is kept by the class...
    '               Do not pass SaveOptions, do not opt for resizing & save to anything but handle or stdPicture
    
    '   SaveOptions :: Optional conversion & rendering effects
    '       - If creating a multiplage TIFF
    '           :: pass a filled MULTIPAGETIFFSTRUCT
    '           :: pass new MULTIIMAGESAVESTRUCT structure supersedes the MUOLTIPAGETIFFSTRUCT and supports
    '                   icon, cursor, GIF, TIFF, PNG
    '       - For all other image formats, pass a filled SAVESTRUCT if desired
    '       - Exception: If converting AVI to animated GIF/PNG, pass a RENDERSTYLESTRUCT2 if you wish to
    '           have a color be transparent on each frame
    '           :: to use the top left color of 1st frame as transparency key, just pass a RENDERSTYLESTRUCT2
    '           :: to use a specific color, fill in the RENDERSTYLESTRUCT2.TransparentColor & .TransparentColorUsed members
    
    '   Return Value and the Destination parameter
    '       If the function fails, return value is False and Destination parameter is unchanged
    '       If function is successful, reutrn value is True and
    '           - If saving to file, Destination parameter unchanged
    '           - If saving to Array, array contained in Destination parameter is filled, zero LBound
    '           - If saving to GDIpImage class, the class contained in Destination now contains the image
    '           - If saving to stdPicture, the stdPicture contained in Destination now contains the image
    '           - If saving to Clipboard, Destination parameter unchanged, image placed on clipboard in Bmp/Metafile format
    '           - If saving to Data Object, Destination parameter unchanged, image placed in data object in Bmp/Metafile format
    '           - If saving to HICON or HBITMAP, destination parameter is the handle. You are responsible for destroying handle
    
    
    ' How to use the SaveOptions parameter.
    ' Let's say you want to save the image to file, but in grayscale & 2x its size
    '
    '   # Dim SS As SAVESTRUCT, bResult As Boolean
    '   # Set SS.RSS.Effects = New GDIpEffects
    '   # RS.RSS.Effects.GrayScale = lvicNTSCPAL
    '   # SS.Width = AlphaImgCtl2.Picture.Width * 2
    '   # SS.Height = AlphaImgCtl2.Picture.Height * 2
    '   # bResult = SavePictureGDIplus(AlphaImgCtl2.Picture, "C:\Sample.png", lvicSaveAsPNG, SS)
    
    ' How to create a multi-page TIFF from GDIpImage objects....(Supported by superseded. See below)
    ' #1. Prepare a MULTIPAGETIFFSTRUCT
    ' #2. Fill in the TIFFPAGESTRUCT for each image to be added as a TIFF page, if custom rendering wanted
    ' #3. Call this function, passing the MULTIPAGETIFFSTRUCT structure, as the SaveOptions parameter
    ' Example, using a small array of AlphaImgCtrl controls
    
    '   # Dim MPTIFF As MULTIPAGETIFFSTRUCT
    '   # Dim lCount As Long, bResult As Boolean
    '   #
    '   # MPTIFF.Pages = 3
    '   # ReDim MPTIFF.TPS(0 To MPTIFF.Pages - 1)
    '   # For lCount = 0 To MPTIFF.Pages - 1
    '   #   Set MPTIFF.TPS(lCount).Picture = AlphaImgCtrls(lCount).Picture
    '   #       ' fill in optional rendering style attributes for each page using the TIFFPAGESTRUCT.SS member
    '   # Next
    '   # bResult = SavePictureGDIplus(Nothing, "C:\Screenshots.tiff", lvicSaveAsTIFF, MPTIFF)


    ' The MULTIIMAGESAVESTRUCT can be used to create animated GIFs/PNGs, multipage TIFFs, multi-image icon/cursor
    
    ' How to create an animated GIF from GDIpImage objects....
    ' #1. Prepare a MULTIIMAGESAVESTRUCT
    ' #2. Fill in the GIFOVERVIEWSTRUCT member
    ' #3. Fill in the GIFFRAMESTRUCT for each image to be added as Gif frame
    ' #4. Call this function, passing the MULTIIMAGESAVESTRUCT structure, as the SaveOptions parameter
    ' Example, using a small array of AlphaImgCtrl controls
    
    '   # Dim MIS As MULTIIMAGESAVESTRUCT
    '   # Dim lCount As Long, bResult As Boolean
    '   #
    '   # MIS.Images = 3
    '   # ReDim MIS.Image(0 To MIS.Images - 1)
    '   # For lCount = 0 To MIS.Images - 1
    '   #   Set MIS.Image(lCount).Picture = AlphaImgCtrls(lCount).Picture
    '   #       ' fill in optional rendering style attributes for each frame using the Image(n).SS member
    '   #       MIS.Image(lCount).SS.AlphaTolerancePct = 25 ' set conversion quality from 32bpp to 8bpp
    '   #   MIS.Image(lCount).GIFFrameInfo.TimeDelay = 1000 ' full second
    '   #   MIS.IMAGE(lCount).GIFFrameInfo.DisposalCode = lvicGIF_Erase ' understanding disposal codes is necessary
    '   #   MIS.Image(lCount).GIFFrameInfo.XOffset = -1 ' center horizontally in overall GIF window
    '   #   MIS.Image(lCount).GIFFrameInfo.YOffset = -1 ' center vertically in overall GIF window
    '   # Next
    '   #
    '   # MIS.GIFOverview.LoopCount = 100
    '   # ' leave MIS.GIFOverview.WindowWidth & WindowHeight zero to auto-calculate, else provide valid values
    '   #
    '   # bResult = SavePictureGDIplus(Nothing, "C:\LogoAnimated.gif", lvicSaveAsGIF, MIS)
    
    ' Similar method used for creating animated PNG
    '   # Dim MIS As MULTIIMAGESAVESTRUCT
    '   # Dim lCount As Long, bResult As Boolean
    '   #
    '   # MIS.Images = 3
    '   # ReDim MIS.Image(0 To MIS.Images - 1)
    '   # For lCount = 0 To MIS.Images - 1
    '   #   Set MIS.Image(lCount).Picture = AlphaImgCtrls(lCount).Picture
    '   #       ' fill in optional rendering style attributes for each frame using the Image(n).SS member
    '   #   MIS.Image(lCount).GIFFrameInfo.TimeDelay = 1000 ' full second
    '   #   MIS.IMAGE(lCount).GIFFrameInfo.DisposalCode = lvicGIF_Erase
    '   #   ' ^^ understanding disposal codes is necessary & more important with APNG.
    '   #   ' See notes I've included with the GIFDisposalCodes enumeration at top of class
    '   #   MIS.Image(lCount).GIFFrameInfo.XOffset = -1 ' center horizontally in overall GIF window
    '   #   MIS.Image(lCount).GIFFrameInfo.YOffset = -1 ' center vertically in overall GIF window
    '   # Next
    '   #
    '   # MIS.GIFOverview.LoopCount = 100
    '   # ' leave MIS.GIFOverview.WindowWidth & WindowHeight zero to auto-calculate, else provide valid values
    '   #
    '   # bResult = SavePictureGDIplus(Nothing, "C:\LogoAnimated.png", lvicSaveAsPNG, MIS)

    ' How to create an multipage TIFF from GDIpImage objects....
    ' #1. Prepare a MULTIIMAGESAVESTRUCT
    ' #2. Call this function, passing the MULTIIMAGESAVESTRUCT structure, as the SaveOptions parameter
    ' Example, using a small array of AlphaImgCtrl controls
    
    '   # Dim MIS As MULTIIMAGESAVESTRUCT
    '   # Dim lCount As Long, bResult As Boolean
    '   #
    '   # MIS.Images = 3
    '   # ReDim MIS.Image(0 To MIS.Images - 1)
    '   # For lCount = 0 To MIS.Images - 1
    '   #   Set MIS.Image(lCount).Picture = AlphaImgCtrls(lCount).Picture
    '   #       ' fill in optional rendering style attributes for each page using the Image(n).SS member
    '   # Next
    '   # bResult = SavePictureGDIplus(Nothing, "C:\Screenshots.tif", lvicSaveAsTIF, MIS)

    ' How to create an multi-image icon from an existing multi-image icon....
    ' #1. Prepare a MULTIIMAGESAVESTRUCT
    ' #2. Call this function, passing the MULTIIMAGESAVESTRUCT structure, as the SaveOptions parameter
    ' Example, using an existing icon resource that contains several images. Reverse the order:
    
    '   # Dim MIS As MULTIIMAGESAVESTRUCT
    '   # Dim lCount As Long, bResult As Boolean, newIndex As Long
    '   #
    '   # MIS.Images = AlphaImgCtl1.ImageCount
    '   # ReDim MIS.Image(0 To MIS.Images - 1)
    '   # For lCount = 0 To MIS.Images - 1
    '   #   Set MIS.Image(lCount).Picture = AlphaImgCtl1.Picture
    '   #       ' fill in optional rendering style attributes for each image using the Image(n).SS member
    '   #   newIndex = MIS.Images - lCount ' image indexes are 1-bound not 0-bound
    '   #   MIS.Image(lCount).FrameNumber = newIndex ' reverse the order
    '   #   If AlphaImgCtl1.Picture.Width(newIndex) > 128 Or AlphaImgCtl1.Picture.Height(newIndex) > 128 Then
    '   #       MIS.Image(lCount).IconEmbeddedAsPNG = True ' option to save icon/cursor in PNG format (Vista+)
    '   #   End If
    '   # Next
    '   # bResult = SavePictureGDIplus(Nothing, "C:\AppIcon.ico", lvicSaveAsICO, MIS)

    ' How to convert an animated GIF to an animated PNG...
    ' The conversion is automatic when the following conditions are met
    ' #1. Source picture contains animated GIF (i.e., more than 1 frame)
    ' #2. Target image format is PNG
    ' #3. No optional SaveOptions parameter passed to this this function
    ' #4. The destination is not a stdPicture object
    ' -- No reverse automatic conversion provided since APNG frames cannot
    '   always be fully converted to GIF without loss of image quality
    
    '   # Dim bResult As Boolean
    '   # ' assuming AlphaImgCtl1.Picture contains an animated GIF resource
    '   # bResult = SavePictureGDIplus(AlphaImgCtl1.Picture, "C:\TrainingImages.png", lvicSaveAsPNG)

    ' How to convert an AVI to either an animated GIF or an animated PNG...
    ' The conversion is automatic when the following conditions are met
    ' #1. Source picture contains an AVI
    ' #2. Target image format is PNG or GIF
    ' #3. No optional SaveOptions parameter passed to this this function
    '       - Exception: If you want a color transparent from every frame of the AVI then
    '       a. Pass a RENDERSTYLESTRUCT2. Leave the .Effects.TransparentColorUsed member False
    '           Doing so will have the control use the top left pixel of each frame as the transparency color
    '       b. Pass a RENDERSTYLESTRUCT2. Set the .Effects.TransparentColor to the RGB color you want made
    '           transparent and then set the .Effects.TransparentColorUsed member to True
    '       c. Passing the SaveOptions member nothing will perform no additional transparency actions on the frames
    ' #4. The destination is not a stdPicture object
    ' -- No reverse automatic conversion provided since saving to AVI is not supported
    ' -- Converting AVI to GIF can result in color loss if AVI frames contain more than 256 colors
    '       Suggestion: If frames are 8bpp or less, convert to GIF else to PNG
    
    '   # Dim bResult As Boolean
    '   # ' assuming AlphaImgCtl1.Picture contains an AVI
    '   # ' to animated GIF:
    '   #   bResult = SavePictureGDIplus(AlphaImgCtl1.Picture, "C:\TrainingImages.gif", lvicSaveAsGIF)
    '   # ' to animated PNG:
    '   #   bResult = SavePictureGDIplus(AlphaImgCtl1.Picture, "C:\TrainingImages.png", lvicSaveAsPNG)
    '   # ' Example of making top left corner pixel color transparent
    '   #   Dim RSS As RENDERSTYLESTRUCT2
    '   #   bResult = SavePictureGDIplus(AlphaImgCtl1.Picture, "C:\TrainingImages.png", lvicSaveAsPNG, RSS)

    If m_Token.Token = 0& Then Exit Function
    
    SavePictureGDIplus = modCommon.SaveImage(Picture, Destination, ImageFormat, SaveOptions)
    
End Function

Public Function PaintPictureGDIplus(Picture As GDIpImage, ByVal hDC As Long, _
                                    Optional ByVal destX As Long = 0&, Optional ByVal destY As Long = 0&, _
                                    Optional ByVal destWidth As Long = 0&, Optional ByVal destHeight As Long = 0&, _
                                    Optional ByVal srcX As Long = 0&, Optional ByVal srcY As Long = 0&, _
                                    Optional ByVal srcWidth As Long = 0&, Optional ByVal srcHeight As Long = 0&, _
                                    Optional ByVal RenderStyle As Variant) As Boolean

    ' Routine renders any GDIpImage to any hDC with optional rendering styles
    ' Note: This is a global wrapper around the GDIpImage class' Render function. That function
    '       offers more advanced users some GDI+ parameters
    
    ' All optional sizing parameters use image size if parameters not provided
    ' If custom attributes wanted while rendering, provide a filled-in RENDERSTYLESTRUCT2 as RenderStyle parameter
    '   Not all RENDERSTYLESTRUCT2 members must be filled
    
    ' An example. We'll render a loaded image to a Picturebox DC with some attributes
    '
    '  #    Dim RSS As RENDERSTYLESTRUCT2
    '  #    With RSS
    '  #        Set .Effects = New GDIpEffects
    '  #        .Effects.GrayScale = lvicNTSCPAL   ' gray scale
    '  #        .Effects.LightnessPct = -10        ' 10% darker
    '  #        ' we'll make top left pixel color transparent throughout
    '  #        .Effects.TransparentColor = GetPixelGDIplus(AlphaImgCtl1.Picture, 0, 0)
    '  #        .Effects.TransparentColorUsed = True
    '  #    End With
    '  #    PaintPictureGDIplus AlphaImgCtl1.Picture, Picture1.hDC, , , , , , , , , RSS
    
    If m_Token.Token = 0& Then Exit Function
    If Picture Is Nothing Then Exit Function
    If Picture.Handle = 0& Then Exit Function
    
    Dim RS As RENDERSTYLESTRUCT2, hBrush As Long
    Dim hGraphics As Long, bClipped As Boolean
    
    If Not IsMissing(RenderStyle) Then
        On Error Resume Next
        If VarType(RenderStyle) = vbUserDefinedType Then
            If LenB(RenderStyle) = LenB(RS) Then RS = RenderStyle
        End If
        If Err Then Exit Function
        On Error GoTo 0
    End If
    If destWidth = 0& Then destWidth = Picture.Width
    If destHeight = 0& Then destHeight = Picture.Height
    If (destWidth < 0& Or srcWidth < 0&) Then
        If Not (srcWidth < 0& And destWidth < 0&) Then RS.Mirrored = RS.Mirrored Xor lvicMirrorHorizontal
        destWidth = Abs(destWidth): srcWidth = Abs(srcWidth)
    End If
    If (destHeight < 0& Or srcHeight < 0&) Then
        If Not (srcHeight < 0& And destHeight < 0&) Then RS.Mirrored = RS.Mirrored Xor lvicMirrorHorizontal
        destHeight = Abs(destHeight): srcHeight = Abs(srcHeight)
    End If
    
    On Error GoTo ExitRoutine
    modCommon.ValidateRenderStyle2 RS, Picture
    If RS.FillColorUsed Then
        If GdipCreateFromHDC(hDC, hGraphics) Then Exit Function
        If GdipCreateSolidFill(RS.FillColorARGB, hBrush) = 0& Then
            GdipFillRectangleI hGraphics, hBrush, destX, destY, destWidth, destHeight
            GdipDeleteBrush hBrush
        End If
    End If
    If RS.FillBrushGDIplus_Handle Then
        If hGraphics = 0& Then
            If GdipCreateFromHDC(hDC, hGraphics) Then Exit Function
        End If
        GdipFillRectangleI hGraphics, RS.FillBrushGDIplus_Handle, destX, destY, destWidth, destHeight
    End If
    With RS
        If Not (.Angle = 0! Or .Angle = 180!) Then
            If hGraphics = 0& Then
                If GdipCreateFromHDC(hDC, hGraphics) Then Exit Function
            End If
            bClipped = (GdipSetClipRectI(hGraphics, destX, destY, destWidth, destHeight, 0&) = 0&)
        End If
        If (RS.Mirrored And lvicMirrorHorizontal) Then destWidth = -destWidth
        If (RS.Mirrored And lvicMirrorVertical) Then destHeight = -destHeight
        PaintPictureGDIplus = Picture.Render(hDC, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, .Angle, .Effects.AttributesHandle, hGraphics, .Effects.EffectsHandle(.EffectType))
        If bClipped Then GdipResetClip hGraphics: bClipped = False
    End With
    PaintPictureGDIplus = True
    
ExitRoutine:
    If bClipped Then GdipResetClip hGraphics
    If hGraphics Then GdipDeleteGraphics hGraphics

End Function

Public Function MergePictureGDIplus(ByRef PictureMergeData As PICTUREMERGESTRUCT) As GDIpImage

    ' Routine enables you to combine 2 or more images into a new single image
    ' Additionally, optional rendering styles can be applied to each image
    
    ' The new image will have the dimensions of PictureMergeData's CanvasWidth and CanvasHeight values
    ' Each image appended will be rendered at PictureMergeData.MIS(n).Left & .Top
    ' Each image will be scaled to PictureMergeData.MIS(n).Width & .Height
    ' Not all structure members need to be filled in, the minimal required are:
    '   PictureMergeData.CanvasWidth, PictureMergeData.CanvasHeight, PictureMergeData.Pictures
    '   For each picture, PictureMergeData.MIS(n).Picture must be set to an existing GDIpImage
    '   For each picture, PictureMergeData.MIS(n).Left & .Top should be provided if not rendering at 0,0
    '   Also, PictureMergeData.MIS(n).Width & .Height should be provided if
    '       you don't want the rendered image full size.
    '   All PictureMergeData.MIS(n).RSS members are optional
    
    ' Here's an example. We'll have a watermark PNG image and want to render it to a resource image
    ' before it is displayed or saved.
    
    '  #   Dim PMS As PICTUREMERGESTRUCT, Cx As Long, Cy As Long
    '  #   With PMS
    '  #       .Pictures = 2
    '  #       ReDim .MIS(0 To .Pictures - 1)
    '  #       ' our source image
    '  #       Set .MIS(0).Picture = LoadPictureGDIplus(LoadResData(101, "Custom"))
    '  #       ' setup the overall dimensions from source image
    '  #       .CanvasWidth = .MIS(0).Picture.Width
    '  #       .CanvasHeight = .MIS(0).Picture.Height
    '  #       With .MIS(1)  ' we'll make the watermark 1/4 size of source image & put it in bottom right corner
    '  #           ' our watermark image
    '  #           Set .Picture = LoadPictureGDIplus(LoadResData(111, "Custom"))
    '  #           ' get scaled size 1/4 of the canvas dimensions & store it in .Width/.Height
    '  #           GetScaledImageSize .Picture.Width, .Picture.Height, _
    '  #               PMS.CanvasWidth \ 4, PMS.CanvasHeight \ 4, Cx, Cy
    '  #           .Width = Cx: .Height = Cy
    '  #           ' now position it to the bottom right of source image
    '  #           .Left = PMS.CanvasWidth - .Width
    '  #           .Top = PMS.CanvasHeight - .Height
    '  #           ' let's make it 40% transparent also
    '  #           Set .RSS.Effects = New GDIpEffects
    '  #           .RSS.Effects.GlobalTransparencyPct = 40
    '  #       End With
    '  #   End With
    '  #   ' finally, let's create it and show it in an alpha image control
    '  #   Set AlphaImgCtl1.Picture = MergePictureGDIplus(PMS)
    '  #   ' we can free up a little memory by not caching the source data
    '  #   AlphaImgCtl1.Picture.KeepOriginalFormat = False
    
    If m_Token.Token = 0& Then Exit Function
    
    Dim cNewImage As GDIpImage
    Set cNewImage = modCommon.OverlayImage(PictureMergeData)
    If cNewImage Is Nothing Then Set MergePictureGDIplus = New GDIpImage Else Set MergePictureGDIplus = cNewImage

End Function

Public Function TilePictureGDIplus(Picture As GDIpImage, Destination As Variant, _
                                ByVal destX As Long, ByVal destY As Long, _
                                ByVal destWidth As Long, ByVal destHeight As Long, _
                                ByVal tileWidth As Long, ByVal tileHeight As Long, _
                                Optional ByVal tileGapWidth As Long = 0&, _
                                Optional ByVal tileGapHeight As Long = 0&, _
                                Optional ByVal TileStyle As TileOrderEnum = lvicTile_NoFlip, _
                                Optional ByVal ScalePictureOnTile As Boolean = True, _
                                Optional ByVal StaggerRows As Boolean = False, _
                                Optional RenderingStyle As Variant) As Boolean
                                
    ' GDI+ tile brushes are very slow in my opinion; not just a little, but very, very slow
    ' So this function is a wrapper for a very fast memory-tiling algorithm. I can tile an entire
    ' screen 100's times faster than GDI+

    '   Picture :: must be a GDIpImage object, either created by you, or a AlphaImgCtrl.Picture object
    '   Desintation :: can be one of the following
    '       1) valid hDC handle. Example. modCommon.TileImage AlphaImgCtl1.Picture, Picture1.hDC, ....
    '           -- destWidth,destHeight determine overall size of the tiled hDC area
    '       2) or an initialized GDIpImage object: Example.
    '               Dim bkgImage As GDIpImage: Set bkgImage = New GDIpImage
    '               modCommon.TileImage AlphaImgCtl1.Picture, bkgImage, ....
    '           -- destWidth,destHeight determine overall size of the returned GDIpImage object
    '   destX,destY :: the starting position on the hDC for tiling
    '   tileWidth,tileHeight :: the size of the individual tiles
    '   tileGapWidth,tileGapHeight :: distance between adjacent tiles.
    '       -- horizontal gap does not apply to the first tile of each row
    '       -- vertical gap does not apply to the 1st row of tiles
    '       -- if RenderingStyle.FillColorUsed is True, gaps are filled with RenderingStyle.FillColorARGB
    '   TileStyle determines how the every other tile will be rendered in relation to previous tile
    '   ScalePictureOnTile :: If true, will scale the Picture to tileWidth,tileHeight else will be stretched
    '       -- exception is if RenderingStyle.Angle is not 0,180,360. No stretching when rotating
    '   StaggerRows:: If true, alternating rows will begin rendering at (tileWidth+tileGapWidth)\2
    '       -- parameter allows a staggered tile appearance
    '   RenderingStyle :: Pass a filled RENDERSTYLESTRUCT2 or don't pass it if no special rendering style needed
    '       - This function does not use rendering attributes utilized by any control maintaining the passed Picture object
    
    ' Note: By supplying gap width same as tile width & using StaggerRows, you can get a checkboard pattern
    
    ' Simple example. Tiling an image as 64x64 squares in grayscale
    
    '  #   Dim RSS As RENDERSTYLESTRUCT2
    '  #   Set RSS.Effects = New GDIpEffects
    '  #   RSS.Effects.GrayScale = lvicNTSCPAL
    '  #   ' Picture1.ScaleMode is assumed to be Pixels else convert ScaleWidth,ScaleHeight to pixels
    '  #   TilePictureGDIplus AlphaImgCtl1.Picture, Picture1.hDC, 0, 0, _
    '  #                        Picure1.ScaleWidth, Picture1.ScaleHeight, 64, 64, , , , , RSS
    
    If m_Token.Token = 0& Then Exit Function
    
    Dim RS As RENDERSTYLESTRUCT2
    If Not IsMissing(RenderingStyle) Then
        If VarType(RenderingStyle) = vbUserDefinedType Then
            If LenB(RenderingStyle) = LenB(RS) Then RS = RenderingStyle
        End If
    End If
    TilePictureGDIplus = modCommon.TileImage(Picture, Destination, destX, destY, destWidth, destHeight, _
                            TileStyle, tileWidth, tileHeight, tileGapWidth, tileGapHeight, _
                            ScalePictureOnTile, StaggerRows, RS)

End Function

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
End Sub

Private Sub Class_Terminate()
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub
