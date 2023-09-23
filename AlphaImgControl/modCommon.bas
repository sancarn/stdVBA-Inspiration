Attribute VB_Name = "modCommon"
Option Explicit

' lots and lots of functions used by the various classes and usercontrol
' anything declared as Public is public to project only, not clients
' Public declarations/methods are used in one or more project classes/usercontrol
' Private declarations/methods are used in this module only

' Internet related APIs for loading image via URL
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByRef hInternet As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenW Lib "wininet.dll" (ByVal sAgent As Long, ByVal lAccessType As Long, ByVal sProxyName As Long, ByVal sProxyBypass As Long, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long
Private Declare Function InternetOpenUrlW Lib "wininet.dll" (ByVal hInternet As Long, ByVal lpszUrl As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Long

' GDI+ and supporting APIs
Public Declare Function GdipBitmapGetPixel Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByVal X As Long, ByVal Y As Long, ByRef pColor As Long) As Long
Public Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTI, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Public Declare Function GdipBitmapSetPixel Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal pColor As Long) As Long
Public Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Public Declare Function GdipCreateBitmapFromGraphics Lib "GdiPlus.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal ptarget As Long, ByRef pbitmap As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal hbm As Long, ByVal hPal As Long, ByRef pbitmap As Long) As Long
Public Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Public Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, hGraphics As Long) As Long
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus.dll" (ByVal pbitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "GdiPlus.dll" (ByRef imgAttr As Long) As Long
Public Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (ByRef pRect As RECTI, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Mode As Long, ByVal WrapMode As Long, ByRef lineGradient As Long) As Long
Private Declare Function GdipCreateRegionHrgn Lib "GdiPlus.dll" (ByVal hRgn As Long, ByRef region As Long) As Long
Public Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Public Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteRegion Lib "GdiPlus.dll" (ByVal region As Long) As Long
Public Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Public Declare Function GdipDisposeImageAttributes Lib "GdiPlus.dll" (ByVal imgAttr As Long) As Long
Public Declare Function GdipDrawImagePointsRect Lib "gdiplus" (ByVal graphics As Long, ByVal pImage As Long, ByRef pPoints As Any, ByVal Count As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As Long, ByVal imageAttributes As Long, Optional ByVal pcallback As Long = 0&, Optional ByVal callbackData As Long = 0&) As Long
Public Declare Function GdipDrawImageRectRectI Lib "GdiPlus.dll" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "GdiPlus.dll" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipEmfToWmfBits Lib "GdiPlus.dll" (ByVal hEMF As Long, ByVal cbData16 As Long, ByVal pData16 As Long, ByVal iMapMode As Long, ByVal eFlags As Long) As Long
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GdipGetDC Lib "GdiPlus.dll" (ByVal graphics As Long, ByRef hDC As Long) As Long
Private Declare Function GdipGetHemfFromMetafile Lib "GdiPlus.dll" (ByVal metafile As Long, ByRef hEMF As Long) As Long
Public Declare Function GdipGetImageBounds Lib "GdiPlus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "GdiPlus.dll" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "GdiPlus.dll" (numEncoders As Long, Size As Long) As Long
Public Declare Function GdipGetImageGraphicsContext Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef graphics As Long) As Long
Private Declare Function GdipGetImageHorizontalResolution Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef resolution As Single) As Long
Public Declare Function GdipGetImagePalette Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef Palette As ColorPalette, ByVal pSize As Long) As Long
Public Declare Function GdipGetImagePaletteSize Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef pSize As Long) As Long
Public Declare Function GdipGetImagePixelFormat Lib "GdiPlus.dll" (ByVal hImage As Long, PixelFormat As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "GdiPlus.dll" (ByVal hImage As Long, ByVal GUID As Long) As Long
Private Declare Function GdipGetImageVerticalResolution Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef resolution As Single) As Long
Public Declare Function GdipGraphicsClear Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pColor As Long) As Long
Public Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, ByRef dimensionID As Any, ByRef Count As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal FileName As Long, hImage As Long) As Long
Public Declare Function GdipLoadImageFromStream Lib "GdiPlus.dll" (ByVal Stream As Long, Image As Long) As Long
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByVal propSize As Long, ByRef buffer As Any) As Long
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByRef Size As Long) As Long
Public Declare Function GdipReleaseDC Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal hDC As Long) As Long
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipRecordMetafileStream Lib "GdiPlus.dll" (ByVal pStream As Long, ByVal referenceHdc As Long, ByVal pType As Long, ByRef frameRect As RECTF, ByVal frameUnit As Long, ByVal description As Long, ByRef metafile As Long) As Long
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal graphics As Long) As Long
Public Declare Function GdipRestoreGraphics Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pState As Long) As Long
Public Declare Function GdipRotateWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipSaveAdd Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mEncoderParams As Any) As Long
Private Declare Function GdipSaveAddImage Lib "GdiPlus.dll" (ByVal pImage As Long, ByVal newImage As Long, ByRef encoderParams As Any) As Long
Public Declare Function GdipSaveGraphics Lib "GdiPlus.dll" (ByVal graphics As Long, ByRef pState As Long) As Long
Public Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal combineMode As Long) As Long
Private Declare Function GdipSetClipRegion Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal hRgn As Long, ByVal combineMode As Long) As Long
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixOffsetMode As Long) As Long
Public Declare Function GdipSaveImageToStream Lib "GdiPlus.dll" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Public Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, ByRef dimensionID As Any, ByVal FrameIndex As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "GdiPlus.dll" (ByVal mImageattr As Long, ByVal mType As Long, ByVal mEnableFlag As Long, ByVal mColorLow As Long, ByVal mColorHigh As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "GdiPlus.dll" (ByVal imgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, ByRef clrMatrix As Any, ByRef grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipSetImageAttributesThreshold Lib "GdiPlus.dll" (ByVal imageattr As Long, ByVal pType As Long, ByVal enableFlag As Long, ByVal threshold As Single) As Long
Public Declare Function GdipSetImagePalette Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef Palette As ColorPalette) As Long
Public Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Public Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Public Declare Function GdipTranslateWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Public Const SmoothingModeAntiAlias As Long = &H4
' GDI+ Path-Related APIs
Public Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mFillMode As Long, ByRef mpath As Long) As Long
Public Declare Function GdipAddPathArc Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Public Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Public Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Public Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Public Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Public Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long


' GDI32/MSImg32 API functions
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32.dll" (ByVal hEMF As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Public Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetClipBox Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpRect As RECTI) As Long
Public Declare Function GetClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Public Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Public Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECTI) As Long
Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Public Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByRef lprcUpdate As RECTI, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Const NULL_BRUSH As Long = 5

' Kernel32 APIs
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileA Lib "kernel32.dll" (ByVal lpFileName As String) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Declare Function EnumResourceNamesA Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByRef lParam As Long) As Long
Public Declare Function EnumResourceNamesW Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByRef lParam As Long) As Long
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal length As Long, ByVal Fill As Byte)
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Any) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

' User32 APIs
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function CreateIconFromResourceEx Lib "user32.dll" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECTI, ByVal hBrush As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As Any) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECTI, ByRef lpSrc1Rect As RECTI, ByRef lpSrc2Rect As RECTI) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECTI, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECTI, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECTI, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Const INVALID_HANDLE_VALUE = -1&
Private Const FILE_ATTRIBUTE_NORMAL = &H80&

' Miscellaneous APIs (VB5 users change msvbvm60 to msvbvm50 below)
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef ptr() As Any) As Long
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GetVersion Lib "kernel32.dll" () As Long
Public Declare Function GetHGlobalFromStream Lib "ole32.dll" (ByVal ppstm As Long, hGlobal As Long) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Sub OleGetClipboard Lib "ole32.dll" (ByRef ppDataObj As Long)
Private Declare Function OleLoadPicture Lib "OLEPRO32.DLL" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Const VT_BYREF As Long = &H4000&            ' Variant type

Private Declare Function SHGetImageListXP Lib "shell32.dll" Alias "#727" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Any) As Long
Private Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Any) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As Any) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIML As Long, ByVal pIndex As Long, ByVal flags As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetFileInfoW Lib "shell32" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, ByVal psfi As Long, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Const MAX_PATH As Long = 260&
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type


Public Enum AlphaReductionLimit     ' how alpha is handled when palettizing
    alpha_None = 0                  ' alpha is not allowed; therefore removed as needed (i.e., JPG)
    alpha_Simple = 1                ' 100% transparent/opaque only (i.e., GIF)
    alpha_Complex = 2               ' palette colors can contain alpha values (i.e., TGA/PNG)
End Enum
Public Enum RawDataOrientation      ' how color reduced data is to be returned
    orient_GDIpHandle = 0           ' as a GDI+ handle
    orient_TopDown = 1              ' array: top-down image data (pcx,gif,png,tga,bmp if to handle)
    orient_BottomUp = 2             ' array: bottom-up image data (icon/cursor/bitmap if to file)
    orient_WantMask = 4             ' return mask in array; array includes indexes,mask,color palette
    orient_8bppIndexes = 8          ' return as 8bpp indexes regardless of color depth (gif/tga)
    orient_SortGrayscale = 16       ' TGA handles grayscale differently, but image indexes must be in correct order
    orient_4bppIndexesMin = 32      ' PCX cannot use 1bpp (non-b&w paletttes); force to 4bpp
    orient_WantPaletteInArray = 64  ' returned array from color reduction will include palette entries; no alpha
    orient_BlackIs1Not0 = 128       ' for WMFs. If image colors are vbBlack & vbWhite, white becomes transparent. Prevent that
    orient_PNMformat = 256          ' pnm cannot use palettized images except for grayscale
End Enum
Public Enum SaveAsMedium
    saveTo_File = 0
    saveTo_Array = 1
    saveTo_GDIplus = 2
    saveTo_stdPicture = 3
    saveTo_Clipboard = 4
    saveTo_DataObject = 5
    saveTo_GDIhandle = 6
End Enum
Public Enum TIFFMultiPageActions    ' see SaveAsTIFF
    TIFF_SingleFrame = 0
    TIFF_MultiFrameStart = 1
    TIFF_MultiFrameAdd = 2
    TIFF_MultiFrameEnd = 3
End Enum
Public Type ColorPalette            ' GDI+ palette object
   flags As Long
   Count As Long
   Entries(1 To 256) As Long
End Type
Public Type RECTF                   ' GDI+ rectangle w/Single vartypes
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type
Public Type RECTI                   ' GDI+ rectangle w/Long vartypes
    nLeft As Long
    nTop As Long
    nWidth As Long
    nHeight As Long
End Type
Public Type POINTAPI                ' GDI structure
    X As Long
    Y As Long
End Type
Public Type SafeArrayBound          ' OLE structure
    cElements As Long
    lLbound As Long
End Type
Public Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' reusable UDT for 1 & 2 dim arrays
End Type
Public Type BITMAPINFOHEADER        ' GDI structure
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
Public Enum LockModeConstants
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum
Public Type BitmapData              ' GDI+ lock/unlock bits structure
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    Scan0Ptr As Long
    ReservedPtr As Long
End Type
Private Type EncoderParameter       ' GDI+ image format encoding structure
    GUID(0 To 3)   As Long
    NumberOfValues As Long
    Type           As Long
    Value          As Long
End Type
'-- Encoder Parameters structure
Private Type EncoderParameters
    Count     As Long
    Parameter(0 To 5) As EncoderParameter
End Type
Private Type ImageCodecInfo         ' GDI+ codec structure
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type
Private Type FILEDESCRIPTOR         ' Clipboard/DataObject structure
  dwFlags As Long
  clsid(0 To 3) As Long
  sizeL As POINTAPI                 ' really SIZEL struct which same size as POINTAPI
  pointL As POINTAPI                ' really POINTL struct which same size as POINTAPI
  dwFileAttributes As Long
  ftCreationTime As POINTAPI        ' really FILETIME struct which same size as POINTAPI
  ftLastAccessTime As POINTAPI      ' really FILETIME struct which same size as POINTAPI
  ftLastWriteTime As POINTAPI       ' really FILETIME struct which same size as POINTAPI
  nFileSizeHigh As Long
  nFileSizeLow As Long
  cFileName(0 To 519) As Byte       ' maxpath*2 for unicode systems (not Win9x)
End Type
Private Type FORMATETC              ' Clipboard/DataObject structure
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lIndex As Long
    TYMED As Long
End Type
Private Type DROPFILES              ' Clipboard/DataObject structure
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type
Private Type STGMEDIUM              ' Clipboard/DataObject structure
    TYMED As Long
    Data As Long
    pUnkForRelease As Long
End Type

'//// usercontrol-only enumerations & structures
Public Enum RenderFlagsEnum        ' Rendering can be paused by user; need to log actions for when rendering is resumed
    render_PrePost = &H1            ' user wants pre & post Paint events
    ' render_RESERVED = &H2         ' reserved for future property
    render_FastRedraw = &H4         ' cached bitmap exists
    render_Shown = &H8              ' control displayed
    render_NoRedraw = &H10          ' do not force refresh until flag reset or usercontrol Refreshed
    render_RedoAutoSize = &H20      '
    render_DoFastRedraw = &H40      ' the FastRedraw image needs to be redone
    render_DoResize = &H80          ' AutoSize and/or Aspect properties changed; control may need to be resized
    render_DoReScale = &H100        ' Other properties/events occured which may require image to be re-scaled
    render_DoHitTest = &H200        ' Requires hit test area to be recalculated
    render_RemoteDrawBkg = &H400    ' flag used during Save/Paint Control/Image As Drawn
    render_RemoteToHGraphics = &H800 ' flag used during Save/Paint Control/Image As Drawn
    render_ShowEvent = &H1000       ' flag indicating the Show event triggered
    render_InitAnimation = &H2000   ' flag indicating animation should begin
    render_AutoSizing = &H4000      ' flag used to prevent recursion when autosizing
End Enum
Public Enum ContainerFlagsEnum
    cnt_AutoSizeMask = &H3          ' AutoSize mask (0=none,2=SingleAngle,3=MultiAngle)
    cnt_AlignCenter = &H4           ' anchored centered
    cnt_Opaque = &H8                ' opaque background uses UserControl.BackColor
    cnt_Border = &H10               ' 1-pixel rectangular border uses UserControl.ForeColor
    cnt_Runtime = &H20              ' container is in runtime else design time
    cnt_MouseDown = &H40            ' left mouse held down over the control
    cnt_RoundBorder = &H80          ' &H100=RoundBdrRough, &H180=Transborder
    cnt_BorderMask = &H180          '
    cnt_GradHorizontal = &H200      ' &H400=Vertical; &H600=Diagonal TL>BR; &H800=Diagonal BL>TR
    cnt_GradientMask = &HE00&       ' mask for gradient style
    cnt_BkgStretch = &H1000&
    cnt_KeyProps = &H80003FFF       ' mask for flags to cache to property bag (up to & including cnt_initLoad-1)
    cnt_InitLoad = &H4000&          ' used during ReadProperties to prevent any PropertyChanged events from triggering
    cnt_MouseValidate = &H8000&     ' flag used to test mouse exit (see cMouseEvent)
    cnt_DatabaseImage = &H400000    ' flag indicating image coming from databound record change & trigger update during Picture SET
    cnt_LastButtonShift = &H1000000 ' contains last button(s) held down on the control
    cnt_DblClicked = &H80000000     ' flags whether dblclick event occurred. See usercontrol's DblClick, MouseDown, MouseUp for more info
End Enum
Public Enum AttributeFlagsEnum
    attr_StretchMask = &H7          ' contains Scaling setting (0-5)
    attr_OLEDragModeShift = &H8     ' bitwise shift (8=dragmodeAuto else manual) OLEDragMOde
    attr_OLEDropShift = &H10        ' bitwise shift (0=none,16=manual,32=auto) OLEDropMode
    attr_OLEDropMask = &H30         ' oleDropMode mask
    attr_HitTestShift = &H40        ' bitwise shift (0=control,64=image,128=imageTrim,256=user-defined) hitTest
    attr_HitTestMask = &H1C0        ' hitTest mask
    attr_MouseEvents = &H200        ' mouse events + enter/exit wanted
    attr_AutoAnimate = &H400        ' image will begin animating when loaded (runtime only)
    attr_KeyProps = &H2FFFFF        ' mask for flags to cache to property bag (up to & including &H200000)
End Enum
Public Type ScalerStruct            ' cached scaled sizes to prevent continuous resizing calculations
    Width As Long                   ' scaled width
    Height As Long                  ' scaled height
    FRdrWidth As Long               ' width of rendered image, including rotation. Used with FastDraw
    FRdrHeight As Long              ' height of rendered image, including rotation. Used with FastDraw
    FixedCx As Long                 ' fixed width. See Aspect property
    FixedCy As Long                 ' fixed height. See Aspect property
    One2One As Boolean              ' if true, then image is actual size or clipped. Faster repaints
    DestDC As Long                  ' set when painting to alternate DC
End Type
Public Type DragDropStruct
    AutoDragPts As POINTAPI         ' X,Y coords to help determine when auto-dragging should start
    Originator As Boolean           ' if dragging, then set to True and prevents dropping contents on itself
    Effect As OLEDropEffectConstants ' see UserControl_OLEDragOver
End Type
'//////////////////////////////////////////////////////////////////

Public Const png_Signature1 As Long = 1196314761    ' PNG signature is 8 bytes
Public Const png_Signature2 As Long = 169478669
Public Const UnitPixel As Long = 2&                 ' GDI+ measurement
Public Const CF_UNICODE As Long = 13&               ' standard unicode text clipboard format
Public Const lvicPicTypeFromBinaries = 30&          ' flag used to identify DLL/EXE image extraction
Public Const lvicPicTypeAsyncDL = 32&               ' flag used when downloading files asynchronously
Public g_UnicodeSystem As Boolean                   ' cached, used often
Public g_ClipboardFormat As Integer                 ' custom clipboard format used for dragging/dropping
Private CF_FILECONTENTS As Long                     ' shell clipboard format (see CreateCustomClipboardFormat)
Public CF_FILEGROUPDESCRIPTORW As Long              ' shell clipboard format (see CreateCustomClipboardFormat)

Public g_TokenClass As cGDIpToken               ' publicly shared class (only 1 instance ever created)
Public g_MouseExitClass As cMouseExit           ' publicly shared class (only 1 instance ever created)
Public g_NewImageData As cGDIpMultiImage        ' staging area between LoadImage & assignment to GDIpImage
Public g_AsyncController As cAsyncController    ' async download manager (only 1 instance ever created)

Public Sub CreateCustomClipboardFormat()

    ' routine creates a clipboard format that the control's use to transfer image data back & forth
    Dim tFormat As Long, sFormat As String, iFormat As Integer
    If g_ClipboardFormat = 0 Then
        ' Create custom clipboard format for OLEStartDrag & OLEDragDrop
        ' convert long to signed integer for VB's usage & SetClipboardCustomFormat which uses Integer
        For iFormat = 1 To 3
            Select Case iFormat
                Case 1: sFormat = "LaVolpeImgCtrlData"
                Case 2: sFormat = "FileContents"
                Case 3: sFormat = "FileGroupDescriptorW"
            End Select
            
            tFormat = RegisterClipboardFormat(sFormat)
            Select Case tFormat
            Case Is < 0&: tFormat = 0&
            Case Is > &H7FFF&
                If tFormat < &H10000 Then tFormat = tFormat - &H10000 Else tFormat = 0&
            Case Else
            End Select
            
            Select Case iFormat
                Case 1: g_ClipboardFormat = tFormat
                Case 2: CF_FILECONTENTS = tFormat
                Case 3: CF_FILEGROUPDESCRIPTORW = tFormat
            End Select
        Next
    End If
    g_UnicodeSystem = IsWindowUnicode(GetDesktopWindow)
    
End Sub

Public Function Color_RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long

    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    Color_RGBtoARGB = (RGBColor And &HFF00&) Or ((RGBColor And &HFF&) * &H10000) Or ((RGBColor And &HFF0000) \ &H10000)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        Color_RGBtoARGB = Color_RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        Color_RGBtoARGB = Color_RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

Public Function Color_ARGBtoRGB(ByVal ARGBcolor As Long, Optional ByRef Opacity As Long) As Long

    ' This routine is the opposite of Color_RGBtoARGB
    ' Returned color is always RGB format, Opacity parameter will contain RGBAcolor opacity (0-255)

   If (ARGBcolor And &H80000000) Then
        Opacity = (ARGBcolor And Not &H80000000) \ &H1000000 Or &H80
    Else
        Opacity = (ARGBcolor \ &H1000000)
    End If
    Color_ARGBtoRGB = (ARGBcolor And &HFF00&) Or ((ARGBcolor And &HFF&) * &H10000) Or ((ARGBcolor And &HFF0000) \ &H10000)

End Function

Public Sub CommonTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal TimerID As Long, ByVal Tick As Long)
    ' common-use timer procedure.
    ' Current customers are the cAnimator & cMouseExit classes, nothing else
    Dim refObject As Object, unrefObject As Object
    
    CopyMemory unrefObject, TimerID, 4&         ' get unreferenced copy of calling class
    Set refObject = unrefObject                 ' reference the copy
    CopyMemory unrefObject, 0&, 4&              ' destroy unreferenced copy
    
    If TypeOf refObject Is Animator Then
        Call refObject.MoveForward              '   tell animator to move to next image
    
    ElseIf TypeOf refObject Is cAsyncClient Then
        KillTimer hWnd, TimerID                 ' window created for async download in GDIpImage
        refObject.Activate                      ' activate download
    
    ElseIf g_MouseExitClass.Owner = 0& Then     ' else can only be a cMouseExit class
        KillTimer hWnd, TimerID                 '   stop timer
    Else
        Call g_MouseExitClass.MouseInControl   ' tell mouse class to test current screen coords for hit test
    End If
    Set refObject = Nothing                     ' destroy referenced copy
End Sub

Public Function ByteAlignOnWord(ByVal BitDepth As Byte, ByVal Width As Long) As Long
    
    ' function to align any bit depth on dWord boundaries
    ByteAlignOnWord = (((Width * BitDepth) + &H1F&) And Not &H1F&) \ &H8&

End Function

Public Function ArrayToPicture(arrayVarPtr As Long, lSize As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty before this was called
    
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    On Error GoTo ExitRoutine
    Set IIStream = IStreamFromArray(arrayVarPtr, lSize)
    
    If Not IIStream Is Nothing Then
        aGUID(0) = &H7BF80980    ' GUID for stdPicture
        aGUID(1) = &H101ABF32
        aGUID(2) = &HAA00BB8B
        aGUID(3) = &HAB0C3000
        Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)
    End If
    
ExitRoutine:
End Function

Public Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As PictureTypeConstants) As IPicture

    ' function creates a stdPicture object from an image handle (bitmap or icon)
    
    'Private Type PictDesc
    '    Size As Long
    '    Type As Long
    '    hHandle As Long
    '    lParam As Long       for bitmaps only: Palette handle
    '                         for WMF only: extentX (integer) & extentY (integer)
    '                         for EMF/ICON: not used
    'End Type
    
    Dim lpPictDesc(0 To 3) As Long, aGUID(0 To 3) As Long
    
    lpPictDesc(0) = 16&
    lpPictDesc(1) = imgType
    lpPictDesc(2) = hImage
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc(0), aGUID(0), True, HandleToStdPicture)
    
End Function

Public Function NormalizeArray(inArray As Variant, outArray() As Byte, Optional WhatSizeOnly As Boolean, Optional ArrayPtr As Long) As Long

    ' When array data is passed to classes/controls from outside this project,
    ' we make a copy of it for our processing/caching when needed. As long as we
    ' have to copy it, we ensure it is 0-LBound and in bytes vs long.
    ' This routine converts byte/long arrays of any dimensions/bounds into
    ' a single dimensional 0-bound byte array

    Dim aVal As Long, aData As Long
    Dim tSA As SafeArray, boundInfo() As SafeArrayBound
    
    ' Valiate passed array is either Byte or Long or Integer
    CopyMemory aVal, ByVal VarPtr(inArray), 2&                      ' get variant type
    If (aVal And vbArray) = 0& Then Exit Function
    Select Case (aVal And &H1F&)
        Case vbByte, vbLong, vbInteger                              ' do nothing, expecting these data types
        Case Else: Exit Function
    End Select
    
    CopyMemory aData, ByVal VarPtr(inArray) + 8&, 4&                ' get pointer to SafeArray structure
    If (aVal And VT_BYREF) = VT_BYREF Then                          ' if array is in Variant ByRef then...
        If aData Then CopyMemory aData, ByVal aData, 4&             ' get pointer to SafeArray structure
    End If
    If aData = 0& Then Exit Function                                ' uninitialized array
    
    On Error GoTo ExitRoutine
    CopyMemory tSA, ByVal aData, 16&                                ' get 16 bytes of SafeArray structure
    If tSA.cDims = 0 Or tSA.pvData = 0& Then Exit Function          ' array with no data, strange
    
    ReDim boundInfo(1 To tSA.cDims)                                 ' size Bounds array & get the bounds
    CopyMemory boundInfo(1), ByVal aData + 16&, tSA.cDims * 8&
    
    If tSA.cDims = 1 And boundInfo(1).lLbound = 0& And tSA.cbElements = 1& Then
        If WhatSizeOnly Then
            NormalizeArray = boundInfo(1).cElements
        Else
            outArray() = inArray                                    ' if already 1-D, 0-bound, byte array, done
        End If
    Else
        aVal = tSA.cbElements                                       ' calculate total bytes used by array
        For aData = 1& To tSA.cDims
            aVal = aVal * boundInfo(aData).cElements
        Next
        If WhatSizeOnly Then
            NormalizeArray = aVal
        Else
            ReDim outArray(0 To aVal - 1&)                          ' resize outArray & copy data
            CopyMemory outArray(0), ByVal tSA.pvData, aVal
        End If
    End If
    
ExitRoutine:
    If Err Then ' only error would be excessively huge array exceeding max Long value in elements
        Err.Clear ' or possibly out of memory for huge arrays
        Erase outArray()
    ElseIf WhatSizeOnly = False Then
        NormalizeArray = True
    Else
        ArrayPtr = tSA.pvData
    End If
    
End Function

Public Function IStreamFromArray(ArrayPtr As Long, length As Long) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, IStreamFromArray
    ElseIf length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, IStreamFromArray)
            End If
        End If
    End If
    
HandleError:
End Function

Public Function IStreamToArray(hStream As Long, arrayBytes() As Byte) As Boolean

    ' Return array of bytes contained in an IUnknown interface (stream)
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If hStream Then
        If GetHGlobalFromStream(ByVal hStream, o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    IStreamToArray = True
                End If
            End If
        End If
    End If
    
End Function

Public Function GetFileHandle(ByVal FileName As String, WriteMode As Boolean) As Long

    ' Function uses APIs to create a file handle to read/write files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    Const FILE_SHARE_WRITE As Long = &H2
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    
    Dim flags As Long, Access As Long
    Dim Disposition As Long, Share As Long
    
    If WriteMode Then
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        If g_UnicodeSystem Then
            flags = GetFileAttributesW(StrPtr(FileName))
            If (flags And FILE_ATTRIBUTE_READONLY) Then
                flags = FILE_ATTRIBUTE_NORMAL
                SetFileAttributesW StrPtr(FileName), flags
            End If
        Else
            flags = GetFileAttributes(FileName)
            If (flags And FILE_ATTRIBUTE_READONLY) Then
                flags = FILE_ATTRIBUTE_NORMAL
                SetFileAttributes FileName, flags
            End If
        End If
        If flags < 0& Then flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    
    Else
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    End If
    
    If g_UnicodeSystem Then
        GetFileHandle = CreateFileW(StrPtr(FileName), Access, Share, ByVal 0&, Disposition, flags, 0&)
    Else
        GetFileHandle = CreateFile(FileName, Access, Share, ByVal 0&, Disposition, flags, 0&)
    End If
    
End Function

Public Function DeleteFileEx(FileName As String) As Boolean

    ' Function uses APIs to delete files :: unicode supported

    If FileOrFolderExists(FileName) = False Then
        DeleteFileEx = True
        
    ElseIf g_UnicodeSystem Then
        If DeleteFileW(StrPtr(FileName)) = 0& Then              ' failed. Read-only?
            If (SetFileAttributesW(StrPtr(FileName), FILE_ATTRIBUTE_NORMAL) = 0&) Then
                DeleteFileEx = Not (DeleteFileW(StrPtr(FileName)) = 0&)
            End If
        Else
            DeleteFileEx = True
        End If
    ElseIf DeleteFileA(FileName) = 0& Then                      ' failed. Read-only?
        If (SetFileAttributes(FileName, FILE_ATTRIBUTE_NORMAL) = 0&) Then
            DeleteFileEx = Not (DeleteFileA(FileName) = 0&)
        End If
    Else
        DeleteFileEx = True
    End If

End Function

Public Function FileOrFolderExists(FileFolderName As String) As Boolean

    If g_UnicodeSystem Then
        FileOrFolderExists = Not (GetFileAttributesW(StrPtr(FileFolderName)) = INVALID_HANDLE_VALUE)
    Else
        FileOrFolderExists = Not (GetFileAttributes(FileFolderName) = INVALID_HANDLE_VALUE)
    End If
    
End Function

Public Sub MoveArrayToVariant(inVariant As Variant, inArray() As Byte, Mount As Boolean)
    
    ' Variants are used a bit in this project to allow functions to receive
    ' multiple variable types (objects, strings, handles, arrays, etc) in a single parameter.
    ' When arrays are passed, don't want to unnecessarily copy the array
    ' if a copy isn't needed. But setting one variant to another that contains
    ' arrays, copies are made. With large arrays, performance suffers.
    ' So... this routine moves an array in/out of a variant and vice versa
    ' without making a copy of the array. We're just swapping pointers
    
    ' When mounting array to variant, the inVariant parameter can contain anything or nothing
    ' When dismounting array from variant, the inVariant parameter MUST contain a return byte array
    
    Dim bDummy() As Byte, srcAddr As Long, dstAddr As Long
    
    If Mount Then                                               ' moving array to variant
        inVariant = bDummy()                                    ' ensure target contains null byte array
        CopyMemory dstAddr, ByVal VarPtr(inVariant), 2&         ' get that null array's pointer
        If (dstAddr And VT_BYREF) Then
            CopyMemory dstAddr, ByVal VarPtr(inVariant) + 8&, 4&
        Else
            dstAddr = VarPtr(inVariant) + 8&
        End If
        srcAddr = VarPtrArray(inArray)                          ' get source array's pointer
        
    Else                                                        ' moving variant's array to array
    
        Erase inArray()                                         ' ensure source is nulled out
        CopyMemory srcAddr, ByVal VarPtr(inVariant), 2&         ' get source array's pointer
        If (srcAddr And VT_BYREF) Then
            CopyMemory srcAddr, ByVal VarPtr(inVariant) + 8&, 4&
        Else
            srcAddr = VarPtr(inVariant) + 8&
        End If
        dstAddr = VarPtrArray(inArray)                          ' get target array's pointer
    End If
    
    CopyMemory ByVal dstAddr, ByVal srcAddr, 4&                 ' swap pointers
    CopyMemory ByVal srcAddr, 0&, 4&                            ' null arrays have a null SafeArray pointer
    
End Sub

Public Function CreateSourcelessHandle(ByVal Handle As Long, Optional Width As Long, Optional Height As Long) As Long

    ' If user-defined KeepOriginalFormat is set to false during LoadPictureGDIplus,
    ' we create a sourceless GDI+ object. In some cases, we will maintain original format
    ' data but create a sourceless image anyway, such as is the case with
    ' metafiles, TGA, PCX, and bitmaps in some cases
    
    ' GDI+ images are wrappers around an image source data. Sourceless images are
    ' 100% maintained by GDI+ and require no source data to hang around

    If Handle = 0& Then Exit Function
    
    Dim returnObject As Long, hGraphics As Long
    Dim dstBoundsI As RECTI, srcBoundsF As RECTF, dstBoundsF As RECTF
    Dim tBMPsrc As BitmapData, tBMPdst As BitmapData
    Dim cPal As ColorPalette, lType As Long
    Dim xRez As Single, yRez As Single
    
    GdipGetImageBounds Handle, srcBoundsF, UnitPixel
    lType = GetImageType(Handle)
    Select Case lType
        ' Metafiles are a bit iffy to convert to stand-alone images
        
        Case lvicPicTypeEMetafile, lvicPicTypeMetafile              ' get horizontal & vertical resolutions
            If (Width = 0& Or Height = 0&) Then
                If lType = lvicPicTypeMetafile Then
                    GdipGetImageHorizontalResolution Handle, xRez
                    GdipGetImageVerticalResolution Handle, yRez
                    If xRez = 0! Then xRez = 96!                            ' should not get zeroes, use default if so
                    If yRez = 0! Then yRez = 96!
                    dstBoundsF.nWidth = srcBoundsF.nWidth * 96! / xRez              ' convert sizes to pixels
                    dstBoundsF.nHeight = srcBoundsF.nHeight * 96! / yRez
                Else
                    dstBoundsF = srcBoundsF
                End If
            Else
                dstBoundsF.nWidth = Width: dstBoundsF.nHeight = Height
            End If
            ' create the target GDI+ image now & get its graphics handle
            If GdipCreateBitmapFromScan0(dstBoundsF.nWidth, dstBoundsF.nHeight, 0&, lvicColor32bppAlpha, ByVal 0&, returnObject) = 0& Then
                If GdipGetImageGraphicsContext(returnObject, hGraphics) = 0& Then
                    ' draw the metafile onto the bitmap
                    GdipDrawImageRectRect hGraphics, Handle, 0!, 0!, dstBoundsF.nWidth, dstBoundsF.nHeight, srcBoundsF.nLeft, srcBoundsF.nTop, srcBoundsF.nWidth, srcBoundsF.nHeight, UnitPixel, 0&, 0&, 0&
                    GdipDeleteGraphics hGraphics
                Else
                    GdipDisposeImage returnObject: returnObject = 0&
                End If
            End If
            
        Case Else
            GdipGetImagePixelFormat Handle, tBMPsrc.PixelFormat
            dstBoundsI.nHeight = srcBoundsF.nHeight: dstBoundsI.nWidth = srcBoundsF.nWidth
            ' create GDI+ image of same bitdepth & color format
            If GdipCreateBitmapFromScan0(dstBoundsI.nWidth, dstBoundsI.nHeight, 0&, tBMPsrc.PixelFormat, ByVal 0&, returnObject) = 0& Then
                ' open the source for reading to get pointer to data
                If GdipBitmapLockBits(Handle, dstBoundsI, ImageLockModeRead, tBMPsrc.PixelFormat, tBMPsrc) = 0& Then
                    ' open the destination for writing using the pointer we just got from source
                    tBMPdst = tBMPsrc
                    If GdipBitmapLockBits(returnObject, dstBoundsI, ImageLockModeWrite Or ImageLockModeUserInputBuf, tBMPdst.PixelFormat, tBMPdst) = 0& Then
                        GdipBitmapUnlockBits returnObject, tBMPdst                  ' done
                        ' for paletted images, need to transfer palette also
                        If tBMPsrc.PixelFormat <= lvicColor8bpp Then
                            GdipGetImagePaletteSize Handle, cPal.Count
                            If lType Then
                                GdipGetImagePalette Handle, cPal, cPal.Count
                                GdipSetImagePalette returnObject, cPal
                            End If
                        End If
                    Else
                        GdipDisposeImage returnObject: returnObject = 0&
                    End If
                    GdipBitmapUnlockBits Handle, tBMPsrc                        ' unlock source now
                Else
                    GdipDisposeImage returnObject: returnObject = 0&
                End If
            End If
    End Select
    CreateSourcelessHandle = returnObject

End Function

Public Sub TrimImage(Handle As Long, Bounds As RECTI, ByVal TransColor As Long)

    ' routine creates a tight rectangle around passed image, trimming off alpha pixels
    '   outside of that image. The image is not actually modified, but its dimensions are

    Dim tBMPsrc As BitmapData, tSizeF As RECTF
    Dim tData() As Long, tSA As SafeArray, X As Long, Y As Long, z As Long
    
    If Handle = 0& Then Exit Sub
    
    GdipGetImageBounds Handle, tSizeF, UnitPixel
    With tSizeF
        Bounds.nHeight = .nHeight
        Bounds.nWidth = .nWidth
    End With
    ' attempt to get bits in premultiplied format. Premultiplied is preferred because
    '   all transparent pixels have a value of 0&. If not premultiplied, then the only
    '   thing guaranteed about transparency is that the Alpha byte is zero, RGB can be anything
    ' The TransColor parameter will be in the format of &HFFrgb. This is the optional usercontrol
    '   setting for transparent color throughout the image
    
    If GdipBitmapLockBits(Handle, Bounds, ImageLockModeRead, lvicColor32bppAlphaMultiplied, tBMPsrc) = 0& Then
        If tBMPsrc.PixelFormat <> lvicColor32bppAlphaMultiplied Or tBMPsrc.Scan0Ptr = 0& Then Exit Sub
        
        With tSA                                                ' create an array overlay on the bits
            .cbElements = 4
            .cDims = 2
            .pvData = tBMPsrc.Scan0Ptr
            If tBMPsrc.stride < 0& Then .pvData = .pvData + (Bounds.nHeight * tBMPsrc.stride) + 1&
            .rgSABound(0).cElements = Bounds.nHeight
            .rgSABound(1).cElements = Bounds.nWidth
        End With
        CopyMemory ByVal VarPtrArray(tData), VarPtr(tSA), 4&
        
        ' test for top of image
        Bounds.nLeft = Bounds.nWidth                            ' force out of bounds values
        For Y = 0& To Bounds.nHeight - 1&
            If tBMPsrc.stride < 0& Then z = Bounds.nHeight - Y - 1& Else z = Y
            For X = 0& To Bounds.nWidth - 1&
                If (tData(X, z) <> TransColor) Then             ' else fully transparent pixel, modify Bounds.Top/Left
                    Bounds.nTop = Y
                    Bounds.nLeft = X
                    Y = Bounds.nHeight
                    Exit For
                End If
            Next
        Next
        If Bounds.nLeft = Bounds.nWidth Then                    ' completely transparent image; return entire area
            Bounds.nLeft = 0&
            Exit Sub
        End If
        ' test for left of image
        If Bounds.nLeft Then
            For Y = Bounds.nTop To Bounds.nHeight - 1&
                For X = 0& To Bounds.nLeft - 1&
                    If (tData(X, Y) <> TransColor) Then          ' else fully transparent pixel, modify Bounds.Left
                        Bounds.nLeft = X
                        If X = 0& Then Y = Bounds.nHeight
                        Exit For
                    End If
                Next
            Next
        End If
        ' test for right of image
        z = Bounds.nLeft + 1&
        If tBMPsrc.stride > 0& Then Y = Bounds.nTop Else Y = 0&
        For Y = Y To Bounds.nHeight - 1&
            For X = Bounds.nWidth - 1& To z Step -1&
                If (tData(X, Y) <> TransColor) Then            ' else fully transparent pixel, modify Bounds.Left
                    z = X + 1&
                    If z = Bounds.nWidth Then Y = Bounds.nHeight
                    Exit For
                End If
            Next
        Next
        Bounds.nWidth = z - Bounds.nLeft
        ' finally test for bottom of image
        For Y = Bounds.nHeight - 1& To Bounds.nTop Step -1
            If tBMPsrc.stride < 0& Then z = Bounds.nHeight - Y - 1& Else z = Y
            For X = Bounds.nLeft To Bounds.nLeft + Bounds.nWidth - 1&
                If (tData(X, z) <> TransColor) Then
                    Bounds.nHeight = Bounds.nHeight - (Bounds.nHeight - Y) - Bounds.nTop + 1&
                    Y = 0&
                    Exit For
                End If
            Next
        Next
        
        CopyMemory ByVal VarPtrArray(tData), 0&, 4&             ' clean up
        GdipBitmapUnlockBits Handle, tBMPsrc
        
    End If

End Sub

Public Function GetDroppedFileNames(OLEDragDrop_DataObject As Variant, ByVal cfType As ClipBoardConstants, _
                            Optional DroppedText As String, Optional cImageData As cGDIpMultiImage, Optional CacheData As Boolean) As Long

    ' Function ensures the passed Data object, upon return, contains unicode filenames as needed,
    ' and is designed to be called from your OLEDragDrop event
    
    ' The function will return number of files dropped and the passed data
    ' object will contain valid unicode/ansi filenames as appropriate.
    
    ' Note: Updated to include getting dropped ANSI/Unicode text
    ' If retrieving text, set cfType=vbCFText & pass the optional DroppedText parameter
    
    If OLEDragDrop_DataObject Is Nothing Then Exit Function
    If Not (TypeOf OLEDragDrop_DataObject Is DataObject) Then Exit Function
    ' only support ANSI/Unicode text & files
    If cfType = vbCFText Then
        If OLEDragDrop_DataObject.GetFormat(CF_UNICODE) = True And g_UnicodeSystem = True Then
            cfType = CF_UNICODE
        Else
            If OLEDragDrop_DataObject.GetFormat(vbCFText) = True Then
                DroppedText = OLEDragDrop_DataObject.GetData(vbCFText)
                GetDroppedFileNames = True
            End If
            Exit Function
        End If
    End If
    
    Dim fmtEtc As FORMATETC, pMedium As STGMEDIUM, fMedium As STGMEDIUM
    Dim dFiles As DROPFILES, fd As FILEDESCRIPTOR
    Dim Vars(0 To 1) As Variant, pVars(0 To 1) As Long, pVartypes(0 To 1) As Integer
    Dim varRtn As Variant
    Dim iFiles As Long, iCount As Long, hDrop As Long
    Dim lLen As Long, sFiles() As String, bText() As Byte
    
    Dim IID_IDataObject As Long ' IDataObject Interface ID
    Const IDataObjVTable_GetData As Long = 12 ' 4th vtable entry
    Const IUnknown_Release As Long = 8&       ' 3rd vtable entry for IUnknown
    Const CC_STDCALL As Long = 4&
    Const TYMED_HGLOBAL = 1
    Const DVASPECT_CONTENT = 1

    With fmtEtc
        .cfFormat = cfType            ' same as CF_DROP
        .lIndex = -1                  ' want all data
        .TYMED = TYMED_HGLOBAL        ' want global ptr to files
        .dwAspect = DVASPECT_CONTENT  ' no rendering
    End With

    ' The IDataObject pointer is 16 bytes after VBs DataObject
    CopyMemory IID_IDataObject, ByVal ObjPtr(OLEDragDrop_DataObject) + 16, 4&
    
    ' Since we know the objPtr of the IDataObject interface, we therefore know
    ' the beginning of the interface's VTable
    ' So, if we know the VTable address and we know which function index we want
    ' to call, we can call it directly using the following OLE API. Otherwise we
    ' would need to use a TLB to define the IDataObject interface since VB doesn't
    ' expose it. This has some really neat implications if you think about it.
    ' The IDataObject function we want is GetData which is the 4th function in
    ' the VTable... ACKNOWLEDGEMENT: http://msdn2.microsoft.com/en-us/library/ms688421.aspx
    
    pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
    pVartypes(1) = vbLong: Vars(1) = VarPtr(pMedium): pVars(1) = VarPtr(Vars(1))

    On Error GoTo ExitRoutine
    ' The variants are required by the OLE API: http://msdn2.microsoft.com/en-us/library/ms221473.aspx
    If DispCallFunc(IID_IDataObject, IDataObjVTable_GetData, CC_STDCALL, _
                        vbLong, 2&, pVartypes(0), pVars(0), varRtn) = 0& Then
        
        If pMedium.Data = 0 Then
            Exit Function ' nothing to do
            
        ElseIf cfType = (CF_FILEGROUPDESCRIPTORW And &HFFFF&) Then
            
            ' dropping from a zip file accessed via Windows compressed folder
            GetDroppedFileNames = pvProcessCompressedFile(IID_IDataObject, pMedium.Data, cImageData, CacheData)
        
        ElseIf cfType = CF_UNICODE Then                     ' unicode text
            CopyMemory hDrop, ByVal pMedium.Data, 4&
            lLen = lstrlenW(ByVal hDrop)
            If lLen Then
                DroppedText = Space$(lLen)
                CopyMemory ByVal StrPtr(DroppedText), ByVal hDrop, lLen * 2&
                GetDroppedFileNames = True
            End If

        ElseIf cfType = vbCFFiles Then
            ' we have a pointer to the files, kinda sorta
            CopyMemory hDrop, ByVal pMedium.Data, 4&
            If hDrop Then
                ' the hDrop is a pointer to a DROPFILES structure
                ' copy the 20-byte structure for our use
                CopyMemory dFiles, ByVal hDrop, 20&
            End If
        
            If dFiles.fWide Then        ' else ansi & nothing to do
                
                ' use the pFiles member to track offsets for file names
                dFiles.pFiles = dFiles.pFiles + hDrop
                ReDim sFiles(1 To OLEDragDrop_DataObject.Files.Count)
            
                For iCount = 1& To UBound(sFiles)
                    ' get the length of the current filename & multiply by 2 because it is unicode
                    ' lstrLenW is supported in Win9x
                    lLen = lstrlenW(ByVal dFiles.pFiles) * 2&
                    sFiles(iCount) = String$(lLen \ 2&, vbNullChar)    ' build a buffer to hold the file name
                    CopyMemory ByVal StrPtr(sFiles(iCount)), ByVal dFiles.pFiles, lLen ' populate the buffer
                    ' move the pointer to location for next file, adding 2 because of a double null separator/delimiter btwn file names
                    dFiles.pFiles = dFiles.pFiles + lLen + 2&
                Next
                
                OLEDragDrop_DataObject.Files.Clear
                For iCount = 1& To iCount - 1&
                    OLEDragDrop_DataObject.Files.Add sFiles(iCount), iCount
                Next
                
            End If
            
            GetDroppedFileNames = iCount - 1&
            
        End If
    End If
    
ExitRoutine:
    If pMedium.Data Then
        If pMedium.pUnkForRelease = 0& Then GlobalFree pMedium.Data
    End If
End Function

Private Function pvProcessCompressedFile(IDataObject As Long, srcDataPtr As Long, cImageData As cGDIpMultiImage, CacheData As Boolean) As ImageFormatEnum

    ' helper function for pvProcessObjectSource

    ' On XP and above, ability to view zipped/compressed files as folders has been added.
    ' However, drag/drop & copy/paste does not return vbCfFiles when
    '   VB's Clipboard.GetFormat() or Data.GetFormat() functions are used.
    ' When dragging or copying out of those zips, accessed as Explorer folders, we will
    '   go low-level COM to get that information...

    ' The variants are required by the OLE API: http://msdn2.microsoft.com/en-us/library/ms221473.aspx
    Dim fmtEtc As FORMATETC, fMedium As STGMEDIUM
    Dim fd As FILEDESCRIPTOR, pMedium As STGMEDIUM
    Dim Vars(0 To 2) As Variant, pVars(0 To 2) As Long, pVartypes(0 To 2) As Integer
    Dim varRtn As Variant
    Dim iCount As Long, hPtr As Long
    Dim lLen As Long, sFile As String, bData() As Byte
    
    Const IDataObjVTable_GetData As Long = 12&  ' 4th vtable entry for IDataObject
    Const IUnknown_Release As Long = 8&         ' 3rd vtable entry for IUnknown
    Const IStream_Read As Long = 12&            ' 4th vtable entry for IStream
    Const CC_STDCALL As Long = 4&
    Const TYMED_HGLOBAL = 1&
    Const TYMED_ISTREAM = 4&
    Const DVASPECT_CONTENT = 1&

    If (srcDataPtr Or IDataObject) = 0& Then    ' will be zeroes when called from GetPastedFileData
        OleGetClipboard IDataObject             ' get data object placed on clipboard & abort if fails
        If IDataObject = 0& Then Exit Function
        pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
        pVartypes(1) = vbLong: Vars(1) = VarPtr(pMedium): pVars(1) = VarPtr(Vars(1))
        With fmtEtc                             ' get what was copied to clipboard
            .cfFormat = &HFFFF& And CF_FILEGROUPDESCRIPTORW
            .lIndex = -1&                 ' want all data
            .TYMED = TYMED_HGLOBAL        ' want global ptr to files
            .dwAspect = DVASPECT_CONTENT  ' no rendering
        End With
        If DispCallFunc(IDataObject, IDataObjVTable_GetData, CC_STDCALL, _
                            vbLong, 2&, pVartypes(0), pVars(0), varRtn) Then Exit Function
    Else
        pMedium.Data = srcDataPtr               ' called from GetDroppedFileNames
    End If

    On Error GoTo ExitRoutine
    If pMedium.Data = 0& Then GoTo ExitRoutine  ' validate & lock pointer
    hPtr = GlobalLock(pMedium.Data)
    If hPtr = 0& Then GoTo ExitRoutine
    
    CopyMemory iCount, ByVal hPtr, 4&          ' get number of files dropped/pasted
    With fmtEtc
        .cfFormat = &HFFFF& And CF_FILECONTENTS
        .lIndex = -1&                 ' want all data
        .TYMED = TYMED_ISTREAM        ' want results to stream
        .dwAspect = DVASPECT_CONTENT  ' no rendering
    End With
    pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
    pVartypes(1) = vbLong: Vars(1) = VarPtr(fMedium): pVars(1) = VarPtr(Vars(1))
    pVartypes(2) = vbLong: Vars(2) = VarPtr(lLen): pVars(2) = VarPtr(Vars(2))
    hPtr = hPtr + 4&                            ' move pointer to begining of filedescriptors
            
    For iCount = 0& To iCount - 1&              ' start the looping
        CopyMemory fd, ByVal hPtr, LenB(fd)     ' get next file descriptor
        hPtr = hPtr + LenB(fd)                  ' move pointer to next file descriptor
        If fd.nFileSizeLow > 0& And fd.nFileSizeHigh = 0& Then
            ' not a folder (0 bytes) and shouldn't be to large of a file with high size @ 0
            
            lLen = lstrlenW(VarPtr(fd.cFileName(0))) ' get length of filename
            If lLen Then                            ' create filename in VB string
                sFile = String$(lLen, 0)
                CopyMemory ByVal StrPtr(sFile), fd.cFileName(0), lLen * 2&
                If InStr(sFile, "\") = 0 Then       ' else a file in a subfolder; not processing subfolders
                
                    fmtEtc.lIndex = iCount          ' identify which file we are to retrieve & get it's stream
                    If DispCallFunc(IDataObject, IDataObjVTable_GetData, CC_STDCALL, vbLong, 2&, pVartypes(0), pVars(0), varRtn) = 0 Then
                        If (fMedium.TYMED And TYMED_ISTREAM) Then   ' sanity check
                            
                            ReDim bData(0 To fd.nFileSizeLow - 1&)  ' prep array to transfer from IStream
                            Vars(1) = fd.nFileSizeLow
                            Vars(0) = VarPtr(bData(0))              ' tell IStream to transfer to our array
                            If DispCallFunc(fMedium.Data, IStream_Read, CC_STDCALL, vbLong, 3&, pVartypes(0), pVars(0), varRtn) = 0& Then
                                If lLen = fd.nFileSizeLow Then      ' validate correct amount of bytes transferred
                                    pvProcessCompressedFile = pvProcessArraySource(cImageData, bData(), CacheData, lvicPicTypeUnknown)
                                End If
                            End If
                            If fMedium.pUnkForRelease = 0& Then  ' are we to release the stream? Release it
                                Call DispCallFunc(fMedium.Data, IUnknown_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
                            End If
                            If pvProcessCompressedFile Then Exit For    ' done?
                            Vars(0) = VarPtr(fmtEtc)            ' adjust these for next IDataObjVTable_GetData call
                            Vars(1) = VarPtr(fMedium)
                        End If
                    End If
                End If
            End If
        End If
    Next
    
ExitRoutine:
    If hPtr Then GlobalUnlock pMedium.Data                      ' unlock the pointer
    If pMedium.Data <> srcDataPtr Then                          ' release data if we initialized it
        If pMedium.pUnkForRelease = 0& Then GlobalFree pMedium.Data ' and also release the clipboard IDataObject
        Call DispCallFunc(IDataObject, IUnknown_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
    End If

End Function

Public Function GetPastedFileData(theFiles As Collection, cfType As Long, Optional cImageData As cGDIpMultiImage, Optional CacheData As Boolean, Optional pastedText As String) As Long

    ' Purpose: support function for pasted unicode filenames.
    
    ' This function will return number of files placed on the clipboard.
    ' The returned collection will contain the unicode/ansi filenames as needed.

    Dim hDrop As Long, lPtr As Long
    Dim lLen As Long
    Dim iCount As Long
    Dim dFiles As DROPFILES
    Dim sFile As String

    On Error GoTo ExitRoutine
    If cfType = vbCFFiles Then
        ' Get handle to CF_HDROP if any:
        If OpenClipboard(0&) = 0 Then Exit Function
             
         hDrop = GetClipboardData(cfType)
         If Not hDrop = 0 Then   ' then copied/cut files exist in memory
             iCount = DragQueryFile(hDrop, -1&, vbNullString, 0)
             If iCount Then
                 ' the hDrop is a pointer to a DROPFILES structure
                 ' copy the 20-byte structure for our use
                 CopyMemory dFiles, ByVal hDrop, 20&
                 ' use the pFiles member to track offsets for file names
                 dFiles.pFiles = dFiles.pFiles + hDrop
                 Set theFiles = New Collection
                 
                 For iCount = 1& To iCount
                     If dFiles.fWide = 0 Then   ' ANSI text, use API to get file name
                        lLen = DragQueryFile(hDrop, iCount - 1&, vbNullString, 0&)   ' query length
                        sFile = String$(lLen, vbNullChar)                            ' set up buffer
                        DragQueryFile hDrop, iCount - 1&, sFile, lLen + 1&           ' populate buffer
                     Else
                        ' get the length of the current file & multiply by 2 because it is unicode
                        ' lstrLenW is supported in Win9x
                        lLen = lstrlenW(ByVal dFiles.pFiles) * 2&
                        sFile = String$(lLen \ 2&, vbNullChar)          ' build a buffer to hold the file name
                        CopyMemory ByVal StrPtr(sFile), ByVal dFiles.pFiles, lLen ' populate the buffer
                        ' move the pointer to location for next file, adding 2 because of a double null separator/delimiter btwn file names
                        dFiles.pFiles = dFiles.pFiles + lLen + 2&
                    End If
                 theFiles.Add sFile
                 Next
                 GetPastedFileData = iCount - 1&
             End If
         End If
         
    ElseIf cfType = CF_FILEGROUPDESCRIPTORW Then
            
        ' dropping from a zip file accessed via Windows compressed folder
        GetPastedFileData = pvProcessCompressedFile(0&, 0&, cImageData, CacheData)
        
    ElseIf cfType = vbCFText Then
    
        If Clipboard.GetFormat(CF_UNICODE) And g_UnicodeSystem = True Then ' unicode text
            If OpenClipboard(0&) Then
                hDrop = GetClipboardData(CF_UNICODE)
                If hDrop Then
                    lLen = GlobalSize(hDrop)
                    If lLen > 0& Then lPtr = GlobalLock(hDrop)
                    If lPtr Then
                        lLen = lstrlenW(ByVal lPtr)
                        If lLen Then
                            pastedText = Space$(lLen)
                            CopyMemory ByVal StrPtr(pastedText), ByVal lPtr, lLen * 2&
                        End If
                    End If
                    If lPtr Then GlobalUnlock hDrop
                End If
                CloseClipboard
            End If
        ElseIf Clipboard.GetFormat(vbCFText) Then
            pastedText = Clipboard.GetText
        End If
    
    End If
    
ExitRoutine:
    If cfType = vbCFFiles Then CloseClipboard

End Function

Public Function SaveAsPNG(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                            Optional SaveOptions As Variant, Optional MIS As Variant, Optional FrameNumber As Long) As Long

    ' saves image as a PNG
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
    
    ' Some notes from results of testing
    ' If image has any transparency, regardless of actual bit depth, GDI+ saves PNG in 32bpp format
    ' If image is paletted and no alpha, any unused palette entries must have high bit set else saved as 32bpp
    
    If SourceHandle = 0& Then Exit Function
    
    Static cFormat As cFunctionsPNG
    
    Dim bOK As Boolean, lResult As Long
    Dim tgtHandle As Long, SS As SAVESTRUCT, APS As MULTIIMAGESAVESTRUCT
    Dim uEncCLSID(0 To 3) As Long, tData() As Byte
    Dim IIStream As IUnknown, tmpPic As StdPicture, tmpDO As DataObject
    Dim cPal As cColorReduction, tGDIpImage As GDIpImage, tObject As Object
    Dim lMedium As Long, tVariant As Variant
    Const MimeType As String = "image/png"
    
    lMedium = returnMedium
    If IsMissing(MIS) Then
        If IsMissing(SaveOptions) Then                               ' will be missing when called from GDIpImage
            SS.RSS.FillColorARGB = Color_RGBtoARGB(vbWindowBackground, 255&)
        Else
            SS = SaveOptions
        End If
    Else
        If cFormat Is Nothing Then Set cFormat = New cFunctionsPNG
        If MIS.Images > 1& Then
            lMedium = saveTo_Array
            APS = MIS
        Else     ' conversion from GIF to APNG
            SS = SaveOptions
            CopyMemory tObject, SS.reserved2, 4&
            Set tGDIpImage = tObject
            CopyMemory tObject, 0&, 4&
            If tGDIpImage.ExtractImageData(tData()) Then
                Set tGDIpImage = Nothing
                If cFormat.ConvertGIF2APNG(tVariant, tData()) = True Then
                    Erase tData()
                    APS.Images = -1&                    ' flag indicating processed
                    SaveAsPNG = True
                End If
            End If
            Set cFormat = Nothing
            If APS.Images <> -1& Then Exit Function
        End If
    End If
    
    If APS.Images <> -1& Then
    
        If SS.ColorDepth > lvicNoColorReduction Then
            Set cPal = New cColorReduction
            If lMedium = saveTo_Clipboard Or lMedium = saveTo_DataObject Then SS.reserved1 = SS.reserved1 Or &H20000000
            If SS.ColorDepth = lvicDefaultReduction Then
                SS.reserved1 = SS.reserved1 Or &H10000000
                tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_Complex, SS)
            ElseIf SS.ColorDepth < lvicConvert_TrueColor24bpp Then
                tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_Complex, SS)
                If tgtHandle = 0& Then Exit Function                ' palettizer failed
            End If
            Set cPal = Nothing
        End If
        
        If tgtHandle = 0& Then
            If SS.reserved2 <> 0& And SS.ExtractCurrentFrameOnly = False And APS.Images = 0& And _
                lMedium <> saveTo_stdPicture And (SS.reserved1 And &HFF00&) \ &H100& = lvicPicTypePNG Then
                CopyMemory tObject, SS.reserved2, 4&
                Set tGDIpImage = tObject
                CopyMemory tObject, 0&, 4&
            End If
            tgtHandle = SourceHandle
        End If
        
        If lMedium = saveTo_stdPicture Then                    ' stdPicture doesn't support PNG; use bitmap
            If GdipCreateHBITMAPFromBitmap(tgtHandle, lResult, SS.RSS.FillColorARGB Or &HFF000000) = 0& Then
                Set tmpPic = HandleToStdPicture(lResult, vbPicTypeBitmap)
                If tmpPic Is Nothing Then
                    DeleteObject lResult
                Else
                    Set returnObject = tmpPic
                    SaveAsPNG = tmpPic.Handle
                End If
            End If
        Else
            If tGDIpImage Is Nothing Then
                If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then
                    Set IIStream = IStreamFromArray(0&, 0&)
                    If Not IIStream Is Nothing Then
                        bOK = (GdipSaveImageToStream(tgtHandle, IIStream, uEncCLSID(0), ByVal 0&) = 0&)
                    End If
                End If
            Else
                bOK = (tGDIpImage.ExtractImageData(tData) = True)
            End If
            If bOK Then
                If lMedium = saveTo_GDIplus Then
                    If tGDIpImage Is Nothing Then
                        If GdipLoadImageFromStream(ObjPtr(IIStream), SaveAsPNG) = 0& Then returnObject = IIStream
                    Else
                        Set tGDIpImage = LoadImage(tData(), , , True)
                        If tGDIpImage.Handle Then
                            SaveAsPNG = lMedium
                            Set returnObject = tGDIpImage
                        End If
                    End If
                Else
                    If tGDIpImage Is Nothing Then
                        bOK = IStreamToArray(ObjPtr(IIStream), tData)
                    Else
                        bOK = True
                    End If
                    If bOK Then
                        If lMedium = saveTo_Array Then
                            SaveAsPNG = UBound(tData) + 1&
                            If APS.Images < 2& Then MoveArrayToVariant returnObject, tData(), True
                        ElseIf lMedium = saveTo_File Then
                            WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
                            SaveAsPNG = (lResult > tData(0))
                        Else
                            GdipCreateHBITMAPFromBitmap tgtHandle, lResult, SS.RSS.FillColorARGB Or &HFF000000
                            If lResult Then
                                Set tmpPic = HandleToStdPicture(lResult, vbPicTypeBitmap)
                                If tmpPic Is Nothing Then
                                    DeleteObject lResult
                                Else
                                    SaveAsPNG = bOK
                                    If lMedium = saveTo_Clipboard Then
                                        Clipboard.SetData tmpPic
                                        If g_ClipboardFormat Then SetClipboardCustomFormat tData(), g_ClipboardFormat
                                    ElseIf lMedium = saveTo_DataObject Then
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
        End If
        
        If tgtHandle <> SourceHandle Then
            If tgtHandle Then GdipDisposeImage tgtHandle
        End If

    End If
    
    '//// APNG tweak related code
    If APS.Images > 1& Or APS.Images = -1& Then
    
        If APS.Images = -1& Then
            APS.Images = 1
        Else
            If FrameNumber = 0& Then
                With APS.Image(LBound(APS.Image))
                    ' by specs, 1st frame of APNG must be same size as overall canvas.
                    ' If it isn't, we'll recreate the image and place it over a canvas-sized blank image
                    If Not (APS.GIFOverview.WindowHeight = .SS.Height And APS.GIFOverview.WindowWidth = .SS.Width) Then
                        Dim srcBmpData As BitmapData, dstBmpData As BitmapData
                        Dim srcRect As RECTI, dstRect As RECTI
                        ' create a blank image of the overall canvas size
                        SS.Width = APS.GIFOverview.WindowWidth                          ' dimensions of canvas size
                        SS.Height = APS.GIFOverview.WindowHeight                        ' no fill color
                        SS.RSS.FillColorUsed = False: SS.RSS.FillBrushGDIplus_Handle = 0&
                        SS.ColorDepth = lvicConvert_TrueColor32bpp_ARGB                 ' 32bpp depth
                        Set tGDIpImage = LoadBlankImage(SS, True)                     ' create canvas sized image (bitmap)
                        If tGDIpImage.Handle Then
                            ' transfer bits from current frame to new image. Define areas for copying
                            dstRect.nHeight = .SS.Height: dstRect.nWidth = .SS.Width
                            srcRect = dstRect
                            dstRect.nLeft = .GIFFrameInfo.XOffset: dstRect.nTop = .GIFFrameInfo.YOffset
                            Set .Picture = LoadImage(tData(), , , True)                 ' create GDI+ image from processed array
                            Erase tData()                                               ' now transfer the image to canvas
                            If GdipBitmapLockBits(.Picture.Handle, srcRect, ImageLockModeRead, lvicColor32bppAlpha, srcBmpData) = 0& Then
                                dstBmpData = srcBmpData
                                If GdipBitmapLockBits(tGDIpImage.Handle, dstRect, ImageLockModeWrite Or ImageLockModeUserInputBuf, lvicColor32bppAlpha, dstBmpData) = 0& Then
                                    GdipBitmapUnlockBits tGDIpImage.Handle, dstBmpData
                                End If
                                GdipBitmapUnlockBits .Picture.Handle, srcBmpData
                            End If
                            SaveAsPNG tData(), tGDIpImage.Handle, saveTo_Array          ' convert bitmap to PNG
                            Set tGDIpImage = Nothing                                    ' no longer needed
                        Else 'failed, cannot create APNG
                            SaveAsPNG = 0&
                            Set cFormat = Nothing
                            Exit Function
                        End If
                    End If
                    .GIFFrameInfo.XOffset = 0&: .GIFFrameInfo.YOffset = 0&              ' reset; 1st frame doesn't have offsets
                End With
            End If
            MoveArrayToVariant tVariant, tData(), True                                  ' place processed array in variant
            SaveAsPNG = cFormat.SaveAsAPNG(tVariant, APS, FrameNumber + 1&)             ' pass off to PNG routines for appending
        End If
        If SaveAsPNG = 0& Then
            Set cFormat = Nothing
        ElseIf FrameNumber = APS.Images - 1& Then                                   ' last frame?
            Set cFormat = Nothing
            SaveAsPNG = 0&
            MoveArrayToVariant tVariant, tData(), False                             ' get processed APNG array bytes
            Select Case returnMedium                                                ' return in requested medium
                Case saveTo_GDIplus
                    Set tGDIpImage = modCommon.LoadImage(tData(), True, , True)
                    If tGDIpImage.Handle <> 0& Then
                        SaveAsPNG = returnMedium
                        Set returnObject = tGDIpImage
                    End If
                Case saveTo_Array
                    SaveAsPNG = UBound(tData) + 1&
                    MoveArrayToVariant returnObject, tData(), True
                Case saveTo_File
                    WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lResult, ByVal 0&
                    SaveAsPNG = (lResult > UBound(tData))
                Case saveTo_Clipboard, saveTo_DataObject
                    GdipCreateHBITMAPFromBitmap APS.Image(LBound(APS.Image)).Picture.Handle, lResult, SS.RSS.FillColorARGB Or &HFF000000
                    If lResult Then
                        Set tmpPic = HandleToStdPicture(lResult, vbPicTypeBitmap)
                        If tmpPic Is Nothing Then
                            DeleteObject lResult
                        Else
                            SaveAsPNG = bOK
                            If lMedium = saveTo_Clipboard Then
                                Clipboard.SetData tmpPic
                                If g_ClipboardFormat Then SetClipboardCustomFormat tData(), g_ClipboardFormat
                            ElseIf lMedium = saveTo_DataObject Then
                                Set tmpDO = returnObject
                                tmpDO.SetData tmpPic, vbCFBitmap
                                If g_ClipboardFormat Then tmpDO.SetData tData(), g_ClipboardFormat
                                Set tmpDO = Nothing
                            End If
                            Set tmpPic = Nothing
                        End If
                    End If
                    Erase tData()
            End Select
        End If
    End If

End Function

Public Function SaveAsBMP(returnObject As Variant, SourceHandle As Long, _
                        ByVal returnMedium As SaveAsMedium, Optional RenderingStyle As Variant) As Long

    ' saves image as a BMP
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    ' returnMedium.
    '   If saveTo_Array then returnObject is the 0-Bound array & function return value is size of array in bytes
    '   If saveTo_GDIpHandle then returnObject is handle's IStream source & function return value is the GDI+ handle
    '   If saveTo_GDIhandle then returnObject is HBITMAP and return value is non-zero if successful
    '   If saveTo_File then returnObject is passed HFILE, return value is non-zero if successful. File is not closed
    '   If saveTo_stdPictureture then returnObject is passed stdPicture, return value is non-zero if successful
    '   If saveTo_Clipboard then return value is non-zero if successful
    '   If saveTo_DataObject then returnObject is passed DataObject, return value is non-zero if successful
    
    Dim cFormat As cFunctionsBMP
    Set cFormat = New cFunctionsBMP
    SaveAsBMP = cFormat.SaveAsBMP(returnObject, SourceHandle, returnMedium, RenderingStyle)
    Set cFormat = Nothing

End Function

Private Function SaveAsAVI(returnObject As Variant, SourceHandle As Long, _
                        ByVal returnMedium As SaveAsMedium, SS As SAVESTRUCT) As Long

    ' saves image as a BMP
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    ' returnMedium.
    '   If saveTo_Array then returnObject is the 0-Bound array & function return value is size of array in bytes
    '   If saveTo_GDIpHandle then returnObject is handle's IStream source & function return value is the GDI+ handle
    '   If saveTo_GDIhandle then returnObject is HBITMAP and return value is non-zero if successful
    '   If saveTo_File then returnObject is passed HFILE, return value is non-zero if successful. File is not closed
    '   If saveTo_stdPictureture then returnObject is passed stdPicture, return value is non-zero if successful
    '   If saveTo_Clipboard then return value is non-zero if successful
    '   If saveTo_DataObject then returnObject is passed DataObject, return value is non-zero if successful
    
    Dim cFormat As cFunctionsAVI
    Set cFormat = New cFunctionsAVI
    SaveAsAVI = cFormat.SaveAsAVI(returnObject, SourceHandle, returnMedium, SS)
    Set cFormat = Nothing

End Function

Private Function pvSaveAsICO(returnObject As Variant, SourceHandle As Long, returnMedium As SaveAsMedium, _
                            Optional SaveOptions As Variant, Optional MIS As Variant, Optional FrameNumber As Long) As Long

    ' saves image as a icon, cursor, multi-image icon/cursor & PNG encoding supported

    Static cFormat As cFunctionsICO
    Dim SS As SAVESTRUCT, ICS As MULTIIMAGESAVESTRUCT
    
    If IsMissing(SaveOptions) Then Exit Function         ' only called by the SaveImage routine
    If IsMissing(MIS) = False Then ICS = MIS
    If cFormat Is Nothing Then Set cFormat = New cFunctionsICO
    SS = SaveOptions
    pvSaveAsICO = cFormat.SaveAsICO(returnObject, SourceHandle, returnMedium, SS, ICS, FrameNumber)
    If pvSaveAsICO = 0& Then
        Set cFormat = Nothing
    ElseIf IsMissing(MIS) = True Then
        Set cFormat = Nothing
    ElseIf FrameNumber = ICS.Images - 1& Then
        Set cFormat = Nothing
    End If

End Function

Private Function pvSaveAsPCX(returnObject As Variant, SourceHandle As Long, returnMedium As SaveAsMedium, _
                            SaveOptions As SAVESTRUCT) As Long

    ' saves image as a PCX

    Dim cFormat As cFunctionsPCX
    Set cFormat = New cFunctionsPCX
    pvSaveAsPCX = cFormat.SaveAsPCX(returnObject, SourceHandle, returnMedium, SaveOptions)
    Set cFormat = Nothing

End Function

Private Function pvSaveAsPNM(returnObject As Variant, SourceHandle As Long, returnMedium As SaveAsMedium, _
                            SaveOptions As SAVESTRUCT, asPAM As Boolean) As Long

    ' saves image as a PNM: PBM,PGM,PPM

    Dim cFormat As cFunctionsPNM
    Set cFormat = New cFunctionsPNM
    pvSaveAsPNM = cFormat.SaveAsPNM(returnObject, SourceHandle, returnMedium, SaveOptions, asPAM)
    Set cFormat = Nothing

End Function

Private Function pvSaveAsTGA(returnObject As Variant, SourceHandle As Long, returnMedium As SaveAsMedium, _
                            Optional SaveOptions As Variant) As Long

    ' saves image as a TGA

    Dim cFormat As cFunctionsTGA
    Dim SS  As SAVESTRUCT
    Set cFormat = New cFunctionsTGA
    SS = SaveOptions
    pvSaveAsTGA = cFormat.SaveAsTGA(returnObject, SourceHandle, returnMedium, SS)
    Set cFormat = Nothing

End Function

Private Function pvSaveAsGIF(returnObject As Variant, SourceHandle As Long, returnMedium As SaveAsMedium, _
                            Optional SaveOptions As Variant, Optional MIS As Variant, Optional FrameNumber As Long) As Long

    ' saves image as a GIF, including animated GIF

    Static cFormat As cFunctionsGIF
    Dim uEncCLSID(0 To 3) As Long
    Dim SS As SAVESTRUCT, AGS As MULTIIMAGESAVESTRUCT
    Const MimeType As String = "image/gif"
    
    If IsMissing(SaveOptions) Then Exit Function         ' only called by the SaveImage routine
    If IsMissing(MIS) = False Then AGS = MIS
    If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then
        If cFormat Is Nothing Then Set cFormat = New cFunctionsGIF
        SS = SaveOptions
        pvSaveAsGIF = cFormat.SaveAsGIF(returnObject, SourceHandle, returnMedium, SS, VarPtr(uEncCLSID(0)), AGS, FrameNumber)
    End If
    If pvSaveAsGIF = 0& Then
        Set cFormat = Nothing
    ElseIf IsMissing(MIS) = True Then
        Set cFormat = Nothing
    ElseIf FrameNumber = AGS.Images - 1& Then
        Set cFormat = Nothing
    End If

End Function

Private Function pvSaveAsMetafile(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, _
                                 SaveOptions As SAVESTRUCT) As Long

    ' saves image as a WMF, EMF or WMF non-placeable
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
    
    ' Some notes from results of testing
    '   some quality is lost if saving to WMF and if source contains alphablending
    '   conversion from complex to simple transparency is automatically done
    '   VB's LoadPicture cannot load non-placeable WMFs. Automatically defaults to WMF in this case
    ' Cannot render alpha to WMF else unexpected results. Trick is to create clip region & convert source to 24bpp
    ' Black and White WMF can cause White to be made transparent, regardless of the color depth
    '   To patch the problem, any 1bpp image will be checked to see if 2 palette entries are black & white & fix black
    '   If more than 1bpp, then image will be default color reduced and problem fixed in the cColorReduction class
    
    If SourceHandle = 0& Then Exit Function
                    
    Dim tmpPic As StdPicture, tmpDO As DataObject
    Dim tgtFormat As SaveAsFormatEnum, bOK As Boolean
    Dim IIStream As IUnknown, sizeF As RECTF, lSize As Long
    Dim hEMF As Long, hWMF As Long, hDC As Long, hGraphics As Long
    Dim outArray() As Byte, tgtHandle As Long, lResult As Long
    Dim cColorReducer As cColorReduction, hClip As Long, hClipGDIp As Long
    Dim tmpGpic As GDIpImage, tObject As Object, cPal As ColorPalette
    Dim tBMP As BitmapData, tBMPsrc As BitmapData, tSize As RECTI
    
    Const MetafileTypeEmf As Long = 3&
    Const MM_ANISOTROPIC As Long = 8&
    
    tgtFormat = (SaveOptions.reserved1 And &HF0&) \ &H10
    
    If (SaveOptions.reserved1 And 1&) = 0& Then      ' 1bpp workaround; if b&w
        If ColorDepthToColorType(lvicNoColorReduction, SourceHandle) = lvicColor1bpp Then
            GdipGetImagePaletteSize SourceHandle, lResult
            If lResult = 16& Then         ' 2 colors + 2 flags
                GdipGetImagePalette SourceHandle, cPal, lResult
                If (cPal.Entries(1) = &HFF000000 And cPal.Entries(2) = &HFFFFFFFF) Then
                    If SaveOptions.reserved2 = 0& Then
                        cPal.Entries(1) = &HFF000001
                        GdipSetImagePalette SourceHandle, cPal
                        SaveOptions.ColorDepth = lvicNoColorReduction    ' no point; already 1bpp & fixed
                    Else
                        SaveOptions.ColorDepth = lvicConvert_16Colors    ' going to force 4bpp
                        SaveOptions.PaletteType = lvicPaletteAdaptive
                    End If
                Else
                    SaveOptions.ColorDepth = lvicNoColorReduction    ' no point; already 1bpp
                End If
            Else
                lResult = 0&
            End If
        End If
    End If
    If lResult = 0& Then
        If SaveOptions.ColorDepth = lvicNoColorReduction Then SaveOptions.ColorDepth = lvicDefaultReduction
    End If
    
    Set cColorReducer = New cColorReduction
    If SaveOptions.ColorDepth > lvicNoColorReduction Then
        If SaveOptions.ColorDepth < lvicConvert_TrueColor24bpp Then
            If SaveOptions.ColorDepth = lvicDefaultReduction Then
                SaveOptions.reserved1 = SaveOptions.reserved1 Or &H10000000
                tgtHandle = cColorReducer.PalettizeToHandle(SourceHandle, alpha_Simple, SaveOptions, orient_BlackIs1Not0)
            ElseIf SaveOptions.ColorDepth < lvicConvert_TrueColor24bpp Then
                tgtHandle = cColorReducer.PalettizeToHandle(SourceHandle, alpha_Simple, SaveOptions, orient_BlackIs1Not0)
                If tgtHandle = 0& Then Exit Function       ' palettizer failed
            End If
        End If
    End If
    If tgtHandle = 0& Then
        Select Case (SaveOptions.reserved1 And &HFF00&) \ &H100&
        Case lvicPicTypeMetafile
            If tgtFormat <> lvicSaveAsEMetafile Then tgtHandle = SourceHandle
        Case lvicPicTypeEMetafile
            If tgtFormat = lvicSaveAsEMetafile Then
                tgtHandle = SourceHandle
                hEMF = tgtHandle
            End If
        End Select
        If tgtHandle Then
            CopyMemory tObject, SaveOptions.reserved2, 4&
            Set tmpGpic = tObject
            CopyMemory tObject, 0&, 4&
            bOK = tmpGpic.ExtractImageData(outArray)
        Else
            tgtHandle = SourceHandle
        End If
    End If
    Set cColorReducer = Nothing
    If tgtHandle = 0& Then Exit Function
    
    On Error GoTo ExitRoutine
    If tmpGpic Is Nothing Then
        If (SaveOptions.reserved1 And 1&) = 0& Then
            bOK = True
        Else    ' has alpha & source format is not in destination format
            hClip = CreateShapedRegion(tgtHandle, SaveOptions.Width, SaveOptions.Height, SaveOptions.AlphaTolerancePct + 1&) ' create clipping region excluding alpha
            tSize.nWidth = SaveOptions.Width: tSize.nHeight = SaveOptions.Height    ' create new 24bpp image
            If GdipCreateBitmapFromScan0(tSize.nWidth, tSize.nHeight, 0&, lvicColor24bpp, ByVal 0&, lResult) = 0& Then
                If GdipBitmapLockBits(lResult, tSize, ImageLockModeWrite, lvicColor24bpp, tBMP) = 0& Then
                    tBMPsrc = tBMP                                                  ' copy 32bpp to 24bpp
                    bOK = (GdipBitmapLockBits(tgtHandle, tSize, ImageLockModeRead Or ImageLockModeUserInputBuf, lvicColor24bpp, tBMPsrc) = 0&)
                End If
                GdipBitmapUnlockBits lResult, tBMP
                If bOK Then                                                         ' replace/set new tgtHandle
                    GdipBitmapUnlockBits tgtHandle, tBMPsrc
                    If tgtHandle = SourceHandle Then
                        If SaveOptions.reserved2 = 0& Then GdipDisposeImage SourceHandle: SourceHandle = lResult
                        tgtHandle = lResult
                    Else
                        GdipDisposeImage tgtHandle: tgtHandle = lResult
                    End If
                Else
                    GdipDisposeImage lResult                                        ' failure
                End If
            End If
        End If
        If bOK Then
            Set IIStream = IStreamFromArray(0&, 0&)
            If Not IIStream Is Nothing Then
                sizeF.nHeight = SaveOptions.Height: sizeF.nWidth = SaveOptions.Width
                hDC = GetDC(GetDesktopWindow)
                If GdipRecordMetafileStream(ObjPtr(IIStream), hDC, MetafileTypeEmf, sizeF, UnitPixel, 0&, hEMF) = 0& Then
                    If GdipGetImageGraphicsContext(hEMF, hGraphics) = 0& Then
                        If hClip Then
                            GdipCreateRegionHrgn hClip, hClipGDIp
                            DeleteObject hClip: hClip = 0&
                            GdipSetClipRegion hGraphics, hClipGDIp, 0&
                            GdipDeleteRegion hClipGDIp: hClipGDIp = 0&
                        End If
                        GdipDrawImageRectRect hGraphics, tgtHandle, 0!, 0!, sizeF.nWidth, sizeF.nHeight, sizeF.nLeft, sizeF.nTop, sizeF.nWidth, sizeF.nHeight, UnitPixel, 0&, 0&, 0&
                        GdipDeleteGraphics hGraphics
                    Else
                        GdipDisposeImage hEMF: hEMF = 0&
                    End If
                End If
                ReleaseDC GetDesktopWindow, hDC
            End If
        End If
        If hEMF Then
            If tgtFormat = lvicSaveAsEMetafile Then
                bOK = True
            ElseIf GdipGetHemfFromMetafile(hEMF, hWMF) = 0& Then
                lSize = GdipEmfToWmfBits(hWMF, 0&, 0&, MM_ANISOTROPIC, 0&)
                GdipDisposeImage hEMF: hEMF = 0&
                Set IIStream = Nothing
                If lSize Then
                    ' add placeable header... ACKNOWLEDGEMENT: http://wvware.sourceforge.net/caolan/ora-wmf.html
                    ' Note that GDI+ can do this for you, but I find better results forcing it myself
                    ReDim outArray(0 To lSize + 21&)
                    If (GdipEmfToWmfBits(hWMF, lSize, VarPtr(outArray(22)), MM_ANISOTROPIC, 0&) <> 0&) Then
                        CopyMemory outArray(0), &H9AC6CDD7, 4&
                        CopyMemory outArray(10), CLng(SaveOptions.Width * Screen.TwipsPerPixelX), 2&
                        CopyMemory outArray(12), CLng(SaveOptions.Height * Screen.TwipsPerPixelY), 2&
                        CopyMemory outArray(14), 1440, 2&   ' ... calc checksum
                        lResult = 22289& Xor CLng(SaveOptions.Width * Screen.TwipsPerPixelX) Xor CLng(SaveOptions.Height * Screen.TwipsPerPixelY) Xor 1440&
                        CopyMemory outArray(20), lResult, 2&
                        bOK = True
                    End If
                End If
                DeleteEnhMetaFile hWMF
            End If
        End If
    End If
    
    If bOK Then
        If returnMedium = saveTo_GDIplus Then
            If tgtFormat = lvicSaveAsMetafile_NonPlaceable Then
                CopyMemory outArray(0), outArray(22), UBound(outArray) - 21&
                pvSaveAsMetafile = pvConvertNonPlaceableWMFtoWMF(outArray(), vbNullString)
                If pvSaveAsMetafile Then
                    Set g_NewImageData = New cGDIpMultiImage
                    g_NewImageData.CacheSourceInfo VarPtrArray(outArray), pvSaveAsMetafile, lvicPicTypeMetafile, True, False
                    Set tmpGpic = New GDIpImage
                    Set returnObject = tmpGpic
                    pvSaveAsMetafile = saveTo_GDIplus
                End If
            ElseIf hEMF = 0& Then                               ' WMF placeable
                Set tmpGpic = LoadImage(outArray(), , , True)
                If tmpGpic.Handle Then
                    Set returnObject = tmpGpic
                    pvSaveAsMetafile = returnMedium
                End If
            ElseIf tmpGpic Is Nothing Then                      ' EMF created manually above
                pvSaveAsMetafile = CreateSourcelessHandle(hEMF)
                GdipDisposeImage hEMF: hEMF = 0&
                If pvSaveAsMetafile Then Set returnObject = IIStream
            Else                                                ' EMF from source
                Set tmpGpic = LoadImage(outArray(), , , True)
                If tmpGpic.Handle Then
                    Set returnObject = tmpGpic
                    pvSaveAsMetafile = returnMedium
                End If
            End If
        Else
            If hEMF Then                                        ' get array data if needed
                If tmpGpic Is Nothing Then
                    bOK = IStreamToArray(ObjPtr(IIStream), outArray())
                    GdipDisposeImage hEMF: hEMF = 0&
                    Set IIStream = Nothing
                Else
                    bOK = True
                End If
            End If
            If bOK Then
                If returnMedium = saveTo_File Then
                    If tgtFormat = lvicSaveAsMetafile_NonPlaceable Then lResult = 22& Else lResult = 0&
                    WriteFile CLng(returnObject), outArray(lResult), UBound(outArray) - lResult + 1&, lResult, ByVal 0&
                    pvSaveAsMetafile = (lResult > UBound(outArray))
                ElseIf returnMedium = saveTo_Array Then
                    If tgtFormat = lvicSaveAsMetafile_NonPlaceable Then
                        CopyMemory outArray(0), outArray(22), UBound(outArray) - 21&
                        ReDim Preserve outArray(0 To UBound(outArray) - 22&)
                    End If
                    pvSaveAsMetafile = UBound(outArray) + 1&
                    MoveArrayToVariant returnObject, outArray, True
                Else
                    Set tmpPic = ArrayToPicture(VarPtr(outArray(0)), UBound(outArray) + 1&)
                    If Not tmpPic Is Nothing Then
                        If returnMedium = saveTo_stdPicture Then
                            Set returnObject = tmpPic
                        Else
                            If tgtFormat = lvicSaveAsMetafile_NonPlaceable Then
                                CopyMemory outArray(0), outArray(22), UBound(outArray) - 21&
                                ReDim Preserve outArray(0 To UBound(outArray) - 22&)
                            End If
                            If returnMedium = saveTo_Clipboard Then
                                If g_ClipboardFormat Then SetClipboardCustomFormat outArray(), g_ClipboardFormat
                                Clipboard.SetData tmpPic
                            ElseIf returnMedium = saveTo_DataObject Then
                                Set tmpDO = returnObject
                                If tgtFormat = lvicSaveAsEMetafile Then
                                    tmpDO.SetData tmpPic, vbCFEMetafile
                                Else
                                    tmpDO.SetData tmpPic, vbCFMetafile
                                End If
                                If g_ClipboardFormat Then tmpDO.SetData outArray(), g_ClipboardFormat
                                Set tmpDO = Nothing
                            End If
                        End If
                        pvSaveAsMetafile = True
                    End If
                    Erase outArray()
                End If
            End If
        End If
    End If
    
ExitRoutine:
    If hEMF Then
        If hEMF <> SourceHandle Then GdipDisposeImage hEMF
    End If
    If tgtHandle <> SourceHandle Then
        If tgtHandle Then GdipDisposeImage tgtHandle
    End If
    If hClipGDIp Then GdipDeleteRegion hClipGDIp
    If hClip Then DeleteObject hClip
End Function

Private Function pvSaveAsJPEG(returnObject As Variant, SourceHandle As Long, _
                           ByVal returnMedium As SaveAsMedium, _
                           SaveOptions As SAVESTRUCT) As Long

    ' saves image as a JPG
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
    
    ' Some notes from results of testing
    ' If image has any transparency, a fill color will be applied before saving PNG
    ' GDI+ seems to refuse to save JPG to any bit depth other than 24bpp
    
    If SourceHandle = 0& Then Exit Function
    If IsMissing(SaveOptions) Then Exit Function
    
    Dim tmpPic As StdPicture, tmpDO As DataObject
    Dim tObject As Object, tGDIpImage As GDIpImage
    Dim lValue As Long, srcDepth As Long, tgtHandle As Long
    Dim uEncCLSID(0 To 3) As Long, tData() As Byte
    Dim IIStream As IUnknown, bOK As Boolean
    Dim uEncParams As EncoderParameters, cPal As cColorReduction
    Dim JPEGQuality As Long
    Const EncoderParameterValueTypeLong As Long = &H4&
    Const JPGEncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    Const MimeType As String = "image/jpeg"
    
    If SaveOptions.ColorDepth > lvicNoColorReduction Then
        Set cPal = New cColorReduction
        If SaveOptions.ColorDepth < lvicConvert_TrueColor24bpp Then
            If SaveOptions.ColorDepth = lvicDefaultReduction Then
                SaveOptions.reserved1 = SaveOptions.reserved1 Or &H10000000
                tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_None, SaveOptions)
            ElseIf SaveOptions.ColorDepth < lvicConvert_TrueColor24bpp Then
                tgtHandle = cPal.PalettizeToHandle(SourceHandle, alpha_None, SaveOptions)
                If tgtHandle = 0& Then Exit Function           ' palettizer failed
            End If
        End If
        Set cPal = Nothing
    End If
    
    If tgtHandle = 0& Then
        If SaveOptions.CompressionJPGQuality = lvicDefaultCompressionQuality And SaveOptions.reserved2 <> 0& _
            And (SaveOptions.reserved1 And &HFF00&) \ &H100& = lvicPicTypeJPEG Then
                CopyMemory tObject, SaveOptions.reserved2, 4&
                Set tGDIpImage = tObject
                CopyMemory tObject, 0&, 4&
                bOK = tGDIpImage.ExtractImageData(tData)
                SaveOptions.CompressionJPGQuality = lvicJPGQuality100pct
        End If
        tgtHandle = SourceHandle
    End If
    
    If tGDIpImage Is Nothing Then
        If pvGetEncoderClsID(MimeType, uEncCLSID) = -1 Then Exit Function
        '-- Set encoder params. (Quality)
        Select Case SaveOptions.CompressionJPGQuality
            Case lvicDefaultCompressionQuality, lvicFormatCompressed: JPEGQuality = 80&
            Case lvicFormatUncompressed: JPEGQuality = 100&
            Case Is > 100&: JPEGQuality = 100&
            Case Is < lvicJPGQuality100pct: JPEGQuality = 100&
            Case Is > lvicDefaultCompressionQuality: JPEGQuality = SaveOptions.CompressionJPGQuality
            Case Else: JPEGQuality = (-SaveOptions.CompressionJPGQuality - 2&) * 10&
        End Select
        uEncParams.Count = 1
        ReDim aEncParams(1 To Len(uEncParams))
        With uEncParams.Parameter(0)
            .NumberOfValues = 1
            .Type = EncoderParameterValueTypeLong
            Call CLSIDFromString(StrPtr(JPGEncoderQuality), .GUID(0))
            .Value = VarPtr(JPEGQuality)
        End With
        Set IIStream = IStreamFromArray(0&, 0&)
        If Not IIStream Is Nothing Then
            bOK = (GdipSaveImageToStream(tgtHandle, IIStream, uEncCLSID(0), uEncParams) = 0&)
        End If
    End If
    If bOK Then
        If returnMedium = saveTo_GDIplus Then
            If tGDIpImage Is Nothing Then
                If GdipLoadImageFromStream(ObjPtr(IIStream), pvSaveAsJPEG) = 0& Then Set returnObject = IIStream
            Else
                Set tGDIpImage = LoadImage(tData(), , , True)
                If tGDIpImage.Handle Then
                    Set returnObject = tGDIpImage
                    pvSaveAsJPEG = returnMedium
                End If
            End If
        Else
            If tGDIpImage Is Nothing Then bOK = IStreamToArray(ObjPtr(IIStream), tData())
            If bOK Then
                Set tGDIpImage = Nothing
                If returnMedium = saveTo_Array Then
                    pvSaveAsJPEG = UBound(tData) + 1&
                    MoveArrayToVariant returnObject, tData(), True
                ElseIf returnMedium = saveTo_File Then
                    WriteFile CLng(returnObject), tData(0), UBound(tData) + 1&, lValue, ByVal 0&
                    pvSaveAsJPEG = (lValue > UBound(tData))
                Else
                    Set tmpPic = ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
                    If Not tmpPic Is Nothing Then
                        If returnMedium = saveTo_stdPicture Then        ' stdPicture supports JPG format
                            Set returnObject = tmpPic
                        ElseIf returnMedium = saveTo_Clipboard Then
                            Clipboard.SetData tmpPic
                            If g_ClipboardFormat Then SetClipboardCustomFormat tData(), g_ClipboardFormat
                        ElseIf returnMedium = saveTo_DataObject Then
                            Set tmpDO = returnObject
                            tmpDO.SetData tmpPic, vbCFBitmap
                            If g_ClipboardFormat Then tmpDO.SetData tData(), g_ClipboardFormat
                            Set tmpDO = Nothing
                        End If
                        pvSaveAsJPEG = tmpPic.Handle
                        Set tmpPic = Nothing
                    End If
                    Erase tData()
                End If
            End If
        End If
    End If
    
    If tgtHandle <> SourceHandle Then
        If tgtHandle Then GdipDisposeImage tgtHandle
    End If

End Function

Public Function SaveAsTIFF(returnObject As Variant, SourceHandle As Long, _
                           ByVal returnMedium As SaveAsMedium, _
                           Optional ByVal Action As TIFFMultiPageActions = TIFF_SingleFrame, _
                           Optional SaveOptions As Variant) As Long

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
    
    ' Some notes from results of testing
    ' TIFF will save any 1 bit image as black and white; even if the image is not black & white
    ' There are 2 workarounds
    '   1) Save TIFF at 24 or 32 bpp to force 1bpp non-b&w image with color
    '   2) Rewrite the 1bpp b&w image as 4 bpp.
    ' I use option #2. 4bpp image is far smaller than 24bpp image
    
    
    ' Single-Page TIFF only
    '   SourceHandle must be a valid GDI+ handle
    '   Action must be TIFF_SingleFrame
    
    ' Multi-Page TIFF only
    '   Action must be in this order
    '       TIFF_MultiFrameStart to begin TIFF & pass SourceHandle for 1st page of TIFF
    '       TIFF_MultiFrameAdd to add additional pages & pass new SourceHandle for each new page
    '       TIFF_MultiFrameEnd to stop adding and return the TIFF GDI+ handle & fill the returnObject
    '           SourceHandle is not used during TIFF_MultiFrameEnd, can be zero
    '           This must be called to free up the stream and handle this function caches
    '  During each phase of TIFF creation, validate the function return value. If the function
    '   returns zero, then abort adding any new images to the failed TIFF
    
    Static hMultiTIFF As Long
    Static IStreamMultiPage As IUnknown

    Dim uEncCLSID(0 To 3) As Long, outData() As Byte
    Dim uEncParams As EncoderParameters, bOK As Boolean
    Dim lParamVal As Long, IIStream As IUnknown, tImage As GDIpImage
    Dim lDepth As Long, lCompress As Long, tObject As Object
    Dim SS As SAVESTRUCT, cDepthReduction As cColorReduction
    Dim tmpPic As StdPicture, tmpDO As DataObject, cPal As ColorPalette
    Dim tgtHandle As Long, srcDepth As Long, sizeF As RECTF

    Const MimeType As String = "image/tiff"
    Const TIFFCompress As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
    Const TIFFBitDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
    Const SaveFlag As String = "{292266fc-ac40-47bf-8cfc-a85b89a655de}"
    Const EncoderValueCompressionLZW As Long = &H2
    Const EncoderValueMultiFrame As Long = 18&
    Const EncoderValuePageDims As Long = 23&
    Const EncoderValueFlush As Long = 20&
    Const EncoderParameterValueTypeLong As Long = &H4&
    
    If SourceHandle <> 0& Or Action = TIFF_MultiFrameEnd Then
        If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then   ' get TIFF encoder class
        
            If SourceHandle Then
                If IsMissing(SaveOptions) Then Exit Function
                SS = SaveOptions
                If (SS.reserved1 And 1&) = 0& Then      ' 1bpp workaround; if b&w will remain b&w
                    If ColorDepthToColorType(lvicNoColorReduction, SourceHandle) = lvicColor1bpp Then
                        GdipGetImagePaletteSize SourceHandle, lParamVal
                        If lParamVal = 16& Then         ' 2 colors + 2 flags
                            GdipGetImagePalette SourceHandle, cPal, lParamVal
                            If Not (cPal.Entries(1) = &HFF000000 And cPal.Entries(2) = &HFFFFFFFF) Then
                                SS.ColorDepth = lvicConvert_16Colors    ' going to force 4bpp
                            Else
                                SS.ColorDepth = lvicNoColorReduction    ' no point; already 1bpp
                            End If
                        End If
                    End If
                End If
                If SS.ColorDepth > lvicNoColorReduction Then
                    If SS.ColorDepth < lvicConvert_TrueColor24bpp Then
                        Set cDepthReduction = New cColorReduction
                        If SS.ColorDepth < lvicConvert_256Colors Then lParamVal = orient_4bppIndexesMin Else lParamVal = 0&
                        If SS.ColorDepth = lvicDefaultReduction Then
                            SS.reserved1 = SS.reserved1 Or &H10000000
                            tgtHandle = cDepthReduction.PalettizeToHandle(SourceHandle, alpha_Complex, SS, orient_GDIpHandle Or lParamVal)
                        ElseIf SS.ColorDepth < lvicConvert_TrueColor24bpp Then
                            tgtHandle = cDepthReduction.PalettizeToHandle(SourceHandle, alpha_Complex, SS, orient_GDIpHandle Or lParamVal)
                        End If
                        Set cDepthReduction = Nothing
                    End If
                End If
                If Action = TIFF_MultiFrameStart Then
                    If hMultiTIFF Then
                        GdipDisposeImage hMultiTIFF
                        Set IStreamMultiPage = Nothing
                    End If
                    If tgtHandle Then                   ' use processed image
                        hMultiTIFF = tgtHandle
                    ElseIf SS.reserved2 = 0& Then
                        hMultiTIFF = SourceHandle        ' use passed image
                        SourceHandle = 0&
                    Else                                ' copy image
                        hMultiTIFF = CreateSourcelessHandle(SourceHandle)
                    End If
                    If hMultiTIFF = 0& Then Exit Function
                    tgtHandle = hMultiTIFF
                ElseIf tgtHandle = 0& Then
                    tgtHandle = SourceHandle
                End If
            End If
            On Error GoTo ExitRoutine                           ' begin error trapping for cleanup
            
            If (tgtHandle = SourceHandle) And (SS.reserved2 <> 0&) And SS.ExtractCurrentFrameOnly = False And _
                ((SS.reserved1 And &HFF00&) \ &H100& = lvicPicTypeTIFF) Then
                ' should the source already be TIFF and not converted in above section; use its data
                CopyMemory tObject, SS.reserved2, 4&
                Set tImage = tObject
                CopyMemory tObject, 0&, 4&
                bOK = tImage.ExtractImageData(outData)
            Else
                If SS.CompressionJPGQuality <> lvicFormatUncompressed Then ' use compression to keep bytes size low
                    With uEncParams.Parameter(0)
                        lCompress = EncoderValueCompressionLZW
                        .NumberOfValues = 1
                        .Type = EncoderParameterValueTypeLong
                         CLSIDFromString StrPtr(TIFFCompress), .GUID(0)
                        .Value = VarPtr(lCompress)
                    End With
                    uEncParams.Count = 1                            ' keep count
                End If
                
                With uEncParams.Parameter(uEncParams.Count)         ' set up bit depth parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong
                    CLSIDFromString StrPtr(TIFFBitDepth), .GUID(0)
                   .Value = VarPtr(lDepth)                          ' will be applied per page
                End With
                uEncParams.Count = uEncParams.Count + 1             ' keep count
                
                If Action <> TIFF_SingleFrame Then                  ' multiple images
                    With uEncParams.Parameter(uEncParams.Count)
                        .NumberOfValues = 1
                        .Type = EncoderParameterValueTypeLong
                        CLSIDFromString StrPtr(SaveFlag), .GUID(0)
                       .Value = VarPtr(lParamVal)
                    End With
                    uEncParams.Count = uEncParams.Count + 1         ' keep count
                End If
                                    
                If Action <> TIFF_MultiFrameEnd Then                ' get color depth of image & update parameter
                    If (SS.reserved1 And 1&) Then
                        lDepth = 32&
                    Else
                        lDepth = (ColorDepthToColorType(0&, tgtHandle) And &HFF00&) \ &H100&
                    End If
                End If
                
                Select Case Action
                    
                    Case TIFF_MultiFrameStart, TIFF_SingleFrame     ' starting single/multipage TIFF
                        Set IIStream = IStreamFromArray(0&, 0&)     ' create IStream for TIFF
                        If IIStream Is Nothing Then GoTo ExitRoutine ' start the TIFF stream
                        lParamVal = EncoderValueMultiFrame          ' ignored for single page TIFF
                        bOK = (GdipSaveImageToStream(tgtHandle, IIStream, uEncCLSID(0&), uEncParams) = 0&)
                        If Not bOK Then
                            If lDepth < 32& Then                    ' failure, try to save as 32bit
                                lDepth = 32&
                                bOK = (GdipSaveImageToStream(tgtHandle, IIStream, uEncCLSID(0&), uEncParams) = 0&)
                            End If
                        End If
                    
                    Case TIFF_MultiFrameAdd                         ' appending to started TIFF
                        If hMultiTIFF = 0& Then Exit Function   ' told you not to diddle with it
                        lParamVal = EncoderValuePageDims
                        bOK = (GdipSaveAddImage(hMultiTIFF, tgtHandle, uEncParams) = 0&)
                        If Not bOK Then
                            If lDepth < 32& Then                    ' failure try again with 32bpp
                                lDepth = 32&
                                bOK = (GdipSaveAddImage(hMultiTIFF, tgtHandle, uEncParams) = 0&)
                            End If
                        End If
                    
                    Case TIFF_MultiFrameEnd                         ' closing off TIFF stream
                        If hMultiTIFF = 0& Then Exit Function       ' told you not to diddle with it
                        lParamVal = EncoderValueFlush
                        bOK = (GdipSaveAdd(hMultiTIFF, uEncParams) = 0&)
                         
                End Select
            End If
            
            If bOK Then                                         ' set return parameters
                If Action = TIFF_MultiFrameStart Then           ' starting a new TIFF stream
                    Set IStreamMultiPage = IIStream                 ' return Stream so it can be passed again & again
                    SaveAsTIFF = hMultiTIFF                     ' return non-zero for success
                
                ElseIf Action = TIFF_SingleFrame Or Action = TIFF_MultiFrameEnd Then
                    ' when ending multi-page stream, get the stream from the returnObject
                    ' else the IStream object created above is the the stream we want
                    bOK = False ' default failure
                    If hMultiTIFF Then GdipDisposeImage hMultiTIFF: hMultiTIFF = 0&
                    
                    If Action = TIFF_MultiFrameEnd Then Set IIStream = IStreamMultiPage
                    If returnMedium = saveTo_Array Or returnMedium = saveTo_File Then
                        If tImage Is Nothing Then               ' else already have data
                            bOK = IStreamToArray(ObjPtr(IIStream), outData())
                        Else
                            bOK = True
                        End If
                        If bOK Then
                            If returnMedium = saveTo_Array Then
                                SaveAsTIFF = UBound(outData) + 1&   ' return size of array
                                MoveArrayToVariant returnObject, outData(), True
                            Else
                                WriteFile CLng(returnObject), outData(0), UBound(outData) + 1&, lDepth, ByVal 0&
                                SaveAsTIFF = (lDepth > UBound(outData))
                            End If
                        End If
                    Else
                        If tImage Is Nothing Then                   ' else already have handle
                            GdipLoadImageFromStream ObjPtr(IIStream), SaveAsTIFF
                        Else
                            SaveAsTIFF = tImage.Handle
                        End If
                        If SaveAsTIFF <> 0& Then
                            If returnMedium = saveTo_GDIplus Then
                                If Not tImage Is Nothing Then
                                    Set IIStream = IStreamFromArray(VarPtr(outData(0)), UBound(outData) + 1&)
                                    If Not IIStream Is Nothing Then bOK = (GdipLoadImageFromStream(ObjPtr(IIStream), SaveAsTIFF) = 0&)
                                Else
                                    bOK = True
                                End If
                                If bOK Then returnObject = IIStream
                            Else
                                Call GdipCreateHBITMAPFromBitmap(SaveAsTIFF, lParamVal, SS.RSS.FillColorARGB Or &HFF000000)
                                If tImage Is Nothing Then GdipDisposeImage SaveAsTIFF: SaveAsTIFF = 0&
                                If lParamVal Then
                                    Set tmpPic = HandleToStdPicture(lParamVal, vbPicTypeBitmap)
                                    If tmpPic Is Nothing Then
                                        DeleteObject hMultiTIFF
                                    ElseIf tmpPic.Handle = 0& Then
                                        DeleteObject hMultiTIFF
                                    Else
                                        bOK = True
                                        If tImage Is Nothing Then bOK = IStreamToArray(ObjPtr(IIStream), outData())
                                        If bOK Then
                                            If returnMedium = saveTo_stdPicture Then
                                                Set returnObject = tmpPic
                                            ElseIf returnMedium = saveTo_Clipboard Then
                                                Clipboard.SetData tmpPic
                                                If g_ClipboardFormat Then SetClipboardCustomFormat outData(), g_ClipboardFormat
                                            ElseIf returnMedium = saveTo_DataObject Then
                                                Set tmpDO = returnObject
                                                tmpDO.SetData tmpPic, vbCFBitmap
                                                If g_ClipboardFormat Then tmpDO.SetData outData(), g_ClipboardFormat
                                            End If
                                        End If
                                    End If
                                    SaveAsTIFF = bOK
                                    hMultiTIFF = 0&
                                End If
                            End If
                        End If
                    End If
                    Set IStreamMultiPage = Nothing
                Else                                            ' appending pages to a TIFF stream
                    SaveAsTIFF = tgtHandle                      ' simply return non-zero for success
                End If
            End If
        End If
    End If
    
ExitRoutine:
    If tgtHandle <> SourceHandle Then
        If tgtHandle <> hMultiTIFF Then
           If tgtHandle Then GdipDisposeImage tgtHandle
        End If
    End If
    If bOK = False Then
        If hMultiTIFF Then GdipDisposeImage hMultiTIFF: hMultiTIFF = 0&
        Set IStreamMultiPage = Nothing
    End If
    
End Function

Public Sub CreateGDIpAttributeHandle(ByRef hAttributes As Long, ByVal GrayScaleType As Long, _
                                        ByVal Lightness As Long, ByVal GlobalTransparency As Long, _
                                        ByVal TransparentARGBColor As Long, _
                                        ByVal BlendColor As Long, ByVal Inverted As Boolean)
                                        
    ' Function creates a GDI+ image attributes handle using following custom martrices
    
    ' Notes about color matrix.
    ' To convert .Net examples from the web
    ' Part 1:
    '   When you see examples like:
    '       (1, 0, 0, 0, 0,
    '        0, 1, 0, 0, 0,
    '        0, 0, 1, 0, 0,
    '        0, 0, 0, 1, 0,
    '        0, 0, 0, 0, 1)
    ' It is translated like so to the clrMatrix array used herein...
    '   consider clrMatrix dimensioned as (columns, rows), so...
    '   1st column above = clrMatrix(0-4,0)
    '   2nd column above = clrMatrix(0-4,1)
    '   3rd column above = clrMatrix(0-4,2)
    '   4th column above = clrMatrix(0-4,3)
    '   5th column above = clrMatrix(0-4,4)
    
    ' Part 2:
    '    When you see examples like:
    '       cm.Matrix00=cm.Matrix11=cm.Matrix22=1
    '       etc...
    '   Consider clrMatrix dimensioned as (columns, rows), so...
    '   Separate the 2 digits in each Matrix## word as row,column
    '       Matrix00 = clrMatrix(0,0)
    '       Matrix11 = clrMatrix(1,1)
    '       Matrix40 = clrMatrix(0,4) and so on
    
    Dim Col As Long, Row As Long
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    ' each column in the 5x5 matrix: 0=DstRed, 1=DstGreen, 2=DstBlue, 3=DstAlpha, 4=not used
    ' DstRed    DstGreen    DstBlue     DstAlpha    n/a
    '-----------------------------------------------------------
    ' RedWeight RedWeight   RedWeight      0        0
    ' GrnWeight GrnWeight   GrnWeight      0        0
    ' BluWeight BluWeight   BluWeight      0        0
    ' TrnWeight TrnWeight   TrnWeight      1*       0   (*global transparency value)
    ' Offset    Offset      Offset         0        1*  (*required)
    '-----------------------------------------------------------
    ' TrnWeight (alpha/transparency) is not used in this function
    ' the DstRed, DstGreen & DstBlue values are calculated by summing the
    '       individual R,G,B component values, per column, multiplied by their weights
    ' Example:
    ' DstRGBcomponent = RedComp*RedWeight + GreenComp*GrnWeight + BlueComp*BluWeight + Offset
    
    '   Identity matrix (no change in R,G, or B) looks like...
    '   clrMatrix(0,0)=1: clrMatrix(1,1)=1: clrMatrix(2,2)=1
    
    Const ColorAdjustTypeBitmap As Long = &H1&
    Const BWLuminanceRatio As Single = 0.475
    ' ^^ used for Black & White grayscale option. Min/Max values: 0.0 / 1.0
    '   Make value larger for more black and smaller for more white
    '   Can adjust this via code by adding/subtracting lightness too and/or inverting colors
    
    If hAttributes Then
        GdipDisposeImageAttributes hAttributes
        hAttributes = 0&
    End If
    
    ' validate passed parameters within bounds
    If Lightness < -100& Then
        Lightness = -100&
    ElseIf Lightness > 100& Then
        Lightness = 100&
    End If
    If GlobalTransparency < 0& Then
        GlobalTransparency = 0&
    ElseIf GlobalTransparency > 100& Then
        GlobalTransparency = 100&
    End If
    
    If GrayScaleType Then
        ' grayscale the image; populate 1st column, 1st 3 rows
        Select Case GrayScaleType
            Case lvicNTSCPAL ' standard weighted average
                clrMatrix(Col, Row) = 0.299: clrMatrix(Col, Row + 1) = 0.587: clrMatrix(Col, Row + 2) = 0.114
            Case lvicCCIR709, lvicBlackWhite ' CCIR709
                clrMatrix(Col, Row) = 0.213: clrMatrix(Col, Row + 1) = 0.715: clrMatrix(Col, Row + 2) = 0.072
            Case lvicSimpleAverage ' pure average
                clrMatrix(Col, Row) = 0.333: clrMatrix(Col, Row + 1) = 0.334: clrMatrix(Col, Row + 2) = clrMatrix(Col, Row)
            Case lvicRedMask ' personal preferences: could be r=1:g=0:b=0 or other weights
                clrMatrix(Col, Row) = 0.8: clrMatrix(Col, Row + 1) = 0.1: clrMatrix(Col, Row + 2) = clrMatrix(Col, Row + 1)
            Case lvicGreenMask ' personal preferences: could be r=0:g=1:b=0 or other weights
                clrMatrix(Col, Row) = 0.1: clrMatrix(Col, Row + 1) = 0.8: clrMatrix(Col, Row + 2) = clrMatrix(Col, Row)
            Case lvicBlueMask ' personal preferences: could be r=0:g=0:b=1 or other weights
                clrMatrix(Col, Row) = 0.1: clrMatrix(Col, Row + 1) = clrMatrix(Col, Row): clrMatrix(Col, Row + 2) = 0.8
            Case lvicRedGreenMask ' personal preferences: could be r=.5:g=.5:b=0 or other weights
                clrMatrix(Col, Row) = 0.45: clrMatrix(Col, Row + 1) = clrMatrix(Col, Row): clrMatrix(Col, Row + 2) = 0.1
            Case lvicBlueGreenMask ' personal preferences: could be r=0:g=.5:b=.5 or other weights
                clrMatrix(Col, Row) = 0.1: clrMatrix(Col, Row + 1) = 0.45: clrMatrix(Col, Row + 2) = clrMatrix(Col, Row + 1)
            Case lvicSepia ' populate 1st 3 columns, 1st 3 rows
                clrMatrix(0, 0) = 0.393: clrMatrix(1, 0) = 0.349: clrMatrix(2, 0) = 0.272
                clrMatrix(0, 1) = 0.769: clrMatrix(1, 1) = 0.686: clrMatrix(2, 1) = 0.534
                clrMatrix(0, 2) = 0.189: clrMatrix(1, 2) = 0.168: clrMatrix(2, 2) = 0.131
            Case Else ' no grayscale
                GrayScaleType = lvicNoGrayScale
        End Select
        If GrayScaleType Then
            clrMatrix(Col, Row + 4) = 0.001         ' add minor offset to prevent GDI+ overflow
            If GrayScaleType = lvicSepia Then
                clrMatrix(Col + 1, Row + 4) = 0.001: clrMatrix(Col + 2, Row + 4) = 0.001
            Else                                    ' fill in columns 2 & 3 of 1st 4 rows
                For Row = 0& To 4&
                    For Col = 1& To 2&
                        clrMatrix(Col, Row) = clrMatrix(0, Row)
                Next: Next
            End If
            clrMatrix(4, 4) = 1!                    ' flag indicating need to create attributes
        End If
    End If
    
    If Lightness Then
        ' add/subtract light intensity by updating the 4th row of 1st 3 columns
        clrMatrix(0, 4) = Lightness / 100! ' red added/subtracted brightness
        clrMatrix(1, 4) = clrMatrix(0, 4) ' same for blue
        clrMatrix(2, 4) = clrMatrix(0, 4) ' same for green
        If clrMatrix(4, 4) = 0! Then
            clrMatrix(0, 0) = 1!: clrMatrix(1, 1) = 1!: clrMatrix(2, 2) = 1!
            clrMatrix(4, 4) = 1!                    ' flag indicating need to create attributes
        End If
    End If
    
    If BlendColor Then
        ' Blending requires a bit more math, but very doable
        ' 1. Get the blend percentage from the hiword/alpha value
        clrMatrix(4, 4) = ((BlendColor And &H7F000000) \ &H1000000) / 100!
        If clrMatrix(4, 4) > 0! Then
            ' 2. Separate the blend color from the blend percentage
            BlendColor = (BlendColor And Not &H7F000000)
            If BlendColor < 0& Then BlendColor = GetSysColor(BlendColor And &HFF&)
            ' 3. Ensure percentage does not exceed max value
            If clrMatrix(4, 4) > 1! Then clrMatrix(4, 4) = 1!
            ' 4. Use custom, weighted, blend algorithm
                ' cannot have any zero values for this algo
                If clrMatrix(0, 0) = 0! Then clrMatrix(0, 0) = 1!
                If clrMatrix(1, 1) = 0! Then clrMatrix(1, 1) = 1!
                If clrMatrix(2, 2) = 0! Then clrMatrix(2, 2) = 1!
                ' calculate percentage of R,G,B within the blend color
                clrMatrix(4, 0) = ((BlendColor And &HFF&) - 127&) / 255! * clrMatrix(4, 4)
                clrMatrix(4, 1) = (((BlendColor \ &H100&) And &HFF) - 127&) / 255! * clrMatrix(4, 4)
                clrMatrix(4, 2) = (((BlendColor \ &H10000) And &HFF) - 127&) / 255! * clrMatrix(4, 4)
                For Row = 0& To 2&
                    For Col = 0& To 2&
                        clrMatrix(Col, Row) = clrMatrix(4, Row) * clrMatrix(Col, Row) + clrMatrix(Col, Row)
                    Next
                    If clrMatrix(0, Row) = 0! Then clrMatrix(0, Row) = 0.001!
                    clrMatrix(4, Row) = 0!
                Next
            clrMatrix(4, 4) = 1!                    ' flag indicating need to create attributes
        End If
    End If
    
    If (Inverted Or GlobalTransparency) Then
        If clrMatrix(4, 4) = 0! Then
            clrMatrix(0, 0) = 1!: clrMatrix(1, 1) = 1!: clrMatrix(2, 2) = 1!
            clrMatrix(4, 4) = 1!                    ' flag indicating need to create attributes
        End If
        If Inverted Then
            For Col = 0& To 2&
                For Row = 0& To 2&
                    clrMatrix(Col, Row) = -clrMatrix(Col, Row)
                 Next
                clrMatrix(Col, 4) = -clrMatrix(Col, 4) + 1!
            Next
        End If
    End If
    
    If clrMatrix(4, 4) = 1! Then                    ' create attributes?
        ' create the image attribute class and set the color matrix
        If GdipCreateImageAttributes(hAttributes) = 0& Then
            ' global blending; value between 0 & 1
            clrMatrix(3, 3) = CSng((100! - GlobalTransparency) / 100!)
            If Not GdipSetImageAttributesColorMatrix(hAttributes, ColorAdjustTypeBitmap, 1&, clrMatrix(0, 0), clrMatrix(0, 0), 0&) = 0& Then
                GdipDisposeImageAttributes hAttributes
                hAttributes = 0&
            ElseIf GrayScaleType = lvicBlackWhite Then
                GdipSetImageAttributesThreshold hAttributes, ColorAdjustTypeBitmap, 1&, Abs(BWLuminanceRatio + (Inverted * 1!))
            End If
        End If
    End If
                                
    If (TransparentARGBColor And &HFF000000) Then 'else no alpha value set, color is already transparent if it exists
        TransparentARGBColor = (TransparentARGBColor And &HFF00FF00) Or (TransparentARGBColor And &HFF0000) \ &H10000 Or (TransparentARGBColor And &HFF) * &H10000
        If hAttributes = 0& Then
            If GdipCreateImageAttributes(hAttributes) = 0& Then
                GdipSetImageAttributesColorKeys hAttributes, ColorAdjustTypeBitmap, 1&, TransparentARGBColor, TransparentARGBColor
            End If
        Else
            GdipSetImageAttributesColorKeys hAttributes, ColorAdjustTypeBitmap, 1&, TransparentARGBColor, TransparentARGBColor
        End If
    End If
    
End Sub

Public Function SetClipboardCustomFormat(inData() As Byte, formatType As Integer) As Long

    ' function adds our image data to the clipboard using a custom-generated format

    Dim lSize As Long, lHandle As Long, lLock As Long
    
    If formatType Then
        If OpenClipboard(0&) Then
            On Error Resume Next
            lSize = UBound(inData) + 1&
            On Error GoTo 0
            If lSize Then
                lHandle = GlobalAlloc(&H2000, lSize)
                If lHandle Then
                    lLock = GlobalLock(lHandle)
                    CopyMemory ByVal lLock, inData(0), lSize
                    GlobalUnlock lHandle
                    SetClipboardCustomFormat = SetClipboardData((&HFFFF& And formatType), lHandle)
                End If
            End If
            Call CloseClipboard
        End If
    End If
ExitRoutine:
End Function

Public Function GetScaledImageSizes(ByVal ImageWidth As Long, ByVal ImageHeight As Long, _
                                    ByVal destWidth As Long, ByVal destHeight As Long, _
                                    ByRef ScaledWidth As Long, ByRef ScaledHeight As Long, _
                                    Optional ByVal Angle As Single = 0!, _
                                    Optional ByVal CanScaleUp As Boolean = True, _
                                    Optional ByVal CanClip As Boolean = True) As Boolean

    ' Function returns scaled (maintaining scale ratio) for passed destination width/height
    ' The CanScaleUp when set to false will never return scaled sizes > than 1:1
    ' The ClipRotation parameter, if False, will ensure rotated image does not exceed destination width/height
    ' The scaled width & height returned in the ScaledWidth & ScaledHeight parameters
    ' If function returns false, return parameters are undefined
    
    Dim xRatio As Double, yRatio As Double
    Dim sinT As Double, cosT As Double, d2r As Double
    Dim h1 As Long, h2 As Long, a As Double

    If (ImageWidth < 1& Or ImageHeight < 1&) Then Exit Function
    If (destWidth < 1& Or destHeight < 1&) Then Exit Function

    xRatio = destWidth / ImageWidth
    yRatio = destHeight / ImageHeight
    If xRatio > yRatio Then xRatio = yRatio
    
    If Angle < 0! Then
        a = 360# + (Angle Mod 360)
    Else
        a = (Angle Mod 360)
    End If
    
    If (a = 0# And Int(Angle) = Angle) Or CanClip = True Then
    
        If xRatio >= 1! And CanScaleUp = False Then
            ScaledWidth = ImageWidth
            ScaledHeight = ImageHeight
        Else
            yRatio = ImageHeight * xRatio
            xRatio = ImageWidth * xRatio
            ScaledWidth = Int(xRatio)
            ScaledHeight = Int(yRatio)
            If xRatio > ScaledWidth Then ScaledWidth = ScaledWidth + 1&
            If yRatio > ScaledHeight Then ScaledHeight = ScaledHeight + 1&
        End If
    
    Else

        yRatio = ImageHeight * xRatio
        ScaledWidth = Int(ImageWidth * xRatio)
        ScaledHeight = Int(yRatio)
        If xRatio * ImageWidth > ScaledWidth Then ScaledWidth = ScaledWidth + 1&
        If yRatio > ScaledHeight Then ScaledHeight = ScaledHeight + 1&
        
        Select Case a
            Case Is < 91
            Case Is < 181: a = 180 - a
            Case Is < 271: a = a - 180
            Case Else: a = 360 - a
        End Select
        d2r = (4& * Atn(1)) / 180   ' conversion factor for degree>radian
        sinT = Sin(a * d2r)
        cosT = Cos(a * d2r)

        h1 = destHeight * destHeight / (ScaledWidth * sinT + ScaledHeight * cosT)
        h2 = destWidth * destHeight / (ScaledWidth * cosT + ScaledHeight * sinT)
        If h1 < h2 Then h2 = h1
        h1 = h2 * destWidth / destHeight
        
        If xRatio >= 1& And CanScaleUp = False Then

            If Not (h1 < ImageWidth Or h2 < ImageHeight) Then
                ScaledWidth = ImageWidth
                ScaledHeight = ImageHeight
                GetScaledImageSizes = True
                Exit Function
            End If

        End If
    
        xRatio = h1 / ImageWidth
        yRatio = h2 / ImageHeight
        If xRatio > yRatio Then xRatio = yRatio

        yRatio = ImageHeight * xRatio
        xRatio = ImageWidth * xRatio
        ScaledWidth = Int(xRatio)
        ScaledHeight = Int(yRatio)
        If xRatio > ScaledWidth Then ScaledWidth = ScaledWidth + 1&
        If yRatio > ScaledHeight Then ScaledHeight = ScaledHeight + 1&

'        ScaledWidth = Int(ImageWidth * xRatio)
'        ScaledHeight = Int(yRatio)
        
    End If
    
    GetScaledImageSizes = True

End Function

Public Function GetScaledCanvasSize(ByVal imgWidth As Long, ByVal imgHeight As Long, _
                                    ByRef CanvasWidth As Long, ByRef CanvasHeight As Long, _
                                    Optional ByVal Angle As Single = 0!) As Boolean

    ' function returns the size of a container/DC required to render the passed dimensions at passed Angle
    
    If (imgWidth < 1& Or imgHeight < 1&) Then Exit Function
    
    Dim sinT As Double, cosT As Double
    Dim a As Double, d2r As Double
    Dim ctrX As Double, ctrY As Double

    If Angle < 0! Then
        a = 360# + (Angle Mod 360)
    Else
        a = (Angle Mod 360)
    End If
    Select Case a
        Case Is < 91#
        Case Is < 181#: a = 180# - a
        Case Is < 271#: a = a - 180#
        Case Else: a = 360# - a
    End Select
    d2r = (4# * Atn(1)) / 180#   ' conversion factor for degree>radian
    sinT = Sin(a * d2r)
    cosT = Cos(a * d2r)
    
    ctrX = imgWidth / 2#
    ctrY = imgHeight / 2#

    a = (-ctrX * sinT) + (-ctrY * cosT)
    d2r = (imgWidth - ctrX) * sinT + (imgHeight - ctrY) * cosT - a
    If d2r - Int(d2r) > 0.00001 Then d2r = d2r + 1#
    CanvasHeight = Int(d2r)
    
    a = ((-ctrX * cosT) - (imgHeight - ctrY) * sinT)
    d2r = (imgWidth - ctrX) * cosT - (-ctrY * sinT) - a
    If d2r - Int(d2r) > 0.00001 Then d2r = d2r + 1#
    CanvasWidth = Int(d2r)

End Function

Private Function pvGetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
  
 ' Routine is a helper function for the various SaveAsxxxx routines
 
  Dim lCount   As Long
  Dim lSize    As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim buffer() As Byte, sMime As String
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(lCount, lSize)
    If (lSize = 0& Or lCount = 0&) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To lCount) As ImageCodecInfo
    ReDim buffer(1 To lSize) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(lCount, lSize, buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), buffer(1), (Len(ICI(1)) * lCount))
    
    lSize = Len(strMimeType)
    sMime = String$(lSize, vbNullChar)
    '-- Loop through all the codecs
    For lIdx = lCount To 1& Step -1&
        '-- Must convert the pointer into a usable string
        With ICI(lIdx)
            If lSize = lstrlenW(ByVal .MimeType) Then
                Call CopyMemory(ByVal StrPtr(sMime), ByVal .MimeType, lSize * 2&)
                If StrComp(sMime, strMimeType, vbTextCompare) = 0 Then
                    CopyMemory ClassID(0), .ClassID(0), 16&
                    Exit For
                End If
            End If
        End With
    Next lIdx
    pvGetEncoderClsID = lIdx
End Function

Public Function LoadImage(ImageSource As Variant, _
                        Optional ByVal KeepOriginalFormat As Boolean = True, _
                        Optional ByVal NoFileLock As Boolean = False, _
                        Optional ByVal ReturnEmptyClassOnFailure As Boolean, _
                        Optional ByVal AsyncClient As GDIpImage = Nothing) As GDIpImage

    ' See AICGlobals.LoadPictureGDIplus for parameter descriptions & restrictions
    ' This is the procedure that loads a resource for the GDIpImage classes
    
    If g_TokenClass.Token = 0& Then
        If ReturnEmptyClassOnFailure Then Set LoadImage = New GDIpImage
        Exit Function
    End If
    
    Dim lResult As Long, tData() As Byte, Size As RECTF
    Dim tSource As Variant, tHandle As Long, tPic As StdPicture
    Dim cImageData As cGDIpMultiImage
    Dim tAI As ASSOCIATEDICON
    
    Set cImageData = New cGDIpMultiImage
    
    If IsEmpty(ImageSource) Then
        lResult = lvicPicTypeUnknown
        
    ElseIf IsObject(ImageSource) Then
        If ImageSource Is Nothing Then
            lResult = lvicPicTypeUnknown
        Else
            lResult = pvProcessObjectSource(cImageData, ImageSource, tData(), KeepOriginalFormat, NoFileLock)
        End If
        
    ElseIf VarType(ImageSource) = vbUserDefinedType Then    ' UDT passed
        If LenB(ImageSource) = LenB(tAI) Then
            On Error Resume Next
            tAI = ImageSource
            If Err Then
                Err.Clear
            Else
                lResult = pvProcessAssociatedIcon(cImageData, KeepOriginalFormat, tAI)
                If lResult Then ImageSource.IndexReturned = tAI.IndexReturned
            End If
            On Error GoTo 0
        End If
        
    ElseIf VarType(ImageSource) = vbString Then        ' assume file name & revert to URL if failure
        If ImageSource = vbNullString Then             ' processing routine tweaked to test for Base64
            lResult = lvicPicTypeUnknown
        Else
            lResult = pvProcessFileSource(cImageData, CStr(ImageSource), KeepOriginalFormat, NoFileLock)
            If lResult = lvicPicTypeNone Then lResult = pvProcessURLSource(cImageData, CStr(ImageSource), KeepOriginalFormat, NoFileLock)
        End If
        
    ElseIf (VarType(ImageSource) And vbArray) Then
        ' following function ensures array is either byte, long or integer
        If NormalizeArray(ImageSource, tData) Then
            lResult = pvProcessArraySource(cImageData, tData(), KeepOriginalFormat, lvicPicTypeUnknown)
        End If
            
    ElseIf TypeOf ImageSource Is IUnknown Then              ' IStream
        If GdipLoadImageFromStream(ObjPtr(ImageSource), tHandle) = 0& Then
            lResult = GetImageType(tHandle)
            cImageData.CacheSourceInfo ImageSource, tHandle, lResult, KeepOriginalFormat, False
        End If
        
    Else ' test for numeric values & assume are handles
        On Error Resume Next
        lResult = CLng(ImageSource)
        If Err Then
            Err.Clear
            On Error GoTo 0
        Else
            On Error GoTo 0
            If lResult = ImageSource Then                   ' testing for non-whole numbers
                If lResult = lvicPicTypeNone Then
                    lResult = lvicPicTypeUnknown
                Else
                    lResult = pvProcessGDIhandleSource(cImageData, lResult, KeepOriginalFormat) ' process as either bitmap/icon handle
                End If
            End If
        End If
    End If
    
    If lResult Then
        Select Case lResult
            Case lvicPicTypeMetafile, lvicPicTypeEMetafile
                ' in order to allow WMF/EMF to be rendered with effects, we convert them to bitmap
                ' The original format is maintained depending on KeepOriginalFormat parameter
                If GetImageType(cImageData.Handle) <> lvicPicTypeBitmap Then
                    ' else wmf non-placeable converted already; nothing to do
                    tHandle = CreateSourcelessHandle(cImageData.Handle)
                    If tHandle Then
                        GdipDisposeImage cImageData.Handle
                        If KeepOriginalFormat = False Then
                            cImageData.CacheSourceInfo Empty, tHandle, lResult, False, False
                        Else
                            cImageData.CacheSourceInfo tSource, 0&, 0&, False, True
                            Set cImageData = New cGDIpMultiImage
                            cImageData.CacheSourceInfo tSource, tHandle, lResult, True, False
                        End If
                    Else
                        lResult = 0&
                        Set cImageData = Nothing
                    End If
                End If
            Case lvicPicTypeJPEG
                Call GdipGetImageBounds(cImageData.Handle, Size, UnitPixel)
                If (Size.nHeight = 0! Or Size.nWidth = 0!) Then
                    ' some known cases where GDI+ loads a valid JPG but reports 0x0 size.
                    ' Patch takes source and allows VB to load JPG, then converts its handle to bitmap.
                    ' The original data is maintained if KeepOriginalFormat=True
                    cImageData.CacheSourceInfo tSource, tHandle, 0&, False, True        ' retrieve image source
                    GdipDisposeImage tHandle: Set cImageData = New cGDIpMultiImage     ' destroy GDI+ image
                    If TypeOf tSource Is IUnknown Then                                      ' IStream?
                        If IStreamToArray(ObjPtr(tSource), tData) Then Set tPic = ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
                        Erase tData()                                                       ' no longer needed
                    ElseIf VarType(tSource) = (vbArray Or vbByte) Then                      ' array?
                        MoveArrayToVariant tSource, tData(), False
                        Set tPic = ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
                        MoveArrayToVariant tSource, tData(), True
                    ElseIf VarType(tSource) = vbString Then                                 ' file name?
                        tHandle = GetFileHandle(CStr(tSource), False)
                        If tHandle <> INVALID_HANDLE_VALUE Then
                            ReDim tData(0 To GetFileSize(tHandle, 0&) - 1&)
                            SetFilePointer tHandle, 0&, 0&, 0&
                            ReadFile tHandle, tData(0), UBound(tData) + 1&, lResult, ByVal 0&
                            CloseHandle tHandle
                            If lResult > UBound(tData) Then Set tPic = ArrayToPicture(VarPtr(tData(0)), UBound(tData) + 1&)
                            If KeepOriginalFormat Then
                                MoveArrayToVariant tSource, tData(), True
                            Else
                                Erase tData()
                            End If
                        End If
                    End If
                    lResult = lvicPicTypeNone                                               ' default to failure
                    If tPic Is Nothing Then                                                 ' VB loaded JPG?
                        tSource = Empty
                    Else
                        If GdipCreateBitmapFromHBITMAP(tPic.Handle, 0&, lResult) = 0& Then  ' create bmp from handle
                            cImageData.CacheSourceInfo tSource, lResult, lvicPicTypeJPEG, KeepOriginalFormat, False
                        End If
                        Set tPic = Nothing
                    End If
                End If
        End Select
    Else
        Set cImageData = Nothing
    End If
    Set g_NewImageData = cImageData
    If AsyncClient Is Nothing Then
        If lResult <> lvicPicTypeNone Or ReturnEmptyClassOnFailure = True Then Set LoadImage = New GDIpImage
        Set g_NewImageData = Nothing
    Else
        If lResult = lvicPicTypeNone Then
            Set g_NewImageData = Nothing
        Else
            Set LoadImage = AsyncClient
        End If
    End If

End Function

Public Function LoadBlankImage(SS As SAVESTRUCT, Optional ByVal ReturnEmptyClassOnFailure As Boolean) As GDIpImage

    ' called only by SaveImage method when attempting to create a blank GDI+ bitmap

    Dim hHandle As Long, lCount As Long
    Dim cPal As cColorReduction, tSS As SAVESTRUCT
    
    If SS.Width > 0& And SS.Height > 0& Then
        SS.ColorDepth = (SS.ColorDepth And Not lvicApplyAlphaTolerance)
        If SS.ColorDepth < lvicConvert_BlackWhite Or SS.ColorDepth > lvicConvert_TrueColor32bpp_pARGB Then SS.ColorDepth = lvicConvert_TrueColor32bpp_RGB
        SS.ColorDepth = ColorDepthToColorType(SS.ColorDepth, 0&)
        
        If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, SS.ColorDepth, ByVal 0&, hHandle) = 0& Then
        
            Select Case (SS.ColorDepth And &HFF00&) \ &H100&
            Case 1&: lCount = 2&
            Case 4&: lCount = 16&
            Case 8&: lCount = 256&
            Case Else
                Dim hGraphics As Long
                If Not (SS.RSS.FillBrushGDIplus_Handle = 0& And SS.RSS.FillColorUsed = False) Then
                    If GdipGetImageGraphicsContext(hHandle, hGraphics) = 0& Then
                        If SS.RSS.FillBrushGDIplus_Handle Then
                            GdipFillRectangleI hGraphics, SS.RSS.FillBrushGDIplus_Handle, 0&, 0&, SS.Width, SS.Height
                        ElseIf (SS.ColorDepth = lvicColor32bppAlpha Or SS.ColorDepth = lvicColor32bppAlphaMultiplied) Then
                            GdipGraphicsClear hGraphics, SS.RSS.FillColorARGB
                        Else
                            GdipGraphicsClear hGraphics, SS.RSS.FillColorARGB Or &HFF000000
                        End If
                        GdipDeleteGraphics hGraphics
                    End If
                End If
            End Select
            If lCount Then
                Set cPal = New cColorReduction
                cPal.ApplyPaletteToHandle lCount, SS.Palette_Handle, hHandle
                Set cPal = Nothing
            End If
        End If
    End If
    If hHandle Then
        Set g_NewImageData = New cGDIpMultiImage
        g_NewImageData.CacheSourceInfo Empty, hHandle, lvicPicTypeBitmap, True, False
        Set LoadBlankImage = New GDIpImage
        Select Case SS.reserved1
            Case Is < lvicSaveAsMetafile
                SS.reserved1 = lvicSaveAsBitmap
            Case Is > lvicSaveAsPAM
                SS.reserved1 = lvicSaveAsBitmap
            Case lvicSaveAs_HCURSOR
                SS.reserved1 = lvicSaveAsCursor
            Case lvicSaveAs_HICON
                SS.reserved1 = lvicSaveAsIcon
        End Select
        If SS.reserved1 > lvicSaveAsBitmap Then
            tSS.CompressionJPGQuality = SS.CompressionJPGQuality
            tSS.CursorHotSpotX = SS.CursorHotSpotX
            tSS.CursorHotSpotY = SS.CursorHotSpotY
            tSS.Height = SS.Height: tSS.Width = SS.Width
            SaveImage LoadBlankImage, LoadBlankImage, SS.reserved1, tSS
        End If
    ElseIf ReturnEmptyClassOnFailure Then
        Set g_NewImageData = Nothing
        Set LoadBlankImage = New GDIpImage
    End If

End Function

Private Function pvProcessArraySource(cImageData As cGDIpMultiImage, inArray() As Byte, CacheData As Boolean, _
                                        Optional KnownFormat As ImageFormatEnum) As ImageFormatEnum

    ' helper function for LoadImage.
    ' This is the core processing routine. This routine is blind to the origination of the data
    '   passed in the inArray parameter. It does not know what file extension it came from (if any),
    '   whether passed from inside/oustide this project, or any details. Therefore, the data is
    '   sent to a ordered list of classes to determine the image format. The classes are designed
    '   to, as quickly as possible, validate whether the data is an image format the class recognizes
    ' If array was preprocessed by another routine, the KnownFormat will identify the image format
    
    Dim cIcons As cFunctionsICO, cBmps As cFunctionsBMP
    Dim cTGA As cFunctionsTGA, cMP3 As cFunctionsMP3
    Dim cGIF As cFunctionsGIF, cPNG As cFunctionsPNG
    Dim cPCX As cFunctionsPCX, cPAM As cFunctionsPNM
    Dim cAVI As cFunctionsAVI
    
    Dim lValue As Long, Cx As Long, Cy As Long, lColorType As Long
    Dim hHandle As Long, IStream As IUnknown, tSource As Variant, tSizeF As RECTF
    
    ' There are several types of image formats we must do manually, so test for those first
    
    ' #1. PNGs. Handled uniquely to support animated PNG (APNG)
    If KnownFormat = lvicPicTypePNG Or KnownFormat = lvicPicTypeUnknown Then
        Set cPNG = New cFunctionsPNG
        Select Case cPNG.IsPNGResource(cImageData, inArray(), CacheData)
        Case 0: ' not a PNG
        Case 1: ' single-frame png; let GDI+ load it
            KnownFormat = lvicPicTypePNG
        Case Else
            pvProcessArraySource = lvicPicTypePNG
            Exit Function
        End Select
        Set cPNG = Nothing
    End If
    
    ' #2. GIFs. Handled manually. GDI+ has issues in some cases
    If KnownFormat = lvicPicTypeGIF Or KnownFormat = lvicPicTypeUnknown Then
        Set cGIF = New cFunctionsGIF
        If cGIF.IsGIFResource(cImageData, inArray(), CacheData) Then
            pvProcessArraySource = lvicPicTypeGIF
            Exit Function
        End If
        Set cGIF = Nothing
    End If
    
    ' #3. Bitmaps that use the alpha channel
    If KnownFormat = lvicPicTypeUnknown Or KnownFormat = lvicPicTypeBitmap Then
        Set cBmps = New cFunctionsBMP
        lValue = cBmps.IsBitmapResource(inArray(), Cx, Cy, lColorType)
        Select Case lValue
            Case 0: KnownFormat = lvicPicTypeUnknown            ' non-bitmap
            Case 1: KnownFormat = lvicPicTypeBitmap             ' non-32bpp, let GDI+ load it
            Case Else                                           ' handle 32bpp alpha-channel bitmaps manually
                cImageData.CacheSourceInfo Empty, lValue, lvicPicTypeBitmap, False, False
                pvProcessArraySource = lvicPicTypeBitmap
                Exit Function
        End Select
        Set cBmps = Nothing
    End If
    
    ' #4. AVIs. GDI+ has no built-in support for these
    If KnownFormat = lvicPicTypeAVI Or KnownFormat = lvicPicTypeUnknown Then
        Set cAVI = New cFunctionsAVI
        If cAVI.IsAVIResource(inArray(), cImageData) Then
            pvProcessArraySource = lvicPicTypeAVI
            Exit Function
        End If
        Set cAVI = Nothing
    End If
    
    ' #5. Icons, cursors & animated cursoSS. GDI+ sucks with these
    If KnownFormat = lvicPicTypeUnknown Or KnownFormat = lvicPicTypeIcon Or _
        KnownFormat = lvicPicTypeCursor Or KnownFormat = lvicPicTypeAnimatedCursor Then
        Set cIcons = New cFunctionsICO
        lValue = cIcons.IsIconResource(cImageData, inArray(), CacheData)
        If lValue Then                                                  ' else not a icon/cursor
            pvProcessArraySource = lValue
            Exit Function
        End If
        Set cIcons = Nothing
        KnownFormat = lvicPicTypeUnknown
    End If
    
    ' #6. PNM (PGM, PBN, PCM files) & also PAM; all of which GDI+ has no built-in support
    If KnownFormat = lvicPicTypeUnknown Or KnownFormat = lvicPicTypePNM Or KnownFormat = lvicPicTypePAM Then
        Set cPAM = New cFunctionsPNM
        lValue = cPAM.IsPNMResource(inArray(), 0&)
        If lValue Then
            hHandle = cPAM.LoadPNMResource(inArray())
            If hHandle Then
                cImageData.CacheSourceInfo VarPtrArray(inArray()), hHandle, lValue, CacheData, False
                pvProcessArraySource = lValue
                Exit Function
            End If
        End If
        Set cPAM = Nothing
    End If
    
    ' #7. MP3 with embedded images. Can contain embedded images
    If KnownFormat = lvicPicTypeUnknown Or KnownFormat = 99& Then
        Set cMP3 = New cFunctionsMP3
        If KnownFormat = 99& Then
            hHandle = cMP3.LoadMP3(inArray(), tSource)
        Else
            If cMP3.IsMP3Resource(inArray, 0&) Then hHandle = cMP3.LoadMP3(inArray(), tSource)
        End If
        If hHandle Then
            ' the return object may be an array (only 1 image from MP3) or TIFF IStream (multiple images)
            If VarType(tSource) = (vbArray Or vbByte) Then
                MoveArrayToVariant tSource, inArray(), False
                pvProcessArraySource = pvProcessArraySource(cImageData, inArray(), CacheData, lvicPicTypeUnknown)
            Else
                Erase inArray()
                cImageData.CacheSourceInfo tSource, hHandle, lvicPicTypeTIFF, CacheData, False
                pvProcessArraySource = lvicPicTypeTIFF
            End If
            Exit Function
        End If
        Set cMP3 = Nothing
    End If

    ' #8. Targa (TGA). Not supported by GDI+
    If KnownFormat = lvicPicTypeUnknown Or KnownFormat = lvicPicTypeTGA Then
        Set cTGA = New cFunctionsTGA
        If KnownFormat = lvicPicTypeUnknown Then
            lValue = cTGA.IsTGAResource(inArray(), 0&)
        Else
            lValue = 1&
        End If
        Select Case lValue
            Case -1: Exit Function ' corrupted format
            Case 0: KnownFormat = lvicPicTypeUnknown
            Case Else ' it's TGA
                hHandle = cTGA.LoadTGAResource(inArray())
                If hHandle = 0& Then Exit Function
                cImageData.CacheSourceInfo VarPtrArray(inArray()), hHandle, lvicPicTypeTGA, CacheData, False
                pvProcessArraySource = lvicPicTypeTGA
                Exit Function
        End Select
        Set cTGA = Nothing
    End If
    
    ' #9. PC Paintbrush (PCX). Not supported by GDI+
    If KnownFormat = lvicPicTypeUnknown Or KnownFormat = lvicPicTypePCX Then
        Set cPCX = New cFunctionsPCX
        If KnownFormat = lvicPicTypeUnknown Then
            lValue = cPCX.IsPCXResource(inArray(), 0&)
        Else
            lValue = 1&
        End If
        Select Case lValue
            Case -1: Exit Function ' corrupted format
            Case 0: KnownFormat = lvicPicTypeUnknown
            Case Else ' it's PCX
                hHandle = cPCX.LoadPCXResource(inArray())
                If hHandle = 0& Then Exit Function
                cImageData.CacheSourceInfo VarPtrArray(inArray()), hHandle, lvicPicTypePCX, CacheData, False
                pvProcessArraySource = lvicPicTypePCX
                Exit Function
        End Select
        Set cPCX = Nothing
    End If
    
    ' All others go to GDI+
    If UBound(inArray) > 56& Then
        Set IStream = IStreamFromArray(VarPtr(inArray(0)), UBound(inArray) + 1&)
        If Not IStream Is Nothing Then
            If GdipLoadImageFromStream(ObjPtr(IStream), hHandle) = 0& Then
                pvProcessArraySource = GetImageType(hHandle)
                
                If pvProcessArraySource = lvicPicTypeEMetafile Then
                    lValue = pvConvertNonPlaceableWMFtoWMF(inArray(), vbNullString)
                    If lValue = lvicPicTypeUnknown Then    ' don't load this image
                        GdipDisposeImage hHandle: hHandle = 0&
                        Set IStream = Nothing
                    Else
                        If lValue <> lvicPicTypeNone Then ' use what was passed
                            GdipDisposeImage hHandle: hHandle = lValue
                            MoveArrayToVariant tSource, inArray(), True
                            pvProcessArraySource = lvicPicTypeMetafile
                        Else
                            tSource = IStream
                        End If
                    End If
                ElseIf pvProcessArraySource = lvicPicTypeBitmap Then ' don't store raw data for bitmaps
                    lValue = CreateSourcelessHandle(hHandle)
                    If lValue Then
                        GdipDisposeImage hHandle
                        hHandle = lValue
                    End If
                Else
                    tSource = IStream
                End If
                If hHandle Then cImageData.CacheSourceInfo tSource, hHandle, pvProcessArraySource, CacheData, False
            End If
        End If
        Erase inArray()
    End If

End Function

Private Function pvProcessAssociatedIcon(cImageData As cGDIpMultiImage, CacheData As Boolean, AI As ASSOCIATEDICON) As ImageFormatEnum

    ' helper function for LoadImage.
    ' This retrieves an associated icon to passed file name or type
    ' If successful the retrieved icon is passed to another routine to convert to GDIpImage class
    
    Const SHGFI_PIDL As Long = &H8
    Const SHGFI_USEFILEATTRIBUTES As Long = &H10
    Const SHGFI_SYSICONINDEX As Long = &H4000&
    Const SHGFI_OPENICON As Long = &H2
    Const ILD_TRANSPARENT As Long = &H1
    Const IID_IImageList    As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"
    'Const IID_IImageList2   As String = "{192B9D83-50FC-457B-90A0-2B82A8B5DAE1}"
    
    Dim SHFI As SHFILEINFO, lFlags As Long, lAttr As Long
    Dim GUID(0 To 37) As Long, hIML As Long, pIML As IUnknown
    Dim lRtn As Long
    
    ' sanity checks first
    If AI.FileName = vbNullString Then
        If (AI.IconType And lvicAssocActualIcon) Then Exit Function
    End If
    
    If IsNumeric(AI.FileName) Then
        ' numeric values in FileName should be either a PIDL or icon index into a system image list
        If (AI.IconType And lvicAssocIconPIDL) Then
            lFlags = SHGFI_PIDL
        ElseIf (AI.IconType And lvicAssocIconIndex) Then
            lFlags = SHGFI_SYSICONINDEX
        End If
    ElseIf StrPtr(AI.FileName) = 0& Then
        AI.FileName = ""
    End If
    
    If lFlags = SHGFI_SYSICONINDEX Then                         ' get image list for small/large icon
        If AI.DesiredSize < lvicSHIL_ExtraLarge_48 Then
            hIML = SHGetFileInfo("", 0&, SHFI, Len(SHFI), lFlags Or SHGFI_USEFILEATTRIBUTES Or AI.DesiredSize)
        End If
        On Error Resume Next
        SHFI.iIcon = CLng(AI.FileName)
        On Error GoTo 0
    
    Else
        If AI.DesiredSize < lvicSHIL_Large_32 Then               ' validate passed icon size
            AI.DesiredSize = lvicSHIL_Large_32
        ElseIf AI.DesiredSize > lvicSHIL_Jumbo_256 Then
            AI.DesiredSize = lvicSHIL_Jumbo_256
        ElseIf AI.DesiredSize > lvicSHIL_ExtraLarge_48 And AI.DesiredSize < lvicSHIL_Jumbo_256 Then
            AI.DesiredSize = lvicSHIL_ExtraLarge_48
        End If
                                                        ' validate icon size supported by O/S (48x48 > )
        If AI.DesiredSize >= lvicSHIL_ExtraLarge_48 Then
            lRtn = GetVersion()
            Select Case (lRtn And &HFF&)
            Case Is > 5                 ' Vista or better
                lRtn = 2&
            Case 5                      ' XP or maybe not
                If ((lRtn And &HFF00&) \ &H100 > 0&) Then lRtn = 1& Else lRtn = 0&
            Case Else                   ' less than XP
                lRtn = 0&
            End Select
        
            If AI.DesiredSize = lvicSHIL_ExtraLarge_48 Then                 ' not applicable for less than XP
                If lRtn = 0 Then AI.DesiredSize = lvicSHIL_Large_32
            ElseIf (AI.DesiredSize = lvicSHIL_Jumbo_256) And (lRtn < 2&) Then ' only for Vista+
                If lRtn = 0 Then AI.DesiredSize = lvicSHIL_Large_32 Else AI.DesiredSize = lvicSHIL_ExtraLarge_48
            End If
        End If
                                                        ' build the flags & attributes API values
        If (AI.IconType And lvicAssocIconOpened) Then lFlags = lFlags Or SHGFI_OPENICON
        If (AI.IconType And lvicAssocActualIcon) Then
            If (lFlags And SHGFI_PIDL) = 0 Then
                If g_UnicodeSystem Then
                    lRtn = GetFileAttributesW(StrPtr(AI.FileName))
                Else
                    lRtn = GetFileAttributes(AI.FileName)
                End If
                If lRtn = INVALID_HANDLE_VALUE Then
                    AI.IconType = (AI.IconType And Not lvicAssocActualIcon)
                    lFlags = lFlags Or SHGFI_USEFILEATTRIBUTES
                    If Right$(AI.FileName, 1) = "\" Then lAttr = vbDirectory
                Else
                    If (lRtn And vbDirectory) = vbDirectory Then lAttr = vbDirectory
                End If
            End If
        Else
            If (lFlags And SHGFI_PIDL) = 0 Then
                If Right$(AI.FileName, 1) = "\" Then lAttr = vbDirectory
            End If
            lFlags = lFlags Or SHGFI_USEFILEATTRIBUTES
        End If
        lFlags = lFlags Or SHGFI_SYSICONINDEX
        If AI.DesiredSize < lvicSHIL_ExtraLarge_48 Then lFlags = lFlags Or AI.DesiredSize
                                                    ' call the API
        If g_UnicodeSystem Then ' unicode calls
            On Error Resume Next
            If (lFlags And SHGFI_PIDL) Then lRtn = CLng(AI.FileName) Else lRtn = StrPtr(AI.FileName)
            If Err Then lRtn = 0&
            On Error GoTo 0
            hIML = SHGetFileInfoW(ByVal lRtn, lAttr, VarPtr(SHFI), Len(SHFI), lFlags)
        Else                        ' ansi system
            If (lFlags And SHGFI_PIDL) Then
                hIML = SHGetFileInfo(ByVal CLng(AI.FileName), lAttr, SHFI, Len(SHFI), lFlags)
            Else
                hIML = SHGetFileInfo(ByVal AI.FileName, lAttr, SHFI, Len(SHFI), lFlags)
            End If
        End If
    End If
    
    ' on XP and above, the image list handle returned by SHGetFileInfo is not the ExtraLarge or Jumbo sized
    ' image lists as expected. We'll use SHGetImageList to get the correct handle
    
    If (Not hIML = 0&) Or (lFlags = SHGFI_SYSICONINDEX) Then
        If AI.DesiredSize >= lvicSHIL_ExtraLarge_48 Or hIML = 0& Then ' XP or greater O/S
            If IIDFromString(StrPtr(IID_IImageList), GUID(0)) = 0 Then
                On Error Resume Next
                lRtn = SHGetImageList(AI.DesiredSize, GUID(0), ByVal VarPtr(pIML))
                If lRtn = 0& Then
                    If Err Then     ' depending on service pack shell32 did not export SHGetImageList correctly
                        Err.Clear   ' so we try again using the ordinal exported
                        lRtn = SHGetImageListXP(AI.DesiredSize, GUID(0), ByVal VarPtr(pIML))
                        If Err Then lRtn = hIML ' assign any non-zero value; will be using the hIML value
                    End If
                End If
                On Error GoTo 0
                If lRtn = 0& Then hIML = ObjPtr(pIML)
            End If
        End If
        If hIML Then
            SHFI.hIcon = ImageList_GetIcon(hIML, SHFI.iIcon, ILD_TRANSPARENT)
            If SHFI.hIcon Then
                pvProcessAssociatedIcon = pvProcessGDIhandleSource(cImageData, SHFI.hIcon, True)
                DestroyIcon SHFI.hIcon
                AI.IndexReturned = SHFI.iIcon Or AI.DesiredSize * &H10000000
            End If
        End If
    End If

End Function

Private Function pvProcessURLSource(cImageData As cGDIpMultiImage, URL As String, CacheData As Boolean, SyncMode As Boolean) As ImageFormatEnum

    ' helper function for LoadImage.
    ' This function attempts to connect to the internet & download byte data from the passed URL
    ' If successful the bytes will be sent to another routine for processing
    ' Valid URL required and can be http, ftp or file

    ' InternetOpenURL api supports only HTTP, FTP, Gopher
    If StrComp(Left$(URL, 4), "http", vbTextCompare) Then
        If StrComp(Left$(URL, 3), "ftp", vbTextCompare) Then
            If StrComp(Left$(URL, 6), "gopher") Then Exit Function
        End If
    ElseIf Not SyncMode Then   ' http & https supported only
        If g_AsyncController Is Nothing Then Set g_AsyncController = New cAsyncController
        Dim cAsync As cAsyncClient
        Set cAsync = g_AsyncController.CreateClient(URL)
        If cAsync Is Nothing Then
            ' required dll/ocx not available, do sync download using async events
            If g_AsyncController.AsyncModeAvailable(True) = True Then
                ' would only fail if controller couldn't create needed delay window (99.9% unlikely)
                Set cAsync = New cAsyncClient
                cAsync.URL = URL            ' assign URL
                cAsync.SyncMode = True      ' set to sync vs async mode
            End If
        End If
        If Not cAsync Is Nothing Then
            cImageData.CacheSourceInfo cAsync, 0&, lvicPicTypeAsyncDL, CacheData, False
            pvProcessURLSource = lvicPicTypeAsyncDL
            Exit Function
        End If
    End If
    
    Dim lBytesRead As Long, bBuffer() As Byte, picData() As Byte
    Dim hInternet As Long, hFile As Long, lRead As Long
    Const INTERNET_FLAG_RELOAD = &H80000000
    Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
    Const CHUNK_SIZE = &H2000&

    On Error GoTo ExitRoutine
    
    If InternetGetConnectedState(lRead, 0&) Then
        If g_UnicodeSystem Then
            hInternet = InternetOpenW(StrPtr(App.EXEName), 0&, 0&, 0&, 0&)
            If hInternet Then hFile = InternetOpenUrlW(hInternet, StrPtr(URL), 0&, 0&, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, ByVal 0&)
        Else
            hInternet = InternetOpen(App.EXEName, 0&, 0&, 0&, 0&)
            If hInternet Then hFile = InternetOpenUrl(hInternet, URL, 0&, 0&, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, ByVal 0&)
        End If
        If hFile Then
            ReDim bBuffer(0 To CHUNK_SIZE)
            Do
                InternetReadFile hFile, bBuffer(0), CHUNK_SIZE, lRead
                If lRead Then
                    ReDim Preserve picData(0 To lBytesRead + lRead - 1&)
                    CopyMemory picData(lBytesRead), bBuffer(0), lRead
                    lBytesRead = lBytesRead + lRead
                    If lRead < CHUNK_SIZE Then Exit Do
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    
ExitRoutine:
    If hFile Then InternetCloseHandle hFile
    If hInternet Then
        InternetCloseHandle hInternet
        If lBytesRead Then pvProcessURLSource = pvProcessArraySource(cImageData, picData(), CacheData, lvicPicTypeUnknown)
    End If
    
End Function

Private Function pvProcessBase64(cImageData As cGDIpMultiImage, ByVal Base64String As String, CacheData As Boolean) As ImageFormatEnum
   
    ' helper function for LoadImage.
    ' Routine processes passed string as Base64. Format must be valid
    ' Once Base64 deciphered, passed to another routine for image processing
    ' ACKNOWLEDGEMENT: based on CVMichael's routine from http://www.vbforums.com/showthread.php?t=498548

    Dim Base64LUT() As Byte
    Dim K4 As Long, K3 As Long, lValue As Long
    Dim tLen As Long, outData() As Byte
    Dim inData() As Byte, tSA As SafeArray
    
    tLen = Len(Base64String)                ' length of string to process
    If tLen < 228& Then Exit Function       ' not long enough to be valid image
    
    ReDim Base64LUT(0 To 255)
    FillMemory Base64LUT(0), 255, 255       ' fill look-up table. All initialized at 255
    For K4 = 0& To 25&                      ' A-Z = 0-25
        Base64LUT(K4 + 65&) = K4
    Next
    For K4 = 26& To 51&                     ' a-z = 26-51
        Base64LUT(K4 + 71&) = K4
    Next
    For K4 = 52& To 61&                     ' 0-9 = 52-61
        Base64LUT(K4 - 4&) = K4
    Next
    Base64LUT(43) = 62                      ' + = 62
    Base64LUT(47) = 63                      ' / = 63
    
    On Error GoTo ExitRoutine
        
    K3 = tLen
    For K4 = 1& To tLen
        Select Case Asc(Mid$(Base64String, K4, 1))
        Case 65 To 90, 97 To 122, 48 To 57, 43, 47, 61
            ' valid characters: [+/=][A-Z][a-z][0-9]
        Case Else
            K3 = K4 - 1&                    ' invalid characters found
            Exit For                        ' process differently
        End Select
    Next
    If K4 < tLen Then                       ' shift out all invalid characters
        For K4 = K4 + 1& To tLen
            Select Case Asc(Mid$(Base64String, K4, 1))
            Case 65 To 90, 97 To 122, 48 To 57, 43, 47, 61
                ' valid characters: [+/=][A-Z][a-z][0-9]
                K3 = K3 + 1&
                Mid$(Base64String, K3, 1) = Mid$(Base64String, K4, 1)
            Case Else
                ' do nothing
            End Select
        Next
        tLen = K3
    End If
    
    If Mid$(Base64String, tLen - 1&, 1) = "=" Then   ' size out array
        ReDim outData(0 To ((tLen \ 4&) * 3& - 3&))
    ElseIf Right$(Base64String, 1) = "=" Then
        ReDim outData(0 To ((tLen \ 4&) * 3& - 2&))
    Else
        ReDim outData(0 To ((tLen \ 4&) * 3& - 1&))
    End If
    
    With tSA                                        ' overlay byte array on unicode string
        .cbElements = 1                             ' speed tweak. Better performance referencing via
        .cDims = 1                                  ' 1D byte array than using StrConv() or using
        .pvData = StrPtr(Base64String)              ' combinations of Asc(Mid$(s,x,n)) in loop below
        .rgSABound(0).cElements = tLen * 2&
    End With                                        ' apply the overlay
    CopyMemory ByVal VarPtrArray(inData), VarPtr(tSA), 4&
    
    K3 = 0&
    For K4 = 0& To (tLen \ 4&) * 8& - 9& Step 8&    ' process string, decode 4 bytes to 3 bytes
        lValue = Base64LUT(inData(K4 + 6&))
        lValue = lValue Or Base64LUT(inData(K4 + 4&)) * &H40&
        lValue = lValue Or Base64LUT(inData(K4 + 2&)) * &H1000&
        lValue = lValue Or Base64LUT(inData(K4)) * &H40000
                                                    ' translate decoded long to bytes
        outData(K3) = (lValue And &HFF0000) \ &H10000
        outData(K3 + 1&) = (lValue And &HFF00&) \ &H100&
        outData(K3 + 2&) = lValue And &HFF&
        K3 = K3 + 3&
    Next
    lValue = 0&                                     ' process final 4 characters
    
    If Base64LUT(inData(K4 + 6&)) <> 255 Then lValue = Base64LUT(inData(K4 + 6&))
    If Base64LUT(inData(K4 + 4&)) <> 255 Then lValue = lValue Or Base64LUT(inData(K4 + 4&)) * &H40&
    If Base64LUT(inData(K4 + 2&)) <> 255 Then lValue = lValue Or Base64LUT(inData(K4 + 2&)) * &H1000&
    If Base64LUT(inData(K4)) <> 255 Then lValue = lValue Or Base64LUT(inData(K4)) * &H40000
    
    tSA.pvData = 0&                                         ' set flag for err handler
    CopyMemory ByVal VarPtrArray(inData), tSA.pvData, 4&    ' remove overlay
    Base64String = vbNullString                             ' release some memory
    
    outData(K3) = (lValue And &HFF0000) \ &H10000
    If UBound(outData) >= (K3 + 1&) Then outData(K3 + 1&) = (lValue And &HFF00&) \ &H100&
    If UBound(outData) >= (K3 + 2&) Then outData(K3 + 2&) = lValue And &HFF&
    
    Erase Base64LUT()
    pvProcessBase64 = pvProcessArraySource(cImageData, outData(), CacheData, lvicPicTypeUnknown)
    
ExitRoutine:
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(inData), 0&, 4&
    
End Function

Private Function pvProcessFileSource(cImageData As cGDIpMultiImage, FileName As String, CacheData As Boolean, UnlockFile As Boolean) As ImageFormatEnum

    ' helper function for LoadImage.
    ' This routine will either allow GDI+ to load image from a file or will read the file contents
    ' into memory and call another routine to process those bytes.
    
    ' There are several types of image formats we must do manually, so test for those first
    ' KnownFormat, when filled in, is passed as result of partial processing from stdPicture or File
    
    ' The parsers called in this routine do not necessarily do complete validation, rather it scans the file
    ' and makes a quick determination whether manual or automatic image processing can be performed
    
    ' Modified July 2011. Will also pass a Base64 encoded string to another function

    Dim hHandle As Long, lResult As Long, lRead As Long
    Dim cBmps As cFunctionsBMP, cIcons As cFunctionsICO
    Dim cTGA As cFunctionsTGA, cMP3 As cFunctionsMP3
    Dim cPNG As cFunctionsPNG, cGIF As cFunctionsGIF
    Dim cPCX As cFunctionsPCX, cPAM As cFunctionsPNM
    Dim cDLL As cFunctionsDLL, cAVI As cFunctionsAVI
    Dim tSource As Variant
    Dim tArray() As Byte, bDefaultToGDIp As Boolean
    
    hHandle = GetFileHandle(FileName, False)        ' open the file
    If hHandle = INVALID_HANDLE_VALUE Then          ' a valid file? If not...
        If InStr(FileName, ".") = 0 Then            ' file may have been passed, but doesn't exist
            If InStr(FileName, "\") = 0 Then        ' if typical file chars found, assume file; but won't be Base64
                If InStr(FileName, ":") = 0 Then pvProcessFileSource = pvProcessBase64(cImageData, FileName, CacheData)
            End If
        ElseIf StrComp(Left$(FileName, 7), "file://", vbTextCompare) = 0 Then
            pvProcessFileSource = pvProcessFileSource(cImageData, Mid$(FileName, 8), CacheData, UnlockFile)
        ElseIf StrComp(Left$(FileName, 7), "file:\\", vbTextCompare) = 0 Then
            pvProcessFileSource = pvProcessFileSource(cImageData, Mid$(FileName, 8), CacheData, UnlockFile)
        End If
        Exit Function
    End If
    
    Set cDLL = New cFunctionsDLL                            ' handled separately to support binary extraction
    lResult = cDLL.IsBinaryResource(FileName, cImageData)
    If lResult Then ' is binary (1=has images, -1=no images)
        CloseHandle hHandle
        If lResult = 1& Then pvProcessFileSource = lvicPicTypeFromBinaries
        Exit Function                                       ' DLLs are never cached to an array
    End If
    Set cDLL = Nothing
    
    Set cAVI = New cFunctionsAVI                            ' handled separately to support AVI
    If cAVI.IsAVIResourceFile(FileName, cImageData) Then
        CloseHandle hHandle
        pvProcessFileSource = lvicPicTypeAVI
        Exit Function                                       ' AVIs are never cached to an array for loading
    End If
    Set cAVI = Nothing
    
    Set cMP3 = New cFunctionsMP3                        ' if MP3, handle manually to extract embedded images
    If cMP3.IsMP3Resource(tArray(), hHandle) Then
        CloseHandle hHandle
        Set cMP3 = Nothing
        pvProcessFileSource = pvProcessArraySource(cImageData, tArray(), CacheData, 99&)
        Exit Function
    End If
    Set cMP3 = Nothing
    
    If UnlockFile Then                                      ' we will be caching the file vs reading it
        lResult = lvicPicTypeUnknown
        bDefaultToGDIp = True
    Else
        Set cPNG = New cFunctionsPNG                        ' handled separately to support animated PNG
        Select Case cPNG.IsPNGResourceFile(hHandle)
        Case 0: ' not a pNG
        Case 1: bDefaultToGDIp = True
        Case Else: ' multi-frame png
            lResult = lvicPicTypePNG
            bDefaultToGDIp = True
        End Select
        Set cPNG = Nothing
    End If
    
    If bDefaultToGDIp = False Then                          ' handled manually. GDI+ has issues with some GIFs
        Set cGIF = New cFunctionsGIF
        If cGIF.IsGIFResourceFile(hHandle) Then
            lResult = lvicPicTypeGIF
            bDefaultToGDIp = True
        End If
        Set cGIF = Nothing
    End If
    
    If bDefaultToGDIp = False Then
        Set cBmps = New cFunctionsBMP                       ' handled manually if alpha channel used
        Select Case cBmps.IsBitmapResourceFile(hHandle)
            Case lvicColor32bppAlpha, lvicColor32bppAlphaMultiplied
                lResult = lvicPicTypeBitmap
                bDefaultToGDIp = True
            Case 0&                                         ' not a bitmap
            Case Else: bDefaultToGDIp = True                ' bitmap GDI+ can handle, let it do that
        End Select
        Set cBmps = Nothing
    End If
    
    If bDefaultToGDIp = False Then
        Set cIcons = New cFunctionsICO                      ' if icon/cursor, handle manually. GDI+ has issues
        If cIcons.IsIconResourceFile(cImageData, hHandle, CacheData) Then
            lResult = lvicPicTypeIcon
            bDefaultToGDIp = True
        End If
        Set cIcons = Nothing
    End If
    
    If bDefaultToGDIp = False Then
        Set cPAM = New cFunctionsPNM                        ' handled manually, GDI+ does not support PBM,PGM,PPM,PAM
        lResult = cPAM.IsPNMResource(tArray(), hHandle)
        If lResult Then bDefaultToGDIp = True
        Set cPAM = Nothing
    End If
    
    If bDefaultToGDIp = False Then
        Set cTGA = New cFunctionsTGA                        ' handled manually, GDI+ does not support TGA
        Select Case cTGA.IsTGAResource(tArray(), hHandle)
        Case -1: Exit Function
        Case 0
        Case Else
            lResult = lvicPicTypeTGA
            bDefaultToGDIp = True
        End Select
        Set cTGA = Nothing
    End If
    
    If bDefaultToGDIp = False Then
        Set cPCX = New cFunctionsPCX                        ' handled manually, GDI+ does not support PCX
        Select Case cPCX.IsPCXResource(tArray(), hHandle)
        Case -1: Exit Function
        Case 0
        Case Else
            lResult = lvicPicTypePCX
            bDefaultToGDIp = True
        End Select
        Set cPCX = Nothing
    End If
    
    If lResult <> lvicPicTypeNone Then   ' those that will be handled manually
        ReDim tArray(0 To GetFileSize(hHandle, 0&) - 1&)    ' or if entire file is cached, we simply
        SetFilePointer hHandle, 0&, 0&, 0&                  ' read the data & pass it along to be processed
        ReadFile hHandle, tArray(0), UBound(tArray) + 1&, lRead, ByVal 0& ' as an array
        CloseHandle hHandle
        If lRead > UBound(tArray) Then
            pvProcessFileSource = pvProcessArraySource(cImageData, tArray, CacheData, lResult)
            Exit Function
        End If
    ElseIf hHandle Then
        CloseHandle hHandle
    End If
    
    If GdipLoadImageFromFile(StrPtr(FileName), hHandle) = 0& Then
        pvProcessFileSource = GetImageType(hHandle)
        tSource = FileName
        If pvProcessFileSource = lvicPicTypeEMetafile Then
            lResult = pvConvertNonPlaceableWMFtoWMF(tArray(), FileName)
            If lResult = lvicPicTypeUnknown Then        ' abort; don't load
                GdipDisposeImage hHandle
            ElseIf lResult <> lvicPicTypeNone Then      ' use what was passed
                MoveArrayToVariant tSource, tArray(), True
                GdipDisposeImage hHandle: hHandle = lResult
                pvProcessFileSource = lvicPicTypeMetafile
            End If
        End If
        If hHandle Then cImageData.CacheSourceInfo tSource, hHandle, pvProcessFileSource, CacheData, False
    End If
    
End Function

Private Function pvProcessGDIhandleSource(cImageData As cGDIpMultiImage, ByVal hHandle As Long, CacheData As Boolean) As ImageFormatEnum

    ' helper function for LoadImage.
    ' Routine takes a GDI image handle (bitmap or icon only) and converts it
    '   to a GDI+ handle. Passed handle is not molested

    Dim tData() As Byte, tBits() As Long, tSA As SafeArray
    Dim lResult As Long, Cx As Long, Cy As Long
    Dim tBMP As BitmapData, tSize As RECTI
    Dim lFormat As Long, tHandle As Long
    Dim cIcons As cFunctionsICO
    
'    Private Type BITMAP
'        bmType As Long             tData(0-3)
'        bmWidth As Long            tData(4-7)
'        bmHeight As Long           tDAta(8-11)
'        bmWidthBytes As Long       tData(12-15)
'        bmPlanes As Integer        tData(16-17)
'        bmBitsPixel As Integer     tData(18-19)
'        bmBits As Long             tData(20-23)
'    End Type
    
    ReDim tData(0 To 23)                                            ' 24 bytes same as BITMAP UDT
    If GetGDIObject(hHandle, 24&, tData(0)) Then                    ' test for bitmap handle first
        ' need to handle 32bpp images separately
        If tData(18) = 32 Then                                      ' bit count position
            CopyMemory lFormat, tData(20), 4&                       ' bits pointer
            If lFormat Then
                CopyMemory Cx, tData(4), 4&                         ' width
                CopyMemory Cy, tData(8), 4&                         ' height
                Erase tData()
                With tSA
                    .cbElements = 4&
                    .cDims = 2
                    .rgSABound(0).cElements = Abs(Cy)
                    .rgSABound(1).cElements = Cx
                    .pvData = lFormat
                End With
                CopyMemory ByVal VarPtrArray(tBits), VarPtr(tSA), 4&
                lFormat = ValidateAlphaChannel(tBits())
                CopyMemory ByVal VarPtrArray(tBits), 0&, 4&
                If lFormat > lvicColor32bpp Then
                    If Cy > 0& Then
                        tSA.pvData = tSA.pvData + (Cy - 1&) * Cx * 4
                        Cx = -Cx
                    End If
                    tSize.nHeight = Abs(Cy): tSize.nWidth = Abs(Cx)
                    If GdipCreateBitmapFromScan0(tSize.nWidth, tSize.nHeight, 0&, lFormat, ByVal 0&, lResult) = 0& Then
                        tBMP.Scan0Ptr = tSA.pvData
                        tBMP.stride = Cx * 4&
                        If GdipBitmapLockBits(lResult, tSize, ImageLockModeWrite Or ImageLockModeUserInputBuf, lFormat, tBMP) Then
                            GdipDisposeImage lResult
                            lResult = 0&
                        Else
                            GdipBitmapUnlockBits lResult, tBMP
                            hHandle = 0&
                        End If
                    End If
                End If
            End If
        End If
        If hHandle Then Call GdipCreateBitmapFromHBITMAP(hHandle, 0&, lResult)
        If lResult Then
            cImageData.CacheSourceInfo Empty, lResult, lvicPicTypeBitmap, CacheData, False
            pvProcessGDIhandleSource = lvicPicTypeBitmap            ' stand-alone image
        End If
    Else                                                            ' test for icon handle
        Set cIcons = New cFunctionsICO
        lResult = cIcons.HICONtoArray(cImageData, hHandle, CacheData)
        If lResult Then                                             ' if successful, array is icon format
            pvProcessGDIhandleSource = lResult
            Set cIcons = Nothing
        End If
    End If

End Function

Private Function pvProcessObjectSource(cImageData As cGDIpMultiImage, Source As Variant, inArray() As Byte, CacheData As Boolean, SyncMode As Boolean) As ImageFormatEnum

    ' helper function for LoadImage.
    ' This function will handle various types of Objects: Clipboard, Data, stdPicture, GDIpImage, Screen
    ' The routine's responsibility is to extract data and pass off to other routines for processing
    '   with sole exception being the Screen object which is processed here

    Dim tPic As StdPicture, hHandle As Long, lSize As Long, sText As String
    Dim tArray() As Byte, colFiles As Collection, lPtr As Long, lCFformat As Long
    
    If TypeOf Source Is StdPicture Then
        Set tPic = Source
        pvProcessObjectSource = pvProcessStdPicSource(cImageData, tPic, CacheData)
    
    ElseIf TypeOf Source Is Clipboard Then
        
        If g_ClipboardFormat Then
            If Clipboard.GetFormat(g_ClipboardFormat) Then
                If OpenClipboard(0&) Then
                    hHandle = GetClipboardData(g_ClipboardFormat And &HFFFF&)
                    If hHandle Then
                        lSize = GlobalSize(hHandle)
                        If lSize > 0& Then
                            ReDim inArray(0 To lSize - 1&)
                            lSize = GlobalLock(hHandle)
                            CopyMemory inArray(0), ByVal lSize, UBound(inArray) + 1&
                        End If
                        GlobalUnlock hHandle
                    End If
                    CloseClipboard
                End If
                If lSize > 0& Then pvProcessObjectSource = pvProcessArraySource(cImageData, inArray(), CacheData, lvicPicTypeUnknown)
            End If
        End If
        
        If pvProcessObjectSource = lvicPicTypeNone Then
            
            On Error Resume Next
            If Clipboard.GetFormat(vbCFFiles) Then
                If GetPastedFileData(colFiles, vbCFFiles) Then
                    For lPtr = 1 To colFiles.Count
                        pvProcessObjectSource = pvProcessFileSource(cImageData, colFiles.Item(lPtr), CacheData, lvicPicTypeUnknown)
                        If pvProcessObjectSource Then Exit For
                    Next
                End If
            ElseIf Clipboard.GetFormat(CF_FILECONTENTS) Then
                If Clipboard.GetFormat(CF_FILEGROUPDESCRIPTORW) Then
                    pvProcessObjectSource = GetPastedFileData(colFiles, CF_FILEGROUPDESCRIPTORW, cImageData, CacheData)
                End If
            End If
            If pvProcessObjectSource = lvicPicTypeNone Then
                ' patch to prevent following scenario:
                ' WordPad-like app that when text copied to clipboard, it also creates metafile of the text
                ' In this case, the metafile is used as the image vs. the text
                If Clipboard.GetFormat(vbCFEMetafile) Then
                    'Set tPic = Clipboard.GetData(vbCFEMetafile)
                    lCFformat = vbCFEMetafile
                ElseIf Clipboard.GetFormat(vbCFMetafile) Then
                    'Set tPic = Clipboard.GetData(vbCFMetafile)
                    lCFformat = vbCFMetafile
                ElseIf Clipboard.GetFormat(vbCFBitmap) Then
                    Set tPic = Clipboard.GetData(vbCFBitmap): lCFformat = -1&
                ElseIf Clipboard.GetFormat(vbCFDIB) Then
                    Set tPic = Clipboard.GetData(vbCFDIB): lCFformat = -1&
                End If
                If lCFformat > -1& Then
                    Call GetPastedFileData(Nothing, vbCFText, Nothing, , sText)
                    If Len(sText) Then
                        pvProcessObjectSource = pvProcessFileSource(cImageData, sText, CacheData, True)
                        If pvProcessObjectSource = lvicPicTypeNone Then pvProcessObjectSource = pvProcessURLSource(cImageData, sText, CacheData, SyncMode)
                    End If
                    If pvProcessObjectSource = lvicPicTypeNone Then
                        If lCFformat > 0& Then Set tPic = Clipboard.GetData(lCFformat)
                    End If
                End If
                If Not tPic Is Nothing Then
                    If tPic.Handle = 0& Then
                        Set tPic = Nothing
                    Else
                        pvProcessObjectSource = pvProcessStdPicSource(cImageData, tPic, CacheData)
                    End If
                End If
            End If
            On Error GoTo 0
            
        End If
        
    ElseIf TypeOf Source Is DataObject Then
        ' get drag/drop object (if file, first file name used)
        ' set sourceObj = that file & next section processes it
        
        ' patch to prevent following scenario:
        ' An app that when text dragged, it may create metafile of the dragged text
        ' In this case, the metafile is used as the image vs. the text
        On Error Resume Next
        If g_ClipboardFormat Then
            If Source.GetFormat(g_ClipboardFormat) Then
                tArray() = Source.GetData(g_ClipboardFormat)
                If Err Then
                    Err.Clear
                ElseIf NormalizeArray(tArray(), inArray()) Then
                    Erase tArray()
                    pvProcessObjectSource = pvProcessArraySource(cImageData, inArray(), CacheData, lvicPicTypeUnknown)
                End If
            End If
        End If
        If pvProcessObjectSource = lvicPicTypeNone Then
            If Source.GetFormat(vbCFFiles) Then
                If GetDroppedFileNames(Source, vbCFFiles) Then
                    For lPtr = 1 To Source.Files.Count
                        pvProcessObjectSource = pvProcessFileSource(cImageData, Source.Files.Item(lPtr), CacheData, lvicPicTypeUnknown)
                        If pvProcessObjectSource Then Exit For
                    Next
                End If
            ElseIf Source.GetFormat(CF_FILECONTENTS) Then
                If Source.GetFormat(CF_FILEGROUPDESCRIPTORW) Then
                    pvProcessObjectSource = GetDroppedFileNames(Source, &HFFFF& And CF_FILEGROUPDESCRIPTORW, , cImageData, CacheData)
                End If
            End If
            If pvProcessObjectSource = lvicPicTypeNone Then
                If Source.GetFormat(vbCFEMetafile) Then
                    'Set tPic = Source.GetData(vbCFEMetafile)
                    lCFformat = vbCFEMetafile
                ElseIf Source.GetFormat(vbCFMetafile) Then
                    'Set tPic = Source.GetData(vbCFMetafile)
                    lCFformat = vbCFMetafile
                ElseIf Source.GetFormat(vbCFBitmap) Then
                    Set tPic = Source.GetData(vbCFBitmap)
                ElseIf Source.GetFormat(vbCFDIB) Then
                    Set tPic = Source.GetData(vbCFDIB)
                End If
                If lCFformat > -1& Then
                    If GetDroppedFileNames(Source, vbCFText, sText) Then
                        pvProcessObjectSource = pvProcessFileSource(cImageData, sText, CacheData, True)
                        If pvProcessObjectSource = lvicPicTypeNone Then pvProcessObjectSource = pvProcessURLSource(cImageData, sText, CacheData, SyncMode)
                    End If
                    If pvProcessObjectSource = lvicPicTypeNone Then
                        If lCFformat > 0& Then Set tPic = Clipboard.GetData(lCFformat)
                    End If
                End If
                If Not tPic Is Nothing Then
                    If tPic.Handle = 0& Then
                        Set tPic = Nothing
                    Else
                        pvProcessObjectSource = pvProcessStdPicSource(cImageData, tPic, CacheData)
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    
    ElseIf TypeOf Source Is GDIpImage Then
        If Source.Handle = 0& Then
            pvProcessObjectSource = -1&
        ElseIf Source.ExtractImageData(inArray(), hHandle) = True Then
            pvProcessObjectSource = pvProcessArraySource(cImageData, inArray(), CacheData, hHandle)
        End If
        
    ElseIf TypeOf Source Is Screen Then
    
        Dim dDC As Long, tDC As Long
        If GdipCreateBitmapFromScan0(Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY, 0&, lvicColor24bpp, ByVal 0&, hHandle) = 0& Then
            If GdipGetImageGraphicsContext(hHandle, lPtr) = 0& Then
                GdipGetDC lPtr, tDC
                dDC = GetDC(GetDesktopWindow())
                BitBlt tDC, 0&, 0&, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY, dDC, 0&, 0&, vbSrcCopy
                ReleaseDC GetDesktopWindow(), dDC
                GdipReleaseDC lPtr, tDC
                GdipDeleteGraphics lPtr
                cImageData.CacheSourceInfo Empty, hHandle, lvicPicTypeBitmap, False, False
                pvProcessObjectSource = lvicPicTypeBitmap
            Else
                GdipDisposeImage hHandle
            End If
        End If
        
    End If

End Function

Private Function pvProcessStdPicSource(cImageData As cGDIpMultiImage, thePic As StdPicture, CacheData As Boolean) As ImageFormatEnum

    ' helper function for LoadImage.
    ' Routine extracts data from a stdPicture object & process that data
    ' NOTE: Sometimes, the stdPicture object can contain original data. For example, when an image control
    '   is assigned a GIF image during design view, that original GIF data is actually kept by the stdPicture
    '   object and can be retrieved and processed as a GIF vs a bitmap. But if same GIF was loaded into a
    '   control during runtime, that original GIF format is not kept & only a bitmap version of the GIF is.

    Dim lResult As Long, tHandle As Long, ipic As IPicture, IStream As IUnknown
    Dim cBmps As cFunctionsBMP
    Dim tArray() As Byte, bSaveMemCopy As Boolean
    
    If thePic.Handle = 0& Then Exit Function
    
    If thePic.Type = vbPicTypeIcon Then             ' load icons from handle
        pvProcessStdPicSource = pvProcessGDIhandleSource(cImageData, thePic.Handle, CacheData)
        Exit Function
    End If
    
    On Error Resume Next
    Set ipic = thePic: If ipic Is Nothing Then Exit Function
    Set IStream = IStreamFromArray(0&, 0&)
    If IStream Is Nothing Then Exit Function
    
    ' have VB's stdPicture save to our stream
    
    ' Don't know why this works the way it does. When a stdPicture object has original data maintained,
    '   passing FALSE in 2nd param, returns that original data, but errors if original data is not maintained.
    '   But passing a boolean variable returns a bitmap version of that original data instead.
    If ipic.KeepOriginalFormat Then
        ipic.SaveAsFile ByVal ObjPtr(IStream), False, lResult
        If lResult = 0& Then ipic.SaveAsFile ByVal ObjPtr(IStream), bSaveMemCopy, lResult
    Else
        ipic.SaveAsFile ByVal ObjPtr(IStream), bSaveMemCopy, lResult
    End If
    On Error GoTo 0
    If lResult = 0& Then Exit Function
    
    ' NOTE: You will see I call GdipLoadImageFromStream twice if 1st call fails. Again, not sure exactly why
    '   function fails first time & only assume it has to do with using IUnknown vs a valid IStream object
    
    If thePic.Type = vbPicTypeBitmap Then
        If IStreamToArray(ObjPtr(IStream), tArray) = False Then Exit Function
        
        ' any bitmap depth coming from a stdPicture is assumed to not use the alpha-channel
        Set cBmps = New cFunctionsBMP
        If cBmps.IsBitmapResource(tArray, 0&, 0&, 0&) Then
            pvProcessStdPicSource = pvProcessArraySource(cImageData, tArray(), CacheData, lvicPicTypeBitmap)
        Else                                                            ' gif or jpg most likely
            Erase tArray()
            If GdipLoadImageFromStream(ObjPtr(IStream), lResult) Then
                Call GdipLoadImageFromStream(ObjPtr(IStream), lResult)
            End If
            If lResult Then
                pvProcessStdPicSource = GetImageType(lResult)
                cImageData.CacheSourceInfo IStream, lResult, pvProcessStdPicSource, CacheData, False
            End If
        End If
        
    ElseIf thePic.Type = vbPicTypeMetafile Or thePic.Type = vbPicTypeEMetafile Then
        Erase tArray()
        If GdipLoadImageFromStream(ObjPtr(IStream), lResult) Then
            Call GdipLoadImageFromStream(ObjPtr(IStream), lResult)
        End If
        If lResult Then
            pvProcessStdPicSource = GetImageType(lResult)
            cImageData.CacheSourceInfo IStream, lResult, pvProcessStdPicSource, CacheData, False
        End If
    End If

End Function

Public Function ColorDepthToColorType(inDepth As Long, Handle As Long) As Long

    ' Routine converts the AICGlobals.ColorDepthEnum to AICGlobals.ImageColorFormatEnum values

    Select Case inDepth
        Case lvicNoColorReduction, lvicDefaultReduction
            GdipGetImagePixelFormat Handle, ColorDepthToColorType
            Select Case ColorDepthToColorType
                Case lvicColor16bpp555Alpha, lvicColor64bppAlpha: ColorDepthToColorType = lvicColor32bppAlpha
                Case lvicColor64bppAlphaMultiplied: ColorDepthToColorType = lvicColor32bppAlphaMultiplied
                Case lvicColor16bppGrayscale: ColorDepthToColorType = lvicColor8bpp
                Case lvicColor16bpp555, lvicColor16bpp565, lvicColor48bpp: ColorDepthToColorType = lvicColor24bpp
            End Select
        Case lvicConvert_BlackWhite: ColorDepthToColorType = lvicColor1bpp
        Case lvicConvert_16Colors: ColorDepthToColorType = lvicColor4bpp
        Case lvicConvert_256Colors: ColorDepthToColorType = lvicColor8bpp
        Case lvicConvert_TrueColor24bpp: ColorDepthToColorType = lvicColor24bpp
        Case lvicConvert_TrueColor32bpp_RGB: ColorDepthToColorType = lvicColor32bpp
        Case lvicConvert_TrueColor32bpp_ARGB: ColorDepthToColorType = lvicColor32bppAlpha
        Case lvicConvert_TrueColor32bpp_pARGB: ColorDepthToColorType = lvicColor32bppAlphaMultiplied
        Case lvicConvert_TrueColor32bpp_pARGB + 1&: ColorDepthToColorType = lvicColor32bppAlpha
    End Select


End Function

Public Function GetImageType(Handle As Long) As Long

    ' Routine returns the format loaded via GDI+

    On Error Resume Next
    If Handle = 0& Then Exit Function
    
    Dim GUID(0 To 3) As Long, sGUID As String, lRet As Long
    ' ACKNOWLEDGEMENT: http://com.it-berater.org/gdiplus/noframes/GdiPlus_constants.htm
    Const ImageFormatBMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatEMF As String = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatGIF As String = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatIcon As String = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatJPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatMemoryBMP As String = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatPNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatTIFF As String = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatUndefined As String = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatWMF As String = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
    ' note to self: haven't seen any images return this yet. Curious
    Const ImageFormatEXIF As String = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"

    If GdipGetImageRawFormat(Handle, VarPtr(GUID(0))) = 0& Then
        sGUID = String$(40, vbNullChar)
        lRet = StringFromGUID2(VarPtr(GUID(0)), StrPtr(sGUID), 40&)
        Select Case Left$(sGUID, lRet - 1)
            Case ImageFormatBMP, ImageFormatMemoryBMP: GetImageType = lvicPicTypeBitmap
            Case ImageFormatEMF: GetImageType = lvicPicTypeEMetafile
            Case ImageFormatGIF: GetImageType = lvicPicTypeGIF
            Case ImageFormatIcon: GetImageType = lvicPicTypeIcon
            Case ImageFormatJPEG: GetImageType = lvicPicTypeJPEG
            Case ImageFormatPNG: GetImageType = lvicPicTypePNG
            Case ImageFormatTIFF: GetImageType = lvicPicTypeTIFF
            Case ImageFormatWMF: GetImageType = lvicPicTypeMetafile
            Case Else: GetImageType = lvicPicTypeUnknown
        End Select
    Else
        GetImageType = lvicPicTypeUnknown
    End If
    
End Function

Public Function SaveImage(Picture As GDIpImage, Destination As Variant, _
                                    ByVal ImageFormat As SaveAsFormatEnum, _
                                    Optional SaveOptions As Variant) As Boolean
                                    
    ' Saves an image to multiple formats and/or objects
    ' A bit lengthy due to the number of possible combinations.
    ' Most of routine is validation and passes saving off to individual format routines

    Dim MTP As MULTIPAGETIFFSTRUCT, MIS As MULTIIMAGESAVESTRUCT, SS As SAVESTRUCT
    Dim tmpDataObj As DataObject
    Dim vData As Variant, bIsMasked As Boolean
    Dim lPage As Long, lResult As Long, hFileHandle As Long
    Dim imgData() As Byte, Cx As Long, Cy As Long
    Dim wrkPic As Long, lParam As Long, lbOffset As Long
    Dim bReleaseMultTiff As Boolean, bProcess As Boolean
    Dim bPP As Long, DestType As SaveAsMedium
    Dim cPal As cColorReduction
    
    Dim tSize As RECTI, hGraphics As Long, srcDepth As Long

    ' validate passed parameters
    If ImageFormat < lvicSaveAsCurrentFormat Or ImageFormat > lvicSaveAsPAM Then
        ImageFormat = lvicSaveAsCurrentFormat
    End If
    
    ' validate Destination object & set tracking flag
    If IsObject(Destination) Then
        If TypeOf Destination Is GDIpImage Then
            DestType = saveTo_GDIplus
        ElseIf TypeOf Destination Is StdPicture Then
            DestType = saveTo_stdPicture
        ElseIf TypeOf Destination Is Clipboard Then
            DestType = saveTo_Clipboard
        ElseIf TypeOf Destination Is DataObject Then
            DestType = saveTo_DataObject
        Else
            Exit Function                                       ' invalid parameter value
        End If
    ElseIf VarType(Destination) = vbString Then
        If Destination = vbNullString Then Exit Function        ' else destType=0
    ElseIf VarType(Destination) = (vbArray Or vbByte) Then
        DestType = saveTo_Array
    ElseIf VarType(Destination) = vbLong Then
        Select Case ImageFormat
        Case lvicSaveAs_HBITMAP, lvicSaveAs_HICON, lvicSaveAs_HCURSOR
            DestType = saveTo_GDIhandle
        Case Else
            Exit Function                                       ' invalid parameter
        End Select
    End If
    If DestType < saveTo_GDIhandle Then
        If ImageFormat = lvicSaveAs_HBITMAP Then
            ImageFormat = lvicSaveAsBitmap
        ElseIf ImageFormat = lvicSaveAs_HCURSOR Then
            ImageFormat = lvicSaveAsCursor
        ElseIf ImageFormat = lvicSaveAs_HICON Then
            ImageFormat = lvicSaveAsIcon
        End If
    End If
    
    ' validate the RenderingStyle or MultiImageTIFF structure
    If Not IsMissing(SaveOptions) Then
        If VarType(SaveOptions) = vbUserDefinedType Then    ' UDT passed
            On Error Resume Next
            If LenB(SaveOptions) = LenB(SS) Then            ' passed a RENDERSTYLESTRUCT
                If Picture Is Nothing Then
                    If DestType = saveTo_GDIplus Then
                        SS = SaveOptions
                        If Err Then
                            Err.Clear
                            Exit Function
                        End If
                        SS.reserved1 = ImageFormat
                        Set Destination = LoadBlankImage(SS, True)
                        SaveImage = (Destination.Handle <> 0&)
                    End If
                    Exit Function
                ElseIf Picture.Handle = 0& Then
                    Exit Function
                Else
                    MIS.Images = 1&                         ' create a MULTIIMAGESAVESTRUCT
                    ReDim MIS.Image(0 To 0)
                    MIS.Image(0).SS = SaveOptions
                    If Err Then
                        Err.Clear
                        Exit Function
                    End If
                    Set MIS.Image(0).Picture = Picture
                End If
            
            Else
                If LenB(SaveOptions) = LenB(MTP) Then
                    MTP = SaveOptions
                    If Err Then
                        Err.Clear
                        Exit Function
                    End If
                    If MTP.Pages < 1& Then Exit Function
                    ImageFormat = lvicSaveAsTIFF
                    MIS.Images = MTP.Pages
                    ReDim MIS.Image(LBound(MTP.TPS) To UBound(MTP.TPS))
                    For lPage = LBound(MTP.TPS) To UBound(MTP.TPS)
                        Set MIS.Image(lPage).Picture = MTP.TPS(lPage).Picture
                        MIS.Image(lPage).SS = MTP.TPS(lPage).SS
                    Next
                    Erase MTP.TPS()
                    
                ElseIf LenB(SaveOptions) = LenB(MIS) Then
                    MIS = SaveOptions
                    If Err Then
                        Err.Clear
                        Exit Function
                    End If
                    If MIS.Images < 1& Then Exit Function
                    If pvValidateMultiImgStruct(MIS, Picture, DestType, ImageFormat) = False Then Exit Function
                    Select Case ImageFormat
                    Case lvicSaveAsCursor, lvicSaveAsIcon, lvicSaveAsTIFF, lvicSaveAsGIF, lvicSaveAsPNG
                        ' expected
                    Case lvicSaveAsCurrentFormat
                        Select Case MIS.Image(LBound(MIS.Image)).Picture.ImageFormat
                        Case lvicSaveAsCursor, lvicSaveAsIcon, lvicSaveAsTIFF, lvicSaveAsGIF, lvicSaveAsPNG ' expected
                            ImageFormat = MIS.Image(LBound(MIS.Image)).Picture.ImageFormat
                        End Select
                    Case Else
                        ImageFormat = lvicSaveAsCurrentFormat
                    End Select
                    If ImageFormat = lvicSaveAsCurrentFormat Then ' if the desttype is file, use the extension else abort
                        If DestType = saveTo_File Then
                            lPage = InStrRev(CStr(Destination), ".")
                            If lPage Then
                                Select Case LCase$(Mid$(CStr(Destination), lPage, 3))
                                Case "tif": ImageFormat = lvicSaveAsTIFF
                                Case "gif": ImageFormat = lvicSaveAsGIF
                                Case "ico": ImageFormat = lvicSaveAsIcon
                                Case "cur": ImageFormat = lvicSaveAsCursor
                                Case "png": ImageFormat = lvicSaveAsPNG
                                Case "pam": ImageFormat = lvicSaveAsPAM
                                Case "pnm": ImageFormat = lvicSaveAsPNM
                                Case "pgm"
                                    With MIS.Image(LBound(MIS.Image)).SS
                                        If .RSS.Effects Is Nothing Then Set .RSS.Effects = New GDIpEffects
                                        If .RSS.Effects.GrayScale = lvicNoGrayScale Or .RSS.Effects.GrayScale = lvicSepia Then
                                            .RSS.Effects.GrayScale = lvicCCIR709
                                        End If
                                    End With
                                    ImageFormat = lvicSaveAsPNM
                                Case "pbm"
                                    With MIS.Image(LBound(MIS.Image)).SS
                                        .Palette_Handle = 0&
                                        .PaletteType = lvicPaletteDefault
                                        .ColorDepth = lvicConvert_BlackWhite
                                    End With
                                    ImageFormat = lvicSaveAsPNM
                                Case "ppm"
                                    MIS.Image(LBound(MIS.Image)).SS.ColorDepth = lvicConvert_TrueColor24bpp
                                    ImageFormat = lvicSaveAsPNM
                                End Select
                            End If
                        End If
                        If ImageFormat = lvicSaveAsCurrentFormat Then ImageFormat = lvicSaveAsBitmap
                    End If
                ElseIf LenB(SaveOptions) = LenB(SS.RSS) Then
                    If (Picture.ImageFormat = lvicPicTypeAVI And Picture.ImageCount > 1&) _
                        And (ImageFormat = lvicSaveAsPNG Or ImageFormat = lvicSaveAsGIF) And DestType <> saveTo_stdPicture Then
                        SS.RSS = SaveOptions
                        ' support converting AVI to png/gif unless creating a stdPicture object
                        If SS.RSS.Effects Is Nothing Then Set SS.RSS.Effects = New GDIpEffects
                        If pvBuildAVIconverter(Picture, MIS, ImageFormat, SS.RSS.Effects) = False Then MIS.Images = 0&
                    End If
                End If
            End If
            On Error GoTo 0
        End If
        lResult = 0&
    ElseIf Not Picture Is Nothing Then                          ' create an array for below loop
        If (Picture.ImageFormat = lvicPicTypeGIF And Picture.ImageCount > 1&) _
            And ImageFormat = lvicSaveAsPNG And DestType <> saveTo_stdPicture Then
            lResult = ImageFormat                               ' non-zero flag indicating alt processing
        ElseIf (Picture.ImageFormat = lvicPicTypeAVI And Picture.ImageCount > 1&) _
            And (ImageFormat = lvicSaveAsPNG Or ImageFormat = lvicSaveAsGIF) And DestType <> saveTo_stdPicture Then
            ' support converting AVI to png/gif unless creating a stdPicture object
            ' allow option to make top left pixel color transparent for each frame
            If pvBuildAVIconverter(Picture, MIS, ImageFormat, Nothing) = False Then MIS.Images = 0&
        Else
            MIS.Images = 1&
            ReDim MIS.Image(0 To 0)
            Set MIS.Image(0).Picture = Picture
            MIS.Image(0).SS.RSS.FillColorARGB = Color_RGBtoARGB(vbWindowBackground, 255&)
        End If
    End If
    If MIS.Images = 0& And lResult = 0& Then Exit Function      ' no valid picture objects passed

    On Error GoTo ExitRoutine
    
    ' if writing to file ensure we can open the file first; else no point in continuing
    If DestType = saveTo_File Then
        Call DeleteFileEx(CStr(Destination))
        hFileHandle = GetFileHandle(Destination, True)
        If hFileHandle = INVALID_HANDLE_VALUE Or hFileHandle = 0& Then Exit Function
    End If

    ' Note: the RENDERSTYLESTRUCT2.reserved1 member currently formatted as so
    ' &H00000001 = image uses transparency
    ' &H000000F0 = destination format (WMF use only for now)
    ' &H00000F00 = source image format
    ' &H10000000 = used by cColorReduction to perform color-less bit reduction
    ' &H20000000 = used by cColorReduction to return proper B&W bitmap if needed
    ' RENDERSTYLESTRUCT2.reserved2 currently holds pointer to the passed GDIpImage class
    
    ' ok, initial validation done, let's process the image
    If lResult = 0& Then
        
        lbOffset = LBound(MIS.Image)
        For lPage = 0& To MIS.Images - 1&
    
            If Not MIS.Image(lPage + lbOffset).Picture Is Nothing Then ' ensure picture exits & not null handle
                bProcess = (MIS.Image(lPage + lbOffset).Picture.Handle <> 0&)
            Else
                bProcess = True
            End If
            If bProcess Then
'/// the primary purpose of this battery of IF statements is to determine if the image requires any
'   pre-processing and if so, at what bit depth the image will be rendered to.
' If preprocessing is required, the only bit depths rendered to are 24bpp or above. The individual
'   SaveAs routines will drop the depth to 8bpp or less as needed.
                bPP = 0&
                With MIS.Image(lPage + lbOffset)
                    If .GroupNumber > 0& Then .Picture.ImageGroup = .GroupNumber
                    If .FrameNumber > 0& Then .Picture.ImageIndex = .FrameNumber
                    SS = .SS: SS.reserved2 = 0&: SS.reserved1 = 0&
                    If .SS.Width = 0& Then .SS.Width = .Picture.Width Else .SS.Width = Abs(.SS.Width)
                    If .SS.Height = 0& Then .SS.Height = .Picture.Height Else .SS.Height = Abs(.SS.Height)
                    If ImageFormat = lvicSaveAsCurrentFormat Then
                        ImageFormat = .Picture.ImageFormat
                        If .Picture.ImageFormat = lvicPicTypeAVI Then ImageFormat = -ImageFormat
                    End If
                    If .Picture.ImageFormat = lvicPicTypeAnimatedCursor Then
                        SS.reserved1 = lvicSaveAsCursor * &H100&
                    Else
                        SS.reserved1 = .Picture.ImageFormat * &H100&
                    End If
                    GdipGetImagePixelFormat .Picture.Handle, srcDepth
'/// call routine that validates structure parameters and makes initial call as to whether image is to be processed
                    SS.reserved1 = Abs(HasTransparency(.Picture.Handle)) Or SS.reserved1
                    bProcess = pvValidateSaveStructure(SS, .Picture, ImageFormat)
                End With
                lResult = (srcDepth And &HFF00&) \ &H100&
'/// saving to stdPicture? If so, enforce 24bpp Bitmap if not supported by VB
                If DestType = saveTo_stdPicture Then
                    Select Case ImageFormat
                        Case lvicSaveAsIcon, lvicSaveAsCursor, lvicSaveAsGIF, lvicSaveAsJPEG, lvicSaveAsMetafile, lvicSaveAsEMetafile, lvicSaveAsMetafile_NonPlaceable
                           ' handled later in logic flow
                        Case Else   ' bitmap or formats not supported by stdPicture
                            If (SS.reserved1 And 1&) Then bProcess = True: bPP = 24&
                            ImageFormat = lvicSaveAsBitmap
                    End Select
                End If
'/// identify formats where opaque color depth (i.e., 24bpp, 32bppRGB),should not affect transparency
'/// to force these to non-transparency, the Opions.FillColorUsed should be set to true
                Select Case ImageFormat
                    Case lvicSaveAsIcon, lvicSaveAsCursor, lvicSaveAs_HICON, lvicSaveAs_HCURSOR
                        bIsMasked = True
                        If lResult + (SS.reserved1 And 1&) = 32& Then bProcess = True
                End Select
'/// prevent premulitplied color depth from being applied to formats that don't support it
'/// if other than default or no reduction or paletted images, ensure we process the image
                Select Case SS.ColorDepth
                    Case lvicConvert_TrueColor32bpp_pARGB
                        If (ImageFormat = lvicSaveAsBitmap Or ImageFormat = lvicSaveAsTGA) Then
                            bPP = 32&
                        Else
                            SS.ColorDepth = lvicConvert_TrueColor32bpp_ARGB
                        End If
                        bProcess = True
                    Case lvicConvert_TrueColor24bpp: bProcess = True
                        If bIsMasked = False Then bPP = 24&
                    Case lvicConvert_TrueColor32bpp_ARGB: bProcess = True
                        If bPP = 0& Then bPP = 32&
                    Case lvicConvert_TrueColor32bpp_RGB: bProcess = True
                        If bPP = 0& Then bPP = 32&
                End Select
'/// some rendering options may change color depth
                With SS.RSS
                    If .FillColorUsed = True And (.FillColorARGB And &HFF000000) = &HFF000000 Then
                        bPP = 24&: bIsMasked = False                    ' force 24bpp minimum
                    ElseIf .Effects.GlobalTransparencyPct Then
                        If bPP <> 24& Then bPP = 32&: SS.reserved1 = SS.reserved1 Or 1& ' will contain transparency
                    ElseIf Not (.Angle = 0! Or .Angle = 180!) Then
                        If bPP = 0& Then
                            If Not (.Angle = 90! Or .Angle = 270!) Then
                                bPP = 32&: SS.reserved1 = SS.reserved1 Or 1& ' will contain transparency
                            ElseIf SS.RotatingCanGrowImage = False Then
                                bPP = 32&: SS.reserved1 = SS.reserved1 Or 1& ' will contain transparency
                            End If
                        End If
                    End If
                    If .Effects.TransparentColorUsed Then
                        If bPP = 0& Then bPP = 32& ' may not contain transparency if the color to be made transaparent doesn't exist
                    End If
                End With
'/// if bpp not determined yet, another check...
                If bPP = 0& Then
                    If (SS.reserved1 And 1&) Then bPP = 32& Else bPP = 24&  ' 32bpp if transparency exists else 24bpp
                End If
'/// handle some specific format cases
                If ImageFormat = lvicSaveAsJPEG Then            ' never supports transparency
                    If bPP = 32& Or lResult = 32& Or (SS.reserved1 And 1&) Then bPP = 24&: bProcess = True
                ElseIf ImageFormat = lvicSaveAs_HBITMAP Then    ' don't allow HBITMAP to include transparency
                    If bPP = 32& Or lResult = 32& Or (SS.reserved1 And 1&) Then bPP = 24&: bProcess = True
                ElseIf ImageFormat = lvicSaveAsPCX Then         ' PCX only allowed transparency if specifically requested
                    If bPP = 32& Or lResult = 32& Or (SS.reserved1 And 1&) Then
                        If MIS.Image(lPage + lbOffset).Picture.ImageFormat = lvicPicTypePCX Then
                            If (SS.ColorDepth And Not lvicApplyAlphaTolerance) < lvicConvert_BlackWhite Then SS.ColorDepth = lvicConvert_TrueColor32bpp_ARGB Or (SS.ColorDepth And lvicApplyAlphaTolerance)
                        End If
                        If (SS.ColorDepth And Not lvicApplyAlphaTolerance) < lvicConvert_TrueColor32bpp_ARGB Then bPP = 24&: bProcess = True
                    End If
                ElseIf ImageFormat = lvicSaveAsBitmap Then      ' no transparency for bitmaps beween 1 & 24 bpp
                    If (SS.ColorDepth And Not lvicApplyAlphaTolerance) < lvicConvert_TrueColor32bpp_RGB Then
                        If (SS.ColorDepth And Not lvicApplyAlphaTolerance) > lvicDefaultReduction Then bPP = 24: bProcess = True
                        ' side note: GDI+ supports 1,4,8 bpp palettes with transparency, but standard BMP format does not
                    End If
                    If (SS.reserved1 And 1&) Then bProcess = True
                ElseIf ImageFormat = lvicSaveAsPNM Then  ' never supports transparency
                    If bPP = 32& Or lResult = 32& Or (SS.reserved1 And 1&) Then bPP = 24&: bProcess = True
                End If
                
                If ImageFormat = -lvicPicTypeAVI Then bProcess = False ' saving modified AVI not supported
                
                If bProcess = False Then
                    wrkPic = MIS.Image(lPage + lbOffset).Picture.Handle
                    SS.reserved2 = ObjPtr(MIS.Image(lPage + lbOffset).Picture) ' informs saving routine that original image is provided
                        
                Else
                    ' create a working copy of the image at desired bit depth
                    If SS.ColorDepth And lvicApplyAlphaTolerance Then
                        SS.ColorDepth = SS.ColorDepth Xor lvicApplyAlphaTolerance
                        If (SS.reserved1 And 1&) Then
'/// if tolerance is to be applied, apply it before drawing the image with any effects
                            Set cPal = New cColorReduction
                            wrkPic = cPal.ApplyAlphaTolerance(MIS.Image(lPage + lbOffset).Picture.Handle, SS.AlphaTolerancePct, SS.Width, SS.Height, False)
                            Set cPal = Nothing
                            If wrkPic Then      ' convert GDI+ handle to a GDIpImage class
                                Set g_NewImageData = New cGDIpMultiImage
                                g_NewImageData.CacheSourceInfo Empty, wrkPic, lvicPicTypeBitmap, False, False
                                Set MIS.Image(lPage + lbOffset).Picture = New GDIpImage
                                Set g_NewImageData = Nothing
                                wrkPic = 0&
                            End If
                        End If
                    End If
                    
'/// last validation for bpp
                    If bPP = 24& And bIsMasked = False Then
                        ' 24bpp target for now unless ico/cur/wmf/emf/gif
                        bPP = lvicColor24bpp: SS.reserved1 = SS.reserved1 And Not 1&
                    ElseIf SS.ColorDepth = lvicConvert_TrueColor32bpp_RGB And bIsMasked = False Then
                        ' 32bpp RGB (no alpha) target for now
                        bPP = lvicColor32bpp: SS.reserved1 = SS.reserved1 And Not 1&
                    ElseIf SS.ColorDepth = lvicConvert_TrueColor32bpp_pARGB Then
                        ' pARGB (PCX/BMP only) and must have been requested
                        bPP = lvicColor32bppAlphaMultiplied
                    Else ' final default
                        bPP = lvicColor32bppAlpha
                    End If
'/// working copy creation
                
                 GetScaledImageSizes MIS.Image(lPage + lbOffset).SS.Width, MIS.Image(lPage + lbOffset).SS.Height, SS.Width, SS.Height, tSize.nWidth, tSize.nHeight, SS.RSS.Angle, False, False
                 If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, bPP, ByVal 0&, wrkPic) = 0& Then _
                     Call GdipGetImageGraphicsContext(wrkPic, hGraphics)         ' get a DC to draw on
                 If (wrkPic = 0& Or hGraphics = 0&) Then
                     If wrkPic Then GdipDisposeImage wrkPic: wrkPic = 0&
                     ' call this next line to allow SaveAsTIFF to destroy image/stream it is maintaining
                     If bReleaseMultTiff = True Then
                         lResult = SaveAsTIFF(vData, wrkPic, DestType, TIFF_MultiFrameStart)
                     Else
                         lResult = wrkPic
                     End If
                     Exit For
                 End If
                 
                SS.reserved1 = (SS.reserved1 And &HFFFF00FF) Or lvicPicTypeBitmap * &H100&
                With SS.RSS                                               ' fill background if applicable
                     If bPP = lvicColor24bpp Or bPP = lvicColor32bpp Then
                         If bIsMasked = False Then GdipGraphicsClear hGraphics, .FillColorARGB Or &HFF000000
                     Else
                         If .FillColorUsed Then GdipGraphicsClear hGraphics, .FillColorARGB
                     End If
                     If .FillBrushGDIplus_Handle Then                        ' if brush provided layer it
                         GdipFillRectangleI hGraphics, .FillBrushGDIplus_Handle, 0&, 0&, SS.Width, SS.Height
                     End If
                     Cx = tSize.nWidth: Cy = tSize.nHeight                   ' handle mirroring
                     If (SS.RSS.Mirrored And lvicMirrorHorizontal) Then Cx = -tSize.nWidth
                     If (SS.RSS.Mirrored And lvicMirrorVertical) Then Cy = -tSize.nHeight
'/// grayscale image if converting to black and white
                        If SS.ColorDepth = lvicConvert_BlackWhite Then
                            SS.RSS.Quality = lvicNearestNeighbor
                            If .Effects.AttributesHandle = 0& And .Effects.EffectsHandle(.EffectType) = 0& Then
                                Set .Effects = New GDIpEffects
                                .Effects.GrayScale = lvicCCIR709
                                MIS.Image(lPage + lbOffset).Picture.Render 0&, (SS.Width - tSize.nWidth) \ 2, (SS.Height - tSize.nHeight) \ 2, Cx, Cy, _
                                        , , , , .Angle, .Effects.AttributesHandle, hGraphics, 0&, .Quality
                                GdipDeleteGraphics hGraphics
                            Else
                                MIS.Image(lPage + lbOffset).Picture.Render 0&, (SS.Width - tSize.nWidth) \ 2, (SS.Height - tSize.nHeight) \ 2, Cx, Cy, _
                                        , , , , .Angle, .Effects.AttributesHandle, hGraphics, .Effects.EffectsHandle(.EffectType), .Quality
                                GdipDeleteGraphics hGraphics
                                lResult = 0&
                                If GdipCreateBitmapFromScan0(SS.Width, SS.Height, 0&, bPP, ByVal 0&, lResult) = 0& Then
                                    Set .Effects = New GDIpEffects
                                    .Effects.GrayScale = lvicCCIR709
                                    If GdipGetImageGraphicsContext(lResult, hGraphics) = 0& Then
                                        GdipDrawImageRectRectI hGraphics, wrkPic, 0&, 0&, SS.Width, SS.Height, 0&, 0&, SS.Width, SS.Height, UnitPixel, .Effects.AttributesHandle, 0&, 0&
                                        GdipDeleteGraphics hGraphics
                                        GdipDisposeImage wrkPic
                                        wrkPic = lResult
                                    Else
                                        GdipDisposeImage lResult: lResult = 0&
                                    End If
                                End If
                                If lResult = 0& Then GdipDisposeImage wrkPic: wrkPic = 0&
                            End If
                        Else
'/// not black and white, render image to working copy
                            MIS.Image(lPage + lbOffset).Picture.Render 0&, (SS.Width - tSize.nWidth) \ 2, (SS.Height - tSize.nHeight) \ 2, Cx, Cy, _
                                    , , , , .Angle, .Effects.AttributesHandle, hGraphics, .Effects.EffectsHandle(.EffectType), .Quality
                            GdipDeleteGraphics hGraphics
                        End If
                    End With
                    If bPP = lvicColor32bppAlpha Or bPP = lvicColor32bppAlphaMultiplied Then
                        If (SS.reserved1 And 1&) = 0& Then SS.reserved1 = SS.reserved1 Or Abs(HasTransparency(wrkPic))
                    End If
                    If (SS.ColorDepth = lvicNoColorReduction) And ((SS.reserved1 And 1&) = 0&) Then
                        If (srcDepth And &HFF00&) \ &H100& < bPP And DestType <> saveTo_GDIhandle Then
                            SS.PaletteType = lvicPaletteAdaptive
                            Select Case (srcDepth And &HFF00&) \ &H100&
                                Case 8&: SS.ColorDepth = lvicConvert_256Colors
                                Case 4&: SS.ColorDepth = lvicConvert_16Colors
                                Case Else: SS.ColorDepth = lvicDefaultReduction
                            End Select
                        End If
                    End If
                    
                End If
'/// send off to other routines to create image format
                If wrkPic Then
                    If DestType = saveTo_File Then
                        vData = hFileHandle
                    ElseIf DestType = saveTo_Array Then
                        vData = imgData()
                    ElseIf DestType <> saveTo_GDIplus Then
                        Set vData = Destination
                    End If
                    Select Case ImageFormat
                        Case lvicSaveAsBitmap, lvicSaveAs_HBITMAP
                            SS.reserved1 = SS.reserved1 Or ImageFormat * &H10&
                            lResult = SaveAsBMP(vData, wrkPic, DestType, SS)
                        Case lvicSaveAsJPEG
                            lResult = pvSaveAsJPEG(vData, wrkPic, DestType, SS)
                        Case lvicSaveAsPNG
                            If MIS.Images < 2& Then ' not animated PNG
                                lResult = SaveAsPNG(vData, wrkPic, DestType, SS)
                            Else                        ' animated
                                lResult = SaveAsPNG(vData, wrkPic, DestType, SS, MIS, lPage)
                                If lResult = 0& Then Exit For
                            End If
                        Case lvicSaveAsGIF
                            If MIS.Images < 2& Then    ' not an animated GIF
                                lResult = pvSaveAsGIF(vData, wrkPic, DestType, SS)
                            Else                        ' animated
                                lResult = pvSaveAsGIF(vData, wrkPic, DestType, SS, MIS, lPage)
                                If lResult = 0& Then Exit For
                            End If
                        Case lvicSaveAsEMetafile, lvicSaveAsMetafile, lvicSaveAsMetafile_NonPlaceable
                            SS.reserved1 = SS.reserved1 Or ImageFormat * &H10
                            lResult = pvSaveAsMetafile(vData, wrkPic, DestType, SS)
                        Case lvicSaveAsPCX
                            lResult = pvSaveAsPCX(vData, wrkPic, DestType, SS)
                        Case lvicSaveAsTGA
                            lResult = pvSaveAsTGA(vData, wrkPic, DestType, SS)
                        Case lvicSaveAsIcon, lvicSaveAsCursor, lvicSaveAs_HICON, lvicSaveAs_HCURSOR
                            SS.reserved1 = SS.reserved1 Or ImageFormat * &H10&
                            If MIS.Images < 2& And MIS.Image(lPage + lbOffset).IconEmbeddedAsPNG = False Then
                                lResult = pvSaveAsICO(vData, wrkPic, DestType, SS)
                            Else
                                lResult = pvSaveAsICO(vData, wrkPic, DestType, SS, MIS, lPage)
                                If lResult = 0& Then Exit For
                            End If
                        Case lvicSaveAsTIFF
                            If MIS.Images < 2& Then
                                lResult = SaveAsTIFF(vData, wrkPic, DestType, TIFF_SingleFrame, SS)
                            Else
                                If bReleaseMultTiff Then lParam = TIFF_MultiFrameAdd Else lParam = TIFF_MultiFrameStart
                                lResult = SaveAsTIFF(vData, wrkPic, DestType, lParam, SS)
                                If lResult = 0& Then Exit For
                                bReleaseMultTiff = True
                            End If
                        Case lvicSaveAsPNM, lvicSaveAsPAM
                            lResult = pvSaveAsPNM(vData, wrkPic, DestType, SS, (ImageFormat = lvicSaveAsPAM))
                        Case Is < 0&
                            ' AVI special handling
                            lResult = SaveAsAVI(vData, wrkPic, DestType, SS)
                    End Select
'/// clean up
                    If wrkPic <> MIS.Image(lPage + lbOffset).Picture.Handle Then
                        If wrkPic Then GdipDisposeImage wrkPic
                    End If
                End If
            End If
        Next
    Else                ' converting animated GIF to APNG
        MIS.Images = 0&
        SS.reserved2 = ObjPtr(Picture)
        If DestType = saveTo_File Then
            vData = hFileHandle
        ElseIf DestType = saveTo_Array Then
            vData = imgData
        ElseIf DestType <> saveTo_GDIplus Then
            Set vData = Destination
        End If
        lResult = SaveAsPNG(vData, Picture.Handle, DestType, SS, MIS, 0&)
        
    End If  ' end of APNG check
    
'/// handle return result
    If lResult <> 0& Then
        If bReleaseMultTiff = True Then
            If DestType = saveTo_File Then
                SaveImage = SaveAsTIFF(hFileHandle, 0&, DestType, TIFF_MultiFrameEnd)
            ElseIf DestType = saveTo_GDIplus Then
                lResult = SaveAsTIFF(vData, 0&, DestType, TIFF_MultiFrameEnd)
            Else
                SaveImage = SaveAsTIFF(Destination, 0&, DestType, TIFF_MultiFrameEnd)
            End If
        ElseIf DestType = saveTo_Array Then
            MoveArrayToVariant vData, imgData(), False
            MoveArrayToVariant Destination, imgData(), True
            SaveImage = True
        End If
        If DestType = saveTo_GDIplus Then               ' create GDIpImage if needed
            If lResult = saveTo_GDIplus Then            ' indicates SaveAs routine created GDIpImage
                Set Destination = vData
            Else
                Set g_NewImageData = New cGDIpMultiImage        ' create GDIpImage class
                If ImageFormat = lvicSaveAsMetafile_NonPlaceable Then ImageFormat = lvicSaveAsMetafile
                g_NewImageData.CacheSourceInfo vData, lResult, ImageFormat, True, False
                Set Destination = New GDIpImage
                Set g_NewImageData = Nothing
            End If
            SaveImage = (Destination.Handle <> 0&)
        ElseIf bReleaseMultTiff = False Then
            If DestType = saveTo_stdPicture Then
                Set Destination = vData
            ElseIf DestType = saveTo_GDIhandle Then
                Destination = vData
            End If
            SaveImage = True
        End If
    End If
    
ExitRoutine:
    If Err Then
        If lPage < MIS.Images - 1& Then
            If wrkPic <> MIS.Image(lPage + lbOffset).Picture.Handle Then
                If wrkPic Then GdipDisposeImage wrkPic
            End If
        End If
        SaveImage = False
        Err.Clear
    End If
    If hFileHandle Then                                                     ' close any open file handles
        CloseHandle hFileHandle
        If SaveImage = False Then DeleteFileEx CStr(Destination)
    End If

End Function

Public Function OverlayImage(picData As PICTUREMERGESTRUCT) As GDIpImage

    ' Routine merges/combines/overlays 2 or more images into one new image

    On Error GoTo ExitRoutine
    If picData.Pictures = 0& Then Exit Function
    If picData.CanvasHeight < 1& Or picData.CanvasWidth < 1& Then Exit Function
    If picData.Pictures <> UBound(picData.MIS) - LBound(picData.MIS) + 1& Then Exit Function
    
    Dim lPic As Long, hBrush As Long, bClipped As Boolean
    Dim lOverlay As Long, hGraphics As Long
    Dim Cx As Long, Cy As Long, RSS As RENDERSTYLESTRUCT2
    
    For lPic = LBound(picData.MIS) To UBound(picData.MIS)
        If Not picData.MIS(lPic).Picture Is Nothing Then
            If picData.MIS(lPic).Picture.Handle Then Exit For
        End If
    Next
    If lPic > UBound(picData.MIS) Then Exit Function          ' no valid images passed
    
    If GdipCreateBitmapFromScan0(picData.CanvasWidth, picData.CanvasHeight, 0&, lvicColor32bppAlpha, ByVal 0&, lOverlay) Then Exit Function
    If GdipGetImageGraphicsContext(lOverlay, hGraphics) Then
        GdipDisposeImage lOverlay
        Exit Function
    End If
    
    For lPic = LBound(picData.MIS) To UBound(picData.MIS)
        With picData.MIS(lPic)
            If Not .Picture Is Nothing Then
                If .Picture.Handle Then
                    RSS = .RSS
                    ValidateRenderStyle2 RSS, .Picture
                    If .Width = 0& Then
                        Cx = .Picture.Width
                    ElseIf .Width < 0& Then
                        Cx = -.Width: RSS.Mirrored = RSS.Mirrored Xor lvicMirrorHorizontal
                    Else
                        Cx = .Width
                    End If
                    
                    If .Height = 0& Then
                        Cy = .Picture.Height
                    ElseIf .Height < 0& Then
                        Cy = -.Height: RSS.Mirrored = RSS.Mirrored Xor lvicMirrorVertical
                    Else
                        Cy = .Height
                    End If
                    
                    If RSS.FillColorUsed Then
                        If GdipCreateSolidFill(RSS.FillColorARGB, hBrush) = 0& Then
                            GdipFillRectangleI hGraphics, hBrush, .Left, .TOp, Cx, Cy
                            GdipDeleteBrush hBrush
                        End If
                    End If
                    If RSS.FillBrushGDIplus_Handle Then
                        GdipFillRectangleI hGraphics, RSS.FillBrushGDIplus_Handle, .Left, .TOp, Cx, Cy
                    End If
                    If Not (RSS.Angle = 0! Or RSS.Angle = 180!) Then _
                        bClipped = (GdipSetClipRectI(hGraphics, .Left, .TOp, Cx, Cy, 0&) = 0&)
                    If (RSS.Mirrored And lvicMirrorHorizontal) Then Cx = -Cx
                    If (RSS.Mirrored And lvicMirrorVertical) Then Cy = -Cy
                    .Picture.Render 0&, .Left, .TOp, Cx, Cy, , , , , RSS.Angle, RSS.Effects.AttributesHandle, hGraphics, RSS.Effects.EffectsHandle(RSS.EffectType), RSS.Quality
                    If bClipped Then GdipResetClip hGraphics: bClipped = False
                End If
            End If
        End With
    Next
    GdipDeleteGraphics hGraphics: hGraphics = 0&
    Set g_NewImageData = New cGDIpMultiImage
    g_NewImageData.CacheSourceInfo Empty, lOverlay, lvicPicTypeBitmap, True, False
    Set OverlayImage = New GDIpImage
    lOverlay = 0&

ExitRoutine:
    If hGraphics Then GdipDeleteGraphics hGraphics
    If lOverlay Then GdipDisposeImage lOverlay
End Function

Private Function pvValidateSaveStructure(SS As SAVESTRUCT, img As GDIpImage, ImageType As SaveAsFormatEnum) As Boolean

    ' routine ensures all members of the SAVESTRUCT are valid

    Dim bResult As Boolean, Cx As Long, Cy As Long
    Dim bScaleUp As Boolean, maxW As Long, maxH As Long
    
    bResult = ValidateRenderStyle2(SS.RSS, img)
    With SS
        If .Width = 0& Then
            .Width = img.Width
        ElseIf .Width < 0& Then
            .Width = -.Width
            .RSS.Mirrored = .RSS.Mirrored Xor lvicMirrorHorizontal
        End If
        If .Height = 0& Then
            .Height = img.Height
        ElseIf .Height < 0& Then
            .Height = -.Height
            .RSS.Mirrored = .RSS.Mirrored Xor lvicMirrorVertical
        End If
        Select Case ImageType                       ' ensure maximum size restrictions
            Case lvicSaveAsMetafile_NonPlaceable    ' use max positive integer values as twips
                maxW = &H7FFF& \ Screen.TwipsPerPixelX: maxH = &H7FFF& \ Screen.TwipsPerPixelY
            Case lvicSaveAsIcon, lvicSaveAsCursor, lvicSaveAs_HICON, lvicSaveAs_HCURSOR
                maxW = 256&: maxH = maxW            ' limited to max of 256x256
            Case Else
                maxW = &H7FFF&: maxH = maxW         ' use max positive integer values
        End Select                                  ' anything larger will most likely fail anyway
        GetScaledCanvasSize .Width, .Height, Cx, Cy, .RSS.Angle ' bounds to display image as requested
        If Cx > maxW Or Cy > maxH Then              ' exceeds max bounds?
            GetScaledImageSizes Cx, Cy, maxW, maxH, Cx, Cy, .RSS.Angle, False, False
            .Width = Cx: .Height = Cy               ' set to maximum allowable; scaled
        ElseIf .RotatingCanGrowImage Then
            If Cx > .Width Or Cy > .Height Then     ' rotating? Up dimensions if requested
                If Not (.RSS.Angle = 0! Or .RSS.Angle = 180!) Then .Width = Cx: .Height = Cy
            End If
        End If
        ' for all other scenarios, .Width & .Height left unchanged
        bResult = bResult Or (.Width <> img.Width) Or (.Height <> img.Height)
        .PaletteType = (.PaletteType And &H3&)
        If .PaletteType > lvicPaletteUserDefined Then .PaletteType = lvicPaletteDefault
        If .AlphaTolerancePct > 99& Then
            .AlphaTolerancePct = 0&
        ElseIf .AlphaTolerancePct < 1& Then
            .AlphaTolerancePct = 254&
        Else
            .AlphaTolerancePct = 255& - (.AlphaTolerancePct * 255&) \ 100&
        End If
        .ColorDepth = (.ColorDepth And &H1F)
        If (.reserved1 And 1&) = 0& Then
            .ColorDepth = (.ColorDepth And Not lvicApplyAlphaTolerance)
        ElseIf (.ColorDepth And lvicApplyAlphaTolerance) Then
            bResult = True
        End If
        Select Case (.ColorDepth And Not lvicApplyAlphaTolerance)
            Case lvicConvert_BlackWhite: bResult = True
            Case Is < lvicNoColorReduction: .ColorDepth = (.ColorDepth And lvicApplyAlphaTolerance)
            Case Is > lvicConvert_TrueColor32bpp_pARGB = (.ColorDepth And lvicApplyAlphaTolerance)
        End Select
        If .ExtractCurrentFrameOnly = True And img.ImageCount < 2& Then .ExtractCurrentFrameOnly = False
    End With
    pvValidateSaveStructure = bResult
    
End Function

Public Function ValidateRenderStyle2(RS As RENDERSTYLESTRUCT2, img As GDIpImage) As Boolean

    ' support function for SaveImage, PaintPictureGDIplus, OverlayImage
    ' sets default values and returns whether special processing is required or not
    
    Dim bResult As Boolean
    With RS
        ' validate passed settings, ignoring any that are out of range
        If .Effects Is Nothing Then
            Set .Effects = New GDIpEffects
        ElseIf .Effects.AttributesHandle Then
            bResult = True                                              ' uses image attributes; render
        ElseIf .Effects.EffectsHandle(.EffectType) Then
            bResult = True                                              ' uses v1.1 effects; render
        End If                                                          ' resizing? If so, render
        .Angle = (Int(.Angle) Mod 360!) + (.Angle - Int(.Angle))
        bResult = bResult Or (.Angle <> 0!)
        .Mirrored = (.Mirrored And lvicMirrorBoth): bResult = bResult Or (.Mirrored > lvicMirrorNone)
        .Quality = (.Quality And &H7)
        If .Quality > lvicHighQualityBicubic Then .Quality = lvicAutoInterpolate
        bResult = bResult Or (.Quality > lvicAutoInterpolate)
        bResult = bResult Or (.FillBrushGDIplus_Handle <> 0&) Or .FillColorUsed
    End With
    ValidateRenderStyle2 = bResult Or img.Segmented

End Function


Public Function HasTransparency(Handle As Long) As Boolean

    ' routine returns whether the raw image contains any levels of transparency

    Dim lValue As Long, bDummy() As Long
    Dim palColors As ColorPalette
    
    ' GDI+ color format doesn't give me everything I need.
    ' In many routines, it is important to know if transparency is used or not
    If Handle = 0& Then Exit Function
    GdipGetImagePixelFormat Handle, lValue
    If (lValue And &HFF00&) \ &H100& <= 8& Then           ' paletted images
        If GdipGetImagePaletteSize(Handle, lValue) = 0& Then
            If GdipGetImagePalette(Handle, palColors, lValue) = 0& Then
                For lValue = 1 To palColors.Count
                    If (palColors.Entries(lValue) And &HFF000000) <> &HFF000000 Then
                        HasTransparency = True
                        Exit For
                    End If
                Next
            End If
        End If
    Else                                        ' non-paletted images
        Select Case lValue
        Case lvicColor24bpp, lvicColor16bpp555, lvicColor16bpp565
            ' these should never have transparency
        Case Else
            HasTransparency = (ValidateAlphaChannel(bDummy(), Handle) <> lvicColor32bpp)
        End Select
    End If
    
End Function

Public Function TileImage(Picture As GDIpImage, Destination As Variant, _
                                ByVal destX As Long, ByVal destY As Long, _
                                ByVal destWidth As Long, ByVal destHeight As Long, _
                                ByVal TileStyle As TileOrderEnum, _
                                ByVal tileWidth As Long, ByVal tileHeight As Long, _
                                ByVal tileGapWidth As Long, _
                                ByVal tileGapHeight As Long, _
                                ByVal ScaledTiles As Boolean, _
                                ByVal AlternateRows As Boolean, _
                                RSS As RENDERSTYLESTRUCT2) As Boolean

    ' Routine tiles an image to a target DC, using various tiling options

    Const OBJ_DC As Long = 3&
    Const OBJ_MEMDC As Long = 10&
    
    If tileWidth < 1& Or tileHeight < 1& Then Exit Function
    If destWidth < 1& Or destHeight < 1& Then Exit Function
    If Picture.Handle = 0& Then Exit Function
    
    If IsObject(Destination) Then
        If Not TypeOf Destination Is GDIpImage Then Exit Function
    ElseIf Not VarType(Destination) = vbLong Then
        Exit Function
    Else
        Select Case GetObjectType(CLng(Destination))
        Case OBJ_DC, OBJ_MEMDC
        Case Else
            Exit Function
        End Select
    End If
    If tileGapHeight < 0& Then tileGapHeight = 0&
    If tileGapWidth < 0& Then tileGapWidth = 0&
    
    Dim tSA As SafeArray, bData() As Long
    Dim X As Long, Y As Long, tileImg As GDIpImage
    Dim tiledCx As Long, tiledCy As Long
    Dim nrRowsToCopy As Long, nrColsToCopy As Long
    Dim tileX As Long, tileY As Long, hBrush As Long
    Dim lHandle As Long, hGraphics As Long
    Dim tBMP As BitmapData, sizeI As RECTI
    Dim Cx As Long, Cy As Long, mirrorX As Long, mirrorY As Long
    
    If GdipCreateBitmapFromScan0(destWidth, destHeight, 0&, lvicColor32bppAlpha, ByVal 0&, lHandle) Then Exit Function
    If GdipGetImageGraphicsContext(lHandle, hGraphics) Then
        GdipDisposeImage lHandle
        Exit Function
    End If
    
    ValidateRenderStyle2 RSS, Picture
    tiledCx = tileWidth + tileGapWidth ' tracks how much we tiled horizontally
    tiledCy = tileHeight + tileGapHeight ' tracks how much we tiled vertically
    If RSS.FillColorUsed Then
        If GdipCreateSolidFill(RSS.FillColorARGB, hBrush) = 0& Then
            GdipFillRectangleI hGraphics, hBrush, 0&, 0&, tiledCx, tiledCy
            If TileStyle > lvicTile_NoFlip Then GdipFillRectangleI hGraphics, hBrush, tiledCx, 0&, tiledCx, tiledCy
            If AlternateRows Then GdipFillRectangleI hGraphics, hBrush, -tiledCx \ 2, tiledCy, tiledCx, tiledCy
            GdipDeleteBrush hBrush
        End If
    End If
    If RSS.FillBrushGDIplus_Handle Then
        GdipFillRectangleI hGraphics, RSS.FillBrushGDIplus_Handle, 0&, 0&, tileWidth, tileHeight
        If TileStyle > lvicTile_NoFlip Then GdipFillRectangleI hGraphics, RSS.FillBrushGDIplus_Handle, tiledCx, 0&, tileWidth, tileHeight
        If AlternateRows Then GdipFillRectangleI hGraphics, RSS.FillBrushGDIplus_Handle, -tiledCx \ 2, tiledCy, tileWidth, tileHeight
    End If
    
    If ScaledTiles = True Or Not (RSS.Angle = 0! Or RSS.Angle = 180!) Then
        GetScaledImageSizes Picture.Width, Picture.Height, tileWidth, tileHeight, Cx, Cy, RSS.Angle, False, False
        X = (tileWidth - Cx) \ 2
        Y = (tileHeight - Cy) \ 2
    Else
        Cx = tileWidth: Cy = tileHeight
    End If
    Picture.Render 0&, X, Y, Cx, Cy, , , , , , RSS.Effects.AttributesHandle, hGraphics, RSS.Effects.EffectsHandle(RSS.EffectType), RSS.Quality
    If TileStyle > lvicTile_NoFlip Then
        Select Case TileStyle
        Case lvicTile_FlipX
            tileX = -Cx: tileY = Cy
        Case lvicTile_FlipY
            tileX = Cx: tileY = -Cy
        Case lvicTile_FlipXY
            tileX = -Cx: tileY = -Cy
        End Select
        Picture.Render 0&, X + tiledCx, Y, tileX, tileY, , , , , , RSS.Effects.AttributesHandle, hGraphics, RSS.Effects.EffectsHandle(RSS.EffectType), RSS.Quality
        If AlternateRows Then
            Picture.Render 0&, X - tiledCx \ 2, Y + tiledCy, tileX, tileY, , , , , , RSS.Effects.AttributesHandle, hGraphics, RSS.Effects.EffectsHandle(RSS.EffectType), RSS.Quality
        End If
        tiledCx = tiledCx + tiledCx
    ElseIf AlternateRows Then
        Picture.Render 0&, X - tiledCx \ 2, tiledCy, Cx, Cy, , , , , , RSS.Effects.AttributesHandle, hGraphics, RSS.Effects.EffectsHandle(RSS.EffectType), RSS.Quality
    End If
    
    GdipDeleteGraphics hGraphics
    
    sizeI.nHeight = destHeight: sizeI.nWidth = destWidth
    If GdipBitmapLockBits(lHandle, sizeI, ImageLockModeWrite, lvicColor32bppAlpha, tBMP) Then
        GdipDisposeImage lHandle
        Exit Function
    End If
    With tSA
        .cbElements = 4
        .cDims = 2
        .pvData = tBMP.Scan0Ptr
        .rgSABound(0).cElements = destHeight
        .rgSABound(1).cElements = destWidth
    End With
    CopyMemory ByVal VarPtrArray(bData), VarPtr(tSA), 4&
    
    If tiledCy < destHeight Then
        nrRowsToCopy = tiledCy
    Else
        nrRowsToCopy = destHeight
    End If
    ' tile the 1st row completely, incrementing the number of pixels rendered on each pass
    Do While destWidth > tiledCx
        If destWidth > tiledCx + tiledCx Then      ' validate width remaining
            nrColsToCopy = tiledCx * 4&
        Else
            nrColsToCopy = (destWidth - tiledCx) * 4&
            If nrColsToCopy < 4 Then Exit Do
        End If
        For Y = 0& To nrRowsToCopy - 1
            CopyMemory bData(tiledCx, Y), bData(0, Y), nrColsToCopy
        Next
        tiledCx = tiledCx + tiledCx            ' increment nr cols tiled
    Loop
    
    If AlternateRows Then
        ' now to handle staggered rows. We will render the 2nd row's 1st tile (1/2 tile really)
        If destHeight > tiledCy + tileHeight Then
            nrRowsToCopy = tileHeight
        Else
            nrRowsToCopy = destHeight - tiledCy
        End If
        If nrRowsToCopy > 0 Then
            tiledCx = tileWidth + tileGapWidth      ' total tile width including gap
            tileX = tiledCx \ 2
            Do While destWidth > tileX
                If tileX + tiledCx < destWidth Then
                    nrColsToCopy = tiledCx * 4&  ' how many pixels can we copy?
                Else
                    nrColsToCopy = (destWidth - tileX) * 4&
                    If nrColsToCopy < 4& Then Exit Do
                End If
                tileY = 0&                              ' set vertical source point
                For Y = tiledCy To tiledCy + nrRowsToCopy - 1
                    CopyMemory bData(tileX, Y), bData(0, tileY), nrColsToCopy
                    tileY = tileY + 1
                Next
                tileX = tileX + nrColsToCopy \ 4&
            Loop
            tiledCy = tiledCy + nrRowsToCopy + tileGapHeight
        End If
    End If
    
    ' The above steps were just to set up a complete row (or 2 rows if staggered).
    ' this is where the speed is realized and kind of scary.
    ' We will be copying entire rows (blocks), in an incrementing manner, each pass
    ' First pass: 1 or 2 rows copied, then double that, then quadruple that, etc.
    ' Just one miscalculation will crash
    Do While destHeight > tiledCy
        If destHeight > tiledCy + tiledCy Then  ' validate height remaining
            nrRowsToCopy = tiledCy
        Else
            nrRowsToCopy = destHeight - tiledCy
        End If
        If nrRowsToCopy < 1 Then Exit Do
        nrColsToCopy = (destWidth * 4) * nrRowsToCopy ' total number of bytes to copy (can be very large)
        CopyMemory bData(0, tiledCy), bData(0, 0), nrColsToCopy
        tiledCy = tiledCy + nrRowsToCopy              ' increment number of rows tiled
    Loop
    CopyMemory ByVal VarPtrArray(bData), 0&, 4&
    GdipBitmapUnlockBits lHandle, tBMP
    
    If IsObject(Destination) Then
        Set g_NewImageData = New cGDIpMultiImage
        g_NewImageData.CacheSourceInfo Empty, lHandle, lvicPicTypeBitmap, True, False
        Set Destination = New GDIpImage
    Else
        If GdipCreateFromHDC(CLng(Destination), hGraphics) Then
            GdipDisposeImage lHandle
            Exit Function
        End If
        GdipDrawImageRectRectI hGraphics, lHandle, destX, destY, destWidth, destHeight, 0&, 0&, destWidth, destHeight, UnitPixel, 0&, 0&, 0&
        GdipDeleteGraphics hGraphics
        GdipDisposeImage lHandle
    End If
    
    TileImage = True

End Function

Public Function FindColor(ByRef PaletteItems() As Long, ByVal Color As Long, ByVal Count As Long, ByRef isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer. Used by cFunctionsBMP & cFunctionsTGA
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    
    ' [in] PaletteItems(). Long Array to search within. Array must be 1-bound
    ' [in] Color. A value to search for. Order is always ascending
    ' [in] Count. Number of items in PaletteItems() to compare against
    ' [out] isNew. If Color not found, isNew is True else False
    ' [out] Return value: The Index where Color was found or where the new Color should be inserted

    Dim UB As Long, lb As Long
    Dim newIndex As Long
    
    If Count = 0& Then
        FindColor = 1&
        isNew = True
        Exit Function
    End If
    
    UB = Count
    lb = 1&
    
    Do Until lb > UB
        newIndex = lb + ((UB - lb) \ 2&)
        If PaletteItems(newIndex) = Color Then
            Exit Do
        ElseIf PaletteItems(newIndex) > Color Then ' new color is lower in sort order
            UB = newIndex - 1&
        Else ' new color is higher in sort order
            lb = newIndex + 1&
        End If
    Loop

    If lb > UB Then  ' color was not found
            
        If Color > PaletteItems(newIndex) Then newIndex = newIndex + 1&
        isNew = True
        
    Else
        isNew = False
    End If
    
    FindColor = newIndex

End Function

Public Function ValidateAlphaChannel(inStream() As Long, Optional Handle As Long) As Long

    ' Purpose: Determine if alpha channel used by passed stream & if so, how it is used
    
    Dim X As Long, Y As Long, lPrevColor As Long
    Dim bAlpha As Boolean, bOpaque As Boolean, bPreMultiplied As Boolean
    Dim tSizeI As RECTI, tSize As RECTF, bValue As Byte
    Dim tSA As SafeArray, tBMP As BitmapData
    
    If Handle Then
        GdipGetImageBounds Handle, tSize, UnitPixel
        If (tSize.nHeight < 1! Or tSize.nWidth < 1!) Then Exit Function
        tSizeI.nHeight = tSize.nHeight: tSizeI.nWidth = tSize.nWidth
        If GdipBitmapLockBits(Handle, tSizeI, ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
        With tSA
            .cbElements = 4
            .cDims = 2
            .pvData = tBMP.Scan0Ptr
            If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
            .rgSABound(0).cElements = tBMP.Height
            .rgSABound(1).cElements = tBMP.Width
        End With
        Erase inStream()
        CopyMemory ByVal VarPtrArray(inStream), VarPtr(tSA), 4&
        On Error GoTo ExitRoutine
    Else
        tSizeI.nHeight = UBound(inStream, 2) + 1&: tSizeI.nWidth = (UBound(inStream, 1) + 1&)
    End If

    ' see if the 32bpp is premultiplied or not and if it is alpha or not
    lPrevColor = inStream(X, Y) Xor 1&
    For Y = 0& To tSizeI.nHeight - 1&
        For X = 0& To tSizeI.nWidth - 1&
            If inStream(X, Y) <> lPrevColor Then
                lPrevColor = inStream(X, Y)
                Select Case (lPrevColor And &HFF000000)
                    Case &HFF000000                             ' fully opaque
                        bOpaque = True
                    Case 0&                                     ' fully transparent
                        bAlpha = True                           ' can't abort; all alpha values may be zero
                    Case Else                                   ' if any of following are triggered: not pre-multiplied
                        bValue = (lPrevColor And &H7F000000) \ &H1000000
                        If lPrevColor < 0& Then bValue = bValue Or &H80
                        If bValue < (lPrevColor And &HFF) Then
                            ValidateAlphaChannel = lvicColor32bppAlpha
                            Y = tSizeI.nHeight: Exit For
                        ElseIf bValue < (lPrevColor And &HFF0000) \ &H10000 Then
                            ValidateAlphaChannel = lvicColor32bppAlpha
                            Y = tSizeI.nHeight: Exit For
                        ElseIf bValue < (lPrevColor And &HFF00&) \ &H100& Then
                            ValidateAlphaChannel = lvicColor32bppAlpha
                            Y = tSizeI.nHeight: Exit For
                        Else
                            bPreMultiplied = True               ' may be or not; don't know yet
                        End If
                End Select
            End If
        Next
    Next
    
    If ValidateAlphaChannel = 0& Then                           ' undetermined
        If bPreMultiplied Then                                  ' either pre-multiplied or not
            ValidateAlphaChannel = lvicColor32bppAlphaMultiplied '   but better to assume premultiplied than not
        ElseIf bAlpha = True And bOpaque = True Then            ' simple transparency
            ValidateAlphaChannel = lvicColor32bppAlpha
        Else                                                    ' all alphas are zero or 255; assume alpha not used
            ValidateAlphaChannel = lvicColor32bpp
        End If
    End If
    
ExitRoutine:
    If tSA.pvData Then
        CopyMemory ByVal VarPtrArray(inStream), 0&, 4&
        GdipBitmapUnlockBits Handle, tBMP
    End If
End Function

Private Function pvConvertNonPlaceableWMFtoWMF(inStream() As Byte, FileName As String) As Long

    ' return value is one of these
    ' lvicPicTypeNone = WMF placeable or EMF image (no modifications)
    ' lvicPicTypeUnknown = Not 100% sure; don't load it if we aren't sure how to manipulate it
    ' GDI+ handle = WMF non_placeable, converted to WMF placeable at screen size scaled to 256x256

    Dim aPtr As Long, lValue As Long, wmfData() As Byte
    Dim Cx As Long, Cy As Long, IIStream As IUnknown
    Dim hHandle As Long
    
    Const magicWMF As Long = &H9AC6CDD7
    Const magicEMF As Long = &H464D4520

    If FileName <> vbNullString Then
        hHandle = GetFileHandle(FileName, False)                ' open the file
        If hHandle = INVALID_HANDLE_VALUE Then
            pvConvertNonPlaceableWMFtoWMF = lvicPicTypeUnknown
            Exit Function
        End If
        ReDim inStream(0 To 44)
        SetFilePointer hHandle, 0&, 0&, 0&
        ReadFile hHandle, inStream(0), 44&, lValue, ByVal 0&
    End If
    
    CopyMemory lValue, inStream(0), 4&
    If lValue = magicWMF Then
        If hHandle Then CloseHandle hHandle
        Erase inStream()
        Exit Function         ' placeable; nothing to do
    End If
    CopyMemory lValue, inStream(40), 4&
    If lValue = magicEMF Then
        If hHandle Then CloseHandle hHandle
        Erase inStream()
        Exit Function         ' EMF, nothing to do
    End If
    
    ' so this is a non-placeable WMF either in WMF or CLP format; validate
    ' if the WMF is not CLP formatted 8 or 16 byte header preceeding WMF standard header, then
    '   the 2nd byte will be 9
    ' For CLP_16bit, value will be between the width (non-zero)
    ' For CLP_32bit the 2nd byte will be zero & 3rd byte will be 1-8
    ' ACKNOWLEDGEMENT: http://wvware.sourceforge.net/caolan/ora-wmf.html#MICMETA-DMYID.2
    
    lValue = 0&: CopyMemory lValue, inStream(2), 2&
    If lValue = 9 Then
        aPtr = 0&
    Else
        If lValue = 0& Then     ' should be 32bpp CLP format, validate
            CopyMemory lValue, inStream(18), 2&
            If lValue = 9& Then aPtr = 16&
        Else                    ' should be 16bpp CLP format, validate
            CopyMemory lValue, inStream(10), 2&
            If lValue = 9& Then aPtr = 8&
        End If
        If aPtr = 0& Then       ' not sure what this is, abort
            pvConvertNonPlaceableWMFtoWMF = lvicPicTypeUnknown
            Exit Function
        End If
    End If
    
    If Screen.Width > Screen.Height Then
        Cx = 256& * Screen.TwipsPerPixelX
        Cy = (Screen.Height / Screen.Width) * Cx
    Else
        Cy = 256& * Screen.TwipsPerPixelY
        Cx = (Screen.Width / Screen.Height) * Cy
    End If
    
    If hHandle Then
        Erase inStream()
        lValue = GetFileSize(hHandle, ByVal 0&)
        ReDim wmfData(0 To lValue - aPtr + 21&)
        SetFilePointer hHandle, aPtr, 0&, 0&
        ReadFile hHandle, wmfData(22), lValue - aPtr, lValue, ByVal 0&
        CloseHandle hHandle
    Else
        ReDim wmfData(0 To UBound(inStream) - aPtr + 22&)
        CopyMemory wmfData(22), inStream(aPtr), UBound(inStream) - aPtr + 1&
    End If
    CopyMemory wmfData(0), &H9AC6CDD7, 4&
    CopyMemory wmfData(10), Cx, 2&
    CopyMemory wmfData(12), Cy, 2&
    CopyMemory wmfData(14), 1440, 2&   ' ... calc checksum
    lValue = 22289& Xor Cx Xor Cy Xor 1440&
    CopyMemory wmfData(20), lValue, 2&
    
    pvConvertNonPlaceableWMFtoWMF = lvicPicTypeUnknown ' default as failure
    
    Set IIStream = IStreamFromArray(VarPtr(wmfData(0)), UBound(wmfData) + 1&)
    If Not IIStream Is Nothing Then
        If GdipLoadImageFromStream(ObjPtr(IIStream), pvConvertNonPlaceableWMFtoWMF) = 0& Then
            lValue = CreateSourcelessHandle(pvConvertNonPlaceableWMFtoWMF)
            GdipDisposeImage pvConvertNonPlaceableWMFtoWMF
            If lValue Then
                pvConvertNonPlaceableWMFtoWMF = lValue
                inStream() = wmfData()
            Else
                pvConvertNonPlaceableWMFtoWMF = lvicPicTypeUnknown ' failure
            End If
        End If
        Set IIStream = Nothing
    End If
End Function


Public Function CreateShapedRegion(SourceHandle As Long, ByVal Width As Long, ByVal Height As Long, ByVal TolerancePct As Long, _
                                Optional ByVal XOffset As Long = 0&, Optional ByVal YOffset As Long = 0&) As Long

    ' Very fast region from bitmap routine. Custom designed by LaVolpe

    ' returns a Windows region handle, or zero if no image assigned or result would be a rectangular region same size as image
    ' the TolerancePct must be a value between 1 and 255
    '   if the pixel opaqueness is <= TolerancePct then that pixel not included in the region
    
    Dim X As Long, Y As Long, a As Long, lAppendFlag As Long
    Dim r() As RECTI, rIndex As Long, rCount As Long
    Dim tSA As SafeArray, tBMP As BitmapData, imgBytes() As Byte
    Dim maxRight As Long, maxLeft As Long
    
    If SourceHandle = 0& Then Exit Function
    If Width < 0& Or Height < 0& Then Exit Function
    
    If TolerancePct < 1& Then
        Exit Function
    ElseIf TolerancePct > 255& Then
        TolerancePct = 255&
    End If
    
    rCount = Height: ReDim r(-2 To rCount - 1&)     ' initialize array with arbitrary amount of rectangle structures
    r(-1).nWidth = Width: r(-1).nHeight = Height
    If GdipBitmapLockBits(SourceHandle, r(-1), ImageLockModeRead, lvicColor32bppAlpha, tBMP) Then Exit Function
    With tSA                                        ' overlay an array on the picture data
        .cbElements = 1
        .cDims = 1
        .pvData = tBMP.Scan0Ptr
        If tBMP.stride < 0& Then .pvData = .pvData + (tBMP.Height - 1&) * tBMP.stride
        .rgSABound(0).cElements = Abs(tBMP.stride) * tBMP.Height
    End With
    CopyMemory ByVal VarPtrArray(imgBytes), VarPtr(tSA), 4&
    On Error GoTo ExitRoutine
    
    maxLeft = Width + XOffset                   ' set invalid max value (nTop & nHeight filled in at end of routine)
    maxRight = XOffset
    For Y = 0& To Height - 1&
        lAppendFlag = 0&                        ' reset for each row
        a = Y * Abs(tBMP.stride) + 3&           ' align to 1st alpha value on current row
        For X = 0& To Width - 1&
            If imgBytes(a) < TolerancePct Then  ' pixel excluded from region
                If (lAppendFlag And 1&) Then    ' currently appending
                    r(rIndex).nWidth = X + XOffset ' close out the RECT
                    rIndex = rIndex + 1&        ' increment
                    If rIndex = rCount Then     ' redim if needed
                        rCount = rCount + Height
                        ReDim Preserve r(-2 To rCount - 1&)
                    End If
                    lAppendFlag = lAppendFlag Xor 1& ' remove appending flag; keep new row flag
                End If
            ElseIf (lAppendFlag And 1&) = 0& Then
                r(rIndex).nLeft = X + XOffset   ' starting new RECT
                r(rIndex).nTop = Y + YOffset: r(rIndex).nHeight = Y + YOffset + 1&
                lAppendFlag = 3&                ' appending and a new row will be added to region
            End If
            a = a + 4&                          ' move array pointer along
        Next
        If (lAppendFlag And 1&) Then            ' handle situations where last pixel in row terminates a RECT
            r(rIndex).nWidth = X + XOffset      ' close out the RECT
            rIndex = rIndex + 1&                ' increment index & redim if necessary
            If rIndex = rCount And Y < Height - 1& Then
                rCount = rCount + Height
                ReDim Preserve r(-2 To rCount - 1&)
            End If
        End If
        If (lAppendFlag And 2&) Then            ' row added: update region bounds
            With r(rIndex - 1&)
                If .nWidth > maxRight Then maxRight = .nWidth
                If maxLeft < .nLeft Then maxLeft = .nLeft
            End With
        End If
    Next
    CopyMemory ByVal VarPtrArray(imgBytes), 0&, 4&
    GdipBitmapUnlockBits SourceHandle, tBMP
    tSA.pvData = 0&
    
    If rIndex Then                          ' we have rectangles; therefore, a region to be created
        ' call function to create region from our byte (RECT) array
        X = pvCreatePartialRegion(r(), 0&, rIndex - 1&, maxLeft, maxRight)
        ' ok, now to test whether or not we are good to go...
        ' if less than 2000 rectangles, region should have been created & if it didn't
        ' it wasn't due to O/S restrictions -- failure
        If X = 0& Then
            If rIndex > 2000& Then
                ' Win98 has limitation of approximately 4000 regional rectangles
                ' In cases of failure, we will create the region in steps of
                ' 2000 vs trying to create the region in one step
                X = pvCreateWin98Region(r(), rIndex, maxLeft, maxRight)
            End If
        End If
        CreateShapedRegion = X
    End If

ExitRoutine:
    If tSA.pvData Then
        CopyMemory ByVal VarPtrArray(imgBytes), 0&, 4&
        GdipBitmapUnlockBits SourceHandle, tBMP
    End If

End Function

Private Function pvCreatePartialRegion(rgnRects() As RECTI, lIndex As Long, uIndex As Long, leftOffset As Long, rightEdge As Long) As Long
    ' Helper function for CreateShapedRegion & pvCreateWin98Region
    ' Called to create a region in its entirety or stepped (see pvCreateWin98Region)

    On Error Resume Next
    ' Note: Ideally contiguous rectangles of equal height & width should be combined
    ' into one larger rectangle. However, thru trial & error I found that Windows
    ' does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region is the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2&) ' bytes 0-15
        .nLeft = 32&                        ' length of region header in bytes
        .nTop = 1&                          ' required cannot be anything else
        .nWidth = uIndex - lIndex + 1&      ' number of rectangles for the region
        .nHeight = .nWidth * 16&            ' byte size used by the rectangles; can be zero
    End With
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .nLeft = leftOffset                 ' left
        .nTop = rgnRects(lIndex).nTop       ' top
        .nWidth = rightEdge                 ' right
        .nHeight = rgnRects(uIndex).nHeight ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    pvCreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2&).nWidth + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear

End Function

Private Function pvCreateWin98Region(rgnRects() As RECTI, rectCount As Long, leftOffset As Long, rightEdge As Long) As Long
    ' Fall-back routine when a very large region fails to be created.
    ' Win98 has problems with regional rectangles over 4000
    ' So, we'll try again in case this is the prob with other systems too.
    ' We'll step it at 2000 at a time which is still very quick

    Dim X As Long, Y As Long ' loop counters
    Dim win98Rgn As Long     ' partial region
    Dim rtnRegion As Long    ' combined region & return value of this function
    Const RGN_OR As Long = 2&
    Const ChunkSize As Long = 2000&

    ' we start with 0 'cause first 2 RECTs are the header
    For X = 0& To rectCount - 1& Step ChunkSize
    
        If X + ChunkSize >= rectCount Then Y = rectCount Else Y = X + ChunkSize
        
        ' attempt to create partial region, scanSize rects at a time
        win98Rgn = pvCreatePartialRegion(rgnRects(), X, Y - 1&, leftOffset, rightEdge)
        
        If win98Rgn = 0& Then    ' failure
            ' cleaup combined region if needed
            If rtnRegion Then
                DeleteObject rtnRegion
                rtnRegion = 0&
            End If
            Exit For ' abort; system won't allow us to create the region
        Else
            If rtnRegion = 0& Then ' first time thru
                rtnRegion = win98Rgn
            Else ' already started
                ' use combineRgn, but only every scanSize times
                CombineRgn rtnRegion, rtnRegion, win98Rgn, RGN_OR
                DeleteObject win98Rgn
            End If
        End If
    Next
    ' done; return result
    pvCreateWin98Region = rtnRegion
    
End Function

Public Function EnumResNamesProcA(ByVal hMod As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByRef lParam As Long) As Long
    
    ' ANSI version. Called to retrieve resource IDs

    If lParam < 0& Then ' counting only
        lParam = ((lParam And &H7FFFF) + 1&) Or &H80000000
        If lParam < &H7FFF& Then EnumResNamesProcA = 1&  ' continue enumerating
    Else                                ' looking for specific item (0 based)
        If (lParam And &H7FFF&) = 0& Then ' got it
            Dim sName As String
            Dim lLen As Long
            Dim b() As Byte
        
            If (lpszName And &HFFFF0000) = 0 Then        ' numeric ID
                sName = "#" & CStr(lpszName And &HFFFF&)
            Else
                lLen = lstrlen(lpszName)           ' string ID
                If (lLen > 0&) Then
                    ReDim b(0 To lLen - 1&)
                    CopyMemory b(0), ByVal lpszName, lLen
                    sName = StrConv(b, vbUnicode)
                Else
                    lParam = 0&: Exit Function
                End If
            End If
            lParam = StrPtr(sName)
            CopyMemory ByVal VarPtr(sName), 0&, 4&
        Else
            lParam = (lParam And &H7FFF&) - 1&
            EnumResNamesProcA = 1&  ' continue enumerating
        End If
    End If

End Function

Public Function EnumResNamesProcW(ByVal hMod As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByRef lParam As Long) As Long
    
    ' UNICODE version. Called to retrieve resource IDs
    
    If lParam < 0& Then ' counting only
        lParam = ((lParam And &H7FFF&) + 1&) Or &H80000000
        If lParam < &H7FFF& Then EnumResNamesProcW = 1& ' continue enumerating
    Else                                ' looking for specific item (0 based)
        If (lParam And &H7FFF&) = 0& Then ' got it
            Dim sName As String
            Dim lLen As Long
        
            If (lpszName And &HFFFF0000) = 0 Then        ' numeric ID
                sName = "#" & CStr(lpszName And &HFFFF&)
            Else
                lLen = lstrlenW(lpszName)           ' string ID
                If (lLen > 0&) Then
                    sName = Space$(lLen)
                    CopyMemory ByVal StrPtr(sName), ByVal lpszName, lLen * 2&
                Else
                    lParam = 0&: Exit Function
                End If
            End If
            lParam = StrPtr(sName)
            CopyMemory ByVal VarPtr(sName), 0&, 4&
        Else
            lParam = (lParam And &H7FFF&) - 1&
            EnumResNamesProcW = 1&              ' continue enumerating
        End If
   End If
   
End Function

Public Function ReverseLong(ByVal inLong As Long) As Long

    ' fast function to reverse a long value from big endian to little endian
    ' PNG files contain reversed longs, as do ID3 v3,4 tags
    ReverseLong = _
      (((inLong And &HFF000000) \ &H1000000) And &HFF&) Or _
      ((inLong And &HFF0000) \ &H100&) Or _
      ((inLong And &HFF00&) * &H100&) Or _
      ((inLong And &H7F&) * &H1000000)
    If (inLong And &H80&) Then ReverseLong = ReverseLong Or &H80000000
End Function

Private Function pvValidateMultiImgStruct(MIS As MULTIIMAGESAVESTRUCT, _
                                          DefaultPicture As GDIpImage, DestType As Long, _
                                          DestFormat As ImageFormatEnum) As Boolean

    ' Helper function for SaveImage routine
    ' SaveImage supports saving multi-image formats: ICO,CUR,GIF,TIF
    ' The new MULTIIMAGESAVESTRUCT structure supports supplying as a image source, an image
    '   the can contain multiple images. This routine validates passed information and clones
    '   source image as necessary. See comments sprinkled about below
    
    ' modified to include some pre-processing to support APNG creation

    Dim f As Long, p As Long, c As Long
    Dim lClone() As Long, lb As Long, tImg As GDIpImage

    On Error Resume Next
    If (UBound(MIS.Image) - LBound(MIS.Image)) + 1& < MIS.Images Then
        If Err Then
            Err.Clear
            Exit Function                           ' passed invalid structure
        End If
        MIS.Images = (UBound(MIS.Image) - LBound(MIS.Image)) + 1&
    End If
    If MIS.Images < 1& Then Exit Function
    On Error GoTo 0
    
    ' skip any frames where user supplied a null image. Identify frames that require cloning
    For lb = (LBound(MIS.Image) + MIS.Images - 1&) To LBound(MIS.Image) Step -1&
        If MIS.Image(lb).Picture Is Nothing Then
            c = 0&                                  ' invalid picture
        Else
            c = (MIS.Image(lb).Picture.Handle <> 0&) ' valid handle?
        End If
        If c Then
            f = f + 1&
            If MIS.Image(lb).GroupNumber > 0& Then
                If MIS.Image(lb).Picture.ImageGroupCount < 2& Then MIS.Image(lb).GroupNumber = 0&
            End If
            If MIS.Image(lb).GroupNumber < 1& Then
                If MIS.Image(lb).FrameNumber > 0& Then                      ' index change requested
                    If MIS.Image(lb).Picture.ImageCount < 2& Then           ' more than 1 image?
                        MIS.Image(lb).FrameNumber = 0&                      ' else already on the index?
                    ElseIf MIS.Image(lb).Picture.ImageIndex = MIS.Image(lb).FrameNumber Then
                        MIS.Image(lb).FrameNumber = 0&
                    End If
                End If
            End If
        Else                                        ' shift images down the array to exclude invalid one
            For p = lb + 1& To (LBound(MIS.Image) + MIS.Images - 1&)
                MIS.Image(p - 1&) = MIS.Image(p)
            Next
        End If
    Next
        
    lb = LBound(MIS.Image)
    If f = 0& Then                                  ' use passed DefaultPicture
        If DefaultPicture Is Nothing Then Exit Function
        Set MIS.Image(lb).Picture = DefaultPicture
        If MIS.Image(lb).GroupNumber > 0& Then
            If DefaultPicture.ImageGroupCount < 2& Then MIS.Image(lb).GroupNumber = 0&
        End If
        If MIS.Image(lb).GroupNumber < 1& Then
            If MIS.Image(lb).FrameNumber > 1& Then      ' index change requested?
                If DefaultPicture.ImageCount < 2& Then  ' more than 1 image?
                    MIS.Image(lb).FrameNumber = 0&      ' else already on the index?
                ElseIf DefaultPicture.ImageIndex = MIS.Image(lb).FrameNumber Then
                    MIS.Image(lb).FrameNumber = 0&
                End If
            End If
        End If
        MIS.Images = 1&
    ElseIf DestType = saveTo_stdPicture Then
        MIS.Images = 1&                             ' no multi-image stdPictures
    Else
        MIS.Images = f
    End If
    
    ' If using a source multi-image format and need to navigate to non-current frame then create clone
    ' Else if we change the image index, any source control using that image will display the change. Not desired
    ReDim lClone(0 To MIS.Images - 1&)
    For f = 0& To MIS.Images - 1&
        MIS.Image(lb + f).SS.ExtractCurrentFrameOnly = True
        If MIS.Image(lb + f).FrameNumber > 0& Then  ' frame change required
            ' Not on the current frame. Did we already clone this image?
            c = ObjPtr(MIS.Image(lb + f).Picture)
            For p = 0& To f - 1&
                If lClone(p) = c Then
                    ' already cloned
                    Set MIS.Image(lb + f).Picture = MIS.Image(p + lb).Picture
                    Exit For
                End If
            Next
            If p = f Then ' didn't clone it, need to do that now
                lClone(f) = c
                Set MIS.Image(lb + f).Picture = LoadImage(MIS.Image(lb + f).Picture, , , True)
            End If
        End If
    Next
    
    If MIS.Images > 1& And DestFormat = lvicPicTypePNG Then
        Dim Cx() As Long, Cy() As Long
        ' see cFunctionsPNG.SaveAsAPNG for more info regarding why we are forcing to 32bpp
        
        ReDim Cx(0 To MIS.Images)
        ReDim Cy(0 To MIS.Images)
        ' Unlike an animated GIF, the overall canvas size is not a separate entry in the animated format.
        '   The overall canvas size is defined by the 1st image in the animated PNG format; therefore,
        '   the overall canvas size needs to be calculated in advance...
        For f = 0& To MIS.Images - 1&
            With MIS.Image(lb + f)
                .SS.ColorDepth = lvicConvert_TrueColor32bpp_ARGB
                If .SS.Width < 1& Then .SS.Width = .Picture.Width
                If .SS.Height < 1& Then .SS.Height = .Picture.Height
                If Not (.SS.RSS.Angle = 0! Or .SS.RSS.Angle = 180!) Then
                    GetScaledCanvasSize .SS.Width, .SS.Height, Cx(f), Cy(f), .SS.RSS.Angle
                    If .SS.RotatingCanGrowImage = False Then
                        ' a bit awkward, but basically need to force .Width/.Height to values that
                        ' will become their eventual actual sizes
                        ' 1. determine size of rotated dimensions with no size restrictions
                        GetScaledCanvasSize .SS.Width, .SS.Height, Cx(f), Cy(f), .SS.RSS.Angle
                        ' 2. scale that bounding box to destination dimensions
                        GetScaledImageSizes Cx(f), Cy(f), .SS.Width, .SS.Height, Cx(f), Cy(f), , False, False
                        ' 3. scale the desired dimensions to that scaled bounding box
                        GetScaledImageSizes .SS.Width, .SS.Height, Cx(f), Cy(f), Cx(MIS.Images), Cy(MIS.Images), .SS.RSS.Angle, False, False
                        ' 4. update the desired width/height & allow rotation to resize overall image size
                        .SS.Width = Cx(MIS.Images): .SS.Height = Cy(MIS.Images)
                        .SS.RotatingCanGrowImage = True
                    End If
                Else
                    Cx(f) = .SS.Width: Cy(f) = .SS.Height
                End If                                              ' adjust canvas size if offets used
                If .GIFFrameInfo.XOffset > 0& Then Cx(f) = Cx(f) + .GIFFrameInfo.XOffset
                If .GIFFrameInfo.YOffset > 0& Then Cy(f) = Cy(f) + .GIFFrameInfo.YOffset
                                                                    ' update/fix overall canvas size as needed
                If Cx(f) > MIS.GIFOverview.WindowWidth Then MIS.GIFOverview.WindowWidth = Cx(f)
                If Cy(f) > MIS.GIFOverview.WindowHeight Then MIS.GIFOverview.WindowHeight = Cy(f)
            End With
        Next
        For f = 0& To MIS.Images - 1&
            With MIS.Image(lb + f)  ' auto-center as needed
                If .GIFFrameInfo.XOffset < 0& Then .GIFFrameInfo.XOffset = (MIS.GIFOverview.WindowWidth - Cx(f)) \ 2&
                If .GIFFrameInfo.YOffset < 0& Then .GIFFrameInfo.YOffset = (MIS.GIFOverview.WindowHeight - Cy(f)) \ 2&
            End With
        Next
    End If
    
    pvValidateMultiImgStruct = True

End Function

Private Function pvBuildAVIconverter(srcImage As GDIpImage, MIS As MULTIIMAGESAVESTRUCT, _
                                    ImageFormat As ImageFormatEnum, Effects As GDIpEffects) As Boolean

    ' function used to set up conversion from AVI to animated GIF/PNG

    Dim cImage As GDIpImage
    Dim f As Long
    
    Set cImage = modCommon.LoadImage(srcImage, , , True)
    If cImage.Handle = 0& Then Exit Function
    
    MIS.Images = cImage.ImageCount
    ReDim MIS.Image(1 To MIS.Images)
    With MIS.Image(1)
        .FrameNumber = 1&
        Set .Picture = cImage
        .SS.Width = .Picture.Width
        .SS.Height = .Picture.Height
        .GIFFrameInfo.DelayTime = .Picture.FrameDuration(.FrameNumber)
        .GIFFrameInfo.DisposalCode = lvicGIF_Erase
        If ImageFormat = lvicPicTypeGIF Then
            .SS.ColorDepth = lvicConvert_256Colors
            .SS.AlphaTolerancePct = 25
            .SS.PaletteType = lvicPaletteAdaptive
        End If
        If Not Effects Is Nothing Then
            Set .SS.RSS.Effects = New GDIpEffects
            If Effects.TransparentColorUsed = False Then
                Call GdipBitmapGetPixel(.Picture.Handle, 0&, 0&, f)
                .SS.RSS.Effects.TransparentColor = Color_ARGBtoRGB(f)
            Else
                .SS.RSS.Effects.TransparentColor = Effects.TransparentColor
            End If
            .SS.RSS.Effects.TransparentColorUsed = True
        End If
    End With
    
    For f = 2& To MIS.Images
        MIS.Image(f) = MIS.Image(1)
        MIS.Image(f).FrameNumber = f
    Next
    MIS.GIFOverview.WindowHeight = MIS.Image(1).SS.Height
    MIS.GIFOverview.WindowWidth = MIS.Image(1).SS.Width

    pvBuildAVIconverter = True
    
End Function

