VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clGdiplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
'*              CLASSE POUR UTILISATION DE GDIPLUS EN VBA                              *
'***************************************************************************************

'***************************************************************************************
' Auteur : Thierry GASPERMENT (Arkham46)
'
' v3.0 (08/04/2018)
'
' Necessite la librairie GDI+
' Le code est libre pour toute utilisation
'***************************************************************************************

'***************************************************************************************
'*                                      LIENS                                          *
'***************************************************************************************
'** Telechargement de GdiPlus.dll
' http://www.microsoft.com/downloads/details.aspx?FamilyID=6a63ab9c-df12-4d41-933c-be590feaa05a&DisplayLang=en
'** Norme Exif
' http://www.exif.org/specifications.html
'** Liste des Tags EXIF pour Gdi+
' http://msdn2.microsoft.com/en-us/library/ms534413.aspx
'**
' Mon Site Web DVP : http://arkham46.developpez.com
'***************************************************************************************
'*                                      EN-TETE                                        *
'***************************************************************************************
Option Explicit
' A = Access
' W = Word
' E = Excel
' P = Powerpoint
#Const AppName = "E"
#Const UseSingle = False
#If AppName = "A" Then
    Option Compare Database
#End If
' Definition des variables non typees en long pour 32 bits ou longptr pour 64 bits
' Les elements des types doivent etre types obligatoirement
#If VBA7 Then
DefLngPtr A-Z
Const PtrNull As LongPtr = 0
#Else
DefLng A-Z
Const PtrNull As Long = 0
#End If

'***************************************************************************************
'*              ScrollBars
'***************************************************************************************
Private goBars As Collection ' collection d'objet pour gestion de barres de defilement
#If AppName = "A" Then
Private WithEvents oImg As Access.Image ' Controle image
Attribute oImg.VB_VarHelpID = -1
Private oForm As Object ' Formulaire
Private WithEvents oBarImage As Access.Image ' Controle image
Attribute oBarImage.VB_VarHelpID = -1
Private oBarForm As Object ' Formulaire
#Else
Private WithEvents oImg As MSForms.Image ' Controle image
Attribute oImg.VB_VarHelpID = -1
Private WithEvents oForm As MSForms.UserForm ' Formulaire
Attribute oForm.VB_VarHelpID = -1
Private WithEvents oBarImage As MSForms.Image ' Controle image
Attribute oBarImage.VB_VarHelpID = -1
Private WithEvents oBarForm As MSForms.UserForm ' Formulaire
Attribute oBarForm.VB_VarHelpID = -1
#End If
Private oObject As Object
Private gImgLastClickRegion As String
Private gImgLastClickX As Single
Private gImgLastClickY As Single
Private gImgLastClickButton As Integer
Private gImgLastClickShift As Integer
Private goImgEventObject As Object
Private gImgEventMouseDown As String
Private gImgEventMouseUp As String
Private gImgEventMouseMove As String
Private gImgEventDblClick As String

Private oBarObject As Object
Private gBarLastClickRegion As String
Private gBarLastClickBar As String
Private gBarLastClickX As Single
Private gBarLastClickY As Single
Private gBarLastClickButton As Integer
Private gBarLastClickShift As Integer
Private gBarKey As String
Private gBarClickX As Single
Private gBarClickY As Single
Private gBarDragRegion As String
Private gBarDragBar As String
Private goBarEventObject As Object
Private gBarEventMouseDown As String
Private gBarEventMouseUp As String
Private gBarEventMouseMove As String
Private gBarEventDblClick As String
Private gBarEventOnRefreshNeeded As String
' Evenements
#If VBA6 Then
Public Event ImgMouseDown(pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ImgMouseUp(pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ImgMouseMove(pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ImgMouseDblClick(pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event BarMouseDown(BarName As String, pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event BarMouseUp(BarName As String, pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event BarMouseMove(BarName As String, pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event BarMouseDblClick(BarName As String, pRegion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event BarOnRefreshNeeded(BarName As String, MouseUp As Boolean)
#End If
'***************************************************************************************
'*                                       API                                           *
'***************************************************************************************
#If VBA7 Then
Private Declare PtrSafe Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean
Private Declare PtrSafe Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As LongPtr, ByVal dwUser As LongPtr, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Private Declare PtrSafe Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As LongPtr, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare PtrSafe Function ResetEvent Lib "kernel32" (ByVal hEvent As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function SHGetFileInfo Lib "SHELL32.DLL" Alias "SHGetFileInfoA" (ByVal pPath As String, _
    ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function ValidateRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function ValidateRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As Rect) As Long
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As Rect) As Long
Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hwnd As LongPtr, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare PtrSafe Function GetDlgCtrlID Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function OleTranslateColor Lib "oleAut32.dll" _
                                           (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, _
                                                                ByRef ppstm As Any) As Long
Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As Object, phglobal As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function SetThemeAppProperties Lib "uxtheme" (ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function SetStretchBltMode Lib "Gdi32" (ByVal hDC As LongPtr, ByVal nStretchMode As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As Rect) As Long
Private Declare PtrSafe Function GetDIBits Lib "Gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function CreateDIBPatternBrush Lib "Gdi32" (ByVal hPackedDIB As LongPtr, ByVal wUsage As Long) As LongPtr
Private Declare PtrSafe Function CreateDIBPatternBrushPt Lib "Gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As LongPtr
Private Declare PtrSafe Function CreatePatternBrush Lib "Gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function SetBrushOrgEx Lib "Gdi32" (ByVal hDC As LongPtr, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
Private Declare PtrSafe Function GetObjectBmp Lib "Gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function SelectClipRgn Lib "Gdi32" (ByVal hDC As LongPtr, ByVal hRgn As LongPtr) As Long
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As Rect, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function GetIconInfo Lib "user32" (ByVal hIcon As LongPtr, piconinfo As ICONINFO) As Long
Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleAut32.dll" _
                        (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal str As LongPtr, id As GUID) As Long
Private Declare PtrSafe Function StringFromCLSID Lib "ole32.dll" (pclsid As GUID, lpszProgID As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DrawEdge Lib "user32" (ByVal hDC As LongPtr, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare PtrSafe Function DrawFrameControl Lib "user32" (ByVal hDC As LongPtr, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare PtrSafe Function SetWorldTransform Lib "Gdi32" (ByVal hDC As LongPtr, lpXform As XForm) As Long
Private Declare PtrSafe Function CreateDIBSection Lib "Gdi32" (ByVal hDC As LongPtr, pBitmapInfo As BitmapInfo, ByVal un As Long, lplpVoid As LongPtr, ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
Private Declare PtrSafe Function SetGraphicsMode Lib "Gdi32" (ByVal hDC As LongPtr, ByVal iMode As Long) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "Gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function DrawIconEx Lib "user32" (ByVal hDC As LongPtr, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As LongPtr, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As LongPtr, ByVal diFlags As Long) As Long
Private Declare PtrSafe Function SetMapMode Lib "Gdi32" (ByVal hDC As LongPtr, ByVal nMapMode As Long) As Long
Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "Gdi32" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function WindowFromDC Lib "user32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "Gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "Gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentObject Lib "Gdi32" (ByVal hDC As LongPtr, ByVal uObjectType As Long) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "Gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "Gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function DPtoLP Lib "Gdi32" (ByVal hDC As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare PtrSafe Function LPtoDP Lib "Gdi32" (ByVal hDC As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare PtrSafe Function SetEnhMetaFileBits Lib "Gdi32" (ByVal cbBuffer As Long, lpData As Byte) As LongPtr
Private Declare PtrSafe Function DeleteEnhMetaFile Lib "Gdi32" (ByVal hemf As LongPtr) As Long
Private Declare PtrSafe Function GetEnhMetaFileBits Lib "Gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Any) As Long
Private Declare PtrSafe Function GetEnhMetaFileHeader Lib "Gdi32" (ByVal hemf As LongPtr, ByVal cbBuffer As Long, lpEMH As ENHMETAHEADER) As Long
Private Declare PtrSafe Function PlayEnhMetaFile Lib "Gdi32" (ByVal hDC As LongPtr, ByVal hemf As LongPtr, lpRect As Rect) As Long
Private Declare PtrSafe Function FrameRgn Lib "Gdi32" (ByVal hDC As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "Gdi32" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
Private Declare PtrSafe Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal fontCollection As LongPtr, ByVal memory As LongPtr, ByVal Length As Long) As Long
Private Declare PtrSafe Function GdipNewInstalledFontCollection Lib "gdiplus" (fontCollection As LongPtr) As Long
Private Declare PtrSafe Function GdipGetFamilyName Lib "gdiplus" (ByVal family As LongPtr, ByVal name As LongPtr, ByVal language As Integer) As Long
Private Declare PtrSafe Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal fontCollection As LongPtr, numFound As Long) As Long
Private Declare PtrSafe Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal fontCollection As LongPtr, ByVal numSought As Long, gpfamilies As LongPtr, numFound As Long) As Long
Private Declare PtrSafe Function GdipNewPrivateFontCollection Lib "gdiplus" (fontCollection As LongPtr) As Long
Private Declare PtrSafe Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal fontCollection As LongPtr, ByVal FileName As LongPtr) As Long
Private Declare PtrSafe Function GdipDeletePrivateFontCollection Lib "gdiplus" (fontCollection As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As LongPtr, ByVal Brush As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateTexture Lib "gdiplus" (ByVal Image As LongPtr, ByVal WrapMd As Long, texture As LongPtr) As Long
Private Declare PtrSafe Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal Brush As LongPtr, ByVal WrapMd As Long) As Long
Private Declare PtrSafe Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal Brush As LongPtr, WrapMd As Long) As Long
Private Declare PtrSafe Function GdipGetTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr, ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipSetTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr, ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipMultiplyTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr, ByVal matrix As LongPtr, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipScaleTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr, ByVal sx As Single, ByVal sy As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipRotateTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr, ByVal angle As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipResetTextureTransform Lib "gdiplus" (ByVal Brush As LongPtr) As Long
Private Declare PtrSafe Function GdipGetTextureImage Lib "gdiplus" (ByVal Brush As LongPtr, Image As LongPtr) As Long
Private Declare PtrSafe Function GdipGetCompositingMode Lib "gdiplus" (ByVal Graphics As LongPtr, CompositingMd As Long) As Long
Private Declare PtrSafe Function GdipSetCompositingMode Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal CompositingMd As Long) As Long
Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal PixOffsetMode As Long) As Long
Private Declare PtrSafe Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As LongPtr, PixOffsetMode As Long) As Long
Private Declare PtrSafe Function GdipSetCompositingQuality Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal CompositingQlty As Long) As Long
Private Declare PtrSafe Function GdipGetCompositingQuality Lib "gdiplus" (ByVal Graphics As LongPtr, CompositingQlty As Long) As Long
Private Declare PtrSafe Function GdipSetClipRegion Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Region As LongPtr, ByVal CombineMd As Long) As Long
Private Declare PtrSafe Function GdipCreateRegion Lib "gdiplus" (Region As LongPtr) As Long
Private Declare PtrSafe Function GdipSetEmpty Lib "gdiplus" (ByVal Region As LongPtr) As Long
Private Declare PtrSafe Function GdipSetInfinite Lib "gdiplus" (ByVal Region As LongPtr) As Long
Private Declare PtrSafe Function GdipResetClip Lib "gdiplus" (ByVal Graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPenMode Lib "gdiplus" (ByVal pen As LongPtr, ByVal penMode As Long) As Long
Private Declare PtrSafe Function GdipCombineRegionRectI Lib "gdiplus" (ByVal Region As LongPtr, Rect As Rect, ByVal CombineMd As Long) As Long
Private Declare PtrSafe Function GdipCombineRegionRegion Lib "gdiplus" (ByVal Region As LongPtr, ByVal region2 As LongPtr, ByVal CombineMd As Long) As Long
Private Declare PtrSafe Function GdipGetRegionHRgn Lib "gdiplus" (ByVal Region As LongPtr, ByVal Graphics As LongPtr, hRgn As LongPtr) As Long
Private Declare PtrSafe Function GdipIsEmptyRegion Lib "gdiplus" (ByVal Region As LongPtr, ByVal Graphics As LongPtr, result As Long) As Long
Private Declare PtrSafe Function GdipDeleteRegion Lib "gdiplus" (ByVal Region As LongPtr) As Long
Private Declare PtrSafe Function GdipCombineRegionPath Lib "gdiplus" (ByVal Region As LongPtr, ByVal Path As LongPtr, ByVal CombineMd As Long) As Long
Private Declare PtrSafe Function GdipCreateRegionPath Lib "gdiplus" (ByVal Path As LongPtr, Region As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateMatrix Lib "gdiplus" (matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateMatrix2 Lib "gdiplus" (ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single, matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipInvertMatrix Lib "gdiplus" (ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipGetMatrixElements Lib "gdiplus" (ByVal matrix As LongPtr, matrixOut As Single) As Long
Private Declare PtrSafe Function GdipCloneMatrix Lib "gdiplus" (ByVal matrix As LongPtr, cloneMatrix As LongPtr) As Long
Private Declare PtrSafe Function GdipMultiplyMatrix Lib "gdiplus" (ByVal matrix As LongPtr, ByVal matrix2 As LongPtr, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipTranslateMatrix Lib "gdiplus" (ByVal matrix As LongPtr, ByVal offsetX As Single, ByVal offsetY As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipRotateMatrix Lib "gdiplus" (ByVal matrix As LongPtr, ByVal angle As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipScaleMatrix Lib "gdiplus" (ByVal matrix As LongPtr, ByVal scaleX As Single, ByVal scaleY As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipShearMatrix Lib "gdiplus" (ByVal matrix As LongPtr, ByVal shearX As Single, ByVal shearY As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipIsMatrixIdentity Lib "gdiplus" (ByVal matrix As LongPtr, result As Long) As Long
Private Declare PtrSafe Function GdipDeleteMatrix Lib "gdiplus" (ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawPath Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, ByVal Path As LongPtr) As Long
Private Declare PtrSafe Function GdipClonePath Lib "gdiplus" (ByVal Path As LongPtr, clonePath As LongPtr) As Long
Private Declare PtrSafe Function GdipCloneRegion Lib "gdiplus" (ByVal Region As LongPtr, cloneRegion As LongPtr) As Long
Private Declare PtrSafe Function GdipTransformPath Lib "gdiplus" (ByVal Path As LongPtr, ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipTransformRegion Lib "gdiplus" (ByVal Region As LongPtr, ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipFillPath Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, ByVal Path As LongPtr) As Long
Private Declare PtrSafe Function GdipFillRegion Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, ByVal Region As LongPtr) As Long
Private Declare PtrSafe Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal Region As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Graphics As LongPtr, result As Long) As Long
Private Declare PtrSafe Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal Region As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Graphics As LongPtr, result As Long) As Long
Private Declare PtrSafe Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As LongPtr, Path As LongPtr) As Long
Private Declare PtrSafe Function GdipDeletePath Lib "gdiplus" (ByVal Path As LongPtr) As Long
Private Declare PtrSafe Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal Path As LongPtr, polyGradient As LongPtr) As Long
Private Declare PtrSafe Function GdipStartPathFigure Lib "gdiplus" (ByVal Path As LongPtr) As Long
Private Declare PtrSafe Function GdipClosePathFigure Lib "gdiplus" (ByVal Path As LongPtr) As Long
Private Declare PtrSafe Function GdipAddPathPath Lib "gdiplus" (ByVal Path As LongPtr, ByVal addingPath As LongPtr, ByVal bConnect As Long) As Long
Private Declare PtrSafe Function GdipAddPathString Lib "gdiplus" (ByVal Path As LongPtr, ByVal str As LongPtr, ByVal Length As Long, ByVal family As LongPtr, ByVal style As Long, ByVal emSize As Single, layoutRect As RECTF, ByVal StringFormat As LongPtr) As Long
Private Declare PtrSafe Function GdipFillPolygonI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, Points As POINTAPI, ByVal count As Long, ByVal FillMd As Long) As Long
Private Declare PtrSafe Function GdipDrawPolygonI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, Points As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipDrawBeziersI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, Points As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal X As Long, ByVal Y As Long, Color As Long) As Long
Private Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal lColor As Long) As Long
Private Declare PtrSafe Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As LongPtr, Format As GUID) As Long
Private Declare PtrSafe Function GdipRecordMetafileI Lib "gdiplus" (ByVal referenceHdc As LongPtr, ByVal etype As Long, frameRect As Rect, ByVal frameUnit As Long, ByVal description As LongPtr, metafile As LongPtr) As Long
Private Declare PtrSafe Function GdipRecordMetafileFileNameI Lib "gdiplus" (ByVal FileName As LongPtr, ByVal referenceHdc As LongPtr, ByVal etype As Long, frameRect As Rect, ByVal frameUnit As LongPtr, ByVal description As LongPtr, metafile As LongPtr) As Long
Private Declare PtrSafe Function GdipGetDC Lib "gdiplus" (ByVal Graphics As LongPtr, hDC As LongPtr) As Long
Private Declare PtrSafe Function GdipReleaseDC Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef Image As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSetImagePalette Lib "gdiplus" (ByVal Image As LongPtr, palette As ColorPalette) As Long
Private Declare PtrSafe Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As LongPtr, ByVal dStyle As Long) As Long
Private Declare PtrSafe Function GdipSetPenDashArray Lib "gdiplus" (ByVal pen As LongPtr, dash As Single, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As Long, ByVal forecolr As Long, ByVal backcolr As Long, Brush As LongPtr) As Long
Private Declare PtrSafe Function GdipBitmapSetResolution Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal xdpi As Single, ByVal ydpi As Single) As Long
Private Declare PtrSafe Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As LongPtr, Graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Graphics As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BitmapInfo, ByVal gdiBitmapData As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCloneImage Lib "gdiplus" (ByVal Image As LongPtr, cloneImage As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As LongPtr, PixelFormat As Long) As Long
Private Declare PtrSafe Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As LongPtr, totalBufferSize As Long, numProperties As Long) As Long
Private Declare PtrSafe Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As LongPtr, ByVal totalBufferSize As Long, ByVal numProperties As Long, allItems As Any) As Long
Private Declare PtrSafe Function GdipRemovePropertyItem Lib "gdiplus" (ByVal Image As LongPtr, ByVal propId As Long) As Long
Private Declare PtrSafe Function GdipAlloc Lib "gdiplus.dll" (ByVal size As LongPtr) As LongPtr
Private Declare PtrSafe Function GdipFree Lib "gdiplus.dll" (ByVal ptr As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As LongPtr, ByVal stream As Object, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As LongPtr, ByVal propId As Long, _
                                                            ByVal propSize As Long, ByRef buffer As Any) As Long
Private Declare PtrSafe Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As LongPtr, ByVal propId As Long, _
                                                                ByRef size As Long) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, LInput As GdiplusStartupInput, Optional ByVal lOutPut As LongPtr = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef hbmReturn As LongPtr, _
                                                                    ByVal Background As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As LongPtr, ByVal hpal As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr) As Long
Private Declare PtrSafe Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As LongPtr, ByVal rfType As Long) As Long
Private Declare PtrSafe Function GdipImageSelectActiveFrame Lib "gdiplus" _
                                                    (ByVal Image As LongPtr, ByRef dimensionID As GUID, _
                                                     ByVal frameIndex As Long) As Long
Private Declare PtrSafe Function GdipImageGetFrameCount Lib "gdiplus" _
                                                (ByVal Image As LongPtr, ByRef dimensionID As GUID, _
                                                 ByRef count As Long) As Long
Private Declare PtrSafe Function GdipGetImageDimension Lib "gdiplus" _
                                               (ByVal Image As LongPtr, ByRef Width As Single, _
                                                ByRef Height As Single) As Long
Private Declare PtrSafe Function GdipSetPropertyItem Lib "gdiplus" (ByVal nImage As LongPtr, item As PropertyItem) As Long
Private Declare PtrSafe Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal Image As LongPtr, resolution As Single) As Long
Private Declare PtrSafe Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal Image As LongPtr, resolution As Single) As Long
Private Declare PtrSafe Function GdipGetPropertyCount Lib "gdiplus" (ByVal Image As LongPtr, numOfProperty As Long) As Long
Private Declare PtrSafe Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As Any, ByRef Image As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As LongPtr, Height As Long) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As LongPtr, Width As Long) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, Rect As Rect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As LongPtr, lockedBitmapData As BitmapData) As Long
Private Declare PtrSafe Function GdipResetWorldTransform Lib "gdiplus" (ByVal Graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipGetWorldTransform Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipSetWorldTransform Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal matrix As LongPtr) As Long
Private Declare PtrSafe Function GdipScaleWorldTransform Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal sx As Single, ByVal sy As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipRotateWorldTransform Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal angle As Single, ByVal order As Long) As Long
Private Declare PtrSafe Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As LongPtr) As Long
Private Declare PtrSafe Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As LongPtr, ByVal align As Long) As Long
Private Declare PtrSafe Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As LongPtr, ByVal align As Long) As Long
Private Declare PtrSafe Function GdipMeasureString Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal str As LongPtr, ByVal Length As Long, ByVal thefont As LongPtr, layoutRect As RECTF, ByVal StringFormat As LongPtr, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As Long
Private Declare PtrSafe Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal mode As Long) As Long
Private Declare PtrSafe Function GdipGetTextRenderingHint Lib "gdiplus" (ByVal Graphics As LongPtr, mode As Long) As Long
Private Declare PtrSafe Function GdipDrawString Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal str As LongPtr, ByVal Length As Long, ByVal thefont As LongPtr, layoutRect As RECTF, ByVal StringFormat As LongPtr, ByVal Brush As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As LongPtr, ByVal emSize As Single, ByVal style As Long, ByVal unit As Long, createdfont As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As LongPtr) As Long
Private Declare PtrSafe Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (nativeFamily As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As LongPtr, ByVal fontCollection As LongPtr, fontFamily As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As LongPtr) As Long
Private Declare PtrSafe Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal imageattr As LongPtr, ByVal ClrAdjType As Long, ByVal enableFlag As Long, ByVal colorLow As Long, ByVal colorHigh As Long) As Long
Private Declare PtrSafe Function GdipSetImageAttributesRemapTable Lib "gdiplus" (ByVal imageattr As LongPtr, ByVal ClrAdjType As Long, ByVal enableFlag As Long, ByVal mapSize As Long, map As ColorMap) As Long
Private Declare PtrSafe Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageattr As LongPtr, ByVal wrap As Long, ByVal argb As Long, ByVal bClamp As Long) As Long
Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipGetInterpolationMode Lib "gdiplus" (ByVal Graphics As LongPtr, interpolation As Long) As Long
Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal interpolation As Long) As Long
Private Declare PtrSafe Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal SmoothingMd As Long) As Long
Private Declare PtrSafe Function GdipGetSmoothingMode Lib "gdiplus" (ByVal Graphics As LongPtr, SmoothingMd As Long) As Long
Private Declare PtrSafe Function GdipDrawLine Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare PtrSafe Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As LongPtr, ByVal startCap As Long) As Long
Private Declare PtrSafe Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As LongPtr, ByVal endCap As Long) As Long
Private Declare PtrSafe Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As LongPtr, ByVal LnJoin As Long) As Long
Private Declare PtrSafe Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal pen As LongPtr, ByVal miterLimit As Single) As Long
Private Declare PtrSafe Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal pen As LongPtr, miterLimit As Single) As Long
Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As LongPtr, Graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, pen As LongPtr) As Long
Private Declare PtrSafe Function GdipDeletePen Lib "gdiplus" (ByVal pen As LongPtr) As Long
Private Declare PtrSafe Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long
Private Declare PtrSafe Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As LongPtr) As Long
Private Declare PtrSafe Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As LongPtr, ByVal ClrAdjType As Long, ByVal enableFlag As Long, colourMatrix As ColorMatrix, grayMatrix As Any, ByVal flags As Long) As Long
Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Image As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, Brush As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawRectangleI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipDrawArcI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal pen As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare PtrSafe Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal metafile As LongPtr, hemf As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPathFillMode Lib "gdiplus" (ByVal Path As LongPtr, ByVal brushmode As Long) As Long
Private Declare PtrSafe Function GdipCreateAdjustableArrowCap Lib "gdiplus" (ByVal Height As Single, ByVal Width As Single, ByVal isFilled As Long, cap As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPenCustomStartCap Lib "gdiplus" (ByVal pen As LongPtr, ByVal customCap As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPenCustomEndCap Lib "gdiplus" (ByVal pen As LongPtr, ByVal customCap As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteCustomLineCap Lib "gdiplus" (ByVal customCap As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Image As LongPtr, ByVal dstx As Long, _
                                                               ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, _
                                                               ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As LongPtr = PtrNull, _
                                                               Optional ByVal CallBack As LongPtr = PtrNull, Optional ByVal callbackData As LongPtr = PtrNull) As Long
Private Declare PtrSafe Function GdipAddPathRectangleI Lib "gdiplus" (ByVal Path As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipAddPathEllipseS Lib "gdiplus" Alias "GdipAddPathEllipse" (ByVal Path As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare PtrSafe Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipCreateRegionRectI Lib "gdiplus" (Rect As Rect, Region As LongPtr) As Long
Private Declare PtrSafe Function GdipTransformMatrixPointsS Lib "gdiplus" Alias "GdipTransformMatrixPoints" (ByVal matrix As LongPtr, pts As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipVectorTransformMatrixPointsS Lib "gdiplus" Alias "GdipVectorTransformMatrixPoints" (ByVal matrix As LongPtr, pts As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipGetRegionBoundsS Lib "gdiplus" Alias "GdipGetRegionBounds" (ByVal Region As LongPtr, ByVal Graphics As LongPtr, Rect As RECTF) As Long
Private Declare PtrSafe Function GdipGetPathWorldBoundsS Lib "gdiplus" Alias "GdipGetPathWorldBounds" (ByVal Path As LongPtr, bounds As RECTF, ByVal matrix As LongPtr, ByVal pen As LongPtr) As Long
Private Declare PtrSafe Function GdipFlattenPath Lib "gdiplus" (ByVal Path As LongPtr, Optional ByVal matrix As LongPtr = 0, Optional ByVal flatness As Single = 0.25) As Long
Private Declare PtrSafe Function GdipWindingModeOutline Lib "gdiplus" (ByVal Path As LongPtr, Optional ByVal matrix As LongPtr = 0, Optional ByVal flatness As Single = 0.25) As Long
Private Declare PtrSafe Function GdipWidenPath Lib "gdiplus" (ByVal Path As LongPtr, Optional ByVal pen As LongPtr = 0, Optional ByVal matrix As LongPtr = 0, Optional ByVal flatness As Single = 0.25) As Long
Private Declare PtrSafe Function GdipGetPointCount Lib "gdiplus" (ByVal Path As LongPtr, count As Long) As Long
Private Declare PtrSafe Function GdipGetPathPoints Lib "gdiplus" (ByVal Path As LongPtr, Points As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipGetPathPointsI Lib "gdiplus" (ByVal Path As LongPtr, Points As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipGetPathTypes Lib "gdiplus" (ByVal Path As LongPtr, types As Any, ByVal count As Long) As Long
#If UseSingle Then
Private Declare PtrSafe Function GdipDrawImageRectRect Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Image As LongPtr, ByVal dstx As Single, _
                                                               ByVal dsty As Single, ByVal dstwidth As Single, ByVal dstheight As Single, _
                                                               ByVal srcx As Single, ByVal srcy As Single, ByVal srcwidth As Single, ByVal srcheight As Single, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As LongPtr = PtrNull, _
                                                               Optional ByVal CallBack As LongPtr = PtrNull, Optional ByVal callbackData As LongPtr = PtrNull) As Long
Private Declare PtrSafe Function GdipAddPathLine Lib "gdiplus" (ByVal Path As LongPtr, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare PtrSafe Function GdipCreateLineBrushFromRect Lib "gdiplus" (Rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal WrapMd As Long, lineGradient As LongPtr) As Long
Private Declare PtrSafe Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal Path As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal pen As LongPtr, ByVal Graphics As LongPtr, result As Long) As Long
Private Declare PtrSafe Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal Path As LongPtr, bounds As RECTF, ByVal matrix As LongPtr, ByVal pen As LongPtr) As Long
Private Declare PtrSafe Function GdipGetRegionBounds Lib "gdiplus" (ByVal Region As LongPtr, ByVal Graphics As LongPtr, Rect As RECTF) As Long
Private Declare PtrSafe Function GdipGetPathGradientCenterPoint Lib "gdiplus" (ByVal Brush As LongPtr, Points As POINTF) As Long
Private Declare PtrSafe Function GdipAddPathRectangle Lib "gdiplus" (ByVal Path As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare PtrSafe Function GdipAddPathArc Lib "gdiplus" (ByVal Path As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare PtrSafe Function GdipAddPathPolygon Lib "gdiplus" (ByVal Path As LongPtr, Points As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipAddPathBeziers Lib "gdiplus" (ByVal Path As LongPtr, Points As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipAddPathCurve2 Lib "gdiplus" (ByVal Path As LongPtr, Points As POINTF, ByVal count As Long, ByVal tension As Single) As Long
Private Declare PtrSafe Function GdipAddPathLine2 Lib "gdiplus" (ByVal Path As LongPtr, Points As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipAddPathEllipse Lib "gdiplus" (ByVal Path As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare PtrSafe Function GdipFillRectangle Lib "gdiplus" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare PtrSafe Function GdipCreateRegionRect Lib "gdiplus" (Rect As RECTF, Region As LongPtr) As Long
Private Declare PtrSafe Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal matrix As LongPtr, pts As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipVectorTransformMatrixPoints Lib "gdiplus" (ByVal matrix As LongPtr, pts As POINTF, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single) As Long
#Else
Private Declare PtrSafe Function GdipDrawImageRectRect Lib "gdiplus" Alias "GdipDrawImageRectRectI" (ByVal Graphics As LongPtr, ByVal Image As LongPtr, ByVal dstx As Long, _
                                                               ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, _
                                                               ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As LongPtr = PtrNull, _
                                                               Optional ByVal CallBack As LongPtr = PtrNull, Optional ByVal callbackData As LongPtr = PtrNull) As Long
Private Declare PtrSafe Function GdipAddPathLine Lib "gdiplus" Alias "GdipAddPathLineI" (ByVal Path As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function GdipCreateLineBrushFromRect Lib "gdiplus" Alias "GdipCreateLineBrushFromRectI" (Rect As Rect, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal WrapMd As Long, lineGradient As LongPtr) As Long
Private Declare PtrSafe Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" Alias "GdipIsOutlineVisiblePathPointI" (ByVal Path As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal pen As LongPtr, ByVal Graphics As LongPtr, result As Long) As Long
Private Declare PtrSafe Function GdipGetPathWorldBounds Lib "gdiplus" Alias "GdipGetPathWorldBoundsI" (ByVal Path As LongPtr, bounds As Rect, ByVal matrix As LongPtr, ByVal pen As LongPtr) As Long
Private Declare PtrSafe Function GdipGetRegionBounds Lib "gdiplus" Alias "GdipGetRegionBoundsI" (ByVal Region As LongPtr, ByVal Graphics As LongPtr, Rect As Rect) As Long
Private Declare PtrSafe Function GdipGetPathGradientCenterPoint Lib "gdiplus" Alias "GdipGetPathGradientCenterPointI" (ByVal Brush As LongPtr, Points As POINTAPI) As Long
Private Declare PtrSafe Function GdipAddPathRectangle Lib "gdiplus" Alias "GdipAddPathRectangleI" (ByVal Path As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipAddPathArc Lib "gdiplus" Alias "GdipAddPathArcI" (ByVal Path As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare PtrSafe Function GdipAddPathPolygon Lib "gdiplus" Alias "GdipAddPathPolygonI" (ByVal Path As LongPtr, Points As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipAddPathBeziers Lib "gdiplus" Alias "GdipAddPathBeziersI" (ByVal Path As LongPtr, Points As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipAddPathCurve2 Lib "gdiplus" Alias "GdipAddPathCurve2I" (ByVal Path As LongPtr, Points As POINTAPI, ByVal count As Long, ByVal tension As Single) As Long
Private Declare PtrSafe Function GdipAddPathLine2 Lib "gdiplus" Alias "GdipAddPathLine2I" (ByVal Path As LongPtr, Points As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipAddPathEllipse Lib "gdiplus" Alias "GdipAddPathEllipseI" (ByVal Path As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipFillRectangle Lib "gdiplus" Alias "GdipFillRectangleI" (ByVal Graphics As LongPtr, ByVal Brush As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipCreateRegionRect Lib "gdiplus" Alias "GdipCreateRegionRectI" (Rect As Rect, Region As LongPtr) As Long
Private Declare PtrSafe Function GdipTransformMatrixPoints Lib "gdiplus" Alias "GdipTransformMatrixPointsI" (ByVal matrix As LongPtr, pts As POINTAPI, ByVal count As Long) As Long
Private Declare PtrSafe Function GdipVectorTransformMatrixPoints Lib "gdiplus" Alias "GdipVectorTransformMatrixPointsI" (ByVal matrix As LongPtr, pts As POINTAPI, ByVal count As Long) As Long
#End If
#Else
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean
Private Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Private Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHGetFileInfo Lib "SHELL32.DLL" Alias "SHGetFileInfoA" (ByVal pPath As String, _
    ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ValidateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As Rect) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
                                           (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, _
                                                                ByRef ppstm As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As Object, phglobal As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetThemeAppProperties Lib "uxtheme" (ByVal dwFlags As Long) As Long
Private Declare Function SetStretchBltMode Lib "Gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetDIBits Lib "Gdi32" _
                                   (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
                                    ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function CreatePatternBrush Lib "Gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateDIBPatternBrush Lib "Gdi32" (ByVal hPackedDIB As Long, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "Gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function SetBrushOrgEx Lib "Gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
Private Declare Function GetObjectBmp Lib "Gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectClipRgn Lib "Gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" _
                        (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function StringFromCLSID Lib "ole32.dll" (pclsid As GUID, lpszProgID As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetWorldTransform Lib "Gdi32" (ByVal hDC As Long, lpXform As XForm) As Long
Private Declare Function CreateDIBSection Lib "Gdi32" (ByVal hDC As Long, pBitmapInfo As BitmapInfo, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetGraphicsMode Lib "Gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "Gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetMapMode Lib "Gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function BitBlt Lib "Gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "Gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "Gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentObject Lib "Gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function CreateCompatibleDC Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DPtoLP Lib "Gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function LPtoDP Lib "Gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SetEnhMetaFileBits Lib "Gdi32" (ByVal cbBuffer As Long, lpData As Byte) As Long
Private Declare Function DeleteEnhMetaFile Lib "Gdi32" (ByVal hemf As Long) As Long
Private Declare Function GetEnhMetaFileBits Lib "Gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Any) As Long
Private Declare Function GetEnhMetaFileHeader Lib "Gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpEMH As ENHMETAHEADER) As Long
Private Declare Function PlayEnhMetaFile Lib "Gdi32" (ByVal hDC As Long, ByVal hemf As Long, lpRect As Rect) As Long
Private Declare Function FrameRgn Lib "Gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetBkMode Lib "Gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal fontCollection As Long, ByVal memory As Long, ByVal length As Long) As Long
Private Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (fontCollection As Long) As Long
Private Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal family As Long, ByVal Name As Long, ByVal language As Integer) As Long
Private Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal fontCollection As Long, numFound As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpfamilies As Long, numFound As Long) As Long
Private Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (fontCollection As Long) As Long
Private Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal fontCollection As Long, ByVal FileName As Long) As Long
Private Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (fontCollection As Long) As Long
Private Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal Brush As Long) As Long
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As Long, texture As Long) As Long
Private Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As Long) As Long
Private Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As Long) As Long
Private Declare Function GdipGetTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal matrix As Long) As Long
Private Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal matrix As Long) As Long
Private Declare Function GdipMultiplyTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal matrix As Long, ByVal order As Long) As Long
Private Declare Function GdipScaleTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sx As Single, ByVal sY As Single, ByVal order As Long) As Long
Private Declare Function GdipRotateTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal angle As Single, ByVal order As Long) As Long
Private Declare Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare Function GdipResetTextureTransform Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipGetTextureImage Lib "gdiplus" (ByVal Brush As Long, Image As Long) As Long
Private Declare Function GdipGetCompositingMode Lib "gdiplus" (ByVal Graphics As Long, CompositingMd As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal CompositingMd As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixOffsetMode As Long) As Long
Private Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, PixOffsetMode As Long) As Long
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal Graphics As Long, ByVal CompositingQlty As Long) As Long
Private Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal Graphics As Long, CompositingQlty As Long) As Long
Private Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal Graphics As Long, ByVal Region As Long, ByVal CombineMd As Long) As Long
Private Declare Function GdipCreateRegion Lib "gdiplus" (Region As Long) As Long
Private Declare Function GdipSetEmpty Lib "gdiplus" (ByVal Region As Long) As Long
Private Declare Function GdipSetInfinite Lib "gdiplus" (ByVal Region As Long) As Long
Private Declare Function GdipResetClip Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal pen As Long, ByVal penMode As Long) As Long
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal Region As Long, Rect As Rect, ByVal CombineMd As Long) As Long
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal Region As Long, ByVal region2 As Long, ByVal CombineMd As Long) As Long
Private Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, hRgn As Long) As Long
Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, result As Long) As Long
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal Region As Long) As Long
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal Region As Long, ByVal path As Long, ByVal CombineMd As Long) As Long
Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal path As Long, Region As Long) As Long
Private Declare Function GdipCreateMatrix Lib "gdiplus" (matrix As Long) As Long
Private Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single, matrix As Long) As Long
Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal matrix As Long) As Long
Private Declare Function GdipGetMatrixElements Lib "gdiplus" (ByVal matrix As Long, matrixOut As Single) As Long
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal matrix As Long, cloneMatrix As Long) As Long
Private Declare Function GdipMultiplyMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal matrix2 As Long, ByVal order As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal order As Long) As Long
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal angle As Single, ByVal order As Long) As Long
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal order As Long) As Long
Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal order As Long) As Long
Private Declare Function GdipIsMatrixIdentity Lib "gdiplus" (ByVal matrix As Long, result As Long) As Long
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal matrix As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal path As Long) As Long
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal path As Long, clonePath As Long) As Long
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal Region As Long, cloneRegion As Long) As Long
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal path As Long, ByVal matrix As Long) As Long
Private Declare Function GdipTransformRegion Lib "gdiplus" (ByVal Region As Long, ByVal matrix As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal path As Long) As Long
Private Declare Function GdipFillRegion Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal Region As Long) As Long
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal Region As Long, ByVal X As Single, ByVal Y As Single, ByVal Graphics As Long, result As Long) As Long
Private Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal Region As Long, ByVal X As Single, ByVal Y As Single, ByVal Graphics As Long, result As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As Long, path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal path As Long, polyGradient As Long) As Long
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal path As Long, ByVal addingPath As Long, ByVal bConnect As Long) As Long
Private Declare Function GdipAddPathString Lib "gdiplus" (ByVal path As Long, ByVal str As Long, ByVal length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, layoutRect As RECTF, ByVal StringFormat As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As POINTAPI, ByVal count As Long, ByVal FillMd As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Graphics As Long, ByVal lColor As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As Long, Format As GUID) As Long
Private Declare Function GdipRecordMetafileI Lib "gdiplus" (ByVal referenceHdc As Long, ByVal etype As Long, frameRect As Rect, ByVal frameUnit As Long, ByVal description As Long, metafile As Long) As Long
Private Declare Function GdipRecordMetafileFileNameI Lib "gdiplus" (ByVal FileName As Long, ByVal referenceHdc As Long, ByVal etype As Long, frameRect As Rect, ByVal frameUnit As Long, ByVal description As Long, metafile As Long) As Long
Private Declare Function GdipGetDC Lib "gdiplus" (ByVal Graphics As Long, hDC As Long) As Long
Private Declare Function GdipReleaseDC Lib "gdiplus" (ByVal Graphics As Long, ByVal hDC As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, bitmap As Long) As Long
Private Declare Function GdipSetImagePalette Lib "gdiplus" (ByVal Image As Long, palette As ColorPalette) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal count As Long) As Long
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As Long, ByVal forecolr As Long, ByVal backcolr As Long, Brush As Long) As Long
Private Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal bitmap As Long, ByVal xdpi As Single, ByVal ydpi As Single) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As Long
Private Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Graphics As Long, bitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BitmapInfo, ByVal gdiBitmapData As Long, bitmap As Long) As Long
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, PixelFormat As Long) As Long
Private Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As Long, totalBufferSize As Long, numProperties As Long) As Long
Private Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As Long, ByVal totalBufferSize As Long, ByVal numProperties As Long, allItems As Any) As Long
Private Declare Function GdipRemovePropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long) As Long
Private Declare Function GdipAlloc Lib "gdiplus.dll" (ByVal Size As Long) As Long
Private Declare Function GdipFree Lib "gdiplus.dll" (ByVal Ptr As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal stream As Object, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef bitmap As Long) As Long
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, _
                                                            ByVal propSize As Long, ByRef buffer As Any) As Long
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, _
                                                                ByRef Size As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, LInput As GdiplusStartupInput, Optional ByVal lOutPut As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, ByRef hbmReturn As Long, _
                                                                    ByVal Background As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" _
                                                    (ByVal Image As Long, ByRef dimensionID As GUID, _
                                                     ByVal frameIndex As Long) As Long
Private Declare Function GdipImageGetFrameCount Lib "gdiplus" _
                                                (ByVal Image As Long, ByRef dimensionID As GUID, _
                                                 ByRef count As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" _
                                               (ByVal Image As Long, ByRef Width As Single, _
                                                ByRef Height As Single) As Long
Private Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal nImage As Long, item As PropertyItem) As Long
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal Image As Long, resolution As Single) As Long
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal Image As Long, resolution As Single) As Long
Private Declare Function GdipGetPropertyCount Lib "gdiplus" (ByVal Image As Long, numOfProperty As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, Rect As Rect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipGetWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal matrix As Long) As Long
Private Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal matrix As Long) As Long
Private Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal sx As Single, ByVal sY As Single, ByVal order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal angle As Single, ByVal order As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As Long) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal Graphics As Long, ByVal str As Long, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, ByVal mode As Long) As Long
Private Declare Function GdipGetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, mode As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal str As Long, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal Brush As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As Long, ByVal unit As Long, createdfont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (nativeFamily As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As Long, ByVal enableFlag As Long, ByVal colorLow As Long, ByVal colorHigh As Long) As Long
Private Declare Function GdipSetImageAttributesRemapTable Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As Long, ByVal enableFlag As Long, ByVal mapSize As Long, map As ColorMap) As Long
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageattr As Long, ByVal wrap As Long, ByVal argb As Long, ByVal bClamp As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, interpolation As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal interpolation As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, SmoothingMd As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal startCap As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal endCap As Long) As Long
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As Long, ByVal LnJoin As Long) As Long
Private Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal pen As Long, ByVal miterLimit As Single) As Long
Private Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal pen As Long, miterLimit As Single) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, Graphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long
Private Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As Long, ByVal enableFlag As Long, colourMatrix As ColorMatrix, grayMatrix As Any, ByVal flags As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal metafile As Long, hemf As Long) As Long
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal path As Long, ByVal brushmode As Long) As Long
Private Declare Function GdipCreateAdjustableArrowCap Lib "gdiplus" (ByVal Height As Single, ByVal Width As Single, ByVal isFilled As Long, cap As Long) As Long
Private Declare Function GdipSetPenCustomStartCap Lib "gdiplus" (ByVal pen As Long, ByVal customCap As Long) As Long
Private Declare Function GdipSetPenCustomEndCap Lib "gdiplus" (ByVal pen As Long, ByVal customCap As Long) As Long
Private Declare Function GdipDeleteCustomLineCap Lib "gdiplus" (ByVal customCap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal dstx As Long, _
                                                               ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, _
                                                               ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = PtrNull, _
                                                               Optional ByVal callback As Long = PtrNull, Optional ByVal callbackData As Long = PtrNull) As Long
Private Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipAddPathEllipseS Lib "gdiplus" Alias "GdipAddPathEllipse" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipCreateRegionRectI Lib "gdiplus" (Rect As Rect, Region As Long) As Long
Private Declare Function GdipTransformMatrixPointsS Lib "gdiplus" Alias "GdipTransformMatrixPoints" (ByVal matrix As Long, pts As POINTF, ByVal count As Long) As Long
Private Declare Function GdipVectorTransformMatrixPointsS Lib "gdiplus" Alias "GdipVectorTransformMatrixPoints" (ByVal matrix As Long, pts As POINTF, ByVal count As Long) As Long
Private Declare Function GdipGetRegionBoundsS Lib "gdiplus" Alias "GdipGetRegionBounds" (ByVal Region As Long, ByVal Graphics As Long, Rect As RECTF) As Long
Private Declare Function GdipGetPathWorldBoundsS Lib "gdiplus" Alias "GdipGetPathWorldBounds" (ByVal path As Long, bounds As RECTF, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipFlattenPath Lib "gdiplus" (ByVal path As Long, Optional ByVal matrix As Long = 0, Optional ByVal flatness As Single = 0.25) As Long
Private Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal path As Long, Optional ByVal matrix As Long = 0, Optional ByVal flatness As Single = 0.25) As Long
Private Declare Function GdipGetPointCount Lib "gdiplus" (ByVal path As Long, count As Long) As Long
Private Declare Function GdipGetPathPoints Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As Long
Private Declare Function GdipGetPathPointsI Lib "gdiplus" (ByVal path As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal path As Long, Optional ByVal pen As Long = 0, Optional ByVal matrix As Long = 0, Optional ByVal flatness As Single = 0.25) As Long
Private Declare Function GdipGetPathTypes Lib "gdiplus" (ByVal path As Long, types As Any, ByVal count As Long) As Long
#If UseSingle Then
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal dstx As Single, _
                                                               ByVal dsty As Single, ByVal dstwidth As Single, ByVal dstheight As Single, _
                                                               ByVal srcx As Single, ByVal srcy As Single, ByVal srcwidth As Single, ByVal srcheight As Single, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = PtrNull, _
                                                               Optional ByVal callback As Long = PtrNull, Optional ByVal callbackData As Long = PtrNull) As Long
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal path As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (Rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal pen As Long, ByVal Graphics As Long, result As Long) As Long
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal path As Long, bounds As RECTF, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, Rect As RECTF) As Long
Private Declare Function GdipGetPathGradientCenterPoint Lib "gdiplus" (ByVal Brush As Long, Points As POINTF) As Long
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As Long
Private Declare Function GdipAddPathBeziers Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As Long
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long, ByVal tension As Single) As Long
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipCreateRegionRect Lib "gdiplus" (Rect As RECTF, Region As Long) As Long
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal matrix As Long, pts As POINTF, ByVal count As Long) As Long
Private Declare Function GdipVectorTransformMatrixPoints Lib "gdiplus" (ByVal matrix As Long, pts As POINTF, ByVal count As Long) As Long
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single) As Long
#Else
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" Alias "GdipDrawImageRectRectI" (ByVal Graphics As Long, ByVal Image As Long, ByVal dstx As Long, _
                                                               ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, _
                                                               ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = PtrNull, _
                                                               Optional ByVal callback As Long = PtrNull, Optional ByVal callbackData As Long = PtrNull) As Long
Private Declare Function GdipAddPathLine Lib "gdiplus" Alias "GdipAddPathLineI" (ByVal path As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" Alias "GdipCreateLineBrushFromRectI" (Rect As Rect, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" Alias "GdipIsOutlineVisiblePathPointI" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal pen As Long, ByVal Graphics As Long, result As Long) As Long
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" Alias "GdipGetPathWorldBoundsI" (ByVal path As Long, bounds As Rect, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipGetRegionBounds Lib "gdiplus" Alias "GdipGetRegionBoundsI" (ByVal Region As Long, ByVal Graphics As Long, Rect As Rect) As Long
Private Declare Function GdipGetPathGradientCenterPoint Lib "gdiplus" Alias "GdipGetPathGradientCenterPointI" (ByVal Brush As Long, Points As POINTAPI) As Long
Private Declare Function GdipAddPathRectangle Lib "gdiplus" Alias "GdipAddPathRectangleI" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" Alias "GdipAddPathArcI" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathPolygon Lib "gdiplus" Alias "GdipAddPathPolygonI" (ByVal path As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipAddPathBeziers Lib "gdiplus" Alias "GdipAddPathBeziersI" (ByVal path As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" Alias "GdipAddPathCurve2I" (ByVal path As Long, Points As POINTAPI, ByVal count As Long, ByVal tension As Single) As Long
Private Declare Function GdipAddPathLine2 Lib "gdiplus" Alias "GdipAddPathLine2I" (ByVal path As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" Alias "GdipAddPathEllipseI" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipFillRectangle Lib "gdiplus" Alias "GdipFillRectangleI" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipCreateRegionRect Lib "gdiplus" Alias "GdipCreateRegionRectI" (Rect As Rect, Region As Long) As Long
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" Alias "GdipTransformMatrixPointsI" (ByVal matrix As Long, pts As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipVectorTransformMatrixPoints Lib "gdiplus" Alias "GdipVectorTransformMatrixPointsI" (ByVal matrix As Long, pts As POINTAPI, ByVal count As Long) As Long
#End If
#End If

'***************************************************************************************
'*                                    Constantes                                       *
'***************************************************************************************
Private Const TIME_PERIODIC = 1
Private Const TIME_CALLBACK_EVENT_SET = &H10
Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&
Private Const INFINITE = &HFFFF
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const SM_CXVSCROLL = 2
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46
Private Const GM_ADVANCED As Long = 2
Private Const SM_CXBORDER = 5
Private Const OBJ_BITMAP As Long = 7
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B&
Private Const HIMETRIC_INCH = 2540          ' Pour conversion Pouce<->Himetric
Private Const MM_HIMETRIC = 3
Private Const STAP_ALLOW_CONTROLS = &H2
Private Const STAP_ALLOW_NONCLIENT = &H1
Private Const STAP_ALLOW_WEBCONTENT = &H4
Private Const GMEM_MOVEABLE = &H2&
Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086
Private Const HALFTONE = 4                  ' Mode pour StretchBlt avec antialiasing
Private Const CF_ENHMETAFILE = 14
Private Const CF_BITMAP = 40
Private Const LOGPIXELSY = 90
Private Const LOGPIXELSX = 88
Private Const DIB_RGB_COLORS As Long = &H0
Private Const BI_RGB As Long = &H0
Private Const PropertyTagFrameDelay As Long = &H5100&
Private Const PropertyTagTypeByte = 1
Private Const PropertyTagTypeASCII = 2
Private Const PropertyTagTypeShort = 3
Private Const PropertyTagTypeLong = 4
Private Const PropertyTagTypeRational = 5
Private Const PropertyTagTypeUndefined = 7
Private Const PropertyTagTypeSLong = 9
Private Const PropertyTagTypeSRational = 10
Private Const WS_EX_COMPOSITED = &H2000000
Private Const GWL_EXSTYLE = &HFFEC
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const MM_TEXT = 1
Private Const SHGFI_ICON = &H100
' Scrollbar
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const SB_CTL = 2
Private Const SB_VERT = 2
Private Const SB_HORZ = 1
' GUID
Private Const gJpegEncoderStr As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Private Const gGifEncoderStr As String = "{557cf402-1a04-11d3-9a73-0000f81ef32e}"
Private Const gBmpEncoderStr As String = "{557cf400-1a04-11d3-9a73-0000f81ef32e}"
Private Const gPngEncoderStr As String = "{557cf406-1a04-11d3-9a73-0000f81ef32e}"
Private Const gTifEncoderStr As String = "{557cf405-1a04-11d3-9a73-0000f81ef32e}"
Private Const gQualityParamStr As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const gEncoderTransformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const WM_NCLBUTTONDOWN = &HA1       ' Click hors du client de la fenetre (barre de titre)
Private Const HTCAPTION = 2                 ' Identifiant : Barre de titre
Private Const cLongOptional As Long = 2147483647
#If UseSingle Then
Private Const cSingleOptional As Single = 3.402823E+38!
#Else
Private Const cSingleOptional As Long = 2147483647
#End If
' DrawFrameControl
Private Const DFC_CAPTION = 1
Private Const DFC_MENU = 2
Private Const DFC_SCROLL = 3
Private Const DFC_BUTTON = 4
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONHELP = &H4
Private Const DFCS_MENUARROW = &H0
Private Const DFCS_MENUCHECK = &H1
Private Const DFCS_MENUBULLET = &H2
Private Const DFCS_MENUARROWRIGHT = &H4
Private Const DFCS_SCROLLUP = &H0
Private Const DFCS_SCROLLDOWN = &H1
Private Const DFCS_SCROLLLEFT = &H2
Private Const DFCS_SCROLLRIGHT = &H3
Private Const DFCS_SCROLLCOMBOBOX = &H5
Private Const DFCS_SCROLLSIZEGRIP = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Private Const DFCS_BUTTONCHECK = &H0
Private Const DFCS_BUTTONRADIOIMAGE = &H1
Private Const DFCS_BUTTONRADIOMASK = &H2
Private Const DFCS_BUTTONRADIO = &H4
Private Const DFCS_BUTTON3STATE = &H8
Private Const DFCS_BUTTONPUSH = &H10
Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const DFCS_CHECKED = &H400
Private Const DFCS_ADJUSTRECT = &H2000
Private Const DFCS_FLAT = &H4000
Private Const DFCS_MONO = &H8000&
' DrawEdge
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
' Constante pour conversion de taille de police
#If AppName = "A" Then
    Private Const cPoints As Long = 20
#Else
    Private Const cPoints As Long = 1
#End If
'***************************************************************************************
'*                                       Types                                         *
'***************************************************************************************
Private Type SECURITY_ATTRIBUTES
    nLength As Long
#If VBA7 Then
    lpSecurityDescriptor As LongPtr
#Else
    lpSecurityDescriptor As Long
#End If
    bInheritHandle As Long
End Type
Private Type XForm
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type
' Pour region auto
Private Type TRgnPoints
    o As Boolean
    T As Boolean
End Type
' Info scrollbar
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Type ColorMap
    oldColor As Long
    newColor As Long
End Type
Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type
' Rectangle pour API
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Point pour API
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type POINTF
    X As Single
    Y As Single
End Type
' Rectangle pour API (single)
Private Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type
' Pour lire les fichiers integres type EMF (enhanced metafile)
Private Type SIZEL
    cx As Long
    cy As Long
End Type
Private Type ENHMETAHEADER
    iType As Long
    nSize As Long
    rclBounds As Rect
    rclFrame As Rect
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As SIZEL
    szlMillimeters As SIZEL
    cbPixelFormat As Long
    offPixelFormat As Long
    bOpenGL As Long
End Type
Private Type ColorPalette
    flags As Long
    count As Long
    Entries(0 To 255) As Long
End Type
Private Type argb
    blue As Byte    ' Bleu
    green As Byte    ' Vert
    red As Byte    ' Rouge
    Alpha As Byte    ' Luminosite
End Type
Private Type BitmapInfoHeader
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
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
#If VBA7 Then
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As LongPtr
    SuppressExternalCodecs As LongPtr
End Type
Private Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    scan0 As LongPtr
    Reserved As Long
End Type
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As LongPtr
    hbmColor As LongPtr
End Type
Private Type SHFILEINFO
   hIcon As LongPtr
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * 260
   szTypeName As String * 80
End Type
Private Type PicBmp
    size As Long
    tType As Long
    hBmp As LongPtr
    hpal As LongPtr
    Reserved As Long
End Type
Private Type bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
End Type
Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
Type As Long
    value As LongPtr
End Type
Private Type PropertyItem
    id As Long
    Length As Long
Type As Integer
    value As LongPtr
End Type
#Else
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    scan0 As Long
    Reserved As Long
End Type
' Information icone
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * 260
   szTypeName As String * 80
End Type
Private Type PicBmp
    size As Long
    tType As Long
    hBmp As Long
    hpal As Long
    Reserved As Long
End Type
Private Type bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    Type As Long
#If VBA7 Then
    value As LongPtr
#Else
    value As Long
#End If
End Type
Private Type PropertyItem
    id As Long
    Length As Long
    Type As Integer
    value As Long
End Type
#End If
Private Type EncoderParameters
    count As Long
    Parameter(0 To 15) As EncoderParameter
End Type
' Type pour info image DIB
Private Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors(0 To 255) As Long
End Type
Private Type tDIB
#If VBA7 Then
    hDC As LongPtr
    hDIB As LongPtr
    DIBPTR As LongPtr
    hOldDIB As LongPtr
    hBrush As LongPtr
#Else
    hDC As Long
    hDIB As Long
    DIBPTR As Long
    hOldDIB As Long
    hBrush As Long
#End If
    BI As BitmapInfo
End Type
'***************************************************************************************
'*                                    Variables                                        *
'***************************************************************************************
Private gIsV11 As Boolean ' Gdi+ 1.1 ?
Private gBitmap     ' Bitmap de travail (image principale)
Private gGraphics     ' Graphique de travail (image principale)
Private gGdipToken     ' Identifiant pour GdiPlus
Private gLib                ' Identifiant de la librairie chargee en memoire
Private gImgList As Collection    ' Collection d'images
Private gTextures As Collection    ' Collection de textures
Private gCollRegions As Collection    ' Collection de regions
Private gCurseurs As Collection    ' Collection de curseurs
' 0 : Handle de la region
' 1 : Identifiant de la region (son nom)
Private gMatrix As Collection  ' Collection de matrices pour matrix* et world*
Private gUserData As Collection ' Collection de donnees utilisateur
Private gCollSave As Collection    ' Collection d'images sauvegardees en memoire (avec Keepimage)
#If UseSingle Then
Private gRepaintCorrection(1 To 8) As Single    ' Corrections pour RepaintFast
#Else
Private gRepaintCorrection(1 To 8) As Long    ' Corrections pour RepaintFast
#End If
Private gClipRegion   ' Clip pour RepaintFast (n'est dessine que l'interieur ou l'exterieur de la region)
Private gFormRegionHwnd As Collection    ' Collection de handle de formulaire a retablir a la fermeture de la classe
Private gPI As Single    ' Valeur de PI
Private gCtrlRef As Object    ' Controle de reference pour coordonnees
Private gCtrlRefToImgMatrix ' Matrice de passage vers gCtrlRef
Private gCancelPaint As Boolean    ' Annule le dessin par windows
Private gActiveFrame As Long    ' Image active d'un gif
Private gPrivateFontCollection     ' Collection de polices de caracteres privees
Private gDrawClipRegion ' Clip pour dessin sur image principale
Private gMatrixForImageReset ' Matrice sauvegardée pour ImageReset
Private gDashUser() As Single ' pattern de pointilles
Private gUseGdi As Boolean
Private gUseGdiWindow As Boolean
Private gUseGdiImg As Boolean
Private gDib As tDIB
Private gId As Long    ' Id du timer
Private ghEvent    ' Id de l'evenement
Private gInterval As Long    ' Intervalle de l'evenement
Private gTransparencyMode As Long ' Mode transparence (over ou copy)
Private gPixelOffsetMode As Long ' Decalage de pixels
Private gSmoothingMode As Long ' Lissage lignes
Private gAntialisingLevel As Long ' Lissage images
Private gTextRenderingHint As Long ' Lissage textes
Private gCompositingQuality As Long ' Qualité transparence
Private gFreq As Currency
Private gTimer As Currency
Private gTimerStart As Currency
Private gOverhead As Currency

#If UseSingle Then
Public LastTextBottom As Single ' Derniere position basse du texte ecrit
Public LastTextRight As Single ' Derniere position droite du texte ecrit (si aligne a gauche)
Public LastTextSize As Single ' Derniere taille du texte ecrit
Public MaxTextSize As Single ' Taille maxi du texte
Public MinTextSize As Single ' Taille mini du texte
Public WrapX As Single, WrapY As Single    ' Enroulement de l'image
Public DrawX1 As Single, DrawY1 As Single, DrawX2 As Single, DrawY2 As Single ' Position du dessin
Public DrawOrigX1 As Single, DrawOrigY1 As Single, DrawOrigX2 As Single, DrawOrigY2 As Single ' Portion du dessin
#Else
Public LastTextBottom As Long ' Derniere position basse du texte ecrit
Public LastTextRight As Long ' Derniere position droite du texte ecrit (si aligne a gauche)
Public LastTextSize As Long ' Derniere taille du texte ecrit
Public MaxTextSize As Long ' Taille maxi du texte
Public MinTextSize As Long ' Taille mini du texte
Public WrapX As Long, WrapY As Long    ' Enroulement de l'image
Public DrawX1 As Long, DrawY1 As Long, DrawX2 As Long, DrawY2 As Long ' Position du dessin
Public DrawOrigX1 As Long, DrawOrigY1 As Long, DrawOrigX2 As Long, DrawOrigY2 As Long ' Portion du dessin
#End If
Public IconBackColor As Long ' Couleur de fond des icones
Public FillTexture As String    ' Texture utilisee pour le remplissage
Public PenTexture As String    ' Texture utilisee pour le crayon
Public RepaintFastTransparency As Boolean    ' Mode transparent pour RepaintFast
Public ImgName As String ' Nom de l'image
Public ConvertMetaFile As Boolean ' Conversion des metafile au chargement
#If VBA6 Then
Public PenAlignMode As EGdipPenAlignMode    ' Alignement du crayon
Public FillMode As EGdipFillMode ' Mode de remplissage
Public LineJoin As EGdipLineJoin ' Type de jointure entre les lignes
Public LineStart As EGdipLineCap ' Type de depart de ligne
Public LineEnd As EGdipLineCap ' Type de fin de ligne
#Else
Public PenAlignMode As Long ' Alignement crayon
Public FillMode As Long ' Mode remplissage
Public LineJoin As Long ' Jointure de lignes
Public LineStart As Long ' Debut de lignes
Public LineEnd As Long ' Fin de lignes
#End If
Public DrawPathClose As Boolean
Public LineArrowLength As Single ' Taille de la fleche
Public LineArrowFill As Boolean ' Remplissage de la fleche
Public LineJoinMiterLimit As Single ' Limite pour jointure de ligne
Public LoadAndClone As Boolean ' Chargement et clonage de l'image (libere le fichier)
Public DrawMatrix As Variant ' Matrice de transformation vers l'image
Public RepaintMatrix As Variant ' Matrice de transformation du contrôle vers l'image principale
Public RepaintMatrixUpdate As Boolean ' Mise a jour de RepaintMatrix
Public RegionAppend As Boolean ' Si Vrai, la region est ajoutee a la region existante
Public TextCreateRegionOutline As Boolean ' Region pour texte = entourage
Public RepaintFastResetCalc As Boolean ' Reset calcul de positionnement de l'image si modif ctrl ou taille img
Public ImgParentPtr ' Image parente => ne pas utilise
Public ImgCopyRendering As Boolean ' copie des modes du parent
Public RegionWorldTransform As Boolean ' Transformation des régions à la création si WorldTransform
'***************************************************************************************
'*                                Enumerations                            *
'***************************************************************************************
#If VBA6 Then
Public Enum EGdipPixelFormat
     PixelFormat1bppIndexed = &H30101
     PixelFormat4bppIndexed = &H30402
     PixelFormat8bppIndexed = &H30803
     PixelFormat16bppGreyScale = &H101004
     PixelFormat16bppRGB555 = &H21005
     PixelFormat16bppRGB565 = &H21006
     PixelFormat16bppARGB1555 = &H61007
     PixelFormat24bppRGB = &H21808
     PixelFormat32bppRGB = &H22009
     PixelFormat32bppARGB = &H26200A
     PixelFormat32bppPARGB = &HE200B
     PixelFormat48bppRGB = &H10300C
     PixelFormat64bppARGB = &H34400D
     PixelFormat64bppPARGB = &H1C400E
End Enum
Public Enum EGdipCursors
    gdipIDC_APPSTARTING = 32650& ' Standard arrow and small hourglass cursor.
    gdipIDC_ARROW = 32512&   ' Standard arrow cursor.
    gdipIDC_CROSS = 32515&   ' Crosshair cursor.
    gdipIDC_HAND = 32649& ' Hand cursor.
    gdipIDC_HELP = 32651& ' Arrow and question mark cursor.
    gdipIDC_IBEAM = 32513&   ' I-beam cursor.
    gdipIDC_NO = 32648&  ' Slashed circle cursor.
    gdipIDC_SIZEALL = 32646& ' Four-pointed arrow cursor pointing north, south, east, and west.
    gdipIDC_SIZENESW = 32643& ' Double-pointed arrow cursor pointing northeast and southwest.
    gdipIDC_SIZENS = 32645&  ' Double-pointed arrow cursor pointing north and south.
    gdipIDC_SIZENWSE = 32642& ' Double-pointed arrow cursor pointing northwest and southeast.
    gdipIDC_SIZEWE = 32644&  ' Double-pointed arrow cursor pointing west and east.
    gdipIDC_UPARROW = 32516& ' Vertical arrow cursor.
    gdipIDC_WAIT = 32514& ' Hourglass cursor.
    gdipIDI_APPLICATION = 32512& ' Application icon.
    gdipIDI_ASTERISK = 32516& ' Asterisk icon.
    gdipIDI_EXCLAMATION = 32515& ' Exclamation point icon.
    gdipIDI_HAND = 32513& ' Stop sign icon.
    gdipIDI_QUESTION = 32514& ' Question-mark icon.
End Enum
Public Enum EGdipImageFormat
        gdipFormatUndefined = 0
        gdipFormatMemoryBMP = 1
        gdipFormatBMP = 2
        gdipFormatEMF = 3
        gdipFormatWMF = 4
        gdipFormatJPEG = 5
        gdipFormatPNG = 6
        gdipFormatGIF = 7
        gdipFormatTIFF = 8
        gdipFormatEXIF = 9
        gdipFormatICON = 10
End Enum
Public Enum EGdipImageFormatSave
        gdipFormatSaveBMP = 2
        gdipFormatSaveEMF = 3
        gdipFormatSaveJPEG = 5
        gdipFormatSavePNG = 6
        gdipFormatSaveGIF = 7
        gdipFormatSaveTIFF = 8
End Enum
Public Enum EGdipImageFormatSaveArray
        gdipFormatSaveArrayBMP = 2
        gdipFormatSaveArrayJPEG = 5
        gdipFormatSaveArrayPNG = 6
        gdipFormatSaveArrayGIF = 7
        gdipFormatSaveArrayTIFF = 8
        gdipFormatSaveArrayPictureDataBMP = 98
        gdipFormatSaveArrayPictureDataEMF = 99
End Enum
Public Enum EGdipLineCap
   LineCapFlat = 0
   LineCapSquare = 1
   LineCapRound = 2
   LineCapTriangle = 3
   'LineCapNoAnchor = &H10
   LineCapSquareAnchor = &H11
   LineCapRoundAnchor = &H12
   LineCapDiamondAnchor = &H13
   LineCapArrowAnchor = &H14
   'LineCapCustom = &HFF
   'LineCapAnchorMask = &HF0
End Enum
Public Enum EGdipLineJoin
   LineJoinMiter = 0
   LineJoinBevel = 1
   LineJoinRound = 2
   LineJoinMiterClipped = 3
End Enum
Public Enum EGdipTransparencyMode
    GdipTransparencySourceOver = 0
    GdipTransparencySourceCopy = 1
End Enum
Public Enum EGdipSmoothingMode
    GdipSmoothingInvalid = -1
    GdipSmoothingDefault = 0
    GdipSmoothingHighSpeed = 1
    GdipSmoothingHighQuality = 2
    GdipSmoothingNone = 3
    GdipSmoothingAntialias = 4
End Enum
Public Enum EGdipPixelOffsetMode
    GdipPixelOffsetInvalid = -1
    GdipPixelOffsetDefault = 0
    GdipPixelOffsetHighSpeed = 1
    GdipPixelOffsetHighQuality = 2
    GdipPixelOffsetNone = 3
    GdipPixelOffsetHalf = 4
End Enum
Public Enum EGdipCompositingQuality
    GdipCompositingQualityInvalid = -1
    GdipCompositingQualityDefault = 0
    GdipCompositingQualityHighSpeed = 1
    GdipCompositingQualityHighQuality = 2
    GdipCompositingQualityGammaCorrected = 3
    GdipCompositingQualityAssumeLinear = 4
End Enum
Public Enum EGdipAntialisingLevel
    GdipAntialiseInvalid = -1
    GdipAntialiseDefault = 0
    GdipAntialiseLowQuality = 1
    GdipAntialiseHighQuality = 2
    GdipAntialiseBilinear = 3
    GdipAntialiseBicubic = 4
    GdipAntialiseNearestNeighbor = 5
    GdipAntialiseBilinearHQ = 6
    GdipAntialiseBicubicHQ = 7
End Enum
Public Enum EGdipTextRenderingHint
   GdipHintSystemDefault = 0
   GdipHintSingleBitPerPixelGridFit
   GdipHintSingleBitPerPixel
   GdipHintAntiAliasGridFit
   GdipHintAntiAlias
   GdipClearTypeGridFit
End Enum
Public Enum EGdipIconType
    GdipNoIcon = 0
    GdipSmallIcon = 1
    GdipLargeIcon = 2
End Enum
Public Enum EGdipCombineMode
    CombineModeReplace = 0
    CombineModeIntersect = 1
    CombineModeUnion = 2
    CombineModeXor = 3
    CombineModeExclude = 4
    CombineModeComplement = 5
End Enum
Public Enum EGdipImgSizeMode
    #If AppName = "A" Then
    GdipSizeModeClip = acOLESizeClip
    GdipSizeModeStretch = acOLESizeStretch
    GdipSizeModeZoom = acOLESizeZoom
    GdipSizeModeAutoSize = acOLESizeAutoSize
    #Else
    GdipSizeModeClip = fmPictureSizeModeClip
    GdipSizeModeStretch = fmPictureSizeModeStretch
    GdipSizeModeZoom = fmPictureSizeModeZoom
    GdipSizeModeAutoSize = 2
    #End If
End Enum
Public Enum EGdipImgPictureAlignment
    GdipAlignTopLeft = 0
    GdipAlignTopRight = 1
    GdipAlignCenter = 2
    GdipAlignBottomLeft = 3
    GdipAlignBottomRight = 4
End Enum
Public Enum EGdipPenAlignMode
    PenAlignmentCenter = 0
    PenAlignmentInset = 1
End Enum
Public Enum EGdipDashStyle
    GdipDashSOLID = 0
    GdipDashDASH = 1           '  -------
    GdipDashDOT = 2            '  .......
    GdipDashDASHDOT = 3        '  _._._._
    GdipDashDASHDOTDOT = 4     '  _.._.._
    GdipDashDASHCUSTOM = 5     ' cf DrawDashCustom
End Enum
Public Enum EGdipHatchStyle
    HatchStyleHorizontal                   ' 0
    HatchStyleVertical                     ' 1
    HatchStyleForwardDiagonal              ' 2
    HatchStyleBackwardDiagonal             ' 3
    HatchStyleCross                        ' 4
    HatchStyleDiagonalCross                ' 5
    HatchStyle05Percent                    ' 6
    HatchStyle10Percent                    ' 7
    HatchStyle20Percent                    ' 8
    HatchStyle25Percent                    ' 9
    HatchStyle30Percent                    ' 10
    HatchStyle40Percent                    ' 11
    HatchStyle50Percent                    ' 12
    HatchStyle60Percent                    ' 13
    HatchStyle70Percent                    ' 14
    HatchStyle75Percent                    ' 15
    HatchStyle80Percent                    ' 16
    HatchStyle90Percent                    ' 17
    HatchStyleLightDownwardDiagonal        ' 18
    HatchStyleLightUpwardDiagonal          ' 19
    HatchStyleDarkDownwardDiagonal         ' 20
    HatchStyleDarkUpwardDiagonal           ' 21
    HatchStyleWideDownwardDiagonal         ' 22
    HatchStyleWideUpwardDiagonal           ' 23
    HatchStyleLightVertical                ' 24
    HatchStyleLightHorizontal              ' 25
    HatchStyleNarrowVertical               ' 26
    HatchStyleNarrowHorizontal             ' 27
    HatchStyleDarkVertical                 ' 28
    HatchStyleDarkHorizontal               ' 29
    HatchStyleDashedDownwardDiagonal       ' 30
    HatchStyleDashedUpwardDiagonal         ' 31
    HatchStyleDashedHorizontal             ' 32
    HatchStyleDashedVertical               ' 33
    HatchStyleSmallConfetti                ' 34
    HatchStyleLargeConfetti                ' 35
    HatchStyleZigZag                       ' 36
    HatchStyleWave                         ' 37
    HatchStyleDiagonalBrick                ' 38
    HatchStyleHorizontalBrick              ' 39
    HatchStyleWeave                        ' 40
    HatchStylePlaid                        ' 41
    HatchStyleDivot                        ' 42
    HatchStyleDottedGrid                   ' 43
    HatchStyleDottedDiamond                ' 44
    HatchStyleShingle                      ' 45
    HatchStyleTrellis                      ' 46
    HatchStyleSphere                       ' 47
    HatchStyleSmallGrid                    ' 48
    HatchStyleSmallCheckerBoard            ' 49
    HatchStyleLargeCheckerBoard            ' 50
    HatchStyleOutlinedDiamond              ' 51
    HatchStyleSolidDiamond                 ' 52
End Enum
Public Enum EGdipTypeEllipse
    TypeEllipseRectangle = 0
    TypeEllipseCenter = 1
End Enum
Public Enum EGdipVertAlign
    VertAlignTop = 0
    VertAlignCenter = 1
    VertAlignBottom = 2
End Enum
Public Enum EGdipHorzAlign
    HorzAlignLeft = 0
    HorzAlignCenter = 1
    HorzAlignRight = 2
End Enum
Public Enum EGdipControlType
    CtrlCaptionClose = 1
    CtrlCaptionMin = 2
    CtrlCaptionMax = 3
    CtrlCaptionRestore = 4
    CtrlCaptionHelp = 5
    CtrlMenuArrow = 6
    CtrlMenuCheck = 7
    CtrlMenuBullet = 8
    CtrlMenuArrowRight = 9
    CtrlScrollUp = 10
    CtrlScrollDown = 11
    CtrlScrollLeft = 12
    CtrlScrollRight = 13
    CtrlScrollComboBox = 14
    CtrlScrollSizeGrip = 15
    CtrlScrollSizeGripright = 16
    CtrlButtonCheck = 17
    CtrlButtonRadio = 18
    CtrlButtonPush = 19
    CtrlFrameSunken = 20
    CtrlFrameRaised = 21
End Enum
Public Enum EGdipFillMode
   GdipFillModeAlternate        ' 0
   GdipFillModeWinding           ' 1
End Enum
#Else
Public Const PixelFormat1bppIndexed = &H30101
Public Const PixelFormat4bppIndexed = &H30402
Public Const PixelFormat8bppIndexed = &H30803
Public Const PixelFormat16bppGreyScale = &H101004
Public Const PixelFormat16bppRGB555 = &H21005
Public Const PixelFormat16bppRGB565 = &H21006
Public Const PixelFormat16bppARGB1555 = &H61007
Public Const PixelFormat24bppRGB = &H21808
Public Const PixelFormat32bppRGB = &H22009
Public Const PixelFormat32bppARGB = &H26200A
Public Const PixelFormat32bppPARGB = &HE200B
public Const PixelFormat48bppRGB = &H10300C
public Const PixelFormat64bppARGB = &H34400D
public Const PixelFormat64bppPARGB = &H1C400E
Private Const GdipNoIcon As Long = &H0
Private Const GdipSmallIcon As Long = 1
Private Const GdipLargeIcon As Long = 2
Private Const GdipSizeModeZoom As Long = 3
Private Const GdipSizeModeClip As Long = 0
Private Const GdipSizeModeStretch As Long = 1
Private Const GdipSizeModeAutoSize As Long = 2
Private Const GdipAntialiseDefault  As Long = 0
Private Const GdipAntialiseHighQuality As Long = 2
Private Const GdipAntialiseNearestNeighbor As Long = 5
Private Const GdipTransparencySourceOver As Long = 0
Private Const GdipPixelOffsetDefault As Long = 0
Private Const GdipSmoothingDefault As Long = 0
Private Const GdipHintAntiAliasGridFit As Long = 3
Private Const GdipCompositingQualityHighSpeed As Long = 1
Private Const CombineModeIntersect As Long = 1
Private Const CombineModeUnion As Long = 2
Private Const LineCapArrowAnchor As Long = &H14
Private Const gdipFormatUndefined = 0
Private Const gdipFormatMemoryBMP = 1
Private Const gdipFormatBMP = 2
Private Const gdipFormatEMF = 3
Private Const gdipFormatWMF = 4
Private Const gdipFormatJPEG = 5
Private Const gdipFormatPNG = 6
Private Const gdipFormatGIF = 7
Private Const gdipFormatTIFF = 8
Private Const gdipFormatEXIF = 9
Private Const gdipFormatICON = 10
Private Const gdipFormatSaveBMP = 2
Private Const gdipFormatSaveEMF = 3
Private Const gdipFormatSaveJPEG = 5
Private Const gdipFormatSavePNG = 6
Private Const gdipFormatSaveGIF = 7
Private Const gdipFormatSaveTIFF = 8
#End If
#If VBA6 Then
Public Enum EGdipJpegTransform
    EncoderValueTransformRotate90 = 13
    EncoderValueTransformRotate180 = 14
    EncoderValueTransformRotate270 = 15
    EncoderValueTransformFlipHorizontal = 16
    EncoderValueTransformFlipVertical = 17
End Enum
#Else
Property Get zJpegTransformRotate90(): zJpegTransformRotate90 = 13: End Property
Property Get zJpegTransformRotate180(): zJpegTransformRotate180 = 14: End Property
Property Get zJpegTransformRotate270(): zJpegTransformRotate270 = 15: End Property
Property Get zJpegTransformFlipHoriz(): zJpegTransformFlipHoriz = 16: End Property
Property Get zJpegTransformFlipVert(): zJpegTransformFlipVert = 17: End Property
#End If
#If VBA6 Then
Public Enum EGdipRotateFlip
    RotateNoneFlipNone = 0
    Rotate90flipnone = 1
    Rotate180flipnone = 2
    Rotate270flipnone = 3
    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7
    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipXY = Rotate90flipnone
    RotateNoneFlipXY = Rotate180flipnone
    Rotate90FlipXY = Rotate270flipnone
    Rotate180FlipXY = RotateNoneFlipNone
End Enum
#Else
Property Get zRotate180FlipXY(): zRotate180FlipXY = zRotateNoneFlipNone: End Property
Property Get zRotate90FlipXY(): zRotate90FlipXY = zRotate270FlipNone: End Property
Property Get zRotateNoneFlipXY(): zRotateNoneFlipXY = zRotate180FlipNone: End Property
Property Get zRotate270FlipXY(): zRotate270FlipXY = zRotate90FlipNone: End Property
Property Get zRotate180FlipY(): zRotate180FlipY = zRotateNoneFlipX: End Property
Property Get zRotate90FlipY(): zRotate90FlipY = zRotate270FlipX: End Property
Property Get zRotateNoneFlipY(): zRotateNoneFlipY = zRotate180FlipX: End Property
Property Get zRotate270FlipX(): zRotate270FlipX = 7: End Property
Property Get zRotate180FlipX(): zRotate180FlipX = 6: End Property
Property Get zRotate90FlipX(): zRotate90FlipX = 5: End Property
Property Get zRotateNoneFlipX(): zRotateNoneFlipX = 4: End Property
Property Get zRotate270FlipNone(): zRotate270FlipNone = 3: End Property
Property Get zRotate180FlipNone(): zRotate180FlipNone = 2: End Property
Property Get zRotate90FlipNone(): zRotate90FlipNone = 1: End Property
Property Get zRotateNoneFlipNone(): zRotateNoneFlipNone = 0: End Property
#End If
'*-----------------------------------------*
'************  Norme Exif 2.2  *************
'*-----------------------------------------*
#If VBA6 Then
Public Enum EGdipTagName
    ' A. Tags Relating to Version
    TagExifVersion = 36864
    TagFlashpixVersion = 40960
    ' B. Tag Relating to Image Data Characteristics
    TagColorSpace = 40961
    ' C. Tags Relating to Image Configuration
    ' D. Tags Relating to User Information
    TagMakerNote = &H927C&
    TagUserComment = &H9286&
    ' E. Tag Relating to Related File Information
    TagRelatedSoundFile = &HA004&
    ' F. Tags Relating to Date and Time
    TagDateTimeOriginal = &H9003&
    TagDateTimeDigitized = &H9004&
    ' G. Tags Relating to Picture-Taking Conditions
    TagExposureTime = &H829A&
    TagFNumber = &H829D&
    TagExposureProgram = &H8822&
    TagSpectralSensitivity = &H8824&
    TagISOSpeedRatings = &H8827&
    TagShutterSpeedValue = &H9201&
    TagApertureValue = &H9202&
    TagBrightnessValue = &H9203&
    TagExposureBiasValue = &H9204&
    TagMaxApertureValue = &H9205&
    TagSubjectDistance = &H9206&
    TagMeteringMode = &H9207&
    TagLightSource = &H9208&
    TagFlash = &H9209&
    TagFocalLength = &H920A&
    TagSubjectArea = &H9214&
    TagFlashEnergy = &HA20B&
    TagSubjectLocation = &HA214&
    TagExposureIndex = &HA215&
    TagSensingMethod = &HA217&
    TagExposureMode = &HA402&
    TagWhiteBalance = &HA403&
    TagDigitalZoomRatio = &HA404&
    TagFocalLengthIn35mmFilm = &HA405&
    TagSceneCaptureType = &HA406&
    TagGainControl = &HA407&
    TagContrast = &HA408&
    TagSaturation = &HA409&
    TagSharpness = &HA40A&
    TagDeviceSettingDescription = &HA40B&
    TagSubjectDistanceRange = &HA40C&
    ' H. Other Tags
    TagImageUniqueID = &HA420&
    TagDocumentName = &H10D&
    TagImageDescription = &H10E&
    TagImageWidth = &H100&
    TagImageHeight = &H101&
    TagEquipModel = &H110&
    TagEquipMake = &H10F&
    TagCompression = &H103&
    TagOrientation = &H112&
    '1 = Horizontal (normal)
    '2 = Mirror horizontal
    '3 = Rotate 180
    '4 = Mirror vertical
    '5 = Mirror horizontal and rotate 270 CW
    '6 = Rotate 90 CW
    '7 = Mirror horizontal and rotate 90 CW
    '8 = Rotate 270 CW
    TagThumbnailOrientation = &H5029&
    TagYResolution = &H11B&
    TagXResolution = &H11A&
    TagThumbnailData = &H501B&
    TagSoftWare = 305
    TagArtist = 315
    TagCopyright = 33432
End Enum
#Else
' A. Tags Relating to Version
Property Get TagExifVersion(): TagExifVersion = 36864: End Property
Property Get TagFlashpixVersion(): TagFlashpixVersion = 40960: End Property
' B. Tag Relating to Image Data Characteristics
Property Get TagColorSpace(): TagColorSpace = 40961: End Property
' C. Tags Relating to Image Configuration
' D. Tags Relating to User Information
Property Get TagMakerNote(): TagMakerNote = &H927C&: End Property
Property Get TagUserComment(): TagUserComment = &H9286&: End Property
' E. Tag Relating to Related File Information
Property Get TagRelatedSoundFile(): TagRelatedSoundFile = &HA004&: End Property
' F. Tags Relating to Date and Time
Property Get TagDateTimeOriginal(): TagDateTimeOriginal = &H9003&: End Property
Property Get TagDateTimeDigitized(): TagDateTimeDigitized = &H9004&: End Property
' G. Tags Relating to Picture-Taking Conditions
Property Get TagExposureTime(): TagExposureTime = &H829A&: End Property
Property Get TagFNumber(): TagFNumber = &H829D&: End Property
Property Get TagExposureProgram(): TagExposureProgram = &H8822&: End Property
Property Get TagSpectralSensitivity(): TagSpectralSensitivity = &H8824&: End Property
Property Get TagISOSpeedRatings(): TagISOSpeedRatings = &H8827&: End Property
Property Get TagShutterSpeedValue(): TagShutterSpeedValue = &H9201&: End Property
Property Get TagApertureValue(): TagApertureValue = &H9202&: End Property
Property Get TagBrightnessValue(): TagBrightnessValue = &H9203&: End Property
Property Get TagExposureBiasValue(): TagExposureBiasValue = &H9204&: End Property
Property Get TagMaxApertureValue(): TagMaxApertureValue = &H9205&: End Property
Property Get TagSubjectDistance(): TagSubjectDistance = &H9206&: End Property
Property Get TagMeteringMode(): TagMeteringMode = &H9207&: End Property
Property Get TagLightSource(): TagLightSource = &H9208&: End Property
Property Get TagFlash(): TagFlash = &H9209&: End Property
Property Get TagFocalLength(): TagFocalLength = &H920A&: End Property
Property Get TagSubjectArea(): TagSubjectArea = &H9214&: End Property
Property Get TagFlashEnergy(): TagFlashEnergy = &HA20B&: End Property
Property Get TagSubjectLocation(): TagSubjectLocation = &HA214&: End Property
Property Get TagExposureIndex(): TagExposureIndex = &HA215&: End Property
Property Get TagSensingMethod(): TagSensingMethod = &HA217&: End Property
Property Get TagExposureMode(): TagExposureMode = &HA402&: End Property
Property Get TagWhiteBalance(): TagWhiteBalance = &HA403&: End Property
Property Get TagDigitalZoomRatio(): TagDigitalZoomRatio = &HA404&: End Property
Property Get TagFocalLengthIn35mmFilm(): TagFocalLengthIn35mmFilm = &HA405&: End Property
Property Get TagSceneCaptureType(): TagSceneCaptureType = &HA406&: End Property
Property Get TagGainControl(): TagGainControl = &HA407&: End Property
Property Get TagContrast(): TagContrast = &HA408&: End Property
Property Get TagSaturation(): TagSaturation = &HA409&: End Property
Property Get TagSharpness(): TagSharpness = &HA40A&: End Property
Property Get TagDeviceSettingDescription(): TagDeviceSettingDescription = &HA40B&: End Property
Property Get TagSubjectDistanceRange(): TagSubjectDistanceRange = &HA40C&: End Property
Property Get TagOrientation(): TagOrientation = &H112&: End Property
Property Get TagThumbnailOrientation(): TagThumbnailOrientation = &H5029&: End Property
' H. Other Tags
Property Get TagImageUniqueID(): TagImageUniqueID = &HA420&: End Property
Property Get TagDocumentName(): TagDocumentName = &H10D&: End Property
Property Get TagImageDescription(): TagImageDescription = &H10E&: End Property
Property Get TagImageWidth(): TagImageWidth = &H100&: End Property
Property Get TagImageHeight(): TagImageHeight = &H101&: End Property
Property Get TagEquipModel(): TagEquipModel = &H110&: End Property
Property Get TagEquipMake(): TagEquipMake = &H10F&: End Property
Property Get TagCompression(): TagCompression = &H103&: End Property
Property Get TagYResolution(): TagYResolution = &H11B&: End Property
Property Get TagXResolution(): TagXResolution = &H11A&: End Property
Property Get TagThumbnailData(): TagThumbnailData = &H501B&: End Property
Property Get TagSoftWare(): TagSoftWare = 305: End Property
Property Get TagArtist(): TagArtist = 315: End Property
Property Get TagCopyright(): TagCopyright = 33432: End Property
#End If
'***************************************************************************************
'*                                    PROPRIETES                                       *
'***************************************************************************************
Public Property Get IsV11() As Boolean
    IsV11 = gIsV11
End Property
Public Property Get Parent() As clGdiplus
    If ImgParentPtr <> 0 Then Set Parent = PtrToObj(ImgParentPtr)
End Property
Public Property Let UseGdi32(pValue As Boolean)
gUseGdiWindow = pValue
gUseGdiImg = pValue
If Not pValue Then gUseGdi = False
End Property
Public Property Get UseGdi32() As Boolean
UseGdi32 = gUseGdi
End Property
Public Property Get hDC()
hDC = gDib.hDC
End Property
Public Property Get hBrush()
hBrush = gDib.hBrush
End Property
Private Property Get PxFmtAlpha(pFmt As Long) As Boolean
Select Case pFmt
    Case PixelFormat16bppARGB1555, PixelFormat32bppARGB, PixelFormat32bppPARGB, PixelFormat64bppARGB, PixelFormat64bppPARGB
        PxFmtAlpha = True
    Case Else
        PxFmtAlpha = False
End Select
End Property
Public Property Get HasTransparency() As Boolean
Dim lPF As Long
GdipGetImagePixelFormat gBitmap, lPF
HasTransparency = PxFmtAlpha(lPF)
End Property
Public Property Let HasTransparency(pValue As Boolean)
If gBitmap <> 0 Then
    If HasTransparency = False Then
        If pValue = True Then
            ConvertBppGraphics gBitmap, True
        End If
    Else
        If pValue = False Then
            ConvertBppGraphics gBitmap, , True
        End If
    End If
End If
End Property
#If VBA6 Then
Public Property Let TransparencyMode(pValue As EGdipTransparencyMode)
#Else
Public Property Let TransparencyMode(pValue As Long)
#End If
gTransparencyMode = pValue
If gGraphics <> 0 Then
    GdipSetCompositingMode gGraphics, pValue
End If
End Property
#If VBA6 Then
Public Property Get TransparencyMode() As EGdipTransparencyMode
#Else
Public Property Get TransparencyMode() As Long
#End If
If gGraphics <> 0 Then
    GdipGetCompositingMode gGraphics, TransparencyMode
Else
    TransparencyMode = -1
End If
End Property
#If VBA6 Then
Public Property Let PixelOffsetMode(pValue As EGdipPixelOffsetMode)
#Else
Public Property Let PixelOffsetMode(pValue As Long)
#End If
gPixelOffsetMode = pValue
If gGraphics <> 0 Then
    GdipSetPixelOffsetMode gGraphics, pValue
End If
End Property
#If VBA6 Then
Public Property Get PixelOffsetMode() As EGdipPixelOffsetMode
#Else
Public Property Get PixelOffsetMode() As Long
#End If
If gGraphics <> 0 Then
    GdipGetPixelOffsetMode gGraphics, PixelOffsetMode
Else
    PixelOffsetMode = -1
End If
End Property
#If VBA6 Then
Public Property Let CompositingQuality(pValue As EGdipCompositingQuality)
#Else
Public Property Let CompositingQuality(pValue As Long)
#End If
gCompositingQuality = pValue
If gGraphics <> 0 Then
    GdipSetCompositingQuality gGraphics, pValue
End If
End Property
#If VBA6 Then
Public Property Get CompositingQuality() As EGdipCompositingQuality
#Else
Public Property Get CompositingQuality() As Long
#End If
If gGraphics <> 0 Then
    GdipGetCompositingQuality gGraphics, CompositingQuality
Else
    CompositingQuality = -1
End If
End Property
#If VBA6 Then
Public Property Let SmoothingMode(pValue As EGdipSmoothingMode)
#Else
Public Property Let SmoothingMode(pValue As Long)
#End If
gSmoothingMode = pValue
If gGraphics <> 0 Then
    GdipSetSmoothingMode gGraphics, pValue
End If
End Property
#If VBA6 Then
Public Property Get SmoothingMode() As EGdipSmoothingMode
#Else
Public Property Get SmoothingMode() As Long
#End If
If gGraphics <> 0 Then
    GdipGetSmoothingMode gGraphics, SmoothingMode
Else
    SmoothingMode = -1
End If
End Property
#If VBA6 Then
Public Property Let AntialisingLevel(pValue As EGdipAntialisingLevel)
#Else
Public Property Let AntialisingLevel(pValue As Long)
#End If
gAntialisingLevel = pValue
If gGraphics <> 0 Then
    GdipSetInterpolationMode gGraphics, pValue
End If
End Property
#If VBA6 Then
Public Property Get AntialisingLevel() As EGdipAntialisingLevel
#Else
Public Property Get AntialisingLevel() As Long
#End If
If gGraphics <> 0 Then
    GdipGetInterpolationMode gGraphics, AntialisingLevel
Else
    AntialisingLevel = -1
End If
End Property
#If VBA6 Then
Public Property Let TextRenderingHint(pValue As EGdipTextRenderingHint)
#Else
Public Property Let TextRenderingHint(pValue As Long)
#End If
gTextRenderingHint = pValue
If gGraphics <> 0 Then
    GdipSetTextRenderingHint gGraphics, pValue
End If
End Property
#If VBA6 Then
Public Property Get TextRenderingHint() As EGdipTextRenderingHint
#Else
Public Property Get TextRenderingHint() As Long
#End If
If gGraphics <> 0 Then
    GdipGetTextRenderingHint gGraphics, TextRenderingHint
Else
    TextRenderingHint = -1
End If
End Property
'---------------------------------------------------------------------------------------
' Cet objet (utile dans un With avec ImgNew par example)
'---------------------------------------------------------------------------------------
Public Property Get Self() As clGdiplus
Set Self = Me
End Property
'---------------------------------------------------------------------------------------
' DPI
'---------------------------------------------------------------------------------------
Public Property Get DpiX() As Single
Dim lResX As Single
On Error GoTo gestion_erreurs
    GdipGetImageHorizontalResolution gBitmap, lResX
    DpiX = lResX
    Exit Property
gestion_erreurs:
    DpiX = 0
End Property
Public Property Get DpiY() As Single
Dim lResY As Single
On Error GoTo gestion_erreurs
    GdipGetImageVerticalResolution gBitmap, lResY
    DpiY = lResY
    Exit Property
gestion_erreurs:
    DpiY = 0
End Property
Public Property Let DpiX(Dpi As Single)
GdipBitmapSetResolution gBitmap, Dpi, DpiY
End Property
Public Property Let DpiY(Dpi As Single)
GdipBitmapSetResolution gBitmap, DpiX, Dpi
End Property
'---------------------------------------------------------------------------------------
' Pointilles personnalises
'---------------------------------------------------------------------------------------
Public Property Get DrawDashCustom() As Variant
DrawDashCustom = gDashUser
End Property
Public Property Let DrawDashCustom(value As Variant)
Dim llb As Long, lub As Long
Dim lcpt As Long
On Error GoTo gestion_erreurs
llb = LBound(value)
lub = UBound(value)
ReDim gDashUser(1 To lub - llb + 1)
For lcpt = 1 To lub - llb + 1
    gDashUser(lcpt) = value(llb + lcpt - 1)
Next
Exit Property
gestion_erreurs:
    Erase gDashUser
End Property
'---------------------------------------------------------------------------------------
' Rectangle limitant le dessin
' En pixels sur l'image
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Sub DrawClipRectangle(Optional ByVal pX1 As Single, Optional ByVal pY1 As Single, Optional ByVal pX2 As Single, Optional ByVal pY2 As Single, Optional pInclude As Boolean = True)
#Else
Public Sub DrawClipRectangle(Optional ByVal pX1 As Long, Optional ByVal pY1 As Long, Optional ByVal pX2 As Long, Optional ByVal pY2 As Long, Optional pInclude As Boolean = True)
#End If
Const sRegion As String = "DummyRegionForDrawClipRectangle"
On Error GoTo gestion_erreurs
If CreateRegionRect(sRegion, pX1, pY1, pX2, pY2, 0) Then
    DrawClipRegion(pInclude) = sRegion
    RegionDelete sRegion
End If
Exit Sub
gestion_erreurs:
'
End Sub
'---------------------------------------------------------------------------------------
' Region limitant le dessin
'---------------------------------------------------------------------------------------
Public Property Let DrawClipRegion(Optional pInclude As Boolean = True, pRegion As String)
Dim lregion As Variant
Dim lRegionPtr, lgraphics
Dim lMatrix
On Error GoTo gestion_erreurs
lgraphics = GetGraphics
If gDrawClipRegion <> 0 Then
    GdipDeleteRegion gDrawClipRegion
    gDrawClipRegion = 0
End If
If pRegion = "" Then
    GdipResetClip lgraphics
Else
    lregion = gCollRegions(pRegion)
    If lregion(2) Then
        GdipCreateRegionPath lregion(0), lRegionPtr
    Else
        GdipCloneRegion lregion(0), lRegionPtr
    End If
    If pInclude Then
        gDrawClipRegion = lRegionPtr
    Else
        GdipCreateRegion gDrawClipRegion
        GdipSetInfinite gDrawClipRegion
        GdipCombineRegionRegion gDrawClipRegion, lRegionPtr, 4
        GdipDeleteRegion lRegionPtr
    End If
    GdipCreateMatrix lMatrix
    GdipGetWorldTransform lgraphics, lMatrix
    GdipResetWorldTransform lgraphics
    GdipSetClipRegion lgraphics, gDrawClipRegion, 0 ' CombineModeReplace
    GdipSetWorldTransform lgraphics, lMatrix
End If
PrivSetClipRegionGdi
Exit Property
gestion_erreurs:
'
End Property
Public Sub PrivSetClipRegionGdi()
    Dim lTempoRegionGdi
    Dim lWorldMatrix
    GdipCreateMatrix lWorldMatrix
    GdipGetWorldTransform gGraphics, lWorldMatrix
    GdipResetWorldTransform gGraphics
    GdipGetRegionHRgn gDrawClipRegion, gGraphics, lTempoRegionGdi
    SelectClipRgn hDC, lTempoRegionGdi
    DeleteObject lTempoRegionGdi
    GdipSetWorldTransform gGraphics, lWorldMatrix
    GdipDeleteMatrix lWorldMatrix
End Sub
'---------------------------------------------------------------------------------------
' Format des pixels de l'image (lecture)
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Property Get ImagePixelFormat() As EGdipPixelFormat
#Else
Public Property Get ImagePixelFormat() As Long
#End If
GdipGetImagePixelFormat gBitmap, ImagePixelFormat
End Property
Public Property Get ImagePixelFormatText() As String
Dim lPxfmt As Long
GdipGetImagePixelFormat gBitmap, lPxfmt
Select Case lPxfmt
    Case PixelFormat1bppIndexed: ImagePixelFormatText = "PixelFormat1bppIndexed"
    Case PixelFormat4bppIndexed: ImagePixelFormatText = "PixelFormat4bppIndexed"
    Case PixelFormat8bppIndexed: ImagePixelFormatText = "PixelFormat8bppIndexed"
    Case PixelFormat16bppGreyScale: ImagePixelFormatText = "PixelFormat16bppGreyScale"
    Case PixelFormat16bppRGB555: ImagePixelFormatText = "PixelFormat16bppRGB555"
    Case PixelFormat16bppRGB565: ImagePixelFormatText = "PixelFormat16bppRGB565"
    Case PixelFormat16bppARGB1555: ImagePixelFormatText = "PixelFormat16bppARGB1555"
    Case PixelFormat24bppRGB: ImagePixelFormatText = "PixelFormat24bppRGB"
    Case PixelFormat32bppRGB: ImagePixelFormatText = "PixelFormat32bppRGB"
    Case PixelFormat32bppARGB: ImagePixelFormatText = "PixelFormat32bppARGB"
    Case PixelFormat32bppPARGB: ImagePixelFormatText = "PixelFormat32bppPARGB"
    Case PixelFormat48bppRGB: ImagePixelFormatText = "PixelFormat48bppRGB"
    Case PixelFormat64bppARGB: ImagePixelFormatText = "PixelFormat64bppARGB"
    Case PixelFormat64bppPARGB: ImagePixelFormatText = "PixelFormat64bppPARGB"
    Case Else: ImagePixelFormatText = "Undefined"
End Select
End Property
'---------------------------------------------------------------------------------------
' Format de l'image (lecture)
'---------------------------------------------------------------------------------------
Public Property Get ImageFormat() As EGdipImageFormat
ImageFormat = PrivImageFormat(gBitmap)
End Property
Private Property Get PrivImageFormat(pBitmap) As EGdipImageFormat
    Dim lGuid As GUID
    Dim lPtr
    Dim lStr As String * 255
    If GdipGetImageRawFormat(pBitmap, lGuid) = 0 Then
        StringFromCLSID lGuid, lPtr
        StringFromPointer lPtr, lStr
        Select Case Left(lStr, 9)
            Case "{B96B3CA9": PrivImageFormat = gdipFormatUndefined
            Case "{B96B3CAA": PrivImageFormat = gdipFormatMemoryBMP
            Case "{B96B3CAB": PrivImageFormat = gdipFormatBMP
            Case "{B96B3CAC": PrivImageFormat = gdipFormatEMF
            Case "{B96B3CAD": PrivImageFormat = gdipFormatWMF
            Case "{B96B3CAE": PrivImageFormat = gdipFormatJPEG
            Case "{B96B3CAF": PrivImageFormat = gdipFormatPNG
            Case "{B96B3CB0": PrivImageFormat = gdipFormatGIF
            Case "{B96B3CB1": PrivImageFormat = gdipFormatTIFF
            Case "{B96B3CB2": PrivImageFormat = gdipFormatEXIF
            Case "{B96B3CB5": PrivImageFormat = gdipFormatICON
            Case Else: PrivImageFormat = gdipFormatUndefined
        End Select
    End If
End Property
Public Property Get ImageFormatText() As String
    Select Case ImageFormat
        Case gdipFormatUndefined: ImageFormatText = "Undefined"
        Case gdipFormatMemoryBMP: ImageFormatText = "MemoryBMP"
        Case gdipFormatBMP: ImageFormatText = "BMP"
        Case gdipFormatEMF: ImageFormatText = "EMF"
        Case gdipFormatWMF: ImageFormatText = "WMF"
        Case gdipFormatJPEG: ImageFormatText = "JPEG"
        Case gdipFormatPNG: ImageFormatText = "PNG"
        Case gdipFormatGIF: ImageFormatText = "GIF"
        Case gdipFormatTIFF: ImageFormatText = "TIFF"
        Case gdipFormatEXIF: ImageFormatText = "EXIF"
        Case gdipFormatICON: ImageFormatText = "ICON"
        Case Else: ImageFormatText = "Undefined"
    End Select
End Property

'---------------------------------------------------------------------------------------
' Hauteur de l'image
'---------------------------------------------------------------------------------------
Public Property Get ImageHeight() As Long
    GdipGetImageHeight gBitmap, ImageHeight
End Property
'---------------------------------------------------------------------------------------
' Largeur de l'image
'---------------------------------------------------------------------------------------
Public Property Get ImageWidth() As Long
    GdipGetImageWidth gBitmap, ImageWidth
End Property
'---------------------------------------------------------------------------------------
' Controle (ou formulaire) de reference pour coordonnees donnees en points/twips
'---------------------------------------------------------------------------------------
Public Property Get RefControl() As Object
    Set RefControl = gCtrlRef
End Property
Public Property Let RefControl(pCtrl As Object)
    Set gCtrlRef = pCtrl
    If gCtrlRefToImgMatrix <> 0 Then
        GdipDeleteMatrix gCtrlRefToImgMatrix
        gCtrlRefToImgMatrix = 0
    End If
End Property

Private Property Get CtrlRefToImgMatrix()
If gCtrlRefToImgMatrix <> 0 Then
    CtrlRefToImgMatrix = gCtrlRefToImgMatrix
Else
    gCtrlRefToImgMatrix = PrivMatrixCtrlToImg(gCtrlRef)
    CtrlRefToImgMatrix = gCtrlRefToImgMatrix
End If
End Property

'---------------------------------------------------------------------------------------
' Bitmap de travail (prive)
'---------------------------------------------------------------------------------------
Private Property Let PrivBitmap(pUseGdi As Boolean, pBitmap)
    Dim lNewBitmap, lNewGraphics
    Dim lWidth As Single, lHeight As Single
    Dim lHdc, lpGraphics, lpHdc
On Error GoTo gestion_erreurs
gUseGdi = pUseGdi
If gCtrlRefToImgMatrix <> PtrNull Then
    GdipDeleteMatrix gCtrlRefToImgMatrix
    gCtrlRefToImgMatrix = PtrNull
End If
If gGraphics <> PtrNull Then GdipDeleteGraphics gGraphics
gGraphics = PtrNull
gActiveFrame = 0
If gDib.hDIB <> PtrNull Then
    DeleteObject (SelectObject(gDib.hDC, gDib.hOldDIB))
    DeleteObject gDib.hBrush
    gDib.hBrush = PtrNull
    gDib.hDIB = PtrNull
End If

If Not gUseGdi Then
    If gBitmap <> PtrNull Then GdipDisposeImage gBitmap
    gBitmap = pBitmap
    DeleteDC gDib.hDC
    gDib.hDC = PtrNull
Else
    If gDib.hDIB <> PtrNull Then
        DeleteObject (SelectObject(gDib.hDC, gDib.hOldDIB))
        DeleteObject gDib.hBrush
        gDib.hBrush = PtrNull
        gDib.hDIB = PtrNull
    End If
    If gBitmap <> PtrNull Then GdipDisposeImage gBitmap
    gBitmap = PtrNull
    If pBitmap = PtrNull Then
        DeleteDC gDib.hDC
        gDib.hDC = PtrNull
        Exit Property
    End If
    GdipGetImageDimension pBitmap, lWidth, lHeight
    gDib.BI.bmiHeader.biBitCount = 32 '24
    gDib.BI.bmiHeader.biSizeImage = 0 '(lWidth * lHeight * 3) + (4 - ((lWidth * 3) Mod 4)) * lHeight
    gDib.BI.bmiHeader.biWidth = lWidth
    gDib.BI.bmiHeader.biHeight = lHeight
    gDib.BI.bmiHeader.biPlanes = 1
    #If Win64 Then
    gDib.BI.bmiHeader.biSize = LenB(gDib.BI.bmiHeader)
    #Else
    gDib.BI.bmiHeader.biSize = Len(gDib.BI.bmiHeader)
    #End If
    If gDib.hDC = PtrNull Then
        If PrivGdipGetImageGraphicsContext(pBitmap, lpGraphics) = 0 Then
            If GdipGetDC(lpGraphics, lpHdc) = 0 Then
                gDib.hDC = CreateCompatibleDC(lpHdc)
                GdipReleaseDC lpGraphics, lpHdc
            End If
            GdipDeleteGraphics lpGraphics
        End If
        If lpHdc = PtrNull Then
            lHdc = GetDC(PtrNull)
            gDib.hDC = CreateCompatibleDC(lHdc)
            ReleaseDC PtrNull, lHdc
        End If
    End If
    gDib.DIBPTR = PtrNull
    gDib.hDIB = CreateDIBSection(gDib.hDC, gDib.BI, DIB_RGB_COLORS, gDib.DIBPTR, 0, 0)
    gDib.hOldDIB = SelectObject(gDib.hDC, gDib.hDIB)
    GdipCreateBitmapFromGdiDib gDib.BI, gDib.DIBPTR, lNewBitmap
    GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
    Call GdipDrawImageRectI(lNewGraphics, pBitmap, 0, 0, lWidth, lHeight)
    GdipDeleteGraphics lNewGraphics
    GdipDisposeImage pBitmap
    If gUseGdiImg And ImgName <> "" Then gDib.hBrush = CreatePatternBrush(ByVal gDib.hDIB)
    gBitmap = lNewBitmap
End If
If gGraphics = PtrNull Then GetGraphics True
Exit Property
gestion_erreurs:
    '
End Property

'***************************************************************************************
'*                                    FONCTIONS EXIF                                   *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Propage l'erreur a l'appelant
'---------------------------------------------------------------------------------------
Private Sub ErrRaise()
    Err.Raise Err.Number, Err.Source, Err.description, Err.HelpFile, Err.HelpContext
End Sub

'---------------------------------------------------------------------------------------
' Lit le tag Exif
'---------------------------------------------------------------------------------------
' pTag  : Tag Exif
' Renvoie la valeur du tag
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function GetExifData(pTag As EGdipTagName) As Variant
#Else
Public Function GetExifData(pTag As Long) As Variant
#End If
' Source http://www.activevb.de/tipps/vb6tipps/tipp0666.html
Dim lret As Long, lPropCount As Long
Dim lBuffer() As Byte, lReturnBuffer() As Byte
Dim lPropSize As Long
Dim lPropertyItem As PropertyItem
Dim lResultStr As String
Dim lResultLong() As Long
Dim lResultShort() As Integer
Dim lResultRational() As Long
Dim lcpt As Integer, lStart As Integer
Dim lNullResult As Boolean
Dim lhMem, lPtr
' Pour Access, renvoit Null si pas de miniature
#If AppName = "A" Then
    lNullResult = True
#End If
On Error GoTo gestion_erreurs:
' Compte le nombre de proprietes Exif
lret = GdipGetPropertyCount(gBitmap, lPropCount)
If lPropCount > 0 Then
    ' Taille de la propriete
    If GdipGetPropertyItemSize(gBitmap, pTag, lPropSize) = 0 Then
        If lPropSize > 0 Then
            ' Lecture de la propriete
            ReDim lBuffer(lPropSize - 1)
            If GdipGetPropertyItem(gBitmap, pTag, lPropSize, lBuffer(0)) = 0 Then
                Call RtlMoveMemory(lPropertyItem, lBuffer(0), LenB(lPropertyItem))
                ' tPropertyItem.length correspond a la taille de la propriete
                If lPropertyItem.Length > 0 Then
                    ' On deplace la valeur dans un tableau (tPropertyItem.Value est un pointeur)
                    ReDim lReturnBuffer(lPropertyItem.Length - 1)
                    Call RtlMoveMemory(lReturnBuffer(0), _
                                       ByVal lPropertyItem.value, _
                                       lPropertyItem.Length)
                    ' Retour de la fonction suivant le type de donnees
                    Select Case lPropertyItem.Type
                    Case PropertyTagTypeASCII
                        ' Traitements specifiques
                        Select Case pTag
                        Case TagDateTimeOriginal, TagDateTimeDigitized
                            ' Date au format Date standard
                            lResultStr = StrConv(lReturnBuffer, vbUnicode)
                            GetExifData = DateSerial(Mid(lResultStr, 1, 4), Mid(lResultStr, 6, 2), Mid(lResultStr, 9, 2)) _
                                        + TimeSerial(Mid(lResultStr, 12, 2), Mid(lResultStr, 15, 2), Mid(lResultStr, 18, 2))
                        Case Else
                            ' Chaine de caracteres
                            GetExifData = StrConv(lReturnBuffer, vbUnicode)
                            GetExifData = Left(GetExifData, Len(GetExifData) - 1)
                        End Select
                    Case PropertyTagTypeUndefined, PropertyTagTypeByte, 6
                        ' Traitements specifiques
                        Select Case pTag
                        Case TagExifVersion, TagFlashpixVersion, TagMakerNote
                            ' Ces donnees sont en Ascii
                            GetExifData = StrConv(lReturnBuffer, vbUnicode)
                        Case TagThumbnailData
                            ' Miniature
                            Dim stream As Object
                            Dim lExifThumb
                            lhMem = GlobalAlloc(GMEM_MOVEABLE, lPropertyItem.Length)
                            If lhMem Then
                                lPtr = GlobalLock(lhMem)
                                If lPtr Then
                                    RtlMoveMemory ByVal lPtr, lReturnBuffer(0), lPropertyItem.Length
                                    GlobalUnlock lhMem
                                    ' Transformation du tableau en flux (stream)
                                    If CreateStreamOnHGlobal(ByVal lhMem, True, stream) = 0 Then
                                        ' Lecture du flux dans un bitmap GDI+
                                        If GdipLoadImageFromStream(stream, lExifThumb) = 0 Then
                                            #If AppName = "A" Then
                                                ' Retour de la fonction = tableau de byte affectable a une propriete PictureData
                                                GetExifData = BitmaptoPictureDataEMF(lExifThumb)
                                            #Else
                                                ' Retour de la fonction = object IPictureDisp
                                                Set GetExifData = PrivGdiPlusToIPictureDisp(, , lExifThumb)
                                            #End If
                                            ' Suppression du bitmap gdi+
                                            GdipDisposeImage lExifThumb
                                        End If
                                    End If
                                End If
                                'GlobalFree lhMem
                            End If
                        Case Else
                            ' Bytes
                            If UBound(lReturnBuffer) = 0 Then
                                GetExifData = lReturnBuffer(0)
                            Else
                                GetExifData = lReturnBuffer
                            End If
                        End Select
                    Case PropertyTagTypeShort, 8
                        ' Entiers
                        ReDim lResultShort(lPropertyItem.Length / 2 - 1)
                        RtlMoveMemory lResultShort(0), lReturnBuffer(0), lPropertyItem.Length
                        ' Traitements specifiques
                        Select Case pTag
                        Case TagFlash
                            For lcpt = 7 To 0 Step -1
                                If lResultShort(0) >= 2 ^ lcpt Then
                                    lResultStr = lResultStr & "1"
                                    lResultShort(0) = lResultShort(0) - 2 ^ lcpt
                                Else
                                    lResultStr = lResultStr & "0"
                                End If
                            Next
                            GetExifData = lResultStr
                        Case Else
                            If UBound(lResultShort) = 0 Then
                                GetExifData = lResultShort(0)
                            Else
                                GetExifData = lResultShort
                            End If
                        End Select
                    Case PropertyTagTypeRational, PropertyTagTypeSRational, 12    ' double
                        ' Rationnel (numerateur + denominateur)
                        ReDim lResultRational(lPropertyItem.Length / 4 - 1)
                        For lcpt = 0 To lPropertyItem.Length / 8 - 1
                            lStart = lcpt * 8
                            ' Numerateur
                            RtlMoveMemory lResultRational(lcpt * 2), lReturnBuffer(lStart), 4
                            ' Denominateur
                            RtlMoveMemory lResultRational(lcpt * 2 + 1), lReturnBuffer(lStart + 4), 4
                        Next
                        GetExifData = lResultRational
                    Case PropertyTagTypeLong, PropertyTagTypeSLong, 11    ' float
                        ' Long
                        ReDim lResultLong(lPropertyItem.Length / 4 - 1)
                        RtlMoveMemory lResultLong(0), lReturnBuffer(0), lPropertyItem.Length
                        If UBound(lResultLong) = 0 Then
                            GetExifData = lResultLong(0)
                        Else
                            GetExifData = lResultLong
                        End If
                    End Select
                End If
            End If
        Else
            ' Initialise le resultat
            If pTag = TagThumbnailData And Not lNullResult Then
                Set GetExifData = Nothing
            Else
                GetExifData = Null
            End If
        End If
    Else
        ' Initialise le resultat
        If pTag = TagThumbnailData And Not lNullResult Then
            Set GetExifData = Nothing
        Else
            GetExifData = Null
        End If
    End If
Else
    ' Initialise le resultat
    If pTag = TagThumbnailData And Not lNullResult Then
        Set GetExifData = Nothing
    Else
        GetExifData = Null
    End If
End If
' Donnees non Exif
' La taille de l'image n'est pas stockee dans les tags Exif en cas de compression
' On lit donc directement la taille
Select Case pTag
Case TagImageWidth
    If IsNull(GetExifData) Then GetExifData = 0: GdipGetImageWidth gBitmap, GetExifData
Case TagImageHeight
    If IsNull(GetExifData) Then GetExifData = 0: GdipGetImageHeight gBitmap, GetExifData
End Select
Exit Function
gestion_erreurs:
    ' Initialise le resultat
    If pTag = TagThumbnailData And Not lNullResult Then
        Set GetExifData = Nothing ' Nothing pour image userform
    Else
        GetExifData = Null
    End If
End Function

'---------------------------------------------------------------------------------------
' Ecrit certains tags Exif
' - TagDateTimeOriginal
' - TagImageDescription
' - TagArtist
' - TagOrientation
' - TagThumbnailOrientation
'---------------------------------------------------------------------------------------
' pTag   : Tag Exif
' pValue : Valeur du tag
'---------------------------------------------------------------------------------------
#If VBA7 Then
Public Function SetExifData(pTag As EGdipTagName, pValue As Variant) As Boolean
Dim lPtr As LongPtr
#Else
Public Function SetExifData(pTag As Long, pValue As Variant) As Boolean
Dim lPtr As Long
#End If
Dim lPropertyItem As PropertyItem
Dim lResultStr As String
Dim lResultShort As Integer
Dim lNeedToFree As Boolean
On Error GoTo gestion_erreurs:
SetExifData = True

lPropertyItem.id = pTag
' Si Null, supprime le tag
If IsNull(pValue) Then
    SetExifData = (GdipRemovePropertyItem(gBitmap, pTag) = 0)
Else
    Select Case pTag
    Case TagThumbnailOrientation, TagOrientation
        ' Orientation de l'image
        lPropertyItem.Length = 2
        lPropertyItem.Type = PropertyTagTypeShort
        lResultShort = pValue
        lPropertyItem.value = GdipAlloc(2)
        lNeedToFree = True
        RtlMoveMemory ByVal lPropertyItem.value, lResultShort, Len(lResultShort)
    Case TagDateTimeOriginal
        ' Date et heure du cliche
        lPropertyItem.Length = 20
        lPropertyItem.Type = PropertyTagTypeASCII
        If Not IsDate(pValue) Then SetExifData = False: Exit Function
        lResultStr = Format(Year(pValue), "0000") & ":" & Format(Month(pValue), "00") & ":" & Format(Day(pValue), "00") _
                   & Chr(&H20) & Format(Hour(pValue), "00") & ":" & Format(Minute(pValue), "00") & ":" & Format(Second(pValue), "00") & Chr(0)
        RtlMoveMemory lPropertyItem.value, lResultStr, LenB(lPtr)
    Case TagImageDescription, TagArtist
        ' Description de l'image
        lPropertyItem.Length = Len(pValue) + 1
        lPropertyItem.Type = PropertyTagTypeASCII
        ' On ajoute le caractere de fin de texte
        lResultStr = pValue & Chr(0)
        RtlMoveMemory lPropertyItem.value, lResultStr, LenB(lPtr)
    Case Else
        SetExifData = False
    End Select
    ' Mise a jour du tag en memoire
    SetExifData = (GdipSetPropertyItem(gBitmap, lPropertyItem) = 0)
    If lNeedToFree Then GdipFree lPropertyItem.value
End If

gestion_erreurs:
If Err.Number <> 0 Then SetExifData = False
End Function

'---------------------------------------------------------------------------------------
' Extrait la miniature dans un bitmap Gdi+
'---------------------------------------------------------------------------------------
Private Function GetThumbNail(pFile As String)
    Dim lret As Long, lPropCount As Long
    Dim lPropSize As Long
    Dim lPropertyItem As PropertyItem
    Dim lBitmap, lhMem, lPtr
    On Error GoTo gestion_erreurs:
    ' Lecture fichier
    GdipCreateBitmapFromFile StrPtr(pFile), lBitmap
    ' Compte le nombre de proprietes Exif
    lret = GdipGetPropertyCount(lBitmap, lPropCount)
    If lPropCount > 0 Then
        ' Taille de la propriete
        If GdipGetPropertyItemSize(lBitmap, TagThumbnailData, lPropSize) = 0 Then
            If lPropSize > 0 Then
                ' Lecture de la propriete
                ReDim lBuffer(lPropSize - 1) As Byte
                If GdipGetPropertyItem(lBitmap, TagThumbnailData, lPropSize, lBuffer(0)) = 0 Then
                    Call RtlMoveMemory(lPropertyItem, lBuffer(0), LenB(lPropertyItem))
                    ' tPropertyItem.length correspond a la taille de la propriete
                    If lPropertyItem.Length > 0 Then
                        ' Miniature
                        Dim stream As Object
                        Dim lExifThumb
                        lhMem = GlobalAlloc(GMEM_MOVEABLE, lPropertyItem.Length)
                        If lhMem Then
                            lPtr = GlobalLock(lhMem)
                            If lPtr Then
                                RtlMoveMemory ByVal lPtr, ByVal lPropertyItem.value, lPropertyItem.Length
                                GlobalUnlock lhMem
                                ' Transformation du tableau en flux (stream)
                                If CreateStreamOnHGlobal(ByVal lhMem, True, stream) = 0 Then
                                    ' Lecture du flux dans un bitmap GDI+
                                    If GdipLoadImageFromStream(stream, lExifThumb) = 0 Then
                                        GetThumbNail = lExifThumb
                                    End If
                                    Set stream = Nothing
                                End If
                            End If
                            'GlobalFree lhMem
                        End If
                    End If
                End If
            End If
        End If
    End If

gestion_erreurs:
    GdipDisposeImage lBitmap
    If Err.Number <> 0 Then GetThumbNail = 0
End Function

'---------------------------------------------------------------------------------------
' Genere une miniature integree a l'image
'---------------------------------------------------------------------------------------
' pSize      : Taille maxi de la miniature
' pQuality   : Qualite de la miniature (0 a 100)
'---------------------------------------------------------------------------------------
Public Function GenerateThumbnail(Optional pSize As Long = 200, Optional pQuality As Long = -1) As Boolean
    Dim lBitmap, lNewBitmap
    Dim lWidth As Single, lHeight As Single
    Dim lNewWidth As Long, lNewHeight As Long
    Dim lgraphics, lNewGraphics
    Dim lGuid As GUID
    Dim lstream As Object
    Dim lhGlobal, lSize, lPtr
    Dim lPropertyItem As PropertyItem
    Dim lRes As Single
    Dim lParams As EncoderParameters
    Dim lret As Long
    On Error GoTo gestion_erreurs
    ' On travaille sur l'image avec graphics
    lBitmap = CloneBppGraphics(gBitmap)
    ' Lecture des dimensions de l'image
    GdipGetImageDimension lBitmap, lWidth, lHeight
    ' Taille de la miniature = pSize pixels sur sa plus grande dimension
    If lWidth > lHeight Then
        lNewHeight = pSize / lWidth * lHeight
        lNewWidth = pSize
    Else
        lNewWidth = pSize / lHeight * lWidth
        lNewHeight = pSize
    End If
    ' Creer un bitmap miniature
    GdipGetImageGraphicsContext lBitmap, lgraphics
    GdipCreateBitmapFromGraphics lNewWidth, lNewHeight, lgraphics, lNewBitmap
    GdipBitmapSetResolution lNewBitmap, 72, 72
    GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
    UpdateGraphics lNewGraphics
    'GdipSetInterpolationMode lNewGraphics, 0
    GdipDrawImageRectI lNewGraphics, lBitmap, 0, 0, lNewWidth, lNewHeight
    ' Supprime les objets graphics
    GdipDeleteGraphics lgraphics
    GdipDeleteGraphics lNewGraphics
    ' Recherche du Guid pour encodeur Jpeg
    CLSIDFromString StrPtr(gJpegEncoderStr), lGuid
    ' Parametre de l'encodeur Jpeg
    If pQuality <> -1 Then
        lParams.count = 1
        With lParams.Parameter(0)
            ' Parametrage de la qualite (0-100)
            CLSIDFromString StrPtr(gQualityParamStr), .GUID
            .NumberOfValues = 1
            .Type = 4    ' Type Long
            .value = VarPtr(pQuality)
        End With
        lParams = lParams
    End If
    ' Cree un flux vide
    If CreateStreamOnHGlobal(ByVal 0, True, lstream) = 0 Then
        ' Sauvegarde l'image miniature dans le flux
        If lParams.count > 0 Then
            lret = GdipSaveImageToStream(ByVal lNewBitmap, lstream, lGuid, lParams)
        Else
            lret = GdipSaveImageToStream(ByVal lNewBitmap, lstream, lGuid, ByVal PtrNull)
        End If
        If lret = 0 Then
            ' Recherche l'emplacement memoire du flux
            If GetHGlobalFromStream(lstream, lhGlobal) = 0 Then
                ' Lecture de la taille du flux en memoire
                lSize = GlobalSize(lhGlobal)
                ' Pointeur vers l'emplacement memoire
                lPtr = GlobalLock(lhGlobal)
                If lPtr Then
                    ' Propriete pour donnees de la miniature
                    lPropertyItem.id = &H501B&
                    lPropertyItem.Length = CLng(lSize)
                    lPropertyItem.Type = 1    'Byte
                    lPropertyItem.value = lPtr
                    ' Affecte la miniature a l'image
                    GenerateThumbnail = (GdipSetPropertyItem(gBitmap, lPropertyItem) = 0)
                End If
                ' Libere la memoire
                GlobalUnlock lhGlobal
            End If
        End If
        ' Libere le flux
        Set lstream = Nothing
        ' Unite de mesure pour resolution d'image de la miniature
        lPropertyItem.id = &H5030&
        lPropertyItem.Length = 2
        lPropertyItem.Type = 3
        ReDim lBuffer(lPropertyItem.Length - 1) As Byte
        lBuffer(0) = 2    ' = Points par pouce
        lPropertyItem.value = GdipAlloc(lPropertyItem.Length)
        Call RtlMoveMemory(ByVal lPropertyItem.value, lBuffer(0), lPropertyItem.Length)
        Call GdipSetPropertyItem(gBitmap, lPropertyItem)
        GdipFree lPropertyItem.value
        ' Resolution horizontale de la miniature
        lPropertyItem.id = &H502D&
        lPropertyItem.Length = 8    ' 8 = tableau de 2 long
        lPropertyItem.Type = 5    ' Type rationnel
        ReDim lBuffer(lPropertyItem.Length - 1) As Byte
        ReDim lbuffervalue(1) As Long
        GdipGetImageHorizontalResolution lNewBitmap, lRes
        lbuffervalue(0) = lRes
        lbuffervalue(1) = 1
        RtlMoveMemory lBuffer(0), lbuffervalue(0), 8
        lPropertyItem.value = GdipAlloc(lPropertyItem.Length)
        Call RtlMoveMemory(ByVal lPropertyItem.value, lBuffer(0), lPropertyItem.Length)
        Call GdipSetPropertyItem(gBitmap, lPropertyItem)
        GdipFree lPropertyItem.value
        ' Resolution horizontale de la miniature
        lPropertyItem.id = &H502E&
        lPropertyItem.Length = 8    '8 = tableau de 2 long
        lPropertyItem.Type = 5    ' Type rationnel
        ReDim lBuffer(lPropertyItem.Length - 1) As Byte
        ReDim lbuffervalue(1) As Long
        GdipGetImageVerticalResolution lNewBitmap, lRes
        lbuffervalue(0) = lRes
        lbuffervalue(1) = 1
        RtlMoveMemory lBuffer(0), lbuffervalue(0), 8
        lPropertyItem.value = GdipAlloc(lPropertyItem.Length)
        Call RtlMoveMemory(ByVal lPropertyItem.value, lBuffer(0), lPropertyItem.Length)
        Call GdipSetPropertyItem(gBitmap, lPropertyItem)
        GdipFree lPropertyItem.value
        ' Compression jpeg
        lPropertyItem.id = &H5023&
        lPropertyItem.Length = 2
        lPropertyItem.Type = 3 ' 3 = Entier
        ReDim lBuffer(lPropertyItem.Length - 1) As Byte
        lBuffer(0) = 6 ' = jpeg avec compression
        lBuffer(1) = 0
        lPropertyItem.value = GdipAlloc(lPropertyItem.Length)
        Call RtlMoveMemory(ByVal lPropertyItem.value, lBuffer(0), lPropertyItem.Length)
        Call GdipSetPropertyItem(gBitmap, lPropertyItem)
        GdipFree lPropertyItem.value
        ' Supprime l'image miniature en memoire
        GdipDisposeImage lNewBitmap
    End If
gestion_erreurs:
    If lBitmap <> gBitmap Then GdipDisposeImage lBitmap
    If Err.Number <> 0 Then GenerateThumbnail = False
End Function

'---------------------------------------------------------------------------------------
' Importe les donnees Exif d'un bitmap
'---------------------------------------------------------------------------------------
' pFromBitmap  : Bitmap source
' ptoBitmap    : Bitmap cible
' pExcludeTags : Tableau de tags Exif a exclure
'---------------------------------------------------------------------------------------
Private Function ImportExifDataFromBitmap(pFromBitmap, pToBitmap, Optional ByRef pExcludeTags As Variant) As Boolean
    Dim lPropCount As Long
    Dim lBuffer() As Byte
    Dim lcpt As Integer
    Dim lCptTag As Integer
    Dim lBitmap
    Dim lBufferSize As Long
    Dim lItem As PropertyItem
    Dim lExclude As Boolean
    On Error GoTo gestion_erreurs:
    If gBitmap = 0 Then GoTo gestion_erreurs
    lBitmap = pFromBitmap
    ' Lecture de la taille des proprietes Exif
    GdipGetPropertySize lBitmap, lBufferSize, lPropCount
    If lPropCount > 0 Then
        ' Redimensionne le buffer
        ReDim lBuffer(lBufferSize)
        ' Lecture de toutes les donnees Exif dans le buffer
        If GdipGetAllPropertyItems(lBitmap, lBufferSize, lPropCount, lBuffer(0)) = 0 Then
            ' Recherche chaque donnee
            For lcpt = 0 To lPropCount - 1
                ' Deplace la donnee dans une structure PropertyItem
                RtlMoveMemory lItem, lBuffer(lcpt * LenB(lItem)), LenB(lItem)
                ' Verifie si le tag est a exclure
                lExclude = False
                If isArray(pExcludeTags) Then
                    For lCptTag = LBound(pExcludeTags) To UBound(pExcludeTags)
                        If lItem.id = pExcludeTags(lCptTag) Then lExclude = True: Exit For
                    Next
                End If
                ' Met a jour la donnee si non exclue
                If Not lExclude Then Call GdipSetPropertyItem(pToBitmap, lItem)
            Next
            ' Retourne vrai si pas d'erreur
            ImportExifDataFromBitmap = True
        End If
    End If
gestion_erreurs:
    If Err.Number <> 0 Then ImportExifDataFromBitmap = False
End Function

'---------------------------------------------------------------------------------------
' Importe les donnees Exif d'un fichier dans l'image en cours
'---------------------------------------------------------------------------------------
' pFile     : Nom du fichier
' pExcludeTags : Tableau de tags Exif a exclure
'---------------------------------------------------------------------------------------
Public Function ImportExifData(ByVal pFile As String, Optional ByRef pExcludeTags As Variant) As Boolean
    Dim lBitmap
    ' Charge l'image a partir du fichier
    If GdipCreateBitmapFromFile(StrPtr(pFile), lBitmap) = 0 Then
        ImportExifData = ImportExifDataFromBitmap(lBitmap, gBitmap, pExcludeTags)
    End If
    ' Supprime l'image temporarire de la memoire
    If lBitmap <> 0 Then GdipDisposeImage lBitmap
End Function

'***************************************************************************************
'*                          FONCTIONS POUR CURSEURS                                    *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Applique un curseur
'---------------------------------------------------------------------------------------
Public Function CursorSet(pName As String) As Boolean
On Error GoTo gestion_erreurs
SetCursor gCurseurs(pName)
gestion_erreurs:
    ' renvoie Vrai si pas d'erreur
    If Err.Number = 0 Then CursorSet = True
End Function
'---------------------------------------------------------------------------------------
' Ajoute un curseur à partir d'une image
'---------------------------------------------------------------------------------------
Public Function CursorAddFromImg(pName As String, pImg As Variant, pHotspotX As Long, pHotspotY As Long) As Boolean
Dim lhico
Dim lhico2
Dim linfo As ICONINFO
Dim lImg As clGdiplus
On Error GoTo gestion_erreurs
Set lImg = ImgFromVariant(pImg)
If GdipCreateHICONFromBitmap(lImg.SaveBitmap, lhico) = 0 Then
    If GetIconInfo(lhico, linfo) <> 0 Then
        linfo.xHotspot = pHotspotX
        linfo.yHotspot = pHotspotY
        linfo.fIcon = 0
        lhico2 = CreateIconIndirect(linfo)
        DeleteObject lhico
        gCurseurs.Add lhico2, pName
        CursorAddFromImg = True
    End If
End If
gestion_erreurs:
If Err.Number = 457 Then CursorDelete pName: Resume
CursorAddFromImg = False
End Function
'---------------------------------------------------------------------------------------
' Ajoute un curseur à partir d'un id standard
'---------------------------------------------------------------------------------------
Public Function CursorAddFromId(pName As String, pId As EGdipCursors) As Boolean
Dim lhico
On Error GoTo gestion_erreurs
lhico = LoadCursor(PtrNull, pId)
If lhico <> 0 Then
    gCurseurs.Add lhico, pName
    CursorAddFromId = True
End If
gestion_erreurs:
If Err.Number = 457 Then CursorDelete pName: Resume
CursorAddFromId = False
End Function
'---------------------------------------------------------------------------------------
' Ajoute un curseur à partir d'un fichier
'---------------------------------------------------------------------------------------
Public Function CursorAddFromFile(pName As String, pFile As String) As Boolean
Dim lhico
On Error GoTo gestion_erreurs
lhico = LoadCursorFromFile(pFile)
If lhico <> 0 Then
    gCurseurs.Add lhico, pName
    CursorAddFromFile = True
End If
gestion_erreurs:
If Err.Number = 457 Then CursorDelete pName: Resume
CursorAddFromFile = False
End Function
'---------------------------------------------------------------------------------------
' Supprime un curseur
'---------------------------------------------------------------------------------------
Public Function CursorDelete(pName As String) As Boolean
    On Error GoTo gestion_erreurs
    ' Supprime le curseur gdi
    DeleteObject gCurseurs(pName)
    ' Retire de la collection
    gCurseurs.Remove pName
gestion_erreurs:
    ' renvoie Vrai si pas d'erreur
    If Err.Number = 0 Then CursorDelete = True
End Function

'***************************************************************************************
'*                          FONCTIONS POUR POLICES DE CARACTERES                       *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Charge une police de caracteres depuis un fichier
'---------------------------------------------------------------------------------------
Public Function FontPrivateAddFromFile(pFile As String) As Boolean
    Dim lFontFile As String
    On Error GoTo gestion_erreurs
    If gPrivateFontCollection = 0 Then
        Call GdipNewPrivateFontCollection(gPrivateFontCollection)
    End If
    If gPrivateFontCollection <> 0 Then
        lFontFile = pFile
        If GdipPrivateAddFontFile(gPrivateFontCollection, StrPtr(lFontFile)) = 0 Then
            FontPrivateAddFromFile = True
        End If
    End If
    Exit Function
gestion_erreurs:
    FontPrivateAddFromFile = False
End Function
'---------------------------------------------------------------------------------------
' Charge une police de caracteres depuis un tableau de byte
'---------------------------------------------------------------------------------------
Public Function FontPrivateAddFromByte(pByte() As Byte) As Boolean
    Dim lByte() As Byte
    On Error GoTo gestion_erreurs
    If gPrivateFontCollection = 0 Then
        Call GdipNewPrivateFontCollection(gPrivateFontCollection)
    End If
    If gPrivateFontCollection <> 0 Then
        lByte = pByte
        If GdipPrivateAddMemoryFont(gPrivateFontCollection, VarPtr(lByte(LBound(lByte))), UBound(lByte) - LBound(lByte) + 1) = 0 Then
            FontPrivateAddFromByte = True
        End If
    End If
    Exit Function
gestion_erreurs:
    FontPrivateAddFromByte = False
End Function
'---------------------------------------------------------------------------------------
' Supprime les polices de caracteres privees
'---------------------------------------------------------------------------------------
Public Function FontsPrivateDelete() As Boolean
    On Error GoTo gestion_erreurs
    If gPrivateFontCollection <> 0 Then
        FontsPrivateDelete = (GdipDeletePrivateFontCollection(gPrivateFontCollection) = 0)
    Else
        FontsPrivateDelete = True
    End If
    Exit Function
gestion_erreurs:
    FontsPrivateDelete = False
End Function
'---------------------------------------------------------------------------------------
' Enumere les polices de caracteres
'---------------------------------------------------------------------------------------
Public Sub FontsEnumerate(pFonts As Collection)
    Dim lNumFound As Long
    Dim lcpt As Long
    Dim lFamilies()
    Dim lFonts As Collection
    Dim lFontCollection
    Dim lName As String
    Dim lStep As Long
    On Error GoTo gestion_erreurs
    Set lFonts = New Collection
    For lStep = 1 To 2
        If lStep = 1 Then
            GdipNewInstalledFontCollection lFontCollection
        Else
            lFontCollection = gPrivateFontCollection
        End If
        GdipGetFontCollectionFamilyCount lFontCollection, lNumFound
        ReDim lFamilies(1 To lNumFound)
        GdipGetFontCollectionFamilyList lFontCollection, lNumFound, lFamilies(1), lNumFound
        For lcpt = 1 To lNumFound
            lName = Space(255)
            GdipGetFamilyName lFamilies(lcpt), StrPtr(lName), ByVal 0
            lFonts.Add Left(lName, InStr(lName, Chr$(0)) - 1)
        Next
    Next
    Set pFonts = lFonts
    Exit Sub
gestion_erreurs:
    Set pFonts = lFonts
End Sub
'***************************************************************************************
'*                                    FONCTIONS                                        *
'***************************************************************************************
'--------------------------------------------------------------------------------------
' Teste si le chemin est celui d'une image valide
'---------------------------------------------------------------------------------------
Public Function PathIsImage(pPath As String) As Boolean
Dim lBitmap
    On Error GoTo gestion_erreurs
    PathIsImage = (GdipLoadImageFromFile(StrPtr(pPath), lBitmap) = 0)
    If lBitmap <> 0 Then GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    PathIsImage = False
End Function
'--------------------------------------------------------------------------------------
' Bitmap gdiplus de l'image
'---------------------------------------------------------------------------------------
Public Function SaveBitmap(Optional pClone As Boolean = False)
    If pClone Then
        SaveBitmap = PrivBitmapClone(gBitmap)
        'GdipCloneImage gBitmap, SaveBitmap
    Else
        SaveBitmap = gBitmap
    End If
End Function
Public Function LoadBitmap(pBitmap, Optional pClone As Boolean = True) As Boolean
Dim lBitmap
    On Error GoTo gestion_erreurs
    If pClone Then
        lBitmap = PrivBitmapClone(pBitmap)
    Else
        lBitmap = pBitmap
    End If
    PrivBitmap(gUseGdiImg) = lBitmap
    LoadBitmap = (gBitmap <> 0)
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    LoadBitmap = False
End Function

'---------------------------------------------------------------------------------------
' Cree une image a partir d'un bitmap et d'un mask gdi32
'---------------------------------------------------------------------------------------
Public Function LoadPictureMask(pPicture, Optional pMask = PtrNull, Optional pBackColor As Long = vbWhite, Optional pBackAlpha As Long = 0, Optional pDelete As Boolean = True) As Boolean
    Dim lOldBmp, lHdc
    Dim lOldMaskBmp, lMaskhdc
    Dim lWidth As Long, lHeight As Long
    Dim lNewBrush
    Dim lImage, lgraphics, lImageDc
    Dim lBmp As bitmap
    Dim lBitmap
    On Error GoTo gestion_erreurs
    ' Creation image
    #If Win64 Then
    Call GetObjectBmp(pPicture, LenB(lBmp), lBmp)
    #Else
    Call GetObjectBmp(pPicture, Len(lBmp), lBmp)
    #End If
    lWidth = lBmp.bmWidth
    lHeight = lBmp.bmHeight
    GdipCreateBitmapFromScan0 lWidth, lHeight, 0, PixelFormat32bppARGB, ByVal PtrNull, lImage
    ' Lecture du contexte graphique
    GdipGetImageGraphicsContext lImage, lgraphics
    ' Remplissage du fond
    GdipCreateSolidFill GdipColor(pBackAlpha, GetColor(pBackColor)), lNewBrush
    GdipFillRectangleI lgraphics, lNewBrush, 0, 0, lWidth, lHeight
    GdipDeleteBrush lNewBrush
    If pMask = PtrNull Then
        ' Utilisation de GetDIBBits pour transférer à gdi+ les pixels avec alpha
        Dim lPixels() As Byte
        Dim lBmpInfo As BitmapInfo
        Dim lcpt As Long
        Dim lHasAlpha As Boolean
        With lBmpInfo.bmiHeader
            .biBitCount = 32
            .biCompression = BI_RGB
            .biPlanes = 1
            #If Win64 Then
            .biSize = LenB(lBmpInfo.bmiHeader)
            #Else
            .biSize = Len(lBmpInfo.bmiHeader)
            #End If
            .biWidth = lWidth
            .biHeight = -lHeight  ' On retourne l'image pour parcourir les pixels de haut en bas
            .biSizeImage = Abs(.biHeight) * (((.biWidth * .biBitCount) + 31) \ 32) * 4
            ReDim lPixels(1 To .biSizeImage)
        End With
        lHdc = CreateCompatibleDC(lImageDc)
        lOldBmp = SelectObject(lHdc, pPicture)
        GetDIBits lHdc, pPicture, 0, lHeight, lPixels(1), lBmpInfo, DIB_RGB_COLORS
        SelectObject lHdc, lOldBmp
        DeleteDC lHdc
        For lcpt = 4 To lBmpInfo.bmiHeader.biSizeImage Step 4
            If lPixels(lcpt) <> 0 Then lHasAlpha = True: Exit For
        Next
        GdipCreateBitmapFromScan0 lWidth, lHeight, Abs(lBmpInfo.bmiHeader.biSizeImage / lBmpInfo.bmiHeader.biHeight), IIf(lHasAlpha, PixelFormat32bppARGB, PixelFormat32bppRGB), ByVal VarPtr(lPixels(1)), lBitmap
        GdipDrawImageRectI lgraphics, lBitmap, 0, 0, lWidth, lHeight
        GdipDisposeImage lBitmap
    Else
         ' Lecture du DC
        GdipGetDC lgraphics, lImageDc
        ' Chargement des objets gdi32
        lHdc = CreateCompatibleDC(lImageDc)
        lOldBmp = SelectObject(lHdc, pPicture)
        If pMask <> PtrNull Then
            lMaskhdc = CreateCompatibleDC(lImageDc)
            lOldMaskBmp = SelectObject(lMaskhdc, pMask)
        End If
        ' Dessine le bitmap et le mask
        BitBlt lImageDc, 0, 0, lWidth, lHeight, lMaskhdc, 0, 0, SRCAND
        BitBlt lImageDc, 0, 0, lWidth, lHeight, lHdc, 0, 0, SRCPAINT
        ' Supprime les objets
        SelectObject lHdc, lOldBmp
        DeleteDC lHdc
        SelectObject lMaskhdc, lOldMaskBmp
        DeleteDC lMaskhdc
        GdipReleaseDC lgraphics, lImageDc
    End If
    If pDelete Then
        DeleteObject pPicture
        DeleteObject pMask
    End If
    GdipDeleteGraphics lgraphics
    LoadPictureMask = (lImage <> 0)
    PrivBitmap(gUseGdiImg) = lImage
gestion_erreurs:
    ' Retourne vrai si pas d'erreur
    LoadPictureMask = (LoadPictureMask And Err.Number = 0)
End Function
'---------------------------------------------------------------------------------------
' Image sous forme de tableau
'---------------------------------------------------------------------------------------
Public Function SaveArray(ByRef pArray() As Byte, Optional pFormat As EGdipImageFormatSaveArray = gdipFormatSaveArrayJPEG, Optional ByVal pQuality As Integer = -1) As Boolean
    Dim lGuid As GUID
    Dim lstream As Object
    Dim lhGlobal
    Dim lSize
    Dim lPtr
    Dim lParams As EncoderParameters
    Dim lret As Long
    Dim lEncoderStr As String

    On Error GoTo gestion_erreurs

    Erase pArray
    ' Format de l'encodeur
    Select Case pFormat
    Case gdipFormatSaveArrayJPEG
        lEncoderStr = gJpegEncoderStr
    Case gdipFormatSaveArrayGIF
        lEncoderStr = gGifEncoderStr
    Case gdipFormatSaveArrayBMP
        lEncoderStr = gBmpEncoderStr
    Case gdipFormatSaveArrayPNG
        lEncoderStr = gPngEncoderStr
    Case gdipFormatSaveArrayTIFF
        lEncoderStr = gTifEncoderStr
    Case gdipFormatSaveArrayPictureDataBMP
        pArray = PrivGdiPlusToPictureData(Nothing, False, False)
        Exit Function
    Case gdipFormatSaveArrayPictureDataEMF
        pArray = PrivGdiPlusToPictureData(Nothing, False, True)
        Exit Function
    End Select
    ' Recherche du Guid pour encodeur Jpeg
    CLSIDFromString StrPtr(lEncoderStr), lGuid
    ' Parametre de l'encodeur Jpeg
    If pQuality <> -1 Then
        lParams.count = 1
        With lParams.Parameter(0)
            ' Parametrage de la qualite (0-100)
            CLSIDFromString StrPtr(gQualityParamStr), .GUID
            .NumberOfValues = 1
            .Type = 4    ' Type Long
            .value = VarPtr(pQuality)
        End With
        lParams = lParams
    End If
    ' Cree un flux vide
    If CreateStreamOnHGlobal(ByVal 0, False, lstream) = 0 Then
        ' Sauvegarde l'image dans le flux
        If lParams.count > 0 Then
            lret = GdipSaveImageToStream(ByVal gBitmap, lstream, lGuid, lParams)
        Else
            lret = GdipSaveImageToStream(ByVal gBitmap, lstream, lGuid, ByVal PtrNull)
        End If
        If lret = 0 Then
            ' Recherche l'emplacement memoire du flux
            If GetHGlobalFromStream(lstream, lhGlobal) = 0 Then
                ' Lecture de la taille du flux en memoire
                lSize = GlobalSize(lhGlobal)
                ' Pointeur vers l'emplacement memoire
                lPtr = GlobalLock(lhGlobal)
                If lPtr Then
                    ReDim pArray(1 To CLng(lSize))
                    RtlMoveMemory pArray(1), ByVal lPtr, lSize
                End If
                ' Libere la memoire
                GlobalUnlock lhGlobal
                SaveArray = True
            End If
        End If
        ' Libere le flux
        Set lstream = Nothing
    End If
    On Error GoTo 0
    Exit Function
gestion_erreurs:
     SaveArray = False
    Erase pArray
End Function
Public Function LoadArray(ByRef pArray() As Byte, Optional pAttachment As Boolean) As Boolean
Dim lBitmap
    On Error GoTo gestion_erreurs
    lBitmap = GetBitmapFromArray(pArray, pAttachment)
    If lBitmap <> 0 Then
        PrivBitmap(gUseGdiImg) = lBitmap
        LoadArray = True
    End If
    Exit Function
gestion_erreurs:
    LoadArray = False
End Function

Private Function GetBitmapFromArray(ByRef pArray() As Byte, Optional pAttachment As Boolean)
    On Error GoTo gestion_erreurs
    Dim stream As Object
    Dim lBitmap
    Dim lhMem
    Dim lPtr
    Dim lSize
    Dim lData() As Byte
    If pAttachment Then
        ReDim lData(1 To UBound(pArray) - LBound(pArray) + 1 - pArray(LBound(pArray)))
        RtlMoveMemory lData(1), pArray(LBound(pArray) + pArray(LBound(pArray))), UBound(lData)
    Else
        lData = pArray
    End If
    lSize = UBound(lData) - LBound(lData) + 1
    lhMem = GlobalAlloc(GMEM_MOVEABLE, lSize)
    If lhMem Then
        lPtr = GlobalLock(lhMem)
        If lPtr Then
            RtlMoveMemory ByVal lPtr, lData(LBound(lData)), lSize
            GlobalUnlock lhMem
            ' Transformation du tableau en flux (stream)
            If CreateStreamOnHGlobal(ByVal lhMem, True, stream) = 0 Then
                ' Lecture du flux dans un bitmap GDI+
                If GdipLoadImageFromStream(stream, lBitmap) = 0 Then
                    If ConvertMetaFile Then ConvertBitmapMetafile lBitmap
                    GetBitmapFromArray = lBitmap
                End If
                Set stream = Nothing
            End If
        End If
    End If
    If GetBitmapFromArray = 0 Then GetBitmapFromArray = PrivLoadBitmapFromPictureData(pArray)
    Exit Function
gestion_erreurs:
    GetBitmapFromArray = 0
End Function

Private Function ConvertBitmapMetafile(pBitmap)
    Dim lFormat As Long
    Dim lLimit As Single
    lFormat = PrivImageFormat(pBitmap)
    If lFormat = 3 Or lFormat = 4 Then ' WMF, EMF
        Dim lNewBitmap1, lNewBitmap2, lNewGraphics
        Dim lWidth As Long, lHeight As Long
        Dim lActualWidth As Single, lActualHeight  As Single
        lLimit = 200 ' Limite 200 Mo
        GdipGetImageWidth pBitmap, lWidth
        GdipGetImageHeight pBitmap, lHeight
        GdipGetImageDimension pBitmap, lActualWidth, lActualHeight
retry:
        Do
            If lActualWidth * lActualHeight * 4 / 1024 / 1024 > lLimit Then
                lActualWidth = lActualWidth * 0.8
                lActualHeight = lActualHeight * 0.8
            Else
                Exit Do
            End If
        Loop
        GdipCreateBitmapFromScan0 lActualWidth, lActualHeight, 0, PixelFormat32bppARGB, ByVal PtrNull, lNewBitmap1
        GdipGetImageGraphicsContext lNewBitmap1, lNewGraphics
        GdipSetInterpolationMode lNewGraphics, 2
        GdipSetSmoothingMode lNewGraphics, 5
        GdipSetTextRenderingHint lNewGraphics, GdipHintAntiAlias
        If GdipDrawImageRectI(lNewGraphics, pBitmap, 0, 0, lActualWidth, lActualHeight) = 2 Then
            ' Trop grop => on abaisse la limite et on retente
            GdipDeleteGraphics lNewGraphics
            GdipDisposeImage lNewBitmap1
            lLimit = lLimit * 0.8
            If lLimit < 10 Then Exit Function
            GoTo retry
        End If
        GdipDeleteGraphics lNewGraphics
        GdipCreateBitmapFromScan0 lWidth, lHeight, 0, PixelFormat32bppARGB, ByVal PtrNull, lNewBitmap2
        GdipGetImageGraphicsContext lNewBitmap2, lNewGraphics
        GdipSetInterpolationMode lNewGraphics, 2
        GdipDrawImageRectI lNewGraphics, lNewBitmap1, 0, 0, lWidth, lHeight
        GdipDisposeImage lNewBitmap1
        GdipDeleteGraphics lNewGraphics
        GdipDisposeImage pBitmap
        pBitmap = lNewBitmap2
        
'        GdipCreateBitmapFromScan0 lWidth, lHeight, 0, PixelFormat32bppARGB, ByVal PtrNull, lNewBitmap1
'        GdipGetImageGraphicsContext lNewBitmap1, lNewGraphics
'        GdipSetInterpolationMode lNewGraphics, 2
'        GdipSetSmoothingMode lNewGraphics, 5
'        GdipSetTextRenderingHint lNewGraphics, GdipHintAntiAlias
'        Debug.Print "draw:" & GdipDrawImageRectRectI(lNewGraphics, pBitmap, 0, 0, lWidth, lHeight, 0, 0, lWidth, lHeight, 0, ByVal PtrNull, ByVal PtrNull, ByVal PtrNull)
'        GdipDeleteGraphics lNewGraphics
'        GdipDisposeImage pBitmap
'        pBitmap = lNewBitmap1
    Else
        pBitmap = pBitmap
    End If
End Function

'---------------------------------------------------------------------------------------
' Renvoit un objet IPictureDisp
'---------------------------------------------------------------------------------------
Public Function SaveIPictureDisp(Optional pBackgroundColor As Long = -1) As Object
    Set SaveIPictureDisp = PrivGdiPlusToIPictureDisp(, , , pBackgroundColor)
End Function

'---------------------------------------------------------------------------------------
' Creation d'un bitmap a partir de l'image contenue dans un controle
' Fonction Privee : renvoi un bitmap gdi+
' https://msdn.microsoft.com/en-us/library/windows/desktop/dd145119(v=vs.85).aspx
'---------------------------------------------------------------------------------------
Private Function PrivLoadBitmapFromPictureData(pData As Variant)
    Dim lBuffer() As Byte
    Dim lBH As BitmapInfoHeader
    Dim lBitmap
    Dim llBound As Long
    Dim lFormat As Long
    Dim lhEMF
    Dim lmfh As ENHMETAHEADER
    Dim lgraphics
    Dim lHdc
    Dim LenRGBQuad As Long
    Dim lPalette As ColorPalette
    Dim lClrBits  As Long
    Dim lHasAlpha As Boolean
    Const lLong16777216 As Long = 16777216
    Dim i As Long
    On Error GoTo gestion_erreurs
    ' Lecture donnees de l'image dans PictureData
    lBuffer = pData
    ' Premier indice du tableau
    llBound = LBound(lBuffer)
    ' Le premier byte contient le type d'image
    Select Case lBuffer(llBound)
    Case CF_BITMAP    ' Bitmap
        ' En-tete du bitmap
        #If Win64 Then
        RtlMoveMemory lBH, lBuffer(llBound), LenB(lBH)
        #Else
        RtlMoveMemory lBH, lBuffer(llBound), Len(lBH)
        #End If
        ' Format = nombre de bits (24, 32, ...)
        lClrBits = lBH.biPlanes * lBH.biBitCount
        If lClrBits = 1 Then
            lClrBits = 1
        ElseIf lClrBits <= 4 Then
            lClrBits = 4
        ElseIf lClrBits <= 8 Then
            lClrBits = 8
        ElseIf lClrBits <= 16 Then
            lClrBits = 16
        ElseIf lClrBits <= 24 Then
            lClrBits = 24
        Else
            lClrBits = 32
        End If
        Select Case lClrBits
        Case 1:     lFormat = PixelFormat1bppIndexed
        Case 4:     lFormat = PixelFormat4bppIndexed
        Case 8:     lFormat = PixelFormat8bppIndexed
        Case 16:    lFormat = PixelFormat16bppRGB555
        Case 24:    lFormat = PixelFormat24bppRGB
        Case 32:    lFormat = PixelFormat32bppPARGB
        Case Else:  lFormat = 0
        End Select
        ' Palette de couleurs
        If lClrBits >= 24 Then
            LenRGBQuad = 0
        Else
            lBH.biClrUsed = 2 ^ lClrBits
            LenRGBQuad = 4 * lBH.biClrUsed
        End If
        If lFormat <> 0 Then
            If lBH.biSizeImage = 0 Then
                lBH.biSizeImage = Abs(lBH.biHeight) * (((lBH.biWidth * lBH.biBitCount) + 31) \ 32) * 4
            End If
            ' Creation de l'image a partir des donnees du buffer
            GdipCreateBitmapFromScan0 lBH.biWidth, lBH.biHeight, lBH.biSizeImage / lBH.biHeight, lFormat, ByVal VarPtr(lBuffer(40 + llBound + LenRGBQuad)), lBitmap
            ' Retourne l'image car sinon l'image est la tete en bas
            Call GdipImageRotateFlip(lBitmap, 6)
            ' Mise a jour de la palette de couleurs
            If LenRGBQuad > 0 Then
                lPalette.count = lBH.biClrUsed
                RtlMoveMemory lPalette.Entries(0), ByVal VarPtr(lBuffer(40 + llBound)), LenRGBQuad
                ' S'assure d'avoir un canal alpha utilisé
                For i = LBound(lPalette.Entries) To UBound(lPalette.Entries)
                    If lPalette.Entries(i) And lLong16777216 Then lHasAlpha = True: Exit For
                Next
                If Not lHasAlpha Then
                    For i = LBound(lPalette.Entries) To UBound(lPalette.Entries)
                        'lPalette.Entries(i) = lPalette.Entries(i) Or lLong16777216
                        lPalette.Entries(i) = GdipColor(255, lPalette.Entries(i))
                    Next
                End If
                lPalette.flags = 1 ' canal alpha
                GdipSetImagePalette lBitmap, lPalette
            End If
        End If
        ' Renvoit le bitmap gdi+ cree
        PrivLoadBitmapFromPictureData = lBitmap
    Case CF_ENHMETAFILE    ' EMF
        ' Copie les donnees de l'image dans l'objet EMF gdi32
        lhEMF = SetEnhMetaFileBits(UBound(lBuffer) - llBound + 1 - 8, lBuffer(8 + llBound))
        ' Recupere l'en-tete de l'image
        #If Win64 Then
        GetEnhMetaFileHeader lhEMF, LenB(lmfh), lmfh
        #Else
        GetEnhMetaFileHeader lhEMF, Len(lmfh), lmfh
        #End If
        ' Conversion des coordonnees
        ' DC de reference
        Dim lHdcref
        lHdcref = GetDC(0)
        lmfh.rclBounds.Right = lmfh.rclFrame.Right / (HIMETRIC_INCH / GetDeviceCaps(lHdcref, LOGPIXELSX))
        lmfh.rclBounds.Bottom = lmfh.rclFrame.Bottom / (HIMETRIC_INCH / GetDeviceCaps(lHdcref, LOGPIXELSY))
        ' Libere le DC
        ReleaseDC 0, lHdcref
        ' Conversion taille image en pixels
        'lmfh.rclBounds.Right = PointsToPixelsX(pCtrl.ImageWidth)
        'lmfh.rclBounds.Bottom = PointsToPixelsY(pCtrl.ImageHeight)
        ' Cree un bitmap 32 bits vierge
        GdipCreateBitmapFromScan0 lmfh.rclBounds.Right - 1, lmfh.rclBounds.Bottom - 1, 0, PixelFormat32bppARGB, ByVal PtrNull, lBitmap
        If lBitmap <> 0 Then
            ' Lecture du contexte graphique et du DC
            GdipGetImageGraphicsContext lBitmap, lgraphics
            GdipGetDC lgraphics, lHdc
            ' Force le mode pixel
            SetMapMode lhEMF, MM_TEXT
            ' Dessine l'EMF dans le bitmap gdiplus par l'intermediaire du DC
            PlayEnhMetaFile lHdc, lhEMF, lmfh.rclBounds
            ' Libere les objets
            GdipReleaseDC lgraphics, lHdc
            GdipDeleteGraphics lgraphics
            DeleteEnhMetaFile lhEMF
            ' Renvoit le bitmap gdi+ cree
            PrivLoadBitmapFromPictureData = lBitmap
        End If
    Case Else    ' Non gere
        PrivLoadBitmapFromPictureData = GetBitmapFromArray(lBuffer)
    End Select
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    PrivLoadBitmapFromPictureData = 0
End Function
#If AppName = "A" Then
Private Function PrivLoadBitmapFromControl(pCtrl As Object)
PrivLoadBitmapFromControl = PrivLoadBitmapFromPictureData(pCtrl.PictureData)
End Function
#Else
Private Function PrivLoadBitmapFromControl(pCtrl As Object)
    Dim lBitmap
    On Error GoTo gestion_erreurs
    ' Transforme le bitmap gdi32 en un bitmap gdip+
    GdipCreateBitmapFromHBITMAP pCtrl.Picture, ByVal PtrNull, lBitmap
    PrivLoadBitmapFromControl = lBitmap
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    PrivLoadBitmapFromControl = 0
End Function
#End If

'---------------------------------------------------------------------------------------
' Creation d'un bitmap a partir de l'image contenue dans un controle
'---------------------------------------------------------------------------------------
' pCtrl : Controle Image
'---------------------------------------------------------------------------------------
Public Function LoadControl(pCtrl As Object) As Boolean
On Error GoTo gestion_erreurs
' Creation du bitmap
PrivBitmap(gUseGdiImg) = PrivLoadBitmapFromControl(pCtrl)
' Renvoit Vrai si bitmap cree
LoadControl = (gBitmap <> 0)
On Error GoTo 0
Exit Function
gestion_erreurs:
LoadControl = False
End Function

'---------------------------------------------------------------------------------------
' Clone de l'image en 32 bits
'---------------------------------------------------------------------------------------
' pImage   : Identifiant gdiplus de l'image
'---------------------------------------------------------------------------------------
' En convertissant l'image en 32 bits, on s'assure que l'image possede un objet Graphics
'---------------------------------------------------------------------------------------
Private Function CloneBppGraphics(ByVal PImage, Optional pForceAlpha As Boolean, Optional pForceNoAlpha As Boolean, Optional pForce32bppARGB As Boolean)
    Dim lPxfmt As Long, lnewpxfmt As Long
    Dim lWidth As Single, lHeight As Single
    Dim lNewImage, lgraphics
    Dim lResX As Single, lResY As Single
    Dim lHasAlpha As Boolean
    Dim lForceAlpha As Boolean
    lForceAlpha = (pForceAlpha Or pForceNoAlpha)
    GdipGetImagePixelFormat PImage, lPxfmt
    If pForce32bppARGB And lPxfmt <> PixelFormat32bppARGB Then
        lnewpxfmt = PixelFormat32bppARGB
    Else
        lHasAlpha = PxFmtAlpha(lPxfmt)
        If lForceAlpha Then ' Demande de forcage alpha
            If lHasAlpha And pForceNoAlpha Then
                lnewpxfmt = PixelFormat32bppRGB
            ElseIf Not lHasAlpha And pForceAlpha Then
                lnewpxfmt = PixelFormat32bppARGB
            End If
        Else ' Demande de graphics
            Select Case lPxfmt
                Case PixelFormat1bppIndexed, PixelFormat4bppIndexed, PixelFormat8bppIndexed, PixelFormat16bppGreyScale, _
                            PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat16bppARGB1555
                    lnewpxfmt = PixelFormat24bppRGB
                Case Else
            End Select
        End If
    End If
    If lnewpxfmt <> 0 Then
        GdipGetImageDimension PImage, lWidth, lHeight
        GdipCreateBitmapFromScan0 lWidth, lHeight, 0, lnewpxfmt, ByVal PtrNull, lNewImage
        GdipGetImageHorizontalResolution PImage, lResX
        GdipGetImageVerticalResolution PImage, lResY
        GdipBitmapSetResolution lNewImage, lResX, lResY
        GdipGetImageGraphicsContext lNewImage, lgraphics
        GdipDrawImageRectRectI lgraphics, PImage, 0, 0, lWidth, lHeight, 0, 0, lWidth, lHeight, 2
        GdipDeleteGraphics lgraphics
    End If
    If lNewImage <> 0 Then
        CloneBppGraphics = lNewImage
    Else
        CloneBppGraphics = PImage
    End If
End Function
'---------------------------------------------------------------------------------------
' Conversion de l'image en 32 bits
'---------------------------------------------------------------------------------------
' pImage   : Identifiant gdiplus de l'image
'---------------------------------------------------------------------------------------
' En convertissant l'image en 32 bits, on s'assure que l'image possede un objet Graphics
'---------------------------------------------------------------------------------------
Private Sub ConvertBppGraphics(ByRef PImage, Optional pForceAlpha As Boolean, Optional pForceNoAlpha As Boolean, Optional pForce32bppARGB As Boolean)
    Dim lOldImage
    Dim lNewImage
    lOldImage = PImage
    lNewImage = CloneBppGraphics(lOldImage, pForceAlpha, pForceNoAlpha, pForce32bppARGB)
    If lNewImage <> lOldImage Then
        If lOldImage = gBitmap Then GdipDeleteGraphics gGraphics: gGraphics = 0
        ImportExifDataFromBitmap lOldImage, lNewImage
        If lOldImage = gBitmap Then
            PrivBitmap(gUseGdiImg) = lNewImage
            lNewImage = gBitmap
        End If
        GdipDisposeImage lOldImage
        PImage = lNewImage
    End If
End Sub

'---------------------------------------------------------------------------------------
' Conversion de l'image en au même format qu'une autre
'---------------------------------------------------------------------------------------
' pImage   : Identifiant gdiplus de l'image
'---------------------------------------------------------------------------------------
Public Function ConvertForImg(pImg As Variant) As Boolean
Dim lgraphics, lNewGraphics
Dim lNewBitmap, lBitmap
Dim lImg As clGdiplus
Set lImg = ImgFromVariant(pImg)
lBitmap = lImg.SaveBitmap
If PrivGdipGetImageGraphicsContext(lBitmap, lgraphics) = 0 Then
    If GdipCreateBitmapFromGraphics(ImageWidth, ImageHeight, lgraphics, lNewBitmap) = 0 Then
        GdipDeleteGraphics lgraphics
        If GdipGetImageGraphicsContext(lNewBitmap, lNewGraphics) = 0 Then
            If GdipDrawImageRectI(lNewGraphics, gBitmap, 0, 0, ImageWidth, ImageHeight) = 0 Then
            'If GdipDrawImageI(lNewGraphics, gBitmap, 0, 0) = 0 Then
                PrivBitmap(gUseGdiImg) = lNewBitmap
                ConvertForImg = True
            End If
            GdipDeleteGraphics lNewGraphics
        End If
    Else
        GdipDeleteGraphics lgraphics
    End If
End If
End Function
Private Function PrivGdipGetImageGraphicsContext(pBitmap, pGraphics) As Long
Dim lPxfmt As Long
GdipGetImagePixelFormat pBitmap, lPxfmt
Select Case lPxfmt
    Case PixelFormat24bppRGB, PixelFormat32bppRGB, PixelFormat32bppARGB, PixelFormat32bppPARGB, _
            PixelFormat48bppRGB, PixelFormat64bppARGB, PixelFormat64bppPARGB
        PrivGdipGetImageGraphicsContext = GdipGetImageGraphicsContext(pBitmap, pGraphics)
    Case Else
        pGraphics = 0
        PrivGdipGetImageGraphicsContext = 1
End Select
End Function
'---------------------------------------------------------------------------------------
' Chargement d'un fichier a partir de son contenu
' le fichier est integralement charge en memoire et n'est pas bloque par l'application
'---------------------------------------------------------------------------------------
Private Function PrivGetClonedBitmapFromFile(pFile As String)
Dim lFile As Integer
Dim lArray() As Byte
On Error GoTo gestion_erreurs
lFile = FreeFile
' Test ouverture Input pour vérifier existence fichier
' Sinon For Binary crée un fichier vide
Open pFile For Input Access Read As #lFile
Close #lFile
Open pFile For Binary As #lFile
ReDim lArray(0 To LOF(lFile) - 1)
Get lFile, , lArray
PrivGetClonedBitmapFromFile = GetBitmapFromArray(lArray)
Erase lArray
Close lFile
Exit Function
gestion_erreurs:
    Close lFile
    PrivGetClonedBitmapFromFile = 0
End Function
'---------------------------------------------------------------------------------------
' Vide l'image (ne tient pas compte des transformations World*)
'---------------------------------------------------------------------------------------
Public Function Clear(Optional ByVal pColor As Long, Optional ByVal pAlpha As Integer = 255) As Boolean
Dim lHdc
Dim lNewBrush, lOldBrush
Dim lRc As Rect
If gDrawClipRegion <> PtrNull Then
    GdipResetClip gGraphics
    SelectClipRgn gDib.hDC, ByVal PtrNull
End If
If pAlpha = 255 And gDib.hDC <> 0 Then
    lHdc = gDib.hDC
    lNewBrush = CreateSolidBrush(GetColor(pColor))
    lOldBrush = SelectObject(lHdc, lNewBrush)
    lRc.Left = 0
    lRc.Top = 0
    lRc.Right = ImageWidth
    lRc.Bottom = ImageHeight
    Clear = (FillRect(lHdc, lRc, lNewBrush) <> 0)
    SelectObject lHdc, lOldBrush
    DeleteObject lNewBrush
End If
If Not Clear Then
    Clear = (GdipGraphicsClear(GetGraphics, GdipColor(pAlpha, GetColor(pColor))) = 0)
End If
If gDrawClipRegion <> PtrNull Then
    GdipSetClipRegion gGraphics, gDrawClipRegion, 0 ' CombineModeReplace
    PrivSetClipRegionGdi
End If
End Function

'---------------------------------------------------------------------------------------
' Ouverture d'un fichier
'---------------------------------------------------------------------------------------
' pFile   : Fichier
' pThumbNail : Mettre a True pour ouvrir la miniature Exif
' pIcon : Mettre a True pour ouvrir l'icone associee au fichier
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function LoadFile(pFile As String, Optional pThumbNail As Boolean = False, Optional pIcon As EGdipIconType = GdipNoIcon) As Boolean
#Else
Public Function LoadFile(pFile As String, Optional pThumbNail As Boolean = False, Optional pIcon As Long = 0) As Boolean
#End If
Dim lBitmap
    On Error GoTo gestion_erreurs
    ' Retour de la fonction
    LoadFile = True
    ' Creation d'un Bitmap Gdi+ a partir du bitmap de l'image
    If Not pThumbNail And Not pIcon <> 0 Then
        If Not LoadAndClone Then
            LoadFile = (GdipLoadImageFromFile(StrPtr(pFile), lBitmap) = 0)
        Else
            lBitmap = PrivGetClonedBitmapFromFile(pFile)
            LoadFile = (lBitmap <> 0)
        End If
        PrivBitmap(gUseGdiImg) = lBitmap
    Else
        If pThumbNail Then
            PrivBitmap(gUseGdiImg) = GetThumbNail(pFile)
        ElseIf pIcon <> 0 Then
            PrivBitmap(gUseGdiImg) = GetIconBitmap(pFile, pIcon)
        End If
        If gBitmap = 0 Then
            If Not LoadAndClone Then
                LoadFile = (GdipLoadImageFromFile(StrPtr(pFile), lBitmap) = 0)
            Else
                lBitmap = PrivGetClonedBitmapFromFile(pFile)
                LoadFile = (lBitmap <> 0)
            End If
            If ConvertMetaFile Then ConvertBitmapMetafile lBitmap
            PrivBitmap(gUseGdiImg) = lBitmap
        Else
            LoadFile = True
        End If
    End If
gestion_erreurs:
    If Err.Number <> 0 Then LoadFile = False
End Function

'---------------------------------------------------------------------------------------
' Creation d'une image vide
'---------------------------------------------------------------------------------------
' pWidth, pHeight : Taille de l'image
' pResolution     : Resolution de l'image (96 par defaut)
'---------------------------------------------------------------------------------------
Public Function CreateBitmap(ByVal pWidth As Long, ByVal pHeight As Long, Optional pResolution As Single = 96) As Boolean
Dim lBitmap
    On Error GoTo Gestion_Erreur
    ' Creation d'un Bitmap Gdi+ 32 bits
    CreateBitmap = (GdipCreateBitmapFromScan0(pWidth, pHeight, 0, PixelFormat32bppARGB, ByVal PtrNull, lBitmap) = 0)
    If CreateBitmap Then
        GdipBitmapSetResolution lBitmap, pResolution, pResolution
        PrivBitmap(gUseGdiImg) = lBitmap
    End If
    On Error GoTo 0
    Exit Function
Gestion_Erreur:
    ' Si pas d'erreur, renvoit Vrai
    CreateBitmap = False
End Function

'---------------------------------------------------------------------------------------
' Creation d'une image vide de la taille d'un controle ou formulaire
'---------------------------------------------------------------------------------------
' pObject        : Controle ou formulaire
'---------------------------------------------------------------------------------------
Public Function CreateBitmapForControl(pObject As Object, Optional pMakeDefaultImg As Boolean = False) As Boolean
#If UseSingle Then
    Dim lLeft As Single, lTop As Single
    Dim lRight As Single, lBottom As Single
#Else
    Dim lLeft As Long, lTop As Long
    Dim lRight As Long, lBottom As Long
#End If
    Dim lWidth As Long, lHeight As Long
    Dim lBitmap
    On Error GoTo Gestion_Erreur
    If pMakeDefaultImg Then ImgObject = pObject
    Call GetControlPos(pObject, lLeft, lTop, lRight, lBottom)
    lWidth = lRight - lLeft + 1
    lHeight = lBottom - lTop + 1
    ' Creation d'un Bitmap Gdi+
    Dim lHwnd, lHdc
    #If AppName = "A" Then
    lHwnd = GetClientHandle(pObject)
    #Else
    If TypeOf pObject Is MSForms.UserForm Then
    lHwnd = GetUserFormHandle(pObject, True)
    Else
    lHwnd = GetUserFormHandle(pObject.Parent, True)
    End If
    #End If
    lHdc = GetDC(lHwnd)
    lBitmap = PrivCreateBitmapForDC(lHdc, lWidth, lHeight)
    ReleaseDC lHwnd, lHdc
    If lBitmap <> 0 Then
        CreateBitmapForControl = True
        PrivBitmap(gUseGdiWindow) = lBitmap
    End If
    Exit Function
Gestion_Erreur:
    ' Si pas d'erreur, renvoit Vrai
    CreateBitmapForControl = False
End Function

Public Function CreateBitmapForImg(pImg As Variant, ByVal pWidth As Long, ByVal pHeight As Long) As Boolean
Dim lBitmap, lgraphics
Dim lImg As clGdiplus
    Set lImg = ImgFromVariant(pImg)
    If PrivGdipGetImageGraphicsContext(lImg.SaveBitmap, lgraphics) = 0 Then
        GdipCreateBitmapFromGraphics pWidth, pHeight, lgraphics, lBitmap
        GdipDeleteGraphics lgraphics
        CreateBitmapForImg = (lBitmap <> 0)
    End If
    If CreateBitmapForImg Then PrivBitmap(gUseGdiImg) = lBitmap
End Function

Public Function CreateBitmapForDC(pDC, Optional ByVal pWidth As Long, Optional ByVal pHeight As Long) As Boolean
Dim lBitmap
    lBitmap = PrivCreateBitmapForDC(pDC, pWidth, pHeight)
    CreateBitmapForDC = (lBitmap <> 0)
    If CreateBitmapForDC Then
        PrivBitmap(gUseGdiWindow) = lBitmap
    End If
End Function
Public Function CreateBitmapForHwnd(pHwnd, Optional ByVal pWidth As Long, Optional ByVal pHeight As Long) As Boolean
Dim lBitmap, lHdc
    lHdc = GetDC(pHwnd)
    lBitmap = PrivCreateBitmapForDC(lHdc, pWidth, pHeight)
    ReleaseDC pHwnd, lHdc
    CreateBitmapForHwnd = (lBitmap <> 0)
    If CreateBitmapForHwnd Then
        PrivBitmap(gUseGdiWindow) = lBitmap
    End If
End Function

Private Function PrivCreateBitmapForDC(pDC, pWidth As Long, pHeight As Long)
Dim lgraphics, lBitmap
Dim lHwnd, lrect As Rect
If pWidth = 0 Or pHeight = 0 Then
    lHwnd = WindowFromDC(pDC)
    GetWindowRect lHwnd, lrect
    If pWidth = 0 Then pWidth = lrect.Right - lrect.Left
    If pHeight = 0 Then pHeight = lrect.Bottom - lrect.Top
End If
GdipCreateFromHDC pDC, lgraphics
GdipCreateBitmapFromGraphics pWidth, pHeight, lgraphics, lBitmap
GdipDeleteGraphics lgraphics
PrivCreateBitmapForDC = lBitmap
End Function

'---------------------------------------------------------------------------------------
' Fermeture du fichier / de l'image
'---------------------------------------------------------------------------------------
Public Sub CloseImage()
PrivBitmap(False) = PtrNull
End Sub

'---------------------------------------------------------------------------------------
' Sauvegarde l'image dans un fichier avec transformation (sans perte Jpeg)
'---------------------------------------------------------------------------------------
' pFile   : Fichier
' pTransformation  : Transformation a appliquer
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function SaveTransformFile(pFile As String, pTransformation As EGdipJpegTransform) As Boolean
#Else
Public Function SaveTransformFile(pFile As String, pTransformation As Long) As Boolean
#End If
Dim lEncoder As GUID
Dim lParams As EncoderParameters
Dim lEncoderStr As String
Dim lTransformation As Long
On Error GoTo gestion_erreurs
' Format de l'encodeur
lEncoderStr = gJpegEncoderStr
' Recherche de l'encodeur Jpeg
CLSIDFromString StrPtr(lEncoderStr), lEncoder
lParams.count = 1
lTransformation = pTransformation
With lParams.Parameter(0)
    CLSIDFromString StrPtr(gEncoderTransformation), .GUID
    .NumberOfValues = 1
    .Type = 4    ' Type Long
    .value = VarPtr(lTransformation)
End With
' Sauvegarde l'image
SaveTransformFile = (GdipSaveImageToFile(gBitmap, StrPtr(pFile), lEncoder, lParams) = 0)
Exit Function
gestion_erreurs:
SaveTransformFile = False
End Function

'---------------------------------------------------------------------------------------
' Sauvegarde l'image dans un fichier sans perte Jpeg
'---------------------------------------------------------------------------------------
' pFile   : Fichier
'---------------------------------------------------------------------------------------
' A utiliser si uniquement modification de tag exif par exemple
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function SaveJpegLossLess(pFile As String) As Boolean
#Else
Public Function SaveJpegLossLess(pFile As String) As Boolean
#End If
On Error GoTo gestion_erreurs

Dim lImage
Dim lEncoder As GUID
Dim lParams As EncoderParameters
Dim lEncoderStr As String
Dim lTransformation As Long
Dim lstream As Object
On Error GoTo gestion_erreurs
' Format de l'encodeur
lEncoderStr = gJpegEncoderStr
' Recherche de l'encodeur Jpeg
CLSIDFromString StrPtr(lEncoderStr), lEncoder
lParams.count = 1
lTransformation = 14 ' Rotate 180
With lParams.Parameter(0)
    CLSIDFromString StrPtr(gEncoderTransformation), .GUID
    .NumberOfValues = 1
    .Type = 4    ' Type Long
    .value = VarPtr(lTransformation)
End With
' Sauvegarde l'image
' Cree un flux vide
If CreateStreamOnHGlobal(ByVal 0, True, lstream) = 0 Then
    SaveJpegLossLess = (GdipSaveImageToStream(gBitmap, lstream, lEncoder, lParams) = 0)
    If SaveJpegLossLess Then
        If GdipLoadImageFromStream(lstream, lImage) = 0 Then
            SaveJpegLossLess = (GdipSaveImageToFile(lImage, StrPtr(pFile), lEncoder, lParams) = 0)
            GdipDisposeImage lImage
        End If
    End If
    Set lstream = Nothing
End If
On Error GoTo 0
Exit Function
gestion_erreurs:
If lImage <> 0 Then GdipDisposeImage lImage
SaveJpegLossLess = False
End Function

'---------------------------------------------------------------------------------------
' Sauvegarde l'image dans un fichier
'---------------------------------------------------------------------------------------
' pFile   : Fichier
' pFormat : format du fichier : JPG,GIF,BMP,PNG,TIF,EMF
' pQuality : Qualite Jpeg (0-100)
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function SaveFile(pFile As String, Optional ByVal pFormat As EGdipImageFormatSave = -1, Optional ByVal pQuality As Integer = -1) As Boolean
#Else
Public Function SaveFile(pFile As String, Optional ByVal pFormat As Long = -1, Optional ByVal pQuality As Integer = -1) As Boolean
#End If
    Dim lEncoder As GUID
    Dim lParams As EncoderParameters
    Dim lEncoderStr As String
    Dim lHdc
    Dim lrect As Rect
    Dim lhMeta, lEmfGraphics
    Dim lDescription As String
    Dim lMetaType As Long
    Dim lPosExt As Long
    On Error GoTo gestion_erreurs
    If pFormat = -1 Then
        lPosExt = Len(pFile) - 3
        If Not Mid(pFile, lPosExt, 1) = "." Then
            lPosExt = Len(pFile) - 4
            If Not Mid(pFile, lPosExt, 1) = "." Then
                lPosExt = 0
            End If
        End If
        If lPosExt <> 0 Then
            Select Case LCase(Mid(pFile, lPosExt + 1))
                Case "jpg", "jpeg"
                    pFormat = gdipFormatJPEG
                Case "emf"
                    pFormat = gdipFormatEMF
                Case "png"
                    pFormat = gdipFormatPNG
                Case "gif"
                    pFormat = gdipFormatGIF
                Case "bmp"
                    pFormat = gdipFormatBMP
                Case "tif", "tiff"
                    pFormat = gdipFormatTIFF
            End Select
        End If
    End If
    If pFormat = gdipFormatSaveEMF Then
        lDescription = "clGdiplus"
        lrect.Right = ImageWidth
        lrect.Bottom = ImageHeight
        GdipGetDC GetGraphics, lHdc
        Select Case pQuality
            Case 1: lMetaType = 3
            Case 2: lMetaType = 4
            Case 3: lMetaType = 5
            Case Else: lMetaType = 3
        End Select
        If GdipRecordMetafileFileNameI(StrPtr(pFile), lHdc, lMetaType, lrect, 2, StrPtr(lDescription), lhMeta) = 0 Then
            GdipGetImageGraphicsContext lhMeta, lEmfGraphics
            GdipDrawImageRectI lEmfGraphics, gBitmap, 0, 0, ImageWidth, ImageHeight
            GdipDeleteGraphics lEmfGraphics
            GdipDisposeImage lhMeta
            SaveFile = True
        End If
        GdipReleaseDC GetGraphics, lHdc
    Else
        ' Format de l'encodeur
        Select Case pFormat
        Case gdipFormatSaveJPEG
            lEncoderStr = gJpegEncoderStr
        Case gdipFormatSaveGIF
            lEncoderStr = gGifEncoderStr
        Case gdipFormatSaveBMP
            lEncoderStr = gBmpEncoderStr
        Case gdipFormatSavePNG
            lEncoderStr = gPngEncoderStr
        Case gdipFormatSaveTIFF
            lEncoderStr = gTifEncoderStr
        End Select
        SaveFile = True
        ' Recherche de l'encodeur
        CLSIDFromString StrPtr(lEncoderStr), lEncoder
        ' Parametre de l'encodeur Jpeg
        If pQuality <> -1 And pFormat = gdipFormatSaveJPEG Then
            lParams.count = 1
            With lParams.Parameter(0)
                ' Parametrage de la qualite (0-100)
                CLSIDFromString StrPtr(gQualityParamStr), .GUID
                .NumberOfValues = 1
                .Type = 4    ' Type Long
                .value = VarPtr(pQuality)
            End With
        End If
        ' Sauvegarde l'image
        If lParams.count > 0 Then
            SaveFile = (GdipSaveImageToFile(gBitmap, StrPtr(pFile), lEncoder, lParams) = 0)
        Else
            SaveFile = (GdipSaveImageToFile(gBitmap, StrPtr(pFile), lEncoder, ByVal PtrNull) = 0)
        End If
    End If
gestion_erreurs:
    If Err.Number <> 0 Then SaveFile = False
End Function

'---------------------------------------------------------------------------------------
' Sauvegarde l'image en memoire
'---------------------------------------------------------------------------------------
' pImage    : nom de la sauvegarde
'---------------------------------------------------------------------------------------
Public Function ImageKeep(Optional PImage As String = "Default") As Boolean
    Dim lBitmap
    On Error GoTo gestion_erreurs
    'GdipCloneImage gBitmap, lBitmap
    lBitmap = PrivBitmapClone(gBitmap)
    'Dim lwidth As Single, lheight As Single
    'GdipGetImageDimension gBitmap, lwidth, lheight
    'GdipCloneBitmapArea 0, 0, lwidth, lheight, 0, gBitmap, lBitmap

    ImageKeep = (lBitmap <> 0)
    If ImageKeep Then
        gCollSave.Add lBitmap, PImage
    End If
gestion_erreurs:
    ' Si l'image existait deja on l'ecrase
    If Err.Number = 457 Then
        ImageDelete PImage: Resume
    ElseIf Err.Number <> 0 Then
        ImageKeep = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Retablit l'image de la memoire
'---------------------------------------------------------------------------------------
' pImage    : nom de la sauvegarde a retablir
'---------------------------------------------------------------------------------------
Public Function ImageReset(Optional PImage As String = "Default") As Boolean
    Dim lBitmap, lImage
    On Error GoTo gestion_erreurs
    lImage = gCollSave.item(PImage)
    GdipCloneImage lImage, lBitmap
    ImageReset = (lBitmap <> 0)
    If ImageReset Then
        If gGraphics <> PtrNull Then
            GdipCreateMatrix gMatrixForImageReset
            GdipGetWorldTransform gGraphics, gMatrixForImageReset
        End If
        PrivBitmap(gUseGdi) = lBitmap
        If gGraphics <> PtrNull Then
            GdipDeleteMatrix gMatrixForImageReset
            gMatrixForImageReset = PtrNull
        End If
    End If
    Exit Function
gestion_erreurs:
    ImageReset = False
End Function
'---------------------------------------------------------------------------------------
' Teste si une image existe en memoire
'---------------------------------------------------------------------------------------
' pImage    : nom de la sauvegarde a tester
'---------------------------------------------------------------------------------------
Public Function ImageExists(Optional PImage As String = "Default") As Boolean
    Dim lImage As Long
    On Error GoTo gestion_erreurs
    lImage = gCollSave.item(PImage)
    ImageExists = True
    Exit Function
gestion_erreurs:
    ImageExists = False
End Function
'---------------------------------------------------------------------------------------
' Supprime toutes les images en memoire
'---------------------------------------------------------------------------------------
Public Function ImagesDelete() As Boolean
    Dim lcpt As Long
    On Error GoTo gestion_erreurs
    ' Vide la collection de sauvegarde
    For lcpt = 1 To gCollSave.count
        GdipDisposeImage gCollSave.item(1)
        gCollSave.Remove 1
    Next
gestion_erreurs:
    If Err.Number = 0 Then ImagesDelete = True
End Function

'---------------------------------------------------------------------------------------
' Supprime l'image en memoire
'---------------------------------------------------------------------------------------
' pImage    : nom de la sauvegarde a supprimer
'---------------------------------------------------------------------------------------
Public Function ImageDelete(Optional PImage As String = "Default") As Boolean
    On Error GoTo gestion_erreurs
    GdipDisposeImage gCollSave.item(PImage)
    gCollSave.Remove PImage
gestion_erreurs:
    If Err.Number = 0 Then ImageDelete = True
End Function
'---------------------------------------------------------------------------------------
' Renvoie le bitmap gdiplus d'une image
'---------------------------------------------------------------------------------------
Public Function ImageGetBitmap(Optional PImage As String = "Default", Optional pClone As Boolean = False) As Long
    Dim lImage As Variant
    On Error Resume Next
    If pClone Then
        GdipCloneImage gCollSave.item(PImage), lImage
    Else
        lImage = gCollSave.item(PImage)
    End If
    On Error GoTo 0
    ImageGetBitmap = lImage
End Function

'---------------------------------------------------------------------------------------
' Clone un bitmap
' Si utilisation de gdi32 (DIB), on ne peut pas utiliser GdipCloneImage
'---------------------------------------------------------------------------------------
Private Function PrivBitmapClone(pBitmap)
    Dim lNewBitmap
    Dim lWidth As Single, lHeight As Single
    Dim lgraphics, lNewGraphics
    On Error GoTo gestion_erreurs
    
    If pBitmap <> gBitmap Or Not gUseGdi Then
        If GdipCloneImage(pBitmap, lNewBitmap) = 0 Then
            PrivBitmapClone = lNewBitmap
        End If
    Else
        lgraphics = GetGraphics
        GdipGetImageDimension pBitmap, lWidth, lHeight
        GdipCreateBitmapFromGraphics lWidth, lHeight, lgraphics, lNewBitmap
        GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
        If GdipDrawImageRectI(lNewGraphics, pBitmap, 0, 0, lWidth, lHeight) = 0 Then
            PrivBitmapClone = lNewBitmap
        End If
        GdipDeleteGraphics lNewGraphics
    End If
    Exit Function
gestion_erreurs:
    PrivBitmapClone = PtrNull
End Function

'---------------------------------------------------------------------------------------
' Initialisation de la classe
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()
Dim lGdiPSI As GdiplusStartupInput ' Structure d'initialisation de gdiplus
    ' Charge la librarie gdiplus
    gLib = LoadLibrary(ApplicationPath & "gdiplus.dll")
    If gLib = 0 Then gLib = LoadLibrary("gdiplus.dll")
        If GetProcAddress(gLib, "GdipDrawImageFX") Then gIsV11 = True Else gIsV11 = False
    lGdiPSI.GdiplusVersion = 1
    Call GdiplusStartup(gGdipToken, lGdiPSI)
    ' Initialise la collection d'images
    Set gImgList = New Collection
    ' Initialise la collection de textures
    Set gTextures = New Collection
    ' Initialise la collection de regions
    Set gCollRegions = New Collection
    ' Initialise la collection de curseurs
    Set gCurseurs = New Collection
    ' Initialise la collection de sauvegarde
    Set gCollSave = New Collection
    ' Initialise la collection de formulaires a retablir
    Set gFormRegionHwnd = New Collection
    ' Initialise les collections de matrices
    Set gMatrix = New Collection
    ' Initialise les donnees utilisateurs
    Set gUserData = New Collection
    ' Utilisation de gdi32 si fenêtre
    gUseGdiWindow = True
    ' Valeur de PI
    gPI = 4 * Atn(1)
    ' Conversion des metafile au chargement
    ConvertMetaFile = True
    ' Fond des icones = blanc
    IconBackColor = vbWhite
    ' Par defaut = chargement sans clonage de l'image
    LoadAndClone = False
    ' Limite jointure = 10 par defaut (cf doc gdi+)
    LineJoinMiterLimit = 10
    ' Modes graphics
    gTransparencyMode = GdipTransparencySourceOver
    gPixelOffsetMode = GdipPixelOffsetDefault
    gSmoothingMode = GdipSmoothingDefault
    gAntialisingLevel = GdipAntialiseHighQuality
    gTextRenderingHint = GdipHintAntiAliasGridFit
    gCompositingQuality = GdipCompositingQualityHighSpeed
    ImgCopyRendering = True
    ' Transformation des régions à la création
    RegionWorldTransform = True
    ' pointilles
    ReDim gDashUser(1 To 1)
    gDashUser(1) = 1
    ' barres de defilement
    Set goBars = New Collection
    DrawMatrix = Null
    RepaintMatrix = Null
End Sub

'---------------------------------------------------------------------------------------
' Liberation de la classe
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Dim lImg As clGdiplus
    Dim lImage As Variant
    Dim lregion As Variant
    Dim lCurseur As Variant
    Dim lFormHwnd As Variant
    ' Supprime les objets de la minuterie
    If gId <> 0 Then timeKillEvent gId
    If ghEvent <> 0 Then ResetEvent ghEvent
    If ghEvent <> 0 Then CloseHandle ghEvent
    ' Vide les donnees utilisateurs
    Set gUserData = Nothing
    ' Ferme un eventuel fichier ouvert
    CloseImage
    ' Vide la collection de matrices de transformation
    Set gMatrix = Nothing
    ' Vide la collection de matrices
    Set gMatrix = Nothing
    ' Vide la collection d'images
    For Each lImg In gImgList
        lImg.CloseImage
        Set lImg = Nothing
    Next
    Set gImgList = Nothing
    ' Vide la collection de textures
    For Each lImage In gTextures
        GdipDeleteBrush lImage
    Next
    Set gTextures = Nothing
    ' Vide la collection de sauvegarde
    For Each lImage In gCollSave
        GdipDisposeImage lImage
    Next
    Set gCollSave = Nothing
    ' Vide la collection de curseurs
    For Each lCurseur In gCurseurs
        DeleteObject lCurseur
    Next
    Set gCurseurs = Nothing
    ' Vide la collection de regions
    For Each lregion In gCollRegions
        If lregion(2) Then
            GdipDeletePath lregion(0)
        Else
            GdipDeleteRegion lregion(0)
        End If
    Next
    Set gCollRegions = Nothing
    ' Supprime les polices de caracteres privees
    If gPrivateFontCollection <> 0 Then
        GdipDeletePrivateFontCollection gPrivateFontCollection
    End If
    ' Retabli les formulaires dont l'affichage a ete reduits
    For Each lFormHwnd In gFormRegionHwnd
        Call SetWindowRgn(lFormHwnd, 0, True)
    Next
    Set gFormRegionHwnd = Nothing
    ' Libere le controle de reference
    Set gCtrlRef = Nothing
    ' Supprime la region pour Clip
    If gClipRegion <> 0 Then GdipDeleteRegion gClipRegion
    If gDrawClipRegion <> 0 Then GdipDeleteRegion gDrawClipRegion
    ' Ferme Gdi+
    If gGdipToken <> 0 Then GdiplusShutdown gGdipToken: gGdipToken = 0
End Sub

'---------------------------------------------------------------------------------------
' Chemin de l'application
'---------------------------------------------------------------------------------------
Public Function ApplicationPath() As String
    #If AppName = "A" Then
        #If VBA6 Then
            ApplicationPath = CurrentProject.Path & "\"
        #Else
            ApplicationPath = Left(CurrentDb.name, Len(CurrentDb.name) - Len(Dir(CurrentDb.name)))
        #End If
    #ElseIf AppName = "E" Then
        ApplicationPath = ThisWorkbook.Path & "\"
    #ElseIf AppName = "W" Then
        ApplicationPath = ThisDocument.Path & "\"
    #ElseIf AppName = "P" Then
        ApplicationPath = ThisPresentation.Path & "\"
    #End If
End Function

'---------------------------------------------------------------------------------------
' Injecte l'image dans un controle ou formulaire
' Repaint temporairement l'image pour reduire les effets de scintillement
'---------------------------------------------------------------------------------------
' pControl          : Controle dans lequel est injecte l'image
' pResizeBefore     : Si Vrai, redimensionne l'image avant
' pResizeAntialise  : Anti-aliasing lors du redimensionnement
' pUseEMF           : Pour Access : Utilisation du type d'image EMF
'         (a tester si le redimensionnement en BMP n'est pas "propre")
'---------------------------------------------------------------------------------------

#If AppName = "A" Then
Public Function RepaintNoFormRepaint(Optional pObject As Object, Optional pResizeBefore As Boolean, Optional pResizeAntialiase As Boolean, Optional pUseEMF As Boolean)
#Else
Public Function RepaintNoFormRepaint(Optional pObject As Object, Optional pResizeBefore As Boolean, Optional pResizeAntialiase As Boolean, Optional pBackgroundColor As Long = -1)
#End If
    ' Injecte l'image dans un controle Image
    #If AppName = "A" Then
    Repaint pObject, pResizeBefore, pResizeAntialiase, pUseEMF
    #Else
    Repaint pObject, pResizeBefore, pResizeAntialiase, pBackgroundColor
    #End If
    ' Dessine l'image sur le formulaire, en annulant le dessin par windows
    gCancelPaint = True
    RepaintFast pObject, , , pResizeAntialiase
    gCancelPaint = False
End Function

'---------------------------------------------------------------------------------------
' Injecte l'image dans un controle Image
'---------------------------------------------------------------------------------------
' pControl          : Controle dans lequel est injecte l'image
' pResizeBefore     : Si Vrai, redimensionne l'image avant
' pResizeAntialise  : Anti-aliasing lors du redimensionnement
' pUseEMF           : Pour Access : Utilisation du type d'image EMF
'         (a tester si le redimensionnement en BMP n'est pas "propre")
'---------------------------------------------------------------------------------------
#If AppName = "A" Then
Public Function Repaint(Optional pObject As Object, Optional pResizeBefore As Boolean, Optional pResizeAntialiase As Boolean, Optional pUseEMF As Boolean)
    If pObject Is Nothing Then Set pObject = oObject
    pObject.PictureData = PrivGdiPlusToPictureData(IIf(pResizeBefore, pObject, Nothing), pResizeAntialiase, pUseEMF)
    If RepaintMatrixUpdate Then PrivRepaintMatrixUpdate pObject
End Function
#Else
Public Function Repaint(Optional pObject As Object, Optional pResizeBefore As Boolean, Optional pResizeAntialiase As Boolean, Optional pBackgroundColor As Long = -1)
    If pObject Is Nothing Then Set pObject = oObject
    Set pObject.Picture = PrivGdiPlusToIPictureDisp(IIf(pResizeBefore, pObject, Nothing), pResizeAntialiase, , pBackgroundColor)
    If RepaintMatrixUpdate Then PrivRepaintMatrixUpdate pObject
End Function
#End If

Private Sub PrivRepaintMatrixUpdate(pObject As Object)
    RepaintMatrix = MatrixCtrlToImg(pObject)
End Sub

'---------------------------------------------------------------------------------------
' Transfere les donnees du Bitmap dans un tableau de type PictureData EMF
'---------------------------------------------------------------------------------------
' pBitmap        : Objet Bitmap Gdiplus
' On peut avoir besoin d'un objet EMF car Access gere tres mal les redimensionnements des Bitmaps
'    ou pour de la transparence
'---------------------------------------------------------------------------------------
Private Function BitmaptoPictureDataEMF(pBitmap) As Variant
    Dim lPicData() As Byte
    Dim rc As Rect
    Dim lhMeta, lngret As Long
    Dim lhEMF
    Dim lEmfGraphics
    Dim lgraphics
    Dim lHdc
    Dim lWidth As Long, lHeight As Long
    Dim lBitmap
    Dim lDescription As String
    Dim pt As POINTAPI
    Dim lHdcref
    On Error GoTo gestion_erreurs:
    lBitmap = pBitmap
    ' Taille EMF
    Call GdipGetImageWidth(lBitmap, lWidth)
    Call GdipGetImageHeight(lBitmap, lHeight)
    ' Recupere la taille en donnees de type OLE_Himetric pour la creation de l'EMF
    lHdcref = GetDC(0)   ' Device contexte temporaire
    lngret = SetMapMode(lHdcref, MM_HIMETRIC)
    ' Rectangle pour creation de l'EMF
    pt.X = (lWidth - 0) * (HIMETRIC_INCH / GetDeviceCaps(lHdcref, LOGPIXELSX))
    pt.Y = (lHeight - 0) * (HIMETRIC_INCH / GetDeviceCaps(lHdcref, LOGPIXELSY))
    ' Conversion de la taille en pixels
    LPtoDP lHdcref, pt, 1
    pt.Y = Abs(pt.Y)
    rc.Right = pt.X
    rc.Bottom = pt.Y
    ' Mode d'origine du dc de reference
    SetMapMode lHdcref, lngret
    ReleaseDC 0, lHdcref
    ' Device context de reference
    PrivGdipGetImageGraphicsContext lBitmap, lgraphics
    ' Si < 24 bits => conversion pour avoir un objet graphics
    If lgraphics = 0 Then
        lBitmap = CloneBppGraphics(pBitmap)
        GdipGetImageGraphicsContext lBitmap, lgraphics
    End If
    GdipGetDC lgraphics, lHdc
    ' Creation EMF
    lDescription = "clGdiplus"
    Call GdipRecordMetafileI(lHdc, 3, rc, 2, StrPtr(lDescription), lhMeta)
    ' Libere objets
    GdipReleaseDC lgraphics, lHdc
    GdipDeleteGraphics lgraphics
    ' Dessine l'image sur l'EMF
    GdipGetImageGraphicsContext lhMeta, lEmfGraphics
    GdipSetInterpolationMode lEmfGraphics, GdipAntialiseNearestNeighbor
    GdipDrawImageRectI lEmfGraphics, lBitmap, 0, 0, rc.Right, rc.Bottom
    GdipDeleteGraphics lEmfGraphics
    ' Conversion en EMF gdi
    GdipGetHemfFromMetafile lhMeta, lhEMF
    ' Supprime l'EMF gdiplus
    GdipDisposeImage lhMeta
    ' Recupere la taille des donnees Meta
    lngret = GetEnhMetaFileBits(lhEMF, 0, ByVal PtrNull)
    ' Redimensionne le tableau de donnees
    ReDim lPicData(0 To (lngret - 1) + 8) As Byte
    ' Recupere les donnees dans le tableau
    lngret = GetEnhMetaFileBits(lhEMF, lngret, lPicData(8))
    ' Supprime le MetaFile
    lngret = DeleteEnhMetaFile(lhEMF)
    ' Type de l'image dans le tableau de donnees
    lPicData(0) = CF_ENHMETAFILE
    ' Renvoie le resultat
    BitmaptoPictureDataEMF = lPicData
    ' Supprime le bitmap temporaire si conversion
    If lBitmap <> pBitmap Then GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    BitmaptoPictureDataEMF = Null  ' Renvoie Null si erreur
End Function
    '---------------------------------------------------------------------------------------
    ' Transfere les donnees du Bitmap dans un tableau de type PictureData Bitmap
    '---------------------------------------------------------------------------------------
    ' pBitmap        : Objet Bitmap Gdiplus
    '---------------------------------------------------------------------------------------
Private Function BitmaptoPictureDataBMP(pBitmap) As Variant
    Dim lPicData() As Byte
    Dim lWidth As Single, lHeight As Single
    Dim lBI As BitmapInfoHeader
    Dim bmpData As BitmapData
    Dim lrect As Rect
    On Error GoTo gestion_erreurs:
    ' Lecture des dimensions de l'image
    GdipGetImageDimension pBitmap, lWidth, lHeight
    ' Rectangle pour GdipBitmapLockBits
    lrect.Right = lWidth
    lrect.Bottom = lHeight
    ' Lecture des donnees du bitmap
    If GdipBitmapLockBits(pBitmap, lrect, &H1, PixelFormat32bppARGB, bmpData) = 0 Then
        ' On rempli l'en-tete
        #If Win64 Then
        lBI.biSize = LenB(lBI)
        #Else
        lBI.biSize = Len(lBI)
        #End If
        ' Taille de l'image = taille de l'objet
        lBI.biWidth = bmpData.Width
        lBI.biHeight = bmpData.Height
        ' Compression RGB (pas de compression en fait)
        lBI.biCompression = &H0
        ' 32bits
        lBI.biBitCount = 32
        ' Nombre de plans : toujours = 1 (d'apres MSDN)
        lBI.biPlanes = 1
        ' Calcul de la taille de l'image
        lBI.biSizeImage = Abs(bmpData.stride) * bmpData.Height
        #If Win64 Then
        ' Redimensionne le tableau pour contenir l'en-tete + l'image
        ReDim lPicData(0 To Abs(bmpData.stride) * bmpData.Height + LenB(lBI) - 1)
        ' Copie des donnees de l'image
        RtlMoveMemory lPicData(LenB(lBI)), ByVal bmpData.scan0, Abs(bmpData.stride) * bmpData.Height
        #Else
        ReDim lPicData(0 To Abs(bmpData.stride) * bmpData.Height + Len(lBI) - 1)
        RtlMoveMemory lPicData(Len(lBI)), ByVal bmpData.scan0, Abs(bmpData.stride) * bmpData.Height
        #End If
        ' Libere le bitmap
        Call GdipBitmapUnlockBits(pBitmap, bmpData)
    End If
    ' Copie des donnees en-tete
    #If Win64 Then
    RtlMoveMemory lPicData(0), lBI, LenB(lBI)
    #Else
    RtlMoveMemory lPicData(0), lBI, Len(lBI)
    #End If
    ' Renvoie le resultat
    BitmaptoPictureDataBMP = lPicData
    Exit Function
gestion_erreurs:
    ' Libere le bitmap
    Call GdipBitmapUnlockBits(pBitmap, bmpData)
    ' Renvoie Null si erreur
    BitmaptoPictureDataBMP = Null
End Function
'---------------------------------------------------------------------------------------
' Recupere un tableau PictureData pour controle ou formulaire Access
'---------------------------------------------------------------------------------------
' pResizeCtrl       : Control d'apres lequel l'image est redimensionnee avant d'etre affichee
' pResizeAntialise  : Anti-aliasing lors du redimensionnement
' pUseEMF           : Utilisation du type d'image EMF (a tester si le redimensionnement en BMP n'est pas "propre")
'---------------------------------------------------------------------------------------
Private Function PrivGdiPlusToPictureData(Optional pResizeCtrl As Object, Optional pResizeAntialise As Boolean = False, Optional pUseEMF As Boolean) As Variant
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lWidthResize As Long, lHeigthResize As Long
    Dim lNewBitmap
    Dim lNewGraphics
    Dim lCtrlWidth As Long, lCtrlHeight As Long
    Dim lFDTGUID As GUID
    Const lFrameDimensionTime As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
    On Error GoTo gestion_erreurs

    If Not IsMissing(pResizeCtrl) And Not pResizeCtrl Is Nothing Then
        ' Redimensionne a la taille du controle
        GdipGetImageDimension gBitmap, lWidth, lHeight
        If lWidth = 0 Then Exit Function
        lCtrlWidth = PointsToPixelsX(pResizeCtrl.Width)
        lCtrlHeight = PointsToPixelsY(pResizeCtrl.Height)
        If ((lWidth - lCtrlWidth)) / lWidth * lHeight < (lHeight - lCtrlHeight) Then
            lWidthResize = lCtrlWidth - (lCtrlWidth - lWidth * (lCtrlHeight) / lHeight)
            lHeigthResize = lWidthResize / lWidth * lHeight
        Else
            lHeigthResize = lCtrlHeight - (lCtrlHeight - lHeight * (lCtrlWidth) / lWidth)
            lWidthResize = lHeigthResize / lHeight * lWidth
        End If
        GdipCreateBitmapFromScan0 lWidthResize, lHeigthResize, lWidthResize * 4, PixelFormat32bppARGB, ByVal PtrNull, lNewBitmap
        GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
        If pResizeAntialise Then GdipSetInterpolationMode lNewGraphics, gAntialisingLevel
        GdipDrawImageRectI lNewGraphics, gBitmap, 0, 0, lWidthResize, lHeigthResize
        GdipDeleteGraphics lNewGraphics
        ' Retour de la fonction = tableau de byte affectable a une propriete PictureData
        If pUseEMF Then
            PrivGdiPlusToPictureData = BitmaptoPictureDataEMF(lNewBitmap)
        Else
            ' Reselectionne l'image gif si necessaire
            If gActiveFrame <> 0 Then
                Call CLSIDFromString(StrPtr(lFrameDimensionTime), lFDTGUID)
                GdipImageSelectActiveFrame lNewBitmap, lFDTGUID, gActiveFrame
            End If
            Call GdipImageRotateFlip(lNewBitmap, 6)    ' Retourne l'image car sinon l'image est la tete en bas
            PrivGdiPlusToPictureData = BitmaptoPictureDataBMP(lNewBitmap)
            GdipDisposeImage lNewBitmap
        End If
    Else
        ' Retour de la fonction = tableau de byte affectable a une propriete PictureData
        If pUseEMF Then
            PrivGdiPlusToPictureData = BitmaptoPictureDataEMF(gBitmap)
        Else
            lNewBitmap = PrivBitmapClone(gBitmap)
            'GdipCloneImage gBitmap, lNewBitmap
            ' Reselectionne l'image gif si necessaire
            If gActiveFrame <> 0 Then
                Call CLSIDFromString(StrPtr(lFrameDimensionTime), lFDTGUID)
                GdipImageSelectActiveFrame lNewBitmap, lFDTGUID, gActiveFrame
            End If
            Call GdipImageRotateFlip(lNewBitmap, 6)    ' Retourne l'image car sinon l'image est la tete en bas
            PrivGdiPlusToPictureData = BitmaptoPictureDataBMP(lNewBitmap)
            GdipDisposeImage lNewBitmap
        End If
    End If

    On Error GoTo 0
    Exit Function
gestion_erreurs:
    If lNewGraphics <> 0 Then GdipDeleteGraphics lNewGraphics
    If lNewBitmap <> 0 Then GdipDisposeImage lNewBitmap
    PrivGdiPlusToPictureData = Null
End Function

'---------------------------------------------------------------------------------------
' Transfere les donnees du Bitmap dans un objet IPictureDisp
'---------------------------------------------------------------------------------------
' pResizeCtrl       : Control d'apres lequel l'image est redimensionnee avant d'etre affichee
' pResizeAntialise  : Anti-aliasing lors du redimensionnement
' pBitmap           : Objet Bitmap
'---------------------------------------------------------------------------------------
#If VBA7 Then
Private Function PrivGdiPlusToIPictureDisp(Optional pResizeCtrl As Object, Optional pResizeAntialise As Boolean = False, Optional ByVal pBitmap As LongPtr, Optional pBackgroundColor As Long = -1) As Object
#Else
Private Function PrivGdiPlusToIPictureDisp(Optional pResizeCtrl As Object, Optional pResizeAntialise As Boolean = False, Optional ByVal pBitmap As Long, Optional pBackgroundColor As Long = -1) As Object
#End If
    Dim PicGUID As GUID
    Dim pic As PicBmp
    With PicGUID
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    Dim lhBitmap
    Dim lBitmap
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lWidthResize As Long, lHeigthResize As Long
    Dim lNewBitmap
    Dim lNewGraphics
    On Error GoTo gestion_erreurs
    If pBitmap = 0 Then
        lBitmap = gBitmap
    Else
        lBitmap = pBitmap
    End If
    ' Redimensionne a la taille du controle
    If Not IsMissing(pResizeCtrl) And Not pResizeCtrl Is Nothing Then
        GdipGetImageDimension lBitmap, lWidth, lHeight
        lWidthResize = PointsToPixelsX(pResizeCtrl.Width)
        lHeigthResize = lWidthResize / lWidth * lHeight
        GdipCreateBitmapFromScan0 lWidthResize, lHeigthResize, lWidthResize * 4, PixelFormat32bppARGB, ByVal PtrNull, lNewBitmap
        GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
        If pResizeAntialise Then GdipSetInterpolationMode lNewGraphics, gAntialisingLevel
        GdipDrawImageRectI lNewGraphics, lBitmap, 0, 0, lWidthResize, lHeigthResize
        GdipDeleteGraphics lNewGraphics
        Call GdipCreateHBITMAPFromBitmap(lNewBitmap, lhBitmap, IIf(pBackgroundColor = -1, 0, GdipColor(255, pBackgroundColor)))
        GdipDisposeImage lNewBitmap
    Else
        Call GdipCreateHBITMAPFromBitmap(lBitmap, lhBitmap, IIf(pBackgroundColor = -1, 0, GdipColor(255, pBackgroundColor)))
    End If
    With pic
        #If Win64 Then
        .size = LenB(pic)
        #Else
        .size = Len(pic)
        #End If
        .tType = 1
        .hBmp = lhBitmap
        .hpal = PtrNull
    End With
    Call OleCreatePictureIndirect(pic, PicGUID, 1, PrivGdiPlusToIPictureDisp)
gestion_erreurs:
    If Err.Number <> 0 Then Set PrivGdiPlusToIPictureDisp = Nothing   ' Renvoie Nothing si erreur
End Function

'---------------------------------------------------------------------------------------
' Recherche de l'icone associee a un fichier
'---------------------------------------------------------------------------------------
' pFile     : Chemin du fichier
'---------------------------------------------------------------------------------------
' Renvoi un bitmap gdiplus
'---------------------------------------------------------------------------------------
#If VBA6 Then
Private Function GetIconBitmap(pPath As String, Optional ByVal pIconType As EGdipIconType = &H0)
#Else
Private Function GetIconBitmap(pPath As String, Optional ByVal pIconType As Long = &H0)
#End If
    Dim lLoadhBmp As bitmap
    Dim lloaddc
    Dim lThumbBmp
    Dim lpicinfo As ICONINFO
    Dim lIcon
    Dim lgraphics
    Dim sfi As SHFILEINFO
    Dim lBrush
    Dim lExtensionOnly As Boolean
    On Error GoTo gestion_erreurs
    ' Pour compatibilite anciennes versions
    If pIconType = -1 Then pIconType = GdipLargeIcon
    If pIconType = GdipSmallIcon Then
        pIconType = &H1
    Else
        pIconType = &H0
    End If
    ' Si uniquement extension
    If Len(pPath) <= 6 And Left(pPath, 1) = "." Then lExtensionOnly = True
    ' Recupere l'icone
    #If Win64 Then
    Call SHGetFileInfo(ByVal pPath, IIf(lExtensionOnly, &H80, 0), sfi, LenB(sfi), IIf(lExtensionOnly, &H10, 0) Or SHGFI_ICON Or pIconType)
    #Else
    Call SHGetFileInfo(ByVal pPath, IIf(lExtensionOnly, &H80, 0), sfi, Len(sfi), IIf(lExtensionOnly, &H10, 0) Or SHGFI_ICON Or pIconType)
    #End If
    lIcon = sfi.hIcon
    ' Informations de l'icone
    GetIconInfo lIcon, lpicinfo
    ' Lecture des infos du bitmap de l'icone
    #If Win64 Then
    Call GetObjectBmp(lpicinfo.hbmColor, LenB(lLoadhBmp), lLoadhBmp)
    #Else
    Call GetObjectBmp(lpicinfo.hbmColor, Len(lLoadhBmp), lLoadhBmp)
    #End If
    ' Creation du bitmap pour recevoir l'image
    GdipCreateBitmapFromScan0 lLoadhBmp.bmWidth, lLoadhBmp.bmHeight, 4 * lLoadhBmp.bmWidth, PixelFormat32bppARGB, ByVal PtrNull, lThumbBmp
    ' Recherche le contexte graphique et le device context associe
    GdipGetImageGraphicsContext lThumbBmp, lgraphics
    GdipGetDC lgraphics, lloaddc
    ' Brosse pour fond
    lBrush = CreateSolidBrush(IconBackColor)
    ' Dessine l'icone sur le bitmap
    DrawIconEx lloaddc, 0, 0, lIcon, lLoadhBmp.bmWidth, lLoadhBmp.bmHeight, 0, ByVal lBrush, &H3
    DeleteObject lBrush
    GdipReleaseDC lgraphics, lloaddc
    GdipDeleteGraphics lgraphics
    GetIconBitmap = lThumbBmp
    DeleteObject lpicinfo.hbmColor
    DeleteObject lpicinfo.hbmMask
gestion_erreurs:
    If Err.Number <> 0 Then GetIconBitmap = 0
    GdipReleaseDC lgraphics, lloaddc
    GdipDeleteGraphics lgraphics
    DeleteObject lBrush
    DestroyIcon lIcon
    DeleteObject lpicinfo.hbmColor
    DeleteObject lpicinfo.hbmMask
End Function

'---------------------------------------------------------------------------------------
' Recupere la couleur systeme si necessaire
'---------------------------------------------------------------------------------------
' pColor     : Numero de la couleur
' Les couleurs definies par defaut sont parfois des couleurs systemes
' Par exemple la couleur le fond d'un formulaire par defaut est -2147483633
' il faut retrouver la couleur definie dans le systeme pour cette valeur
' Renvoie la couleur dans un Long
'---------------------------------------------------------------------------------------
Private Function GetColor(ByVal pColor As Long) As Long
    If pColor < 0 Then
        Call OleTranslateColor(pColor, 0, pColor)
    End If
    GetColor = pColor
End Function

'---------------------------------------------------------------------------------------
' Conversion code couleur Long vers RGB
'---------------------------------------------------------------------------------------
' pLong     : Numero de la couleur
' pRed      : Composante Rouge
' pGreen    : Composante Verte
' pBlue     : Composante Bleue
'---------------------------------------------------------------------------------------
Public Function LongToRGB(ByVal pLong As Long, pRed As Long, pGreen As Long, pBlue As Long) As Boolean
    Const lLong65536 As Long = 65536
    Const lLong256 As Long = 256
    ' On recupere la couleur systeme si necessaire
    pLong = GetColor(pLong)
    On Error GoTo gestion_erreurs:
    pBlue = Int(pLong / lLong65536)
    pGreen = Int((pLong - (lLong65536 * pBlue)) / lLong256)
    pRed = pLong - ((pBlue * lLong65536) + (pGreen * lLong256))
gestion_erreurs:
    If Err.Number = 0 Then LongToRGB = True    ' Renvoie Vrai si pas d'erreur
End Function

'---------------------------------------------------------------------------------------
' Conversion taille police en pixel
'---------------------------------------------------------------------------------------
' pSize     : Taille de police
' pObject  : Controle de reference
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function FontSizeToPixel(ByVal pSize As Single, Optional pObject As Object) As Single
#Else
Public Function FontSizeToPixel(ByVal pSize As Long, Optional pObject As Object) As Long
#End If
    On Error GoTo gestion_erreurs
    If pObject Is Nothing Then
        FontSizeToPixel = PointsToPixelsY(cPoints * pSize)
    Else
        FontSizeToPixel = CtrlToImgY(cPoints * pSize, pObject, False)
    End If
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    FontSizeToPixel = 0
End Function

'---------------------------------------------------------------------------------------
' Conversion pixel en taille police
'---------------------------------------------------------------------------------------
' pPixel     : Taille en pixel
' pObject  : Controle de reference
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function PixelToFontSize(ByVal pSize As Single, Optional pObject As Object) As Long
#Else
Public Function PixelToFontSize(ByVal pSize As Long, Optional pObject As Object) As Long
#End If
    On Error GoTo gestion_erreurs
    If pObject Is Nothing Then
        PixelToFontSize = PixelsToPointsY(pSize) / cPoints
    Else
        PixelToFontSize = ImgToCtrlY(pSize, pObject, False) / cPoints
    End If
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    PixelToFontSize = 0
End Function
'---------------------------------------------------------------------------------------
' Conversion code couleur Long vers ARGB
'---------------------------------------------------------------------------------------
' pAlpha     : Luminosite
' pColor     : Numero de la couleur
'---------------------------------------------------------------------------------------
'@Modified by Sancarn to remove unnecessary calls and avoid RetlMoveMemory
'  taken from LaVolpe https://www.vbforums.com/showthread.php?513275-Byte-array-gt-Long&s=95b16d2076684258a809968d89a43f20&p=3165168&viewfull=1#post3165168
Private Function GdipColor(ByVal pAlpha As Long, ByVal pColor As Long) As Long
      If pAlpha > 127 Then
        GdipColor = ((pAlpha And &H7F&) * &H1000000 Or &H80000000) Or pColor
      Else
        GdipColor = (pAlpha * &H1000000) Or pColor
      End If
End Function


'---------------------------------------------------------------------------------------
' Verifie les coordonnees et les remet dans le bon ordre si necessaire
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Sub CheckCoordinates(pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single)
#Else
Private Sub CheckCoordinates(pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long)
#End If
    Dim lTempo As Long
    If pX1 > pX2 Then
        lTempo = pX1
        pX1 = pX2
        pX2 = lTempo
    End If
    If pY1 > pY2 Then
        lTempo = pY1
        pY1 = pY2
        pY2 = lTempo
    End If
End Sub
'---------------------------------------------------------------------------------------
' Definit le refresh du formulaire avec la technique du double buffer
'---------------------------------------------------------------------------------------
' Permet de reduire le scintillement de l'image (surtout WinXP/Acc2003+)
'---------------------------------------------------------------------------------------
#If AppName = "A" Then
Public Function SetDoubleBufferXP(pForm As Access.Form, Optional pActive As Boolean = True) As Boolean
#Else
Public Function SetDoubleBufferXP(pForm As Object, Optional pActive As Boolean = True) As Boolean
#End If
Dim lNewStyle
Dim lOldStyle
Dim lComposited As Boolean
Dim lHwnd
On Error GoTo gestion_erreurs
#If AppName = "A" Then
    If pForm.PopUp Then
        lHwnd = pForm.hwnd
    Else
        lHwnd = Application.hWndAccessApp
    End If
#Else
    lHwnd = GetUserFormHandle(pForm)
#End If

lOldStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)

lComposited = (lOldStyle And WS_EX_COMPOSITED) = WS_EX_COMPOSITED
lNewStyle = lOldStyle
If Not pActive And lComposited Then
    lNewStyle = lOldStyle Xor WS_EX_COMPOSITED
ElseIf pActive And Not lComposited Then
    lNewStyle = lOldStyle Or WS_EX_COMPOSITED
End If
If lNewStyle <> lOldStyle Then
    SetWindowLong lHwnd, GWL_EXSTYLE, lNewStyle
End If
gestion_erreurs:
If Err.Number = 0 Then SetDoubleBufferXP = True    ' renvoie vrai si pas d'erreur
End Function
'---------------------------------------------------------------------------------------
' Active/Desactive le Theme XP pour l'application (controles uniquement)
'---------------------------------------------------------------------------------------
' Permet de reduire le scintillement de l'image (surtout WinXP/Acc2003)
'---------------------------------------------------------------------------------------
Public Function SetXPTheme(Optional pActive As Boolean = True) As Boolean
    On Error GoTo gestion_erreurs
    If pActive Then
        SetThemeAppProperties STAP_ALLOW_WEBCONTENT Or STAP_ALLOW_NONCLIENT Or STAP_ALLOW_CONTROLS
    Else
        SetThemeAppProperties STAP_ALLOW_WEBCONTENT Or STAP_ALLOW_NONCLIENT
    End If
gestion_erreurs:
    If Err.Number = 0 Then SetXPTheme = True    ' renvoie vrai si pas d'erreur
End Function

'---------------------------------------------------------------------------------------
' Deplacement du formulaire en deplacant la souris
' (Uniquement sur formulaire independant)
'---------------------------------------------------------------------------------------
Public Function DragForm(pForm As Object) As Boolean
    Dim lHwnd
    On Error GoTo gestion_erreurs
    #If AppName = "A" Then
        lHwnd = pForm.hwnd
    #Else
        lHwnd = GetUserFormHandle(pForm)
    #End If
    Call ReleaseCapture
    ' Simule le click sur la barre de titre
    SendMessage lHwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal PtrNull
gestion_erreurs:
    If Err.Number = 0 Then DragForm = True    ' Renvoie Vrai si pas d'erreur
End Function

'---------------------------------------------------------------------------------------
' Calcul la position et la taille de l'image dessinee dans un rectangle en fonction de
'   pImgsizeMode et pImgPictureAlignment
'---------------------------------------------------------------------------------------
' Fonction privee a usage interne uniquement
'
' En entree :
' -----------
' pX1,pY1,pX2,pY2       : Rectangle dans lequel l'image est positionnee
'      Si pImgSizeMode = GdipSizeModeAutoSize (pX1,pY1 : Centre de l'image)
' pWidth                : Largeur de l'image a positionner et redimensionner
' pHeight               : Hauteur de l'image a positionner et redimensionner
' pImgSizeMode          : Type d'image (echelle/decoupage/zoom/centree
' pImgPictureAlignment  : Position de l'image dans le controle (Centree/En Haut a Droite,..)
' pImgCtrl              :   - Si Vrai l'image est a positionner dans l'objet image
'                           - Si Faux on positionne une image dans le rectangle defini
'
' En sortie :
' -----------
' pX1,pY1,pX2,pY2       : Coordonnees de l'image dans le rectangle defini en entree
'                           (ou dans l'objet image si pImgSizeMode = GdipSizeModeAutoSize)
' lDecalageX,lDecalageY : Decalage de l'image
'   Si on travaille sur le controle (pImgCtrl = Vrai) les decalages sont l'espace
'     entre le bord de l'objet image et le bord de l'image dessinees par Excel
'   Si on travaille sur une image en memoire (pImgCtrl = Faux) les decalages sont
'       calcules pour la fonction BitBlt; c'est le decalage entre le bord de l'image
'       et le debut de l'image visible (par ex. si GdipSizeModeClip avec positionnement
'       en bas a droite on ne voit pas la partie en haut a gauche de l'image)
' pWidth                : Largeur de l'image apres positionnement et redimensionnement
' pHeight               : Hauteur de l'image apres positionnement et redimensionnement
'---------------------------------------------------------------------------------------

' Calculs en fonction du type d'image et de la position
#If UseSingle Then
Private Sub CalcXY(pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single, pWidth As Single, pHeight As Single, pDecalageX As Single, pDecalageY As Single, pImgSizeMode As Long, pImgPictureAlignment As Long, Optional pImgCtrl As Object, Optional pdX As Single, Optional pdY As Single)
    Dim lLeft As Single, lTop As Single
    Dim lRight As Single, lBottom As Single
    Dim lControlWidthPix As Single
    Dim lControlHeightPix As Single
    Dim lDecalage As Single
#Else
Private Sub CalcXY(pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long, pWidth As Long, pHeight As Long, pDecalageX As Long, pDecalageY As Long, pImgSizeMode As Long, pImgPictureAlignment As Long, Optional pImgCtrl As Object, Optional pdX As Long, Optional pdY As Long)
    Dim lLeft As Long, lTop As Long
    Dim lRight As Long, lBottom As Long
    Dim lControlWidthPix As Long
    Dim lControlHeightPix As Long
    Dim lDecalage As Long
#End If
On Error GoTo gestion_erreurs
    Dim lWidth As Single, lHeight As Single
' Cas du calcul sur l'objet image
If Not pImgCtrl Is Nothing Then
    lDecalage = GetControlPos(pImgCtrl, lLeft, lTop, lRight, lBottom)
    pdX = lDecalage
    pdY = lDecalage
    pX1 = lDecalage
    pY1 = lDecalage
    pX2 = pX1 + lRight - lLeft
    pY2 = pY1 + lBottom - lTop
    lControlWidthPix = lRight - lLeft + 1
    lControlHeightPix = lBottom - lTop + 1
    ' Dimensions de l'image
    GdipGetImageDimension gBitmap, lWidth, lHeight
    If pWidth = 0 Then pWidth = lWidth
    If pHeight = 0 Then pHeight = lHeight
End If
' Initialise les decalages
pDecalageX = 0
pDecalageY = 0
Select Case pImgSizeMode
Case GdipSizeModeAutoSize  ' Image centree sur un point
    pX1 = pX1 - pWidth \ 2
    pY1 = pY1 - pHeight \ 2
    pX2 = pX1 + pWidth - 1
    pY2 = pY2 + pHeight - 1
Case GdipSizeModeStretch
    pX1 = pX1
    pY1 = pY1
    pX2 = pX2
    pY2 = pY2
    pWidth = pX2 - pX1 + 1
    pHeight = pY2 - pY1 + 1
    pDecalageX = 0
    pDecalageY = 0
Case GdipSizeModeZoom
    If pWidth <> 0 And pHeight <> 0 Then
        If ((pWidth - pX2 + pX1 - 1)) / pWidth * pHeight < (pHeight - pY2 + pY1 - 1) Then
            pDecalageX = pX2 - pX1 + 1 - pWidth * (pY2 - pY1 + 1) / pHeight
        Else
            pDecalageY = pY2 - pY1 + 1 - pHeight * (pX2 - pX1 + 1) / pWidth
        End If
    End If
    Select Case pImgPictureAlignment
    Case 0    ' Haut-Gauche
        pX1 = pX1
        pX2 = pX2 - pDecalageX
        pY1 = pY1
        pY2 = pY2 - pDecalageY
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = 0
            pDecalageY = 0
        End If
    Case 1    ' Haut-Droiste
        pX1 = pX1 + pDecalageX
        pX2 = pX2
        pY1 = pY1
        pY2 = pY2 - pDecalageY
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = pDecalageX
            pDecalageY = 0
        End If
    Case 2    ' Centre
        pX1 = pX1 + pDecalageX \ 2
        pX2 = pX2 - pDecalageX \ 2 - pDecalageX Mod 2
        pY1 = pY1 + pDecalageY \ 2
        pY2 = pY2 - pDecalageY \ 2 - pDecalageY Mod 2
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = pDecalageX \ 2
            pDecalageY = pDecalageY \ 2
        End If
    Case 3    ' Bas-Gauche
        pX1 = pX1
        pX2 = pX2 - pDecalageX
        pY1 = pY1 + pDecalageY
        pY2 = pY2
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = 0
            pDecalageY = pDecalageY
        End If
    Case 4    ' Bas-Droite
        pX1 = pX1 + pDecalageX
        pX2 = pX2
        pY1 = pY1 + pDecalageY
        pY2 = pY2
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = pDecalageX
            pDecalageY = pDecalageY
        End If
    End Select
    pWidth = pX2 - pX1 + 1
    pHeight = pY2 - pY1 + 1
Case GdipSizeModeClip
    Select Case pImgPictureAlignment
    Case 0    ' Haut-Gauche
        pX1 = pX1
        pX2 = fMin(pX2, pX1 + pWidth - 1)
        pY1 = pY1
        pY2 = fMin(pY2, pY1 + pHeight - 1)
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = 0
            pDecalageY = 0
        Else
            ' Decalage pour bitblit
            pDecalageX = 0
            pDecalageY = 0
        End If
    Case 1    ' Haut-Droite
        pX1 = fMax(pX1, pX2 - pWidth + 1)
        pX2 = pX2
        pY1 = pY1
        pY2 = fMin(pY2, pY1 + pHeight - 1)
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = lControlWidthPix - pWidth
            pDecalageY = 0
        Else
            ' Decalage pour bitblit
            pDecalageX = pWidth - (pX2 - pX1 + 1)
            pDecalageY = 0
        End If
    Case 2    ' Centre
        ' Decalage pour calculs
        pDecalageX = pWidth - (pX2 - pX1 + 1)
        pDecalageY = pHeight - (pY2 - pY1 + 1)
        pX1 = fMax(pX1, (pX1) - pDecalageX \ 2)
        pX2 = fMin((pX2), (pX2) + pDecalageX \ 2 + pDecalageX Mod 2)
        pY1 = fMax((pY1), (pY1) - pDecalageY \ 2)
        pY2 = fMin((pY2), (pY2) + pDecalageY \ 2 + pDecalageY Mod 2)
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = (lControlWidthPix - pWidth) \ 2
            pDecalageY = (lControlHeightPix - pHeight) \ 2
        Else
            ' Decalage pour bitblit
            pDecalageX = IIf(pDecalageX > 0, pDecalageX \ 2, 0)
            pDecalageY = IIf(pDecalageY > 0, pDecalageY \ 2, 0)
        End If
    Case 3    ' Bas-Gauche
        pX1 = (pX1)
        pX2 = fMin((pX2), pX1 + pWidth - 1)
        pY2 = (pY2)
        pY1 = fMax((pY1), pY2 - pHeight + 1)
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = 0
            pDecalageY = lControlHeightPix - pHeight
        Else
            ' Decalage pour bitblit
            pDecalageX = 0
            pDecalageY = pHeight - (pY2 - pY1 + 1)
        End If
    Case 4    ' Bas-Droite
        pX1 = fMax(pX1, pX2 - pWidth + 1)
        pX2 = pX2
        pY2 = pY2
        pY1 = fMax(pY1, pY2 - pHeight + 1)
        If Not pImgCtrl Is Nothing Then
            ' Decalage pour controle
            pDecalageX = lControlWidthPix - pWidth
            pDecalageY = lControlHeightPix - pHeight
        Else
            ' Decalage pour bitblit
            pDecalageX = pWidth - (pX2 - pX1 + 1)
            pDecalageY = pHeight - (pY2 - pY1 + 1)
        End If
    End Select
    pWidth = pX2 - pX1 + 1
    pHeight = pY2 - pY1 + 1
End Select
gestion_erreurs:
If Err.Number <> 0 Then ErrRaise    'Propage l'erreur
End Sub

'---------------------------------------------------------------------------------------
' Converti les points en Pixels sur l'axe horizontal
'---------------------------------------------------------------------------------------
' pPointsX : Valeur a convertir en points
' Renvoie la valeur convertie en Pixels
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function PointsToPixelsX(ByVal pPointsX As Single) As Single
#Else
Public Function PointsToPixelsX(ByVal pPointsX As Single) As Long
#End If
    Static Mult As Single
    Dim hDC
    If Mult = 0 Then
        hDC = GetDC(0)
        #If AppName = "A" Then
            Mult = 1440 / GetDeviceCaps(hDC, LOGPIXELSX)
        #Else
            Mult = 72 / GetDeviceCaps(hDC, LOGPIXELSX)
        #End If
        ReleaseDC 0, hDC
    End If
#If UseSingle Then
    PointsToPixelsX = pPointsX / Mult
#Else
    PointsToPixelsX = CLng(pPointsX / Mult)
#End If
End Function

'---------------------------------------------------------------------------------------
' Converti les Points en Pixels sur l'axe vertical
'---------------------------------------------------------------------------------------
' pPointsY : Valeur a convertir en Points
' Renvoie la valeur convertie en Pixels
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function PointsToPixelsY(ByVal pPointsY As Single) As Single
#Else
Public Function PointsToPixelsY(ByVal pPointsY As Single) As Long
#End If
    Static Mult As Single
    Dim hDC
    If Mult = 0 Then
        hDC = GetDC(0)
        #If AppName = "A" Then
            Mult = 1440 / GetDeviceCaps(hDC, LOGPIXELSY)
        #Else
            Mult = 72 / GetDeviceCaps(hDC, LOGPIXELSY)
        #End If
        ReleaseDC 0, hDC
    End If
#If UseSingle Then
    PointsToPixelsY = pPointsY / Mult
#Else
    PointsToPixelsY = CLng(pPointsY / Mult)
#End If
End Function

'---------------------------------------------------------------------------------------
' Converti les Pixels en Points sur l'axe horizontal
'---------------------------------------------------------------------------------------
' pPixelsX : Valeur a convertir en Pixels
' Renvoie la valeur convertie en Points
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function PixelsToPointsX(ByVal pPixelsX As Single) As Single
#Else
Public Function PixelsToPointsX(ByVal pPixelsX As Long) As Single
#End If
    Static Mult As Single
    Dim hDC
    If Mult = 0 Then
        hDC = GetDC(0)
        #If AppName = "A" Then
            Mult = 1440 / GetDeviceCaps(hDC, LOGPIXELSX)
        #Else
            Mult = 72 / GetDeviceCaps(hDC, LOGPIXELSX)
        #End If
        ReleaseDC 0, hDC
    End If
    PixelsToPointsX = pPixelsX * Mult
End Function

'---------------------------------------------------------------------------------------
' Converti les Pixels en Points sur l'axe vertical
'---------------------------------------------------------------------------------------
' pPixelsY : Valeur a convertir en Pixels
' Renvoie la valeur convertie en Points
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function PixelsToPointsY(ByVal pPixelsY As Single) As Single
#Else
Public Function PixelsToPointsY(ByVal pPixelsY As Long) As Single
#End If
    Static Mult As Single
    Dim hDC
    If Mult = 0 Then
        hDC = GetDC(0)
        #If AppName = "A" Then
            Mult = 1440 / GetDeviceCaps(hDC, LOGPIXELSY)
        #Else
            Mult = 72 / GetDeviceCaps(hDC, LOGPIXELSY)
        #End If
        ReleaseDC 0, hDC
    End If
    PixelsToPointsY = pPixelsY * Mult
End Function

'---------------------------------------------------------------------------------------
' Valeur Maxi
'---------------------------------------------------------------------------------------
' Renvoie la valeur maxi entre a et b
'---------------------------------------------------------------------------------------
Private Function fMax(A As Variant, B As Variant) As Variant
    If A > B Then fMax = A Else fMax = B
End Function

'---------------------------------------------------------------------------------------
' Valeur Mini
'---------------------------------------------------------------------------------------
' Renvoie la valeur mini entre a et b
'---------------------------------------------------------------------------------------
Private Function fMin(A As Variant, B As Variant) As Variant
    If A < B Then fMin = A Else fMin = B
End Function

'---------------------------------------------------------------------------------------
' Converti une abscisse des coordonnees image vers les coordonnees controle
'  de Pixels vers points
'---------------------------------------------------------------------------------------
' pX             : Coordonnee a convertir en pixel
' pObject       : Controle Image
' pDecalage      : Calcul avec decalage (Position de l'image dans le controle)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function ImgToCtrlX(ByVal pX As Single, pObject As Object, Optional pDecalage As Boolean = True) As Single
#Else
Public Function ImgToCtrlX(ByVal pX As Long, pObject As Object, Optional pDecalage As Boolean = True) As Single
#End If
    Dim lMatrix
    Dim lX As Single
    On Error GoTo Gestion_Erreur
    lMatrix = PrivMatrixImgToCtrl(pObject)
    lX = pX
    If pDecalage Then
        PrivMatrixTransformPointSingle lX, 0, lMatrix
    Else
        PrivMatrixTransformPointSingle lX, 0, lMatrix, True
    End If
    GdipDeleteMatrix lMatrix
    ImgToCtrlX = lX
    Exit Function
Gestion_Erreur:
        ImgToCtrlX = 0    ' Renvoie 0 s'il y a une erreur
End Function
'---------------------------------------------------------------------------------------
' Converti une ordonnee des coordonnees image vers les coordonnees controle
'  de pixels vers points
'---------------------------------------------------------------------------------------
' pX             : Coordonnee a convertir en pixels
' pObject       : Controle Image
' pDecalage      : Calcul avec decalage (Position de l'image dans le controle)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function ImgToCtrlY(ByVal pY As Single, pObject As Object, Optional pDecalage As Boolean = True) As Single
#Else
Public Function ImgToCtrlY(ByVal pY As Long, pObject As Object, Optional pDecalage As Boolean = True) As Single
#End If
    Dim lMatrix
    Dim lY As Single
    On Error GoTo Gestion_Erreur
    lMatrix = PrivMatrixImgToCtrl(pObject)
    lY = pY
    If pDecalage Then
        PrivMatrixTransformPointSingle 0, lY, lMatrix
    Else
        PrivMatrixTransformPointSingle 0, lY, lMatrix, True
    End If
    GdipDeleteMatrix lMatrix
    ImgToCtrlY = lY
    Exit Function
Gestion_Erreur:
        ImgToCtrlY = 0    ' Renvoie 0 s'il y a une erreur
End Function

'---------------------------------------------------------------------------------------
' Converti une abscisse des coordonnees controle vers les coordonnees image
'  de points vers Pixels
'---------------------------------------------------------------------------------------
' pX             : Coordonnee a convertir en points
' pObject        : Controle Image
' pDecalage      : Calcul avec decalage (Position de l'image dans le controle)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CtrlToImgX(ByVal pX As Single, pObject As Object, Optional pDecalage As Boolean = True) As Single
#Else
Public Function CtrlToImgX(ByVal pX As Single, pObject As Object, Optional pDecalage As Boolean = True) As Long
#End If
    Dim lMatrix
    Dim lX As Single
    On Error GoTo Gestion_Erreur
    lMatrix = PrivMatrixCtrlToImg(pObject)
    lX = pX
    If pDecalage Then
        PrivMatrixTransformPointSingle lX, 0, lMatrix
    Else
        PrivMatrixTransformPointSingle lX, 0, lMatrix, True
    End If
    GdipDeleteMatrix lMatrix
    CtrlToImgX = lX
    Exit Function
Gestion_Erreur:
        CtrlToImgX = 0    ' Renvoie 0 s'il y a une erreur
End Function
'---------------------------------------------------------------------------------------
' Converti une ordonnee des coordonnees controle vers les coordonnees image
'  de points vers Pixels
'---------------------------------------------------------------------------------------
' pY             : Coordonnee a convertir en points
' pObject       : Controle Image
' pDecalage      : Calcul avec decalage (Position de l'image dans le controle)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CtrlToImgY(ByVal pY As Single, pObject As Object, Optional pDecalage As Boolean = True) As Single
#Else
Public Function CtrlToImgY(ByVal pY As Single, pObject As Object, Optional pDecalage As Boolean = True) As Long
#End If
    Dim lMatrix
    Dim lY As Single
    On Error GoTo Gestion_Erreur
    lMatrix = PrivMatrixCtrlToImg(pObject)
    lY = pY
    If pDecalage Then
        PrivMatrixTransformPointSingle 0, lY, lMatrix
    Else
        PrivMatrixTransformPointSingle 0, lY, lMatrix, True
    End If
    GdipDeleteMatrix lMatrix
    CtrlToImgY = lY
    Exit Function
Gestion_Erreur:
        CtrlToImgY = 0    ' Renvoie 0 s'il y a une erreur
End Function

'***************************************************************************************
'*                             DONNEES UTILISATEUR
'***************************************************************************************
Public Property Get UserData(pName As String, Optional pLevel As String) As Variant
Dim lUserData As Collection
On Error GoTo gestion_erreurs
Set lUserData = gUserData(pLevel)
If IsObject(lUserData(pName)) Then
    Set UserData = lUserData(pName)
Else
    UserData = lUserData(pName)
End If
Exit Property
gestion_erreurs:
UserData = Null
End Property
Public Property Let UserData(pName As String, Optional pLevel As String, pValue As Variant)
Dim lUserData As Collection
On Error GoTo gestion_erreurs
Set lUserData = gUserData(pLevel)
If IsNull(pValue) Then
    On Error Resume Next
    lUserData.Remove pName
    On Error GoTo gestion_erreurs
Else
    lUserData.Add pValue, pName
End If
Exit Property
gestion_erreurs:
If Err.Number = 457 Then
    lUserData.Remove pName
    Resume
End If
If Err.Number = 5 Then
    gUserData.Add New Collection, pLevel
    Resume
End If
End Property
'***************************************************************************************
'*                             TEXTURES
'***************************************************************************************
Public Property Get TextureWrap(pName As String) As Boolean
Dim lTextureWrap As Long
On Error GoTo gestion_erreurs
GdipGetTextureWrapMode gTextures(pName), lTextureWrap
TextureWrap = (lTextureWrap = 4&)
Exit Property
gestion_erreurs:
End Property
Public Property Let TextureWrap(pName As String, pValue As Boolean)
On Error GoTo gestion_erreurs
GdipSetTextureWrapMode gTextures(pName), IIf(pValue, 0&, 4&)
Exit Property
gestion_erreurs:
End Property
'---------------------------------------------------------------------------------------
' Ajoute une texture a partir d'un fichier
'---------------------------------------------------------------------------------------
' pName         : Identifiant de la texture
' pFile         : Chemin du fichier de la texture
'---------------------------------------------------------------------------------------
Public Function TextureAddFromFile(pName As String, pFile As String) As Boolean
    Dim lBitmap
    Dim lTexture
    On Error GoTo gestion_erreurs
    GdipLoadImageFromFile StrPtr(pFile), lBitmap
    GdipCreateTexture lBitmap, 0, lTexture
    GdipDisposeImage lBitmap
    If lTexture <> 0 Then
        gTextures.Add lTexture, pName
        TextureAddFromFile = True
    End If
    Exit Function
gestion_erreurs:
    ' Si la texture existait deja on l'ecrase
    If Err.Number = 457 Then TextureDelete pName: Resume
    If lBitmap <> 0 Then GdipDisposeImage lBitmap
    If lBitmap <> 0 Then GdipDeleteBrush lTexture
    TextureAddFromFile = False
End Function

'---------------------------------------------------------------------------------------
' Ajoute une texture a partir d'une image
'---------------------------------------------------------------------------------------
' pName         : Identifiant de la texture
' pImg          : Image contenant la texture (nom ou objet)
'---------------------------------------------------------------------------------------
Public Function TextureAddFromImg(pName As String, pImg As Variant) As Boolean
    Dim lBitmap
    Dim lTexture
    Dim lImg As clGdiplus
    On Error GoTo gestion_erreurs
    Set lImg = ImgFromVariant(pImg)
    lBitmap = lImg.SaveBitmap
    GdipCreateTexture lBitmap, 0, lTexture
    If lTexture <> 0 Then
        gTextures.Add lTexture, pName
        TextureAddFromImg = True
    End If
    'If lBitmap <> 0 Then GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    ' Si la texture existait deja on l'ecrase
    If Err.Number = 457 Then TextureDelete pName: Resume
    'If lBitmap <> 0 Then GdipDisposeImage lBitmap
    If lTexture <> 0 Then GdipDeleteBrush lTexture
    TextureAddFromImg = False
End Function

'---------------------------------------------------------------------------------------
' Ajoute une texture a partir de l'image integree a un controle
'---------------------------------------------------------------------------------------
Public Function TextureAddFromControl(pName As String, pCtrl As Object) As Boolean
Dim lBitmap
Dim lTexture
On Error GoTo gestion_erreurs
lBitmap = PrivLoadBitmapFromControl(pCtrl)
If lBitmap <> 0 Then
    GdipCreateTexture lBitmap, 0, lTexture
    gTextures.Add lTexture, pName
    TextureAddFromControl = True
End If
If lBitmap <> 0 Then GdipDisposeImage lBitmap
Exit Function
gestion_erreurs:
' Si la texture existait deja on l'ecrase
If Err.Number = 457 Then TextureDelete pName: Resume
' Retourne vrai si pas d'erreur
If lBitmap <> 0 Then GdipDisposeImage lBitmap
If lTexture <> 0 Then GdipDeleteBrush lTexture
TextureAddFromControl = False
End Function

'---------------------------------------------------------------------------------------
' Supprime une texture
'---------------------------------------------------------------------------------------
' pName         : Identifiant de la texture
'---------------------------------------------------------------------------------------
Public Function TextureDelete(pName As String) As Boolean
    On Error GoTo gestion_erreurs
    ' Supprime l'image gdi+
    GdipDeleteBrush gTextures.item(pName)
    ' Retire de la collection
    gTextures.Remove pName
gestion_erreurs:
    ' renvoie Vrai si pas d'erreur
    If Err.Number = 0 Then TextureDelete = True
End Function

'---------------------------------------------------------------------------------------
' Deplace une texture
'---------------------------------------------------------------------------------------
' pName         : Identifiant de la texture
' pDx, pDy      : Deplacements en pixels
'---------------------------------------------------------------------------------------
Public Function TextureTranslate(pName As String, ByVal pdX As Single, ByVal pdY As Single, Optional pRight As Boolean = True) As Boolean
    Dim lTexture As Long
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    TextureTranslate = (GdipTranslateTextureTransform(lTexture, pdX, pdY, -pRight) = 0)
    Exit Function
gestion_erreurs:
    TextureTranslate = False
End Function
Public Function TextureScale(pName As String, ByVal psX As Single, ByVal psY As Single, Optional pRight As Boolean = True) As Boolean
    Dim lTexture As Long
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    TextureScale = (GdipScaleTextureTransform(lTexture, psX, psY, -pRight) = 0)
    Exit Function
gestion_erreurs:
    TextureScale = False
End Function
Public Function TextureReset(pName As String) As Boolean
    Dim lTexture As Long
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    TextureReset = (GdipResetTextureTransform(lTexture) = 0)
    Exit Function
gestion_erreurs:
    TextureReset = False
End Function
Public Function TextureRotate(pName As String, ByVal pAngle As Single, Optional pRight As Boolean = True) As Boolean
    Dim lTexture As Long
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    TextureRotate = (GdipRotateTextureTransform(lTexture, pAngle, -pRight) = 0)
    Exit Function
gestion_erreurs:
    TextureRotate = False
End Function
Public Function TextureTransform(pName As String, Optional pMatrix As Variant, Optional pRight As Boolean = True) As Boolean
    Dim lTexture As Long
    Dim lMatrix
    Dim lgraphics
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    ' Recherche la matrice
    If IsMissing(pMatrix) Then
        lgraphics = GetGraphics()
        GdipCreateMatrix lMatrix
        GdipGetWorldTransform lgraphics, lMatrix
    Else
        lMatrix = GetMatrix(pMatrix)
    End If
    ' Applique les transformations
    TextureTransform = (GdipMultiplyTextureTransform(lTexture, lMatrix, -pRight) = 0)
    ' Supprime la matrice temporaire
    GdipDeleteMatrix lMatrix
    Exit Function
gestion_erreurs:
    TextureTransform = False
End Function
#If UseSingle Then
Public Function DrawTexture(pName As String, ByVal pX As Single, ByVal pY As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pRegion As String = "") As Boolean
#Else
Public Function DrawTexture(pName As String, ByVal pX As Long, ByVal pY As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pRegion As String = "") As Boolean
#End If
Dim lWidth As Long
Dim lHeight As Long
Dim lTexture
Dim lgraphics
Dim lBitmap
On Error GoTo gestion_erreurs
lgraphics = GetGraphics
lTexture = gTextures.item(pName)
If pX2 = cSingleOptional Or pY2 = cSingleOptional Then
    GdipGetTextureImage lTexture, lBitmap
    If pX2 = cSingleOptional Then GdipGetImageWidth lBitmap, lWidth: pX2 = pX + lWidth - 1 Else lWidth = pX2 - pX + 1
    If pY2 = cSingleOptional Then GdipGetImageHeight lBitmap, lHeight: pY2 = pY + lHeight - 1 Else lHeight = pY2 - pY + 1
    GdipDisposeImage lBitmap
Else
    lWidth = pX2 - pX + 1
    lHeight = pY2 - pY + 1
End If
GdipTranslateTextureTransform lTexture, pX, pY, 1
DrawTexture = (GdipFillRectangle(lgraphics, lTexture, pX, pY, lWidth - 1, lHeight - 1) = 0)
GdipTranslateTextureTransform lTexture, -pX, -pY, 1

' Ajout de region
If pRegion <> "" Then
    PrivDrawRectangle pX, pY, pX2, pY2, , -1, 0, , , pRegion, , , True
End If
gestion_erreurs:
    DrawTexture = False
End Function
'---------------------------------------------------------------------------------------
' Largeur d'une texture
'---------------------------------------------------------------------------------------
' pName         : Identifiant de la texture
'---------------------------------------------------------------------------------------
Public Function TextureWidth(pName As String) As Long
    Dim lTexture
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    GdipGetTextureImage lTexture, lBitmap
    GdipGetImageWidth lBitmap, TextureWidth
    GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    TextureWidth = 0
End Function

'---------------------------------------------------------------------------------------
' Hauteur d'une texture
'---------------------------------------------------------------------------------------
' pName         : Identifiant de la texture
'---------------------------------------------------------------------------------------
Public Function TextureHeight(pName As String) As Long
    Dim lTexture
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lTexture = gTextures.item(pName)
    GdipGetTextureImage lTexture, lBitmap
    GdipGetImageHeight lBitmap, TextureHeight
    GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    TextureHeight = 0
End Function

'***************************************************************************************
'*                             LISTE D'IMAGES                                          *
'***************************************************************************************
Private Function ImgFromVariant(pImg As Variant) As clGdiplus
On Error GoTo gestion_erreurs
    If IsObject(pImg) Then
        Set ImgFromVariant = pImg
    Else
        Set ImgFromVariant = img(CStr(pImg))
    End If
Exit Function
gestion_erreurs:
Set ImgFromVariant = Nothing
End Function
Public Function ImgNewTemp(Optional pName As String, Optional ByVal pWidth As Long = cLongOptional, Optional ByVal pHeight As Long = cLongOptional) As clGdiplus
Set ImgNewTemp = PrivImgNew(pName, pWidth, pHeight, , , True)
End Function
Public Function ImgNew(Optional pName As String, Optional ByVal pWidth As Long = cLongOptional, Optional ByVal pHeight As Long = cLongOptional, Optional pBefore As Variant, Optional pAfter As Variant) As clGdiplus
Set ImgNew = PrivImgNew(pName, pWidth, pHeight, pBefore, pAfter, False)
End Function
Public Function ImgNewForImage(Optional pName As String, Optional pImgRef As Variant, Optional ByVal pWidth As Long = cLongOptional, Optional ByVal pHeight As Long = cLongOptional, Optional pBefore As Variant, Optional pAfter As Variant) As clGdiplus
    Dim lImg As clGdiplus
    Set lImg = ImgFromVariant(pImgRef)
    Set ImgNewForImage = PrivImgNew(pName, pWidth, pHeight, pBefore, pAfter, False, lImg)
End Function
Private Function PrivImgNew(Optional ByVal pName As String, Optional ByVal pWidth As Long = cLongOptional, Optional ByVal pHeight As Long = cLongOptional, Optional pBefore As Variant, Optional pAfter As Variant, Optional pTemporary As Boolean = False, Optional pImg As clGdiplus) As clGdiplus
On Error GoTo gestion_erreurs
    Set PrivImgNew = New clGdiplus
    PrivImgNew.LoadAndClone = True
    If pName = "" Then pName = CStr(ObjPtr(PrivImgNew))
    PrivImgNew.ImgName = pName
    PrivImgNew.ImgParentPtr = ObjPtr(Me)
    If ImgCopyRendering Then
        PrivImgNew.AntialisingLevel = gAntialisingLevel
        PrivImgNew.SmoothingMode = gSmoothingMode
        PrivImgNew.TransparencyMode = gTransparencyMode
        PrivImgNew.TextRenderingHint = gTextRenderingHint
        PrivImgNew.PixelOffsetMode = gPixelOffsetMode
        PrivImgNew.CompositingQuality = gCompositingQuality
    End If
    If pWidth <> cLongOptional And pHeight <> cLongOptional Then
        If pImg Is Nothing Then
            PrivImgNew.CreateBitmap pWidth, pHeight
        Else
            PrivImgNew.CreateBitmapForImg pImg, pWidth, pHeight
        End If
    End If
    If Not pTemporary Then gImgList.Add PrivImgNew, pName, pBefore, pAfter
    Exit Function
gestion_erreurs:
' Si l'image existait deja on l'ecrase
If Err.Number = 457 Then ImgDelete pName: Resume
Set PrivImgNew = Nothing
End Function

Public Property Get Imgs() As Collection
Set Imgs = gImgList
End Property

Public Property Get img(Optional pName As String) As clGdiplus
On Error GoTo gestion_erreurs
If Len(pName) = 0 Then
    Set img = Me
Else
    Set img = gImgList(pName)
End If
    Exit Property
gestion_erreurs:
    Set img = Nothing
End Property

'---------------------------------------------------------------------------------------
' Teste si une image existe dans la liste d'images
'---------------------------------------------------------------------------------------
' pName         : Identifiant de l'image
'---------------------------------------------------------------------------------------
Public Function ImgExists(pName As String) As Boolean
    On Error GoTo gestion_erreurs
    ImgExists = (Not img(pName) Is Nothing)
    Exit Function
gestion_erreurs:
    ImgExists = False
End Function

'---------------------------------------------------------------------------------------
' Supprime plusieurs images de la liste d'images
'  (toutes par defaut)
'---------------------------------------------------------------------------------------
Public Function ImgsDelete(Optional pName As String = "*") As Boolean
Dim loImg As clGdiplus
On Error GoTo gestion_erreurs
    ' Vide la collection d'images
    For Each loImg In gImgList
        If loImg.ImgName Like pName Then
            gImgList.Remove loImg.ImgName
        End If
    Next
    ImgsDelete = True
    Exit Function
gestion_erreurs:
    ' renvoie Vrai si pas d'erreur
    ImgsDelete = False
End Function

'---------------------------------------------------------------------------------------
' Supprime une image de la liste d'images
'---------------------------------------------------------------------------------------
' pName         : Identifiant de l'image
'---------------------------------------------------------------------------------------
Public Function ImgDelete(pName As String) As Boolean
On Error GoTo gestion_erreurs
    gImgList.Remove pName
    ImgDelete = True
gestion_erreurs:
    ImgDelete = False
End Function

'---------------------------------------------------------------------------------------
' Clone une image de la liste d'image
'---------------------------------------------------------------------------------------
' pImgSource    : image source
' pNameDest     : image destination
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function ImgClone(pImgSource As Variant, pNameDest As String, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, _
            Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional) As clGdiplus
#Else
Public Function ImgClone(pImgSource As Variant, pNameDest As String, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, _
            Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional) As clGdiplus
#End If
    Dim lClone As clGdiplus, lImgSrc As clGdiplus
    On Error GoTo gestion_erreurs
    Set lImgSrc = ImgFromVariant(pImgSource)
    If pX1 <> cSingleOptional Or pY1 <> cSingleOptional Then
        If pX2 = cSingleOptional Then pX2 = lImgSrc.ImageWidth - 1
        If pY2 = cSingleOptional Then pY2 = lImgSrc.ImageHeight - 1
        Set lClone = ImgNew(pNameDest)
        lClone.CreateBitmapForImg pImgSource, pX2 - pX1 + 1, pY2 - pY1 + 1
        lClone.DrawImg pImgSource, -pX1, -pY1, pX2 - pX1 + 1, pY2 - pY1 + 1, , GdipSizeModeClip, GdipAlignTopLeft
        Set ImgClone = lClone
    Else
        Set lClone = ImgNew(pNameDest)
        lClone.LoadBitmap lImgSrc.SaveBitmap
        Set ImgClone = lClone
    End If
    Exit Function
gestion_erreurs:
    ImgDelete pNameDest
    Set ImgClone = Nothing
End Function

'---------------------------------------------------------------------------------------
' Affiche une image de la liste d'images
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawImg(pImg As Variant, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Public Function DrawImg(pImg As Variant, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawImg(pImg As Variant, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Public Function DrawImg(pImg As Variant, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#End If
If Not gCtrlRef Is Nothing Then
    PrivMatrixTransformPoint pX1, pY1, CtrlRefToImgMatrix
    PrivMatrixTransformPoint pX2, pY2, CtrlRefToImgMatrix
End If
DrawImg = PrivImgDraw(pImg, pX1, pY1, pX2, pY2, pTranspColor, pImgSizeMode, pImgPictureAlignment, pPercent, pRegion, pRegionColor, pRegionAlpha)
End Function

'---------------------------------------------------------------------------------------
' Affiche cette image sur une autre image
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function ImgDraw(pTarget As clGdiplus, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Public Function ImgDraw(pTarget As clGdiplus, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function ImgDraw(pTarget As clGdiplus, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Public Function ImgDraw(pTarget As clGdiplus, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                          Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#End If
' Conversion si controle de reference
If Not gCtrlRef Is Nothing Then
    PrivMatrixTransformPoint pX1, pY1, CtrlRefToImgMatrix
    PrivMatrixTransformPoint pX2, pY2, CtrlRefToImgMatrix
End If
ImgDraw = pTarget.DrawImg(Me, pX1, pY1, pX2, pY2, pTranspColor, pImgSizeMode, pImgPictureAlignment, pPercent, pRegion, pRegionColor, pRegionAlpha)
End Function

'---------------------------------------------------------------------------------------
' Affiche une image de la liste d'images
'---------------------------------------------------------------------------------------
' pName                 : Identifiant de l'image
' Si pImgSizeMode = GdipSizeModeAutoSize :
'       pX1,pY1         : Position de l'image
'       pX2             : Largeur de l'image
'       pY2             : Hauteur de l'image
' Si pImgSizeMode = GdipSizeModeZoom/GdipSizeModeStretch/GdipSizeModeClip :
'       pX1,Y1          : Point Haut-Gauche du rectangle contenant l'image
'       pX2,Y2          : Point Bas-Droite du rectangle contenant l'image
' pTranspcolor          : Couleur de transparence : les points de cette couleur ne
'                           seront pas dessines
' pImgSizeMode          : Type d'affichage de l'image
' pImgPictureAlignment  : Position de l'image
' pPercent              : Pourcentage pour affichage transclucide (0:invisible;255:normal)
' pRegion               : Nom d'une region a creer (= rectangle entourant l'image)
' pRegionColor    : Couleur pour creation de region
' pRegionAlpha    : Transparence pour creation de region
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Private Function PrivImgDraw(pImg As Variant, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                                   Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Private Function PrivImgDraw(pImg As Variant, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                                   Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#Else
    #If VBA6 Then
Private Function PrivImgDraw(pImg As Variant, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                                   Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Private Function PrivImgDraw(pImg As Variant, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                                   Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#End If

Dim lImgBitmap
Dim lImg As clGdiplus
On Error GoTo gestion_erreurs
' Bitmap de l'image demandee
Set lImg = ImgFromVariant(pImg)
lImgBitmap = lImg.SaveBitmap
PrivImgDraw = PrivDrawBitmap(lImgBitmap, pX1, pY1, pX2, pY2, pTranspColor, pImgSizeMode, pImgPictureAlignment, pPercent, pRegion, pRegionColor, pRegionAlpha, lImg)
Exit Function
gestion_erreurs:
If Err.Number <> 0 Then PrivImgDraw = False   ' renvoie Vrai si pas d'erreur
End Function

'***************************************************************************************
'*                                       REGIONS                                       *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Ajoute un chemin ou une region dans la collection de regions
'---------------------------------------------------------------------------------------
' Si chemin, les transformation World* sont appliquees
' Si region, elle ne le sont pas
'---------------------------------------------------------------------------------------
Private Function AddToRegion(pRegion As String, Optional pPath, Optional pRgn, Optional pWorldTransform As Boolean = True, _
                Optional pPenWidth As Variant = 0, Optional pLineJoin As Variant = Null, _
                Optional pLineJoinMiterLimit As Variant = Null, Optional pLineStart As Variant = Null, _
                Optional pLineEnd As Variant = Null, Optional pLineArrowLength As Variant = Null, _
                Optional pLineArrowFill As Variant = Null, Optional pPenAlignMode As Variant = Null _
                ) As Boolean
Dim lMatrix
Dim lgraphics
Dim lregion(0 To 10) As Variant
Dim lToBeAdded
Dim lRegionName As String
Const cDummyRegion As String = "ThisIsADummyRegionForAppend"
On Error GoTo gestion_erreurs
If pWorldTransform And pPath <> 0 And RegionWorldTransform Then
    lgraphics = GetGraphics
    If lgraphics <> 0 Then
        GdipCreateMatrix lMatrix
        GdipGetWorldTransform lgraphics, lMatrix
        GdipTransformPath pPath, lMatrix
        GdipDeleteMatrix lMatrix
    End If
End If
lRegionName = pRegion
'If RegionAppend And RegionExists(pRegion) Then
'    lRegionName = cDummyRegion
'Else
'    lRegionName = pRegion
'End If
If pPath <> 0 Then lToBeAdded = pPath
If pRgn <> 0 Then lToBeAdded = pRgn
lregion(0) = lToBeAdded        ' Handle de la region
lregion(1) = lRegionName     ' Identifiant de la region
lregion(2) = (pRgn = 0)     ' Objet path si vrai
lregion(3) = pPenWidth ' Taille du crayon utilise si dessin
If IsNull(pLineJoin) Then lregion(4) = LineJoin Else lregion(4) = pLineJoin
If IsNull(pLineJoinMiterLimit) Then lregion(5) = LineJoinMiterLimit Else lregion(5) = pLineJoinMiterLimit
If IsNull(pLineStart) Then lregion(6) = LineStart Else lregion(6) = pLineStart
If IsNull(pLineEnd) Then lregion(7) = LineEnd Else lregion(7) = pLineEnd
If IsNull(pLineArrowLength) Then lregion(8) = LineArrowLength Else lregion(8) = pLineArrowLength
If IsNull(pLineArrowFill) Then lregion(9) = LineArrowFill Else lregion(9) = pLineArrowFill
If IsNull(pPenAlignMode) Then lregion(10) = PenAlignMode Else lregion(10) = pPenAlignMode
gCollRegions.Add lregion, lRegionName    ' Ajout a la collection
If RegionAppend Then
    RegionAppend = False
    RegionCombine pRegion, cDummyRegion, CombineModeUnion
    RegionDelete cDummyRegion
    RegionAppend = True
End If
AddToRegion = True
Exit Function
gestion_erreurs:
If Err.Number = 457 Then
    If RegionAppend Then
        lRegionName = cDummyRegion
    Else
        RegionDelete lRegionName
    End If
    Resume
End If
If lMatrix <> 0 Then GdipDeleteMatrix lMatrix
AddToRegion = False
End Function

'---------------------------------------------------------------------------------------
' Teste si la region est un chemin
'---------------------------------------------------------------------------------------
Public Function RegionIsPath(pRegion As String) As Boolean
On Error GoTo gestion_erreurs
    RegionIsPath = gCollRegions(pRegion)(2)
    Exit Function
gestion_erreurs:
RegionIsPath = False
End Function
'---------------------------------------------------------------------------------------
' Ajoute une region line
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CreateRegionLine(pRegion As String, ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionLine(pRegion As String, ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
Dim lPath
On Error GoTo gestion_erreurs
GdipCreatePath 0, lPath
If GdipAddPathLine(lPath, pX1, pY1, pX2, pY2) = 0 Then
    CreateRegionLine = AddToRegion(pRegion, lPath, , , pPenWidth)
End If
Exit Function
gestion_erreurs:
CreateRegionLine = False
End Function

'---------------------------------------------------------------------------------------
' Ajoute une region elliptique
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a creer
' Si Ptype = 0 : On passe un rectangle en parametre, l'ellipse remplit ce rectangle
'       pX1,Y1          : Point Haut-Gauche du rectangle
'       pX2,Y2          : Point Bas-Droite du rectangle
' Si Ptype = 1 : on passe le centre et les rayons de l'ellipse en parametre
'       pX1,pY1         : Centre de l'ellipse
'       pX2             : Rayon horizontal
'       pY2             : Rayon Vertical
' pX1, pX2, pY1, pY2 : Coordonnees d'une ellipse
' pStartAngle       : Angle de demarrage d'un arc en degre
' pSweepAngle       : Angle de balayage d'un arc en degre
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function CreateRegionEllipse(pRegion As String, ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pType As EGdipTypeEllipse = 0, Optional ByVal pStartAngle As Variant, Optional ByVal pSweepAngle As Variant, Optional ByVal pPenWidth As Single = 1) As Boolean
    #Else
Public Function CreateRegionEllipse(pRegion As String, ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pType As Long = 0, Optional ByVal pStartAngle As Variant, Optional ByVal pSweepAngle As Variant, Optional ByVal pPenWidth As Single = 1) As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function CreateRegionEllipse(pRegion As String, ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pType As EGdipTypeEllipse = 0, Optional ByVal pStartAngle As Variant, Optional ByVal pSweepAngle As Variant, Optional ByVal pPenWidth As Long = 1) As Boolean
    #Else
Public Function CreateRegionEllipse(pRegion As String, ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pType As Long = 0, Optional ByVal pStartAngle As Variant, Optional ByVal pSweepAngle As Variant, Optional ByVal pPenWidth As Long = 1) As Boolean
    #End If
#End If
 On Error GoTo gestion_erreurs
   CreateRegionEllipse = PrivDrawEllipse(pX1, pY1, pX2, pY2, pType, , -1, pPenWidth, , pStartAngle, pSweepAngle, , pRegion, True)
    Exit Function
gestion_erreurs:
    CreateRegionEllipse = False
End Function

'---------------------------------------------------------------------------------------
' Ajoute une region rectangulaire
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a creer
' pX1, pX2, pY1, pY2 : Coordonnees d'un rectangle
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CreateRegionRect(pRegion As String, ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionRect(pRegion As String, ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
On Error GoTo gestion_erreurs
    CreateRegionRect = PrivDrawRectangle(pX1, pY1, pX2, pY2, , -1, pPenWidth, , , pRegion, , , True)
    Exit Function
gestion_erreurs:
    CreateRegionRect = False
End Function
'---------------------------------------------------------------------------------------
' Ajoute une region rectangulaire avec coins arrondis
'---------------------------------------------------------------------------------------
' pX3, pY3   : Taille de l'ellipse utilise pour les coins
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CreateRegionRoundRect(pRegion As String, ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, ByVal pX3 As Single, ByVal pY3 As Single, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionRoundRect(pRegion As String, ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, ByVal pX3 As Long, ByVal pY3 As Long, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
On Error GoTo gestion_erreurs
    CreateRegionRoundRect = PrivDrawRectangle(pX1, pY1, pX2, pY2, , -1, pPenWidth, , , pRegion, pX3, pY3, True)
    Exit Function
gestion_erreurs:
    CreateRegionRoundRect = False
End Function

'---------------------------------------------------------------------------------------
' Ajoute une region polygon
'---------------------------------------------------------------------------------------
' pRegion       : Nom de la region a creer
' pPoints       : Tableau de points formant le polygone : (X1,Y1,X2,Y2,X3,Y3,...)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CreateRegionPolygon(pRegion As String, ByVal pPoints As Variant, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionPolygon(pRegion As String, ByVal pPoints As Variant, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
On Error GoTo gestion_erreurs
    CreateRegionPolygon = PrivDrawPolygon(pPoints, , , pPenWidth, , , pRegion, False, True)
    Exit Function
gestion_erreurs:
    CreateRegionPolygon = False
End Function
'---------------------------------------------------------------------------------------
' Ajoute une region courbe
'---------------------------------------------------------------------------------------
' pRegion       : Nom de la region a creer
' pPoints       : Tableau de points formant la courbe : (X1,Y1,X2,Y2,X3,Y3,...)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CreateRegionCurve(pRegion As String, ByVal pPoints As Variant, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionCurve(pRegion As String, ByVal pPoints As Variant, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
On Error GoTo gestion_erreurs
    CreateRegionCurve = PrivDrawPolygon(pPoints, , , pPenWidth, , , pRegion, True, True)
    Exit Function
gestion_erreurs:
    CreateRegionCurve = False
End Function

#If UseSingle Then
Public Function CreateRegionCardinal(pRegion As String, ByVal pPoints As Variant, Optional pTension As Single = 0.5, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionCardinal(pRegion As String, ByVal pPoints As Variant, Optional pTension As Single = 0.5, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
On Error GoTo gestion_erreurs
    CreateRegionCardinal = PrivDrawPolygon(pPoints, , , pPenWidth, , , pRegion, False, True, , True, pTension)
    Exit Function
gestion_erreurs:
    CreateRegionCardinal = False
End Function

'---------------------------------------------------------------------------------------
' Ajoute une region a partir d'un chemin SVG
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a creer
' pX1, pX2, pY1, pY2 : Coordonnees d'un rectangle
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function CreateRegionSVG(pRegion As String, pPath As String, Optional ByVal pPenWidth As Single = 1) As Boolean
#Else
Public Function CreateRegionSVG(pRegion As String, pPath As String, Optional ByVal pPenWidth As Long = 1) As Boolean
#End If
On Error GoTo gestion_erreurs
    CreateRegionSVG = PrivDrawSVG(pPath, , , pPenWidth, , , pRegion, True)
    Exit Function
gestion_erreurs:
    CreateRegionSVG = False
End Function

'---------------------------------------------------------------------------------------
' Liste des regions
'---------------------------------------------------------------------------------------
Public Function Regions() As Collection
Dim lregion As Variant
Set Regions = New Collection
For Each lregion In gCollRegions
    Regions.Add CStr(lregion(1))
Next
End Function

'---------------------------------------------------------------------------------------
' Region pour clone inter classe (ne pas utiliser hors de la classe)
'---------------------------------------------------------------------------------------
Public Function Region(pRegion As String) As Variant
On Error GoTo gestion_erreurs
Region = gCollRegions(pRegion)
Exit Function
gestion_erreurs:
Region = Null
End Function

'---------------------------------------------------------------------------------------
' Remplit une region
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a remplir
' pColor             : Couleur
' pColorGradient     : Deuxieme couleur pour degrade
' pGradientVert      : Mettre a Vrai pour un degrade vertical
' pAlpha            : Transparence (0 --> 255)
'---------------------------------------------------------------------------------------
Public Function RegionFill(pRegion As String, Optional pColor As Long, Optional pColorGradient As Long = -1, Optional pGradientVert As Boolean = False, Optional pAlpha As Integer = 255) As Boolean
    Dim lBrush
    Dim lgraphics
    Dim lWidth As Single, lHeight As Single
#If UseSingle Then
    Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
    Dim lrect As RECTF
#Else
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
    Dim lrect As Rect
#End If
    Dim lPath
    On Error GoTo gestion_erreurs
    If pRegion = "" Then
        RegionFill = False
        Exit Function
    End If
    lPath = gCollRegions.item(pRegion)(0)
    ' Contexte graphique
    lgraphics = GetGraphics()
    ' Dimensions de l'image
    GdipGetImageDimension gBitmap, lWidth, lHeight
    ' Creation de la brosse
    If FillTexture <> "" Then
        On Error Resume Next
        lBrush = gTextures.item(FillTexture)
        On Error GoTo gestion_erreurs
    End If
    If lBrush = 0 Then
        If pColorGradient = -1 Then
            GdipCreateSolidFill GdipColor(pAlpha, GetColor(pColor)), lBrush
        Else
            RegionGetRect pRegion, lx1, ly1, lx2, ly2
            lrect.Left = lx1
            lrect.Top = ly1
            lrect.Right = lx2 - lx1 + 1
            lrect.Bottom = ly2 - ly1 + 1
            GdipCreateLineBrushFromRect lrect, GdipColor(pAlpha, GetColor(pColor)), GdipColor(pAlpha, GetColor(pColorGradient)), -pGradientVert, 0, lBrush
        End If
    End If
    If gCollRegions.item(pRegion)(2) Then
        GdipSetPathFillMode lPath, FillMode
        GdipFillPath lgraphics, lBrush, lPath
    Else
        GdipFillRegion lgraphics, lBrush, lPath
    End If
gestion_erreurs:
    If Err.Number = 0 Then RegionFill = True
    If FillTexture = "" Then GdipDeleteBrush lBrush
End Function

'---------------------------------------------------------------------------------------
' Hachure une region
'---------------------------------------------------------------------------------------
' pRegion       : Nom de la region a hachurer
' pColor        : Couleur
' pBackColor    : Deuxieme couleur (fond)
' pStyle        : Type de hachure
' pAlpha        : Transparence (0 --> 255)
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function RegionHatch(pRegion As String, pColor As Long, Optional pBackColor As Long = -1, Optional pStyle As EGdipHatchStyle, Optional pAlpha As Integer = 255) As Boolean
#Else
Public Function RegionHatch(pRegion As String, pColor As Long, Optional pBackColor As Long = -1, Optional pStyle As Long, Optional pAlpha As Integer = 255) As Boolean
#End If
Dim lBrush
Dim lgraphics
Dim lBitmap
Dim lPath
On Error GoTo gestion_erreurs
If pRegion = "" Then
    RegionHatch = False
    Exit Function
End If
lPath = gCollRegions.item(pRegion)(0)
' Contexte graphique
lgraphics = GetGraphics()
lBitmap = gBitmap
' Creaation de la brosse
If pBackColor <> -1 Then
    GdipCreateHatchBrush pStyle, GdipColor(pAlpha, GetColor(pColor)), GdipColor(pAlpha, GetColor(pBackColor)), lBrush
Else
    GdipCreateHatchBrush pStyle, GdipColor(pAlpha, GetColor(pColor)), GdipColor(0, 0), lBrush
End If
' Remplit la region
If gCollRegions.item(pRegion)(2) Then
    GdipSetPathFillMode lPath, FillMode
    GdipFillPath lgraphics, lBrush, lPath
Else
    GdipFillRegion lgraphics, lBrush, lPath
End If
gestion_erreurs:
If Err.Number = 0 Then RegionHatch = True
GdipDeleteBrush lBrush
End Function

'---------------------------------------------------------------------------------------
' Supprime une region
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a supprimer
'---------------------------------------------------------------------------------------
Public Function RegionDelete(pRegion As String) As Boolean
    Dim lPath
    On Error GoTo gestion_erreurs:
    lPath = gCollRegions.item(pRegion)(0)
    If gCollRegions.item(pRegion)(2) Then
        GdipDeletePath lPath
    Else
        GdipDeleteRegion lPath
    End If
    gCollRegions.Remove pRegion
gestion_erreurs:
    If Err.Number = 0 Then RegionDelete = True
End Function

'---------------------------------------------------------------------------------------
' Teste si une region existe
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a tester
'---------------------------------------------------------------------------------------
Public Function RegionExists(pRegion As String) As Boolean
    Dim lregion As Variant
    On Error GoTo gestion_erreurs:
    lregion = gCollRegions.item(pRegion)
    RegionExists = True
    Exit Function
gestion_erreurs:
    RegionExists = False
End Function

'---------------------------------------------------------------------------------------
' Supprime toutes les regions
'---------------------------------------------------------------------------------------
' pRegion   : Nom de la region a supprimer (joker * possible)
'---------------------------------------------------------------------------------------
Public Function RegionsDelete(Optional pRegion As String = "*") As Boolean
    Dim lregion As Variant
    On Error GoTo gestion_erreurs:
    ' Vide la collection de regions
    For Each lregion In gCollRegions
        If lregion(1) Like pRegion Then
            If lregion(2) Then
                GdipDeletePath lregion(0)
            Else
                GdipDeleteRegion lregion(0)
            End If
            gCollRegions.Remove lregion(1)
        End If
    Next
gestion_erreurs:
    If Err.Number = 0 Then RegionsDelete = True
End Function

'---------------------------------------------------------------------------------------
' Teste si deux regions ont une intersection
'---------------------------------------------------------------------------------------
Public Function RegionsIntersect(pRegion1 As String, pRegion2 As String, Optional pRegionIntersect As String) As Boolean
    Dim lregion1
    Dim lregion2
    Dim lgraphics
    Dim lRgn1
    Dim lResult As Long
    Dim lcpt As Long
    On Error GoTo gestion_erreurs:
    If pRegionIntersect <> "" Then RegionDelete pRegionIntersect
    lResult = False
    lregion1 = gCollRegions.item(pRegion1)(0)
    If pRegion2 <> "" Then
        lregion2 = gCollRegions.item(pRegion2)(0)
        If gCollRegions.item(pRegion1)(2) Then
            GdipCreateRegionPath lregion1, lRgn1
        Else
            GdipCloneRegion lregion1, lRgn1
        End If
        If lRgn1 <> 0 Then
            If gCollRegions.item(pRegion2)(2) Then
                Call GdipCombineRegionPath(lRgn1, lregion2, 1)
            Else
                Call GdipCombineRegionRegion(lRgn1, lregion2, 1)
            End If
            lgraphics = GetGraphics()
            Call GdipIsEmptyRegion(lRgn1, lgraphics, lResult)
            If pRegionIntersect = "" Then
                GdipDeleteRegion lRgn1
            Else
                AddToRegion pRegionIntersect, , lRgn1, False
            End If
        End If
    Else
        For lcpt = 1 To gCollRegions.count
            If gCollRegions.item(lcpt)(1) <> pRegion1 Then
                lregion2 = gCollRegions.item(lcpt)(0)
                If gCollRegions.item(pRegion1)(2) Then
                    GdipCreateRegionPath lregion1, lRgn1
                Else
                    GdipCloneRegion lregion1, lRgn1
                End If
                If lRgn1 <> 0 Then
                    If gCollRegions.item(lcpt)(2) Then
                        Call GdipCombineRegionPath(lRgn1, lregion2, 1)
                    Else
                        Call GdipCombineRegionRegion(lRgn1, lregion2, 1)
                    End If
                    lgraphics = GetGraphics()
                    Call GdipIsEmptyRegion(lRgn1, lgraphics, lResult)
                    If Not CBool(lResult) Then
                        If pRegionIntersect = "" Then
                            GdipDeleteRegion lRgn1
                        Else
                            AddToRegion pRegionIntersect, , lRgn1, False
                        End If
                        pRegion2 = gCollRegions.item(lcpt)(1)
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    RegionsIntersect = Not CBool(lResult)
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    RegionsIntersect = False
    GdipDeleteRegion lRgn1
End Function

'---------------------------------------------------------------------------------------
' Recherche la region sous un point
'---------------------------------------------------------------------------------------
' pX, pY    : Coordonnees d'un point
' La fonction renvoie le nom de la region
'---------------------------------------------------------------------------------------
Public Function GetRegionXY(ByVal pX As Single, ByVal pY As Single, Optional pInclude As Variant, Optional pExclude As Variant, Optional pObject As Variant, Optional pMatrix As Variant) As String
GetRegionXY = PrivGetRegionXY(pX, pY, pInclude, pExclude, pObject, 0, pMatrix)
End Function
#If UseSingle Then
Public Function GetRegionXYLine(ByVal pX As Single, ByVal pY As Single, Optional pLineWidth As Single = -1, Optional pInclude As Variant, Optional pExclude As Variant, Optional pObject As Variant, Optional pMatrix As Variant) As String
#Else
Public Function GetRegionXYLine(ByVal pX As Single, ByVal pY As Single, Optional pLineWidth As Long = -1, Optional pInclude As Variant, Optional pExclude As Variant, Optional pObject As Variant, Optional pMatrix As Variant) As String
#End If
GetRegionXYLine = PrivGetRegionXY(pX, pY, pInclude, pExclude, pObject, pLineWidth, pMatrix)
End Function
'---------------------------------------------------------------------------------------
' Recherche la region sous un point
'---------------------------------------------------------------------------------------
' pX, pY    : Coordonnees d'un point
' La fonction renvoie le nom de la region
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivGetRegionXY(ByVal pX As Single, ByVal pY As Single, Optional pInclude As Variant, Optional pExclude As Variant, Optional pObject As Variant, Optional ByVal pLineWidth As Single = -1, Optional pMatrix As Variant) As String
#Else
Private Function PrivGetRegionXY(ByVal pX As Single, ByVal pY As Single, Optional pInclude As Variant, Optional pExclude As Variant, Optional pObject As Variant, Optional ByVal pLineWidth As Long = -1, Optional pMatrix As Variant) As String
#End If
    Dim lregion As Variant
    Dim lcpt As Long
    Dim lCptExclude As Long
    Dim lExclude As Boolean
    Dim lret As Long
    Dim lCtrl As Object
    Dim lpen
#If UseSingle Then
    Dim lLineWidth As Single
#Else
    Dim lLineWidth As Long
#End If
    On Error GoTo gestion_erreurs
    ' Conversion si controle de reference
    If IsMissing(pObject) Then
        Set lCtrl = gCtrlRef
    Else
        Set lCtrl = pObject
    End If
    If Not lCtrl Is Nothing Then
        pX = CtrlToImgX(pX, lCtrl, True)
        pY = CtrlToImgY(pY, lCtrl, True)
    End If
    If Not IsMissing(pMatrix) Then
        MatrixTransformPointSingle pX, pY, pMatrix
    End If
    For lcpt = gCollRegions.count To 1 Step -1
        lregion = gCollRegions.item(lcpt)
        lExclude = False
        If Not IsMissing(pInclude) Then
            lExclude = True
            For lCptExclude = LBound(pInclude) To UBound(pInclude)
                If lregion(1) Like pInclude(lCptExclude) Then
                    lExclude = False
                    Exit For
                End If
            Next
        End If
        If Not lExclude Then
            If Not IsMissing(pExclude) Then
                lExclude = False
                For lCptExclude = LBound(pExclude) To UBound(pExclude)
                    If lregion(1) Like pExclude(lCptExclude) Then
                        lExclude = True
                        Exit For
                    End If
                Next
            Else
                lExclude = False
            End If
        End If
        If Not lExclude Then
            lret = 0
            If lregion(2) Then
                lLineWidth = pLineWidth
                If lLineWidth = -1 Then lLineWidth = lregion(3)
                If lLineWidth > 0 Then
                    GdipCreatePen1 0, lLineWidth, 2, lpen
                    If Not IsNull(lregion(4)) Then GdipSetPenLineJoin lpen, lregion(4)
                    If Not IsNull(lregion(5)) Then GdipSetPenMiterLimit lpen, lregion(5)
                    ' Extremites
                    SetLineCap lpen, lregion(6), lregion(7), lregion(8), lregion(9)
                    ' Alignement du crayon
                    If Not IsNull(lregion(10)) Then If lregion(10) = 1 Then GdipSetPenMode lpen, 1
                    GdipIsOutlineVisiblePathPoint lregion(0), pX, pY, lpen, ByVal PtrNull, lret
                    GdipDeletePen lpen
                Else
                    GdipIsVisiblePathPoint lregion(0), pX, pY, ByVal PtrNull, lret
                End If
            Else
                GdipIsVisibleRegionPoint lregion(0), pX, pY, ByVal PtrNull, lret
            End If
            If lret <> 0 Then
                PrivGetRegionXY = lregion(1)
                Exit For
            End If
        End If
    Next
gestion_erreurs:
    If Err.Number <> 0 Then PrivGetRegionXY = ""
End Function
'---------------------------------------------------------------------------------------
' Ajoute un chemin a un autre chemin
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region
' pRegionToAdd : Nom de la region a ajouter
'---------------------------------------------------------------------------------------
Private Function RegionAddRegion(pRegion As String, pRegionToAdd As String) As Boolean
    On Error GoTo gestion_erreurs
    If Not RegionExists(pRegion) Then
        RegionAddRegion = RegionClone(pRegionToAdd, pRegion)
    ElseIf gCollRegions.item(pRegion)(2) And gCollRegions.item(pRegionToAdd)(2) Then
        RegionAddRegion = (GdipAddPathPath(gCollRegions.item(pRegion)(0), gCollRegions.item(pRegionToAdd)(0), 0) = 0)
    End If
gestion_erreurs:
    If Err.Number <> 0 Then RegionAddRegion = False
End Function

'---------------------------------------------------------------------------------------
' Optimize une region
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region (de type chemin)
' pSeuil : seuil en pixel
'---------------------------------------------------------------------------------------
Public Function RegionOptimize(pRegion As String, Optional pSeuil As Long = 5) As Boolean
Dim lregion
Dim lNewPath
Dim lRegionIsPath As Boolean
Dim lCount As Long
Dim lNewCount As Long
Dim lNewSubCount As Long
Dim lPoints() As POINTAPI
#If UseSingle Then
Dim lNewPoints() As POINTF
Dim lNewSubPoints() As POINTF
#Else
Dim lNewPoints() As POINTAPI
Dim lNewSubPoints() As POINTAPI
#End If
Dim lPathTypes() As Byte
Dim lcpt As Long
Dim lX As Long, lY As Long
Dim lNextx As Long, lNexty As Long
Dim lLastx As Long, lLasty As Long
Dim lOldx As Long, lOldy As Long
Dim lOldPrevx As Long, lOldPrevy As Long
Dim lRegionArray As Variant
Dim lPointNeeded As Boolean
Dim lDiffAngle As Double
Dim lLastDiffAngleSign As Double
Dim lOldDiffAngleSign As Double
Dim lSeuilAngle As Long
On Error GoTo gestion_erreurs
lSeuilAngle = 10
lregion = gCollRegions.item(pRegion)(0)
lRegionIsPath = gCollRegions.item(pRegion)(2)
If lRegionIsPath Then
    GdipGetPointCount lregion, lCount
    ReDim lPoints(1 To lCount)
    ReDim lPathTypes(1 To lCount)
    GdipGetPathPointsI lregion, lPoints(1), lCount
    GdipGetPathTypes lregion, lPathTypes(1), lCount
    lOldx = cLongOptional: lOldy = cLongOptional
    lOldPrevx = cLongOptional: lOldPrevy = cLongOptional
    lOldDiffAngleSign = cLongOptional
    GdipCreatePath 0, lNewPath
    For lcpt = 1 To lCount
        lPointNeeded = False
        If lPathTypes(lcpt) = 0 Then ' PathPointTypeStart
            lNewSubCount = 0
            Erase lNewSubPoints
        ElseIf (lPathTypes(lcpt) And &H80) Then  ' PathPointTypeCloseSubpath
            lPointNeeded = True
        ElseIf (lPathTypes(lcpt) And &H1) Then ' PathPointTypeLine
            lX = lPoints(lcpt).X
            lY = lPoints(lcpt).Y
            lNextx = lPoints(lcpt - (lcpt < lCount)).X
            lNexty = lPoints(lcpt - (lcpt < lCount)).Y
            lLastx = lPoints(lcpt + (lcpt > 1)).X
            lLasty = lPoints(lcpt + (lcpt > 1)).Y
            If lcpt = 1 Or lcpt = lCount Or lPathTypes(lcpt + (lcpt > 1)) = &H0 Then
                lPointNeeded = True
            Else
                ' Angle entre les deux derniers points retenus et le point en cours
                lOldDiffAngleSign = (atan2((lOldx - lX), (lOldy - lY)) - atan2((lOldx - lOldPrevx), (lOldy - lOldPrevy)))
                ' Angle entre le point précédent, le point en cours et le point suivant
                lLastDiffAngleSign = (atan2((lX - lNextx), (lY - lNexty)) - atan2((lX - lLastx), (lY - lLasty)))
                lDiffAngle = Abs(lLastDiffAngleSign)
                ' Test si les points précédent ou suivant sont éloignés du point en cours
                If (Abs(lX - lNextx) > 3 Or Abs(lY - lNexty) > 3) Or (Abs(lX - lLastx) > 3 Or Abs(lY - lLasty) > 3) Then
                    ' Points suffisament éloignés pour que lLastDiffAngleSign soit pertinent
                    If Abs(lLastDiffAngleSign) < gPI * lSeuilAngle / 180 Or (Abs(lLastDiffAngleSign) > gPI * (180 - lSeuilAngle) / 180 And Abs(lLastDiffAngleSign) < gPI * (180 + lSeuilAngle) / 180) Then
                    Else
                        ' Si angle grand => changement de direction = on ajoute un point mais seulement si on dépasse le seuil demandé
                        If (Abs(lX - lOldx) > pSeuil Or Abs(lY - lOldy) > pSeuil) Then
                            lPointNeeded = True
                        End If
                    End If
                Else
                    ' Points rapprochés
                    ' Teste si on n'est pas en ligne droite verticale ou horizontale
                    If Abs(lX - lOldx) = 0 Or Abs(lY - lOldy) = 0 Then
                    Else
                        ' Si dépassement du seuil
                        If (Abs(lX - lOldx) > pSeuil Or Abs(lY - lOldy) > pSeuil) Then
                            ' Recherche les points précédent et suivant (delta = seuil)
                            lNextx = lPoints(lcpt - pSeuil * (lcpt < lCount - pSeuil)).X
                            lNexty = lPoints(lcpt - pSeuil * (lcpt < lCount - pSeuil)).Y
                            lLastx = lPoints(lcpt + pSeuil * (lcpt > 1 + pSeuil)).X
                            lLasty = lPoints(lcpt + pSeuil * (lcpt > 1 + pSeuil)).Y
                            ' Calcule l'angle entre les points
                            lLastDiffAngleSign = (atan2((lX - lNextx), (lY - lNexty)) - atan2((lX - lLastx), (lY - lLasty)))
                            ' Si petit angle
                            If Abs(lLastDiffAngleSign) < gPI * lSeuilAngle / 180 Or (Abs(lLastDiffAngleSign) > gPI * (180 - lSeuilAngle) / 180 And Abs(lLastDiffAngleSign) < gPI * (180 + lSeuilAngle) / 180) Then
                                ' Vérifie qu'on est toujours dans l'alignement des deux points précédents
                                lLastDiffAngleSign = (lLastDiffAngleSign - lOldDiffAngleSign)
                                If Abs(lLastDiffAngleSign) < gPI * lSeuilAngle / 180 Or (Abs(lLastDiffAngleSign) > gPI * (180 - lSeuilAngle) / 180 And Abs(lLastDiffAngleSign) < gPI * (180 + lSeuilAngle) / 180) Then
                                    ' Sinon => nouveau point
                                    lPointNeeded = True
                                End If
                            Else ' Si grand angle => nouveau point
                                lPointNeeded = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'lPointNeeded = True
        If lPointNeeded Then
            lNewCount = lNewCount + 1
            lNewSubCount = lNewSubCount + 1
            ReDim Preserve lNewPoints(1 To lNewCount)
            ReDim Preserve lNewSubPoints(1 To lNewSubCount)
            lNewPoints(lNewCount).X = lX
            lNewPoints(lNewCount).Y = lY
            lNewSubPoints(lNewSubCount).X = lX
            lNewSubPoints(lNewSubCount).Y = lY
            lOldPrevx = lOldx
            lOldPrevy = lOldy
            lOldx = lX
            lOldy = lY
        End If
        If (lPathTypes(lcpt) And &H80) Or lcpt = lCount Then
            GdipStartPathFigure lNewPath
            GdipAddPathLine2 lNewPath, lNewSubPoints(1), lNewSubCount
            GdipClosePathFigure lNewPath
        End If
    'Debug.Print lPathTypes(lcpt), lPointNeeded, lx, ly
    Next
    'For lcpt = 1 To lNewCount
    '    GdipAddPathEllipse  lNewPath, lNewPoints(lcpt).X, lNewPoints(lcpt).Y, 3, 3
    'Next
    'GdipAddPathPolygon lNewPath, lNewPoints(1), lNewCount
    lRegionArray = gCollRegions.item(pRegion)
    lRegionArray(0) = lNewPath
    GdipDeletePath lregion
    gCollRegions.Remove pRegion
    gCollRegions.Add lRegionArray, pRegion
    RegionOptimize = True
End If
Exit Function
gestion_erreurs:
    RegionOptimize = False
End Function

'---------------------------------------------------------------------------------------
' Nombre de points d'une region (de type chemin)
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region
'---------------------------------------------------------------------------------------
Public Function RegionGetPointsCount(pRegion As String) As Long
Dim lregion
Dim lRegionIsPath As Boolean
Dim lCount As Long
On Error GoTo gestion_erreurs
lregion = gCollRegions.item(pRegion)(0)
lRegionIsPath = gCollRegions.item(pRegion)(2)
If lRegionIsPath Then
    GdipGetPointCount lregion, lCount
    RegionGetPointsCount = lCount
End If
Exit Function
gestion_erreurs:
    RegionGetPointsCount = 0
End Function

'---------------------------------------------------------------------------------------
' Transforme une region (de type chemin) en un chemin qui encadre la région
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region
'---------------------------------------------------------------------------------------
Public Function RegionOutline(pRegion As String, Optional flatness As Single = 0.25) As Boolean
Dim lregion
Dim lRegionIsPath As Boolean
On Error GoTo gestion_erreurs
lregion = gCollRegions.item(pRegion)(0)
lRegionIsPath = gCollRegions.item(pRegion)(2)
If lRegionIsPath Then
    RegionOutline = (GdipWindingModeOutline(lregion, 0, flatness) = 0)
End If
Exit Function
gestion_erreurs:
    RegionOutline = False
End Function

'---------------------------------------------------------------------------------------
' Transforme les courbes d'une region (de type chemin) en segments
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region
'---------------------------------------------------------------------------------------
Public Function RegionFlatten(pRegion As String, Optional flatness As Single = 0.25) As Boolean
Dim lregion
Dim lRegionIsPath As Boolean
On Error GoTo gestion_erreurs
lregion = gCollRegions.item(pRegion)(0)
lRegionIsPath = gCollRegions.item(pRegion)(2)
If lRegionIsPath Then
    RegionFlatten = (GdipFlattenPath(lregion, 0, flatness) = 0)
End If
Exit Function
gestion_erreurs:
    RegionFlatten = False
End Function

'---------------------------------------------------------------------------------------
' Combine deux regions
'---------------------------------------------------------------------------------------
' pRegion1   : nom de la region
' pRegion2 : Nom de la region
' pCombineMode : Mode de compbinaison
'---------------------------------------------------------------------------------------
' RegionFrame ne fonctionnera plus sur pRegion1
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function RegionCombine(pRegion1 As String, pRegion2 As String, pCombineMode As EGdipCombineMode) As Boolean
#Else
Public Function RegionCombine(pRegion1 As String, pRegion2 As String, pCombineMode As Long) As Boolean
#End If
On Error GoTo gestion_erreurs
Dim lregion1
Dim lRegion1bis
Dim lregion2
Dim lRegion1IsPath As Boolean
Dim lRegion2IsPath As Boolean
If Not RegionExists(pRegion1) Then
    GdipCreateRegion lregion1
    GdipSetEmpty lregion1
    AddToRegion pRegion1, , lregion1
    lRegion1IsPath = False
Else
    lregion1 = gCollRegions.item(pRegion1)(0)
    lRegion1IsPath = gCollRegions.item(pRegion1)(2)
End If
lregion2 = gCollRegions.item(pRegion2)(0)
lRegion2IsPath = gCollRegions.item(pRegion2)(2)
' Si union de deux path => regionaddregion
If pCombineMode = 2 And lRegion1IsPath And lRegion2IsPath Then   ' CombineModeUnion = 2
    RegionCombine = RegionAddRegion(pRegion1, pRegion2)
    If RegionCombine = True Then Exit Function
End If
If lRegion1IsPath Then
    GdipCreateRegionPath lregion1, lRegion1bis
    GdipDeletePath lregion1
    gCollRegions.Remove pRegion1
    AddToRegion pRegion1, , lRegion1bis, False
    lregion1 = lRegion1bis
End If
If lRegion2IsPath Then
    RegionCombine = (GdipCombineRegionPath(lregion1, lregion2, pCombineMode) = 0)
Else
    RegionCombine = (GdipCombineRegionRegion(lregion1, lregion2, pCombineMode) = 0)
End If
gestion_erreurs:
If Err.Number <> 0 Then RegionCombine = False
End Function

'---------------------------------------------------------------------------------------
' Test si une region est vide
'---------------------------------------------------------------------------------------
Public Function RegionIsEmpty(pRegion As String) As Boolean
Dim lgraphics
Dim lResult As Long
On Error GoTo gestion_erreurs
If Not gCollRegions.item(pRegion)(2) Then
    lgraphics = GetGraphics()
    Call GdipIsEmptyRegion(gCollRegions.item(pRegion)(0), lgraphics, lResult)
End If
RegionIsEmpty = CBool(lResult)
Exit Function
gestion_erreurs:
RegionIsEmpty = False
End Function

'---------------------------------------------------------------------------------------
' Clone une region
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region a cloner
' pRegionClone : Nom de la region clone
'---------------------------------------------------------------------------------------
Public Function RegionClone(pRegion As String, pRegionclone As String, Optional pSourceGdi As clGdiplus) As Boolean
    Dim lregion As Variant
    Dim lregionClone
    On Error GoTo gestion_erreurs
    If pSourceGdi Is Nothing Then
        lregion = gCollRegions.item(pRegion)
    Else
        lregion = pSourceGdi.Region(pRegion)
    End If
    ' On supprime la region si elle existait deja
    RegionDelete pRegionclone
    ' Clone la path ou la region
    If lregion(2) Then
        GdipClonePath lregion(0), lregionClone
        AddToRegion pRegionclone, lregionClone, , False
        'gCollRegions.Add Array(lregionClone, pRegionclone, True, gCollRegions.item(pRegion)(3)), pRegionclone
        RegionClone = True
    Else
        GdipCloneRegion lregion(0), lregionClone
        AddToRegion pRegionclone, , lregionClone, False
        'gCollRegions.Add Array(lregionClone, pRegionclone, False, -1), pRegionclone
        RegionClone = True
    End If
gestion_erreurs:
    If Err.Number <> 0 Then RegionClone = False
End Function

'---------------------------------------------------------------------------------------
' Renvoi les coordonnees d'une region rectangulaire
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region
' pX1, pY1, pX2, pY2    : Coordonnees du rectangle
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function RegionGetRect(pRegion As String, pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single) As Boolean
    Dim lrect As RECTF
#Else
Public Function RegionGetRect(pRegion As String, pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long) As Boolean
    Dim lrect As Rect
#End If
    Dim lgraphics

    On Error GoTo gestion_erreurs
    If gCollRegions.item(pRegion)(2) Then
        GdipGetPathWorldBounds gCollRegions.item(pRegion)(0), lrect, ByVal PtrNull, ByVal PtrNull
    Else
        ' Contexte graphique
        lgraphics = GetGraphics()
        GdipGetRegionBounds gCollRegions.item(pRegion)(0), lgraphics, lrect
    End If
    pX1 = lrect.Left
    pX2 = lrect.Left + lrect.Right - 1
    pY1 = lrect.Top
    pY2 = lrect.Top + lrect.Bottom - 1
gestion_erreurs:
    If Err.Number = 0 Then RegionGetRect = True
End Function

#If UseSingle Then
Public Function RegionGetCenter(pRegion As String, pX As Single, pY As Single) As Boolean
Dim lPt As POINTF
#Else
Public Function RegionGetCenter(pRegion As String, pX As Long, pY As Long) As Boolean
Dim lPt As POINTAPI
#End If
Dim lBrush
On Error GoTo gestion_erreurs
If gCollRegions.item(pRegion)(2) Then
    If GdipCreatePathGradientFromPath(gCollRegions.item(pRegion)(0), lBrush) = 0 Then
        If GdipGetPathGradientCenterPoint(lBrush, lPt) = 0 Then
            pX = lPt.X
            pY = lPt.Y
            RegionGetCenter = True
        End If
    End If
End If
If RegionGetCenter = False Then
#If UseSingle Then
    Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
#Else
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
#End If
    RegionGetRect pRegion, lx1, ly1, lx2, ly2
    pX = (lx2 + lx1) / 2
    pY = (ly2 + ly1) / 2
    RegionGetCenter = True
End If
Exit Function
gestion_erreurs:
    RegionGetCenter = False
End Function
'---------------------------------------------------------------------------------------
' Teste si un point est dans une region
'---------------------------------------------------------------------------------------
' pX, pY    : Coordonnees d'un point
' pRegion   : nom de la region a tester
'---------------------------------------------------------------------------------------
Public Function PointInRegion(ByVal pX As Single, ByVal pY As Single, pRegion As String, Optional pObject As Variant) As Boolean
    Dim lret As Long
    Dim lCtrl As Object
    On Error GoTo gestion_erreurs
    ' Conversion si controle de reference
    If IsMissing(pObject) Then
        Set lCtrl = gCtrlRef
    Else
        Set lCtrl = pObject
    End If
    If Not lCtrl Is Nothing Then
        pX = CtrlToImgX(pX, lCtrl, True)
        pY = CtrlToImgY(pY, lCtrl, True)
    End If
    If gCollRegions.item(pRegion)(2) Then
        GdipIsVisiblePathPoint gCollRegions.item(pRegion)(0), pX, pY, ByVal PtrNull, lret
    Else
        GdipIsVisibleRegionPoint gCollRegions.item(pRegion)(0), pX, pY, ByVal PtrNull, lret
    End If
    PointInRegion = (lret <> 0)
gestion_erreurs:
    If Err.Number <> 0 Then PointInRegion = False
End Function
'---------------------------------------------------------------------------------------
' Teste si un point est sur la ligne d'une region
'---------------------------------------------------------------------------------------
' pX, pY    : Coordonnees d'un point
' pRegion   : nom de la region a tester
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function PointInRegionLine(ByVal pX As Single, ByVal pY As Single, pRegion As String, Optional pLineWidth As Single = -1, Optional pObject As Variant) As Boolean
    Dim lLineWidth As Single
#Else
Public Function PointInRegionLine(ByVal pX As Single, ByVal pY As Single, pRegion As String, Optional pLineWidth As Long = -1, Optional pObject As Variant) As Boolean
    Dim lLineWidth As Long
#End If
    Dim lregion As Variant
    Dim lret As Long
    Dim lpen
    Dim lCtrl As Object
    On Error GoTo gestion_erreurs
    ' Conversion si controle de reference
    If IsMissing(pObject) Then
        Set lCtrl = gCtrlRef
    Else
        Set lCtrl = pObject
    End If
    If Not lCtrl Is Nothing Then
        pX = CtrlToImgX(pX, lCtrl, True)
        pY = CtrlToImgY(pY, lCtrl, True)
    End If
    lregion = gCollRegions.item(pRegion)
    If lregion(2) Then
        lLineWidth = pLineWidth
        If lLineWidth = -1 Then If Not IsNull(lregion(3)) Then lLineWidth = lregion(3)
        If lLineWidth > 0 Then
            GdipCreatePen1 0, lLineWidth, 2, lpen
            If Not IsNull(lregion(4)) Then GdipSetPenLineJoin lpen, lregion(4)
            If Not IsNull(lregion(5)) Then GdipSetPenMiterLimit lpen, lregion(5)
            ' Extremites
            SetLineCap lpen, lregion(6), lregion(7), lregion(8), lregion(9)
            ' Alignement du crayon
            If Not IsNull(lregion(10)) Then If lregion(10) = 1 Then GdipSetPenMode lpen, 1
            GdipIsOutlineVisiblePathPoint lregion(0), pX, pY, lpen, ByVal PtrNull, lret
            GdipDeletePen lpen
        End If
    End If
    PointInRegionLine = (lret <> 0)
gestion_erreurs:
    If Err.Number <> 0 Then PointInRegionLine = False
End Function
'---------------------------------------------------------------------------------------
' Encadre une region
'---------------------------------------------------------------------------------------
' pRegion   : nom de la region a encadrer
' pPenColor : couleur du trait
' pPenWidth : taille du trait
' pAlpha    : transparence du trait (0-255)
' pDash     : pointilles
' pInside   : encadre a l'interieur si Vrai
'---------------------------------------------------------------------------------------
#If UseSingle Then
#If VBA6 Then
Public Function RegionFrame(pRegion As String, Optional pPenColor As Long, Optional ByVal pPenWidth As Single = 1, Optional pAlpha As Integer = 255, Optional pDash As EGdipDashStyle, Optional pInside As Boolean = False) As Boolean
#Else
Public Function RegionFrame(pRegion As String, Optional pPenColor As Long, Optional ByVal pPenWidth As Single = 1, Optional pAlpha As Integer = 255, Optional pDash As Long, Optional pInside As Boolean = False) As Boolean
#End If
#Else
#If VBA6 Then
Public Function RegionFrame(pRegion As String, Optional pPenColor As Long, Optional ByVal pPenWidth As Long = 1, Optional pAlpha As Integer = 255, Optional pDash As EGdipDashStyle, Optional pInside As Boolean = False) As Boolean
#Else
Public Function RegionFrame(pRegion As String, Optional pPenColor As Long, Optional ByVal pPenWidth As Long = 1, Optional pAlpha As Integer = 255, Optional pDash As Long, Optional pInside As Boolean = False) As Boolean
#End If
#End If
    Dim lgraphics, lpen, lPath
    Dim lhrgn, lNewBrush, lOldMode As Long, lHdc
    On Error GoTo gestion_erreurs
    If pRegion = "" Then
        RegionFrame = False
        Exit Function
    End If
    lPath = gCollRegions.item(pRegion)(0)
    If gCollRegions.item(pRegion)(2) Then
        ' Contexte graphique
        lgraphics = GetGraphics()
        ' Cree le crayon
        GdipCreatePen1 GdipColor(pAlpha, GetColor(pPenColor)), pPenWidth, 2, lpen
        ' Pointilles
        If pDash > 0 Then GdipSetPenDashStyle lpen, pDash
        If pDash = 5 Then GdipSetPenDashArray lpen, gDashUser(1), UBound(gDashUser) - LBound(gDashUser) + 1
        ' Alignement du crayon
        If PenAlignMode = 1 Then GdipSetPenMode lpen, 1
        ' Texture du crayon
        If PenTexture <> "" Then
            On Error Resume Next
            GdipSetPenBrushFill lpen, gTextures.item(PenTexture)
            On Error GoTo gestion_erreurs
        End If
        ' Dessine la region
        If pInside Then
            GdipSetPenMode lpen, 1
        End If
        GdipDrawPath lgraphics, lpen, lPath
        RegionFrame = True
    Else ' Si region => utilise gdi32 => pas de pAlpha ni pDash
        ' Creation d'une region gdi32
        If GdipGetRegionHRgn(lPath, GetGraphics, lhrgn) = 0 Then
            ' Cree la brosse de la couleur desiree
            lNewBrush = CreateSolidBrush(GetColor(pPenColor))
            ' DC du Graphics
            GdipGetDC GetGraphics, lHdc
            ' Mode transparent on ne veut pas tout remplir, juste le tour
            lOldMode = SetBkMode(lHdc, 1) ' 1 = TRANSPARENT
            ' Encadre la region avec un trait de taille pWidth
            FrameRgn lHdc, lhrgn, lNewBrush, pPenWidth, pPenWidth
            ' Libere les objets
            SetBkMode lHdc, lOldMode
            DeleteObject lNewBrush
            GdipReleaseDC GetGraphics, lHdc
            RegionFrame = True
        End If
    End If
gestion_erreurs:
    If Err.Number <> 0 Then RegionFrame = False
    ' Supprime objets temporaires
    GdipDeletePen lpen
End Function

'---------------------------------------------------------------------------------------
' Determine une region definie par la couleur des points
' Function privee pour le controle image et les images de la liste d'images
'---------------------------------------------------------------------------------------
' pRegion       : Identifiant de la region a creer
' pColor        : Couleur determinant la region
' pAlpha        : Transparence determinant la region
' pIncludeColor : - Vrai : La region contient les points de couleur pColor
'                 - Faux : La region contient les points de couleur differente de pColor
' pPointX, pPointY : Point de depart pour remplissage style pot de peinture
'---------------------------------------------------------------------------------------
Public Function CreateRegionFromColor(pRegion As String, Optional ByVal pColor As Long = -1, Optional ByVal pAlpha As Long = -1, Optional pIncludeColor As Boolean = False, Optional pPointX As Long = -1, Optional pPointY As Long = -1) As Boolean
Dim lBitmap
    On Error GoTo gestion_erreurs
    lBitmap = gBitmap 'CloneBppGraphics(gBitmap, , , True)
    CreateRegionFromColor = PrivCreateRegionFromColor(lBitmap, pRegion, pColor, pAlpha, pIncludeColor, pPointX, pPointY)
    If lBitmap <> gBitmap Then GdipDisposeImage lBitmap
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    CreateRegionFromColor = False
End Function

'---------------------------------------------------------------------------------------
' Redimensionne une region
'---------------------------------------------------------------------------------------
' pRegion        : Nom de la region
' pScaleX, pScaleY : Facteurs de redimensionnement
'---------------------------------------------------------------------------------------
Public Function RegionScale(pRegion As String, ByVal pScaleX As Single, ByVal pScaleY As Single, Optional pCenter As Boolean = True) As Boolean
    Dim lPathMatrix
#If UseSingle Then
    Dim lRectBefore As RECTF
#Else
    Dim lRectBefore As Rect
#End If
    Dim lgraphics
    On Error GoTo gestion_erreurs
    GdipCreateMatrix lPathMatrix
    If pCenter Then
        ' Taille region
        If gCollRegions(pRegion)(2) Then
            GdipGetPathWorldBounds gCollRegions(pRegion)(0), lRectBefore, ByVal PtrNull, ByVal PtrNull
        Else
            lgraphics = GetGraphics()
            GdipGetRegionBounds gCollRegions(pRegion)(0), lgraphics, lRectBefore
        End If
        ' Positionne la region centree sur l'origine
        GdipTranslateMatrix lPathMatrix, -lRectBefore.Left - lRectBefore.Right / 2, -lRectBefore.Top - lRectBefore.Bottom / 2, 0
    End If
    GdipScaleMatrix lPathMatrix, pScaleX, pScaleY, 1
    If pCenter Then
        ' Repositionne la region
        GdipTranslateMatrix lPathMatrix, lRectBefore.Left + lRectBefore.Right / 2, lRectBefore.Top + lRectBefore.Bottom / 2, 1
    End If
    If gCollRegions(pRegion)(2) Then
        RegionScale = (GdipTransformPath(gCollRegions(pRegion)(0), lPathMatrix) = 0)
    Else
        RegionScale = (GdipTransformRegion(gCollRegions(pRegion)(0), lPathMatrix) = 0)
    End If
    GdipDeleteMatrix lPathMatrix
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    RegionScale = False
End Function

'---------------------------------------------------------------------------------------
' Deplace une region
'---------------------------------------------------------------------------------------
' pRegion        : Nom de la region
' pdX, pdY : Nombre de pixels de deplacement
'---------------------------------------------------------------------------------------
Public Function RegionTranslate(pRegion As String, ByVal pdX As Single, ByVal pdY As Single) As Boolean
    Dim lPathMatrix
    On Error GoTo gestion_erreurs
    GdipCreateMatrix lPathMatrix
    GdipTranslateMatrix lPathMatrix, pdX, pdY, 0
    If gCollRegions(pRegion)(2) Then
        RegionTranslate = (GdipTransformPath(gCollRegions(pRegion)(0), lPathMatrix) = 0)
    Else
        RegionTranslate = (GdipTransformRegion(gCollRegions(pRegion)(0), lPathMatrix) = 0)
    End If
    GdipDeleteMatrix lPathMatrix
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    RegionTranslate = False
End Function
'---------------------------------------------------------------------------------------
' Applique une rotation a une region
'---------------------------------------------------------------------------------------
' pRegion        : Nom de la region
' pAngle         : Angle en degre
' pCenter        : Si vrai, le rectangle contenant la region est centre avant rotation
'---------------------------------------------------------------------------------------
Public Function RegionRotate(pRegion As String, ByVal pAngle As Single, Optional pCenter As Boolean = True) As Boolean
    Dim lPathMatrix
#If UseSingle Then
    Dim lRectBefore As RECTF
#Else
    Dim lRectBefore As Rect
#End If
    Dim lgraphics
    On Error GoTo gestion_erreurs
    GdipCreateMatrix lPathMatrix
    If pCenter Then
        ' Taille region
        If gCollRegions(pRegion)(2) Then
            GdipGetPathWorldBounds gCollRegions(pRegion)(0), lRectBefore, ByVal PtrNull, ByVal PtrNull
        Else
            lgraphics = GetGraphics()
            GdipGetRegionBounds gCollRegions(pRegion)(0), lgraphics, lRectBefore
        End If
        ' Position la region centree sur l'origine
        GdipTranslateMatrix lPathMatrix, -lRectBefore.Left - lRectBefore.Right / 2, -lRectBefore.Top - lRectBefore.Bottom / 2, 0
    End If
    ' Rotation
    GdipRotateMatrix lPathMatrix, pAngle, 1
    If pCenter Then
        ' Repositionne la region
        GdipTranslateMatrix lPathMatrix, lRectBefore.Left + lRectBefore.Right / 2, lRectBefore.Top + lRectBefore.Bottom / 2, 1
    End If
    If gCollRegions(pRegion)(2) Then
        RegionRotate = (GdipTransformPath(gCollRegions(pRegion)(0), lPathMatrix) = 0)
    Else
        RegionRotate = (GdipTransformRegion(gCollRegions(pRegion)(0), lPathMatrix) = 0)
    End If
    GdipDeleteMatrix lPathMatrix
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    RegionRotate = False
End Function

'---------------------------------------------------------------------------------------
' Transforme une region en fonction d'une matrice
'---------------------------------------------------------------------------------------
' pRegion        : Nom de la region
'---------------------------------------------------------------------------------------
Public Function RegionTransform(pRegion As String, Optional pMatrix As Variant) As Boolean
    Dim lMatrix
    Dim lgraphics
    On Error GoTo gestion_erreurs
    ' Recherche la matrice
    If IsMissing(pMatrix) Then
        lgraphics = GetGraphics()
        GdipCreateMatrix lMatrix
        GdipGetWorldTransform lgraphics, lMatrix
    Else
        lMatrix = GetMatrix(pMatrix)
    End If
    ' Applique les transformations
    If gCollRegions(pRegion)(2) Then
        RegionTransform = (GdipTransformPath(gCollRegions(pRegion)(0), lMatrix) = 0)
    Else
        RegionTransform = (GdipTransformRegion(gCollRegions(pRegion)(0), lMatrix) = 0)
    End If
    ' Supprime la matrice temporaire
    GdipDeleteMatrix lMatrix
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    RegionTransform = False
End Function

'---------------------------------------------------------------------------------------
' Determine une region definie par la couleur des points
' Function privee pour le controle image et les images de la liste d'images
'---------------------------------------------------------------------------------------
' pBitmap       : Bitmap de l'image
' pColor        : Couleur determinant la region
' pAlpha        : Transparence determinant la region
' pRegion       : Identifiant de la region a creer
' pIncludeColor : - Vrai : La region contient les points de couleur pColor
'                 - Faux : La region contient les points de couleur differente de pColor
' pPointX, pPointY : Point de depart pour remplissage style pot de peinture
'---------------------------------------------------------------------------------------
Private Function PrivCreateRegionFromColor(pBitmap, pRegion As String, Optional ByVal pColor As Long = -1, Optional ByVal pAlpha As Long = -1, Optional pIncludeColor As Boolean = False, Optional ByVal pPointX As Long = -1, Optional ByVal pPointY As Long = -1) As Boolean
'   **** Code d'origine VB : Jean-Philippe AMBROSINO (cf developpez.com) ****
    Dim lngX As Long, lngY As Long
    Dim lngStartX As Long, lngStartY As Long
    Dim lngEndX As Long, lngEndY As Long
    Dim lStatus As Boolean
    Dim lPixels() As Byte
    Dim lPoints() As TRgnPoints
    Dim lRestant() As Variant
    Dim lCount As Long
    Dim lCalc As Long
    Dim lColor As Long
    Dim lAlpha As Long
    Dim lBitmap
    Dim lWidthSng As Single
    Dim lHeightSng As Single
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lrect As Rect
    Dim bmpData As BitmapData
    Dim lPath
    Dim lbytestruct As argb
    Dim lColorChanged As Boolean
    On Error GoTo gestion_erreurs
    lBitmap = pBitmap
    ' Dimensions de l'image
    If GdipGetImageDimension(lBitmap, lWidthSng, lHeightSng) = 0 Then
        lWidth = CLng(lWidthSng)
        lHeight = CLng(lHeightSng)
    End If
    ' Lecture des couleurs des pixels
    lrect.Right = lWidth
    lrect.Bottom = lHeight
    If GdipBitmapLockBits(lBitmap, lrect, &H1, PixelFormat32bppARGB, bmpData) = 0 Then
        ReDim lPixels(1 To bmpData.stride * bmpData.Height)
        RtlMoveMemory lPixels(1), ByVal bmpData.scan0, bmpData.stride * bmpData.Height
        Call GdipBitmapUnlockBits(lBitmap, bmpData)
    End If
    If pPointX = -1 Or pPointY = -1 Then
        ' On demarre a (0,0)
        lngStartX = 0
        lngStartY = 0
        ' Status de depart = on est hors de la couleur demandee
        lStatus = False
        ' Boucle sur les pixels
        For lngX = 1 To lWidth
            lStatus = False
            For lngY = 1 To lHeight
                lCalc = (lWidth * (lngY - 1) + lngX) * 4 - 3
                If pColor = -1 Then
                    lColor = pColor
                Else
                    lColor = RGB(lPixels(lCalc + 2), lPixels(lCalc + 1), lPixels(lCalc))
                End If
                If pAlpha = -1 Then
                    lAlpha = pAlpha
                Else
                    lAlpha = lPixels(lCalc + 3)
                End If
                If lStatus Then
                    ' Si la couleur ne correspond pas a celle demandee ou si on arrive en bas de l'image
                    lColorChanged = (Not pIncludeColor And lColor = pColor And lAlpha = pAlpha) _
                                 Or (pIncludeColor And (lColor <> pColor Or lAlpha <> pAlpha))
                    If lColorChanged Or lngY >= lHeight Then
                        ' Fin de la sous-region
                        lngEndX = lngX - 1
                        lngEndY = lngY - 2 - (Not lColorChanged)
                        If lPath = 0 Then
                            GdipCreatePath 0, lPath
                        End If
                        Call GdipAddPathRectangleI(lPath, lngStartX, lngStartY, lngEndX - lngStartX + 1, lngEndY - lngStartY + 1)
                        ' Status = on est hors de la couleur demandee
                        lStatus = False
                    End If
                Else
                    ' Si la couleur correspond a celle demande
                    If (Not pIncludeColor And (lColor <> pColor Or lAlpha <> pAlpha)) _
                       Or (pIncludeColor And lColor = pColor And lAlpha = pAlpha) _
                       Then
                        ' On commence une nouvelle sous-region
                        lngStartX = lngX - 1
                        lngStartY = lngY - 1
                        lngEndX = lngX - 1
                        lngEndY = lngY - 1
                        ' Si fin de la colonne => region de 1 pixel
                        If lngY >= lHeight Then
                            If lPath = 0 Then
                                GdipCreatePath 0, lPath
                            End If
                            Call GdipAddPathRectangleI(lPath, lngStartX, lngStartY, lngEndX - lngStartX + 1, lngEndY - lngStartY + 1)
                            lStatus = False
                        Else
                            ' Status = on est dans la couleur demandee
                            lStatus = True
                        End If
                    End If
                End If
            Next
        Next
    Else
        ' Conversion si controle de reference
        Dim lPointX As Single, lPointY As Single
        If Not gCtrlRef Is Nothing Then
            lPointX = pPointX
            lPointY = pPointY
            PrivMatrixTransformPointSingle lPointX, lPointY, CtrlRefToImgMatrix
            pPointX = lPointX
            pPointY = lPointY
        End If
        ' Initialise le tableau de points
        ReDim lPoints(0 To lWidth + 1, 0 To lHeight + 1)
        ' Ajoute le point de depart
        ReDim lRestant(0 To 1) As Variant
        pPointX = pPointX + 1
        pPointY = pPointY + 1
        lRestant(1) = Array(pPointX, pPointY)
        lPoints(pPointX, pPointY).o = True
        ' Boucle de recherche des points a inclure dans la region
        Do
            ' Si pas de couleur et pas de transparence precisees, on prend celles du point
            If pColor = -1 And pAlpha = -1 Then
                Call GdipBitmapGetPixel(lBitmap, pPointX - 1, pPointY - 1, lColor)
                RtlMoveMemory lbytestruct, lColor, 4
                pColor = RGB(lbytestruct.red, lbytestruct.green, lbytestruct.blue)
                pAlpha = lbytestruct.Alpha
            End If
            ' Verifie qu'il reste des points a traiter
            lCount = UBound(lRestant)
            If lCount = 0 Then Exit Do
            ' Coordonnees du premier point a traiter
            lngX = lRestant(1)(0)
            lngY = lRestant(1)(1)
            ' Supprime le point en cours de traitement
            If UBound(lRestant) > 1 Then
                lRestant(1) = lRestant(UBound(lRestant))
            End If
            ReDim Preserve lRestant(0 To UBound(lRestant) - 1)
            ' Verifie qu'on est dans les limites du tableau
            If lngX > 0 And lngX <= lWidth And lngY > 0 And lngY <= lHeight Then
                lCalc = (lWidth * (lngY - 1) + lngX) * 4 - 3
                If pColor = -1 Then
                    lColor = pColor
                Else
                    lColor = RGB(lPixels(lCalc + 2), lPixels(lCalc + 1), lPixels(lCalc))
                End If
                If pAlpha = -1 Then
                    lAlpha = pAlpha
                Else
                    lAlpha = lPixels(lCalc + 3)
                End If
                ' Verifie la couleur du pixel
                If (lColor = pColor And lAlpha = pAlpha And True) Or _
                   ((lColor <> pColor Or lAlpha <> pAlpha) And Not True) Then
                    ' Premier parametre indique que le point a ete traite
                    lPoints(lngX, lngY).T = True
                    ' Ajoute a la liste les points adjacents
                    If Not lPoints((lngX + 1), lngY).o Then
                        lPoints(lngX + 1, lngY).o = True
                        ReDim Preserve lRestant(0 To UBound(lRestant) + 1)
                        lRestant(UBound(lRestant)) = Array(lngX + 1, lngY)
                    End If
                    If Not lPoints((lngX), (lngY + 1)).o Then
                        lPoints(lngX, lngY + 1).o = True
                        ReDim Preserve lRestant(0 To UBound(lRestant) + 1)
                        lRestant(UBound(lRestant)) = Array(lngX, lngY + 1)
                    End If
                    If Not lPoints((lngX - 1), (lngY)).o Then
                        lPoints(lngX - 1, lngY).o = True
                        ReDim Preserve lRestant(0 To UBound(lRestant) + 1)
                        lRestant(UBound(lRestant)) = Array(lngX - 1, lngY)
                    End If
                    If Not lPoints((lngX), (lngY - 1)).o Then
                        lPoints(lngX, lngY - 1).o = True
                        ReDim Preserve lRestant(0 To UBound(lRestant) + 1)
                        lRestant(UBound(lRestant)) = Array(lngX, lngY - 1)
                    End If
                End If
            End If
        Loop
        ' Creation de la region
        lStatus = False
        lngStartX = 0
        lngStartY = 0
        For lngX = 1 To lWidth
            lStatus = False
            For lngY = 1 To lHeight
                If lStatus Then
                    If Not lPoints(lngX, lngY).T = pIncludeColor Or lngY >= lHeight Then
                        ' Fin de la sous-region
                        lngEndX = lngX - 1
                        lngEndY = lngY - 2 - ((lPoints(lngX, lngY).T) = pIncludeColor)
                        If lPath = 0 Then
                            GdipCreatePath 0, lPath
                        End If
                        Call GdipAddPathRectangleI(lPath, lngStartX, lngStartY, lngEndX - lngStartX + 1, lngEndY - lngStartY + 1)
                        ' Status = on est hors de la couleur demandee
                        lStatus = False
                    End If
                Else
                    If lPoints(lngX, lngY).T = pIncludeColor Then
                        ' On commence une nouvelle sous-region
                        lngStartX = lngX - 1
                        lngStartY = lngY - 1
                        lngEndX = lngX - 1
                        lngEndY = lngY - 1
                        ' Si fin de la colonne => region de 1 pixel
                        If lngY >= lHeight Then
                            If lPath = 0 Then
                                GdipCreatePath 0, lPath
                            End If
                            Call GdipAddPathRectangleI(lPath, lngStartX, lngStartY, lngEndX - lngStartX + 1, lngEndY - lngStartY + 1)
                            lStatus = False
                        Else
                            lStatus = True
                        End If
                    End If
                End If
            Next
        Next
    End If
    ' Renvoit Vrai si OK
    PrivCreateRegionFromColor = AddToRegion(pRegion, lPath, , False)
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    PrivCreateRegionFromColor = False
End Function

'---------------------------------------------------------------------------------------
' Limite l'affichage de formulaire a une region
'---------------------------------------------------------------------------------------
' pForm         : Le formulaire a limiter
' pRegion       : Nom de la region limite
' pCtrlRef      : Controle image de reference
'---------------------------------------------------------------------------------------
#If AppName = "A" Then
Public Function SetFormRegion(pForm As Access.Form, pRegion As String, Optional pCtrlRef As Access.Control, Optional pInverseRegion As Boolean, Optional ByVal pCorrectionLeft As Single, Optional ByVal pCorrectionTop As Single, Optional ByVal pCorrectionRight As Single, Optional ByVal pCorrectionBottom As Single) As Boolean
#Else
Public Function SetFormRegion(pForm As Object, pRegion As String, Optional pCtrlRef As MSForms.Control, Optional pInverseRegion As Boolean, Optional ByVal pCorrectionLeft As Single, Optional ByVal pCorrectionTop As Single, Optional ByVal pCorrectionRight As Single, Optional ByVal pCorrectionBottom As Single) As Boolean
#End If
Dim lFormHwnd
Dim lTempoRegionGdiplus
Dim lTempoRegionGdi
Dim lgraphics
Dim lPathMatrix
#If UseSingle Then
Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
Dim lDecalageX As Single, lDecalageY As Single
Dim lrefX1 As Single, lrefY1 As Single
#Else
Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
Dim lDecalageX As Long, lDecalageY As Long
Dim lrefX1 As Long, lrefY1 As Long
#End If
Dim lImgWidth As Single, lImgHeight As Single
Dim lScaleX As Single, lScaleY As Single
Dim lTranslateX As Single, lTranslateY As Single
Dim lrect As RECTF
Dim lFormRect As Rect

On Error GoTo gestion_erreurs

#If AppName = "A" Then
    lFormHwnd = pForm.hwnd
#Else
    lFormHwnd = GetUserFormHandle(pForm)
#End If

' Transforme le chemin (=path) en region gdiplus, ou clone la region
If gCollRegions.item(pRegion)(2) Then
    GdipCreateRegionPath gCollRegions.item(pRegion)(0), lTempoRegionGdiplus
Else
    GdipCloneRegion gCollRegions.item(pRegion)(0), lTempoRegionGdiplus
End If
' Cree une matrix de tranformation
GdipCreateMatrix lPathMatrix
' Decalage controle et formulaire
lScaleX = 1
lScaleY = 1
If pCtrlRef Is Nothing Then
    ' Pas de controle defini, decalage formulaire uniquement
    lTranslateX = FormDecalageX(pForm) - 1
    lTranslateY = FormDecalageY(pForm) - 1
Else
    ' Si un controle est defini, on decale en fonction de ce controle
    GdipGetImageDimension gBitmap, lImgWidth, lImgHeight
    lx2 = lImgWidth
    ly2 = lImgHeight
    CalcXY lx1, ly1, lx2, ly2, CLng(lImgWidth), CLng(lImgHeight), lDecalageX, lDecalageY, PictureSizeMode(pCtrlRef), PictureAlignment(pCtrlRef), pCtrlRef
    ' Decalage controle + formulaire
    GetControlPos pCtrlRef, lrefX1, lrefY1, 0, 0, False
    lTranslateX = lrefX1 + FormDecalageX(pForm) - 1
    lTranslateY = lrefY1 + FormDecalageY(pForm) - 1
    ' Redimensionne la region si mode d'affichage zoom ou etendu
    If PictureSizeMode(pCtrlRef) <> GdipSizeModeClip Then
        lScaleX = (lx2 - lx1 + 1) / lImgWidth
        lScaleY = (ly2 - ly1 + 1) / lImgHeight
    End If
    ' Decalage du a la position de l'image dans le controle
    lTranslateX = lTranslateX + lDecalageX
    lTranslateY = lTranslateY + lDecalageY
End If
' Applique les transformations
GdipScaleMatrix lPathMatrix, lScaleX, lScaleY, 0
GdipTranslateMatrix lPathMatrix, lTranslateX, lTranslateY, 1
GdipTransformRegion lTempoRegionGdiplus, lPathMatrix
' Supprime la matrice
GdipDeleteMatrix lPathMatrix
' Lecture du contexte d'affichage
lgraphics = GetGraphics()
' Corrections manuelles
If pCorrectionRight <> 0 Or pCorrectionBottom <> 0 Or _
   pCorrectionLeft <> 0 Or pCorrectionTop <> 0 Then
    ' Dimensions de la region
    GdipGetRegionBoundsS lTempoRegionGdiplus, lgraphics, lrect
    ' Cree une matrix de tranformation
    GdipCreateMatrix lPathMatrix
    lScaleX = (lrect.Right + pCorrectionRight - pCorrectionLeft) / lrect.Right
    lScaleY = (lrect.Bottom + pCorrectionBottom - pCorrectionTop) / lrect.Bottom
    lTranslateX = pCorrectionLeft
    lTranslateY = pCorrectionTop
    If lScaleX <> 0 Or lScaleY <> 0 Then
        GdipTranslateMatrix lPathMatrix, -lrect.Left, -lrect.Top, 0
        GdipScaleMatrix lPathMatrix, lScaleX, lScaleY, 1
        GdipTranslateMatrix lPathMatrix, lrect.Left, lrect.Top, 1
    End If
    GdipTranslateMatrix lPathMatrix, lTranslateX, lTranslateY, 1
    ' Applique les transformations
    GdipTransformRegion lTempoRegionGdiplus, lPathMatrix
    ' Supprime la matrice
    GdipDeleteMatrix lPathMatrix
End If
If pInverseRegion Then
    GetWindowRect lFormHwnd, lFormRect
    lFormRect.Right = lFormRect.Right - lFormRect.Left + 1
    lFormRect.Bottom = lFormRect.Bottom - lFormRect.Top + 1
    lFormRect.Left = 0
    lFormRect.Top = 0
    GdipCombineRegionRectI lTempoRegionGdiplus, lFormRect, 3
End If
' Creation d'une region gdi32
GdipGetRegionHRgn lTempoRegionGdiplus, lgraphics, lTempoRegionGdi
' Applique la region au formulaire
SetFormRegion = (SetWindowRgn(lFormHwnd, lTempoRegionGdi, True) > 0)
' Suppression des objets
DeleteObject lTempoRegionGdi
GdipDeleteRegion lTempoRegionGdiplus
' Ajoute le formulaire a la collection de formulaires a retablir a la fermeture
On Error Resume Next
If SetFormRegion Then gFormRegionHwnd.Add lFormHwnd
On Error GoTo 0
Exit Function
gestion_erreurs:
    ' Suppression des objets
    If lTempoRegionGdi <> 0 Then DeleteObject lTempoRegionGdi
    If lTempoRegionGdiplus <> 0 Then GdipDeleteRegion lTempoRegionGdiplus
    If lPathMatrix <> 0 Then GdipDeleteMatrix lPathMatrix
    SetFormRegion = False
End Function

'---------------------------------------------------------------------------------------
' Retablit l'affichage du formulaire complet
'---------------------------------------------------------------------------------------
' pForm         : Le formulaire a retablir
' Le formulaire est automatiquement retabli lorsque la classe est liberee
' Renvoie Vrai si pas d'erreur
'---------------------------------------------------------------------------------------
#If AppName = "A" Then
Public Function ResetFormRegion(pForm As Access.Form) As Boolean
#Else
Public Function ResetFormRegion(pForm As Object) As Boolean
#End If
Dim lFormHwnd
On Error GoTo gestion_erreurs
#If AppName = "A" Then
    lFormHwnd = pForm.hwnd
#Else
    lFormHwnd = GetUserFormHandle(pForm)
#End If
' on met 0 comme region pour retablir l'affichage complet du formulaire
ResetFormRegion = SetWindowRgn(lFormHwnd, 0, True)
' Retire le formulaire de la collection de formulaires a retablir a la fermeture
On Error Resume Next
If ResetFormRegion Then gFormRegionHwnd.Remove lFormHwnd
On Error GoTo 0
Exit Function
gestion_erreurs:
'gFormRegion = True
ResetFormRegion = False
End Function

' Recherche le vecteur normal a un point d'une region
' Le point de reference (pRefPoint*) sert a choisir la normale dirige vers ce point
#If UseSingle Then
Public Function NormalRegionPoint(ByVal pX As Single, ByVal pY As Single, pRegion As String, _
                    pNormalX As Single, pNormalY As Single, _
                    Optional ByVal pRefPointX As Single = cSingleOptional, Optional ByVal pRefPointY As Single = cSingleOptional, _
                    Optional pRayon As Single, Optional pScale As Single) As Boolean
#Else
Public Function NormalRegionPoint(ByVal pX As Long, ByVal pY As Long, pRegion As String, _
                    pNormalX As Single, pNormalY As Single, _
                    Optional ByVal pRefPointX As Long = cSingleOptional, Optional ByVal pRefPointY As Long = cSingleOptional, _
                    Optional pRayon As Long, Optional pScale As Long) As Boolean
#End If
Dim lPtHit As POINTF
Dim lRefPt As POINTF
#If UseSingle Then
Dim lR1 As RECTF, lR2 As RECTF
Dim lR1X1 As Single, lR1Y1 As Single, lR1X2 As Single, lR1Y2 As Single
Dim lR2X1 As Single, lR2Y1 As Single, lR2X2 As Single, lR2Y2 As Single
Dim lHitX1 As Single, lHitY1 As Single, lHitX2 As Single, lHitY2 As Single
#Else
Dim lR1 As Rect, lR2 As Rect
Dim lR1X1 As Long, lR1Y1 As Long, lR1X2 As Long, lR1Y2 As Long
Dim lR2X1 As Long, lR2Y1 As Long, lR2X2 As Long, lR2Y2 As Long
Dim lHitX1 As Long, lHitY1 As Long, lHitX2 As Long, lHitY2 As Long
#End If
Dim lScal As Single
Dim lMatrix, lMatrixHit
Dim lPath, lregion, lHitRegion, lHitRegion1, lHitRegion2
Dim lPt As POINTF
Dim lgraphics, lImage
Dim lBool As Boolean
Dim lret As Long
Dim lRayon As Long
Dim lScale As Single
On Error GoTo gestion_erreurs
lRayon = pRayon
If lRayon = 0 Then lRayon = 5
lScale = pScale
If lScale <= 0 Then lScale = 1
lPtHit.X = pX: lPtHit.Y = pY
lRefPt.X = pRefPointX: lRefPt.Y = pRefPointY

GdipCreateMatrix lMatrixHit
GdipScaleMatrix lMatrixHit, lScale, lScale, 1
If gCollRegions.item(pRegion)(2) Then
    GdipClonePath gCollRegions.item(pRegion)(0), lregion
    GdipTransformPath lregion, lMatrixHit
Else
    GdipCloneRegion gCollRegions.item(pRegion)(0), lregion
    GdipTransformRegion lregion, lMatrixHit
End If
GdipTransformMatrixPointsS lMatrixHit, lPtHit, 1

GdipCreatePath 0, lPath
GdipAddPathEllipseS lPath, lPtHit.X - (lRayon * lScale), lPtHit.Y - (lRayon * lScale), 2 * (lRayon * lScale), 2 * (lRayon * lScale)
GdipCreateRegionPath lPath, lHitRegion
GdipDeletePath lPath
 
GdipCloneRegion lHitRegion, lHitRegion1
GdipCloneRegion lHitRegion, lHitRegion2
If gCollRegions.item(pRegion)(2) Then
    GdipCombineRegionPath lHitRegion1, lregion, 1
    GdipCombineRegionPath lHitRegion2, lregion, 4
Else
    GdipCombineRegionRegion lHitRegion1, lregion, 1
    GdipCombineRegionRegion lHitRegion2, lregion, 4
End If

If GdipCreateBitmapFromScan0(1, 1, 4, PixelFormat32bppARGB, ByVal PtrNull, lImage) = 0 Then
    If GdipGetImageGraphicsContext(lImage, lgraphics) = 0 Then
        If GdipGetRegionBounds(lHitRegion1, lgraphics, lR1) = 0 Then
            lR1X1 = lR1.Left
            lR1X2 = lR1.Left + lR1.Right
            lR1Y1 = lR1.Top
            lR1Y2 = lR1.Top + lR1.Bottom
        End If
        If GdipGetRegionBounds(lHitRegion2, lgraphics, lR2) = 0 Then
            lR2X1 = lR2.Left
            lR2X2 = lR2.Left + lR2.Right
            lR2Y1 = lR2.Top
            lR2Y2 = lR2.Top + lR2.Bottom
        End If
        GdipDeleteGraphics lgraphics
    End If
    GdipDisposeImage lImage
End If

' Recherche les points d'intersection entre les 2 rectangles
If lR1X1 = lR2X1 And lR1X2 <> lR2X2 Then
    lR1X1 = lR1X1 - 1
ElseIf lR1X2 = lR2X2 And lR1X1 <> lR2X1 Then
    lR1X2 = lR1X2 - 1
End If
If lR1Y1 = lR2Y1 And lR1Y2 <> lR2Y2 Then
    lR1Y1 = lR1Y1 - 1
ElseIf lR1Y2 = lR2Y2 And lR1Y1 <> lR2Y1 Then
    lR1Y2 = lR1Y2 - 1
End If
If lR1Y2 < lR2Y1 Or lR1Y1 > lR2Y2 Then
    lHitX1 = lR1X1: lHitX2 = lR1X2: lHitY1 = lR1Y1: lHitY2 = lR1Y1
ElseIf lR1X2 < lR2X1 Or lR1X1 > lR2X2 Then
    lHitX1 = lR1X1: lHitX2 = lR1X1: lHitY1 = lR1Y1: lHitY2 = lR1Y2
Else
    If lR1X1 >= lR2X1 And lR1X1 <= lR2X2 Then
        If lR1Y1 <= lR2Y1 And lR1Y2 >= lR2Y1 Then RectangleAddIntersect lR1X1, lR2Y1, lHitX1, lHitY1, lHitX2, lHitY2
        If lR1Y1 <= lR2Y2 And lR1Y2 >= lR2Y2 Then RectangleAddIntersect lR1X1, lR2Y2, lHitX1, lHitY1, lHitX2, lHitY2
    End If
    If lR1X2 >= lR2X1 And lR1X2 <= lR2X2 Then
        If lR1Y1 <= lR2Y1 And lR1Y2 >= lR2Y1 Then RectangleAddIntersect lR1X2, lR2Y1, lHitX1, lHitY1, lHitX2, lHitY2
        If lR1Y1 <= lR2Y2 And lR1Y2 >= lR2Y2 Then RectangleAddIntersect lR1X2, lR2Y2, lHitX1, lHitY1, lHitX2, lHitY2
    End If
    If lR1Y1 >= lR2Y1 And lR1Y1 <= lR2Y2 Then
        If lR1X1 <= lR2X1 And lR1X2 >= lR2X1 Then RectangleAddIntersect lR2X1, lR1Y1, lHitX1, lHitY1, lHitX2, lHitY2
        If lR1X1 <= lR2X2 And lR1X2 >= lR2X2 Then RectangleAddIntersect lR2X2, lR1Y1, lHitX1, lHitY1, lHitX2, lHitY2
    End If
    If lR1Y2 >= lR2Y1 And lR1Y2 <= lR2Y2 Then
        If lR1X1 <= lR2X1 And lR1X2 >= lR2X1 Then RectangleAddIntersect lR2X1, lR1Y2, lHitX1, lHitY1, lHitX2, lHitY2
        If lR1X1 <= lR2X2 And lR1X2 >= lR2X2 Then RectangleAddIntersect lR2X2, lR1Y2, lHitX1, lHitY1, lHitX2, lHitY2
    End If
End If
' Choisi la normale parmi les 2 possibles
If lRefPt.X <> cSingleOptional And lRefPt.Y <> cSingleOptional Then
    GdipTransformMatrixPointsS lMatrixHit, lRefPt, 1
    lScal = (lPtHit.X - lRefPt.X) * (lHitY2 - lHitY1) + (lPtHit.Y - lRefPt.Y) * -(lHitX2 - lHitX1)
    If lScal < 0 Then
        pNormalX = lHitY2 - lHitY1
        pNormalY = -(lHitX2 - lHitX1)
    Else
        pNormalX = -(lHitY2 - lHitY1)
        pNormalY = (lHitX2 - lHitX1)
    End If
Else
    If gCollRegions.item(pRegion)(2) Then
        GdipIsVisiblePathPoint lregion, lPtHit.X - (lHitY2 - lHitY1), lPtHit.Y + (lHitX2 - lHitX1), ByVal PtrNull, lret
        lBool = (lret <> 0)
    Else
        GdipIsVisibleRegionPoint lregion, lPtHit.X - (lHitY2 - lHitY1), lPtHit.Y + (lHitX2 - lHitX1), ByVal PtrNull, lret
        lBool = (lret <> 0)
    End If
    If lBool Then
        pNormalX = lHitY2 - lHitY1
        pNormalY = -(lHitX2 - lHitX1)
    Else
        pNormalX = -(lHitY2 - lHitY1)
        pNormalY = (lHitX2 - lHitX1)
    End If
End If

GdipDeleteRegion lHitRegion1
GdipDeleteRegion lHitRegion2
GdipDeleteRegion lHitRegion
GdipDeleteRegion lregion
GdipDeleteMatrix lMatrixHit

' Normalise le vecteur (longueur = 1)
GdipCreateMatrix lMatrix
GdipScaleMatrix lMatrix, 1 / Sqr(pNormalX ^ 2 + pNormalY ^ 2), 1 / Sqr(pNormalX ^ 2 + pNormalY ^ 2), 0
lPt.X = pNormalX
lPt.Y = pNormalY
GdipVectorTransformMatrixPointsS lMatrix, lPt, 1
GdipDeleteMatrix lMatrix
pNormalX = lPt.X
pNormalY = lPt.Y
NormalRegionPoint = True
Exit Function

gestion_erreurs:
NormalRegionPoint = False
End Function

#If UseSingle Then
Private Sub RectangleAddIntersect(pX As Single, pY As Single, pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single)
#Else
Private Sub RectangleAddIntersect(pX As Long, pY As Long, pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long)
#End If
    If pX1 = 0 And pY1 = 0 Then
        pX1 = pX: pY1 = pY
    ElseIf pX2 = 0 And pY2 = 0 Then
        pX2 = pX: pY2 = pY
    End If
End Sub

' Recherche le premier point d'intersection entre un segment et une region
#If UseSingle Then
Public Function RayCast(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, _
                        pRegion As String, pRayX As Long, pRayY As Long _
                        ) As Boolean
#Else
Public Function RayCast(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, _
                        pRegion As String, pRayX As Long, pRayY As Long _
                        ) As Boolean
#End If
' Parcours le segment avec un algorithm de Bresenham
' Adapte de http://www.codeproject.com/Articles/15604/Ray-casting-in-a-2D-tile-based-environment

Dim lsteep As Boolean
Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
Dim ldeltax As Long, ldeltay As Long, lerror As Long, lxstep As Long, lystep As Long, lY As Long, lX As Long
Dim lPtX As Long, lPtY As Long
On Error GoTo gestion_erreurs
lsteep = Abs(pY2 - pY1) > Abs(pX2 - pX1)
If lsteep Then
    lx1 = pY1:    lx2 = pY2:    ly1 = pX1:    ly2 = pX2
Else
    lx1 = pX1:    lx2 = pX2:    ly1 = pY1:    ly2 = pY2
End If
If lx1 > lx2 Then lxstep = -1 Else lxstep = 1
ldeltax = lx2 - lx1
ldeltay = Abs(ly2 - ly1)
lerror = 0
lystep = 0
lY = ly1
If (ly1 < ly2) Then lystep = 1 Else lystep = -1
For lX = lx1 To lx2 Step lxstep
    If lsteep Then
        lPtX = lY:        lPtY = lX
    Else
        lPtX = lX:        lPtY = lY
    End If
    If PointInRegion(lPtX, lPtY, pRegion) Then
        pRayX = lPtX
        pRayY = lPtY
        RayCast = True
        Exit Function
    End If
    lerror = lerror + ldeltay
    If (2 * lerror >= ldeltax) Then
        lY = lY + lystep
        lerror = lerror - ldeltax
    End If
Next
Exit Function
gestion_erreurs:
RayCast = False
End Function

'***************************************************************************************
'*                             DESSIN / TRANSFORMATIONS                                *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Prive : defini le type de debut et fin des lignes pour un crayon
'---------------------------------------------------------------------------------------
Private Sub SetLineCap(pPen, Optional ByVal pLineStart As Variant = Null, Optional ByVal pLineEnd As Variant = Null, _
                        Optional ByVal pLineArrowLength As Variant = Null, Optional ByVal pLineArrowFill As Variant = Null)
Dim lccap
If IsNull(pLineStart) Then pLineStart = LineStart
If IsNull(pLineEnd) Then pLineEnd = LineEnd
If IsNull(pLineArrowLength) Then pLineArrowLength = LineArrowLength
If IsNull(pLineArrowFill) Then pLineArrowFill = LineArrowFill
GdipSetPenStartCap pPen, pLineStart
If pLineStart = LineCapArrowAnchor And pLineArrowLength > 0 Then
    GdipCreateAdjustableArrowCap pLineArrowLength, pLineArrowLength, pLineArrowFill, lccap
    GdipSetPenCustomStartCap pPen, lccap
    GdipDeleteCustomLineCap lccap
End If
GdipSetPenEndCap pPen, pLineEnd
If pLineEnd = LineCapArrowAnchor And LineArrowLength > 0 Then
    GdipCreateAdjustableArrowCap pLineArrowLength, pLineArrowLength, pLineArrowFill, lccap
    GdipSetPenCustomEndCap pPen, lccap
    GdipDeleteCustomLineCap lccap
End If
End Sub
'---------------------------------------------------------------------------------------
' Dessine un bitmap gdiplus sur l'image
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawBitmap(pBitmap, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                           Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Private Function DrawBitmap(pBitmap, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                            Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawBitmap(pBitmap, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                           Optional pImgSizeMode As EGdipImgSizeMode = GdipSizeModeZoom, Optional pImgPictureAlignment As EGdipImgPictureAlignment = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #Else
Private Function DrawBitmap(pBitmap, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                            Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1) As Boolean
    #End If
#End If
' Conversion si controle de reference
If Not gCtrlRef Is Nothing Then
    PrivMatrixTransformPoint pX1, pY1, CtrlRefToImgMatrix
    PrivMatrixTransformPoint pX2, pY2, CtrlRefToImgMatrix
End If
DrawBitmap = PrivDrawBitmap(pBitmap, pX1, pY1, pX2, pY2, pTranspColor, pImgSizeMode, pImgPictureAlignment, pPercent, pRegion, pRegionColor, pRegionAlpha)
End Function

'---------------------------------------------------------------------------------------
' Dessine un bitmap gdiplus
'---------------------------------------------------------------------------------------
' pBitmap               : Identifiant du bitmap
' Si pImgSizeMode = GdipSizeModeAutoSize :
'       pX1,pY1         : Position de l'image
'       pX2             : Largeur de l'image
'       pY2             : Hauteur de l'image
' Si pImgSizeMode = GdipSizeModeZoom/GdipSizeModeStretch/GdipSizeModeClip :
'       pX1,Y1          : Point Haut-Gauche du rectangle contenant l'image
'       pX2,Y2          : Point Bas-Droite du rectangle contenant l'image
' pTranspcolor          : Couleur de transparence : les points de cette couleur ne
'                           seront pas dessines
' pImgSizeMode          : Type d'affichage de l'image
' pImgPictureAlignment  : Position de l'image
' pPercent              : Pourcentage pour affichage transclucide (0:invisible;255:normal)
' pRegion               : Nom d'une region a creer (= rectangle entourant l'image)
' pRegionColor    : Couleur pour creation de region
' pRegionAlpha    : Transparence pour creation de la region
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawBitmap(pBitmap, ByVal pX1 As Single, ByVal pY1 As Single, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pTranspColor As Long = -1, _
                                Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, _
                                Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1, Optional pImg As clGdiplus) As Boolean
Dim lX1Src As Single, lY1Src As Single, lX2Src As Single, lY2Src As Single
Dim lWidth As Single
Dim lHeight As Single
Dim lDecalageX As Single, lDecalageY As Single
#Else
Private Function PrivDrawBitmap(pBitmap, ByVal pX1 As Long, ByVal pY1 As Long, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pTranspColor As Long = -1, _
                                Optional pImgSizeMode As Long = GdipSizeModeZoom, Optional pImgPictureAlignment As Long = 2, Optional pPercent As Integer = 255, _
                                Optional pRegion As String, Optional pRegionColor As Long = -1, Optional pRegionAlpha As Long = -1, Optional pImg As clGdiplus) As Boolean
Dim lX1Src As Long, lY1Src As Long, lX2Src As Long, lY2Src As Long
Dim lWidth As Long
Dim lHeight As Long
Dim lDecalageX As Long, lDecalageY As Long
#End If
Dim lWidthSng As Single
Dim lHeightSng As Single
Dim lgraphics
Dim lImgBitmap
Dim lMatrix As ColorMatrix
Dim limgattr
Dim lPathMatrix
Dim lWorldMatrix
Dim lScaleX As Single, lScaleY As Single
Dim loImg As clGdiplus
Dim lGdi32 As Boolean
On Error GoTo gestion_erreurs
If pImg Is Nothing Then
    Set loImg = Me
Else
    Set loImg = pImg
    lGdi32 = (pImg.hBrush <> PtrNull And hDC <> PtrNull)
End If
If lGdi32 Then
    If pX2 = cSingleOptional Then pX2 = ImageWidth - 1
    If pY2 = cSingleOptional Then pY2 = ImageHeight - 1
    If WrapX <> 0 Or WrapY <> 0 Then
        Dim lrect As Rect
        Dim lPt As POINTAPI
        lrect.Left = pX1
        lrect.Top = pY1
        lrect.Right = pX2 - pX1 + 1
        lrect.Bottom = pY2 - pY1 + 1
        SetBrushOrgEx hDC, pX1 - WrapX, pY1 - WrapY, lPt
        FillRect hDC, lrect, pImg.hBrush
        SetBrushOrgEx hDC, lPt.X, lPt.Y, lPt
    Else
        'TransparentBlt hDC, pImg.hDC, pX1, pY1, pImg.ImageWidth, pImg.ImageHeight, vbWhite
        BitBlt hDC, pX1, pY1, pX2 - pX1 + 1, pY2 - pY1 + 1, pImg.hDC, 0, 0, SRCCOPY
    End If
    PrivDrawBitmap = True
Else
    lX1Src = pX1: lY1Src = pY1: lX2Src = pX2: lY2Src = pY2
    ' Contexte graphique
    lgraphics = GetGraphics()
    ' Bitmap a dessiner
    lImgBitmap = pBitmap
    ' Dimensions de l'image
    If GdipGetImageDimension(lImgBitmap, lWidthSng, lHeightSng) = 0 Then
    #If UseSingle Then
        lWidth = lWidthSng
        lHeight = lHeightSng
    #Else
        lWidth = CLng(lWidthSng)
        lHeight = CLng(lHeightSng)
    #End If
        'MatrixTransformVector lWidth, lHeight, MatrixInvert(WorldGetMatrix)
        ' Calcul des coordonnees manquantes
        If pX2 = cSingleOptional And pY2 = cSingleOptional Then
            pX2 = lWidth - 1
            pY2 = lHeight - 1
        ElseIf pX2 = cSingleOptional Then    ' Largeur en fonction de la hauteur
            pX2 = pX1 + pY2 * lWidth / lHeight
        ElseIf pY2 = cSingleOptional Then    ' Hauteur en fonction de la largeur
            pY2 = pY1 + pX2 * lHeight / lWidth
        End If
        ' Placement sur un point
        If pImgSizeMode = GdipSizeModeAutoSize Then
            ' Calcule les nouvelles positions en fonction de pImgPictureAlignment
            Select Case pImgPictureAlignment
            Case 0    ' Haut-Gauche
                pX1 = pX1
                pY1 = pY1
            Case 1    ' Haut-Droite
                pX1 = pX1 - lWidth + 1
                pY1 = pY1
            Case 2    ' Centre
                pX1 = pX1 - lWidth \ 2
                pY1 = pY1 - lHeight \ 2
            Case 3    ' Bas-Gauche
                pX1 = pX1
                pY1 = pY1 - lHeight + 1
            Case 4    ' Bas-Droite
                pX1 = pX1 - lWidth + 1
                pY1 = pY1 - lHeight + 1
            End Select
            pX2 = pX1 + lWidth - 1
            pY2 = pY1 + lHeight - 1
        Else
            CalcXY pX1, pY1, pX2, pY2, lWidth, lHeight, lDecalageX, lDecalageY, pImgSizeMode, pImgPictureAlignment
        End If
        ' Transparence de l'image
        If pPercent <> 255 Then
            ' Cree l'objet ImageAttributes
            Call GdipCreateImageAttributes(limgattr)
            ' Cree la matrice
            lMatrix.m(0, 0) = 1: lMatrix.m(1, 0) = 0: lMatrix.m(2, 0) = 0: lMatrix.m(3, 0) = 0: lMatrix.m(4, 0) = 0
            lMatrix.m(0, 1) = 0: lMatrix.m(1, 1) = 1: lMatrix.m(2, 1) = 0: lMatrix.m(3, 1) = 0: lMatrix.m(4, 1) = 0
            lMatrix.m(0, 2) = 0: lMatrix.m(1, 2) = 0: lMatrix.m(2, 2) = 1: lMatrix.m(3, 2) = 0: lMatrix.m(4, 2) = 0
            lMatrix.m(0, 3) = 0: lMatrix.m(1, 3) = 0: lMatrix.m(2, 3) = 0: lMatrix.m(3, 3) = pPercent / 255: lMatrix.m(4, 3) = 0
            lMatrix.m(0, 4) = 0: lMatrix.m(1, 4) = 0: lMatrix.m(2, 4) = 0: lMatrix.m(3, 4) = 0: lMatrix.m(4, 4) = 1
            ' Applique la matrice de couleurs
            Call GdipSetImageAttributesColorMatrix(limgattr, 1, True, lMatrix, ByVal 0, 0)
        End If
        ' Couleur de transparence
        If pTranspColor <> -1 Then
            ' Cree l'objet ImageAttributes
            If limgattr = 0 Then Call GdipCreateImageAttributes(limgattr)
            ' Applique les couleurs de transparence
            GdipSetImageAttributesColorKeys limgattr, 1, True, GdipColor(0, GetColor(pTranspColor)), GdipColor(255, GetColor(pTranspColor))
        End If
        ' Enroulement
        If WrapX <> 0 Or WrapY <> 0 Then
            ' Pour un bon fonctionnement WrapX et WrapY ne doivent pas être inférieurs (valeur absolue) à la taille de l'image dessinée
            If WrapY > 0 And WrapY <= lHeightSng Then
                WrapY = WrapY + lHeightSng + 1
            ElseIf WrapY < 0 And WrapY >= -lHeightSng Then
                WrapY = WrapY - lHeightSng - 1
            End If
            If WrapX > 0 And WrapX <= lWidthSng Then
                WrapX = WrapX + lWidthSng + 1
            ElseIf WrapX < 0 And WrapX >= -lWidthSng Then
                WrapX = WrapX - lWidthSng - 1
            End If
            ' Cree l'objet ImageAttributes
            If limgattr = 0 Then Call GdipCreateImageAttributes(limgattr)
            ' Applique l'"enroulement" de l'image
            GdipSetImageAttributesWrapMode limgattr, 0, 0, True
        End If
        ' Matrice de dessin
        Dim lDrawMatrix
        Dim lDrawMatrixTab(1 To 6) As Single
        GdipCreateMatrix lDrawMatrix
        GdipCreateMatrix lWorldMatrix
        GdipGetWorldTransform lgraphics, lWorldMatrix
        If pImgSizeMode = GdipSizeModeClip Then
            GdipTranslateMatrix lDrawMatrix, -pX1, -pY1, 1
            GdipTranslateMatrix lDrawMatrix, -lX1Src, -lY1Src, 1
            GdipTranslateMatrix lDrawMatrix, lDecalageX, lDecalageY, 1
        Else
            GdipTranslateMatrix lDrawMatrix, -lX1Src, -lY1Src, 1
            GdipTranslateMatrix lDrawMatrix, (lX1Src - pX1), (lY1Src - pY1), 1
            GdipScaleMatrix lDrawMatrix, lWidthSng / (pX2 - pX1 + 1), lHeightSng / (pY2 - pY1 + 1), 1
        End If
        GdipInvertMatrix lWorldMatrix
        GdipMultiplyMatrix lDrawMatrix, lWorldMatrix, 0
        GdipGetMatrixElements lDrawMatrix, lDrawMatrixTab(1)
        GdipDeleteMatrix lDrawMatrix
        GdipDeleteMatrix lWorldMatrix
        loImg.DrawMatrix = lDrawMatrixTab
        ' Dessin de l'image
        If pImgSizeMode = GdipSizeModeClip Then
            ' Dessine l'image dans l'image cible, en appliquant les attributs si necessaire
            PrivDrawBitmap = (GdipDrawImageRectRect(lgraphics, lImgBitmap, pX1, pY1, pX2 - pX1 + 1, pY2 - pY1 + 1, WrapX + lDecalageX, WrapY + lDecalageY, pX2 - pX1 + 1, pY2 - pY1 + 1, 2, limgattr) = 0)
        Else
            ' Dessine l'image redimensionnee dans l'image cible, en appliquant les attributs
            PrivDrawBitmap = (GdipDrawImageRectRect(lgraphics, lImgBitmap, pX1, pY1, pX2 - pX1 + 1, pY2 - pY1 + 1, WrapX, WrapY, CLng(lWidthSng), CLng(lHeightSng), 2, limgattr) = 0)
        End If
        ' Mis a jour de la position
        loImg.DrawX1 = pX1: loImg.DrawY1 = pY1: loImg.DrawX2 = pX2: loImg.DrawY2 = pY2
        If pImgSizeMode = GdipSizeModeClip Or pImgSizeMode = GdipSizeModeAutoSize Then
            loImg.DrawOrigX1 = lDecalageX: loImg.DrawOrigY1 = lDecalageY: loImg.DrawOrigX2 = lDecalageX + (pX2 - pX1 + 1) - 1: loImg.DrawOrigY2 = lDecalageY + (pY2 - pY1 + 1) - 1
        Else
            loImg.DrawOrigX1 = 0: loImg.DrawOrigY1 = 0: loImg.DrawOrigX2 = CLng(lWidthSng) - 1: loImg.DrawOrigY2 = CLng(lHeightSng) - 1
        End If
        ' Ajout de region
        If pRegion <> "" Then
            If pRegionColor = -1 And pRegionAlpha = -1 Then
                PrivDrawRectangle pX1, pY1, pX2, pY2, , -1, 0, , , pRegion, , , True
            Else
                PrivCreateRegionFromColor lImgBitmap, pRegion, pRegionColor, pRegionAlpha, False
                GdipCreateMatrix lWorldMatrix
                GdipGetWorldTransform lgraphics, lWorldMatrix
                GdipCreateMatrix lPathMatrix
                lScaleX = (pX2 - pX1 + 1) / lWidthSng
                lScaleY = (pY2 - pY1 + 1) / lHeightSng
                GdipScaleMatrix lPathMatrix, lScaleX, lScaleY, 0
                GdipTranslateMatrix lPathMatrix, pX1, pY1, 1
                GdipTransformPath gCollRegions(pRegion)(0), lPathMatrix
                GdipTransformPath gCollRegions(pRegion)(0), lWorldMatrix
                GdipDeleteMatrix lWorldMatrix
                GdipDeleteMatrix lPathMatrix
            End If
        End If
    Else
        PrivDrawBitmap = False
    End If
End If
gestion_erreurs:
If Err.Number <> 0 Then PrivDrawBitmap = False   ' renvoie Vrai si pas d'erreur
' Supprime objets temporaires
If limgattr <> 0 Then GdipDisposeImageAttributes limgattr
End Function

'---------------------------------------------------------------------------------------
' Dessine une ligne
'---------------------------------------------------------------------------------------
' pX1,pY1,pX2,pY2   : Coordonnees de la ligne
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pDash             : pointilles si Vrai
' pRegion           : Nom de la region
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawLine(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String) As Boolean
    #Else
Public Function DrawLine(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String) As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawLine(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String) As Boolean
    #Else
Public Function DrawLine(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String) As Boolean
    #End If
#End If
' Conversion si controle de reference
If Not gCtrlRef Is Nothing Then
    PrivMatrixTransformPoint pX1, pY1, CtrlRefToImgMatrix
    PrivMatrixTransformPoint pX2, pY2, CtrlRefToImgMatrix
End If
DrawLine = PrivDrawLine(pX1, pY1, pX2, pY2, pPenColor, pPenWidth, pDash, pAlpha, pRegion)
End Function

'---------------------------------------------------------------------------------------
' Dessine une ligne
'---------------------------------------------------------------------------------------
' pX1,pY1,pX2,pY2   : Coordonnees de la ligne
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawLine(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String) As Boolean
#Else
Private Function PrivDrawLine(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String) As Boolean
#End If
    Dim lgraphics
    Dim lpen, lPath
    On Error GoTo gestion_erreurs
    ' Contexte graphique
    lgraphics = GetGraphics()
    ' Cree le crayon
    GdipCreatePen1 GdipColor(pAlpha, GetColor(pPenColor)), pPenWidth, 2, lpen
    ' Alignement du crayon
    If PenAlignMode = 1 Then GdipSetPenMode lpen, 1
    ' Pointilles
    If pDash > 0 Then GdipSetPenDashStyle lpen, pDash
    If pDash = 5 Then GdipSetPenDashArray lpen, gDashUser(1), UBound(gDashUser) - LBound(gDashUser) + 1
    ' Texture du crayon
    If PenTexture <> "" Then
        On Error Resume Next
        GdipSetPenBrushFill lpen, gTextures.item(PenTexture)
        On Error GoTo gestion_erreurs
    End If
    ' Extremites
    SetLineCap lpen
    ' Dessine la ligne
    GdipDrawLine lgraphics, lpen, pX1, pY1, pX2, pY2
    ' Region
    If pRegion <> "" Then
        GdipCreatePath 0, lPath
        GdipAddPathLine lPath, pX1, pY1, pX2, pY2
        AddToRegion pRegion, lPath, , , pPenWidth, LineJoin
    End If
gestion_erreurs:
    If Err.Number <> 0 Then PrivDrawLine = False    ' renvoie Vrai si pas d'erreur
    ' Supprime objets temporaires
    GdipDeletePen lpen
End Function

'---------------------------------------------------------------------------------------
' Dessine un pixel
'---------------------------------------------------------------------------------------
' pX,pY         : Position du pixel
' pColor        : Couleur du pixel
' pAlpha        : Transparence (0 --> 255)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function DrawPixel(ByVal pX As Single, ByVal pY As Single, pColor As Long, Optional pAlpha As Integer = 255) As Boolean
#Else
Public Function DrawPixel(ByVal pX As Long, ByVal pY As Long, pColor As Long, Optional pAlpha As Integer = 255) As Boolean
#End If
' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        PrivMatrixTransformPoint pX, pY, CtrlRefToImgMatrix
    End If
    DrawPixel = PrivDrawPixel(pX, pY, pColor, pAlpha)
End Function

'---------------------------------------------------------------------------------------
' Dessine un pixel
'---------------------------------------------------------------------------------------
' pX,pY         : Position du pixel
' pColor        : Couleur du pixel
' pAlpha        : Transparence (0 --> 255)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawPixel(ByVal pX As Single, ByVal pY As Single, pColor As Long, Optional pAlpha As Integer = 255) As Boolean
#Else
Private Function PrivDrawPixel(ByVal pX As Long, ByVal pY As Long, pColor As Long, Optional pAlpha As Integer = 255) As Boolean
#End If
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lBitmap = gBitmap
    ' Dessine le point
    WorldTransformPoint pX, pY
    
    PrivDrawPixel = (GdipBitmapSetPixel(lBitmap, pX, pY, GdipColor(pAlpha, GetColor(pColor))) = 0)
    
gestion_erreurs:
    If Err.Number <> 0 Then PrivDrawPixel = False    ' renvoie Vrai si pas d'erreur
End Function

'---------------------------------------------------------------------------------------
' Dessine un controle
'---------------------------------------------------------------------------------------
' pX1,pY1,pX2,pY2   : Coordonnees du rectangle
' pRegion           : Nom de la region a creer
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawControl(pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single, pControlType As EGdipControlType, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawControl(pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single, pControlType As Long, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawControl(pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long, pControlType As EGdipControlType, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawControl(pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long, pControlType As Long, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #End If
#End If
' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        pX1 = CtrlToImgX(pX1, gCtrlRef, True)
        pY1 = CtrlToImgY(pY1, gCtrlRef, True)
        pX2 = CtrlToImgX(pX2, gCtrlRef, True)
        pY2 = CtrlToImgY(pY2, gCtrlRef, True)
    End If
    DrawControl = PrivDrawControl(pX1, pY1, pX2, pY2, pControlType, pPushed, pChecked, pInactive, pFlat, pAdjustRect, pRegion)
End Function

'---------------------------------------------------------------------------------------
' Dessine un controle
'---------------------------------------------------------------------------------------
' pX1,pY1,pX2,pY2   : Coordonnees du rectangle
' pRegion           : Nom de la region a creer
' http://support.microsoft.com/kb/141999
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Private Function PrivDrawControl(pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single, pControlType As EGdipControlType, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #Else
Private Function PrivDrawControl(pX1 As Single, pY1 As Single, pX2 As Single, pY2 As Single, pControlType As Long, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Private Function PrivDrawControl(pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long, pControlType As EGdipControlType, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #Else
Private Function PrivDrawControl(pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long, pControlType As Long, Optional pPushed As Boolean, Optional pChecked As Boolean, Optional pInactive As Boolean, Optional pFlat As Boolean, Optional pAdjustRect As Boolean, Optional pRegion As String = "") As Boolean
    #End If
#End If
    Dim lrect As Rect
    Dim lgraphics
    Dim lHdc
    Dim lMatrix
    Dim lun1 As Long, lun2 As Long
    Dim lEdge As Long
    Dim lMatrixGdi As XForm
    On Error GoTo gestion_erreurs
    ' Graphic de l'image
    lgraphics = GetGraphics()
    ' Lecture des transformations de gdiplus
    GdipCreateMatrix lMatrix
    GdipGetWorldTransform lgraphics, lMatrix
    GdipGetMatrixElements lMatrix, lMatrixGdi.eM11
    GdipDeleteMatrix lMatrix
    ' Coordonnees du controle
    lrect.Left = pX1 - (PenAlignMode = 1)
    lrect.Top = pY1 - (PenAlignMode = 1)
    lrect.Right = pX2 + (PenAlignMode = 1) + 1
    lrect.Bottom = pY2 + (PenAlignMode = 1) + 1
    ' Cree une region rectangulaire
    If pRegion <> "" Then
        CreateRegionRect pRegion, pX1, pY1, pX2 - 1, pY2 - 1
    End If
    ' Device context gdi32
    GdipGetDC lgraphics, lHdc
    ' Applique les transformations a gdi32
    'If pImage = "" Then
        SetGraphicsMode lHdc, GM_ADVANCED
        SetWorldTransform lHdc, lMatrixGdi
    'End If
    ' Type de controle
    Select Case pControlType
        Case 1, 2, 3, 4, 5 ' Caption
            lun1 = DFC_CAPTION
        Case 6, 7, 8, 9 ' Menu
            lun1 = DFC_MENU
        Case 10, 11, 12, 13, 14, 15, 16 ' Scroll
            lun1 = DFC_SCROLL
        Case 17, 18, 19 ' Button
            lun1 = DFC_BUTTON
        Case 20 ' Sunken Frame
            lun1 = EDGE_SUNKEN
            lun2 = BF_RECT
        Case 21 ' Raised Frame
            lun1 = EDGE_RAISED
            lun2 = BF_RECT
    End Select
    ' Sous-type
    Select Case pControlType
        Case 1: lun2 = DFCS_CAPTIONCLOSE
        Case 2: lun2 = DFCS_CAPTIONMIN
        Case 3: lun2 = DFCS_CAPTIONMAX
        Case 4: lun2 = DFCS_CAPTIONRESTORE
        Case 5: lun2 = DFCS_CAPTIONHELP
        Case 6: lun2 = DFCS_MENUARROW
        Case 7: lun2 = DFCS_MENUCHECK
        Case 8: lun2 = DFCS_MENUBULLET
        Case 9: lun2 = DFCS_MENUARROWRIGHT
        Case 10: lun2 = DFCS_SCROLLUP
        Case 11: lun2 = DFCS_SCROLLDOWN
        Case 12: lun2 = DFCS_SCROLLLEFT
        Case 13: lun2 = DFCS_SCROLLRIGHT
        Case 14: lun2 = DFCS_SCROLLCOMBOBOX
        Case 15: lun2 = DFCS_SCROLLSIZEGRIP
        Case 16: lun2 = DFCS_SCROLLSIZEGRIPRIGHT
        Case 17: lun2 = DFCS_BUTTONCHECK
        Case 18: lun2 = DFCS_BUTTONRADIO
        Case 19: lun2 = DFCS_BUTTONPUSH
    End Select
    ' Flags
    lun2 = lun2 Or DFCS_ADJUSTRECT * -pAdjustRect
    lun2 = lun2 Or DFCS_CHECKED * -pChecked
    lun2 = lun2 Or DFCS_FLAT * -pFlat
    lun2 = lun2 Or DFCS_INACTIVE * -pInactive
    lun2 = lun2 Or DFCS_PUSHED * -pPushed
    ' Dessine le controle
    Select Case pControlType
        Case 20, 21 ' Frame
            DrawEdge lHdc, lrect, lun1, lun2
        Case Else
            DrawFrameControl lHdc, lrect, lun1, lun2
    End Select
    ' Taille des bords
    lEdge = GetSystemMetrics(SM_CXBORDER)
    ' Retour coordonnees si pAdjustRect
    If pAdjustRect Then
        Select Case pControlType
            Case 10, 11, 12, 13, 14 ' scrollbutton
                pX1 = pX1 + lEdge
                pX2 = pX2 + lEdge
                pY1 = pY1 - lEdge
                pY2 = pY2 - lEdge
            Case 19, 20, 21 ' pushbutton & frame
                pX1 = lrect.Left
                pX2 = lrect.Right
                pY1 = lrect.Top
                pY2 = lrect.Bottom
        End Select
    End If
gestion_erreurs:
    If Err.Number = 0 Then PrivDrawControl = True   ' renvoie Vrai si pas d'erreur
    If lHdc <> 0 Then GdipReleaseDC lgraphics, lHdc
End Function

'---------------------------------------------------------------------------------------
' Dessine un rectangle
'---------------------------------------------------------------------------------------
' pX1,pY1,pX2,pY2   : Coordonnees du rectangle
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawRectangle(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawRectangle(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawRectangle(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawRectangle(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#End If
    DrawRectangle = PrivDrawRectangle(pX1, pY1, pX2, pY2, pBackColor, pPenColor, pPenWidth, pDash, pAlpha, pRegion)
End Function
'---------------------------------------------------------------------------------------
' Dessine un rectangle avec coins arrondis
'---------------------------------------------------------------------------------------
' pX3, pY3   : Taille de l'ellipse utilise pour les coins
'---------------------------------------------------------------------------------------
#If UseSingle Then
#If VBA6 Then
Public Function DrawRoundRectangle(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, ByVal pX3 As Single, ByVal pY3 As Single, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
#Else
Public Function DrawRoundRectangle(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, ByVal pX3 As Single, ByVal pY3 As Single, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
#End If
#Else
#If VBA6 Then
Public Function DrawRoundRectangle(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, ByVal pX3 As Long, ByVal pY3 As Long, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
#Else
Public Function DrawRoundRectangle(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, ByVal pX3 As Long, ByVal pY3 As Long, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
#End If
#End If
    DrawRoundRectangle = PrivDrawRectangle(pX1, pY1, pX2, pY2, pBackColor, pPenColor, pPenWidth, pDash, pAlpha, pRegion, pX3, pY3)
End Function

'---------------------------------------------------------------------------------------
' Dessine un rectangle
'---------------------------------------------------------------------------------------
' pX1,pY1,pX2,pY2   : Coordonnees du rectangle
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
' pX3, pY3   : Taille de l'ellipse utilise pour les coins
' pOnlyCreateRegion : Vrai si creation de region sans dessin
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawRectangle(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pX3 As Single = cSingleOptional, Optional pY3 As Single = cSingleOptional, Optional pOnlyCreateRegion As Boolean) As Boolean
#Else
Private Function PrivDrawRectangle(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pX3 As Long = cSingleOptional, Optional pY3 As Long = cSingleOptional, Optional pOnlyCreateRegion As Boolean) As Boolean
#End If
    Dim lgraphics
    Dim lpen
    Dim lBrush
    Dim lPath
    Dim lret As Boolean
    On Error GoTo gestion_erreurs
    ' Contexte graphique
    lgraphics = GetGraphics()
    ' Verifie que les coordonnees sont dans le bon ordre
    Call CheckCoordinates(pX1, pY1, pX2, pY2)
    ' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        PrivMatrixTransformPoint pX1, pY1, CtrlRefToImgMatrix
        PrivMatrixTransformPoint pX2, pY2, CtrlRefToImgMatrix
        If pX3 <> cSingleOptional And pY3 <> cSingleOptional Then PrivMatrixTransformPoint pX3, pY3, CtrlRefToImgMatrix, True
    End If
    ' Cree le crayon
    GdipCreatePen1 GdipColor(pAlpha, GetColor(pPenColor)), pPenWidth, 2, lpen
    ' Alignement du crayon
    If PenAlignMode = 1 Then GdipSetPenMode lpen, 1
    ' Texture du crayon
    If PenTexture <> "" Then
        On Error Resume Next
        GdipSetPenBrushFill lpen, gTextures.item(PenTexture)
        On Error GoTo gestion_erreurs
    End If
    ' Pointilles
    If pDash > 0 Then GdipSetPenDashStyle lpen, pDash
    If pDash = 5 Then GdipSetPenDashArray lpen, gDashUser(1), UBound(gDashUser) - LBound(gDashUser) + 1
    ' Cree un chemin
    GdipCreatePath 0, lPath
    If pX3 = cSingleOptional Or pY3 = cSingleOptional Then
        GdipAddPathRectangle lPath, pX1, pY1, pX2 - pX1 - (pPenWidth = 0), pY2 - pY1 - (pPenWidth = 0)
    Else
        GdipAddPathLine lPath, pX1 + pX3 \ 2, pY1, pX2 - pX3 \ 2, pY1
        GdipAddPathArc lPath, pX2 - pX3, pY1, pX3, pY3, 270, 90
        GdipAddPathLine lPath, pX2, pY1 + pY3 \ 2, pX2, pY2 - pY3 \ 2
        GdipAddPathArc lPath, pX2 - pX3, pY2 - pY3, pX3, pY3, 0, 90
        GdipAddPathLine lPath, pX2 - pX3 \ 2, pY2, pX1 + pX3 \ 2, pY2
        GdipAddPathArc lPath, pX1, pY2 - pY3, pX3, pY3, 90, 90
        GdipAddPathLine lPath, pX1, pY2 - pY3 \ 2, pX1, pY1 + pY3 \ 2
        GdipAddPathArc lPath, pX1, pY1, pX3, pY3, 180, 90
    End If
    ' Si dessin demande
    If Not pOnlyCreateRegion Then
        ' Rempli le rectangle
        If FillTexture <> "" Then
            On Error Resume Next
            lBrush = gTextures.item(FillTexture)
            On Error GoTo gestion_erreurs
            If lBrush <> 0 Then GdipFillPath lgraphics, lBrush, lPath
        End If
        If lBrush = 0 Then
            If pBackColor <> -1 Then
                GdipCreateSolidFill GdipColor(pAlpha, GetColor(pBackColor)), lBrush
                GdipFillPath lgraphics, lBrush, lPath
                GdipDeleteBrush lBrush
            End If
        End If
        ' Dessine le rectangle
        If pPenWidth <> 0 And pPenColor <> -1 Then
            GdipSetPenLineJoin lpen, LineJoin
            GdipSetPenMiterLimit lpen, LineJoinMiterLimit
            PrivDrawRectangle = (GdipDrawPath(lgraphics, lpen, lPath) = 0)
        Else
            PrivDrawRectangle = True
        End If
    End If
    ' Creation region
    If pRegion <> "" Then
        lret = AddToRegion(pRegion, lPath, , , pPenWidth)
    Else
        ' Supprime le chemin
        GdipDeletePath lPath
    End If
gestion_erreurs:
    If Err.Number <> 0 Then PrivDrawRectangle = False    ' renvoie Vrai si pas d'erreur
    If pOnlyCreateRegion Then PrivDrawRectangle = lret
    ' Supprime objets temporaires
    If lpen <> 0 Then GdipDeletePen lpen
End Function

'---------------------------------------------------------------------------------------
' Dessine un polygone
'---------------------------------------------------------------------------------------
' pPoints           : Tableau de points formant le polygone : (X1,Y1,X2,Y2,X3,Y3,...)
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawPolygon(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawPolygon(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawPolygon(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawPolygon(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#End If
    DrawPolygon = PrivDrawPolygon(pPoints, pBackColor, pPenColor, pPenWidth, pDash, pAlpha, pRegion)
End Function

'---------------------------------------------------------------------------------------
' Dessine une courbe
'---------------------------------------------------------------------------------------
' pPoints           : Tableau de points formant la courbe : (X1,Y1,X2,Y2,X3,Y3,...)
'           Necessite 1 point de depart + 3 points par courbe
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawCurve(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawCurve(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawCurve(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawCurve(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#End If
    DrawCurve = PrivDrawPolygon(pPoints, pBackColor, pPenColor, pPenWidth, pDash, pAlpha, pRegion, True)
End Function

#If UseSingle Then
    #If VBA6 Then
Public Function DrawCardinal(pPoints As Variant, Optional pTension As Single = 0.5, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawCardinal(pPoints As Variant, Optional pTension As Single = 0.5, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawCardinal(pPoints As Variant, Optional pTension As Single = 0.5, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawCardinal(pPoints As Variant, Optional pTension As Single = 0.5, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#End If
    DrawCardinal = PrivDrawPolygon(pPoints, pBackColor, pPenColor, pPenWidth, pDash, pAlpha, pRegion, False, False, , True, pTension)
End Function
'---------------------------------------------------------------------------------------
' Dessine un polygone ou une courbe
'---------------------------------------------------------------------------------------
' pPoints           : Tableau de points formant le polygone : (X1,Y1,X2,Y2,X3,Y3,...)
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
' pBezier           : Vrai si courbe de bezier
' pOnlyCreateRegion : Vrai si creation de region sans dessin
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawPolygon(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pBezier As Boolean, Optional pOnlyCreateRegion As Boolean, Optional pClosePath As Boolean = False, Optional pCardinal As Boolean, Optional pTension As Single) As Boolean
    Dim lPolyPoints() As POINTF
#Else
Private Function PrivDrawPolygon(pPoints As Variant, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pBezier As Boolean, Optional pOnlyCreateRegion As Boolean, Optional pClosePath As Boolean = False, Optional pCardinal As Boolean, Optional pTension As Single) As Boolean
    Dim lPolyPoints() As POINTAPI
#End If
    Dim lcpt As Long
    Dim lgraphics
    Dim lpen
    Dim lBrush
    Dim lPath
    Dim lret As Boolean
    On Error GoTo gestion_erreurs
    pClosePath = DrawPathClose
    ' Contexte graphique
    lgraphics = GetGraphics()
    ' Redimensionne le tableau de points pour l'API
    ReDim lPolyPoints(0 To ((UBound(pPoints) - LBound(pPoints) + 1) / 2) - 1)
    ' Boucle sur toutes les coordonnees du tableau passe en parametre
    For lcpt = 0 To (UBound(pPoints) - LBound(pPoints) - 1) / 2
        ' Abcisse du point
        lPolyPoints(lcpt).X = pPoints(2 * lcpt + LBound(pPoints))
        ' Ordonnee du point
        lPolyPoints(lcpt).Y = pPoints(2 * lcpt + 1 + LBound(pPoints))
    Next
    ' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        ' Boucle sur toutes les coordonnees du tableau passe en parametre
        For lcpt = LBound(lPolyPoints) To UBound(lPolyPoints)
            ' Abcisse du point
            lPolyPoints(lcpt).X = CtrlToImgX(lPolyPoints(lcpt).X, gCtrlRef, True)
            ' Ordonnee du point
            lPolyPoints(lcpt).Y = CtrlToImgY(lPolyPoints(lcpt).Y, gCtrlRef, True)
        Next
    End If
    ' Cree le crayon
    GdipCreatePen1 GdipColor(pAlpha, GetColor(pPenColor)), pPenWidth, 2, lpen
    ' Alignement du crayon
    If PenAlignMode = 1 Then GdipSetPenMode lpen, 1
    ' Extremites
    SetLineCap lpen
    ' Cree un chemin
    GdipCreatePath 0, lPath
    If pCardinal Then
        GdipAddPathCurve2 lPath, lPolyPoints(0), UBound(lPolyPoints) + 1, pTension
    ElseIf pBezier Then
        GdipAddPathBeziers lPath, lPolyPoints(0), UBound(lPolyPoints) + 1
    Else
        If pClosePath Then
            GdipAddPathPolygon lPath, lPolyPoints(0), UBound(lPolyPoints) + 1
        Else
            GdipAddPathLine2 lPath, lPolyPoints(0), UBound(lPolyPoints) + 1
        End If
    End If
    ' Si dessin demande
    If Not pOnlyCreateRegion Then
        ' Pointilles
        If pDash > 0 Then GdipSetPenDashStyle lpen, pDash
        If pDash = 5 Then GdipSetPenDashArray lpen, gDashUser(1), UBound(gDashUser) - LBound(gDashUser) + 1
        ' Texture du crayon
        If PenTexture <> "" Then
            On Error Resume Next
            GdipSetPenBrushFill lpen, gTextures.item(PenTexture)
            On Error GoTo gestion_erreurs
        End If
        ' Remplit le polygon
        GdipSetPathFillMode lPath, FillMode
        If FillTexture <> "" Then
            On Error Resume Next
            lBrush = gTextures.item(FillTexture)
            On Error GoTo gestion_erreurs
            If lBrush <> 0 Then GdipFillPath lgraphics, lBrush, lPath
        Else
            If pBackColor <> -1 Then
                GdipCreateSolidFill GdipColor(pAlpha, GetColor(pBackColor)), lBrush
                GdipFillPath lgraphics, lBrush, lPath
                GdipDeleteBrush lBrush
            End If
        End If
        ' Dessine le polygon
        If pPenWidth <> 0 And pPenColor <> -1 Then
            GdipSetPenLineJoin lpen, LineJoin
            GdipSetPenMiterLimit lpen, LineJoinMiterLimit
            PrivDrawPolygon = (GdipDrawPath(lgraphics, lpen, lPath) = 0)
        Else
            PrivDrawPolygon = True
        End If
    End If
    ' Creation d'une region
    If pRegion <> "" Then
        lret = AddToRegion(pRegion, lPath, , , pPenWidth)
    Else
        ' Supprime le chemin
        GdipDeletePath lPath
    End If
gestion_erreurs:
    If Err.Number <> 0 Then PrivDrawPolygon = False    ' renvoie Vrai si pas d'erreur
    If pOnlyCreateRegion Then PrivDrawPolygon = lret
    ' Supprime objets temporaires
    If lpen <> 0 Then GdipDeletePen lpen
End Function

'---------------------------------------------------------------------------------------
' Dessine un chemin SVG (A V H M L C Q a v h m l c q z)
'---------------------------------------------------------------------------------------
' pPath             : Chemin au format SVG
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawSVG(pPath As String, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawSVG(pPath As String, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawSVG(pPath As String, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawSVG(pPath As String, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#End If
    DrawSVG = PrivDrawSVG(pPath, pBackColor, pPenColor, pPenWidth, pDash, pAlpha, pRegion)
End Function

'---------------------------------------------------------------------------------------
' Dessine un chemin SVG
'---------------------------------------------------------------------------------------
' pOnlyCreateRegion : Vrai si creation de region sans dessin
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawSVG(pPath As String, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pOnlyCreateRegion As Boolean) As Boolean
    Dim lLastx As Single, lLasty As Single
    Dim lX As Single, lY As Single
    Dim lPolyPoints() As POINTF
#Else
Private Function PrivDrawSVG(pPath As String, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pOnlyCreateRegion As Boolean) As Boolean
    Dim lLastx As Long, lLasty As Long
    Dim lX As Long, lY As Long
    Dim lPolyPoints() As POINTAPI
#End If
    Dim lCoord As String ' Coordonnees
    Dim lCoordArray As Variant ' Coordonnees en tableau
    Dim lCptCoord As Long ' Compteur pour parcourir les coordonnees
    Dim lcptpoint As Long
    Dim lgraphics
    Dim lpen
    Dim lBrush
    Dim lPath
    Dim lRef As Boolean
    Dim lRel As Boolean
    Dim lret As Boolean
    Dim lType As String
    On Error GoTo gestion_erreurs
    ' Contexte graphique
    lgraphics = GetGraphics()
    lRef = Not gCtrlRef Is Nothing
    ' Mise en forme des coordonnees
    lCoord = Replace(pPath, ",", " ")
    ' Cree un tableau a partir de la chaine de caracteres
    lCoordArray = Split(lCoord, " ")
    ' Cree le crayon
    GdipCreatePen1 GdipColor(pAlpha, GetColor(pPenColor)), pPenWidth, 2, lpen
    ' Alignement du crayon
    If PenAlignMode = 1 Then GdipSetPenMode lpen, 1
    ' Cree un chemin
    GdipCreatePath 0, lPath
    ' Initialise le compteur
    lCptCoord = LBound(lCoordArray)
    Do
        If lCptCoord > UBound(lCoordArray) Then Exit Do
        lType = lCoordArray(lCptCoord)
        If Len(lType) > 1 Then
            lCptCoord = lCptCoord - 1
            lCoordArray(lCptCoord + 1) = Mid(lCoordArray(lCptCoord + 1), 2)
            lType = Left(lType, 1)
        End If
        lRel = (StrComp(LCase(lType), lType, vbBinaryCompare) = 0)
        Select Case lType
            Case "M", "m" ' Point de depart
                lX = val(lCoordArray(lCptCoord + 1))
                lY = val(lCoordArray(lCptCoord + 2))
                If lRef Then
                    lX = CtrlToImgX(lX, gCtrlRef, Not lRel)
                    lY = CtrlToImgY(lX, gCtrlRef, Not lRel)
                End If
                If lRel Then
                    lX = lLastx + lX
                    lY = lLasty + lY
                End If
                lLastx = lX
                lLasty = lY
                GdipStartPathFigure lPath
                lCptCoord = lCptCoord + 3
            Case "A", "a" ' Arc
                ' (rx ry x-axis-rotation large-arc-flag sweep-flag x y)+
                ' a 25,25 -30 0,1 50,-25
                lX = val(lCoordArray(lCptCoord + 6))
                lY = val(lCoordArray(lCptCoord + 7))
                If lRef Then
                    lX = CtrlToImgX(lX, gCtrlRef, Not lRel)
                    lX = CtrlToImgY(lX, gCtrlRef, Not lRel)
                End If
                If lRel Then
                    lX = lLastx + lX
                    lY = lLasty + lY
                End If
                ' Calcul pour ellipse entre lLastX, lLastY et lX,lY
                ' Inspire de : http://www.charlespetzold.com/blog/2008/01/Mathematics-of-ArcSegment.html
                Dim lMatrix
                Dim lAngle As Single, lRX As Single, lRY As Single
                Dim lpt1 As POINTF, lpt2 As POINTF
                Dim lMidPoint As POINTF, lCenter As POINTF
                Dim lhalfChord As Single
                Dim lvect As POINTF, lvectRotated As POINTF
                Dim lvectRotatedLength As Single
                Dim lisLargeArc As Boolean, lisCounterclockwise As Boolean
                Dim lcenterDistance As Single
                Dim langle1 As Single, langle2 As Single
                Dim lPath2, lMatrix2
                ' Point de depart et d'arrivee
                lpt1.X = lLastx
                lpt1.Y = lLasty
                lpt2.X = lX
                lpt2.Y = lY
                ' Angle = x-axis-rotation
                lAngle = val(lCoordArray(lCptCoord + 3))
                ' Rayons = rx & ry
                lRX = val(lCoordArray(lCptCoord + 1))
                lRY = val(lCoordArray(lCptCoord + 2))
                ' Choix de l'arc = large-arc-flag
                lisLargeArc = CBool(lCoordArray(lCptCoord + 4))
                ' Sens de l'arc = sweep-flag
                lisCounterclockwise = Not CBool(lCoordArray(lCptCoord + 5))
                ' Transformation des points en fonction de la rotation
                '   et du rapport Y/X des rayons
                GdipCreateMatrix lMatrix
                GdipRotateMatrix lMatrix, -lAngle, 1
                GdipScaleMatrix lMatrix, lRY / lRX, 1, 1
                GdipVectorTransformMatrixPointsS lMatrix, lpt1, 1
                GdipVectorTransformMatrixPointsS lMatrix, lpt2, 1
                ' point median de la corde
                lMidPoint.X = (lpt1.X + lpt2.X) / 2
                lMidPoint.Y = (lpt1.Y + lpt2.Y) / 2
                ' vecteur = corde
                lvect.X = lpt2.X - lpt1.X
                lvect.Y = lpt2.Y - lpt1.Y
                ' demi-longueur de la corde
                lhalfChord = Sqr(lvect.X ^ 2 + lvect.Y ^ 2) / 2
                ' Vecteur perpendiculaire
                If (lisLargeArc = lisCounterclockwise) Then
                    lvectRotated.X = -lvect.Y
                    lvectRotated.Y = lvect.X
                Else
                    lvectRotated.X = lvect.Y
                    lvectRotated.Y = -lvect.X
                End If
                ' taille du vecteur
                lvectRotatedLength = Sqr((lvectRotated.X) ^ 2 + (lvectRotated.Y) ^ 2)
                lcenterDistance = Math.Sqr(Abs(lRY * lRY - lhalfChord * lhalfChord))
                ' Centre de l'ellipse
                lCenter.X = lMidPoint.X + lcenterDistance * lvectRotated.X / lvectRotatedLength
                lCenter.Y = lMidPoint.Y + lcenterDistance * lvectRotated.Y / lvectRotatedLength
                ' Recalcule la position des points
                GdipInvertMatrix lMatrix
                GdipVectorTransformMatrixPointsS lMatrix, lCenter, 1
                GdipVectorTransformMatrixPointsS lMatrix, lpt1, 1
                GdipVectorTransformMatrixPointsS lMatrix, lpt2, 1
                GdipDeleteMatrix lMatrix
                ' Calcul des angles
                langle1 = atan2(lpt1.Y - lCenter.Y, lpt1.X - lCenter.X)
                langle2 = atan2(lpt2.Y - lCenter.Y, lpt2.X - lCenter.X)
                ' Choix de l'arc
                If (lisLargeArc = (Math.Abs(langle2 - langle1) < gPI)) Then
                    If (langle1 < langle2) Then
                        langle1 = langle1 + 2 * gPI
                    Else
                        langle2 = langle2 + 2 * gPI
                    End If
                End If
                ' Ajout de l'arc au chemin
                If lAngle = 0 Then
                    GdipAddPathArc lPath, lCenter.X - lRX, lCenter.Y - lRY, _
                         2 * lRX, 2 * lRY, _
                         langle1 * 180 / gPI, (langle2 - langle1) * 180 / gPI
                Else
                    ' si rotation => cree un chemin temporaire pour l'appliquer
                    GdipCreatePath 0, lPath2
                    GdipAddPathArc lPath2, lCenter.X - lRX, lCenter.Y - lRY, _
                        2 * lRX, 2 * lRY, _
                        -lAngle + langle1 * 180 / gPI, (langle2 - langle1) * 180 / gPI
                    GdipCreateMatrix lMatrix2
                    GdipTranslateMatrix lMatrix2, -lCenter.X, -lCenter.Y, 1
                    GdipRotateMatrix lMatrix2, lAngle, 1
                    GdipTranslateMatrix lMatrix2, lCenter.X, lCenter.Y, 1
                    GdipTransformPath lPath2, lMatrix2
                    GdipAddPathPath lPath, lPath2, True
                    GdipDeletePath lPath2
                End If
                lLastx = lX
                lLasty = lY
                lCptCoord = lCptCoord + 8
            Case "V", "v" ' Ligne verticale
                lY = val(lCoordArray(lCptCoord + 1))
                If lRef Then
                    lY = CtrlToImgY(lY, gCtrlRef, Not lRel)
                End If
                If lRel Then
                    lY = lLasty + lY
                End If
                GdipAddPathLine lPath, lLastx, lLasty, lLastx, lY
                lLasty = lY
                lCptCoord = lCptCoord + 2
            Case "H", "h" ' Ligne horizontale
                lX = val(lCoordArray(lCptCoord + 1))
                If lRef Then
                    lX = CtrlToImgX(lX, gCtrlRef, Not lRel)
                End If
                If lRel Then
                    lX = lLastx + lX
                End If
                GdipAddPathLine lPath, lLastx, lLasty, lX, lLasty
                lLastx = lX
                lCptCoord = lCptCoord + 2
            Case "L", "l" ' Segment
                lX = val(lCoordArray(lCptCoord + 1))
                lY = val(lCoordArray(lCptCoord + 2))
                If lRef Then
                    lX = CtrlToImgX(lX, gCtrlRef, Not lRel)
                    lY = CtrlToImgY(lY, gCtrlRef, Not lRel)
                End If
                If lRel Then
                    lX = lLastx + lX
                    lY = lLasty + lY
                End If
                GdipAddPathLine lPath, lLastx, lLasty, lX, lY
                lLastx = lX
                lLasty = lY
                lCptCoord = lCptCoord + 3
            Case "C", "c", "Q", "q" ' Courbe
                ReDim lPolyPoints(0 To 3)
                lPolyPoints(0).X = lX
                lPolyPoints(0).Y = lY
                For lcptpoint = 1 To 3
                    lX = val(lCoordArray(lCptCoord + 1 + (lcptpoint + (lCoordArray(lCptCoord) = "q" And lcptpoint > 1)) * 2 - 2))
                    lY = val(lCoordArray(lCptCoord + 2 + (lcptpoint + (lCoordArray(lCptCoord) = "q" And lcptpoint > 1)) * 2 - 2))
                    ' Conversion si controle de reference
                    If lRef Then
                        lX = CtrlToImgX(lX, gCtrlRef, Not lRel)
                        lY = CtrlToImgY(lY, gCtrlRef, Not lRel)
                    End If
                    If lRel Then
                        lX = lLastx + lX
                        lY = lLasty + lY
                    End If
                    lLastx = lX
                    lLasty = lY
                    lPolyPoints(lcptpoint).X = lX
                    lPolyPoints(lcptpoint).Y = lY
                Next
                GdipAddPathBeziers lPath, lPolyPoints(0), UBound(lPolyPoints) + 1
                Erase lPolyPoints
                lCptCoord = lCptCoord + 7
            Case "Z", "z" ' Ferme la figure
                GdipClosePathFigure lPath
                lCptCoord = lCptCoord + 1
            Case Else
                lCptCoord = lCptCoord + 1
        End Select
    Loop
    ' Si dessin demande
    If Not pOnlyCreateRegion Then
        ' Pointilles
        If pDash > 0 Then GdipSetPenDashStyle lpen, pDash
        If pDash = 5 Then GdipSetPenDashArray lpen, gDashUser(1), UBound(gDashUser) - LBound(gDashUser) + 1
        ' Texture du crayon
        If PenTexture <> "" Then
            On Error Resume Next
            GdipSetPenBrushFill lpen, gTextures.item(PenTexture)
            On Error GoTo gestion_erreurs
        End If
        ' Remplit le dessin svg
        GdipSetPathFillMode lPath, FillMode
        If FillTexture <> "" Then
            On Error Resume Next
            lBrush = gTextures.item(FillTexture)
            On Error GoTo gestion_erreurs
            If lBrush <> 0 Then GdipFillPath lgraphics, lBrush, lPath
        Else
            If pBackColor <> -1 Then
                GdipCreateSolidFill GdipColor(pAlpha, GetColor(pBackColor)), lBrush
                GdipFillPath lgraphics, lBrush, lPath
                GdipDeleteBrush lBrush
            End If
        End If
        ' Dessine le polygon
        If pPenWidth <> 0 And pPenColor <> -1 Then
            GdipSetPenLineJoin lpen, LineJoin
            GdipSetPenMiterLimit lpen, LineJoinMiterLimit
            ' Extremites
            SetLineCap lpen
            PrivDrawSVG = (GdipDrawPath(lgraphics, lpen, lPath) = 0)
        Else
            PrivDrawSVG = True
        End If
    End If
    ' Creation d'une region
    If pRegion <> "" Then
        lret = AddToRegion(pRegion, lPath, , , pPenWidth)
    Else
        ' Supprime le chemin
        GdipDeletePath lPath
    End If
gestion_erreurs:
    If Err.Number <> 0 Then PrivDrawSVG = False   ' renvoie Vrai si pas d'erreur
    If pOnlyCreateRegion Then PrivDrawSVG = lret
    ' Supprime objets temporaires
    If lpen <> 0 Then GdipDeletePen lpen
End Function
' http://www.developpez.net/forums/d26505/autres-langages/general-visual-basic-6-vbscript/vb-6-anterieur/fonctions-mathematiques-vb/
Private Function atan2(dx As Double, dy As Double, _
  Optional positive_only As Boolean = True) As Double
 
  Dim PI As Double
  Dim PI_2 As Double
 
  PI = Atn(1) * 4    'PI
  PI_2 = Atn(1) * 2  'PI/2
 
  Select Case True
    Case Abs(dx) < 0.0000001
      '-pi(-180°) / 0(0°)
      atan2 = IIf(dy < 0#, -PI, 0#)
 
    Case Abs(dy) < 0.0000001
      '-pi/2(-90°) / +pi/2(+90°)
      atan2 = IIf(dx < 0#, -PI_2, PI_2)
 
    Case Else
      atan2 = Atn(dx / dy)
      atan2 = atan2 + IIf(atan2 < 0#, IIf(dy < 0#, _
        PI, 0#), IIf(dx < 0#, -PI, 0#))
  End Select
 
  If positive_only And (atan2 < 0) Then _
    atan2 = atan2 + 2 * PI
End Function
'---------------------------------------------------------------------------------------
' Dessine une ellipse
'---------------------------------------------------------------------------------------
' Si Ptype = 0 : On passe un rectangle en parametre, l'ellipse remplit ce rectangle
'       pX1,Y1          : Point Haut-Gauche du rectangle
'       pX2,Y2          : Point Bas-Droite du rectangle
' Si Ptype = 1 : on passe le centre et les rayons de l'ellipse en parametre
'       pX1,pY1         : Centre de l'ellipse
'       pX2             : Rayon horizontal
'       pY2             : Rayon Vertical
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pStartAngle       : Angle de demarrage d'un arc en degre
' pSweepAngle       : Angle de balayage d'un arc en degre
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawEllipse(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pType As EGdipTypeEllipse = 0, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As EGdipDashStyle, Optional pStartAngle As Variant, Optional pSweepAngle As Variant, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawEllipse(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pType As Long = 0, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional pStartAngle As Variant, Optional pSweepAngle As Variant, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#Else
    #If VBA6 Then
Public Function DrawEllipse(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pType As EGdipTypeEllipse = 0, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As EGdipDashStyle, Optional pStartAngle As Variant, Optional pSweepAngle As Variant, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawEllipse(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pType As Long = 0, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional pStartAngle As Variant, Optional pSweepAngle As Variant, Optional pAlpha As Integer = 255, Optional pRegion As String = "") As Boolean
    #End If
#End If
    DrawEllipse = PrivDrawEllipse(pX1, pY1, pX2, pY2, pType, pBackColor, pPenColor, pPenWidth, pDash, pStartAngle, pSweepAngle, pAlpha, pRegion)
End Function

'---------------------------------------------------------------------------------------
' Dessine une ellipse
'---------------------------------------------------------------------------------------
' Si Ptype = 0 : On passe un rectangle en parametre, l'ellipse remplit ce rectangle
'       pX1,Y1          : Point Haut-Gauche du rectangle
'       pX2,Y2          : Point Bas-Droite du rectangle
' Si Ptype = 1 : on passe le centre et les rayons de l'ellipse en parametre
'       pX1,pY1         : Centre de l'ellipse
'       pX2             : Rayon horizontal
'       pY2             : Rayon Vertical
' pBackColor        : Couleur de fond (-1 si transparent)
' pPenColor         : Couleur du trait
' pPenWidth         : Largeur du trait
' pDash             : Pointilles
' pStartAngle       : Angle de demarrage d'un arc en degre
' pSweepAngle       : Angle de balayage d'un arc en degre
' pAlpha            : Transparence (0 --> 255)
' pRegion           : Nom de la region a creer
' pOnlyCreateRegion : Vrai si creation de region sans dessin
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawEllipse(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, Optional pType As Long = 0, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Single = 1, Optional pDash As Long, Optional ByVal pStartAngle As Variant, Optional ByVal pSweepAngle As Variant, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pOnlyCreateRegion As Boolean) As Boolean
    Dim lRX As Long, lRY As Long
#Else
Private Function PrivDrawEllipse(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional pType As Long = 0, Optional pBackColor As Long = -1, Optional pPenColor As Long = 0, Optional ByVal pPenWidth As Long = 1, Optional pDash As Long, Optional ByVal pStartAngle As Variant, Optional ByVal pSweepAngle As Variant, Optional pAlpha As Integer = 255, Optional pRegion As String = "", Optional pOnlyCreateRegion As Boolean) As Boolean
    Dim lRX As Single, lRY As Single
#End If
    Dim lgraphics
    Dim lpen
    Dim lBrush
    Dim lPath
    Dim lret As Boolean
    On Error GoTo gestion_erreurs
    ' Contexte graphique
    lgraphics = GetGraphics()
    ' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        PrivMatrixTransformPoint pX1, pY1, CtrlRefToImgMatrix
        PrivMatrixTransformPoint pX2, pY2, CtrlRefToImgMatrix, pType = 1
    End If
    ' Cree le crayon
    GdipCreatePen1 GdipColor(pAlpha, GetColor(pPenColor)), pPenWidth, 2, lpen
    ' Alignement du crayon
    If PenAlignMode = 1 Then GdipSetPenMode lpen, 1
    ' Coordonnees du rectangle si centre+rayon
    If pType = 1 Then
        lRX = pX2
        lRY = pY2
        pX2 = pX1 + lRX - 1
        pY2 = pY1 + lRY - 1
        pX1 = pX1 - lRX + 1
        pY1 = pY1 - lRY + 1
    End If
    ' Cree un chemin
    GdipCreatePath 0, lPath
    If Not IsMissing(pStartAngle) And Not IsMissing(pSweepAngle) Then
        ' Extremites
        SetLineCap lpen
        PrivDrawEllipse = (GdipAddPathArc(lPath, pX1, pY1, pX2 - pX1, pY2 - pY1, CSng(pStartAngle), CSng(pSweepAngle)) = 0)
    Else
        PrivDrawEllipse = (GdipAddPathEllipse(lPath, pX1, pY1, pX2 - pX1, pY2 - pY1) = 0)
    End If
    ' Si dessin demande
    If Not pOnlyCreateRegion Then
        ' Pointilles
        If pDash > 0 Then GdipSetPenDashStyle lpen, pDash
        If pDash = 5 Then GdipSetPenDashArray lpen, gDashUser(1), UBound(gDashUser) - LBound(gDashUser) + 1
        ' Texture du crayon
        If PenTexture <> "" Then
            On Error Resume Next
            GdipSetPenBrushFill lpen, gTextures.item(PenTexture)
            On Error GoTo gestion_erreurs
        End If
        ' Rempli l'ellipse
        If FillTexture <> "" Then
            On Error Resume Next
            lBrush = gTextures.item(FillTexture)
            On Error GoTo gestion_erreurs
             If lBrush <> 0 Then GdipFillPath lgraphics, lBrush, lPath
        Else
            If pBackColor <> -1 Then
                GdipCreateSolidFill GdipColor(pAlpha, GetColor(pBackColor)), lBrush
                GdipFillPath lgraphics, lBrush, lPath
                GdipDeleteBrush lBrush
            End If
        End If
        ' Dessine l'ellipse
        If pPenWidth <> 0 And pPenColor <> -1 Then
            PrivDrawEllipse = (GdipDrawPath(lgraphics, lpen, lPath) = 0)
        Else
            PrivDrawEllipse = True
        End If
    End If
    ' Creation d'une region
    If pRegion <> "" Then
        lret = AddToRegion(pRegion, lPath, , , pPenWidth)
    Else
        ' Supprime le chemin
        GdipDeletePath lPath
    End If
gestion_erreurs:
    If Err.Number <> 0 Then PrivDrawEllipse = False    ' renvoie Vrai si pas d'erreur
    If pOnlyCreateRegion Then PrivDrawEllipse = lret
    ' Supprime objets temporaires
    GdipDeletePen lpen
End Function

'---------------------------------------------------------------------------------------
' Dessine du texte
'---------------------------------------------------------------------------------------
' pText             : Texte a ecrire
' pSize             : Taille du texte en pixel
' pFontName         : Nom de la police de caractere
' pX1,pY1,pX2,pY2   : Coordonnees pour positionner le texte
' pAlignHoriz       : Alignement horizontal (0 = gauche; 1 = centre; 2 = droite)
' pAlignVert        : Alignement vertical (0 = haut; 1 = centre; 2 = bas)
' pPenColor         : Couleur du texte
' pPenAlpha         : Transparence du texte (0 a 255)
' pBackColor        : Couleur de fond (-1 si transparent)
' pBackAlpha        : Transparence du fond (0 a 255)
' pItalic           : Texte en italique (vrai/faux)
' pBold             : Texte en gras (vrai/faux)
' pUnderline        : Texte souligne (vrai/faux)
' pItalic           : Texte barre (vrai/faux)
' pOnlyGetSize      : N'ecrit pas le texte, renvoit la taille du texte dans pX1,pY1,pX2,PY2
' pRegion           : Ajoute une region correspondant au rectangle englobant le texte ecrit
'---------------------------------------------------------------------------------------
#If UseSingle Then
    #If VBA6 Then
Public Function DrawText(pText As String, pSize As Single, Optional pFontName As String, _
                         Optional pX1 As Single, Optional pY1 As Single, Optional pX2 As Single = cSingleOptional, Optional pY2 As Single = cSingleOptional, _
                         Optional pAlignHoriz As EGdipHorzAlign = 1, Optional pAlignVert As EGdipVertAlign = 1, _
                         Optional pPenColor As Long = 0, Optional pPenAlpha As Integer = 255, _
                         Optional pBackColor As Long = -1, Optional pBackAlpha As Integer = 255, _
                         Optional pItalic As Boolean, Optional pBold As Boolean, _
                         Optional pUnderline As Boolean, Optional pStrikeOut As Boolean, _
                         Optional pOnlyGetSize As Boolean = False, _
                         Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawText(pText As String, pSize As Single, Optional pFontName As String, _
                         Optional pX1 As Single, Optional pY1 As Single, Optional pX2 As Single = cSingleOptional, Optional pY2 As Single = cSingleOptional, _
                         Optional pAlignHoriz As Long = 1, Optional pAlignVert As Long = 1, _
                         Optional pPenColor As Long = 0, Optional pPenAlpha As Integer = 255, _
                         Optional pBackColor As Long = -1, Optional pBackAlpha As Integer = 255, _
                         Optional pItalic As Boolean, Optional pBold As Boolean, _
                         Optional pUnderline As Boolean, Optional pStrikeOut As Boolean, _
                         Optional pOnlyGetSize As Boolean = False, _
                         Optional pRegion As String = "") As Boolean
    #End If
Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
#Else
    #If VBA6 Then
Public Function DrawText(pText As String, pSize As Long, Optional pFontName As String, _
                         Optional pX1 As Long, Optional pY1 As Long, Optional pX2 As Long = cSingleOptional, Optional pY2 As Long = cSingleOptional, _
                         Optional pAlignHoriz As EGdipHorzAlign = 1, Optional pAlignVert As EGdipVertAlign = 1, _
                         Optional pPenColor As Long = 0, Optional pPenAlpha As Integer = 255, _
                         Optional pBackColor As Long = -1, Optional pBackAlpha As Integer = 255, _
                         Optional pItalic As Boolean, Optional pBold As Boolean, _
                         Optional pUnderline As Boolean, Optional pStrikeOut As Boolean, _
                         Optional pOnlyGetSize As Boolean = False, _
                         Optional pRegion As String = "") As Boolean
    #Else
Public Function DrawText(pText As String, pSize As Long, Optional pFontName As String, _
                         Optional pX1 As Long, Optional pY1 As Long, Optional pX2 As Long = cSingleOptional, Optional pY2 As Long = cSingleOptional, _
                         Optional pAlignHoriz As Long = 1, Optional pAlignVert As Long = 1, _
                         Optional pPenColor As Long = 0, Optional pPenAlpha As Integer = 255, _
                         Optional pBackColor As Long = -1, Optional pBackAlpha As Integer = 255, _
                         Optional pItalic As Boolean, Optional pBold As Boolean, _
                         Optional pUnderline As Boolean, Optional pStrikeOut As Boolean, _
                         Optional pOnlyGetSize As Boolean = False, _
                         Optional pRegion As String = "") As Boolean
    #End If
Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
#End If
    lx1 = pX1: lx2 = pX2: ly1 = pY1: ly2 = pY2
' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        PrivMatrixTransformPoint lx1, ly1, CtrlRefToImgMatrix
        PrivMatrixTransformPoint lx2, ly2, CtrlRefToImgMatrix
    End If
    DrawText = PrivDrawText(pText, pSize, pFontName, lx1, ly1, lx2, ly2, _
                            pAlignHoriz, pAlignVert, pPenColor, pPenAlpha, pBackColor, pBackAlpha, _
                            pItalic, pBold, pUnderline, pStrikeOut, pOnlyGetSize, pRegion)
    If DrawText And pOnlyGetSize Then
        pX1 = lx1: pX2 = lx2: pY1 = ly1: pY2 = ly2
    End If
End Function

'---------------------------------------------------------------------------------------
' Dessine du texte
'---------------------------------------------------------------------------------------
' pText             : Texte a ecrire
' pFontName         : Nom de la police de caractere
' pX1,pY1,pX2,pY2   : Coordonnees pour positionner le texte
' pAlignHoriz       : Alignement horizontal (0 = gauche; 1 = centre; 2 = droite)
' pAlignVert        : Alignement vertical (0 = haut; 1 = centre; 2 = bas)
' pPenColor         : Couleur du texte
' pPenAlpha         : Transparence du texte (0 a 255)
' pBackColor        : Couleur de fond (-1 si transparent)
' pBackAlpha        : Transparence du fond (0 a 255)
' pItalic           : Texte en italique (vrai/faux)
' pBold             : Texte en gras (vrai/faux)
' pUnderline        : Texte souligne (vrai/faux)
' pStrikeout        : Texte barre (vrai/faux)
' pOnlyGetSize      : N'ecrit pas le texte, renvoit la taille du texte dans pX1,pY1,pX2,PY2
' pRegion           : Ajoute une region correspondant au rectangle englobant le texte ecrit
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivDrawText(pText As String, pSize As Single, Optional ByVal pFontName As String, _
                              Optional pX1 As Single, Optional pY1 As Single, Optional pX2 As Single = cSingleOptional, Optional pY2 As Single = cSingleOptional, _
                              Optional ByVal pAlignHoriz As Long = 1, Optional ByVal pAlignVert As Long = 1, _
                              Optional pPenColor As Long = 0, Optional pPenAlpha As Integer = 255, _
                              Optional pBackColor As Long = -1, Optional pBackAlpha As Integer = 255, _
                              Optional pItalic As Boolean, Optional pBold As Boolean, _
                              Optional pUnderline As Boolean, Optional pStrikeOut As Boolean, _
                              Optional pOnlyGetSize As Boolean = False, _
                              Optional pRegion As String = "") As Boolean
    Dim lSize As Single
    Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
#Else
Private Function PrivDrawText(pText As String, pSize As Long, Optional ByVal pFontName As String, _
                              Optional pX1 As Long, Optional pY1 As Long, Optional pX2 As Long = cSingleOptional, Optional pY2 As Long = cSingleOptional, _
                              Optional ByVal pAlignHoriz As Long = 1, Optional ByVal pAlignVert As Long = 1, _
                              Optional pPenColor As Long = 0, Optional pPenAlpha As Integer = 255, _
                              Optional pBackColor As Long = -1, Optional pBackAlpha As Integer = 255, _
                              Optional pItalic As Boolean, Optional pBold As Boolean, _
                              Optional pUnderline As Boolean, Optional pStrikeOut As Boolean, _
                              Optional pOnlyGetSize As Boolean = False, _
                              Optional pRegion As String = "") As Boolean
    Dim lSize As Long
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
#End If
    Dim lgraphics
    Dim lFontFamily, lFont
    Dim lTextBrush, lBackBrush
    Dim lTextRect As RECTF
    Dim lCalcRect As RECTF
    Dim lStrFormat
    Dim lBitmap
    Dim lWidth As Single, lHeight As Single
    Dim lCenter As Boolean
    Dim lStyle As Long
    Dim lTextRect2 As RECTF
    Dim lCalcRect2 As RECTF
    Dim lShrink As Boolean
    Dim lEnlarge As Boolean
    Dim lPath, lpen
    On Error GoTo gestion_erreurs
    ' Contexte graphique
    lgraphics = GetGraphics()
    lBitmap = gBitmap
    ' Rectangle de positionnement du texte
    ' Dimensions de l'image
    GdipGetImageDimension lBitmap, lWidth, lHeight
    lx1 = pX1: lx2 = pX2: ly1 = pY1: ly2 = pY2
    If lx1 = 0 And lx2 = 0 And ly1 = 0 And ly2 = 0 Then
        'pour compatibilite avec version precedente
        lx2 = -1
        ly2 = -1
    Else
        If pX2 = cSingleOptional And pY2 = cSingleOptional Then
            lCenter = True
        Else
            CheckCoordinates lx1, ly1, lx2, ly2
        End If
        If lx2 = cSingleOptional Then lx2 = lWidth
        If ly2 = cSingleOptional Then ly2 = lHeight
    End If
    lTextRect.Left = lx1
    lTextRect.Top = ly1
    lTextRect.Right = lx2 - lx1 ' + 1
    lTextRect.Bottom = ly2 - ly1 '+ 1
    ' Position sur un point
    If lCenter Then
        Select Case pAlignHoriz
        Case 0
            lTextRect.Left = lx1 - lWidth
            lTextRect.Right = lWidth
            pAlignHoriz = 2
        Case 1
            lTextRect.Left = lx1 - lWidth \ 2
            lTextRect.Right = lWidth
        Case 2
            lTextRect.Left = lx1
            lTextRect.Right = lWidth
            pAlignHoriz = 0
        End Select
        Select Case pAlignVert
        Case 0
            lTextRect.Top = ly1 - lHeight
            lTextRect.Bottom = lHeight
            pAlignVert = 2
        Case 1
            lTextRect.Top = ly1 - lHeight \ 2
            lTextRect.Bottom = lHeight
        Case 2
            lTextRect.Top = ly1
            lTextRect.Bottom = lHeight
            pAlignVert = 0
        End Select
    End If
    ' Cree une famille de police
    If pFontName = "" Then
        Call GdipGetGenericFontFamilySansSerif(lFontFamily)
    Else
        If GdipCreateFontFamilyFromName(StrPtr(pFontName), 0, lFontFamily) <> 0 Then
            Call GdipCreateFontFamilyFromName(StrPtr(pFontName), gPrivateFontCollection, lFontFamily)
        End If
    End If
    If lFontFamily = 0 Then
        PrivDrawText = False
        Exit Function
    End If
    ' Cree une police de caracteres
    lStyle = -pBold - pItalic * 2 - pUnderline * 4 - pStrikeOut * 8
    lSize = pSize
    Call GdipCreateFont(lFontFamily, lSize, lStyle, 2, lFont)
    If lFont = 0 Then
        PrivDrawText = False
        GdipDeleteFontFamily lFontFamily
        Exit Function
    End If
    ' Brosse pour texte
    If FillTexture <> "" Then
        On Error Resume Next
        lTextBrush = gTextures.item(FillTexture)
        On Error GoTo gestion_erreurs
    Else
        GdipCreateSolidFill GdipColor(pPenAlpha, GetColor(pPenColor)), lTextBrush
    End If
    ' Format du texte
    GdipCreateStringFormat 0, 0, lStrFormat
    GdipSetStringFormatAlign lStrFormat, pAlignHoriz
    GdipSetStringFormatLineAlign lStrFormat, pAlignVert
    If Not lCenter And (MaxTextSize > 0 Or MinTextSize > 0) Then
        Do
            If lSize <= 0 Then Exit Do
            ' Reduit ou agrandi le texte pour rentrer dans la boite
            lTextRect2.Right = 0
            lTextRect2.Bottom = 0
            lTextRect2.Left = lTextRect.Left
            lTextRect2.Top = lTextRect.Top
            lCalcRect2.Left = 0
            lCalcRect2.Right = 0
            lCalcRect2.Bottom = 0
            lCalcRect2.Top = 0
            GdipMeasureString lgraphics, StrPtr(pText), Len(pText), lFont, lTextRect2, lStrFormat, lCalcRect2, ByVal 0, ByVal 0
            If lCalcRect2.Bottom > lTextRect.Bottom Or lCalcRect2.Right > lTextRect.Right Then
                If lSize = MinTextSize Then Exit Do
                lSize = lSize - 1
                lShrink = True
                GdipDeleteFont lFont
                lFont = 0
                Call GdipCreateFont(lFontFamily, lSize, lStyle, 2, lFont)
                If lEnlarge Then Exit Do
            ElseIf lCalcRect2.Bottom < lTextRect.Bottom Or lCalcRect2.Right < lTextRect.Right Then
                If lShrink Then Exit Do
                If lSize = MaxTextSize Then Exit Do
                lSize = lSize + 1
                lEnlarge = True
                GdipDeleteFont lFont
                lFont = 0
                Call GdipCreateFont(lFontFamily, lSize, lStyle, 2, lFont)
            Else
                Exit Do
            End If
        Loop
        If lFont = 0 Then
            PrivDrawText = False
            GdipDeleteFontFamily lFontFamily
            Exit Function
        End If
    End If
    ' Dessine sauf si uniquement taille demandee
    If Not pOnlyGetSize Then
        ' Position et Taille du texte
        GdipMeasureString lgraphics, StrPtr(pText), Len(pText), lFont, lTextRect, lStrFormat, lCalcRect, ByVal 0, ByVal 0
        LastTextBottom = lCalcRect.Top + lCalcRect.Bottom
        LastTextRight = lCalcRect.Left + lCalcRect.Right
        LastTextSize = lSize
        ' Couleur de fond
        If pBackColor <> -1 Then
            GdipCreateSolidFill GdipColor(pBackAlpha, GetColor(pBackColor)), lBackBrush
            GdipFillRectangle lgraphics, lBackBrush, lCalcRect.Left, lCalcRect.Top, lCalcRect.Right + 1, lCalcRect.Bottom + 1
        End If
        ' Si remplissage de texture, dessine le contour
        If FillTexture <> "" Then
            GdipCreatePath 0, lPath
            lTextRect.Right = lTextRect.Right - 1
            lTextRect.Bottom = lTextRect.Bottom - 1
            GdipAddPathString lPath, StrPtr(pText), Len(pText), lFontFamily, 0, lSize, lTextRect, lStrFormat
            If pPenColor <> -1 Then GdipCreatePen1 GdipColor(pPenAlpha, GetColor(pPenColor)), 1, 2, lpen
            If lTextBrush <> 0 Then GdipFillPath lgraphics, lTextBrush, lPath
            If lpen <> 0 Then GdipDrawPath lgraphics, lpen, lPath
            If pRegion <> "" And TextCreateRegionOutline Then
                Call AddToRegion(pRegion, lPath, , , 1)
            Else
                GdipDeletePath lPath
            End If
            If lpen <> 0 Then GdipDeletePen lpen
        Else
            ' Ecrit le texte
            Call GdipDrawString(lgraphics, StrPtr(pText), Len(pText), lFont, lTextRect, lStrFormat, lTextBrush)
            If pRegion <> "" And TextCreateRegionOutline Then
                GdipCreatePath 0, lPath
                GdipAddPathString lPath, StrPtr(pText), Len(pText), lFontFamily, 0, lSize, lTextRect, lStrFormat
                Call AddToRegion(pRegion, lPath, , , 1)
            End If
        End If
    Else
        ' Position et Taille du texte
        GdipMeasureString lgraphics, StrPtr(pText), Len(pText), lFont, lTextRect, lStrFormat, lCalcRect, ByVal 0, ByVal 0
        ' Renvoit l'emplacement du texte
        pX1 = lCalcRect.Left
        pY1 = lCalcRect.Top
        pX2 = pX1 + lCalcRect.Right
        pY2 = pY1 + lCalcRect.Bottom
    End If
    If pRegion <> "" And Not TextCreateRegionOutline Then
        CreateRegionRect pRegion, lCalcRect.Left, lCalcRect.Top, lCalcRect.Left + lCalcRect.Right, lCalcRect.Top + lCalcRect.Bottom
    End If
gestion_erreurs:
    If Err.Number = 0 Then PrivDrawText = True    ' renvoie Vrai si pas d'erreur
    ' Supprime objets temporaires
    GdipDeleteFont lFont
    GdipDeleteFontFamily lFontFamily
    If FillTexture = "" Then GdipDeleteBrush lTextBrush
    GdipDeleteBrush lBackBrush
    GdipDeleteStringFormat lStrFormat
End Function
'---------------------------------------------------------------------------------------
' Renvoi un tableau de long contenant les pixels de l'image
'---------------------------------------------------------------------------------------
' Ex :    Dim lPixels() As Byte
'           climg.GetPixels lPixels
'           b = lPixels(1, cptx, cpty)
'           g = lPixels(2, cptx, cpty)
'           r = lPixels(3, cptx, cpty)
'           a = lPixels(4, cptx, cpty)
'---------------------------------------------------------------------------------------
Public Function SavePixels() As Byte()
    SavePixels = PrivGetPixels()
End Function

'---------------------------------------------------------------------------------------
' Renvoi la couleur d'un pixel
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivGetPixel(pX As Single, pY As Single) As Long
#Else
Private Function PrivGetPixel(pX As Long, pY As Long) As Long
#End If

    Dim bmpData As BitmapData
    Dim lrect As Rect
    Dim lPixels() As Byte
    Dim lBitmap
    On Error GoTo gestion_erreurs
    
    lBitmap = gBitmap
    lrect.Left = pX
    lrect.Top = pY
    lrect.Right = 1
    lrect.Bottom = 1
    If GdipBitmapLockBits(lBitmap, lrect, &H1, PixelFormat32bppARGB, bmpData) = 0 Then
        ReDim lPixels(1 To 4, 0 To bmpData.Width - 1, 0 To bmpData.Height - 1)
        RtlMoveMemory lPixels(1, 0, 0), ByVal bmpData.scan0, 4 ' bmpData.stride * bmpData.Height
        Call GdipBitmapUnlockBits(lBitmap, bmpData)
        PrivGetPixel = RGB(lPixels(3, 0, 0), lPixels(2, 0, 0), lPixels(1, 0, 0))
    End If
    If lBitmap <> gBitmap Then GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    Call GdipBitmapUnlockBits(lBitmap, bmpData)
End Function

'---------------------------------------------------------------------------------------
' Renvoi la couleur d'un pixel
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function GetPixel(ByVal pX As Single, ByVal pY As Single) As Long
#Else
Public Function GetPixel(ByVal pX As Long, ByVal pY As Long) As Long
#End If
' Conversion si controle de reference
    If Not gCtrlRef Is Nothing Then
        PrivMatrixTransformPoint pX, pY, CtrlRefToImgMatrix
    End If
    GetPixel = PrivGetPixel(pX, pY)
End Function


'---------------------------------------------------------------------------------------
' Renvoi un tableau de long contenant les pixels
'---------------------------------------------------------------------------------------
' Ex :    Dim lPixels() As Byte
'           climg.GetPixels lPixels
'           b = lPixels(1, cptx, cpty) ' bleu
'           g = lPixels(2, cptx, cpty) ' vert
'           r = lPixels(3, cptx, cpty) ' rouge
'           a = lPixels(4, cptx, cpty) ' transparence
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Private Function PrivGetPixels() As Byte()
    Dim bmpData As BitmapData
    Dim lrect As Rect
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lPixels() As Byte
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lBitmap = gBitmap 'CloneBppGraphics(gBitmap, , , True)
    GdipGetImageDimension lBitmap, lWidth, lHeight
    lrect.Right = lWidth
    lrect.Bottom = lHeight
    If GdipBitmapLockBits(lBitmap, lrect, &H1, PixelFormat32bppARGB, bmpData) = 0 Then
        ReDim lPixels(1 To 4, 0 To bmpData.Width - 1, 0 To bmpData.Height - 1)
        RtlMoveMemory lPixels(1, 0, 0), ByVal bmpData.scan0, bmpData.stride * bmpData.Height
        Call GdipBitmapUnlockBits(lBitmap, bmpData)
        PrivGetPixels = lPixels
    End If
    If lBitmap <> gBitmap Then GdipDisposeImage lBitmap
    Exit Function
gestion_erreurs:
    Call GdipBitmapUnlockBits(lBitmap, bmpData)
    Erase PrivGetPixels
End Function

'---------------------------------------------------------------------------------------
' Injecte un tableau de long contenant les pixels dans l'image
'---------------------------------------------------------------------------------------
' pPixels : Tableau de points ARGB
'---------------------------------------------------------------------------------------
' Ex :  Dim lPixels() As Byte
'       ReDim lPixels(1 To 4, 1 to lWidth,1 to lHeight )
'       ......
'       climg.SetPixels lPixels
'---------------------------------------------------------------------------------------
Public Function LoadPixels(pPixels() As Byte) As Boolean
    LoadPixels = PrivSetPixels(pPixels)
End Function

'---------------------------------------------------------------------------------------
' Injecte un tableau de long contenant les pixels dans l'image
'---------------------------------------------------------------------------------------
' pPixels : Tableau de points ARGB
'---------------------------------------------------------------------------------------
' Ex :  Dim lPixels() As Byte
'       ReDim lPixels(1 To 4, 1 to lWidth,1 to lHeight )
'       ......
'       climg.SetPixels lPixels
'---------------------------------------------------------------------------------------
Private Function PrivSetPixels(pPixels() As Byte) As Boolean
    Dim lbmpdata As BitmapData
    Dim lrect As Rect
    Dim lBitmap
    Dim lNewBitmap As Boolean
    On Error GoTo gestion_erreurs
    lbmpdata.Width = UBound(pPixels, 2) - LBound(pPixels, 2) + 1
    lbmpdata.Height = UBound(pPixels, 3) - LBound(pPixels, 3) + 1
    If gBitmap <> 0 Then
        ConvertBppGraphics gBitmap, , , True
        lBitmap = gBitmap
    Else
        CreateBitmap lbmpdata.Width, lbmpdata.Height
        lBitmap = gBitmap
        lNewBitmap = True
    End If
    lbmpdata.stride = 4 * lbmpdata.Width
    lbmpdata.PixelFormat = PixelFormat32bppARGB
    lbmpdata.scan0 = VarPtr(pPixels(LBound(pPixels, 1), LBound(pPixels, 2), LBound(pPixels, 3)))
    lrect.Left = LBound(pPixels, 2) * -(Not lNewBitmap)
    lrect.Top = LBound(pPixels, 3) * -(Not lNewBitmap)
    lrect.Right = lbmpdata.Width
    lrect.Bottom = lbmpdata.Height
    If GdipBitmapLockBits(lBitmap, lrect, &H2 Or &H4, PixelFormat32bppARGB, lbmpdata) = 0 Then
        Call GdipBitmapUnlockBits(lBitmap, lbmpdata)
        PrivSetPixels = True
    End If
    Exit Function
gestion_erreurs:
    If lBitmap <> 0 Then Call GdipBitmapUnlockBits(lBitmap, lbmpdata)
End Function

'---------------------------------------------------------------------------------------
' Rotation de l'image
'---------------------------------------------------------------------------------------
' pAngle    : angle de rotation en degres
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
Public Function Rotate(ByVal pAngle As Single, Optional pKeepExifData As Boolean) As Boolean
    Rotate = PrivRotate(pAngle, pKeepExifData)
End Function

'---------------------------------------------------------------------------------------
' Rotation de l'image
'---------------------------------------------------------------------------------------
' pAngle    : angle de rotation en degres
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
Private Function PrivRotate(ByVal pAngle As Single, Optional pKeepExifData As Boolean) As Boolean
    Dim lgraphics
    Dim lNewBitmap, lNewGraphics
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lBrush
    Dim lPoint1x As Long, lPoint1y As Long, lPoint2x As Long, lPoint2y As Long, lPoint3x As Long, lPoint3y As Long
    Dim lMinX As Long, lMaxX As Long, lMinY As Long, lMaxY As Long
    Dim lAngle As Double, CA As Double, SA As Double
    Dim lsX As Long, lsY As Long, lDx As Long, lDy As Long
    Dim lOldBitmap
    On Error GoTo gestion_erreurs
    ' La rotation s'effectue autout du point 0 (en haut a gauche)
    ' Le point 1 est en haut a droite, le 2 en bas a droite, le 3 en bas a gauche
    
    ' Contexte graphique
    lgraphics = GetGraphics()
    lOldBitmap = gBitmap

    GdipGetImageDimension lOldBitmap, lWidth, lHeight

    lsX = lWidth
    lsY = lHeight
    ' Conversion angle en radian
    lAngle = (pAngle * 2 * 4 * Atn(1)) / 360
    ' Calcul les cosinus et sinus de l'angle
    CA = Cos(lAngle)
    SA = Sin(lAngle)
    ' Calcul les nouvelles positions des points (le point 0 ne bouge pas)
    lPoint1x = -lsY * SA
    lPoint1y = lsY * CA
    lPoint2x = lsX * CA - lsY * SA
    lPoint2y = lsY * CA + lsX * SA
    lPoint3x = lsX * CA
    lPoint3y = lsX * SA
    ' Recherche les mini et maxi
    lMinX = fMin(0, fMin(lPoint1x, fMin(lPoint2x, lPoint3x)))
    lMinY = fMin(0, fMin(lPoint1y, fMin(lPoint2y, lPoint3y)))
    lMaxX = fMax(0, fMax(lPoint1x, fMax(lPoint2x, lPoint3x)))
    lMaxY = fMax(0, fMax(lPoint1y, fMax(lPoint2y, lPoint3y)))
    ' Nouvelle taille
    lDx = Abs(lMaxX) - lMinX
    lDy = Abs(lMaxY) - lMinY
    ' On cree un nouveau bitmap pour contenir l'image transformee
    GdipCreateBitmapFromGraphics lDx, lDy, lgraphics, lNewBitmap
    GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
    UpdateGraphics lNewGraphics
    GdipCreateSolidFill GdipColor(0, vbBlack), lBrush
    GdipFillRectangleI lNewGraphics, lBrush, 0, 0, CLng(lWidth), CLng(lHeight)
    GdipDeleteBrush lBrush
    ' Transformations a appliquer (rotation + translation pour centrage)
    GdipRotateWorldTransform lNewGraphics, pAngle, 0
    GdipTranslateWorldTransform lNewGraphics, -lMinX, -lMinY, 1
    ' Dessine l'image sur le nouveau bitmap en appliquant les transformations
    PrivRotate = (GdipDrawImageRectRectI(lNewGraphics, lOldBitmap, 0, 0, lWidth, lHeight, 0, 0, lWidth, lHeight, 2, 0) = 0)
    ' Desactive les transformations pour la suite
    GdipResetWorldTransform lNewGraphics
    ' Reimporte les donnees Exif
    If pKeepExifData Then ImportExifDataFromBitmap lOldBitmap, lNewBitmap
gestion_erreurs:
    ' Supprime objets temporaires
    GdipDeleteGraphics lNewGraphics
    PrivBitmap(gUseGdi) = lNewBitmap
    If Err.Number <> 0 Then PrivRotate = False
End Function

'---------------------------------------------------------------------------------------
' Change l'image courante d'un gif anime
'---------------------------------------------------------------------------------------
' pFrame : numero de l'image
'---------------------------------------------------------------------------------------
Public Function GifSetFrame(pFrame As Long) As Boolean
    GifSetFrame = PrivGifSetFrame(pFrame)
End Function

'---------------------------------------------------------------------------------------
' Change l'image courante d'un gif anime
'---------------------------------------------------------------------------------------
' pFrame : numero de l'image
'---------------------------------------------------------------------------------------
Private Function PrivGifSetFrame(pFrame As Long) As Boolean
    Dim lFDTGUID As GUID
    Const lFrameDimensionTime As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lBitmap = gBitmap
    Call CLSIDFromString(StrPtr(lFrameDimensionTime), lFDTGUID)
    PrivGifSetFrame = (GdipImageSelectActiveFrame(lBitmap, lFDTGUID, pFrame) = 0)
    If PrivGifSetFrame Then gActiveFrame = pFrame
    Exit Function
gestion_erreurs:
    If Err.Number <> 0 Then PrivGifSetFrame = False
End Function

'---------------------------------------------------------------------------------------
' Nombre d'images d'un gif anime
'---------------------------------------------------------------------------------------
Public Function GifGetFrameCount() As Long
    GifGetFrameCount = PrivGifGetFrameCount
End Function

'---------------------------------------------------------------------------------------
' Nombre d'images d'un gif anime
'---------------------------------------------------------------------------------------
Private Function PrivGifGetFrameCount() As Long
    Dim lFDTGUID As GUID
    Dim lBitmap
    Const lFrameDimensionTime As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
    On Error GoTo gestion_erreurs
    lBitmap = gBitmap
    Call CLSIDFromString(StrPtr(lFrameDimensionTime), lFDTGUID)
    GdipImageGetFrameCount lBitmap, lFDTGUID, PrivGifGetFrameCount
    Exit Function
gestion_erreurs:
    If Err.Number <> 0 Then PrivGifGetFrameCount = 0
End Function

'---------------------------------------------------------------------------------------
' Delais d'affichage des images d'un gif anime
' Renvoit un tableau contenant autant de delais que d'images dans le gif
'---------------------------------------------------------------------------------------
Public Function GifGetFrameDelay() As Variant
    GifGetFrameDelay = PrivGifGetFrameDelay
End Function

'---------------------------------------------------------------------------------------
' Delais d'affichage des images d'un gif anime
' Renvoit un tableau contenant autant de delais que d'images dans le gif
'---------------------------------------------------------------------------------------
Private Function PrivGifGetFrameDelay() As Variant
    Dim lPropSize As Long
    Dim lBuffer() As Byte
    Dim lReturnBuffer() As Byte
    Dim lPropertyItem As PropertyItem
    Dim lResultLong() As Long
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lBitmap = gBitmap
    If GdipGetPropertyItemSize(lBitmap, PropertyTagFrameDelay, lPropSize) = 0 Then
        ReDim lBuffer(lPropSize - 1)
        If GdipGetPropertyItem(lBitmap, PropertyTagFrameDelay, lPropSize, lBuffer(0)) = 0 Then
            Call RtlMoveMemory(lPropertyItem, lBuffer(0), LenB(lPropertyItem))
            If lPropertyItem.Length > 0 Then
                ' On deplace la valeur dans un tableau (tPropertyItem.Value est un pointeur)
                ReDim lReturnBuffer(lPropertyItem.Length - 1)
                Call RtlMoveMemory(lReturnBuffer(0), _
                                   ByVal lPropertyItem.value, _
                                   lPropertyItem.Length)
                ' Pour chaque image, le delai est stocke dans un Long, donc taille = 4
                ReDim lResultLong(lPropertyItem.Length / 4 - 1)
                RtlMoveMemory lResultLong(0), lReturnBuffer(0), lPropertyItem.Length
                PrivGifGetFrameDelay = lResultLong
            End If
        End If
    End If
    Exit Function
gestion_erreurs:
    If Err.Number <> 0 Then PrivGifGetFrameDelay = Null
End Function

'---------------------------------------------------------------------------------------
' Rotation/Miroir de l'image
' pType     : type de transformation
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
#If VBA6 Then
Public Function RotateFlip(pType As EGdipRotateFlip, Optional pKeepExifData As Boolean = False) As Boolean
#Else
Public Function RotateFlip(pType As Long, Optional pKeepExifData As Boolean = False) As Boolean
#End If
RotateFlip = PrivRotateFlip(pType, pKeepExifData)
End Function

'---------------------------------------------------------------------------------------
' Rotation/Miroir de l'image
' pType     : type de transformation
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
#If VBA6 Then
Private Function PrivRotateFlip(pType As EGdipRotateFlip, Optional pKeepExifData As Boolean = False) As Boolean
#Else
Private Function PrivRotateFlip(pType As Long, Optional pKeepExifData As Boolean = False) As Boolean
#End If
Dim lWidth As Single
Dim lHeight As Single
Dim lNewBitmap
Dim lOldBitmap
Dim lgraphics, lNewGraphics
On Error GoTo gestion_erreurs
lgraphics = GetGraphics()
lOldBitmap = gBitmap
GdipGetImageDimension lOldBitmap, lWidth, lHeight

' On cree un nouveau bitmap pour contenir l'image transformee
GdipCreateBitmapFromGraphics lWidth, lHeight, lgraphics, lNewBitmap
GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
UpdateGraphics lNewGraphics
' Dessine l'image sur le nouveau bitmap
PrivRotateFlip = (GdipDrawImageRectRectI(lNewGraphics, lOldBitmap, 0, 0, lWidth, lHeight, 0, 0, lWidth, lHeight, 2, 0) = 0)
' Applique la transformation
If PrivRotateFlip Then PrivRotateFlip = (GdipImageRotateFlip(lNewBitmap, pType) = 0)
If pKeepExifData Then ImportExifDataFromBitmap lOldBitmap, lNewBitmap

gestion_erreurs:
If Err.Number <> 0 Then PrivRotateFlip = False
On Error Resume Next
GdipDeleteGraphics lNewGraphics
PrivBitmap(gUseGdi) = lNewBitmap
End Function

'---------------------------------------------------------------------------------------
' Redimensionne l'image
'---------------------------------------------------------------------------------------
' pWidth        : Largeur
' pHeight       : Hauteur
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
Public Function Resize(Optional ByVal pWidth As Long = 0, Optional ByVal pHeight As Long = 0, Optional pKeepExifData As Boolean = False) As Boolean
    Resize = PrivResizeImage(pWidth, pHeight, pKeepExifData)
End Function

'---------------------------------------------------------------------------------------
' Redimensionne l'image par un facteur
'---------------------------------------------------------------------------------------
Public Function ScaleI(Optional ByVal psX As Single = 1, Optional ByVal psY As Single = 1, Optional pKeepExifData As Boolean = False) As Boolean
    ScaleI = PrivResizeImage(ImageWidth * psX, ImageHeight * psY, pKeepExifData)
End Function

'---------------------------------------------------------------------------------------
' Redimensionne l'image
'---------------------------------------------------------------------------------------
' pWidth        : Largeur
' pHeight       : Hauteur
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
Private Function PrivResizeImage(Optional pWidth As Long = 0, Optional pHeight As Long = 0, Optional pKeepExifData As Boolean = False) As Boolean
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lNewBitmap
    Dim lgraphics
    Dim lNewGraphics
    Dim lOldBitmap
    Dim limgattr
    On Error GoTo gestion_erreurs
    
    lgraphics = GetGraphics()
    lOldBitmap = gBitmap
    GdipGetImageDimension lOldBitmap, lWidth, lHeight
    If pWidth = 0 And pHeight = 0 Then
        pWidth = lWidth
        pHeight = lHeight
    Else
        If pWidth = 0 Then
            pWidth = pHeight / lHeight * lWidth
        ElseIf pHeight = 0 Then
            pHeight = pWidth / lWidth * lHeight
        End If
    End If
    GdipCreateBitmapFromGraphics pWidth, pHeight, lgraphics, lNewBitmap
    GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
    UpdateGraphics lNewGraphics
    ' Applique un WrapMode = WrapModeTileFlipXY pour éviter la création de bordures au redimensionnements
    GdipCreateImageAttributes limgattr
    GdipSetImageAttributesWrapMode limgattr, 3, 0, True ' 3 = WrapModeTileFlipXY
    PrivResizeImage = (GdipDrawImageRectRectI(lNewGraphics, lOldBitmap, 0, 0, pWidth, pHeight, 0, 0, lWidth, lHeight, 2, limgattr, PtrNull, PtrNull) = 0)
    'PrivResizeImage = (GdipDrawImageRectI(lNewGraphics, lOldBitmap, 0, 0, pWidth, pHeight) = 0)
    If pKeepExifData Then ImportExifDataFromBitmap lOldBitmap, lNewBitmap

gestion_erreurs:
    If Err.Number <> 0 Then PrivResizeImage = False
    On Error Resume Next
    GdipDisposeImageAttributes limgattr
    GdipDeleteGraphics lNewGraphics
    PrivBitmap(gUseGdi) = lNewBitmap
End Function

'---------------------------------------------------------------------------------------
' Decoupe l'image
'---------------------------------------------------------------------------------------
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function Crop(Optional ByVal pX1 As Single, Optional ByVal pY1 As Single, _
            Optional ByVal pX2 As Single, Optional ByVal pY2 As Single, Optional pKeepExifData As Boolean = False) As Boolean
#Else
Public Function Crop(Optional ByVal pX1 As Long, Optional ByVal pY1 As Long, _
            Optional ByVal pX2 As Long, Optional ByVal pY2 As Long, Optional pKeepExifData As Boolean = False) As Boolean
#End If
    Crop = PrivCrop(pX1, pY1, pX2, pY2, pKeepExifData)
End Function

'---------------------------------------------------------------------------------------
' Decoupe l'image
'---------------------------------------------------------------------------------------
' pKeepExifData : Conserve les donnees Exif
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivCrop(Optional ByVal pX1 As Single, Optional ByVal pY1 As Single, _
            Optional ByVal pX2 As Single, Optional ByVal pY2 As Single, Optional pKeepExifData As Boolean = False) As Boolean
#Else
Private Function PrivCrop(Optional ByVal pX1 As Long, Optional ByVal pY1 As Long, _
            Optional ByVal pX2 As Long, Optional ByVal pY2 As Long, Optional pKeepExifData As Boolean = False) As Boolean
#End If
    Dim lCropWidth As Long, lCropHeight As Long
    Dim lNewBitmap, lOldBitmap
    Dim lgraphics, lNewGraphics
    On Error GoTo gestion_erreurs
    CheckCoordinates pX1, pY1, pX2, pY2
    lgraphics = GetGraphics()
    lOldBitmap = gBitmap
    lCropWidth = pX2 - pX1 + 1
    lCropHeight = pY2 - pY1 + 1
    GdipCreateBitmapFromGraphics lCropWidth, lCropHeight, lgraphics, lNewBitmap
    GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
    UpdateGraphics lNewGraphics
    PrivCrop = (GdipDrawImageRectRect(lNewGraphics, lOldBitmap, _
                                      0, 0, lCropWidth, lCropHeight, pX1, pY1, lCropWidth, lCropHeight, 2, 0) = 0)
    If pKeepExifData Then ImportExifDataFromBitmap lOldBitmap, lNewBitmap

gestion_erreurs:
    If Err.Number <> 0 Then PrivCrop = False
    On Error Resume Next
    GdipDeleteGraphics lNewGraphics
    PrivBitmap(gUseGdi) = lNewBitmap
End Function

'---------------------------------------------------------------------------------------
' Remplit l'image d'une couleur
'---------------------------------------------------------------------------------------
' pColor             : Couleur
' pColorGradient     : Deuxieme couleur pour degrade
' pGradientVert      : Mettre a Vrai pour un degrade vertical
' pX1, pY1, pX2, pY2 : Coordonnees du rectangle a traiter
' pAlpha            : Transparence (0 --> 255)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function FillColor(Optional pColor As Long = 0, Optional pColorGradient As Long = -1, Optional pGradientVert As Boolean = False, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pAlpha As Integer = 255) As Boolean
#Else
Public Function FillColor(Optional pColor As Long = 0, Optional pColorGradient As Long = -1, Optional pGradientVert As Boolean = False, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pAlpha As Integer = 255) As Boolean
#End If
    PrivFillColor pColor, pColorGradient, pGradientVert, pX1, pY1, pX2, pY2, pAlpha
End Function

'---------------------------------------------------------------------------------------
' Remplit l'image d'une couleur
'---------------------------------------------------------------------------------------
' pColor             : Couleur
' pColorGradient     : Deuxieme couleur pour degrade
' pGradientVert      : Mettre a Vrai pour un degrade vertical
' pX1, pY1, pX2, pY2 : Coordonnees du rectangle a traiter
' pAlpha            : Transparence (0 --> 255)
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivFillColor(Optional pColor As Long = 0, Optional pColorGradient As Long = -1, Optional pGradientVert As Boolean = False, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional, Optional pAlpha As Integer = 255) As Boolean
    Dim lrect As RECTF
    Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
#Else
Private Function PrivFillColor(Optional pColor As Long = 0, Optional pColorGradient As Long = -1, Optional pGradientVert As Boolean = False, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional, Optional pAlpha As Integer = 255) As Boolean
    Dim lrect As Rect
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
#End If
    Dim lBrush
    Dim lgraphics
    Dim lWidth As Single, lHeight As Single
    Dim lBitmap
    On Error GoTo gestion_erreurs
    lgraphics = GetGraphics()
    lBitmap = gBitmap
    ' Dimensions de l'image
    GdipGetImageDimension lBitmap, lWidth, lHeight
    ' Coordonnees du rectangle a traiter
    lx1 = pX1: lx2 = pX2: ly1 = pY1: ly2 = pY2
    If lx1 = cSingleOptional Then lx1 = 0
    If lx2 = cSingleOptional Then lx2 = lWidth - 1
    If ly1 = cSingleOptional Then ly1 = 0
    If ly2 = cSingleOptional Then ly2 = lHeight - 1
    CheckCoordinates lx1, ly1, lx2, ly2
    ' Rempli la zone
    If FillTexture <> "" Then
        On Error Resume Next
        lBrush = gTextures.item(FillTexture)
        On Error GoTo gestion_erreurs
    End If
    If lBrush = 0 Then
        If pColorGradient = -1 Then
            GdipCreateSolidFill GdipColor(pAlpha, GetColor(pColor)), lBrush
        Else
            lrect.Left = lx1
            lrect.Top = ly1
            lrect.Right = lx2 - lx1 + 1
            lrect.Bottom = ly2 - ly1 + 1
            GdipCreateLineBrushFromRect lrect, GdipColor(pAlpha, GetColor(pColor)), GdipColor(pAlpha, GetColor(pColorGradient)), -pGradientVert, 0, lBrush
        End If
    End If
    GdipFillRectangle lgraphics, lBrush, lx1, ly1, lx2 - lx1 + 1, ly2 - ly1 + 1
gestion_erreurs:
    If Err.Number = 0 Then PrivFillColor = True
    If FillTexture = "" Then GdipDeleteBrush lBrush
End Function

'---------------------------------------------------------------------------------------
' Remplace une couleur de l'image
'---------------------------------------------------------------------------------------
' pOldColor : Couleur a remplacer
' pNewColor : Couleur de remplacement
' pOldAlpha : Ancienne transparence (0 --> 255)
' pNewAlpha : Nouvelle transparence (0 --> 255)
' pX1, pY1, pX2, pY2 : Coordonnees du rectangle a traiter
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function ReplaceColor(ByVal pOldColor As Long, ByVal pNewColor As Long, Optional ByVal pOldAlpha As Long = -1, Optional ByVal pNewAlpha As Long = -1, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional) As Boolean
#Else
Public Function ReplaceColor(ByVal pOldColor As Long, ByVal pNewColor As Long, Optional ByVal pOldAlpha As Long = -1, Optional ByVal pNewAlpha As Long = -1, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional) As Boolean
#End If
    ReplaceColor = PrivReplaceColor(pOldColor, pNewColor, pOldAlpha, pNewAlpha, pX1, pY1, pX2, pY2)
End Function

'---------------------------------------------------------------------------------------
' Remplace une couleur de l'image
'---------------------------------------------------------------------------------------
' pOldColor : Couleur a remplacer
' pNewColor : Couleur de remplacement
' pOldAlpha : Ancienne transparence (0 --> 255)
' pNewAlpha : Nouvelle transparence (0 --> 255)
' pX1, pY1, pX2, pY2 : Coordonnees du rectangle a traiter
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivReplaceColor(ByVal pOldColor As Long, ByVal pNewColor As Long, Optional ByVal pOldAlpha As Long = -1, Optional ByVal pNewAlpha As Long = -1, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional) As Boolean
    Dim lpt1 As POINTF, lpt2 As POINTF
    Dim lSrcWidth As Single, lSrcHeight As Single
#Else
Private Function PrivReplaceColor(ByVal pOldColor As Long, ByVal pNewColor As Long, Optional ByVal pOldAlpha As Long = -1, Optional ByVal pNewAlpha As Long = -1, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional) As Boolean
    Dim lpt1 As POINTAPI, lpt2 As POINTAPI
    Dim lSrcWidth As Long, lSrcHeight As Long
#End If
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lgraphics, lNewGraphics
    Dim lNewBitmap
    Dim limgattr
    Dim lcpt As Long
    Dim lColorMap() As ColorMap
    Dim lOldColorARGB As Long
    Dim lNewColorARGB As Long
    Dim lBitmap
    Dim lMatrix, lResult As Long
    Dim lFull As Boolean
    On Error GoTo gestion_erreurs
    lgraphics = GetGraphics()
    lBitmap = gBitmap
    ' Dimensions de l'image
    GdipGetImageDimension lBitmap, lWidth, lHeight
    ' Coordonnees du rectangle a traiter
    lpt1.X = pX1: lpt1.Y = pY1
    lpt2.X = pX2: lpt2.Y = pY2
    lFull = (pX1 = cSingleOptional And pX2 = cSingleOptional And pY1 = cSingleOptional And pY2 = cSingleOptional)
    If lpt1.X = cSingleOptional Then lpt1.X = 0
    If lpt2.X = cSingleOptional Then lpt2.X = lWidth - 1
    If lpt1.Y = cSingleOptional Then lpt1.Y = 0
    If lpt2.Y = cSingleOptional Then lpt2.Y = lHeight - 1
    lSrcWidth = lpt2.X - lpt1.X + 1
    lSrcHeight = lpt2.Y - lpt1.Y + 1
    GdipCreateMatrix lMatrix
    GdipGetWorldTransform lgraphics, lMatrix
    GdipIsMatrixIdentity lMatrix, lResult
    If Not lResult Then
        If Not lFull Then
            GdipTransformMatrixPoints lMatrix, lpt1, 1
            GdipTransformMatrixPoints lMatrix, lpt2, 1
        End If
        GdipResetWorldTransform lgraphics
    Else
        GdipDeleteMatrix lMatrix
    End If
    ' Clone l'image (on va dessiner le clone sur l'image d'origine)
    If lFull Then
        lNewBitmap = PrivBitmapClone(lBitmap)
        'GdipCloneImage lBitmap, lNewBitmap
    Else
        GdipCreateBitmapFromGraphics lSrcWidth, lSrcHeight, lgraphics, lNewBitmap
        GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
        UpdateGraphics lNewGraphics
        GdipDrawImageRectRect lNewGraphics, lBitmap, 0, 0, lSrcWidth, lSrcHeight, lpt1.X, lpt1.Y, lSrcWidth, lSrcHeight, 2, ByVal PtrNull, ByVal PtrNull, ByVal PtrNull
    End If
    ' Definit la matrice de couleurs
    If pOldAlpha = -1 Then
        ReDim lColorMap(0 To 255)
        For lcpt = 0 To 255
            lOldColorARGB = GdipColor(lcpt, GetColor(pOldColor))
            lColorMap(lcpt).oldColor = lOldColorARGB
            lNewColorARGB = GdipColor(IIf(pNewAlpha <> -1, pNewAlpha, lcpt), GetColor(pNewColor))
            lColorMap(lcpt).newColor = lNewColorARGB
        Next
    Else
        ReDim lColorMap(0 To 0)
        lOldColorARGB = GdipColor(pOldAlpha, GetColor(pOldColor))
        lColorMap(lcpt).oldColor = lOldColorARGB
        lNewColorARGB = GdipColor(IIf(pNewAlpha <> -1, pNewAlpha, lcpt), GetColor(pNewColor))
        lColorMap(lcpt).newColor = lNewColorARGB
    End If
    ' Cree l'objet ImageAttributes
    Call GdipCreateImageAttributes(limgattr)
    ' Applique la matrice de couleurs
    Call GdipSetImageAttributesRemapTable(limgattr, 1, True, UBound(lColorMap) + 1, lColorMap(0))
    ' Dessine l'image clone dans l'image cible, en appliquant les attributs
    GdipSetCompositingMode lgraphics, GdipTransparencySourceCopy
    PrivReplaceColor = (GdipDrawImageRectRect(lgraphics, lNewBitmap, lpt1.X, lpt1.Y, lSrcWidth, lSrcHeight, 0, 0, lSrcWidth, lSrcHeight, 2, limgattr) = 0)
    GdipSetCompositingMode lgraphics, gTransparencyMode
gestion_erreurs:
    If Err.Number <> 0 Then PrivReplaceColor = False
    On Error Resume Next
    If Not lResult Then
        GdipSetWorldTransform lgraphics, lMatrix
        GdipDeleteMatrix lMatrix
    End If
    ' Supprime les objets temporaires
    GdipDisposeImageAttributes limgattr
    GdipDeleteGraphics lNewGraphics
    GdipDisposeImage lNewBitmap
End Function

'---------------------------------------------------------------------------------------
' Applique une matrice de couleurs a l'image
'---------------------------------------------------------------------------------------
' pRed, pGreen, pBlue, pAlpha : Tableaux de single de dimension 5
'   1er a 3eme : Multiplicateurs des composantes rouge, vert, bleu
'   4eme       : Constante
'   5eme       : Valeur a ajouter
' pEmptyColor : Couleur de remplissage (-1 pour dessiner par dessus l'image)
' pX1, pY1, pX2, pY2 : Coordonnees du rectangle a traiter
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function ApplyColorMatrix(ByVal pRed As Variant, ByVal pGreen As Variant, ByVal pBlue As Variant, ByVal pAlpha As Variant, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional) As Boolean
#Else
Public Function ApplyColorMatrix(ByVal pRed As Variant, ByVal pGreen As Variant, ByVal pBlue As Variant, ByVal pAlpha As Variant, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional) As Boolean
#End If
    ApplyColorMatrix = PrivApplyColorMatrix(pRed, pGreen, pBlue, pAlpha, pX1, pY1, pX2, pY2)
End Function

'---------------------------------------------------------------------------------------
' Applique une matrice de couleurs a l'image
'---------------------------------------------------------------------------------------
' pRed, pGreen, pBlue, pAlpha : Tableaux de single de dimension 5
'   1er a 3eme : Multiplicateurs des composantes rouge, vert, bleu
'   4eme       : Constante
'   5eme       : Valeur a ajouter
' pEmptyColor : Couleur de remplissage (-1 pour dessiner par dessus l'image)
' pX1, pY1, pX2, pY2 : Coordonnees du rectangle a traiter
'---------------------------------------------------------------------------------------
#If UseSingle Then
Private Function PrivApplyColorMatrix(ByVal pRed As Variant, ByVal pGreen As Variant, ByVal pBlue As Variant, ByVal pAlpha As Variant, Optional ByVal pX1 As Single = cSingleOptional, Optional ByVal pY1 As Single = cSingleOptional, Optional ByVal pX2 As Single = cSingleOptional, Optional ByVal pY2 As Single = cSingleOptional) As Boolean
    Dim lpt1 As POINTF, lpt2 As POINTF
#Else
Private Function PrivApplyColorMatrix(ByVal pRed As Variant, ByVal pGreen As Variant, ByVal pBlue As Variant, ByVal pAlpha As Variant, Optional ByVal pX1 As Long = cSingleOptional, Optional ByVal pY1 As Long = cSingleOptional, Optional ByVal pX2 As Long = cSingleOptional, Optional ByVal pY2 As Long = cSingleOptional) As Boolean
    Dim lpt1 As POINTAPI, lpt2 As POINTAPI
#End If
    Dim lWidth As Single
    Dim lHeight As Single
    Dim lgraphics, lNewGraphics
    Dim lNewBitmap
    Dim limgattr
    Dim lMatrix As ColorMatrix
    Dim lRed0 As Integer
    Dim lGreen0 As Integer
    Dim lBlue0 As Integer
    Dim lAlpha0 As Integer
    Dim lOldBitmap
    Dim lMatrixSrc, lResult As Long
    Dim lSrcWidth As Long, lSrcHeight As Long
    Dim lFull As Boolean
    On Error GoTo gestion_erreurs
    lgraphics = GetGraphics()
    lOldBitmap = gBitmap
    ' Dimensions de l'image
    GdipGetImageDimension lOldBitmap, lWidth, lHeight
    ' Coordonnees du rectangle a traiter
    lpt1.X = pX1: lpt1.Y = pY1
    lpt2.X = pX2: lpt2.Y = pY2
    lFull = (pX1 = cSingleOptional And pX2 = cSingleOptional And pY1 = cSingleOptional And pY2 = cSingleOptional)
    If lpt1.X = cSingleOptional Then lpt1.X = 0
    If lpt2.X = cSingleOptional Then lpt2.X = lWidth - 1
    If lpt1.Y = cSingleOptional Then lpt1.Y = 0
    If lpt2.Y = cSingleOptional Then lpt2.Y = lHeight - 1
    lSrcWidth = lpt2.X - lpt1.X + 1
    lSrcHeight = lpt2.Y - lpt1.Y + 1
    GdipCreateMatrix lMatrixSrc
    GdipGetWorldTransform lgraphics, lMatrixSrc
    GdipIsMatrixIdentity lMatrixSrc, lResult
    If Not lResult Then
        If Not lFull Then
            GdipTransformMatrixPoints lMatrixSrc, lpt1, 1
            GdipTransformMatrixPoints lMatrixSrc, lpt2, 1
        End If
        GdipResetWorldTransform lgraphics
    Else
        GdipDeleteMatrix lMatrixSrc
    End If
    ' Clone l'image (on va dessiner le clone sur l'image d'origine)
    If lFull Then
        lNewBitmap = PrivBitmapClone(lOldBitmap)
        'GdipCloneImage lOldBitmap, lNewBitmap
    Else
        GdipCreateBitmapFromGraphics lSrcWidth, lSrcHeight, lgraphics, lNewBitmap
        GdipGetImageGraphicsContext lNewBitmap, lNewGraphics
        GdipDrawImageRectRect lNewGraphics, lOldBitmap, 0, 0, lSrcWidth, lSrcHeight, lpt1.X, lpt1.Y, lSrcWidth, lSrcHeight, 2, ByVal PtrNull, ByVal PtrNull, ByVal PtrNull
    End If
    ' Definit la matrice de couleurs
    lRed0 = LBound(pRed)
    ReDim Preserve pRed(4 - LBound(pRed))
    lGreen0 = LBound(pGreen)
    ReDim Preserve pGreen(4 - LBound(pGreen))
    lBlue0 = LBound(pBlue)
    ReDim Preserve pBlue(4 - LBound(pBlue))
    lAlpha0 = LBound(pAlpha)
    ReDim Preserve pAlpha(4 - LBound(pAlpha))
    lMatrix.m(0, 0) = pRed(lRed0): lMatrix.m(1, 0) = pGreen(lGreen0): lMatrix.m(2, 0) = pBlue(lBlue0): lMatrix.m(3, 0) = pAlpha(lAlpha0): lMatrix.m(4, 0) = 0
    lMatrix.m(0, 1) = pRed(lRed0 + 1): lMatrix.m(1, 1) = pGreen(lGreen0 + 1): lMatrix.m(2, 1) = pBlue(lBlue0 + 1): lMatrix.m(3, 1) = pAlpha(lAlpha0 + 1): lMatrix.m(4, 1) = 0
    lMatrix.m(0, 2) = pRed(lRed0 + 2): lMatrix.m(1, 2) = pGreen(lGreen0 + 2): lMatrix.m(2, 2) = pBlue(lBlue0 + 2): lMatrix.m(3, 2) = pAlpha(lAlpha0 + 2): lMatrix.m(4, 2) = 0
    lMatrix.m(0, 3) = pRed(lRed0 + 3): lMatrix.m(1, 3) = pGreen(lGreen0 + 3): lMatrix.m(2, 3) = pBlue(lBlue0 + 3): lMatrix.m(3, 3) = pAlpha(lAlpha0 + 3): lMatrix.m(4, 3) = 0
    lMatrix.m(0, 4) = pRed(lRed0 + 4): lMatrix.m(1, 4) = pGreen(lGreen0 + 4): lMatrix.m(2, 4) = pBlue(lBlue0 + 4): lMatrix.m(3, 4) = pAlpha(lAlpha0 + 4): lMatrix.m(4, 4) = 1
    ' Cree l'objet ImageAttributes
    Call GdipCreateImageAttributes(limgattr)
    ' Applique la matrice de couleurs
    Call GdipSetImageAttributesColorMatrix(limgattr, 1, True, lMatrix, ByVal 0, 0)
    ' Dessine l'image clone dans l'image cible, en appliquant les attributs
    GdipSetCompositingMode lgraphics, GdipTransparencySourceCopy
    Call GdipDrawImageRectRect(lgraphics, lNewBitmap, lpt1.X, lpt1.Y, lSrcWidth, lSrcHeight, 0, 0, lSrcWidth, lSrcHeight, 2, limgattr)
    GdipSetCompositingMode lgraphics, gTransparencyMode
gestion_erreurs:
    If Err.Number <> 0 Then PrivApplyColorMatrix = False
    On Error Resume Next
    If Not lResult Then
        GdipSetWorldTransform lgraphics, lMatrixSrc
        GdipDeleteMatrix lMatrixSrc
    End If
    ' Supprime les objets temporaires
    GdipDisposeImageAttributes limgattr
    GdipDeleteGraphics lNewGraphics
    GdipDisposeImage lNewBitmap
End Function


'***************************************************************************************
'*                             DESSIN DIRECT SUR FORMULAIRE                            *
'***************************************************************************************

'---------------------------------------------------------------------------------------
' Corrections pour RepaintFast
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Sub RepaintFastSetCorrection(Optional ByVal pLeftDest As Single, Optional ByVal pTopDest As Single, Optional ByVal pWidthDest As Single, Optional ByVal pHeightDest As Single, _
                                    Optional ByVal pLeftSrc As Single, Optional ByVal pTopSrc As Single, Optional ByVal pWidthSrc As Single, Optional ByVal pHeightSrc As Single)
#Else
Public Sub RepaintFastSetCorrection(Optional ByVal pLeftDest As Long, Optional ByVal pTopDest As Long, Optional ByVal pWidthDest As Long, Optional ByVal pHeightDest As Long, _
                                    Optional ByVal pLeftSrc As Long, Optional ByVal pTopSrc As Long, Optional ByVal pWidthSrc As Long, Optional ByVal pHeightSrc As Long)
#End If
    gRepaintCorrection(1) = pLeftDest
    gRepaintCorrection(2) = pTopDest
    gRepaintCorrection(3) = pWidthDest
    gRepaintCorrection(4) = pHeightDest
    gRepaintCorrection(5) = pLeftSrc
    gRepaintCorrection(6) = pTopSrc
    gRepaintCorrection(7) = pWidthSrc
    gRepaintCorrection(8) = pHeightSrc
End Sub

'---------------------------------------------------------------------------------------
' Clip pour RepaintFast
' N'est redessine que l'interieur du controle ou l'exterieur si pInclude = False
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function RepaintFastSetClipControl(pCtrl As Control, Optional pInclude As Boolean = True, Optional pCorrectLeft As Single, Optional pCorrectTop As Single, Optional pCorrectRight As Single, Optional pCorrectBottom As Single) As Boolean
    Dim lrect As RECTF
#Else
Public Function RepaintFastSetClipControl(pCtrl As Control, Optional pInclude As Boolean = True, Optional pCorrectLeft As Long, Optional pCorrectTop As Long, Optional pCorrectRight As Long, Optional pCorrectBottom As Long) As Boolean
    Dim lrect As Rect
#End If
    Dim lregion
    On Error GoTo gestion_erreurs
    GetControlPos pCtrl, lrect.Left, lrect.Top, lrect.Right, lrect.Bottom, True
    lrect.Left = lrect.Left + pCorrectLeft
    lrect.Top = lrect.Top + pCorrectTop
    lrect.Right = lrect.Right - 1 - lrect.Left + 1 + pCorrectRight
    lrect.Bottom = lrect.Bottom - 1 - lrect.Top + 1 + pCorrectBottom
    GdipCreateRegionRect lrect, lregion
    AddClipRegion lregion, pInclude
    GdipDeleteRegion lregion
    RepaintFastSetClipControl = True
    Exit Function
gestion_erreurs:
    RepaintFastSetClipControl = False
End Function
'---------------------------------------------------------------------------------------
' Clip pour RepaintFast
' N'est redessine que l'interieur du rectangle ou l'exterieur si pInclude = False
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function RepaintFastSetClipRectangle(pObject As Object, _
            Optional ByVal pX1 As Single, Optional ByVal pY1 As Single, _
            Optional ByVal pX2 As Single, Optional ByVal pY2 As Single, _
            Optional pInclude As Boolean = True, _
            Optional ByVal pCorrectLeft As Single, Optional ByVal pCorrectTop As Single, _
            Optional ByVal pCorrectRight As Single, Optional ByVal pCorrectBottom As Single) As Boolean
    Dim lrect As RECTF
#Else
Public Function RepaintFastSetClipRectangle(pObject As Object, _
            Optional ByVal pX1 As Long, Optional ByVal pY1 As Long, _
            Optional ByVal pX2 As Long, Optional ByVal pY2 As Long, _
            Optional pInclude As Boolean = True, _
            Optional ByVal pCorrectLeft As Long, Optional ByVal pCorrectTop As Long, _
            Optional ByVal pCorrectRight As Long, Optional ByVal pCorrectBottom As Long) As Boolean
    Dim lrect As Rect
#End If
    Dim lregion
On Error GoTo gestion_erreurs
    lrect.Left = PointsToPixelsX(pObject.Left + ImgToCtrlX(pX1, pObject)) - 1 + pCorrectLeft
    lrect.Top = PointsToPixelsY(pObject.Top + ImgToCtrlY(pY1, pObject)) - 1 + pCorrectTop
    lrect.Right = PointsToPixelsX(pObject.Left + ImgToCtrlX(pX2, pObject)) - 1 - lrect.Left + 1 + pCorrectRight
    lrect.Bottom = PointsToPixelsY(pObject.Top + ImgToCtrlY(pY2, pObject)) - 1 - lrect.Top + 1 + pCorrectBottom
    GdipCreateRegionRect lrect, lregion
    AddClipRegion lregion, pInclude
    GdipDeleteRegion lregion
    RepaintFastSetClipRectangle = True
    Exit Function
gestion_erreurs:
    RepaintFastSetClipRectangle = False
End Function
'---------------------------------------------------------------------------------------
' Clip pour RepaintFast
' N'est redessine que le rectangle entourant la region ou l'exterieur si pInclude = False
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function RepaintFastSetClipRegion(pObject As Object, _
            pRegion As String, _
            Optional pInclude As Boolean = True, _
            Optional ByVal pCorrectLeft As Single, Optional ByVal pCorrectTop As Single, _
            Optional ByVal pCorrectRight As Single, Optional ByVal pCorrectBottom As Single) As Boolean
    Dim lrect As RECTF
#Else
Public Function RepaintFastSetClipRegion(pObject As Object, _
            pRegion As String, _
            Optional pInclude As Boolean = True, _
            Optional ByVal pCorrectLeft As Long, Optional ByVal pCorrectTop As Long, _
            Optional ByVal pCorrectRight As Long, Optional ByVal pCorrectBottom As Long) As Boolean
    Dim lrect As Rect
#End If
    Dim lregion
    Dim lgraphics
    Dim lRight As Long, lBottom As Long
On Error GoTo gestion_erreurs
    If gCollRegions.item(pRegion)(2) Then
        GdipGetPathWorldBounds gCollRegions.item(pRegion)(0), lrect, ByVal PtrNull, ByVal PtrNull
    Else
        ' Contexte graphique
        lgraphics = GetGraphics()
        GdipGetRegionBounds gCollRegions.item(pRegion)(0), lgraphics, lrect
    End If
    lRight = lrect.Left + lrect.Right - 1
    lBottom = lrect.Top + lrect.Bottom - 1
    lrect.Left = PointsToPixelsX(pObject.Left + ImgToCtrlX(lrect.Left - 1, pObject)) - 1 + pCorrectLeft
    lrect.Top = PointsToPixelsY(pObject.Top + ImgToCtrlY(lrect.Top - 1, pObject)) - 1 + pCorrectTop
    lrect.Right = PointsToPixelsX(pObject.Left + ImgToCtrlX(lRight + 1, pObject)) - 1 - lrect.Left + 1 + pCorrectRight
    lrect.Bottom = PointsToPixelsY(pObject.Top + ImgToCtrlY(lBottom + 1, pObject)) - 1 - lrect.Top + 1 + pCorrectBottom
    GdipCreateRegionRect lrect, lregion
    AddClipRegion lregion, pInclude
    GdipDeleteRegion lregion
    RepaintFastSetClipRegion = True
    Exit Function
gestion_erreurs:
    RepaintFastSetClipRegion = False
End Function
Public Sub RepaintFastResetClip()
    If gClipRegion <> PtrNull Then
        GdipDeleteRegion gClipRegion
        gClipRegion = PtrNull
    End If
End Sub
Private Sub AddClipRegion(pRegion, pInclude As Boolean)
    If gClipRegion <> 0 Then
        GdipCombineRegionRegion gClipRegion, pRegion, IIf(pInclude, 2, 4)
    Else
        If pInclude Then
            GdipCloneRegion pRegion, gClipRegion
        Else
            GdipCreateRegion gClipRegion
            GdipSetInfinite gClipRegion
            GdipCombineRegionRegion gClipRegion, pRegion, 4
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Dessine directement sur le formulaire
' (Plus rapide mais temporaire)
'---------------------------------------------------------------------------------------
' pObject  : Controle determinant l'emplacement de l'image
'---------------------------------------------------------------------------------------
#If VBA7 Then
Public Function RepaintFast(Optional pObject As Object, Optional pTranspColor As Long = -1, Optional pPercent As Integer = 255, Optional pAntialiase As Boolean = False, Optional phDC As LongPtr, Optional pHwnd As LongPtr)
#Else
Public Function RepaintFast(Optional pObject As Object, Optional pTranspColor As Long = -1, Optional pPercent As Integer = 255, Optional pAntialiase As Boolean = False, Optional phDC As Long, Optional pHwnd As Long)
#End If
#If UseSingle Then
    Static lrect As RECTF
    Dim lContainerX As Single, lContainerY As Single
    Static lTop As Single, lLeft As Single
    Static lWidth As Single, lHeight As Single
    Dim lx1 As Single, ly1 As Single, lx2 As Single, ly2 As Single
    Static lDecalageX As Single, lDecalageY As Single
    Dim lImgWidth As Single, lImgHeight As Single
#Else
    Static lrect As Rect
    Dim lContainerX As Long, lContainerY As Long
    Static lTop As Long, lLeft As Long
    Static lWidth As Long, lHeight As Long
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
    Static lDecalageX As Long, lDecalageY As Long
    Dim lImgWidth As Long, lImgHeight As Long
#End If
    Dim lhrgn
    Dim lRectI As Rect
    Dim lgraphics
    Dim lImgWidthS As Single, lImgHeightS As Single
    Dim lHdc
    Static lHwnd
    Dim siX As SCROLLINFO
    Dim siY As SCROLLINFO
    Dim lMatrix As ColorMatrix
    Dim limgattr
    Dim lRgnValidate
    Dim lGdiRegion
    Dim lScrollHwndX
    Dim lScrollHwndY
    Dim lWindowRect As Rect
    Dim lSizeMode As Long
    Dim lParentForm As Object
    Dim lParentHwnd
    Dim lOldCM As Long, lNewCM As Long
    Dim lUseGdi As Boolean
    Dim lObjPtr
    
    Static sObjPtr
    Static sBitmap
    If pObject Is Nothing Then Set pObject = oObject
    lObjPtr = ObjPtr(pObject) ' conserve le pointeur, le passage dans GetClientHandle peut modifier le pointeur
    If gDib.hDC <> 0 And Not RepaintFastTransparency Then lUseGdi = True
    If RepaintFastResetCalc Then
        sBitmap = PtrNull
        sObjPtr = PtrNull
        RepaintFastResetCalc = False
    End If
    If Not pObject Is Nothing And (sObjPtr <> lObjPtr Or sBitmap <> gBitmap) Then
       #If AppName = "A" Then
            If TypeOf pObject Is Access.Form Then
                Set lParentForm = pObject
            Else
                Set lParentForm = pObject.Parent
            End If
            ' Remonte jusqu'au formulaire si controle dans onglets
            If TypeOf lParentForm Is Page Then
                On Error Resume Next
                Do
                    Set lParentForm = lParentForm.Parent
                    If Err.Number <> 0 Then Exit Do
                Loop
                On Error GoTo 0
            End If
            If Not lParentForm Is Nothing Then
                lParentHwnd = lParentForm.hwnd
                Set lParentForm = Nothing
            Else
                Exit Function
            End If
            lHwnd = GetClientHandle(pObject)
            ' Position de la barre de defilement horizontale
            #If Win64 Then
            siX.cbSize = LenB(siX)
            #Else
            siX.cbSize = Len(siX)
            #End If
            siX.fMask = SIF_ALL
            lScrollHwndX = GetScrollBarHwnd(lParentHwnd, SB_HORZ)
            Call GetScrollInfo(lScrollHwndX, SB_CTL, siX)
            ' Position de la barre de defilement verticale
            lScrollHwndY = GetScrollBarHwnd(lParentHwnd, SB_VERT)
            #If Win64 Then
            siY.cbSize = LenB(siY)
            #Else
            siY.cbSize = Len(siY)
            #End If
            siY.fMask = SIF_ALL
            Call GetScrollInfo(lScrollHwndY, SB_CTL, siY)
        #Else
            If TypeOf pObject Is MSForms.UserForm And Not TypeOf pObject Is MSForms.Frame Then
                Set lParentForm = pObject
            Else
                Set lParentForm = pObject.Parent
            End If
            lHwnd = GetUserFormHandle(lParentForm, True)
            If TypeOf lParentForm Is Frame Then
                GetControlPos lParentForm
                lContainerX = CLng(GetControlPos(lParentForm)) - 1
                lContainerY = CLng(lContainerX)
            End If
        #End If
        On Error Resume Next
        siX.nPos = Me.PointsToPixelsX(lParentForm.ScrollLeft)
        siY.nPos = Me.PointsToPixelsY(lParentForm.ScrollTop)
        On Error GoTo 0
    End If
    If lHwnd = PtrNull Then Exit Function
    If pHwnd <> 0 Then lHwnd = pHwnd
    GdipGetImageDimension gBitmap, lImgWidthS, lImgHeightS
    lImgWidth = CLng(lImgWidthS)
    lImgHeight = CLng(lImgHeightS)
    If phDC <> 0 Then
        lHdc = phDC
    Else
        lHdc = GetWindowDC(lHwnd)
    End If
    
    GdipCreateFromHDC lHdc, lgraphics
    
    If (sObjPtr <> lObjPtr Or sBitmap <> gBitmap) Then
        If Not pObject Is Nothing Then
            #If AppName = "A" Then
            If TypeOf pObject Is Access.Form Then
            #Else
            If TypeOf pObject Is MSForms.UserForm And Not TypeOf pObject Is MSForms.Frame Then
            #End If
                lLeft = 1
                lTop = 1
                lWidth = Me.PointsToPixelsX(pObject.InsideWidth)
                lHeight = Me.PointsToPixelsY(pObject.InsideHeight)
            Else
                lLeft = Me.PointsToPixelsX(pObject.Left)
                lTop = Me.PointsToPixelsY(pObject.Top)
                lWidth = Me.PointsToPixelsX(pObject.Width)
                lHeight = Me.PointsToPixelsY(pObject.Height)
            End If
            lLeft = CLng(lLeft) + lContainerX
            lTop = CLng(lTop) + lContainerY
            lx2 = lImgWidth - 1
            ly2 = lImgHeight - 1
            CalcXY lx1, ly1, lx2, ly2, lImgWidth, lImgHeight, lDecalageX, lDecalageY, PictureSizeMode(pObject), PictureAlignment(pObject), pObject
            lLeft = lLeft + lx1 - siX.nPos - 1
            lTop = lTop + ly1 - siY.nPos - 1
            lWidth = CLng(lx2) - lx1 + 1
            lHeight = CLng(ly2) - ly1 + 1
            ' Correction d'affichage si double-buffer
            #If AppName = "A" Then
                If (GetWindowLong(lParentHwnd, GWL_EXSTYLE) And WS_EX_COMPOSITED) = WS_EX_COMPOSITED Then
                    GetWindowRect lParentHwnd, lWindowRect
                    lLeft = lLeft + lWindowRect.Left
                    lTop = lTop + lWindowRect.Top
                End If
            #End If
        Else
            lLeft = 0
            lTop = 0
            lWidth = lImgWidth
            lHeight = lImgHeight
        End If
    End If
    ' Annule le dessin du formulaire
    If gCancelPaint Then
        If sObjPtr <> lObjPtr Then
            GetControlPos pObject, lrect.Left, lrect.Top, lrect.Right, lrect.Bottom
            lrect.Left = lrect.Left - 1
            lrect.Top = lrect.Top - 1
            lrect.Right = lrect.Right ' + 1
            lrect.Bottom = lrect.Bottom '+ 1
        End If
        If gClipRegion <> 0 Then
            ' Si le dessin par RepaintFast est limite a une region, on en tient compte
            GdipCreateRegionRect lrect, lRgnValidate
            GdipCombineRegionRegion lRgnValidate, gClipRegion, CombineModeIntersect
            GdipGetRegionHRgn lRgnValidate, lgraphics, lGdiRegion
            ValidateRgn lHwnd, lGdiRegion
            DeleteObject lGdiRegion
            GdipDeleteRegion lRgnValidate
        Else
            lRectI.Left = lrect.Left
            lRectI.Top = lrect.Top
            lRectI.Right = lrect.Right
            lRectI.Bottom = lrect.Bottom
            ValidateRect lHwnd, lRectI
        End If
    End If
    ' Transparence de l'image
    If pPercent <> 255 Then
        ' Cree l'objet ImageAttributes
        Call GdipCreateImageAttributes(limgattr)
        ' Cree la matrice
        lMatrix.m(0, 0) = 1: lMatrix.m(1, 0) = 0: lMatrix.m(2, 0) = 0: lMatrix.m(3, 0) = 0: lMatrix.m(4, 0) = 0
        lMatrix.m(0, 1) = 0: lMatrix.m(1, 1) = 1: lMatrix.m(2, 1) = 0: lMatrix.m(3, 1) = 0: lMatrix.m(4, 1) = 0
        lMatrix.m(0, 2) = 0: lMatrix.m(1, 2) = 0: lMatrix.m(2, 2) = 1: lMatrix.m(3, 2) = 0: lMatrix.m(4, 2) = 0
        lMatrix.m(0, 3) = 0: lMatrix.m(1, 3) = 0: lMatrix.m(2, 3) = 0: lMatrix.m(3, 3) = pPercent / 255: lMatrix.m(4, 3) = 0
        lMatrix.m(0, 4) = 0: lMatrix.m(1, 4) = 0: lMatrix.m(2, 4) = 0: lMatrix.m(3, 4) = 0: lMatrix.m(4, 4) = 1
        ' Applique la matrice de couleurs
        Call GdipSetImageAttributesColorMatrix(limgattr, 1, True, lMatrix, ByVal 0, 0)
    End If
    ' Couleur de transparence
    If pTranspColor <> -1 Then
        ' Cree l'objet ImageAttributes
        If limgattr = 0 Then Call GdipCreateImageAttributes(limgattr)
        ' Applique les couleurs de transparence
        GdipSetImageAttributesColorKeys limgattr, 1, True, GdipColor(0, GetColor(pTranspColor)), GdipColor(255, GetColor(pTranspColor))
    End If
    ' Antialiase
    If pAntialiase Then
        GdipSetInterpolationMode lgraphics, gAntialisingLevel
    Else
        GdipSetInterpolationMode lgraphics, 5 'GdipAntialiseNearestNeighbor
    End If
    ' Limite la region d'affichage
    If gClipRegion <> PtrNull Then
        If lUseGdi Then
            GdipGetRegionHRgn gClipRegion, lgraphics, lhrgn
            SelectClipRgn lHdc, lhrgn
        Else
            GdipSetClipRegion lgraphics, gClipRegion, 0
        End If
    End If
    ' Mode remplace
    GdipGetCompositingMode lgraphics, lOldCM
    ' Dessine l'image sur le formulaire
    If Not pObject Is Nothing Then
        lSizeMode = PictureSizeMode(pObject)
    Else
        lSizeMode = GdipSizeModeClip
    End If
    #If AppName = "A" Then
    #Else
    If Not lParentForm Is Nothing Then
        If lLeft + lWidth - 1 > Me.PointsToPixelsX(lParentForm.InsideWidth) Then lWidth = Me.PointsToPixelsX(lParentForm.InsideWidth) - lLeft + 2
        If lTop + lHeight - 1 > Me.PointsToPixelsY(lParentForm.InsideHeight) Then lHeight = Me.PointsToPixelsY(lParentForm.InsideHeight) - lTop + 2
    End If
    #End If
    If lSizeMode = GdipSizeModeClip Then
        If lUseGdi Then
            BitBlt lHdc, lLeft + gRepaintCorrection(1), lTop + gRepaintCorrection(2), lWidth, lHeight, gDib.hDC, IIf(lDecalageX < 0, -lDecalageX, 0) + gRepaintCorrection(5), IIf(lDecalageY < 0, -lDecalageY, 0) + gRepaintCorrection(6), SRCCOPY
        Else
            lNewCM = (Not RepaintFastTransparency) * -1
            If lNewCM <> lOldCM Then GdipSetCompositingMode lgraphics, lNewCM
            Call GdipDrawImageRectRect(lgraphics, gBitmap, lLeft + gRepaintCorrection(1), lTop + gRepaintCorrection(2), lWidth + gRepaintCorrection(3), lHeight + gRepaintCorrection(4), _
                                    IIf(lDecalageX < 0, -lDecalageX, 0) + gRepaintCorrection(5), IIf(lDecalageY < 0, -lDecalageY, 0) + gRepaintCorrection(6), lWidth + gRepaintCorrection(7), lHeight + gRepaintCorrection(8), 2, limgattr)
        End If
    Else
        lNewCM = (Not RepaintFastTransparency) * -1
        If lNewCM <> lOldCM Then GdipSetCompositingMode lgraphics, lNewCM
        Call GdipDrawImageRectRect(lgraphics, gBitmap, lLeft + gRepaintCorrection(1), lTop + gRepaintCorrection(2), lWidth + gRepaintCorrection(3), lHeight + gRepaintCorrection(4), _
                                    gRepaintCorrection(5), gRepaintCorrection(6), lImgWidthS + gRepaintCorrection(7), lImgHeightS + gRepaintCorrection(8), 2, limgattr)
    End If
    GdipSetCompositingMode lgraphics, lOldCM
    If gClipRegion <> PtrNull Then GdipResetClip lgraphics
    If lhrgn <> PtrNull Then DeleteObject lhrgn
    If limgattr <> PtrNull Then GdipDisposeImageAttributes limgattr
    GdipDeleteGraphics lgraphics
    If phDC = PtrNull Then ReleaseDC lHwnd, lHdc
    If Not gCancelPaint And RepaintMatrixUpdate Then PrivRepaintMatrixUpdate pObject
    
    sObjPtr = lObjPtr: sBitmap = gBitmap
End Function

#If AppName = "A" Then
        '---------------------------------------------------------------------------------------
        ' Handle du client de la section du formulaire qui contient l'image
        '---------------------------------------------------------------------------------------
Public Function GetClientHandle(pImgCtrl As Control)
    Dim lHwnd
    Dim lret As Long
    Dim lClassName As String
    Dim lHeaderFooter As Boolean
    Dim lcpt As Integer
    Dim lParentForm As Object
    Dim lParentHwnd
    Set lParentForm = pImgCtrl.Parent
    ' Remonte jusqu'au formulaire si controle dans onglets
    If TypeOf lParentForm Is Page Then
        On Error Resume Next
        Do
            Set lParentForm = lParentForm.Parent
            If Err.Number <> 0 Then Exit Do
        Loop
        On Error GoTo 0
    End If
    If Not lParentForm Is Nothing Then
        lParentHwnd = lParentForm.hwnd
        Set lParentForm = Nothing
    Else
        Exit Function
    End If
    lHwnd = GetWindow(lParentHwnd, GW_CHILD)
    ' Compte le nombre de client oFormSub
    Do While lHwnd
        lClassName = Space(255)
        lret = GetClassName(lHwnd, lClassName, 255)
        lClassName = Left(lClassName, lret)
        If lClassName = "OFormSub" Then
            lcpt = lcpt + 1
        End If
        lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
    Loop
    If lcpt > 1 Then lHeaderFooter = True
    lcpt = 0
    lHwnd = GetWindow(lParentHwnd, GW_CHILD)
    Do While lHwnd
        lClassName = Space(255)
        lret = GetClassName(lHwnd, lClassName, 255)
        lClassName = Left(lClassName, lret)
        If lClassName = "OFormSub" Then
            lcpt = lcpt + 1
            Select Case pImgCtrl.Section
            Case acDetail
                If (lHeaderFooter And lcpt = 2) Or (Not lHeaderFooter And lcpt = 1) Then
                    GetClientHandle = lHwnd
                    Exit Do
                End If
            Case acHeader
                If lcpt = 1 Then
                    GetClientHandle = lHwnd
                    Exit Do
                End If
            Case acFooter
                If lcpt = 3 Then
                    GetClientHandle = lHwnd
                    Exit Do
                End If
            End Select
        End If
        lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
    Loop
End Function

'---------------------------------------------------------------------------------------
' Handle de la barre de defilement
'---------------------------------------------------------------------------------------
Private Function GetScrollBarHwnd(ByVal FormhWnd, ByVal BarType As Long)
    Dim lret As Long
    Dim lClassName As String
    Dim lWindowText As String
    Dim CurrenthWnd
    CurrenthWnd = GetWindow(FormhWnd, GW_CHILD)
    Do Until CurrenthWnd = 0
        Call GetScrollBarHwnd(CurrenthWnd, BarType)
        lClassName = Space(255)
        lret = GetClassName(CurrenthWnd, lClassName, 255)
        lClassName = Left(lClassName, lret)
        If lClassName = "Scrollbar" Or lClassName = "NUIScrollbar" Then
            If val(SysCmd(acSysCmdAccessVer)) >= 12 Then
                lWindowText = Space(255)
                lret = GetWindowText(CurrenthWnd, lWindowText, 255)
                lWindowText = Left(lWindowText, lret)
                If lWindowText = "Horizontal" Then
                    lret = SB_HORZ
                ElseIf lWindowText = "Vertical" Then
                    lret = SB_VERT
                End If
            Else
                lret = GetDlgCtrlID(CurrenthWnd)
            End If
            If lret = BarType Then
                GetScrollBarHwnd = CurrenthWnd
                Exit Function
            End If
        End If
        CurrenthWnd = GetWindow(CurrenthWnd, GW_HWNDNEXT)
    Loop
End Function
#Else
'---------------------------------------------------------------------------------------
' Handle d'un UserForm (ou d'un cadre)
'---------------------------------------------------------------------------------------
' pForm : Formulaire, cadre, image
'---------------------------------------------------------------------------------------
Public Function GetUserFormHandle(pObject As Object, Optional pClientArea As Boolean = False)
    Dim lUserFormHwnd
    Dim lCaption As String
    Dim lVar As Control
    Dim lFrame As Control
    Dim lParentForm As Object
On Error GoTo gestion_erreurs
    If TypeOf pObject Is MSForms.Frame Or TypeOf pObject Is MSForms.Page Then
        Set lFrame = pObject
        If pClientArea Then
            GetUserFormHandle = lFrame.[_GethWnd]
            'GetUserFormHandle = FindClient(GetUserFormHandle, lFrame)
        Exit Function
        End If
    End If
    ' Recherche userform
    If Not TypeOf pObject Is MSForms.UserForm Then
        Set lParentForm = pObject.Parent
        Do
            If TypeOf lParentForm Is MSForms.Frame Then Set lFrame = lParentForm
            If TypeOf lParentForm Is MSForms.UserForm And Not TypeOf lParentForm Is MSForms.Page And Not TypeOf lParentForm Is MSForms.MultiPage Then Exit Do
            Set lParentForm = lParentForm.Parent
        Loop
    Else
        Set lParentForm = pObject
    End If
    ' Recherche Frame ou page
    If lFrame Is Nothing Then
    For Each lVar In lParentForm.Controls
        If TypeOf lVar Is MSForms.Frame Or TypeOf lVar Is MSForms.Page Then
            If lVar.Parent Is lParentForm Then
                Set lFrame = lVar
                Exit For
            End If
        End If
    Next
    End If
    lCaption = lParentForm.Caption

    If lFrame Is Nothing Then
        If val(Application.Version) < 9 Then
            ' Excel 97 or earlier
            lUserFormHwnd = FindWindow("ThunderXFrame", lCaption)
        Else
            ' Excel 2000 or later
            lUserFormHwnd = FindWindow("ThunderDFrame", lCaption)
        End If
    Else
        lUserFormHwnd = lFrame.[_GethWnd]
        Do
            If GetParent(GetParent(lUserFormHwnd)) = 0 Then Exit Do
            lUserFormHwnd = GetParent(lUserFormHwnd)
        Loop
    End If
    ' Si zone cliente
    If pClientArea Then
        lUserFormHwnd = FindClient(lUserFormHwnd, lParentForm)
    End If
    GetUserFormHandle = lUserFormHwnd
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    Select Case Err.Number
    Case -2147418105    ' Formulaire deconnecte
    Case Else
        MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procedure GetUserFormHandle du module ClGdiplus"
    End Select
End Function

Private Function FindClient(pHwndParent, pFrame As Object)
    Dim lCurrenthWnd
    Dim lClientHwnd
    Dim lret As Long
    Dim lClassName As String
    Dim lRectForm As Rect
    Dim lrect As Rect
    ' Habituellement, fenetre de classe "F3 Server 60000000"
    If Not TypeOf pFrame Is Frame Then
        lClientHwnd = FindWindowEx(pHwndParent, 0, "F3 Server 60000000", vbNullString)
    End If
    ' Le numero derriere "F3 Serveur" peut varier
    If lClientHwnd = 0 Then
        lCurrenthWnd = GetWindow(pHwndParent, GW_CHILD)
        Do Until lCurrenthWnd = 0
            lClassName = Space(255)
            lret = GetClassName(lCurrenthWnd, lClassName, 255)
            lClassName = Left(lClassName, lret)
            If lClassName Like "F3 Server*" Then
                If Not TypeOf pFrame Is Frame Then
                    lClientHwnd = lCurrenthWnd
                    Exit Do
                Else
                    GetWindowRect pHwndParent, lRectForm
                    GetWindowRect lCurrenthWnd, lrect
                    If Abs(PixelsToPointsX(lrect.Left - lRectForm.Left - 1) - pFrame.Left - 1) <= 2 And _
                            Abs(PixelsToPointsY(lrect.Top - lRectForm.Top - 1) - pFrame.Top - 1) <= 2 And _
                            Abs(PixelsToPointsY(lrect.Right - lRectForm.Left - 1) - (pFrame.Left + pFrame.Width + 1)) <= 2 And _
                            Abs(PixelsToPointsY(lrect.Bottom - lRectForm.Top - 1) - (pFrame.Top + pFrame.Height + 1)) <= 2 Then
                        lClientHwnd = lCurrenthWnd
                        Exit Do
                    Else
                    
                        lClientHwnd = FindClient(lCurrenthWnd, pFrame)
                        If lClientHwnd <> 0 Then
                            Exit Do
                        End If
                        '
                    End If
                End If
            End If
            lCurrenthWnd = GetWindow(lCurrenthWnd, GW_HWNDNEXT)
        Loop
    End If
    FindClient = lClientHwnd
End Function
#End If

' http://support.microsoft.com/kb/183544/fr
Private Sub StringFromPointer(pOLESTR, strOut As String)
    Dim ByteArray(255) As Byte
    Dim intTemp As Integer
    Dim intCount As Integer
    Dim i As Integer

    intTemp = 1

    'Walk the string and retrieve the first byte of each WORD.
    While intTemp <> 0
        RtlMoveMemory intTemp, ByVal pOLESTR + i, 2
        ByteArray(intCount) = intTemp
        intCount = intCount + 1
        i = i + 2
    Wend

    'Copy the byte array to our string.
    RtlMoveMemory ByVal strOut, ByteArray(0), intCount
End Sub

'---------------------------------------------------------------------------------------
' Calcul le decalage entre la gauche de la fenetre du formulaire et la gauche
'   de l'interieur du formulaire
'---------------------------------------------------------------------------------------
' pForm     : Le formulaire
' pSelector : Decaler egalement de la largeur du selecteur (uniquement access)
' Le decalage correspond a la taille de la bordure + eventuellement le selecteur
'---------------------------------------------------------------------------------------
#If AppName = "A" Then
Private Function FormDecalageX(pForm As Access.Form, Optional pSelector As Boolean = True) As Long
#Else
Private Function FormDecalageX(pForm As Object, Optional pSelector As Boolean = True) As Long
#End If
Dim lWrc As Rect
Dim lCrc As Rect
Dim lFormHwnd
On Error GoTo gestion_erreurs
#If AppName = "A" Then
    lFormHwnd = pForm.hwnd
#Else
    lFormHwnd = GetUserFormHandle(pForm)
#End If
' Coordonnees de la fenetre complete
GetWindowRect lFormHwnd, lWrc
' Taille de l'interieur de la fenetre
GetClientRect lFormHwnd, lCrc
' Calcul du decalage (=taille bordure)
FormDecalageX = (lWrc.Right - lWrc.Left - lCrc.Right) / 2
#If AppName = "A" Then
    ' On decale de la largeur du selecteur
    If pSelector Then FormDecalageX = FormDecalageX + PointsToPixelsX(pForm.CurrentSectionLeft)
    'FormDecalageX = FormDecalageX - 1
#End If
On Error GoTo 0
Exit Function
gestion_erreurs:
FormDecalageX = 0    ' Renvoie 0 s'il y a une erreur
End Function

'---------------------------------------------------------------------------------------
' Calcul le decalage entre le haut de la fenetre du formulaire et le haut
'   de l'interieur du formulaire
'---------------------------------------------------------------------------------------
' pForm    : Le formulaire
' Le decalage correspond a la taille de la bordure + la taille de la barre de titre
'---------------------------------------------------------------------------------------
#If AppName = "A" Then
Private Function FormDecalageY(pForm As Access.Form) As Long
#Else
Private Function FormDecalageY(pForm As Object) As Long
#End If
Dim lWrc As Rect
Dim lCrc As Rect
Dim lborder As Long
Dim lFormHwnd
On Error GoTo gestion_erreurs
#If AppName = "A" Then
    lFormHwnd = pForm.hwnd
#Else
    lFormHwnd = GetUserFormHandle(pForm)
#End If
' Coordonnees de la fenetre complete
GetWindowRect lFormHwnd, lWrc
' Taille de l'interieur de la fenetre
GetClientRect lFormHwnd, lCrc
' Calcul de la taille de bordure
lborder = (lWrc.Right - lWrc.Left - lCrc.Right) / 2
' Calcul du decalage complet
FormDecalageY = lWrc.Bottom - lWrc.Top - lborder - lCrc.Bottom
#If AppName = "A" Then
    'FormDecalageY = FormDecalageY - 1
#End If
On Error GoTo 0
Exit Function
gestion_erreurs:
FormDecalageY = 0    ' Renvoie 0 s'il y a une erreur
End Function
'***************************************************************************************
'*                                  TRANSFORMATIONS                                    *
'***************************************************************************************
Public Function WorldGetMatrix() As Variant
Dim lgraphics
Dim lMatrix
Dim lOutMatrix(1 To 6) As Single
lgraphics = GetGraphics()
  ' Sauvegarde les transformations
GdipCreateMatrix lMatrix
GdipGetWorldTransform lgraphics, lMatrix
GdipGetMatrixElements lMatrix, lOutMatrix(1)
GdipDeleteMatrix lMatrix
WorldGetMatrix = lOutMatrix
End Function
Public Function WorldSetMatrix(pMatrix As Variant) As Boolean
Dim lMatrix
Dim lgraphics
On Error GoTo gestion_erreurs
lMatrix = GetMatrix(pMatrix)
lgraphics = GetGraphics()
WorldSetMatrix = (GdipSetWorldTransform(lgraphics, lMatrix)) = 0
GdipDeleteMatrix lMatrix
PrivWorldTrandformGdi
Exit Function
gestion_erreurs:
WorldSetMatrix = False
End Function
Public Sub WorldRotate(ByVal pAngle As Single, Optional pRight As Boolean = True)
Dim lgraphics
    lgraphics = GetGraphics()
    GdipRotateWorldTransform lgraphics, pAngle, -pRight
    PrivWorldTrandformGdi
End Sub
Public Sub WorldTranslate(Optional ByVal pdX As Single = 0, Optional ByVal pdY As Single = 0, Optional pRight As Boolean = True)
Dim lgraphics
    lgraphics = GetGraphics()
    GdipTranslateWorldTransform lgraphics, pdX, pdY, -pRight
    PrivWorldTrandformGdi
End Sub
Public Sub WorldScale(Optional ByVal psX As Single = 1, Optional ByVal psY As Single = 1, Optional pRight As Boolean = True)
Dim lgraphics
    lgraphics = GetGraphics()
    GdipScaleWorldTransform lgraphics, psX, psY, -pRight
    PrivWorldTrandformGdi
End Sub
Public Sub WorldView(ByVal pSrcX1 As Single, ByVal pSrcY1 As Single, ByVal pSrcX2 As Single, ByVal pSrcY2 As Single, _
                Optional ByVal pDstX1 As Single, Optional ByVal pDstY1 As Single, Optional ByVal pDstX2 As Single, Optional ByVal pDstY2 As Single)
Dim lgraphics
If pSrcX2 = pSrcX1 Or pSrcY2 = pSrcY1 Then Exit Sub
'If pDstX2 = 0 Then pDstX2 = ImageWidth - 1
'If pDstY2 = 0 Then pDstY2 = ImageHeight - 1
    lgraphics = GetGraphics()
    GdipTranslateWorldTransform lgraphics, -pSrcX1, -pSrcY1, 1
    GdipScaleWorldTransform lgraphics, (pDstX2 - pDstX1) / (pSrcX2 - pSrcX1), _
                                    (pDstY2 - pDstY1) / (pSrcY2 - pSrcY1), 1
    'GdipTranslateWorldTransform lGraphics, pSrcX1, pSrcY1, 1
    GdipTranslateWorldTransform lgraphics, pDstX1, pDstY1, 1
    PrivWorldTrandformGdi
End Sub

Public Sub WorldReset()
Dim lgraphics
    lgraphics = GetGraphics()
      ' Desactive les transformations
    GdipResetWorldTransform lgraphics
    PrivWorldTrandformGdi
End Sub
Public Function WorldPush(Optional pMatrix As String = "default") As Boolean
Dim lgraphics
Dim lMatrix
Dim lMatrixTab(1 To 6) As Single
On Error GoTo gestion_erreurs
    lgraphics = GetGraphics()
      ' Sauvegarde les transformations
    GdipCreateMatrix lMatrix
    GdipGetWorldTransform lgraphics, lMatrix
    GdipGetMatrixElements lMatrix, lMatrixTab(1)
    GdipDeleteMatrix lMatrix
    gMatrix.Add lMatrixTab, pMatrix
    WorldPush = True
Exit Function
gestion_erreurs:
    ' Si la matrice existait deja on l'ecrase
    If Err.Number = 457 Then
        gMatrix.Remove pMatrix: Resume
    Else
        WorldPush = False
    End If
End Function
Public Function WorldPop(Optional pMatrix As String = "default") As Boolean
Dim lgraphics
Dim lMatrix
On Error GoTo gestion_erreurs
    lgraphics = GetGraphics()
    lMatrix = GetMatrix(pMatrix)
      ' Restaure les transformations
    GdipSetWorldTransform lgraphics, lMatrix
    GdipDeleteMatrix lMatrix
    PrivWorldTrandformGdi
    WorldPop = True
Exit Function
gestion_erreurs:
    WorldPop = False
End Function
#If UseSingle Then
Public Function WorldTransformPoint(pX As Single, pY As Single, Optional pMatrix As String = "") As Boolean
#Else
Public Function WorldTransformPoint(pX As Long, pY As Long, Optional pMatrix As String = "") As Boolean
#End If
On Error GoTo gestion_erreurs
    WorldTransformPoint = PrivWorldTransformPoint(pX, pY, pMatrix, False)
Exit Function
gestion_erreurs:
    WorldTransformPoint = False
End Function
#If UseSingle Then
Public Function WorldTransformVector(pX As Single, pY As Single, Optional pMatrix As String = "") As Boolean
#Else
Public Function WorldTransformVector(pX As Long, pY As Long, Optional pMatrix As String = "") As Boolean
#End If
On Error GoTo gestion_erreurs
    WorldTransformVector = PrivWorldTransformPoint(pX, pY, pMatrix, True)
Exit Function
gestion_erreurs:
    WorldTransformVector = False
End Function
#If UseSingle Then
Private Function PrivWorldTransformPoint(pX As Single, pY As Single, Optional pMatrix As String = "", Optional pVector As Boolean) As Boolean
Dim lPt As POINTF
#Else
Private Function PrivWorldTransformPoint(pX As Long, pY As Long, Optional pMatrix As String = "", Optional pVector As Boolean) As Boolean
Dim lPt As POINTAPI
#End If
Dim lgraphics
Dim lMatrix
On Error GoTo gestion_erreurs
    ' Recherche la matrice
    If Len(pMatrix) > 0 Then
        lMatrix = GetMatrix(pMatrix)
    Else
        lgraphics = GetGraphics()
        GdipCreateMatrix lMatrix
        GdipGetWorldTransform lgraphics, lMatrix
    End If
    ' Applique les transformations
    lPt.X = pX
    lPt.Y = pY
    If pVector Then
        PrivWorldTransformPoint = (GdipVectorTransformMatrixPoints(lMatrix, lPt, 1) = 0)
    Else
        PrivWorldTransformPoint = (GdipTransformMatrixPoints(lMatrix, lPt, 1) = 0)
    End If
    pX = lPt.X
    pY = lPt.Y
    GdipDeleteMatrix lMatrix
Exit Function
gestion_erreurs:
    PrivWorldTransformPoint = False
End Function
Private Sub PrivWorldTrandformGdi()
Dim lMatrixGdip
Dim lMatrixGdi32 As XForm
If UseGdi32 Then
    SetGraphicsMode hDC, GM_ADVANCED
    GdipCreateMatrix lMatrixGdip
    GdipGetWorldTransform gGraphics, lMatrixGdip
    GdipGetMatrixElements lMatrixGdip, lMatrixGdi32.eM11
    GdipDeleteMatrix lMatrixGdip
    SetWorldTransform hDC, lMatrixGdi32
End If
End Sub
'***************************************************************************************
'*                                  MATRICES
'***************************************************************************************
Private Function GetMatrixTab(pMatrix) As Variant
Dim lMatrixTab(1 To 6) As Single
GdipGetMatrixElements pMatrix, lMatrixTab(1)
GetMatrixTab = lMatrixTab
End Function
Private Function GetMatrix(pMatrix As Variant)
Dim lMatrix
Dim lMatrixTab As Variant
Dim lb As Long
' Recherche la matrice
If isArray(pMatrix) Then
    lMatrixTab = pMatrix
Else
    lMatrixTab = gMatrix(pMatrix)
End If
lb = LBound(lMatrixTab)
GdipCreateMatrix2 lMatrixTab(lb), lMatrixTab(lb + 1), lMatrixTab(lb + 2), lMatrixTab(lb + 3), lMatrixTab(lb + 4), lMatrixTab(lb + 5), lMatrix
GetMatrix = lMatrix
End Function
Private Sub UpdateMatrix(pName As String, pMatrix As Variant)
On Error GoTo gestion_erreurs
    gMatrix.Add pMatrix, pName
Exit Sub
gestion_erreurs:
If Err.Number = 457 Then
    gMatrix.Remove pName
    Resume
End If
End Sub
Public Function MatrixExists(pName As String) As Boolean
Dim lValue As Single
On Error GoTo gestion_erreurs
    lValue = gMatrix(pName)(1)
    MatrixExists = True
Exit Function
gestion_erreurs:
    MatrixExists = False
End Function
Public Function matrix(pName As String) As Variant
On Error GoTo gestion_erreurs
    matrix = gMatrix(pName)
Exit Function
gestion_erreurs:
    matrix = Null
End Function
Public Function MatrixPointsToPixels() As Variant
Dim lMatrix
lMatrix = PrivMatrixPointsToPixels
MatrixPointsToPixels = GetMatrixTab(lMatrix)
GdipDeleteMatrix lMatrix
End Function
Private Function PrivMatrixPointsToPixels()
Dim lHdc
lHdc = GetDC(PtrNull)
#If AppName = "A" Then
GdipCreateMatrix2 GetDeviceCaps(lHdc, LOGPIXELSX) / 1440, 0, 0, GetDeviceCaps(lHdc, LOGPIXELSY) / 1440, 0, 0, PrivMatrixPointsToPixels
#Else
GdipCreateMatrix2 GetDeviceCaps(lHdc, LOGPIXELSX) / 72, 0, 0, GetDeviceCaps(lHdc, LOGPIXELSY) / 72, 0, 0, PrivMatrixPointsToPixels
#End If
ReleaseDC PtrNull, lHdc
End Function
Public Function MatrixCmToPixels() As Variant
Dim lMatrix
On Error GoTo gestion_erreurs
    GdipCreateMatrix2 1 / 2.54 * DpiX, 0, 0, 1 / 2.54 * DpiY, 0, 0, lMatrix
    MatrixCmToPixels = GetMatrixTab(lMatrix)
    GdipDeleteMatrix lMatrix
Exit Function
gestion_erreurs:
    MatrixCmToPixels = Null
End Function
Public Function MatrixPixelsToCm() As Variant
Dim lMatrix
On Error GoTo gestion_erreurs
    GdipCreateMatrix2 2.54 / DpiX, 0, 0, 2.54 / DpiY, 0, 0, lMatrix
    MatrixPixelsToCm = GetMatrixTab(lMatrix)
    GdipDeleteMatrix lMatrix
Exit Function
gestion_erreurs:
    MatrixPixelsToCm = Null
End Function
Public Function MatrixPixelsToPoints() As Variant
Dim lMatrix
lMatrix = PrivMatrixPixelsToPoints
MatrixPixelsToPoints = GetMatrixTab(lMatrix)
GdipDeleteMatrix lMatrix
End Function
Private Function PrivMatrixPixelsToPoints()
Dim lHdc
lHdc = GetDC(PtrNull)
#If AppName = "A" Then
GdipCreateMatrix2 1440 / GetDeviceCaps(lHdc, LOGPIXELSX), 0, 0, 1440 / GetDeviceCaps(lHdc, LOGPIXELSY), 0, 0, PrivMatrixPixelsToPoints
#Else
GdipCreateMatrix2 72 / GetDeviceCaps(lHdc, LOGPIXELSX), 0, 0, 72 / GetDeviceCaps(lHdc, LOGPIXELSY), 0, 0, PrivMatrixPixelsToPoints
#End If
ReleaseDC PtrNull, lHdc
End Function
Public Function MatrixCtrlToImg(pObject As Object) As Variant
Dim lMatrix
lMatrix = PrivMatrixCtrlToImg(pObject)
MatrixCtrlToImg = GetMatrixTab(lMatrix)
GdipDeleteMatrix lMatrix
End Function
Public Function MatrixImgToCtrl(pObject As Object) As Variant
Dim lMatrix
lMatrix = PrivMatrixImgToCtrl(pObject)
MatrixImgToCtrl = GetMatrixTab(lMatrix)
GdipDeleteMatrix lMatrix
End Function
Private Function PrivMatrixCtrlToImg(pObject As Object)
    Dim lWidth As Single
    Dim lHeight As Single
#If UseSingle Then
    Dim lx1 As Single, lx2 As Single, ly1 As Single, ly2 As Single
    Dim lDx As Single, lDy As Single
    Dim lDecalageX As Single, lDecalageY As Single
#Else
    Dim lx1 As Long, lx2 As Long, ly1 As Long, ly2 As Long
    Dim lDx As Long, lDy As Long
    Dim lDecalageX As Long, lDecalageY As Long
#End If
    Dim lMatrix, lMatrix2
    On Error GoTo Gestion_Erreur
    GdipCreateMatrix lMatrix
    ' Dimensions de l'image
        ' Dimensions de l'image
    GdipGetImageDimension gBitmap, lWidth, lHeight
    ' Renvoie 0 si l'objet n'est pas initialise
    If pObject Is Nothing Then PrivMatrixCtrlToImg = PtrNull: Exit Function
    ' Recalcule la position de l'image dans l'objet
    CalcXY lx1, ly1, lx2, ly2, 0, 0, lDecalageX, lDecalageY, PictureSizeMode(pObject), PictureAlignment(pObject), pObject, lDx, lDy
    GdipCreateMatrix lMatrix
    Select Case PictureSizeMode(pObject)
            ' Mode echelle
        Case GdipSizeModeStretch
            ' On fait une regle de trois entre la taille de l'image et celle du controle
            lMatrix2 = PrivMatrixPointsToPixels
            GdipMultiplyMatrix lMatrix, lMatrix2, 1
            GdipDeleteMatrix lMatrix2
            GdipTranslateMatrix lMatrix, -lDx, -lDy, 1
            GdipScaleMatrix lMatrix, lWidth / (lx2 - lx1 + 1), lHeight / (ly2 - ly1 + 1), 1
            ' Mode decoupage
        Case GdipSizeModeClip
            ' On opere un decalage au cas ou l'image ne soit pas positionnee tout en haut du controle
            lMatrix2 = PrivMatrixPointsToPixels
            GdipMultiplyMatrix lMatrix, lMatrix2, 1
            GdipDeleteMatrix lMatrix2
            GdipTranslateMatrix lMatrix, -lDx - lDecalageX, -lDy - lDecalageY, 1
            ' Mode zoom
        Case GdipSizeModeZoom
            ' On opere un decalage plus une regle de trois entre la taille de l'image et celle de l'image visible
            lMatrix2 = PrivMatrixPointsToPixels
            GdipMultiplyMatrix lMatrix, lMatrix2, 1
            GdipDeleteMatrix lMatrix2
            GdipTranslateMatrix lMatrix, -lDx - lDecalageX, -lDy - lDecalageY, 1
            GdipScaleMatrix lMatrix, lWidth / (lx2 - lx1 + 1), lHeight / (ly2 - ly1 + 1), 1
    End Select
    PrivMatrixCtrlToImg = lMatrix
    Exit Function
Gestion_Erreur:
        PrivMatrixCtrlToImg = PtrNull
End Function
Private Function PrivMatrixImgToCtrl(pObject As Object)
Dim lMatrix
On Error GoTo Gestion_Erreur
    lMatrix = PrivMatrixCtrlToImg(pObject)
    GdipInvertMatrix lMatrix
    PrivMatrixImgToCtrl = lMatrix
    Exit Function
Gestion_Erreur:
        PrivMatrixImgToCtrl = PtrNull
End Function
Public Function MatrixNew(Optional pMatrix As String, Optional pElements As Variant) As Variant
Dim lb As Long, lcpt As Long
Dim lMatrixTab(1 To 6) As Single
On Error GoTo gestion_erreurs
If Not IsMissing(pElements) Then
    lb = LBound(pElements)
    For lcpt = 1 To 6
        lMatrixTab(lcpt) = pElements(lcpt - 1 + lb)
    Next
Else
   lMatrixTab(1) = 1
   lMatrixTab(4) = 1
End If
If Len(pMatrix) > 0 Then gMatrix.Add lMatrixTab, pMatrix
MatrixNew = lMatrixTab
Exit Function
gestion_erreurs:
    If Err.Number = 457 Then
        gMatrix.Remove pMatrix: Resume
    Else
        MatrixNew = Null
    End If
End Function

Public Function MatrixDelete(pMatrix As String) As Boolean
On Error GoTo gestion_erreurs
gMatrix.Remove pMatrix
MatrixDelete = True
Exit Function
gestion_erreurs:
    MatrixDelete = False
End Function

Public Function MatrixTranslate(pMatrix As Variant, Optional ByVal pdX As Single = 0, Optional ByVal pdY As Single = 0, Optional pRight As Boolean = True) As Variant
Dim lMatrix
On Error GoTo gestion_erreurs
lMatrix = GetMatrix(pMatrix)
MatrixTranslate = (GdipTranslateMatrix(lMatrix, pdX, pdY, -pRight) = 0)
If MatrixTranslate Then
    MatrixTranslate = GetMatrixTab(lMatrix)
    If Not isArray(pMatrix) Then UpdateMatrix CStr(pMatrix), MatrixTranslate
End If
gestion_erreurs:
    If Err.Number <> 0 Then MatrixTranslate = Null
    If lMatrix <> 0 Then GdipDeleteMatrix lMatrix
End Function
Public Function MatrixScale(pMatrix As Variant, Optional ByVal psX As Single = 1, Optional ByVal psY As Single = 1, Optional pRight As Boolean = True) As Variant
Dim lMatrix
On Error GoTo gestion_erreurs
lMatrix = GetMatrix(pMatrix)
MatrixScale = (GdipScaleMatrix(lMatrix, psX, psY, -pRight) = 0)
If MatrixScale Then
    MatrixScale = GetMatrixTab(lMatrix)
    If Not isArray(pMatrix) Then UpdateMatrix CStr(pMatrix), MatrixScale
End If
gestion_erreurs:
    If Err.Number <> 0 Then MatrixScale = Null
    If lMatrix <> 0 Then GdipDeleteMatrix lMatrix
End Function
Public Function MatrixRotate(pMatrix As Variant, ByVal pAngle As Single, Optional pRight As Boolean = True) As Variant
Dim lMatrix
On Error GoTo gestion_erreurs
lMatrix = GetMatrix(pMatrix)
MatrixRotate = (GdipRotateMatrix(lMatrix, pAngle, -pRight) = 0)
If MatrixRotate Then
    MatrixRotate = GetMatrixTab(lMatrix)
    If Not isArray(pMatrix) Then UpdateMatrix CStr(pMatrix), MatrixRotate
End If
gestion_erreurs:
    If Err.Number <> 0 Then MatrixRotate = Null
    If lMatrix <> 0 Then GdipDeleteMatrix lMatrix
End Function
Public Function MatrixMultiply(pMatrix1 As Variant, pMatrix2 As Variant, Optional pRight As Boolean = True) As Variant
Dim lMatrix1, lMatrix2, lMatrix
On Error GoTo gestion_erreurs
lMatrix1 = GetMatrix(pMatrix1)
lMatrix2 = GetMatrix(pMatrix2)
GdipCloneMatrix lMatrix1, lMatrix
MatrixMultiply = (GdipMultiplyMatrix(lMatrix, lMatrix2, -pRight) = 0)
If MatrixMultiply Then
    MatrixMultiply = GetMatrixTab(lMatrix)
    If Not isArray(pMatrix1) Then UpdateMatrix CStr(pMatrix1), MatrixMultiply
End If
gestion_erreurs:
    If Err.Number <> 0 Then MatrixMultiply = Null
    If lMatrix <> 0 Then GdipDeleteMatrix lMatrix
End Function

Public Function MatrixInvert(pMatrix As Variant) As Variant
Dim lMatrix
On Error GoTo gestion_erreurs
lMatrix = GetMatrix(pMatrix)
MatrixInvert = (GdipInvertMatrix(lMatrix) = 0)
If MatrixInvert Then
    MatrixInvert = GetMatrixTab(lMatrix)
    If Not isArray(pMatrix) Then UpdateMatrix CStr(pMatrix), MatrixInvert
End If
gestion_erreurs:
    If Err.Number <> 0 Then MatrixInvert = Null
    If lMatrix <> 0 Then GdipDeleteMatrix lMatrix
End Function
#If UseSingle Then
Public Function MatrixTransformPoint(pX As Single, pY As Single, ParamArray pMatrix() As Variant) As Boolean
#Else
Public Function MatrixTransformPoint(pX As Long, pY As Long, ParamArray pMatrix() As Variant) As Boolean
#End If
Dim lcpt As Long
On Error GoTo gestion_erreurs
For lcpt = LBound(pMatrix) To UBound(pMatrix)
    MatrixTransformPoint = PrivMatrixTransformPoint(pX, pY, GetMatrix(pMatrix(lcpt)), False, True)
Next
Exit Function
gestion_erreurs:
    MatrixTransformPoint = False
End Function
#If UseSingle Then
Public Function MatrixTransformVector(pX As Single, pY As Single, ParamArray pMatrix() As Variant) As Boolean
#Else
Public Function MatrixTransformVector(pX As Long, pY As Long, ParamArray pMatrix() As Variant) As Boolean
#End If
Dim lcpt As Long
On Error GoTo gestion_erreurs
For lcpt = LBound(pMatrix) To UBound(pMatrix)
    MatrixTransformVector = PrivMatrixTransformPoint(pX, pY, GetMatrix(pMatrix(lcpt)), True, True)
Next
Exit Function
gestion_erreurs:
    MatrixTransformVector = False
End Function
Public Function MatrixTransformPointSingle(pX As Single, pY As Single, ParamArray pMatrix() As Variant) As Boolean
Dim lcpt As Long
On Error GoTo gestion_erreurs
For lcpt = LBound(pMatrix) To UBound(pMatrix)
    MatrixTransformPointSingle = PrivMatrixTransformPointSingle(pX, pY, GetMatrix(pMatrix(lcpt)), False, True)
Next
Exit Function
gestion_erreurs:
    MatrixTransformPointSingle = False
End Function
Public Function MatrixTransformVectorSingle(pX As Single, pY As Single, ParamArray pMatrix() As Variant) As Boolean
Dim lcpt As Long
On Error GoTo gestion_erreurs
For lcpt = LBound(pMatrix) To UBound(pMatrix)
    MatrixTransformVectorSingle = PrivMatrixTransformPointSingle(pX, pY, GetMatrix(pMatrix(lcpt)), True, True)
Next
Exit Function
gestion_erreurs:
    MatrixTransformVectorSingle = False
End Function
#If UseSingle Then
Private Function PrivMatrixTransformPoint(pX As Single, pY As Single, pMatrix, Optional pVector As Boolean, Optional pDeleteMatrix As Boolean = False) As Boolean
Dim lPt As POINTF
#Else
Private Function PrivMatrixTransformPoint(pX As Long, pY As Long, pMatrix, Optional pVector As Boolean, Optional pDeleteMatrix As Boolean = False) As Boolean
Dim lPt As POINTAPI
#End If
On Error GoTo gestion_erreurs
    lPt.X = pX
    lPt.Y = pY
    If pVector Then
        PrivMatrixTransformPoint = (GdipVectorTransformMatrixPoints(pMatrix, lPt, 1) = 0)
    Else
        PrivMatrixTransformPoint = (GdipTransformMatrixPoints(pMatrix, lPt, 1) = 0)
    End If
    pX = lPt.X
    pY = lPt.Y
gestion_erreurs:
    If Err.Number <> 0 Then PrivMatrixTransformPoint = False
    If pDeleteMatrix Then GdipDeleteMatrix pMatrix
End Function
Private Function PrivMatrixTransformPointSingle(pX As Single, pY As Single, pMatrix, Optional pVector As Boolean, Optional pDeleteMatrix As Boolean = False) As Boolean
Dim lPt As POINTF
On Error GoTo gestion_erreurs
    lPt.X = pX
    lPt.Y = pY
    If pVector Then
        PrivMatrixTransformPointSingle = (GdipVectorTransformMatrixPointsS(pMatrix, lPt, 1) = 0)
    Else
        PrivMatrixTransformPointSingle = (GdipTransformMatrixPointsS(pMatrix, lPt, 1) = 0)
    End If
    pX = lPt.X
    pY = lPt.Y
gestion_erreurs:
    If Err.Number <> 0 Then PrivMatrixTransformPointSingle = False
    If pDeleteMatrix Then GdipDeleteMatrix pMatrix
End Function

'***************************************************************************************
'*              Divers
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Procedure de lecture de temps écoulé
'http://support.microsoft.com/kb/172338/fr
'---------------------------------------------------------------------------------------
Private Sub TimeInit()
Dim lCtr1 As Currency, lCtr2 As Currency
    QueryPerformanceFrequency gFreq
    QueryPerformanceCounter lCtr1
    QueryPerformanceCounter lCtr2
    gOverhead = lCtr2 - lCtr1
End Sub
Public Sub TimeStart()
    TimeInit
    QueryPerformanceCounter gTimerStart
End Sub
Public Function TimeElapsedMs(Optional pSinceLastTime As Boolean = True) As Single
Dim lCtr2 As Currency
If pSinceLastTime Then
    If gTimer = 0 Then
        TimeInit
        QueryPerformanceCounter gTimer
    Else
        QueryPerformanceCounter lCtr2
        TimeElapsedMs = (lCtr2 - gTimer - gOverhead) / gFreq * 1000
        gTimer = lCtr2
  End If
Else
    If gTimerStart = 0 Then
        TimeStart
    Else
        QueryPerformanceCounter lCtr2
        TimeElapsedMs = (lCtr2 - gTimerStart - gOverhead) / gFreq * 1000
  End If
End If
End Function
'---------------------------------------------------------------------------------------
' Procedure d'attente
'---------------------------------------------------------------------------------------
Public Sub Wait(ByVal pIntervalMs As Long, Optional pDoEvents As Boolean = True)
    Dim lret As Long
    Dim ls As SECURITY_ATTRIBUTES
    ' Cree un evenement
    If ghEvent = PtrNull Then
#If VBA64 Then
        ls.nLength = LenB(ls)
#Else
        ls.nLength = Len(ls)
#End If
        ls.lpSecurityDescriptor = PtrNull
        ls.bInheritHandle = 0
        ghEvent = CreateEvent(ls, False, False, CStr(ObjPtr(Me)))
    End If
    ' Arrêt si intervalle nul
    If pIntervalMs = 0 Then
        If gId <> 0 Then timeKillEvent gId: gId = 0
        If ghEvent <> PtrNull Then ResetEvent ghEvent
        If ghEvent <> PtrNull Then CloseHandle ghEvent: ghEvent = PtrNull
        gInterval = 0
        Exit Sub
    End If
    ' On s'assure que l'intervalle est positif
    If pIntervalMs <= 1 Then pIntervalMs = 1
    ' Si non lance ou si changement d'intervalle => on (re)lance le timer
    If pIntervalMs <> gInterval Then
        gInterval = pIntervalMs
        If gId <> 0 Then timeKillEvent gId
        gId = timeSetEvent(pIntervalMs, 0, ghEvent, 0, TIME_PERIODIC Or TIME_CALLBACK_EVENT_SET)
    End If
    ' Boucle tant que le timer n'a pas declenche d'evenement
    Do
        ' Cette fonction retourne si :
        '- le timer a ete declenche => lRet = WAIT_OBJECT_0
        '- n'importe quel autre message a ete envoye a l'application => on traite ces message avec un DoEvents
        lret = MsgWaitForMultipleObjects(1, ghEvent, False, _
                                         INFINITE, IIf(pDoEvents, QS_ALLINPUT&, 0))
        ' Traite les messages
        If pDoEvents Then DoEvents
    Loop Until lret = WAIT_OBJECT_0
End Sub

Private Function PtrToObj(ByVal lPtr) As Object
Dim oUnk As Object
 
   RtlMoveMemory oUnk, lPtr, LenB(lPtr)
   Set PtrToObj = oUnk
   RtlMoveMemory oUnk, 0&, LenB(lPtr)
             
End Function

Private Function GetGraphics(Optional pDoNotConvert As Boolean)
Dim lgraphics
On Error GoTo gestion_erreurs
    If gGraphics = 0 Then
        If Not pDoNotConvert Then
            If Not PrivGdipGetImageGraphicsContext(gBitmap, lgraphics) = 0 Then
                ConvertBppGraphics gBitmap
                GdipGetImageGraphicsContext gBitmap, lgraphics
            End If
        Else
            PrivGdipGetImageGraphicsContext gBitmap, lgraphics
        End If
        gGraphics = lgraphics
        If gGraphics <> 0 Then
            UpdateGraphics gGraphics
        End If
    Else
        lgraphics = gGraphics
    End If
    GetGraphics = lgraphics
Exit Function
gestion_erreurs:
    GetGraphics = 0
End Function

Private Sub UpdateGraphics(pGraphics)
    GdipSetCompositingMode pGraphics, gTransparencyMode
    GdipSetSmoothingMode pGraphics, gSmoothingMode
    GdipSetPixelOffsetMode pGraphics, gPixelOffsetMode
    GdipSetInterpolationMode pGraphics, gAntialisingLevel
    GdipSetTextRenderingHint pGraphics, gTextRenderingHint
    GdipSetCompositingQuality pGraphics, gCompositingQuality
    If gDrawClipRegion <> PtrNull Then
        GdipSetClipRegion gGraphics, gDrawClipRegion, 0 ' CombineModeReplace
        PrivSetClipRegionGdi
    End If
    If gMatrixForImageReset <> PtrNull Then
        GdipSetWorldTransform gGraphics, gMatrixForImageReset
        PrivWorldTrandformGdi
    End If
End Sub

'---------------------------------------------------------------------------------------
' Position d'un controle
'---------------------------------------------------------------------------------------
#If UseSingle Then
Public Function GetControlPos(pObject As Object, Optional pX1 As Single, Optional pY1 As Single, Optional pX2 As Single, Optional pY2 As Single, Optional pIncludeBorders As Boolean) As Single
Dim lBorderSize As Single
Dim lDecalage As Single
Dim lBorderStyle As Single
Dim lBorderWidth As Single
#Else
Public Function GetControlPos(pObject As Object, Optional pX1 As Long, Optional pY1 As Long, Optional pX2 As Long, Optional pY2 As Long, Optional pIncludeBorders As Boolean) As Long
Dim lBorderSize As Long
Dim lDecalage As Long
Dim lBorderStyle As Long
Dim lBorderWidth As Long
#End If
Dim lSpecialEffect As Long
#If AppName = "A" Then
If TypeOf pObject Is Access.Form Then
#Else
If TypeOf pObject Is MSForms.UserForm And Not TypeOf pObject Is MSForms.Frame And Not TypeOf pObject Is MSForms.MultiPage Then
#End If
    pX1 = 0
    pY1 = 0
    pX2 = PointsToPixelsX(pObject.InsideWidth) - 1
    pY2 = PointsToPixelsY(pObject.InsideHeight) - 1
    Exit Function
End If
On Error Resume Next
lSpecialEffect = pObject.SpecialEffect
lBorderStyle = pObject.BorderStyle
lBorderWidth = pObject.BorderWidth
On Error GoTo gestion_erreurs
#If AppName = "A" Then
If lSpecialEffect > 0 And lSpecialEffect <> 4 Then
    lBorderSize = GetSystemMetrics(SM_CXEDGE)
    lDecalage = 1
Else
    If lBorderStyle = 0 Then
        lBorderSize = 0
        lDecalage = 0
    Else
        Select Case lBorderWidth
            Case 0: lBorderSize = 1: lDecalage = 1
            Case 1: lBorderSize = 1: lDecalage = 1
            Case 2: lBorderSize = 3: lDecalage = 2
            Case 3: lBorderSize = 4: lDecalage = 2
            Case 4: lBorderSize = 5: lDecalage = 3
            Case 5: lBorderSize = 7: lDecalage = 4
            Case 6: lBorderSize = 8: lDecalage = 4
        End Select
    End If
End If
If pIncludeBorders Then
    lBorderSize = -lBorderSize
    lDecalage = lBorderSize + lDecalage
End If
#Else
If lSpecialEffect > 0 Then
    lBorderSize = GetSystemMetrics(SM_CXEDGE) * 2
    lDecalage = lBorderSize - 1
Else
    lBorderSize = -(lBorderStyle <> 0) * 2
    lDecalage = -(lBorderStyle <> 0) + 1
End If
If pIncludeBorders Then
    lBorderSize = -1 '-lBorderSize
    lDecalage = 0 'lBorderSize + lDecalage
End If
#End If
pX1 = PointsToPixelsX(pObject.Left) + lDecalage
pY1 = PointsToPixelsY(pObject.Top) + lDecalage
pX2 = PointsToPixelsX(pObject.Left + pObject.Width) - lBorderSize + lDecalage - 1
pY2 = PointsToPixelsY(pObject.Top + pObject.Height) - lBorderSize + lDecalage - 1
GetControlPos = lDecalage
Exit Function
gestion_erreurs:
    GetControlPos = 0
End Function

Private Function PictureSizeMode(pObject As Object) As Long
On Error GoTo gestion_erreurs
If pObject Is Nothing Then PictureSizeMode = GdipSizeModeClip: Exit Function
#If AppName = "A" Then
    PictureSizeMode = pObject.SizeMode
#Else
    PictureSizeMode = pObject.PictureSizeMode
#End If
Exit Function
gestion_erreurs:
    PictureSizeMode = GdipSizeModeZoom
End Function
Private Function PictureAlignment(pObject As Object) As Long
On Error GoTo gestion_erreurs
If pObject Is Nothing Then PictureAlignment = 2: Exit Function ' GdipAlignCenter
#If AppName = "A" Then
    PictureAlignment = pObject.PictureAlignment
#Else
    PictureAlignment = pObject.PictureAlignment
#End If
Exit Function
gestion_erreurs:
    PictureAlignment = 2 ' GdipAlignCenter
End Function

'***************************************************************************************
'*              Image principale
'***************************************************************************************
Public Property Let ImgObject(pValue As Object)
    Set oImg = Nothing
    Set oForm = Nothing
    On Error Resume Next
    #If AppName = "A" Then
    If TypeOf pValue Is Access.Form Then
    #Else
    If TypeOf pValue Is MSForms.Frame Then
    #End If
        Set oForm = pValue
    Else
        Set oImg = pValue
    End If
    On Error GoTo 0
    Set oObject = pValue
    RepaintFastResetCalc = True
End Property

#If AppName = "A" Then
Private Sub oForm_DblClick(ByVal Cancel As Integer)
#Else
Private Sub oForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
#End If
oImg_DblClick Cancel
End Sub

Private Sub oForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oImg_MouseDown Button, Shift, X, Y
End Sub

Private Sub oForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oImg_MouseMove Button, Shift, X, Y
End Sub

Private Sub oForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oImg_MouseUp Button, Shift, X, Y
End Sub
' Sur double-click
#If AppName = "A" Then
Private Sub oImg_DblClick(Cancel As Integer)
#Else
Private Sub oImg_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
#End If
    If Not goImgEventObject Is Nothing And gImgEventDblClick <> "" Then
        CallByName goImgEventObject, gImgEventDblClick, vbMethod, gImgLastClickRegion, gImgLastClickButton, gImgLastClickShift, gImgLastClickX, gImgLastClickY, Self
    Else
#If VBA6 Then
        RaiseEvent ImgMouseDblClick(gImgLastClickRegion, gImgLastClickButton, gImgLastClickShift, gImgLastClickX, gImgLastClickY)
#End If
    End If
    gImgLastClickRegion = ""
End Sub

' Sur click
#If AppName = "A" Then
Private Sub oImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub oImg_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    MatrixTransformPointSingle X, Y, MatrixCtrlToImg(oImg)
    lregion = GetRegionXY(X, Y)
    gImgLastClickRegion = lregion
    gImgLastClickX = X
    gImgLastClickY = Y
    gImgLastClickButton = Button
    gImgLastClickShift = Shift
    If Not goImgEventObject Is Nothing And gImgEventMouseDown <> "" Then
        CallByName goImgEventObject, gImgEventMouseDown, vbMethod, lregion, Button, Shift, X, Y, Self
    Else
#If VBA6 Then
        RaiseEvent ImgMouseDown(lregion, Button, Shift, X, Y)
#End If
    End If
End Sub

' Sur souris relachee
#If AppName = "A" Then
Private Sub oImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub oImg_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    MatrixTransformPointSingle X, Y, MatrixCtrlToImg(oImg)
    lregion = GetRegionXY(X, Y)
    If Not goImgEventObject Is Nothing And gImgEventMouseUp <> "" Then
        CallByName goImgEventObject, gImgEventMouseUp, vbMethod, lregion, Button, Shift, X, Y, Self
    Else
#If VBA6 Then
        RaiseEvent ImgMouseUp(lregion, Button, Shift, X, Y)
#End If
    End If
End Sub
' Sur souris deplacee
#If AppName = "A" Then
Private Sub oImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub oImg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    MatrixTransformPointSingle X, Y, MatrixCtrlToImg(oImg)
    lregion = GetRegionXY(X, Y)
    If Not goImgEventObject Is Nothing And gImgEventMouseMove <> "" Then
        CallByName goImgEventObject, gImgEventMouseMove, vbMethod, lregion, Button, Shift, X, Y, Self
    Else
#If VBA6 Then
        RaiseEvent ImgMouseMove(lregion, Button, Shift, X, Y)
#End If
    End If
End Sub

'***************************************************************************************
'*              ScrollBars
'***************************************************************************************
Private Property Get BarValue(pBarName As String, pValueName As String) As Variant
On Error GoTo gestion_erreurs
BarValue = goBars(pBarName).UserData(pValueName)
Exit Property
gestion_erreurs:
BarValue = Null
End Property
Private Property Let BarValue(Optional pBarName As String = "Default", Optional pValueName As String, value As Variant)
goBars(pBarName).UserData(pValueName) = value
End Property

Public Sub BarEvents(pObject As Object, Optional pEventOnRefreshNeeded As String, Optional pEventMouseDown As String, Optional pEventMouseUp As String, Optional pEventMouseMove As String, Optional pEventMouseDblClick As String)
Set goBarEventObject = pObject
gBarEventMouseDown = pEventMouseDown
gBarEventMouseUp = pEventMouseUp
gBarEventMouseMove = pEventMouseMove
gBarEventDblClick = pEventMouseDblClick
gBarEventOnRefreshNeeded = pEventOnRefreshNeeded
End Sub

Private Sub RaiseEventRefreshNeeded(pName As String, pMouseUp As Boolean)
    If Not goBarEventObject Is Nothing And gBarEventOnRefreshNeeded <> "" Then
        CallByName goBarEventObject, gBarEventOnRefreshNeeded, vbMethod, Self, pName, pMouseUp
    Else
#If VBA6 Then
        RaiseEvent BarOnRefreshNeeded(pName, pMouseUp)
#End If
    End If
End Sub

#If AppName = "A" Then
Private Sub oBarForm_DblClick(ByVal Cancel As Integer)
#Else
Private Sub oBarForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
#End If
oBarImage_DblClick Cancel
End Sub

Private Sub oBarForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oBarImage_MouseDown Button, Shift, X, Y
End Sub

Private Sub oBarForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oBarImage_MouseMove Button, Shift, X, Y
End Sub

Private Sub oBarForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
oBarImage_MouseUp Button, Shift, X, Y
End Sub

' Sur double-click
#If AppName = "A" Then
Private Sub oBarImage_DblClick(Cancel As Integer)
#Else
Private Sub oBarImage_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
#End If
If Not gBarLastClickRegion Like gBarKey & "Bar*" Then
        If Not goBarEventObject Is Nothing And gBarEventDblClick <> "" Then
            CallByName goBarEventObject, gBarEventDblClick, vbMethod, gBarLastClickBar, gBarLastClickRegion, gBarLastClickButton, gBarLastClickShift, gBarLastClickX, gBarLastClickY, Self
        Else
#If VBA6 Then
            RaiseEvent BarMouseDblClick(gBarLastClickBar, gBarLastClickRegion, gBarLastClickButton, gBarLastClickShift, gBarLastClickX, gBarLastClickY)
#End If
        End If
        gBarLastClickRegion = ""
    End If
End Sub

' Sur click
#If AppName = "A" Then
Private Sub oBarImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub oBarImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    Dim pName As String
    If Not BarTransformPoint(X, Y) Then Exit Sub
    lregion = BarGetRegion(X, Y, pName)
    gBarLastClickRegion = lregion
    gBarLastClickX = X
    gBarLastClickY = Y
    gBarLastClickButton = Button
    gBarLastClickShift = Shift
    If lregion <> "" Then
        gBarClickX = X
        gBarClickY = Y
        gBarDragRegion = lregion
        gBarDragBar = pName
        If lregion Like gBarKey & "BarY*" Then
            If lregion = gBarKey & "BarYUp" Then
                BarStartY(pName) = BarStartY(pName) + BarValue(pName, "BarElementHeight")
                RaiseEventRefreshNeeded pName, False
            ElseIf lregion = gBarKey & "BarYDown" Then
                BarStartY(pName) = BarStartY(pName) - BarValue(pName, "BarElementHeight")
                RaiseEventRefreshNeeded pName, False
            ElseIf lregion = gBarKey & "BarYFront" Then
            Else
                BarY(pName) = Y
            End If
        ElseIf lregion Like gBarKey & "BarX*" Then
            If lregion = gBarKey & "BarXUp" Then
                BarStartX(pName) = BarStartX(pName) - BarValue(pName, "BarElementWidth")
                RaiseEventRefreshNeeded pName, False
            ElseIf lregion = gBarKey & "BarXDown" Then
                BarStartX(pName) = BarStartX(pName) + BarValue(pName, "BarElementWidth")
                RaiseEventRefreshNeeded pName, False
            ElseIf lregion = gBarKey & "BarXFront" Then
            Else
                BarX(pName) = X
            End If
        End If
    End If
    If Not lregion Like gBarKey & "Bar*" Then
        If Not goBarEventObject Is Nothing And gBarEventMouseDown <> "" Then
            CallByName goBarEventObject, gBarEventMouseDown, vbMethod, pName, lregion, Button, Shift, X, Y, Self
        Else
#If VBA6 Then
            RaiseEvent BarMouseDown(pName, lregion, Button, Shift, X, Y)
#End If
        End If
    End If
End Sub

' Sur souris relachee
#If AppName = "A" Then
Private Sub oBarImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub oBarImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    Dim pName As String
    If Not BarTransformPoint(X, Y) Then Exit Sub
    lregion = BarGetRegion(X, Y, pName)
    If Not lregion Like gBarKey & "Bar*" Then
        If Not goBarEventObject Is Nothing And gBarEventMouseUp <> "" Then
            CallByName goBarEventObject, gBarEventMouseUp, vbMethod, pName, lregion, Button, Shift, X, Y, Self
        Else
#If VBA6 Then
            RaiseEvent BarMouseUp(pName, lregion, Button, Shift, X, Y)
#End If
        End If
    Else
        RaiseEventRefreshNeeded pName, True
    End If
    gBarDragRegion = ""
    gBarDragBar = ""
End Sub
' Sur souris deplacee
#If AppName = "A" Then
Private Sub oBarImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub oBarImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    Dim pName As String
    If Not BarTransformPoint(X, Y) Then Exit Sub
    lregion = BarGetRegion(X, Y, pName)
    If Button = vbKeyLButton Then
        If gBarDragBar <> "" Then
            If gBarDragRegion Like gBarKey & "Bar*" Then
                If gBarDragRegion Like gBarKey & "BarY*" Then
                    BarY(gBarDragBar, True) = Y
                    gBarClickY = Y
                ElseIf gBarDragRegion Like gBarKey & "BarX*" Then
                    BarX(gBarDragBar, True) = X
                    gBarClickX = X
                End If
            End If
        End If
    End If
    If Not lregion Like gBarKey & "Bar*" Then
        If Not goBarEventObject Is Nothing And gBarEventMouseMove <> "" Then
            CallByName goBarEventObject, gBarEventMouseMove, vbMethod, pName, lregion, Button, Shift, X, Y, Self
        Else
#If VBA6 Then
            RaiseEvent BarMouseMove(pName, lregion, Button, Shift, X, Y)
#End If
        End If
    End If
End Sub

Private Function BarGetRegion(ByVal pX As Single, ByVal pY As Single, pBarName As String) As String
    Dim lregion As String
    Dim loBar As clGdiplus
    Dim lBarName As String
    On Error GoTo gestion_erreurs
    For Each loBar In goBars
        lBarName = loBar.UserData("BarName")
        loBar.CreateRegionRect "bar", BarLeft(lBarName), BarTop(lBarName), BarRight(lBarName), BarBottom(lBarName)
        If loBar.PointInRegion(pX, pY, "bar") Then
            loBar.RegionDelete "bar"
            pBarName = lBarName
            lregion = loBar.GetRegionXY(pX, pY)
            If lregion <> "" Then
                BarGetRegion = lregion
                Exit For
            End If
        Else
            loBar.RegionDelete "bar"
        End If
    Next
    If BarGetRegion = "" Then BarGetRegion = GetRegionXY(pX, pY)
gestion_erreurs:
    If Err.Number <> 0 Then BarGetRegion = ""
End Function

Private Function BarTransformPoint(pX As Single, pY As Single) As Boolean
    BarTransformPoint = True
    If Not IsNull(RepaintMatrix) Then
        MatrixTransformPointSingle pX, pY, RepaintMatrix
    ElseIf Not oBarObject Is Nothing Then
        If ImgParentPtr <> 0 Then
            MatrixTransformPointSingle pX, pY, Parent.MatrixCtrlToImg(oBarObject)
        Else
            MatrixTransformPointSingle pX, pY, MatrixCtrlToImg(oBarObject)
        End If
    End If
    If Not IsNull(DrawMatrix) Then
        MatrixTransformPointSingle pX, pY, DrawMatrix
        If Not (pX >= DrawOrigX1 And pX <= DrawOrigX2 And pY >= DrawOrigY1 And pY <= DrawOrigY2) Then
            BarTransformPoint = False
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Objet pour barres de defilement
'---------------------------------------------------------------------------------------
Public Property Get BarObject() As Object
Set BarObject = oBarObject
End Property

Public Property Let BarObject(value As Object)
Const lMessage As String = "Definissez la propriete '[Property]' de l'image a :" & vbCrLf & "[Procedure evenementielle]"
On Error GoTo gestion_erreurs
Set oBarImage = Nothing
Set oBarForm = Nothing
On Error Resume Next
Set oBarImage = value
Set oBarForm = value
Set oBarObject = value
If value Is Nothing Then
    Set goBars = Nothing
    Exit Property
End If
#If AppName = "A" Then
If oBarObject.OnMouseDown = "" Then MsgBox Replace(lMessage, "[Property]", "Sur souris appuyee")
If oBarObject.OnMouseMove = "" Then MsgBox Replace(lMessage, "[Property]", "Sur souris deplacee")
If oBarObject.OnMouseUp = "" Then MsgBox Replace(lMessage, "[Property]", "Sur souris relachee")
If oBarObject.OnDblClick = "" Then MsgBox Replace(lMessage, "[Property]", "Sur double clic")
#End If
Exit Property
gestion_erreurs:
    '
End Property

'---------------------------------------------------------------------------------------
' Nouvelle barre de defilement
'---------------------------------------------------------------------------------------
Public Function BarNew(Optional pName As String = "Default") As Boolean
Dim lBarGdi As clGdiplus
On Error GoTo gestion_erreurs
Set lBarGdi = New clGdiplus
lBarGdi.CreateBitmap 1, 1 ' utile pour transformations
lBarGdi.UserData("BarName") = pName
lBarGdi.UserData("BarLeft") = 0
lBarGdi.UserData("BarTop") = 0
lBarGdi.UserData("BarRight") = ImageWidth - 1
lBarGdi.UserData("BarBottom") = ImageHeight - 1
lBarGdi.UserData("BarSize") = GetSystemMetrics(SM_CXVSCROLL)
If lBarGdi.UserData("BarSize") = 0 Then lBarGdi.UserData("BarSize") = 16
lBarGdi.UserData("BarWidth") = lBarGdi.UserData("BarSize")
lBarGdi.UserData("BarHeight") = lBarGdi.UserData("BarSize")
lBarGdi.UserData("BarY") = 1 + lBarGdi.UserData("BarSize")
lBarGdi.UserData("BarX") = 1 + lBarGdi.UserData("BarSize")
lBarGdi.UserData("BarKey") = CStr(ObjPtr(Me))
lBarGdi.UserData("BarBackColor") = RGB(230, 230, 230)
lBarGdi.UserData("UserBarVertVisible") = Null
lBarGdi.UserData("UserBarHorzVisible") = Null
lBarGdi.UserData("BarVertVisible") = False
lBarGdi.UserData("BarHorzVisible") = False
lBarGdi.UserData("BarStartX") = 0
lBarGdi.UserData("BarStartY") = 0
lBarGdi.UserData("BarMaxX") = 0
lBarGdi.UserData("BarMaxY") = 0
lBarGdi.UserData("BarElementWidth") = 0
lBarGdi.UserData("BarElementHeight") = 0
goBars.Add lBarGdi, pName
BarNew = True
Exit Function
gestion_erreurs:
If Err.Number = 457 Then
    BarDelete pName
    Resume
End If
BarNew = False
End Function

'---------------------------------------------------------------------------------------
' Supprime une barre
'---------------------------------------------------------------------------------------
Public Function BarDelete(Optional pName As String = "Default") As Boolean
On Error GoTo gestion_erreurs
goBars.Remove pName
BarDelete = True
Exit Function
gestion_erreurs:
BarDelete = False
End Function
' La barre est initialisee
Public Property Get BarExists(Optional pName As String = "Default") As Boolean
On Error GoTo gestion_erreurs
BarExists = Not goBars(pName) Is Nothing
Exit Property
gestion_erreurs:
BarExists = False
End Property
' Taille des barres
Public Property Get BarSize(Optional pName As String = "Default") As Long
BarSize = goBars(pName).UserData("BarSize")
End Property
Public Property Let BarSize(Optional pName As String = "Default", pSize As Long)
goBars(pName).UserData("BarSize") = pSize
BarReScaleX pName
BarReScaleY pName
End Property

' Positions
Public Property Get BarLeft(Optional pName As String = "Default") As Long
    BarLeft = goBars(pName).UserData("BarLeft")
End Property
Public Property Get BarTop(Optional pName As String = "Default") As Long
    BarTop = goBars(pName).UserData("BarTop")
End Property
Public Property Get BarRight(Optional pName As String = "Default") As Long
    BarRight = goBars(pName).UserData("BarRight")
End Property
Public Property Get BarBottom(Optional pName As String = "Default") As Long
    BarBottom = goBars(pName).UserData("BarBottom")
End Property

' Largeur interieure
Public Property Get BarInsideWidth(Optional pName As String = "Default") As Long
Dim loBar As clGdiplus
Set loBar = goBars(pName)
BarInsideWidth = loBar.UserData("BarRight") - loBar.UserData("BarLeft") + 1 + loBar.UserData("BarSize") * BarVertVisible(pName)
If BarInsideWidth = 0 Then BarInsideWidth = ImageWidth
End Property
' Hauteur interieure
Public Property Get BarInsideHeight(Optional pName As String = "Default") As Long
Dim loBar As clGdiplus
Set loBar = goBars(pName)
    BarInsideHeight = loBar.UserData("BarBottom") - loBar.UserData("BarTop") + 1 + loBar.UserData("BarSize") * BarHorzVisible(pName)
    If BarInsideHeight = 0 Then BarInsideHeight = ImageHeight
End Property

' Affiche ou masque la barre verticale
Public Property Let BarVertVisible(Optional pName As String = "Default", pVisible As Boolean)
    BarValue(pName, "UserBarVertVisible") = pVisible
    BarReScaleX pName
    BarReScaleY pName
    If Not pVisible And goBars(pName).UserData("BarStartY") <> 0 Then
        goBars(pName).UserData("BarStartY") = 0
        RaiseEventRefreshNeeded pName, False
    End If
End Property
Public Property Get BarVertVisible(Optional pName As String = "Default") As Boolean
If IsNull(BarValue(pName, "UserBarVertVisible")) Then
    BarVertVisible = goBars(pName).UserData("BarVertVisible")
Else
    BarVertVisible = goBars(pName).UserData("UserBarVertVisible")
End If
End Property

' Affiche ou masque la barre horizontale
Public Property Let BarHorzVisible(Optional pName As String = "Default", pVisible As Boolean)
    BarValue(pName, "UserBarHorzVisible") = pVisible
    BarReScaleY pName
    BarReScaleX pName
    If Not pVisible And goBars(pName).UserData("BarStartX") <> 0 Then
        goBars(pName).UserData("BarStartX") = 0
        RaiseEventRefreshNeeded pName, False
    End If
End Property
Public Property Get BarHorzVisible(Optional pName As String = "Default") As Boolean
If IsNull(BarValue(pName, "UserBarHorzVisible")) Then
    BarHorzVisible = goBars(pName).UserData("BarHorzVisible")
Else
    BarHorzVisible = goBars(pName).UserData("UserBarHorzVisible")
End If
End Property

' Echelle de la barre horizontale
Public Sub BarScaleX(pMax As Single, Optional pElementSize As Single = 1, Optional pLeft As Long = cLongOptional, Optional pRight As Long = cLongOptional, Optional pName As String = "Default")
If BarValue(pName, "BarLeft") = 0 And BarValue(pName, "BarTop") = 0 And BarValue(pName, "BarRight") = 0 And BarValue(pName, "BarBottom") = 0 Then
    BarValue(pName, "BarRight") = ImageWidth - 1
    BarValue(pName, "BarBottom") = ImageHeight - 1
End If
BarValue(pName, "BarMaxX") = pMax
BarValue(pName, "BarElementWidth") = pElementSize
BarValue(pName, "BarHorzVisible") = (BarValue(pName, "BarMaxX") <> 0 And BarValue(pName, "BarMaxX") > BarCountVisibleX(pName))
If pLeft <> cLongOptional Then BarValue(pName, "BarLeft") = pLeft Else BarValue(pName, "BarLeft") = 0
If pRight <> cLongOptional Then BarValue(pName, "BarRight") = pRight Else BarValue(pName, "BarRight") = ImageWidth - 1
BarReScaleY pName
BarReScaleX pName
End Sub

Private Sub BarReScaleX(Optional pName As String = "Default")
BarValue(pName, "BarBackWidth") = BarInsideWidth(pName) - 2 * BarValue(pName, "BarSize") - 2
If BarValue(pName, "BarElementWidth") <> 0 And BarValue(pName, "BarMaxX") <> 0 Then
    BarValue(pName, "BarWidth") = BarValue(pName, "BarBackWidth") / BarValue(pName, "BarMaxX") * (BarValue(pName, "BarBackWidth") / BarValue(pName, "BarElementWidth"))
Else
    BarValue(pName, "BarWidth") = 0
End If
If BarValue(pName, "BarMaxX") <> 0 And BarValue(pName, "BarMaxX") > BarCountVisibleX(pName) And BarValue(pName, "BarElementWidth") <> 0 Then
    BarValue(pName, "BarWidth") = BarValue(pName, "BarBackWidth") / BarValue(pName, "BarMaxX") * (BarValue(pName, "BarBackWidth") / BarValue(pName, "BarElementWidth"))
    If BarValue(pName, "BarWidth") > BarValue(pName, "BarBackWidth") Then BarValue(pName, "BarWidth") = BarValue(pName, "BarBackWidth")
    If BarValue(pName, "BarWidth") < BarValue(pName, "BarSize") Then BarValue(pName, "BarWidth") = BarValue(pName, "BarSize")
    BarValue(pName, "BarHorzVisible") = True
Else
    BarValue(pName, "BarHorzVisible") = False
End If
BarStartX(pName) = BarValue(pName, "BarStartX")
End Sub

' Echelle de la barre verticale
Public Sub BarScaleY(pMax As Single, Optional pElementSize As Single = 1, Optional pTop As Long = cLongOptional, Optional pBottom As Long = cLongOptional, Optional pName As String = "Default")
If BarValue(pName, "BarLeft") = 0 And BarValue(pName, "BarTop") = 0 And BarValue(pName, "BarRight") = 0 And BarValue(pName, "BarBottom") = 0 Then
    BarValue(pName, "BarRight") = ImageWidth - 1
    BarValue(pName, "BarBottom") = ImageHeight - 1
End If
BarValue(pName, "BarMaxY") = pMax
BarValue(pName, "BarElementHeight") = pElementSize
BarValue(pName, "BarVertVisible") = (BarValue(pName, "BarMaxY") <> 0 And BarValue(pName, "BarMaxY") > BarCountVisibleY(pName))
If pTop <> cLongOptional Then BarValue(pName, "BarTop") = pTop Else BarValue(pName, "BarTop") = 0
If pBottom <> cLongOptional Then BarValue(pName, "BarBottom") = pBottom Else BarValue(pName, "BarBottom") = ImageHeight - 1
BarReScaleX pName
BarReScaleY pName
End Sub

Private Sub BarReScaleY(Optional pName As String = "Default")
BarValue(pName, "BarBackHeight") = BarInsideHeight(pName) - 2 * BarValue(pName, "BarSize") - 2
If BarValue(pName, "BarElementHeight") <> 0 And BarValue(pName, "BarMaxY") <> 0 Then
    BarValue(pName, "BarHeight") = BarValue(pName, "BarBackHeight") / BarValue(pName, "BarMaxY") * (BarValue(pName, "BarBackHeight") / BarValue(pName, "BarElementHeight"))
Else
    BarValue(pName, "BarHeight") = 0
End If
If BarValue(pName, "BarMaxY") <> 0 And BarValue(pName, "BarMaxY") > BarCountVisibleY(pName) And BarValue(pName, "BarElementHeight") <> 0 Then
    BarValue(pName, "BarHeight") = BarValue(pName, "BarBackHeight") / BarValue(pName, "BarMaxY") * (BarValue(pName, "BarBackHeight") / BarValue(pName, "BarElementHeight"))
    If BarValue(pName, "BarHeight") > BarValue(pName, "BarBackHeight") Then BarValue(pName, "BarHeight") = BarValue(pName, "BarBackHeight")
    If BarValue(pName, "BarHeight") < BarValue(pName, "BarSize") Then BarValue(pName, "BarHeight") = BarValue(pName, "BarSize")
    BarValue(pName, "BarVertVisible") = True
Else
    BarValue(pName, "BarVertVisible") = False
End If
BarStartY(pName) = BarValue(pName, "BarStartY")
End Sub
' Premiere ligne visible
Public Property Get BarFirstVisibleRow(Optional pName As String = "Default") As Long
If BarValue(pName, "BarElementHeight") = 0 Then
    BarFirstVisibleRow = 0
Else
    BarFirstVisibleRow = -BarStartY / BarValue(pName, "BarElementHeight") + 1
End If
End Property
Public Property Let BarFirstVisibleRow(Optional pName As String = "Default", pRow As Long)
BarStartY(pName) = -(pRow - 1) * BarValue(pName, "BarElementHeight")
End Property
' Derniere ligne visible
Public Property Get BarLastVisibleRow(Optional pName As String = "Default") As Long
If BarValue(pName, "BarElementHeight") = 0 Then
    BarLastVisibleRow = 0
Else
    BarLastVisibleRow = BarFirstVisibleRow + BarInsideHeight(pName) \ BarValue(pName, "BarElementHeight")
End If
End Property
' Nombre d'elements maxi visibles
Public Property Get BarCountVisibleY(Optional pName As String = "Default") As Double
    If BarValue(pName, "BarElementHeight") = 0 Then
        BarCountVisibleY = 0
    Else
        BarCountVisibleY = BarInsideHeight(pName) / BarValue(pName, "BarElementHeight")
    End If
End Property
Public Property Get BarCountVisibleX(Optional pName As String = "Default") As Double
    If BarValue(pName, "BarElementWidth") = 0 Then
        BarCountVisibleX = 0
    Else
        BarCountVisibleX = BarInsideWidth(pName) / BarValue(pName, "BarElementWidth")
    End If
End Property
' Position du dessin en X (en pixels)
Public Property Get BarStartX(Optional pName As String = "Default") As Long
    BarStartX = BarValue(pName, "BarStartX")
End Property
Public Property Let BarStartX(Optional pName As String = "Default", pStartX As Long)
    If Not BarValue(pName, "BarHorzVisible") Or BarValue(pName, "BarBackWidth") = BarValue(pName, "BarWidth") Then BarValue(pName, "BarX") = BarValue(pName, "BarLeft") + BarValue(pName, "BarSize") + 1: BarValue(pName, "BarStartX") = 0: Exit Property
    If pStartX > 0 Then pStartX = 0
    If pStartX < -(BarValue(pName, "BarMaxX") - BarCountVisibleX(pName)) * BarValue(pName, "BarElementWidth") Then pStartX = -(BarValue(pName, "BarMaxX") - BarCountVisibleX(pName)) * BarValue(pName, "BarElementWidth")
    BarValue(pName, "BarStartX") = pStartX
    If BarValue(pName, "BarElementWidth") <> 0 And (BarValue(pName, "BarMaxX") - BarCountVisibleX(pName)) <> 0 Then
        BarValue(pName, "BarX") = BarValue(pName, "BarLeft") + BarValue(pName, "BarSize") + 1 - BarValue(pName, "BarStartX") / (BarValue(pName, "BarMaxX") - BarCountVisibleX(pName)) / BarValue(pName, "BarElementWidth") * (BarValue(pName, "BarBackWidth") - BarValue(pName, "BarWidth"))
    Else
        BarValue(pName, "BarX") = BarValue(pName, "BarLeft") + BarValue(pName, "BarSize") + 1
    End If
End Property
' Position du dessin en Y (en pixels)
Public Property Get BarStartY(Optional pName As String = "Default") As Long
    BarStartY = BarValue(pName, "BarStartY")
End Property
Public Property Let BarStartY(Optional pName As String = "Default", pStartY As Long)
    If Not BarValue(pName, "BarVertVisible") Or BarValue(pName, "BarBackHeight") = BarValue(pName, "BarHeight") Then BarValue(pName, "BarY") = BarValue(pName, "BarTop") + BarValue(pName, "BarSize") + 1: BarValue(pName, "BarStartY") = 0: Exit Property
    If pStartY > 0 Then pStartY = 0
    If pStartY < -(BarValue(pName, "BarMaxY") - BarCountVisibleY(pName)) * BarValue(pName, "BarElementHeight") Then pStartY = -(BarValue(pName, "BarMaxY") - BarCountVisibleY(pName)) * BarValue(pName, "BarElementHeight")
    BarValue(pName, "BarStartY") = pStartY
    If BarValue(pName, "BarElementHeight") <> 0 And (BarValue(pName, "BarMaxY") - BarCountVisibleY(pName)) <> 0 Then
        BarValue(pName, "BarY") = BarValue(pName, "BarTop") + BarValue(pName, "BarSize") + 1 - BarValue(pName, "BarStartY") / (BarValue(pName, "BarMaxY") - BarCountVisibleY(pName)) / BarValue(pName, "BarElementHeight") * (BarValue(pName, "BarBackHeight") - BarValue(pName, "BarHeight"))
    Else
        BarValue(pName, "BarY") = BarValue(pName, "BarTop") + BarValue(pName, "BarSize") + 1
    End If
End Property
' Position de la barre verticale (en twips/points)
Private Property Let BarY(Optional pName As String = "Default", Optional pRelative As Boolean, Y As Single)
    Dim lDelta As Long
    Dim lOldBarStartY As Long
    If BarValue(pName, "BarBackHeight") = BarValue(pName, "BarHeight") Then BarValue(pName, "BarStartY") = 0: Exit Property
    lDelta = BarValue(pName, "BarSize") + 1 + BarValue(pName, "BarTop")
    If pRelative Then
        BarValue(pName, "BarY") = (BarValue(pName, "BarY") - (gBarClickY - Y))
    Else
        BarValue(pName, "BarY") = Y - BarValue(pName, "BarHeight") / 2
    End If
    If BarValue(pName, "BarY") < lDelta Then BarValue(pName, "BarY") = lDelta
    If BarValue(pName, "BarY") > lDelta + BarValue(pName, "BarBackHeight") - BarValue(pName, "BarHeight") Then BarValue(pName, "BarY") = lDelta + BarValue(pName, "BarBackHeight") - BarValue(pName, "BarHeight")
    lOldBarStartY = BarValue(pName, "BarStartY")
    BarValue(pName, "BarStartY") = -(BarValue(pName, "BarY") - lDelta) / (BarValue(pName, "BarBackHeight") - BarValue(pName, "BarHeight")) * (BarValue(pName, "BarMaxY") - BarCountVisibleY(pName)) * BarValue(pName, "BarElementHeight")
    If lOldBarStartY <> BarValue(pName, "BarStartY") Then RaiseEventRefreshNeeded pName, False
End Property
' Position de la barre horizontale (en twips/points)
Private Property Let BarX(Optional pName As String = "Default", Optional pRelative As Boolean, X As Single)
    Dim lDelta As Long
    Dim lOldBarStartX As Long
    If BarValue(pName, "BarBackWidth") = BarValue(pName, "BarWidth") Then BarValue(pName, "BarStartX") = 0: Exit Property
    lDelta = BarValue(pName, "BarSize") + 1 + BarValue(pName, "BarLeft")
    If pRelative Then
        BarValue(pName, "BarX") = (BarValue(pName, "BarX") - (gBarClickX - X))
    Else
        BarValue(pName, "BarX") = X - BarValue(pName, "BarWidth") / 2
    End If
    If BarValue(pName, "BarX") < lDelta Then BarValue(pName, "BarX") = lDelta
    If BarValue(pName, "BarX") > lDelta + BarValue(pName, "BarBackWidth") - BarValue(pName, "BarWidth") Then BarValue(pName, "BarX") = lDelta + BarValue(pName, "BarBackWidth") - BarValue(pName, "BarWidth")
    lOldBarStartX = BarValue(pName, "BarStartX")
    BarValue(pName, "BarStartX") = -(BarValue(pName, "BarX") - lDelta) / (BarValue(pName, "BarBackWidth") - BarValue(pName, "BarWidth")) * (BarValue(pName, "BarMaxX") - BarCountVisibleX(pName)) * BarValue(pName, "BarElementWidth")
    If lOldBarStartX <> BarValue(pName, "BarStartX") Then RaiseEventRefreshNeeded pName, False
End Property

Private Sub BarDeleteRegions(pName As String)
On Error GoTo gestion_erreurs
goBars(pName).RegionsDelete
Exit Sub
gestion_erreurs:
'
End Sub

Private Sub BarCreateRegion(pBarName As String, pRegion As String, pX1 As Long, pY1 As Long, pX2 As Long, pY2 As Long)
Dim loBar As clGdiplus
On Error GoTo gestion_erreurs
    Set loBar = goBars(pBarName)
    loBar.WorldSetMatrix WorldGetMatrix
    loBar.CreateRegionRect pRegion, pX1, pY1, pX2, pY2
    Exit Sub
gestion_erreurs:
'
End Sub

' Dessine les barres
Public Sub BarDraw(Optional pName As String = "Default")
Dim lOldPenAlignMode As Long
Dim lBarVertVisible As Boolean, lBarHorzVisible As Boolean
    BarDeleteRegions pName
    lBarVertVisible = BarValue(pName, "BarVertVisible")
    lBarHorzVisible = BarValue(pName, "BarHorzVisible")
    If BarVertVisible(pName) Or BarHorzVisible(pName) Then
        lOldPenAlignMode = PenAlignMode
        PenAlignMode = 0 'PenAlignmentCenter
    End If
    If BarVertVisible(pName) Then
        ' Fond
        DrawRectangle BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarSize") + BarValue(pName, "BarTop"), BarValue(pName, "BarRight"), BarValue(pName, "BarSize") + BarValue(pName, "BarTop") + BarValue(pName, "BarBackHeight") + 1, BarValue(pName, "BarBackColor"), RGB(100, 100, 100), 0
        If lBarVertVisible Then BarCreateRegion pName, gBarKey & "BarYBack", BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarSize") + BarValue(pName, "BarTop"), BarValue(pName, "BarRight"), BarValue(pName, "BarSize") + BarValue(pName, "BarTop") + BarValue(pName, "BarBackHeight") + 1
        ' Barre
        If lBarVertVisible Then DrawControl BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarY"), BarValue(pName, "BarRight"), BarValue(pName, "BarY") + BarValue(pName, "BarHeight"), 19 ' CtrlButtonPush
        If lBarVertVisible Then BarCreateRegion pName, gBarKey & "BarYFront", BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarY"), BarValue(pName, "BarRight"), BarValue(pName, "BarY") + BarValue(pName, "BarHeight")
        ' Carre haut
        DrawControl BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarTop"), BarValue(pName, "BarRight"), BarValue(pName, "BarSize") + BarValue(pName, "BarTop"), 10 'CtrlScrollUp
        If lBarVertVisible Then BarCreateRegion pName, gBarKey & "BarYUp", BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarTop"), BarValue(pName, "BarRight"), BarValue(pName, "BarSize") + BarValue(pName, "BarTop")
        ' Carre bas
        DrawControl BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + BarValue(pName, "BarSize") * BarHorzVisible(pName), BarValue(pName, "BarRight"), BarValue(pName, "BarBottom") + BarValue(pName, "BarSize") * BarHorzVisible(pName), 11 'CtrlScrollDown
        If lBarVertVisible Then BarCreateRegion pName, gBarKey & "BarYDown", BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + BarValue(pName, "BarSize") * BarHorzVisible(pName), BarValue(pName, "BarRight"), BarValue(pName, "BarBottom") + BarValue(pName, "BarSize") * BarHorzVisible(pName)
    End If
    If BarHorzVisible(pName) Then
        ' Fond
        DrawRectangle BarValue(pName, "BarSize") + BarValue(pName, "BarLeft"), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarSize") + BarValue(pName, "BarLeft") + BarValue(pName, "BarBackWidth") + 1, BarValue(pName, "BarBottom"), BarValue(pName, "BarBackColor"), RGB(100, 100, 100), 0
        If lBarHorzVisible Then BarCreateRegion pName, gBarKey & "BarXBack", BarValue(pName, "BarSize") + BarValue(pName, "BarLeft"), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarSize") + BarValue(pName, "BarLeft") + BarValue(pName, "BarBackWidth") + 1, BarValue(pName, "BarBottom")
        ' Barre
        If lBarHorzVisible Then DrawControl BarValue(pName, "BarX"), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarX") + BarValue(pName, "BarWidth"), BarValue(pName, "BarBottom"), 19 'CtrlButtonPush
        If lBarHorzVisible Then BarCreateRegion pName, gBarKey & "BarXFront", BarValue(pName, "BarX"), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarX") + BarValue(pName, "BarWidth"), BarValue(pName, "BarBottom")
        ' Carre gauche
        DrawControl BarValue(pName, "BarLeft"), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarSize") + BarValue(pName, "BarLeft"), BarValue(pName, "BarBottom"), 12 'CtrlScrollLeft
        If lBarHorzVisible Then BarCreateRegion pName, gBarKey & "BarXDown", BarValue(pName, "BarLeft"), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarSize") + BarValue(pName, "BarLeft"), BarValue(pName, "BarBottom")
        ' Carre droite
        DrawControl BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + BarValue(pName, "BarSize") * BarVertVisible(pName), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarRight") + BarValue(pName, "BarSize") * BarVertVisible(pName), BarValue(pName, "BarBottom"), 13 'CtrlScrollRight
        If lBarHorzVisible Then BarCreateRegion pName, gBarKey & "BarXUp", BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + BarValue(pName, "BarSize") * BarVertVisible(pName), BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarRight") + BarValue(pName, "BarSize") * BarVertVisible(pName), BarValue(pName, "BarBottom")
    End If
    ' Carre en bas a droite
    If BarVertVisible(pName) And BarHorzVisible(pName) Then
        DrawRectangle BarValue(pName, "BarRight") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarBottom") - BarValue(pName, "BarSize") + 1, BarValue(pName, "BarRight"), BarValue(pName, "BarBottom"), BarValue(pName, "BarBackColor"), , 0
    End If
    If BarVertVisible(pName) Or BarHorzVisible(pName) Then
        PenAlignMode = lOldPenAlignMode
    End If
End Sub


