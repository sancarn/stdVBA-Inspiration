VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPNGwriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Write PNG image format when use of GDI+ not available ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' This class is only used on non-GDI+ systems


' This class is a partial version of a full-blown PNG creation class.
' It creates a PNG with many options, but not nearly all PNG options avaialble
' and is specifically modified to creating PNGs from pre-multiplied 32bpp DIBs.
' There still may be references to Interlacing. Interlacing options do not
' exist in this class and any such references are results of extracting the
' routines from the full-version PNG creation class.
' Note that GDI+ does not offer any PNG options when creating PNGs, this class
' exposes several options and can be modified to support all PNG options.

' CUSTOM TAILORED FOR PRE-MULTIPLIED 32bpp DIBS. Routines not portable for normal DIBs.

' Required is a version of the zLIB DLL which can be found at www.zlib.net.
' zLIB comes in at least two varieties: C calling convention (_cdecl) and
' VB/PASCAL calling convention (_stdcall). This routine can use either of those
' conventions, but the zLIB file must be named one of the two following,
' not case sensitive, both are original filenames:  zLib.dll or zLib1.dll

' Key highlights:
' 1. PNGs can be created without GDI+ as long as zlib or zlib1 is present
' 2. Using bit reduction algorithms, a 32bpp DIB can be converted to one of the
'       the following:  8 bpp paletted, 24 bpp or 32 bpp PNGs; supporting full alpha
' 3. The PNG compression filtering mechanism in this routine is user-selected.
'       Filters assist in reorganizing byte information to make it compress better
'       Speed vs Size tradeoffs: filter type None is fastest while type Paeth is smallest (generally)
'       See notes in FilterImage routine, set filter in c32bppDIB.PngPropertySet routine
' 4. Over a dozen options available when creating PNGs, see Me.AddProperty
' 5. This class almost always creates smaller PNG files than GDI+ when default filtering is used
' 6. PNGs can be saved to file or saved to an array



' array mapping structures
Private Type SafeArray
    cDims As Integer                ' numer of dimensions
    fFeatures As Integer            ' not used
    cbElements As Long              ' byte size of each element (byte=1,Int=2,Long=4)
    cLocks As Long                  ' not used
    pvData As Long                  ' pointer to memory space containing array
    rgSABound(0 To 3) As Long
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1

Private Enum eColorTypes    ' internal use only
    clrGrayScale = 0
    clrTrueColor = 2
    clrPalette = 3
    clrGrayAlpha = 4
    clrTrueAlpha = 6
End Enum

Private Const zlibMaxCompression = 9 ' when zlib compress2 is available

'----------------------------------------------------------------------------
' following are optional PNG properties only
Private m_Filter As eFilterMethods
Private m_bKGD As Long          ' default PNG background color if a view opts to render against solid bkg
Private m_Captions() As String  ' see c32bppDIB.ePngProperties
Private m_PNGprops As Long      ' indicates which, if any, m_Captions are used
'----------------------------------------------------------------------------

Private cCfunction As cCDECL        ' class to allow using C calling convention
Private m_ZLIBadv As Boolean        ' which version of zLIB? true=use compress2 else use compress
Private m_zLIBname As String        ' either zlib1 or zlib
Private m_Palette() As Byte         ' PNG palette if image can be palettized
Private m_transPal() As Byte        ' alpha values for PNG palettes as needed
Private m_Uncompressed() As Byte    ' initialized, contains uncompressed DIB bytes in 8,24,32 bit formats
Private m_Stream() As Byte          ' never initialized, overlay to host 32bpp DIB
Private m_Trans As Long             ' flag indicating whether or not transparency is used in DIB
Private m_ColorType As eColorTypes  ' the color type the PNG will be created in

Friend Function AddProperty(ByVal PropertyID As ePngProperties, Optional PropertyValue As Variant, Optional ByVal Caption_Misc As String) As Boolean
    
    ' Sets up to 14 properties: the properties are described below
    
    ' [PropertyID]
    ' pngProp_Title - Short (one line) title or caption for image
    ' pngProp_Author - Name of image's creator
    ' pngProp_Description - Description of image
    ' pngProp_Copyright - Copyright notice
    ' pngProp_CreationTime - Creation Time of original image creation. Should be GMT/Zulu time
    ' pngProp_Software - Software used to create the image
    ' pngProp_Disclaimer - Legal disclaimer
    ' pngProp_Warning - Warning or nature of content
    ' pngProp_Source - Device used to create the image
    ' pngProp_Comment - Comment
    ' ^^::^^ the above are registered/recognized keywords usable in PNGs
    
    ' pngProp_Miscellaneous - miscellaneous, can contain most any text and size.
    '   When used, you must supply a Caption/Keyword for the text in Caption_Misc parameter
    
    ' pngProp_DefaultBkgColor - background color used if PNG viewer opts to compose against only solid colors
    
    ' pngProp_DateTimeModified - date/time when the PNG was last modified. Should be GMT/Zulu time
    
    ' pngProp_FilterMethod - changes the precompression filter method used.
    '   See EncodeFilter_Avg, EncodeFilter_None, EncodeFilter_Paeth, EncodeFilter_Sub, EncodeFilter_Up
    
    '[PropertyValue]
    ' If not provided, the property associated with PropertyID is removed
    '   for pngProp_DefaultBkgColor, this parameter is a valid RGB color
    '       Note: For paletted this must be one of the
    '       palette colors otherwise the chunk may not be written.
    '   for pngProp_DateTimeModified, this parameter is a valid date/time. IsDate(PropertyValue) must be True
    '   for pngProp_FilterMethod, one of the c32bppDIB.eFilterMethods enumeration values
    '   for all others, this is text
    
    '[Caption_Misc]
    '   Only used when PropertyID is pngProp_Miscellaneous. This is the caption/keyword to
    '   be used with the supplied text. Cannot be one of the registered keywords
    
    ' attempt to disqualify bogus parameter values
    If IsObject(PropertyValue) Then Exit Function
    If IsNull(PropertyValue) Then Exit Function
    ' no combining propertyIDs
    If Not IsValidProperty(PropertyID) Then Exit Function
    
    On Error GoTo ExitRoutine
    Dim sText As String, X As Long, CaptionID As Long
    
    ' initialize array if needed
    If iparseIsArrayEmpty(VarPtrArray(m_Captions)) = 0& Then ReDim m_Captions(0 To 10)
        
    Select Case PropertyID
    Case ePngProperties.pngProp_DefaultBkgColor
        m_bKGD = (CLng(PropertyValue) And &HFFFFFF)  ' strip high byte if necessary
        
    Case ePngProperties.pngProp_DateTimeModified
        If IsDate(PropertyValue) = False Then Exit Function
        m_Captions(10) = CStr(PropertyValue)
        
    Case ePngProperties.pngProp_Miscellaneous   ' free-form text and captions/keywords
    
        CaptionID = FormatCaption(Caption_Misc)
        If Caption_Misc = vbNullString Then
            Exit Function ' must have a caption
        ElseIf Not CaptionID = -1& Then
            ' the caption passed is a reserved caption, reroute to appropriate property
            AddProperty = AddProperty(CaptionID, PropertyValue)
            Exit Function
        End If
        
        ' send to routine to format/validate the text
        sText = FormatText(CStr(PropertyValue))
        For CaptionID = 11& To UBound(m_Captions)
            ' LargeBlockText are formatted as:  Caption Chr$(0) Text
            X = InStr(m_Captions(CaptionID), Chr$(0))
            If StrComp(Caption_Misc, Left$(m_Captions(CaptionID), X - 1&), vbTextCompare) = 0& Then Exit For
        Next
        ' if appending, resize array
        If CaptionID > UBound(m_Captions) Then ReDim Preserve m_Captions(0 To CaptionID)
        m_Captions(CaptionID) = Caption_Misc & Chr$(0) & sText
        
    Case ePngProperties.pngProp_FilterMethod      ' filter type being modified
        m_Filter = CLng(PropertyValue)
        
    Case Else                           ' reserved keyword being modified
        m_Captions(PropertyIndex(PropertyID)) = FormatText(CStr(PropertyValue))
    End Select
    
    m_PNGprops = m_PNGprops Or PropertyID
    AddProperty = True
        
ExitRoutine:
End Function

Friend Function RemoveProperty(ByVal PropertyID As ePngProperties, ByVal Caption_Misc As String) As Boolean

    ' Parameters: See AddProperty
    ' [Caption_Misc] only used when PropertyID = pngProp_Miscellaneous
    
    ' no combining propertyIDs
    If Not IsValidProperty(PropertyID) Then Exit Function
    
    Dim sText As String, X As Long, CaptionID As Long
    
    ' deleting a property, do we have that property?
    If (m_PNGprops And PropertyID) = PropertyID Then
        
        If PropertyID = pngProp_Miscellaneous Then  ' miscellaneous/non-standard caption
            ' first see if passed caption is reserved or not
            CaptionID = FormatCaption(Caption_Misc)
            If Caption_Misc = vbNullString Then
                Exit Function
            ElseIf Not CaptionID = -1& Then ' reserved caption, reroute
                RemoveProperty = RemoveProperty(CaptionID, vbNullString)
                Exit Function
            Else
                CaptionID = PropertyIndex(PropertyID, Caption_Misc)
                If CaptionID = -1& Then
                    Exit Function
                Else
                    If CaptionID < UBound(m_Captions) Then
                        ' not the last, move the last into this one's position
                        m_Captions(CaptionID) = m_Captions(UBound(m_Captions))
                    End If
                    ' resize the captions array
                    ReDim Preserve m_Captions(0 To UBound(m_Captions) - 1&)
                End If
                ' if less than 11 captions, we no longer have an large block captions
                If UBound(m_Captions) = 10& Then m_PNGprops = (m_PNGprops And Not PropertyID)
            End If
            
        ElseIf PropertyID = pngProp_DefaultBkgColor Then
            m_PNGprops = (m_PNGprops And Not PropertyID)
            
        ElseIf PropertyID = pngProp_FilterMethod Then
            m_Filter = filterDefault
            m_PNGprops = (m_PNGprops And Not PropertyID)
            
        Else
            m_Captions(PropertyIndex(PropertyID)) = vbNullString
            m_PNGprops = (m_PNGprops And Not PropertyID)
        End If
    
        RemoveProperty = True
        
    End If

End Function

Friend Function GetProperty(ByVal PropertyID As ePngProperties, Optional ByVal Caption_Misc As String) As Variant
    
    ' VB5/4 users, can functions return variants? If not, move the return value
    ' into the parameters and set the return value below as needed
    
    ' Parameters: See AddProperty
    ' [Caption_Misc] only used when PropertyID = pngProp_Miscellaneous
    
    ' If property was not set, the following are returned
    ' if PropertyID=pngProp_DefaultBkgColor, -1
    ' if PropertyID=pngProp_DateTimeModified, nullDate; CDate(0)
    ' if PropertyID=pngProp_FilterMethod, 0&
    ' for all others, vbNullString
    
    ' no combining propertyIDs
    If Not IsValidProperty(PropertyID) Then Exit Function
    
    Dim CaptionID As Long, X As Long
    
    Select Case PropertyID
    Case ePngProperties.pngProp_DefaultBkgColor
        If (m_PNGprops And PropertyID) = 0& Then
            GetProperty = -1&
        Else
            GetProperty = m_bKGD
        End If
        
    Case ePngProperties.pngProp_DateTimeModified
        If (m_PNGprops And PropertyID) = 0& Then
            GetProperty = CDate(0#)
        Else
            GetProperty = CDate(m_Captions(10))
        End If
        
    Case ePngProperties.pngProp_Miscellaneous
        If (m_PNGprops And PropertyID) = 0& Then
            GetProperty = vbNullString
        Else
            CaptionID = FormatCaption(Caption_Misc)
            If Caption_Misc = vbNullString Then
                GetProperty = vbNullString
            ElseIf Not CaptionID = -1& Then
                GetProperty = GetProperty(CaptionID)
            Else
                For CaptionID = 11& To UBound(m_Captions)
                    X = InStr(m_Captions(CaptionID), Chr$(0))
                    If StrComp(Caption_Misc, Left$(m_Captions(CaptionID), X - 1&), vbTextCompare) = 0& Then
                        GetProperty = Mid$(m_Captions(CaptionID), X + 1&)
                        Exit For
                    End If
                Next
                If CaptionID > UBound(m_Captions) Then GetProperty = vbNullString
            End If
        End If
        
    Case ePngProperties.pngProp_FilterMethod
        GetProperty = m_Filter
        
    Case Else
        If (m_PNGprops And PropertyID) = 0& Then
            GetProperty = vbNullString
        Else
            GetProperty = m_Captions(PropertyIndex(PropertyID))
        End If
    End Select

End Function

Friend Function SavePNGex(cHost As c32bppDIB, FileName As String, outStream() As Byte) As Boolean

    If cHost.Handle = 0& Then Exit Function
    
    Dim tSA As SafeArray    ' overlay onto our DIB as needed
    Dim bSuccess As Boolean
    Dim FileNum As Integer
    Dim hFile As Long
    Dim bSkipBKGD As Boolean
    
    ' if we don't have Zlib, we can't continue with this class
    If zValidateZLIBversion = False Then Exit Function
    
    If Not FileName = vbNullString Then
        hFile = iparseGetFileHandle(FileName, False)
        If (hFile = INVALID_HANDLE_VALUE) Then Exit Function
    End If
    iparseOverlayHost_Byte m_Stream, VarPtr(tSA), 2, cHost.Height, cHost.scanWidth, cHost.BitsPointer ' overlay DMA array
    
    ' optimizations to reduce bit depth and reduce palette data
     On Error GoTo ExitRoutine
    ' Can image be palettized (smallest PNG size)?
    m_Trans = -1&
    If PalettizeImage(cHost.Alpha) = False Then
        ' if not, can we reduce to 24bpp from 32bpp?
        OptimizeTrueColor cHost.Alpha
    End If
    '  The above functions converted 32bpp DIB to necessary format for PNG creation
    ' The conversion is in the m_uncompress array. Remove overlay now
    iparseOverlayHost_Byte m_Stream, 0, 0, 0, 0, 0  ' remove DMA overlay
    tSA.cDims = 0
    
''   CREATE THE PNG using following rules
''   ------------------------------------
''    Critical chunks (must appear in this order):
''
''               Name  Multiple Ok?  Ordering constraints
''               IHDR    No          Must be first
''               PLTE    No          Before first IDAT (chunk is optional)
''               IDAT    Yes         Multiple IDATs must be consecutive
''               IEND    No          Must be last
''
''    Ancillary chunks (among other ancilliary chunks, order is not dictated):
''
''               Name  Multiple OK?  Ordering constraints relative to Critical chunks
''               cHRM    No          Before PLTE and IDAT
''               gAMA    No          Before PLTE and IDAT
''               iCCP    No          Before PLTE and IDAT
''               sBIT    No          Before PLTE and IDAT
''               sRGB    No          Before PLTE and IDAT
''               bKGD    No          After PLTE; before IDAT
''               hIST    No          After PLTE; before IDAT
''               tRNS    No          After PLTE; before IDAT
''               pHYs    No          Before IDAT
''               sPLT    Yes         Before IDAT
''               tIME    No          None
''               iTXt    Yes         None
''               tEXt    Yes         None
''               zTXt    Yes         None
    
    ' Write the PNG header
    If Write_IHDR(hFile, outStream, cHost, False) = False Then GoTo ExitRoutine
    If Write_tEXt(hFile, outStream, True) = False Then GoTo ExitRoutine ' write the Author & Title if needed
    If Write_PLTE(hFile, outStream, bSkipBKGD) = False Then GoTo ExitRoutine ' write the palette
    If Not bSkipBKGD Then   ' < may be set when bkgd color is not part of palette (Color Type 3 only)
        If Write_bKGD(hFile, outStream) = False Then GoTo ExitRoutine ' write bkgd color
    End If
    If Write_tRNS(hFile, outStream) = False Then GoTo ExitRoutine ' write transparency info
    ' Here we are going to filter & compress the DIB data & then write the IDAT chunk
    If FilterImage(hFile, outStream, cHost, m_Filter) = False Then GoTo ExitRoutine ' write data
    If Write_tIMe(hFile, outStream) = False Then GoTo ExitRoutine ' write last modified timestamp
    If Write_tEXt(hFile, outStream, False) = False Then GoTo ExitRoutine ' write other text (i.e., description, etc)
    If Write_zTXt(hFile, outStream) = True Then ' write any miscellaneous text
        ' Add the IEND termination to the PNG
        bSuccess = Write_IEND(hFile, outStream)   ' write the end flag
    End If

ExitRoutine:
    ' clean up as needed

    If Not tSA.cDims = 0 Then iparseOverlayHost_Byte m_Stream, 0, 0, 0, 0, 0  ' remove DMA overlay
    If Not hFile = 0& Then CloseHandle hFile
    
    Erase m_transPal()
    Erase m_Palette()
    Erase m_Uncompressed()
    Set cCfunction = Nothing
    
    If Err Then Err.Clear
    On Error Resume Next
    If bSuccess = False Then
        If hFile = 0& Then Erase outStream() Else iparseDeleteFile FileName
    Else
        SavePNGex = bSuccess
    End If
    If Err Then Err.Clear

End Function

Private Function PropertyIndex(PropertyID As ePngProperties, Optional Caption_Misc As String) As Long

    ' Helper function. Returns the m_Captions() array index for the passed PropertyID

    Dim X As Long, CaptionID As Long
    
    If PropertyID = pngProp_Miscellaneous Then
        For CaptionID = 11 To UBound(m_Captions)
            X = InStr(m_Captions(CaptionID), Chr$(0))
            If StrComp(Left$(m_Captions(CaptionID), X - 1&), Caption_Misc) = 0& Then Exit For
        Next
        If CaptionID > UBound(m_Captions) Then CaptionID = -1&
    Else
        X = PropertyID
        Do Until X = 1&
            X = X \ 2&
            CaptionID = CaptionID + 1&
        Loop
    End If

    PropertyIndex = CaptionID

End Function

Private Function PalettizeImage(isAlpha As AlphaTypeEnum) As Boolean

    ' Function determines if image can be palettized vs 24/32 bpp true color
    ' Once determined it can be paletted, it will optimize to include the following:
    ' 1. Convert to PNG grayscale palette if possible, saves at least 768 bytes vs color palette
    ' 2. Rearrange palette to reduce alpha/palette entries, saves up to 200+ bytes if alpha is used
    ' 3. Converts per-color grayscale to a modified color palette, reducing size at least 50%
    ' This modified version does not reduce to 1,2,or 4 bits per pixel
    '   -- Any paletted image is 256 colors, but only the needed palette entries are cached in PNG
    
    Dim X As Long, Y As Long, scanWidth As Long
    Dim palCount As Long, Index As Long
    Dim Color As Long, newColor As Boolean
    Dim palXRef() As Byte, palAlpha() As Byte
    Dim tSortIndex() As Long, tPalette() As Long
    
    On Error GoTo ExitRoutine
    
    ' count unique colors (maximum of 256 if we are to palettize)
    ' Note that alphas are included in the tSortIndex. This is because any color
    ' using more than one alpha value would require separate palette entries:
    ' Example: Red @ Alpha 255 & Red @ Alpha 128 requires two palette entries
    ReDim m_transPal(1 To 256)          ' array to hold alpha values only
    ReDim tSortIndex(1 To 256)          ' sort indexes
    ReDim tPalette(1 To 256) As Long    ' palette
    For Y = 0& To UBound(m_Stream, 2)
        For X = 0& To UBound(m_Stream, 1) Step 4&
        
            CopyMemory Color, m_Stream(X, Y), 4&
            Index = iparseFindColor(tSortIndex, Color, palCount, newColor)   ' use binary search routine
            If newColor = True Then
                If palCount = 256& Then Exit Function       ' exceeded palette entries limit
                palCount = palCount + 1&                    ' increment entry count & shift palette to maintain asc sort
                If Index < palCount Then
                    CopyMemory tSortIndex(Index + 1&), tSortIndex(Index), (palCount - Index) * 4&
                    CopyMemory tPalette(Index + 1&), tPalette(Index), (palCount - Index) * 4&
                End If
                tSortIndex(Index) = Color                    ' add new color to the palette
                CopyMemory tPalette(Index), Color, 3&
            End If
        
        Next
    Next
    
    ' if we got here, then image can be palettized, but to which of the following?
    ' 1. Palette - no transparency? like non-transparent GIFs (isAlpha=False)
    ' 2. Palette - simple transparency? like transparent GIFs (grayscale handled differently in PNGs)
    ' 3. Palette - per-color index transparency?
    
    Y = 0&
    If isAlpha Then
        ' separate alpha from color and count how many non-opaque alpha values
        For X = 1& To palCount
            If (tSortIndex(X) And &H7FFFFFFF) = tSortIndex(X) Then ' high bit not set
                m_transPal(X) = tSortIndex(X) \ &H1000000
            Else                                                 ' high bit is set
                m_transPal(X) = ((tSortIndex(X) And &H7FFFFFFF) \ &H1000000) Or &H80
            End If
            If Not m_transPal(X) = 255 Then
                Y = Y + 1&   ' count different levels of transparency
                Index = X    ' track last palette entry with alpha value <> 255
            End If
        Next
    Else
        FillMemory m_transPal(1), 256&, 255 ' all alphas are opaque
    End If
    
    Select Case Y
    Case 0&
    ' 1. Palette - no transparency? like non-transparent GIFs (isAlpha=False)
        m_Trans = -1& ' no transparency
    Case 1&
    ' 2. Palette - simple transparency? like transparent GIFs
        m_Trans = Index ' flag & may be changed later in this routine
    Case Else
    ' 3. Palette - per-color transparency?
        m_Trans = 0& ' > -1 means we have transparency at some level
        ' alphas are kept in the m_transPal() array
    End Select
    
'     Now for the last optimization attempt: check for grayscale but only for non per-color
'     alpha images. Why restrict grayscale to non per-color alpha when PNG can support grayscale
'     per-color alpha? Here's why: per-color grayscale alpha is ColorType 4. ColorType 4
'     always requires 16 bits per pixel (bpp), regardless of grayscale bit depth, but
'     ColorType 3 requires 8 bpp (max) + 768 palette bytes (max) + 256 bytes (max) for alpha info:
'       ColorType 4, 256x256 image: 256*256*2=131072 bytes for color information (grayscale has no palette in PNGs)
'       ColorType 3 (8bpp), 256x256 image: 256*256*1+768+256=66560 bytes for color information
'           note: ColorType 4 is always 16bpp, but ColorType 3 can be 1,2,4,8 bpp
'           and palette/alpha arrays can be reduced too
    
    m_ColorType = clrPalette    ' Color Type 3 (color palette)
    If Not m_Trans = 0& Then
        ' check each palette entry to see if grayscale or not. When not, abort loop
        For Index = 1& To palCount
            If Not (tPalette(Index) And &HFF) = ((tPalette(Index) \ &H100&) And &HFF) Then ' compare B to G
                Exit For
            ElseIf Not (tPalette(Index) And &HFF) = ((tPalette(Index) \ &H10000) And &HFF) Then ' compare B to R
                Exit For
            End If
        Next
        
        If Index > palCount Then    ' need to tweak transparency possibly
        
            m_ColorType = clrGrayScale ' Color Type 0
            
            If isAlpha > AlphaNone Then
                ' we only got here because just 1 color was transparent & with a pre-multiplied DIB
                ' that color is always black. But if non-transparent black was used elsewhere in the
                ' grayscale then we need to change the transparency. Non-transparent black is very
                ' common in grayscales
                ReDim palXRef(1 To 256) ' track which grayscales are used
                For X = 1& To palCount
                    If tPalette(X) = 0& Then             ' this is black
                        If Not m_transPal(X) = 0& Then   ' and not our transparent black
                            palXRef(1) = 1     ' mark black as used
                        End If
                    Else
                        palXRef(Index + 1&) = 1 ' non-black, mark as used
                    End If
                Next
                If palXRef(1) = 1 Then
                    ' non-transparent black is used in the grayscale, so we must change our
                    ' tranparent black - Locate a grayscale not in use
                    For Index = 2& To palCount
                        If palXRef(Index) = 0 Then
                            ' bingo, we'll use this one
                            m_Trans = Index - 1&
                            Exit For
                        End If
                    Next
                Else    ' black was not in the image, we can use black as transparency
                    m_Trans = 0&
                End If
                Erase palXRef
            End If
        End If
    End If
    
    scanWidth = UBound(m_Stream, 1) \ 4& + 1&         ' width of image
    X = (UBound(m_Stream, 2) + 1) * scanWidth - 1&   ' calculate size of total image bytes
    ReDim m_Uncompressed(0 To X)    ' the Filter function expects 1D arrays
    
    If m_ColorType = clrGrayScale Then
        ' grayscale is easy enough, transfer 32bpp info to 8bpp info
        ' Remember, PNG grayscale color types do not use palettes.
        ' Grayscale palettes are PNG decoders responsibility
        For Y = 0& To UBound(m_Stream, 2)
            Index = Y * scanWidth
            For X = 0& To UBound(m_Stream, 1) Step 4&
                If m_Stream(X + 3&, Y) = 0& Then     ' transparency index needed
                    m_Uncompressed(Index) = m_Trans ' use modified transparency index as necessary
                Else
                    m_Uncompressed(Index) = m_Stream(X, Y)  ' use grayscale index
                End If
                Index = Index + 1&
            Next
        Next
        
    Else
        ' for color palettes, we want to re-order entries when per-color alpha is used.
        ' Why the hassle? Shrink PNG a bit more. When color palettes have transparency,
        ' you must have a 1 byte Alpha value for each palette entry. But, that 1 byte
        ' alpha value, when = 255, is optional and implied. Therefore, if we move all
        ' palette entries with transparency to top of array, then all those 255s at the
        ' bottom of the array don't need to be cached in the PNG; not being there, PNG
        ' decoders must assume value is 255. We can save anywhere up to 200+ bytes
        ' depending on the image.
        ReDim palXRef(0 To 1, 0 To palCount - 1)
        
        If m_Trans = -1& Then   ' no transparencies and not grayscale
            For X = 0& To palCount - 1& ' all entries are opaque, no cross-referencing needed
                palXRef(0, X) = X
                palXRef(1, X) = X
            Next
        Else                    ' per-color alpha being used
            ' since we are re-ordering, we also need to build a cross-reference so
            ' we can reference palette locations, old to new and vice versa
            Y = 0&: X = palCount - 1&  ' starting points for top & bottom of array
            For Index = 0& To palCount - 1&
                If m_transPal(Index + 1&) = 255 Then
                    palXRef(1, X) = Index ' keep full opaque entries at bottom of array
                    palXRef(0, Index) = X ' double link reference
                    X = X - 1&
                Else
                    palXRef(0, Index) = Y ' move non-opaque entries near top of array
                    palXRef(1, Y) = Index ' double link reference
                    Y = Y + 1&
                End If
            Next
        End If
        
        ' now we build our 8 bpp paletted image, referencing the re-sorted palette entires
        For Y = 0& To UBound(m_Stream, 2)
            Index = Y * scanWidth
            For X = 0& To UBound(m_Stream, 1) Step 4&
                ' get 32bit color from DIB
                CopyMemory Color, m_Stream(X, Y), 4&
                ' locate it in our temp palette using binary search algorithm
                Color = iparseFindColor(tSortIndex, Color, palCount, False)
                ' now cache its re-sorted reference
                m_Uncompressed(Index) = palXRef(0, Color - 1&)
                Index = Index + 1&
            Next
        Next
        
        ' good, now we need to build the palette the PNG will use,
        ' but we will be using 3 byte values, not 4 byte values & colors need to be RGB vs BGR
        ReDim m_Palette(1 To palCount * 3& + 4&) ' extra 4 bytes are used during Write_PLTE
        For X = 1& To palCount
            ' calculate new index for this palette entry
            Index = palXRef(0, X - 1&) * 3& + 5&   ' offset that extra 4 bytes too
            ' simultaneously remove pre-multiplication
            Select Case m_transPal(X)
            Case 255    ' full opaque
                m_Palette(Index) = (tPalette(X) \ &H10000) And &HFF&
                m_Palette(Index + 1&) = (tPalette(X) \ &H100&) And &HFF&
                m_Palette(Index + 2&) = tPalette(X) And &HFF&
            Case 0: ' do nothing, color is always 0,0,0
            Case Else
                m_Palette(Index) = (((tPalette(X) \ &H10000) And &HFF&) * m_transPal(X) \ 255)
                m_Palette(Index + 1&) = (((tPalette(X) \ &H100&) And &HFF&) * m_transPal(X) \ 255)
                m_Palette(Index + 2&) = ((tPalette(X) And &HFF&) * m_transPal(X) \ 255)
            End Select
        Next
        Erase tPalette()
        
        If m_Trans = -1& Then
            Erase m_transPal()
        Else
            ' now we are going to double check how many non-opaque palette entries we have
            For X = 0& To palCount - 1&
                If m_transPal(palXRef(1, X) + 1&) = 255 Then
                    palCount = X    ' we saved 256-X bytes at least
                    Exit For
                End If
            Next
            ReDim palAlpha(1 To palCount + 4&)   ' extra 4 bytes used in Write_tRNS
            ' rewrite the m_transPal array, only caching non-opaque palette entries
            For X = 0& To palCount - 1&
                palAlpha(X + 5&) = m_transPal(palXRef(1, X) + 1&)
            Next
            m_transPal = palAlpha
        
        End If
    End If
    
    PalettizeImage = True

ExitRoutine:
End Function

Private Function OptimizeTrueColor(ByVal isAlpha As AlphaTypeEnum) As Boolean

    ' Function attempts to reduce 32bpp DIB to 24bpp DIB.  Check is not brute force
    ' Reduction to Palette already tried before this routine was called
    ' Reduction can occur when:
    '   1. No transparency is used (all Alpha values are 255), or
    '   2. Only simple transparency is used (if alpha between 1 & 254 then no reduction)
    
    Dim X As Long, Y As Long
    Dim bAbort As Boolean, tOffset As Long
    Dim palAlpha(0 To 255) As Byte, palCount As Long
    Dim scanWidth As Long, Color As Long
    
    m_Trans = -1&       ' flag indicating no simple transparency. ColorType 6 (clrTrueAlpha) implies transparency
    If isAlpha Then
        m_ColorType = clrTrueAlpha  ' default color type for this DIB
        
        ' can reduction to 24bpp be done?
        
        ' This routine will run quickly. We will not make the effort to check every possible
        ' color in the 16 million color range, rather, we will be looking at a max of 768
        ' colors: 256 Reds, 256 Greens, & 256 Blues (grayscales only). If we find one we
        ' can use, bingo, else write as 32bpp
        
        ' Now here's the catch. When full transparency is in play, black is always the transparent
        ' color in premultiplied DIBs, but if non-transparent black (i.e., RGB=0,0,0:Alpha>0) is
        ' used anywhere else in the image, then we can't leave black as the transparent color;
        ' we'll need to change it.
        
        For Color = 0& To 2&  ' Color = B, G, R
        
            tOffset = 3& - Color ' location of the alpha byte relative to "Color"
            palCount = 0&       ' number of "Color" shades used; from 1 to 256
            
            For Y = 0& To UBound(m_Stream, 2)
                For X = Color To UBound(m_Stream, 1) Step 4&
                
                    Select Case m_Stream(X + tOffset, Y)
                    Case 255    ' fully opaque
                        If palAlpha(m_Stream(X, Y)) = 0 Then    ' has it been counted?
                            palCount = palCount + 1&            ' up the count & abort if we maxed out
                            If palCount = 256& Then             ' all 256 shades of "Color" used
                                Y = UBound(m_Stream, 2)         ' force early termination of loop Y
                                Exit For
                            End If
                            palAlpha(m_Stream(X, Y)) = 1        ' flag it
                        End If
                    Case 0      ' fully transparent
                    Case Else   ' partial transparency              ' can't reduce to 24bpp
                        Y = UBound(m_Stream, 2)                     ' force early termination of loop Y
                        palCount = 256&                             ' prevent next IF statement from executing
                        Color = 3&                                  ' force early termination of loop Color
                        Exit For
                    End Select
                Next
            Next
            
            If Not palCount = 256& Then ' did we find a color we can use? It could be black
                For X = 0& To 255&  ' lets find out which it is
                    If palAlpha(X) = 0 Then
                        ' since the X-shade of the R, G, or B isn't used in the image,
                        ' we can safely state that RGB(X,X,X) is also not in the image
                        m_Trans = X Or (X * &H100&) Or (X * &H10000)
                        Exit For
                    End If
                Next
                m_ColorType = clrTrueColor  ' reduce to 24bpp vs 32bpp
                Exit For
            End If
            
            Erase palAlpha()    ' reset to zeros
        Next
    Else
        m_ColorType = clrTrueColor      ' no transparency, reduction to 24bpp
    End If
    
    ' Remove premultiplication & convert to 24bpp if needed
    ' Use separate loops vs adding an IF statement for every pixel to test for color type
    
    If m_ColorType = clrTrueAlpha Then  ' 32bpp (ColorType 6)
        scanWidth = UBound(m_Stream, 1) + 1&
        ReDim m_Uncompressed(0 To scanWidth * (UBound(m_Stream, 2) + 1&) - 1&)
        For Y = 0& To UBound(m_Stream, 2)
            tOffset = Y * scanWidth
            For X = 0& To UBound(m_Stream, 1) Step 4&
                ' remove pre-multiplication
                Select Case m_Stream(X + 3&, Y)
                Case 255
                    m_Uncompressed(tOffset) = m_Stream(X + 2&, Y)
                    m_Uncompressed(tOffset + 1&) = m_Stream(X + 1&, Y)
                    m_Uncompressed(tOffset + 2&) = m_Stream(X, Y)
                    m_Uncompressed(tOffset + 3&) = 255
                Case 0 ' do nothing
                Case Else
                    Color = m_Stream(X + 3&, Y)
                    m_Uncompressed(tOffset) = (255& * m_Stream(X + 2&, Y) \ Color)
                    m_Uncompressed(tOffset + 1&) = (255& * m_Stream(X + 1&, Y) \ Color)
                    m_Uncompressed(tOffset + 2&) = (255& * m_Stream(X, Y) \ Color)
                    m_Uncompressed(tOffset + 3&) = Color
                End Select
                tOffset = tOffset + 4&
            Next
        Next
    Else            ' 24bpp (Color Type 2) with or without simple transparency
        scanWidth = iparseByteAlignOnWord(24, UBound(m_Stream, 1) \ 4 + 1&)
        ' convert BGR to RGB, the Filter function expects 1D arrays
        ReDim m_Uncompressed(0 To scanWidth * (UBound(m_Stream, 2) + 1&) - 1&)
        For Y = 0& To UBound(m_Stream, 2)
            tOffset = Y * scanWidth
            For X = 0& To UBound(m_Stream, 1) Step 4&
                ' simultaneously remove pre-multiplication. Don't carry over any alpha values
                Select Case m_Stream(X + 3&, Y)
                Case 255
                    m_Uncompressed(tOffset) = m_Stream(X + 2&, Y)
                    m_Uncompressed(tOffset + 1&) = m_Stream(X + 1&, Y)
                    m_Uncompressed(tOffset + 2&) = m_Stream(X, Y)
                Case 0 ' uses simple transparency (1 color is transparent)
                    CopyMemory m_Uncompressed(tOffset), m_Trans, 3&
                Case Else
                    Color = m_Stream(X + 3&, Y)
                    m_Uncompressed(tOffset) = (255& * m_Stream(X + 2&, Y) \ Color)
                    m_Uncompressed(tOffset + 1&) = (255& * m_Stream(X + 1&, Y) \ Color)
                    m_Uncompressed(tOffset + 2&) = (255& * m_Stream(X, Y) \ Color)
                End Select
                tOffset = tOffset + 3&
            Next
        Next
    End If
End Function

Private Function ByteAlignOnByte(ByVal totalWidth As Long, ByVal btsPerPixel As Byte) As Long
' // LaVolpe, Dec 1 thru 10
    ' returns number of bytes required to display n pixels at p color depth (byte aligned)
    ByteAlignOnByte = (totalWidth * btsPerPixel + 7&) \ 8&
End Function

Private Function Write_IHDR(FileNum As Long, Stream() As Byte, Host As c32bppDIB, isInterlaced As Boolean) As Boolean

    Const png_Signature1 As Long = 1196314761
    Const png_Signature2 As Long = 169478669
    Const chnk_IHDR As Long = &H52444849 'Image header
    
    On Error GoTo eh
    Dim pngData(0 To 16) As Byte ' 13 byte header + 4 byte chunk name
    Dim gpLong As Long           ' general purpose variable
    Dim rwLen As Long
    
    ' build header
    CopyMemory pngData(0), chnk_IHDR, 4&    ' chunk name
    gpLong = iparseReverseLong(Host.Width)   ' png width
    CopyMemory pngData(4), gpLong, 4&
    gpLong = iparseReverseLong(Host.Height)  ' png height
    CopyMemory pngData(8), gpLong, 4&
    
    ' bit depth, 16bit (PNG 16 bytes per R,G,B element or 48 bytes per pixel)
    ' not supported via this class
    pngData(12) = 8 ' only 1,2,4,48 bpp are different, 8,24,32 bpp is 8
    
    pngData(13) = m_ColorType
    ' pngData(14) & (15) will always be zero (compression/filter methods)
    ' next byte is 1 if interlaced
    pngData(16) = Abs(isInterlaced)

    If FileNum = 0& Then ' writing to array vs file
        ReDim Stream(0 To 32) ' png signature, header len, header, crc value (33 bytes)
        CopyMemory Stream(0), png_Signature1, 4&
        CopyMemory Stream(4), png_Signature2, 4&
        gpLong = iparseReverseLong(13&) ' len of header
        CopyMemory Stream(8), gpLong, 4&
        CopyMemory Stream(12), pngData(0), 17&
        gpLong = zCreateCRC(VarPtr(pngData(0)), 17&)
        CopyMemory Stream(29), gpLong, 4&
        Write_IHDR = True
    Else
    
        WriteFile FileNum, png_Signature1, 4&, rwLen, ByVal 0&
        If rwLen = 4& Then
            WriteFile FileNum, png_Signature2, rwLen, rwLen, ByVal 0&
            If rwLen = 4& Then
                WriteFile FileNum, iparseReverseLong(13&), rwLen, rwLen, ByVal 0&
                If rwLen = 4& Then
                    WriteFile FileNum, pngData(0), 17&, rwLen, ByVal 0&
                    If rwLen = 17& Then
                        WriteFile FileNum, zCreateCRC(VarPtr(pngData(0)), rwLen), 4&, rwLen, ByVal 0&
                        Write_IHDR = (rwLen = 4&)
                    End If
                End If
            End If
        End If
    End If
eh:
    If Err Then Err.Clear
End Function


Private Function Write_PLTE(FileNum As Long, Stream() As Byte, Invalid_bKGD As Boolean) As Boolean

    ' Note: the palette is preprocessed before it arrives here: BGR>RGB
    On Error GoTo eh
    
    If m_ColorType = clrPalette Then ' paletted images only
    
        Const chnk_PLTE As Long = &H45544C50 'Palette
        
        Dim gpLong As Long          ' general purpose variable
        Dim Index As Long
        Dim rwLen As Long
        
        ' when paletted, the bKGD chunk comes after the palette, but for palettes the
        ' bkgd chunk must be one of the palette entries, therefore, we will attempt to
        ' find the color in the palette, add it to the palette if possible, or skip
        ' the optional chunk if color is not in the palette
        If (m_PNGprops And ePngProperties.pngProp_DefaultBkgColor) = ePngProperties.pngProp_DefaultBkgColor Then
            Dim bkg(0 To 2) As Byte
            CopyMemory bkg(0), m_bKGD, 3&
            For Index = 5& To UBound(m_Palette) Step 3&
                If bkg(0) = m_Palette(Index) Then
                    If bkg(1) = m_Palette(Index + 1&) Then
                        If bkg(2) = m_Palette(Index + 2&) Then Exit For
                    End If
                End If
            Next
            If Index < UBound(m_Palette) Then   ' found it, ref the index
                m_bKGD = (Index - 5&) \ 3&
            ElseIf UBound(m_Palette) < 772& Then ' we can add it, let's do that
                ' ^^ 772 is 256*3+4
                ReDim Preserve m_Palette(1 To UBound(m_Palette) + 3&)
                m_bKGD = (UBound(m_Palette) - 5&) \ 3&
                CopyMemory m_Palette(UBound(m_Palette) - 2&), bkg(0), 3&
            Else
                Invalid_bKGD = True ' do not write the bkgd chunk
            End If
        End If
            
        CopyMemory m_Palette(1), chnk_PLTE, 4&
        gpLong = UBound(m_Palette)
        
        If FileNum = 0& Then 'writing to array vs file
            Index = UBound(Stream) + 1&
            ReDim Preserve Stream(0 To Index + gpLong + 7&)
            rwLen = iparseReverseLong(gpLong - 4&)
            CopyMemory Stream(Index), rwLen, 4&                 ' size of chunk
            CopyMemory Stream(Index + 4&), m_Palette(1), gpLong ' palette
            rwLen = zCreateCRC(VarPtr(m_Palette(1)), gpLong)
            CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&   ' crc
            Write_PLTE = True
        Else
            WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
            If rwLen = 4& Then
                WriteFile FileNum, m_Palette(1), gpLong, rwLen, ByVal 0&
                If rwLen = gpLong Then
                    WriteFile FileNum, zCreateCRC(VarPtr(m_Palette(1)), gpLong), 4&, rwLen, ByVal 0&
                    Write_PLTE = (rwLen = 4&)
                End If
            End If
        End If
        Erase m_Palette()   ' no longer needed
    Else
        Write_PLTE = True
    End If
eh:
    If Err Then Err.Clear

End Function

Private Function Write_tEXt(FileNum As Long, Stream() As Byte, bTitleAuthorOnly As Boolean) As Boolean

    ' Function writes uncompressed standard Keywords & text to the PNG
    
    ' Note. Per PNG specs, some text should be written near top of the file while others
    ' should be written near the end. There is no requirement for text to appear in
    ' any specific location. The logic for writing some near the top is for search
    ' engines only. It would be faster to find that text if nearer the top.
    ' Therefore, this routine is called twice, once near the top of the PNG and
    ' and again just before the IEND chunk is written
    
    Const chnk_tEXt As Long = &H74584574 'Text - uncompressed
    
    On Error GoTo ExitRoutine
    
    Dim pngData() As Byte   ' data to be written to PNG file
    Dim txtData() As Byte
    Dim gpLong As Long
    Dim lenKeyword As Long
    Dim lenText As Long
    
    Dim Index As Long
    Dim CaptionID As Long
    Dim tProps As Long
    Dim lastCaption As Long
    Dim keyWord As String
    Dim rwLen As Long
    
    If bTitleAuthorOnly Then    ' called after writing IHDR
        CaptionID = ePngProperties.pngProp_Title
        lastCaption = ePngProperties.pngProp_Description
    Else                        ' called before writing IEND
        CaptionID = ePngProperties.pngProp_Description
        lastCaption = ePngProperties.pngProp_Miscellaneous
    End If
    tProps = m_PNGprops
    Do Until CaptionID = lastCaption
        If (tProps And CaptionID) = CaptionID Then
            tProps = tProps And Not CaptionID
            Select Case CaptionID
            Case pngProp_Title: Index = 0
                keyWord = "Title" & Chr$(0)
            Case pngProp_Author: Index = 1&
                keyWord = "Author" & Chr$(0)
            Case pngProp_Comment: Index = 9&
                keyWord = "Comment" & Chr$(0)
            Case pngProp_Copyright: Index = 3&
                keyWord = "Copyright" & Chr$(0)
            Case pngProp_CreationTime: Index = 4&
                keyWord = "Creation Time" & Chr$(0)
            Case pngProp_Description: Index = 2&
                keyWord = "Description" & Chr$(0)
            Case pngProp_Disclaimer: Index = 6&
                keyWord = "Disclaimer" & Chr$(0)
            Case pngProp_Software: Index = 5&
                keyWord = "Software" & Chr$(0)
            Case pngProp_Source: Index = 8&
                keyWord = "Source" & Chr$(0)
            Case pngProp_Warning: Index = 7&
                keyWord = "Warning" & Chr$(0)
            End Select
                
            ' tXTt chunk format::
            'Keyword 1-79 bytes (character string)
            'Null separator 1 byte (null character)
            'Text string 0 or more bytes (character string)
            
            lenKeyword = Len(keyWord)
            txtData() = StrConv(keyWord, vbFromUnicode)
            If Len(m_Captions(Index)) > 0& Then
                lenText = Len(m_Captions(Index))
                ReDim pngData(1 To lenKeyword + lenText + 4&)
                CopyMemory pngData(5), txtData(0), lenKeyword
                txtData() = StrConv(m_Captions(Index), vbFromUnicode)
                CopyMemory pngData(5& + lenKeyword), txtData(0), lenText
                
            Else ' handle zero-length chunks.
                ' Note: I would prefer to just skip these, but maybe you might
                ' decide to use one as a flag for something else?
                ReDim pngData(1 To lenKeyword + 4&)
                CopyMemory pngData(5), txtData(0), lenKeyword
            End If
            CopyMemory pngData(1), chnk_tEXt, 4&
            gpLong = lenKeyword + lenText + 4&
            
            If FileNum = 0& Then ' writing to stream
                Index = UBound(Stream) + 1&
                ReDim Preserve Stream(0 To Index + gpLong + 7&)
                rwLen = iparseReverseLong(gpLong - 4&)
                CopyMemory Stream(Index), rwLen, 4&
                CopyMemory Stream(Index + 4), pngData(1), gpLong
                rwLen = zCreateCRC(VarPtr(pngData(1)), gpLong)
                CopyMemory Stream(Index + 4& + gpLong), rwLen, 4&
                Write_tEXt = True
            Else
                WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
                If rwLen = 4& Then
                    WriteFile FileNum, pngData(1), gpLong, rwLen, ByVal 0&
                    If rwLen = gpLong Then
                        WriteFile FileNum, zCreateCRC(VarPtr(pngData(1)), gpLong), 4&, rwLen, ByVal 0&
                        Write_tEXt = (rwLen = 4&)
                    End If
                End If
            End If
        End If
        CaptionID = CaptionID * 2&
    Loop
    
ExitRoutine:
    If Err Then
        Err.Clear
    Else
        If lenKeyword = 0& Then Write_tEXt = True
    End If

End Function

Private Function Write_tIMe(FileNum As Long, Stream() As Byte) As Boolean

    ' Note: the time stamp should be Universal Time, not local area
    
    If (m_PNGprops And ePngProperties.pngProp_DateTimeModified) = ePngProperties.pngProp_DateTimeModified Then

        Const chnk_tIME As Long = &H454D4974 'Timestamp
        
        On Error GoTo eh
        Dim pngData(0 To 10) As Byte ' 7 byte date/time + 4 byte chunk name
        Dim gpLong As Long
        Dim gpInt As Integer
        Dim dtStamp As Date
        Dim rwLen As Long
        
        dtStamp = CDate(m_Captions(10))
        
        CopyMemory pngData(0), chnk_tIME, 4&
            gpInt = Year(dtStamp)
        CopyMemory pngData(5), gpInt, 2&
        pngData(4) = pngData(6)             ' swap endian of integer
            gpInt = Month(dtStamp)
        CopyMemory pngData(6), gpInt, 1&
            gpInt = Day(dtStamp)
        CopyMemory pngData(7), gpInt, 1&
            gpInt = Hour(dtStamp)
        CopyMemory pngData(8), gpInt, 1&
            gpInt = Minute(dtStamp)
        CopyMemory pngData(9), gpInt, 1&
            gpInt = Second(dtStamp)
        CopyMemory pngData(10), gpInt, 1&
        
        If FileNum = 0& Then ' writing to stream
            gpLong = UBound(Stream) + 1&
            ReDim Preserve Stream(0 To gpLong + 18&)
            rwLen = iparseReverseLong(7)
            CopyMemory Stream(gpLong), rwLen, 4&
            CopyMemory Stream(gpLong + 4&), pngData(0), 11&
            rwLen = zCreateCRC(VarPtr(pngData(0)), 11&)
            CopyMemory Stream(gpLong + 15&), rwLen, 4&
            Write_tIMe = True
        Else
            WriteFile FileNum, iparseReverseLong(7), 4&, rwLen, ByVal 0&
            If rwLen = 4& Then
                WriteFile FileNum, pngData(0), 11&, rwLen, ByVal 0&
                If rwLen = 11& Then
                    WriteFile FileNum, zCreateCRC(VarPtr(pngData(0)), rwLen), 4&, rwLen, ByVal 0&
                    Write_tIMe = (rwLen = 4&)
                End If
            End If
        End If
    Else
        Write_tIMe = True
    End If
eh:
    If Err Then Err.Clear

End Function

Private Function Write_tRNS(FileNum As Long, Stream() As Byte) As Boolean

    ' For paletted/grayscale images, tRNS is the palette index, otherwise RGB value
    On Error GoTo eh
        
    If m_Trans = -1& Then
        Write_tRNS = True
    
    Else ' transparency not used
    
        Const chnk_tRNS As Long = &H534E5274 'Simple Transparency & palette transparency
        Dim Index As Long
        Dim gpLong As Long
        Dim rwLen As Long
    
        Select Case m_ColorType
            
            Case clrPalette ' Paletted (palette count * 3 + 4 byte chunk name)
                ' nothing to do; done during PalettizeImage
            
            Case clrGrayScale ' grayscale
                ReDim m_transPal(1 To 6)   ' 2 bytes + 4 byte chunk name
                m_transPal(6) = m_Trans
                ' Note: m_transPal(5) used with 48bit per pixel images (not supported)
            
            Case clrTrueColor ' we have simple transparency for true color
                ReDim m_transPal(1 To 10)   ' 6 bytes + 4 byte chunk name
                m_transPal(6) = m_Trans And &HFF
                m_transPal(8) = (m_Trans \ &H100&) And &HFF
                m_transPal(10) = (m_Trans \ &H10000) And &HFF
                ' Note: m_transPal(5,7,9) used with 48bit per pixel images (not supported)
        
            Case Else
                ' Color Types 4 & 6 are prohibited from having a tRNS chunk
                Write_tRNS = True
                Exit Function
        End Select
        
        CopyMemory m_transPal(1), chnk_tRNS, 4&
        gpLong = UBound(m_transPal)
    
        ' write the chunk
        If FileNum = 0& Then ' writing to array vs file
            Index = UBound(Stream) + 1&
            ReDim Preserve Stream(0 To Index + gpLong + 7&)
            rwLen = iparseReverseLong(gpLong - 4&)
            CopyMemory Stream(Index), rwLen, 4&                 ' chunk size
            CopyMemory Stream(Index + 4&), m_transPal(1), gpLong ' palette
            gpLong = zCreateCRC(VarPtr(m_transPal(1)), gpLong)
            CopyMemory Stream(Index + UBound(m_transPal) + 4&), gpLong, 4&       ' crc value
            Write_tRNS = True
        Else
            WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
            If rwLen = 4& Then
                WriteFile FileNum, m_transPal(1), gpLong, rwLen, ByVal 0&
                If rwLen = gpLong Then
                    WriteFile FileNum, zCreateCRC(VarPtr(m_transPal(1)), gpLong), 4&, rwLen, ByVal 0&
                    Write_tRNS = (rwLen = 4&)
                End If
            End If
        End If
        Erase m_transPal()
    End If
eh:
    If Err Then Err.Clear

End Function

Private Function Write_zTXt(FileNum As Long, Stream() As Byte) As Boolean

    ' Function writes non-reserved keyword compressed/uncompressed text to the PNG
    
    If (m_PNGprops And ePngProperties.pngProp_Miscellaneous) = ePngProperties.pngProp_Miscellaneous Then

        On Error GoTo eh

        Const chnk_tEXt As Long = &H74584574 'Text - uncompressed
        Const chnk_zTXt As Long = &H7458547A 'Text - compressed
        
        Dim txtData() As Byte   ' comments/text in bytes
        Dim pngData() As Byte   ' data to be written to PNG file
        Dim sText As String
        Dim gpLong As Long
        Dim Index As Long
        Dim rwLen As Long
        Dim lenKeyword As Long
        Dim lenText As Long
        Dim bWritten As Boolean
    
        For Index = 11& To UBound(m_Captions)
            ' convert keyword to bytes
            lenKeyword = InStr(m_Captions(Index), Chr$(0))
            lenText = Len(m_Captions(Index)) - lenKeyword
            txtData() = StrConv(m_Captions(Index), vbFromUnicode)
            
            ' per PNG specs....
            ' It is recommended that text items less than 1K (1024 bytes)
            ' in size should be output using uncompressed text chunks
            If lenText > 1024& Then
            
                ' IMPORTANT: This portion of the routine is not equipped to write
                ' zero-length text block. That is only handled below where the
                ' .Text length is < 1025... DO NOT modify that IF statement to
                ' allow zero-length chunks to fall thru to this portion of IF
                
                ' zTXt chunk format::
                'Keyword 1-79 bytes (character string)
                'Null separator 1 byte (null character)
                'Compression method 1 byte
                'Compressed text datastream n bytes
                
                ' Note that the compression byte of zero needs to be included too,
                ' but we don't add it to the txtData conversion above cause zero
                ' would be converted to 48 -- Asc("0").
                
                gpLong = lenText * 0.01 + 12& + lenText
                '^^ Text won't always compress smaller; it should, but may not
                ' That is why it is recommended to allow 1024 bytes as uncompressed
                ReDim pngData(1 To gpLong + (lenKeyword + 5&))
                ' ^^ include 4 bytes for chunk name + keyword length + 1 byte compression method
                    
                If zDeflate(VarPtr(pngData(6& + lenKeyword)), gpLong, VarPtr(txtData(lenKeyword)), lenText) = True Then
                    ' ^^ store compression after chunk name, after keyword and after compression method
                    ' ^^ begin compression on 1st byte of the text, not the caption or compression method
                    
                    CopyMemory pngData(1), chnk_zTXt, 4&
                    CopyMemory pngData(5), txtData(0), lenKeyword
                    gpLong = gpLong + lenKeyword + 5&

                    If FileNum = 0& Then    ' writing to array
                        Index = UBound(Stream) + 1&
                        ReDim Preserve Stream(0 To Index + gpLong + 7&)
                        rwLen = iparseReverseLong(gpLong - 4&)
                        CopyMemory Stream(Index), rwLen, 4&
                        CopyMemory Stream(Index + 4&), pngData(1), gpLong
                        rwLen = zCreateCRC(VarPtr(pngData(1)), gpLong)
                        CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&
                        Write_zTXt = True
                    Else
                        WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
                        If rwLen = 4& Then
                            WriteFile FileNum, pngData(1), gpLong, rwLen, ByVal 0&
                            If rwLen = gpLong Then
                                WriteFile FileNum, zCreateCRC(VarPtr(pngData(1)), gpLong), 4&, rwLen, ByVal 0&
                                Write_zTXt = (rwLen = 4&)
                            End If
                        End If
                    End If
                    bWritten = True
                Else    ' failed to compress. Which means our buffer was too small
                        ' Therefore we will add it as uncompressed instead of
                        ' making the buffer even bigger
                End If
            End If
            
            If bWritten Then  'either len<1025 or compression failed
                bWritten = False    ' reset
            Else
                ' tXTt chunk format::
                'Keyword 1-79 bytes (character string)
                'Null separator 1 byte (null character)
                'Text string 0 or more bytes (character string)

                gpLong = lenText + lenKeyword + 4&  ' size of chunk
                ReDim pngData(1 To gpLong)
                CopyMemory pngData(1), chnk_tEXt, 4&
                CopyMemory pngData(5), txtData(0), lenKeyword
                
                If Not lenText = 0& Then ' zero-length text; not prohibited by PNG specs
                    CopyMemory pngData(5 + lenKeyword), txtData(lenKeyword), lenText
                End If
                
                If FileNum = 0& Then ' writing to array
                    Index = UBound(Stream) + 1&
                    ReDim Preserve Stream(0 To Index + gpLong + 7&)
                    rwLen = iparseReverseLong(gpLong - 4&)
                    CopyMemory Stream(Index), rwLen, 4&
                    CopyMemory Stream(Index + 4&), pngData(1), gpLong
                    rwLen = zCreateCRC(VarPtr(pngData(1)), gpLong)
                    CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&
                    Write_zTXt = True
                Else
                    WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
                    If rwLen = 4& Then
                        WriteFile FileNum, pngData(1), gpLong, rwLen, ByVal 0&
                        If rwLen = gpLong Then
                            WriteFile FileNum, zCreateCRC(VarPtr(pngData(1)), gpLong), 4&, rwLen, ByVal 0&
                            Write_zTXt = (rwLen = 4&)
                        End If
                    End If
                End If
            End If
        Next
    Else
        Write_zTXt = True
    End If
    
eh:
    If Err Then Err.Clear

End Function

Private Function Write_bKGD(FileNum As Long, Stream() As Byte) As Boolean

    ' For paletted/grayscale images, this is the palette index, otherwise RGB value
    On Error GoTo eh
    Const chnk_bKGD As Long = &H44474B62 'Window Background Color
    
    If (m_PNGprops And ePngProperties.pngProp_DefaultBkgColor) = ePngProperties.pngProp_DefaultBkgColor Then
    
        Dim pngData() As Byte
        Dim gpLong As Long
        Dim rwLen As Long
        Dim Index As Long
            
        ' Per PNG specs, bKGD chunk must come before IDAT and after PLTE
        Select Case m_ColorType
        Case clrPalette  ' 1 byte + 4 byte chunk name
            ReDim pngData(0 To 4)
            pngData(4) = CByte(m_bKGD)
        
        Case clrGrayScale, clrGrayAlpha   ' grayscales, 2 bytes + 4 byte chunk name
            ReDim pngData(0 To 6)
            pngData(5) = (m_bKGD And &HFF)
            ' pngData(4) used with 48bit per pixel images (not supported)
            
        Case Else ' true color, RGB format
            ReDim pngData(0 To 9)   ' 6 bytes + 4 byte chunk name
            pngData(5) = m_bKGD And &HFF
            pngData(7) = (m_bKGD \ &H100&) And &HFF
            pngData(9) = (m_bKGD \ &H10000) And &HFF
            ' Note: pngData(4,6,8) used with 48bit per pixel images (not supported)
        End Select
        
        CopyMemory pngData(0), chnk_bKGD, 4&
        gpLong = UBound(pngData) + 1&
        
        If FileNum = 0& Then 'writing to array
            Index = UBound(Stream) + 1&
            ReDim Preserve Stream(0 To Index + gpLong + 7&)
            rwLen = iparseReverseLong(gpLong - 4&)
            CopyMemory Stream(Index), rwLen, 4&
            CopyMemory Stream(Index + 4&), pngData(0), gpLong
            rwLen = zCreateCRC(VarPtr(pngData(0)), gpLong)
            CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&
            Write_bKGD = True
        Else
            WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
            If rwLen = 4& Then
                WriteFile FileNum, pngData(0), gpLong, rwLen, ByVal 0&
                If rwLen = gpLong Then
                    WriteFile FileNum, zCreateCRC(VarPtr(pngData(0)), gpLong), 4&, rwLen, ByVal 0&
                    Write_bKGD = (rwLen = 4&)
                End If
            End If
        End If
    Else
        Write_bKGD = True
    End If
eh:
    If Err Then Err.Clear
        
End Function

Private Function Write_IDAT(FileNum As Long, outStream() As Byte, imgData() As Byte, filterLen As Long) As Boolean

    ' Function writes the IDAT chunk(s). This routine writes IDAT as one chunk
    ' Note: IDATs can be written in multiple chunks; if so, chunks must be consecutive

    Const chnk_IDAT As Long = &H54414449 'Image data
    
    On Error GoTo eh
    Dim gpLong As Long, Index As Long
    Dim rwLen As Long
    
    CopyMemory imgData(0), chnk_IDAT, 4&
    If FileNum = 0& Then ' writing to array vs file
        Index = UBound(outStream) + 1&
        ReDim Preserve outStream(0 To Index + filterLen + 11&)
        gpLong = iparseReverseLong(filterLen)
        CopyMemory outStream(Index), gpLong, 4&  ' add chunk size
        CopyMemory outStream(Index + 4&), imgData(0), filterLen + 4& ' add compressed data
        gpLong = zCreateCRC(VarPtr(imgData(0)), filterLen + 4&)
        CopyMemory outStream(Index + 8& + filterLen), gpLong, 4& ' add crc value
        Write_IDAT = True
    Else
        
        WriteFile FileNum, iparseReverseLong(filterLen), 4&, rwLen, ByVal 0&
        If rwLen = 4& Then
            WriteFile FileNum, imgData(0), filterLen + 4&, rwLen, ByVal 0&
            If rwLen = filterLen + 4& Then
                WriteFile FileNum, zCreateCRC(VarPtr(imgData(0)), rwLen), 4&, rwLen, ByVal 0&
                Write_IDAT = (rwLen = 4&)
            End If
        End If
    End If
eh:
    If Err Then Err.Clear

End Function

Private Function Write_IEND(FileNum As Long, Stream() As Byte) As Boolean

    ' This must be the final chunk
    Const chnk_IEND As Long = &H444E4549 'End of Image
    
    On Error GoTo eh
    Dim Index As Long
    Dim gpLong As Long
    Dim rwLen As Long
    
    If FileNum = 0 Then ' writing to array vs file
    
        Index = UBound(Stream) + 1&
        ReDim Preserve Stream(0 To Index + 11&)
        CopyMemory Stream(Index), 0&, 4&
        CopyMemory Stream(Index + 4), chnk_IEND, 4&   ' chunk name, chunk length is zero
        gpLong = zCreateCRC(VarPtr(chnk_IEND), 4&)
        CopyMemory Stream(Index + 8&), gpLong, 4&    ' crc value
        Write_IEND = True
    Else
    
        WriteFile FileNum, rwLen, 4&, rwLen, ByVal 0&
        If rwLen = 4& Then
            WriteFile FileNum, chnk_IEND, 4&, rwLen, ByVal 0&
            If rwLen = 4& Then
                WriteFile FileNum, zCreateCRC(VarPtr(chnk_IEND), 4&), 4&, rwLen, ByVal 0&
                Write_IEND = (rwLen = 4&)
            End If
        End If
    End If
eh:
    If Err Then Err.Clear

End Function

Private Sub EncodeFilter_None(pngData() As Byte, _
                        ByVal RowNr As Long, dibRowNr As Long, _
                        ByVal scanLineDIB As Long, scanLinePNG As Long, _
                        stepVal As Byte, AdptValue As Long)

    ' this routine is only called when adapative filter method is used
    
    Dim X As Long
    Dim startByte As Long, locDIB As Long
    Dim lTest As Long
    
    If scanLineDIB > -1 Then ' processing interlaced image
        ' for interlaced, m_Uncompressed will be a top-down calculated array
        ' and the scanLineDIB parameter is an offset into the interlaced array
        startByte = scanLineDIB + 1
        locDIB = startByte
    Else
        ' for non-interlaced, m_Uncompressed will be a bottom-up DIB
        locDIB = dibRowNr * -scanLineDIB
        startByte = RowNr * scanLinePNG + RowNr + 1
   End If
    
    For X = locDIB To locDIB + scanLinePNG - 1
        lTest = lTest + m_Uncompressed(X)
        If lTest > AdptValue Then Exit Sub
    Next
    
    If lTest = 0 Then lTest = 1
    AdptValue = lTest
    pngData(startByte - 1) = 0

End Sub


Private Sub EncodeFilter_Up(pngData() As Byte, _
                        ByVal RowNr As Long, dibRowNr As Long, _
                        ByVal scanLineDIB As Long, scanLinePNG As Long, _
                        stepVal As Byte, AdptValue As Long)

' this is Filter Type 2
'http://www.w3.org/TR/PNG/#9-table91
    
    Dim ppTop As Integer
    Dim X As Long
    Dim startByte As Long, locDIB As Long
    Dim lTest As Long, prevRow As Long
    
    If scanLineDIB > -1 Then ' processing interlaced image
        ' for interlaced, m_Uncompressed will be a top-down calculated array
        ' and the scanLineDIB parameter is an offset into the interlaced array
        startByte = scanLineDIB + 1
        scanLineDIB = scanLinePNG + 1
        locDIB = startByte
    Else
        ' for non-interlaced, m_Uncompressed will be a bottom-up DIB
        locDIB = dibRowNr * -scanLineDIB
        startByte = RowNr * scanLinePNG + RowNr + 1
    End If
        
    If AdptValue Then
        
        If RowNr Then
            For X = locDIB To locDIB + scanLinePNG - 1
                lTest = lTest + Abs(0 + m_Uncompressed(X) - m_Uncompressed(X - scanLineDIB))
                If lTest > AdptValue Then Exit Sub
            Next
            
            If lTest = 0 Then lTest = 1
            AdptValue = lTest
            pngData(startByte - 1) = 2
        End If
        
    Else
        For X = 0 To scanLinePNG - 1
            If RowNr Then ppTop = m_Uncompressed(locDIB + X - scanLineDIB)
            ' VB workaround for C++ unsigned math
            If ppTop > m_Uncompressed(locDIB + X) Then
                pngData(startByte + X) = 256 - ppTop + m_Uncompressed(locDIB + X)
            Else
                pngData(startByte + X) = m_Uncompressed(locDIB + X) - ppTop
            End If
        Next
        pngData(startByte - 1) = 2
    End If

End Sub

Private Sub EncodeFilter_Sub(pngData() As Byte, _
                        ByVal RowNr As Long, dibRowNr As Long, _
                        ByVal scanLineDIB As Long, scanLinePNG As Long, _
                        stepVal As Byte, AdptValue As Long)

' This is Filter Type 1
'http://www.w3.org/TR/PNG/#9-table91

    Dim X As Long
    Dim startByte As Long, locDIB As Long
    Dim lTest As Long
    
    If scanLineDIB > -1 Then ' processing interlaced image
        ' for interlaced, m_Uncompressed will be a top-down calculated array
        ' and the scanLineDIB parameter is an offset into the interlaced array
        startByte = scanLineDIB + 1
        scanLineDIB = scanLinePNG + 1
        locDIB = startByte
    Else
        ' for non-interlaced, m_Uncompressed will be a bottom-up DIB
        locDIB = dibRowNr * -scanLineDIB
        startByte = RowNr * scanLinePNG + RowNr + 1
    End If
        
    If AdptValue Then
    
        ' 1st n bytes for 1st pixel are unfiltered
        For X = locDIB To stepVal + locDIB - 1
            lTest = lTest + m_Uncompressed(X)
        Next
        
        For X = locDIB + stepVal To scanLinePNG - 1
            lTest = lTest + Abs(0 + m_Uncompressed(X) - m_Uncompressed(X - stepVal))
            If lTest > AdptValue Then Exit Sub
        Next
        
        If lTest = 0 Then lTest = 1
        AdptValue = lTest
        
    Else
        ' 1st n bytes for 1st pixel are unfiltered
        CopyMemory pngData(startByte), m_Uncompressed(locDIB), stepVal
        
        For X = stepVal To scanLinePNG - 1
            ' VB workaround for C++ unsigned math
            If m_Uncompressed(locDIB + X - stepVal) > m_Uncompressed(locDIB + X) Then
                pngData(startByte + X) = 256 - m_Uncompressed(locDIB + X - stepVal) + m_Uncompressed(locDIB + X)
            Else
                pngData(startByte + X) = m_Uncompressed(locDIB + X) - m_Uncompressed(locDIB + X - stepVal)
            End If
        Next
    End If
    pngData(startByte - 1) = 1

End Sub

Private Sub EncodeFilter_Avg(pngData() As Byte, _
                        ByVal RowNr As Long, dibRowNr As Long, _
                        ByVal scanLineDIB As Long, scanLinePNG As Long, _
                        stepVal As Byte, AdptValue As Long)


' This is Filter Type 3
'http://www.w3.org/TR/PNG/#9-table91

    Dim ppLeft As Integer, ppTop As Integer
    Dim X As Long, pReturn As Integer
    Dim locDIB As Long, startByte As Long
    Dim lTest As Long
    
    If scanLineDIB > -1 Then ' processing interlaced image
        ' for interlaced, m_Uncompressed will be a top-down calculated array
        ' and the scanLineDIB parameter is an offset into the interlaced array
        startByte = scanLineDIB + 1
        scanLineDIB = scanLinePNG + 1
        locDIB = startByte
    Else
        ' for non-interlaced, m_Uncompressed will be a bottom-up DIB
        locDIB = dibRowNr * -scanLineDIB
        startByte = RowNr * scanLinePNG + RowNr + 1
    End If
    
    If AdptValue Then
    
        If RowNr Then
        
            For X = locDIB To locDIB + scanLinePNG - 1
                ppTop = m_Uncompressed(X - scanLineDIB)
                If X >= locDIB + stepVal Then ppLeft = m_Uncompressed(X - stepVal)
                lTest = lTest + Abs((m_Uncompressed(X) - (ppLeft + ppTop) \ 2))
                If lTest > AdptValue Then Exit Sub
            Next
            If lTest = 0 Then lTest = 1
            AdptValue = lTest
            pngData(startByte - 1) = 3
            
        End If
        
    Else
    
        For X = 0 To scanLinePNG - 1
        
            If RowNr Then ppTop = m_Uncompressed(locDIB - scanLineDIB + X)
            If X >= stepVal Then ppLeft = m_Uncompressed(locDIB - stepVal + X)
            
            pReturn = (ppLeft + ppTop) \ 2
            ' VB workaround for C++ unsigned math
            If pReturn > m_Uncompressed(locDIB + X) Then
                pngData(X + startByte) = 256 - pReturn + m_Uncompressed(X + locDIB)
            Else
                pngData(X + startByte) = m_Uncompressed(X + locDIB) - pReturn
            End If
        
        Next
        pngData(startByte - 1) = 3
    
    End If
    
End Sub

Private Sub EncodeFilter_Paeth(pngData() As Byte, _
                        ByVal RowNr As Long, dibRowNr As Long, _
                        ByVal scanLineDIB As Long, scanLinePNG As Long, _
                        stepVal As Byte, AdptValue As Long)

' This is Filter Type 4
'http://www.w3.org/TR/PNG/#9-table91

    Dim ppLeft As Integer, ppTop As Integer, ppTopLeft As Integer
    Dim X As Long, pReturn As Integer
    Dim startByte As Long, locDIB As Long
    Dim lTest As Long, prevRow As Long
    
    If scanLineDIB > -1 Then ' processing interlaced image
        ' for interlaced, m_Uncompressed will be a top-down calculated array
        ' and the scanLineDIB parameter is an offset into the interlaced array
        startByte = scanLineDIB + 1
        scanLineDIB = scanLinePNG + 1
        locDIB = startByte
    Else
        ' for non-interlaced, m_Uncompressed will be a bottom-up DIB
        locDIB = dibRowNr * -scanLineDIB
        startByte = RowNr * scanLinePNG + RowNr + 1
    End If
        
    If AdptValue Then
    
        If RowNr Then
            
            For X = locDIB To locDIB + scanLinePNG - 1
                If X >= stepVal + locDIB Then ' we are not on the 1st pixel
                    ppLeft = m_Uncompressed(X - stepVal)
                    ppTopLeft = m_Uncompressed(X - scanLineDIB - stepVal)
                End If
                ppTop = m_Uncompressed(X - scanLineDIB)
                ' get the Paeth closest neighbor
                lTest = lTest + Abs((m_Uncompressed(X) - PaethPredictor(ppLeft, ppTop, ppTopLeft)))
                If lTest > AdptValue Then Exit Sub
            Next
            
            If lTest = 0 Then lTest = 1
            AdptValue = lTest
            pngData(startByte - 1) = 4
        End If
    Else
    
        For X = 0 To scanLinePNG - 1
            
            If X >= stepVal Then ' we are not on the 1st pixel
                ppLeft = m_Uncompressed(locDIB + X - stepVal)
                If RowNr Then
                    prevRow = locDIB + X - scanLineDIB
                    ppTop = m_Uncompressed(prevRow)
                    ppTopLeft = m_Uncompressed(prevRow - stepVal)
                End If
            Else
                If RowNr Then ppTop = m_Uncompressed(locDIB + X - scanLineDIB)
            End If
            ' get the Paeth closest neighbor
            pReturn = PaethPredictor(ppLeft, ppTop, ppTopLeft)
            
            ' VB workaround for C++ unsigned math
            If pReturn > m_Uncompressed(locDIB + X) Then
                pngData(startByte + X) = 256 - pReturn + m_Uncompressed(locDIB + X)
            Else
                pngData(startByte + X) = m_Uncompressed(locDIB + X) - pReturn
            End If
            
        Next
        pngData(startByte - 1) = 4
    End If

End Sub

Private Function PaethPredictor(Left As Integer, Above As Integer, UpperLeft As Integer) As Integer

' http://www.w3.org/TR/PNG/#9-table91
' algorithm is used for both encoding & decoding the png paeth filter
' Based off of the formula created by Alan W. Paeth & provided fully in url above


    Dim pa As Integer, pb As Integer, pC As Integer, p As Integer
    p = Left + Above - UpperLeft
    pa = Abs(p - Left)
    pb = Abs(p - Above)
    pC = Abs(p - UpperLeft)
    
    ' tie breaker
    ' The order in which the comparisons are performed is critical and shall not be altered
    If (pa <= pb) And (pa <= pC) Then
        PaethPredictor = Left
    ElseIf pb <= pC Then
        PaethPredictor = Above
    Else
        PaethPredictor = UpperLeft
    End If

End Function

Private Function FilterImage(FileNum As Long, Stream() As Byte, Host As c32bppDIB, ByVal FilterMethod As eFilterMethods) As Boolean

    ' Routine will Filter the image in one of the 5 types of authorized PNG Filters.
    ' The Adaptive filter approach will select a best-guess filter to use for each
    ' scan line of the image. Otherwise, the same filter is applied to every scan line.
    
    ' Note about filters. Binary data compresses very poorly. Filters are a way to
    '   rewrite the binary data so that it will compress better. That is its only purpose.


    Dim scanWidth_DIB As Long, scanWidth_PNG As Long    ' scanwidths of 2 images
    Dim compressedData() As Byte            ' filtered PNG data
    Dim filteredData() As Byte              ' unfiltered PNG data
    Dim gpLong As Long                      ' general purpose Long value
    Dim arrayPtr As Long, pIndex As Long    ' array/loop variables
    Dim bytePP As Byte                      ' DIB/PNG bytes per pixel
    
    If FilterMethod < filterDefault Or FilterMethod > filterAdaptive Then FilterMethod = filterDefault

    Select Case m_ColorType
    Case clrGrayScale, clrPalette
        scanWidth_DIB = Host.Width
        scanWidth_PNG = scanWidth_DIB
        ' paletted images. Almost always, filter type zero is best (no filters)
        If FilterMethod = filterDefault Then FilterMethod = filterNone
        bytePP = 1
        
    Case Else ' true color, true color w/alpha (grayscale w/Alpha is converted to clrPalette in PalettizeImage)
        ' the best, non-adapative method for 24/32 bit is usually Paeth
        If FilterMethod = filterDefault Then FilterMethod = filterPaeth
        If m_ColorType = clrTrueAlpha Then
            ' get scan width for PNG file: byte aligned
            scanWidth_DIB = Host.scanWidth
            scanWidth_PNG = scanWidth_DIB
            bytePP = 4
        Else
            scanWidth_DIB = iparseByteAlignOnWord(24, Host.Width)
            scanWidth_PNG = ByteAlignOnByte(Host.Width, 24)
            bytePP = 3
        End If

    End Select
    
    ' Size raw data to be compressed and include 1 filter byte per line of image
    ReDim filteredData(0 To scanWidth_PNG * Host.Height + Host.Height - 1)
    If Err Then
        ' about the only possible error would be not enough memory to process the image file
        Err.Clear
        Exit Function
    End If
    
    
        For pIndex = 0 To Host.Height - 1
            
            arrayPtr = pIndex * scanWidth_PNG + pIndex   ' position of scanline
            
            If FilterMethod = filterAdaptive Then
                ' adaptive filtering
                ' although this can sequeeze an extra couple kb out of the png, I am finding
                ' that using Paeth appears to be either better or very close to adaptive filtering
                ' in most cases. Paeth is slowest of the top 5 filters (0-4). But adaptive filtering
                ' is significantly slower than Paeth. The deciding factor for adapative
                ' outdoing the others is the number of colors in the image. The more colors,
                ' the better chances adaptive has of being smaller size. The least amount
                ' of colors, the better chances Paeth has of being smaller.
                
                ' More testing needed though. I wouldn't imagine the PNG specs would recommend
                ' adaptive filtering unless it had some huge advantage over Paeth. What
                ' I am avoiding at all costs is a brute force routine to definitively
                ' find the best scanline filter method. That brute force can literally
                ' take minutes on full size 24/32bpp images.
                
                filteredData(arrayPtr) = 0
                gpLong = scanWidth_PNG * 254&  ' max value
                
                ' listed in order of quickest
                EncodeFilter_None filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, gpLong
                EncodeFilter_Sub filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, gpLong
                EncodeFilter_Up filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, gpLong
                EncodeFilter_Avg filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, gpLong
                EncodeFilter_Paeth filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, gpLong
                ' ^^ the -scanWidth_DIB is a flag indicating we are not filtering interlaced PNG
            Else
                filteredData(arrayPtr) = FilterMethod - 1
            End If
                    
            Select Case filteredData(arrayPtr) + 1 ' cache filter method into PNG data
            Case filterNone
                gpLong = (Host.Height - pIndex - 1) * scanWidth_DIB ' get current row relative to upside down DIB
                CopyMemory filteredData(arrayPtr + 1), m_Uncompressed(gpLong), scanWidth_PNG
            Case filterAdjLeft
                EncodeFilter_Sub filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, 0
            Case filterAdjTop
                EncodeFilter_Up filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, 0
            Case filterAdjAvg
                EncodeFilter_Avg filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, 0
            Case filterPaeth
                EncodeFilter_Paeth filteredData, pIndex, Host.Height - pIndex - 1, -scanWidth_DIB, scanWidth_PNG, bytePP, 0
            End Select
            ' ^^ the -scanWidth_DIB is a flag indicating we are not filtering interlaced PNG
        
        Next
    
    Erase m_Uncompressed    ' array no longer needed
    If Err Then
        Err.Clear
    Else
        gpLong = UBound(filteredData) + 1
        gpLong = gpLong * 0.01 + 12 + gpLong ' < sizing algorithm for compressed array per zLIB specs
        ReDim compressedData(0 To gpLong + 3) ' include 4 bytes for Write_IDAT
        If zDeflate(VarPtr(compressedData(4)), gpLong, VarPtr(filteredData(0)), UBound(filteredData) + 1) = True Then
            FilterImage = Write_IDAT(FileNum, Stream(), compressedData(), gpLong)
        End If
    
    End If

End Function

Private Function FormatText(txt2format As String) As String
    
    ' Function ensures text added to a PNG file meets PNG specs
    Dim Index As Integer
    Dim rtnString As String
    Const maxLength As Long = 32700&
    ' not a PNG restriction, but self-imposed. Keep text to Integer length
    
    ' per PNG specs, text and captions are limited to Latin1 character set and the line feed
    ' Latin1 character set is Chr$(32) & above.
    ' http://www.w3.org/TR/PNG/iso_8859-1.txt
    
    If txt2format = vbNullString Then
        rtnString = txt2format
    Else
        rtnString = Left$(txt2format, maxLength)
        
        ' per PNG specs, the only character allowed lower than a SPACE is the line feed character
        ' therefore we will replace vbCrLf with the line feed chr$(10)/vbLf
        rtnString = Replace$(rtnString, vbCrLf, vbLf)
        rtnString = Replace$(rtnString, vbCr, vbLf)
        
        For Index = 1 To Len(rtnString)
            Select Case Asc(Mid$(rtnString, Index, 1))
            Case 9, 10   ' LF is valid. Tab too, depending on sources, is ok
            Case Is < 32 ' all these are not allowed
                rtnString = vbNullString
                Exit For
            Case Else       ' otherwise, all other characters are allowed
            End Select
        Next
    End If
    
    FormatText = rtnString

End Function

Private Function FormatCaption(ByRef Caption As String) As Long

    ' per PNG specs, text and captions are limited to Latin1 character set and the line feed
    ' Latin1 character set is Chr$(32) & above.
    ' http://www.w3.org/TR/PNG/iso_8859-1.txt
    
    ' Return values
    ' If caption is a reserved caption,
    '       then FormatCaption=PropertyID of reserved caption
    '   else FormatCaption=-1 to indicate not reserved
    ' If caption is disqualified, then Caption is returned as vbNullString

    Dim Index As Long, CaptionID As Long
    
    If Not Caption = vbNullString Then
        
        ' we are formatting/validating a keyword/caption for miscellaneous text entries
            
        ' when adding miscellaneous text, a keyword is required, but must meet specs
        Caption = Trim$(Left$(Caption, 79)) ' absolute requirement
        Do Until InStr(Caption, "  ") = 0   ' absolute requirement, no double spacing
            Caption = Replace$(Caption, "  ", " ")
        Loop
        
        'only character codes 32-126 and 161-255 are allowed
        For Index = 1 To Len(Caption)
            Select Case Asc(Mid$(Caption, Index, 1))
            Case 160: Mid$(Caption, Index, 1) = Chr$(32) ' suggested: convert hard space to soft space
            Case 32 To 126  ' valid
            Case 161 To 255  ' valid
            Case Else       ' otherwise, all other characters are NOT allowed, invalidating caption
                Caption = vbNullString
                Exit For
            End Select
        Next
        
        ' now the last check, cannot use a reserved keyword
        If Not Caption = vbNullString Then CaptionID = isKeyWord(Caption)
    
    End If

    FormatCaption = CaptionID

End Function

Private Function isKeyWord(inCaption As String) As Long

    ' compares passed caption to PNG reserved keywords
    Dim Index As Long, keyWord As String, keyID As Long
        For Index = 1 To 11
            Select Case Index
            Case 1: keyWord = "Title": keyID = pngProp_Title
            Case 2: keyWord = "Author": keyID = pngProp_Author
            Case 3: keyWord = "Description": keyID = pngProp_Description
            Case 4: keyWord = "Copyright": keyID = pngProp_Copyright
            Case 5: keyWord = "Creation Time": keyID = pngProp_CreationTime
            Case 6: keyWord = "Software": keyID = pngProp_Software
            Case 7: keyWord = "Disclaimer": keyID = pngProp_Disclaimer
            Case 8: keyWord = "Warning": keyID = pngProp_Warning
            Case 9: keyWord = "Source'": keyID = pngProp_Source
            Case 10: keyWord = "Comment": keyID = pngProp_Comment
            Case 11: keyID = -1: Exit For
            End Select
            If StrComp(keyWord, inCaption, vbTextCompare) = 0 Then Exit For
        Next
    isKeyWord = keyID   ' return value of -1 indicates caption is not reserved
    
End Function

Private Function IsValidProperty(PropertyID As ePngProperties) As Boolean
    Select Case PropertyID
    Case pngProp_Author: IsValidProperty = True
    Case pngProp_Comment: IsValidProperty = True
    Case pngProp_Copyright: IsValidProperty = True
    Case pngProp_CreationTime: IsValidProperty = True
    Case pngProp_Description: IsValidProperty = True
    Case pngProp_Disclaimer: IsValidProperty = True
    Case pngProp_Miscellaneous: IsValidProperty = True
    Case pngProp_Software: IsValidProperty = True
    Case pngProp_Source: IsValidProperty = True
    Case pngProp_Title: IsValidProperty = True
    Case pngProp_Warning: IsValidProperty = True
    Case pngProp_DefaultBkgColor: IsValidProperty = True
    Case pngProp_FilterMethod: IsValidProperty = True
    Case pngProp_DateTimeModified: IsValidProperty = True
    End Select
    
End Function

' =======================================
' FOLLOWING 3 FUNCTIONS ARE ZLIB RELATED
' =======================================

Private Function zValidateZLIBversion() As Boolean

    ' Test for zlib availability & compatibility
    
    Dim procName As String, i As Integer
    
    procName = "compress2"
    m_ZLIBadv = True
    For i = 1 To 0 Step -1
        m_zLIBname = "zlib1.dll" ' try this filename first
        If iparseValidateDLL(m_zLIBname, procName) = True Then Exit For
        m_zLIBname = "zlib.dll"  ' try this filename next
        If iparseValidateDLL(m_zLIBname, procName) = True Then Exit For
        ' can't find latest version, look for earlier one (won't have compress2 function)
        procName = "compress"
        m_ZLIBadv = False
    Next
    
    If Not i = -1 Then ' else zlib or zlib1 not found
        Set cCfunction = New cCDECL
        zValidateZLIBversion = True
    End If
    
End Function

Private Function zCreateCRC(crcSrcRef As Long, srcLength As Long) As Long

    ' function returns zLIB's CRC value for passed crcTestRef value.
    Dim lReturn As Long
    If Not cCfunction Is Nothing Then
        lReturn = cCfunction.CallFunc(m_zLIBname, retLong, "crc32", 0&, crcSrcRef, srcLength)
    End If
    If Not lReturn = 0& Then zCreateCRC = iparseReverseLong(lReturn)
    
End Function

Private Function zDeflate(destRef As Long, destSize As Long, srcRef As Long, srcSizeRef As Long) As Boolean

    ' function compresses/deflates passed srcRef into passed destRef and modifies the destSizeRef to indicate byte count of destRef
    
    ' earliest versions of DLL do not have Compress2 which newer versions have.
    ' Newer versions allow a compression parameter to allow deeper compression.
    ' When compress is called in newer DLL it just reroutes to the compress2 method
    
    Dim procName As String
    If Not cCfunction Is Nothing Then
        If m_ZLIBadv Then
            zDeflate = (cCfunction.CallFunc(m_zLIBname, retLong, "compress2", destRef, VarPtr(destSize), srcRef, srcSizeRef, zlibMaxCompression) = 0&)
        Else    ' older version has 1 less parameter
            zDeflate = (cCfunction.CallFunc(m_zLIBname, retLong, "compress", destRef, VarPtr(destSize), srcRef, srcSizeRef) = 0&)
        End If
    End If

End Function




