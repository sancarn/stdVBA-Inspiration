VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read MP3/WMA image formats ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' Class attempts to extract image formats from an MP3 or WMA file.
' The image format returned by this class depends on the number of images
' extracted from the MP3, if any.

' When only a single image is extracted, that image format is returned
' When 2 or more images are extracted, each are combined to a multi-page TIFF

' When multiple images do exist, any image that is tagged as "Front Cover" will
' always be the 1st image in the returned multi-page TIFF

' ////// MP3 \\\\\\
' Using 3 separate parsing routines, one for each version of the id3 tags; though versions are similar
'   but there are enough subtle differences that would require a lot of IF Version=xxx statements
' =================================================================================================
' id3v2.x uses 3 byte length values & 3 character frame types & no option for extended headers
' id3v3.x uses 4 byte length values & 4 character frame types. Extended header size is not sync-safe
' id3v4.x similar to id3v3 except Extended header size is synch-safe & additional frame flag values
'   id3v4 offers 4 types of text encoding vs just the 2 types in earlier versions
'   id3v4 allows individual frames to be unsynchronized vs. just the entire tag
'   id3v4 uses sync-safe frame lengths, earlier versions do not

' Parsing routines will undo unsynchronization for all 3 versions. See pvSynchronizeTag for more
' If any compression is used (all versions) or encryption (id3v4) then frames are skipped
' Though I can include decompression routines, I haven't found any mp3s that use zLib compression

' fyi. id3v1 not supported simply because it has no capability to store images
' there may never be a id3v5 since id3v4 is having troubles catching on

'/////// WMA \\\\\\\\
' Only aware of 1 format for images in WMA files.
' As far as I know, any images will always be contained within the Extended Header of the WMA
' and within the Extended header, images will only be contained within the metadata library records
' See pvParseWMA for more
 
Private Type MP3Offsets     ' used to track image data found within the passed tag
    Start As Long           ' where image starts in the tag
    length As Long          ' number of bytes for the image data
    Category As Long        ' what category the image belongs to
End Type
Private m_Token As cGDIpToken

Public Function IsMP3Resource(theStream() As Byte, FileHandle As Long) As Long

    ' quickly determines if MP3 format possibility
    ' updated to include WMA format also
    
    Dim lResult As Long, lSize As Long, sGUID As String
    Const tagWMAheader As String = "{75B22630-668E-11CF-A6D9-00AA0062CE6C}"
    
    If FileHandle Then
        lSize = GetFileSize(FileHandle, 0&)
        If lSize < 30& Then Exit Function
        ReDim theStream(0 To 29)                        ' read 29 bytes; enough to check for mp3 or wma
        SetFilePointer FileHandle, 0&, 0&, 0&
        ReadFile FileHandle, theStream(0), 30&, lResult, ByVal 0&
    Else
        lSize = UBound(theStream) + 1&
        If lSize < 30& Then Exit Function
    End If
    
    ' test for ID3 header
    If (theStream(0) = 73 And theStream(1) = 68 And theStream(2) = 51) Then
        If theStream(3) > 1 And theStream(3) < 5 Then       ' supports versions 2,3,4
            If theStream(4) < 255 Then                      ' revision will never be 255
                lResult = pvSyncSafeLong(theStream, 6&)
                If lResult <= lSize And lResult > 60& Then  ' else invalid or too small for image data
                    If FileHandle Then                      ' extract just the tag portion
                        ReDim theStream(0 To lResult - 1&)
                        SetFilePointer FileHandle, 0&, 0&, 0&
                        ReadFile FileHandle, theStream(0), lResult, lSize, ByVal 0&
                        If lSize = lResult Then IsMP3Resource = theStream(3)
                    Else
                        IsMP3Resource = theStream(3)
                    End If
                End If
            End If
        End If
    ElseIf theStream(3) = 117 Then  ' will be 117 if the 1st 16 bytes are the proper GUID
        ' test for wma format
        sGUID = String$(38, vbNullChar)
        If StringFromGUID2(VarPtr(theStream(0)), StrPtr(sGUID), 39) > 0& Then
            If sGUID = tagWMAheader Then
                CopyMemory lResult, theStream(20), 4&
                If lResult = 0& Then ' else the file is either very huge or malformatted wma
                    CopyMemory lResult, theStream(16), 4&
                    If lResult > 60& And lResult <= lSize And theStream(29) = 2 Then
                        ' This field (29) must be set to the value 0x02.
                        ' If the value is different when read, the application should fail to source the content.
                        If FileHandle Then
                            ReDim theStream(0 To lResult - 1&)
                            SetFilePointer FileHandle, 0&, 0&, 0&
                            ReadFile FileHandle, theStream(0), lResult, lSize, ByVal 0&
                            If lResult = lSize Then IsMP3Resource = 1& ' non-zero value
                        Else
                            IsMP3Resource = 1& 'non-zero return value
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function

Public Function LoadMP3(theStream() As Byte, sourceData As Variant) As Long

    ' Function passes off parsing to individual routines and then creates
    ' the image(s) extracted within those routines
    
    ' IsMP3Resource must first be called and return a non-zero value before this function is ever called
    
    If m_Token.Version = 0& Then Exit Function
    
    Dim mp3Images() As MP3Offsets, bBits() As Byte, lIndexes() As Long
    Dim tSA As SafeArray, I As Long, gPicAdd As GDIpImage
    Dim bReleaseTIFF As Boolean, SS As SAVESTRUCT
    
    Const tagFRONTCOVER As Long = &H3&
    Const tagBACKCOVER As Long = &H4&
    
    Select Case theStream(3) ' mp3 version or wma flag
        Case 2: I = pvParse3v2(theStream, mp3Images())
        Case 3: I = pvParse3v3(theStream, mp3Images())
        Case 4: I = pvParse3v4(theStream, mp3Images())
        Case 117: I = pvParseWMA(theStream, mp3Images)
    End Select
    If I = 0& Then Exit Function                            ' failed to parse any images
    
    If I = 1& Then                                          ' just 1 image
        ReDim bBits(0 To mp3Images(0).length - 1&)          ' will be using original format
        CopyMemory bBits(0), theStream(mp3Images(0).Start), mp3Images(0).length
        modCommon.MoveArrayToVariant sourceData, bBits(), True
        LoadMP3 = I                                         ' return non-zero
    Else
    
        ' create a pecking order, where we'll ensure any cover art images are first in the TIFF
        ReDim lIndexes(0 To I - 1&)
        For I = 1& To I - 1&: lIndexes(I) = I: Next
        For I = 0& To I - 1&
            If mp3Images(I).Category = tagFRONTCOVER Or mp3Images(I).Category = tagBACKCOVER Then
                If I > 0& Then      ' cover art; we want this first in our indexes
                    CopyMemory lIndexes(1), lIndexes(0), I * 4&
                    lIndexes(0) = I
                End If
                If mp3Images(I).Category = tagFRONTCOVER Then Exit For
            End If
        Next
    
        SS.RSS.FillColorARGB = modCommon.Color_RGBtoARGB(vbWindowBackground, 0&)
        tSA.cbElements = 1
        tSA.cDims = 1                                       ' we'll use the source array for our images
        For I = 0& To UBound(mp3Images)                     ' since LoadImage makes copy of array if needed
            With mp3Images(lIndexes(I))                     ' no point in copying it to pass to the function
                tSA.pvData = VarPtr(theStream(.Start))
                tSA.rgSABound(0).cElements = .length
            End With
            
            CopyMemory ByVal VarPtrArray(bBits()), VarPtr(tSA), 4&  ' overlay temp array on source
            Set gPicAdd = modCommon.LoadImage(bBits(), , , True)    ' Send the hacked array for processing
            CopyMemory ByVal VarPtrArray(bBits()), 0&, 4&           ' undo hacking
            
            If gPicAdd.Handle Then                          ' add image to a TIFF
                SS.reserved2 = ObjPtr(gPicAdd)
                SS.reserved1 = Abs(modCommon.HasTransparency(gPicAdd.Handle))
                SS.Width = gPicAdd.Width: SS.Height = gPicAdd.Height
                If bReleaseTIFF Then                        ' done adding?
                    If modCommon.SaveAsTIFF(sourceData, gPicAdd.Handle, saveTo_GDIplus, TIFF_MultiFrameAdd, SS) = 0& Then
                        bReleaseTIFF = False                ' flag indicating failure
                        Exit For
                    End If
                Else
                    If modCommon.SaveAsTIFF(sourceData, gPicAdd.Handle, saveTo_GDIplus, TIFF_MultiFrameStart, SS) Then
                        bReleaseTIFF = True                 ' flag indicating success
                    Else
                        bReleaseTIFF = False                ' failure
                        Exit For
                    End If
                End If
            End If
        
        Next
        Set gPicAdd = Nothing                               ' don't need this any longer
        If bReleaseTIFF Then                                ' successful? If so, close out TIFF
            LoadMP3 = modCommon.SaveAsTIFF(sourceData, 0&, saveTo_GDIplus, TIFF_MultiFrameEnd, SS)
        End If
    End If

End Function

Private Function pvReverseLong4(inStream() As Byte, lOffset As Long) As Long

    ' Convert 4 byte big endian to little endian
    pvReverseLong4 = inStream(lOffset + 3&) Or inStream(lOffset + 2&) * &H100& _
        Or inStream(lOffset + 1&) * &H10000 Or (inStream(lOffset) And &H7F) * &H1000000
    If (inStream(lOffset) And &H80) Then pvReverseLong4 = pvReverseLong4 Or &H80000000

End Function

Private Function pvReverseLong3(inStream() As Byte, lOffset As Long) As Long

    ' Convert 3 byte big endian to 4 byte little endian
    pvReverseLong3 = inStream(lOffset + 2&) Or inStream(lOffset + 1&) * &H100& Or inStream(lOffset) * &H10000

End Function

Private Function pvSyncSafeLong(inStream() As Byte, lOffset As Long) As Long

    ' Convert 4 byte big endian sync-safe to 4 byte little endian
    ' Sync-safe values are 28 bits vs 32 bits. The high bit of each byte is zero/ignored
    ' -- concatenate the 1st 7 bits of each byte to create a long value
    pvSyncSafeLong = (inStream(lOffset + 3&) And &H7F) Or (inStream(lOffset + 2&) And &H7F) * &H80& _
        Or (inStream(lOffset + 1&) And &H7F) * &H4000& Or (inStream(lOffset) And &H7F) * &H200000

End Function

Private Function pvSynchronizeTag(inStream() As Byte, Offset As Long, nrBytes As Long) As Long

    Dim aPtr As Long, zPtr As Long
    Dim lResult As Long, aPtrAnchor As Long

    ' if unsync'd, then need to fix it up before parsing
    ' unsynchronizing a tag was done for much older mp3 players, that didn't recognize ID3 tags
    
    ' 1. the unsync encoding (done by whoever created the id3 tag) looks for any 11 consecutive bits that are all on
    ' 2. then it splits the 1st 8 bits from the next 3 bits by inserting a null byte
    ' 3. but the splitting affects any naturally occurring 255 0 byte combinations, so it first must fix those
    '   by inserting a null byte between them: so 11111111 00000000 will end up as 11111111 00000000 00000000
    ' 4. Now for sync'ing:
    '   these 2 bytes: 11111111 111xxxxx will end up as 3 bytes: 11111111 00000000 111xxxxx
    ' 5. to remove the scheme, we need to simply remove any null byte that follows a 255 byte
    
    ' id3v4 has ability to unsync individual frames, less their 10 byte frame header
    ' earlier versions applied unsync to entire tag, less the 6/10 byte tag header
    '   id3v4 can also apply this to the entire tag, same logic
    
    ' Note: Any array passed to this function is owned by this control. Am not messing with user data
    
    zPtr = Offset: aPtrAnchor = Offset
    For aPtr = Offset + 1& To Offset + nrBytes - 1&
        If inStream(aPtr) = 0& Then                 ' see if previous byte was 255
            If inStream(aPtr - 1&) = 255 Then       ' shift array left
                If lResult Then CopyMemory inStream(zPtr), inStream(aPtrAnchor), aPtr - aPtrAnchor
                zPtr = zPtr + aPtr - aPtrAnchor     ' adjust position for next write
                aPtrAnchor = aPtr + 1&              ' adjust position for first inclusive byte
                lResult = lResult + 1&              ' keep track of number of bytes removed
            End If
        End If
    Next
    If (aPtrAnchor < aPtr) And (lResult > 0&) Then  ' perform last shift if needed
        CopyMemory inStream(zPtr), inStream(aPtrAnchor), aPtr - aPtrAnchor
        inStream(Offset + nrBytes - lResult - 1&) = 0 ' ensure adjusted tag terminates with zero
    End If
    pvSynchronizeTag = lResult

End Function

Private Function pvParse3v2(inStream() As Byte, tImages() As MP3Offsets) As Long

    ' http://www.id3.org/id3v2-00
    
    ' id3v2.3.0 header block format
    ' offset 0 : 3 byte magic number: ID3
    ' offset 3 : 1 byte major version
    ' offset 4 : 1 byte revision
    ' offset 5 : 1 byte flag   (bit 7=unsync'd, 6=compression)
    ' offset 6 : 4 byte sync-safe length

    Dim lMax As Long, aPtr As Long, lSize As Long
    Dim lValue As Long, imgOffset As Long, lCount As Long
    Const tagPIC As Long = 4409680                      ' only care about these frames
    
    ' frame header (6 bytes always)
    ' aPtr + 0 : 3 byte header type (i.e., PIC)
    ' aPtr + 3 : 3 byte header size as big endian long
    If (inStream(5) And &H40) Then Exit Function        ' standard compression was never implemented
    
    aPtr = 10&
    lMax = pvSyncSafeLong(inStream(), 6&)
    If (inStream(5) And &H80) Then lMax = lMax - pvSynchronizeTag(inStream(), aPtr, lMax - aPtr)
    If lMax < aPtr Then Exit Function
    
    Do While aPtr < lMax
        If inStream(aPtr) = 0& Then Exit Do             ' padding; done
        For lValue = aPtr To aPtr + 2&                  ' validate header type (ASCII A-Z 0-9)
            Select Case inStream(lValue)
            Case 65 To 90, 48 To 57
            Case Else
                Exit Do                                 ' malformatted id3 tag
            End Select
        Next
        lSize = pvReverseLong3(inStream(), aPtr + 3&)   ' and its overall size
        If lSize + aPtr <= lMax Then
            CopyMemory lValue, inStream(aPtr), 3&       ' get frame type
            If (lValue And &HFFFFFF) = tagPIC And lSize > 0& Then ' zero-length frames are not uncommon
                ' PIC header
                ' aPtr + 6      : 1 byte text type (0=ASCII, 1=unicode)
                ' aPtr + 7      : 3 byte image format (PNG, JPG, etc)
                ' aPtr + 10     : 1 byte image category (i.e., album front, back, etc)
                ' aPtr + 11     : t bytes image description (depends on text type byte)
                ' aptr + 11+t   : binary data for the image
                '   Note: if image format is "-->" then image data is a URL & excluded from processing here
                If Not (inStream(aPtr + 7&) = 45 And inStream(aPtr + 8&) = 45) Then ' else dashes in the image type field
                    ' determine length of image description (max of 64)
                    Select Case inStream(aPtr + 6&)
                        Case 0: imgOffset = lstrlen(VarPtr(inStream(aPtr + 11&))) + 1&
                        Case 1: imgOffset = lstrlenW(ByVal VarPtr(inStream(aPtr + 11&))) * 2& + 2&
                        Case Else: imgOffset = 0&
                    End Select
                    If imgOffset Then                     ' else malformed tag
                        imgOffset = imgOffset + aPtr + inStream(aPtr + 6&) + 11&
                        ' ^^ +11 = 6-10 in header above + 6 byte frame header
                        ReDim Preserve tImages(0 To lCount)
                        With tImages(lCount)
                            .Category = inStream(aPtr + 10&)
                            .Start = imgOffset
                            .length = lSize - (imgOffset - aPtr - 6&) ' determine actual size of just image data
                            If .length + .Start > aPtr + lSize + 6& Or .length < 57& Then lCount = lCount - 1& ' abort this frame
'                            Debug.Print "(3v2)PIC found "
'                            Debug.Print vbTab; "Size: "; .Length; " found at pos: "; .Start; inStream(aPtr + 9); inStream(5)
                        End With
                        lCount = lCount + 1&
                    End If
                End If
            End If
        End If
        aPtr = aPtr + lSize + 6&
    Loop
    pvParse3v2 = lCount

End Function

Private Function pvParse3v3(inStream() As Byte, tImages() As MP3Offsets) As Long

    ' http://www.id3.org/id3v2.3.0
    
    ' id3v2.3.0 header block format
    ' offset 0 : 3 byte magic number: ID3
    ' offset 3 : 1 byte major version
    ' offset 4 : 1 byte revision
    ' offset 5 : 1 byte flag   (bit 7=unsync'd, 6=extended header)
    ' offset 6 : 4 byte sync-safe length

    Dim lMax As Long, aPtr As Long, lSize As Long
    Dim lValue As Long, imgOffset As Long, lCount As Long
    Const tagAPIC As Long = 1128878145                  ' only care about these frames
    
    ' frame header (10 bytes always)
    ' aPtr + 0 : 4 byte header type (i.e., APIC)
    ' aPtr + 4 : 4 byte header size as big endian long
    ' aPtr + 8 : 1 byte flag (not important for this parser)
    ' aPtr + 9 : 1 byte flag (includes whether frame is encrypted and/or comnpressed)
    
    aPtr = 10&
    If (inStream(5) And &H40) Then          ' extended header (length is not sync-safe)
        aPtr = aPtr + pvReverseLong4(inStream(), aPtr)
        If aPtr < 10& Then Exit Function    ' indicates malformed or my parsing engine is wrong
    End If
    lMax = pvSyncSafeLong(inStream(), 6&)
    If (inStream(5) And &H80) Then lMax = lMax - pvSynchronizeTag(inStream(), aPtr, lMax - 10&)
    If lMax < aPtr Then Exit Function
    
    Do While aPtr < lMax
        If inStream(aPtr) = 0& Then Exit Do             ' padding; done
        For lValue = aPtr To aPtr + 3&                  ' validate header type (ASCII A-Z 0-9)
            Select Case inStream(lValue)
            Case 65 To 90, 48 To 57
            Case 0
                If lValue < aPtr + 3& Then Exit Do      ' if not 4th character then abort
                ' these shouldn't really be allowed. Per specs frame names are 4 characters, not 3 + null
                ' but since we are not processing them anyway; just let it go
            Case Else
                Exit Do                                 ' malformatted id3 tag
            End Select
        Next
        lSize = pvReverseLong4(inStream(), aPtr + 4&)   ' and its overall size
        If lSize < 0& Then Exit Do                      ' something is corrupted shouldn't have negative values
        If lSize + aPtr <= lMax Then
            CopyMemory lValue, inStream(aPtr), 4&       ' get frame type
            If lValue = tagAPIC And lSize > 0& Then     ' zero-length frames are not uncommon
                ' APIC header
                ' aPtr + 10     : 1 byte text type (0=ASCII, 1=unicode with BOM)
                ' aPtr + 11     : n bytes image MIME type (null terminated non-unicode string)
                ' aPtr + 11+n   : 1 byte image category (i.e., album front, back, etc)
                ' aPtr + 12+n   : t bytes image description (depends on text type byte)
                ' aptr + 12+n+t : binary data for the image
                '   Note: if MIME type is "-->" then image data is a URL & excluded from processing here
                If (inStream(aPtr + 9&) And &HC0) = 0& Then             ' else compressed and/or encrypted
                    imgOffset = lstrlen(VarPtr(inStream(aPtr + 11&)))   ' get length of MIME type
                    If imgOffset = 3& Then                              ' malformed? should be MIME type or --> or nothing
                        If inStream(aPtr + 12&) = 45 Then imgOffset = -1& ' do not process URL links
                    End If
                    If imgOffset = 0& Or imgOffset > 2& Then
                        imgOffset = imgOffset + aPtr + 13&   ' 13 = 10 header, 1 null terminator, 1 text type, 1 category
                        ' determine length of image description (should be max of 64)
                        Select Case inStream(aPtr + 10&)
                            Case 0: lValue = lstrlen(VarPtr(inStream(imgOffset))) + 1&
                            Case 1: lValue = lstrlenW(ByVal VarPtr(inStream(imgOffset))) * 2& + 2&
                            Case Else: imgOffset = 0&
                        End Select
                        If imgOffset Then                               ' else malformated id3 tag
                            ReDim Preserve tImages(0 To lCount)
                            With tImages(lCount)
                                .Category = inStream(imgOffset - 1&)
                                .Start = imgOffset + lValue
                                .length = lSize - (.Start - aPtr - 10&) ' determine actual size of just image data
                                If .length + .Start > aPtr + lSize + 10& Or .length < 57& Then lCount = lCount - 1& ' abort this frame
'                                Debug.Print "(3v3)APIC found "
'                                Debug.Print vbTab; "Size: "; .Length; " found at pos: "; .Start; inStream(aPtr + 9); inStream(5)
                            End With
                            lCount = lCount + 1&
                        End If
                    End If
                End If
            End If
        End If
        aPtr = aPtr + lSize + 10&
    Loop
    pvParse3v3 = lCount

End Function

Private Function pvParse3v4(inStream() As Byte, tImages() As MP3Offsets) As Long

    ' http://www.id3.org/id3v2.4.0-structure
    
    ' id3v2.4.0 header block format
    ' offset 0 : 3 byte magic number: ID3
    ' offset 3 : 1 byte major version
    ' offset 4 : 1 byte revision
    ' offset 5 : 1 byte flag   (bit 7=unsync'd, 6=extended header, 4=footer)
    ' offset 6 : 4 byte sync-safe length

    Dim lMax As Long, aPtr As Long, lSize As Long, lSizeOffset As Long
    Dim lValue As Long, imgOffset As Long, lCount As Long
    Const tagAPIC As Long = 1128878145                  ' only care about these frames
    
    ' frame header (10 bytes always)
    ' aPtr + 0 : 4 byte header type (i.e., APIC)
    ' aPtr + 4 : 4 byte header size as sync-safe big endian long
    ' aPtr + 8 : 1 byte flag (not important for this parser)
    ' aPtr + 9 : 1 byte flag (includes whether frame is unsync'd, encrypted and/or comnpressed + other stuff)
    '   if data length bit set then 4 bytes added to end of frame
    '   if group bit set then 1 byte added to end of frame
    '   if encryption bit set then 1 byte added to end of frame (note that encrypted frames are not processed here)
    
    aPtr = 10&
    If (inStream(5) And &H40) Then          ' extended header (length is sync-safe)
        aPtr = aPtr + pvSyncSafeLong(inStream(), aPtr)
    End If
    lMax = pvSyncSafeLong(inStream(), 6&)
    If (inStream(5) And &H10) Then lMax = lMax - 10& ' exclude processing the footer
    If (inStream(5) And &H80) Then lMax = lMax - pvSynchronizeTag(inStream(), aPtr, lMax - 10&)
    If lMax < aPtr Then Exit Function
    
    Do While aPtr < lMax
        If inStream(aPtr) = 0& Then Exit Do             ' padding; done
        For lValue = aPtr To aPtr + 3&                  ' validate header type (ASCII A-Z 0-9)
            Select Case inStream(lValue)
            Case 65 To 90, 48 To 57
            Case 0
                If lValue < aPtr + 3& Then Exit Do      ' if not 4th character then abort
                ' these shouldn't really be allowed. Per specs frame names are 4 characters, not 3 + null
                ' but since we are not processing them anyway; just let it go
            Case Else
                Exit Do                                 ' malformatted id3 tag
            End Select
        Next
        
        lSize = pvSyncSafeLong(inStream(), aPtr + 4&)   ' and its overall size
        If lSize + aPtr <= lMax Then
            CopyMemory lValue, inStream(aPtr), 4&       ' get frame type
            If lValue = tagAPIC And lSize > 0& Then     ' zero-length frames are not uncommon
                ' APIC header (no change from v2.0)
                ' aPtr + 10     : 1 byte text type (0=ASCII, 1=unicode with BOM, 2=unicode w/o BOM, 3=UTF-8)
                ' aPtr + 11     : n bytes image MIME type (null terminated non-unicode string)
                ' aPtr + 11+n   : 1 byte image category (i.e., album front, back, etc)
                ' aPtr + 12+n   : t bytes image description (depends on text type byte)
                ' aptr + 12+n+t : binary data for the image
                '   Note: if MIME type is "-->" then image data is a URL & excluded from processing here
                If (inStream(aPtr + 9&) And &HC) = 0& Then  ' else compressed and/or encrypted
                    If (inStream(aPtr + 9&) And &H2) Then   ' frame is unsync'd
                        lSizeOffset = pvSynchronizeTag(inStream(), aPtr + 10&, lSize)
                    Else
                        lSizeOffset = 0&
                    End If
                    If (inStream(aPtr + 9&) And &H40) Then lSizeOffset = lSizeOffset - 1&
                    ' ^^ grouping flag; extra byte added to frame
                    If (inStream(aPtr + 9&) And &H1) Then lSizeOffset = lSizeOffset - 4&
                    ' ^^ data length flag; extra 4 byets added to frame
                    
                    imgOffset = lstrlen(VarPtr(inStream(aPtr + 11&)))   ' get length of MIME type
                    If imgOffset = 3& Then                              ' malformed? should be MIME type or --> or nothing
                        If inStream(aPtr + 12&) = 45 Then imgOffset = -1& ' do not process URL links
                    End If
                    If imgOffset = 0& Or imgOffset > 2& Then            ' else can't be an official MIME value
                        imgOffset = imgOffset + aPtr + 13&   ' 13 = 10 header, 1 null terminator, 1 text type, 1 category
                        ' determine length of image description (max of 64)
                        Select Case inStream(aPtr + 10&)
                        Case 0, 3                                       ' single null terminator
                            lValue = lstrlen(VarPtr(inStream(imgOffset))) + 1&
                        Case 1, 2                                       ' double null terminator
                            lValue = lstrlenW(ByVal VarPtr(inStream(imgOffset))) * 2& + 2&
                        Case Else
                            imgOffset = 0&
                        End Select
                        If imgOffset Then
                            ReDim Preserve tImages(0 To lCount)
                            With tImages(lCount)
                                .Category = inStream(imgOffset - 1&)
                                .Start = imgOffset + lValue
                                .length = lSize - (.Start - aPtr - 10&) - lSizeOffset ' determine actual size of just image data
                                If .length + .Start > aPtr + lSize + 10& - lSizeOffset Or .length < 57& Then lCount = lCount - 1&                                  ' abort this frame
'                                Debug.Print "(3v4)APIC found "
'                                Debug.Print vbTab; "Size: "; .Length; " found at pos: "; .Start; inStream(aPtr + 9); inStream(5), lSizeOffset; lSize
                            End With
                            lCount = lCount + 1&
                        End If
                    End If
                End If
            End If
        End If
        aPtr = aPtr + lSize + 10&
    Loop
    pvParse3v4 = lCount

End Function

Private Function pvParseWMA(inStream() As Byte, tImages() As MP3Offsets) As Long

    ' WMA mini-parser
    ' http://msdn.microsoft.com/en-us/library/bb643323.aspx

    Dim aPtr As Long, sGUID As String, lCount As Long
    Dim lSize As Long, lMax As Long, nrHdrObj As Long
    
    ' ASF header introduction
    ' offset 0  : 16 byte guid
    ' offset 16 : 8 byte block size
    
    ' First header is the Object Header
    ' 24 byte introduction
    ' offset 24 : 4 byte number of blocks within this header
    ' offset 28 : 1 byte reserved (value should always be 1)
    ' offset 29 : 1 byte reserved (value should always be 2)
    ' -- of all the types of blocks, we are only concerned with the Extended header & Extended content
    
    
    ' WMA format has GUIDs as block/section headers. Here are one's were intereseted in
    Const tagASF_Header_Extension_Object As String = "{5FBF03B5-A92E-11CF-8EE3-00C00C205365}"
    Const tagASF_Extended_Content_Description_Object = "{D2D0A440-E307-11D2-97F0-00A0C95EA850}"
    
    CopyMemory lMax, inStream(16), 4&           ' the overall size of the header
    CopyMemory nrHdrObj, inStream(24), 4&
    If nrHdrObj < 1& Then Exit Function
    
    aPtr = 30&: lMax = lMax - 24&
    sGUID = String$(38, vbNullChar)
    For nrHdrObj = 0 To nrHdrObj - 1&
        ' get the GUID and the size of the block. Abort if any blocks are over 2gb
        If StringFromGUID2(VarPtr(inStream(aPtr)), StrPtr(sGUID), 39&) = 0& Then Exit For
        CopyMemory lSize, inStream(aPtr + 20&), 4&: If lSize <> 0& Then Exit For
        CopyMemory lSize, inStream(aPtr + 16&), 4& ' size of the block
        'Debug.Print "guid "; sGUID; lSize
        
        Select Case sGUID
        
        Case tagASF_Extended_Content_Description_Object
            If pvParseWMA_ExtContent(lSize + aPtr, aPtr + 24&, inStream(), tImages(), lCount) = False Then Exit Function
        Case tagASF_Header_Extension_Object
            Call pvParseWMA_ExtHeader(lSize + aPtr, aPtr + 24&, inStream(), tImages(), lCount)
        Case Else
            ' do nothing; skip the section
        End Select
        
        aPtr = aPtr + lSize
        If aPtr > lMax Then Exit For
    Next
    pvParseWMA = lCount

End Function

Private Function pvParseWMA_ExtContent(Ceiling As Long, Offset As Long, inStream() As Byte, tImages() As MP3Offsets, ImgCount As Long) As Boolean

    ' Extended Content header
    ' 24 byte introduction
    ' + 0 bytes : 2 byte count of descriptors
    ' each descriptor will be in this format. No padding, all are consecutive
    ' + 0 : 2 byte description length
    ' + 2 : unicode descriptor name
    ' + 2 + name length : 2 byte data type
    ' + 4 + name length : 2 byte data length
    ' + 8 + name length : the data in format defned by parsed data type
    '   -- of all the possible records, the only ones we care about are WM/Picture
    
    ' The WM/Picture record looks like this
    '   + 0 : 1 byte image category. Same 20+ categories that ID3 uses for MP3s
    '   + 1 : 4 byte image data length
    '   + 5 : double null terminated, unicode MIME type
    '   + 5 + MIME size : double null terminated, unicode image description
    '   + 5 + MIME size + Image description : the image binary data
    
    Dim lRecs As Long, lSize As Long
    Dim sDescriptor As String
    Const tagPicture As String = "WM/Picture" & vbNullChar
    
    CopyMemory lRecs, inStream(Offset), 2&
    Offset = Offset + 2&
    For lRecs = 0& To lRecs - 1&
        lSize = 0&: CopyMemory lSize, inStream(Offset), 2&      ' record's name length
        sDescriptor = String$(lSize \ 2&, vbNullChar)           ' get the name
        CopyMemory ByVal StrPtr(sDescriptor), inStream(Offset + 2&), lSize - (lSize And 1&)
        Offset = Offset + lSize + 2&
        If sDescriptor = tagPicture Then                        ' is it WM/Picture?
            CopyMemory lSize, inStream(Offset), 2&              ' get data type
            If lSize = 1& Then
                CopyMemory lSize, inStream(Offset + 2&), 2&     ' data size
                Offset = Offset + 4&
                ReDim Preserve tImages(0 To ImgCount)
                With tImages(ImgCount)
                    .Category = inStream(Offset)                ' cache image category (same as MP3)
                    CopyMemory .length, inStream(Offset + 1&), 4& ' cache length of image
                    ' skip the MIME, double null terminated string
                    lSize = lstrlenW(ByVal VarPtr(inStream(Offset + 5&))) * 2& + 2&
                    Offset = Offset + lSize + 5& ' 5 = cat byte + size bytes
                    ' skip the image description, double null terminated string
                    lSize = lstrlenW(ByVal VarPtr(inStream(Offset))) * 2& + 2&
                    .Start = Offset + lSize                     ' cache where image starts
                    Offset = .Start + .length                   ' move pointer to end of block
                    If Offset > Ceiling Then ImgCount = ImgCount - 1&
                End With
                lSize = -1&
                ImgCount = ImgCount + 1&
            End If
        End If
        If lSize > -1& Then ' skipping this record
            CopyMemory lSize, inStream(Offset + 2&), 2&         ' record's data length
            Offset = Offset + lSize + 4&
        End If
        If Offset > Ceiling Then Exit Function
    Next
    pvParseWMA_ExtContent = True
    
End Function

Private Function pvParseWMA_ExtHeader(Ceiling As Long, Offset As Long, inStream() As Byte, tImages() As MP3Offsets, ImgCount As Long) As Boolean

    ' Extended header
    ' 24 byte introduction
    ' +  0 : 16 byte GUID & must be 5FBF03B5-A92E-11CF-8EE3-00C00C205365
    ' + 16 : 2 byte reserved (value should always be 6)
    ' + 18 : 4 byte data length
    ' + 22 : the extended header data
    '   - this data consists of 0 or more consecutive blocks, in any order
    '   - of all the blocks that can exist, we only care about the MetaData & Extended Content blocks
    
    ' MetaData header
    ' 24 byte introduction
    ' + 0 bytes : 2 byte record count
    ' each metadata record will be in this format. No padding, all records are consecutive
    '   + 0  : 2 byte language list index
    '   + 2  : 2 byte stream number
    '   + 4  : 2 byte metadata name length
    '   + 6  : 2 byte data type flag (byte array, integer, long, etc)
    '   + 8  : 4 byte data length
    '   + 12 : metadata name where number of bytes defined @ offset 4
    '   + 12 + name length : the metadata in format defined @ offset 6
    '   total record length is 12 + value @ offset 4 + value @ offset 8
    '   -- of all the possible records, the only ones we care about are WM/Picture
    
    ' The WM/Picture metadata record looks like this
    '   + 0 : 1 byte image category. Same 20+ categories that ID3 uses for MP3s
    '   + 1 : 4 byte image data length
    '   + 5 : double null terminated, unicode MIME type
    '   + 5 + MIME size : double null terminated, unicode image description
    '   + 5 + MIME size + Image description : the image binary data
    
    Dim sGUID As String, sTag As String, lExtSize As Long
    Dim lSize As Long, lRecs As Long, lRecSize As Long
    
    Const tagASF_Reserved_1 As String = "{ABD3D211-A9BA-11CF-8EE6-00C00C205365}"
    Const tagASF_Metadata_Library_Object As String = "{44231C94-9498-49D1-A141-1D134E457054}"
    Const tagASF_Extended_Content_Description_Object = "{D2D0A440-E307-11D2-97F0-00A0C95EA850}"
    ' not a header, but a section 'name' within a metadata record
    Const tagPicture As String = "WM/Picture" & vbNullChar
    

    sGUID = String$(38, vbNullChar)
    Call StringFromGUID2(VarPtr(inStream(Offset)), StrPtr(sGUID), 39&)
    If sGUID <> tagASF_Reserved_1 Then Exit Function  ' per specs, this must be the next 16 bytes
    CopyMemory lExtSize, inStream(Offset + 18&), 4&   ' size of extended header
    Offset = Offset + 22&                             ' move pointer amount of bytes read in block
    lExtSize = lExtSize + Offset                      ' set max read size
    
    Do While Offset < lExtSize
        If StringFromGUID2(VarPtr(inStream(Offset)), StrPtr(sGUID), 39&) = 0& Then Exit Function
        CopyMemory lSize, inStream(Offset + 20&), 4&: If lSize <> 0& Then Exit Function
        CopyMemory lSize, inStream(Offset + 16&), 4&  ' size of extended header sub-block
'       Debug.Print vbTab; sGUID; lSize
        
        If sGUID = tagASF_Metadata_Library_Object Then
            CopyMemory lRecs, inStream(Offset + 24&), 2&: If lRecs = 0& Then Exit Function
            lRecSize = Offset + lSize                   ' set max read size & read nr of meta records
            Offset = Offset + 26&                       ' move pointer amount of bytes read in block
            For lRecs = 0& To lRecs - 1&
                lSize = 0&: CopyMemory lSize, inStream(Offset + 6&), 2&   ' data type
                If lSize = 1& Then          ' byte array which is what we are after
                    CopyMemory lSize, inStream(Offset + 4&), 2&   ' record's name length
                    If lSize Then
                        sTag = String$(lSize \ 2&, vbNullChar)     ' get the name
                        CopyMemory ByVal StrPtr(sTag), inStream(Offset + 12&), lSize - (lSize And 1&)
                        If sTag = tagPicture Then                  ' is it WM/Picture?
                            Offset = Offset + 12& + lSize          ' move pointer to where it starts
                            
                            ReDim Preserve tImages(0 To ImgCount)
                            With tImages(ImgCount)
                                .Category = inStream(Offset)       ' cache image category (same as MP3)
                                CopyMemory .length, inStream(Offset + 1&), 4& ' cache length of image
                                ' skip the MIME, double null terminated string
                                lSize = lstrlenW(ByVal VarPtr(inStream(Offset + 5&))) * 2& + 2&
                                Offset = Offset + lSize + 5& ' 5 = cat byte + size bytes
                                ' skip the image description, double null terminated string
                                lSize = lstrlenW(ByVal VarPtr(inStream(Offset))) * 2& + 2&
                                .Start = Offset + lSize            ' cache where image starts
                                Offset = .Start + .length          ' move pointer to end of block
                                If Offset > lRecSize Then ImgCount = ImgCount - 1&
                            End With
                            
                            ImgCount = ImgCount + 1&
                            lSize = -1&                            ' flag indicating block 100% processed
                        End If
                    End If
                End If
                If lSize > -1& Then ' skipping this metadata record
                    lSize = 0&: CopyMemory lSize, inStream(Offset + 4&), 2& ' record's name length
                    Offset = Offset + lSize + 12&
                    CopyMemory lSize, inStream(Offset - lSize - 4&), 4& ' record's data length
                    Offset = Offset + lSize
                End If
                If Offset > lRecSize Then Exit Function
            Next
        Else
            Offset = Offset + lSize
            ' haven't seen any extended content contained within the extended header; but
            ' I couldn't find any documentation to say it is not possible -- so, just in case:
            If sGUID = tagASF_Extended_Content_Description_Object Then
                If pvParseWMA_ExtContent((Offset), Offset - lSize + 24&, inStream(), tImages(), ImgCount) = False Then Exit Function
            End If
        End If
        If Offset > Ceiling Then Exit Function
    Loop
    pvParseWMA_ExtHeader = True

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

