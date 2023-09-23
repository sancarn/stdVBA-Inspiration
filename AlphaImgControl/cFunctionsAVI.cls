VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionsAVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Play AVI Files ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' Note that AVI files are iffy at best, when using the avifil32 dll. DirectShow appears a better choice
' However, DirectShow support requires some really complicated solutions if not drawing directly to a window.

' In order to play AVIs, the AVI will be extracted to the user's Temp or MyDocuments folder and
'   played from there. These files will be deleted when no longer needed; except if you terminate
'   the project abnormally (while control is uncompiled) which leaves the file remaining

' So, if an AVI fails to load properly, it can mean one of four things
' 1) The file is not an AVI or is corrupted
' 2) The AVI does not contain images. Expected VIDS block not found
' 3) You do not have the correct codec installed to play the AVI
'       -- codecs come in a few flavors. The one used by Media Player may not be same as used for avifil32
'       -- to determine codec required, follow these easy steps
'       a. Open AVI file in notepad or wordpad
'       b. search for: VIDS
'       c. If found, the next 4 characters is the codec needed and known as the FourCC
'       d. You can use sites like www.fourcc.org to find & install codecs. But beware
' 4) The AVI loads but image is wrong or corrupted. Find a better codec

' Storing an AVI in an AlphaImgCtl is probably not a good idea for a couple reasons
' 1) file can be quite large, depending on compression codec used
' 2) if project/app passed to another pc, that pc needs to have the codec installed to view the AVI
' You may be better off converting the AVI to an animated GIF or animated PNG
'   - there are several free apps on the net that will convert AVI to GIF
'   - this control can save an AVI to either animated GIF or animated PNG
'       -- animated GIF conversion can cause color loss if AVI is 16 bpp or greater
'       -- animated PNG conversion will result in no color loss
'   - See SavePictureGdip.RTF that came with this project

Private Declare Function GdipCreateBitmapFromGdiDib Lib "GdiPlus.dll" (bminfohdr As Any, ByVal pBits As Long, ByRef pbitmap As Long) As Long

Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempPathW Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function GetTempFileNameA Lib "kernel32.dll" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempFileNameW Lib "kernel32.dll" (ByVal lpszPath As Long, ByVal lpPrefixString As Long, ByVal wUnique As Long, ByVal lpTempFileName As Long) As Long
Private Declare Function SHGetFolderPathA Lib "shfolder.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetFolderPathW Lib "shfolder.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long

Private Type TAVISTREAMINFO ' this is the ANSI version (however only used herein to retrieve rate & scale)
    fccType As Long
    fccHandler As Long
    dwFlags As Long         '/* Contains AVITF_* flags */
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long ' /* dwRate / dwScale == samples/second */
    dwStart As Long
    dwLength As Long '; /* In units above... */
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As RECTI
    dwEditCount As Long
    dwFormatChangeCount As Long
    szName(0 To 63) As Byte
End Type

Private Declare Function AVIStreamOpenFromFile Lib "avifil32.dll" Alias "AVIStreamOpenFromFileA" (ppavi As Any, ByVal szFile As String, ByVal fccType As Long, ByVal lParam As Long, ByVal Mode As Long, pclsidHandler As Any) As Long
Private Declare Function AVIStreamOpenFromFileW Lib "avifil32.dll" (ppavi As Any, ByVal szFile As Long, ByVal fccType As Long, ByVal lParam As Long, ByVal Mode As Long, pclsidHandler As Any) As Long
Private Declare Function AVIStreamStart Lib "avifil32.dll" (pavi As Any) As Long
Private Declare Function AVIStreamRelease Lib "avifil32.dll" (pavi As Any) As Long
Private Declare Function AVIStreamLength Lib "avifil32.dll" (pavi As Any) As Long
Private Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (pavi As Any, lpbiWanted As Any) As Long
Private Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (pg As Any) As Long
Private Declare Function AVIStreamSampleToTime Lib "avifil32.dll" (pavi As Any, ByVal lSample As Long) As Long
Private Declare Sub AVIStreamInfo Lib "avifil32.dll" Alias "AVIStreamInfoA" (pavi As Any, psi As TAVISTREAMINFO, ByVal lSize As Long)
Private Declare Function AVIStreamRead Lib "avifil32.dll" (pavi As Any, ByVal lStart As Long, ByVal lSamples As Long, lpBuffer As Any, ByVal cbBuffer As Long, plBytes As Long, plSamples As Long) As Long
Private Declare Function AVIStreamReadFormat Lib "avifil32.dll" (pavi As Any, ByVal lPos As Long, lpFormat As Any, ByRef lpcbFormat As Long) As Long
Private Declare Function AVIStreamGetFrame Lib "avifil32.dll" (pg As Any, ByVal lPos As Long) As Long

Private m_BitmapInfo As BITMAPINFOHEADER    ' frame attributes
Private m_FileName As String                ' file name avi is being accessed from
Private m_Stream As Long                    ' AVI stream object pointer
Private m_FrameObject As Long               ' AVI frame object pointer
Private m_Frames As Long                    ' number of AVI frames
Private m_FrameNr As Long                   ' current frame number
Private m_State As Long                     ' 1=initialized, 2=owned file
Private m_Token As cGDIpToken               ' GDI+ cGDIpToken reference

Public Function IsAVIResource(inStream() As Byte, cImageData As cGDIpMultiImage) As Boolean

    ' function verifies the passed array is an AVI format; and attempts to load 1st frame

    If (m_State And 1) = 0& Then Exit Function
    If m_Token.Token = 0& Then Exit Function
    
    Dim hHandle As Long
    If pvIsAVIformat(inStream()) Then
        Call pvUnload
        If pvCacheAVItoFile(inStream) Then
            If pvLoadAVI() = True Then
                hHandle = Me.MoveToFrame(m_Frames)
                If hHandle Then
                    pvCacheFormat hHandle, cImageData
                    IsAVIResource = True
                End If
            End If
        End If
    End If

End Function

Public Function IsAVIResourceFile(FileName As String, cImageData As cGDIpMultiImage) As Boolean
    
    ' function verifies the passed file is an AVI format; and attempts to load 1st frame
    
    If (m_State And 1) = 0& Then Exit Function
    If m_Token.Token = 0& Then Exit Function
    
    Dim fHandle As Long, hHandle As Long
    fHandle = GetFileHandle(FileName, False)
    If fHandle = INVALID_HANDLE_VALUE Then Exit Function
    If GetFileSize(fHandle, 0&) > 60& Then
        Dim bData(0 To 59) As Byte, lRead As Long
        SetFilePointer fHandle, 0&, 0&, 0&
        ReadFile fHandle, bData(0), 60&, lRead, ByVal 0&
        CloseHandle fHandle
        If pvIsAVIformat(bData()) Then
            Call pvUnload
            m_FileName = FileName
            If pvLoadAVI() = True Then
                If cImageData Is Nothing Then           ' testing, not caching
                    IsAVIResourceFile = True
                Else
                    hHandle = Me.MoveToFrame(m_Frames)
                    If hHandle Then
                        pvCacheFormat hHandle, cImageData
                        IsAVIResourceFile = True
                    End If
                End If
            End If
        End If
    Else
        CloseHandle fHandle
    End If
    
End Function

Public Property Get FrameCount() As Long
    FrameCount = m_Frames   ' returns number of AVI frames
End Property

Public Function SaveAsAVI(returnObject As Variant, SourceHandle As Long, ByVal returnMedium As SaveAsMedium, SS As SAVESTRUCT) As Long

    ' This routine does not create AVI formats. It simply saves existing AVI to a target destination

    Dim lResult As Long, imgData() As Byte, hBitmap As Long
    Dim tPic As StdPicture, tmpDO As DataObject
    Dim tObject As Object, tGDIpImage As GDIpImage
    
    If SS.reserved2 = 0& Then Exit Function
    CopyMemory tObject, SS.reserved2, 4&
    Set tGDIpImage = tObject
    CopyMemory tObject, 0&, 4&
    lResult = tGDIpImage.ExtractImageData(imgData())
    
    If returnMedium = saveTo_File And lResult <> 0& Then
        WriteFile CLng(returnObject), imgData(0), UBound(imgData) + 1&, lResult, ByVal 0&
        SaveAsAVI = (lResult > UBound(imgData))
    ElseIf returnMedium = saveTo_Array And lResult <> 0& Then
        modCommon.MoveArrayToVariant returnObject, imgData(), True
        SaveAsAVI = True
    ElseIf returnMedium = saveTo_GDIplus And lResult <> 0& Then
        Set returnObject = LoadImage(imgData(), , , True)
        If returnObject.Handle = 0& Then Set returnObject = Nothing Else SaveAsAVI = returnMedium
    Else
        GdipCreateHBITMAPFromBitmap SourceHandle, hBitmap, -1&
        If hBitmap Then
            SaveAsAVI = True
            Set tPic = modCommon.HandleToStdPicture(hBitmap, vbPicTypeBitmap)
            If tPic Is Nothing Then
                DeleteObject hBitmap: SaveAsAVI = False
            ElseIf returnMedium = saveTo_stdPicture Then
                Set returnObject = tPic
                SaveAsAVI = (tPic.Handle <> 0&)
            ElseIf returnMedium = saveTo_Clipboard Then
                Clipboard.SetData tPic
                If g_ClipboardFormat And lResult <> 0& Then modCommon.SetClipboardCustomFormat imgData(), g_ClipboardFormat
            ElseIf returnMedium = saveTo_DataObject Then
                Set tmpDO = returnObject
                tmpDO.SetData tPic, vbCFBitmap
                If g_ClipboardFormat And lResult <> 0& Then tmpDO.SetData imgData(), g_ClipboardFormat
                Set tmpDO = Nothing
            End If
            Set tPic = Nothing
        End If
    End If

End Function


Private Sub pvUnload()

   ' clean up routine
   
   If m_FrameObject Then ' close frame interface
        Call AVIStreamGetFrameClose(ByVal m_FrameObject)
        m_FrameObject = 0&
   End If

   If m_Stream Then ' release avi stream
        AVIStreamRelease ByVal m_Stream
        m_Stream = 0&
   End If
   m_Frames = 0&
   m_FrameNr = -1&
   If (m_State And 2&) Then ' destroy file if we created it
        If m_FileName <> vbNullString Then DeleteFileEx m_FileName
        m_FileName = vbNullString
        m_State = m_State Xor 2&
    End If
   
End Sub

Private Function pvLoadAVI() As Boolean
    
    ' function attempts to load the AVI
    
    Dim lSize As Long, bFailure As Boolean
   
    Const STREAM_TYPE_VIDEO = &H73646976 ' "VIDS"
    Const OF_READ = &H0
    Const OF_SHARE_DENY_WRITE As Long = &H20
   
    If g_UnicodeSystem Then
        bFailure = (AVIStreamOpenFromFileW(m_Stream, StrPtr(m_FileName), STREAM_TYPE_VIDEO, 0&, OF_READ Or OF_SHARE_DENY_WRITE, ByVal 0&))
    Else
        bFailure = (AVIStreamOpenFromFile(m_Stream, m_FileName, STREAM_TYPE_VIDEO, 0&, OF_READ Or OF_SHARE_DENY_WRITE, ByVal 0&))
    End If
    If Not bFailure Then
        ' open frames
        m_FrameObject = AVIStreamGetFrameOpen(ByVal m_Stream, ByVal 0&)
        If m_FrameObject Then
            m_Frames = AVIStreamLength(ByVal m_Stream) ' get number of frames
            If m_Frames Then Call AVIStreamReadFormat(ByVal m_Stream, 0, ByVal 0&, lSize) ' Get the AVI format:
        End If
    End If
    If lSize < 40& Then Call pvUnload Else pvLoadAVI = True

End Function

Public Function MoveToFrame(ByVal FrameNr As Long) As Long

    ' navigation to other frames. Passed FrameNr is zero-bound

    Dim tHandle As Long
    If m_Frames = 0& Then Exit Function
    
    If FrameNr < 0& Then
        FrameNr = m_Frames - 1&
    ElseIf FrameNr >= m_Frames Then
        FrameNr = 0&
    End If
    m_FrameNr = FrameNr
    MoveToFrame = pvGetFrameImage(m_FrameNr)
    
End Function

Private Function pvCacheAVItoFile(inStream() As Byte) As Boolean
    
    ' find a path/file we can dump our AVI bytes too
    ' The AVI dll library uses files
    
    Const ssfPERSONAL  As Long = 5&

    Dim strBuffer As String, strFile As String
    Dim lResult As Long, hHandle As Long
    
    On Error Resume Next
    
    ' 1. Try GetTempPath first
    strBuffer = Space$(MAX_PATH)
    If g_UnicodeSystem Then
        lResult = GetTempPathW(MAX_PATH, StrPtr(strBuffer))
        If lResult > MAX_PATH Then
            strBuffer = Space$(lResult + 1&)
            lResult = GetTempPathW(lResult, StrPtr(strBuffer))
        End If
    Else
        lResult = GetTempPath(MAX_PATH, strBuffer)
        If lResult > MAX_PATH Then
            strBuffer = Space$(lResult + 1&)
            lResult = GetTempPath(lResult, strBuffer)
        End If
    End If
    If lResult Then
        ' GetTempPath does not verify the path exists; nor does it verify user has write-access
        ' We need to make those verifications
        strFile = Left$(strBuffer, lResult)
        If modCommon.FileOrFolderExists(strFile) Then
            ' path exists, can we write to it? Create random file name
            strBuffer = strFile & "_aviTestFile" & ObjPtr(Me) & ".tmp"
            lResult = modCommon.GetFileHandle(strBuffer, True)
            If lResult = INVALID_HANDLE_VALUE Then
                strFile = vbNullString              ' nope can't write to it
            Else
                CloseHandle lResult                 ' all is ok
                DeleteFileEx strBuffer
            End If
        Else
            strFile = vbNullString                  ' temp path doesn't exist; that sucks
        End If
    End If
    If strFile = vbNullString Then
    ' If we couldn't get or write to the temp path, then we will default to user's documents folder
    ' 2. Try using the shell script.
        With CreateObject("Shell.Application").NameSpace(ssfPERSONAL).self
            strFile = .Path
        End With
        If Err Then Err.Clear
        If strFile = vbNullString Then
    ' 3. Hmmm. Ok, let's try good old-fashioned API
            strBuffer = Space$(MAX_PATH)
            If g_UnicodeSystem Then
                lResult = SHGetFolderPathW(0&, ssfPERSONAL, 0&, 0&, StrPtr(strBuffer))
            Else
                lResult = SHGetFolderPathA(0, ssfPERSONAL, 0, 0, strBuffer)
            End If
            If lResult = 0& Then
                strFile = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
                If Err Then Err.Clear
            End If
    ' 4. There is no 4th attempt here. If we failed at this point, then need some info to troubleshoot
        End If
    End If
    If strFile = vbNullString Then Exit Function
    
    If Right$(strFile, 1) <> "\" Then
        If Right$(strFile, 1) <> "/" Then strFile = strFile & "\"
    End If
    
    If g_UnicodeSystem Then
        lResult = GetTempFileNameW(StrPtr(strFile), StrPtr("AVI"), 0&, StrPtr(strBuffer))
    Else
        lResult = GetTempFileNameA(strFile, "AVI", 0&, strBuffer)
    End If
    If lResult = 0& Then Exit Function
    
    ' the temp file is created with a .tmp extension.
    ' we want an avi extension, so let's destroy the 0-byte file the api created
    strFile = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    DeleteFileEx strFile
    ' create a new file with the .avi extension
    lResult = InStr(1, strFile, ".tmp", vbTextCompare)
    If lResult Then
        strFile = Replace$(strFile, ".tmp", "_tmp.avi", , , vbTextCompare)
    Else
        strFile = strFile & ".avi"
    End If
    ' and now open it and we'll use that
    hHandle = GetFileHandle(strFile, True)
    If hHandle = INVALID_HANDLE_VALUE Or hHandle = 0& Then Exit Function
    
    WriteFile hHandle, inStream(0), UBound(inStream) + 1&, lResult, ByVal 0&
    CloseHandle hHandle
    
    If (lResult > UBound(inStream)) Then
        m_FileName = strFile
        m_State = m_State Or 2&      ' we own this and it should be deleted when class exits
        pvCacheAVItoFile = True
    Else
        DeleteFileEx strFile         ' failed; could be out of disk space?
    End If

End Function

Private Function pvGetDIBHeaderSize(ByRef inHeader As BITMAPINFOHEADER) As Long
    
    ' function validates the bitmap header is one we are capable of handling
    
    Const BI_BITFIELDS As Long = 3&
    Const BI_RGB As Long = 0&
    
    pvGetDIBHeaderSize = -1&            ' default value = failure
    Select Case inHeader.biBitCount
        Case 1, 4, 8
        Case 16, 24, 32
            inHeader.biClrUsed = 0&
        Case Else
            Exit Function
    End Select
    Select Case inHeader.biCompression
        Case BI_BITFIELDS, BI_RGB
            ' if BI_RLE, it should have been converted before getting here
        Case Else
            Exit Function
    End Select
    Select Case inHeader.biClrUsed
        Case 0 To 256
        Case Else ' not valid bitmap header
            Exit Function
    End Select
    If inHeader.biWidth < 1& Or inHeader.biHeight = 0& Then Exit Function
    
    On Error Resume Next
    inHeader.biWidth = (inHeader.biWidth * Abs(inHeader.biHeight)) \ Abs(inHeader.biHeight)
    If Err Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    If (inHeader.biSize > 39& And inHeader.biPlanes = 1) Then
        With inHeader ' Set palette size if undefined for a paletted image
            If ((.biBitCount < 9) And (.biClrUsed < 1)) Then .biClrUsed = 2 ^ .biBitCount
            pvGetDIBHeaderSize = .biSize + (.biClrUsed * 4) + _
                IIf((.biCompression = BI_BITFIELDS) And (.biSize = Len(inHeader)), 12, 0)
        End With
    End If
End Function

Private Function pvGetFrameInfo(FrameNr As Long, frameObject As Long, streamObject As Long, frameData As BITMAPINFOHEADER) As Long

    ' function returns size & bitcount info via a BitmapInfoHeader

    If frameObject = 0& Then Exit Function
    
    Dim lResult As Long, lMemAddr As Long
    
    frameData.biBitCount = 0&
    Call AVIStreamReadFormat(ByVal streamObject, FrameNr, ByVal 0&, lResult)
    If lResult > 39& Then
    
        lMemAddr = AVIStreamGetFrame(ByVal frameObject, FrameNr)
        If lMemAddr Then
            CopyMemory frameData, ByVal lMemAddr, 40&
            If pvGetDIBHeaderSize(frameData) < 0& Then
                frameData.biBitCount = 0&
            Else
                frameData.biXPelsPerMeter = lResult
                pvGetFrameInfo = lMemAddr
            End If
        End If
    End If

End Function

Private Function pvGetFrameImage(FrameNr As Long) As Long

    ' function gets memory address of FrameNr (passed as zero-bound)
    ' The GDI+ bitmap will be created from that memory address (if no alpha channel is used)
    '   If alpha channel used, then bitmap is created manually
    ' Note that the memory address remains valid unless the frame object is cleared
    '   with a call to AVIStreamGetFrame or the stream is released
    
    If m_FrameObject = 0& Then Exit Function
    
    Dim lResult As Long, hPtr As Long, lHeight As Long, lOffset As Long
    Dim tBMP As BitmapData, tSize As RECTI, bData() As Long
    
    Call AVIStreamReadFormat(ByVal m_Stream, FrameNr, ByVal 0&, lOffset)
    If lOffset < 40& Then Exit Function ' not expected format; should be bitmapinfoheader (min 40 bytes)
    
    hPtr = AVIStreamGetFrame(ByVal m_FrameObject, FrameNr)
    If hPtr = 0& Then Exit Function

    CopyMemory m_BitmapInfo, ByVal hPtr, 40&        ' may be larger than 40, but all we need is 40
    If pvGetDIBHeaderSize(m_BitmapInfo) < 0& Then Exit Function
    
    If m_BitmapInfo.biBitCount = 32 Then
        If m_BitmapInfo.biHeight < 0& Then
            lHeight = -m_BitmapInfo.biHeight
        Else
            lHeight = m_BitmapInfo.biHeight
        End If
        m_BitmapInfo.biSizeImage = lHeight * m_BitmapInfo.biWidth * 4&
        ReDim bData(0 To m_BitmapInfo.biWidth - 1&, 0& To lHeight - 1&)
        CopyMemory bData(0, 0), hPtr + lOffset, m_BitmapInfo.biSizeImage
        tBMP.PixelFormat = modCommon.ValidateAlphaChannel(bData(), 0&)
        If tBMP.PixelFormat <> lvicColor32bpp Then
            ' create bitmap and copy bits manually when alpha channel is used
            If GdipCreateBitmapFromScan0(m_BitmapInfo.biWidth, lHeight, 0&, tBMP.PixelFormat, ByVal 0&, lResult) = 0& Then
                If m_BitmapInfo.biHeight < 0& Then
                    tBMP.stride = m_BitmapInfo.biWidth * 4&
                    tBMP.Scan0Ptr = hPtr + lOffset
                Else
                    tBMP.stride = -m_BitmapInfo.biWidth * 4&
                    tBMP.Scan0Ptr = hPtr + lOffset + tBMP.stride
                End If
                tSize.nWidth = m_BitmapInfo.biWidth: tSize.nHeight = lHeight
                tBMP.Width = tSize.nWidth: tBMP.Height = tSize.nHeight
                If GdipBitmapLockBits(lResult, tSize, ImageLockModeUserInputBuf Or ImageLockModeWrite, tBMP.PixelFormat, tBMP) = 0& Then
                    GdipBitmapUnlockBits lResult, tBMP
                    pvGetFrameImage = lResult: lResult = 0&
                End If
                If lResult Then
                    GdipDisposeImage lResult
                    tBMP.Scan0Ptr = 0&
                End If
            End If
        End If
        Erase bData()
    End If
    If tBMP.Scan0Ptr = 0& Then ' not processed as 32bpp above
        GdipCreateBitmapFromGdiDib ByVal hPtr, hPtr + lOffset, pvGetFrameImage
    End If

End Function

Private Sub pvCacheFormat(gHandle As Long, cImageData As cGDIpMultiImage)

    Dim SeqDelay() As Long, i As Long
    Dim lInterval As Long
    
    ' there are two ways to get the average frame intervals
    ' 1st - use one api & calculate based off of rate & scale
    Dim tAVI As TAVISTREAMINFO
    AVIStreamInfo ByVal m_Stream, tAVI, LenB(tAVI)
    If tAVI.dwScale > 0& Then
        On Error Resume Next
        lInterval = 1000! / CSng(tAVI.dwRate / tAVI.dwScale)
    End If
    If lInterval = 0& Then
        ' 2nd - use a few api calls
        lInterval = AVIStreamStart(ByVal m_Stream) + AVIStreamLength(ByVal m_Stream)
        lInterval = AVIStreamSampleToTime(ByVal m_Stream, lInterval) \ m_Frames
    End If
    cImageData.InitializeContent ObjPtr(Me), lvicPicTypeAVI, m_Frames
    ReDim SeqDelay(0 To 1, 1 To m_Frames)
    For i = 1 To m_Frames
        SeqDelay(0, i) = i - 1&
        SeqDelay(1, i) = lInterval
    Next
    cImageData.SetSequenceDuration SeqDelay()            ' send off & done
    cImageData.CacheSourceInfo m_FileName, gHandle, lvicPicTypeAVI, True, False

End Sub

Private Function pvIsAVIformat(inStream() As Byte) As Boolean

    ' AVIs are a RIFF format witha magic number of AVI+space

    Dim lValue As Long
    Const aviRIFF As Long = 1179011410 ' magic number for RIFF
    Const aviMagicNr As Long = 541677121 ' magic number for AVI RIFF
    
    If UBound(inStream) > 57 Then
        CopyMemory lValue, inStream(0), 4&
        If lValue = aviRIFF Then
            CopyMemory lValue, inStream(8), 4&
            pvIsAVIformat = (lValue = aviMagicNr)
        End If
    End If

End Function

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
    If g_TokenClass.AVIClassesInitiated = True Then m_State = 1&
End Sub

Private Sub Class_Terminate()
   Call pvUnload
   m_Token.RemoveRef
   Set m_Token = Nothing
End Sub


