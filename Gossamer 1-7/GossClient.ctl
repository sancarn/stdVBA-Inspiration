VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl GossClient 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   405
   ScaleWidth      =   405
   ToolboxBitmap   =   "GossClient.ctx":0000
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   70
      Left            =   0
      Top             =   270
      Width           =   340
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Width           =   340
   End
End
Attribute VB_Name = "GossClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'GossClient
'==========
'
'A wrapper for a Winsock control used for an HTTP client connection.
'
'A constituent part of the Gossamer HTTP server control, one or more
'GossClients are used as a control array for active HTTP client
'connections.
'
'This is where the bulk of the HTTP protocol and server behavior is
'implemented.
'
'Changes
'-------
'
'Version: 1.1
'
' o Enhanced parsing of HTTP headers to eliminate spaces and tabs before
'   and after : separator.
' o Marked Public members "hidden" since they're only for use by Gossamer.
' o LogEvent is now RaiseLogEvent, DynamicRequest is now
'   RaiseDynamicRequest.
' o Simplified VDir path handling and enabled legal relative paths (those
'   that stay within VDir).
' o Increased I/O buffer size to VB native I/O maximum (32,767 bytes).
'
'Version: 1.2
'
' o Handle If-Modified-Since headers in GET requests, returning 304 if
'   the file has not changed.
'
'Version: 1.3
'
' o Realized that I'm using binary I/O for static resources and abandoned
'   the inapplicable buffer size parameter.  Handling buffering properly
'   (though it wasn't broken) but more importantly moving to a larger
'   buffer size.
'
'Version: 1.4
'
' o Changed buffer size back down to 8192.  Concerned about possibly
'   blocking on SendData with a slow client.
' o Corrected handling of headers in SendResponse.
' o Added "Accept-Ranges: none" header to responses.
' o Added server headers to responses.
'
'Version: 1.5
'
' o Changed buffer size back to 8192.  Was supposed to have been done by
'   1.4 but somehow didn't actually get done!
' o Bad Content-Length headers were reported as bad Content-Type. Fixed.
' o New gesServerError added for reporting errors like "out of file
'   numbers."
' o Addressed a timing problem where SendComplete could occur before
'   file sending was completely set up.
' o Handle case where static resource (file) was sent, had no leftover
'   partial block to send, and client requested "close."
'
'Version 1.6
'
' o Treat sckWouldBlock as a getWSSoftError: log it as such and
'   then ignore it.
' o Stop setting .Timestamp since it gets set when GE instances are
'   created.
'
'Version: 1.7
'
' o Fixed a nasty (silly) bug in SendResponse() where no space was sent
'   between Status and StatusText.  Amazing this wasn't discovered
'   sooner!
' o Removed sending an extraneous space after header ":" separators.  Not
'   important but a silly waste of bandwidth.
' o In ProcessRequest() upcase both strParts(0) [method] and strParts(2)
'   [HTTP version] from the HTTP request line of text in keeping with the
'   Internet convention of "being liberal in what you accept."  Almost
'   never an issue but a lost-cost thing to do anyway.
' o In SendResponse() ensure that a CRLF gets sent after any ExtraHeaders
'   even if the client (our container) code neglected to.
' o Lowercased the HTML tags in SendCannedFormatted().
'

Private Enum RequestStates
    rqsIdle = 0
    rqsHeadersComplete
    rqsReqLineComplete
    rqsRequestComplete
End Enum

Private Const STATIC_BUFFER_SZ As Long = 8192
Private Const DOUBLE_CRLF As String = vbCrLf & vbCrLf

Private mBlnConnClose As Boolean      'Received a "Connection: close" header.
Private mBlnInUse As Boolean          'In-use status of this GossClient.
Private mBytBuffer() As Byte          'Buffer for reading static resource requested.
Private mBytResponse() As Byte        'Response data from dynamic request.
Private mLngResponseLen As Long       'Valid bytes in mBytResponse.
Private mColReqHeaders As Collection  'Request headers Collection.
                                      'Items are String arrays (0)=Name, (1)=Value.
                                      'Keys are stored UPPERCASED.
Private mIntFile As Integer           'Native I/O file number of static resource requested.
                                      'If non-0 a file is open.
Private mIntFullBlocks As Integer     'Remaining count of full blocks to send in static resource.
Private mIntIndex As Integer          'Index in control array of this GossClient.
Private mIntLastBlockSize As Integer  'Size in bytes of final block in static resource.
Private mLngContentLen                'Request content length from header.
Private mReqState As RequestStates    'Where we are receiving a request.
Private mStrBuffer As String          'Buffered incoming request text.
Private mStrContent As String         'Request Content body.
Private mStrHTTPVersion As String     'Version of request.
Private mStrReqLine As String         'HTTP request line.
Private mStrRespBuffer As String      'We want to buffer responses so we can handle
                                      '  ReqCloseConn = True properly.
Private mLngRespUsed As Long

Public Property Get InUse() As Boolean
Attribute InUse.VB_Description = "Only for use by Gossamer"
Attribute InUse.VB_MemberFlags = "40"
    InUse = mBlnInUse
End Property

Public Sub Accept(ByVal requestID As Long)
Attribute Accept.VB_Description = "Only for use by Gossamer"
Attribute Accept.VB_MemberFlags = "40"
    mBlnInUse = True
    mStrBuffer = ""
    wskClient.Accept requestID
End Sub

Public Sub Init(ByVal IndexValue As Integer)
Attribute Init.VB_Description = "Only for use by Gossamer"
Attribute Init.VB_MemberFlags = "40"
    mIntIndex = IndexValue
End Sub

Public Sub Shutdown()
Attribute Shutdown.VB_Description = "Only for use by Gossamer"
Attribute Shutdown.VB_MemberFlags = "40"
    wskClient_Close
End Sub

Private Sub AppendResp(ByVal Text As String)
    Dim Length As Long
    
    Length = Len(Text)
    If mLngRespUsed + Length > Len(mStrRespBuffer) Then
        If Len(mStrRespBuffer) < 1 Then
            mStrRespBuffer = Space$(Length + 200)
        Else
            mStrRespBuffer = mStrRespBuffer & Space$(Length + 100)
        End If
    End If
    Mid$(mStrRespBuffer, mLngRespUsed + 1, Length) = Text
    mLngRespUsed = mLngRespUsed + Length
End Sub

Private Function ExtractResp() As String
    ExtractResp = Left$(mStrRespBuffer, mLngRespUsed)
    mLngRespUsed = 0
End Function

Private Sub ProcessRequest()
    Dim blnReturnFile As Boolean
    Dim dtSince As Date
    Dim GE As GossEvent
    Dim lngFileBytes As Long
    Dim lngInStr As Long
    Dim sngStatus As Single
    Dim strFile As String
    Dim strKeepAlive As String
    Dim strMIME As String
    Dim strParts() As String
    Dim strRespExtraHeaders As String
    Dim strStatusText As String

    strParts = Split(mStrReqLine, " ", 3)
    strParts(0) = UCase$(strParts(0))
    Set GE = New GossEvent
    With GE
        .EventType = getHTTP
        .IP = wskClient.RemoteHostIP
        .Port = wskClient.RemotePort
    End With
    If UBound(strParts) <> 2 Then
        With GE
            .EventSubtype = gesHTTPError
            .Method = "ERROR"
            .Text = "Bad Request Line: " & mStrReqLine
        End With
        Parent.RaiseLogEvent GE, mIntIndex
        
        mReqState = rqsIdle
        wskClient.Close
        mStrBuffer = ""
        mBlnInUse = False
    Else
        'Process "Connection: close" headers.
        On Error Resume Next
        strKeepAlive = mColReqHeaders("CONNECTION")(1)
        If Err.Number = 0 Then
            On Error GoTo 0
            If UCase$(strKeepAlive) = "CLOSE" Then mBlnConnClose = True
        End If
        
        strParts(2) = UCase$(strParts(2))
        mStrHTTPVersion = strParts(2)
        With GE
            .Method = strParts(0)
            .Text = strParts(1)
            .HTTPVersion = strParts(2)
        End With
        strParts(1) = Replace$(strParts(1), "/", "\")
        Select Case strParts(0)
            Case "GET", "HEAD"
                lngInStr = InStr(strParts(1), "?")
                If lngInStr > 0 Then
                    'Request parameters present, assume dynamic content request.
                    If strParts(0) = "HEAD" Then
                        SendCanned 501
                    Else
                        'GET dynamic content.
                        GE.EventSubtype = gesGETDynamic
                        Parent.RaiseDynamicRequest strParts(0), _
                                                   Left$(strParts(1), lngInStr - 1), _
                                                   Mid$(strParts(1), lngInStr + 1), _
                                                   mColReqHeaders, _
                                                   sngStatus, _
                                                   strStatusText, _
                                                   strMIME, _
                                                   strRespExtraHeaders, _
                                                   mBytResponse, _
                                                   mIntIndex
                        If sngStatus = 0 Then
                            SendCanned 501
                        Else
                            SendResponse sngStatus, _
                                         strStatusText, _
                                         strMIME, _
                                         strRespExtraHeaders
                        End If
                    End If
                Else
                    'Request for static resource.
                    GE.EventSubtype = gesGETStatic
                    strFile = strParts(1)
                    If Right$(strFile, 1) = "\" Then strFile = strFile & Parent.DefaultPage
                    strFile = Parent.ResolvePath(Parent.VDirPath & strFile)
                    If Left$(strFile, Len(Parent.VDirPath)) <> Parent.VDirPath Then
                        'Bad request, trying to snoop outside VDir?
                        SendCanned 403
                    Else
                        'Locate file.
                        On Error Resume Next
                        GetAttr strFile
                        If Err.Number Then
                            'No such file.
                            On Error GoTo 0
                            SendCanned 404
                        Else
                            'Found file.
                            On Error GoTo 0
                            If strParts(0) = "HEAD" Then
                                'Return only HEADers.
                                SendStaticHeader FileLen(strFile), strFile
                            Else
                                'GET of static content.
                                On Error Resume Next
                                dtSince = _
                                    Parent.UTCParseString(mColReqHeaders("IF-MODIFIED-SINCE")(1))
                                If Err.Number Then
                                    On Error GoTo 0
                                    blnReturnFile = True
                                Else
                                    On Error GoTo 0
                                    blnReturnFile = _
                                        dtSince < Parent.UTCDateTime(FileDateTime(strFile))
                                End If
                                
                                If blnReturnFile Then
                                    On Error Resume Next
                                    mIntFile = Parent.GetFreeFile()
                                    If Err.Number Then
                                        On Error GoTo 0
                                        Parent.RaiseLogEvent GE, mIntIndex
                                        GE.EventSubtype = gesServerError
                                        GE.Text = "Ran out of file numbers"
                                        SendCanned 500.13
                                    Else
                                        'Open file, format and send headers and prime for
                                        'transmission of content.
                                        On Error GoTo 0
                                        Open strFile For Binary Access Read As #mIntFile
                                        ReDim mBytBuffer(STATIC_BUFFER_SZ - 1)
                                        lngFileBytes = LOF(mIntFile)
                                        mIntFullBlocks = lngFileBytes \ STATIC_BUFFER_SZ
                                        mIntLastBlockSize = lngFileBytes Mod STATIC_BUFFER_SZ
                                        SendStaticHeader lngFileBytes, strFile
                                        'File content will be sent via SendComplete handler.
                                    End If
                                Else
                                    SendCanned 304
                                End If
                            End If
                        End If
                    End If
                End If
            
            Case "POST"
                GE.EventSubtype = gesPOST
                Parent.RaiseDynamicRequest strParts(0), _
                                           strParts(1), _
                                           mStrContent, _
                                           mColReqHeaders, _
                                           sngStatus, _
                                           strStatusText, _
                                           strMIME, _
                                           strRespExtraHeaders, _
                                           mBytResponse, _
                                           mIntIndex
                If sngStatus = 0 Then
                    SendCanned 501
                Else
                    SendResponse sngStatus, _
                                 strStatusText, _
                                 strMIME, _
                                 strRespExtraHeaders
                End If
            
            Case Else
                GE.EventSubtype = gesUnknown
                wskClient_Close
        End Select

        mStrContent = ""
        Parent.RaiseLogEvent GE, mIntIndex
        mReqState = rqsIdle
        
        Set mColReqHeaders = Nothing
    End If
End Sub

Private Sub SendCanned(ByVal Status As Single)
    Const MSG304 As String = "304 Not Modified"
    Const MSG403 As String = "403 Forbidden"
    Const MSG404 As String = "404 Not Found"
    Const MSG500 As String = "500 Internal Server Error"
    Const MSG500_13 As String = "500.13 Server busy"
    Const MSG501 As String = "501 Not Implemented"
    
    Select Case Status
        Case 304
            SendCannedFormatted MSG304
        
        Case 403
            SendCannedFormatted MSG403
        
        Case 404
            SendCannedFormatted MSG404
        
        Case 500.13
            SendCannedFormatted MSG500_13
        
        Case 501
            SendCannedFormatted MSG501
        
        Case Else
            SendCannedFormatted MSG500
    End Select
End Sub

Private Sub SendCannedFormatted(ByVal StatusText As String)
    AppendResp mStrHTTPVersion & " " & StatusText & vbCrLf
    AppendResp "Date:" & Parent.UTCString(Now()) & vbCrLf
    AppendResp "Content-Type:text/html" & vbCrLf
    AppendResp "Content-Length:" & CStr(Len(StatusText) + 35) & vbCrLf
    AppendResp "Accept-Ranges:none" & vbCrLf
    AppendResp Parent.ServerHeader & DOUBLE_CRLF
    AppendResp "<html><body><h1>" & StatusText & "</h1></body></html>"
    wskClient.SendData ExtractResp()
End Sub

Private Sub SendResponse(ByVal Status As Single, _
                         ByVal StatusText As String, _
                         ByVal MIME As String, _
                         ByVal ExtraHeaders As String)
    
    AppendResp mStrHTTPVersion & " " & CStr(Status) & " " & StatusText & vbCrLf
    AppendResp "Date:" & Parent.UTCString(Now()) & vbCrLf
    If InStr(1, ExtraHeaders, "Last-Modified:", vbTextCompare) = 0 Then
        AppendResp "Last-Modified:" & Parent.UTCString(Now()) & vbCrLf
    End If
    If Len(MIME) > 0 Then AppendResp "Content-Type:" & MIME & vbCrLf
    mLngResponseLen = 0
    On Error Resume Next
    mLngResponseLen = UBound(mBytResponse) + 1
    On Error GoTo 0
    AppendResp "Content-Length:" & CStr(mLngResponseLen) & vbCrLf
    AppendResp "Accept-Ranges:none" & vbCrLf
    AppendResp Parent.ServerHeader & vbCrLf
    If Len(ExtraHeaders) > 0 Then
        AppendResp ExtraHeaders
        If Right$(ExtraHeaders, 2) <> vbCrLf Then AppendResp vbCrLf
    End If
    AppendResp vbCrLf 'Second CRLF, terminating headers.
    wskClient.SendData ExtractResp() 'Body if any will be sent via SendComplete handler.
End Sub

Private Sub SendStaticHeader(ByVal Length As Long, _
                             ByVal Resource As String)
    Dim strMIME As String
    
    strMIME = Parent.ExtensionToMIME(Mid$(Resource, InStrRev(Resource, ".") + 1))
    AppendResp mStrHTTPVersion & " 200 Ok" & vbCrLf
    AppendResp "Date:" & Parent.UTCString(Now()) & vbCrLf
    AppendResp "Last-Modified:" & Parent.UTCString(FileDateTime(Resource)) & vbCrLf
    AppendResp "Content-Type:" & strMIME & vbCrLf
    AppendResp "Content-Length:" & CStr(Length) & vbCrLf
    AppendResp "Accept-Ranges:none" & vbCrLf
    AppendResp Parent.ServerHeader & DOUBLE_CRLF
    wskClient.SendData ExtractResp()
End Sub

Private Sub UserControl_Resize()
    Const DIMENSIONS As Single = 420
    
    With UserControl
        .Height = DIMENSIONS
        .Width = DIMENSIONS
    End With
End Sub

Private Sub wskClient_Close()
    If wskClient.State <> sckClosed Then
        wskClient.Close
        wskClient.RemotePort = 0
    End If
    
    If mIntFile Then
        Close #mIntFile
        mIntFile = 0
    End If
    
    mReqState = rqsIdle
    mStrBuffer = ""
    mStrRespBuffer = ""
    mLngRespUsed = 0
    mBlnInUse = False
    mBlnConnClose = False
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    Dim GE As GossEvent
    Dim lngHeader As Long
    Dim lngInStr As Long
    Dim strChar As String
    Dim strContentLen As String
    Dim strFragment As String
    Dim strHeadBlock As String
    Dim strHeaders() As String
    Dim strParts() As String
    
    wskClient.GetData strFragment, vbString
    mStrBuffer = mStrBuffer & strFragment
    
    'Look for the Request Line if we don't have it.
    If mReqState = rqsIdle Then
        'Erratic POST cleanup:
        'There is not supposed to be anything after the POST content but many
        'clients submit an extra CRLF.  Delete them if found here, which will
        'be a leftover from a previous request on a persistent connection.
        Do
            lngInStr = InStr(mStrBuffer, vbCrLf)
            If lngInStr > 0 Then
                If lngInStr > 1 Then
                    'We found a complete Request Line.
                    mStrReqLine = Left$(mStrBuffer, lngInStr - 1)
                    mReqState = rqsReqLineComplete
                End If
                mStrBuffer = Mid$(mStrBuffer, lngInStr + 2)
            End If
        Loop Until lngInStr = 0 Or lngInStr > 1
    End If
    
    'Look for the Headers block if we have the Request Line.
    If mReqState = rqsReqLineComplete Then
        lngInStr = InStr(mStrBuffer, DOUBLE_CRLF)
        If lngInStr > 0 Then
            'We have the Headers.
            strHeadBlock = Left$(mStrBuffer, lngInStr - 1)
            mStrBuffer = Mid$(mStrBuffer, lngInStr + 4)
            
            'Parse Headers into Collection. Keys are stored UPPERCASED.
            Set mColReqHeaders = New Collection
            strHeaders = Split(strHeadBlock, vbCrLf)
            For lngHeader = 0 To UBound(strHeaders)
                strParts = Split(strHeaders(lngHeader), ":", 2)
                'Strip whitespace from Attribute.
                strChar = Right$(strParts(0), 1)
                Do While strChar = vbTab Or strChar = " "
                    strParts(1) = Left$(strParts(0), Len(strParts(0)) - 1)
                    strChar = Right$(strParts(0), 1)
                Loop
                If UBound(strParts) > 0 Then
                    'Strip whitespace from Value.
                    strChar = Left$(strParts(1), 1)
                    Do While strChar = vbTab Or strChar = " "
                        strParts(1) = Mid$(strParts(1), 2)
                        strChar = Left$(strParts(1), 1)
                    Loop
                End If
                'Watch for and remove duplicate headers (keep last one).
                On Error Resume Next
                mColReqHeaders.Add strParts, UCase$(strParts(0))
                If Err.Number Then
                    mColReqHeaders.Remove strParts(0)
                    mColReqHeaders.Add strParts, UCase$(strParts(0))
                End If
                On Error GoTo 0
            Next
            
            'Look for Content-Length.
            On Error Resume Next
            strContentLen = mColReqHeaders("CONTENT-LENGTH")(1)
            If Err.Number Then
                'No Content-Length header.  Bypass checking for it.
                On Error GoTo 0
                mReqState = rqsRequestComplete
            Else
                'Process Content-Length.
                On Error GoTo 0
                If IsNumeric(strContentLen) Then
                    mLngContentLen = CLng(strContentLen)
                    mReqState = rqsHeadersComplete
                Else
                    'Bad Content-Length error.
                    Set GE = New GossEvent
                    With GE
                        .EventType = getHTTP
                        .EventSubtype = gesHTTPError
                        .IP = wskClient.RemoteHostIP
                        .Port = wskClient.RemotePort
                        .Method = "ERROR"
                        .Text = "Bad Content-Length header value: " & strContentLen
                    End With
                    Parent.RaiseLogEvent GE, mIntIndex
                    
                    wskClient_Close
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Look for the end of the Request if we have processed the Headers.
    If mReqState = rqsHeadersComplete Then
        If Len(mStrBuffer) >= mLngContentLen Then
            mStrContent = Left$(mStrBuffer, mLngContentLen)
            mStrBuffer = Mid$(mStrBuffer, mLngContentLen + 1)
            mReqState = rqsRequestComplete
        End If
    End If
    
    'Process completed Request (all of Content-Length rcvd or no Content-Length header).
    If mReqState = rqsRequestComplete Then ProcessRequest
End Sub

Private Sub wskClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim EventType As GossEventTypes
    Dim GE As GossEvent
    
    CancelDisplay = True
    
    If Number = sckWouldBlock Then
        EventType = getWSSoftError
    Else
        EventType = getWSError
        wskClient_Close
    End If
    
    If Number <> sckConnectAborted Then
        Set GE = New GossEvent
        With GE
            .EventType = EventType
            .EventSubtype = Number
            .IP = wskClient.RemoteHostIP
            .Port = wskClient.RemotePort
            .Text = Description
        End With
        Parent.RaiseLogEvent GE, mIntIndex
    'Else
        'Persistent connection closed/aborted, which is not logged.
    End If
End Sub

Private Sub wskClient_SendComplete()
    If mLngResponseLen > 0 Then
        mLngResponseLen = 0
        wskClient.SendData mBytResponse
        Erase mBytResponse
        Exit Sub 'Bypass CheckClose until next SendComplete.
    End If
    
    If mIntFile Then
        'We're sending a static (file) resource.  Continue.
        If mIntFullBlocks > 0 Then
            Get #mIntFile, , mBytBuffer
            mIntFullBlocks = mIntFullBlocks - 1
        Else
            If mIntLastBlockSize > 0 Then
                ReDim mBytBuffer(mIntLastBlockSize - 1)
                Get #mIntFile, , mBytBuffer
            End If
            Close #mIntFile
            mIntFile = 0
            If mIntLastBlockSize <= 0 Then GoTo CheckClose
        End If
        wskClient.SendData mBytBuffer
        Exit Sub 'Bypass CheckClose until next SendComplete.
    End If

CheckClose:
    If mBlnConnClose Then
        'Request had a "Connection: close" header.
        wskClient_Close
    End If
End Sub
