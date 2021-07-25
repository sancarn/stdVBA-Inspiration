VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Gossamer 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   405
   ScaleWidth      =   405
   ToolboxBitmap   =   "Gossamer.ctx":0000
   Begin MSWinsockLib.Winsock wskRequest 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin GossDemo1.GossClient gcClients 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   400
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   370
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillStyle       =   7  'Diagonal Cross
      Height          =   400
      Left            =   0
      Top             =   0
      Width           =   400
   End
End
Attribute VB_Name = "Gossamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Gossamer
'========
'
'A tiny HTTP server control.
'
'This control uses a Winsock control as a listener for incoming HTTP
'client connections which are handed off to a control array of
'GossClient controls.
'
'Changes
'-------
'
'Version: 1.1
'
' o Added Port property.
' o Changed StartListening so Port parameter is optional, using
'   current value of the Port property setting if not supplied.
' o Added procedure attributes, set LogEvent as default event.
' o Marked several Public members "hidden" since they're only meant
'   for use by GossClients.
' o Methods: LogEvent is now RaiseLogEvent, DynamicRequest is now
'   RaiseDynamicRequest.
' o New handling of VDir, property can now only be set while not
'   listening.
'
'Version: 1.2
'
' o Split out function UTCDateTime from UTCString.
' o Added function UTCParseString to convert HTTP timestamps to
'   Date values.
' o VDirPath R/O property is no longer hidden, since it can be useful
'   in handling dynamic requests.
'
'Version: 1.3
'
' no change
'
'Version: 1.4
'
' o Added new (hidden) property ServerHeader for use by GossClient.
' o Added EntityEncode method.
'
'Version: 1.5
'
' o Corrected ResolvePath() so it no longer allocates a buffer 1 char too
'   long nor needs to then truncate prior to return.
'
'Version 1.6
'
' o Treat sckWouldBlock as a getWSSoftError: log it as such and then
'   ignore it.
' o Stop setting .Timestamp since it gets set when GE instances are
'   created.
'
'Version: 1.7
'
' o Added GOSS_VERSION_MAJOR and GOSS_VERSION_MINOR Consts for use in the
'   default ServerHeader value.
' o Removed sending an extraneous space after header ":" separators.  Not
'   important, and quite commonly done but a silly waste of bandwidth.
'

Private Const GOSS_VERSION_MAJOR As String = "1"
Private Const GOSS_VERSION_MINOR As String = "7"

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetFullPathName Lib "kernel32" _
    Alias "GetFullPathNameW" ( _
    ByVal lpFileName As Long, _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As Long, _
    ByVal lpFilePart As Long) As Long

Private Declare Function GetTimeZoneInformation Lib "kernel32" ( _
    lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Private mMaxConnections As Integer
Private mVDir As String
Private mVDirPath As String
Private mServerHeader As String

Public DefaultPage As String 'Default page for directory requests.
Attribute DefaultPage.VB_VarDescription = "Simple file name of Gossamer's default page to return on directory requests"
Public Port As Long          'Server listen port.
Attribute Port.VB_VarDescription = "HTTP port to listen on"

Public Event DynamicRequest(ByVal Method As String, _
                            ByVal URI As String, _
                            ByVal Params As String, _
                            ByVal ReqHeaders As Collection, _
                            ByRef RespStatus As Single, _
                            ByRef RespStatusText As String, _
                            ByRef RespMIME As String, _
                            ByRef RespExtraHeaders As String, _
                            ByRef RespBody() As Byte, _
                            ByVal ClientIndex As Integer)
                            
Public Event LogEvent(ByVal GossEvent As GossEvent, _
                      ByVal ClientIndex As Integer) '-1 for Gossamer events.
Attribute LogEvent.VB_Description = "Raised when a loggable event has occurred"
Attribute LogEvent.VB_MemberFlags = "200"

Public Property Get MaxConnections() As Long
Attribute MaxConnections.VB_Description = "Maximum number of active client connections to accept"
    MaxConnections = mMaxConnections
End Property

Public Property Let MaxConnections(ByVal Max As Long)
    If 0 < Max Or Max <= 1000 Then
        mMaxConnections = Max
    Else
        Err.Raise &H8004A702, "Gossamer", "MaxConnections must be 1 to 1000"
    End If
End Property

Public Property Get ServerHeader() As String
Attribute ServerHeader.VB_Description = "Returns versioned server header string for responses"
Attribute ServerHeader.VB_MemberFlags = "40"
    ServerHeader = mServerHeader
End Property

Public Property Get State() As Integer
Attribute State.VB_Description = "Returns State of the listener Winsock control"
    State = wskRequest.State
End Property

Public Property Get VDir() As String
Attribute VDir.VB_Description = "Directory containing static resources of Gossamer site"
    VDir = mVDir
End Property

Public Property Let VDir(ByVal Directory As String)
    If wskRequest.State <> sckListening Then
        mVDir = Directory
        If Len(mVDir) = 0 Then
            mVDirPath = CurDir$()
        Else
            mVDirPath = ResolvePath(mVDir)
        End If
    Else
        Err.Raise &H8004A704, "Gossamer", "Can't change VDir while listening"
    End If
End Property

Public Property Get VDirPath() As String
Attribute VDirPath.VB_Description = "Returns fully qualified path for current VDir setting"
    VDirPath = mVDirPath
End Property

Public Function EntityEncode(ByVal Text As String) As String
Attribute EntityEncode.VB_Description = "Encode a text string to be inserted into HTML as text with HTML entity encoding"
    EntityEncode = Join$(Split(Text, "&"), "&amp;")
    EntityEncode = Join$(Split(EntityEncode, """"), "&quot;")
    EntityEncode = Join$(Split(EntityEncode, "<"), "&lt;")
    EntityEncode = Join$(Split(EntityEncode, ">"), "&gt;")
End Function

Public Function ExtensionToMIME(ByVal Extension As String) As String
Attribute ExtensionToMIME.VB_Description = "Return MIME type corresponding to the supplied file extension value (without .)"
    Extension = UCase$(Extension)
    Select Case Extension
        Case "CSS"
            ExtensionToMIME = "text/css"
        Case "GIF"
            ExtensionToMIME = "image/gif"
        Case "HTM", "HTML"
            ExtensionToMIME = "text/html"
        Case "ICO"
            ExtensionToMIME = "image/vnd.microsoft.icon"
        Case "JPG", "JPEG"
            ExtensionToMIME = "image/jpeg"
        Case "JS", "JSE"
            ExtensionToMIME = "application/javascript"
        Case "PNG"
            ExtensionToMIME = "image/png"
        Case "RTF"
            ExtensionToMIME = "application/rtf"
        Case "TIF", "TIFF"
            ExtensionToMIME = "image/tiff"
        Case "TXT"
            ExtensionToMIME = "text/plain"
        Case "VBS", "VBE"
            ExtensionToMIME = "application/vbscript"
        Case "XML", "XSD"
            ExtensionToMIME = "text/xml"
        Case "ZIP"
            ExtensionToMIME = "application/zip"
        Case Else
            ExtensionToMIME = "application/octet-stream"
    End Select
End Function

Public Function GetFreeFile() As Integer
Attribute GetFreeFile.VB_Description = "Calls FreeFile(0) to get file number, if exhausted tries FreeFile(1)"
    On Error Resume Next
    GetFreeFile = FreeFile(0)
    If Err.Number Then
        On Error GoTo 0
        GetFreeFile = FreeFile(1)
    End If
End Function

Public Sub RaiseDynamicRequest(ByVal Method As String, _
                               ByVal URI As String, _
                               ByVal Params As String, _
                               ByVal ReqHeaders As Collection, _
                               ByRef RespStatus As Single, _
                               ByRef RespStatusText As String, _
                               ByRef RespMIME As String, _
                               ByRef RespExtraHeaders As String, _
                               ByRef RespBody() As Byte, _
                               ByVal Index As Integer)
    RaiseEvent DynamicRequest(Method, _
                              URI, _
                              Params, _
                              ReqHeaders, _
                              RespStatus, _
                              RespStatusText, _
                              RespMIME, _
                              RespExtraHeaders, _
                              RespBody, _
                              Index)
End Sub

Public Sub RaiseLogEvent(ByVal GossEvent As GossEvent, ByVal Index As Integer)
Attribute RaiseLogEvent.VB_Description = "Only for use by GossClient"
Attribute RaiseLogEvent.VB_MemberFlags = "40"
    RaiseEvent LogEvent(GossEvent, Index)
End Sub

Public Function ResolvePath(ByVal RelativePath As String) As String
Attribute ResolvePath.VB_Description = "Only for use by GossClient"
Attribute ResolvePath.VB_MemberFlags = "40"
    'Returns full path to RelativePath, "" if any error.
    Dim strFullPath As String
    Dim lngLen As Long
    Dim lngFilePart As Long
    
    lngLen = GetFullPathName(StrPtr(RelativePath), 0, StrPtr(strFullPath), lngFilePart)
    If lngLen Then
        'If the lpBuffer buffer is too small to contain the path, the return value
        'is the size, in TCHARs, of the buffer that is required to hold the path
        'and the terminating null character.
        strFullPath = String$(lngLen - 1, 0)
        lngLen = GetFullPathName(StrPtr(RelativePath), lngLen, StrPtr(strFullPath), lngFilePart)
        If lngLen Then
            'If the function succeeds, the return value is the length, in TCHARs,
            'of the string copied to lpBuffer, not including the terminating null
            'character.
            ResolvePath = strFullPath
        End If
    End If
End Function

Public Sub StartListening(Optional ByVal Port As Long = -1, Optional ByVal AdapterIP As String = "")
Attribute StartListening.VB_Description = "Begin accepting HTTP connections, may specify listen port and adapter IP to bind to"
    Dim GE As GossEvent
    
    If Port > -1 Then Me.Port = Port
    
    Set GE = New GossEvent
    With GE
        .EventType = getServer
        .EventSubtype = gesStarted
        .IP = AdapterIP
        .Port = Me.Port
        .Text = "Gossamer service started"
    End With
    RaiseEvent LogEvent(GE, -1)
    
    With wskRequest
        If wskRequest.State = sckListening Then
            Err.Raise &H8004A700, "Gossamer", "Already listening"
        End If
        
        If Len(AdapterIP) > 0 Then
            .Bind Me.Port, AdapterIP
        Else
            .Bind Me.Port
        End If
        .Listen
    End With
End Sub

Public Sub StopListening()
Attribute StopListening.VB_Description = "Shuts down any active GossClients and stops listening for client connection requests"
    Dim GE As GossEvent
    Dim Index As Integer
    
    wskRequest.Close
    For Index = 0 To gcClients.Count - 1
        If gcClients(Index).InUse Then gcClients(Index).Shutdown
    Next
    
    Set GE = New GossEvent
    With GE
        .EventType = getServer
        .EventSubtype = gesStopped
        .Text = "Gossamer service stopped"
    End With
    RaiseEvent LogEvent(GE, -1)
End Sub

Public Function URLDecode(ByVal URLEncoded As String) As String
Attribute URLDecode.VB_Description = "Converts URLEncoded string to plaintext string"
    Dim intPart As Integer
    Dim strParts() As String
    
    URLDecode = Replace$(URLEncoded, "+", " ")
    strParts = Split(URLDecode, "%")
    For intPart = 1 To UBound(strParts)
        strParts(intPart) = _
                Chr$(CLng("&H" & Left$(strParts(intPart), 2))) _
              & Mid$(strParts(intPart), 3)
    Next
    URLDecode = Join$(strParts, "")
End Function

Public Function UTCDateTime(ByVal DateTime As Date) As Date
Attribute UTCDateTime.VB_Description = "Convert Date value from local time to UTC equivalent"
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lngRet As Long
    Dim lngOffsetMinutes As Long
    
    'Return the time difference between local & GMT time in minutes.
    lngRet = GetTimeZoneInformation(tzi)
    lngOffsetMinutes = -tzi.Bias
    
    'If we are in daylight saving time, apply the bias if applicable.
    If lngRet = TIME_ZONE_ID_DAYLIGHT Then
        If tzi.DaylightDate.wMonth Then
            lngOffsetMinutes = lngOffsetMinutes - tzi.DaylightBias
        End If
    End If
    
    UTCDateTime = DateAdd("n", lngOffsetMinutes, DateTime)
End Function

Public Function UTCParseString(ByVal UTCString As String) As Date
Attribute UTCParseString.VB_Description = "Convert HTTP UTC timestamp string to Date value, if badly formatted return Now() value"
    On Error Resume Next
    UTCParseString = CDate(Mid$(UTCString, 6, 20))
    If Err.Number Then UTCParseString = UTCDateTime(Now())
End Function

Public Function UTCString(ByVal DateTime As Date) As String
Attribute UTCString.VB_Description = "Converts Date value in local time zone to HTTP timestamp in GMT form"
    UTCString = Format$(UTCDateTime(DateTime), _
                        "Ddd, dd Mmm YYYY HH:NN:SS \G\M\T")
End Function

Private Sub UserControl_Initialize()
    mServerHeader = "Server:BVO.Gossamer/" & GOSS_VERSION_MAJOR & "." & GOSS_VERSION_MINOR _
                  & vbCrLf _
                  & "X-Powered-By:Microsoft.Visual.Basic.6.0"
    gcClients(0).Init 0
End Sub

Private Sub UserControl_InitProperties()
    DefaultPage = "index.htm"
    MaxConnections = 32
    Port = 8080
    VDir = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DefaultPage = PropBag.ReadProperty("DefaultPage", "index.htm")
    MaxConnections = PropBag.ReadProperty("MaxConnections", 32)
    Port = PropBag.ReadProperty("Port", 8080)
    VDir = PropBag.ReadProperty("VDir", "")
End Sub

Private Sub UserControl_Resize()
    Const DIMENSIONS As Single = 420
    
    With UserControl
        .Height = DIMENSIONS
        .Width = DIMENSIONS
    End With
    With Shape1
        .Height = DIMENSIONS
        .Width = DIMENSIONS
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DefaultPage", DefaultPage, "index.htm"
    PropBag.WriteProperty "MaxConnections", MaxConnections, 32
    PropBag.WriteProperty "Port", Port, 8080
    PropBag.WriteProperty "VDir", VDir, ""
End Sub

Private Sub wskRequest_ConnectionRequest(ByVal requestID As Long)
    Dim Index As Integer
    
    For Index = 0 To gcClients.Count - 1
        If Not gcClients(Index).InUse Then
            gcClients(Index).Accept requestID
            Exit Sub
        End If
    Next
    
    'Refuse connections over the limit.
    If Index + 1 > mMaxConnections Then Exit Sub
    
    'Load another instance of GossClient.
    Load gcClients(Index)
    gcClients(Index).Init Index
    gcClients(Index).Accept requestID
End Sub

Private Sub wskRequest_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim EventType As GossEventTypes
    Dim GE As GossEvent
    
    CancelDisplay = True
    
    If Number = sckWouldBlock Then
        EventType = getWSSoftError
    Else
        EventType = getWSError
        wskRequest.Close 'Stop accepting requests.
    End If
    
    Set GE = New GossEvent
    With GE
        .EventType = EventType
        .EventSubtype = Number
        .IP = wskRequest.RemoteHostIP
        .Port = wskRequest.RemotePort
        .Text = Description
    End With
    RaiseEvent LogEvent(GE, -1)
End Sub
