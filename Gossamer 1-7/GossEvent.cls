VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GossEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'GossEvent
'=========
'
'This class is used for passing loggable HTTP and status events
'to the application.
'
'Changes
'-------
'
'Version: 1.1
'Version: 1.2
'Version: 1.3
'Version: 1.4
'
' no change
'
'Version: 1.5
'
' o New gesServerError value added.
'
'Version 1.6
'
' o New getWSSoftError value added.
' o New ClientIndex property added, -1 for wskRequest.
' o Set Timestamp upon creation.
'
'Version: 1.7
'
' no change
'

Public Enum GossEventSubtypes
    'EventType = getHTTP.
    gesUnknown = -1
    gesGETStatic = 0
    gesGETDynamic
    gesHTTPError
    gesPOST
    gesServerError
    
    'EventType = getServer.
    gesStarted = 0
    gesStopped
End Enum

Public Enum GossEventTypes
    getWSError = 1
    getHTTP
    getServer
    getWSSoftError
End Enum

Public EventSubtype As GossEventSubtypes 'Or Winsock error Number.
Attribute EventSubtype.VB_VarDescription = "Event subtype or Winsock error number"
Public EventType As GossEventTypes
Attribute EventType.VB_VarDescription = "Event type"
Public HTTPVersion As String 'Normally "HTTP/m.n" or empty.
Attribute HTTPVersion.VB_VarDescription = "HTTP version from request line of the request or empty"
Public IP As String
Attribute IP.VB_VarDescription = "Source IP of client"
Public Method As String 'Empty for Winsock error events.
Attribute Method.VB_VarDescription = "HTTP method (GET, HEAD, POST) of request"
Public Port As Long
Attribute Port.VB_VarDescription = "Local port of client connection"
Public Text As String
Attribute Text.VB_VarDescription = "Description of event"
Public Timestamp As Date
Attribute Timestamp.VB_VarDescription = "Date and time client request was received"

Private Sub Class_Initialize()
    Timestamp = Now()
End Sub
