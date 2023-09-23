VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Create & Track Async Downloads ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' This class is a collection of cAsyncClient classes.
' The class initializes clients, activates them, and manages a queue. The queue is used
'   to cache downloads awaiting a free slot. At most 10 simultaneous downloads are allowed.

Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long

Private m_Hwnd As Long                      ' handle window used for delay starting async downloads
Private m_InetType As Long                  ' XMLHTTP type
Private m_Clients As Collection             ' collection of client classes
Private m_Queue As Collection               ' collection of queued client classes
Private Const MAXCLIENTS As Long = 10&      ' max number of active downloads at any one time

Public Function AddRef(Client As cAsyncClient) As Boolean

    If AsyncModeAvailable = False Then Exit Function
    
    If m_Clients Is Nothing Then
        Set m_Clients = New Collection
        Set m_Queue = New Collection
    End If
    
    If m_Clients.Count = MAXCLIENTS Then
        m_Queue.Add Client, CStr(ObjPtr(Client))
    Else
        m_Clients.Add Client, CStr(ObjPtr(Client))
        If m_Queue.Count Then pvPopQueuedClient
        AddRef = True
    End If

End Function

Public Function RemoveRef(Client As cAsyncClient, bQueued As Boolean) As Boolean

    If Not m_Clients Is Nothing Then
        
        On Error Resume Next
        
        If bQueued Then
            If m_Queue.Count Then
                m_Queue.Remove CStr(ObjPtr(Client))
                If Err Then Err.Clear
            End If
        ElseIf m_Clients.Count Then
            m_Clients.Remove CStr(ObjPtr(Client))
            If Err Then Err.Clear
            If m_Queue.Count Then pvPopQueuedClient
        End If
        If (m_Clients.Count Or m_Queue.Count) = 0& Then
            If m_Hwnd Then
                DestroyWindow m_Hwnd
                m_Hwnd = 0&
            End If
        End If
        RemoveRef = True
        
'        Debug.Print "after removeref: "; m_Clients.Count; m_Queue.Count
        
    End If

End Function

Public Property Get DelayWindow() As Long
    DelayWindow = m_Hwnd
End Property

Public Property Get AsyncModeAvailable(Optional AlwaysDelayWindow As Boolean = False) As Boolean

    If m_InetType = 0& Then
        
        Dim tTestObj As Object
        On Error Resume Next
        ' version 3.0 of XMLHTTP
        m_InetType = 1&
        Set tTestObj = CreateObject("MSXML2.ServerXMLHTTP")
        If Err Then
            Err.Clear
            m_InetType = m_InetType + 1&
            Set tTestObj = CreateObject("Microsoft.XMLHTTP")
            If Err Then
                m_InetType = -1&
                Err.Clear
            End If
        End If
        Set tTestObj = Nothing
        On Error GoTo 0
    End If
    If m_InetType > 0& Or AlwaysDelayWindow = True Then AsyncModeAvailable = pvCreateDelayWindow()
    
    
End Property

Public Sub Abort(ActiveClients As Boolean, QueuedClients As Boolean)
    
    ' global abort routine for async downloads
    ' can abort just acitve downloads or just queued or both
    
    If Not (ActiveClients Or QueuedClients) Then Exit Sub
    
    Dim v As Variant
    Dim colClients As Collection
    
    ' copy active clients & set clients to new collection so we have a zero count
    ' with a zero count, any attempted calls to RemoveRef from clients will not modify collection
    Set colClients = m_Clients
    Set m_Clients = New Collection
    
    On Error Resume Next
    If ActiveClients Then                   ' abort all active downloads
        For Each v In colClients
            v.Abort
        Next
        Set colClients = New Collection
    End If
    If QueuedClients Then                   ' abort all queued downloads
        For Each v In m_Queue
            v.Abort
        Next
        Set m_Queue = New Collection
    End If
    Set m_Clients = colClients
    Set colClients = Nothing
    If ActiveClients = False Then
        ' it may be possible while we were purging the queue, a client finished
        ' and attempted to remove itself from the active list. If so, it didn't
        ' so we do it now
        For Each v In m_Clients
            Select Case v.State
            Case lvicAsync_Downloading, lvicAsync_Queued, lvicAsync_WaitingOnReply
                ' do nothing
            Case Else
                m_Clients.Remove CStr(ObjPtr(v))
            End Select
        Next
    End If
    If m_Clients.Count = 0& Then            ' if no more active clients, then
        If m_Queue.Count Then               ' activate queued if we have any
            pvPopQueuedClient
        ElseIf m_Hwnd Then
            DestroyWindow m_Hwnd
            m_Hwnd = 0&
        End If
    End If
        
End Sub

Public Sub GetStatus(NrDownloading As Long, NrWaiting As Long, NrQueued As Long)

    Dim v As Variant
    If Not m_Clients Is Nothing Then
        On Error Resume Next
            For Each v In m_Clients
                Select Case v.State
                Case lvicAsync_Downloading: NrDownloading = NrDownloading + 1&
                Case lvicAsync_WaitingOnReply: NrWaiting = NrWaiting + 1&
                Case Else
                End Select
            Next
        On Error GoTo 0
        NrQueued = m_Queue.Count
    End If

End Sub

Public Function CreateClient(URL As String) As cAsyncClient

    If Me.AsyncModeAvailable = False Then Exit Function
    
    Set CreateClient = New cAsyncClient
    CreateClient.URL = URL

End Function

Private Sub pvPopQueuedClient()

    If m_Clients.Count < MAXCLIENTS Then
        Dim v As Variant
        Set v = m_Queue.Item(1)
        m_Queue.Remove 1
        v.Activate
    End If

End Sub

Private Function pvCreateDelayWindow() As Boolean

    If m_Hwnd = 0& Then
        m_Hwnd = CreateWindowEx(0&, "Static", "lvicDelayRelay", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    End If
    pvCreateDelayWindow = (m_Hwnd <> 0&)

End Function

Private Sub Class_Initialize()
    Randomize Timer
End Sub

Private Sub Class_Terminate()
    
    Me.Abort True, True

End Sub



