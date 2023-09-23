VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Asyncrhonous download of an image via URL ========-----
' Class Instancing = Privzte (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' Note about sending events and VB blocking rules for uncompiled code
' If an event is sent from code while a modal window is displayed, the event will be blocked.
' Since this class is asyncrhonous, it is possible such a situation can occur. The major downside
' is that if this control is not compiled, the event to pass image data to the owner GDIpImage
' class won't be received. This means the data stays in this class' m_Data() array and no way
' of letting the owner class know it is ready to be processed. We could set up a timer and continue
' to resend the data, but since this an issue only when this control is uncompiled, we will not
' go that far. Instead, we will simply pass a boolean to the owner class and have it set its
' value to true. Whenever the user calls this class' State property and owner class was already
' notified once, it will resend the notification. This becomes the user's responsibility knowing
' how VB treats raised events in uncompiled controls/projects. This issue is not applicable once
' the control and project are both compiled.

' Suggestion for users. If you will be displaying any modal window (i.e., MsgBox, CommmonDialog, etc),
' and you have async downloads in progress, you should call this class' State event to ensure any
' downloaded image data is processed. Examples
'   From control: Debug.Assert (AlphaImgCtl1.Picture.AsyncDownloadStatus Or 1)
'   From a GDIpImage class: Debug.Assert (myImage.AsyncDownloadStatus Or 1)
' By using Debug.Assert calls, they will not be included in any compiled project & since this
' 'bug' is only present in uncompiled projects, win-win

' Note that this is typical behavior. Think of a timer control
' - While your project is uncompiled, the timer events stop coming while a modal window is open
' - When project is compiled, the timer events are no longer blocked.
' If in doubt, you can easily prove this...
' 1. Add a timer and command button to a form
' 2. In form load...
'   Timer1.Interval = 250: Timer1.Enabled = True
' 3. In the Timer1_Timer() event
'   Me.Caption = Timer
' 4. In the Command1_Click() event
'   MsgBox "Did caption stop updating?", vbOkOnly + vbInformation
' Now run the project uncompiled & compiled and you'll see the statements made about events
' being blocked are true, after clicking the command button


Public Event DownloadComplete(DataPointer As Long, ErrorCode As Long, Rcvd As Boolean)
' Error codes defined at http://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html

Private Enum StateMasksEnum
    eNotInit = &H80000000   ' initialized, not started
    eQueued = &H8F000000    ' queued, not active yet
    eComplete = &HFF000000  ' finished
    eExiting = &HF000000    ' download in progress
    eIDEfix = &HF00000      ' sent result.
    eSyncMode = &HF0000     ' sync mode from async request
End Enum

Private m_Data() As Byte
Private m_State As Long
' ^^ high byte will be one of the StateMasksEnum above
' ^^ low word = completed state/error code
' ^^ value of -1& indicates class terminating; don't forward events

Private m_URL As String
Private m_Inet As Object

Public Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
Attribute OnReadyStateChange.VB_MemberFlags = "40"

    ' must be default property & should be hidden
    
    If (m_State And eExiting) = eExiting Then
        Exit Sub ' recursive; already have result
    End If
    
    m_State = m_Inet.readyState

    Select Case m_State
    Case 0, 1, 2 ' object initialization messages
        
    Case 3 ' interactive
    
    Case 4 ' completed
        m_State = (eExiting Or m_Inet.Status)
        If (m_State And &HFFFF&) = lvicAsync_Ok Then
            Dim vResults As Variant
            vResults = m_Inet.responseBody
            modCommon.MoveArrayToVariant vResults, m_Data(), False
            Call pvCompleted
        Else
            Call Me.Abort
        End If
    
    End Select

End Sub

Public Function Activate() As Boolean

    If (m_State And eNotInit) = eNotInit Then ' initialized, not started or queued
    
    ' How activation occurs
    ' 1. New async request started, this class is created and URL cached; nothing more done
    ' 2. The owner GDIpImage class activates a timer that calls this function
    ' 3. The class calls the cAsyncController to add it to its client collection
    ' 4. If the controller refuses, this class is cached instead, by the controller
    ' 5. If not refused, this routine starts the download
    ' 6. If cached (controller refused it), when the controller's active client list has an
    '   open slot, it will call this function and Steps 3 thru 6 repeat as needed
    
        Dim bOK As Boolean, sURL As String, lDelim As Long
    
        On Error Resume Next
        bOK = g_AsyncController.AddRef(Me)
        
        If bOK Then
            Set m_Inet = CreateObject("MSXML2.ServerXMLHTTP")
            m_State = 1&
            If Err Then
                Err.Clear
                m_State = 2&
                Set m_Inet = CreateObject("Microsoft.XMLHTTP")
                If Err Then
                    Err.Clear
                    m_State = 0&
                End If
            End If
            
            If m_State = 0& Then
                m_State = eExiting Or 503& ' service unavailable
                Call Me.Abort
                Exit Function
            Else
                ' per msdn documentation, to prevent inet object from retrieving from
                ' the users cache, use POST vs GET. However, not all sites will accept
                ' a POST command when trying to retrieve data. Therefore, alternate
                ' workaround used which is to provide a bogus query within the URL.
                ' This randomized query string should prevent cache retrieval
                If m_State = 2& Then lDelim = InStrRev(m_URL, "/")
                If lDelim = 0& Then
                    sURL = m_URL
                Else
                    If InStr(lDelim, m_URL, "?") Then
                        sURL = m_URL & "&rnd" & ObjPtr(Me) & "=" & Int(Rnd * vbWhite)
                    Else
                        sURL = m_URL & "?rnd" & ObjPtr(Me) & "=" & Int(Rnd * vbWhite)
                    End If
                End If
                m_State = 16&
                m_Inet.open "GET", sURL, True
                m_Inet.OnReadyStateChange = Me
                m_Inet.send
                If Err Then
                    If m_URL <> sURL Then                   ' try again with original URL
                        Err.Clear
                        m_Inet.open "GET", m_URL, True
                        m_Inet.OnReadyStateChange = Me
                        m_Inet.send                         ' errored again? if so, abort
                    End If
                    If Err Then
                        Err.Clear
                        m_State = eExiting Or 400& ' bad request
                        Call Me.Abort
                        Exit Function
                    End If
                End If
            End If
        Else
            m_State = eQueued Or lvicAsync_Queued ' flag for queued
            
        End If
        
        Activate = True
        
    ElseIf (m_State And eSyncMode) Then
    
        Dim tImg As New GDIpImage
        If modCommon.LoadImage(m_URL, True, True, , tImg) Is Nothing Then
            m_State = eComplete Or lvicAsync_NoImage Or eIDEfix
            RaiseEvent DownloadComplete(0&, (m_State And &HFFFF&), True)
        Else
            m_State = eComplete Or lvicAsync_Ok Or eIDEfix
            RaiseEvent DownloadComplete(-1&, 0&, True)
        End If
        Set g_NewImageData = Nothing
    End If

End Function

Public Sub Abort()

    On Error Resume Next
    Dim bNoNotify As Boolean
    
    Select Case (m_State And eComplete)
    
    Case eQueued                 ' queued
        m_State = eComplete Or lvicAsync_UserAborted
        g_AsyncController.RemoveRef Me, True
        
    Case eComplete              ' already done or class terminating
        bNoNotify = True
        
    Case eNotInit
        m_State = eComplete Or lvicAsync_UserAborted
    
    Case Else
        ' called from the download or activate procedures
        If (m_State And eComplete) = eExiting Then
            m_State = eComplete Or m_State
        Else
            m_State = eComplete Or lvicAsync_UserAborted ' aborted by user, not failed download
        End If
        If Not m_Inet Is Nothing Then
            m_Inet.Abort
            Set m_Inet = Nothing
        End If
        g_AsyncController.RemoveRef Me, False
    End Select
    
    If Not bNoNotify Then
        bNoNotify = False
        RaiseEvent DownloadComplete(0&, (m_State And &HFFFF&), bNoNotify)
        If bNoNotify Then m_State = m_State Or eIDEfix
    End If
    
End Sub

Public Property Get State() As Long

    Dim bReceived As Boolean

    Select Case (m_State And &HFF000000)
    Case eNotInit ' not activated yet
        State = lvicAsync_Queued
    Case 0&         ' currently downloading
        If (m_State And &HFFFF&) = 16& Then
            State = lvicAsync_WaitingOnReply
        Else
            State = lvicAsync_Downloading
        End If
    Case Else
        If (m_State And &HFFFF&) = lvicAsync_Ok Then
            State = lvicAsync_Processing
            If (m_State And eIDEfix) = 0& Then
                RaiseEvent DownloadComplete(VarPtrArray(m_Data()), (m_State And &HFFFF&), bReceived)
                If bReceived Then m_State = m_State Or eIDEfix
            End If
        Else
            State = (m_State And &HFFFF&)
            If (m_State And eIDEfix) = 0& Then
                RaiseEvent DownloadComplete(0&, (m_State And &HFFFF&), bReceived)
                If bReceived Then m_State = m_State Or eIDEfix
            End If
        End If
        
    End Select
End Property
Public Property Let SyncMode(newValue As Boolean)
    If newValue Then m_State = (m_State And Not eNotInit) Or eSyncMode
End Property

Public Property Get URL() As String
    URL = m_URL
End Property
Public Property Let URL(newURL As String)
    If m_URL = vbNullString Then m_URL = newURL
End Property

Private Sub pvCompleted()

    Dim bReceived As Boolean

    m_State = m_State Or eComplete
    
    Set m_Inet = Nothing
    g_AsyncController.RemoveRef Me, False
    
    If (m_State And &HFFFF&) = lvicAsync_Ok Then
        RaiseEvent DownloadComplete(VarPtrArray(m_Data()), (m_State And &HFFFF&), bReceived)
    Else
        RaiseEvent DownloadComplete(0&, (m_State And &HFFFF&), bReceived)
    End If
    If bReceived Then m_State = m_State Or eIDEfix
    
End Sub

Private Sub Class_Initialize()
    m_State = eNotInit
End Sub

Private Sub Class_Terminate()
    If Not m_Inet Is Nothing Then
        m_State = -1&
        m_Inet.Abort
        Set m_Inet = Nothing
    End If
End Sub
