VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Basic implementation of Chrome Devtools Protocol (CDP)
' using VBA






Private objBrowser As clsExec

' every message sent to edge has an id
Private lngLastID As Long

Private strSessionId As String

'this buffer holds messages that are not yet processed
Private strBuffer As String

Public Function serialize() As String

    Dim objSerialize As Dictionary
    Set objSerialize = New Dictionary
    Call objSerialize.Add("objBrowser", objBrowser.serialize())
    Call objSerialize.Add("lngLastID", lngLastID)
    Call objSerialize.Add("strSessionId", strSessionId)
    serialize = ConvertToJson(objSerialize)


End Function

Public Sub deserialize(strSerialized As String)
    Dim objSerialize As Dictionary
    Set objSerialize = ParseJson(strSerialized)
    
    Set objBrowser = New clsExec
    Call objBrowser.deserialize(objSerialize.Item("objBrowser"))
    
    lngLastID = objSerialize.Item("lngLastID")
    strSessionId = objSerialize.Item("strSessionId")

End Sub

' CDP messages received from chrome are null-terminated
' It seemed to me you cant search for vbnull in a string
' in vba. Thats why i re-implemented the search function

Private Function searchNull() As Long
    Dim i As Long
    
    Dim lngBufferLength As Long
    lngBufferLength = Len(strBuffer)
    searchNull = 0
    
    If lngBufferLength = 0 Then
        Exit Function
    End If
    
    For i = 1 To lngBufferLength
        If Mid(strBuffer, i, 1) = vbNullChar Then
            searchNull = i
            Exit Function
        End If
    Next i

End Function

Private Function sendMessage(strMessage As String, Optional objAllMessages As Dictionary) As Dictionary

    Dim intRes As Long
    Dim strRes As String
    
    Dim lngCurrentId As Long
    lngCurrentId = lngLastID
    
    'We increase the global ID counter
    lngLastID = lngLastID + 1
    
    
    ' Before sending a message the messagebuffer is emptied
    ' All messages that we have received sofar cannot be an answer
    ' to the message that we will send
    ' So they can be safely discarded
    
    If objAllMessages Is Nothing Then
        intRes = 1
        Do Until intRes < 1
            intRes = objBrowser.readProcCDP(strRes)
            
            If intRes > 0 Then
                strBuffer = strBuffer & strRes
            End If
        Loop
        
        Dim lngNullCharPos As Long
        lngNullCharPos = searchNull()
        
        Do Until lngNullCharPos = 0
        
            'Debug.Print (Left(strBuffer, lngNullCharPos))
            strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
                        
            lngNullCharPos = searchNull()
    
        Loop
    End If

    ' sometimes edge writes to stdout
    ' we clear stdout here, too.
    
    intRes = objBrowser.readProcSTD(strRes)
    

    ' We add the currentID and sessionID to the message
    
    strMessage = Left(strMessage, Len(strMessage) - 1)
    
    If strSessionId <> "" Then
        strMessage = strMessage & ", ""sessionId"":""" & strSessionId & """"
    End If
    
    strMessage = strMessage & ", ""id"":" & lngCurrentId & "}" & vbNullChar
    
    Call objBrowser.writeProc(strMessage)
    
    
    
    ' We have some failsafe counter in order not to
    ' loop forever
    Dim intCounter As Integer
    intCounter = 0
    
    
    Do Until intCounter > 1000
        intRes = 1
        ' We read from edge and process messages until we receive a
        ' message with our ID
        Do Until intRes < 1
            intRes = objBrowser.readProcCDP(strRes)
            
            If intRes > 0 Then
                strBuffer = strBuffer & strRes
            End If
        Loop
        
        lngNullCharPos = searchNull()
        
        Do Until lngNullCharPos = 0
        
            
            strRes = Left(strBuffer, lngNullCharPos - 1)
            
            strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
                        
            Dim objRes As Dictionary
            Dim objDic2 As Dictionary
            Dim objDic3 As Dictionary
            
            Dim boolFound As Boolean
            boolFound = False
            
            'Debug.Print (strRes)
            
            
            If strRes <> "" Then
        
                'Set objRes = getJSONCollection(strRes)
                Set objRes = JsonConverter.ParseJson(strRes)
                
                If Not (objAllMessages Is Nothing) Then
                    objAllMessages.Add CStr(objAllMessages.Count), objRes
                End If
                
                
                If objRes.Exists("id") Then
        
                
                    If objRes.Item("id") = lngCurrentId Then
                        Set sendMessage = objRes
                        Exit Function
                    End If
                End If
            End If
            lngNullCharPos = searchNull()

        Loop
        DoEvents
        Call Sleep(0.1)
        intCounter = intCounter + 1
    Loop
    
    Debug.Print ("-----")
    Debug.Print ("timeout")
    Debug.Print (strMessage)
    Debug.Print ("-----")
    
    Call err.Raise(-900, , "timeout")


End Function

' This function allows to evaulate a javascript expression
' in the context of the  page
Public Function jsEval(strString As String, Optional boolRetry = True) As Variant

    Dim objRes As Dictionary
    
    Dim strMessage As String
    
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    
    Dim objMessage As Dictionary
    Set objMessage = JsonConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    
    strMessage = JsonConverter.ConvertToJson(objMessage)
    
    Set objRes = sendMessage(strMessage)
    
    If objRes Is Nothing And boolRetry Then
        Stop
        Set objRes = sendMessage(strMessage)
    End If
    
    
    If (objRes.Exists("error")) Then
        ' Oops, there was an error in out javascript expression
        Stop
        Exit Function
    End If
    
    ' If the return type has a specific type
    ' we can return the result
    
    If objRes.Item("result").Item("result").Item("type") = "string" Or objRes.Item("result").Item("result").Item("type") = "boolean" Or objRes.Item("result").Item("result").Item("type") = "number" Then
        jsEval = objRes.Item("result").Item("result").Item("value")
    End If

End Function


' This function must be calles after start and before all other methods
' This function attaches to a session of the browser
Public Function attach(strUrl As String) As Integer
    

    Dim objRes As Dictionary
    
    Dim objAllMessages As Dictionary
    Set objAllMessages = New Dictionary
     
    Set objRes = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":true}}", objAllMessages)
    
    Dim i As Integer
    Dim boolFound As Boolean
    
    Dim strKey As Variant
    
    Dim objDic2 As Dictionary
    Dim objDic3 As Dictionary
    
    
    For Each strKey In objAllMessages.Keys
    
        Set objRes = objAllMessages.Item(strKey)
        
        
        If Not objRes.Exists("params") Then GoTo nextloop1
        Set objDic2 = objRes.Item("params")
        
        If Not objDic2.Exists("targetInfo") Then GoTo nextloop1
        Set objDic3 = objDic2.Item("targetInfo")
        
        If objDic3.Item("type") <> "page" Then GoTo nextloop1
        
        If objDic3.Item("url") <> strUrl And strUrl <> "" Then
            GoTo nextloop1
        End If
        
        boolFound = True
        Exit For

nextloop1:
    Next strKey
    
    If Not boolFound Then
        attach = -1
        Exit Function
    End If
    
    'Stop
    
    Set objRes = sendMessage("{""method"":""Target.attachToTarget"",""params"":{""targetId"":""" & objDic3.Item("targetId") & """,""flatten"":true}}")
    
    strSessionId = objRes.Item("result").Item("sessionId")
    
    Set objRes = sendMessage("{""method"":""Runtime.enable"",""params"":{}}")
    
    Set objRes = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":false}}")

    
    attach = 0
    
    Call Sleep
    

End Function

' This function makes edhe naviagte to a given URL
Public Sub navigate(strUrl As String)
    Dim objRes As Dictionary
    Set objRes = sendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strUrl & """}}")
    Call Sleep
End Sub



' This method starts up the browser
Public Sub start(Optional boolSerializable As Boolean = False)
    Set objBrowser = New clsExec
    
    Dim strCall As String
    strCall = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"" --remote-debugging-pipe --enable-automation --enable-logging"
    
    Dim intRes As Integer
    
    intRes = objBrowser.init(strCall, boolSerializable)
    
    If intRes <> 0 Then
       Call err.Raise(-99, , "error start browser")
    End If

    Call Sleep
    lngLastID = 1
    
    Dim strRes As String
    
    intRes = 0
    
    Dim intCounter As Integer
    intCounter = 0
    
    Do Until intRes > 0 Or intCounter > 1000
        intRes = objBrowser.readProcSTD(strRes)
        DoEvents
        Call Sleep(0.1)
        intCounter = intCounter + 1
    Loop
    
End Sub

Public Sub closeBrowser()

    Dim objRes As Dictionary
    
    On Error Resume Next
    Set objRes = sendMessage("{""method"":""Browser.close"",""params"":{}}")
    
    'it seems without waitng a bit the browser crashes and the next time wants ro recover from a crash
    Call Sleep(5)

End Sub

Public Function connectionAlive() As Boolean
    On Error GoTo err
    Dim strLoc As String
    strLoc = jsEval("window.location.href")
    
    connectionAlive = True
    Exit Function
    
err:

    connectionAlive = False
    
End Function

Public Sub waitCompletion()
    Dim strState As String
    strState = "x"
    Call Sleep(1)
    Do Until strState = "complete"
        strState = Me.jsEval("document.readyState")
        Call Sleep(1)
    Loop
    
End Sub
