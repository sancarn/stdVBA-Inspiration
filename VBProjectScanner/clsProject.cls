VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Rs As ADODB.Recordset
Dim m_External As Boolean

Public Sub SetSource(ICallback As IEvents, FileName As String)

    ' called to process vbp,vbg files
    
    Dim bOk As Boolean, lErr As ValidationConstants
    Dim hHandle As Long, vBookMark As Variant
    Dim rs As ADODB.Recordset, n&
    Dim cProjExt As clsProject, cSource As clsSourceFile
    Dim cEvents As IEvents
    
    If m_External = False Then modMain.GlobalsInitialize
    Set m_Rs = gParsedItems
    Set cEvents = ICallback

    If m_External = False Then
        pvProcessVBG FileName, hHandle, ICallback, lErr
        If lErr = 0 Then
            If hHandle = 0 Then hHandle = modMain.SetSourceData(FileName, 0, lErr)
        End If
    Else
        hHandle = modMain.SetSourceData(FileName, &H80000000, lErr)
    End If
    If lErr <> 0 Then
        If m_External = False Then ICallback.ParseError lErr
        GoTo exitRoutine
    End If
    
    n = InStrRev(FileName, chrSlash)
    Call ICallback.Status("Processing " & Mid$(FileName, n + 1))
    bOk = modMain.ScanProject(m_External)
    modMain.CloseHandle hHandle
    Set gSourceFile.Owner = Nothing
    If bOk = False Then
        ICallback.ParseError vnFileInvalid
        GoTo exitRoutine
    End If
    
    If m_External = False Then Call ICallback.ParsedBegin(False)
    Set rs = m_Rs.Clone
    ' get listing of all code files for the project & parse them
    rs.Bookmark = gSourceFile.ProjBookMark
    rs.Filter = modMain.SetQuery(recType, qryIs, itSourceFile, _
                    qryAnd, recParent, qryIs, rs.Fields(recID).Value)
    Do While rs.EOF = False
        Set cSource = New clsSourceFile
        Set cEvents = ICallback
        cSource.SetSource cEvents, rs.Fields(recAttr).Value, rs.Fields(recID).Value, m_External, n
        If n <> 0 Then lErr = n: n = 0
        rs.MoveNext
    Loop
    Set cSource = Nothing
    
    If m_External = False Then
        ' process any external projects
        rs.Filter = modMain.SetQuery(recType, qryIs, itReference, qryAnd, _
                                          recFlags, qryIs, iaExternalProj)
        If rs.EOF = True Then
            bOk = False
        Else
            Set cProjExt = New clsProject
            cProjExt.IsExternal = True: bOk = True
            vBookMark = gSourceFile.ProjBookMark
            Do
                gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recID).Value), , , 1&
                gSourceFile.ProjBookMark = gParsedItems.Bookmark
                Set cEvents = ICallback
                cProjExt.SetSource cEvents, rs.Fields(recAttr).Value
                rs.MoveNext
            Loop Until rs.EOF = True
            Set cProjExt = Nothing
            gSourceFile.ProjBookMark = vBookMark: vBookMark = Empty
        End If
        Call ICallback.Status("Resolving events for the project")
        modMain.ResolveEvents bOk
        Call ICallback.ParseComplete
    End If
    rs.Close: Set rs = Nothing
    If lErr <> 0 Then
        If m_External = False Then ICallback.ParseError lErr
    End If
    
exitRoutine:
    Set ICallback = Nothing
    Set cEvents = Nothing
    modMain.GlobalsRelease
    
End Sub

Public Property Let IsExternal(Value As Boolean)
    m_External = Value
End Property

Private Sub pvClear()
    
    gSourceFile.ProjBookMark = Empty
    If Not m_Rs Is Nothing Then
        m_Rs.Filter = modMain.SetQuery(recType, qryIs, itSourceFile, qryAnd, recParent, qryNot, -1&)
        Do Until m_Rs.EOF = True
            modMain.DeleteFileW StrPtr(m_Rs.Fields(recAttr2).Value)
            m_Rs.MoveNext
        Loop
        m_Rs.Close: Set m_Rs = Nothing
    End If
    
End Sub

Private Sub pvProcessVBG(FileName As String, hHandle As Long, ICallback As IEvents, lErr As ValidationConstants)

    ' look for vbg format
    ' if exists, only process the startup project

    Dim n&, p&, sPath$
    Dim lStart&, lMax&
    
    hHandle = modMain.SetSourceData(FileName, &H80000000, lErr)
    If lErr <> 0 Then GoTo abortRoutine
    
    Do
        modMain.ParseNextLine lMax + 1, lStart, lMax
        If lStart = lMax Then Exit Do
        If p = 0 Then
            If Mid$(gSourceFile.Text, lStart, lMax - lStart) <> "VBGROUP 5.0" Then
                modMain.CreateRecord 0, vbNullString, itProject, 0, 0, FileName
                Exit Sub
            End If
            p = lStart
        Else
            p = lStart
            modMain.ParseNextWordEx p, lMax, n, p, wbpEqualEOW
            If gSourceFile.Data(n) = vbKeyS Then
                If Mid$(gSourceFile.Text, n, p - n) = "StartupProject" Then
                    If ICallback.ParsedBegin(True) = False Then
                        lErr = vnAborted
                        GoTo abortRoutine
                    End If
                    modMain.CreateRecord -1, vbNullString, itProject, 0, 0, FileName
                    modMain.ParseNextWordEx p, lMax, n, 0&, wbpEqualEOW
                    p = InStrRev(FileName, chrSlash)
                    If p <> 0 Then sPath = Left$(FileName, p)
                    FileName = Mid$(gSourceFile.Text, n, lMax - n)
                    FileName = modMain.ResolveRelativePath(FileName, sPath)
                    modMain.CloseHandle hHandle: hHandle = 0
                    Exit Sub
                End If
            End If
        End If
    Loop
    lErr = vnFileInvalid

abortRoutine:
    If hHandle <> 0 And hHandle <> -1 Then
        modMain.CloseHandle hHandle: hHandle = 0
    End If

End Sub

Private Sub Class_Terminate()
    If m_External = False Then
        Call pvClear
        If Not gParsedItems Is Nothing Then
            If gParsedItems.State <> adStateClosed Then gParsedItems.Close
            Set gParsedItems = Nothing
        End If
    End If
End Sub
