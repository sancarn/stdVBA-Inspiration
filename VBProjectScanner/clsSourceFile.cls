VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSourceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempPathW& Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long)
Private Declare Function GetTempFileNameW& Lib "kernel32" (ByVal lpPathName As Long, ByVal lpPrefixString As Long, ByVal uUnique As Long, ByVal lpTempFileName As Long)

Dim m_HandleOffsets&, lError&
Dim m_Name$, m_Type&, m_Attrs&, m_RecID&
Dim m_DefTypes() As Byte, m_DefTypeDone As Boolean

Public Function SetSource(ICallback As IEvents, FileName As String, lParent As Long, _
                    IsExternal As Boolean, lErr As ValidationConstants) As Boolean

    ' called to parse code pages

    Dim hHandle&, n&
    
    hHandle = modMain.SetSourceData(FileName, lParent, lErr)
    If hHandle = 0 Or hHandle = -1 Then Exit Function
    
    If IsExternal = True Then m_Attrs = iaExternalProj
    n = InStrRev(FileName, chrSlash)
    Call ICallback.Status("Processing " & Mid$(FileName, n + 1))
    Set gSourceFile.Owner = Me
    If modMain.ScanHeader(IsExternal) = True Then
        If IsExternal = False Then
            gParsedItems.Bookmark = gSourceFile.CPBookMark
            m_Name = gParsedItems.Fields(recName).Value
            m_Type = gParsedItems.Fields(recType).Value
            m_Attrs = CLng(gParsedItems.Fields(recFlags).Value)
            m_RecID = gParsedItems.Fields(recID).Value
            modMain.ScanSource
        End If
    ElseIf IsExternal = False Then
        lError = vnFileInvalid
    End If
    modMain.CloseHandle hHandle
    Set gSourceFile.Owner = Nothing
    Set ICallback = Nothing
    lErr = lError

End Function

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get FileType() As ItemTypeEnum
    FileType = m_Type
End Property

Public Property Get FileAttrs() As ItemTypeAttrEnum
    FileAttrs = m_Attrs
End Property

Public Property Get RecordID() As Long
    RecordID = m_RecID
End Property

Public Sub SetError(lErr As ValidationConstants)
    lError = lErr
End Sub

Public Function IsDefTyped(c As Integer) As Boolean
    If m_DefTypeDone = False Then
        m_DefTypeDone = True
        modMain.GetDefTypes m_RecID, m_DefTypes()
    End If
    Select Case c
    Case vbKeyA To vbKeyZ
        IsDefTyped = CBool(m_DefTypes(c - 65) = 1)
    Case 97 To 122
        IsDefTyped = CBool(m_DefTypes(c - 97) = 1)
    End Select
    
End Function

Public Sub LogOffsets(ptrOffsets As Long, Size As Long)

    ' passed array are file offsets to statements that are
    ' not cached during initial scan. These offsets prevent
    ' having to reparse the code page for the statements if
    ' a validation check is wanted. Simply prevents having
    ' to do the same thing twice
    
    Dim lValue&, sFile$
    
    If m_HandleOffsets = 0 Then
        pvGetTempFileName sFile
        m_HandleOffsets = modMain.GetFileHandle(sFile, True)
        If m_HandleOffsets = 0 Then
            m_HandleOffsets = -1: Exit Sub
        End If
        With gParsedItems
            .Bookmark = gSourceFile.CPBookMark
            .Find modMain.SetQuery(recID, qryIs, .Fields(recParent).Value), , , 1&
            .Fields(recAttr2).Value = sFile
            .Update
        End With
    End If
    If m_HandleOffsets <> -1 Then
        modMain.WriteFile m_HandleOffsets, ByVal ptrOffsets, Size, lValue
    End If
End Sub

Public Sub LogStats(ParamArray vStats() As Variant)

    Dim n&
    
    For n = 0 To UBound(vStats) Step 2
        modMain.CreateRecord m_RecID, "(Stats)", itStats, 0, 0, , , , CLng(vStats(n)), CLng(vStats(n + 1))
    Next

End Sub

Private Sub pvGetTempFileName(sFile As String)

    ' simply return a temp file name in user's temp folder
    
    Dim lLen&, sFolder$
    Const MAX_PATH& = 260
    
    lLen = MAX_PATH
    sFolder = Space$(lLen)
    lLen = GetTempPathW(lLen, StrPtr(sFolder))
    sFolder = Left$(sFolder, lLen)
    
    sFile = String$(MAX_PATH, vbNullChar)
    If GetTempFileNameW(StrPtr(sFolder), StrPtr("scn"), 0, StrPtr(sFile)) <> 0 Then
        lLen = InStr(sFile, vbNullChar)
        sFile = Left$(sFile, lLen - 1)
    End If

End Sub

Private Sub Class_Terminate()
    If m_HandleOffsets <> 0 Then
        If m_HandleOffsets <> -1 Then modMain.CloseHandle m_HandleOffsets
    End If
    
End Sub
