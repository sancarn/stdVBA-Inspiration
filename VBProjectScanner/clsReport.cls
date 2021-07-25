VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ParTabTab = vbNewLine & vbTab & vbTab
Const chrsColonSpc = ": ", chrsInMthd = " in method ", chrsInDec = " in Declarations"
Const chrsMthdColon = "Method: ", chrsVarColon = "Variable: ", chrsApiColon = "API: "
Const chrsConstColon = "Constant: ", chrsEnumColon = "Enum: ", chrsDecs = " Declarations"
Const chrsIn = " in ", chrsDupes = " duplicates "

Public Sub CreateReport_Standard(rtfCanvas As RichTextBox, _
                        tvDisplay As TreeView, tParent As Node, _
                        ICallback As IEvents, CallbackParam As Long)

    ' either creates a report in RTF format or summarizes findings
    ' if tvDisplay is nothing, then create report else summarize

    Dim sList$, n&, t&
    Dim sBlock$(), sTemplate$()
    Dim rsFiles As ADODB.Recordset
    Dim rsCodePg As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Const ParTab = vbNewLine & vbTab
    
    sTemplate() = Split(vbTab & "Option Explicit not used" & ParTab & _
            "Methods having no executable statements" & ParTab & _
            "Methods having Stop or End statements" & ParTab & _
            "Items not VarTyped" & ParTab & _
            "Items in zombie state" & ParTab & _
            "Variables declared WithEvents not trapping events" & ParTab & _
            "Variables using ReDim on non-declared variables" & ParTab & _
            "Variant function vs. String function instances" & ParTab & _
            "Duplicated declaration names" & ParTab & _
            "Duplicated literals", vbNewLine)

    Set rsFiles = gParsedItems.Clone
    rsFiles.Bookmark = gSourceFile.ProjBookMark
    Set rsCodePg = gParsedItems.Clone
    rsCodePg.Filter = modMain.SetQuery(recType, qryIs, itCodePage, qryAnd, recFlags, qryGT, -1)
    Set rs = gParsedItems.Clone
    rs.Sort = recType & chrComma & recParent & chrComma & recIdxName
    
    rsCodePg.MoveFirst
    Do Until rsCodePg.EOF = True
        If (rsCodePg.Fields(recFlags) And iaOpExplicit) = 0 Then t = t + 1
        rsCodePg.MoveNext
    Loop
    
    sBlock() = sTemplate
    sBlock(0) = sBlock(0) & chrsColonSpc & CStr(t)
    rs.Filter = modMain.SetQuery(recType, qryIs, itMethod, qryAnd, recDiscrep, qryLike, "%E%")
    sBlock(1) = sBlock(1) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itMethod, qryAnd, recDiscrep, qryLike, "%X%")
    sBlock(2) = sBlock(2) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recDiscrep, qryLike, "%V%")
    sBlock(3) = sBlock(3) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recDiscrep, qryLike, "%Z%")
    sBlock(4) = sBlock(4) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itVariable, qryAnd, recDiscrep, qryLike, "%W%")
    sBlock(5) = sBlock(5) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itVariable, qryAnd, recDiscrep, qryLike, "%R%")
    sBlock(6) = sBlock(6) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itDiscrep, qryAnd, recFlags, qryIs, vtVarFunc)
    sBlock(7) = sBlock(7) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itDiscrep, qryAnd, recFlags, qryIs, vtDupeDecs)
    sBlock(8) = sBlock(8) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itDiscrep, qryAnd, recFlags, qryIs, vtDupeLiterals)
    sBlock(9) = sBlock(9) & chrsColonSpc & CStr(rs.RecordCount)
    
    If Not tvDisplay Is Nothing Then
        For t = 0 To UBound(sBlock)
            tvDisplay.Nodes.Add tParent, tvwChild, , sBlock(t)
        Next
        tParent.Expanded = True
        GoTo exitRoutine
    End If
    
    ' the report is initialized with stub statements that are replaced
    ' one stub per code page
    
    For n = 1 To rsCodePg.RecordCount
        sList = sList & LoadResString(6)
    Next
    rtfCanvas.TextRTF = Replace(LoadResString(5), "%NL%", vbNewLine) & sList & "\par\par}"
    rtfCanvas.SelStart = rtfCanvas.Find("%Project%")
    rtfCanvas.SelLength = 9
    rtfCanvas.SelText = "Project: " & rsFiles.Fields(recName).Value & vbNewLine & _
                        "Location: " & rsFiles.Fields(recAttr).Value & vbNewLine & vbNewLine & _
                        "Files Scanned: " & CStr(rsCodePg.RecordCount)
                        
    rtfCanvas.SelStart = rtfCanvas.Find("%Summary%", rtfCanvas.SelStart)
    rtfCanvas.SelLength = 9
    rtfCanvas.SelText = Join(sBlock(), vbNewLine)
    
    rsCodePg.Sort = recIdxName
    Do Until rsCodePg.EOF = True
        t = InStr(rsCodePg.Fields(recName).Value, chrDot)
        ICallback.Status "Build Standard Report: " & Mid$(rsCodePg.Fields(recName).Value, t + 1)
        
        sBlock() = sTemplate
        pvLog_OptionExplicit rsCodePg, sBlock(0)
        pvLog_EmptyCode rs, sBlock(1), rsCodePg.Fields(recID).Value
        pvLog_StopEnd rs, sBlock(2), rsCodePg.Fields(recID).Value
        pvLog_VarType rs, sBlock(3), rsCodePg.Fields(recID).Value
        pvLog_Zombies rs, sBlock(4), rsCodePg.Fields(recID).Value
        pvLog_WithNoEvents rs, sBlock(5), rsCodePg.Fields(recID).Value
        pvLog_ReDim rs, sBlock(6), rsCodePg.Fields(recID).Value
        pvLog_VarFuncs rs, sBlock(7), rsCodePg.Fields(recID).Value
        pvLog_DupeDecs rs, sBlock(8), rsCodePg.Fields(recID).Value
        pvLog_DupeLiterals rs, sBlock(9), rsCodePg.Fields(recID).Value
        
        rtfCanvas.SelStart = rtfCanvas.Find("%Page%", rtfCanvas.SelStart, , rtfWholeWord)
        rtfCanvas.SelLength = 6
        t = InStr(rsCodePg.Fields(recName).Value, chrDot)
        rtfCanvas.SelText = Mid$(rsCodePg.Fields(recName).Value, t + 1)
        
        sList = Join(sBlock(), vbNullString)
        If LenB(sList) = 0 Then
            sList = vbTab & "No issues noted"
        Else
            sList = Left$(sList, Len(sList) - 2)
        End If
        rtfCanvas.SelStart = rtfCanvas.Find("%Results%", rtfCanvas.SelStart, , rtfWholeWord)
        rtfCanvas.SelLength = 9
        rtfCanvas.SelText = sList
        
        rsCodePg.MoveNext
    Loop
    rtfCanvas.SelStart = 0
    ICallback.ReportComplete 0, CallbackParam
    
exitRoutine:
    rsFiles.Close: Set rsFiles = Nothing
    rsCodePg.Close: Set rsCodePg = Nothing
    rs.Close: Set rs = Nothing
    
End Sub

Public Sub CreateReport_Safety(rtfCanvas As RichTextBox, _
                        tvDisplay As TreeView, tParent As Node, _
                        ICallback As IEvents, CallbackParam As Long)

    ' Safety check, malicious code check
    ' either creates a report in RTF format or summarizes findings
    ' if tvDisplay is nothing, then create report else summarize

    Dim sList$, n&, t&
    Dim sBlock$(), sTemplate$()
    Dim rsFiles As ADODB.Recordset
    Dim rsCodePg As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Const ParTab = vbNewLine & vbTab
    
    sTemplate() = Split(vbTab & _
            "VB Function Instances" & ParTab & _
            "Windows API Instances", vbNewLine)

    Set rsFiles = gParsedItems.Clone
    rsFiles.Bookmark = gSourceFile.ProjBookMark
    Set rsCodePg = gParsedItems.Clone
    rsCodePg.Filter = modMain.SetQuery(recType, qryIs, itCodePage, qryAnd, recFlags, qryGT, -1)
    Set rs = gParsedItems.Clone
    rs.Sort = recType & chrComma & recParent & chrComma & recIdxName
    
    sBlock() = sTemplate
    rs.Filter = modMain.SetQuery(recType, qryIs, itDiscrep, qryAnd, recFlags, qryIs, vtMalicious)
    sBlock(0) = sBlock(0) & chrsColonSpc & CStr(rs.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itAPI, qryAnd, recDiscrep, qryLike, "%M%")
    sBlock(1) = sBlock(1) & chrsColonSpc & CStr(rs.RecordCount)
    
    sList = Replace(rsFiles.Fields(recAttr).Value, ".vbp", ".vbw", , , vbTextCompare)
    t = modMain.GetFileLastModDate(sList, 0&, 0&, t)
    If Not tvDisplay Is Nothing Then
        If t = 0 Then
            tvDisplay.Nodes.Add tParent, tvwChild, , "VBW file exists: Yes"
        Else
            tvDisplay.Nodes.Add tParent, tvwChild, , "VBW file exists: No"
        End If
        For n = 0 To UBound(sBlock)
            tvDisplay.Nodes.Add tParent, tvwChild, , sBlock(n)
        Next
        tParent.Expanded = True
        GoTo exitRoutine
    End If
    
    ' the report is initialized with stub statements that are replaced
    ' one stub per code page
    
    sList = vbNullString
    For n = 1 To rsCodePg.RecordCount
        sList = sList & LoadResString(6)
    Next
    rtfCanvas.TextRTF = Replace(LoadResString(5), "%NL%", vbNewLine) & sList & "\par\par}"
    n = rtfCanvas.Find("%Project%")
    rtfCanvas.SelStart = n
    rtfCanvas.SelLength = 9
    rtfCanvas.SelText = "Project: " & rsFiles.Fields(recName).Value & vbNewLine & _
                        "Location: " & rsFiles.Fields(recAttr).Value & vbNewLine & vbNewLine & _
                        "Files Scanned: " & CStr(rsCodePg.RecordCount)
                        
    n = rtfCanvas.Find("%Summary%", n)
    rtfCanvas.SelStart = n
    rtfCanvas.SelLength = 9
    If t = 0 Then
        rtfCanvas.SelText = Join(sBlock(), vbNewLine) & ParTab & "VBW file exists: Yes"
    Else
        rtfCanvas.SelText = Join(sBlock(), vbNewLine) & ParTab & "VBW file exists: No"
    End If
    
    rsCodePg.Sort = recIdxName
    Do Until rsCodePg.EOF = True
        t = InStr(rsCodePg.Fields(recName).Value, chrDot)
        ICallback.Status "Build Safety Report: " & Mid$(rsCodePg.Fields(recName).Value, t + 1)
        
        sBlock() = sTemplate
        pvLog_MalFuncs rs, sBlock(0), rsCodePg.Fields(recID).Value
        pvLog_MalAPIs rs, sBlock(1), rsCodePg.Fields(recID).Value
        
        n = rtfCanvas.Find("%Page%", n)
        rtfCanvas.SelStart = n
        rtfCanvas.SelLength = 6
        t = InStr(rsCodePg.Fields(recName).Value, chrDot)
        rtfCanvas.SelText = Mid$(rsCodePg.Fields(recName).Value, t + 1)
        
        sList = Join(sBlock(), vbNullString)
        If LenB(sList) = 0 Then
            sList = vbTab & "No issues noted"
        Else
            sList = Left$(sList, Len(sList) - 2)
        End If
        n = rtfCanvas.Find("%Results%", n)
        rtfCanvas.SelStart = n
        rtfCanvas.SelLength = 9
        rtfCanvas.SelText = sList
        
        rsCodePg.MoveNext
    Loop
    rtfCanvas.SelStart = 0
    ICallback.ReportComplete 1, CallbackParam
    
exitRoutine:
    rsFiles.Close: Set rsFiles = Nothing
    rsCodePg.Close: Set rsCodePg = Nothing
    rs.Close: Set rs = Nothing
    
End Sub

Private Sub pvLog_OptionExplicit(rs As ADODB.Recordset, sResults As String)

    If (rs.Fields(recFlags) And iaOpExplicit) = 0 Then
        sResults = sResults & vbNewLine
    Else
        sResults = vbNullString
    End If

End Sub

Private Sub pvLog_EmptyCode(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList As String
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recDiscrep, qryLike, "%E%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            If (rs.Fields(recType) = itCodePage) Then
                sList = sList & ParTabTab & "Code Page: "
            Else
                sList = sList & ParTabTab & chrsMthdColon
            End If
            sList = sList & rs.Fields(recName).Value
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_StopEnd(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList As String
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recDiscrep, qryLike, "%X%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            sList = sList & ParTabTab & chrsMthdColon & rs.Fields(recName).Value
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_VarType(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList As String, t As Long
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recDiscrep, qryLike, "%V%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            t = rs.Fields(recType).Value
            Select Case t
            Case itAPI
                sList = sList & ParTabTab & chrsApiColon & rs.Fields(recName).Value
            Case itMethod
                sList = sList & ParTabTab & chrsMthdColon & rs.Fields(recName).Value
            Case itVariable
                sList = sList & ParTabTab & chrsVarColon & rs.Fields(recName).Value
                If rs.Fields(recParent).Value = lCodePg Then
                    sList = sList & chrsInDec
                Else
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
                    sList = sList & chrsInMthd & gParsedItems.Fields(recName).Value
                End If
            Case itParameter
                sList = sList & ParTabTab & "Parameter: " & rs.Fields(recName).Value
                gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
                t = gParsedItems.Fields(recType).Value
                If t = itMethod Then
                    sList = sList & chrsInMthd & gParsedItems.Fields(recName).Value
                ElseIf t = itEvent Then
                    sList = sList & " in public event " & gParsedItems.Fields(recName).Value
                ElseIf t = itClassEvent Then
                    sList = sList & chrsInMthd & gParsedItems.Fields(recName).Value
                ElseIf t = itAPI Then
                    sList = sList & " in API " & gParsedItems.Fields(recName).Value
                End If
            End Select
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_Zombies(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList$, sEnum$, lEnum&
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recDiscrep, qryLike, "%Z%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            Select Case rs.Fields(recType).Value
            Case itAPI
                sList = sList & ParTabTab & chrsApiColon & rs.Fields(recName).Value
            Case itMethod
                sList = sList & ParTabTab & chrsMthdColon & rs.Fields(recName).Value
            Case itVariable
                sList = sList & ParTabTab & chrsVarColon & rs.Fields(recName).Value
                If rs.Fields(recParent).Value = lCodePg Then
                    sList = sList & chrsInDec
                Else
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
                    sList = sList & chrsInMthd & gParsedItems.Fields(recName).Value
                End If
            Case itConstant
                sList = sList & ParTabTab & chrsConstColon & rs.Fields(recName).Value
                If rs.Fields(recParent).Value = lCodePg Then
                    sList = sList & chrsInDec
                Else
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
                    sList = sList & chrsInMthd & gParsedItems.Fields(recName).Value
                End If
            Case itType
                sList = sList & ParTabTab & "Type/UDT: " & rs.Fields(recName).Value
            Case itEnum
                sList = sList & ParTabTab & "Enumeration: " & rs.Fields(recName).Value & " (entirety)"
            Case itEnumMember
                If lEnum <> rs.Fields(recParent).Value Then
                    lEnum = rs.Fields(recParent).Value
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, lEnum), , , 1&
                    sEnum = gParsedItems.Fields(recName).Value
                End If
                sList = sList & ParTabTab & "Enum Member: " & sEnum & chrDot & rs.Fields(recName).Value
            Case itParameter
                sList = sList & ParTabTab & "Parameter: " & rs.Fields(recName).Value
                gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
                sList = sList & chrsInMthd & gParsedItems.Fields(recName).Value
            Case itCodePage
                sList = sList & ParTabTab & "Code Page: This code page in its entirety"
            End Select
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    
    End If
    
End Sub

Private Sub pvLog_WithNoEvents(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList As String
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recDiscrep, qryLike, "%W%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            sList = sList & ParTabTab & chrsVarColon & rs.Fields(recName).Value & chrsInDec
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_ReDim(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList As String
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recDiscrep, qryLike, "%R%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
            sList = sList & ParTabTab & chrsVarColon & rs.Fields(recName).Value & _
                        chrsInMthd & gParsedItems.Fields(recName).Value
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_VarFuncs(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList$
    Dim rsMbrs As ADODB.Recordset
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itDiscrep, _
                            qryAnd, recFlags, qryIs, vtVarFunc)
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Set rsMbrs = gParsedItems.Clone
        rsMbrs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg)
        Do
            rsMbrs.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
            sList = sList & ParTabTab & "Function: " & rs.Fields(recName).Value & _
                        chrsInMthd & rsMbrs.Fields(recName).Value
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_DupeDecs(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList$, sDupe$, sTarget$, n&, t&
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itDiscrep, _
                            qryAnd, recFlags, qryIs, vtDupeDecs)
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recOffset).Value), , , 1&
            sTarget = gParsedItems.Fields(recName).Value
            Select Case gParsedItems.Fields(recType).Value
                Case itVariable: sTarget = chrsVarColon & sTarget
                Case itConstant: sTarget = chrsConstColon & sTarget
                Case itAPI: sTarget = chrsApiColon & sTarget
                Case itEnum: sTarget = chrsEnumColon & sTarget
                Case itEnumMember
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , , 1&
                    sTarget = "Member " & gParsedItems.Fields(recName).Value & chrDot & sTarget
                Case itType: sTarget = "Type/UDT " & sTarget
            End Select
            gParsedItems.Find modMain.SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , , 1&
            n = InStr(gParsedItems.Fields(recName).Value, chrDot)
            sTarget = sTarget & chrsIn & Mid$(gParsedItems.Fields(recName).Value, n + 1) & chrsDecs
            
            gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
            sDupe = gParsedItems.Fields(recName).Value
            Select Case gParsedItems.Fields(recType).Value
                Case itVariable: sDupe = chrsVarColon & sDupe
                Case itConstant: sDupe = chrsConstColon & sDupe
                Case itAPI: sDupe = chrsApiColon & sDupe
                Case itEnum: sDupe = chrsEnumColon & sDupe
                Case itEnumMember
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , , 1&
                    sDupe = "Enum Member " & gParsedItems.Fields(recName).Value & "." & sDupe
                Case itType: sDupe = "Type/UDT " & sDupe
            End Select
            If rs.Fields(recStart).Value = 0 Then
                gParsedItems.Find modMain.SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , , 1&
                sDupe = sDupe & chrsInMthd & gParsedItems.Fields(recName).Value
            End If
            
            sList = sList & ParTabTab & sDupe & chrsDupes & sTarget
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub

Private Sub pvLog_DupeLiterals(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList$, sDupe$, sTarget$, n&, t&
    Dim rsMbrs As ADODB.Recordset, sConst$
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itDiscrep, _
                            qryAnd, recFlags, qryIs, vtDupeLiterals)
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itDiscrep, _
                            qryAnd, recFlags, qryIs, vtDupeLiterals, _
                            qryAnd, recOffset, qryNot, 0)
        If rs.EOF = False Then
            sConst = ParTabTab & "Duplicates among constants"
            Do
                gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recOffset).Value), , , 1&
                sTarget = chrsConstColon & gParsedItems.Fields(recName).Value
                gParsedItems.Find modMain.SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , , 1&
                n = InStr(gParsedItems.Fields(recName).Value, chrDot)
                sTarget = sTarget & chrsIn & Mid$(gParsedItems.Fields(recName).Value, n + 1) & chrsDecs
                
                gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
                If rs.Fields(recStart).Value = 2 Then
                    sDupe = rs.Fields(recName).Value & " in Constant " & gParsedItems.Fields(recName).Value
                Else
                    sDupe = rs.Fields(recName).Value & chrsInMthd & gParsedItems.Fields(recName).Value
                End If
                sConst = sConst & ParTabTab & vbTab & sDupe & chrsDupes & sTarget
                rs.MoveNext
            Loop Until rs.EOF = True
        Else
            sConst = ParTabTab & "Duplicates among constants: none"
        End If
        
        Set rsMbrs = gParsedItems.Clone
        rsMbrs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itDiscrep, _
                            qryAnd, recFlags, qryIs, vtDupeLiterals, _
                            qryAnd, recOffset, qryIs, 0)
        If rsMbrs.EOF = False Then
            rsMbrs.Sort = recOffset2 & " DESC" & chrComma & recName
            sList = ParTabTab & "Duplicate literals within code page"
            Do
                sDupe = "x" & CStr(rsMbrs.Fields(recOffset2).Value) & ":" & vbTab & rsMbrs.Fields(recName).Value
                sList = sList & ParTabTab & vbTab & sDupe
                rsMbrs.MoveNext
            Loop Until rsMbrs.EOF = True
        Else
            sList = ParTabTab & "Duplicate literals within code page: none"
        End If
        sResults = sResults & sConst & sList & vbNewLine
        rsMbrs.Close: Set rsMbrs = Nothing
    End If
    
End Sub

Private Sub pvLog_MalFuncs(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList$
    Dim rsMbrs As ADODB.Recordset
    
    Set rs = gParsedItems.Clone
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itDiscrep, _
                            qryAnd, recFlags, qryIs, vtMalicious)
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Set rsMbrs = gParsedItems.Clone
        rsMbrs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg)
        Do
            rsMbrs.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
            sList = sList & ParTabTab & "Function: " & rs.Fields(recName).Value & _
                        chrsInMthd & rsMbrs.Fields(recName).Value
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
        rsMbrs.Close: Set rsMbrs = Nothing
    End If

End Sub

Private Sub pvLog_MalAPIs(rs As ADODB.Recordset, sResults As String, lCodePg As Long)

    Dim sList$
    
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                            qryAnd, recType, qryIs, itAPI, _
                            qryAnd, recDiscrep, qryLike, "%M%")
    If rs.RecordCount = 0 Then
        sResults = vbNullString
    Else
        Do
            sList = sList & ParTabTab & chrsApiColon & rs.Fields(recName).Value & _
                        " Alias: " & rs.Fields(recAttr).Value
            rs.MoveNext
        Loop Until rs.EOF = True
        sResults = sResults & sList & vbNewLine
    End If

End Sub


