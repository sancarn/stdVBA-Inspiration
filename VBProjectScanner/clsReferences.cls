VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKeyExW Lib "advapi32.dll" _
                (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, _
                ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExW Lib "advapi32.dll" _
                (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, _
                lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumKeyExW Lib "advapi32.dll" (ByVal hKey As Long, _
                ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, _
                ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, _
                ByVal lpftLastWriteTime As Long) As Long

Public Enum RefStateEnum
    rstCanLoad = 0                  ' TLB is loadable
    rstCustom = 1                   ' custom, see modMain.pvAddDesignerEvents
    rstNoLoad = 2                   ' failed to find/load, events will be unresolved
End Enum
Public Enum TlbLoadRequestEnum
    tlbNoLoad = 0                   ' search for fully qualified name, don't cache events
    tlbLoadEvents = 1               ' search for FQN & cache events
    tlbLoadMethods = 2              ' search for FQN & cache methods
End Enum

Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_CLASSES_ROOT = -2147483648#
Const ACCESS_READ = &H20019
Const REG_SZ = 1&

Const chrHex = "&H"
Const recFQN = "fFQN", recGUID = "fGUID", recMajor = "fMajor", recMinor = "fMinor"
Const recCRC = "fCRC", recCRC2 = "fCRC2", recLang = "fLang"
Const recMthd = "fMethods", recEvnts = "fEvents", recRefID = "fRefID"

Dim m_Rs As ADODB.Recordset
Dim m_Events As ADODB.Recordset
Dim TLInfo As TypeLibInfo
Dim m_RecID&

Public Sub ParseAndLoad(sStatement As String, bIsObj As Boolean)

    ' statements expected in these types of format...
    ' text after the equal sign is sent here
    
    ' Object lines, after equal sign, formatted like so:
    ' format #1
    '   Object={3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0; RICHTX32.OCX
    '   1st token: GUID
    '   2nd token: #majorVersion.minorVersion. Values are decimal, not hex
    '   3rd token: #LCID. Value is decimal, not hex
    '   4th token: object name, no path
    ' format #2
    '   Object=*\AC:\...\...\someProject.vbp
    '   1st token: *\A, *\C, *\B, *\D (B,D indicates path is MAC vs Windows)
    '   2nd token: path to the project
    ' format #3
    '   this format appears in code page files
    '   is either of the first 2 formats above, but wrapped in quotes
    '   Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
    
    ' References lines, after equal sign, formatted like so:
    ' Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\..\Windows\System32\stdole2.tlb#OLE Automation
    '   1st token: *\G, *\H (H indicates path is MAC vs Windows)
    '   2nd token: GUID
    '   3rd token: #majorVersion.minorVersion. Values are in hex
    '   4th token: #LCID. Value is in hex
    '   5th token: path to the reference

    Dim n&, p&, lLang&, sTokens$()
    Dim sValue$, sGuid$, sLib$
    Dim iMinor%, iMajor%, lFlags&
    Dim lState As RefStateEnum, vBkMk As Variant
    Dim rs As ADODB.Recordset
    
    If AscW(sStatement) = vbKeyQuote Then _
        sStatement = Replace$(sStatement, Chr$(vbKeyQuote), vbNullString)
    
    If AscW(sStatement) = 123 Then              ' curly bracket
        sGuid = Left$(sStatement, 38): n = 40
    Else
        Select Case Left$(sStatement, 3)
        Case "*\G", "*\H"
            n = AscW(Mid$(sStatement, 4, 1))
            If n <> 123 Then Exit Sub           ' curly bracket open
            p = AscW(Mid$(sStatement, 41, 1))
            If p <> 125 Then Exit Sub           ' curly bracket closed
            sGuid = Mid$(sStatement, 4, 38)
            n = 43
        Case "*\A", "*\C", "*\B", "*\D"
            sStatement = Mid$(sStatement, 4)
            If GetFileLastModDate(sStatement, 0&, 0&, 0&) = 0 Then
                ' 0 ParentID is a flag = external project
                Set rs = gParsedItems.Clone
                rs.Filter = modMain.SetQuery(recParent, qryIs, 0, qryAnd, recName, qryIs, sStatement)
                If rs.EOF = True Then
                    vBkMk = gSourceFile.ProjBookMark: gSourceFile.ProjBookMark = Empty
                    modMain.CreateRecord 0, vbNullString, itReference, 0, 0, sStatement, , , iaExternalProj
                    gSourceFile.ProjBookMark = vBkMk: vBkMk = Empty
                End If
                rs.Close: Set rs = Nothing
            End If
            Exit Sub
        Case Else
            Exit Sub
        End Select
    End If
    
    If m_Rs Is Nothing Then pvCreateRecordsets
    
    m_Rs.Find modMain.SetQuery(recGUID, qryIs, sGuid), , , 1&
    If m_Rs.EOF = False Then Exit Sub
    
    ' get text left of semicolon if one exists
    p = InStr(n, sStatement, chrSemi)
    If p = 0 Then p = Len(sStatement) + 1
    sTokens() = Split(Mid$(sStatement, n, p - n), chrHash)
    If UBound(sTokens) > 0 Then
        p = InStr(sTokens(0), chrDot)
        If p <> 0 Then
            If bIsObj = False Then
                iMajor = CInt(chrHex & Left$(sTokens(0), p - 1))
                iMinor = CInt(chrHex & Mid$(sTokens(0), p + 1))
                lLang = CLng(chrHex & sTokens(1))
            Else
                iMajor = CInt(Left$(sTokens(0), p - 1))
                iMinor = CInt(Mid$(sTokens(0), p + 1))
                lLang = CLng(sTokens(1))
            End If
        End If
    End If
    Erase sTokens()
    
    On Error Resume Next            ' attempt to load this version
    Set TLInfo = TLI.TypeLibInfoFromRegistry(sGuid, iMajor, iMinor, lLang)
    If Err.Number = 0 Then          ' success
        lState = rstCanLoad
    Else
        On Error GoTo 0
        lState = rstNoLoad
        sValue = pvSearchRegistry("TypeLib\" & sGuid, "win32", vbNullString)
        If LenB(sValue) <> 0 Then         ' split the return value into tokens
            n = InStr(sValue, chrSlash)
            sTokens() = Split(Mid$(sValue, n + 1), chrSlash)
            n = InStr(sTokens(1), chrDot)
            iMajor = CInt(chrHex & Left$(sTokens(1), n - 1))
            iMinor = CInt(chrHex & Mid$(sTokens(1), n + 1))
            lLang = CLng(chrHex & sTokens(2))
            On Error Resume Next            ' attempt to load this version
            Set TLInfo = TLI.TypeLibInfoFromRegistry(sGuid, iMajor, iMinor, lLang)
            If Err.Number = 0 Then lState = rstCanLoad
        End If
    End If
    On Error GoTo 0
    
    If lState = rstNoLoad Then
        sValue = "Unknown.Unregistered"
        lFlags = -1
    Else
        sValue = TLInfo.Name
        sLib = TLInfo.ContainingFile
    End If
    
    sStatement = sGuid & chrHash & CStr(iMajor And &HFFFF&) & chrDot & CStr(iMinor And &HFFFF&) & chrDot & lLang
    Set rs = gParsedItems.Clone
    rs.Filter = modMain.SetQuery(recParent, qryIs, 0, qryAnd, recAttr2, qryIs, sStatement)
    If rs.EOF = True Then
        vBkMk = gSourceFile.ProjBookMark: gSourceFile.ProjBookMark = Empty
        modMain.CreateRecord 0, sValue, itReference, 0, 0, sLib, sStatement, , lFlags
        gSourceFile.ProjBookMark = vBkMk: vBkMk = Empty
    
        m_RecID = m_RecID + 1
        n = modMain.CRCItem(LCase$(sValue), True)
        m_Rs.AddNew Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10), Array(m_RecID, sValue, sGuid, Hex$(iMajor), Hex$(iMinor), lLang, n, 0, lState, 0, 0)
        m_Rs.Update
    End If
    rs.Close: Set rs = Nothing
    Set TLInfo = Nothing
    
End Sub

Public Function AddCustomReference(sFQN As String, State As RefStateEnum) As Boolean
                        
    ' create a TLB reference that will not be queried
    ' State param indicates whether or not event will be tagged resolved/unresolved
    ' see modMain.pvAddDesignerEvents
                        
    Dim n&, p&
    
    If m_Rs Is Nothing Then pvCreateRecordsets
    
    If m_Rs.BOF = False Then            ' already added?
        m_Rs.Find modMain.SetQuery(recFQN, qryIs, sFQN), , , 1&
    End If
    If m_Rs.EOF = True Then
        m_RecID = m_RecID + 1
        n = modMain.CRCItem(LCase$(sFQN), True)
        p = modMain.CRCItem(LCase$(Mid$(sFQN, InStr(sFQN, chrDot) + 1)), True)
        m_Rs.AddNew Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10), Array(m_RecID, sFQN, "{none}", 0, 0, 0, n, p, State, 0, 0)
        m_Rs.Update
        AddCustomReference = True
    End If
    
End Function

Public Sub AddCustomEvents(sFQN As String, fEvents As Variant)
    
    ' create custom events for a custom reference
    ' see modMain.pvAddDesignerEvents
    
    If m_Rs Is Nothing Then Exit Sub
    
    Dim lID&, n&
    
    m_Rs.Find modMain.SetQuery(recFQN, qryIs, sFQN), , , 1&
    If m_Rs.EOF = True Then Exit Sub
    lID = m_Rs.Fields(0).Value
    For n = 0 To UBound(fEvents)
        m_Events.AddNew Array(0, 1, 2), Array(lID, tlbLoadEvents Or tlbLoadMethods, LCase$(fEvents(n)))
        m_Events.Update
    Next

End Sub

Public Function GetFQN(FQN As String, ByVal Action As TlbLoadRequestEnum, _
                        Optional isControl As Boolean) As Long
    
    ' routine tries to locate a class within a TLB
    ' passed Item is either fully qualified or not
    '   if not qualified, then refs are looked at first come,
    '   first served until the item is found
    
    Dim lOffsetEvt&, lOffsetMthd&, n&, p&
    Dim lFlags&, sName$, sLib$
    '0x00   not found, search TLBs top to bottom
    '0x01   found by name
    '0x02   found library (from FQN), search that TLB only
    '0x03   found after search, cache results
                        
    If m_Rs Is Nothing Then Exit Function
    If LenB(m_Rs.Sort) = 0 Then m_Rs.Sort = recID & " Desc"

    n = modMain.CRCItem(LCase$(FQN), True)
    p = InStr(FQN, chrDot)
    If p = 0 Then                   ' search by unqualified name
        sName = FQN
        m_Rs.Find modMain.SetQuery(recCRC2, qryIs, n), , , 1&
        If m_Rs.EOF = False Then    ' found
            lFlags = 1: FQN = m_Rs.Fields(1).Value
            p = InStr(FQN, chrDot)
            If p <> 0 Then sLib = Left$(FQN, p - 1)
        Else
            m_Rs.MoveFirst          ' manual search
        End If
    Else                            ' search for fully qualified name
        sName = Mid$(FQN, p + 1): sLib = Left$(FQN, p - 1)
        m_Rs.Find modMain.SetQuery(recCRC, qryIs, n), , , 1&
        If m_Rs.EOF = False Then
            lFlags = 1              ' found by FQN
        Else                        ' else search for library
            n = modMain.CRCItem(LCase$(sLib), True)
            m_Rs.Find modMain.SetQuery(recCRC, qryIs, n), , , 1&
            If m_Rs.EOF = True Then
                m_Rs.Filter = adFilterNone
                m_Rs.MoveFirst      ' manual search
            Else
                lFlags = 2          ' search this library only
            End If
        End If
    End If
    
    If lFlags = 1 Then              ' object previously processed, abort early
        If Action = tlbNoLoad Then
            GetFQN = 1: Exit Function
        End If
        n = m_Rs.Fields(recFlags).Value
        If n = rstNoLoad Then       ' failed to load this on previous attempt
            Exit Function
        ElseIf n = rstCustom Then
            GetFQN = m_Rs.Fields(0).Value: Exit Function
        End If
        lOffsetMthd = m_Rs.Fields(recMthd).Value
        lOffsetEvt = m_Rs.Fields(recEvnts).Value
    Else
        If Not TLInfo Is Nothing Then   ' unload TLB if needed
            If LenB(sLib) <> 0 Then
                If TLInfo.Name <> sLib Then Set TLInfo = Nothing
            End If
        End If
        lOffsetMthd = -1
        Do
        If m_Rs.Fields(recFlags).Value = rstCanLoad Then
            With m_Rs   ' load the library and search
                If .Fields(recFlags).Value <> rstNoLoad Then
                    If TLInfo Is Nothing Then
                        Set TLInfo = TLI.TypeLibInfoFromRegistry(.Fields(recGUID).Value, _
                            CInt(chrHex & .Fields(recMajor).Value), CInt(chrHex & .Fields(recMinor).Value), .Fields(recLang).Value)
                    End If
                    On Error Resume Next
                    lOffsetMthd = TLInfo.GetTypeInfoNumber(sName)
                    If Err Then
                        Err.Clear
                    Else ' found, get events index now that we have methods index
                        lOffsetEvt = lOffsetMthd
                        With TLInfo.GetTypeInfo(lOffsetMthd)
                            lOffsetMthd = .DefaultInterface.TypeInfoNumber
                            lOffsetEvt = lOffsetMthd
                            If .TypeKind = TKIND_COCLASS Then
                                If Not .DefaultEventInterface Is Nothing Then
                                    If .DefaultEventInterface.Members.Count <> 0 Then _
                                        lOffsetEvt = .DefaultEventInterface.TypeInfoNumber
                                End If
                            End If
                            FQN = TLInfo.Name & chrDot & .Name
                        End With
                    End If
                    On Error GoTo 0
                    If lOffsetMthd <> -1 Then
                        If Action = tlbNoLoad Then
                            GetFQN = 1: Exit Function
                        End If
                        lFlags = 3: Exit Do         ' found
                    End If
                End If
            End With
        End If
        If lFlags = 2 Then Exit Do              ' done if seaching only 1 lib
        Set TLInfo = Nothing
        m_Rs.MoveNext
        Loop Until m_Rs.EOF = True
    End If
    
    If (lFlags And 1) = 1 Then
        If lFlags = 3 Then
            m_RecID = m_RecID + 1: FQN = LCase$(FQN)
            n = modMain.CRCItem(FQN, True)
            p = InStr(FQN, chrDot)
            If p <> 0 Then p = modMain.CRCItem(Mid$(FQN, p + 1), True)
            m_Rs.AddNew Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10), _
                        Array(m_RecID, FQN, m_Rs.Fields(recGUID).Value, _
                              m_Rs.Fields(recMajor).Value, m_Rs.Fields(recMinor).Value, _
                              m_Rs.Fields(recLang).Value, n, p, _
                              0, lOffsetEvt, lOffsetMthd)
            m_Rs.Update
        End If
        p = m_Rs.Fields(0).Value
        If Action <> tlbNoLoad Then                 ' cache events/methods
            m_Events.Filter = modMain.SetQuery(recRefID, qryIs, p, qryAnd, recType, qryIs, Action)
            If m_Events.EOF = True Then             ' unless already cached
                On Error Resume Next
                If TLInfo Is Nothing Then           ' load TLB if needed
                    Set TLInfo = TLI.TypeLibInfoFromRegistry(m_Rs.Fields(recGUID).Value, _
                        CInt(chrHex & m_Rs.Fields(recMajor).Value), CInt(chrHex & m_Rs.Fields(recMinor).Value), m_Rs.Fields(recLang).Value)
                End If
                If Err.Number = 0 Then
                    If Action = tlbLoadEvents Then n = lOffsetEvt Else n = lOffsetMthd
                    With TLInfo.GetTypeInfo(n).Members  ' for each, load/log
                        For n = 1 To .Count
                            sName = LCase$(.Item(n).Name)
                            If LenB(sName) <> 0 Then
                                m_Events.AddNew Array(0, 1, 2), Array(p, Action, sName)
                                m_Events.Update
                            End If
                        Next
                    End With
                Else
                    m_Rs.Fields(recFlags).Value = rstNoLoad
                    m_Rs.Update: Err.Clear
                End If
                On Error GoTo 0
            End If
            m_Events.Filter = adFilterNone
        End If
        GetFQN = p
    End If

End Function

Public Function IsEvent(sName As String, lRefID As Long) As Boolean

    ' returns if event is found in its related TLB
    
    If lRefID <> 0 Then
        m_Events.Filter = modMain.SetQuery(recRefID, qryIs, lRefID, qryAnd, _
                                            recEvnts, qryIs, LCase$(sName))
        IsEvent = (m_Events.EOF = False)
    End If

End Function

Private Sub pvCreateRecordsets()

    ' temp recordsets used to track which library methods/events
    ' have been loaded vs. reloading each one, each time for the
    ' various controls and code pages

    Set m_Rs = New ADODB.Recordset
    With m_Rs
        .Fields.Append recID, adInteger
        .Fields.Append recFQN, adVarChar, 1023  ' fully qualified name
        .Fields.Append recGUID, adVarChar, 38   ' GUID
        .Fields.Append recMajor, adVarChar, 4   ' version
        .Fields.Append recMinor, adVarChar, 4   ' version
        .Fields.Append recLang, adInteger       ' LCID, language
        .Fields.Append recCRC, adInteger        ' case-insensitive lookup (FQN)
        .Fields.Append recCRC2, adInteger       ' case-insensitive lookup (name)
        .Fields.Append recFlags, adInteger      ' RefStateEnum value
        .Fields.Append recEvnts, adInteger      ' TLB handle for events collection
        .Fields.Append recMthd, adInteger       ' TLB handle for methods collection
        .Open
    End With
    Set m_Events = New ADODB.Recordset
    With m_Events
        .Fields.Append recRefID, adInteger       ' cross-ref to m_Rs.RecID
        .Fields.Append recType, adInteger        ' event or method
        .Fields.Append recEvnts, adVarWChar, 1023 ' event/method name
        .Open
    End With

End Sub

Private Function pvSearchRegistry(InitKey As String, TargetKey As String, TargetValue As String) As String

    ' used to locate TLB in registry as needed, read-only access requested
    ' InitKey is the root folder to begin searching within the registry, i.e., "TypeLib"
    ' TargetKey is the key to find within the root, i.e., "win32"
    ' TargetValue, not currently used, is to compare the content of the target key
    
    ' Return value is the path where the TargetKey was found, i.e.,
    '   TypeLib\{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}\1.0\0\win32

    Dim hKeyTgt&, lSize&, lIndex&
    Dim hKey&, sValue$, sBuffer$
    Dim bOk As Boolean

    RegOpenKeyExW HKEY_CLASSES_ROOT, StrPtr(InitKey), 0, ACCESS_READ, hKey
    If hKey <> 0 Then       ' collect various versions
        sBuffer = Space$(256): lIndex = 0
        Do                  ' enumerate versions of TLB
            lSize = 256
            If RegEnumKeyExW(hKey, lIndex, StrPtr(sBuffer), lSize, 0&, 0&, 0&, 0&) = _
                            ERROR_NO_MORE_ITEMS Or lSize = 0 Then Exit Do
            sValue = Left$(sBuffer, lSize)
            If StrComp(sValue, TargetKey, vbTextCompare) = 0 Then
                bOk = True
                sValue = InitKey & chrSlash & sValue
            Else
                sValue = InitKey & chrSlash & sValue
                bOk = pvFindSubKey(sValue, sBuffer, TargetKey)
            End If
            If bOk Then
                If LenB(TargetValue) = 0 Then
                    pvSearchRegistry = sValue
                    Exit Do
                Else
                    RegOpenKeyExW HKEY_CLASSES_ROOT, StrPtr(sValue), 0, ACCESS_READ, hKeyTgt
                    If hKeyTgt <> 0 Then
                        lSize = 256
                        If RegQueryValueExW(hKeyTgt, 0&, 0&, REG_SZ, StrPtr(sBuffer), lSize) = 0& Then
                            RegCloseKey hKeyTgt
                            If InStr(1, Left$(sBuffer, lSize - 1), chrSlash & TargetValue, vbTextCompare) <> 0 Then
                                pvSearchRegistry = sValue
                                Exit Do
                            End If
                        Else
                            RegCloseKey hKeyTgt
                        End If
                    End If
                End If
            End If
            lIndex = lIndex + 1
        Loop
        RegCloseKey hKey
    End If

End Function

Private Function pvFindSubKey(sPath As String, sBuffer As String, sKey As String) As Boolean

    ' recursive search into registry child keys, called by SearchRegistry
    
    Dim hKey&, lIndex&, lSize&, sValue$
    
    RegOpenKeyExW HKEY_CLASSES_ROOT, StrPtr(sPath), 0, ACCESS_READ, hKey
    If hKey <> 0 Then
        Do                  ' enumerate versions of TLB
            lSize = Len(sBuffer)
            If RegEnumKeyExW(hKey, lIndex, StrPtr(sBuffer), lSize, 0&, 0&, 0&, 0&) = _
                            ERROR_NO_MORE_ITEMS Or lSize = 0 Then Exit Do
            sValue = Left$(sBuffer, lSize)
            If sValue = sKey Then
                sPath = sPath & chrSlash & sKey
                pvFindSubKey = True
                Exit Do
            Else
                sValue = sPath & chrSlash & sValue
                If pvFindSubKey(sValue, sBuffer, sKey) = True Then
                    sPath = sValue
                    pvFindSubKey = True
                    Exit Do
                End If
            End If
            lIndex = lIndex + 1
        Loop
        RegCloseKey hKey
    End If

End Function

Private Sub Class_Terminate()
    If Not m_Rs Is Nothing Then
        m_Rs.Close: Set m_Rs = Nothing
        m_Events.Close: Set m_Events = Nothing
    End If
    Set TLInfo = Nothing
End Sub
