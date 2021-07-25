VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type WordListStructEx   ' used during validation checks
    Count As Long
    Item() As Long              ' see RecordSlotEnum
    pReserved As Long
End Type
Private Type LocalsNameStruct   ' used for private enum members/local declarations
    Count As Long
    fName() As String
End Type
Private Type ZombieListStruct
    Count As Long               ' size of List() array
    Names As LocalsNameStruct   ' used when adding local scope items to List(0)
    List() As WordListStructEx  ' up to 12 individual lists, see ZombieListEnum
End Type
' notes: reason why we split these lists into multiple separate categories
'   In VB, it is possible to have the same name within different categories
'   For example, "Public Enum Test" and "Public Test As Long" is legitimate
'       where the enum & variable have the same name in the same module.
'   When this occurs, there is a heirarchy if the item were referenced
'   solely by its name. Also, not all categories can duplicate other categories

' Allowed duplication is shown by this table
' M=method/API, E=enum, Em=enum mbr, T=UDT, C=const, V=variable, F=code file
'           F   M   E   Em  C   V   T
'   F           =   =   =   =   =   =       << referenced directly or via: As/New
'   M       =       =               =       << referenced directly
'   E       =   =       =   =   =           << only referenced via: As/New
'   Em      =   =   =               =       << referenced directly or via parent
'   C       =       =               =       << referenced directly
'   V       =       =               =       << referenced directly
'   T       =   =       =   =   =           << only referenced via: As/New
' (*) order listed is heirarchy when only the name is referenced, i.e. X vs modMain.X
' however, heirarchy is also referenced from lowest scope to highest scope.
' The order below is by scope
Private Enum ZombieListEnum
    zlLocals = 0        ' Dim,ReDim,Const statements & parameters within methods
    zlCpShadow = 1      ' VB shadow methods, i.e., Left,Width,hWnd,etc
    zlCpRaised = 2      ' non-bas Public Events
    zlCpMethods = 3     ' APIs/methods, not class events
    zlCpEnums = 4       ' private enums & enum members
    zlCpVars = 5        ' private module-level variables,constants
    zlCpUDTs = 6        ' private UDTs/Types
    zlGblPages = 7      ' code pages
    zlGblMethods = 8    ' global methods,APIs
    zlGblEnums = 9      ' global enums & enum members
    zlGblVars = 10      ' global variables,constants
    zlGblUDTs = 11      ' global UDTs/Types
End Enum
Private Enum RecordSlotEnum
    rteCRC = 0
    rteRecID = 1
    rteParent = 2
    rteType = 3
    rteDiscreps = 4
End Enum

Const chrsAs = "As"

' lots of lists for various types of validation checks
Dim m_ZombieList As ZombieListStruct
    ' up to 12 lists containing parsed project declarations & objects
Dim m_WordList As WordListStruct
    ' parsed words to compare against the zombie lists & others
Dim m_DecsList As WordListStruct
    ' parsed declarations only that use the keywords: As or New
Dim m_Keywords As WordListStruct
    ' VB reserved keywords
Dim m_VarFuncs As WordListStruct
    ' known VB methods that have both Variant & String versions
Dim m_Malicious As WordListStruct
    ' common APIs and VB commands that can be used for malicious intent
Dim m_ParseChecks As WordListStruct
    ' used to minimize processing of m_VarFuncs items over & over again in same method
Dim m_Literals As ZombieListStruct
    ' string/date literals compared for duplication as needed
Dim m_DupeExclusions(0 To 1) As WordListStructEx
    ' list of code pages excluded in some checks: Dupe declartions and/or literals
Dim m_WithTracker As clsWithTracking    ' used to normalize With Block objects
Dim m_rsLiterals As ADODB.Recordset     ' used for tracking duplicate literals
Dim m_TLB As TLI.TypeLibInfo            ' used to load VB shadow methods/events
Dim m_DefTypes() As Byte                ' used to vartype items via DefType statements
Dim m_ValidationChecks As ValidationTypeEnum
Dim m_crcCpg&, m_crcMe&

Public Sub ValidateProject(ICallback As IEvents, lOptions As ValidationTypeEnum)

    Dim lOffsets&(), c&, p&, n&, hHandle&
    Dim lValue&, lCount&, nrFiles&, fnr&
    Dim sName$, sCodeFile$, sTokens$()
    Dim rs As ADODB.Recordset, vBkMk As Variant
    Dim rsMethods As ADODB.Recordset
    Dim rsGlobals As ADODB.Recordset
    
    ' get a list of all code pages to be validated
    Set rs = gParsedItems.Clone
    rs.Filter = modMain.SetQuery(recType, qryIs, itCodePage, qryAnd, recFlags, qryGT, -1)
    If rs.EOF = True Then GoTo exitRoutine
    
    ICallback.ValidationBegin
    m_ValidationChecks = lOptions
    
    ' create lists for excluded code pages in some validations
    Set rsGlobals = gParsedItems.Clone
    If (lOptions And vtDupeDecs) <> 0 Then
        rsGlobals.Filter = modMain.SetQuery(recType, qryIs, itValidation, qryAnd, recGrp, qryIs, 0)
        If rsGlobals.EOF = False Then
            rsGlobals.Sort = recFlags: n = 1
            ReDim m_DupeExclusions(0).Item(0 To 0, 1 To rsGlobals.RecordCount)
            Do
                m_DupeExclusions(0).Item(0, n) = rsGlobals.Fields(recFlags).Value
                n = n + 1
                rsGlobals.MoveNext
            Loop Until rsGlobals.EOF = True
            m_DupeExclusions(0).Count = rsGlobals.RecordCount
        End If
    End If
    If (lOptions And vtDupeLiterals) <> 0 Then
        rsGlobals.Filter = modMain.SetQuery(recType, qryIs, itValidation, qryAnd, recGrp, qryIs, 1)
        If rsGlobals.EOF = False Then
            rsGlobals.Sort = recFlags: n = 1
            ReDim m_DupeExclusions(1).Item(0 To 0, 1 To rsGlobals.RecordCount)
            Do
                m_DupeExclusions(1).Item(0, n) = rsGlobals.Fields(recFlags).Value
                n = n + 1
                rsGlobals.MoveNext
            Loop Until rsGlobals.EOF = True
            m_DupeExclusions(1).Count = rsGlobals.RecordCount
        End If
    End If
    rsGlobals.Sort = vbNullString
    rsGlobals.Filter = adFilterNone
    
    ' load the VB library if zombie checks are active
    If (m_ValidationChecks And vtZombie) <> 0 Then
        rsGlobals.Find modMain.SetQuery(recName, qryIs, "VB"), , , 1&
        If rsGlobals.EOF = False Then
            On Error Resume Next
            sName = rsGlobals.Fields(recAttr2).Value
            n = InStr(sName, chrHash)
            sTokens() = Split(Mid$(sName, n + 1), chrDot)
            If Err Then
                Err.Clear
            Else
                sName = Left$(sName, n - 1)
                Set m_TLB = TLI.TypeLibInfoFromRegistry(sName, CInt(sTokens(0)), CInt(sTokens(1)), CLng(sTokens(2)))
                If Err Then Err.Clear
            End If
            Erase sTokens()
            On Error GoTo 0
        End If
    End If
    
    ' get list of global declarations
    rsGlobals.Filter = modMain.SetQuery(recScope, qryIs, scpGlobal)
    rsGlobals.Sort = recCodePg & chrComma & recType & chrComma & recID
    rs.Sort = recIdxAttr
    nrFiles = rs.RecordCount
    
    Set m_WithTracker = New clsWithTracking
    ReDim m_ZombieList.List(zlLocals To zlGblUDTs)
    m_ZombieList.Count = zlGblUDTs + 1
    
    m_Keywords.List = Split(LoadResString(0), chrComma)
    m_Keywords.Count = UBound(m_Keywords.List)
    
    If (m_ValidationChecks And vtVarFunc) <> 0 Then
        m_VarFuncs.List = Split(LoadResString(1), chrComma)
        m_VarFuncs.Count = UBound(m_VarFuncs.List)
    End If
    If (m_ValidationChecks And vtMalicious) <> 0 Then
        m_Malicious.List = Split(LoadResString(2), chrComma)
        m_Malicious.Count = UBound(m_Malicious.List)
    End If
    If (m_ValidationChecks And vtDupeLiterals) <> 0 Then pvCreateLiteralsRS
    
    pvAddDeclares_Globals rs, rsGlobals ' add globals to m_ZombieList
    rs.MoveFirst                        ' first code page
    Set rsMethods = gParsedItems.Clone
    rsMethods.Sort = recID
    
    n = 0
    Do
        ' set host's statusbar text & cache code page name
        sCodeFile = rs.Fields(recName).Value
        fnr = fnr + 1: c = InStr(sCodeFile, chrDot)
        sCodeFile = CStr(fnr) & chrSlash & CStr(nrFiles) & " Validating " & Mid$(sCodeFile, c + 1)
        ICallback.Status sCodeFile
        sCodeFile = sCodeFile & chrDot

        ' prep for parsing code page mehtods
        gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recID).Value), , , 1&
        vBkMk = gParsedItems.Bookmark
        ' get the code page's file
        gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
        sName = gParsedItems.Fields(recAttr).Value
        hHandle = modMain.SetSourceData(sName, &H80000000, 0)
        If hHandle = 0 Or hHandle = -1 Then GoTo exitRoutine
        modMain.CloseHandle hHandle
        gSourceFile.CPBookMark = vBkMk: vBkMk = Empty
        
        ' get the cached offsets file for this code page
        sName = gParsedItems.Fields(recAttr2).Value
        If LenB(sName) = 0 Then
            lCount = 0
        Else
            hHandle = modMain.GetFileHandle(sName, False)
            If hHandle = 0 Or hHandle = -1 Then GoTo exitRoutine
                
            modMain.GetFileLastModDate sName, 0&, 0&, lCount
            ReDim lOffsets(0 To lCount \ 4 - 1)
            modMain.ReadFile hHandle, lOffsets(0), lCount, lValue
            modMain.CloseHandle hHandle
            lCount = lCount \ 4
        End If
        ' parse out parts of the declarations section not parsed during initial scan
        c = 0: pvAddDeclares_CodePage rs, c, lOffsets(), lCount
        ' get list of methods, events, and APIs
        rsMethods.Filter = modMain.SetQuery(recCodePg, qryIs, rs.Fields(recID).Value, _
                                            qryAnd, recType, qryLT, itEnum, _
                                            qryAnd, recScope, qryNot, scpLocal)
        If rsMethods.EOF = False Then
            Do
                Select Case rsMethods.Fields(recType).Value
                Case itMethod, itClassEvent
                    ' for each of these, locate cached offsets for method's content
                    ICallback.Status sCodeFile & rsMethods.Fields(recName).Value
                    n = rsMethods.Fields(recStart).Value
                    p = rsMethods.Fields(recOffset2).Value
                    lValue = 0
                    For c = c To lCount - 1 Step 2
                        If lOffsets(c) > n Then     ' start of method content
                            If lValue = 0 Then lValue = c + 1
                        End If                      ' end of method content?
                        If lOffsets(c) > p Then Exit For
                    Next
                    If lValue = 0 Then              ' method with no code
                        Call pvDoEmptyCode(rsMethods)
                    Else
                        ' parse method params then the method itself
                        pvParseName_Params rsMethods
                        pvParseWords rsMethods, lOffsets(), lValue - 1, c - lValue + 1
                        If m_WordList.Count <> 0 Then
                            If (m_ValidationChecks And vtZombie) = 0 Then
                                m_WordList.Count = 0: Erase m_WordList.List()
                            Else
                                m_WordList.pReserved = rsMethods.Fields(recID).Value
                                pvDoZombieCheck m_WordList, False
                                m_WordList.pReserved = 0
                            End If
                            pvResolveZombies rsMethods, zlLocals, zlLocals
                        End If
                        Erase m_ParseChecks.List(): m_ParseChecks.Count = 0
                    End If
                Case Else
                End Select
                rsMethods.MoveNext
            Loop Until rsMethods.EOF = True         ' get next method
            If m_DecsList.Count <> 0 Then           ' zombie check any items declared with As or New
                If (m_ValidationChecks And vtZombie) <> 0 Then pvDoZombieCheck m_DecsList, True
            End If
        End If
        Erase lOffsets()
        For c = zlCpMethods To zlCpUDTs             ' declarations section?
            If m_ZombieList.List(c).Count <> 0 Then Exit For
        Next
        If c <= zlCpUDTs Then
            ICallback.Status sCodeFile & ".. log findings"
            rsMethods.Filter = modMain.SetQuery(recCodePg, qryIs, rs.Fields(recID).Value)
            pvResolveZombies rsMethods, zlCpRaised, zlCpUDTs
        End If
        If m_Literals.Count <> 0 Then pvUpdateLiteralsRS
        rs.MoveNext
        m_WithTracker.Reset
    Loop Until rs.EOF = True                        ' get next code page
    
    ICallback.Status "Performing post-validation checks"
    If (m_ValidationChecks And vtZombie) <> 0 Then
        rsGlobals.MoveFirst: pvDoPostZombieCheck rsGlobals
    End If
    If (m_ValidationChecks And vtMalicious) <> 0 Then pvListMaliciousDLLs rs
    If (m_ValidationChecks And vtDupeDecs) <> 0 Then
        ICallback.Status "Checking for duplicated declarations"
        pvResolveDupeDecs rs
    End If
    If m_Literals.Count <> 0 Then
        ICallback.Status "Checking for duplicated literals"
        pvResolveDupeLiterals rs
    End If
    
exitRoutine:
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not rsMethods Is Nothing Then rsMethods.Close: Set rsMethods = Nothing
    If Not rsGlobals Is Nothing Then rsGlobals.Close: Set rsGlobals = Nothing
    Erase m_ZombieList.List()
    Erase m_ZombieList.Names.fName()
    Erase m_DupeExclusions
    Set m_WithTracker = Nothing
    Set m_TLB = Nothing
    If hHandle <> 0 And hHandle <> -1 Then
        ICallback.Status "Validation complete"
        ICallback.ValidationComplete Me, m_ValidationChecks
    End If
    
End Sub

Private Sub pvParseWords(rs As ADODB.Recordset, _
                    arrOffsets() As Long, arrStart As Long, arrCount As Long)

    ' Primarily a word-processor based on pre-parsed statements
    ' Routine parses all code contained within a specific method.
    ' Declarations were already processed at this point
    
    ' As words are parsed out, they are collected to be used for
    '   the various validation checks. Two words are always
    '   cached, so that the previous word is available if needed.
    ' Code within parentheses initiate recursion into this routine.
    '   That code is immediately processed and when the parentheses
    '   are closed, recursion exits and pre-recursion parsing continues.
    
    Dim lStartWd&, lMaxWd&, lStmt&, lStartStmt&
    Dim c%, lMaxStmt&, wdIdx&, IdxKeyWord&
    Dim sWord$(0 To 1), arrParents&(0 To 1)
    Dim lFlagsStmt&, pFlags&
    ' 0x00001   With statement begins
    ' 0x00002   With statement name parsed
    ' 0x00004   As or New keyword parsed
    
    pFlags = wbpMarkBrackets
    For lStmt = arrStart To arrStart + arrCount - 1 Step 2
        lStartStmt = arrOffsets(lStmt)
        lMaxStmt = arrOffsets(lStmt + 1)
        If lMaxStmt < 0 Then                    ' continuation exists?
            lMaxStmt = -lMaxStmt                ' join the lines for easier parsing
            pvJoinContinuation lStartStmt, lMaxStmt
        End If
        lMaxWd = lStartStmt: lStartWd = lMaxWd: lFlagsStmt = 0 ' initialize for loop
        Do
            modMain.ParseNextWord lMaxWd, lMaxStmt, lStartWd, lMaxWd, pFlags
            If lStartWd = lMaxWd Then Exit Do   ' sanity check, shouldn't happen
            c = gSourceFile.Data(lStartWd)      ' 1st char of parsed word
            Select Case c
            Case vbKey0 To vbKey9, 38: lStartWd = 0 ' numeric,&   <- don't care
            Case vbKeyQuote, vbKeyHash          ' string/date literals
                If m_Literals.Count <> 0 Then
                    pvAddLiteralToList rs, Mid$(gSourceFile.Text, lStartWd, lMaxWd - lStartWd)
                End If
            Case vbKeyParenthesis
                If lMaxWd - lStartWd > 2 Then   ' else parsed: ()
                    ' recursion when parentheses are encountered
                    arrParents(0) = lStartWd + 1: arrParents(1) = lMaxWd - 1
                    pvParseWords rs, arrParents(), 0, 2
                End If
                wdIdx = wdIdx Xor 1
            Case Else
                If modMain.IsVarTyped(gSourceFile.Data(lMaxWd - 1)) = 0 Then
                    sWord(wdIdx) = Mid$(gSourceFile.Text, lStartWd, lMaxWd - lStartWd)
                Else
                    sWord(wdIdx) = Mid$(gSourceFile.Text, lStartWd, lMaxWd - lStartWd - 1)
                End If
                If (pFlags And wbpBracketsMarked) <> 0 Then
                    sWord(wdIdx) = Replace(sWord(wdIdx), vbLf, vbNullString)
                    pFlags = pFlags Xor wbpBracketsMarked
                    If AscW(sWord(wdIdx)) = 38 Then lStartWd = 0: IdxKeyWord = -1
                End If
                
                If lFlagsStmt <> 0 Then
                    If lFlagsStmt = 1 Then      ' previous word: With
                        lFlagsStmt = 2          ' triggered at end of statement
                        If c = vbKeyDot Then
                            sWord(wdIdx) = m_WithTracker.WithName & sWord(wdIdx)
                        End If
                        m_WithTracker.BeginWithStatement sWord(wdIdx)
                        IdxKeyWord = 0
                    ElseIf (lFlagsStmt And 4) = 4 Then ' previous word: As or New
                        IdxKeyWord = modMain.FindWord(m_Keywords, sWord(wdIdx), vbTextCompare, False)
                        If IdxKeyWord = 0 Then
                            modMain.FindWord m_DecsList, sWord(wdIdx), vbBinaryCompare, True
                            IdxKeyWord = -1
                        End If
                        lFlagsStmt = lFlagsStmt Xor 4
                    Else
                        IdxKeyWord = 0
                    End If
                ElseIf c = vbKeyDot Then
                    If gSourceFile.Data(lStartWd - 1) = vbKeyParenthesis2 Then
                        sWord(wdIdx) = sWord(wdIdx Xor 1) & sWord(wdIdx)
                    Else
                        sWord(wdIdx) = m_WithTracker.WithName & sWord(wdIdx)
                    End If
                    IdxKeyWord = 0
                Else
                    IdxKeyWord = modMain.FindWord(m_Keywords, sWord(wdIdx), vbTextCompare, False)
                End If
                If IdxKeyWord = 0 Then
                    pvProcessParsedWord rs, sWord(wdIdx), sWord(wdIdx Xor 1), lStartStmt, lStartWd, lMaxWd
                ElseIf IdxKeyWord > 0 Then
                    Select Case c
                    Case vbKeyA             ' As, AddressOf
                        If m_Keywords.List(IdxKeyWord) = chrsAs Then
                            lFlagsStmt = lFlagsStmt Or 4
                        ElseIf m_Keywords.List(IdxKeyWord) = "AddressOf" Then
                            pvDoAddressOf lMaxWd, lMaxStmt
                        End If
                    Case vbKeyW             ' With
                        If m_Keywords.List(IdxKeyWord) = "With" Then
                            If lMaxWd = lMaxStmt Then   ' End With statement
                                m_WithTracker.EndWithBlock
                            Else                        ' With statement
                                lFlagsStmt = 1  ' flag triggered on next word
                            End If
                        End If
                    Case vbKeyR             ' ReDim, RaiseEvent
                        If m_Keywords.List(IdxKeyWord) = "ReDim" Then
                            pvParseDeclares rs, lStartWd, lMaxWd, lMaxStmt
                        ElseIf m_Keywords.List(IdxKeyWord) = "RaiseEvent" Then
                            pvDoRaiseEvent lMaxWd, lMaxStmt
                        End If
                    Case vbKeyD             ' Dim, Date
                        If m_Keywords.List(IdxKeyWord) = "Dim" Then
                            pvParseDeclares rs, lStartWd, lMaxWd, lMaxStmt
                        ElseIf m_Keywords.List(IdxKeyWord) = "Date" Then
                            If lStartStmt <> lStartWd Then
                                If sWord(wdIdx Xor 1) <> chrsAs Then
                                    pvProcessParsedWord rs, sWord(wdIdx), sWord(wdIdx Xor 1), lStartStmt, lStartWd, lMaxWd
                                End If
                            End If
                        End If
                    Case vbKeyS             ' Static, Stop, String
                        If m_Keywords.List(IdxKeyWord) = "String" Then
                            If sWord(wdIdx Xor 1) <> chrsAs Then
                                pvProcessParsedWord rs, sWord(wdIdx), sWord(wdIdx Xor 1), lStartStmt, lStartWd, lMaxWd
                            End If
                        ElseIf m_Keywords.List(IdxKeyWord) = ParseTypeStat Then
                            pvParseDeclares rs, lStartWd, lMaxWd, lMaxStmt
                        ElseIf m_Keywords.List(IdxKeyWord) = "Stop" Then
                            pvDoStopEndUsage rs
                        End If
                    Case vbKeyC             ' Const
                        If m_Keywords.List(IdxKeyWord) = ParseTypeCnt Then
                            pvParseDeclares rs, lStartWd, lMaxWd, lMaxStmt
                        End If
                    Case vbKeyE             ' End
                        If m_Keywords.List(IdxKeyWord) = "End" Then
                            If lStartWd = lStartStmt And lMaxWd = lMaxStmt Then
                                pvDoStopEndUsage rs
                            End If
                        End If
                    Case vbKeyG             ' GoSub, GoTo can ignore remainder of statement
                        If m_Keywords.List(IdxKeyWord) = "GoTo" Then lMaxWd = lMaxStmt
                        If m_Keywords.List(IdxKeyWord) = "GoSub" Then lMaxWd = lMaxStmt
                    Case vbKeyN             ' New
                        If m_Keywords.List(IdxKeyWord) = "New" Then
                            lFlagsStmt = lFlagsStmt Or 4
                        End If
                    End Select
                End If
            End Select
            wdIdx = wdIdx Xor 1                 ' toggle active index
        Loop While lMaxWd < lMaxStmt
        If lFlagsStmt <> 0 Then
            If (lFlagsStmt And 2) = 2 Then      ' clean up "End With" statement
                m_WithTracker.EndWithStatement
            End If
            lFlagsStmt = 0
        End If
    Next                                        ' parse next statement

End Sub

Private Sub pvAddDeclares_CodePage(rsCpg As ADODB.Recordset, lLastStmt As Long, _
                                lStatements() As Long, nrStatements As Long)
                                
    ' Routine creates a sorted array with code-page related items.
    ' Public enums are excluded herein because they were handled earlier.
    ' Private enum members are were not parsed during initial scan.
    ' We have to parse them here.

    Dim sValue$, lStart&, p&, n&, lCpg&
    Dim lDecsOffset&(0 To 1)
    Dim lRecord&(rteCRC To rteDiscreps)
    Dim rs As ADODB.Recordset
    
    For n = zlCpRaised To zlCpUDTs                 ' reset
        If m_ZombieList.List(n).Count <> 0 Then
            Erase m_ZombieList.List(n).Item()
            m_ZombieList.List(n).Count = 0
        End If
    Next
    If m_ZombieList.Names.Count <> 0 Then           ' reset
        m_ZombieList.Names.Count = 0
        Erase m_ZombieList.Names.fName()
    End If
    
    If (rsCpg.Fields(recFlags).Value And iaPredeclared) = 0 Then
        m_crcCpg = 0
    Else    ' some "private" items can be referenced with code page, i.e., Form1.Text1
        sValue = rsCpg.Fields(recName).Value & chrDot
        sValue = Mid$(sValue, InStr(sValue, chrDot) + 1)
        m_crcCpg = modMain.CRCItem(sValue, True)
    End If
    
    lCpg = rsCpg.Fields(recID)
    If m_Literals.Count <> 0 Then
        If m_Literals.Names.Count <> 0 Then         ' reset
            m_Literals.Names.Count = 0
            Erase m_Literals.Names.fName()
        End If
        m_Literals.List(0).Count = 0
        m_Literals.List(0).pReserved = lCpg
    End If
    
    ' get a list of items that would have offsets cached to file and those
    ' items will be private enum members and all UDT members
    Set rs = gParsedItems.Clone
    rs.Filter = chrParentO & modMain.SetQuery(recParent, qryIs, lCpg, _
                                qryAnd, recType, qryIs, itEnum, _
                                qryAnd, recScope, qryIs, scpPrivate) & ") Or (" & _
                    modMain.SetQuery(recParent, qryIs, lCpg, _
                                qryAnd, recType, qryIs, itType) & chrParentC
    If rs.EOF = False Then
        rs.Sort = recID                             ' sort first seen
        Do                                          ' get enum/type block start/end
            n = rs.Fields(recStart).Value: p = rs.Fields(recOffset2).Value
            For lLastStmt = lLastStmt To nrStatements - 1 Step 2
                If lStatements(lLastStmt) > n Then  ' find start/end of block
                    If lStart = 0 Then lStart = lLastStmt + 1
                    If lStatements(lLastStmt) > p Then Exit For
                End If
            Next                                    ' process enum/UDT members
            pvParseName_Members rs, lStatements(), lStart - 1, lLastStmt - lStart + 1
            rs.MoveNext: lStart = 0
        Loop Until rs.EOF = True
    End If
    ' get list of all declarations that are not global (which were already handled)
    rs.Filter = modMain.SetQuery(recCodePg, qryIs, lCpg)
    If rs.EOF = True Then GoTo exitRoutine
    
    ' get list of any Def[xxx] statements for locally declared variables/params
    modMain.GetDefTypes rsCpg.Fields(recID).Value, m_DefTypes()
    rs.Sort = recType
    Do
        n = rs.Fields(recType).Value
        Select Case n                               ' set collection slot for type
            Case itMethod, itClassEvent: p = zlCpMethods
                ' these are processed after routine exits
            Case itAPI: p = zlCpMethods
                pvParseName_Params rs
            Case itEvent: p = zlCpRaised
                pvParseName_Params rs
            Case itEnum: p = zlCpEnums
            Case itVariable, itConstant, itEnumMember
                lDecsOffset(0) = rs.Fields(recOffset).Value
                If n = itEnumMember Then
                    lDecsOffset(1) = rs.Fields(recEnd).Value
                    pvParseWords rs, lDecsOffset(), 0, 2
                    p = 0: rs.MoveNext  ' already added when pvParseName_Members
                Else
                    lDecsOffset(1) = rs.Fields(recOffset2).Value
                    pvParseWords rs, lDecsOffset(), 0, 2
                    p = zlCpVars
                End If
            Case itControl: p = zlCpVars
            Case itType: p = zlCpUDTs
            Case itImplements
                modMain.FindWord m_WordList, rs.Fields(recAttr).Value, vbBinaryCompare, True
                rs.MoveNext: p = 0
            Case Else: p = 0                        ' n/a, skip past these
                rs.Find modMain.SetQuery(recType, qryNot, n)
        End Select
        If p <> 0 Then
            If rs.Fields(recScope).Value > scpLocal Then
                lRecord(rteDiscreps) = 0
                sValue = rs.Fields(recDiscrep).Value
                For lStart = 1 To Len(sValue)
                    Select Case Mid$(sValue, lStart, 1)
                    Case chrV: lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtVarType
                    Case chrZ: lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtZombie
                    Case chrW: lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtWithNoEvents
                    End Select
                Next
                sValue = rs.Fields(recName).Value   ' item name
                lRecord(rteCRC) = modMain.CRCItem(sValue, True)
                lRecord(rteRecID) = rs.Fields(recID).Value
                lRecord(rteParent) = rs.Fields(recParent).Value
                lRecord(rteType) = n
                                                    ' add as: Item
                lStart = pvFindWordEx(m_ZombieList.List(p), lRecord(), True)
                If lStart <> 0 Then GoSub preferPropGet_
                If p = zlCpMethods Or p = zlCpVars Then ' may be a public method/variable
                    If rs.Fields(recScope).Value = scpPublic Then
                                                    ' add as: Me.Item
                        lRecord(rteCRC) = modMain.CRCItem(sValue, True, m_crcMe)
                        lStart = pvFindWordEx(m_ZombieList.List(p), lRecord(), True)
                        If lStart <> 0 Then GoSub preferPropGet_
                    End If
                    If m_crcCpg <> 0 Then           ' predeclared class/module?
                        If rs.Fields(recScope).Value > scpPrivate Then
                                                    ' add as: CodePage.Item
                            lRecord(rteCRC) = modMain.CRCItem(sValue, True, m_crcCpg)
                            lStart = pvFindWordEx(m_ZombieList.List(p), lRecord(), True)
                            If lStart <> 0 Then GoSub preferPropGet_
                        End If
                    End If
                End If
            End If
            rs.MoveNext
        End If
    Loop Until rs.EOF = True
    
exitRoutine:
    If lLastStmt <> 0 Then
        ' parse enum/UDT values. These may zombie check other items
        ' for example EnumMbrC = EnumMbrA Or EnumMbrB
        pvParseWords rs, lStatements(), 0, lLastStmt
    End If
    If (m_WordList.Count Or m_DecsList.Count) <> 0 Then
        If (m_ValidationChecks And vtZombie) = 0 Then
            m_WordList.Count = 0: Erase m_WordList.List()
        Else
            If Not m_TLB Is Nothing Then pvLoadShadowMethods rsCpg
            pvDoZombieCheck m_WordList, False
        End If
    End If
    Erase m_ParseChecks.List(): m_ParseChecks.Count = 0
    If m_Literals.Count <> 0 Then pvUpdateLiteralsRS
    rs.Close: Set rs = Nothing
    Exit Sub
    
preferPropGet_:
    If n = itMethod Then ' when Prop Let,Set parsed, use Get for zombie checks
        If (rs.Fields(recFlags).Value And iaPropGet) = iaPropGet Then
            m_ZombieList.List(p).Item(rteRecID, lStart) = rs.Fields(recID).Value
        End If
    End If
    Return
End Sub

Private Sub pvAddDeclares_Globals(rsCpg As ADODB.Recordset, rsGlobals As ADODB.Recordset)

    ' routine creates a sorted array with the various ways any
    ' project-global item can be referenced. Deepest example,
    ' an enum member of a public enum within a bas-module:
    '   scpLocal                                    Member
    '   ItemScopeEnum.scpLocal                      Enum.Member
    '   modMain.scpLocal                            Module.Member
    '   modMain.ItemScopeEnum.scpLocal              Module.Enum.Member
    '   prjScanner.scpLocal                         Project.Member
    '   prjScanner.ItemScopeEnum.scpLocal           Project.Enum.Member
    '   prjScanner.modMain.ItemScopeEnum.scpLocal   Project.Module.Enum.Member
    '   prjScanner.modMain.scpLocal                 Project.Module.Member
    
    ' Global items include
    '   Public scope declarations within a bas module
    '   Any publicly declared enumerations/members
    '   Any code page file
    
    Dim rsEnums As ADODB.Recordset
    Dim lValue&, lType&, lCpg&, lID&, n&, crcPrj&
    Dim crcPrjCpg&, crcEnm&, eParent&, sName$, sEnum$
    Dim lRecord&(rteCRC To rteDiscreps)
    
    For lID = zlGblPages To zlGblUDTs         ' sanity checks
        If m_ZombieList.List(lID).Count <> 0 Then
            Erase m_ZombieList.List(lID).Item()
            m_ZombieList.List(lID).Count = 0
        End If
    Next
    
    If rsGlobals.EOF = True Then GoTo exitRoutine
    
    m_crcMe = modMain.CRCItem("Me.", True)      ' CRC value of "Me." & project name
    crcPrj = modMain.CRCItem(gParsedItems.Fields(recName).Value & chrDot, True)
    
    Set rsEnums = gParsedItems.Clone            ' get list of global Enum members
    rsEnums.Filter = modMain.SetQuery(recType, qryIs, itEnum, _
                                    qryAnd, recScope, qryIs, scpGlobal)
    rsEnums.Sort = recID
    
    Do Until rsGlobals.EOF = True                   ' for each global item, add to array
        lValue = rsGlobals.Fields(recCodePg).Value  ' changed code page?
        If lValue <> lCpg Then
            lCpg = lValue: lType = 0                ' locate code page
            rsCpg.Find modMain.SetQuery(recID, qryIs, lCpg), , , 1&
            sName = rsCpg.Fields(recName).Value
            lValue = InStr(sName, chrDot) + 1      ' CRC value of code page
            crcPrjCpg = modMain.CRCItem(sName & chrDot, True)
            m_crcCpg = modMain.CRCItem(Mid$(sName, lValue) & chrDot, True)
        End If
        lValue = rsGlobals.Fields(recType).Value
        If lValue <> lType Then                 ' changed item type?
            lType = lValue
            Select Case lType                   ' set collection slot for this type
                Case itMethod, itAPI: lID = zlGblMethods
                Case itEvent, itClassEvent: lID = zlGblMethods
                Case itEnum, itEnumMember: lID = zlGblEnums
                Case itVariable, itConstant: lID = zlGblVars
                Case itType: lID = zlGblUDTs
                Case Else: lID = 0              ' n/a, skip past these
                    rsGlobals.Find modMain.SetQuery(recType, qryNot, lType), 1&
            End Select
        End If
        If lID <> 0 Then
            lRecord(rteDiscreps) = 0
            sName = rsGlobals.Fields(recDiscrep).Value
            For n = 1 To Len(sName)
                Select Case Mid$(sName, n, 1)
                Case chrV: lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtVarType
                Case chrZ: lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtZombie
                Case chrW: lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtWithNoEvents
                End Select
            Next
            lRecord(rteRecID) = rsGlobals.Fields(recID).Value
            lRecord(rteParent) = rsGlobals.Fields(recParent).Value
            lRecord(rteType) = lType
            sName = rsGlobals.Fields(recName).Value    ' item name, create record for array
            lRecord(rteCRC) = modMain.CRCItem(sName, True)
                                                ' add item as: Item
            lValue = pvFindWordEx(m_ZombieList.List(lID), lRecord(), True)
            If lValue <> 0 Then GoSub preferPropGet_
                                                ' add item as: CodePage.Item
            lRecord(rteCRC) = modMain.CRCItem(sName, True, m_crcCpg)
            lValue = pvFindWordEx(m_ZombieList.List(lID), lRecord(), True)
            If lValue <> 0 Then GoSub preferPropGet_
                                                ' add item as: Project.CodePage.Item
            lRecord(rteCRC) = modMain.CRCItem(sName, True, crcPrjCpg)
            lValue = pvFindWordEx(m_ZombieList.List(lID), lRecord(), True)
            If lValue <> 0 Then GoSub preferPropGet_
                                                ' add item as: Project.Item
            lRecord(rteCRC) = modMain.CRCItem(sName, True, crcPrj)
            lValue = pvFindWordEx(m_ZombieList.List(lID), lRecord(), True)
            If lValue <> 0 Then GoSub preferPropGet_
            
            If lType = itEnumMember Then        ' handle enum members
                lValue = rsGlobals.Fields(recParent).Value
                If eParent <> lValue Then
                    eParent = lValue
                    rsEnums.Find modMain.SetQuery(recID, qryIs, eParent)
                    sEnum = rsEnums.Fields(recName).Value & chrDot
                    crcEnm = modMain.CRCItem(sEnum, True)
                End If
                sName = sEnum & sName
                lRecord(rteCRC) = modMain.CRCItem(sName, True)
                pvFindWordEx m_ZombieList.List(lID), lRecord(), True
                                            ' add item as: Enum.Item
                lRecord(rteCRC) = modMain.CRCItem(sName, True, crcEnm)
                pvFindWordEx m_ZombieList.List(lID), lRecord(), True
                                            ' add item as: CodePage.Enum.Item
                lRecord(rteCRC) = modMain.CRCItem(sName, True, m_crcCpg)
                pvFindWordEx m_ZombieList.List(lID), lRecord(), True
                                            ' add item as: Project.CodePage.Enum.Item
                lRecord(rteCRC) = modMain.CRCItem(sName, True, crcPrjCpg)
                pvFindWordEx m_ZombieList.List(lID), lRecord(), True
                                            ' add item as: Project.Enum.Item
                lRecord(rteCRC) = modMain.CRCItem(sName, True, crcPrj)
                pvFindWordEx m_ZombieList.List(lID), lRecord(), True
            End If
            rsGlobals.MoveNext                         ' next item
        End If
    Loop
    rsEnums.Close: Set rsEnums = Nothing
    
    rsCpg.MoveFirst                             ' add code pages to the list
    Do                                          ' add code page as: Project.CodePage
        sName = rsCpg.Fields(recDiscrep).Value
        lRecord(rteDiscreps) = 0
        If InStr(sName, chrV) <> 0 Then lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtVarType
        If InStr(sName, chrZ) <> 0 Then lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtZombie
        If InStr(sName, chrW) <> 0 Then lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtWithNoEvents
        lRecord(rteCRC) = modMain.CRCItem(sName, True)
        lRecord(rteRecID) = rsCpg.Fields(recID).Value
        lRecord(rteType) = itCodePage
        lRecord(rteParent) = 0
        sName = rsCpg.Fields(recName).Value
        lValue = InStr(sName, chrDot) + 1
        pvFindWordEx m_ZombieList.List(zlGblPages), lRecord(), True
                                                ' add code page as: CodePage
        lRecord(rteCRC) = modMain.CRCItem(Mid$(sName, lValue), True)
        pvFindWordEx m_ZombieList.List(zlGblPages), lRecord(), True
        rsCpg.MoveNext
    Loop Until rsCpg.EOF = True
    
exitRoutine:
    Exit Sub
    
preferPropGet_:
    If lType = itMethod Then
        If (rsGlobals.Fields(recFlags).Value And iaPropGet) = iaPropGet Then
            m_ZombieList.List(lID).Item(rteRecID, lValue) = rsGlobals.Fields(recID).Value
        End If
    End If
    Return
End Sub

Private Function pvIsDefTyped(c As Integer) As Boolean
    ' checks 1st character of a word against Def[xxx] statements
    Select Case c
    Case vbKeyA To vbKeyZ
        pvIsDefTyped = CBool(m_DefTypes(c - 65) = 1)
    Case 97 To 122
        pvIsDefTyped = CBool(m_DefTypes(c - 97) = 1)
    End Select
    
End Function

Private Sub pvParseName_Members(rs As ADODB.Recordset, lStmts() As Long, _
                                lStart As Long, lCount As Long)

    ' all UDT members and private Type members come to this routine
    ' Enum members are logged to recordset, their vartype processed for zombie checks
    ' UDT members are not logged/cached at all, their vartype processed for zombie checks

    Dim pFlags&, lOffset&, n&, p&, lCrc&, lID&
    Dim lRecord&(rteCRC To rteDiscreps), sName$

    pFlags = wbpMarkBrackets
    For n = lStart To lStart + lCount - 1 Step 2
        p = lStmts(n)                       ' start position in text, which is updated
        modMain.ParseNextWord p, Abs(lStmts(n + 1)), p, lStmts(n), pFlags
        
        ' private enum members cached as EnumMember & Enum.EnumMember
        If rs.Fields(recType).Value = itEnum Then
            If lID = 0 Then
                lID = rs.Fields(recID).Value
                sName = rs.Fields(recName).Value & chrDot  ' CRC of enum name
                lCrc = modMain.CRCItem(sName, True)
            End If
            lOffset = modMain.IsVarTyped(gSourceFile.Data(p - 1))
            If (pFlags And wbpBracketsMarked) = 0 Then
                sName = Mid$(gSourceFile.Text, p, lStmts(n) - p - lOffset)
            Else
                sName = Replace(Mid$(gSourceFile.Text, p, lStmts(n) - p - lOffset), vbLf, vbNullString)
            End If
            lRecord(rteCRC) = modMain.CRCItem(sName, True)
            lRecord(rteRecID) = modMain.CreateRecord(lID, sName, itEnumMember, 0, 0, , , , , , , scpPrivate, chrZ)
            lRecord(rteType) = itEnumMember
            lRecord(rteDiscreps) = vtZombie
            lRecord(rteParent) = lID
            pvFindWordEx m_ZombieList.List(zlCpEnums), lRecord(), True
            lRecord(rteCRC) = modMain.CRCItem(sName, True, lCrc)
            pvFindWordEx m_ZombieList.List(zlCpEnums), lRecord(), True
        End If
    Next
    
End Sub

Private Sub pvParseName_Params(rs As ADODB.Recordset)

    ' APIs, Events, Methods are passed here to vartype check params & add to locals array
    
    Dim n&, p&, lMax&, c%, lType As ItemTypeEnum
    Dim sName$, lMode&, lOffset&, lFlags&
    Dim lRecord&(rteCRC To rteDiscreps), pFlags&
    Const ParseParamArray = "ParamArray"
    Const ParseOptional = "Optional"
    Const ParseByVal = "ByVal"
    Const ParseByRef = "ByRef"
    
    lType = rs.Fields(recType).Value Or (rs.Fields(recFlags) And iaImplemented)
    p = rs.Fields(recOffset).Value          ' locate start of params
    lMax = rs.Fields(recEnd).Value
    If (rs.Fields(recFlags).Value And iaSplitIdent) <> 0 Then
        lMax = -lMax
        pvJoinContinuation p, lMax
    End If
    If gSourceFile.Data(p) <> vbKeyParenthesis Then
        modMain.ParseNextWordEx p, lMax, n, p, wbpFirstChar
    End If
    If gSourceFile.Data(p + 1) = vbKeyParenthesis2 Then
        If p + 2 < lMax Then Exit Sub
    End If                                  ' add to statements for zombie checks
    lRecord(0) = p: lRecord(1) = lMax
    pvParseWords rs, lRecord(), 0, 2
    If gSourceFile.Data(p + 1) = vbKeyParenthesis2 Then Exit Sub
    
    p = p + 1                               ' skip past the open parenthesis
    lRecord(rteType) = itParameter
    lRecord(rteParent) = rs.Fields(recID).Value
    pFlags = wbpParentsCntr Or wbpNotCommaEOW Or wbpSepNamedParams Or wbpMarkBrackets
    Do                                          ' get next token from params
        modMain.ParseNextWordEx p, lMax, n, p, pFlags
        If n = lMax Then Exit Do
        
        c = gSourceFile.Data(n)
        If c <> vbKeyParenthesis Then
            If lMode = 0 Then               ' looking for param name
                lMode = 1                   ' look for comma
                If c = vbKeyP Then          ' skip keyword: ParamArray
                    If p - n = 10 Then
                        If Mid$(gSourceFile.Text, n, 10) = ParseParamArray Then
                            lMode = 0: lFlags = 1
                        End If
                    End If
                ElseIf c = vbKeyB Then      ' skip keywords: ByVal,ByRef
                    If p - n = 5 And gSourceFile.Data(n + 1) = 121 Then
                        If Mid$(gSourceFile.Text, n, 5) = ParseByVal Then
                            lMode = 0
                        ElseIf Mid$(gSourceFile.Text, n, 5) = ParseByRef Then
                            lMode = 0
                        End If
                    End If
                ElseIf c = vbKeyO Then      ' skip keyword: Optional
                    If p - n = 8 Then
                        If Mid$(gSourceFile.Text, n, 8) = ParseOptional Then lMode = 0
                    End If
                End If
                If lMode = 1 Then           ' parse out param name
                    If gSourceFile.Data(p - 1) = vbKeyComma Then
                        lMode = 2: lOffset = 1 ' end of this param
                    ElseIf gSourceFile.Data(p - 1 - lOffset) = vbKeyParenthesis2 Then
                        lMode = 3: lOffset = lOffset + 1
                    End If                  ' end of params
                    If modMain.IsVarTyped(gSourceFile.Data(p - 1 - lOffset)) = 1 Then
                        lFlags = 1: lOffset = lOffset + 1
                    ElseIf lType = itClassEvent Then
                        lFlags = 1          ' event params are not vartype checked
                    End If
                    If (pFlags And wbpBracketsMarked) = 0 Then
                        sName = Mid$(gSourceFile.Text, n, p - n - lOffset)
                        gSourceFile.Data(n) = 38 ' prevents parser from processing
                    Else
                        sName = Replace(Mid$(gSourceFile.Text, n + 1, p - n - lOffset - 2), vbLf, vbNullString)
                        Mid$(gSourceFile.Text, n, p - n - lOffset) = Space$(p - n - lOffset)
                        pFlags = pFlags Xor wbpBracketsMarked
                    End If
                    pFlags = pFlags Xor wbpMarkBrackets
                    If lMode = 3 Then p = lMax: lMode = 2
                End If
            ElseIf lMode = 1 Then           ' looking for comma or end of params
                If c = vbKeyComma Then
                    lMode = 2
                ElseIf gSourceFile.Data(p - 1) = vbKeyComma Then
                    lMode = 2
                ElseIf gSourceFile.Data(p - 1) = vbKeyParenthesis2 Then
                    lMode = 2: p = lMax     ' closing parenthesis
                ElseIf p = lMax Then
                    lMode = 2               ' end of line, no commas
                ElseIf lFlags = 0 Then
                    If c = vbKeyA Then      ' look for "As" if not yet vartyped
                        If gSourceFile.Data(n + 1) = 115 And p - n = 2 Then lFlags = 1
                    End If
                End If
            End If
            If lMode = 2 Then               ' end of parameter parsed
                lRecord(rteDiscreps) = 0
                If (m_ValidationChecks And vtZombie) <> 0 Then
                    If lFlags = 0 Then          ' still need vartype checked?
                        If pvIsDefTyped(AscW(sName)) = True Then
                            lFlags = 1
                        ElseIf lType = itAPI Or lType = itEvent Then
                            ' API/Event parameters are not zombie checked so log now
                            lRecord(rteRecID) = modMain.CreateRecord(lRecord(rteParent), sName, itParameter, 0, 0, , , , , , , scpLocal, chrV & chrZ)
                        End If
                    End If                      ' track for zombie checks
                End If
                If lType <> itAPI And lType <> itEvent Then
                    pvAddNameToList sName
                    lRecord(rteRecID) = -m_ZombieList.Names.Count
                    If (m_ValidationChecks And vtZombie) <> 0 Then
                        If lFlags = 0 Then lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtVarType
                        If lType = itMethod Then lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtZombie
                    End If
                    lRecord(rteCRC) = modMain.CRCItem(sName, True)
                    pvFindWordEx m_ZombieList.List(zlLocals), lRecord(), True
                End If
                lFlags = 0: lMode = 0: lOffset = 0
                pFlags = pFlags Or wbpMarkBrackets
            End If
        End If
    Loop

End Sub

Private Sub pvParseDeclares(rs As ADODB.Recordset, lDecStart As Long, _
                            lStart As Long, lMax As Long)
                            
    ' individual Dim,ReDim,Const statements within methods are processed here

    Dim n&, p&, lFlags&, lType&, pFlags&
    Dim sName$, lMode&, lOffset&
    Dim lRecord&(rteCRC To rteDiscreps)
    Const ParsePreserve = "Preserve"
    
    If gSourceFile.Data(lDecStart) = vbKeyC Then        ' Const statement
        lType = 1
    ElseIf gSourceFile.Data(lDecStart) = vbKeyR Then    ' ReDim statement
        lType = 3
    Else
        lType = 2                                       ' Dim/Static statement
    End If
    
    p = lStart
    lRecord(rteParent) = rs.Fields(recID).Value
    If lType = 1 Then
        lRecord(rteType) = itConstant
    Else
        lRecord(rteType) = itVariable
    End If
    pFlags = wbpNotCommaEOW Or wbpParentsCntr Or wbpMarkBrackets
    Do
        modMain.ParseNextWordEx p, lMax, n, p, pFlags
        If n = p Then Exit Do
        If lMode = 0 Then                       ' looking for name
            lMode = 1
            If lType = 3 Then                   ' ReDim
                lType = 5                       ' only hit this once
                If gSourceFile.Data(n) = vbKeyP Then
                    If p - n = 8 And gSourceFile.Data(n + 1) = 114 Then
                        ' keyword Preserve implicity means variable is vartyped
                        If Mid$(gSourceFile.Text, n, 8) = ParsePreserve Then
                            lMode = 0: lFlags = 1: lType = 4
                        End If
                    End If
                End If
            End If
            If lMode = 1 Then                   ' looking for name
                If gSourceFile.Data(p - 1) = vbKeyComma Then
                    lOffset = 1: lMode = 2      ' ex: Dim X, Y
                End If
                If modMain.IsVarTyped(gSourceFile.Data(p - 1 - lOffset)) = 1 Then
                    lOffset = lOffset + 1: lFlags = 1
                End If
                If (pFlags And wbpBracketsMarked) = 0 Then
                    sName = Mid$(gSourceFile.Text, n, p - n - lOffset)
                    gSourceFile.Data(n) = 38
                Else
                    sName = Replace(Mid$(gSourceFile.Text, n + 1, p - n - lOffset - 2), vbLf, vbNullString)
                    Mid$(gSourceFile.Text, n, p - n - lOffset) = Space$(p - n - lOffset)
                    pFlags = pFlags Xor wbpBracketsMarked
                End If
                If lType = 1 Then lFlags = 1    ' consts are vartyped
                pFlags = pFlags Xor wbpMarkBrackets
            End If
        ElseIf lMode = 1 Then                   ' looking for comma
            If gSourceFile.Data(p - 1) = vbKeyComma Then
                lMode = 2
            ElseIf gSourceFile.Data(n) = vbKeyComma Then
                lMode = 2
            ElseIf p = lMax Then
                lMode = 2
            ElseIf lFlags = 0 Then              ' looking for chrsAs if vartype needed
                If gSourceFile.Data(n) = vbKeyA Then
                    If gSourceFile.Data(n + 1) = 115 And p - n = 2 Then lFlags = 1
                End If
            End If
        End If
        If lMode = 2 Then                       ' process variable/const
            lRecord(rteDiscreps) = vtZombie
            lRecord(rteCRC) = modMain.CRCItem(sName, True)
            If lType > 2 Then                   ' ReDim: check against class/globals
                lMode = -1
                If pvFindWordEx(m_ZombieList.List(zlLocals), lRecord(), False) = 0 Then
                    If pvFindWordEx(m_ZombieList.List(zlCpVars), lRecord(), False) = 0 Then
                        If pvFindWordEx(m_ZombieList.List(zlGblVars), lRecord(), False) = 0 Then
                            If InStr(sName, chrDot) = 0 Then
                                ' ReDim without prior Dim, track for zombie checks
                                lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtRedim
                                lMode = lFlags
                            End If
                        End If
                    End If
                End If
                lFlags = lMode
            End If
            If lFlags <> -1 Then                ' else not a local variable
                If lFlags = 0 Then              ' need vartype still?
                    If pvIsDefTyped(AscW(sName)) = False Then
                        lRecord(rteDiscreps) = lRecord(rteDiscreps) Or vtVarType
                    End If
                End If
                lRecord(rteRecID) = -(m_ZombieList.Names.Count + 1)
                If (m_ValidationChecks And vtDupeDecs) <> 0 Then
                    If pvFindWordEx(m_ZombieList.List(zlCpVars), lRecord(), False) = 0 Then
                        lRecord(rteRecID) = modMain.CreateRecord(lRecord(rteParent), sName, lRecord(rteType), 0, 0)
                    End If
                End If
                If lRecord(rteRecID) < 0 Then pvAddNameToList sName
                pvFindWordEx m_ZombieList.List(zlLocals), lRecord, True
            End If
            lMode = 0: lOffset = 0: lFlags = 0
            pFlags = pFlags Or wbpMarkBrackets
        End If
    Loop

End Sub

Private Sub pvAddNameToList(sName As String)

    ' helper function to resize local's Name() array
    
    With m_ZombieList.Names
        .Count = .Count + 1
        If .Count = 1 Then
            ReDim .fName(1 To 10)
        ElseIf .Count > UBound(.fName) Then
            ReDim Preserve .fName(1 To .Count + 20)
        End If
        .fName(.Count) = sName
    End With

End Sub

Private Sub pvDoStopEndUsage(rs As ADODB.Recordset)

    ' log Stop or End found in code

    If (m_ValidationChecks And vtZombie) = 0 Then Exit Sub

    Dim lRecord&(0 To 0), lID&
    lRecord(rteCRC) = modMain.CRCItem(rs.Fields(recName).Value, True)
    lID = pvFindWordEx(m_ZombieList.List(zlCpMethods), lRecord(), False)
    If lID <> 0 Then
        With m_ZombieList.List(zlCpMethods)
            .Item(rteDiscreps, lID) = .Item(rteDiscreps, lID) Or vtStopEnd
        End With
    End If

End Sub

Private Sub pvDoEmptyCode(rs As ADODB.Recordset)

    ' log method that contains no excutable statements
    
    If (m_ValidationChecks And vtZombie) = 0 Then Exit Sub

    Dim lRecord&(0 To 0), lID&, p&
    If rs.Fields(recType).Value <> itClassEvent Then
        lRecord(rteCRC) = modMain.CRCItem(rs.Fields(recName).Value, True)
        lID = zlCpMethods
        p = pvFindWordEx(m_ZombieList.List(lID), lRecord(), False)
        If p = 0 Then                       ' not a private method, try global
            lID = zlGblMethods
            p = pvFindWordEx(m_ZombieList.List(zlGblMethods), lRecord(), False)
            If p = 0 Then Exit Sub          ' sanity check, should not happen
        End If
        m_ZombieList.List(lID).Item(rteDiscreps, p) = _
            m_ZombieList.List(lID).Item(rteDiscreps, p) Or vtEmptyCode
    End If

End Sub

Private Sub pvDoRaiseEvent(lStart As Long, lMax As Long)

    ' zombie check RaiseEvent statements

    If (m_ValidationChecks And vtZombie) = 0 Then Exit Sub

    Dim lRecord&(0 To 0), p&, sName$
    
    modMain.ParseNextWord lStart, lMax, p, lStart, 0&
    sName = Mid$(gSourceFile.Text, p, lStart - p)
    lRecord(rteCRC) = modMain.CRCItem(sName, True)
    p = pvFindWordEx(m_ZombieList.List(zlCpRaised), lRecord(), False)
    If p <> 0 Then                       ' not a private method, try global
        m_ZombieList.List(zlCpRaised).Item(rteDiscreps, p) = _
            m_ZombieList.List(zlCpRaised).Item(rteDiscreps, p) And Not vtZombie
    End If

End Sub

Private Sub pvDoAddressOf(lStart As Long, lMax As Long)

    ' ensure AddressOf target is properly identified to prevent false positive zombies

    If (m_ValidationChecks And vtZombie) = 0 Then Exit Sub

    Dim lRecord&(0 To 0), lID&
    Dim p&, n&, pFlags&, sName$
    Dim rs As ADODB.Recordset
    
    pFlags = wbpMarkBrackets
    p = lStart
    modMain.ParseNextWordEx p, lMax, p, n, pFlags
    sName = Mid$(gSourceFile.Text, p, n - p)
    If modMain.IsVarTyped(gSourceFile.Data(p - 1)) Then
        sName = Mid$(gSourceFile.Text, p, n - p - 1)
    Else
        sName = Mid$(gSourceFile.Text, p, n - p)
    End If
    If (pFlags And wbpBracketsMarked) <> 0 Then
        sName = Replace(sName, vbLf, vbNullString)
    End If
    
    lRecord(rteCRC) = modMain.CRCItem(sName, True)
    pFlags = 0: lStart = n
    lID = zlCpMethods
    p = pvFindWordEx(m_ZombieList.List(lID), lRecord(), False)
    If p = 0 Then                       ' not a private method, try global
        lID = zlGblMethods
        p = pvFindWordEx(m_ZombieList.List(zlGblMethods), lRecord(), False)
        If p = 0 Then Exit Sub          ' sanity check, should not happen
    End If
    If m_ZombieList.List(lID).Item(rteType, p) = itClassEvent Then Exit Sub
    
    m_ZombieList.List(lID).Item(rteType, p) = itClassEvent
    m_ZombieList.List(lID).Item(rteDiscreps, p) = _
        m_ZombieList.List(lID).Item(rteDiscreps, p) And Not (vtZombie Or vtEmptyCode)
    
    Set rs = gParsedItems.Clone
    n = m_ZombieList.List(lID).Item(rteRecID, p)
    rs.Find modMain.SetQuery(recID, qryIs, n), , , 1&
    If rs.EOF = True Then GoTo exitRoutine
    If rs.Fields(recType).Value = itClassEvent Then GoTo exitRoutine
    rs.Fields(recType).Value = itClassEvent
    rs.Update
    
    ' any logged method parameters need to have zombie flag removed
    rs.Filter = modMain.SetQuery(recType, qryIs, itParameter, qryAnd, _
                                recParent, qryIs, n)
    Do Until rs.EOF = True
        If InStr(rs.Fields(recDiscrep).Value, chrZ) <> 0 Then
            rs.Fields(recDiscrep).Value = Replace(rs.Fields(recDiscrep).Value, chrZ, vbNullString)
            rs.Update
        End If
        rs.MoveNext
    Loop
exitRoutine:
    rs.Close: Set rs = Nothing

End Sub

Private Function pvFindWordEx(wordList As WordListStructEx, _
                         lCriteria() As Long, bAutoAppend As Boolean, _
                         Optional cIndex As Long = 0, _
                         Optional AllowDupes As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must be kept sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    ' note: in this routine, passed array must be one-bound, not zero-bound
    
    Dim UB&, LB&, newIndex&
    
    If wordList.Count = 0 Then         ' initialize as neeed
        If bAutoAppend Then
            wordList.Count = 1
            ReDim wordList.Item(rteCRC To rteDiscreps, 1 To 50)
            CopyMemory wordList.Item(rteCRC, 1), lCriteria(rteCRC), 20&
        End If
    Else
        With wordList
            LB = 1: UB = .Count         ' set 1-bound range
            Do
                newIndex = LB + ((UB - LB) \ 2)
                If lCriteria(cIndex) = .Item(cIndex, newIndex) Then
                    pvFindWordEx = newIndex
                    Exit Do
                ElseIf lCriteria(cIndex) < .Item(cIndex, newIndex) Then
                    UB = newIndex - 1
                Else
                    LB = newIndex + 1
                End If
            Loop Until LB > UB
            
            If pvFindWordEx = 0 Or AllowDupes = True Then
                If bAutoAppend = True Then ' insert word into sorted list
                    If pvFindWordEx = 0 Then
                        If newIndex < LB Then newIndex = newIndex + 1
                    Else
                        newIndex = newIndex + 1
                    End If
                    .Count = .Count + 1
                    If .Count > UBound(.Item, 2) Then
                        ReDim Preserve .Item(rteCRC To rteDiscreps, 1 To .Count + 99)
                    End If
                    If newIndex < .Count Then
                        CopyMemory ByVal VarPtr(.Item(rteCRC, newIndex + 1)), _
                                   ByVal VarPtr(.Item(rteCRC, newIndex)), (.Count - newIndex) * 20&
                    End If
                    CopyMemory .Item(rteCRC, newIndex), lCriteria(rteCRC), 20&
                End If
            End If
        End With
    End If
    
End Function

Private Sub pvListMaliciousDLLs(rs As ADODB.Recordset)

    ' look for specific DLLs and VB functions that could be used for malintent

    Dim p&, n&, sName$, bOk As Boolean
    Dim v As Variant, vBkMk As Variant
    
    rs.Filter = modMain.SetQuery(recType, qryIs, itAPI)
    If rs.EOF = True Then Exit Sub
    rs.Sort = recIdxAttr & chrComma & recAttr2
    For n = 0 To 6
        Select Case n
        Case 0: sName = "oleaut32.dll"
            v = Array("DispCallFunc")
        Case 1: sName = "shell32.dll"
            v = Array("ShellExecute")
        Case 2: sName = "kernel32.dll"
            v = Array("CreateProcess", "HeapAlloc", "LoadResource", "OpenProcess", "SetFileTime", "SetLocalTime", "SetSystemTime", "VirtualAlloc", "WriteProcessMemory")
        Case 3: sName = "user32.dll"
            v = Array("keybd_event", "SendInput")
        Case 4: sName = "ole32.dll"
            v = Array("CoCreateInstance", "CoGetClassObject", "CoGetInstanceFromFile", "CoGetInstanceFromIStorage", "DllRegisterServer", "DllUnregisterServer")
        Case 5: sName = "advapi32.dll"
            If (m_ValidationChecks And vtRegReadDLLs) = 0 Then
                v = Array("RegCopy", "RegFlush", "RegReplace", "RegRestore", "RegSave", "RegSet", "RegUnload")
            Else
                v = Array("RegCopy", "RegEnum", "RegFlush", "RegGet", "RegLoad", "RegOpen", "RegQuery", "RegReplace", "RegRestore", "RegSave", "RegSet", "RegUnload")
            End If
        Case 6: sName = "shlwapi.dll"
            If (m_ValidationChecks And vtRegReadDLLs) = 0 Then
                v = Array("SHCopyKey", "SHDelete", "SHRegCreate", "SHRegDelete", "SHRegDuplicate", "SHRegSet", "SHRegWrite", "SHSetValue")
            Else
                v = Array("SHCopyKey", "SHDelete", "SHEnum", "SHGetValue", "SHOpenReg", "SHQuery", "SHRegCreate", "SHRegDelete", _
                          "SHRegDuplicate", "SHRegEnum", "SHRegGet", "SHRegOpen", "SHRegQuery", "SHRegSet", "SHRegWrite", "SHSetValue")
            End If
        End Select
        rs.Filter = modMain.SetQuery(recType, qryIs, itAPI, qryAnd, recAttr2, qryIs, sName)
        If rs.EOF = False Then
            vBkMk = Empty
            For p = 0 To UBound(v)
                If IsEmpty(vBkMk) = True Then rs.MoveFirst Else rs.Bookmark = vBkMk
                rs.Find modMain.SetQuery(recAttr, qryLike, v(p) & chrPct)
                If rs.EOF = False Then
                    Do
                        vBkMk = rs.Bookmark
                        If (m_ValidationChecks And vtZombie) = 0 Then
                            bOk = True  ' zombie checks not done, assume API is not zombie
                        Else            ' else only report if API is not a zombie
                            bOk = CBool(InStr(rs.Fields(recDiscrep).Value, chrZ) = 0)
                        End If
                        If bOk = True Then
                            rs.Fields(recDiscrep).Value = rs.Fields(recDiscrep).Value & chrM
                            rs.Update: bOk = False
                        End If
                        rs.MoveNext
                        If rs.EOF = True Then Exit Do
                        If (rs.Fields(recAttr).Value Like v(p) & chrAsterisk) = False Then Exit Do
                    Loop
                End If
            Next
        End If
    Next
    
End Sub

Private Sub pvDoPostZombieCheck(rs As ADODB.Recordset)

    Dim rsMbrs As ADODB.Recordset
    Dim vBkMk As Variant
    Dim n&, p&, c&, sValue$

    pvResolveZombies rs, zlGblPages, zlGblUDTs
    
    ' /// Step 1. Zombie resolution for properties
    '     for any property Get that is in zombie state and has a
    '     property Let/Set that is not, then remove zombie state for that Get
    
    rs.Sort = vbNullString              ' get list of all Properties
    rs.Filter = modMain.SetQuery(recType, qryIs, itMethod, _
                                qryAnd, recFlags, qryLT, 0)
    If rs.EOF = False Then
        rs.Sort = recGrp                ' CRC of property name & code page name
        Do                              ' for each zombie
            rs.Find modMain.SetQuery(recDiscrep, qryLike, "%Z%")
            If rs.EOF = True Then Exit Do
            ' locate the first & last property in that group. mark the first position
            n = rs.Fields(recGrp).Value: p = 0
            sValue = modMain.SetQuery(recGrp, qryNot, n)
            rs.Find sValue, , adSearchBackward
            rs.MoveNext: c = rs.Fields(recID).Value ' c = first property of group
            rs.Find sValue: rs.MovePrevious         ' find & mark last in group
            vBkMk = rs.Bookmark: n = rs.Fields(recID).Value ' n = last property in group
            Do                          ' move from bottom up and see if any are not zombies
                If p = 0 Then
                    If InStr(rs.Fields(recDiscrep).Value, chrZ) = 0 Then p = 1
                End If              ' continue until we get to 1st in group
                If rs.Fields(recID).Value = c Then Exit Do
                rs.MovePrevious
            Loop
            If p = 1 Then               ' non-zombie found
                Do                      ' move from top down, removing zombie flag
                    rs.Fields(recDiscrep).Value = Replace(rs.Fields(recDiscrep).Value, chrZ, vbNullString)
                    If rs.Fields(recID).Value = n Then Exit Do
                    rs.MoveNext
                Loop
                If rs.EditMode <> adEditNone Then rs.UpdateBatch
            Else
                rs.Bookmark = vBkMk
            End If
            rs.MoveNext: sValue = vbNullString
        Loop Until rs.EOF = True
    End If
    
    ' /// Step 2. Zombie resolution for enums/members
    '     If enum has zombie flag, then if any member does not, remove the flag
    '     If all members have zombie flag, remove flag from each member and
    '     apply the flag to the enum instead
    '     variable p used for tracking:
    '       0x001   enum itself is not a zombie, members are therefore not zombies
    '       0x002   at least one member is not a zombie
    '       0x004   at least one member is a zombie
    '       0x008   all members are excluded
    rs.Filter = modMain.SetQuery(recType, qryIs, itEnum)
    If rs.EOF = False Then
        Set rsMbrs = gParsedItems.Clone
        rsMbrs.Filter = modMain.SetQuery(recType, qryIs, itEnumMember)
        rsMbrs.Sort = recParent
        Do
            n = rs.Fields(recID).Value
            If InStr(rs.Fields(recDiscrep).Value, chrZ) = 0 Then p = 1 Else p = 0
            If (m_ValidationChecks And vtExcludeEMbrZombies) <> 0 Then p = p Or 8
            
            rsMbrs.Find modMain.SetQuery(recParent, qryIs, n), , , 1&
            If rsMbrs.EOF = False Then
                vBkMk = rsMbrs.Bookmark
                Do
                    If rsMbrs.Fields(recParent).Value <> n Then Exit Do
                    If InStr(rsMbrs.Fields(recDiscrep).Value, chrZ) = 0 Then
                        p = p Or 2          ' member not a zombie
                    Else
                        ' if parent not a zombie or excluding members, remove flag
                        If (p And 9) <> 0 Then
                            rsMbrs.Fields(recDiscrep).Value = Replace(rsMbrs.Fields(recDiscrep).Value, chrZ, vbNullString)
                            rsMbrs.Update
                        Else
                            p = p Or 4
                        End If
                    End If
                    rsMbrs.MoveNext
                Loop Until rsMbrs.EOF = True
                If (p And 3) = 2 Then       ' enum is zombie, but at least one member is not
                    If InStr(rs.Fields(recDiscrep).Value, chrZ) <> 0 Then
                        rs.Fields(recDiscrep).Value = Replace(rs.Fields(recDiscrep).Value, chrZ, vbNullString)
                        rs.Update
                    End If
                ElseIf p = 4 Then           ' enum & all members are zombies
                    rsMbrs.Bookmark = vBkMk ' remove zombie state from all members
                    Do
                        If rsMbrs.Fields(recParent).Value <> n Then Exit Do
                        rsMbrs.Fields(recDiscrep).Value = Replace(rsMbrs.Fields(recDiscrep).Value, chrZ, vbNullString)
                        rsMbrs.Update
                        rsMbrs.MoveNext
                    Loop Until rsMbrs.EOF = True
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF = True
        rsMbrs.Close: Set rsMbrs = Nothing
    End If
    
    ' /// Step 3. Zombie resolution for uncompiled usercontrols
    rs.Filter = modMain.SetQuery(recType, qryIs, itCodePage)
    Do Until rs.EOF = True
        If (rs.Fields(recFlags).Value And iaMaskCodePage) = iaUC Then
            If InStr(rs.Fields(recDiscrep).Value, chrZ) <> 0 Then
                If rsMbrs Is Nothing Then
                    Set rsMbrs = gParsedItems.Clone
                    rsMbrs.Filter = modMain.SetQuery(recType, qryIs, itControl)
                End If
                rsMbrs.Find SetQuery(recAttr, qryIs, rs.Fields(recName).Value), , , 1&
                If rsMbrs.EOF = False Then
                    rs.Fields(recDiscrep).Value = Replace(rs.Fields(recDiscrep).Value, chrZ, vbNullString)
                    rs.Update
                End If
            End If
        End If
        rs.MoveNext
    Loop
    If Not rsMbrs Is Nothing Then
        rsMbrs.Close: Set rsMbrs = Nothing
    End If
    
    ' /// Step 4. Zombie resolution for startup form
    gParsedItems.Bookmark = gSourceFile.ProjBookMark
    sValue = gParsedItems.Fields(recAttr2).Value
    n = InStr(sValue, chrSemi) + 1
    sValue = Mid$(sValue, n)
    If sValue <> "(None)" Then
        With gParsedItems
            If LCase$(sValue) = "sub main" Then
                .Find modMain.SetQuery(recGrp, qryIs, vbKeyM), , , 1&
                If .EOF = False Then
                    .Fields(recDiscrep).Value = Replace(.Fields(recDiscrep), chrZ, vbNullString)
                    .Update
                End If
                .Find modMain.SetQuery(recID, qryIs, .Fields(recParent).Value), , , 1&
            Else
                sValue = .Fields(recName).Value & chrDot & sValue
                .Find modMain.SetQuery(recName, qryIs, sValue), , , 1&
            End If
            If .EOF = False Then
                .Fields(recDiscrep).Value = Replace(.Fields(recDiscrep), chrZ, vbNullString)
                .Update
            End If
        End With
    End If
    
End Sub

Private Sub pvJoinContinuation(ByVal lStart As Long, lMax As Long)

    ' Only called when an object, or its property/method, is separated
    '   from the object identifier (. or !) by a continuation character.
    ' Any other scenario where a split line exists is not sent here.
    
    ' goal of this routine is NOT to join lines exactly the same way VB
    '   would join the line. Rather the goal is to ensure objects are
    '   attached to their property/method after join is finished.
    '   For example, ensure this line: Me _
    '                                     .Left = 0
    '                is joined as: Me.Left = 0
    ' ------------------------------------------------------------------
    ' Join rules used by this routine, first come - first served
    ' [Prev Line Last Char] [Cur Line First Char]   Result
    ' ------------------------------------------------------------------
    '           .           and     any             join w/o space btwn chars
    '           (           or      )               join w/o space btwn chars
    '           )           and     (               join w/o space btwn chars
    '          any          and     , or ; or !     join w/o space btwn chars
    '       not keyword     and     .               join w/o space btwn chars
    ' all other scenarios                           join with space btwn chars
    '(*) keyword includes operands/operators

    Dim p&, n&, lPos&
    
    p = lStart: lMax = Abs(lMax)
    Do
        modMain.ParseNextWordEx p, lMax, n, p, wbpContinuation Or wbpNotCommaEOW
        If n = lStart Then
            lPos = p
        ElseIf n <> p Then
            If gSourceFile.Data(lPos - 1) = vbKeyDot Or _
               gSourceFile.Data(lPos - 1) = vbKeyParenthesis Then
                    ' previous line ends in dot or parenthesis
                    lPos = lPos - 1         ' removes trailing space from prev line
            Else
                Select Case gSourceFile.Data(n)      ' current line starts with...
                Case vbKeyParenthesis2, vbKeyComma, vbKeySemicolon, vbKeyBang
                    lPos = lPos - 1
                Case vbKeyDot
                    Select Case gSourceFile.Data(lPos - 1)
                    Case vbKeyParenthesis2, 35, 37, 38, 42 To 47, 59, 60, 61, 62, 64, 92
                        ' previous line last char is one of these: )#%&*+,-./;<=>@\
                        gSourceFile.Data(lPos) = vbKeySpace
                    Case Else
                        ' test for previous line ending in keyword
                        For lStart = lPos - 1 To lStart + 1 Step -1
                            If modMain.IsWhiteSpace(gSourceFile.Data(lStart)) = 1 Then Exit For
                            If modMain.IsEndOfLine(gSourceFile.Data(lStart)) = 1 Then Exit For
                        Next
                        If FindWord(m_Keywords, Mid$(gSourceFile.Text, lStart + 1, lPos - lStart - 1), _
                           vbBinaryCompare, False) = 0 Then
                            lPos = lPos - 1 ' if not keyword, then adjacent join
                        Else
                            gSourceFile.Data(lPos) = vbKeySpace
                        End If
                    End Select
                End Select
            End If
            ' shift next line to end of previous line
            CopyMemory gSourceFile.Data(lPos + 1), gSourceFile.Data(n), (p - n) * 2
            ' set position @ end of line, include trailng space
            lPos = lPos + p - n + 1
        End If
    Loop While p < lMax
    lMax = lPos
    
End Sub

Private Sub pvDoZombieCheck(wordList As WordListStruct, bDecsOnly As Boolean)

    ' Checks each parsed word in m_WordList against project items in m_ZombieList.
    ' Parsed words may have object identifiers (. or !). If so, then the "words"
    ' are separated by identifier and checked until a hit found or no more identifiers
    
    ' lFail & sFail are shortcircuits to help zombie checks move along faster.
    ' Here's an example of a hypothetical word list portion. Always sorted.
    ' ListView1
    ' ListView1.ListItems
    ' ListView1.ListItems.Text
    ' ListView1.ListItems.SubItem
    
    ' Basically, if a previous "word" failed and current "word" is a method/property
    ' of that the previous "word", then the current word fails the check also...
    ' ListView1                     -- found in zombie list
    ' ListView1.ListItems           -- not found
    ' ListView1.ListItems.Text      -- skipped, method/property of failed word
    ' ListView1.ListItems.SubItem   -- skipped, method/property of failed word

    Dim n&, p&, lID&, lRecord&(0 To 0)
    Dim sName$, sFail$, lFail&, c&
    
    For n = 1 To wordList.Count
        sName = wordList.List(n)      ' current parsed word
        If lFail <> 0 Then              ' check against previously failed search
            If sFail <> Left$(sName, lFail) Then lFail = 0
        End If
        If lFail = 0 Then               ' if not previously failed...
            Do
                lRecord(rteCRC) = modMain.CRCItem(sName, True)
                For p = 0 To m_ZombieList.Count - 1
                    If bDecsOnly = True Then
                        Select Case p
                        Case zlCpEnums, zlCpUDTs, zlGblPages, zlGblEnums, zlGblUDTs
                            lID = pvFindWordEx(m_ZombieList.List(p), lRecord(), False)
                        Case 0: lID = 0
                        End Select
                    Else
                        lID = pvFindWordEx(m_ZombieList.List(p), lRecord(), False)
                    End If
                    If lID <> 0 Then
                        If p = zlCpShadow Then Exit For
                        If p = zlCpMethods Then
                            ' prevent a method from zombie checking itself
                            If m_ZombieList.List(p).Item(rteRecID, lID) = _
                                m_ZombieList.List(p).pReserved Then Exit For
                        End If
                        m_ZombieList.List(p).Item(rteDiscreps, lID) = _
                            m_ZombieList.List(p).Item(rteDiscreps, lID) And Not vtZombie
                        Exit For
                    End If
                Next
                If lID <> 0 Then Exit Do    ' found
                c = InStrRev(sName, chrDot) ' else get previous word portion
                If c = 0 Then Exit Do
                If lFail = 0 Then           ' set fail state & fail string
                    lFail = Len(sName) + 1: sFail = sName & chrDot
                End If
                sName = Left$(sName, c - 1) ' next word portion to check
            Loop
        End If
    Next
    If bDecsOnly = True Then
        m_DecsList.Count = 0: Erase m_DecsList.List()
    Else
        m_WordList.Count = 0: Erase m_WordList.List()
    End If
    
End Sub

Private Sub pvResolveZombies(rs As ADODB.Recordset, lStart As Long, lMax As Long)
'
    ' called at three points
    ' 1. After a method has been processed, resolving only method declarations
    ' 2. After all code-page methods have been processed, resolving declarations
    ' 3. After entire project processed, resolving global declarations
    ' FYI. m_ZombieList.List(zlCpShadow) is never processed

    Dim n&, p&, lID&, lMask&, lLast&, sValue$
    Dim lRecord&(rteCRC To rteDiscreps), vBkMark As Variant
    Dim tList As WordListStructEx
    
    vBkMark = rs.Bookmark: lLast = vtRedim + vtRedim
    For n = lStart To lMax                          ' range of zombie lists to check
        If m_ZombieList.List(n).Count <> 0 Then
            With m_ZombieList.List(n)
                ' since items can be listed & referenced various ways, we
                ' need to group same items (referenced differently) on recordID
                For p = 1 To .Count
                    CopyMemory lRecord(rteCRC), .Item(rteCRC, p), 20
                    lID = pvFindWordEx(tList, lRecord, True, rteRecID)
                    If lID <> 0 Then                ' already added to grouping
                        tList.Item(rteDiscreps, lID) = ((tList.Item(rteDiscreps, lID) Or .Item(rteDiscreps, p)) And &HFFFFFFFC) _
                                Or (tList.Item(rteDiscreps, lID) And .Item(rteDiscreps, p))
                    End If
                Next
                Erase .Item(): .Count = 0           ' done grouping
            End With
            
            For p = 1 To tList.Count                ' collect various discrepancies
                If tList.Item(rteDiscreps, p) <> 0 Then
                    lMask = vtVarType
                    Do Until lMask = lLast          ' max discpancies to collect
                        Select Case (tList.Item(rteDiscreps, p) And lMask)
                        Case vtVarType: sValue = sValue & chrV
                        Case vtZombie: sValue = sValue & chrZ
                        Case vtRedim: sValue = sValue & chrR
                        Case vtEmptyCode: sValue = sValue & "E"
                        Case vtStopEnd: sValue = sValue & "X"
                        Case vtWithNoEvents: sValue = sValue & chrW
                        Case vtMalicious: sValue = sValue & chrM
                        End Select
                        lMask = lMask + lMask
                    Loop
                End If
                If tList.Item(rteRecID, p) < 0 Then ' cache-only reference
                    If LenB(sValue) <> 0 Then       ' log if has discrepancies
                        modMain.CreateRecord tList.Item(rteParent, p), _
                            m_ZombieList.Names.fName(-tList.Item(rteRecID, p)), _
                            tList.Item(rteType, p), 0, 0, , , , , , , scpLocal, sValue
                    End If
                ElseIf (m_ValidationChecks And vtZombie) <> 0 Then
                    ' logged in rs, update discrepancies
                    rs.Find modMain.SetQuery(recID, qryIs, tList.Item(rteRecID, p)), , , 1&
                    If rs.EOF = False Then
                        rs.Fields(recDiscrep).Value = sValue
                        rs.Update
                    Else
                        gParsedItems.Find modMain.SetQuery(recID, qryIs, tList.Item(rteRecID, p)), , , 1&
                        gParsedItems.Fields(recDiscrep).Value = sValue
                        gParsedItems.Update
                    End If
                End If
                sValue = vbNullString
            Next
            tList.Count = 0: Erase tList.Item()
        End If
    Next
    rs.Bookmark = vBkMark

End Sub

Private Sub pvProcessParsedWord(rs As ADODB.Recordset, _
                               sName As String, sPrevWord As String, _
                               lStmt As Long, lStart As Long, lMax As Long)
                            
    ' routine looks for variant usage of string functions, i.e., Mid vs Mid$
    ' Date,String were pre-filtered and not sent here if used as a vartype
    ' Note: false negatives possible when a custom method has the same name
    '   as one of the functions tested.
    
    If Len(sName) < 3 Then GoTo addWordToList_
    If (m_ValidationChecks And (vtVarFunc Or vtMalicious)) = 0 Then GoTo addWordToList_
    
    Dim sWord$, sProcessed$
    Dim lFlags&, p&, n&, lID&, lIndex&
    '   0x001   has $ suffix
    '   0x002   Left method with parameters
    '   0x004   may be statement vs. method
    '   0x008   date/time keyword
    '   0x010   potential malicious statement
    '   0x020   has VB. or VBA. prefix
    Select Case (gSourceFile.Data(lStart) And Not 32)
    Case vbKeyC, vbKeyD, vbKeyI, vbKeyL, vbKeyS, vbKeyT
        ' methods can be in vtVarFuncs or vtMalicious
        If (m_ValidationChecks And vtVarFunc) <> 0 Then
            lIndex = modMain.FindWord(m_VarFuncs, sName, vbTextCompare, False)
        End If
        If lIndex = 0 Then
            If (m_ValidationChecks And vtMalicious) = 0 Then GoTo addWordToList_
            lIndex = modMain.FindWord(m_Malicious, sName, vbTextCompare, False)
            If lIndex = 0 Then GoTo addWordToList_
            lFlags = lFlags Or 16
        End If
    Case vbKeyG
        ' methods only in vtMalicious
        If (m_ValidationChecks And vtMalicious) = 0 Then GoTo addWordToList_
        lIndex = modMain.FindWord(m_Malicious, sName, vbTextCompare, False)
        If lIndex = 0 Then GoTo addWordToList_
        lFlags = lFlags Or 16
    Case vbKeyE, vbKeyF, vbKeyH, vbKeyM, vbKeyO, vbKeyR, vbKeyU
        ' methods only in vtVarFuncs
        If (m_ValidationChecks And vtVarFunc) = 0 Then GoTo addWordToList_
        lIndex = modMain.FindWord(m_VarFuncs, sName, vbTextCompare, False)
        If lIndex = 0 Then GoTo addWordToList_
    Case vbKeyV
        ' methods staring with VB. or VBA.
        If (m_ValidationChecks And vtVarFunc) <> 0 Then
            lIndex = modMain.FindWord(m_VarFuncs, sName, vbTextCompare, False)
        End If
        If lIndex = 0 Then
            If (m_ValidationChecks And vtMalicious) = 0 Then GoTo addWordToList_
            lIndex = modMain.FindWord(m_Malicious, sName, vbTextCompare, False)
            If lIndex = 0 Then GoTo addWordToList_
            lFlags = lFlags Or 16
        End If
        lFlags = lFlags Or 32
    Case Else: GoTo addWordToList_
    End Select
    
    If (lFlags And 16) = 0 Then
        If modMain.IsVarTyped(gSourceFile.Data(lMax - 1)) = 1 Then
            If gSourceFile.Data(lMax - 1) <> 36 Then GoTo addWordToList_
            lFlags = 1                          ' has $ suffix
        End If
        sWord = m_VarFuncs.List(lIndex)
    Else
        sWord = m_Malicious.List(lIndex)
    End If
    
    n = InStrRev(sWord, chrDot)
    If n <> 0 Then
        sWord = Mid$(sWord, n + 1): lIndex = 0  ' will get preferred index later, if needed
    End If
    If (lFlags And 32) = 32 Then GoTo logDiscrep_
    
    Select Case AscW(sWord)
    Case vbKeyE
        ' Error 13 is a statement, not a method
        ' MsgBox Error(13) is a method
        If sWord = "Error" Then
            If sPrevWord = "On" Then GoTo exitRoutine
            modMain.ParseNextWordEx lMax, gSourceFile.length, n, p, wbpFirstChar
            If gSourceFile.Data(n) <> vbKeyParenthesis Then lFlags = lFlags Or 4
        End If
    Case vbKeyL
        If sWord = "Left" Then
            ' Left w/o parameters or on LHS of expression is not a VBA method
            If lStmt = lStart Then GoTo addWordToList_
            If sPrevWord = "Then" Then GoTo addWordToList_
            If sPrevWord = "Else" Then GoTo addWordToList_
            modMain.ParseNextWordEx lMax, gSourceFile.length, n, p, wbpFirstChar
            If gSourceFile.Data(n) <> vbKeyParenthesis Then GoTo addWordToList_
            modMain.ParseNextWordEx p + 1, gSourceFile.length, n, p, wbpFirstChar
            If gSourceFile.Data(n) = vbKeyParenthesis2 Then GoTo addWordToList_
            lFlags = lFlags Or 6
        End If
    Case vbKeyD
        ' Date = #1/1/2020# is a statement, not a method
        If sWord = "Date" Then lFlags = lFlags Or 8
    Case vbKeyT
        ' Time = #1:00:00 PM# is a statement, not a method
        If sWord = "Time" Then lFlags = lFlags Or 8
    End Select
    
    If (lFlags And 1) = 0 Then sProcessed = chrSlash & sName Else sProcessed = chrHash & sName
    If (lFlags And 4) = 0 Then
        If modMain.FindWord(m_ParseChecks, sProcessed, vbBinaryCompare, True) <> 0 Then
            GoTo exitRoutine
        End If
    End If
    
    ' check to see if this is a custom method (or declared item) vs a VBA function
    Dim lRecord&(0 To 0)
    lRecord(rteCRC) = modMain.CRCItem(sName, True)
    n = zlLocals: lID = pvFindWordEx(m_ZombieList.List(n), lRecord(), False)
    If lID = 0 Then
        For n = zlCpMethods To zlGblUDTs
            lID = pvFindWordEx(m_ZombieList.List(n), lRecord(), False)
            If lID <> 0 Then Exit For
        Next
    End If
    If lID <> 0 Then
        If (lFlags And 2) = 2 Then
            If n = zlCpMethods Or n = zlGblMethods Then
                gParsedItems.Find modMain.SetQuery(recID, qryIs, m_ZombieList.List(n).Item(rteRecID, lID)), , , 1&
                If (gParsedItems.Fields(recFlags).Value And iaLeftParams) = 0 Then
                    GoTo logDiscrep_
                End If
            End If
        End If
        m_ZombieList.List(n).Item(rteDiscreps, lID) = _
            m_ZombieList.List(n).Item(rteDiscreps, lID) And Not vtZombie
        If (lFlags And 8) = 8 Then
            If n = zlCpMethods Or n = zlGblMethods Then
                m_ZombieList.List(n).Item(rteDiscreps, lID) = _
                    m_ZombieList.List(n).Item(rteDiscreps, lID) Or vtMalicious
            End If
        End If
        GoTo exitRoutine
    End If
    
    ' if we got here, the target is not overridden, handle statements vs methods
    If (lFlags And 12) <> 0 Then
        If (lFlags And 2) = 0 Then
            If lStmt = lStart Then
                lFlags = lFlags Or 16
            ElseIf sPrevWord = "Then" Then
                lFlags = lFlags Or 16
            ElseIf sPrevWord = "Else" Then
                lFlags = lFlags Or 16
            End If
            If (lFlags And 16) = 16 Then
                If (m_ValidationChecks And vtMalicious) = 0 Then GoTo exitRoutine
                If AscW(sWord) = vbKeyE Then GoTo exitRoutine
                Mid$(sProcessed, 1, 1) = "!"
            Else
                Mid$(sProcessed, 1, 1) = chrColon
            End If
        End If
        If modMain.FindWord(m_ParseChecks, sProcessed, vbBinaryCompare, True) <> 0 Then
            GoTo exitRoutine
        End If
    End If
    
logDiscrep_:
    If (lFlags And 17) = 0 Then              ' if not $ suffix, then...
        If lIndex = 0 Then lIndex = modMain.FindWord(m_VarFuncs, sWord, vbTextCompare, False)
        lFlags = vtVarFunc
    ElseIf (lFlags And 16) = 0 Then
        GoTo exitRoutine
    Else
        If (m_ValidationChecks And vtExcludeDateTime) <> 0 Then
            If (lFlags And 8) <> 0 Then GoTo exitRoutine
        End If
        If lIndex = 0 Then
            If (lFlags And 4) = 0 Then
                lIndex = modMain.FindWord(m_Malicious, sWord, vbTextCompare, False)
            Else
                lIndex = modMain.FindWord(m_VarFuncs, sWord, vbTextCompare, False)
            End If
        End If
        lFlags = vtMalicious
    End If
    modMain.CreateRecord rs.Fields(recID).Value, sWord, itDiscrep, 0, 0, , , , lFlags, , , scpLocal, , lIndex
    Exit Sub
    
addWordToList_:
    modMain.FindWord m_WordList, sName, vbBinaryCompare, True
exitRoutine:

End Sub

Private Sub pvLoadShadowMethods(rs As ADODB.Recordset)

    ' many VB code pages have shadow methods, like hWnd, Left, Width, etc
    ' these are loaded when zombie checks are being processed

    Dim n&, sClass$
    
    n = (rs.Fields(recFlags).Value And iaMaskCodePage)
    If m_ZombieList.List(zlCpShadow).pReserved = n Then Exit Sub
    
    With m_ZombieList.List(zlCpShadow)
        .pReserved = (rs.Fields(recFlags).Value And iaMaskCodePage)
        Erase .Item(): .Count = 0
        Select Case .pReserved
            Case iaForm, iaMDI, iaUC, iaPPG, iaUserDoc
                sClass = rs.Fields(recAttr).Value
        End Select
    End With
    If LenB(sClass) = 0 Then Exit Sub
    
    Dim lRecord&(rteCRC To rteDiscreps)
    Dim sValue$, p&, lCdPg&
    
    On Error GoTo exitRoutine
    n = m_TLB.GetTypeInfoNumber(Mid$(sClass, InStr(sClass, chrDot) + 1))
    p = m_TLB.GetTypeInfo(n).DefaultInterface.TypeInfoNumber
    
    lCdPg = modMain.CRCItem("Me.", True)
    With m_TLB.GetTypeInfo(p).Members  ' for each, load/log
        For n = 1 To .Count
            sValue = .Item(n).Name
            lRecord(rteCRC) = modMain.CRCItem(sValue, True)
            pvFindWordEx m_ZombieList.List(zlCpShadow), lRecord(), True
            lRecord(rteCRC) = modMain.CRCItem(sValue, True, lCdPg)
            pvFindWordEx m_ZombieList.List(zlCpShadow), lRecord(), True
        Next
    End With
    
exitRoutine:
    On Error GoTo 0

End Sub

Private Sub pvResolveDupeDecs(rs As ADODB.Recordset)

    Dim lRecord&(rteCRC To rteDiscreps)
    Dim lID&, n&, p&, lCodePg&, wl As ZombieListStruct
    Dim rsMbrs As ADODB.Recordset, lLooper&
    
    Const DUPES = "(DupeDecs)"
    
    rs.Filter = modMain.SetQuery(recType, qryIs, itCodePage, qryAnd, recFlags, qryGT, -1)
    If rs.EOF = True Then GoTo exitRoutine
    
    ReDim wl.List(0 To 11)
    
    ' load global declares and log any duplication among globals
    Set rsMbrs = gParsedItems.Clone
    rsMbrs.Filter = modMain.SetQuery(recScope, qryIs, scpGlobal)
    rsMbrs.Sort = recType & chrComma & recCodePg
    
    Do Until rsMbrs.EOF = True
        lID = 0
        If m_DupeExclusions(0).Count <> 0 Then
            ' don't load public items from excluded code pages
            lRecord(0) = rsMbrs.Fields(recCodePg).Value
            lID = pvFindWordEx(m_DupeExclusions(0), lRecord(), False)
        End If
        If lID = 0 Then
            lRecord(rteCRC) = modMain.CRCItem(rsMbrs.Fields(recName).Value, True)
            lRecord(rteParent) = rsMbrs.Fields(recCodePg).Value
            lRecord(rteRecID) = rsMbrs.Fields(recID).Value
            Select Case rsMbrs.Fields(recType).Value
            Case itAPI: p = 0
            Case itEnum: p = 1
            Case itEnumMember: p = 2
            Case itConstant: p = 3
            Case itVariable: p = 4
            Case itType: p = 5
            Case Else: p = -1
                rsMbrs.Find modMain.SetQuery(recType, qryNot, rsMbrs.Fields(recType).Value)
            End Select
            If p <> -1 Then
                For n = 0 To 5
                    lID = pvFindWordEx(wl.List(n), lRecord(), False)
                    If lID <> 0 Then
                        If lCodePg <> rsMbrs.Fields(recCodePg).Value Then
                            lCodePg = rsMbrs.Fields(recCodePg).Value
                            gParsedItems.Find modMain.SetQuery(recID, qryIs, lCodePg), , , 1&
                            gSourceFile.CPBookMark = gParsedItems.Bookmark
                        End If
                        modMain.CreateRecord lRecord(rteRecID), DUPES, _
                            itDiscrep, 2, 0, , , , vtDupeDecs, wl.List(n).Item(rteRecID, lID)
                        Exit For
                    End If
                Next
                pvFindWordEx wl.List(p), lRecord(), True
                rsMbrs.MoveNext
            End If
        Else
            rsMbrs.MoveNext
        End If
    Loop
    
    Do
        If lLooper = 0 Then
            If (rs.Fields(recFlags).Value And iaMaskCodePage) = iaBAS Then lID = 0 Else lID = 1
        Else
            If (rs.Fields(recFlags).Value And iaMaskCodePage) <> iaBAS Then lID = 0 Else lID = 1
        End If
        If lID = 0 Then
            ' load module/class level constants, log any duplication among them or globals
            If m_DupeExclusions(0).Count <> 0 Then
                ' don't load items from excluded code pages
                lRecord(0) = rs.Fields(recID).Value
                lID = pvFindWordEx(m_DupeExclusions(0), lRecord(), False)
            End If
            If lID = 0 Then
                If lCodePg <> rs.Fields(recID).Value Then
                    lCodePg = rs.Fields(recID).Value
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, lCodePg), , , 1&
                    gSourceFile.CPBookMark = gParsedItems.Bookmark
                End If
                rsMbrs.Filter = chrParentO & modMain.SetQuery(recParent, qryIs, lCodePg) & ") Or (" & _
                                modMain.SetQuery(recCodePg, qryIs, lCodePg, qryAnd, recType, qryIs, itEnumMember) & chrParentC
                For n = 6 To 11
                    Erase wl.List(n).Item()
                    wl.List(n).Count = 0
                Next
                Do Until rsMbrs.EOF = True
                    lRecord(rteCRC) = modMain.CRCItem(rsMbrs.Fields(recName).Value, True)
                    lRecord(rteParent) = rsMbrs.Fields(recCodePg).Value
                    lRecord(rteRecID) = rsMbrs.Fields(recID).Value
                    Select Case rsMbrs.Fields(recType).Value
                    Case itAPI: p = 6
                    Case itEnum: p = 7
                    Case itEnumMember: p = 8
                    Case itConstant: p = 9
                    Case itVariable: p = 10
                    Case itType: p = 11
                    Case itMethod: p = 99
                    Case Else: p = -1
                        rsMbrs.Find modMain.SetQuery(recType, qryNot, rsMbrs.Fields(recType).Value)
                    End Select
                    If p <> -1 Then
                        For n = 6 To 11
                            lID = pvFindWordEx(wl.List(n), lRecord(), False)
                            If lID <> 0 Then
                                modMain.CreateRecord lRecord(rteRecID), DUPES, _
                                    itDiscrep, 1, 0, , , , vtDupeDecs, wl.List(n).Item(rteRecID, lID)
                                Exit For
                            End If
                        Next
                        If p <> 99 Then
                            If lID = 0 Then
                                For n = 0 To 5
                                    lID = pvFindWordEx(wl.List(n), lRecord(), False)
                                    If lID <> 0 Then
                                        If lRecord(rteRecID) <> wl.List(n).Item(rteRecID, lID) Then
                                            modMain.CreateRecord lRecord(rteRecID), DUPES, _
                                                itDiscrep, 2, 0, , , , vtDupeDecs, wl.List(n).Item(rteRecID, lID)
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            pvFindWordEx wl.List(p), lRecord(), True
                            If (m_ValidationChecks And vtDupeDecsNoScope) <> 0 Then
                                If p <> 10 Then pvFindWordEx wl.List(p - 6), lRecord(), True
                            End If
                        End If
                        rsMbrs.MoveNext
                    End If
                Loop
                rsMbrs.Filter = modMain.SetQuery(recCodePg, qryIs, lCodePg, _
                                                qryAnd, recParent, qryNot, lCodePg)
                n = 0
                Do Until rsMbrs.EOF = True
                    Select Case rsMbrs.Fields(recType).Value
                    Case itConstant: p = 9: n = n Or 1
                    Case itVariable: p = 10: n = n Or 2
                    Case Else
                        If n = 3 Then Exit Do Else p = -1
                        rsMbrs.Find modMain.SetQuery(recType, qryNot, rsMbrs.Fields(recType).Value)
                    End Select
                    If p <> -1 Then
                        lRecord(rteCRC) = modMain.CRCItem(rsMbrs.Fields(recName).Value, True)
                        lID = pvFindWordEx(wl.List(p), lRecord(), False)
                        If lID <> 0 Then
                            modMain.CreateRecord rsMbrs.Fields(recID).Value, DUPES, _
                                itDiscrep, 0, 0, , , , vtDupeDecs, wl.List(p).Item(rteRecID, lID)
                        End If
                        rsMbrs.MoveNext
                    End If
                Loop
            End If
        End If
        rs.MoveNext
        If rs.EOF = True Then
            If lLooper = 1 Then Exit Do
            lLooper = 1: rs.MoveFirst
        End If
    Loop

exitRoutine:
    Erase wl.List()
    rsMbrs.Close: Set rsMbrs = Nothing

End Sub

Private Sub pvResolveDupeLiterals(rs As ADODB.Recordset)

    Dim lRecord&(rteCRC To rteDiscreps)
    Dim lMin&, lID&, n&, lCodePg&, lLooper&
    
    rs.Filter = modMain.SetQuery(recType, qryIs, itCodePage, qryAnd, recFlags, qryGT, -1)
    If rs.EOF = True Then GoTo exitRoutine
    
    ReDim m_Literals.List(0 To 1)
    lMin = (m_ValidationChecks And vtLitMinCountMask) \ vtLitMinCountShift
    
    ' load global constants and log any duplication among global const values
    m_rsLiterals.Filter = modMain.SetQuery(recFlags, qryIs, 2)
    If m_rsLiterals.EOF = False Then
        Do
            lID = 0
            If m_DupeExclusions(1).Count <> 0 Then
                ' don't load public constant values from excluded code pages
                lRecord(0) = m_rsLiterals.Fields(recCodePg).Value
                lID = pvFindWordEx(m_DupeExclusions(1), lRecord(), False)
            End If
            If lID = 0 Then
                lRecord(rteCRC) = m_rsLiterals.Fields(recGrp).Value
                lRecord(rteParent) = m_rsLiterals.Fields(recParent).Value
                lID = pvFindWordEx(m_Literals.List(0), lRecord(), True)
                If lID <> 0 Then
                    If lCodePg <> m_rsLiterals.Fields(recCodePg).Value Then
                        lCodePg = m_rsLiterals.Fields(recCodePg).Value
                        gParsedItems.Find modMain.SetQuery(recID, qryIs, lCodePg), , , 1&
                        gSourceFile.CPBookMark = gParsedItems.Bookmark
                    End If
                    modMain.CreateRecord lRecord(rteParent), m_rsLiterals.Fields(recName).Value, _
                        itDiscrep, 2, 0, , , lRecord(rteCRC), vtDupeLiterals, m_Literals.List(0).Item(rteParent, lID)
                End If
            End If
            m_rsLiterals.MoveNext
        Loop Until m_rsLiterals.EOF = True
    End If
    
    Do
        If lLooper = 0 Then
            If (rs.Fields(recFlags).Value And iaMaskCodePage) = iaBAS Then lID = 0 Else lID = 1
        Else
            If (rs.Fields(recFlags).Value And iaMaskCodePage) <> iaBAS Then lID = 0 Else lID = 1
        End If
        If lID = 0 Then
            ' load module/class level constants, log any duplication among them or globals
            If m_DupeExclusions(1).Count <> 0 Then
                ' don't load items from excluded code pages
                lRecord(0) = rs.Fields(recID).Value
                lID = pvFindWordEx(m_DupeExclusions(1), lRecord(), False)
            End If
            If lID = 0 Then
                m_rsLiterals.Filter = modMain.SetQuery(recCodePg, qryIs, rs.Fields(recID).Value, _
                                                        qryAnd, recFlags, qryIs, 1)
                If lCodePg <> rs.Fields(recID).Value Then
                    lCodePg = rs.Fields(recID).Value
                    gParsedItems.Find modMain.SetQuery(recID, qryIs, lCodePg), , , 1&
                    gSourceFile.CPBookMark = gParsedItems.Bookmark
                End If
                Do Until m_rsLiterals.EOF = True
                    lRecord(rteCRC) = m_rsLiterals.Fields(recGrp).Value
                    lRecord(rteParent) = m_rsLiterals.Fields(recParent).Value
                    n = 1: lID = pvFindWordEx(m_Literals.List(n), lRecord(), True)
                    If lID = 0 Then
                        n = 0: lID = pvFindWordEx(m_Literals.List(0), lRecord(), False)
                    End If
                    If lID = 0 Then
                        If (m_ValidationChecks And vtDupeLitsNoScope) <> 0 Then
                            pvFindWordEx m_Literals.List(0), lRecord(), True
                        End If
                    Else
                        modMain.CreateRecord lRecord(rteParent), m_rsLiterals.Fields(recName).Value, _
                            itDiscrep, 2, 0, , , lRecord(rteCRC), vtDupeLiterals, m_Literals.List(n).Item(rteParent, lID)
                    End If
                    m_rsLiterals.MoveNext
                Loop
            
                ' load all literals parsed from method blocks, log any that exceed minimal
                ' duplication count or that duplicate module/global constants
                m_rsLiterals.Filter = modMain.SetQuery(recCodePg, qryIs, rs.Fields(recID).Value, _
                                                        qryAnd, recFlags, qryIs, 0)
                If m_rsLiterals.EOF = False Then
                    Do
                        lRecord(rteCRC) = m_rsLiterals.Fields(recGrp).Value
                        lRecord(rteParent) = m_rsLiterals.Fields(recParent).Value
                        For n = 1 To 0 Step -1
                            lID = pvFindWordEx(m_Literals.List(n), lRecord(), False)
                            If lID <> 0 Then
                                modMain.CreateRecord lRecord(rteParent), m_rsLiterals.Fields(recName).Value, _
                                    itDiscrep, 1, 0, , , lRecord(rteCRC), vtDupeLiterals, m_Literals.List(n).Item(rteParent, lID), m_rsLiterals.Fields(recOffset2).Value
                                m_rsLiterals.Fields(recOffset2).Value = 1
                                m_rsLiterals.Update
                                Exit For
                            End If
                        Next
                        m_rsLiterals.MoveNext
                    Loop Until m_rsLiterals.EOF = True
                    m_rsLiterals.MoveFirst
                    Do
                        If m_rsLiterals.Fields(recOffset2).Value > lMin Or _
                            AscW(m_rsLiterals.Fields(recName)) = vbKeyHash Then
                            With m_rsLiterals
                                modMain.CreateRecord .Fields(recParent).Value, .Fields(recName).Value, _
                                itDiscrep, 0, 0, , , .Fields(recGrp).Value, vtDupeLiterals, , .Fields(recOffset2).Value
                            End With
                        End If
                        m_rsLiterals.MoveNext
                    Loop Until m_rsLiterals.EOF = True
                End If
                Erase m_Literals.List(1).Item
                m_Literals.List(1).Count = 0
            End If
        End If
        rs.MoveNext
        If rs.EOF = True Then
            If lLooper = 1 Then Exit Do
            lLooper = 1: rs.MoveFirst
        End If
    Loop

exitRoutine:
    Erase m_Literals.List()
    m_rsLiterals.Close: Set m_rsLiterals = Nothing

End Sub

Private Sub pvAddLiteralToList(rs As ADODB.Recordset, sName As String)

    ' this routine not called if duplicate literals option not selected
    
    If Len(sName) > ((m_ValidationChecks And vtLitMinSizeMask) \ vtLitMinSizeShift) - 1 Then
        Exit Sub
    End If
    
    Dim lRecord&(rteCRC To rteDiscreps)
    Dim bAllowDupes As Boolean, lID&
    
    lRecord(rteParent) = rs.Fields(recID).Value
    lRecord(rteCRC) = modMain.CRCItem(sName, True)
    lRecord(rteRecID) = m_Literals.Names.Count + 1
    lRecord(rteDiscreps) = 1
    If rs.Fields(recType).Value = itConstant Then
        If rs.Fields(recParent).Value = m_Literals.List(0).pReserved Then
            lRecord(rteType) = 1
            If rs.Fields(recScope).Value = scpGlobal Then
                lRecord(rteType) = lRecord(rteType) Or 2
            End If
            bAllowDupes = True
        End If
    End If

    lID = pvFindWordEx(m_Literals.List(0), lRecord(), True, , bAllowDupes)
    If lID = 0 Or bAllowDupes = True Then
        With m_Literals.Names
            .Count = .Count + 1
            If .Count = 1 Then
                ReDim .fName(1 To 10)
            ElseIf .Count > UBound(.fName) Then
                ReDim Preserve .fName(1 To .Count + 10)
            End If
            .fName(.Count) = sName
        End With
    Else
        m_Literals.List(0).Item(rteDiscreps, lID) = _
            m_Literals.List(0).Item(rteDiscreps, lID) + 1
    End If

End Sub

Private Sub pvUpdateLiteralsRS()

    ' track/update count of a hardcoded literal
    
    Dim n&
    With m_Literals.List(0)
        For n = 1 To .Count
            m_rsLiterals.AddNew Array(0, 1, 2, 3, 4, 5), _
                Array(m_Literals.Names.fName(.Item(rteRecID, n)), _
                .Item(rteCRC, n), .Item(rteParent, n), _
                m_Literals.List(0).pReserved, .Item(rteType, n), .Item(rteDiscreps, n))
            m_rsLiterals.Update
        Next
    End With
    m_Literals.List(0).Count = 0
    Erase m_Literals.List(0).Item()
    Erase m_Literals.Names.fName()
    m_Literals.Names.Count = 0

End Sub

Private Sub pvCreateLiteralsRS()

    ' temp recordset used to track duplicate literals

    Set m_rsLiterals = New ADODB.Recordset
    With m_rsLiterals
        .Fields.Append recName, adVarWChar, 1023
        .Fields.Append recGrp, adInteger
        .Fields.Append recParent, adInteger
        .Fields.Append recCodePg, adInteger
        .Fields.Append recFlags, adInteger
        .Fields.Append recOffset2, adInteger
        .Open
    End With
    ReDim m_Literals.List(0 To 0)
    m_Literals.Count = 1
    
End Sub
