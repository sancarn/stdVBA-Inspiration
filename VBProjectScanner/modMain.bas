Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function PathGetArgsW Lib "shlwapi.dll" (ByVal pszPath As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal length As Long)
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Public Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesExW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal fInfoLevelId As Long, ByRef lpFileInformation As Any) As Long
Private Declare Function GetFullPathNameW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long

Public Type SourceDataStruct    ' used when parsing any source file
    Text As String              ' source file, converted from DBCS if applies
    Data() As Integer           ' array overlayed onto text
    First As Long               ' typically, start of declarations section
    length As Long              ' size of array, less one
    aOverlay(0 To 5) As Long    ' Faux SafeArray structure
    CPBookMark As Variant       ' gParsedItems recordset bookmark for overall source file
    ItemBookMark As Variant     ' gParsedItems recordset bookmark for current record
    ProjBookMark As Variant     ' gParsedItems recordset bookmark for overall project
    Owner As clsSourceFile      ' reference to class doing the parsing
End Type
Public Type WordListStruct      ' used during validation checks
    Count As Long
    List() As String
    pReserved As Long
End Type

Public Enum ValidationTypeEnum
    ' 0x0000FFFF    used for primary checks
    ' 0x000F0000    used for required length of literal to report as duplicate
    ' 0x00F00000    used for required duplication count of literals
    ' 0xFF000000    used for sub-option checks & exclusions
    vtVarType = 1
    vtZombie = 2
    vtEmptyCode = 4
    vtStopEnd = 8
    vtWithNoEvents = &H10
    vtMalicious = &H20
    vtRedim = &H40
    vtVarFunc = &H100
    vtDupeDecs = &H200
    vtDupeLiterals = &H400
    vtLitMinSizeShift = &H10000
    vtLitMinSizeMask = &HF0000
    vtLitMinCountMask = &HF00000
    vtLitMinCountShift = &H100000
    vtDupeDecsNoScope = &H1000000
    vtDupeLitsNoScope = &H2000000
    vtExcludeEMbrZombies = &H20000000
    vtRegReadDLLs = &H40000000
    vtExcludeDateTime = &H80000000
End Enum

Public Enum StatementParseFlags
    fCompIf = 1                 ' compiler directive #If found
    fCompElseIf = 2             ' compiler directive #ElseIf found
    fCompElse = 3               ' compiler directive #Else found
    fCompEnd = 4                ' compiler directive #End If found
    fCompConst = 5              ' compiler directive #Const found
    fCompMask = 7               ' mask for above flags
    fCompIgnore = 8             ' directive prevents processing contained statements
    fLiteralStr = &H10000       ' parsing string literal
    fComment = &H20000          ' parsing comment/remark
    fLiteralDt = &H40000        ' parsing date literal
    fLabel = &H80000            ' processing a label
    fContd = &H80000000         ' processing continuation character
End Enum
Public Enum WordBreakOptEnum    ' see ParseNextWordEx
    wbpParentsCntr = 1          ' paired parentheses treated as single word
    wbpParentsEOW = 2           ' open parenthesis treated as wordbreak
    wbpNotCommaEOW = 4          ' comma not a wordbreak else it is
    wbpSepNamedParams = 8       ' split named parameters, return param value only
    wbpFirstChar = 16           ' stop at first character of next word
    wbpContinuation = 32        ' all characters up to line continuation
    wbpEqualEOW = 64            ' equal sign treated as wordbreak
    wbpNotSemcolonEOW = 128     ' semicolon not a wordbreak else it is
    wbpLineSep = 256
    wbpMarkBrackets = 512
    wbpBracketsMarked = 1024
End Enum
Public Enum ItemTypeEnum
    itMethod = 1                ' function,sub,property
    itAPI = 2                   ' API
    itEvent = 3                 ' Public Event()
    itClassEvent = 4            ' non-BAS event, i.e., Click, Terminate, etc
    itEnum = 5                  ' Enum
    itEnumMember = 6            ' Enum member (during validation only)
    itVariable = 7              ' variable
    itConstant = 8              ' constant
    itControl = 9               ' control
    itType = 10                 ' Type/UDT
    itTypeMember = 11           ' Type member (during validation only)
    itParameter = 12            ' Method/API/Event parameter (during validation only)
    itImplements = 13           ' Implements statement
    itProject = &H20            ' project overall
    itSourceFile = &H21         ' source file: .cls, .frm, etc
    itCodePage = &H22           ' parent record for source file code entries
    itDefType = &H23            ' Def[xxx] statement
    itHelpFile = &H30           ' help file usage
    itMiscFile = &H31           ' included documents
    itResFile = &H32            ' res file usage
    itStats = &H40              ' parsing stats
    itDiscrep = &H1000          ' used primarily for dupe decs/literals & malicious code checks
    itValidation = &H2000       ' excluded code pages during validation & options
    itReference = &H10000000    ' external reference
End Enum
Public Enum ItemTypeAttrEnum
    '0x0000001F                 code file categories (itCodeFile)
    '0x000003E0                 code file properties (itCodeFile)
    '0x0F000000                 code file option statments (itCodeFile)
    '0x80000000                 external code file (itCodeFile)
    iaBAS = 1                   ' source file is BAS module else "class module"
    iaForm = 2                  ' source file: form
    iaMDI = 3                   ' source file: MDI form
    iaClass = 4                 ' source file: class
    iaUC = 5                    ' source file: usercontrol
    iaPPG = 6                   ' source file: property page
    iaUserDoc = 7               ' source file: user document
    iaDesigner = 8              ' source file: designer
    iaUnkSource = 9             ' source file: unknown
    iaMaskCodePage = &H1F       ' mask for code page attributes
    iaMdiChild = &H20           ' form is a MDI child
    iaDBCS = &H40               ' source file is non-ANSI
    iaExposed = &H80            ' source file exposed to user else private to project
    iaPredeclared = &H100       ' class is predeclared
    iaGblNameSpace = &H200      ' global namespace
    iaFileError = &H400         ' file access error
    iaOpExplicit = &H8000000    ' Option statement type
    iaOpText = &H4000000        ' Option statement type
    iaOpBase1 = &H2000000       ' Option statement type
    iaOpPrivate = &H1000000     ' Option statement type
    iaMaskOptions = &HF000000   ' mask for above options
    iaExternalProj = &H80000000 ' external project reference
    
    '0x8000000F                 variable/enum related attrs (itVariable/itEnumMember)
    iaBrackets = 1              ' item (enum member) defined with [ ]
    iaHidden = 2                ' designer hardcoded/hidden "WithEvents" variables
    iaWithEvents = &H80000000   ' variable uses WithEvents
    
    '0x00001000                 statement related attrs
    iaSplitIdent = &H1000       ' statement has object identifier split among lines
    
    '0xFFFE0000                 method/event related attrs (itMethod,itEvent,itClassEvent)
    iaUnresolved = &H400000     ' events were not resolved
    iaImplemented = &H200000    ' class & its public methods are implemented
    iaLeftParams = &H100000     ' custom Left property has parameters
    iaSub = &H10000000          ' method type: sub
    iaFunction = &H20000000     ' method type: function
    iaProperty = &H80000000     ' method type: property
    iaPropGet = &H81000000      ' method type: Property Get
    iaPropLet = &H82000000      ' method type: Property Let
    iaPropSet = &H84000000      ' method type: Property Set
    
    '0x0000000F                 stats related attrs (itStats)
    iaStatements = 1: iaExclusions = 2: iaComments = 3
    
    '0x000000FF                 custom use for designers only, not cached in recordset
    iaWebClass = 16: iaDataEnv = 32
End Enum
Public Enum ItemScopeEnum
    scpGlobal = -1&             ' project level, includes all Public Enum & source files
    scpLocal = 0                ' method level & params
    scpPrivate = 1              ' module level
    scpFriend = 2               ' module level
    scpPublic = 3               ' module level
End Enum
Public Enum ValidationConstants
    vnFileNotFound = 1          ' file error
    vnFileEmpty = 2             ' file error (0 bytes)
    vnFileTooBig = 3            ' file error ( >2GB )
    vnFileInvalid = 4           ' file error (unexpected content)
    vnFileAccess = 5            ' file error (file access error)
    vnAborted = 6               ' user aborted processing
    vnOpenBracket = 10
    vnOpenQuote = 11
    vnOpenDate = 12
    vnOpenMethodBlk = 13
    vnOpenEnumUDT = 14
    vnOpenCompIF = 15
End Enum
Public Enum QueryOperandConstants ' used in SetQuery
    qryIs = 0: qryNot = 1
    qryGT = 2: qryGTE = 3: qryLT = 4: qryLike = 5
    qryAnd = 6: qryOr = 7
End Enum

Public Const recParent = "fParent", recCodePg = "fCodePage"
Public Const recID = "fID", recName = "fName", recType = "fType"
Public Const recAttr = "fAttr", recAttr2 = "fAttr2"
Public Const recGrp = "fGroup", recFlags = "fFlags"
Public Const recOffset = "fOffset", recOffset2 = "fOffset2"
Public Const recStart = "fSoS", recEnd = "fEoS"
Public Const recScope = "fScope", recDiscrep = "fDiscrep"
Public Const recIdxName = "fSort1", recIdxAttr = "fSort2"

Public Const chrApos = "'", chrParentO = "(", chrParentC = ")"
Public Const chrAsterisk = "*", chrColon = ":", chrSemi = ";"
Public Const chrDot = ".", chrComma = ",", chrSlash = "\"
Public Const chrHash = "#", chrPct = "%"
Public Const chrV = "V", chrZ = "Z", chrW = "W"
Public Const chrR = "R", chrM = "M", chrI = "i"

Public Const vbKeyQuote = 34, vbKeyDot = 46, vbKeyBang = 33
Public Const vbKeyRemark = 39, vbKeyBracket = 91, vbKeyBracket2 = 93
Public Const vbKeyHash = 35, vbKeyUnderscore = 95, vbKeyColon = 58, vbKeySlash = 47
Public Const vbKeyLineFeed = 10, vbKeyParenthesis = 40, vbKeyParenthesis2 = 41
Public Const vbKeyComma = 44, vbKeyEqual = 61, vbKeySemicolon = 59

Public Const WM_USER = &H400

Public Const ParseTypeCnt = "Const", ParseTypeStat = "Static"
Const ParseTypeSub = "Sub", ParseTypeFnc = "Function", ParseTypePpy = "Property"
Const ParseTypePub = "Public", ParseTypePriv = "Private"
Const ParseTypeFrnd = "Friend"
Const ParseTypeTyp = "Type", ParseTypeEnm = "Enum", ParseTypeAtt = "Attribute"
Const ParsePropBegin = "BeginProperty", ParsePropEnd = "EndProperty"
Const ParseObject = "Object", ParseName = "Name"
Const chrClass = "Class", chrClassVB = "VB.Class", chrPctWild = "_%"


Public gParsedItems As ADODB.Recordset  ' see GlobalsInitialize & CreateRecord
Public gSourceFile As SourceDataStruct  ' see SetSourceData

Dim m_CompDirs As clsCompDir            ' compiler directive processing
Dim m_RefsEvents As clsReferences       ' project references & class events
Dim m_CRC32LUT&(0 To 255)               ' CRC32 lookup table
Dim m_Offsets&()                        ' method internal statement offsets, 2 entries per offset
Dim m_IdxOffset&                        ' current index into m_Offsets
Dim m_RecordID&                         ' incrementing unique ID

Public Sub GlobalsInitialize()
    ' called by clsProject when about to start parsing new project
    Set m_CompDirs = New clsCompDir
    Set m_RefsEvents = New clsReferences
    ReDim m_Offsets(0 To 999)
    pvCreateCRC32LUT
    m_RecordID = 0
    
    ' Jump to end of module for complete descripton of how field values are used
    ' record hierarchy (parent/child relationships) looks like this:
    '    vbg file
    '    vbp file
    '        project-wide stuff (compiler constants, references, file list, etc)
    '        code files, aka modules (.frm, .bas, .Cls, etc)
    '            code page (similar to code file, different attrs/props)
    '                Def[xxx] statements
    '                APIs
    '                    params
    '                Constants, Variables(standard)
    '                Variables using WithEvents
    '                    Events
    '                Enums, Types
    '                    members (not displayed in tree)
    '                Implements
    '                    Events
    '                Class Events (n/a for bas-modules)
    '                Class/Module Methods (Sub,Function,Property)
    '                    params, constants, variables
    '                Public Events (used with RaiseEvent)
    '                    params
    '           *Option statements are a property value, not an individual record
    '   external uncompiled vbp projects (if applies)
    '       code files that are exposed (public/global in scope)
    '           code page (similar to code file, different attrs/props)
    Set gParsedItems = New ADODB.Recordset
    With gParsedItems
        .Fields.Append recID, adInteger             ' m_RecordID value
        .Fields.Append recCodePg, adInteger         ' code page/file grouping index
        .Fields.Append recParent, adInteger         ' child/parent relationship
        .Fields.Append recName, adVarWChar, 1023    ' parsed item's name
        .Fields.Append recAttr, adVarWChar, 1023    ' attribute for parsed item
        .Fields.Append recAttr2, adVarWChar, 1023   ' addtl attribute as needed
        .Fields.Append recType, adInteger           ' ItemTypeEnum values
        .Fields.Append recGrp, adInteger            ' as needed, typically CRC
        .Fields.Append recFlags, adInteger          ' ItemTypeAttrEnum values
        .Fields.Append recOffset, adInteger         ' per-item, statement offset
        .Fields.Append recOffset2, adInteger        ' per-item, statement offset
        .Fields.Append recStart, adInteger          ' typically statement start
        .Fields.Append recEnd, adInteger            ' typically statement end
        .Fields.Append recScope, adInteger          ' ItemScopeEnum values
        .Fields.Append recDiscrep, adVarChar, 25    ' list of:
            'Z = zombie check failed    E = empty method (no code)
            'V = vartype check failed   D = duplicate declaration
            'U = undeclared variable    R = array created without Dim
            'X,x = End/Stop check
        .Fields.Append recIdxName, adInteger        ' sorting index for recName
        .Fields.Append recIdxAttr, adInteger        ' sorting index for recAttr
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub

Public Sub GlobalsRelease()
    ' called by clsProject when it finishes parsing project code files
    Set m_CompDirs = Nothing
    Set m_RefsEvents = Nothing
    Erase m_Offsets()
    SetSourceData vbNullString, 0, 0&
End Sub

Public Function CreateRecord(pParent&, pName$, pType&, pStart&, pEnd&, _
                        Optional pAttr$, Optional pAttr2$, Optional pGroup&, _
                        Optional pFlags&, Optional pOffset&, Optional pOffset2&, _
                        Optional pScope&, Optional pDiscrep$, Optional pSortIdx&) As Long
                        
    ' helper function to log parsed items during initial scan & validation

    Dim pCodePg&
    With gParsedItems
        If IsEmpty(gSourceFile.CPBookMark) = False Then
            If IsEmpty(gSourceFile.ProjBookMark) = False Then
                .Bookmark = gSourceFile.CPBookMark
                pCodePg = .Fields(recCodePg).Value
            End If
        End If
        m_RecordID = m_RecordID + 1
        .AddNew Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), _
                Array(m_RecordID, pCodePg, pParent, pName, pAttr, pAttr2, _
                      pType, pGroup, pFlags, pOffset, pOffset2, _
                      pStart, pEnd, pScope, pDiscrep, pSortIdx, 0)
        .Update
        gSourceFile.ItemBookMark = .Bookmark
    End With
    CreateRecord = m_RecordID
    
End Function

Public Function SetSourceData(FileName As String, lParent As Long, errCode As ValidationConstants) As Long

    ' attempts to load the passed FileName and overlay an array on it

    With gSourceFile
        If .aOverlay(3) <> 0 Then       ' overlay exists, remove it
            CopyMemory ByVal VarPtrArray(.Data), 0&, 4
            .Text = vbNullString
            Erase .aOverlay()
        End If
        If LenB(FileName) = 0 Then      ' releasing source
            .length = 0: .First = 0
            .CPBookMark = Empty
        Else
            Dim aData() As Byte
            Dim lSize&, lRead&, hHandle&
                                        ' validate file exists
            FileName = ResolveRelativePath(FileName, vbNullString)
            errCode = GetFileLastModDate(FileName, 0&, 0&, lSize)
            If errCode <> 0 Then Exit Function
                                        ' attempt to access file
            hHandle = GetFileHandle(FileName, False)
            If hHandle = 0 Or hHandle = -1 Then
                errCode = vnFileAccess
                Exit Function
            End If
                                        ' read file into array
            ReDim aData(0 To lSize + 1)
            ReadFile hHandle, aData(0), lSize, lRead
            If lSize <> lRead Then
                errCode = vnFileAccess: CloseHandle hHandle
                Erase aData()
                Exit Function
            End If
            aData(lSize) = vbKeyReturn: aData(lSize + 1) = vbKeyLineFeed
                                        ' convert file as needed (handles DBCS)
            .Text = StrConv(aData(), vbUnicode)
            Erase aData()
            If lSize + 2 <> Len(.Text) Then lRead = 0
                                        ' overlay integer array
            .aOverlay(0) = &H10001: .aOverlay(1) = 2: .aOverlay(3) = StrPtr(.Text)
            .aOverlay(4) = Len(.Text): .aOverlay(5) = 1
            CopyMemory ByVal VarPtrArray(.Data), VarPtr(.aOverlay(0)), 4
            .First = 0: .length = .aOverlay(4)
                                        ' create itSourceFile record
            If lParent = &H80000000 Then ' flag = don't create a new record
                .CPBookMark = Empty
            Else
                CreateRecord lParent, vbNullString, itSourceFile, 0&, lSize, FileName
                If lRead = 0 Then           ' include DBCX flag as needed
                    gParsedItems.Fields(recFlags).Value = iaDBCS
                    gParsedItems.Update
                End If
                .CPBookMark = gParsedItems.Bookmark
            End If
            m_IdxOffset = 0             ' used in ScanSource routine
            SetSourceData = hHandle     ' return handle, caller is responsible for it
        End If
        .ItemBookMark = Empty
    End With

End Function

Public Function ScanHeader(IsExternal As Boolean) As Boolean

    ' Routine scans the header portion of a code file
    ' This routine is looking for key properties in the file along
    ' with controls and/or hidden designer variables

    Dim lMode&, lBlocks&, lStart&, n&, p&, cAttrs&
    Dim lMax&, lParent&, lCrc&, sClass$, sName$
    Dim rs As ADODB.Recordset, sProject$
    Dim bExposed As Boolean, bIsUC As Boolean
    
    On Error GoTo abortRoutine
    With gParsedItems       ' update record properties
        .Bookmark = gSourceFile.ProjBookMark
        sProject = .Fields(recName).Value
        bIsUC = (.Fields(recDiscrep).Value = "Control")
        .Bookmark = gSourceFile.CPBookMark
        .Fields(recAttr).Value = vbNullString
        If IsExternal = True Then .Fields(recFlags).Value = iaExternalProj
        .Fields(recType).Value = itCodePage
        .Fields(recCodePg).Value = .Fields(recID).Value
        .Update
    End With
    lParent = m_RecordID
    
    Do  ' parse header, move to first non-whitespace character
        ParseNextLine lMax + 1, lStart, lMax
        If lStart = lMax Then Exit Do
        p = lStart
        
        ' parse the header statement
        ParseNextWordEx p, lMax, n, p
        
        '/// Begin/BEGIN block?
        If gSourceFile.Data(n) = vbKeyB Then
            If p - n = 5 Then
                If LCase$(Mid$(gSourceFile.Text, n, p - n)) = "begin" Then
                    lBlocks = lBlocks + 1
                    If p = lMax Then                ' VB class
                        sClass = chrClassVB: cAttrs = 1 ' temp lib/class name
                    Else
                        ParseNextWordEx p, lMax, n, p
                        sClass = Mid$(gSourceFile.Text, n, p - n)
                        If IsExternal = False Then _
                            cAttrs = pvAddDesignerEvents(sClass) ' returns designer or not
                    End If
                    If lMode = 0 Then               ' 1st Block, cache class name
                        With gParsedItems
                            .Bookmark = gSourceFile.CPBookMark
                            .Fields(recAttr).Value = sClass
                            .Fields(recScope).Value = scpGlobal
                            If IsExternal = False Then .Fields(recDiscrep).Value = chrZ
                            .Update
                        End With
                        lMode = 1
                        If cAttrs = iaDataEnv Then  ' custom parsing for hidden variables
                            lMax = pvScanHeader_DataEnv(lParent, p, IsExternal)
                        ElseIf cAttrs = iaWebClass Then ' custom parsing for hidden variables
                            lMax = pvScanHeader_WebClass(lParent, p, IsExternal)
                        End If
                    ElseIf p < lMax And IsExternal = False Then ' control, parse name & log
                        If rs Is Nothing Then Set rs = gParsedItems.Clone
                        ParseNextWordEx p, lMax, n, p
                        sName = Mid$(gSourceFile.Text, n, p - n)
                        lCrc = CRCItem(sName, True)
                        rs.Filter = SetQuery(recGrp, qryIs, lCrc, _
                                            qryAnd, recParent, qryIs, lParent)
                        If rs.EOF = True Then       ' 1st instance
                            CreateRecord lParent, sName, itControl, 0, 0, sClass, , lCrc, , , , scpPrivate
                        Else                        ' 2nd+ instance
                            rs.Fields(recOffset).Value = rs.Fields(recOffset).Value + 1
                            rs.Update
                        End If
                    End If
                End If
            End If
            
        '/// End/END block?
        ElseIf gSourceFile.Data(n) = vbKeyE Then
            If p - n = 3 Then
                If LCase$(Mid$(gSourceFile.Text, n, p - n)) = "end" Then
                    lBlocks = lBlocks - 1
                End If
            End If
            
        '/// Attributes section?? Triggers end of header & start of declarations
        ElseIf lBlocks = 0 Then
            If gSourceFile.Data(n) = vbKeyA Then
                If p - n = 9 Then
                    If Mid$(gSourceFile.Text, n, p - n) = ParseTypeAtt Then
                        ParseNextWordEx p, lMax, n, p
                        gParsedItems.Bookmark = gSourceFile.CPBookMark
                        With gParsedItems
                            Select Case Mid$(gSourceFile.Text, n, p - n)
                            Case "VB_Name"
                                ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                                sName = sProject & chrDot & Mid$(gSourceFile.Text, n + 1, p - n - 2)
                                .Fields(recName).Value = sName
                                If lMode = 0 Then   ' modules don't have Begin/End blocks
                                    .Fields(recAttr).Value = "VB.Module"
                                    .Fields(recScope).Value = scpGlobal
                                    .Fields(recFlags).Value = .Fields(recFlags).Value Or iaBAS
                                Else
                                    Select Case .Fields(recAttr).Value
                                    Case chrClassVB: .Fields(recFlags).Value = .Fields(recFlags).Value Or iaClass
                                    Case "VB.Form": .Fields(recFlags).Value = .Fields(recFlags).Value Or iaForm
                                    Case "VB.MDIForm": .Fields(recFlags).Value = .Fields(recFlags).Value Or iaMDI
                                    Case "VB.UserControl": .Fields(recFlags).Value = .Fields(recFlags).Value Or iaUC
                                        If bIsUC = True Then .Fields(recDiscrep).Value = Replace(.Fields(recDiscrep).Value, chrZ, vbNullString)
                                    Case "VB.PropertyPage": .Fields(recFlags).Value = .Fields(recFlags).Value Or iaPPG
                                                            .Fields(recDiscrep).Value = vbNullString
                                    Case "VB.UserDocument": .Fields(recFlags).Value = .Fields(recFlags).Value Or iaUserDoc
                                    Case Else: .Fields(recFlags).Value = .Fields(recFlags).Value Or iaDesigner
                                                            .Fields(recDiscrep).Value = vbNullString
                                    End Select
                                End If
                                .Fields(recGrp).Value = CRCItem(LCase$(sName), True)
                                .Update
                                lMode = 2               ' flag, valid header section
                            Case "VB_Exposed"           ' exposed outside of project
                                ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                                bExposed = CBool(Mid$(gSourceFile.Text, n, p - n))
                            Case "VB_PredeclaredId"
                                .Fields(recFlags).Value = .Fields(recFlags).Value Or iaPredeclared
                                .Update
                            Case "VB_GlobalNameSpace"
                                ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                                If CBool(Mid$(gSourceFile.Text, n, p - n)) = True Then
                                    .Fields(recFlags).Value = .Fields(recFlags).Value Or iaGblNameSpace
                                    .Fields(recDiscrep).Value = Replace(.Fields(recDiscrep).Value, chrZ, vbNullString)
                                    .Update
                                End If
                            End Select
                        End With
                    End If
                End If
                
            '/// End of Attributes section, declarations start here
            ElseIf lMode = 2 Then
                GoTo exitRoutine
            
            '/// Parse "Object" statement
            ElseIf gSourceFile.Data(n) = vbKeyO And IsExternal = False Then
                If p - n = 6 Then
                    If lMode <> 2 And IsExternal = False Then
                        If Mid$(gSourceFile.Text, n, p - n) = ParseObject Then
                            ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                            m_RefsEvents.ParseAndLoad Mid$(gSourceFile.Text, n, lMax - 1), True
                        End If
                    End If
                End If
            End If
            
        '/// MDI child & MultiUse checks
        ElseIf gSourceFile.Data(n) = vbKeyM And IsExternal = False Then
            If p - n = 8 Then
                If Mid$(gSourceFile.Text, n, p - n) = "MultiUse" Then
                    ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                    If CBool(Mid$(gSourceFile.Text, n, p - n)) = True Then
                        gParsedItems.Bookmark = gSourceFile.CPBookMark
                        gParsedItems.Fields(recFlags).Value = gParsedItems.Fields(recFlags).Value Or iaMdiChild
                        gParsedItems.Update
                    End If
                ElseIf Mid$(gSourceFile.Text, n, p - n) = "MDIChild" Then
                    ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                    If CBool(Mid$(gSourceFile.Text, n, p - n)) = True Then
                        gParsedItems.Bookmark = gSourceFile.CPBookMark
                        gParsedItems.Fields(recFlags).Value = gParsedItems.Fields(recFlags).Value Or iaMdiChild
                        gParsedItems.Update
                    End If
                End If
            End If
            
        '/// VB.Class parsing for key properties to determine what class events are used
        ElseIf (cAttrs And 7) <> 0 Then
            If gSourceFile.Data(n) = vbKeyP Then        ' Persistable
                If p - n = 11 Then
                    If Mid$(gSourceFile.Text, n, p - n) = "Persistable" Then
                        ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                        If CBool(Mid$(gSourceFile.Text, n, p - n)) = True Then cAttrs = cAttrs Or 2
                    End If
                End If
            ElseIf gSourceFile.Data(n) = vbKeyD Then    ' DataSourceBehavior
                If p - n = 18 Then
                    If Mid$(gSourceFile.Text, n, p - n) = "DataSourceBehavior" Then
                        ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                        If CBool(Mid$(gSourceFile.Text, n, p - n)) = True Then cAttrs = cAttrs Or 4
                    End If
                End If
            End If
        End If
    Loop
    
exitRoutine:
    ' following check can only fail when parsing code files from an external project
    gParsedItems.Bookmark = gSourceFile.CPBookMark
    If IsExternal = False Or bExposed = True Then
        If (cAttrs And 7) <> 0 Then
            Select Case cAttrs
            Case 5: sClass = "VBRUN.IDataProviderClassEvt"
                ' class events: Initialize,Terminate,GetDataMember,DataConnection
            Case 3: sClass = "VBRUN.IPersistableClassEvt"
                ' class events: Initialize,Terminate,InitProperties,ReadProperties,WriteProperties
            Case 7: sClass = "VBRUN.IPersistableDataProviderClassEvt"
                ' call events: same as above + GetDataMember,DataConnection
            Case Else: sClass = "VBRUN.IClassModuleEvt"
                ' class events: Initialize,Terminate
            End Select
            gParsedItems.Fields(recAttr).Value = sClass
            gParsedItems.Update
        End If
        gSourceFile.First = lStart              ' start of declarations section
        ScanHeader = CBool(lMode = 2)
    Else    ' remove this record
        gParsedItems.Delete: gParsedItems.Update
        gSourceFile.CPBookMark = Empty
        ScanHeader = True
    End If
    
abortRoutine:
    If Err Then Err.Clear
    If Not rs Is Nothing Then
        rs.Close: Set rs = Nothing
    End If
    On Error GoTo 0
    
End Function

Private Function pvScanHeader_DataEnv(lParent As Long, lStart As Long, IsExternal As Boolean) As Long

    ' specifically used to parse out hidden DataEnvironment variables

    Dim lBlocks&, n&, p&
    Dim lMax&, sName$, nrVars&
    Dim cmdState&
    ' cmdState used for parsing DataEnvironment variables
    ' 0x01  connection being processed
    ' 0x02  recordset being processed
    ' 0x03  recordset Grouping = -1
    ' 0x08+ looking for EndProperty
    
    ' DataEnvironment has several hidden WitheEvents-like variables
    ' there are two types: Adodb.Connection & Adodb.Recordset
    
    ' The key sections look like the following.
    ' Recordset variables always have "rs" prefixed to its name
    ' -----------------------------------------
    '   NumConnections = 2                      << nr BeginProperty statements
    '   BeginProperty Connection1               << start of connection #1
    '      ConnectionName = "BaseConnection"    << variable name
    '      ...
    '   EndProperty
    ' -----------------------------------------
    '   NumRecordsets = 4                       << nr BeginProperty statements
    '   BeginProperty Recordset1                << start of recordset #1
    '      CommandName = "MasterList"           << variable name unless a grouping
    '      CommDispId = 1002                    << if -1 then not WithEvents
    '      RsDispId = 1018                      << if -1 then not WithEvents
    '      ...
    '      Grouping = -1                        << if -1 then variable name follows
    '      GroupingName = "MasterList_Grouping" << used only when Grouping = -1
    '      ...
    '   EndProperty
    ' -----------------------------------------
    
    lMax = lStart
    Do  ' parse header, move to first non-whitespace character
        For lStart = lMax + 1 To gSourceFile.length - 1
            If IsEndOfLine(gSourceFile.Data(lStart)) = 0 Then
                If IsWhiteSpace(gSourceFile.Data(lStart)) = 0 Then Exit For
            End If
        Next
        If lStart = gSourceFile.length Then Exit Do
        ' move to end of statement & RTrim statement
        For lMax = lStart + 1 To gSourceFile.length
            If IsEndOfLine(gSourceFile.Data(lMax)) = 1 Then
                For n = lMax - 1 To lStart + 1 Step -1
                    If IsWhiteSpace(gSourceFile.Data(n)) = 0 Then
                        lMax = n + 1: Exit For
                    End If
                Next
                Exit For
            End If
        Next
        p = lStart
        
        ' parse the header statement
        ParseNextWordEx p, lMax, n, p

        If gSourceFile.Data(n) = vbKeyB Then
            If p - n = 13 Then                      ' BeginProperty
                If Mid$(gSourceFile.Text, n, p - n) = ParsePropBegin Then
                    If lBlocks = 0 And nrVars = 0 Then
                        ' sanity check if NumConnections,NumRecordsets not found
                        lMax = lStart - 1: Exit Do
                    Else ' increment
                        lBlocks = lBlocks + 1
                    End If
                End If
            End If
        ElseIf lBlocks = 0 Then
            If gSourceFile.Data(n) = vbKeyN Then    ' NumConnections,NumRecordsets
                If p - n = 14 Then                  ' NumConnections
                    If Mid$(gSourceFile.Text, n, p - n) = "NumConnections" Then
                        ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                        nrVars = CLng(Mid$(gSourceFile.Text, n, p - n))
                        If nrVars <> 0 Then cmdState = 1
                    End If
                ElseIf p - n = 13 Then              ' NumRecordsets
                    If Mid$(gSourceFile.Text, n, p - n) = "NumRecordsets" Then
                        ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                        nrVars = CLng(Mid$(gSourceFile.Text, n, p - n))
                        If nrVars <> 0 Then cmdState = 2
                    End If
                End If
            ElseIf gSourceFile.Data(n) = vbKeyE Then
                If p - n = 3 Then                   ' End
                    If Mid$(gSourceFile.Text, n, p - n) = "End" Then
                        ' sanity check if NumConnections,NumRecordsets not found
                        lMax = lStart - 1: Exit Do
                    End If
                End If
            End If
        
        ElseIf cmdState > 8 Then    ' looking for the end of variable's block
            If gSourceFile.Data(n) = vbKeyE Then    ' EndProperty
                If p - n = 11 Then
                    If Mid$(gSourceFile.Text, n, p - n) = ParsePropEnd Then
                        lBlocks = lBlocks - 1       ' decrement
                        If lBlocks = 0 Then         ' done if last recordset
                            nrVars = nrVars - 1: cmdState = cmdState Xor 8
                            If nrVars = 0 And cmdState = 2 Then Exit Do
                        End If
                    End If
                End If
            End If
        ElseIf cmdState = 1 Then                    ' parsing connections
            If gSourceFile.Data(n) = vbKeyC Then
                If p - n = 14 Then                  ' ConnectionName
                    If Mid$(gSourceFile.Text, n, p - n) = "ConnectionName" Then
                        If IsExternal = False Then
                            ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                            sName = Mid$(gSourceFile.Text, n + 1, p - n - 2)
                            CreateRecord lParent, sName, itVariable, 0, 0, "ADODB.Connection", , , iaWithEvents Or iaHidden, , , scpPrivate
                        End If
                        cmdState = cmdState Or 8    ' skip to EndProperty
                    End If
                End If
            End If
        
        '/// these are checked for parsing recordset-related statements
        ElseIf gSourceFile.Data(n) = vbKeyC Then    ' CommandName,CommDispId
            If p - n = 11 Then                      ' CommandName
                If Mid$(gSourceFile.Text, n, p - n) = "CommandName" Then
                    ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                    ' name used if Grouping <> -1
                    sName = Mid$(gSourceFile.Text, n + 1, p - n - 2)
                End If
            ElseIf p - n = 10 Then                  ' CommDispId
                If Mid$(gSourceFile.Text, n, p - n) = "CommDispId" Then
                    ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                    If CLng(Mid$(gSourceFile.Text, n, p - n)) = -1 Then
                        ' variable not loaded if value is -1
                        cmdState = cmdState Or 8    ' skip to EndProperty
                    End If
                End If
            End If
        ElseIf gSourceFile.Data(n) = vbKeyR Then
            If p - n = 8 Then                       ' RsDispId
                If Mid$(gSourceFile.Text, n, p - n) = "RsDispId" Then
                    ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                    If CLng(Mid$(gSourceFile.Text, n, p - n)) = -1 Then
                        ' variable not loaded if value is -1
                        cmdState = cmdState Or 8    ' skip to EndProperty
                    End If
                End If
            End If
        ElseIf gSourceFile.Data(n) = vbKeyG Then    ' Grouping,GroupingName
            If p - n = 8 Then                       ' Grouping
                If Mid$(gSourceFile.Text, n, p - n) = "Grouping" Then
                    ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                    ' if value = -1, then GroupingName used as variable's name
                    If CLng(Mid$(gSourceFile.Text, n, p - n)) = -1 Then cmdState = 3
                End If
            ElseIf p - n = 12 Then                  ' GroupingName
                If Mid$(gSourceFile.Text, n, p - n) = "GroupingName" Then
                    If IsExternal = False Then
                        If cmdState = 3 Then        ' get variable's name
                            ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                            sName = Mid$(gSourceFile.Text, n + 1, p - n - 2)
                        End If
                        CreateRecord lParent, "rs" & sName, itVariable, 0, 0, "ADODB.Recordset", , , iaWithEvents Or iaHidden, , , scpPrivate
                    End If
                    cmdState = 10                   ' skip to EndProperty
                End If
            End If
        End If
    Loop
    pvScanHeader_DataEnv = lMax

End Function

Private Function pvScanHeader_WebClass(lParent As Long, lStart As Long, IsExternal As Boolean) As Long

    ' specifically used to parse out hidden WebClass variables

    Dim lBlocks&, n&, p&
    Dim lMax&, sName$, nrVars&
    Dim cmdState&
    ' cmdState used for parsing WebClass variables
    ' 0x01  WebItems being processed, looking for count
    ' 0x02  WebItem being processed, looking for its Name
    ' 0x08+ looking for EndProperty
    
    ' WebClass has several hidden WitheEvents-like variables
    
    ' The key section looks like the following.
    ' -----------------------------------------
    '   BeginProperty WebItems {...}            << container for all WebItems
    '       WebItemCount = 4                    << nr BeginProperty statements
    '       BeginProperty WebItem1 {...}        << start of a WebItem
    '           MajorVersion = 0
    '           MinorVersion = 8
    '           Name = "Template1"              << variable name
    '           ...
    '       EndProperty
    '   EndProperty
    ' -----------------------------------------
    
    lMax = lStart
    Do  ' parse header, move to first non-whitespace character
        For lStart = lMax + 1 To gSourceFile.length - 1
            If IsEndOfLine(gSourceFile.Data(lStart)) = 0 Then
                If IsWhiteSpace(gSourceFile.Data(lStart)) = 0 Then Exit For
            End If
        Next
        If lStart = gSourceFile.length Then Exit Do
        ' move to end of statement & RTrim statement
        For lMax = lStart + 1 To gSourceFile.length
            If IsEndOfLine(gSourceFile.Data(lMax)) = 1 Then
                For n = lMax - 1 To lStart + 1 Step -1
                    If IsWhiteSpace(gSourceFile.Data(n)) = 0 Then
                        lMax = n + 1: Exit For
                    End If
                Next
                Exit For
            End If
        Next
        p = lStart
        
        ' parse the header statement
        ParseNextWordEx p, lMax, n, p

        If gSourceFile.Data(n) = vbKeyB Then
            If p - n = 13 Then                      ' BeginProperty
                If Mid$(gSourceFile.Text, n, p - n) = ParsePropBegin Then
                    If lBlocks = 0 Then
                        ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                        If Mid$(gSourceFile.Text, n, p - n) = "WebItems" Then
                            cmdState = 1
                        Else
                            ' sanity check if WebItemCount not found
                            lMax = lStart - 1: Exit Do
                        End If
                    ElseIf lBlocks = 1 Then
                        cmdState = 2
                    End If
                    lBlocks = lBlocks + 1
                End If
            End If
        ElseIf lBlocks = 1 Then
            If gSourceFile.Data(n) = vbKeyW Then
                If p - n = 12 Then                  ' WebItemCount
                    If Mid$(gSourceFile.Text, n, p - n) = "WebItemCount" Then
                        ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                        nrVars = CLng(Mid$(gSourceFile.Text, n, p - n))
                        If nrVars = 0 Then
                            lMax = lStart - 1: Exit Do
                        End If
                    End If
                End If
            Else    ' sanity check if WebItemCount not found
                lMax = lStart - 1: Exit Do
            End If
        
        ElseIf cmdState > 8 Then    ' looking for the end of variable's block
            If gSourceFile.Data(n) = vbKeyE Then    ' EndProperty
                If p - n = 11 Then
                    If Mid$(gSourceFile.Text, n, p - n) = ParsePropEnd Then
                        lBlocks = lBlocks - 1       ' decrement
                        If lBlocks = 1 Then         ' done if last web item
                            nrVars = nrVars - 1
                            If nrVars = 0 Then Exit Do
                            cmdState = cmdState Xor 8
                        End If
                    End If
                End If
            End If
        ElseIf cmdState = 2 Then
            If gSourceFile.Data(n) = vbKeyN Then
                If p - n = 4 Then                       ' Name
                    If Mid$(gSourceFile.Text, n, p - n) = ParseName Then
                        If IsExternal = False Then
                            ParseNextWordEx p, lMax, n, p, wbpEqualEOW
                            sName = Mid$(gSourceFile.Text, n + 1, p - n - 2)
                            CreateRecord lParent, sName, itVariable, 0, 0, "WebClassLibrary.WebItem", , , iaWithEvents Or iaHidden, , , scpPrivate
                        End If
                        cmdState = 9                    ' skip to EndProperty
                    End If
                End If
            End If
        End If
    Loop
    pvScanHeader_WebClass = lMax

End Function

Public Function ScanProject(IsExternal As Boolean) As Boolean

    ' Routine scans the VBP file
    ' This routine is looking for code files in use and
    '   key properties in the file along

    Dim lStart&, n&, p&, lValidation&, c%
    Dim lMax&, lParent&, lDtHigh&, lDtLow&
    Dim sName$, sPath$, sVersion$
    Dim lType As ItemTypeEnum, lAttr As ItemTypeAttrEnum
    
    ParseNextLine lMax + 1, lStart, lMax
    If InStr(Mid$(gSourceFile.Text, lStart, lMax - lStart), "=") = 0 Then
        Exit Function
    End If
    
    With gParsedItems                   ' update record properties
        sPath = .Fields(recAttr).Value
        GetFileLastModDate sPath, lDtHigh, lDtLow, 0&
        If IsExternal = False Then .Fields(recType).Value = itProject
        .Fields(recOffset).Value = lDtHigh
        .Fields(recOffset2).Value = lDtLow
        .Update
    End With
    n = InStrRev(sPath, chrSlash)            ' cache the VBP's base path
    sPath = Left$(sPath, n)
    sVersion = "M.m.R"
    
    If IsExternal = False Then          ' all VB projects use these, cache them
        lParent = m_RecordID
        gSourceFile.ProjBookMark = gParsedItems.Bookmark
        m_RefsEvents.ParseAndLoad "{000204EF-0000-0000-C000-000000000046}#6.0#9", False 'VBA
        m_RefsEvents.ParseAndLoad "{EA544A21-C82D-11D1-A3E4-00A0C90AEA82}#6.0#9", False 'VBRUN
        m_RefsEvents.ParseAndLoad "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}#6.0#9", False 'VB
    Else
        lParent = gParsedItems.Fields(recID).Value
    End If
    
    Do  ' parse vbp, move to first non-whitespace character
        p = lStart
        
        ' parse the vbp line entry
        ParseNextWordEx p, lMax, lStart, p, wbpEqualEOW
        c = gSourceFile.Data(lStart)
        Select Case p - lStart
        Case 4              ' Type,Form,Name
            If c = vbKeyF Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "Form" Then
                    lType = itSourceFile: lAttr = iaForm
                End If
            ElseIf c = vbKeyN Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = ParseName Then
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    gParsedItems.Bookmark = gSourceFile.ProjBookMark
                    gParsedItems.Fields(recName).Value = Mid$(gSourceFile.Text, lStart + 1, p - lStart - 2)
                    gParsedItems.Update
                    lValidation = lValidation + 1
                End If
            ElseIf c = vbKeyT And IsExternal = False Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = ParseTypeTyp Then
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    gParsedItems.Bookmark = gSourceFile.ProjBookMark
                    gParsedItems.Fields(recDiscrep).Value = Mid$(gSourceFile.Text, lStart, p - lStart)
                    gParsedItems.Update
                    lValidation = lValidation + 1
                End If
            End If
        Case 5              ' Class
            If c = vbKeyC Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = chrClass Then
                    lType = itSourceFile: lAttr = iaClass
                End If
            End If
        Case 6              ' Object,Module
            If IsExternal = False Then
                If c = vbKeyM Then
                    If Mid$(gSourceFile.Text, lStart, p - lStart) = "Module" Then
                        lType = itSourceFile: lAttr = iaBAS
                    End If
                ElseIf c = vbKeyO Then
                    If Mid$(gSourceFile.Text, lStart, p - lStart) = ParseObject Then lType = itReference
                End If
            End If
        Case 7              ' Startup
            If c = vbKeyS And IsExternal = False Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "Startup" Then
                    ParseNextWordEx p, lMax, lStart, p, wbpEqualEOW
                    gParsedItems.Bookmark = gSourceFile.ProjBookMark
                    gParsedItems.Fields(recAttr2).Value = Mid$(gSourceFile.Text, lStart + 1, p - lStart - 2)
                    gParsedItems.Update
                    lValidation = lValidation + 1
                End If
            End If
        Case 8              ' Designer,CondComp,HelpFile,MajorVer,MinorVer
            If c = vbKeyM Then
                Select Case Mid$(gSourceFile.Text, lStart, p - lStart)
                Case "MajorVer"
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    sVersion = Replace(sVersion, chrM, Mid$(gSourceFile.Text, lStart, p - lStart))
                    lValidation = lValidation + 1
                Case "MinorVer"
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    sVersion = Replace(sVersion, "m", Mid$(gSourceFile.Text, lStart, p - lStart))
                    lValidation = lValidation + 1
                End Select
            ElseIf c = vbKeyD Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "Designer" Then
                    lType = itSourceFile: lAttr = iaDesigner
                End If
            ElseIf IsExternal = False Then
                If c = vbKeyC Then
                    If Mid$(gSourceFile.Text, lStart, p - lStart) = "CondComp" Then
                        ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                        m_CompDirs.GlobalConstants = Mid$(gSourceFile.Text, lStart + 1, p - lStart - 2)
                    End If
                ElseIf c = vbKeyH Then
                    If Mid$(gSourceFile.Text, lStart, p - lStart) = "HelpFile" Then
                        ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                        If lStart <> lMax Then
                            sName = Mid$(gSourceFile.Text, lStart + 1, lMax - lStart - 2)
                            lType = itHelpFile
                        End If
                    End If
                End If
            End If
        Case 9              ' ResFile32,Reference
            If c = vbKeyR And IsExternal = False Then
                Select Case Mid$(gSourceFile.Text, lStart, p - lStart)
                Case "ResFile32"
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    If lStart <> lMax Then
                        sName = Mid$(gSourceFile.Text, lStart + 1, lMax - lStart - 2)
                        lType = itResFile
                    End If
                Case "Reference": lType = itReference
                End Select
            End If
        Case 10             ' RelatedDoc
            If c = vbKeyR And IsExternal = False Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "RelatedDoc" Then
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    If lStart <> lMax Then
                        sName = Mid$(gSourceFile.Text, lStart, lMax - lStart)
                        lType = itMiscFile
                    End If
                End If
            End If
        Case 11             ' RevisionVer,UserControl
            If c = vbKeyR Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "RevisionVer" Then
                    ParseNextWordEx p + 1, lMax, lStart, p    ' next word after equal sign
                    sVersion = Replace(sVersion, chrR, Mid$(gSourceFile.Text, lStart, p - lStart))
                    lValidation = lValidation + 1
                End If
            ElseIf c = vbKeyU Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "UserControl" Then
                    lType = itSourceFile: lAttr = iaUC
                End If
            End If
        Case 12             ' UserDocument,PropertyPage
            If c = vbKeyP Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "PropertyPage" Then
                    lType = itSourceFile: lAttr = iaPPG
                End If
            ElseIf c = vbKeyU And IsExternal = False Then
                If Mid$(gSourceFile.Text, lStart, p - lStart) = "UserDocument" Then
                    lType = itSourceFile: lAttr = iaUserDoc
                End If
            End If
        End Select
        
        Select Case lType
        Case itResFile, itMiscFile, itHelpFile      ' support file
            If LenB(sName) <> 0 Then
                sName = ResolveRelativePath(sName, sPath)
                n = InStrRev(sName, chrSlash) + 1
                CreateRecord lParent, Mid$(sName, n), lType, 0, 0, sName
            End If
            lType = 0
        Case itSourceFile                           ' code file
            ParseNextWordEx p + 1, lMax, lStart, p, wbpNotSemcolonEOW
            If gSourceFile.Data(p - 1) = vbKeySemicolon Then
                ParseNextWordEx p, lMax, lStart, p, wbpFirstChar ' get next token
            End If
            ' validate file is accessible
            sName = ResolveRelativePath(Mid$(gSourceFile.Text, lStart, lMax - lStart), sPath)
            If IsExternal = True Then lAttr = lAttr Or iaExternalProj
            n = GetFileLastModDate(sName, lDtHigh, lDtLow, 0&)
            If n <> 0 Then lAttr = lAttr Or iaFileError
            CreateRecord lParent, vbNullString, lType, 0, 0, sName, , , lAttr, lDtHigh, lDtLow
            lType = 0
        Case itReference                            ' external reference
            ParseNextWordEx p + 1, lMax, lStart, p
            sName = Mid$(gSourceFile.Text, lStart, lMax - lStart)
            m_RefsEvents.ParseAndLoad sName, False
            lType = 0
        End Select
        
        ParseNextLine lMax + 1, lStart, lMax
    Loop Until lStart = lMax
    
    If lValidation = 6 Or IsExternal = True Then
        ' include the Version if it was fully parsed
        If InStr(1, sVersion, chrM, vbTextCompare) = 0 Then
            If InStr(1, sVersion, chrR) = 0 Then
                gParsedItems.Bookmark = gSourceFile.ProjBookMark
                gParsedItems.Fields(recAttr2).Value = sVersion & chrSemi & gParsedItems.Fields(recAttr2).Value
                gParsedItems.Update
            End If
        End If
        ScanProject = True
    End If
    
End Function

Public Sub ScanSource()

    ' Routine is a statement parser. Typically "words" are not parsed with some exceptions
    ' Goal is to quickly scan source files to log declarations and method locations
    
    ' EOS (end of statement) occurs with colon & end of line chars (see IsEndOfLine)
    ' EOS is also the start of an inline comment, i.e., X = 0 ' zeroize
    '   colon exception #1: when used in named parameters, i.e., ParamName:=123
    '   colon exception #2: when exists inside a container
    ' If a line continuation occurs, EOS parsing continues until no more continuations
    
    ' Character containers can contain colons. Some of those can include sub-containers
    ' Since those can produce EOS false-positives they are specially handled
    '   "..."   String literals, characters between quotes (ASCII 34)
    '           example: MsgBox "Error: Permission Denied"
    '   #...#   Date literals, characters between hash tags (ASCII 35)
    '           example: Const MinDate = #1/1/1999 12:59:59 PM#
    '   [...]   Characters between square [brackets] (ASCII 91 & 93)
    '           example: Private Enum Crazy: [{)#"/:\"#(}] = 0: End Enum
    '   '...    Characters contained in comments (ASCII 39 or Rem statement)
    '           example: ' << unexpected if IPicture:Render errors here >>
    
    ' Compiler directives (#Const, #If, #Else, etc) are processed in this routine.
    ' The directive can result in subsequent code being skipped for processing.
    
    ' Declaration statements are minimally processed and logged to include positions and names
    ' Methods (Sub,Function,Property) have start and end positions logged along with their names
    
    Dim c%, lMode&                  ' 0=declarations, 1=methods
    Dim lStart&, n&, p&, lType&     ' item type started, i.e., Enum, Type, Method, etc
    Dim nrRemarks&, nrStatements&, nrExcluded&
    Dim bContd As Boolean           ' a line is split between object identifiers
    Dim bHasColon As Boolean        ' used to help distinguish label from statement
    
    Dim lCompileFlags As StatementParseFlags
    Dim lFlags As StatementParseFlags ' LoWord used for bracket counting
    ' notes: comments, brackets, string literals & date literals specially
    '   handled because they can contain EOS/directive characters
    
    m_CompDirs.Reset                                    ' compiler directive tracking
    For n = gSourceFile.First To gSourceFile.length - 1 ' gSourceFile.First will be start of Declarations section
        c = gSourceFile.Data(n)
        
        ' /// special handling
        If lFlags <> 0 Then
            If lFlags = fLiteralStr Then
                If c = vbKeyQuote Then lFlags = 0       ' end of string literal
            
            ElseIf lFlags = fComment Then               ' test for end of comment
                If IsEndOfLine(c) <> 0 Then
                    nrRemarks = nrRemarks + 1: lFlags = 0
                    If gSourceFile.Data(n - 1) = vbKeyUnderscore Then
                        If IsWhiteSpace(gSourceFile.Data(n - 2)) = 1 Then lFlags = fComment
                    End If
                    If lFlags = 0 Then bHasColon = False
                End If
            
            ElseIf lFlags = fContd Then                 ' handle unique continuation scenarios
                If c <> vbKeySpace Then
                    If c = vbKeyRemark Or c = vbKeyColon Then
                        ' non-whitespace char is comment or colon, EOS either way
                        GoSub RTrimStatement_
                        If c = vbKeyRemark Then lFlags = fComment Else bHasColon = True
                    ElseIf IsEndOfLine(c) = 1 Then
                        ' double carriage return after continuation is EOS
                        GoSub RTrimStatement_
                        bHasColon = False
                    ElseIf IsWhiteSpace(c) = 0 Then   ' typical continuation?
                        n = n - 1: lFlags = 0         ' relook at character, could be start of container
                        If bContd = False Then
                            If c = vbKeyDot Or c = vbKeyBang Then
                                bContd = True: n = n + 1
                            End If
                        End If
                    End If
                End If
            ElseIf lFlags = fLabel Then                 ' numeric label, i.e.,
                If c < vbKey0 Or c > vbKey9 Then        ' 100   someStatement
                    lFlags = 0
                    If c <> 32 Then n = n - 1           ' relook at character
                End If
            ElseIf lFlags = fLiteralDt Then
                If c = vbKeyHash Then lFlags = 0
            ElseIf c = vbKeyBracket Then                ' another open bracket
                lFlags = lFlags + 1
            ElseIf c = vbKeyBracket2 Then               ' a matched closing bracket
                lFlags = lFlags - 1                     ' will be zero when all are closed
            End If
            
        ' /// continuation of statement, looking for EOS colon
        ElseIf lStart <> 0 Then
            Select Case c
            Case vbKeySpace                             ' most common whitespace, exit quickly
            Case vbKeyColon                             ' colon
                If gSourceFile.Data(n + 1) = vbKeyEqual Then ' check for named parameter usage: ParamName:=123
                    n = n + 1
                Else
                    If bHasColon = False Then
                        For p = lStart To n - 1         ' test for labels
                            Select Case gSourceFile.Data(p)
                            Case 32 To 38, 40 To 46, 60 To 64, 91 To 94
                                Exit For
                            Case Else   ' if not one of above, then...
                                If IsWhiteSpace(gSourceFile.Data(p)) = 1 Then Exit For
                            End Select
                        Next
                        If p = n Then
                            lStart = 0: nrStatements = nrStatements + 1
                        End If
                        bHasColon = True
                    End If
                    If lStart <> 0 Then
                        If lCompileFlags = 0 Then GoSub logStatementType_ Else lStart = 0
                    End If
                End If
                
            Case vbKeyReturn, vbKeyLineFeed, &H2028, &H2029
                If gSourceFile.Data(n - 1) = vbKeyUnderscore Then ' continuation?
                    If IsWhiteSpace(gSourceFile.Data(n - 2)) = 1 Then
                        ' ignore if continuation is start of current statement
                        ' example... X = 1: _
                        '            Y = 0        ' set continuation flag as needed
                        If n - 1 = lStart Then GoTo n_
                        lFlags = lFlags Or fContd
                        If gSourceFile.Data(n - 3) = vbKeyDot Then bContd = True
                        If c = vbKeyReturn Then         ' prevent vbCrLf combo triggering two EOS events
                            If gSourceFile.Data(n + 1) = vbKeyLineFeed Then n = n + 1
                        End If
                    End If
                End If
                If (lFlags And fContd) = 0 Then         ' else a continuation
                    If lCompileFlags = 0 Then           ' else is directive or directive-ignored statement
                        GoSub logStatementType_         ' EOS, test block for decs/methods
                    Else
                        GoSub evaluateCompDir_
                    End If
                    bHasColon = False
                End If
                
            ' following will trigger special handling, if character matches
            Case vbKeyQuote: lFlags = fLiteralStr       ' string literal
            Case vbKeyRemark                            ' comment/remark
                GoSub RTrimStatement_: lFlags = fComment
            Case vbKeyHash                              ' possible date literal container
                lFlags = pvIsDateLiteral(n)
            Case vbKeyBracket: lFlags = &H1             ' open bracket - set number closing ones needed
            End Select
            
        ' /// statement started?
        ElseIf c <> vbKeySpace Then                     ' most common whitespace character
            If IsEndOfLine(c) = 0 Then
                If IsWhiteSpace(c) = 0 Then
                    lStart = n                          ' statement start
                    Select Case c
                    Case vbKeyRemark                    ' comment
                        lFlags = fComment: lStart = 0
                    Case vbKey0 To vbKey9               ' should be a numeric label
                        lFlags = fLabel: lStart = 0
                    Case vbKeyColon                     ' statements don't start with colons
                        lStart = 0: bHasColon = True
                    Case vbKeyR                             ' looking for Rem statement
                        If n + 3 <= gSourceFile.length Then
                            If Mid$(gSourceFile.Text, n, 3) = "Rem" Then
                                Select Case gSourceFile.Data(n + 3)  ' check 4th character
                                Case vbKey0 To vbKey9   ' number, not REM statement
                                Case vbKeyA To vbKeyZ   ' alpha char, not REM statement
                                Case 97 To 122          ' alpha char, not REM statement
                                Case Else: lFlags = fComment: lStart = 0: n = n + 3
                                End Select
                            End If
                        End If
                    Case vbKeyHash
                        ' if statement starts with hash, then statement
                        ' should be a compilation directive or constant
                        Select Case gSourceFile.Data(n + 1)
                        Case vbKeyI, vbKeyE, vbKeyC
                            ' get end of word
                            For p = n + 1 To gSourceFile.length
                                If IsWhiteSpace(gSourceFile.Data(p)) = 1 Then Exit For
                                If IsEndOfLine(gSourceFile.Data(p)) = 1 Then Exit For
                            Next
                            If p - n > 2 Then
                                ' note: if lCompileFlags, at this point, is non-zero
                                ' then its only possible value is fCompIgnore
                                Select Case Mid$(gSourceFile.Text, n + 1, p - n - 1)
                                Case "If":      lCompileFlags = lCompileFlags Or fCompIf
                                Case "Else":    lCompileFlags = lCompileFlags Or fCompElse
                                Case "End"
                                    If Mid$(gSourceFile.Text, p + 1, 2) = "If" Then
                                        lCompileFlags = lCompileFlags Or fCompEnd
                                        p = p + 2
                                    End If
                                Case "ElseIf":  lCompileFlags = lCompileFlags Or fCompElseIf
                                Case ParseTypeCnt: lCompileFlags = lCompileFlags Or fCompConst
                                End Select
                                If (lCompileFlags And fCompMask) <> 0 Then
                                    lFlags = 0: n = p - 1 ' move pointer past parsed word
                                End If
                            End If
                        End Select
                    Case vbKeyA ' ignore Attribute statements, verify word
                        If n + 8 < gSourceFile.length Then
                            If gSourceFile.Data(n + 8) = 101 Then ' e
                                If IsWhiteSpace(gSourceFile.Data(n + 9)) <> 0 Then
                                    lStart = 0
                                ElseIf IsEndOfLine(gSourceFile.Data(n + 9)) <> 0 Then
                                    lStart = 0
                                End If
                                If lStart = 0 Then      ' validate
                                    If Mid$(gSourceFile.Text, n, 9) = ParseTypeAtt Then
                                        lFlags = fComment: nrRemarks = nrRemarks - 1
                                    Else
                                        lStart = n
                                    End If
                                End If
                            End If
                        End If
                    ' sanity check on following; doubtful they can start a statement
                    Case vbKeyBracket: lFlags = &H1       ' open bracket - set number closing brackets needed
                    Case vbKeyQuote: lFlags = fLiteralStr ' string literal
                    End Select
                End If
            End If
        End If
n_: Next n

    If lCompileFlags <> 0 Then
        gParsedItems.Source.SetError vnOpenCompIF
    ElseIf lFlags <> 0 Then
        If lFlags = fLiteralStr Then
            gSourceFile.Owner.SetError vnOpenQuote
        ElseIf lFlags = fContd Or lFlags = fComment Or lFlags = fLabel Then
            ' n/a
        ElseIf lFlags = fLiteralDt Then
            gSourceFile.Owner.SetError vnOpenDate
        Else
            gSourceFile.Owner.SetError vnOpenBracket
        End If
    ElseIf lType <> 0 Then
        If lMode = 0 Then
            gSourceFile.Owner.SetError vnOpenEnumUDT
        Else
            gSourceFile.Owner.SetError vnOpenMethodBlk
        End If
    End If
    
    With gParsedItems           ' update decs start & methods end
        .Bookmark = gSourceFile.CPBookMark
        .Fields(recStart).Value = gSourceFile.First
        .Fields(recOffset2).Value = n
        .Update
    End With
    If m_IdxOffset <> 0 Then    ' if any statements cached, log them
        gSourceFile.Owner.LogOffsets VarPtr(m_Offsets(0)), m_IdxOffset * 4
    End If
    gSourceFile.Owner.LogStats iaStatements, nrStatements, iaComments, nrRemarks, iaExclusions, nrExcluded
    Exit Sub
    
RTrimStatement_:
    For p = n - 1 To lStart Step -1
        If IsWhiteSpace(gSourceFile.Data(p)) = 0 Then
            If gSourceFile.Data(p) = vbKeyUnderscore Then
                If IsWhiteSpace(gSourceFile.Data(p - 1)) = 0 Then Exit For
            Else
                Exit For
            End If
        End If
    Next
    lFlags = n: n = p + 1
    If lCompileFlags = 0 Then
        GoSub logStatementType_
    Else
        GoSub evaluateCompDir_
    End If
    n = lFlags: lFlags = 0
    Return

evaluateCompDir_:
    ' minimally at this point, lCompileFlags will include fCompIgnore (ignoring statements)
    Select Case (lCompileFlags And fCompMask)
    Case fCompConst
        If (lCompileFlags And fCompIgnore) = 0 Then
            m_CompDirs.AddLocalConstant Mid$(gSourceFile.Text, lStart, n - lStart)
        End If
        lCompileFlags = lCompileFlags Xor fCompConst
    Case fCompIf, fCompElseIf
        lCompileFlags = m_CompDirs.IsBlockValid(Mid$(gSourceFile.Text, lStart, n - lStart), lCompileFlags)
    Case fCompElse, fCompEnd
        lCompileFlags = m_CompDirs.IsBlockValid(vbNullString, lCompileFlags)
    Case Else
        nrExcluded = nrExcluded + 1
    End Select
    lStart = 0
    Return
    
logStatementType_:
    p = 0
    If lType = itEnumMember Then
        lType = pvParseName_EnumMbr(lStart, n)
    ElseIf lType = 0 Then       ' look for declarations statement or method
        If lMode = 0 Then       ' declarations?
            lType = pvIsDecsStart(lStart, n, bContd)
            If lType = 0 Then   ' done with Declarations, methods next
                With gParsedItems   ' log decs end & methods start
                    .Bookmark = gSourceFile.CPBookMark
                    .Fields(recEnd).Value = lStart - 1
                    .Fields(recOffset).Value = lStart
                    .Update
                End With
                lMode = 1
                GoSub logStatementType_
                Return
            End If
        Else                    ' method?
            lType = pvIsMethodStart(lStart, n, bContd)
            If lType = 0 Then p = 1
        End If
        If lMode = 0 Then       ' need to look for the end of a block?
            If lType <> itEnumMember Then
                If lType <> itEnum And lType <> itType Then lType = 0
            End If
        End If
    ElseIf pvIsEndBlock(lStart, n, lType, p) = True Then
        lType = 0               ' done with current block of code
    Else
        p = p Xor 1
    End If
    If p = 1 Then
        m_Offsets(m_IdxOffset) = lStart
        m_IdxOffset = m_IdxOffset + 2
        If bContd = True Then   ' flag offset as having continuation character
            m_Offsets(m_IdxOffset - 1) = -n
        Else
            m_Offsets(m_IdxOffset - 1) = n
        End If
        If m_IdxOffset > UBound(m_Offsets) Then
            gSourceFile.Owner.LogOffsets VarPtr(m_Offsets(0)), m_IdxOffset * 4
            m_IdxOffset = 0
        End If
    End If
    nrStatements = nrStatements + 1
    lStart = 0: bContd = False
    Return

End Sub

Public Function GetFileHandle(ByVal FileName As String, WriteMode As Boolean) As Long

    ' Function uses APIs to create a file handle to read/write files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    
    Dim Flags&, access&, Disposition&, Share&
    
    If WriteMode Then
        access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        Flags = GetFileAttributesW(StrPtr(FileName))
        If (Flags And FILE_ATTRIBUTE_READONLY) Then
            Flags = FILE_ATTRIBUTE_NORMAL
            SetFileAttributesW StrPtr(FileName), Flags
        End If
        If Flags < 0& Then Flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    
    Else
       access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    End If
    
    GetFileHandle = CreateFileW(StrPtr(FileName), access, Share, ByVal 0&, Disposition, Flags, 0&)
    
End Function

Public Function ResolveRelativePath(sRelPath As String, sBasePath As String) As String

    ' easy method to expand relative paths & ensure backslash continuity, i.e., \ vs /
    ' VB uses relative paths often in its vbp & code files

    Dim sBuffer$, n&, p&
    
    If Left$(sRelPath, 2) = "\\" Then
        p = 1
    ElseIf Left$(sRelPath, 2) = "//" Then
        p = 1
    Else
        p = InStr(2, sRelPath, chrColon)
    End If
    n = 256: sBuffer = Space$(n)
    If p = 0 Then
        n = GetFullPathNameW(StrPtr(sBasePath & sRelPath), n, StrPtr(sBuffer), 0&)
    Else
        n = GetFullPathNameW(StrPtr(sRelPath), n, StrPtr(sBuffer), 0&)
    End If
    If n = 0 Then
        ResolveRelativePath = sRelPath
    Else
        ResolveRelativePath = Left$(sBuffer, n)
    End If
        
End Function

Public Function GetFileLastModDate(FileName As String, _
                                lDateHigh As Long, lDateLow As Long, lSize As Long) As Long

    Dim WFAD&(0 To 8)       ' faux WIN32_FILE_ATTRIBUTE_DATA structure
        ' 0     = dwFileAttributes
        ' 1-2   = ftCreationTime
        ' 3-4   = ftLastAccessTime
        ' 5-6   = ftLastWriteTime
        ' 7-8   = nFileSize (low/high)
    
    If GetFileAttributesExW(StrPtr(FileName), 0&, WFAD(0)) = 0 Then
        GetFileLastModDate = vnFileNotFound
        
    ElseIf WFAD(7) <> 0 Then
        GetFileLastModDate = vnFileTooBig
    
    ElseIf WFAD(8) = 0 Then
        GetFileLastModDate = vnFileEmpty
        
    Else
        lDateHigh = WFAD(5)
        lDateLow = WFAD(6)
        lSize = WFAD(8)
    End If

End Function

Public Function SetQuery(ParamArray queryParams() As Variant) As String

    ' Lots of recordset queries (rs.Filter & rs.Find) are used througout
    '   the project. Just wanted a simple routine to create these dynamically that
    '   can be adapted to passing constants, field names, etc and almost as easy to
    '   read (from the context of calling routine) as having plain text filters.
    '   i.e., instead of seeing: rs.Filter = "RecID=" & lID
    '                  we'd see: rs.Filter = SetQuery(recID, qryIs, lID)
    ' main advantage: change or add field names? Don't need to make dozens upon dozens of
    ' changes in code. Simply update the constFld[xxx] value, in declarations section, and done.
    
    Dim n&, lMax&, sValue$, sQuery$
    
    lMax = UBound(queryParams)
    Do
        Select Case queryParams(n + 1)
            Case qryIs: sQuery = sQuery & chrParentO & queryParams(n) & "="
            Case qryGT: sQuery = sQuery & chrParentO & queryParams(n) & ">"
            Case qryLT: sQuery = sQuery & chrParentO & queryParams(n) & "<"
            Case qryNot: sQuery = sQuery & chrParentO & queryParams(n) & "<>"
            Case qryLike: sQuery = sQuery & chrParentO & queryParams(n) & " Like "
            Case qryGTE: sQuery = sQuery & chrParentO & queryParams(n) & ">="
        End Select
        If VarType(queryParams(n + 2)) = vbString Then
            sValue = queryParams(n + 2)
            If InStr(sValue, chrApos) = 0 Then '  ensure doubled apostrophes if needed
                sQuery = sQuery & chrApos & sValue & chrApos
            Else
                sQuery = sQuery & chrApos & Replace(sValue, chrApos, "''") & chrApos
            End If
        Else
            sQuery = sQuery & CStr(queryParams(n + 2))  ' numeric only, no dates are used in this project
        End If
        n = n + 3
        If n > lMax Then
            sQuery = sQuery & chrParentC
            Exit Do
        End If
        If queryParams(n) = qryAnd Then
            sQuery = sQuery & ") And "
        Else
            sQuery = sQuery & ") Or "
        End If
        n = n + 1
    Loop
    SetQuery = sQuery
    
End Function

Public Sub ResolveEvents(ExternalProjs As Boolean)

    ' the last routine called during initial scan of a project

    ' An event is a method associated with classes, implemented objects,
    '   controls and variables using WithEvents. An Event statement
    '   defined within the declarations section isn't one of these.
    ' Routine will verify that methods that look like events are actually
    '   events. And if verified, then the event becomes a child record
    '   of the object that is using the event. At this point, all methods
    '   are child records of the code page.

    Dim rsObjs As ADODB.Recordset, rsMethods As ADODB.Recordset
    Dim rsCodePg As ADODB.Recordset, rsEvents As ADODB.Recordset
    Dim n&

    Set rsObjs = gParsedItems.Clone
    Set rsMethods = gParsedItems.Clone
    
    With gParsedItems                   ' sort for faster rs queries
        .Sort = recAttr: .MoveFirst
        Do                              ' sort by Attr1 field
            n = n + 1
            .Fields(recIdxAttr).Value = n
            .MoveNext
        Loop Until .EOF = True
        .Sort = recName: .MoveFirst: n = 1
        Do                              ' sort by Name field next
            .Fields(recIdxName).Value = n
            n = n + 1
            .MoveNext
        Loop Until .EOF = True
        .Filter = 0
    End With
    
    ' get listing of all code pages, excluding external ones
    rsObjs.Filter = SetQuery(recType, qryIs, itCodePage, qryAnd, _
                             recParent, qryNot, 0)
    If rsObjs.EOF = False Then
        Call pvResolveEvents_Classes(rsObjs, rsMethods)
    End If
    
    ' get listing of all code pages as potential implemented,withevents,control targets
    Set rsCodePg = gParsedItems.Clone
    rsCodePg.Filter = SetQuery(recType, qryIs, itCodePage)
    If rsCodePg.EOF = False Then
        Set rsEvents = gParsedItems.Clone
        
        ' get list of all Implementations
        rsObjs.Filter = SetQuery(recType, qryIs, itImplements)
        If rsObjs.EOF = False Then
            If pvResolveEvents_Implements(ExternalProjs, rsCodePg, rsObjs, rsMethods, rsEvents) = True Then
                rsCodePg.MoveFirst
                Do
                    If (rsCodePg.Fields(recFlags).Value And iaImplemented) <> 0 Then
                        rsObjs.Filter = SetQuery(recType, qryIs, itMethod, _
                                            qryAnd, recParent, qryIs, rsCodePg.Fields(recID).Value, _
                                            qryAnd, recScope, qryNot, scpPrivate)
                        Do Until rsObjs.EOF = True
                            rsObjs.Fields(recType).Value = itClassEvent
                            rsObjs.Fields(recDiscrep).Value = Replace(rsObjs.Fields(recDiscrep).Value, chrZ, vbNullString)
                            rsObjs.MoveNext
                        Loop
                    End If
                    rsCodePg.MoveNext
                Loop Until rsCodePg.EOF = True
            End If
        End If
        ' get list of all WithEvent variables
        rsObjs.Filter = SetQuery(recType, qryIs, itVariable, qryAnd, _
                                 recFlags, qryLT, 0)
        If rsObjs.EOF = False Then
            Call pvResolveEvents_WithEvents(ExternalProjs, rsCodePg, rsObjs, rsMethods, rsEvents)
        End If
        ' get list of all controls
        rsObjs.Filter = SetQuery(recType, qryIs, itControl)
        If rsObjs.EOF = False Then
            Call pvResolveEvents_Controls(ExternalProjs, rsCodePg, rsObjs, rsMethods, rsEvents)
        End If
        rsEvents.Close: Set rsEvents = Nothing
    End If
    rsObjs.Close: Set rsObjs = Nothing
    rsMethods.Close: Set rsMethods = Nothing
    rsCodePg.Close: Set rsCodePg = Nothing
    
End Sub

Private Function pvIsDecsStart(ByVal lStart As Long, lMax As Long, _
                                isSplit As Boolean) As ItemTypeEnum
    
    ' routine parses out declaration section statements
    ' if routine returns zero, then end of declarations
    
    Dim lType As ItemTypeEnum, c%
    Const ParseTypeDim = "Dim", ParseTypeApi = "Declare", ParseTypeGbl = "Global"
    Const ParseTypeEvt = "Event", ParseTypeOpt = "Option", ParseTypeImp = "Implements"
    Const ParseTypeDInt = "DefInt", ParseTypeDLng = "DefLng", ParseTypeDBol = "DefBool"
    Const ParseTypeDByt = "DefByte", ParseTypeDCur = "DefCur", ParseTypeDSng = "DefSng"
    Const ParseTypeDDt = "DefDate", ParseTypeDStr = "DefStr", ParseTypeDObj = "DefObj"
    Const ParseTypeDVar = "DefVar", ParseTypeDDec = "DefDec"
    
    c = gSourceFile.Data(lStart + 1)     ' 2nd character
    Select Case gSourceFile.Data(lStart) ' 1st character
    Case vbKeyP                 ' looking for Public,Private & 2nd char is u or v
        If c <> 114 And c <> 117 Then lStart = 0
    Case vbKeyD                 ' looking for Dim, Declare, Defxxx & 2nd char is i or e
        If c <> 105 And gSourceFile.Data(lStart + 1) <> 101 Then lStart = 0
    Case vbKeyC                 ' looking for Const & 2nd char is o
        If c <> 111 Then lStart = 0
    Case vbKeyE                 ' looking for Enum, Event & 2nd char is n or v
        If c <> 110 And c <> 118 Then lStart = 0
    Case vbKeyT                 ' looking for Type & 2nd char is y
        If c <> 121 Then lStart = 0
    Case vbKeyO                 ' looking for Option & 2nd char is p
        If c <> 112 Then lStart = 0
    Case vbKeyI                 ' looking for Implements & 2nd char is m
        If c <> 109 Then lStart = 0
    Case vbKeyW                 ' looking for WithEvents & 2nd char is i
        If c <> 105 Then lStart = 0
    Case vbKeyG                 ' looking for Global & 2nd char is l
        If c <> 108 Then lStart = 0
    Case vbKeyA                 ' looking for Attribute & 2nd char is t
        If c <> 116 Then lStart = 0
    Case Else: lStart = 0
    End Select
        
    If lStart <> 0 Then
        Dim lScope As ItemScopeEnum, p&, n&
        
        p = lStart
        Do
            ParseNextWordEx p, lMax, n, p
            Select Case Mid$(gSourceFile.Text, n, p - n)         ' get the word
            Case ParseTypePub, ParseTypePriv, ParseTypeGbl
                ' words that can precede declarations
                If gSourceFile.Data(n + 1) = 114 Then
                    lScope = scpPrivate
                Else
                    lScope = scpPublic
                End If
            Case ParseTypeDim: lType = itVariable
            Case ParseTypeApi: lType = itAPI
                If lScope = scpLocal Then lScope = scpPublic
            Case ParseTypeCnt: lType = itConstant
            Case ParseTypeTyp: lType = itType
                If lScope = scpLocal Then lScope = scpPublic
            Case ParseTypeEnm: lType = itEnum
                If lScope = scpLocal Then lScope = scpPublic
            Case ParseTypeEvt: lType = itEvent
                If lScope = scpLocal Then lScope = scpPublic
            Case ParseTypeOpt: lType = -1
                pvParseName_Option p, lMax: GoTo exitRoutine
            Case ParseTypeImp: lType = itImplements
                pvParseName_Misc gSourceFile.Owner.RecordID, lStart, p, lMax, itImplements, scpPrivate, isSplit
                GoTo exitRoutine
            Case ParseTypeDInt, ParseTypeDLng, ParseTypeDBol, ParseTypeDByt
                lType = itDefType
            Case ParseTypeDCur, ParseTypeDSng, ParseTypeDDt, ParseTypeDStr
                lType = itDefType
            Case ParseTypeDObj, ParseTypeDVar, ParseTypeDDec
                lType = itDefType
            Case ParseTypeAtt: lType = -1
                ' statement is non-header Attribute, no additional processing
                GoTo exitRoutine
            Case ParseTypeFnc, ParseTypePpy, ParseTypeSub: GoTo exitRoutine ' end of declarations
            Case ParseTypeFrnd, ParseTypeStat: GoTo exitRoutine ' end of declarations
            Case Else: lType = itVariable: p = n
            End Select
            
            If lType = itDefType Then
                pvParseName_DEF n, lMax
            ElseIf lType <> 0 Then
                If lScope = scpLocal Then
                    lScope = scpPrivate
                ElseIf lScope = scpPublic Then
                    Select Case lType
                    Case itEnum, itType
                        lScope = scpGlobal
                    Case itAPI, itVariable, itConstant
                        If (gSourceFile.Owner.FileAttrs And iaMaskCodePage) = iaBAS Then lScope = scpGlobal
                    End Select
                End If
                Select Case lType
                Case itEnum, itType: pvParseName_Misc gSourceFile.Owner.RecordID, lStart, p, lMax, lType, lScope, isSplit
                    If lScope = scpGlobal And lType = itEnum Then lType = itEnumMember
                Case itAPI:         pvParseName_API lStart, p, lMax, lScope, isSplit
                Case itConstant:    pvParseName_Constant gSourceFile.Owner.RecordID, lStart, p, lMax, lScope, isSplit
                Case itVariable:    pvParseName_Variable gSourceFile.Owner.RecordID, lStart, p, lMax, lScope, isSplit
                Case itEvent:       pvParseName_Method lStart, p, lMax, lType, lScope, isSplit
                End Select
            End If
        Loop While lType = 0
    End If
    
exitRoutine:
    pvIsDecsStart = lType

End Function

Private Function pvIsMethodStart(ByVal lStart As Long, lMax As Long, isSplit As Boolean) As Long
    
    ' Function identifies passed statement as a method signature
    '   i.e., Private|Public|Friend|Static Function|Sub|Property
    ' Return value is method type or zero if not a method signature
    
    Dim c%, n&
    
    c = gSourceFile.Data(lStart + 1)
    Select Case gSourceFile.Data(lStart)
    Case vbKeyP, vbKeyS, vbKeyF  ' P,gSourceFile.Text,F: looking for Public,Private,Property,Sub,Static,Friend,Function
        If c <> 117 Then    ' P,gSourceFile.Text,F: second character u: Public,Sub,Function
            If gSourceFile.Data(lStart) = vbKeyS Then  ' gSourceFile.Text: second character t: Static
                If c <> 116 Then lStart = 0
            ElseIf c <> 114 Then ' F,P: second character r: Friend,Private,Property
                lStart = 0
            End If
        End If
    End Select
            
    If lStart <> 0 Then
        Dim lType As ItemTypeAttrEnum, lScope As ItemScopeEnum
        Dim p&, lParent&
        
        p = lStart
        Do
            ParseNextWordEx p, lMax, n, p
            Select Case Mid$(gSourceFile.Text, n, p - n)   ' get the word
            Case ParseTypePub, ParseTypePriv, ParseTypeFrnd, ParseTypeStat
                ' words that can precede method type, i.e., Private Static MethodName()
                If gSourceFile.Data(n) = vbKeyF Then
                    lScope = scpFriend
                ElseIf gSourceFile.Data(n + 1) = 114 Then
                    lScope = scpPrivate
                Else
                    lScope = scpPublic
                End If
            Case ParseTypeFnc: lType = iaFunction
            Case ParseTypeSub: lType = iaSub
            Case ParseTypePpy: lType = iaProperty
            Case Else: Exit Do                          ' none of the above
            End Select
        Loop While lType = 0
        
        If lType <> 0 Then
            If lScope = scpLocal Then lScope = scpPublic
            If lScope = scpPublic Then
                If (gSourceFile.Owner.FileAttrs And iaMaskCodePage) = iaBAS Then lScope = scpGlobal
            End If
            pvParseName_Method lStart, p, lMax, lType, lScope, isSplit
            pvIsMethodStart = lType
        End If
    End If

End Function

Private Function pvIsEndBlock(ByVal lStart As Long, lMax As Long, lType As Long, lAttr As Long) As Boolean

    ' function looks for "End xxx" statement for Function,Sub,Property,Type,Enum statements
    ' Returns True if is end statement

    Dim lLen&
    If gSourceFile.Data(lStart) = vbKeyE Then                ' looking for End
        If gSourceFile.Data(lStart + 1) = 110 Then           ' n
            If gSourceFile.Data(lStart + 2) = 100 Then       ' d
                If lType = iaSub Then
                    lLen = 3                        ' Len("Sub")
                ElseIf lType = itEnum Or lType = itType Then
                    lLen = 4                        ' Len("Type"),Len("Enum")
                Else
                    lLen = 8                        ' Len("Function"),Len("Property")
                End If
                If lMax - lStart > lLen + 3 Then
                    ParseNextWordEx lStart + 3, lMax, lStart, 0&, wbpFirstChar
                    If lType = iaFunction Then      ' looking for End Function
                        If gSourceFile.Data(lStart) = vbKeyF Then
                            pvIsEndBlock = (Mid$(gSourceFile.Text, lStart, lLen) = ParseTypeFnc)
                        End If
                    ElseIf lType = iaSub Then       ' looking for End Sub
                        If gSourceFile.Data(lStart) = vbKeyS Then
                            pvIsEndBlock = (Mid$(gSourceFile.Text, lStart, lLen) = ParseTypeSub)
                        End If
                    ElseIf lType = itType Then      ' looking for End Type
                        If gSourceFile.Data(lStart) = vbKeyT Then
                            pvIsEndBlock = (Mid$(gSourceFile.Text, lStart, lLen) = ParseTypeTyp)
                        End If
                    ElseIf lType = itEnum Then ' looking for End Enum
                        If gSourceFile.Data(lStart) = vbKeyE Then
                            pvIsEndBlock = (Mid$(gSourceFile.Text, lStart, lLen) = ParseTypeEnm)
                        End If
                    ElseIf lType = iaProperty Then  ' looking for End Property
                        If gSourceFile.Data(lStart) = vbKeyP Then
                            pvIsEndBlock = (Mid$(gSourceFile.Text, lStart, lLen) = ParseTypePpy)
                        End If
                    End If
                    If pvIsEndBlock = True Then     ' else likely End With
                        With gParsedItems
                            .Bookmark = gSourceFile.ItemBookMark
                            .Fields(recOffset2).Value = lMax
                            .Update
                        End With
                    End If
                End If
            
            End If
        End If
    ElseIf gSourceFile.Data(lStart) = vbKeyA Then
        If gSourceFile.Data(lStart + 1) = 116 Then  ' t
            ParseNextWordEx lStart, lMax, lStart, lLen
            If lLen - lStart = 9 Then
                If Mid$(gSourceFile.Text, lStart, 9) = ParseTypeAtt Then lAttr = 1
            End If
        End If
    End If

End Function

Public Function CRCItem(sText As String, bUnicode As Boolean, _
                        Optional InitCRC As Long = -1) As Long

    ' Standard CRC algorithm
    ' Returns -1 failure or 0 success, or the CRC value depending on parameters...
    
    ' Params:
    '   sText: non-null string
    '   bUnicode: true to CRC 2-bytes per character, else 1-byte for ANSI text
    
    Dim i&, crc32val&, iLookup&
    Dim bData() As Byte, tSA&(0 To 5)
    
    crc32val = InitCRC
    If LenB(sText) <> 0 Then
        tSA(0) = 1: tSA(1) = 1: tSA(3) = StrPtr(sText): tSA(4) = LenB(sText)
        CopyMemory ByVal VarPtrArray(bData), VarPtr(tSA(0)), 4&
        If bUnicode = True Then iLookup = 1 Else iLookup = 2
        For i = 0 To UBound(bData) Step iLookup
            iLookup = (crc32val And &HFF&) Xor bData(i)
            crc32val = (((crc32val And &HFFFFFF00) \ &H100&) And &HFFFFFF) Xor m_CRC32LUT(iLookup)
        Next i
        CopyMemory ByVal VarPtrArray(bData()), 0&, 4&
    End If
    CRCItem = crc32val

End Function

Public Sub FauxDoEvents()
    ' pulled from this posting
    ' http://www.vbforums.com/showthread.php?315416-Ok-noobies-DoEvents-is-slow!!!-Here-s-are-faster-methods
    
    ' only calls DoEvents when absolutely necessary.
    ' potential side-effect: if form is marked by Windows as "Not Responding",
    '   it should clear relatively quickly but in doing so, form visibly repaints
    Const QS_KEY As Long = &H1
    Const QS_MOUSEBUTTON As Long = &H4
    Const QS_POSTMESSAGE As Long = &H8
    Const QS_SENDMESSAGE As Long = &H40
    If GetQueueStatus(QS_KEY Or QS_MOUSEBUTTON Or QS_POSTMESSAGE Or QS_SENDMESSAGE) <> 0 Then DoEvents
    
End Sub

Private Sub pvCreateCRC32LUT()

    ' initialize the CRC lookup table/array

    Const CRCpolynomial = &HEDB88320
    ' &HEDB88320 is the official polynomial used by CRC32 in PKZip & zLIB.
    
    Dim i&, j&, lValue&
    
    ' create a CRC lookup table (LUT)
    For i = 0 To 255
        lValue = i
        For j = 0 To 7
            If (lValue And 1&) Then
                lValue = (((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF) Xor CRCpolynomial
            Else
                lValue = ((lValue And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        m_CRC32LUT(i) = lValue
    Next i

End Sub

Public Sub ParseNextLine(ByVal lStart As Long, wdStart As Long, wdEnd As Long)

    ' common use method to return a single line of text, carriage-return delimited
    ' used to parse vbp, vbg, and header sections of code pages
    ' on return, wdStart & wdEnd are bounds of the line, left & right trimmed

    Dim n&
    For wdStart = lStart To gSourceFile.length - 1
        If IsEndOfLine(gSourceFile.Data(wdStart)) = 0 Then
            If IsWhiteSpace(gSourceFile.Data(wdStart)) = 0 Then Exit For
        End If
    Next
    If wdStart < gSourceFile.length Then
        ' move to end of statement & RTrim statement
        For wdEnd = wdStart + 1 To gSourceFile.length
            If IsEndOfLine(gSourceFile.Data(wdEnd)) = 1 Then
                For n = wdEnd - 1 To lStart + 1 Step -1
                    If IsWhiteSpace(gSourceFile.Data(n)) = 0 Then
                        wdEnd = n + 1: Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Else
        wdEnd = wdStart
    End If

End Sub

Public Sub ParseNextWordEx(ByVal lStart As Long, ByVal lMax As Long, _
                          wdStart As Long, wdEnd As Long, _
                          Optional ByVal WordBreakOptions As WordBreakOptEnum)
                            
    ' Routine parses words from a previously parsed statement
    ' Parsed statements are guaranteed to be trimmed and do not contain comments/remarks
    
    ' EOW (end of word) character is any white space value (ASCII 32 and lower)
    '   also included as EOW are unicode line breaks: &H2028,&H2029
    '   also included are commas and semicolons
    '   Exception: when the EOW occurs inside a string literal or inside square brackets
    '   WordBreakOptions parameter can add other exceptions
    
    ' if wdEnd is lMax on return, end of statement occurred
    
    ' lStart: character position to begin parsing for a "word"
    ' lMax: length of statement + 1
    ' wdStart: on return, position for 1st character of parsed word
    ' wdEnd: on return, position immediately after last character of parsed word
    ' WordBreakOptions:
    '   wbpParentsCntr      parenthesis pairs treated as character container
    '   wbpParentsEOW       open parenthesis treated as EOW
    '   wbpEqualEOW         equal sign treated as EOW
    '   wbpNotCommaEOW      comma is not a wordbreak
    '   wbpNotSemcolonEOW   semicolon is not a wordbreak
    '   wbpSepNamedParams   skip named parameter, return parameter value
    '   wbpContinuation     return entire line segment before continuation character
    '   wbpFirstChar        first character acts EOW
    
    ' usage examples
    ' - keep code inside parentheses togeher, use wbpParentsCntr:
    '   statement is: strArray() = Split(Text1.Text, vbCrLf)(0)
    '   when routine is run to extract all words, then:
    '   word #1: strArray
    '   word 32: ()
    '   word #3: =
    '   word #4: Split
    '   word #5: (Text1.Text, vbCrLf)
    '   word #6: (0)
    ' - want commas returned to help identify comma-delimited items, use wbpNotCommaEOW:
    '   statement is: Const vbkeyColon = 58, vbKeyEqual = 61
    '   when routine is run to extract all words, then:
    '   word #1: Const
    '   word #2: vbkeyColon
    '   word #3: =
    '   word #4: 58,
    '   word #5: vbKeyEqual
    '   word #6: =
    '   word #7: 61
    
    Dim lFlags As StatementParseFlags
    Dim c%, lParents&, lOffset&, lBrkOffset&, lSplitIdent&
    
    For wdStart = lStart To lMax - 1                ' look for start of a word
        c = gSourceFile.Data(wdStart)
        If c <> vbKeySpace Then
            If IsWhiteSpace(c) = 0 Then
                If IsEndOfLine(c) = 0 Then
                    Select Case c
                    Case vbKeyParenthesis           ' parentheses specially handled
                        If (WordBreakOptions And 3) <> 0 Then
                            If (WordBreakOptions And wbpParentsCntr) <> 0 Then lParents = 1
                        End If                      ' in any case, start of a "word"
                        Exit For
                    Case vbKeyComma                 ' comma included in word?
                        If (WordBreakOptions And wbpNotCommaEOW) <> 0 Then Exit For ' if so, set offset
                    Case vbKeySemicolon
                        If (WordBreakOptions And wbpNotSemcolonEOW) <> 0 Then Exit For ' if so, set offset
                    Case vbKeyUnderscore            ' continuation characters treated as white space
                        Select Case gSourceFile.Data(wdStart + 1)
                        Case vbKeyReturn, vbKeyLineFeed, &H2028, &H2029
                        Case Else: Exit For
                        End Select
                    Case vbKeyColon                 ' named parameter, verify & skip
                        If gSourceFile.Data(wdStart + 1) = vbKeyEqual Then wdStart = wdStart + 1
                    Case vbKeyEqual
                        If (WordBreakOptions And wbpEqualEOW) = 0 Then Exit For
                    Case vbKeyQuote: lFlags = fLiteralStr: Exit For
                    Case vbKeyBracket: lFlags = 1: lBrkOffset = wdStart: Exit For
                    Case vbKeyHash
                        If (WordBreakOptions And wbpFirstChar) = 0 Then
                            lFlags = pvIsDateLiteral(wdEnd)
                            If lFlags = fLiteralDt Then lOffset = 8
                        End If
                        Exit For
                    Case Else: Exit For             ' any other character is start of a "word"
                    End Select
                End If
            End If
        End If
    Next
    If wdStart = lMax Then                          ' abort if no word start found
        wdEnd = lMax: Exit Sub
    ElseIf (WordBreakOptions And wbpFirstChar) <> 0 Then
        wdEnd = wdStart: Exit Sub
    End If
    
    If (WordBreakOptions And wbpLineSep) <> 0 Then lFlags = fContd
    
    For wdEnd = wdStart + 1 + lOffset To lMax - 1   ' now find end of word
        c = gSourceFile.Data(wdEnd)
        If lFlags = 0 Then                          ' else special handling for containers
            If c = vbKeyQuote Then
                lFlags = fLiteralStr                ' start string literal container
            ElseIf c = vbKeyBracket Then
                lFlags = 1: lBrkOffset = wdEnd      ' start square bracket container
            ElseIf c = vbKeyHash Then
                lFlags = pvIsDateLiteral(wdEnd)
            ElseIf lParents <> 0 Then               ' parentheses container in effect
                If c = vbKeyParenthesis2 Then
                    lParents = lParents - 1             ' decrease container count
                    If lParents = 0 Then
                        wdEnd = wdEnd + 1               ' place after closing parenthesis
                        If (WordBreakOptions And wbpNotCommaEOW) <> 0 Then
                            If wdEnd < lMax - 1 Then    ' include trailing , or ; as needed
                                If gSourceFile.Data(wdEnd) = vbKeyComma Then
                                    wdEnd = wdEnd + 1: Exit For
                                End If
                            End If
                        End If
                        If (WordBreakOptions And wbpNotSemcolonEOW) <> 0 Then
                            If wdEnd < lMax - 1 Then    ' include trailing , or ; as needed
                                If gSourceFile.Data(wdEnd) = vbKeySemicolon Then wdEnd = wdEnd + 1
                            End If
                        End If
                        Exit For
                    End If
                ElseIf c = vbKeyParenthesis Then
                    lParents = lParents + 1             ' increase container count
                End If
            ElseIf IsWhiteSpace(c) = 1 Then
                If (WordBreakOptions And wbpContinuation) = 0 Then Exit For
            Else
                Select Case c
                Case vbKeyComma                 ' done if comma not EOW
                    If (WordBreakOptions And wbpNotCommaEOW) = 0 Then Exit For
                Case vbKeySemicolon
                    If (WordBreakOptions And wbpNotSemcolonEOW) = 0 Then Exit For
                Case vbKeyParenthesis           ' done if parentheses EOW
                    If (WordBreakOptions And 3) <> 0 Then Exit For
                Case vbKeyUnderscore
                    If (WordBreakOptions And wbpContinuation) <> 0 Then
                        If IsEndOfLine(gSourceFile.Data(wdEnd + 1)) = 1 Then
                            If IsWhiteSpace(gSourceFile.Data(wdEnd - 1)) = 1 Then
                                wdEnd = wdEnd - 1: Exit For
                            End If
                        End If
                    End If
                Case vbKeyColon                 ' separate named params?
                    If (WordBreakOptions And wbpSepNamedParams) <> 0 Then Exit For
                Case vbKeyEqual
                    If (WordBreakOptions And wbpEqualEOW) <> 0 Then Exit For
                Case vbKeySemicolon: Exit For
                End Select
            End If
        ElseIf lFlags = fLiteralStr Then            ' string literal container end?
            If c = vbKeyQuote Then lFlags = 0
        ElseIf lFlags = fLiteralDt Then
            If c = vbKeyHash Then lFlags = 0
        ElseIf lFlags = fContd Then
            If lOffset = 0 Then lOffset = wdStart
            If c = vbKeyUnderscore Then
                If lOffset < wdEnd - 1 Then         ' previous char was whitespace
                    If IsEndOfLine(gSourceFile.Data(wdEnd + 1)) = 1 Then
                        wdEnd = lOffset + 1: Exit For
                    End If
                End If
                lOffset = wdEnd
            ElseIf IsWhiteSpace(c) = 0 Then
                lOffset = wdEnd
            End If
        ElseIf c = vbKeyBracket Then                ' increase container count
            lFlags = lFlags + 1
        ElseIf c = vbKeyBracket2 Then               ' square bracket container end?
            lFlags = lFlags - 1
            If (WordBreakOptions And wbpMarkBrackets) <> 0 Then
                If (lFlags Or lParents) = 0 Then
                    WordBreakOptions = WordBreakOptions Or wbpBracketsMarked
                    gSourceFile.Data(lBrkOffset) = vbKeyLineFeed
                    gSourceFile.Data(wdEnd) = vbKeyLineFeed
                End If
            End If
        End If
    Next
    
End Sub

Public Sub ParseNextWord(ByVal lStart As Long, ByVal lMax As Long, _
                          wdStart As Long, wdEnd As Long, _
                          pFlags As WordBreakOptEnum)
                            
    ' Routine parses words from a previously parsed statement
    ' This routine is different from ParseNextWordEx in the following ways
    ' - less options, more efficient for validation routines
    ' - grouped parentheses always returned with their contained code
    ' - bracketed identifiers have brackets changed to vbLf for replacement
    '       in addition, pFlags set to fBrkIdentifier in this case
    '   otherwise, brackets are returned
    ' - only item prefixed with hash tag will be a date literal
    ' - named parameters returned with trailing colon
    ' - operands are ignored, as are commas, semicolons
    ' - ampersand ignored if not part of a word
    
    Dim lFlags As StatementParseFlags
    Dim c%, lParents&, lOffset&, lBrkOffset&
    
    For wdStart = lStart To lMax - 1                ' look for start of a word
        c = gSourceFile.Data(wdStart)
        If c <> vbKeySpace Then
            If IsWhiteSpace(c) = 0 Then
                If IsEndOfLine(c) = 0 Then
                    Select Case c
                    Case 42 To 45, 60 To 62, 47, 92, 94 ' *+,-=/\^<>
                    Case vbKeySemicolon, vbKeyParenthesis2
                    Case 38                         ' &
                        If IsWhiteSpace(gSourceFile.Data(wdStart) + 1) = 0 Then Exit For
                        wdStart = wdStart + 1
                    Case vbKeyQuote
                        lFlags = fLiteralStr: Exit For
                    Case vbKeyBracket
                        lFlags = 1: lBrkOffset = wdStart: Exit For
                    Case vbKeyHash
                        If pvIsDateLiteral((wdStart)) = 0 Then
                            wdStart = wdStart + 1: Exit For
                        Else
                            lOffset = 8: lFlags = fLiteralDt: Exit For
                        End If
                    Case vbKeyUnderscore
                        If IsEndOfLine(gSourceFile.Data(wdStart + 1)) <> 1 Then Exit For
                    Case vbKeyParenthesis
                        lParents = 1: Exit For
                    Case Else: Exit For             ' any other character is start of a "word"
                    End Select
                End If
            End If
        End If
    Next
    If wdStart = lMax Then                          ' abort if no word start found
        wdEnd = lMax: Exit Sub
    End If
    
    For wdEnd = wdStart + lOffset + 1 To lMax - 1     ' now find end of word
        c = gSourceFile.Data(wdEnd)
        If lFlags = 0 Then                          ' else special handling for containers
            Select Case c
            Case 42 To 45, 92, 47, 94, vbKeySemicolon, vbKeyColon ' *+,-\/^;:
                If lParents = 0 Then Exit For
            Case vbKeyQuote: lFlags = fLiteralStr
            Case vbKeyBracket: lFlags = 1: lBrkOffset = wdEnd
            Case vbKeyParenthesis
                If lParents = 0 Then Exit For
                lParents = lParents + 1
            Case vbKeyParenthesis2
                lParents = lParents - 1
                If lParents = 0 Then
                    If gSourceFile.Data(wdStart) = vbKeyParenthesis Then wdEnd = wdEnd + 1
                    Exit For
                End If
            Case Else
                If lParents = 0 Then
                    If IsWhiteSpace(c) = 1 Then Exit For
                End If
            End Select
        ElseIf lFlags = fLiteralStr Then            ' string literal container end?
            If c = vbKeyQuote Then lFlags = 0
        ElseIf lFlags = fLiteralDt Then             ' end of date literal
            If c = vbKeyHash Then lFlags = 0
        ElseIf c = vbKeyBracket Then                ' increase container count
            lFlags = lFlags + 1
        ElseIf c = vbKeyBracket2 Then               ' square bracket container end?
            lFlags = lFlags - 1
            If (pFlags And wbpMarkBrackets) <> 0 Then
                If (lFlags Or lParents) = 0 Then
                    pFlags = pFlags Or wbpBracketsMarked
                    gSourceFile.Data(lBrkOffset) = vbKeyLineFeed
                    gSourceFile.Data(wdEnd) = vbKeyLineFeed
                End If
            End If
        End If
    Next
    
End Sub

Private Function pvParseName_EnumMbr(ByVal lStart As Long, lMax As Long) As ItemTypeEnum

    Dim n&, p&, lOffset&, lParent&
    Dim sName$, pFlags&
    
    p = lStart: pFlags = wbpEqualEOW Or wbpMarkBrackets
    ParseNextWordEx lStart, lMax, n, p, pFlags ' get the enum member name or End Enum
    If gSourceFile.Data(n) = vbKeyE Then
        If p - n = 3 Then                   ' looking for "End"
            If Mid$(gSourceFile.Text, n, 3) = "End" Then
                gParsedItems.Find SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , adSearchBackward
                gParsedItems.Fields(recOffset2).Value = p
                gParsedItems.Update
                Exit Function
            End If
        End If
    End If
                                            ' get position after equal sign
    ParseNextWordEx p, lMax, n, 0&, pFlags
    If gParsedItems.Fields(recType).Value = itEnumMember Then
        lParent = gParsedItems.Fields(recParent).Value
    Else
        lParent = gParsedItems.Fields(recID).Value
    End If
    lOffset = IsVarTyped(gSourceFile.Data(p - 1))
    sName = Mid$(gSourceFile.Text, lStart, p - lStart - lOffset)
    If (pFlags And wbpBracketsMarked) <> 0 Then
        sName = Replace(sName, vbLf, vbNullString)
    End If
    CreateRecord lParent, sName, itEnumMember, lStart, lMax, , , , , n, , scpGlobal, chrZ
    pvParseName_EnumMbr = itEnumMember

End Function

Private Sub pvParseName_Misc(lParent As Long, blkStart As Long, _
                            lStart As Long, lMax As Long, _
                            cType As ItemTypeEnum, lScope As ItemScopeEnum, isSplit As Boolean)

    ' on entrance, lStart is next character after word: Type,Enum,Implements
    ' ex: Enum HitResultEnum
    '     Type RECT
    '     Implements ISubclass
    
    ' easy to get name, it is always last word of statement & can't be arrayed
    
    Dim n&, p&, lFlags&, pFlags&, sName$
    
    If isSplit = True Then lFlags = iaSplitIdent
    pFlags = wbpMarkBrackets
    ParseNextWordEx lStart, lMax, p, n, pFlags
'    On Error GoTo exitRoutine
    sName = Mid$(gSourceFile.Text, p, n - p)
    If (pFlags And wbpBracketsMarked) <> 0 Then
        sName = Replace(sName, vbLf, vbNullString)
    End If
    If cType = itImplements Then
        CreateRecord lParent, "(Implements)", cType, blkStart, lMax, sName, , , lFlags, , , lScope
    Else
        CreateRecord lParent, sName, cType, blkStart, lMax, , , , lFlags, , , lScope, chrZ
    End If
    gParsedItems.Update

exitRoutine:
    On Error GoTo 0

End Sub

Private Sub pvParseName_Method(ByVal blkStart As Long, ByVal lStart As Long, lMax As Long, _
                            ByVal cType As ItemTypeAttrEnum, lScope As ItemScopeEnum, isSplit As Boolean)

    ' on entrance, lStart is next character after word: Sub,Function,Property
    '   or if a public event, then after word: Event
    ' ex: MethodName(), MethodName(params)
    
    Dim p&, n&, lType&, lCrc&, lOffset&, lFlags&
    Dim pFlags&, sName$, sDiscrep$
    
    If (gSourceFile.Owner.FileAttrs And iaMaskCodePage) = iaBAS Then
        sDiscrep = chrZ
    ElseIf lScope = scpPrivate Or cType = itEvent Then
        sDiscrep = chrZ
    End If
    
    If isSplit = True Then lFlags = iaSplitIdent
    pFlags = wbpParentsCntr Or wbpMarkBrackets
    ParseNextWordEx lStart, lMax, n, p, pFlags
'    On Error GoTo exitRoutine
    If cType = itEvent Then
        sName = Mid$(gSourceFile.Text, n, p - n)
        lType = cType: cType = 0
    Else
        If cType = iaProperty Then      ' this word is Get,Let,Set
            If gSourceFile.Data(n) = vbKeyG Then
                cType = iaPropGet: sDiscrep = sDiscrep & chrV
            ElseIf gSourceFile.Data(n) = vbKeyL Then
                cType = iaPropLet
            Else
                cType = iaPropSet
            End If                      ' move to property name
            ParseNextWordEx p, lMax, n, p, pFlags
        ElseIf cType = iaFunction Then
            sDiscrep = sDiscrep & chrV
        End If
        lOffset = IsVarTyped(gSourceFile.Data(p - 1))
        sName = Mid$(gSourceFile.Text, n, p - n - lOffset)
        If (pFlags And wbpBracketsMarked) <> 0 Then
            pFlags = pFlags Xor wbpBracketsMarked
            sName = Replace(sName, vbLf, vbNullString)
        End If
        If lOffset <> 0 Then sDiscrep = Replace(sDiscrep, chrV, vbNullString)
        If cType < 0 Then lCrc = CRCItem(sName & chrHash & gSourceFile.Owner.RecordID, True)
        lType = itMethod
    End If
    lStart = p: pFlags = pFlags Xor wbpMarkBrackets
    
    If cType = iaSub Then
        If (gSourceFile.Owner.FileAttrs And iaMaskCodePage) = iaBAS Then
            If Len(sName) = 4 Then
                If LCase$(sName) = "main" Then lCrc = vbKeyM
            End If
        End If
    Else
        If LCase$(sName) = "left" Then
            ParseNextWord p, lMax, n, 0&, wbpFirstChar    ' move to beginning of parameters
            If gSourceFile.Data(n) = vbKeyParenthesis Then
                ParseNextWordEx n + 1, lMax, 0&, n, wbpFirstChar
                If gSourceFile.Data(n) <> vbKeyParenthesis2 Then
                    lFlags = lFlags Or iaLeftParams
                End If
            End If
        End If
        If InStr(sDiscrep, chrV) <> 0 Then
            ParseNextWord p, lMax, n, p, 0&     ' move past params
            If p <> lMax Then
                ParseNextWord p, lMax, n, p, 0& ' get next word, if any
                If gSourceFile.Data(n) = vbKeyA Then ' looking for: As
                    If gSourceFile.Data(n + 1) = 115 And p - n = 2 Then
                        sDiscrep = Replace(sDiscrep, chrV, vbNullString)
                    End If
                End If
            End If
            If InStr(sDiscrep, chrV) <> 0 Then
                If gSourceFile.Owner.IsDefTyped(AscW(sName)) = True Then
                    sDiscrep = Replace(sDiscrep, chrV, vbNullString)
                End If
            End If
        End If
    End If
    
    CreateRecord gSourceFile.Owner.RecordID, sName, lType, blkStart, lMax, , , lCrc, cType Or lFlags, lStart, , lScope, sDiscrep

exitRoutine:
    On Error GoTo 0

End Sub

Private Sub pvParseName_API(blkStart As Long, lStart As Long, lMax As Long, _
                            lScope As ItemScopeEnum, isSplit As Boolean)
                           
    ' on entrance, lStart is next character after word: Declare
    ' ex: Declare APIname Lib "user32.dll" () As ...
    '     Declare APIname Lib "user32.dll" Alias "exportedName" () As ...

    Dim p&, lType&, lMode&, lOffset&, lParams&, pFlags&
    Dim sName$, sAlias$, sDLL$, sDiscrep$
    
'    On Error GoTo exitRoutine
    p = lStart
    Do
        ParseNextWordEx p, lMax, lStart, p, pFlags
        If lMode = 0 Then               ' looking for method type
            Select Case Mid$(gSourceFile.Text, lStart, p - lStart) ' get method type
                Case ParseTypeFnc: lType = iaFunction: lMode = 1: sDiscrep = "VZ"
                Case ParseTypeSub: lType = iaSub: lMode = 1: sDiscrep = chrZ
            End Select
            pFlags = wbpParentsCntr Or wbpMarkBrackets
        
        ElseIf lMode = 1 Then           ' looking for API/method name
            lOffset = IsVarTyped(gSourceFile.Data(p - 1))
            If lOffset = 0 Then
                sName = Mid$(gSourceFile.Text, lStart, p - lStart)
            Else
                sName = Mid$(gSourceFile.Text, lStart, p - lStart - 1)
                sDiscrep = chrZ
            End If
            If (pFlags And wbpBracketsMarked) <> 0 Then
                pFlags = pFlags Xor wbpBracketsMarked
                sName = Replace(sName, vbLf, vbNullString)
            End If
            pFlags = wbpParentsCntr: lMode = 2
            
        ElseIf lMode = 2 Then           ' looking for dll name
            If gSourceFile.Data(lStart) = vbKeyQuote Then
                sDLL = Mid$(gSourceFile.Text, lStart + 1, p - lStart - 2) ' strip out DLL
                If InStr(sDLL, chrDot) = 0 Then sDLL = sDLL & ".dll"
                lStart = InStr(sDLL, chrSlash)
                If lStart <> 0 Then sDLL = Mid$(sDLL, lStart + 1)
                lMode = 3
            End If
            
        ' looking for alias, if any. If exists will be a string literal
        ' and if not exists, required parentheses will be parsed
        Else
            If gSourceFile.Data(lStart) = vbKeyQuote Then
                sAlias = Mid$(gSourceFile.Text, lStart + 1, p - lStart - 2)
                lParams = p
                If LenB(sDiscrep) = 2 Then ParseNextWord p, lMax, lStart, p, 0&
            ElseIf gSourceFile.Data(lStart) = vbKeyParenthesis Then
                sAlias = sName: lParams = lStart
            End If
            If lParams <> 0 Then
                If Len(sDiscrep) = 2 Then
                    ParseNextWord p, lMax, lStart, p, 0&
                    If lStart < p Then
                        sDiscrep = chrZ
                    ElseIf gSourceFile.Owner.IsDefTyped(AscW(sName)) = True Then
                        sDiscrep = chrZ
                    End If
                End If
                Exit Do
            End If
        End If
    Loop While p < lMax
    
    sDLL = LCase$(sDLL)
    p = CRCItem(sDLL, True)
    If isSplit = True Then lMode = iaSplitIdent Else lMode = 0
    CreateRecord gSourceFile.Owner.RecordID, sName, itAPI, blkStart, lMax, sAlias, sDLL, p, lMode, lParams, , lScope, sDiscrep

exitRoutine:
    On Error GoTo 0
    
End Sub

Private Sub pvParseName_Constant(lParent As Long, blkStart As Long, _
                            ByVal lStart As Long, lMax As Long, _
                            lScope As ItemScopeEnum, isSplit As Boolean)
                               
    ' on entrance, lStart is next character after word: Const
    ' ex: WM_PAINT = 5
    '     WM_PAINT = 5, WM_DESTROY = 2
    '     MODNAME = "modMain"
    '     WM_PAINT As Long = 5
    '     xMASK = (2 Or 8 Or 16)
    
    Dim p&, n&, lMode&, lFlags&, pFlags&
    Dim sName$, sValue$
    
'    On Error GoTo exitRoutine
    p = lStart: pFlags = wbpParentsCntr Or wbpMarkBrackets
    If isSplit = True Then lFlags = iaSplitIdent
    Do
        ParseNextWordEx p, lMax, n, p, pFlags
        If lMode = 0 Then           ' looking for name
            If IsVarTyped(gSourceFile.Data(p - 1)) = 0 Then
                sName = Mid$(gSourceFile.Text, n, p - n)
            Else
                sName = Mid$(gSourceFile.Text, n, p - n - 1)
            End If
            If (pFlags And wbpBracketsMarked) <> 0 Then
                pFlags = pFlags Xor wbpBracketsMarked
                sName = Replace(sName, vbLf, vbNullString)
            End If
            lMode = 1: lStart = p: pFlags = pFlags Xor wbpMarkBrackets
        ElseIf lMode = 1 Then
            If gSourceFile.Data(n) = vbKeyEqual Then
                lStart = p + 1: lMode = 2: pFlags = pFlags Or wbpNotCommaEOW
            End If
        ElseIf lMode = 2 Then
            If gSourceFile.Data(p - 1) = vbKeyComma Then
                lMode = 0
            ElseIf gSourceFile.Data(n) = vbKeyComma Then
                lMode = 0
            ElseIf p = n Then
                lMode = 0
            End If
            If lMode = 0 Then
                CreateRecord lParent, sName, itConstant, blkStart, lMax, , , , lFlags, lStart, p, lScope, chrZ
                If p = n Then Exit Do
                pFlags = (pFlags Or wbpMarkBrackets) Xor wbpNotCommaEOW
            End If
        End If
    Loop
    
exitRoutine:
    On Error GoTo 0
End Sub

Private Sub pvParseName_Variable(lParent As Long, blkStart As Long, _
                            ByVal lStart As Long, lMax As Long, _
                            lScope As ItemScopeEnum, isSplit As Boolean)

    ' on entrance, lStart is next character after word: Dim,Public,Private,Global
    ' ex: Dim n As Long
    '     Dim v As Double, i, j, WithEvents oForm As Form
    '     Dim x() As Long
    '     Dim y(0 To 3, 0 To 20) As Byte
    
    Dim p&, n&, sName$, sObj$, sDiscrep$
    Dim lOffset&, lWEvents&, lMode&, lFlags&, pFlags&
    Const ParseTypeWevt = "WithEvents"
    
'    On Error GoTo exitRoutine
    If isSplit = True Then lFlags = iaSplitIdent
    pFlags = wbpNotCommaEOW Or wbpParentsCntr Or wbpMarkBrackets
    p = lStart
    Do
        ParseNextWordEx p, lMax, n, p, pFlags
        If n = p Then Exit Do
        
        If lMode = 0 Then
            If gSourceFile.Data(n) = vbKeyW And p - n = 10 Then
                sName = Mid$(gSourceFile.Text, n, p - n)
                If sName = ParseTypeWevt Then
                    lWEvents = iaWithEvents
                    ParseNextWordEx p, lMax, n, p, pFlags
                End If
            End If
            If gSourceFile.Data(p - 1) = vbKeyComma Then         ' ex: Dim X, Y
                lOffset = IsVarTyped(gSourceFile.Data(p - 2))
                sName = Mid$(gSourceFile.Text, n, p - n - 1 - lOffset)
                lMode = 3: lStart = p - 1
            Else
                lOffset = IsVarTyped(gSourceFile.Data(p - 1))
                sName = Mid$(gSourceFile.Text, n, p - n - lOffset)
                lMode = 2: lStart = p
            End If
            If (pFlags And wbpBracketsMarked) <> 0 Then
                pFlags = pFlags Xor wbpBracketsMarked
                sName = Replace(sName, vbLf, vbNullString)
            End If
            pFlags = pFlags Xor wbpMarkBrackets
        ElseIf lMode = 2 Then   ' looking for comma
            If gSourceFile.Data(n) = vbKeyComma Then
                lMode = 3
            ElseIf gSourceFile.Data(p - 1) = vbKeyComma Then
                lMode = 3
            ElseIf gSourceFile.Data(n) = vbKeyA Then
                If p - n = 2 And gSourceFile.Data(n + 1) = 115 Then
                    lOffset = 1
                    If (lWEvents And iaWithEvents) <> 0 Then
                        ParseNextWordEx p, lMax, n, p, wbpNotCommaEOW Or wbpParentsCntr
                        If gSourceFile.Data(p - 1) = vbKeyComma Then
                            sObj = Mid$(gSourceFile.Text, n, p - n - 1)
                            lMode = 3
                        Else
                            sObj = Mid$(gSourceFile.Text, n, p - n)
                        End If
                    End If
                End If
                If p = lMax Then lMode = 3
            ElseIf p = lMax Then
                lMode = 3
            End If
        End If
        If lMode = 3 Then
            If lOffset = 0 Then
                If gSourceFile.Owner.IsDefTyped(AscW(sName)) = True Then
                    sDiscrep = chrZ
                Else
                    sDiscrep = "VZ"
                End If
            Else
                sDiscrep = chrZ
            End If
            If lScope = scpPublic Then  ' public in non-bas module else would be global
                sDiscrep = Replace(sDiscrep, chrZ, vbNullString)
            End If
            CreateRecord lParent, sName, itVariable, blkStart, lMax, sObj, , , lWEvents Or lFlags, lStart, p, lScope, sDiscrep
            lWEvents = 0: sObj = vbNullString: lMode = 0
            pFlags = pFlags Or wbpMarkBrackets
        End If
    Loop
    
exitRoutine:
    On Error GoTo 0
End Sub

Private Sub pvParseName_Option(lStart As Long, lMax As Long)

    ' on entrance, lStart is next character after word: Option
    ' ex: Option Explicit, Option Compare Text, Option Base 1
    
    Dim lAttr&
    
'    On Error GoTo exitRoutine
    With gParsedItems
        .Bookmark = gSourceFile.CPBookMark
        lAttr = .Fields(recFlags).Value
        ParseNextWordEx lStart, lMax, lStart, 0&, wbpFirstChar
        Select Case gSourceFile.Data(lStart)
        Case vbKeyE: lAttr = lAttr Or iaOpExplicit
        Case vbKeyC: lAttr = lAttr Or iaOpText ' fyi: Compare Database is VBA not VB
        Case vbKeyB: lAttr = lAttr Or iaOpBase1
        Case vbKeyP: lAttr = lAttr Or iaOpPrivate
        End Select
        .Fields(recFlags).Value = lAttr
        .Update
        If IsEmpty(gSourceFile.ItemBookMark) = False Then .Bookmark = gSourceFile.ItemBookMark
    End With
    
exitRoutine:
    On Error GoTo 0
End Sub

Private Sub pvParseName_DEF(lStart As Long, lMax As Long)
                               
    ' on entrance, lStart begins the Def[xxx] statement
    ' ex: DefInt I-K
    '     DefLng A, E, X-Z
    ' routine compresses/normalizes the statement
    ' i.e. "DefLng:A,E,X-Z" returned from something like this:
    '   DefLng A, _
    '       E, X _
    '       - _
    '       Z
    
    Dim p&, n&, sName$
    
'    On Error GoTo exitRoutine
    ParseNextWordEx lStart, lMax, n, p
    sName = Mid$(gSourceFile.Text, lStart, p - lStart) & chrColon
    Do
        ParseNextWordEx p + 1, lMax, lStart, p, wbpNotCommaEOW
        sName = sName & Mid$(gSourceFile.Text, lStart, p - lStart)
    Loop While p < lMax
    CreateRecord gSourceFile.Owner.RecordID, "(DefType)", itDefType, lStart, lMax, sName
    
exitRoutine:
    On Error GoTo 0
End Sub

Public Function IsWhiteSpace(ByVal c As Integer) As Long
    If c <> vbKeySpace Then
        Select Case c
        Case 0 To &H1F
        Case &HA0, &H3000, &H202F, &H205F
        Case &H1680, &H2000 To &H200A
        Case Else: Exit Function
        End Select
    End If
    IsWhiteSpace = 1
End Function

Public Function IsEndOfLine(ByVal c As Integer) As Long
    If c <> vbKeyReturn Then
        If c <> vbKeyLineFeed Then
            If c <> &H2028 Then
                If c <> &H2029 Then Exit Function
            End If
        End If
    End If
    IsEndOfLine = 1
End Function

Private Function pvIsDateLiteral(n As Long) As StatementParseFlags

    If n + 8 < gSourceFile.length Then
        ' next 3 characters will fit one of these patterns: #/# or ##/
        If gSourceFile.Data(n + 1) > vbKeySlash Then             ' numeric needed
            If gSourceFile.Data(n + 1) < vbKeyColon Then
                If gSourceFile.Data(n + 2) = vbKeySlash Then     ' #/# pattern check
                    If gSourceFile.Data(n + 3) > vbKeySlash Then ' numeric needed
                        If gSourceFile.Data(n + 3) < vbKeyColon Then _
                            pvIsDateLiteral = fLiteralDt: n = n + 8
                    End If
                ElseIf gSourceFile.Data(n + 2) > vbKeySlash Then ' ##/ pattern check
                    If gSourceFile.Data(n + 2) < vbKeyColon Then ' numeric needed
                        If gSourceFile.Data(n + 3) = vbKeySlash Then _
                            pvIsDateLiteral = fLiteralDt: n = n + 8
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function IsVarTyped(ByVal c As Integer) As Long

    ' tests passed character value for known vartype sybmols

    Select Case c
    Case vbKeyBang, vbKeyHash To 38, 64, 94     ' !,#,%,^,&,@,$
        IsVarTyped = 1
    End Select
    
End Function

Private Function pvAddDesignerEvents(sClass As String) As ItemTypeAttrEnum

    ' Designers events are customized by VB to include Initialize,Terminate
    '   events with some events from an implemented TLB class
    ' This routine hardcodes those static events for Event Resolution

    Const EvtInitialize = "Initialize"
    Const EvtTerminate = "Terminate"

    Select Case UCase$(sClass)
    Case "{C0E45035-5775-11D0-B388-00A0C9055D8E}"
        pvAddDesignerEvents = iaDataEnv     ' requires special parsing for variables
        sClass = "DEDesignerObjects.DataEnvironment" ' actual implemented class
        If m_RefsEvents.AddCustomReference(sClass, rstCustom) = True Then
            m_RefsEvents.AddCustomEvents sClass, _
                    Array(EvtInitialize, EvtTerminate)
        End If
    Case "{AC0714F6-3D04-11D1-AE7D-00A0C90F26F4}"
        sClass = "AddInDesignerObjects.AddinInstance" ' actual implemented class
        If m_RefsEvents.AddCustomReference(sClass, rstCustom) = True Then
            m_RefsEvents.AddCustomEvents sClass, _
                    Array("OnAddInsUpdate", "OnBeginShutdown", "OnConnection", _
                          "OnDisconnection", "OnStartupComplete", EvtInitialize, EvtTerminate)
        End If
    Case "{78E93846-85FD-11D0-8487-00A0C90DC8A9}"
        sClass = "MSDataReportLib.DataReport"       ' actual implemented class
        If m_RefsEvents.AddCustomReference(sClass, rstCustom) = True Then
            m_RefsEvents.AddCustomEvents sClass, _
                    Array("Activate", "AsyncProgress", "Deactivate", _
                          "Error", EvtInitialize, "ProcessingTimeout", _
                          "QueryClose", "Resize", EvtTerminate)
        End If
    Case "{90290CCD-F27D-11D0-8031-00C04FB6C701}"
        Dim lID&
        sClass = "MSHMTL.DHTMLPage"                 ' actual implemented class
        ' DHTML designer has 2 hidden WithEvents variables
        lID = gParsedItems.Fields(recID).Value
        CreateRecord lID, "BaseWindow", itVariable, 0, 0, "MSHTML.HTMLWindow2", , , iaWithEvents Or iaHidden, , , scpPrivate
        CreateRecord lID, "Document", itVariable, 0, 0, "MSHTML.HTMLDocument", , , iaWithEvents Or iaHidden, , , scpPrivate
        If m_RefsEvents.AddCustomReference(sClass, rstCustom) = True Then
            m_RefsEvents.AddCustomEvents sClass, _
                    Array(EvtInitialize, "Load", EvtTerminate, "Unload")
        End If
    Case "{17016CEE-E118-11D0-94B8-00A0C91110ED}"
        pvAddDesignerEvents = iaWebClass    ' requires special parsing for variables
        sClass = "WebClassLibrary.WebClass" ' actual implemented class
        If m_RefsEvents.AddCustomReference(sClass, rstCustom) = True Then
            ' note: until I can parse these with full confidence, we will
            ' not perform Event Resolution on parsed WebItem variables and
            ' assume any event prefixed with the WebItem name is resolved.
            m_RefsEvents.AddCustomReference "WebClassLibrary.WebItem", rstNoLoad
            m_RefsEvents.AddCustomEvents sClass, _
                    Array("BeginRequest", "EndRequest", "FatalErrorResponse", _
                          "Start", EvtInitialize, EvtTerminate)
        End If
    End Select

End Function

Private Sub pvResolveEvents_Controls(ExternalProjs As Boolean, rsCodePg As ADODB.Recordset, _
                                    rsObjs As ADODB.Recordset, rsMethods As ADODB.Recordset, _
                                    rsEvents As ADODB.Recordset)
                                    
    ' Controls are always defined with Library.ClassName
    ' First step is to determine where the control orignates from:
    '   Does control belong to TLB?
    '       If so, we query the TLB for all possible control events
    '   If not, the Public Event declarations within the code file are used

    Dim sValue$, sObj$, sName$, sProject$
    Dim o&, n&, lRefID&, lEventID&, bEvent As Boolean
    
    gParsedItems.Bookmark = gSourceFile.ProjBookMark
    sProject = gParsedItems.Fields(recName).Value
    rsObjs.Sort = recIdxAttr
    Do
        sName = rsObjs.Fields(recName).Value    ' check for events
        rsMethods.Filter = SetQuery(recCodePg, qryIs, rsObjs.Fields(recCodePg).Value, qryAnd, _
                                recType, qryIs, itMethod, qryAnd, _
                                recName, qryLike, sName & chrPctWild)
        If rsMethods.EOF = False Then
            sValue = rsObjs.Fields(recAttr).Value
            If sObj <> sValue Then
                sObj = sValue
                o = InStr(sObj, chrDot)
                If o = 0 Then
                    lRefID = 0                  ' should never happen
                ElseIf Left$(sObj, o - 1) = sProject Or ExternalProjs = True Then
                    ' uncompiled usercontrol or external project's control
                    rsCodePg.Find SetQuery(recGrp, qryIs, CRCItem(LCase$(sObj), True)), , , 1&
                    If rsCodePg.EOF = False Then
                        If (rsCodePg.Fields(recFlags).Value And iaExternalProj) <> 0 Then
                            lRefID = 0
                        Else
                            rsEvents.Filter = SetQuery(recCodePg, qryIs, rsCodePg.Fields(recID).Value, _
                                                qryAnd, recType, qryIs, itEvent)
                            If rsEvents.EOF = True Then lRefID = 0 Else lRefID = -1
                        End If
                    Else                        ' check TLBs
                        lRefID = m_RefsEvents.GetFQN(sValue, tlbLoadEvents)
                    End If
                Else                            ' check TLBs
                    lRefID = m_RefsEvents.GetFQN(sValue, tlbLoadEvents)
                End If
            End If
            If lRefID = 0 Then                  ' can't resolve, assume all are events
                rsObjs.Fields(recFlags).Value = rsObjs.Fields(recFlags).Value Or iaUnresolved
                rsObjs.Update
            ElseIf LCase$(Left$(sObj, 3)) <> "vb." Then
                lEventID = m_RefsEvents.GetFQN("VB.VBControlExtenderEvents", tlbLoadEvents)
            End If
            n = Len(sName) + 2                  ' prefix of event name
            rsMethods.MoveLast                  ' loop from bottom up
            Do
                If lRefID = 0 Then
                    bEvent = True
                Else
                    If lRefID = -1 Then
                        rsEvents.Find SetQuery(recName, qryIs, Mid$(rsMethods.Fields(recName).Value, n)), , , 1&
                        bEvent = (rsEvents.EOF = False)
                    Else
                        bEvent = m_RefsEvents.IsEvent(Mid$(rsMethods.Fields(recName).Value, n), lRefID)
                    End If
                    If bEvent = False And lEventID <> 0 Then
                        bEvent = m_RefsEvents.IsEvent(Mid$(rsMethods.Fields(recName).Value, n), lEventID)
                    End If
                End If
                If bEvent = True Then
                    With rsMethods
                        .Fields(recParent).Value = rsObjs.Fields(recID).Value
                        .Fields(recDiscrep).Value = vbNullString
                        .Fields(recType).Value = itClassEvent
                        If lRefID = 0 Then .Fields(recFlags).Value = .Fields(recFlags).Value Or iaUnresolved
                        .Update
                    End With
                Else
                    ' faux event, i.e., Command1_TestSub is not a CommandButton event
                End If
                rsMethods.MovePrevious
            Loop Until rsMethods.BOF = True
        End If
        rsObjs.MoveNext
    Loop Until rsObjs.EOF = True

End Sub

Private Sub pvResolveEvents_Classes(rsObjs As ADODB.Recordset, rsMethods As ADODB.Recordset)

    ' All code files, except bas modules, are considered classes
    ' All classes are defined with fully qualified names (FQN)
    ' TLBs are always used to verify class events

    Dim sValue$, sObj$, sName$
    Dim o&, n&, lAttrs&, lRefID&
    Dim bEvent As Boolean
    
    rsObjs.Sort = recIdxAttr                ' sort by FQN
    Do
        lAttrs = rsObjs.Fields(recFlags).Value And iaMaskCodePage
        If lAttrs <> iaBAS Then             ' bas modules have no events
            sName = rsObjs.Fields(recAttr).Value
            o = InStr(sName, chrDot)
            If o = 0 Then
                sValue = sName              ' no FQN
            ElseIf lAttrs = iaClass Then
                sValue = chrClass            ' generic Class
                rsObjs.Fields(recAttr).Value = chrClassVB
                rsObjs.Update
            Else
                sValue = Mid$(sName, o + 1) ' class name from FQN
            End If
            rsMethods.Filter = SetQuery(recCodePg, qryIs, rsObjs.Fields(recCodePg).Value, qryAnd, _
                                    recType, qryIs, itMethod, qryAnd, _
                                    recName, qryLike, sValue & chrPctWild)
            If rsMethods.EOF = False Then   ' any class events to verify?
                If sObj <> sName Then
                    sObj = sName
                    n = Len(sValue) + 2     ' offset from name prefix
                    lRefID = m_RefsEvents.GetFQN(sName, tlbLoadEvents)
                    If lRefID = 0 Then      ' TLB not found or cannot be loaded
                        rsObjs.Fields(recFlags).Value = rsObjs.Fields(recFlags).Value Or iaUnresolved
                        rsObjs.Update
                    End If
                End If
                rsMethods.MoveLast          ' loop from bottom up
                Do
                    If lRefID = 0 Then
                        bEvent = True
                    Else                    ' verify from TLB
                        bEvent = m_RefsEvents.IsEvent(Mid$(rsMethods.Fields(recName).Value, n), lRefID)
                    End If
                    If bEvent = True Then
                        With rsMethods
                            .Fields(recParent).Value = rsObjs.Fields(recID).Value
                            .Fields(recDiscrep).Value = vbNullString
                            .Fields(recType).Value = itClassEvent
                            .Update
                        End With
                    Else
                        ' faux event, i.e., Class_Exit is not a Class event
                    End If
                    rsMethods.MovePrevious
                Loop Until rsMethods.BOF = True
            End If
        End If
        rsObjs.MoveNext
    Loop Until rsObjs.EOF = True

End Sub

Private Function pvResolveEvents_Implements(ExternalProjs As Boolean, rsCodePg As ADODB.Recordset, _
                                    rsObjs As ADODB.Recordset, rsMethods As ADODB.Recordset, _
                                    rsEvents As ADODB.Recordset) As Boolean

    ' Implement statements includes the object we want methods from
    ' The implemented object may not be fully qualified, i.e., Implements ISubclass
    ' First job is to determine whether implemented object is a project class
    '   or a TLB item. Check project code files first & assume TLB otherwise.

    Dim sValue$, sObj$, sName$, sProject$
    Dim o&, n&, lRefID&, bEvent As Boolean
    
    gParsedItems.Bookmark = gSourceFile.ProjBookMark
    sProject = gParsedItems.Fields(recName).Value
    rsObjs.Sort = recIdxAttr                    ' sort on implemented object
    Do
        sName = rsObjs.Fields(recAttr).Value
        o = InStr(sName, chrDot)                   ' extract name from FQN if applies
        If o = 0 Then sValue = sName Else sValue = Mid$(sName, o + 1)
        rsMethods.Filter = SetQuery(recCodePg, qryIs, rsObjs.Fields(recCodePg).Value, qryAnd, _
                                recType, qryIs, itMethod, qryAnd, _
                                recName, qryLike, sValue & chrPctWild)
        If rsMethods.EOF = False Then           ' any events to verify?
            If sObj <> sName Then
                sObj = sName: lRefID = 0
                If o = 0 Then                   ' unqualified name lookup needed
                    If ExternalProjs = False Then
                        sName = sProject & chrDot & sObj
                        rsCodePg.Find SetQuery(recGrp, qryIs, CRCItem(LCase$(sName), True)), , , 1&
                        If rsCodePg.EOF = False Then lRefID = -1
                    Else                        ' target can be external project
                        sName = chrDot & sObj: n = Len(sName)
                        rsCodePg.Find SetQuery(recName, qryLike, chrPct & sName & chrPct), , , 1&
                        Do Until rsCodePg.EOF  ' verify not partial hit, i.e., target is clsBase but hit is clsBaseX
                            If Right$(rsCodePg.Fields(recName).Value, n) = sName Then
                                lRefID = -1: Exit Do
                            End If
                            rsCodePg.Find SetQuery(recName, qryLike, chrPct & sName & chrPct), 1&
                        Loop
                    End If
                Else
                    rsCodePg.Find SetQuery(recGrp, qryIs, CRCItem(LCase$(sObj), True)), , , 1&
                    If rsCodePg.EOF = False Then lRefID = -1
                End If
                If lRefID = 0 Then              ' not project code file, test TLBs
                    lRefID = m_RefsEvents.GetFQN((sObj), tlbLoadMethods)
                ElseIf (rsCodePg.Fields(recFlags).Value And iaExternalProj) <> 0 Then
                    lRefID = 0
                Else                            ' get listing of possible events
                    rsEvents.Filter = SetQuery(recCodePg, qryIs, rsCodePg.Fields(recID).Value, _
                                        qryAnd, recType, qryIs, itMethod, _
                                        qryAnd, recScope, qryNot, scpPrivate)
                    If rsEvents.EOF = True Then
                        lRefID = 0
                    ElseIf (rsCodePg.Fields(recFlags).Value And iaImplemented) = 0 Then
                        rsCodePg.Fields(recFlags).Value = rsCodePg.Fields(recFlags).Value Or iaImplemented
                        rsCodePg.Update
                        Do
                            rsEvents.Fields(recFlags).Value = rsEvents.Fields(recFlags).Value Or iaImplemented
                            rsEvents.MoveNext
                        Loop Until rsEvents.EOF = True
                        pvResolveEvents_Implements = True
                    End If
                End If
                n = Len(sValue) + 2             ' prefix offset for event name
            End If
            If lRefID = 0 Then
                rsObjs.Fields(recFlags).Value = rsObjs.Fields(recFlags).Value Or iaUnresolved
                rsObjs.Update
            End If
            rsMethods.MoveLast                  ' loop from bottom up
            Do
                If lRefID = 0 Then
                    bEvent = True
                ElseIf lRefID = -1 Then
                    rsEvents.Find SetQuery(recName, qryIs, Mid$(rsMethods.Fields(recName).Value, n)), , , 1&
                    If rsEvents.EOF = True Then
                        bEvent = False
                    Else
                        bEvent = (rsEvents.Fields(recScope).Value <> scpFriend)
                    End If
                Else
                    bEvent = m_RefsEvents.IsEvent(Mid$(rsMethods.Fields(recName).Value, n), lRefID)
                End If
                If bEvent = True Then
                    With rsMethods
                        .Fields(recParent).Value = rsObjs.Fields(recID).Value
                        .Fields(recDiscrep).Value = vbNullString
                        .Fields(recType).Value = itClassEvent
                        .Update
                    End With
                Else
                    ' faux event, i.e., Command1_TestSub is not a CommandButton event
                End If
                rsMethods.MovePrevious
            Loop Until rsMethods.BOF = True
        End If
        rsObjs.MoveNext
    Loop Until rsObjs.EOF = True

End Function

Private Sub pvResolveEvents_WithEvents(ExternalProjs As Boolean, rsCodePg As ADODB.Recordset, _
                                    rsObjs As ADODB.Recordset, rsMethods As ADODB.Recordset, _
                                    rsEvents As ADODB.Recordset)

    ' WithEvents statements includes the object we want events from
    ' The target object may not be fully qualified, i.e., Dim WithEvents f As Form
    ' First job is to determine whether target object is a project class
    '   or a TLB item. Check project code files first & assume TLB otherwise.

    Dim sObj$, sName$, sProject$
    Dim o&, n&, lRefID&, bEvent As Boolean
    
    gParsedItems.Bookmark = gSourceFile.ProjBookMark
    sProject = gParsedItems.Fields(recName).Value
    rsObjs.Sort = recIdxAttr                    ' sort on implemented object
    Do
        sName = rsObjs.Fields(recName).Value    ' target object's name
        rsMethods.Filter = SetQuery(recCodePg, qryIs, rsObjs.Fields(recCodePg).Value, qryAnd, _
                                recType, qryIs, itMethod, qryAnd, _
                                recName, qryLike, sName & chrPctWild)
        If rsMethods.EOF = True Then            ' any events to verify?
            rsObjs.Fields(recDiscrep).Value = rsObjs.Fields(recDiscrep).Value & chrW
            rsObjs.Update
        Else
            n = Len(sName) + 2                  ' prefix offset for event names
            sName = rsObjs.Fields(recAttr).Value
            If sObj <> sName Then
                sObj = sName: lRefID = 0
                o = InStr(sName, chrDot)
                If o = 0 Then                   ' unqualified name lookup needed
                    If ExternalProjs = False Then
                        sName = sProject & chrDot & sObj
                        rsCodePg.Find SetQuery(recGrp, qryIs, CRCItem(LCase$(sName), True)), , , 1&
                        If rsCodePg.EOF = False Then lRefID = -1
                    Else                        ' target can be external project
                        sName = chrDot & sName: o = Len(sName)
                        rsCodePg.Find SetQuery(recName, qryLike, chrPct & sName & chrPct), , , 1&
                        Do Until rsCodePg.EOF  ' verify not partial hit, i.e., target is Form but hit is FormEx
                            If Right$(rsCodePg.Fields(recName).Value, o) = sName Then
                                lRefID = -1: Exit Do
                            End If
                            rsCodePg.Find SetQuery(recName, qryLike, chrPct & sName & chrPct), 1&
                        Loop
                    End If
                Else
                    rsCodePg.Find SetQuery(recGrp, qryIs, CRCItem(LCase$(sObj), True)), , , 1&
                    If rsCodePg.EOF = False Then lRefID = -1
                End If
                If lRefID = 0 Then              ' not project code file, test TLBs
                    lRefID = m_RefsEvents.GetFQN((sObj), tlbLoadEvents)
                ElseIf (rsCodePg.Fields(recFlags).Value And iaExternalProj) <> 0 Then
                    lRefID = 0
                Else                            ' else get listing of possible events
                    rsEvents.Filter = SetQuery(recCodePg, qryIs, rsCodePg.Fields(recID).Value, _
                                            qryAnd, recType, qryIs, itEvent)
                    If rsEvents.EOF = True Then lRefID = 0
                End If
            End If
            If lRefID = 0 Then
                rsObjs.Fields(recFlags).Value = rsObjs.Fields(recFlags).Value Or iaUnresolved
                rsObjs.Update
            End If
            rsMethods.MoveLast                  ' loop from bottom up
            Do
                If lRefID = 0 Then
                    bEvent = True
                ElseIf lRefID = -1 Then
                    rsEvents.Find SetQuery(recName, qryIs, Mid$(rsMethods.Fields(recName).Value, n)), , , 1&
                    bEvent = (rsEvents.EOF = False)
                Else
                    bEvent = m_RefsEvents.IsEvent(Mid$(rsMethods.Fields(recName).Value, n), lRefID)
                End If
                If bEvent = True Then
                    With rsMethods
                        .Fields(recParent).Value = rsObjs.Fields(recID).Value
                        .Fields(recDiscrep).Value = vbNullString
                        .Fields(recType).Value = itClassEvent
                        .Update
                    End With
                Else
                    ' faux event, i.e., Form_Exit is not a VB.Form event
                End If
                rsMethods.MovePrevious
            Loop Until rsMethods.BOF = True
        End If
        rsObjs.MoveNext
    Loop Until rsObjs.EOF = True

End Sub

Public Function FindWord(wordList As WordListStruct, _
                         sCriteria As String, CompareMode As VbCompareMethod, _
                         bAutoAppend As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must be kept sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    ' note: in this routine, passed array must be one-bound, not zero-bound
    
    Dim UB&, LB&, newIndex&
    
    If LenB(sCriteria) = 0 Then Exit Function
    
    If wordList.Count = 0 Then         ' initialize as neeed
        If bAutoAppend Then
            wordList.Count = 1
            ReDim wordList.List(1 To 10)
            wordList.List(1) = sCriteria
        End If
    Else
        With wordList
            LB = 1: UB = .Count         ' set 1-bound range
            Do
                newIndex = LB + ((UB - LB) \ 2)
                Select Case StrComp(sCriteria, .List(newIndex), CompareMode)
                Case 0              ' match & done
                    FindWord = newIndex: Exit Do
                Case Is < 0         ' criteria is lower in sort order
                    UB = newIndex - 1
                Case Else           ' criteria is higher in sort order
                    LB = newIndex + 1
                End Select
            Loop Until LB > UB
            
            If FindWord = 0 Then
                If bAutoAppend = True Then ' insert word into sorted list
                    If newIndex < LB Then newIndex = newIndex + 1
                    .Count = .Count + 1
                    If .Count > UBound(.List) Then
                        ReDim Preserve .List(LBound(.List) To .Count + 25)
                    End If
                    If newIndex < .Count Then
                        CopyMemory ByVal VarPtr(.List(newIndex + 1)), _
                                   ByVal VarPtr(.List(newIndex)), (.Count - newIndex) * 4&
                        CopyMemory ByVal VarPtr(.List(newIndex)), 0&, 4& ' ensure StrPtr removed
                    End If
                    .List(newIndex) = sCriteria
                End If
            End If
        End With
    End If
    
End Function

Public Sub GetDefTypes(lCpg As Long, aTypes() As Byte)

    Dim sTypes$(), p&, n&
    Dim rs As ADODB.Recordset
    
    Set rs = gParsedItems.Clone
    rs.Filter = SetQuery(recType, qryIs, itDefType, qryAnd, recCodePg, qryIs, lCpg)
    ReDim aTypes(0 To 25)
    Do Until rs.EOF = True
        p = InStr(rs.Fields(recAttr).Value, chrColon)
        sTypes() = Split(Mid$(rs.Fields(recAttr).Value, p + 1), chrComma)
        For n = 0 To UBound(sTypes)
            If LenB(sTypes(n)) = 1 Then
                aTypes(AscW(sTypes(n)) - 65) = 1
            Else
                For p = AscW(sTypes(n)) To AscW(Mid$(sTypes(n), 3))
                    aTypes(p - 65) = 1
                Next
            End If
        Next
        rs.MoveNext
    Loop
    Erase sTypes()
    rs.Close: Set rs = Nothing

End Sub

Public Function ProcessCommandLine(Parameters As String) As Boolean
    
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    Dim lSize As Long, lPtr As Long
    lPtr = PathGetArgsW(GetCommandLineW)
    If lPtr Then
        lSize = lstrlenW(lPtr): Parameters = Space$(lSize)
        CopyMemory ByVal StrPtr(Parameters), ByVal lPtr, lSize * 2
        Parameters = Trim$(Parameters)
        If LenB(Parameters) <> 0 Then
            If AscW(Parameters) = vbKeyQuote Then
                Parameters = Mid$(Parameters, 2, Len(Parameters) - 2)
                ProcessCommandLine = True
            Else
                ProcessCommandLine = (InStr(Parameters, " ") = 0)
            End If
        End If
    End If
    
End Function

'///////////////////////////////////////////////////////////////////////////
'   Recordset detail usage per parsed item type
'   If the field is not annotated, then its value is not applicable

'   Fields recIdxName & recIdxAttr are custom (see ResolveEvents)
'   Field RecID is a simple incremental value, non-zero

'   Fields recType has the following values
'   itProject. The base project
'       recParent       0
'           recName     parsed from VBP file "Name" statement
'           recAttr     path/filename
'           recAttr2    Version;Startup Object
'           recOffset   file date (low part)
'           recOffset2  file date (high part)
'           recEnd      file size
'           recDiscrep  parsed from VBP file "Type" statement
'       recParent       -1
'           recAttr     group project file name
'           recOffset   file date (low part)
'           recOffset2  file date (high part)
'           recEnd      file size

'   itReference. External TLB reference or external project
'       recParent       0
'       recFlags        TLB/OCX: 0,-1 (unregistered)
'           recName     library name
'           recAttr     library path/filename
'           recAttr2    TLB GUID & version data
'       recFlags        VB project: iaExternalProj
'           recName     parsed from VBP file "Name" statement
'           recAttr     path/filename
'           recAttr2    Version
'           recOffset   file date (low part)
'           recOffset2  file date (high part)
'           recEnd      file size

'   itResFile,itHelpFile,itMiscFile. Project support files
'       recParent       base project RecID
'       recName         file name
'       recAttr         path/filename

'   itStats. Code page related statistics
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         hardcoded: (Stats)
'       recFlags        iaStatements, iaExclusions, or iaComments
'       recOffset       statistic's value

'   itSourceFile. Code files related to a project
'       recParent       base project RecID
'       recAttr2        file path/name
'       recFlags        code file type & iaExternalProj (if applies)
'       recOffset       file date (low part)
'       recOffset2      file date (high part)

'   itCodePage. Container record for parsed items in a source file
'       recParent       source file's RecID
'       recName         parsed from file "VB_Name" attribute
'                       result is prefixed with project's name, i.e., Project1.Form1
'       recAttr         Fully qualified name: Library.Class, i.e., VB.Form
'       recFlags        various ItemTypeAttrEnum values, negative if external project related
'       recOffset       end of Declarations section
'       recOffset2      start of Methods section
'       recStart        start of Declarations section
'       recEnd          end of Methods section (end of file)
'       recScope        scpGlobal
'       recDiscrep      initially: Z

'   itControl. Controls
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed from file "VB_Name" attribute
'       recAttr         fully qualified name, i.e., VB.PictureBox
'       recGrp          CRC of recName value
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       UBound if control is arrayed
'       recScope        scpPrivate

'   itDefType. Def[xxx] statements, i.e., DefInt I-N, P-R,Z
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         hardcoded: (DefType)
'       recAttr         formatted ex: DefInt:I-N,P-R,Z

'   itClassEvent. Methods that are events or implemented, i.e., Form_Unload
'       recCodePg       code page's RecID
'       recParent       RecID of event owner (Code Page,Implements,Variable,Control)
'       recName         parsed name from statement
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       start of parameters
'       recOffset2      end of method block
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value

'   itAPI. API declarations
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recAttr         API's Alias value if any, else recName value
'       recAttr2        API's DLL name, LCase()
'       recGrp          CRC of DLL name
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       start of parameters
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep      initially: Z if Sub else VZ

'   itConstant. Constant declarations, can be a comma-delimited multi-statement
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recGrp          validation only: CRC of name & value
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       postion after name
'       recOffset2      end of constant's statement
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep:     initially: Z

'   itEnum. Enumeration statements
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recGrp          validation only: CRC of enum name & all members
'       recFlags        various ItemTypeAttrEnum values
'       recOffset2      end of enum block
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep      initially: Z

'   itEnumMember. Enum members during validation only else not parsed in initial scan
'       recCodePg       code page's RecID
'       recParent       Enum's RecID
'       recName         parsed name from statement
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       postion after name
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        same as Enum record
'       recDiscrep      initially: Z

'   itEvent. Event statements found in Declarations section only
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       start of parameters
'       recStart        start of statement, negative if from external project
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep      initially: Z

'   itImplements. Implements statements
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         hardcoded: (Implements)
'       recAttr         object name/FQN that is implemented
'       recFlags        various ItemTypeAttrEnum values
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value

'   itMethod. Sub,Function,Property that is not an event
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recGrp          vbKeyM only if Sub Main, properties: CRC(codePage.MethodName)
'       recFlags        various ItemTypeAttrEnum values
'       recOffset       start of parameters
'       recOffset2      end of method block
'       recStart        start of statement, negative if from external project
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep      initially: Z if Sub else VZ

'   itParameter. Method/Event parameters, during validation only
'       recCodePg       code page's RecID
'       recParent       RecID of related method/event
'       recName         parsed name from statement
'       recFlags        various ItemTypeAttrEnum values
'       recStart        start of parameter statement
'       recEnd          end of paramter's statement
'       recScope        ItemScopeEnum value
'       recDiscrep:     initially: V

'   itType: Type/UDT statements
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recGrp          validation only: CRC of type name & all members
'       recFlags        various ItemTypeAttrEnum values
'       recOffset2      end of Type block
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep      initially: Z

'   itTypeMember. not currently used

'   itVariable. Variable declarations, can be a comma-delimited multi-statement
'       recCodePg       code page's RecID
'       recParent       code page's RecID
'       recName         parsed name from statement
'       recGrp          validation only: CRC of name & value
'       recFlags        various ItemTypeAttrEnum values, negative if WithEvents usage
'       recOffset       postion after variable name
'       recOffset2      end of variable statement
'       recStart        start of statement
'       recEnd          end of statement
'       recScope        ItemScopeEnum value
'       recDiscrep:     initially: VZ

'///////////////////////////////////////////////////////////////////////////
