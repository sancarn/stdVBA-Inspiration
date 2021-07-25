VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Project Scanner"
   ClientHeight    =   6945
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView tvScan 
      Height          =   6285
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   11086
      _Version        =   327682
      Indentation     =   392
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar sbarStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6570
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView tvValidation 
      Height          =   5670
      Left            =   120
      TabIndex        =   2
      Top             =   690
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   10001
      _Version        =   327682
      Indentation     =   392
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Rescan Project"
         Index           =   1
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Vali&dation"
      Index           =   1
      Begin VB.Menu mnuValidate 
         Caption         =   "Project &Validation"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuValidate 
         Caption         =   "&Malicious Code Safety Check Only"
         Index           =   1
         Begin VB.Menu mnuMalIntent 
            Caption         =   "&Standard"
            Index           =   0
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuMalIntent 
            Caption         =   "Include Registry &Reading APIs"
            Index           =   1
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnuValidate 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuValidate 
         Caption         =   "&Review Validation Report(s)"
         Index           =   3
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&View"
      Index           =   2
      Begin VB.Menu mnuView 
         Caption         =   "&Validation Descriptions "
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "Entire Code P&age"
         Index           =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Declarations Section"
         Index           =   3
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Methods Section"
         Index           =   4
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuView 
         Caption         =   "V&BP File"
         Index           =   6
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuView 
         Caption         =   "VB&G File"
         Enabled         =   0   'False
         Index           =   7
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Window"
      Index           =   3
      Begin VB.Menu mnuWindow 
         Caption         =   "&Scan Results"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "&Validation Results"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDpiPmAssistant                  ' required to receive DPI changes
Dim WithEvents cDpiAssist As clsDpiPmAssist ' required to react to DPI changes
Attribute cDpiAssist.VB_VarHelpID = -1
'/// consider using class MsgBox and InputBox methods. See method comments for more.
'    Example: cDpiAssist.MsgBox "Hello World", vbOKOnly + vbInformation, "Test"
'    Example: cDpiAssist.InputBox "Enter some value", "Test"

'/// If clsDpiAsstCmnCtrls class is not used, remove it from your project
'    and also change the #Const below to False. No other code changes needed.
#Const UseDpiAsstCmnCtrls = True
#If UseDpiAsstCmnCtrls Then
    Dim cCCtlsAssist As clsDpiAsstCmnCtrls
#End If

' /// above declarations are for DPI-awareness

Implements IEvents

Dim m_Project As clsProject
Dim m_Busy As Boolean               ' prevent re-entrance while validating/processing
Dim m_MaxNodes As Long              ' maximum loaded node count
Const MaxNewNodes As Long = 1000
'   Treeviews are limited to 32k nodes. Having large numbers of nodes, say
' in the thousands, also slows down population of the tree significantly.
' The majority of projects will never hit that max number, but it is possible
' that very large projects can hit that number and large projects slow down
' tree population regardless.
'   This project will initially load the node headers for every parsed code
' page and then set an arbritrary limit of 1000 more nodes. Parsed items are
' not initially loaded into the treeview, they are loaded on demand as parent
' nodes are expanded. If that additional 1000 node limit is reached, then nodes
' will be culled as needed. In many cases, nodes are automatically culled when
' they are collapsed. The limit is set at end of: IEvents_ParseComplete
'   Dynamic node loading and culling have two purposes: 1) speed in initializing
' the treeview and 2) prevent exceeding treeview maximum limits. To assist with
' culling, nodes that can be auto-filled/culled on demand have a key prefixed
' with "a" for auto-fill on expansion. When the node is populated, the key prefix
' changes to "x" for auto-cull on collapse. As needed auto-culling occurs after
' auto-fill nodes are expanded. See tvScan_Expand & tvScan_Collapse
Const KeyAutoExpand = "a", KeyAutoCull = "x"

Private Sub pvDisplayCodePageScan(tRoot As Node)

    ' routine called for each source file (except vbp,vbg)
    ' when called, tRoot is the "Code Files" treeviewnode
    ' and gParsedItems is set to the code page record
    
    Dim rs As ADODB.Recordset
    Dim tNode As Node, tSection As Node
    Dim sValue$, lParent&, lFlags&, n&, lStats&(0 To 2)
    
    Set rs = gParsedItems.Clone
    rs.Bookmark = gParsedItems.Bookmark
    lParent = rs.Fields(recID).Value
    lFlags = rs.Fields(recFlags).Value
    
    ' determine code page category
    Select Case rs.Fields(recFlags).Value And iaMaskCodePage
    Case iaBAS:         sValue = "Modules"
    Case iaClass:       sValue = "Classes"
    Case iaForm, iaMDI: sValue = "Forms"
    Case iaPPG:         sValue = "Property Pages"
    Case iaUC:          sValue = "User Controls"
    Case iaDesigner:    sValue = "Designers"
    Case iaUserDoc:     sValue = "User Documents"
    End Select
    
    Set tNode = tRoot.Child             ' append code page category as needed
    Do Until tNode Is Nothing
        If tNode.Text = sValue Then Exit Do
        Set tNode = tNode.Next
    Loop
    If tNode Is Nothing Then
        Set tRoot = tvScan.Nodes.Add(tRoot, tvwChild, , sValue)
        tRoot.Sorted = True
    Else
        Set tRoot = tNode
    End If
    
    sValue = rs.Fields(recName).Value   ' extract code page name & update statusbar
    sValue = Mid$(sValue, InStr(sValue, chrDot) + 1)
    sbarStatus.SimpleText = "Loading parsed items for " & sValue
    modMain.FauxDoEvents
    
    ' tweak code page title as needed
    Set tSection = tvScan.Nodes.Add(tRoot, tvwChild, "C" & rs.Fields(recID).Value, sValue)
    n = (rs.Fields(recFlags).Value And iaMaskCodePage)
    If n = iaMDI Then
        tSection.Text = tSection.Text & " (MDI)"
    ElseIf n = iaDesigner Then
        tSection.Text = tSection.Text & " (" & rs.Fields(recAttr).Value & chrParentC
    End If
    tSection.Sorted = True  ' have its child nodes sorted
    
    ' append categories for each item within this code page
    rs.Filter = modMain.SetQuery(recParent, qryIs, lParent)
    If rs.EOF = False Then
        rs.Sort = recType
        Do
            n = rs.Fields(recType).Value
            Select Case n
            Case itVariable:    sValue = "Variables"
            Case itMethod:      sValue = "Methods"
            Case itControl:     sValue = "Controls"
            Case itAPI:         sValue = "APIs"
            Case itConstant:    sValue = "Constants"
            Case itEnum:        sValue = "Enumerations"
            Case itType:        sValue = "User-Defined Types"
            Case itEvent:       sValue = "Events - Raised"
            Case itClassEvent:  sValue = "Events - Class"
                If (lFlags And iaUnresolved) <> 0 Then sValue = chrAsterisk & sValue
            Case itImplements:  sValue = "Implementations"
            Case itStats:       sValue = vbNullString
                Select Case rs.Fields(recFlags).Value
                Case iaStatements: lStats(0) = lStats(0) + rs.Fields(recOffset).Value
                Case iaComments: lStats(2) = lStats(2) + rs.Fields(recOffset).Value
                Case iaExclusions: lStats(1) = lStats(1) + rs.Fields(recOffset).Value
                End Select
            Case Else:          sValue = vbNullString
            End Select
            If LenB(sValue) <> 0 Then
                Set tNode = tSection.Child  ' find existing category
                Do Until tNode Is Nothing
                    If tNode.Text = sValue Then Exit Do
                    Set tNode = tNode.Next
                Loop
                If tNode Is Nothing Then    ' create category if didn't exist
                    Set tNode = tvScan.Nodes.Add(tSection, tvwChild, KeyAutoExpand & n & chrSlash & lParent, sValue)
                    tvScan.Nodes.Add tNode, tvwChild, , "(auto-expand)"
                End If                      ' add placeholder
            End If
            If n = itStats Then             ' want each record
                rs.MoveNext
            Else                            ' skip to next category
                rs.Find modMain.SetQuery(recType, qryNot, n)
            End If
        Loop Until rs.EOF = True
    End If
    tSection.Sorted = False
    rs.Close: Set rs = Nothing
    
    Set tNode = tSection.Child              ' add the code page source
    gParsedItems.Find modMain.SetQuery(recID, qryIs, gParsedItems.Fields(recParent).Value), , , 1&
    If tNode Is Nothing Then
        Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , "Source: " & gParsedItems.Fields(recAttr).Value)
    Else
        Set tNode = tvScan.Nodes.Add(tNode.FirstSibling, tvwPrevious, , "Source: " & gParsedItems.Fields(recAttr).Value)
    End If
    If (lFlags And iaMaskOptions) <> 0 Then ' add code page Option statements
        Set tNode = tvScan.Nodes.Add(tNode, tvwNext, , "Options: ")
        sValue = vbNullString
        If (lFlags And iaOpExplicit) <> 0 Then sValue = ", Explicit"
        If (lFlags And iaOpBase1) <> 0 Then sValue = sValue & ", Base 1"
        If (lFlags And iaOpText) <> 0 Then sValue = sValue & ", Compare Text"
        If (lFlags And iaOpPrivate) <> 0 Then sValue = sValue & ", Private"
        tNode.Text = tNode.Text & Mid$(sValue, 3)
    End If
                                            ' add the stats
    Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , "Statistics")
    tvScan.Nodes.Add tNode, tvwChild, , "Statements: " & FormatNumber$(lStats(0), 0)
    tvScan.Nodes.Add tNode, tvwChild, , "Statements Excluded: " & FormatNumber$(lStats(1), 0)
    tvScan.Nodes.Add tNode, tvwChild, , "Comments: " & FormatNumber$(lStats(2), 0)

End Sub

Private Sub pvLoadCommon(tSection As Node, rs As ADODB.Recordset)

    ' routine handles generic loading of items for code page category
    ' Enums, Types, Constants, APIs, Public Events

    Dim sAttrs$

    rs.Sort = recIdxName
    Do Until rs.EOF = True
        If rs.Fields(recType) <> itEvent Then
            Select Case rs.Fields(recScope).Value
            Case scpPrivate, scpLocal: sAttrs = " {Private}"
            Case Else: sAttrs = " {Public}"
            End Select
        End If
        tvScan.Nodes.Add tSection, tvwChild, chrI & rs.Fields(recID).Value, rs.Fields(recName).Value & sAttrs
        rs.MoveNext
    Loop
    
End Sub

Private Sub pvLoadEvents(tSection As Node, rs As ADODB.Recordset, lParent As Long)

    ' routine handles loading class events for a code page
    
    Dim lOffset&, sName$

    ' locate the code page record to get its name
    gParsedItems.Find modMain.SetQuery(recID, qryIs, lParent), , , 1&
    sName = gParsedItems.Fields(recAttr).Value
    If (gParsedItems.Fields(recFlags).Value And iaImplemented) = 0 Then
        lOffset = InStr(sName, chrDot)
        lOffset = Len(sName) - lOffset + 2
    Else
        lOffset = 1
    End If
    rs.Sort = recIdxName
    Do Until rs.EOF = True      ' strip the name from the event
        sName = Mid$(rs.Fields(recName).Value, lOffset)
        tvScan.Nodes.Add tSection, tvwChild, chrI & rs.Fields(recID).Value, sName
        rs.MoveNext
    Loop
    
End Sub

Private Sub pvLoadVars(tSection As Node, rs As ADODB.Recordset)

    ' routine handles loading variables for a code page
    ' also loaded are events for variables declared "WithEvents"

    Dim sName$, sAttrs$, lFlags&, lAttrs&
    Dim rsEvents As ADODB.Recordset, tNode As Node
    
    rs.Sort = recIdxName
    Do Until rs.EOF = True
        ' for WithEvent variables, customize the display
        sName = rs.Fields(recName).Value
        lFlags = rs.Fields(recFlags).Value
        If (lFlags And iaBrackets) <> 0 Then sName = "[" & sName & "]"
        If (lFlags And iaWithEvents) <> 0 Then
            lAttrs = Len(sName) + 2
            If (lFlags And iaHidden) <> 0 Then
                sName = sName & " [hidden] (WithEvents)"
            Else
                sName = sName & " (WithEvents)"
            End If
        End If
        If (lFlags And iaUnresolved) <> 0 Then sName = chrAsterisk & sName
        
        Select Case rs.Fields(recScope).Value
        Case scpPrivate, scpLocal: sAttrs = " {Private}"
        Case Else: sAttrs = " {Public}"
        End Select
        Set tNode = tvScan.Nodes.Add(tSection, tvwChild, chrI & rs.Fields(recID).Value, sName & sAttrs)
        If lAttrs <> 0 Then                 ' WithEvents in play
            Set tNode = tvScan.Nodes.Add(tNode, tvwChild, , rs.Fields(recAttr).Value & " events")
            If rsEvents Is Nothing Then Set rsEvents = modMain.gParsedItems.Clone
            rsEvents.Filter = modMain.SetQuery(recParent, qryIs, rs.Fields(recID).Value, _
                                                qryAnd, recType, qryIs, itClassEvent)
            Do Until rsEvents.EOF = True    ' load each event, stripping variable from name
                sName = Mid$(rsEvents.Fields(recName).Value, lAttrs)
                tvScan.Nodes.Add tNode, tvwChild, chrI & rsEvents.Fields(recID).Value, sName
                rsEvents.MoveNext
            Loop
            lAttrs = 0
        End If
        rs.MoveNext
    Loop
    If Not rsEvents Is Nothing Then
        rsEvents.Close: Set rsEvents = Nothing
    End If

End Sub

Private Sub pvLoadCtrls(tSection As Node, rs As ADODB.Recordset)

    ' routine handles loading controls for a code page
    ' also loaded are any of its events

    Dim tNode As Node, sName As String, lID As Long
    Dim rsEvents As ADODB.Recordset, lOffset As Long
    
    rs.Sort = recIdxName
    Set rsEvents = gParsedItems.Clone
    Do Until rs.EOF = True
        lID = rs.Fields(recID).Value            ' cache control's record ID
        sName = rs.Fields(recName).Value
        sName = sName & " (" & rs.Fields(recAttr).Value & chrParentC
        If (rs.Fields(recFlags) And iaUnresolved) <> 0 Then sName = chrAsterisk & sName
        Set tNode = tvScan.Nodes.Add(tSection, tvwChild, chrI & lID, sName)
        If rs.Fields(recOffset).Value <> 0 Then ' customize if multiple instances
            tNode.Text = tNode.Text & " x" & CStr(rs.Fields(recOffset).Value) + 1
        End If
        
        ' query for any of its events
        rsEvents.Filter = modMain.SetQuery(recParent, qryIs, lID, _
                                    qryAnd, recType, qryIs, itClassEvent)
        If rsEvents.EOF = False Then
            If LenB(rsEvents.Sort) = 0 Then rsEvents.Sort = recIdxName
            lOffset = Len(rs.Fields(recName).Value) + 2 ' used to strip name from event
            Do
                sName = Mid$(rsEvents.Fields(recName).Value, lOffset)
                tvScan.Nodes.Add tNode, tvwChild, chrI & rsEvents.Fields(recID).Value, sName
                rsEvents.MoveNext
            Loop Until rsEvents.EOF = True
        End If
        rs.MoveNext
    Loop
    
End Sub

Private Sub pvLoadMethods(tSection As Node, rs As ADODB.Recordset)

    ' routine handles loading methods for a code page
    ' not included are methods that are events & Public Events

    Dim lAttrs&, lID&
    Dim sName$, sValue$, sAttrs$

    rs.Sort = recIdxName
    Do Until rs.EOF = True
        sValue = rs.Fields(recName).Value       ' method name
        If sName = sValue Then                  ' previous was same name (Prop Let,Get,Set)
            lAttrs = lAttrs Or (rs.Fields(recFlags).Value And &HFFFF0000)
        Else
            If lAttrs <> 0 Then GoSub postMethod_
            lAttrs = (rs.Fields(recFlags).Value And &HFFFF0000) Or rs.Fields(recScope).Value
            sName = sValue: lID = rs.Fields(recID).Value
        End If
        rs.MoveNext
    Loop
    If lAttrs <> 0 Then GoSub postMethod_
    Exit Sub
    
postMethod_:    ' customize the name displayed for the method, then append it
    If (lAttrs And iaFunction) = iaFunction Then
        sAttrs = " {Fnc "
    ElseIf (lAttrs And iaSub) = iaSub Then
        sAttrs = " {Sub "
    Else
        sAttrs = " {"
        If (lAttrs And iaPropGet) = iaPropGet Then sAttrs = sAttrs & "Get,"
        If (lAttrs And iaPropLet) = iaPropLet Then sAttrs = sAttrs & "Let,"
        If (lAttrs And iaPropSet) = iaPropSet Then sAttrs = sAttrs & "Set,"
        Mid$(sAttrs, Len(sAttrs), 1) = " "
    End If
    Select Case (lAttrs And &HF)
        Case scpFriend: sAttrs = sAttrs & "Friend}"
        Case scpPrivate: sAttrs = sAttrs & "Private}"
        Case Else: sAttrs = sAttrs & "Public}"
    End Select
    tvScan.Nodes.Add tSection, tvwChild, chrI & lID, sName & sAttrs
    Return
End Sub

Private Sub pvLoadImps(tSection As Node, rs As ADODB.Recordset)

    ' routine handles loading Implementations for a code page

    Dim tNode As Node, rsEvents As ADODB.Recordset
    Dim lID&, lAttrs&, lOffset&
    Dim sName$, sValue$, sAttrs$
    
    Set rsEvents = modMain.gParsedItems.Clone
    Do Until rs.EOF = True
        lID = rs.Fields(recID).Value        ' cache record ID
        sName = rs.Fields(recAttr).Value    ' Implemented object's name
        lOffset = InStr(sName, chrDot)         ' used to strip library from name
        If lOffset = 0 Then lOffset = Len(sName) + 2 Else lOffset = Len(sName) - lOffset + 2
        If (rs.Fields(recFlags).Value And iaUnresolved) <> 0 Then sName = chrAsterisk & sName
        
        ' append the item & then append each of its events
        Set tNode = tvScan.Nodes.Add(tSection, tvwChild, chrI & lID, sName)
        rsEvents.Filter = modMain.SetQuery(recParent, qryIs, lID, _
                                           qryAnd, recType, qryIs, itClassEvent)
        If rsEvents.EOF = False Then
            rsEvents.Sort = recIdxName: sName = vbNullString
            Do
                sValue = Mid$(rsEvents.Fields(recName).Value, lOffset)
                If sValue = sName Then      ' previous was same, Prop Let,Get,Set
                    lAttrs = lAttrs Or rsEvents.Fields(recFlags).Value
                Else
                    If lAttrs <> 0 Then GoSub postMethod_
                    lAttrs = lAttrs Or rsEvents.Fields(recFlags).Value
                    sName = sValue: lID = rsEvents.Fields(recID).Value
                End If
                rsEvents.MoveNext
            Loop Until rsEvents.EOF = True
            If lAttrs <> 0 Then GoSub postMethod_
        End If
        rs.MoveNext
    Loop
    rsEvents.Close: Set rsEvents = Nothing
    Exit Sub
    
postMethod_:
    If (lAttrs And iaFunction) = iaFunction Then
        sAttrs = " {Fnc "
    ElseIf (lAttrs And iaSub) = iaSub Then
        sAttrs = " {Sub "
    Else
        sAttrs = " {"
        If (lAttrs And iaPropGet) = iaPropGet Then sAttrs = sAttrs & "Get,"
        If (lAttrs And iaPropLet) = iaPropLet Then sAttrs = sAttrs & "Let,"
        If (lAttrs And iaPropSet) = iaPropSet Then sAttrs = sAttrs & "Set,"
    End If
    Mid$(sAttrs, Len(sAttrs), 1) = "}"
    tvScan.Nodes.Add tNode, tvwChild, chrI & lID, sName & sAttrs
    lAttrs = 0
    Return
End Sub

Private Sub Form_Load()
    Dim sFile As String
    Me.Caption = Me.Caption & " v" & CStr(App.Major) & chrDot & CStr(App.Minor)
    cDpiAssist.Activate Me ' do not move this line to Form_Activate
    sbarStatus.SimpleText = "Select a project to scan"
    If ProcessCommandLine(sFile) = True Then
        Me.Show: FauxDoEvents
        pvLoadFile sFile, False
    End If
End Sub

Private Sub Form_Resize()
    If cDpiAssist.IsScalingCycleActive = False Then pvDoResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f As Form
    Set cDpiAssist = Nothing
    #If UseDpiAsstCmnCtrls Then
        Set cCCtlsAssist = Nothing
    #End If
    Set m_Project = Nothing
    For Each f In Forms
        Unload f
        Set f = Nothing
    Next
End Sub

Private Sub pvLoadFile(sFile As String, bRescan As Boolean)

    Dim n&
    Dim rs As ADODB.Recordset, vBkMk As Variant
    Dim colCPexclusions(0 To 1) As Collection

    m_Busy = True
    tvScan.Nodes.Clear
    tvValidation.Nodes.Clear
    Call mnuWindow_Click(0)
    mnuWindow(1).Enabled = False
    
    If bRescan = True Then
        Set rs = gParsedItems.Clone
        rs.Filter = modMain.SetQuery(recType, qryIs, itValidation)
        If rs.EOF = False Then
            Set colCPexclusions(0) = New Collection
            Set colCPexclusions(1) = New Collection
            rs.Sort = recGrp
            n = rs.Fields(recFlags).Value: rs.MoveNext
            Do Until rs.EOF = True
                If rs.Fields(recGrp).Value <> 0 Then Exit Do
                colCPexclusions(0).Add rs.Fields(recName).Value
                rs.MoveNext
            Loop
            Do Until rs.EOF = True
                colCPexclusions(1).Add rs.Fields(recName).Value
                rs.MoveNext
            Loop
        End If
        rs.Close: Set rs = Nothing
    End If
    Set m_Project = New clsProject
    m_Project.SetSource Me, sFile
    If Not gParsedItems Is Nothing Then
        If Not colCPexclusions(0) Is Nothing Then
            vBkMk = gSourceFile.ProjBookMark
            gSourceFile.ProjBookMark = Empty
            Set rs = gParsedItems.Clone
            rs.Filter = modMain.SetQuery(recType, qryIs, itCodePage)
            modMain.CreateRecord 0, vbNullString, itValidation, 0, 0, , , -1, n
            For n = 1 To colCPexclusions(0).Count
                rs.Find modMain.SetQuery(recName, qryIs, colCPexclusions(0).Item(n)), , , 1&
                If rs.EOF = False Then
                    modMain.CreateRecord 0, vbNullString, itValidation, 0, 0, , , , rs.Fields(recID).Value
                End If
            Next
            For n = 1 To colCPexclusions(1).Count
                rs.Find modMain.SetQuery(recName, qryIs, colCPexclusions(1).Item(n)), , , 1&
                If rs.EOF = False Then
                    modMain.CreateRecord 0, vbNullString, itValidation, 0, 0, , , 1, rs.Fields(recID).Value
                End If
            Next
            rs.Close: Set rs = Nothing
            gSourceFile.ProjBookMark = vBkMk: vBkMk = Empty
            Erase colCPexclusions()
        End If
    End If
    m_Busy = False

End Sub

Private Sub mnuView_Click(Index As Integer)

    If Index = 0 Then                       ' F1
        Load frmValidationHelp
        frmValidationHelp.mnuHelp.Visible = True
        frmValidationHelp.rtb.TextRTF = StrConv(LoadResData(101, "RTF"), vbUnicode)
        frmValidationHelp.ShowAdjustedForDPI
        Exit Sub
    End If
    
    If m_Busy = True Then Exit Sub
    If gParsedItems Is Nothing Then Exit Sub
    
    Dim aData() As Byte
    Dim f As frmValidationHelp
    Dim rs As ADODB.Recordset
    Dim n&, lRead&, hHandle&
    Dim tNode As Node, sFile$, sName$
    
    If Index = mnuView.UBound Then          ' VBG
        sName = mnuView(Index).Tag
    Else
        If Index = mnuView.UBound - 1 Then  ' VBP
            gParsedItems.Bookmark = gSourceFile.ProjBookMark
        Else                                ' Code page, decs section or methods
            If tvScan.Visible = False Then Call mnuWindow_Click(0)
            Set tNode = tvScan.SelectedItem
            If Not tNode Is Nothing Then
                Do Until tNode Is Nothing
                    If LenB(tNode.Key) <> 0 Then
                        If AscW(tNode.Key) = vbKeyC Then Exit Do
                    End If
                    Set tNode = tNode.Parent
                Loop
            End If
            If tNode Is Nothing Then
                cDpiAssist.MsgBox "First select any node related to a code page", vbInformation + vbOKOnly, "No Action Taken"
                Exit Sub
            End If
            Set rs = gParsedItems.Clone
            rs.Find modMain.SetQuery(recID, qryIs, CLng(Mid$(tNode.Key, 2))), , , 1&
            gParsedItems.Find modMain.SetQuery(recID, qryIs, rs.Fields(recParent).Value), , , 1&
        End If
        sName = gParsedItems.Fields(recAttr).Value
    End If
    
    hHandle = modMain.GetFileHandle(sName, False)
    If hHandle = 0 Or hHandle = -1 Then
        cDpiAssist.MsgBox "Failed to access file", vbExclamation + vbOKOnly, "Error"
    Else
        modMain.GetFileLastModDate sName, 0&, 0&, n
        ReDim aData(0 To n - 1)
        ReadFile hHandle, aData(0), n, lRead
        modMain.CloseHandle hHandle
        sFile = StrConv(aData(), vbUnicode)
        Erase aData()
        If Index = 3 Then                   ' declarations section
            n = rs.Fields(recStart).Value: lRead = rs.Fields(recEnd).Value
            sFile = Mid$(sFile, n, lRead - n)
        ElseIf Index = 4 Then               ' methods section
            n = rs.Fields(recOffset).Value: lRead = rs.Fields(recOffset2).Value
            sFile = Mid$(sFile, n, lRead - n)
        End If
        Set f = New frmValidationHelp
        f.Caption = "File: " & Mid$(sName, InStrRev(sName, chrSlash) + 1)
        f.rtb.Text = sFile: sFile = vbNullString
        f.Tag = chrDot: f.SetWordWrap False
        f.Show: Set f = Nothing
    End If
    If Not rs Is Nothing Then
        rs.Close: Set rs = Nothing
    End If
    
End Sub

Private Sub mnuMalIntent_Click(Index As Integer)

    If m_Busy = True Then Exit Sub
    If tvScan.Nodes.Count = 0 Then Exit Sub
    
    If tvValidation.Nodes.Count <> 0 Then
        If tvValidation.Visible = False Then Call mnuWindow_Click(1)
        cDpiAssist.MsgBox "Validations performed. Rescan project to perform new validations", vbInformation + vbOKOnly, "No Action"
        Exit Sub
    End If
    Dim cValidation As clsValidation
    Set cValidation = New clsValidation
    If Index = 0 Then
        cValidation.ValidateProject Me, vtMalicious
    Else
        cValidation.ValidateProject Me, vtMalicious Or vtRegReadDLLs
    End If
    Set cValidation = Nothing
    
End Sub

Private Sub mnuOpen_Click(Index As Integer)
    
    If m_Busy = True Then Exit Sub
    
    Dim f As Form, cBrowser As CmnDialogEx
    Dim bWarn As Boolean, sFile$, sGUID$
    
    For Each f In Forms
        If LenB(f.Tag) <> 0 Then
            If bWarn = False Then
                If cDpiAssist.MsgBox("Any reports/files you are viewing will be closed." & vbCrLf & _
                    "Click cancel to abort and then save the report if desired", vbInformation + vbOKCancel + vbDefaultButton2, "Confirmation") = vbCancel Then
                    Set f = Nothing
                    Exit Sub
                End If
                bWarn = True
            End If
            Unload f
        End If
    Next
    Set f = Nothing
    
    If Index = 1 Then                   ' rescan
        If gParsedItems Is Nothing Then
            Index = 0
        Else
            gParsedItems.Bookmark = gSourceFile.ProjBookMark
            sFile = gParsedItems.Fields(recAttr).Value
        End If
    End If
    If LenB(sFile) = 0 Then
        Set cBrowser = New CmnDialogEx
        With cBrowser
            sGUID = .CreateClientDataGUID("Import")
            .DialogTitle = "Select VB Project"
            .Filter = "Project Files|*.vbp;*.vbg"
        End With
        If cBrowser.ShowOpen(Me.hWnd, , , sGUID) = False Then
            GoTo exitRoutine
        End If
        sFile = cBrowser.FileName
    End If
    
    pvLoadFile sFile, (Index = 1)
    
exitRoutine:
    Set cBrowser = Nothing
End Sub

Private Sub mnuValidate_Click(Index As Integer)
    
    If Index = 0 Then
        If m_Busy = True Then Exit Sub
        If tvScan.Nodes.Count = 0 Then Exit Sub
        If tvValidation.Nodes.Count <> 0 Then
            If tvValidation.Visible = False Then Call mnuWindow_Click(1)
            cDpiAssist.MsgBox "Validations performed. Rescan project to perform new validations", vbInformation + vbOKOnly, "No Action"
            Exit Sub
        End If
        frmValidate2.Show , Me
        
    ElseIf Index = 3 Then
        If m_Busy = True Then Exit Sub
        If tvValidation.Nodes.Count = 0 Then
            If tvScan.Nodes.Count <> 0 Then
                If cDpiAssist.MsgBox("Start validations?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
                    Call mnuValidate_Click(0)
                End If
            End If
            Exit Sub
        End If
        Dim cReport As clsReport, f As frmValidationHelp
        Set cReport = New clsReport
        gParsedItems.Bookmark = gSourceFile.ProjBookMark
        If tvValidation.Nodes(1).Key = chrI & vtZombie Then
            Set f = New frmValidationHelp
            f.Tag = gParsedItems.Fields(recName) & "_ScanReportStd.rtf"
            f.Caption = "Report - Standard Scan"
            cReport.CreateReport_Standard f.rtb, Nothing, Nothing, Me, f.hWnd
            If tvValidation.Nodes(1).Next Is Nothing Then
                Set f = Nothing
            Else
                Set f = New frmValidationHelp
            End If
        Else
            Set f = New frmValidationHelp
        End If
        If Not f Is Nothing Then
            f.Tag = gParsedItems.Fields(recName) & "_ScanReportSafety.rtf"
            f.Caption = "Report - Safety Scan"
            cReport.CreateReport_Safety f.rtb, Nothing, Nothing, Me, f.hWnd
            Set f = Nothing
        End If
        Set cReport = Nothing
    End If
        
End Sub

Private Sub mnuWindow_Click(Index As Integer)
    If Index = 0 Then
        tvScan.Enabled = True: tvScan.Visible = True
        tvValidation.Enabled = False: tvValidation.Visible = False
    Else
        tvValidation.Enabled = True: tvValidation.Visible = True
        tvScan.Enabled = False: tvScan.Visible = False
    End If
    mnuWindow(Index).Checked = True
    mnuWindow(Index Xor 1).Checked = False
End Sub

Private Sub tvScan_Collapse(ByVal Node As ComctlLib.Node)
    
    If Node.Key <> vbNullString Then
        If AscW(Node.Key) = 120 Then            ' x = can be auto-culled
            Dim tNode As Node
            Node.Key = KeyAutoExpand & Mid$(Node.Key, 2)  ' a = auto-fill
            tvScan.Nodes.Add Node.Child, tvwPrevious, , "(auto-expand)"
            Set tNode = Node.Child.LastSibling  ' cull from bottom up
            Do Until tNode Is Node.Child
                tvScan.Nodes.Remove tNode.Index
                Set tNode = Node.Child.LastSibling
            Loop
        End If
    End If
    
End Sub

Private Sub tvScan_Expand(ByVal Node As ComctlLib.Node)
    
    If Node.Key <> vbNullString Then
        If AscW(Node.Key) = 97 Then             ' a = can be auto-filled
            Dim tNode As Node, pNode As Node
            Dim lParent As Long, lType As Long
            Dim rs As ADODB.Recordset
            Set rs = gParsedItems.Clone
            With Node                           ' get item type & code page
                lType = InStr(.Key, chrSlash)
                lParent = CLng(Mid$(.Key, lType + 1))
                lType = CLng(Mid$(.Key, 2, lType - 2))
                .Key = KeyAutoCull & Mid$(.Key, 2) ' x = can be auto-culled
            End With
            rs.Filter = modMain.SetQuery(recParent, qryIs, lParent, _
                                        qryAnd, recType, qryIs, lType, _
                                        qryAnd, recStart, qryGT, -1)
            Select Case lType
            Case itEnum:            pvLoadCommon Node, rs
            Case itType:            pvLoadCommon Node, rs
            Case itImplements:      pvLoadImps Node, rs
            Case itControl:         pvLoadCtrls Node, rs
            Case itAPI:             pvLoadCommon Node, rs
            Case itEvent:           pvLoadCommon Node, rs
            Case itVariable:        pvLoadVars Node, rs
            Case itConstant:        pvLoadCommon Node, rs
            Case itMethod:          pvLoadMethods Node, rs
            Case itClassEvent:      pvLoadEvents Node, rs, lParent
            End Select
            rs.Close: Set rs = Nothing  ' remove placeholder node
            tvScan.Nodes.Remove Node.Child.Index
            
            ' auto-cull as needed, bottom up
            If tvScan.Nodes.Count > m_MaxNodes Then
                For lType = 0 To 1                  ' 2 passes, as needed
                ' Pass #1, cull hidden nodes belonging to collapsed ancestor nodes
                ' Pass #2, cull visible nodes as a last resort
                ' the current code page will not be culled
                    Set pNode = Node.Parent.Parent  ' set to code page category
                    Set tNode = pNode.LastSibling   ' set to last category
                    Do
                        If Not tNode Is pNode Then
                            If lType = 0 Then       ' cull hidden nodes first
                                pvCullNodes tNode, tNode.Expanded
                            Else                    ' cull visible nodes
                                pvCullNodes tNode, False
                            End If                  ' abort when max no longer exceeded
                            If tvScan.Nodes.Count < m_MaxNodes Then Exit For
                        End If
                        Set tNode = tNode.Previous
                    Loop Until tNode Is Nothing
                    Set pNode = Node.Parent         ' set to this code page
                    Set tNode = pNode.LastSibling   ' set to last code page in this category
                    Do
                        If Not tNode Is pNode Then
                            If lType = 0 Then
                                pvCullNodes tNode, tNode.Expanded
                            Else
                                pvCullNodes tNode, False
                            End If
                            If tvScan.Nodes.Count < m_MaxNodes Then Exit For
                        End If
                        Set tNode = tNode.Previous
                    Loop Until tNode Is Nothing
                Next
                Set pNode = Nothing
                Node.EnsureVisible
                Set tvScan.SelectedItem = Node
                ' if we got here and max count still exceeded, set a new max count
                If tvScan.Nodes.Count > m_MaxNodes Then m_MaxNodes = tvScan.Nodes.Count
            End If
            Set tNode = Nothing
        End If
    End If
End Sub

Private Sub pvCullNodes(pNode As Node, isVisible As Boolean)

    ' Recursive routine to cull nodes. Called from tvScan_Expand
    ' isVisible must be false for a node to be culled

    If LenB(pNode.Key) <> 0 Then
        If AscW(pNode.Key) = 120 Then       ' x = can be auto-culled
            If isVisible = False Then pNode.Expanded = False
            Exit Sub                        ' collapsing triggers culling
        ElseIf Asc(pNode.Key) = 97 Then     ' a = auto-fill (node not filled, not culled)
            Exit Sub
        End If
    End If
    
    If Not pNode.Child Is Nothing Then      ' recurse through child nodes
        Dim tNode As Node
        Set tNode = pNode.Child.LastSibling
        Do
            If isVisible = True Then
                ' all ancestors up to this point are currently expanded
                pvCullNodes tNode, tNode.Expanded
            Else
                ' at least one ancestor up to this point is collapsed
                pvCullNodes tNode, False
            End If  ' abort recursion if max no longer exceeded
            If tvScan.Nodes.Count < m_MaxNodes Then Exit Do
            Set tNode = tNode.Previous
        Loop Until tNode Is Nothing
    End If
    
End Sub

Private Sub pvDoResize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    On Error GoTo exitRoutine
    With tvScan
        .Move .Left, .Top, _
            Me.ScaleWidth - .Left - .Left, _
            Me.ScaleHeight - .Top - .Left - sbarStatus.Height
    End With
    ' external ocx dimensions should be validated when DPI aware...
    cDpiAssist.SyncOcxToParent tvScan
    tvValidation.Move tvScan.Left, tvScan.Top, tvScan.Width, tvScan.Height
    cDpiAssist.SyncOcxToParent tvValidation
    
exitRoutine:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub IEvents_ParseComplete()

    ' called when the project has been completely parsed for display
    
    Dim rsBase As ADODB.Recordset, rs As ADODB.Recordset
    Dim tNode As Node, tSection As Node
    Dim tMisc As Node, tRoot As Node
    Dim sValue$, n&, lParent&
    Dim lStmts&, lComments&, lExclusions&
    
    sbarStatus.SimpleText = "Populating tree and collecting imported DLLs"
    modMain.FauxDoEvents
    
    Set rsBase = gParsedItems.Clone
    Set rs = gParsedItems.Clone
    rsBase.Find SetQuery(recParent, qryIs, -1), , , 1&
    With mnuView(mnuView.UBound)
        If rsBase.EOF = False Then
            .Enabled = True
            .Tag = rsBase.Fields(recAttr).Value
        Else
            .Enabled = False
        End If
    End With
    rsBase.Bookmark = gSourceFile.ProjBookMark
    lParent = rsBase.Fields(recID).Value
    Set tRoot = tvScan.Nodes("P" & lParent)
    
    ' update the project node with its name vs its path/filename
    tRoot.Text = "Project: " & rsBase.Fields(recName).Value
    sValue = rsBase.Fields(recAttr2).Value: n = InStr(sValue, chrSemi) + 1
    
    ' append the Properties/Dependencies child node
    Set tSection = tvScan.Nodes.Add(tRoot, tvwChild, , "Properties and Dependencies")
    ' append the Source, Type & Version nodes
    tvScan.Nodes.Add tSection, tvwChild, , "Source: " & rs.Fields(recAttr).Value
    tvScan.Nodes.Add tSection, tvwChild, , "Type: " & rs.Fields(recDiscrep).Value
    If n <> 1 Then tvScan.Nodes.Add tSection, tvwChild, , "Version: " & Left$(sValue, n - 2)
    
    ' append the References node & then each reference
    rs.Filter = modMain.SetQuery(recType, qryIs, itReference)
    If rs.EOF = False Then
        Set tSection = tvScan.Nodes.Add(tSection, tvwChild, , "References")
        Do
            If rs.Fields(recFlags).Value = iaExternalProj Then
                sValue = "Project: " & rs.Fields(recName).Value
                n = InStr(rs.Fields(recAttr2).Value, chrSemi)
                If n <> 0 Then
                    sValue = sValue & " version " & Left$(rs.Fields(recAttr2).Value, n - 1)
                End If
                Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , sValue)
                tvScan.Nodes.Add tNode, tvwChild, , rs.Fields(recAttr).Value
            Else
                sValue = rs.Fields(recAttr2).Value
                sValue = Mid$(sValue, InStr(sValue, chrHash) + 1)
                n = InStrRev(sValue, chrDot)
                sValue = " version " & Left$(sValue, n - 1)
                Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , rs.Fields(recName).Value & sValue)
                If rs.Fields(recFlags).Value = -1 Then
                    n = InStr(rs.Fields(recAttr2).Value, chrHash)
                    tvScan.Nodes.Add tNode, tvwChild, , Left$(rs.Fields(recAttr2).Value, n - 1)
                Else
                    tvScan.Nodes.Add tNode, tvwChild, , rs.Fields(recAttr).Value
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF = True
        Set tSection = tSection.Parent
    End If
    
    ' append the DLLs node, then each DLL while skipping duplicates
    rs.Filter = modMain.SetQuery(recType, qryIs, itAPI)
    If rs.EOF = False Then
        Set tSection = tvScan.Nodes.Add(tSection, tvwChild, , "DLLs Imported")
        tSection.Sorted = True: n = 0
        rs.Sort = recGrp & chrComma & recIdxAttr
        Do Until rs.EOF = True
            If rs.Fields(recGrp).Value <> n Then        ' different API DLL?
                n = rs.Fields(recGrp).Value
                Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , rs.Fields(recAttr2).Value)
                sValue = vbNullString
            End If
            If rs.Fields(recAttr).Value <> sValue Then  ' different DLL function?
                sValue = rs.Fields(recAttr).Value
                tvScan.Nodes.Add tNode, tvwChild, , sValue
            End If
            rs.MoveNext
        Loop
        tSection.Sorted = False
        Set tSection = tSection.Parent
    End If
    
    ' append support files nodes and the path/filename of the file(s)
    rs.Filter = modMain.SetQuery(recType, qryNot, itSourceFile, qryAnd, recParent, qryIs, lParent)
    If rs.EOF = False Then
        rs.Sort = recType
        Do
            Select Case rs.Fields(recType)
            Case itResFile      ' only ever one
                Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , "Resource File")
                tvScan.Nodes.Add tNode, tvwChild, , rs.Fields(recAttr).Value
            Case itHelpFile     ' only ever one
                Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , "Help File")
                tvScan.Nodes.Add tNode, tvwChild, , rs.Fields(recAttr).Value
            Case itMiscFile     ' can be multiple
                If tMisc Is Nothing Then
                    Set tMisc = tvScan.Nodes.Add(tSection, tvwChild, , "Miscellaneous File(s)")
                End If
                tvScan.Nodes.Add tMisc, tvwChild, , rs.Fields(recAttr).Value
            End Select
            rs.MoveNext
        Loop Until rs.EOF = True
        rs.Sort = vbNullString
    End If
    
    ' if the project starts with Sub Main, identify which module Main exists in
    sValue = rsBase.Fields(recAttr2).Value: n = InStr(sValue, chrSemi) + 1
    Set tNode = tvScan.Nodes.Add(tSection, tvwChild, , "Startup Object: " & Mid$(sValue, n))
    n = InStr(rsBase.Fields(recAttr2).Value, chrSemi) + 1
    If LCase$(Mid$(rsBase.Fields(recAttr2).Value, n)) = "sub main" Then
        ' locate the module where Sub Main exists
        rs.Filter = modMain.SetQuery(recType, qryIs, itMethod, qryAnd, _
                                    recGrp, qryIs, vbKeyM)
        If rs.EOF = False Then
            rs.Filter = modMain.SetQuery(recID, qryIs, rs.Fields(recCodePg).Value)
            n = InStr(rs.Fields(recName).Value, chrDot)
            tNode.Text = tNode.Text & " (Module: " & Mid$(rs.Fields(recName).Value, n + 1) & chrParentC
        End If
    End If
    
    ' filter for all code pages, excluding any external project code pages
    rsBase.Filter = modMain.SetQuery(recType, qryIs, itSourceFile, qryAnd, _
                                    recFlags, qryGT, -1)
    
    ' append the statistics node, then tally statistics from each code page
    Set tSection = tvScan.Nodes.Add(tSection, tvwChild, , "Statistics -- see legend")
    rs.Filter = modMain.SetQuery(recType, qryIs, itStats)
    If rs.EOF = False Then
        Do
            Select Case rs.Fields(recFlags).Value
            Case iaStatements: lStmts = lStmts + rs.Fields(recOffset).Value
            Case iaComments: lComments = lComments + rs.Fields(recOffset).Value
            Case iaExclusions: lExclusions = lExclusions + rs.Fields(recOffset).Value
            End Select
            rs.MoveNext
        Loop Until rs.EOF = True
        sValue = FormatNumber$(lStmts, 0)   ' cached, used later
        tvScan.Nodes.Add tSection, tvwChild, , "Statements: " & sValue
        tvScan.Nodes.Add tSection, tvwChild, , "Statements Excluded: " & FormatNumber$(lExclusions, 0)
        tvScan.Nodes.Add tSection, tvwChild, , "Comments: " & FormatNumber$(lComments, 0)
    End If
    
    ' add other statistics
    tvScan.Nodes.Add tSection, tvwChild, , "Source Files: " & CStr(rsBase.RecordCount)
    rs.Filter = modMain.SetQuery(recType, qryIs, itMethod, _
                                    qryAnd, recOffset2, qryGT, 0)
    n = rs.RecordCount
    rs.Filter = modMain.SetQuery(recType, qryIs, itClassEvent)
    tvScan.Nodes.Add tSection, tvwChild, , "Methods/Events: " & FormatNumber$(n + rs.RecordCount, 0)
    rs.Filter = modMain.SetQuery(recType, qryIs, itAPI)
    tvScan.Nodes.Add tSection, tvwChild, , "API Declarations: " & FormatNumber$(rs.RecordCount, 0)
    rs.Filter = modMain.SetQuery(recType, qryIs, itEnum)
    tvScan.Nodes.Add tSection, tvwChild, , "Enum Declarations: " & FormatNumber$(rs.RecordCount, 0)
    rs.Filter = modMain.SetQuery(recType, qryIs, itType)
    tvScan.Nodes.Add tSection, tvwChild, , "Type Declarations: " & FormatNumber$(rs.RecordCount, 0)
    rs.Close: Set rs = Nothing
    
    ' append the Code Files node, then process each code page
    Set tSection = tvScan.Nodes.Add(tRoot, tvwChild, "c", "Code Files")
    If rsBase.EOF = False Then
        Set tMisc = Nothing
        tSection.Sorted = True: tSection.Expanded = True
        Do Until rsBase.EOF = True
            Set tNode = tSection
            If (rsBase.Fields(recFlags) And iaFileError) = 0 Then
                gParsedItems.Find modMain.SetQuery(recParent, qryIs, rsBase.Fields(recID).Value), , , 1&
                pvDisplayCodePageScan tNode
            Else
                If tMisc Is Nothing Then
                    Set tMisc = tvScan.Nodes.Add(tSection, tvwChild, , "(File Access Errors)")
                    tMisc.Expanded = True
                End If
                tvScan.Nodes.Add tMisc, tvwChild, , rsBase.Fields(recAttr).Value
            End If
            rsBase.MoveNext
        Loop
        tSection.Sorted = False
        Set tMisc = Nothing
    End If
    sValue = "Done. Parsed " & CStr(rsBase.RecordCount) & " source files having " & sValue & " statements"
    sbarStatus.SimpleText = sValue
    rsBase.Close: Set rsBase = Nothing
    
    ' add finishing touches. For each code page category, tally number of files
    Set tNode = tSection.Child
    Do Until tNode Is Nothing
        n = tNode.Children
        If n <> 0 Then tNode.Text = tNode.Text & " (" & CStr(n) & chrParentC
        Set tNode = tNode.Next
    Loop
    
    ' add the legend
    Set tSection = tvScan.Nodes.Add(tRoot, tvwNext, , "Legend")
    tvScan.Nodes.Add tSection, tvwChild, , "Fnc = Function"
    tvScan.Nodes.Add tSection, tvwChild, , "Sub = Subroutine"
    tvScan.Nodes.Add tSection, tvwChild, , "Get = Property Get"
    tvScan.Nodes.Add tSection, tvwChild, , "Let = Property Let"
    tvScan.Nodes.Add tSection, tvwChild, , "Set = Property Set"
    tvScan.Nodes.Add tSection, tvwChild, , "* = Events only; were not resolved. Missing/unregistered reference or external VBP project"
    tvScan.Nodes.Add tSection, tvwChild, , "Statements Stat includes executable only"
    tvScan.Nodes.Add tSection, tvwChild, , "Statements Excluded Stat due to compiler directive(s) evaluating False"
    tvScan.Nodes.Add tSection, tvwChild, , "Comments Stat includes in-line and/or separate line comments/remarks"
    
    tRoot.Selected = True: tRoot.EnsureVisible
    Set tNode = Nothing: Set tSection = Nothing: Set tRoot = Nothing
    m_MaxNodes = tvScan.Nodes.Count + MaxNewNodes
    m_Busy = False

End Sub

Private Function IEvents_ParsedBegin(ByVal bGroupProject As Boolean) As Boolean
    
    ' called when the VBP/VBG file is accessed and about to be parsed
    
    If bGroupProject = True Then
        If cDpiAssist.MsgBox("Only the Startup Project in group projects are scanned. Continue?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
            Set m_Project = Nothing
        Else
            IEvents_ParsedBegin = True
        End If
        Exit Function
    End If
    
    Dim sValue$, sKey$, n&
    
    gParsedItems.Bookmark = gSourceFile.ProjBookMark
    sKey = "P" & gParsedItems.Fields(recID).Value
    sValue = gParsedItems.Fields(recAttr).Value
    n = InStr(sValue, chrSemi) + 1
    sValue = "Project: " & Mid$(sValue, n)
    
    If tvScan.Nodes.Count = 0 Then
        tvScan.Nodes.Add(, , sKey, sValue).Expanded = True
    Else
        tvScan.Nodes.Add(tvScan.Nodes(1), tvwChild, sKey, sValue).Expanded = True
    End If
    modMain.FauxDoEvents
    
End Function

Private Sub IEvents_ParseError(errCode As ValidationConstants)
    ' called when an error occurs that will abort processing
    Dim sError As String
    If errCode <= vnAborted Then
        Select Case errCode
        Case vnFileAccess: sError = "File Access Error"
        Case vnFileEmpty: sError = "Empty File"
        Case vnFileInvalid: sError = "Unexpected file format or corrupt file"
        Case vnFileNotFound: sError = "File Not Found"
        Case vnFileTooBig: sError = "File exceeds 2 GB"
        Case vnAborted: sbarStatus.SimpleText = "Aborted by user": Exit Sub
        End Select
        sbarStatus.SimpleText = "Error. " & sError
        cDpiAssist.MsgBox "Critical error." & vbCrLf & sError, vbExclamation + vbOKOnly, "Done"
    Else
        Select Case errCode
        Case vnOpenBracket: sError = "No closing bracket [ ]"
        Case vnOpenQuote: sError = "No closing quote in string literal"
        Case vnOpenDate: sError = "No closing hash in date literal"
        Case vnOpenMethodBlk: sError = "Missing End Sub, Function, or Property"
        Case vnOpenEnumUDT: sError = "Missing End Enum or End Type"
        Case vnOpenCompIF: sError = "Missing #End If statement"
        Case Else: Exit Sub
        End Select
        cDpiAssist.MsgBox "Parsing errors existed in one or more source files." & vbCrLf & sError & vbCrLf & vbCrLf & _
                    "Ensure project is syntax free. Suggest starting project in IDE with Ctrl+F5", vbInformation + vbOKOnly, "FYI"
    End If
End Sub

Private Sub IEvents_ReportComplete(ReportType As Long, lParam As Long)

    ' called when validation report is ready for display

    If lParam <> 0 Then
        Dim f As Form
        For Each f In Forms
            If f.hWnd = lParam Then
                f.SetWordWrap False
                f.ShowAdjustedForDPI
                Exit For
            End If
        Next
        Set f = Nothing
    End If
    sbarStatus.SimpleText = vbNullString
    
End Sub

Private Sub IEvents_Status(Msg As String)
    
    ' called to update the parsing status
    sbarStatus.SimpleText = Msg: FauxDoEvents
    
End Sub

Private Sub IEvents_ValidationBegin()

    ' called when validation is about to start
    m_Busy = True
    
End Sub

Private Sub IEvents_ValidationComplete(cValidation As clsValidation, lOptions As ValidationTypeEnum)
    
    ' called when validation has completed
    
    Dim tNode As Node, cReport As clsReport
    
    mnuWindow(1).Enabled = True
    Set cReport = New clsReport
    m_Busy = True
    If (lOptions And vtZombie) <> 0 Then
        Set tNode = tvValidation.Nodes.Add(, , chrI & vtZombie, "Summary - Standard Validation")
        cReport.CreateReport_Standard Nothing, tvValidation, tNode, Me, 0
    End If
    If (lOptions And vtMalicious) <> 0 Then
        Set tNode = tvValidation.Nodes.Add(, , chrI & vtMalicious, "Summary - Safety Validation")
        cReport.CreateReport_Safety Nothing, tvValidation, tNode, Me, 0
    End If
    Set cReport = Nothing
    Call mnuWindow_Click(1)
    m_Busy = False
    
End Sub


'/// Following routines are callbacks related to DPI awareness

Private Sub IDpiPmAssistant_Attach(oDpiPmAssist As clsDpiPmAssist)
    '/// this is required else no DPI scaling will occur
    Set cDpiAssist = oDpiPmAssist
    #If UseDpiAsstCmnCtrls Then
        Set cCCtlsAssist = New clsDpiAsstCmnCtrls
        cCCtlsAssist.Attach Me, oDpiPmAssist
    #End If
End Sub

Private Function IDpiPmAssistant_DpiScalingCycle(ByVal Reason As DpiScaleReason, _
                                ByVal Action As DpiScaleCycleEnum, _
                                ByVal OldDPI As Long, ByVal NewDPI As Long, _
                                ByRef userParams As Variant) As Long
    '/// use this event to prep for rescaling and handle any post-scaling actions you need
    #If UseDpiAsstCmnCtrls Then
        cCCtlsAssist.DpiScalingCycle Reason, Action, OldDPI, NewDPI
    #End If
    If Action = dpiAsst_EndCycleHost Then pvDoResize
End Function

Private Function IDpiPmAssistant_ScaleControlVB(theControl As Control, _
                        ByVal Reason As DpiScaleReason, ByVal Action As DpiActionEnum, _
                        ByVal ScaleRatio As Single, newX As Single, newY As Single, _
                        newCx As Single, newCy As Single, userParams As Variant) As Long
    '/// identify any controls that should not be scaled by returning non-zero
    '/// for controls with picture properties, scale images separately as needed
End Function

Private Function IDpiPmAssistant_ScaleControlOCX(theControl As Control, _
                        ByVal Reason As DpiScaleReason, ByVal Action As DpiActionEnum, _
                        ByVal ScaleRatio As Single, newX As Single, newY As Single, _
                        newCx As Single, newCy As Single, fontProperties As String, _
                        userParams As Variant) As Long
    '/// identify any controls that should not be scaled by returning non-zero
    '/// for controls with picture properties, scale images separately as needed
    
    '/// if using the common controls assist class, then sample code looks like:
    If Action = dpiAsst_BeginEvent Then
        #If UseDpiAsstCmnCtrls Then
            IDpiPmAssistant_ScaleControlOCX = cCCtlsAssist.ScaleControlOCX(theControl, Reason, Action, ScaleRatio, newX, newY, newCx, newCy)
        #End If
    End If
End Function

Private Sub IDpiPmAssistant_ScaleHost(ByVal Reason As DpiScaleReason, _
                          ByRef TwipsWidth As Single, ByRef TwipsHeight As Single, _
                          ByVal OldDPI As Long, ByVal NewDPI As Long, _
                          ByRef IncludeSplashControl As Boolean)
    '/// if overriding passed size parameters, change them relative to new DPI
    '/// to display a splash control while scaling:
    IncludeSplashControl = Me.Visible
End Sub

Private Sub IDpiPmAssistant_IncludeSetParentControls(ByRef theControls As VBA.Collection, _
                                    ByVal Reason As DpiScaleReason, _
                                    ByVal ScaleRatio As Single, ByRef userParams As Variant)
    '/// respond to this if you use SetParent to add controls from other forms
    '/// i.e., If you used: SetParent Form2.Text1.hWnd, Me.hWnd
    '          then include that control here: theControls.Add Form2.Text1
End Sub

Private Function IDpiPmAssistant_Subclasser(EventValue As Long, ByVal BeforeHwnd As Boolean, _
                        ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '/// respond to this if you have called cDpiAssist.SubclassHwnd
End Function

Private Sub cDpiAssist_ResyncAlignedOcx()
    '/// Event requires declaring cDpiAssist using WithEvents keyword
    '   If you have any non-intrinsic controls with Align property <> vbAlignNone
    '   then include them below like so, for each of those controls.
    '   If any is a coolbar control, pass 3rd (optional) parameter as True
       cDpiAssist.SyncOcxToParent sbarStatus
End Sub
