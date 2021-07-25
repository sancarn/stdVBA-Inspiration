VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmValidate2 
   Caption         =   "Validation Options"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
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
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   90
      ScaleHeight     =   1215
      ScaleWidth      =   7455
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5805
      Width           =   7455
      Begin VB.CommandButton cmdGo 
         Caption         =   "Perform Validations"
         Height          =   375
         Left            =   4095
         TabIndex        =   5
         Top             =   705
         Width           =   3240
      End
      Begin VB.ComboBox cboLitOpts 
         Height          =   330
         Index           =   1
         ItemData        =   "frmValidate2.frx":0000
         Left            =   2805
         List            =   "frmValidate2.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   735
         Width           =   1065
      End
      Begin VB.ComboBox cboLitOpts 
         Height          =   330
         Index           =   0
         ItemData        =   "frmValidate2.frx":0004
         Left            =   2805
         List            =   "frmValidate2.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1065
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "View Validation Details"
         Height          =   375
         Left            =   4095
         TabIndex        =   4
         Top             =   300
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Minimum duplication count is"
         Height          =   210
         Index           =   4
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Minimum literal length is"
         Height          =   210
         Index           =   3
         Left            =   60
         TabIndex        =   10
         Top             =   375
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "When reporting duplicate hardcoded literals..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   0
         Width           =   3825
      End
   End
   Begin ComctlLib.TreeView tvChecksExt 
      Height          =   3060
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   5398
      _Version        =   327682
      Indentation     =   587
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin ComctlLib.TreeView tvChecksStd 
      Height          =   1860
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   3281
      _Version        =   327682
      Indentation     =   587
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Extended Optional Checks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2355
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Standard Validation Checks - Always Performed"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   3915
   End
End
Attribute VB_Name = "frmValidate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDpiPmAssistant                  ' required to receive DPI changes
Dim cDpiAssist As clsDpiPmAssist ' required to react to DPI changes
Attribute cDpiAssist.VB_VarHelpID = -1
#Const UseDpiAsstCmnCtrls = True
#If UseDpiAsstCmnCtrls Then
    Dim cCCtlsAssist As clsDpiAsstCmnCtrls
#End If

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Type TVITEM
    mask As Long
    hItem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type
Const TV_FIRST As Long = &H1100
Const TVM_GETITEMW As Long = (TV_FIRST + 62)
Const TVM_SETITEMW As Long = (TV_FIRST + 63)
Const TVIS_STATEIMAGEMASK As Long = &HF000&
Const TVIF_HANDLE As Long = &H10
Const TVIF_STATE As Long = &H8
Const TV_NodeUnchecked = &H1000&
Const TV_NodeChecked = &H2000&
Const XREF_NodeHItem = 68&

Dim m_NodeDblClk As Long
Dim udtItem As TVITEM

Private Sub Form_Load()
    cDpiAssist.Activate Me ' do not move this line to Form_Activate
    
    pvInitControls
End Sub

Private Sub Form_Resize()
    If cDpiAssist.IsScalingCycleActive = False Then pvDoResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' per MSDN, when checkbox style used on v5 of treeview,
    ' must destroy the imagelist associated with checkmark bitmaps
    Dim hIml As Long
    Const TVSIL_STATE As Long = 2
    Const TVM_GETIMAGELIST As Long = (TV_FIRST + 8)
    
    hIml = SendMessage(tvChecksStd.hWnd, TVM_GETIMAGELIST, TVSIL_STATE, ByVal 0&)
    If hIml <> 0 Then ImageList_Destroy hIml
    hIml = SendMessage(tvChecksExt.hWnd, TVM_GETIMAGELIST, TVSIL_STATE, ByVal 0&)
    If hIml <> 0 Then ImageList_Destroy hIml

    Set cDpiAssist = Nothing
    #If UseDpiAsstCmnCtrls Then
        Set cCCtlsAssist = Nothing
    #End If
    
End Sub

Private Sub cmdHelp_Click()
    Load frmValidationHelp
    frmValidationHelp.rtb.TextRTF = StrConv(LoadResData(101, "RTF"), vbUnicode)
    frmValidationHelp.ShowAdjustedForDPI
End Sub

Private Sub cmdGo_Click()

    ' determine which items are checked and build options from those
    ' then begin validation

    Dim tNode As Node, cNode As Node, sPrjName$
    Dim lOptions As ValidationTypeEnum
    Dim cValidation As clsValidation
    Dim rs As ADODB.Recordset, vBkMk As Variant
    
    ' delete any recordset entries for previous validation options
    Set rs = gParsedItems.Clone
    vBkMk = gSourceFile.ProjBookMark
    rs.Bookmark = vBkMk
    sPrjName = rs.Fields(recName).Value & chrDot
    gSourceFile.ProjBookMark = Empty
    rs.Filter = modMain.SetQuery(recType, qryIs, itValidation)
    If rs.EOF = False Then rs.Delete adAffectCurrent
    rs.Close: Set rs = Nothing
    
    Set cValidation = New clsValidation
    lOptions = vtEmptyCode Or vtStopEnd Or vtVarType Or vtWithNoEvents Or vtZombie
    With tvChecksExt
        ' /// Get ReDim option
        Set tNode = .Nodes(1)
        If pvGetItemState(tNode, .hWnd) = TV_NodeChecked Then lOptions = lOptions Or vtRedim
        ' /// Get Var/Str Functions option
        Set tNode = tNode.Next
        If pvGetItemState(tNode, .hWnd) = TV_NodeChecked Then
            lOptions = lOptions Or vtVarFunc
            If pvGetItemState(tNode.Child, .hWnd) = TV_NodeChecked Then lOptions = lOptions Or vtExcludeDateTime
        End If
        ' /// Get Dupe Decs option
        Set tNode = tNode.Next
        If pvGetItemState(tNode, .hWnd) = TV_NodeChecked Then
            lOptions = lOptions Or vtDupeDecs
            If pvGetItemState(tNode.Child, .hWnd) = TV_NodeChecked Then lOptions = lOptions Or vtDupeDecsNoScope
            ' get list of excluded code pages; persist selection to recordset
            If Not tNode.Child.Next Is Nothing Then
                Set cNode = tNode.Child.Next.Child.Next
                Do
                    If pvGetItemState(cNode, .hWnd) = TV_NodeChecked Then
                        modMain.CreateRecord 0, sPrjName & cNode.Text, itValidation, 0, 0, , , 0, CLng(Mid$(cNode.Key, 2))
                    End If
                    Set cNode = cNode.Next
                Loop Until cNode Is Nothing
            End If
        End If
        ' /// Get Dupe Literals option
        Set tNode = tNode.Next
        If pvGetItemState(tNode, .hWnd) = TV_NodeChecked Then
            lOptions = lOptions Or vtDupeLiterals
            If pvGetItemState(tNode.Child, .hWnd) = TV_NodeChecked Then lOptions = lOptions Or vtDupeLitsNoScope
            ' get list of excluded code pages; persist selection to recordset
            If Not tNode.Child.Next Is Nothing Then
                Set cNode = tNode.Child.Next.Child.Next
                Do
                    If pvGetItemState(cNode, .hWnd) = TV_NodeChecked Then
                        modMain.CreateRecord 0, sPrjName & cNode.Text, itValidation, 0, 0, , , 1, CLng(Mid$(cNode.Key, 2))
                    End If
                    Set cNode = cNode.Next
                Loop Until cNode Is Nothing
            End If
        End If
        Set cNode = Nothing
        Set tNode = tNode.Next
        ' /// Get Malicious Code option
        If pvGetItemState(tNode, .hWnd) = TV_NodeChecked Then
            lOptions = lOptions Or vtMalicious
            If pvGetItemState(tNode.Child, .hWnd) = TV_NodeChecked Then lOptions = lOptions Or vtRegReadDLLs
        End If
    End With
    
    ' /// update options to exclude enum members
    If pvGetItemState(tvChecksStd.Nodes(chrZ), tvChecksStd.hWnd) = TV_NodeChecked Then
        lOptions = lOptions Or vtExcludeEMbrZombies
    End If
    ' /// Get minimum literal requirements
    lOptions = lOptions Or cboLitOpts(0).ListIndex * vtLitMinSizeShift
    lOptions = lOptions Or (cboLitOpts(1).ListIndex + 1) * vtLitMinCountShift
    
    modMain.CreateRecord 0, vbNullString, itValidation, 0, 0, , , -1, lOptions
    gSourceFile.ProjBookMark = vBkMk: vBkMk = Empty
    Me.Visible = False
    cValidation.ValidateProject frmMain, lOptions
    Unload Me

End Sub

Private Sub pvInitControls()

    Const GWL_STYLE As Long = -16
    Const TVS_CHECKBOXES As Long = &H100
    Dim n As Long, tNode As Node
    Dim rs As ADODB.Recordset
    
    '  set 2 combobox lists
    cboLitOpts(0).AddItem "0"
    For n = 1 To 10
        cboLitOpts(0).AddItem CStr(n)
        cboLitOpts(1).AddItem CStr(n)
    Next
    cboLitOpts(0).ListIndex = 2
    cboLitOpts(1).ListIndex = 0
    
    ' add checkbox style to treeviews
    n = GetWindowLong(tvChecksExt.hWnd, GWL_STYLE)
    n = n Or TVS_CHECKBOXES
    SetWindowLong tvChecksExt.hWnd, GWL_STYLE, n
    SetWindowLong tvChecksStd.hWnd, GWL_STYLE, n
    
    ' get list of code pages
    Set rs = gParsedItems.Clone
    rs.Filter = modMain.SetQuery(recType, qryIs, itCodePage, qryAnd, recFlags, qryGT, -1)
    rs.Sort = recIdxName
    
    ' for each standard check, set checkbox value
    udtItem.stateMask = TVIS_STATEIMAGEMASK
    udtItem.mask = TVIF_HANDLE Or TVIF_STATE
    
    ' set the standard checks
    With tvChecksStd
        pvSetItemState .Nodes.Add(, , , "Report Option Explicit not used"), .hWnd, 0
        pvSetItemState .Nodes.Add(, , , "Report Methods without Executable Statements"), .hWnd, 0
        pvSetItemState .Nodes.Add(, , , "Report Methods with Stop or End statements"), .hWnd, 0
        pvSetItemState .Nodes.Add(, , , "Report Items with no VarType"), .hWnd, 0
        Set tNode = .Nodes.Add(, , , "Report Zombie Items")
        pvSetItemState tNode, .hWnd, 0: tNode.Expanded = True
        ' set checkbox for its child node
        pvSetItemState .Nodes.Add(tNode, tvwChild, chrZ, "Exclude enumeration members"), .hWnd, TV_NodeChecked
        pvSetItemState .Nodes.Add(, , , "Report Variables Declared WithEvents having no Events"), .hWnd, 0
    End With
    
    ' set the extended checks
    With tvChecksExt
        ' /// ReDim check
        pvSetItemState .Nodes.Add(, , , "Report ReDim Statements on Undeclared Variables"), .hWnd, TV_NodeChecked
        ' /// Variant vs. String VB Function check
        Set tNode = .Nodes.Add(, , , "Report use of Variant vs. String VB Functions")
        tNode.Expanded = True
        pvSetItemState .Nodes.Add(tNode, tvwChild, , "Exclude use of variant version of Date & Time"), .hWnd, TV_NodeChecked
        ' /// Dupe Dec Names check
        Set tNode = .Nodes.Add(, , chrI & vtDupeDecs, "Report Duplicated Declaration Names")
            tNode.Expanded = True
            .Nodes.Add tNode, tvwChild, , "Include all scopes. Do not restrict duplication to just global item names"
            ' add list of code pages as children to last child node
            If rs.EOF = False Then
                Set tNode = .Nodes.Add(tNode, tvwChild, , "Do not compare item names from these code pages...")
                pvSetItemState tNode, .hWnd, 0
                pvSetItemState .Nodes.Add(tNode, tvwChild, , "Double click here to copy exclusions from duplicated literals"), .hWnd, 0
                n = InStr(rs.Fields(recName).Value, chrDot) + 1
                Do Until rs.EOF = True
                    .Nodes.Add tNode, tvwChild, chrW & rs.Fields(recID).Value, Mid$(rs.Fields(recName).Value, n)
                    rs.MoveNext
                Loop
                rs.MoveFirst
            End If
        ' /// Dupe Literals check
        Set tNode = .Nodes.Add(, , chrI & vtDupeLiterals, "Report Duplicated Literals")
            tNode.Expanded = True
            .Nodes.Add tNode, tvwChild, , "Include all scopes. Do not restrict duplication to just global constant values"
            ' add list of code pages as children to last child node
            If rs.EOF = False Then
                Set tNode = .Nodes.Add(tNode, tvwChild, , "Do not compare constant values from these code pages...")
                pvSetItemState tNode, .hWnd, 0
                pvSetItemState .Nodes.Add(tNode, tvwChild, , "Double click here to copy exclusions from duplicated declarations"), .hWnd, 0
                Do Until rs.EOF = True
                    .Nodes.Add tNode, tvwChild, chrM & rs.Fields(recID).Value, Mid$(rs.Fields(recName).Value, n)
                    rs.MoveNext
                Loop
            End If
        ' /// Malicious Code check
        Set tNode = .Nodes.Add(, , , "Peform Safety Check (Malicious Code Check)")
        tNode.Expanded = True
        .Nodes.Add tNode, tvwChild, , "Also report common registry reading APIs"
    End With
    
    ' /// reapply any previous validation options (applies if project was rescanned)
    rs.Filter = modMain.SetQuery(recType, qryIs, itValidation)
    If rs.EOF = False Then
        rs.Sort = recGrp
        n = rs.Fields(recFlags).Value       ' validation options
        rs.MoveNext
        With tvChecksExt
            Do Until rs.EOF = True
                If rs.Fields(recGrp).Value <> 0 Then Exit Do
                pvSetItemState .Nodes(chrW & rs.Fields(recFlags).Value), .hWnd, TV_NodeChecked
                rs.MoveNext
            Loop
            Do Until rs.EOF = True
                pvSetItemState .Nodes(chrM & rs.Fields(recFlags).Value), .hWnd, TV_NodeChecked
                rs.MoveNext
            Loop
            Set tNode = .Nodes(1)
            If (n And vtRedim) = 0 Then pvSetItemState tNode, .hWnd, TV_NodeUnchecked
            Set tNode = tNode.Next
            If (n And vtVarFunc) <> 0 Then pvSetItemState tNode, .hWnd, TV_NodeChecked
            If (n And vtExcludeDateTime) = 0 Then pvSetItemState tNode.Child, .hWnd, TV_NodeUnchecked
            Set tNode = tNode.Next
            If (n And vtDupeDecs) <> 0 Then pvSetItemState tNode, .hWnd, TV_NodeChecked
            If (n And vtDupeDecsNoScope) <> 0 Then pvSetItemState tNode.Child, .hWnd, TV_NodeChecked
            Set tNode = tNode.Next
            If (n And vtDupeLiterals) <> 0 Then pvSetItemState tNode, .hWnd, TV_NodeChecked
            If (n And vtDupeLitsNoScope) <> 0 Then pvSetItemState tNode.Child, .hWnd, TV_NodeChecked
            Set tNode = tNode.Next
            If (n And vtMalicious) <> 0 Then pvSetItemState tNode, .hWnd, TV_NodeChecked
            If (n And vtRegReadDLLs) <> 0 Then pvSetItemState tNode.Child, .hWnd, TV_NodeChecked
        End With
        If (n And vtExcludeEMbrZombies) = 0 Then
            pvSetItemState tvChecksStd.Nodes(chrZ), tvChecksStd.hWnd, TV_NodeUnchecked
        End If
        cboLitOpts(0).ListIndex = (n And vtLitMinSizeMask) \ vtLitMinSizeShift
        cboLitOpts(1).ListIndex = ((n And vtLitMinCountMask) \ vtLitMinCountShift) - 1
    End If
    rs.Close: Set rs = Nothing: Set tNode = Nothing

End Sub

Private Sub pvDoResize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    picContainer.Top = Me.ScaleHeight - picContainer.Height
    If picContainer.Top > tvChecksExt.Top Then
        On Error Resume Next
        tvChecksExt.Height = picContainer.Top - tvChecksExt.Top - tvChecksExt.Left
        If Err Then
            Err.Clear
        Else ' external ocx dimensions should be validated when DPI aware...
            cDpiAssist.SyncOcxToParent tvChecksExt
        End If
    End If

End Sub

Private Sub pvSetItemState(tNode As Node, hWnd As Long, lState As Long)
    ' hack to get treeview hItem with known offset from a node's ObjPtr
    CopyMemory udtItem.hItem, ByVal ((ObjPtr(tNode) Xor &H80000000) + XREF_NodeHItem) Xor &H80000000, 4
    udtItem.State = lState
    SendMessage hWnd, TVM_SETITEMW, 0, udtItem
End Sub

Private Function pvGetItemState(tNode As Node, hWnd As Long) As Long
    ' hack to get treeview hItem with known offset from a node's ObjPtr
    CopyMemory udtItem.hItem, ByVal ((ObjPtr(tNode) Xor &H80000000) + XREF_NodeHItem) Xor &H80000000, 4
    SendMessage hWnd, TVM_GETITEMW, 0, udtItem
    pvGetItemState = (udtItem.State And udtItem.stateMask)
End Function

Private Sub tvChecksExt_DblClick()
    ' determine if a specific node was double clicked
    If m_NodeDblClk <> 0 Then
        Dim tNode As Node, srcNode As Node
        Set tNode = tvChecksExt.Nodes(m_NodeDblClk).Parent
        m_NodeDblClk = 0
        If Not tNode Is Nothing Then
            Set tNode = tNode.Parent
            If Not tNode Is Nothing Then
                If LenB(tNode.Key) <> 0 Then
                    If tNode.Key = chrI & vtDupeLiterals Then
                        Set srcNode = tvChecksExt.Nodes(chrI & vtDupeDecs).Child.Next.Child.Next
                    ElseIf tNode.Key = chrI & vtDupeDecs Then
                        Set srcNode = tvChecksExt.Nodes(chrI & vtDupeLiterals).Child.Next.Child.Next
                    End If
                    If Not srcNode Is Nothing Then
                        Set tNode = tNode.Child.Next.Child
                        With tvChecksExt
                            Do Until srcNode Is Nothing
                                Set tNode = tNode.Next
                                pvSetItemState tNode, .hWnd, pvGetItemState(srcNode, .hWnd)
                                Set srcNode = srcNode.Next
                            Loop
                            .Refresh
                        End With
                        Set srcNode = Nothing
                    End If
                End If
                Set tNode = Nothing
            End If
        End If
    End If
End Sub

Private Sub tvChecksExt_NodeClick(ByVal Node As ComctlLib.Node)
    m_NodeDblClk = Node.Index   ' last node clicked on
End Sub

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
