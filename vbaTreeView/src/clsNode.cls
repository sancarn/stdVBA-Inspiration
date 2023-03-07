VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Build 026
'***************************************************************************
'
' Authors:  JKP Application Development Services, info@jkp-ads.com, https://www.jkp-ads.com
'           Peter Thornton, pmbthornton@gmail.com
'
' (c)2013-2015, all rights reserved to the authors
'
' You are free to use and adapt the code in these modules for
' your own purposes and to distribute as part of your overall project.
' However all headers and copyright notices should remain intact
'
' You may not publish the code in these modules, for example on a web site,
' without the explicit consent of the authors
'***************************************************************************

'-------------------------------------------------------------------------
' Module    : clsNode
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Holds all information of a node of the tree
'-------------------------------------------------------------------------
Option Explicit

Private mbExpanded As Boolean

Private mcolChildNodes As Collection

Private moParentNode As clsNode
Private moLastActiveNode As clsNode
Private moTree As clsTreeView

Private msKey As String
Private mvCaption
Private msControlTipText As Variant

Private mlChecked As Long        ' PT checkbox tristate boolean 0/-1 or 1 for null
Private mnIndex As Long             ' PT order added to Treeview's mcolNodes, won't change
Private mlVisIndex As Long          ' PT the visible order in the current view, changes with expand/collapse
Private mvIconMainKey               ' PT string name or numeric index as icon Key for the Image collection
Private mvIconExpandedKey           ' PT ditto for expanded icon
Private mlIconCnt As Long           ' PT number of icons availabel for this node 0, 1 or 2
Private msngTextWidth As Single     ' PT autosized text width before the node is widened beyond the frame
Private mlBackColor As Long         ' PT
Private mbBold As Boolean           ' PT
Private mlForeColor As Long         ' PT
Private mvTag

Private WithEvents mctlControl As MSForms.Label
Attribute mctlControl.VB_VarHelpID = -1
Private WithEvents mctlExpander As MSForms.Label
Attribute mctlExpander.VB_VarHelpID = -1
Private WithEvents moEditBox As MSForms.TextBox     ' PT editbox
Attribute moEditBox.VB_VarHelpID = -1
Private WithEvents mctlCheckBox As MSForms.Label    ' PT checkbox
Attribute mctlCheckBox.VB_VarHelpID = -1

Private mctlExpanderBox As MSForms.Label
Private mctlVLine As MSForms.Label  ' PT vertical line, only the first child node with children will have a vertical line
Private mctlHLine As MSForms.Label  ' PT horizontal line
Private mctlIcon As MSForms.Image   ' PT separate icon image control

Public Enum ndSortOrder
    ndAscending = 1
    ndDescending = 2
End Enum
Public Enum ndCompareMethod
    ndBinaryCompare = 0
    ndTextCompare = 1
End Enum
Public Enum ndMouse
    ndDown = 1
    ndUp = 2
    ndMove = 3
    ndBeforeDragOver = 4
    ndBeforeDropOrPaste = 5
End Enum

#If Mac Then
    Const mcFullWidth As Long = 800
#Else
    Const mcFullWidth As Long = 600
#End If

Private moLabSizer As MSForms.Label
Private Const mcBreak As String = "¶ " '"• "

'*********************
'* Public Properties *
'*********************

Public Property Get BackColor() As Long

    BackColor = mlBackColor ' if zero the treecaller will apply the frame container's backcolor

End Property

Public Property Let BackColor(lColor As Long)
'PT if lColor is written as 0/black, change it to 1 as 0 means default
    mlBackColor = lColor
    If mlBackColor = 0 Then mlBackColor = 1
    If Not mctlControl Is Nothing Then
        mctlControl.BackColor = lColor
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = mbBold
End Property

Public Property Let Bold(bBold As Boolean)
    mbBold = bBold
    If Not mctlControl Is Nothing Then
        mctlControl.Font.Bold = mbBold
    End If
End Property

Public Property Get Caption()
    Caption = mvCaption
End Property

Public Property Let Caption(ByVal vCaption)
Dim sngWd As Single
    mvCaption = vCaption
    If Not mctlControl Is Nothing Then
        
        mctlControl.Caption = CStr(vCaption)
        If mctlControl.AutoSize Then
            msngTextWidth = mctlControl.Width
        Else
            sngWd = mctlControl.Width
            mctlControl.AutoSize = True
            msngTextWidth = mctlControl.Width
            mctlControl.AutoSize = False
        End If
    End If
End Property

Public Property Get Checked()    ' PT
     ' Checked values are -1 true, 0 false, +1 mixed
     ' If TriState is enabled be careful not to return a potential +1 to a boolean or it'll coerce to True
    Checked = mlChecked
End Property

Public Property Let Checked(vChecked)  ' PT
    Dim bFlag As Boolean, bTriState As Boolean
    Dim lChecked As Long
    Dim cChild As clsNode

    ' Checked values are -1 true, 0 false, +1 mixed
    ' if vChecked is a boolean Checked will coerce to -1 or 0
    ' if vChecked is Null Checked is set as +1

    If VarType(vChecked) = vbBoolean Then
        lChecked = vChecked
    ElseIf IsNull(vChecked) Then
        lChecked = 1
    ElseIf vChecked >= -1 And vChecked <= 1 Then
        lChecked = vChecked
    End If

    bFlag = lChecked <> mlChecked
    mlChecked = lChecked

    If Not mctlCheckBox Is Nothing And bFlag Then
        moTree.Changed = True
        UpdateCheckbox
    End If
    
    If Not moTree Is Nothing Then    ' eg during clone
        bFlag = moTree.CheckBoxes(bTriState)
        If bTriState Then
            If ParentNode.Caption <> "RootHolder" Then
                ParentNode.CheckTriStateParent
            End If
            
            If Not ChildNodes Is Nothing Then
                For Each cChild In ChildNodes
                    cChild.CheckTriStateChildren mlChecked
                Next
            End If
        End If
    End If
    
End Property

Public Property Get Child() As clsNode
' PT Returns a reference to the first Child node, if any
    On Error Resume Next
    Set Child = mcolChildNodes(1)
End Property

Public Property Get ChildNodes() As Collection
    Set ChildNodes = mcolChildNodes
End Property

Public Property Set ChildNodes(colChildNodes As Collection)
    Set mcolChildNodes = colChildNodes
End Property

Public Property Get ControlTipText() As String
    ControlTipText = msControlTipText
End Property

Public Property Let ControlTipText(ByVal sControlTipText As String)
    msControlTipText = sControlTipText
    If Not mctlControl Is Nothing Then
        mctlControl.ControlTipText = msControlTipText
    End If
End Property

Public Property Get Expanded() As Boolean
    Expanded = mbExpanded
End Property

Public Property Let Expanded(ByVal bExpanded As Boolean)
    mbExpanded = bExpanded
    If Not Me.Expander Is Nothing Then
        UpdateExpanded bControlOnly:=False
    ElseIf Not Me.Control Is Nothing Then
        UpdateExpanded bControlOnly:=True
    End If
End Property

Public Property Get ForeColor() As Long
    ForeColor = mlForeColor
End Property

Public Property Let ForeColor(lColor As Long)
'PT if lColor is written as 0/black, change it to 1 as 0 means default
    mlForeColor = lColor
    If mlForeColor = 0 Then mlForeColor = 1
    If Not mctlControl Is Nothing Then
        mctlControl.ForeColor = lColor
    End If
End Property

Public Property Get FirstSibling() As clsNode
    If Not moParentNode Is Nothing Then    ' PT Root has no parent
        Set FirstSibling = moParentNode.GetChild(1)
    End If
End Property

Public Property Get LastSibling() As clsNode
    If Not moParentNode Is Nothing Then    ' PT Root has no parent
        Set LastSibling = moParentNode.GetChild(-1)    ' -1 flags GetChild to return the last Child
    End If
End Property

Public Property Get ImageExpanded()
' PT string name or numeric index for the main icon key
    ImageExpanded = mvIconExpandedKey
End Property

Public Property Let ImageExpanded(vImageExpanded)
' PT string name or numeric index for an expanded icon key
    On Error GoTo errExit
    If Not IsMissing(vImageExpanded) Then
        If Not IsEmpty(vImageExpanded) Then
            If Len(mvIconMainKey) = 0 Then
                mvIconMainKey = vImageExpanded
            End If
            mvIconExpandedKey = vImageExpanded
            mlIconCnt = 2
        End If
    End If
errExit:
End Property

Public Property Get ImageMain()
' PT string name or numeric index for the main icon key
    ImageMain = mvIconMainKey
End Property

Public Property Let ImageMain(vImageMain)
' PT string name or numeric index for the main icon key
    On Error GoTo errExit
    If Not IsMissing(vImageMain) Then
        If Not IsEmpty(vImageMain) Then
            mvIconMainKey = vImageMain
            If mlIconCnt = 0 Then mlIconCnt = 1
        End If
    End If
errExit:
End Property

Public Sub ImageUpdate()
Dim bFullWidth As Boolean
Dim vKey
Dim pic As StdPicture

    On Error GoTo errH
    If Me.hasIcon(vKey) Then
    '  error if Icon or Control is not yet created, silent abort, it'll update in the next BuildRoot
        If moTree.GetNodeIcon(vKey, pic, bFullWidth) Then
            If bFullWidth Then
                Me.Icon.Picture = pic
            Else
                Me.Control.Picture = pic
            End If
        ElseIf bFullWidth Then
            Set Me.Icon.Picture = Nothing
        Else
            Set Me.Control.Picture = Nothing
        End If
    End If

    Exit Sub
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
End Sub


Public Property Get Key() As String
    Key = msKey
End Property

Public Property Let Key(ByVal sKey As String)
    Dim bIsInMainCol As Boolean
    Dim i As Long
    Dim cTmp As clsNode

    On Error GoTo errH

    If Tree Is Nothing Then
        msKey = sKey
        Exit Property
    ElseIf msKey = sKey Or Len(sKey) = 0 Then
        Exit Property
    End If

    On Error Resume Next
    Set cTmp = Tree.Nodes.Item(sKey)
    On Error GoTo errH

    If Not cTmp Is Nothing Then
        Err.Raise 457    ' standard duplicate key error
    End If

    ' to change the Key, remove Me and add Me back where it was with the new key
    For Each cTmp In Tree.Nodes
        i = i + 1
        If cTmp Is Me Then
            bIsInMainCol = True
            Exit For
        End If
    Next

    If bIsInMainCol Then
        With Tree.Nodes
            .Remove i
            If i <= .Count Then
                .Add Me, sKey, i
            Else
                .Add Me
            End If
        End With
    Else
        ' Let Key  called by via move/copy
    End If

    msKey = sKey

    Exit Property
errH:
    Err.Raise Err.Number, "Let Key", Err.Description
End Property

Public Property Get Level() As Long
    Dim lLevel As Long
    Dim cNode As clsNode

    On Error GoTo errH
    lLevel = -1
    Set cNode = Me.ParentNode
    While Not cNode Is Nothing
        lLevel = lLevel + 1
        Set cNode = cNode.ParentNode
    Wend
    Level = lLevel
    Exit Property
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
End Property

Public Property Get NextNode() As clsNode    ' can't name this proc 'Next' in VBA
' PT return the next sibling if there is one
    Dim i As Long
    Dim cNode As clsNode

    With Me.ParentNode
        For Each cNode In .ChildNodes
            i = i + 1
            If cNode Is Me Then
                Exit For
            End If
        Next
        If .ChildNodes.Count > i Then
            Set NextNode = .ChildNodes(i + 1)
        End If
    End With
End Property

Public Property Get ParentNode() As clsNode
    Set ParentNode = moParentNode
End Property

Public Property Set ParentNode(oParentNode As clsNode)
    Set moParentNode = oParentNode
End Property

Public Property Get Previous() As clsNode
' PT return the previous sibling if there is one
    Dim i As Long
    Dim cNode As clsNode

    With Me.ParentNode
        For Each cNode In Me.ParentNode.ChildNodes
            i = i + 1
            If cNode Is Me Then
                Exit For
            End If
        Next
        If i > 1 Then
            Set Previous = .ChildNodes(i - 1)
        End If
    End With
End Property

Public Property Get Root() As clsNode
    Dim cTmp As clsNode
    Set cTmp = Me
    Do While Not cTmp.ParentNode.ParentNode Is Nothing
        Set cTmp = cTmp.ParentNode
    Loop
    Set Root = cTmp
End Property

Public Property Get Tag()
    Tag = mvTag
End Property

Public Property Let Tag(vTag)
    mvTag = vTag
End Property


'*****************************
'* Public subs and functions *
'*****************************

Public Function Sort(Optional ByVal ndOrder As ndSortOrder = ndAscending, _
                     Optional ByVal ndCompare As ndCompareMethod = ndTextCompare) As Boolean
' PT Sorts the child nodes,
'    returns True if the order has changed to flag Refresh should be called
    Dim sCaptions() As String
    Dim lStart As Long, lLast As Long, i As Long
    Dim colNodes As New Collection
    Dim bIsUnSorted As Boolean

    On Error GoTo errExit
    lStart = 1
    lLast = ChildNodes.Count    ' error if no childnodes to sort

    If lLast = 1 Then
        ' nothing to sort
        Exit Function
    End If

    ReDim idx(lStart To lLast) As Long
    ReDim sCaptions(lStart To lLast) As String
    For i = lStart To lLast
        idx(i) = i
        sCaptions(i) = ChildNodes.Item(i).Caption
    Next

    If ndOrder <> ndAscending Then ndOrder = -1    ' descending
    If ndCompare <> ndTextCompare Then ndCompare = ndBinaryCompare

    Call BinarySortIndexText(sCaptions(), lStart, lLast, idx, ndOrder, ndCompare)

    For i = lStart To lLast - 1
        If idx(i) <> idx(i + 1) - 1 Then
            bIsUnSorted = True
            Exit For
        End If
    Next

    If bIsUnSorted Then
        For i = lStart To lLast
            colNodes.Add ChildNodes(idx(i))
        Next
        Set ChildNodes = colNodes
        Sort = True
    End If

errExit:
'   Probably(?) any error was because there were no childnodes, no need to raise an error
End Function

Public Function AddChild(Optional sKey As String, _
                         Optional vCaption, _
                         Optional vImageMain, _
                         Optional vImageExpanded) As clsNode
    Dim bTriState As Boolean
    Dim cChild As clsNode

    On Error GoTo errH
    Set cChild = New clsNode

    With moTree.Nodes

        If Len(sKey) Then
100         .Add cChild, sKey
101
            cChild.Key = sKey
        Else
            .Add cChild
        End If

        cChild.Index = .Count
    End With

    If mcolChildNodes Is Nothing Then
        Set mcolChildNodes = New Collection
    End If

    mcolChildNodes.Add cChild

    With cChild
        If Not IsMissing(vImageMain) Then
            If Len(vImageMain) Then
                .ImageMain = vImageMain
            End If
        End If

        If Not IsMissing(vImageExpanded) Then
            If Len(vImageExpanded) Then
                .ImageExpanded = vImageExpanded
            End If
        End If

        .Caption = vCaption
        
        If mlChecked = -1 Then ' -1 = true, +1 = mixed
            If moTree.CheckBoxes(bTriState) Then
                If bTriState Then
                    .Checked = True
                End If
            End If
        End If
        
        Set .Tree = moTree
        Set .ParentNode = Me
    End With

    Set AddChild = cChild

    Exit Function
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If

    If Erl = 100 And Err.Number = 457 Then
        Err.Raise vbObjectError + 1, "clsNode.AddChild", "Duplicate key: '" & sKey & "'"
    Else
        Err.Raise Err.Number, "clsNode.AddChild", Err.Description
    End If
End Function

Public Function ChildIndex(sKey As String) As Long
'-------------------------------------------------------------------------
' Procedure : ChildIndex
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Returns the index of a childnode using its key
'-------------------------------------------------------------------------
    Dim cNode As clsNode
    Dim lCt As Long
    For Each cNode In mcolChildNodes
        lCt = lCt + 1
        If sKey = cNode.Key Then
            ChildIndex = lCt
            Set cNode = Nothing
            Exit Function
        End If
    Next
    Set cNode = Nothing
End Function

Public Function FullPath() As String
' PT, get all the grand/parent keys
' assumes use of key

    Dim s As String
    Dim cNode As clsNode

    On Error GoTo errDone
    s = Me.Key
    Set cNode = Me

    While Err.Number = 0
        Set cNode = cNode.ParentNode
        s = cNode.Key & "\" & s
    Wend

errDone:
    FullPath = s
End Function

Public Function GetChild(vKey As Variant) As clsNode
'-------------------------------------------------------------------------
' Procedure : GetChild
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Returns a childnode using its key
'-------------------------------------------------------------------------
    Dim cNode As clsNode
    Dim lIdx As Long

    If VarType(vKey) = vbString Then

        For Each cNode In mcolChildNodes
            If vKey = cNode.Key Then
                Set GetChild = cNode
                Set cNode = Nothing
                Exit Function
            End If
        Next

    ElseIf Not mcolChildNodes Is Nothing Then
        lIdx = vKey
        If lIdx = -1 Then
            lIdx = mcolChildNodes.Count
        End If
        If lIdx > 0 Then
            Set GetChild = mcolChildNodes(lIdx)
        Else: Set mcolChildNodes = Nothing
        End If
    End If

    Set cNode = Nothing
End Function


'*************************************************************************
'*    Friend Properties, Subs & Funtions                                 *
'*    ** these procedures are visible throughout the project but should  *
'*    ** only be used to communicate with the TreeView, ie clsTreeView   *
'*************************************************************************

Friend Property Get Control() As MSForms.Label
    Set Control = mctlControl
End Property

Friend Property Set Control(ctlControl As MSForms.Label)
    Set mctlControl = ctlControl
    If Not mctlControl Is Nothing Then
        If Not moTree Is Nothing Then
            Set mctlControl.Font = moTree.TreeControl.Font
        Else
            Stop
        End If
    End If
End Property

Friend Property Get Index() As Long    ' PT
    Index = mnIndex
End Property

Friend Property Let Index(idx As Long)
' PT Index: the order this node was added to Treeview's collection mcolNodes
'    Index will never increase but may decrement if previously added nodes are removed
    mnIndex = idx
End Property

Friend Property Let VisIndex(lVisIndex As Long)
    mlVisIndex = lVisIndex
End Property

Friend Property Get VisIndex() As Long    ' PT
    VisIndex = mlVisIndex
End Property

Friend Property Get Tree() As clsTreeView
    Set Tree = moTree
End Property

Friend Property Set Tree(oTree As clsTreeView)
    Set moTree = oTree
End Property

Friend Property Get Checkbox() As MSForms.Control
    Set Checkbox = mctlCheckBox
End Property

Friend Property Set Checkbox(oCtl As MSForms.Control)
    Set mctlCheckBox = oCtl
End Property

Friend Property Get Expander() As MSForms.Label
    Set Expander = mctlExpander
End Property

Friend Property Set Expander(ctlExpander As MSForms.Label)
    Set mctlExpander = ctlExpander
End Property

Friend Property Get ExpanderBox() As MSForms.Label
    Set ExpanderBox = mctlExpanderBox
End Property

Friend Property Set ExpanderBox(ctlExpanderBox As MSForms.Label)
    Set mctlExpanderBox = ctlExpanderBox
End Property

Friend Property Set HLine(ctlHLine As MSForms.Label)
    Set mctlHLine = ctlHLine
End Property

Friend Property Get HLine() As MSForms.Label
    Set HLine = mctlHLine
End Property

Friend Property Set Icon(ctlIcon As MSForms.Image)
    Set mctlIcon = ctlIcon
End Property

Friend Property Get Icon() As MSForms.Image
    Set Icon = mctlIcon
End Property

Friend Property Get TextWidth() As Single
    TextWidth = msngTextWidth
End Property

Friend Property Let TextWidth(sngTextWidth As Single)
    msngTextWidth = sngTextWidth
End Property

Friend Property Get VLine() As MSForms.Label
    Set VLine = mctlVLine
End Property

Friend Property Set VLine(ctlVLine As MSForms.Label)
    Set mctlVLine = ctlVLine
End Property

Friend Sub CheckTriStateParent()
' PT set triState value of parent according to its childnodes' values
    Dim alChecked(-1 To 1) As Long
    Dim cChild As clsNode

    If Not ChildNodes Is Nothing Then
        For Each cChild In ChildNodes
            alChecked(cChild.Checked) = alChecked(cChild.Checked) + 1
        Next
        If alChecked(1) Then
            alChecked(1) = 1
        ElseIf alChecked(-1) = ChildNodes.Count Then
            alChecked(1) = -1
        ElseIf alChecked(0) = ChildNodes.Count Then
            alChecked(1) = 0
        Else
            alChecked(1) = 1
        End If
        
        If mlChecked <> alChecked(1) Then
            mlChecked = alChecked(1)
            UpdateCheckbox
        End If

    End If
    
    If Not Me.Caption = "RootHolder" Then
        If Not ParentNode.ParentNode Is Nothing Then
            ParentNode.CheckTriStateParent
        End If
    End If

End Sub

Friend Sub CheckTriStateChildren(lChecked As Long)
' PT, make checked values of children same as parent's
'     only called if triState is enabled
Dim cChild As clsNode

    If mlChecked <> lChecked Then
        mlChecked = lChecked
        UpdateCheckbox
    End If
    
    If Not ChildNodes Is Nothing Then
        For Each cChild In ChildNodes
            cChild.CheckTriStateChildren lChecked
        Next
    End If
End Sub

Friend Function hasIcon(vKey) As Boolean
' PT get the appropriate icon key/index, if any
    If mlIconCnt = 2 And mbExpanded Then
        vKey = mvIconExpandedKey
        hasIcon = True    'Not IsEmpty(vKey) '(True
    ElseIf mlIconCnt Then
        vKey = mvIconMainKey
        hasIcon = True    'Not IsEmpty(vKey)
    End If
End Function

'Friend Sub EditBox(bEnterEdit As Boolean)    '  PT new in 006PT2 ,,move to clsTreView?
''-------------------------------------------------------------------------
'' Procedure : moCtl_Click
'' Author    : Peter Thornton
'' Created   : 20-01-2013
'' Purpose   : Enter/exit Editmode, show/hide the edit textbox
''-------------------------------------------------------------------------
'    On Error Resume Next
'    Set moEditBox = moTree.TreeControl.Controls("EditBox")
'    On Error GoTo 0
'
'    If bEnterEdit Then
'
'        If moEditBox Is Nothing Then
'            Set moEditBox = moTree.TreeControl.Controls.Add("forms.textbox.1", False)
'            moEditBox.Name = "EditBox"
'        End If
'
'        With moEditBox
'            .Left = Control.Left - 3
'            .Top = Control.Top - 1.5
'            .AutoSize = True
'            .BorderStyle = fmBorderStyleSingle
'            .Text = Caption
'            Control.Visible = False    ' hide the node label while editing
'            .ZOrder 0
'            .Visible = True
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'        End With
'
'    ElseIf Not moEditBox Is Nothing Then
'        ' exit editmode
'        If Not moEditBox Is Nothing Then
'            ' error if moEditBox has already been removed
'            On Error Resume Next
'            moEditBox.Visible = False
'            moEditBox.Text = ""
'            Set moEditBox = Nothing
'        End If
'        Control.Visible = True
'
'    End If
'End Sub
Friend Sub EditBox(bEnterEdit As Boolean)    '  PT new in 006PT2 ,,move to clsTreView?
'-------------------------------------------------------------------------
' Procedure : moCtl_Click
' Author    : Peter Thornton
' Created   : 20-01-2013. Ammended 026 to resize/reposition the editbox within the confines of the frame container
' Purpose   : Enter/exit Editmode, show/hide the edit textbox
'-------------------------------------------------------------------------

Dim wd As Single, ht As Single
    On Error Resume Next
    Set moEditBox = moTree.TreeControl.Controls("EditBox")
    Set moLabSizer = moTree.TreeControl.Controls("LabSizer")
    On Error GoTo 0

    If bEnterEdit Then

        If moEditBox Is Nothing Then
            Set moEditBox = moTree.TreeControl.Controls.Add("forms.textbox.1", "EditBox")
            Set moLabSizer = moTree.TreeControl.Add("forms.Label.1", "LabSizer")
            moLabSizer.Left = moTree.TreeControl.Left
            moLabSizer.Top = moTree.TreeControl.Top - 12
            moLabSizer.WordWrap = False
            moLabSizer.AutoSize = True
            moLabSizer.Visible = False
        End If

        With moEditBox
            .Left = mctlControl.Left - 3
            .Top = mctlControl.Top - 1.5
            .AutoSize = False
            .Width = moTree.TreeControl.Width
            .BorderStyle = fmBorderStyleSingle
            If Len(Me.Caption) Then
                .Text = Me.Caption
            Else
                moEditBox_Change
            End If
            wd = .Width
            ht = .Height
            .ZOrder 0
            .Visible = True
            mctlControl.Visible = False
            .SelStart = 0
            .SelLength = Len(.Text)

            .SetFocus
            .Width = wd
            .Height = ht
        End With

    ElseIf Not moEditBox Is Nothing Then
        ' exit editmode
        If Not moEditBox Is Nothing Then
            ' error if moEditBox has already been removed
            On Error Resume Next
            moEditBox.Visible = False
            moEditBox.Text = ""
'            If Len(moEditBox.Tag) Then
'                moTree.TreeControl.ScrollTop = CSng(moEditBox.Tag)
'                moEditBox.Tag = ""
'            End If
            Set moEditBox = Nothing
            Set moLabSizer = Nothing
        End If
        mctlControl.Visible = True
    End If

End Sub

Friend Function RemoveChild(cNode As clsNode, Optional bReverse As Boolean) As Boolean
'PT remove a node from the collection,
'   note, this is only one part of the process of removing a node

Dim lCt As Long
Dim i As Long
Dim cTmp As clsNode
    On Error GoTo errH

    If bReverse Then
        For i = mcolChildNodes.Count To 1 Step -1
            Set cTmp = mcolChildNodes(i)
            If mcolChildNodes(i) Is cNode Then
                mcolChildNodes.Remove i
                RemoveChild = True
                Exit For
            End If
        Next
    Else
        For Each cTmp In mcolChildNodes
            lCt = lCt + 1
            If cTmp Is cNode Then
                mcolChildNodes.Remove lCt
                RemoveChild = True
                Exit For
            End If
        Next
    End If
    
    If mcolChildNodes.Count = 0 Then
        Set mcolChildNodes = Nothing
        Me.Expanded = False
    End If

    Exit Function
errH:
    Err.Raise vbObjectError, "RemoveChild", Err.Description
End Function

Friend Sub RemoveNodeControls()
    Dim cChild As clsNode
    If Not ChildNodes Is Nothing Then
        For Each cChild In ChildNodes
            cChild.RemoveNodeControls
        Next
    End If
    DeleteNodeControls False
End Sub

Friend Sub TerminateNode(Optional bDeleteNodeControls As Boolean)
'-------------------------------------------------------------------------
' Procedure : TerminateNode
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Terminates the class instance
'-------------------------------------------------------------------------
    Dim cChild As clsNode
    'Instead of the Terminate event of the class we use this public
    'method so it can be explicitly called by parent classes.
    'This is done because to break the two way or circular references
    'between the parent child classes.
    
    'The most important call in this routine is to destroy the reference
    'between this node class and the parent treeview class -
    '    < Set moTree = Nothing >
    'Once all the moTree references to have been destroyed everything else will
    ' 'tear down' normally

    If Not ChildNodes Is Nothing Then
        For Each cChild In ChildNodes
            ' recursively drill down to all child nodes in this branch
            cChild.TerminateNode bDeleteNodeControls
        Next
    End If

    ' If deleting individual nodes while the treeview is running we also want to
    ' remove all associated controls as well as removing references
    
    If bDeleteNodeControls Then
        DeleteNodeControls True
        If bDeleteNodeControls Then
            Index = -1
        End If
    End If

    Set mcolChildNodes = Nothing
    Set moTree = Nothing
End Sub


'******************************
'* Private subs and functions *
'******************************

Private Sub BinarySortIndexText(sCaptions() As String, ByVal lStart As Long, ByVal lEnd As Long, ByRef idx() As Long, ndOrder As Long, ndCompare As ndCompareMethod)
' PT sorts the index array based on the string array
    Dim lSmall As Long, lLarge As Long, sMid As String, lTmp As Long

    lSmall = lStart
    lLarge = lEnd
    sMid = sCaptions(idx((lSmall + lLarge) / 2))

    Do While lSmall <= lLarge
        Do While (StrComp(sCaptions(idx(lSmall)), sMid, ndCompare) = -ndOrder And lSmall < lEnd)
            lSmall = lSmall + 1
        Loop
        Do While (StrComp(sCaptions(idx(lLarge)), sMid, ndCompare) = ndOrder And lLarge > lStart)
            lLarge = lLarge - 1
        Loop
        If lSmall <= lLarge Then
            lTmp = idx(lSmall)
            idx(lSmall) = idx(lLarge)
            idx(lLarge) = lTmp
            lSmall = lSmall + 1
            lLarge = lLarge - 1
        End If
    Loop

    If lStart <= lLarge Then
        Call BinarySortIndexText(sCaptions(), lStart, lLarge, idx, ndOrder, ndCompare)
    End If
    If lSmall <= lEnd Then
        Call BinarySortIndexText(sCaptions(), lSmall, lEnd, idx, ndOrder, ndCompare)
    End If
End Sub

Private Sub DeleteNodeControls(bClearIndex As Boolean)
'PT Delete all controls linked to this node

    On Error GoTo errH

    With moTree.TreeControl.Controls
        If Not mctlControl Is Nothing Then
            .Remove mctlControl.Name
            Set mctlControl = Nothing
            If Not mctlHLine Is Nothing Then
                .Remove mctlHLine.Name
                Set mctlHLine = Nothing
            End If
            If Not mctlIcon Is Nothing Then
                .Remove mctlIcon.Name
                Set mctlIcon = Nothing
            End If
            If Not mctlIcon Is Nothing Then
                .Remove mctlIcon.Name
                Set mctlIcon = Nothing
            End If
        End If

        If Not mctlExpander Is Nothing Then
            .Remove mctlExpander.Name
            Set mctlExpander = Nothing
        End If
        If Not mctlExpanderBox Is Nothing Then
            .Remove mctlExpanderBox.Name
            Set mctlExpanderBox = Nothing
        End If
        If Not mctlVLine Is Nothing Then
            .Remove mctlVLine.Name
            Set mctlVLine = Nothing
        End If


        If Not moEditBox Is Nothing Then
            .Remove moEditBox.Name
            Set moEditBox = Nothing
        End If
        If Not mctlCheckBox Is Nothing Then
            .Remove mctlCheckBox.Name
            Set mctlCheckBox = Nothing
        End If

        If Not Me.ParentNode Is Nothing Then
            ' if Me is the last child delete parent's expander and VLine (if it has one)
            If FirstSibling Is LastSibling Then

                If Not Me.ParentNode.VLine Is Nothing Then
                    .Remove Me.ParentNode.VLine.Name
                    Set Me.ParentNode.VLine = Nothing
                End If
                
                If Not Me.ParentNode.ExpanderBox Is Nothing Then
                    .Remove Me.ParentNode.ExpanderBox.Name
                    Set Me.ParentNode.ExpanderBox = Nothing
                End If

                If Not Me.ParentNode.Expander Is Nothing Then
                    .Remove Me.ParentNode.Expander.Name
                    Set Me.ParentNode.Expander = Nothing
                End If

                Me.ParentNode.Expanded = False

            End If

        End If

    End With

    If bClearIndex Then
        Me.Index = -1  ' flag this node to be removed from mcolNodes in NodeRemove
    End If

    Exit Sub
errH:
    ' Stop
    Resume Next
End Sub

Private Function UpdateCheckbox()
Dim pic As StdPicture

    If Not mctlCheckBox Is Nothing Then
        With mctlCheckBox
            If moTree.GetCheckboxIcon(mlChecked, pic) Then
                .Picture = pic
            Else
                .Caption = IIf(mlChecked, "a", "")
                If (mlChecked = 1) <> (.ForeColor = RGB(180, 180, 180)) Then
                    .ForeColor = IIf(mlChecked = 1, RGB(180, 180, 180), vbWindowText)
                End If
            End If

        End With
    End If
    
    moTree.NodeEventRouter Me, "Checkbox", 1
    
End Function

Private Sub UpdateExpanded(bControlOnly As Boolean)
'-------------------------------------------------------------------------
' Procedure : UpdateExpanded
' Author    : Peter Thornton
' Created   : 27-01-2013
' Purpose   : Called via an Expander click or arrow keys
'             Updates the Expanded property and changes +/- caption
'-------------------------------------------------------------------------
    Dim bFullWidth As Boolean
    Dim vKey
    Dim pic As StdPicture

    If Not bControlOnly Then
        With Me.Expander
            If moTree.GetExpanderIcon(mbExpanded, pic) Then
                .Picture = pic
            Else
                If mbExpanded Then
                    .Caption = "-"
                Else
                    .Caption = "+"
                End If
            End If
        End With
    End If

    On Error GoTo errExit
    If Me.hasIcon(vKey) Then
        If moTree.GetNodeIcon(vKey, pic, bFullWidth) Then
            If bFullWidth Then
                Me.Icon.Picture = pic   ' potential error if Icon is nothing, let error abort
            Else
                Me.Control.Picture = pic
            End If
        End If
    End If
errExit:
End Sub


'***********************
'*   Node Events       *
'***********************

Private Sub mctlCheckBox_Click()    ' PT new in 006PT2
'-------------------------------------------------------------------------
' Procedure : moCtl_Click
' Author    : Peter Thornton
' Created   : 20-01-2013
' Purpose   : Event fires when a Checkbox label is clicked
'-------------------------------------------------------------------------
    If moTree.EditMode(Me) Then
        ' exit editmode if in editmode
        moTree.EditMode(Me) = False
    End If
    If mlChecked = 0 Then

        Checked = -1
    Else
        Checked = 0
    End If

    If Not moTree.ActiveNode Is Me Then
        Set moTree.ActiveNode = Me

        moTree.NodeEventRouter Me, "Caption", 1    ' share the checkbox click event
    End If
End Sub

Private Sub mctlControl_Click()
'-------------------------------------------------------------------------
' Procedure : mctlControl_Click
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Event fires when a treebranch is clicked
'-------------------------------------------------------------------------

' PT the call to NodeClick will raise the click event to the form
Dim bFlag As Boolean

    If Not moLastActiveNode Is Nothing Then
        moLastActiveNode.Control.BorderStyle = fmBorderStyleNone
        Set moLastActiveNode = Nothing
        bFlag = True
    End If

    If moTree.ActiveNode Is Nothing Then
        Set moTree.ActiveNode = Me
        bFlag = True
    ElseIf Not bFlag Then
        bFlag = mctlControl.BorderStyle <> fmBorderStyleNone
    End If
    
    If Not moTree.ActiveNode Is Me Or bFlag Then
        ' only raise the event the first time the node is activated
         moTree.NodeEventRouter Me, "Caption", 1 ' tvClick
         ' if preferred the click event is always raised to the form (even if the
         ' node was previously active) simply comment or remove this If/EndIf check
    End If

End Sub

Private Sub mctlControl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' PT  a node label has been double-clicked, enter edit-mode if manual editing is enabled
    Dim bDummy As Boolean

        If moTree.EnableLabelEdit(bDummy) Then
            moTree.EditMode(Me) = True
            EditBox bEnterEdit:=True
        Else
            ' moTree.NodeEventRouter Me, "Caption", tvDblClick
        End If

End Sub

Private Sub mctlControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PT temporarily activate and highlight the MouseDown node and a grey border to the previous activenode
'   MouseUp and Click events will confirm the action or reset the previous active node
Dim bFlag As Boolean

    If moTree.ActiveNode Is Me Then
        bFlag = Me.Control.BackColor = vbHighlight
       ' bFlag = bFlag Or Me.Control.BorderStyle = fmBorderStyleSingle ' in Access this should be uncommented
    End If
    
    If Not bFlag Then
        Set moLastActiveNode = moTree.ActiveNode
        Set moTree.ActiveNode = Me
        If Not moLastActiveNode Is Nothing Then
            moLastActiveNode.Control.BorderStyle = fmBorderStyleSingle
            moLastActiveNode.Control.BorderColor = RGB(200, 200, 200)
        End If
    End If

    If moTree.EditMode(Me) Then
        ' if any node is in edit mode exit edit mode
        moTree.EditMode(Me) = False
    End If
    
    ' moTree.NodeEventRouter Me, "Caption", tvDown, Button, Shift, X, Y
End Sub

Private Sub mctlControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
' PT MouseUp fires before the Click event, at this point we don't know 100% if user
'    definately wants to activate the MouseDown node. If user drags the mouse off the MouseDown node the
'    Click event will not fire which means user wants to cancel and revert to the previous activenode.
'
'    If MouseUp occurs with the cursor not over the node reset the original activenode

Dim bFlag As Boolean
Dim bMouseIsOver As Boolean
Dim bMoveCopy As Boolean

    If Not moLastActiveNode Is Nothing Then
        With Me.Control
            ' is the mouse over the node or within a pixel of it
            bMouseIsOver = (X >= -1 And X <= .Width + 1) And (Y >= -1 And Y <= .Height + 1)
        End With
        
        If Not bMouseIsOver Then
            ' if the last-activenode was marked for MoveCopy we will need to reset it
            bFlag = moLastActiveNode Is moTree.MoveCopyNode(bMoveCopy)

            ' reset the original activenode
            moLastActiveNode.Control.BorderStyle = fmBorderStyleNone
            Set moTree.ActiveNode = moLastActiveNode

            If bFlag Then
                Set moTree.MoveCopyNode(bMoveCopy) = moLastActiveNode
            End If

            Set moLastActiveNode = Nothing
            
        ElseIf Button = 2 Then
            ' the click event doesn't fire with right click so explicitly call it
            mctlControl_Click
        End If
    End If
    
    ' moTree.NodeEventRouter Me, "Caption", tvUp, Button, Shift, X, Y

End Sub

Private Sub mctlExpander_Click()
'
    Expanded = Not Expanded
    If moTree.EditMode(Me) Then
        ' if any node is in edit mode exit edit mode
        moTree.EditMode(Me) = False
    End If
    moTree.NodeEventRouter Me, "Expander", 1
End Sub

Private Sub moEditBox_Change()
' PT resize horizontally & vertically to accommodate the changed text but restrict size to a bit less than the container frame
'    Autosize is good for width only if multiline=false, but it's problematic to size both width & height to text without this fix
Dim i As Long, ub As Long
Dim lSelSt As Long
Dim lRows As Long
Dim ht As Single, wd As Single
Dim sngMaxW As Single, sngMaxHt As Single, sngMaxR As Single, sngVisR
Dim sngIdealL, sngIdealTop As Single
Dim sEdit As String
Dim splt, v, splt2
Static bExit As Boolean
Const cPad As Long = 6

    If bExit Then Exit Sub

    On Error GoTo errH

    With moTree.TreeControl
        sngMaxW = .InsideWidth - cPad
        sngMaxHt = IIf(.ScrollHeight, .ScrollHeight, .InsideHeight) - cPad
        sngMaxR = IIf(.ScrollWidth > .InsideWidth, .ScrollWidth, .InsideWidth) - cPad
        sngVisR = .ScrollLeft + .InsideWidth
    End With

    sngIdealL = mctlControl.Left - 1.5
    sngIdealTop = mctlControl.Top - 1.5

    sEdit = moEditBox

    If InStr(1, sEdit, vbNewLine) Then
        splt = Split(sEdit, vbNewLine)
    ElseIf InStr(1, sEdit, vbCr) Then
        splt = Split(sEdit, vbCr)
    ElseIf InStr(1, sEdit, vbLf) Then
        splt = Split(sEdit, vbLf)
    Else
        ReDim splt(0)
        splt(0) = sEdit
    End If
    ub = UBound(splt)
    
    With moLabSizer

        .Caption = "A"
        ht = .Height
        .Caption = ""
        lRows = 1

        For i = 0 To ub
            splt2 = Split(splt(i), " ")
            For Each v In splt2
                .Caption = .Caption & v & " "
                .AutoSize = True
                If .Width > sngMaxW - 9 Then
                    lRows = lRows + 1
                    .Caption = v & " "
                    If .Width > wd Then
                        wd = .Width
                        If .Width > sngMaxW Then
                            lRows = lRows + 1
                        End If
                    End If
                    
                ElseIf .Width > wd Then
                    wd = .Width
                End If
            Next
            If i < ub Then
                lRows = lRows + 1
                .Caption = ""
            End If
        Next
        If .Width > wd Then wd = .Width
        If wd + 15 > sngMaxW Then wd = sngMaxW Else wd = wd + 15
    End With

    bExit = True

    With moEditBox

        If lRows > 1 Then
            .AutoSize = True
            .MultiLine = True
            lSelSt = .SelStart
            .SelStart = 0
            .SelStart = lSelSt

            If ht * lRows + 9 > sngMaxHt Then
                .Height = sngMaxHt - 3
            Else
                .Height = ht * lRows + 9
            End If

            If sngIdealTop + .Height > sngMaxHt - 3 Then
                .Top = sngMaxHt - .Height + 3
            Else
                .Top = sngIdealTop
            End If
            .Width = wd

        Else
            .MultiLine = False
            .AutoSize = True
            .MultiLine = True
            .Top = mctlControl.Top - 1.5
        End If

        
        If sngIdealL + .Width > sngMaxR Then
            .Left = sngMaxR - .Width + 3
        Else
            .Left = sngIdealL
        End If
        sngVisR = (.Left + .Width) - sngVisR
        If sngVisR > 0 Then
            moTree.TreeControl.ScrollLeft = moTree.TreeControl.ScrollLeft + sngVisR
        End If
          
    End With
done:
    bExit = False
    Exit Sub
    
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
    
    EditBox False
    Resume done
End Sub


Private Sub moEditBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)    'PT
' PT Textbox key events to Enter or Esc the Editbox,   006PT2
' PT ammended 026 to resize/reposition the editbox within the confines of the frame container
'    remove any CR & LF to ensure the node label is a single line but retain line breaks in the property if bMultiLine

    Dim bCancel As Boolean
    Dim bSort As Boolean, bMultiLine As Boolean
    Dim sNewText As String
    

    If KeyCode = vbKeyReturn And Shift = 0 Then
        sNewText = moEditBox
        If sNewText = Me.Caption Then
            KeyCode = vbKeyEscape
        Else
            ' note: sNewText could be parsed where the RaiseAfterLabelEdit is received
            bCancel = moTree.RaiseAfterLabelEdit(Me, sNewText)
            If Not bCancel Then
                bCancel = moTree.EnableLabelEdit(bSort, bMultiLine)
                
                If InStr(1, sNewText, vbCr) Or InStr(1, sNewText, vbLf) Then

                    #If Mac Then
                        '' Mac doesn't seem to use vbCrLf/vbNewLine, only vbCr or vbLf (?)
                        If bMultiLine Then
                            Me.Caption = Replace(Replace(sNewText, vbCr, mcBreak), vbLf, mcBreak)
                        Else
                            Me.Caption = Replace(Replace(sNewText, vbCr, "  "), vbLf, "  ")
                        End If
                    #Else
                        If bMultiLine Then
                            Me.Caption = Replace(Replace(sNewText, vbCrLf, mcBreak), vbLf, mcBreak)
                        Else
                            Me.Caption = Replace(Replace(sNewText, vbCrLf, "  "), vbLf, "  ")
                        End If
                    #End If

                    If bMultiLine Then
                    ' Drectly write the original multiline to the Caption, but don't use Let-Caption Property
                        mvCaption = sNewText
                    End If
                    
                Else
                    Me.Caption = sNewText
                    
                End If

                mctlControl.WordWrap = False
                mctlControl.AutoSize = True
                TextWidth = mctlControl.Width
                
                If TextWidth < mcFullWidth And moTree.FullWidth Then
                    Control.Width = mcFullWidth
                End If
                
                moTree.Changed = True

                If bSort Then
                    If Me.ParentNode.Sort Then
                        moTree.Refresh
                    End If
                End If
                
                moTree.SetScrollbars True ' replaced moTree.SetScrollbars in 026
            End If
            EditBox False
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        moTree.EditMode(Me) = False
        EditBox False
    End If

End Sub


Private Sub Class_Initialize()
' default properties
    mbExpanded = True  ' default
    
    #If DebugMode = 1 Then
        gClsNodeInit = gClsNodeInit + 1    ' PT, for testing only, remove, see ClassCounts() in modDemo
    #End If
End Sub

Private Sub Class_Terminate()
    #If DebugMode = 1 Then
        gClsNodeTerm = gClsNodeTerm + 1    ' PT, for testing,
    #End If
    Set moTree = Nothing
End Sub


