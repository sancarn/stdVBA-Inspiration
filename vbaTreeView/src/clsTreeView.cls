VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Build 026.4
'***************************************************************************
'
' Authors:  JKP Application Development Services, info@jkp-ads.com, https://www.jkp-ads.com
'           Peter Thornton, pmbthornton@gmail.com
'
' (c)2013-2017, all rights reserved to the authors
'
' You are free to use and adapt the code in these modules for
' your own purposes and to distribute as part of your overall project.
' However all headers and copyright notices should remain intact
'
' You may not publish the code in these modules, for example on a web site,
' without the explicit consent of the authors
'***************************************************************************

'-------------------------------------------------------------------------
' Module    : clsTreeView
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Creates a VBA Treeview control in a frame on your UserForm
'-------------------------------------------------------------------------
Option Explicit
#Const HostProject = "Excel" ', or Access or Word

Public WithEvents TreeControl As MSForms.Frame
Attribute TreeControl.VB_VarHelpID = -1

Private mbInActive                  'PT the treeview is not in focus
Private mbAlwaysRedesign As Boolean 'PT temporary flag to force mbRedesign=true, see Move()
Private mbAutoSort As Boolean       'PT sort siblings after manual edit
Private mbChanged As Boolean        'PT "dirty", user has edited node(s)
Private mbCheckboxes As Boolean     'PT show checkboxes
Private mbLabelEdit As Boolean      'PT allow manual editing with F2 and double click
Private mbTriState As Boolean       'PT enable tripple state checkboxes
Private mbCheckboxImage As Boolean  'PT determins if icons are used for checkboxes
Private mbEditMode As Boolean       'PT flag if in editmode
Private mbFullWidth As Boolean      'PT use separate image controls for icons, can highlight nodes to full width
Private mbGotIcons As Boolean       'PT got a collection of images
Private mbExpanderImage As Boolean  'PT determines if icons will be used for collapse/expand controls
Private mbKeyDown As Boolean        'PT Enter-keyup in a Textbox occurs when next control gets focus
Private mbMove As Boolean           'PT flag intention of the MoveCopyNode
Private mbMultiLine As Boolean
Private mbRedesign As Boolean       'PT flag to reset all dim's after changing NodeHeight or Indentation at runtime
Private mbRootButton As Boolean     'PT Root has an expander button
Private mbShowExpanders As Boolean  'PT Show +/- buttons
Private mbShowLines As Boolean      'PT determines if lines will be created and shown
Private mlBackColor As Long         'PT frameholder's backcolor
Private mlForeColor As Long         'PT frameholder's ForeColor
Private mlLabelEdit As Long         'PT 0-Automatic, 1-Manual can't be edited
Private mlLineColor As Long         'PT Lines color
Private mlNodesCreated As Long      'PT in/de-cremented as nodes are added/deleted from mcolNodes
Private mlNodesDeleted As Long      'PT incremented as node.controls are deleted, purpose to give unique id for control names
Private mlVisCount As Long          'PT incremented from zero as each node is displayed
Private mlVisOrder() As Long        'PT an index array to identify displayed nodes in the order as displayed
Private msAppName As String         'JKP: Title of messageboxes
Private msngChkBoxPad As Single     'PT offset if using checkboxes
Private msngChkBoxSize As Single    'PT checkbox size
Private msngIndent As Single        'PT default 15
Private msngLineLeft As Single      'PT Left pos of Root H & V lines, 3 + alpha
Private msngNodeHeight As Single    'JKP: vertical distance between nodes
Private msngRootLine As Single      'PT if mbRootButton, same as msngIndent, else 0
Private msngTopChk As Single        'PT top checkbox  (these "tops" are offsets down from the top a given node)
Private msngTopExpB As Single       'PT top expander button (a label)
Private msngTopExpT As Single       'PT top expander text (a label)
Private msngTopHV As Single         'PT top for Horiz' & Vert' lines (mid height of a node + top padding))
Private msngTopIcon As Single       'PT top icon
Private msngTopLabel As Single      'PT top node label, if font height less than NodeHeight
Private msngVisTop As Single        'PT activenode top relative to scroll-top
Private msngMaxWidths() As Single   'PT array, max width of text in each level, helps determine scroll-width
Private moActiveNode As clsNode     'JKP: refers to the selected node
Private moEditNode As clsNode       'PT the node in EditMode
Private moMoveNode As clsNode       'PT node waiting to be moved
Private moRootHolder As clsNode     'PT parent for the root node(s), although a clsNode it's not a real node
Private mcolIcons As Collection     'PT collection of stdPicture objects, their names as keys
Private mcolNodes As Collection     'JKP: global collection of all the nodes
Private moCheckboxImage(-1 To 1) As StdPicture   'PT checkbox true/false/triState icons
Private moExpanderImage(-1 To 0) As StdPicture   'PT collapse/expand icons
#If HostProject = "Access" Then
  Private moForm As Access.Form     'PT the main form, eg to return debug stats to the caption
#Else
  Private moForm As MSForms.UserForm
#End If
''-----------------------------------------------------------

Public Enum tvMouse
    tvClick = 1
    tvDblClick = 2
    tvDown = 3
    tvMove = 4
    tvUp = 5
End Enum

Public Enum tvTreeRelationship
    tvFirst = 0
    tvLast = 1
    tvNext = 2
    tvPrevious = 3
    tvChild = 4
End Enum

Event Click(cNode As clsNode)       'Node click event
' MouseAction generic example for all mouse events. See notes in NodeEventRouter
' Event MouseAction(cNode As clsNode, Action As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event NodeCheck(cNode As clsNode)   'Checkbox change event
Event AfterLabelEdit(ByRef Cancel As Boolean, NewString As String, cNode As clsNode)
Event KeyDown(cNode As clsNode, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
                    
Private Type POINTAPI
    X As Long
    Y As Long
End Type

#If Mac Then
 ' don't compile the APIs in Mac
#ElseIf VBA7 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
            ByRef lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SetCursorPos Lib "user32.dll" ( _
            ByVal X As Long, _
            ByVal Y As Long) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
#Else
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                          ByRef lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32.dll" ( _
                                          ByVal X As Long, _
                                          ByVal Y As Long) As Long
    Private Declare Function getTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
#End If

' Mac displays at 72 pixels per 72 points (1 pt / pixel) vs 96/72  or 0.75 pt / pixel in most Windows systems
' The respective constants help size and position node controls appropriatelly in for Mac 1 and Win 0.75
' However sizing may not adapt correctly in Windows OS with Large Fonts, contact for support

#If Mac Then
    Private Const mcCheckboxFont As Long = 13
    Private Const mcCheckboxPad As Single = 19
    Private Const mcCheckboxPadImg As Single = 15
    Private Const mcChkBoxSize As Single = 13
    Private Const mcExpanderFont As Long = 13
    Private Const mcExpButSize As Single = 15
    Private Const mcExpBoxSize As Long = 12
    Private Const mcFullWidth As Long = 800
    Private Const mcIconPad As Single = 17
    Private Const mcIconSize As Long = 16
    Private Const mcTLpad As Long = 4
    Private Const mcLineLeft As Single = mcTLpad + 10
    Private Const mcPtPxl As Single = 1
#Else
    Private Const mcCheckboxFont As Long = 10
    Private Const mcCheckboxPad As Single = 15
    Private Const mcCheckboxPadImg As Single = 11.25
    Private Const mcChkBoxSize As Single = 10.5
    Private Const mcExpanderFont As Long = 10
    Private Const mcExpButSize As Single = 11.25
    Private Const mcExpBoxSize As Long = 9
    Private Const mcFullWidth As Long = 600
    Private Const mcIconPad As Single = 14.25
    Private Const mcIconSize As Long = 12
    Private Const mcTLpad As Long = 3
    Private Const mcLineLeft As Single = mcTLpad + 7.5
    Private Const mcPtPxl As Single = 0.75
    Private Const SYMBOL_CHARSET As Long = 2
#End If

Private Const mcSource As String = "clsTreeView"

'***************************
'*    Public Properties    *
'***************************

Public Property Get ActiveNode() As clsNode
    Set ActiveNode = moActiveNode
End Property

Public Property Set ActiveNode(oActiveNode As clsNode)
'-------------------------------------------------------------------------
' Procedure : ActiveNode
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 17-01-2013
' Purpose   : Setting the activenode also updates the node colors
'             and ensures the node is scrolled into view
'-------------------------------------------------------------------------

    Dim cTmp As clsNode
    If oActiveNode Is MoveCopyNode(False) Then
        Set MoveCopyNode(False) = Nothing
    End If

    If moActiveNode Is oActiveNode Then
        SetActiveNodeColor
        Exit Property
    End If
    
    ResetActiveNodeColor ActiveNode

    If oActiveNode.Control Is Nothing Then
        Set cTmp = oActiveNode.ParentNode
        While Not cTmp.Caption = "RootHolder"
            cTmp.Expanded = True
            Set cTmp = cTmp.ParentNode
        Wend

        If mlNodesCreated Then
            BuildRoot False
        End If

    End If

    Set moActiveNode = oActiveNode
    SetActiveNodeColor

End Property

Public Sub ExpandNode(cNode As clsNode)
    Dim cTmp As clsNode

    Set cTmp = cNode.ParentNode
    While Not cTmp.Caption = "RootHolder"
        cTmp.Expanded = True
    Wend
    
End Sub

Public Property Get AppName() As String
    AppName = msAppName
End Property

Public Property Let AppName(ByVal sAppName As String)
    msAppName = sAppName
End Property

Public Property Get Changed() As Boolean
'PT user has edited node(s) and/or changed Checked value(s)
    Changed = mbChanged
End Property

Public Property Let Changed(ByVal bChanged As Boolean)
' called after manual node edit and Checked change
    mbChanged = bChanged
End Property

Public Property Get CheckBoxes(Optional bTriState As Boolean) As Boolean    ' PT
    CheckBoxes = mbCheckboxes
    bTriState = mbTriState
End Property

Public Property Let CheckBoxes(Optional bTriState As Boolean, ByVal bCheckboxes As Boolean)   ' PT
    Dim bOrig As Boolean
    Dim bOrigTriState As Boolean

    bOrig = mbCheckboxes
    mbCheckboxes = bCheckboxes

    bOrigTriState = mbTriState
    mbTriState = bTriState
    If bCheckboxes Then
        msngChkBoxPad = mcCheckboxPad
        If msngNodeHeight < mcExpButSize Then msngNodeHeight = mcExpButSize
    Else
        msngChkBoxPad = 0
    End If

    If Not TreeControl Is Nothing Then

        If TreeControl.Controls.Count And (bOrig <> mbCheckboxes Or bOrigTriState <> mbTriState) Then
            ' Checkboxes added changed after start-up so update the treeview
            mbRedesign = True
            Refresh
        End If
    End If

End Property

#If HostProject = "Access" Then
    Public Property Set Form(frm As Access.Form)
        Set moForm = frm
    End Property
#Else
    Public Property Set Form(frm As MSForms.UserForm)
        Set moForm = frm
    End Property
#End If

Public Property Get FullWidth() As Boolean
    FullWidth = mbFullWidth
End Property

Public Property Let FullWidth(bFullWidth As Boolean)
    mbFullWidth = bFullWidth
End Property

Public Sub ImageAdd(pic As StdPicture, sName As String) ' PT
' added v026
    On Error GoTo errH
    If mcolIcons Is Nothing Then
        Set mcolIcons = New Collection
        mbGotIcons = True
    End If
    mcolIcons.Add pic, sName ' error if sName is not a unique key
    Exit Sub
errH:
    Err.Raise Err.Number, "ImageAdd", "Error in ImageAdd" & vbNewLine & Err.Description
End Sub

Public Property Set Images(objImages As Object)
    Dim sDesc As String
    Dim pic As stdole.StdPicture
    Dim obj As Object
    ' PT  objImages can be a collection of StdPicture objects
    '     a Frame containing only Image controls (or controls with an image handle)
    '     stdole.IPictureDisp or stdole.StdPicture  objects
    
    On Error GoTo errH
    If TypeName(objImages) = "Collection" Then
        Set mcolIcons = objImages
100     For Each pic In mcolIcons
            ' if not a valid picture let the error abort
        Next
    Else
        Set mcolIcons = New Collection

        '#If HostProject = "Access" Then
            '' if the frame is on an Access form include .Object
            'For Each obj In objImages.Object.Controls

200         For Each obj In objImages.Controls
                mcolIcons.Add obj.Picture, obj.Name
            Next
    End If

    ' Flag we have a valid collection of images
    mbGotIcons = mcolIcons.Count >= 1
    
    Exit Property
errH:
    Set mcolIcons = Nothing
    If Erl = 100 Then
        sDesc = "The obImages collection includes an invalue StdPicture object"
    ElseIf Erl = 200 Then
        sDesc = "A control in objImages does not contain a valid Picture object"
    End If
    sDesc = sDesc & vbNewLine & Err.Description
        
    Err.Raise Err.Number, "Images", sDesc

End Property

Public Property Get Indentation() As Single
    Indentation = msngIndent
End Property

Public Property Let Indentation(sngIndent As Single)
    Dim cNode As clsNode
    Dim sngOld As Single

    sngOld = msngIndent

    #If Mac Then
        If sngIndent < 16 Then
            msngIndent = 16    ' min indent ?
        ElseIf sngIndent > 80 Then
            msngIndent = 80    ' max indent
        Else
            msngIndent = Int(sngIndent)
        End If
    #Else
        If sngIndent < 12 Then
            msngIndent = 12    ' min indent ?
        ElseIf sngIndent > 60 Then
            msngIndent = 60    ' max indent
        Else
            msngIndent = Int((sngIndent * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
        End If
    #End If

    If mbRootButton Then msngRootLine = msngIndent

    If Not TreeControl Is Nothing And Not (sngOld = msngIndent) Then
        ' changed after start-up so update the treeview
        If TreeControl.Controls.Count Then
            Set cNode = Me.ActiveNode
            Refresh
            If Not cNode Is Nothing Then
                Set ActiveNode = cNode
            End If
        End If
    End If
End Property
'Public Property Get EnableLabelEdit(Optional bAutoSort As Boolean) As Boolean  del
'    EnableLabelEdit = mbLabelEdit
'    bAutoSort = mbAutoSort
'End Property
'
'Public Property Let EnableLabelEdit(Optional bAutoSort As Boolean, ByVal bLabelEdit As Boolean)    ' PT
'' optional bAutoSort: automatically resort siblings after a manual edit
'    mbLabelEdit = bLabelEdit
'    mbAutoSort = bAutoSort
'End Property
Public Property Get EnableLabelEdit(Optional bAutoSort As Boolean, Optional bMultiLine As Boolean) As Boolean
    EnableLabelEdit = mbLabelEdit
    bAutoSort = mbAutoSort
    bMultiLine = mbMultiLine
End Property

Public Property Let EnableLabelEdit(Optional bAutoSort As Boolean, Optional bMultiLine As Boolean, ByVal bLabelEdit As Boolean)   ' PT
'PT optional bAutoSort: automatically sort siblings after a manual edit
'   optional bMultiLine: do not remove line breaks from manual edits
    mbLabelEdit = bLabelEdit
    mbAutoSort = bAutoSort
    mbMultiLine = bMultiLine
End Property

Public Property Get LabelEdit(Optional bAutoSort As Boolean) As Long    ' PT
' PT,  equivalent to Treeview.LabelEdit
' 0/tvwAutomatic nodes can be manually edited
' optional bAutoSort: automatically resort siblings after a manual edit

    LabelEdit = mlLabelEdit
    bAutoSort = mbAutoSort
End Property

Public Property Let LabelEdit(Optional bAutoSort As Boolean, ByVal nLabelEdit As Long)    ' PT
    mlLabelEdit = nLabelEdit
    mbLabelEdit = (nLabelEdit = 0)
    mbAutoSort = bAutoSort
End Property

Public Property Get LineColor() As Long
    'default mlLineColor=vbScrollBars applied in initialize
    LineColor = mlLineColor
End Property

Public Property Let LineColor(lColor As Long)
    If lColor <= vbInfoBackground Or (lColor >= 0 And lColor <= vbWhite) Then
        mlLineColor = lColor
    Else ' invalid input
         mlLineColor = vbScrollBars
    End If
End Property

Public Property Get MoveCopyNode(Optional bMove As Boolean, Optional lColor As Long) As clsNode
    bMove = mbMove
    Set MoveCopyNode = moMoveNode
End Property
Public Property Set MoveCopyNode(Optional bMove As Boolean, Optional lColor As Long, cNode As clsNode)
    Static lOrigBackcolor As Long

    mbMove = bMove
    If lColor = 0 Then
        If bMove Then
            lColor = RGB(255, 231, 162)
        Else: lColor = RGB(159, 249, 174)
        End If
    End If

    If Not moMoveNode Is Nothing Then
        moMoveNode.BackColor = lOrigBackcolor
        moMoveNode.Control.BackColor = lOrigBackcolor
        Set moMoveNode = Nothing
    Else

    End If

    If Not cNode Is Nothing Then
        lOrigBackcolor = cNode.BackColor
        If lOrigBackcolor = 0 Then lOrigBackcolor = mlBackColor
        cNode.BackColor = lColor

        cNode.Control.BackColor = cNode.BackColor
        cNode.Control.ForeColor = cNode.ForeColor
        Set moMoveNode = cNode
    Else

    End If
End Property

'Public Property Get MultiSelect() As Boolean
'    MultiSelect = mbMultiSelect
'End Property
'Public Property Let MultiSelect(mbMultiSelect As Boolean)
'    mbMultiSelect = MultiSelect
'End Property

Public Property Get NodeHeight() As Single
    If msngNodeHeight = 0 Then msngNodeHeight = 12
    NodeHeight = msngNodeHeight
End Property

Public Property Let NodeHeight(ByVal sngNodeHeight As Single)
    Dim cNode As clsNode
    Dim sngOld As Single

    sngOld = msngNodeHeight

    #If Mac Then
        If sngNodeHeight < 12 Then  ' height of expander-box is 12 in Mac
            msngNodeHeight = 12
        ElseIf sngNodeHeight > 60 Then
            msngNodeHeight = 60
        Else
            msngNodeHeight = Int(sngNodeHeight)
        End If
    #Else
        If sngNodeHeight < 9 Then  ' height of expander-box is 9 in Windows normal font
            msngNodeHeight = 9
        ElseIf sngNodeHeight > 45 Then
            msngNodeHeight = 45
        Else
            msngNodeHeight = Int((sngNodeHeight * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
        End If

    #End If
    If mbRootButton Then msngRootLine = msngIndent
    If Not TreeControl Is Nothing And Not (sngOld = msngNodeHeight) Then
        If TreeControl.Controls.Count Then
            Set cNode = Me.ActiveNode
            Refresh
            If Not cNode Is Nothing Then
                Set ActiveNode = cNode
            End If
        End If
    End If
End Property

Public Property Get Nodes() As Collection
' Global collection of the nodes
' *DO NOT USE* its Nodes.Add and Nodes.Remove methods
' To add & remove nodes use clsNode.AddChild() or clsTreeView.NodeAdd and clsTeevView.NodeRemove()
    If mcolNodes Is Nothing Then Set mcolNodes = New Collection
    Set Nodes = mcolNodes
End Property

Public Property Get RootButton() As Boolean
    If mbRootButton Then RootButton = 1
End Property

Public Property Let RootButton(lRootLeader As Boolean)
' PT The Root nodes have expanders and lines (if mbShowlines)

    mbRootButton = lRootLeader
    If mbRootButton Then
        msngRootLine = msngIndent
    Else
        msngRootLine = 0
    End If

    If Not Me.TreeControl Is Nothing Then
        If Not moRootHolder Is Nothing Then
            If Not moRootHolder.ChildNodes Is Nothing Then
                Refresh
            End If
        End If
    End If
End Property

Public Property Get RootNodes() As Collection
'PT returns the collection of Root-nodes
' **should be treated as read only. Use AddRoot and NodeRemove to add/remove a root node**
    Set RootNodes = moRootHolder.ChildNodes
End Property

Public Property Get ShowExpanders() As Boolean
    ShowExpanders = mbShowExpanders
End Property

Public Property Let ShowExpanders(bShowExpanders As Boolean)

    mbShowExpanders = bShowExpanders
    
    If Not TreeControl Is Nothing Then
        If TreeControl.Controls.Count Then
            Refresh
        End If
    End If
End Property

Public Property Get ShowLines() As Boolean
    ShowLines = mbShowLines
End Property

Public Property Let ShowLines(bShowLines As Boolean)
' PT Show horizontal & vertical lines
Dim bOrig As Boolean

    bOrig = mbShowLines
    mbShowLines = bShowLines

    If Not TreeControl Is Nothing Then
        If TreeControl.Controls.Count Then
            If bOrig <> mbShowLines Then
                ' ShowLines added after start-up so update the treeview
                Refresh
            End If
        End If
    End If

End Property

'***********************************
'*    Public functions and subs    *
'***********************************

Public Function AddRoot(Optional sKey As String, Optional vCaption, Optional vImageMain, _
                        Optional vImageExpanded) As clsNode

    On Error GoTo errH

    If moRootHolder Is Nothing Then
        Set moRootHolder = New clsNode
        Set moRootHolder.ChildNodes = New Collection
        Set moRootHolder.Tree = Me
        moRootHolder.Caption = "RootHolder"
        If mcolNodes Is Nothing Then
            Set mcolNodes = New Collection
        End If
    End If

    Set AddRoot = moRootHolder.AddChild(sKey, vCaption, vImageMain, vImageExpanded)

    Exit Function
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
    Err.Raise Err.Number, "AddRoot", Err.Description

End Function

Public Sub CheckboxImage(picFalse As StdPicture, picTrue As StdPicture, Optional picTriState As StdPicture)
    On Error GoTo errExit:
    Set moCheckboxImage(0) = picFalse
    Set moCheckboxImage(-1) = picTrue
    If Not IsMissing(picTriState) Then
        Set moCheckboxImage(1) = picTriState
    End If

    mbCheckboxImage = True
errExit:
End Sub

Public Sub EnterExit(bExit As Boolean)
'PT WithEvents can't trap Enter/Exit events, if we need them here they can be
'   called from the TreeControl's Enter/Exit events in the form
    mbInActive = bExit
    SetActiveNodeColor bExit ' apply appropriate vbInactiveCaptionText / vbHighlight

End Sub

Public Sub ExpanderImage(picMinus As StdPicture, picPlus As StdPicture)
    On Error GoTo errExit:
    Set moExpanderImage(0) = picPlus
    Set moExpanderImage(-1) = picMinus
    mbExpanderImage = True
errExit:
End Sub

Public Sub ExpandToLevel(lExpansionLevel As Long, Optional bReActivate As Boolean = True)
' PT call SetTreeExpansionLevel and reactivates the closest expanded parent if necessary
'    eg, if activeNode.level = 4 and lExpansionLevel = 2, the activenode's grandparent will be activated
    Dim cTmp As clsNode

    Call SetTreeExpansionLevel(lExpansionLevel - 1)

    If bReActivate Then
        If ActiveNode.Level > lExpansionLevel Then
            Set cTmp = ActiveNode.ParentNode
            While cTmp.Level > lExpansionLevel
                Set cTmp = cTmp.ParentNode
            Wend
            Set ActiveNode = cTmp
        End If
    End If
    
End Sub

Public Sub Copy(cSource As clsNode, cDest As clsNode, _
                Optional vBefore, Optional ByVal vAfter, _
                Optional ByVal bShowError As Boolean)
    Dim lParentChecked As Long
    Set MoveCopyNode(False) = Nothing
    If mbTriState Then lParentChecked = cDest.Checked
    Clone cDest, cSource, vBefore, vAfter
    If lParentChecked = -1 And cDest.Checked = 0 Then
        
        cDest.Checked = -1
    End If
    SetActiveNodeColor
    
End Sub

Public Sub Move(cSource As clsNode, cDest As clsNode, _
                Optional vBefore, Optional ByVal vAfter, _
                Optional ByVal bShowError As Boolean)
' PT Move source node + children to destination node
'    cannot move the Root and cannot move to a descendant
'   vBefore/vAfter work as for normal collection; error if invalid, eg a new collection, after the last item, etc
'
Dim sErrDesc As String
Dim bIsParent As Boolean
Dim bRemoveReverseOrder As Boolean
Dim lParentChecked As Long
Dim idx As Long
Dim cNode As clsNode
Dim cSourceParent As clsNode

    Set MoveCopyNode(False) = Nothing
    On Error GoTo errH

    If cSource Is Nothing Or cDest Is Nothing Or cSource Is cDest Then   ' Or cSource Is Root
        Exit Sub
    End If

    Set cNode = cDest
    bIsParent = False
    Do
        Set cNode = cNode.ParentNode
        bIsParent = cNode Is cSource
    Loop Until cNode Is Nothing Or bIsParent

    If bIsParent Then
        Err.Raise vbObjectError + 110
    End If

    If cDest.ChildNodes Is Nothing Then
        ' the child becomes a parent for the first time
        Set cDest.ChildNodes = New Collection
        ' expander & VLine will get created automatically if necessary
    End If

    If cDest Is cSource.ParentNode Then
        ' changing order in the original parent
        For Each cNode In cDest.ChildNodes
            idx = idx + 1
            If cNode Is cSource Then
                Exit For
            End If
        Next
        If Not IsMissing(vBefore) Then
            bRemoveReverseOrder = Val(vBefore) < idx
        ElseIf Not IsMissing(vAfter) Then
            bRemoveReverseOrder = Val(vAfter) < idx
        ElseIf cSource Is cSource.LastSibling Then
            Exit Sub ' no point to move to the same place
        End If
    End If

    AddNodeToCol cDest.ChildNodes, cSource, False, vBefore, vAfter

    Set cSourceParent = cSource.ParentNode
    
    With cSourceParent
        .RemoveChild cSource, bRemoveReverseOrder
        ' if the old parent has no more children remove its expander & VLine

        If .ChildNodes Is Nothing Then

            If Not .Expander Is Nothing Then
                Me.TreeControl.Controls.Remove .Expander.Name
                Set .Expander = Nothing
            End If

            If Not .ExpanderBox Is Nothing Then
                Me.TreeControl.Controls.Remove .ExpanderBox.Name
                Set .ExpanderBox = Nothing
            End If

            If Not .VLine Is Nothing Then
                Me.TreeControl.Controls.Remove .VLine.Name
                Set .VLine = Nothing
            End If

            .Expanded = False

        End If
    End With

    Set cSource.ParentNode = cDest
    cDest.Expanded = True

    If mbTriState Then
        lParentChecked = cDest.Checked
        cDest.CheckTriStateParent
        cSourceParent.CheckTriStateParent
        If lParentChecked = -1 And cDest.Checked = 0 Then
            cDest.Checked = -1
        End If
    End If

    SetActiveNodeColor

    mbAlwaysRedesign = True    ' ensure Left's get recalc'd during next refresh

    Exit Sub
errH:

    Select Case Err.Number
    Case vbObjectError + 110
        sErrDesc = "Cannot cut and move a Node to a descendant node"
    Case Else
        sErrDesc = "Move: " & Err.Description
    End Select

    If bShowError Then
        MsgBox sErrDesc, , AppName
    Else
        Err.Raise Err.Number, mcSource, "Move: " & sErrDesc
    End If

End Sub

Public Function NodeAdd(Optional vRelative, _
                        Optional vRelationship, _
                        Optional sKey As String, _
                        Optional vCaption, _
                        Optional vImageMain, _
                        Optional vImageExpanded) As clsNode    '  As tvTreevRelationship

'PT, similar to the ocx treeview's nodes.add method
'    main difference is vRelative can be a Node object as well as a key or index
'    see also clsNode.AddChild

    Dim i As Long
    Dim cNode As clsNode
    Dim cRelative As clsNode
    Dim cParent As clsNode
    Dim cTmp As clsNode
    '    tvFirst = 0  tvlast = 1 tvNext = 2 tvprevious = 3  tvChild = 4

    If IsMissing(vRelative) Then

        Set NodeAdd = Me.AddRoot(sKey, vCaption, vImageMain, vImageExpanded)
        Exit Function
    Else

        On Error Resume Next
        Set cRelative = vRelative
        If cRelative Is Nothing Then
            Set cRelative = mcolNodes(vRelative)
        End If

        On Error GoTo errH
        If cRelative Is Nothing Then
            Err.Raise vbObjectError + 100, "NodeAdd", "vRelative is not a valid node or a node.key"
        End If
    End If

    If IsMissing(vRelationship) Then
        vRelationship = tvTreeRelationship.tvNext    ' default
    End If

    If vRelationship = tvChild Or cRelative Is cRelative.Root Then
        Set cParent = cRelative
    Else
        Set cParent = cRelative.ParentNode
    End If

    Set cNode = New clsNode

    If Len(sKey) Then
100     mcolNodes.Add cNode, sKey
101
    Else
        mcolNodes.Add cNode
    End If

    If cParent.ChildNodes Is Nothing Then
        Set cParent.ChildNodes = New Collection
    End If

    With cParent.ChildNodes
        If .Count = 0 Then
            .Add cNode
        Else
            i = 0
            If vRelationship = tvNext Or vRelationship = tvPrevious Then
                For Each cTmp In cParent.ChildNodes
                    i = i + 1
                    If cTmp Is cRelative Then
                        Exit For
                    End If
                Next
            End If
            Select Case vRelationship
            Case tvFirst: .Add cNode, , 1
            Case tvLast: .Add cNode, after:=.Count
            Case tvNext: .Add cNode, after:=i
            Case tvPrevious: .Add cNode, before:=i
            Case tvChild: .Add cNode
            End Select
        End If
    End With

    With cNode
        .Key = sKey
        .Caption = CStr(vCaption)
        .ImageMain = vImageMain
        .ImageExpanded = vImageExpanded
        .Index = mcolNodes.Count
        
        Set .ParentNode = cParent
        Set .Tree = Me
    End With

    Set cNode.Tree = Me    ' do this after let key = skey
    Set NodeAdd = cNode

    Exit Function
errH:
    If mcolNodes Is Nothing Then
        Set mcolNodes = New Collection
        Resume
    End If
    If Erl = 100 And Err.Number = 457 Then
        Err.Raise vbObjectError + 1, "clsNode.AddChild", "Duplicate key: '" & sKey & "'"
    Else
        #If DebugMode = 1 Then
            Stop
            Resume
        #End If
        Err.Raise Err.Number, "clsNode.AddChild", Err.Description
    End If
End Function

Public Sub NodeRemove(cNode As clsNode)
' PT Remove a Node, its children and grandchildrem
'    remove all associated controls and tear down class objects
'    Call Refresh() when done removing nodes

    Dim lIdx As Long
    Dim lNodeCtlsOrig As Long
    Dim cParent As clsNode
    Dim cNodeAbove As clsNode, cNd As clsNode
    
    On Error GoTo errH

    Set cNodeAbove = NextVisibleNodeInTree(cNode, bUp:=True)
    Set cParent = cNode.ParentNode

    cNode.TerminateNode True

    cParent.RemoveChild cNode
    
    cNode.Index = -1    ' flag to get removed from mcolNodes in the loop below
    If ActiveNode Is cNode Then
        Set moActiveNode = Nothing
    End If
    Set moEditNode = Nothing

    lIdx = 0
    lNodeCtlsOrig = mlNodesCreated
    mlNodesCreated = 0
    
    For Each cNd In mcolNodes
        lIdx = lIdx + 1
        If cNd.Index = -1 Then
            mcolNodes.Remove lIdx
            lIdx = lIdx - 1 ' decrement the collection index
        Else
            mlNodesCreated = mlNodesCreated - CLng(Not cNd.Control Is Nothing)
            cNd.Index = lIdx
        End If
    Next

    mlNodesDeleted = mlNodesDeleted + lNodeCtlsOrig - mlNodesCreated

    Set cNode = Nothing    ' should terminate the class

    If mlNodesCreated Then
        If Not cNodeAbove Is Nothing Then
            Set Me.ActiveNode = cNodeAbove
        ElseIf mcolNodes.Count Then
            Set Me.ActiveNode = mcolNodes(1)
        End If
    Else
        'all nodes deleted
        Erase mlVisOrder
        Erase msngMaxWidths
        mlVisCount = 0
        mlNodesCreated = 0
        mlNodesDeleted = 0
    End If

    Exit Sub
errH:
    #If DebugMode = 1 Then
        Debug.Print Err.Description, Err.Number
        Stop
        Resume
    #End If
End Sub

Public Sub NodesClear()
' PT,  similar to Treeview.Nodes.Clear
    Dim i As Long
    On Error GoTo errH

    If Not TreeControl Is Nothing Then
        With TreeControl
            For i = TreeControl.Controls.Count - 1 To 0 Step -1
                TreeControl.Controls.Remove i
            Next
            .ScrollBars = fmScrollBarsNone
            .ScrollWidth = 0
            .ScrollHeight = 0
            .ScrollTop = 0
        End With
    End If
    
    Erase mlVisOrder
    Erase msngMaxWidths
    mlVisCount = 0
    mlNodesCreated = 0
    mlNodesDeleted = 0
    
    TerminateTree

    mbChanged = False

    Exit Sub
errH:
    #If DebugMode = 1 Then
        Stop
        Resume
    #End If
End Sub

Public Sub Refresh()
' Create node controls as required the first time respective parent's Expanded property = true
' hide or show and (re)position node controls as required
' Call Refresh after changing any Treeview properties or after adding/removing/moving any nodes
' or making any change that will alter placement of nodes in the treeview
    Dim bInit As Boolean

    If Me.TreeControl Is Nothing Then
        TerminateTree
        ' a Frame (container for the treeview) should have been referrenced to me.TreeControl
        Err.Raise vbObjectError + 10, mcSource, "Refresh: 'TreeControl' frame is not referenced"
        
    ElseIf moRootHolder Is Nothing Then
        '
        Err.Raise vbObjectError + 11, mcSource, "Refresh: No Root nodes have been created"
        
    ElseIf moRootHolder.ChildNodes Is Nothing Then
        ' nothing to do
        mlVisCount = 0
        mlNodesCreated = 0
        mlNodesDeleted = 0
        Erase mlVisOrder
        Erase msngMaxWidths
        Exit Sub
        
    ElseIf Me.TreeControl.Controls.Count = 0 Then
        ' display the treeview for first time
        bInit = True
    Else
        ' ensure all node properties are checked, eg after changing indentation or nodeheight during runtime
        mbRedesign = True
    End If

    On Error GoTo errExit

    BuildRoot bInit

    Exit Sub

errExit:
    HourGlassCursor False
    Err.Raise Err.Number, mcSource, "Error in BuildRoot: " & Err.Description
End Sub

Public Sub ScrollToView(Optional cNode As clsNode, _
                        Optional Top1Bottom2 As Long, _
                        Optional bCollapseOthers As Boolean)
' PT scrolls the treeview to position the node in view
' Top1Bottom2= 0 roughly 1/3 from the top
' Top1Bottom2= 1 or -1 at the top
' Top1Bottom2= 2 or -2 at the bottom

    Dim bIsVisible As Boolean
    Dim bWasCollapsed As Boolean
    Dim lVisIndex As Long
    Dim sngTop As Single
    Dim sngBot As Single
    Dim sngVisHt As Single
    Dim sngScrollTop As Single
    Dim cTmp As clsNode

    If cNode Is Nothing Then
        Set cNode = ActiveNode
    End If

    If bCollapseOthers Then
        SetTreeExpansionLevel 0
    End If
    
    Set cTmp = cNode.ParentNode
    While Not cTmp.Caption = "RootHolder"
        If Not cTmp.Expanded Then
            bWasCollapsed = True
            cTmp.Expanded = True
        End If
        Set cTmp = cTmp.ParentNode
    Wend
    
    If bWasCollapsed Then
        BuildRoot False
    End If

    lVisIndex = cNode.VisIndex
    sngBot = mcTLpad + lVisIndex * NodeHeight
    sngTop = sngBot - NodeHeight

    With TreeControl
        sngVisHt = .InsideHeight
        If .ScrollBars = fmScrollBarsBoth Or .ScrollBars = fmScrollBarsHorizontal Then
            sngVisHt = sngVisHt - 15    ' roughly(?) width of a scrollbar
        End If

        bIsVisible = sngTop > .ScrollTop And _
                     sngBot < .ScrollTop + sngVisHt
        
        If Not bIsVisible Or Top1Bottom2 > 0 Then
        
            If Top1Bottom2 < 0 Then Top1Bottom2 = Top1Bottom2 * -1
            
            If Top1Bottom2 = 0 Then  ' place about 1/3 from top
                sngScrollTop = lVisIndex * NodeHeight - .InsideHeight / 3

            ElseIf Top1Bottom2 = 1 Then  ' scroll to top
                sngScrollTop = sngTop - mcTLpad
            Else
                sngScrollTop = sngBot - sngVisHt + mcTLpad    ' scroll to bottom
            End If

            If sngScrollTop < 0 Then
                sngScrollTop = 0
            End If

            .ScrollTop = sngScrollTop
        End If
    End With
End Sub

Public Sub SetScrollbars(bRecalcWidths As Boolean, Optional bSetScrollDims As Boolean)
' PT set scrollbars as required to ensure all nodes are viewable
'    called internally by BuildRoot
'    call externally to reset scrollbars while resizing the Treeview container Frame with bSetScrollDims:=False
'    or with bSetScrollDims:=True after changing or updating text in node captions
Dim i As Long
Dim bars As Long

    If bRecalcWidths Then
        RecalcMaxTextWidth
        bSetScrollDims = True
    End If

    With Me.TreeControl
    
        If bSetScrollDims Then
            .ScrollWidth = MaxNodeWidth + mcTLpad
            .ScrollHeight = mlVisCount * NodeHeight + mcTLpad
        End If
        
        For i = 0 To 1
            bars = 0
            If .InsideWidth < .ScrollWidth Then
                bars = fmScrollBarsHorizontal
            End If

            If .InsideHeight < .ScrollHeight Then
                bars = bars + fmScrollBarsVertical
            End If

            If bars <> .ScrollBars Then
                .ScrollBars = bars
            Else
                Exit For
            End If
        Next

    End With
    
End Sub

Public Sub TerminateTree()
'-------------------------------------------------------------------------
' Procedure : TerminateTree
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 15-01-2013
' Purpose   : Terminates this class' instance
'-------------------------------------------------------------------------
Dim cNode As clsNode
    'Instead of the terminate event of the class
    'we use this public method so it can be
    'explicitly called by parent classes
    'this is done because we'll end up having multiple circular references
    'between parent and child classes, which may cause the terminate events to be ignored.

    If Not moRootHolder Is Nothing Then
        If Not moRootHolder.ChildNodes Is Nothing Then
            For Each cNode In moRootHolder.ChildNodes

                cNode.TerminateNode
            Next
        End If
        moRootHolder.TerminateNode
    End If
    
    Set moMoveNode = Nothing
    Set moEditNode = Nothing
    Set moActiveNode = Nothing
    Set moRootHolder = Nothing
    Set mcolNodes = Nothing
    
    '** TerminateTree does NOT reset treeview properties or remove
    '** the TreeControl reference to the treeview's Frame control
    '
    '   If the form is being unloaded it's enough to call TerminateTree in it's close event,
    '   node controls will automatically unload with the form.
    '   However if the treeview is to be cleared or moved but the main form is not being unloaded
    '   call the NodesRemove method which will remove all node controls, then call TerminateTree
End Sub

'***********************************************************************************************
'*    Friend properties, functions and subs                                                    *
'*    although visible throughout the project these are only intended to be called by clsNodes *
'***********************************************************************************************

Friend Property Get EditMode(cNode As clsNode) As Boolean  ' PT
    EditMode = mbEditMode
End Property

Friend Property Let EditMode(cNode As clsNode, ByVal bEditMode As Boolean)  ' PT

    Set MoveCopyNode(False) = Nothing
    mbEditMode = bEditMode

    If Not moEditNode Is Nothing Then
        moEditNode.EditBox False
    End If


    If bEditMode Then
        Set moEditNode = cNode
    Else
        Set moEditNode = Nothing
    End If
End Property

Friend Function GetExpanderIcon(bExpanded As Boolean, pic As StdPicture) As Boolean
    If mbExpanderImage Then
        Set pic = moExpanderImage(bExpanded)
        GetExpanderIcon = True
    End If
End Function
Friend Function GetCheckboxIcon(lChecked As Long, pic As StdPicture) As Boolean
    If mbCheckboxImage Then
        Set pic = moCheckboxImage(lChecked)
        GetCheckboxIcon = True
    End If
End Function

Friend Function GetNodeIcon(vKey, pic As StdPicture, bFullWidth As Boolean) As Boolean
    On Error GoTo errExit
    Set pic = mcolIcons(vKey)
    GetNodeIcon = True
errExit:
    bFullWidth = mbFullWidth   ' after error in case no picture
End Function

Friend Function RaiseAfterLabelEdit(cNode As clsNode, sNewText As String) As Boolean
' PT called from moEditBox_KeyDown after vbKeyEnter
'
    Dim Cancel As Boolean
    RaiseEvent AfterLabelEdit(Cancel, sNewText, cNode)
    RaiseAfterLabelEdit = Cancel
End Function

Friend Sub NodeEventRouter(cNode As clsNode, sControl As String, lAction As tvMouse, _
                           Optional Button As Integer, Optional Shift As Integer, Optional X As Single, Optional Y As Single)
Dim bFlag As Boolean
Dim lngViewable As Long
Dim cLastChild As clsNode

    If sControl = "Caption" Then

        If Not ActiveNode Is cNode Then
            Set ActiveNode = cNode
        Else
            SetActiveNodeColor
        End If

        If lAction = 1 Then  ' tvClick
            RaiseEvent Click(cNode)
            
        Else ' tvDblClick, tvDown, tvMove, tvUp
        
            ' To enable DblClick, MouseDown and MouseUp events:
            '   uncomment moTree.NodeEventRouter in mctlControl_MouseDown, _MouseUp and _DblClick in clsNode
            '   add a mctlControl_MouseMove stub if MouseMove is required and adapt the call to NodeEventRouter with lAction = tvMove
            '   uncomment Event MouseAction in the declarations in this module
            '   uncomment RaiseEvent MouseAction below
            '   add the treeview_MouseAction stub in the main from that holds the reference to the treeview
            
            ' RaiseEvent MouseAction(cNode, lAction, Button, Shift, X, Y)
        End If

    ElseIf sControl = "Expander" Then

        bFlag = Not ActiveNode Is cNode

        If bFlag Then
            Set ActiveNode = cNode
        End If

        BuildRoot False

        If cNode.Expanded Then
            If Not cNode.ChildNodes Is Nothing Then
                Set cLastChild = cNode.ChildNodes(cNode.ChildNodes.Count)
                If Not NodeIsVisible(cLastChild, lngViewable) Then
                    If lngViewable > cNode.ChildNodes.Count Then
                        ScrollToView cLastChild, Top1Bottom2:=2
                    Else
                        ScrollToView cNode, Top1Bottom2:=1
                    End If
                End If
            End If
        End If

        If bFlag Then
            ' Skip unnecesary click event if user clicks the already active node,
            ' only raise the click event if a new Node has just been made active.
            ' Remove this If bFlag if always want to trap the click event
            RaiseEvent Click(cNode)
        End If

    ElseIf sControl = "Checkbox" Then

        RaiseEvent NodeCheck(cNode)

    End If

End Sub

Friend Function UniqueKey(sKey As String) As String
    Dim cNode As clsNode
    For Each cNode In Nodes
        If cNode.Key = sKey Then
            Err.Raise vbObjectError + 1, "clsTreeView", "Duplicate key: '" & sKey & "'"
        End If
    Next
    UniqueKey = sKey
End Function


'*********************************************************************************************
'*    Private events    *
'**********************************************************************************************

Private Sub TreeControl_Click()
' PT exit editmode if an empty part of the treeview is clicked
    EditMode(ActiveNode) = False
End Sub

'************************************
'*    Private functions and subs    *
'************************************

Private Sub Class_Initialize()

' default properties

    mbRootButton = True
    mbShowExpanders = True
    mbShowLines = True
    mlLineColor = vbScrollBars
    
    #If Mac Then
        msngIndent = 20
        msngNodeHeight = 16
    #Else
        msngIndent = 15
        msngNodeHeight = 12
    #End If
    
    msngRootLine = msngIndent
    msAppName = "TreeView"
    
    #If DebugMode = 1 Then
        gClsTreeViewInit = gClsTreeViewInit + 1    'for testing only
    #End If
    
End Sub

Private Sub Class_Terminate()
    #If DebugMode = 1 Then
        gClsTreeViewTerm = gClsTreeViewTerm + 1
    #End If
End Sub

Private Function AddNodeToCol(colNodes As Collection, cAddNode As clsNode, bTreeCol As Boolean, Optional vBefore, Optional vAfter)
    Dim i As Long
    Dim sKey As String
    Dim cTmp As clsNode
    Dim pos As Long

    If bTreeCol Then sKey = cAddNode.Key

    If Len(sKey) Then
        On Error Resume Next
        i = 0
        Set cTmp = colNodes(sKey)
        If Not cTmp Is Nothing Then
            pos = InStr(1, sKey, "_copy:")
            If pos Then
                sKey = Left$(sKey, pos - 1)
            End If
            sKey = sKey & "_copy:"
            While Not cTmp Is Nothing
                Set cTmp = Nothing
                i = i + 1
                Set cTmp = colNodes(sKey & i)
            Wend
            sKey = sKey & i

            If bTreeCol Then
                cAddNode.Key = sKey
            End If

        End If

        On Error GoTo 0    ' error returns to caller

        If IsMissing(vBefore) And IsMissing(vAfter) Then
            colNodes.Add cAddNode, sKey
        ElseIf IsMissing(vAfter) Then
            colNodes.Add cAddNode, sKey, Val(vBefore)
        Else
            colNodes.Add cAddNode, sKey, , Val(vAfter)
        End If

    Else    ' no key
        If IsMissing(vBefore) And IsMissing(vAfter) Then
            colNodes.Add cAddNode
        ElseIf IsMissing(vAfter) Then
            colNodes.Add cAddNode, , Val(vBefore)
        Else
            colNodes.Add cAddNode, , , Val(vAfter)
        End If
    End If
    
End Function

Private Sub BuildRoot(bInit As Boolean)
Dim bCursorWait As Boolean
Dim bTriStateOrig As Boolean
Dim i As Long
Dim bars As Long
Dim lLastRootVisIndex As Long
Dim sngActiveNodeScrollTop As Single    ' PT distance activenode was from scrolltop top before refresh, if visible
Dim sngChkBoxPad As Single
Dim sngHeightAllNodes As Single
Dim sngIconPad As Single
Dim cRoot As clsNode
Dim objCtrl As MSForms.Control

Dim vIconKey

Dim sCap As String
Dim sngTickCnt As Single

    On Error GoTo locErr

    #If DebugMode Then
        #If Mac Then
            sngTickCnt = Timer
        #Else  ' Windows
            sngTickCnt = getTickCount
        #End If
    #End If

    bInit = TreeControl.Count = 0

    'TODO find equivalent for cancel key in Access & Word
    #If HostProject = "Access" Then
    #ElseIf HostProject = "Word" Then
    #Else
        Application.EnableCancelKey = xlErrorHandler
    #End If

    If mbAlwaysRedesign Then mbRedesign = True

    '    mcChkBoxSize = 10.5    ' 11.25
    '    mcLineLeft = 3 + 7.5    'msngIndent / 2

    ' PT if these arrays aren't large enough Redim Preserve is done in error handler
    ReDim mlVisOrder(1 To mlNodesCreated + 100)
    If bInit Or mbRedesign Then
        ReDim msngMaxWidths(0 To 7)
    End If

    If mcolNodes.Count - mlNodesCreated > 400 Then
        ' creating many controls might take a while
        HourGlassCursor bApply:=True
        bCursorWait = True
    End If
    
    If Not bInit Then
        If NodeIsVisible Then
            sngActiveNodeScrollTop = (ActiveNode.VisIndex - 1) * NodeHeight - Me.TreeControl.ScrollTop
        End If
    End If

    mlVisCount = 0
    bTriStateOrig = mbTriState
    mbTriState = False

    If CheckBoxes Then
        If mbCheckboxImage Then
            sngChkBoxPad = mcCheckboxPadImg
        Else
            sngChkBoxPad = mcCheckboxPad
        End If
        If mcChkBoxSize > msngNodeHeight Then
            msngNodeHeight = mcChkBoxSize
        End If
    End If

    ' work out respective offsets to various node controls from node tops
    msngTopExpB = mcTLpad + (msngNodeHeight - mcExpButSize) / 2 + 1.5
    If mbExpanderImage Then
        msngTopExpT = mcTLpad + (msngNodeHeight - (mcExpButSize - 4)) / 2
    Else
        msngTopExpT = mcTLpad + (msngNodeHeight - mcExpButSize) / 2
    End If

    msngTopChk = mcTLpad + (msngNodeHeight - mcChkBoxSize) / 2
    msngTopIcon = mcTLpad + (msngNodeHeight - mcIconSize) / 2
    msngTopHV = mcTLpad + msngNodeHeight / 2
    Call Round75


    With TreeControl
        mlBackColor = .BackColor    ' default colours for node labels
        mlForeColor = .ForeColor

        If bInit Then
            
        Else
            ' PT, refresh, start by hiding all the controls
            For Each objCtrl In .Controls
                objCtrl.Visible = False
            Next
        End If

        For Each cRoot In moRootHolder.ChildNodes
            sngIconPad = 0
            If mbFullWidth Then
                If mbGotIcons And cRoot.hasIcon(vIconKey) Then
                    sngIconPad = mcIconPad
                End If
            End If

            If cRoot.Control Is Nothing Then
                mlNodesCreated = mlNodesCreated + 1
                'Add the rootnode to the tree
                Set cRoot.Control = TreeControl.Controls.Add("Forms.label.1", "Node" & mlNodesDeleted + mlNodesCreated, False)
                With cRoot.Control

                    If Not mbFullWidth And mbGotIcons Then
                        If cRoot.hasIcon(vIconKey) Then
                            .PicturePosition = fmPicturePositionLeftCenter
                            .Picture = mcolIcons(vIconKey)
                        End If
                    End If

                    .Top = mcTLpad + mlVisCount * msngNodeHeight
                    .Left = mcTLpad + msngRootLine + sngIconPad + msngChkBoxPad

                    If cRoot.BackColor Then
                        .BackColor = cRoot.BackColor
                    End If
                    If cRoot.ForeColor Then
                        .ForeColor = cRoot.ForeColor
                    End If

                    If cRoot.Bold Then .Font.Bold = True
                    .Caption = cRoot.Caption
                    .AutoSize = True
                    .WordWrap = False

                    cRoot.TextWidth = .Width

                    If .Width + sngIconPad > msngMaxWidths(0) Then
                        msngMaxWidths(0) = .Width + sngIconPad
                    End If

                    ' calc msngTopLabel to align node label to mid NodeHeight
                    ' first calc min NodeHeight if not set higher by user
                    If .Height > msngNodeHeight Then
                        ' optimal HodeHeight for the current font
                        msngNodeHeight = .Height    ' 'don't use the Property method or Refresh will be called
                    ElseIf .Height < msngNodeHeight Then
                        #If Mac Then
                            msngTopLabel = Int(msngNodeHeight - .Height) / 2
                        #Else
                            msngTopLabel = Int((msngNodeHeight - .Height + mcPtPxl) / 3 * 2) * mcPtPxl
                        #End If
                        .Top = mcTLpad + msngTopLabel + mlVisCount * msngNodeHeight
                    End If

                    If mbFullWidth Then
                        If msngTopLabel < mcFullWidth Then
                            .Width = mcFullWidth
                            .AutoSize = False
                        End If
                    End If

                    If Len(cRoot.ControlTipText) Then
                        .ControlTipText = cRoot.ControlTipText
                    End If

                    .WordWrap = False
                    .ZOrder 0
                    .Visible = True

                End With
            Else

                With cRoot.Control

                    If mbRedesign Then
                        .Left = mcTLpad + msngRootLine + sngIconPad + msngChkBoxPad

                        If cRoot.TextWidth + sngIconPad > msngMaxWidths(0) Then
                            msngMaxWidths(0) = cRoot.TextWidth + sngIconPad
                        End If
                    End If

                    If .Height > msngNodeHeight Then
                        msngNodeHeight = .Height
                    ElseIf .Height < msngNodeHeight Then
                        #If Mac Then
                            msngTopLabel = Int(msngNodeHeight - .Height) / 2
                        #Else
                            msngTopLabel = Int((msngNodeHeight - .Height + mcPtPxl) / 3 * 2) * mcPtPxl
                        #End If
                    End If

                    .Top = mcTLpad + msngTopLabel + mlVisCount * msngNodeHeight

                    .Visible = True

                End With
            End If

            ' horizontal line
            If mbRootButton And mbShowLines Then
                If cRoot.HLine Is Nothing Then
                    Set cRoot.HLine = TreeControl.Controls.Add("Forms.label.1", "HLine" & cRoot.Control.Name, False)
                    With cRoot.HLine
                        .Top = msngTopHV + mlVisCount * msngNodeHeight
                        .Left = mcLineLeft
                        .Caption = ""
                        .BorderStyle = fmBorderStyleSingle
                        .BorderColor = LineColor    'vbScrollBars
                        .Width = msngIndent
                        .Height = mcPtPxl
                        .TextAlign = fmTextAlignCenter
                        .BackStyle = fmBackStyleTransparent
                        .ZOrder 1
                        .Visible = True
                    End With
                Else
                    With cRoot.HLine
                        .Width = msngIndent
                        .Top = msngTopHV + mlVisCount * msngNodeHeight  ' 3 + NodeHeight/2 (to nearest 0.75)
                        .Visible = True
                    End With
                End If
            End If

            ' Checkbox
            If CheckBoxes Then
                If cRoot.Checkbox Is Nothing Then
                    Set cRoot.Checkbox = TreeControl.Controls.Add("Forms.label.1", "CheckBox" & cRoot.Control.Name, False)
                    With cRoot.Checkbox
                        .Left = mcTLpad + msngRootLine
                        .Top = msngTopChk + mlVisCount * msngNodeHeight

                        If mbCheckboxImage Then
                            'Use an image
                            .BorderStyle = fmBorderStyleNone
                            .Picture = moCheckboxImage(cRoot.Checked)
                            .PicturePosition = fmPicturePositionLeftTop
                            .AutoSize = True

                        Else
                            .Width = mcChkBoxSize
                            .Height = mcChkBoxSize
                            .Font.Name = "Marlett"  ' "a" is a tick
                            #If Mac = False Then
                                .Font.Charset = SYMBOL_CHARSET
                            #End If
                            .FontSize = mcCheckboxFont     '9
                            .BorderStyle = fmBorderStyleSingle
                            .BackColor = vbWindowBackground
                            .ForeColor = vbWindowText
                            If cRoot.Checked Then
                                .Caption = "a"
                                If cRoot.Checked = 1 Then
                                    .ForeColor = RGB(180, 180, 180)
                                End If
                            End If
                        End If

                        .Visible = True
                    End With
                Else
                    With cRoot.Checkbox
                        .Left = mcTLpad + msngRootLine
                        .Top = msngTopChk + mlVisCount * msngNodeHeight
                        .Visible = True
                    End With
                End If
            End If

            ' Icon
            If mbFullWidth And mbGotIcons Then
                If cRoot.hasIcon(vIconKey) Then
                    If cRoot.Icon Is Nothing Then
                        Set cRoot.Icon = TreeControl.Controls.Add("Forms.Image.1", "Icon" & cRoot.Control.Name, False)
                        With cRoot.Icon
                            .BackStyle = fmBackStyleTransparent
                            .BorderStyle = fmBorderStyleNone
                            '.AutoSize
                            .Width = mcIconSize
                            .Height = mcIconSize
                            .Left = mcTLpad + msngRootLine + msngChkBoxPad
                            .Top = msngTopIcon + mlVisCount * msngNodeHeight
                            .Picture = mcolIcons(vIconKey)
                            .BackStyle = fmBackStyleTransparent
                            .Visible = True
                        End With
                    Else
                        With cRoot.Icon
                            .Left = mcTLpad + msngRootLine + msngChkBoxPad
                            .Top = msngTopIcon + mlVisCount * msngNodeHeight
                            .Visible = True
                        End With
                    End If
                Else
                    sngIconPad = 0
                End If
            End If

            mlVisCount = mlVisCount + 1
            mlVisOrder(mlVisCount) = cRoot.Index
            cRoot.VisIndex = mlVisCount

            lLastRootVisIndex = mlVisCount

            'Now add this root's children
            If Not cRoot.ChildNodes Is Nothing Then
                BuildTree cRoot, 1, True
            End If

        Next

        'Vertical line for multiple roots
        If mbRootButton And mbShowLines Then
            If moRootHolder.ChildNodes.Count > 1 Then

                If moRootHolder.VLine Is Nothing Then
                    Set moRootHolder.VLine = TreeControl.Controls.Add("forms.label.1", "VLine_Roots")
                    With moRootHolder.VLine
                        .ZOrder 1
                        .Width = mcPtPxl
                        .Caption = ""
                        .BorderColor = mlLineColor    ' vbScrollBars
                        .BorderStyle = fmBorderStyleSingle
                        .Top = msngTopHV
                        .Left = mcLineLeft
                        .Height = (lLastRootVisIndex - 1) * msngNodeHeight
                    End With

                Else

                    With moRootHolder.VLine
                        .Top = msngTopHV
                        .Height = (lLastRootVisIndex - 1) * msngNodeHeight
                        .Visible = True
                    End With
                End If

            End If
        End If

        sngHeightAllNodes = mlVisCount * NodeHeight + mcTLpad     ' mcTLpad for top/bottom padding
        If bInit Then
            .ScrollHeight = 0
            .ScrollLeft = 0
        End If

        SetScrollbars False, True

        If bInit Then    ' startup
            '' make the first root node active but don't highlight it
            Set moActiveNode = moRootHolder.ChildNodes(1)
            '' or if preferred highlighted at startup
            'Set ActiveNode = moRootHolder.ChildNodes(1)
        ElseIf Not ActiveNode Is Nothing Then
            If Not NodeIsVisible Then
                .ScrollTop = (ActiveNode.VisIndex - 1) * NodeHeight - sngActiveNodeScrollTop
            End If
        End If

    End With

    #If DebugMode Then
        On Error Resume Next
        #If Mac Then
            sngTickCnt = Timer - sngTickCnt
        #Else
            sngTickCnt = (getTickCount - sngTickCnt) / 1000
        #End If

        sCap = "Seconds: " & Format(sngTickCnt, "0.00") & _
               "    Nodes: " & mcolNodes.Count & _
               "  created: " & mlNodesCreated & _
               "  visible: " & mlVisCount & _
               "    Total controls: " & TreeControl.Controls.Count

        #If HostProject = "Access" Then
            If Not moForm Is Nothing Then
                moForm.Caption = sCap
            End If
        #Else
            Me.TreeControl.Parent.Caption = sCap
        #End If
        On Error GoTo locErr
    #End If

    mbRedesign = False
    mbTriState = bTriStateOrig

done:

    If bCursorWait Then
        HourGlassCursor False
    End If


    'TODO: implement API equivalent for cancel key in Access & Word ?
    #If HostProject = "Access" Then
    #ElseIf HostProject = "Word" Then
    #Else
        Application.EnableCancelKey = xlInterrupt
    #End If

    Exit Sub

locErr:
    If bCursorWait Then HourGlassCursor False
    mbRedesign = False
    mbTriState = bTriStateOrig

    If Err.Number = 9 And (mlVisCount = UBound(mlVisOrder) + 1) Then
        ' most likely an array needs enlarging
        If mlVisCount = UBound(mlVisOrder) + 1 Then
            ReDim Preserve mlVisOrder(LBound(mlVisOrder) To mlVisCount + 100)
            Resume
        End If
    ElseIf Err.Number = 18 Then
        ' user pressed ctrl-break in Excel
        MsgBox "Loading/refreshing Treeview aborted", , AppName
        NodesClear
        Resume done
    End If

    #If DebugMode = 1 Then
        Debug.Print Err.Number, Err.Description
        Stop
        Resume
    #End If

    Err.Raise Err.Number, "BuildRoot", Err.Description
End Sub

Sub HourGlassCursor(bApply As Boolean)

    #If HostProject = "Access" Then
        Application.DoCmd.Hourglass bApply
    #ElseIf HostProject = "Word" Then
        System.Cursor = IIf(bApply, wdCursorWait, wdCursorNormal)
    #Else
        Application.Cursor = IIf(bApply, xlWait, xlDefault)
    #End If

    #If Mac = False Then
        Dim pt As POINTAPI
        If Not bApply Then
            ' in some systems the cursor may fail to reset to default, this forces it
            GetCursorPos pt
            SetCursorPos pt.X, pt.Y
        End If
    #End If

End Sub

Private Sub BuildTree(cNode As clsNode, ByVal lLevel As Long, Optional lMaxLevel As Long = -1)
    Dim cChild As clsNode
    Dim lVLineTopIdx As Long

    ' disable error handling, let error return to caller
   ' On Error GoTo locErr

    If (lLevel > 1 Or mbRootButton) And mbShowExpanders Then

        'Expand/collapse button box (not needed if we use icons are used for expanders)
        If Not mbExpanderImage Then
            If cNode.ExpanderBox Is Nothing Then
                Set cNode.ExpanderBox = TreeControl.Controls.Add("Forms.label.1", "ExpBox" & cNode.Control.Name, False)
                With cNode.ExpanderBox
                    .Top = (mlVisCount - 1) * NodeHeight + msngTopExpB
                    .Left = mcTLpad * 2 + (lLevel - 2) * msngIndent + msngRootLine
                    .Width = mcExpBoxSize
                    .Height = mcExpBoxSize
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = mlLineColor
                    .BackStyle = fmBackStyleOpaque
                    .Visible = True
                End With
            Else
                With cNode.ExpanderBox
                    If mbRedesign Then .Left = mcTLpad * 2 + (lLevel - 2) * msngIndent + msngRootLine
                    .Top = (mlVisCount - 1) * NodeHeight + msngTopExpB
                    .Visible = True
                End With
            End If
        End If

        'Expand/collapse button text (or icon)
        If cNode.Expander Is Nothing Then
            Set cNode.Expander = TreeControl.Controls.Add("Forms.label.1", "ExpText" & cNode.Control.Name, False)
            With cNode.Expander
                .Left = mcTLpad * 2 + (lLevel - 2) * msngIndent + msngRootLine
                .Top = (mlVisCount - 1) * NodeHeight + msngTopExpT

                If mbExpanderImage Then
                    'Use an image
                    .AutoSize = True
                    .Width = 7.5
                    .Height = 7.5
                    .BorderStyle = fmBorderStyleNone
                    .PicturePosition = fmPicturePositionLeftTop
                    .Picture = moExpanderImage(cNode.Expanded)
                    #If Mac Then
                        .BackStyle = fmBackStyleTransparent
                    #End If
                Else
                    'use +/- text
                    .Width = mcExpButSize
                    .Height = mcExpButSize

                    If cNode.Expanded = True Then
                        .Caption = "-"
                        .Font.Bold = True
                    Else
                        .Caption = "+"
                        .Font.Bold = False
                    End If

                    .Font.Size = mcExpanderFont
                    .TextAlign = fmTextAlignCenter
                    .BackStyle = fmBackStyleTransparent
                End If
                .Visible = True
            End With
        Else
            With cNode.Expander
                If mbRedesign Then .Left = mcTLpad * 2 + (lLevel - 2) * msngIndent + msngRootLine
                .Top = (mlVisCount - 1) * NodeHeight + msngTopExpT
                .Visible = True
            End With
        End If

    End If  ' lLevel > 1 Or mbRootButton) And mbShowExpanders

    If cNode.Expanded And (lMaxLevel < lLevel Or lMaxLevel = -1) Then

        'Vertical line
        If mbShowLines Then
            If cNode.VLine Is Nothing Then
                Set cNode.VLine = TreeControl.Controls.Add("Forms.label.1", "VLine" & cNode.Control.Name, False)
                lVLineTopIdx = mlVisCount
                With cNode.VLine
                    .ZOrder 1
                    .Top = msngTopHV + (lVLineTopIdx - 1) * NodeHeight
                    .Left = mcLineLeft + msngRootLine + msngIndent * (lLevel - 1)
                    .Width = mcPtPxl
                    .Height = NodeHeight
                    .Caption = ""
                    .BorderColor = mlLineColor
                    .BorderStyle = fmBorderStyleSingle
                    .Visible = True
                End With

            Else
                lVLineTopIdx = mlVisCount
                With cNode.VLine
                    .Top = msngTopHV + (lVLineTopIdx - 1) * NodeHeight
                    If mbRedesign Then
                        .Left = mcLineLeft + msngRootLine + msngIndent * (lLevel - 1)
                        .Visible = True
                    End If
                End With
            End If
        End If

        For Each cChild In cNode.ChildNodes

            ' extend the vertical line
            If mbShowLines Then
                With cNode.VLine
                    .Height = (mlVisCount - lVLineTopIdx + 1) * msngNodeHeight
                    .Visible = True
                End With
            End If

            BuildNodeControls cChild, lLevel

            If Not cChild.ChildNodes Is Nothing Then
                BuildTree cChild, lLevel + 1
            End If

        Next

    End If    ' cNode.Expanded And (lMaxLevel < lLevel Or lMaxLevel = -1)

    Exit Sub

'locErr:
'    #If DebugMode = 1 Then
'        Stop
'        Resume
'    #End If
End Sub

Private Sub BuildNodeControls(cNode As clsNode, ByVal lLevel As Long)
' PT, create or (un)hide the controls, size & position to suit
' all created nodes have a caption, and optionally a horizontal line, checkbox and seperate icon

    Dim sngIconPad As Single
    Dim sName As String
    Dim vKey

    On Error GoTo locErr

  '  Application.EnableCancelKey = xlErrorHandler

    If cNode.Control Is Nothing Then
        mlNodesCreated = mlNodesCreated + 1
        sName = "Node" & mlNodesDeleted + mlNodesCreated
    ElseIf mbRedesign Then
         sName = cNode.Control.Name
    End If

    'Horizontal line
    If mbShowLines Then
        If cNode.HLine Is Nothing Then
            Set cNode.HLine = TreeControl.Controls.Add("Forms.label.1", "HLine" & sName, False)
            With cNode.HLine
                .Left = mcLineLeft + msngRootLine + msngIndent * (lLevel - 1)
                .Top = msngTopHV + mlVisCount * NodeHeight
                .Width = msngIndent
                .Height = mcPtPxl
                .Caption = ""
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = mlLineColor 'vbScrollBars
                 If mbRedesign Then
                    .ZOrder 1
                 End If
                .Visible = True
            End With
        Else
            With cNode.HLine
                If mbRedesign Then
                    .Left = mcLineLeft + msngRootLine + msngIndent * (lLevel - 1)
                    .Width = msngIndent
                End If
                .Top = msngTopHV + mlVisCount * NodeHeight
                .Visible = True
            End With
        End If
    End If

    ' Checkbox
    If CheckBoxes Then
        If cNode.Checkbox Is Nothing Then
            Set cNode.Checkbox = TreeControl.Controls.Add("Forms.label.1", "CheckBox" & sName, False)
            With cNode.Checkbox
                .Left = mcTLpad + msngRootLine + msngIndent * lLevel
                .Top = mlVisCount * NodeHeight + msngTopChk

                If mbCheckboxImage Then
                    'Use an image
                    .BorderStyle = fmBorderStyleNone
                    .Picture = moCheckboxImage(cNode.Checked)
                    .PicturePosition = fmPicturePositionLeftBottom
                    .AutoSize = True
                Else

                    .Width = mcChkBoxSize
                    .Height = mcChkBoxSize
                    .Font.Name = "Marlett"
                    #If Mac = False Then
                        .Font.Charset = SYMBOL_CHARSET
                    #End If
                    .Font.Size = 10
                    .TextAlign = fmTextAlignCenter
                    .BorderStyle = fmBorderStyleSingle
                    If cNode.Checked Then
                        .Caption = "a"
                        If cNode.Checked = 1 Then
                            .ForeColor = RGB(180, 180, 180)
                        End If
                    End If
                End If

                .Visible = True
            End With
        Else
            With cNode.Checkbox
                If mbRedesign Then .Left = mcTLpad + msngRootLine + msngIndent * lLevel
                .Top = mlVisCount * NodeHeight + msngTopChk
                .Visible = True
            End With
        End If
    End If

    ' Icon, in its own image control if using FullWidth, otherwise it goes in the label
    If mbFullWidth And mbGotIcons Then
        If cNode.hasIcon(vKey) Then
            sngIconPad = mcIconPad
            If cNode.Icon Is Nothing Then
                Set cNode.Icon = TreeControl.Controls.Add("Forms.Image.1", "Icon" & sName, False)
                With cNode.Icon
                    .BorderStyle = fmBorderStyleNone
                    .Left = mcTLpad + msngRootLine + msngIndent * lLevel + msngChkBoxPad
                    .Top = mlVisCount * NodeHeight + msngTopIcon
                    '.AutoSize
                    .Width = mcIconSize
                    .Height = mcIconSize
                    .BackStyle = fmBackStyleTransparent
                    .Picture = mcolIcons(vKey)
                    .BackStyle = fmBackStyleTransparent
                    .Visible = True
                End With
            Else
                With cNode.Icon
                    If mbRedesign Then
                        .Left = mcTLpad + msngRootLine + msngIndent * lLevel + msngChkBoxPad
                    End If
                    .Top = mlVisCount * NodeHeight + msngTopIcon
                    .Visible = True
                End With
            End If
        Else
            sngIconPad = 0
        End If
    End If
    
    'The node itself
    If cNode.Control Is Nothing Then
        
        Set cNode.Control = TreeControl.Controls.Add("Forms.label.1", sName, False)
        With cNode.Control
            .WordWrap = False
            .AutoSize = True
            .Left = mcTLpad + msngRootLine + msngIndent * lLevel + msngChkBoxPad + sngIconPad
            .Top = mcTLpad + msngTopLabel + mlVisCount * NodeHeight

            If Not mbFullWidth And mbGotIcons Then
                If cNode.hasIcon(vKey) Then
                    .PicturePosition = fmPicturePositionLeftCenter
                    .Picture = mcolIcons(vKey)
                End If
            End If

            If cNode.Bold Then .Font.Bold = True
            .WordWrap = False
            .AutoSize = True
            .Caption = cNode.Caption
            cNode.TextWidth = .Width

            If cNode.TextWidth + sngIconPad > msngMaxWidths(lLevel) Then
                msngMaxWidths(lLevel) = cNode.TextWidth + sngIconPad
            End If

            If mbFullWidth Then
                .AutoSize = False
                If .Width <= mcFullWidth Then .Width = mcFullWidth
            End If
            If cNode.BackColor Then
                .BackColor = cNode.BackColor
            End If
            If cNode.ForeColor Then
                .ForeColor = cNode.ForeColor
            End If
            
            If Len(cNode.ControlTipText) Then
                .ControlTipText = cNode.ControlTipText
            End If
            
            .Visible = True
        End With

    Else
        With cNode.Control
            If mbRedesign Then
                .Left = mcTLpad + msngRootLine + msngIndent * lLevel + sngIconPad + msngChkBoxPad

                If cNode.TextWidth + sngIconPad > msngMaxWidths(lLevel) Then
                    msngMaxWidths(lLevel) = cNode.TextWidth + sngIconPad
                End If
            End If

            .Top = mlVisCount * NodeHeight + mcTLpad + msngTopLabel
            .Visible = True
        End With

    End If

    mlVisCount = mlVisCount + 1
    mlVisOrder(mlVisCount) = cNode.Index
    cNode.VisIndex = mlVisCount

    Exit Sub

locErr:
    If Err.Number = 9 Then
        ' most likely an array needs enlarging
        If mlVisCount = UBound(mlVisOrder) + 1 Then
            ReDim Preserve mlVisOrder(LBound(mlVisOrder) To mlVisCount + 100)
            Resume
        ElseIf lLevel > UBound(msngMaxWidths) Then
            ReDim Preserve msngMaxWidths(LBound(msngMaxWidths) To lLevel + 5)
            Resume
        End If
    ElseIf Err.Number = 18 Then
        Err.Raise 18    ' user pressed ctrl-break, pass to BuildRoot
    Else
        #If DebugMode = 1 Then
            Stop
            Resume
        #End If
        Err.Raise Err.Number, "BuildNodeControls", Err.Description
    End If

End Sub

Private Sub Clone(cParent As clsNode, cNode As clsNode, Optional vBefore, Optional ByVal vAfter)
' PT clone a node and add the 4-way references
    Dim bTriStateOrig As Boolean
    Dim cClone As clsNode
    Dim cChild As clsNode
     
    On Error GoTo errH

    If cParent Is Nothing Or cNode Is Nothing Then
        Exit Sub
    End If

    bTriStateOrig = mbTriState
    mbTriState = False
    
    Set cClone = New clsNode

    With cNode
        If .BackColor = 0 Then .BackColor = mlBackColor
        cClone.BackColor = .BackColor
        cClone.Caption = .Caption
        cClone.Checked = .Checked
        cClone.Expanded = .Expanded
        If .ForeColor = 0 Then .ForeColor = mlForeColor
        cClone.ImageExpanded = .ImageExpanded
        cClone.ImageMain = .ImageMain
        cClone.ForeColor = .ForeColor
        cClone.Key = .Key
    End With

    If cParent.ChildNodes Is Nothing Then
        Set cParent.ChildNodes = New Collection
    End If

    Set cClone.ParentNode = cParent

    If Not cNode.ChildNodes Is Nothing Then
        For Each cChild In cNode.ChildNodes
            Clone cClone, cChild    ' don't pass vBefore/vAfter
        Next
    End If

    AddNodeToCol cParent.ChildNodes, cClone, False, vBefore, vAfter

    Set cClone.Tree = Me
    
    AddNodeToCol mcolNodes, cClone, bTreeCol:=True

    cClone.Index = Nodes.Count
    
    mbTriState = bTriStateOrig
    If mbTriState Then
        cClone.ParentNode.CheckTriStateParent
    End If
    
    Exit Sub

errH:
    #If DebugMode = 1 Then
        Debug.Print Err.Description
        Stop
        Resume
    #End If
    mbTriState = bTriStateOrig
End Sub

Private Function MaxNodeWidth() As Single
'-------------------------------------------------------------------------
' Procedure : MaxNodeWidth
' Author    : Peter Thornton
' Created   : 27-01-2013
' Purpose   : Get the max right for horizontal scroll
'-------------------------------------------------------------------------
    Dim lLevel As Long
    Dim sngMax As Single

    ''' msngMaxWidths(), contains maximum text-width + additional icon width (if any) in each level
    '  tot-width = 3 + msngRootLine + msngIndent * lLevel + msngChkBoxPad + [ msngIconPad + text-width]

    For lLevel = 0 To UBound(msngMaxWidths)
        If msngMaxWidths(lLevel) Then
            If mcTLpad + msngRootLine + msngIndent * lLevel + msngChkBoxPad + msngMaxWidths(lLevel) > sngMax Then
                sngMax = mcTLpad + msngRootLine + msngIndent * lLevel + msngChkBoxPad + msngMaxWidths(lLevel)
            End If
        End If
    Next
    MaxNodeWidth = sngMax
    
End Function

Private Function NextVisibleNodeInTree(ByRef cStartNode As clsNode, Optional bUp As Boolean = True) As clsNode
'-------------------------------------------------------------------------
' Procedure : NextVisibleNodeInTree
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 16-01-2013
' Purpose   : Function that returns either the next or the previous node adjacent to the active node
'-------------------------------------------------------------------------

    Dim lStep As Long
    Dim lNextVis As Long    'PT

    On Error GoTo errH
    If bUp Then lStep = -1 Else lStep = 1

    If cStartNode Is Nothing Then
        Set NextVisibleNodeInTree = mcolNodes(1)
    Else

        lNextVis = cStartNode.VisIndex
        lNextVis = lNextVis + lStep
        If lNextVis >= 1 And lNextVis <= mlVisCount Then
            lNextVis = mlVisOrder(lNextVis)
            Set NextVisibleNodeInTree = mcolNodes(lNextVis)
        End If
    End If
    Exit Function
errH:
    #If DebugMode = 1 Then
        Stop
        Debug.Print Err.Description
        Resume
    #End If
End Function

Private Function NodeIsVisible(Optional cNode As clsNode, Optional lngCntVisible As Long) As Boolean
Dim idxFirstVis As Long
Dim idxLastVis As Long

    If TreeControl Is Nothing Then
        Exit Function
    End If

    With TreeControl
        idxFirstVis = .ScrollTop / NodeHeight + 1
        lngCntVisible = (.InsideHeight - mcTLpad) / NodeHeight
        idxLastVis = lngCntVisible + idxFirstVis - 1
    End With

    If cNode Is Nothing Then
        If Not ActiveNode Is Nothing Then

            Set cNode = ActiveNode
        Else
            Exit Function
        End If
    End If

    If idxLastVis > mlVisCount Then idxLastVis = mlVisCount

    If Not cNode Is Nothing Then
        NodeIsVisible = cNode.VisIndex >= idxFirstVis And cNode.VisIndex <= idxLastVis
    End If

End Function

Private Sub RecalcMaxTextWidth()
Dim i As Long
Dim lLevel As Long
Dim cNode As clsNode
Dim sngIconPad As Single
Dim v

    For i = LBound(msngMaxWidths) To UBound(msngMaxWidths)
        msngMaxWidths(i) = 0
    Next
    
    For i = 1 To mlVisCount
        Set cNode = mcolNodes(mlVisOrder(i))
        lLevel = cNode.Level
        If Me.FullWidth Then
            sngIconPad = 0
            If cNode.hasIcon(v) Then
                sngIconPad = mcIconPad
            End If
        End If

        If cNode.TextWidth + sngIconPad > msngMaxWidths(lLevel) Then
            msngMaxWidths(lLevel) = cNode.TextWidth + sngIconPad
        End If
    Next

End Sub

Private Sub ResetActiveNodeColor(cNode As clsNode)
    Dim lBColor As Long
    Dim lFColor As Long
    If Not cNode Is Nothing Then
        lBColor = cNode.BackColor
        lFColor = cNode.ForeColor
        With cNode.Control
            .BorderStyle = fmBorderStyleNone
            .BackColor = IIf(lBColor, lBColor, mlBackColor)
            .ForeColor = IIf(lFColor, lFColor, mlForeColor)
        End With
    End If
End Sub

Private Sub Round75()
'-------------------------------------------------------------------------
' Procedure : Round75
' Author    : Peter Thornton
' Created   : 29-01-2013
' Purpose   : Make size & position dims a factor of 0.75 points (units of 1 pixel)
'-------------------------------------------------------------------------
#If Mac Then
    msngTopExpB = Int(msngTopExpB)
    msngTopExpT = Int(msngTopExpT)
    msngTopHV = Int(msngTopHV)
    msngTopIcon = Int(msngTopIcon)
    msngTopChk = Int(msngTopChk)
    msngTopLabel = Int(msngTopLabel)
#Else
    msngTopExpB = Int((msngTopExpB * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
    msngTopExpT = Int((msngTopExpT * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
    msngTopHV = Int((msngTopHV * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
    msngTopIcon = Int((msngTopIcon * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
    msngTopChk = Int((msngTopChk * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
    msngTopLabel = Int((msngTopLabel * 2 + mcPtPxl) / 3 * 2) * mcPtPxl
#End If
End Sub

Private Sub SetActiveNodeColor(Optional bInactive)

    If Not ActiveNode Is Nothing Then

        If IsMissing(bInactive) Then
            On Error Resume Next
            #If HostProject = "Access" Then
                bInactive = mbInActive
            #Else
                bInactive = Not Me.TreeControl Is Me.TreeControl.Parent.ActiveControl
            #End If
            On Error GoTo 0
        End If

         ' system highlight colours, bInactive set in Sub EnterExit called from TreeControl's Enter/Exit events (Excel) / subForm's OnEnter/OnExit events (Access)

        With ActiveNode.Control
            If bInactive Then
            ''' when treeeview not in focus
            
                ResetActiveNodeColor moActiveNode
                '' just a grey border
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = RGB(190, 190, 190)
                
                '' inactive colours
'                .BackColor = vbInactiveTitleBar
'                .ForeColor = vbWindowText
            Else
                ' in focus
                .BorderStyle = fmBorderStyleNone
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End If
        End With

    End If
End Sub

Private Sub SetTreeExpansionLevel(lLevel As Long, Optional lCurLevel As Long, _
                                          Optional cNode As clsNode, Optional bExit As Boolean = False)
'-------------------------------------------------------------------------
' Procedure : SetTreeExpansionLevel
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 17-01-2013
' Purpose   : Updates the expanded properties according to lLevel
'             Called recursively.
'-------------------------------------------------------------------------
    Dim cChild As clsNode
    If bExit Then Exit Sub
    If cNode Is Nothing Then

        For Each cNode In moRootHolder.ChildNodes
            If lLevel > -1 Then
                cNode.Expanded = True
            Else
                cNode.Expanded = False
            End If
            If Not cNode.ChildNodes Is Nothing Then
                For Each cChild In cNode.ChildNodes
                    cChild.Expanded = (lLevel > lCurLevel)
                    SetTreeExpansionLevel lLevel, lCurLevel + 1, cChild, False
                Next
            End If
        Next

    ElseIf Not cNode.ChildNodes Is Nothing Then
        For Each cChild In cNode.ChildNodes
            cChild.Expanded = (lLevel > lCurLevel)
            SetTreeExpansionLevel lLevel, lCurLevel + 1, cChild, False
        Next
    End If
    
End Sub

Private Sub TreeControl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim sngVisTop As Single
    Dim cNode As clsNode

    ' PT toggle expand/collapse with key Enter
    If KeyCode = vbKeyReturn Then
        If ActiveNode.Expanded Then
            KeyCode = vbKeyLeft
        Else
            KeyCode = vbKeyRight
        End If
    End If

    If Not ActiveNode Is Nothing Then
        Select Case KeyCode
        Case vbKeyLeft
            If ActiveNode.Level = 0 And Not mbRootButton Then
                ' don't attempt to collapse the Root if it doesn't have a button

            ElseIf Not ActiveNode.ChildNodes Is Nothing Then
                If ActiveNode.Expanded Then
                    ActiveNode.Expanded = False
                     NodeEventRouter ActiveNode, "Expander", 1
                Else
                    If Not ActiveNode.ParentNode Is Nothing Then
                        If ActiveNode.ParentNode.Expanded Then
                            'If Not ActiveNode.ParentNode.Level = 0 And mbRootButton Then
                            If ActiveNode.ParentNode.Level >= 0 Then
                                Set ActiveNode = ActiveNode.ParentNode
                                ScrollToView , -1
                                NodeEventRouter ActiveNode, "Caption", 1
                            End If
                        End If
                    End If
                End If
            Else
                If Not ActiveNode.ParentNode Is Nothing Then
                    If ActiveNode.ParentNode.Level = 0 And Not mbRootButton Then
                        ' don't attempt to collapse the Root if it doesn't have a button
                        ' redundant ?
                    ElseIf ActiveNode.ParentNode.Expanded Then
                        If ActiveNode.ParentNode.Caption <> "RootHolder" Then
                            Set ActiveNode = ActiveNode.ParentNode
                            ScrollToView ActiveNode, -1
                            NodeEventRouter ActiveNode, "Caption", 1
                        End If
                    End If
                End If
            End If

        Case vbKeyRight
            If Not ActiveNode.ChildNodes Is Nothing Then
                If ActiveNode.Expanded = False Then
                    ActiveNode.Expanded = True
                    If Not ActiveNode.Expander Is Nothing Then
                        NodeEventRouter ActiveNode, "Expander", 1
                    Else
                        ' a Root node and mbRootButton = False
                        BuildRoot False
                    End If
                Else
                    Set ActiveNode = ActiveNode.ChildNodes(1)
                    NodeEventRouter ActiveNode, "Caption", 1
                End If

            End If

        Case vbKeyUp, vbKeyDown
            If ActiveNode.VisIndex = mlVisCount And KeyCode = vbKeyDown Then
                ' if the activenode is the last node and collaped, expand it and activate the 1st childnode
                If Not ActiveNode.ChildNodes Is Nothing Then
                    If ActiveNode.Expanded = False Then
                        ActiveNode.Expanded = True
                        BuildRoot False
                    End If
                End If
            End If

            Set cNode = NextVisibleNodeInTree(ActiveNode, (KeyCode = vbKeyUp))

            If Not cNode Is Nothing Then
                Set ActiveNode = cNode
                ScrollToView ActiveNode, IIf(KeyCode = vbKeyUp, -1, -2) ' the -ve means will scroll won't change if the node is visible

                NodeEventRouter ActiveNode, "Caption", 1
            End If

        Case vbKeyPageUp, vbKeyPageDown
            'store the activenode's vertical position to reset a similar in the keyup
            If Not mbKeyDown Then
                sngVisTop = (ActiveNode.VisIndex - 1) * NodeHeight - TreeControl.ScrollTop
                If sngVisTop > 0 And sngVisTop < TreeControl.InsideHeight Then
                    msngVisTop = sngVisTop
                Else
                    msngVisTop = 0
                End If
            End If

        Case vbKeyEscape
            Set MoveCopyNode(False) = Nothing

        Case vbKeySpace  ' PT toggle checkbox with space
            If CheckBoxes Then
                ActiveNode.Checked = Not ActiveNode.Checked
                'NodeEventRouter called in node.Checked/UpdateCheckbox
            End If
        End Select

        mbKeyDown = True    ' PT

        RaiseEvent KeyDown(ActiveNode, KeyCode, Shift)
    Else
        If Not mcolNodes Is Nothing Then
            If mcolNodes.Count Then
                Set ActiveNode = mcolNodes(1)
            End If
        End If
    End If

End Sub

Private Sub TreeControl_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'-------------------------------------------------------------------------
' Procedure : TreeControl_KeyUp
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 17-01-2013
' Purpose   : Handles collapsing and expanding of the tree using left and right arrow keys
'             and moving up/down the tree using up/down arrow keys
'             Also handles folding of the tree when you use the numeric keys.
'-------------------------------------------------------------------------
    Dim lIdx As Long
    Dim sngNewScrollTop As Single

    If Not mbKeyDown Then    'PT
        ' PT KeyDown was initiated in some other control,
        '   eg Key Enter in the Editbox or tabbing to the treecontrol (enter event)
        Exit Sub
    Else

        mbKeyDown = False
    End If

    If Not ActiveNode Is Nothing Then

        #If Mac Then
            If mbLabelEdit And Shift = 2 And KeyCode = 85 Then  ' 18/5/21 moved this #If Mac construct from KeyDown and below Select Case KeyCode
                KeyCode = vbKeyF2  ' Ctrl-U = F2 edit
            End If
        #End If

        Select Case KeyCode

        ' PT look into moving more key events into KeyDown

        Case 48 To 57, 96 To 105
            If KeyCode >= 96 Then KeyCode = KeyCode - 48
            If (KeyCode > vbKey0 Or mbRootButton) And Shift = 0 Then

                ExpandToLevel (KeyCode - 48)
                BuildRoot False
            End If

        Case vbKeyF2, 93   ' F2 & key right/context menu (?) PT
            If mbLabelEdit Then
                If Not ActiveNode Is Nothing Then
                    EditMode(ActiveNode) = True
                    ActiveNode.EditBox True
                End If
            End If
        Case vbKeyPageUp, vbKeyPageDown
            ' PT activate node in the same position as previous activenode when scrolling
            With Me.TreeControl
                sngNewScrollTop = .ScrollTop
                lIdx = (sngNewScrollTop + msngVisTop) / NodeHeight + 1

                If (lIdx - 1) * NodeHeight < .ScrollTop Then
                    lIdx = lIdx + 1

                ElseIf lIdx * NodeHeight > .InsideHeight + .ScrollTop Then
                    lIdx = lIdx - 1
                End If
            End With

            If lIdx > 1 And lIdx <= mlVisCount Then
                lIdx = mlVisOrder(lIdx)
                Set ActiveNode = mcolNodes(lIdx)
            End If
            
        Case vbKeyHome, vbKeyEnd
            If KeyCode = vbKeyHome Then lIdx = 1 Else lIdx = mlVisCount
            lIdx = mlVisOrder(lIdx)
            If ActiveNode.Index <> lIdx Then
                Set ActiveNode = mcolNodes(lIdx)
            End If
        Case Else

        End Select
    Else
        If Not mcolNodes Is Nothing Then
            If mcolNodes.Count Then
                Set ActiveNode = mcolNodes(1)
            End If
        End If
    End If
    
End Sub

