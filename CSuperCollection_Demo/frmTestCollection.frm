VERSION 5.00
Begin VB.Form frmTestCollection 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test the CSuperCollection class"
   ClientHeight    =   6645
   ClientLeft      =   2805
   ClientTop       =   1455
   ClientWidth     =   9495
   Icon            =   "frmTestCollection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   Begin VB.Frame Frame1 
      Caption         =   "key name or index"
      Height          =   1005
      Index           =   8
      Left            =   60
      TabIndex        =   29
      Top             =   5595
      Width           =   1815
      Begin VB.CommandButton cmdRaiseEvent 
         Caption         =   "Raise Event!"
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox txRaiseEvent 
         Height          =   300
         Left            =   330
         TabIndex        =   30
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdIterateKeys 
      Caption         =   "Iterate Keys"
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   5100
      Width           =   1650
   End
   Begin VB.CheckBox ckAllowAssignments 
      Caption         =   "Allow Item Assignments"
      Height          =   420
      Left            =   210
      TabIndex        =   48
      Top             =   2760
      Width           =   1425
   End
   Begin VB.OptionButton optCompare 
      Caption         =   "Binary Compare"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   46
      Top             =   1425
      Value           =   -1  'True
      Width           =   1485
   End
   Begin VB.OptionButton optCompare 
      Caption         =   "Text Compare"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   47
      Top             =   2010
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "change item name"
      Height          =   1365
      Index           =   7
      Left            =   7650
      TabIndex        =   23
      Top             =   4050
      Width           =   1815
      Begin VB.TextBox txAssign 
         Height          =   300
         Index           =   0
         Left            =   510
         TabIndex        =   25
         Top             =   225
         Width           =   1155
      End
      Begin VB.TextBox txAssign 
         Height          =   300
         Index           =   1
         Left            =   510
         TabIndex        =   27
         Top             =   585
         Width           =   1155
      End
      Begin VB.CommandButton cmdAssign 
         Caption         =   "Assign To Item"
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Key: "
         Height          =   225
         Index           =   9
         Left            =   75
         TabIndex        =   24
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "New:"
         Height          =   225
         Index           =   8
         Left            =   75
         TabIndex        =   26
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "item name and key"
      Height          =   1365
      Index           =   6
      Left            =   1950
      TabIndex        =   5
      Top             =   4050
      Width           =   1815
      Begin VB.TextBox txAddSingleItem 
         Height          =   300
         Index           =   0
         Left            =   570
         TabIndex        =   7
         Top             =   210
         Width           =   1155
      End
      Begin VB.TextBox txAddSingleItem 
         Height          =   300
         Index           =   1
         Left            =   570
         TabIndex        =   9
         Top             =   585
         Width           =   1155
      End
      Begin VB.CommandButton cmdAddSingleItem 
         Caption         =   "Add Item"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Name: "
         Height          =   225
         Index           =   6
         Left            =   75
         TabIndex        =   6
         Top             =   255
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Key:"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "key name"
      Height          =   1005
      Index           =   5
      Left            =   7650
      TabIndex        =   43
      Top             =   5595
      Width           =   1815
      Begin VB.TextBox txKeyExists 
         Height          =   300
         Left            =   330
         TabIndex        =   44
         Top             =   225
         Width           =   1155
      End
      Begin VB.CommandButton cmdKeyExists 
         Caption         =   "Key Exists?"
         Height          =   315
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "key name or index"
      Height          =   1005
      Index           =   4
      Left            =   3855
      TabIndex        =   36
      Top             =   5595
      Width           =   1815
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Item"
         Height          =   315
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox txRemove 
         Height          =   300
         Left            =   330
         TabIndex        =   37
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "enter *ANY* number"
      Height          =   1005
      Index           =   3
      Left            =   1950
      TabIndex        =   32
      Top             =   5595
      Width           =   1815
      Begin VB.TextBox txBaseIndex 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   225
         Width           =   540
      End
      Begin VB.CommandButton cmdBaseIndex 
         Caption         =   "Set Base Index"
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Index: "
         Height          =   225
         Index           =   2
         Left            =   315
         TabIndex        =   33
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "enter index number"
      Height          =   1005
      Index           =   2
      Left            =   5745
      TabIndex        =   39
      Top             =   5595
      Width           =   1815
      Begin VB.CommandButton cmdGetKey 
         Caption         =   "Get Key"
         Height          =   315
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox txGetKey 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Index: "
         Height          =   225
         Index           =   7
         Left            =   315
         TabIndex        =   40
         Top             =   255
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "old and new key name"
      Height          =   1365
      Index           =   1
      Left            =   5745
      TabIndex        =   17
      Top             =   4050
      Width           =   1815
      Begin VB.TextBox txChangeKey 
         Height          =   300
         Index           =   0
         Left            =   510
         TabIndex        =   19
         Top             =   225
         Width           =   1155
      End
      Begin VB.TextBox txChangeKey 
         Height          =   300
         Index           =   1
         Left            =   510
         TabIndex        =   21
         Top             =   585
         Width           =   1155
      End
      Begin VB.CommandButton cmdChangeKey 
         Caption         =   "Change Key"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "From: "
         Height          =   225
         Index           =   5
         Left            =   75
         TabIndex        =   18
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "To:"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "key name or index"
      Height          =   1365
      Index           =   0
      Left            =   3855
      TabIndex        =   11
      Top             =   4050
      Width           =   1815
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move Item"
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1380
      End
      Begin VB.TextBox txMove 
         Height          =   300
         Index           =   1
         Left            =   510
         TabIndex        =   15
         Top             =   585
         Width           =   1155
      End
      Begin VB.TextBox txMove 
         Height          =   300
         Index           =   0
         Left            =   510
         TabIndex        =   13
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "To:"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From: "
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Collection"
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   3885
      Width           =   1650
   End
   Begin VB.CommandButton cmdNestedWalk 
      Caption         =   "Nested For...Each"
      Height          =   315
      Left            =   105
      TabIndex        =   3
      Top             =   4695
      Width           =   1650
   End
   Begin VB.TextBox txDisplay 
      Height          =   3945
      Left            =   1860
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   45
      Width           =   7590
   End
   Begin VB.CommandButton cmdWalk 
      Caption         =   "For...Each"
      Height          =   315
      Left            =   105
      TabIndex        =   2
      Top             =   4290
      Width           =   1650
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Items"
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   3480
      Width           =   1650
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   5
      X1              =   12
      X2              =   114
      Y1              =   220
      Y2              =   220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   13
      X2              =   115
      Y1              =   219
      Y2              =   219
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   12
      X2              =   114
      Y1              =   84
      Y2              =   84
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   13
      X2              =   115
      Y1              =   83
      Y2              =   83
   End
   Begin VB.Label Label1 
      Caption         =   "(case insensitive)"
      Height          =   240
      Index           =   11
      Left            =   465
      TabIndex        =   51
      Top             =   2250
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "(case sensitive)"
      Height          =   240
      Index           =   10
      Left            =   465
      TabIndex        =   50
      Top             =   1665
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   13
      X2              =   115
      Y1              =   174
      Y2              =   174
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   12
      X2              =   114
      Y1              =   175
      Y2              =   175
   End
   Begin VB.Label Label1 
      Caption         =   "This form allows you to test drive the properties and methods of the Super Collection object."
      Height          =   1005
      Index           =   12
      Left            =   120
      TabIndex        =   52
      Top             =   210
      Width           =   1710
   End
End
Attribute VB_Name = "frmTestCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

  Private Const MSGBX_TITLE As String = "Super Collection"

  Private Const KEY_NOT_FOUND As Long = (-&HEFFFFFFF)

  Private WithEvents m_oCol As CSuperCollection
Attribute m_oCol.VB_VarHelpID = -1

Private Sub cmdRaiseEvent_Click()
  
  Dim sKeyOrIndex$, nIndex&
  
  On Error GoTo EH
  
  sKeyOrIndex = txRaiseEvent.Text
  
  txDisplay.Text = txDisplay.Text & vbCrLf & "Raising Event In Item: " & sKeyOrIndex & vbCrLf
  txDisplay.SelStart = Len(txDisplay.Text)
  
  
  If IsNumeric(sKeyOrIndex) Then
    nIndex = CLng(sKeyOrIndex)
    
    m_oCol(nIndex).FireEvent
    
  Else
    m_oCol(sKeyOrIndex).FireEvent
  End If
  
ExitNow:
  txDisplay.Text = txDisplay.Text & vbCrLf & "End Raising Event In Item: " & sKeyOrIndex & vbCrLf
  txDisplay.SelStart = Len(txDisplay.Text)
  
  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  txDisplay.Text = txDisplay.Text & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow
  
End Sub

Private Sub m_oCol_GenericEvent(ByVal Index As Long, ByVal Key As String)
  
  Dim sText$
  
  sText = "  Event received from item number: " & CStr(Index) & "  Key: " & IIf(Key <> vbnullstring, Key, "No Key Entered")
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)
  
  MsgBox Trim$(sText), vbInformation, MSGBX_TITLE
  
End Sub

Private Sub Form_DblClick()
  
  Dim i&, oCol As Collection, sString As Variant, sString2 As Variant, sString3 As Variant
  Set oCol = New Collection
  
  oCol.Add "Apple", "Apple"
  oCol.Add "Pear", "Pear"
  oCol.Add "Banana", "Banana"
  oCol.Add "Zuchini", "Zuchini"
  
  For Each sString In oCol
    Debug.Print sString
    
    For Each sString2 In oCol
      'oCol.Remove "Apple"
'      For Each sString3 In oCol
'        Debug.Print sString3
'      Next
      
      Debug.Print sString2
      'oCol.Remove "Apple"
      'Set oCol = Nothing
      'oCol.Add "Poppie", "Poppie", 1
    Next
  Next

  Set oCol = Nothing

End Sub

Private Sub optCompare_Click(Index As Integer)
  
  Static bInSub As Boolean
  Dim i&

  If bInSub = False Then
    bInSub = True
    
    On Error GoTo EH
    
    Select Case True
      Case optCompare(0).Value = True
        i = 1
        m_oCol.CompareMode = BinaryCompare
        
      Case optCompare(1).Value = True
        i = 0
        m_oCol.CompareMode = TextCompare
        
    End Select
    
ExitNow:
    
    bInSub = False
  End If

  Exit Sub
  
EH:
  Select Case Err.Number
    Case AlreadyInitialized
      MsgBox Err.Description & vbCrLf & "Clear the collection and try again.", vbInformation, MSGBX_TITLE
  
    Case Else
      MsgBox Err.Description, vbInformation, MSGBX_TITLE
  End Select
  
  Err.Clear
  optCompare(i).Value = True
  Resume ExitNow
  
End Sub
  
Private Sub ckAllowAssignments_Click()
  ' the property that allows us to do item issignments
  m_oCol.AllowItemAsignments = (ckAllowAssignments.Value = vbChecked)
End Sub

Private Sub cmdAdd_Click()
  ' add some items to the collection
  Dim i&, sItemKey$, eCategory As FRUIT_CATEGORY, oFruit As cFruit, sText$
  
  With m_oCol
    
    ' we first have to clear the collection
    cmdClear_Click
    
    sText = vbCrLf & "Adding Items" & vbCrLf & "  Items/Keys: "
    
    For i = 0 To 9
      Select Case i
        Case 0: sItemKey = "Apple": eCategory = Pome
        Case 1: sItemKey = "Orange": eCategory = Citrus
        Case 2: sItemKey = "Pear": eCategory = Pome
        Case 3: sItemKey = "Banana": eCategory = Other
        Case 4: sItemKey = "Peach": eCategory = Pome
        Case 5: sItemKey = "Grape": eCategory = Berry
        Case 6: sItemKey = "Pomegranate": eCategory = Berry
        Case 7: sItemKey = "Lemon": eCategory = Citrus
        Case 8: sItemKey = "Tangerine": eCategory = Citrus
        Case 9: sItemKey = "Mango": eCategory = Other
      End Select
        
      Set oFruit = New cFruit
      With oFruit
        .Key = sItemKey
        .Name = sItemKey
        .CollectionPointer = ObjPtr(m_oCol)
        .Category = eCategory
      End With
      
      If i Then sText = sText & ", "
      sText = sText & sItemKey & "/" & sItemKey

      .Add oFruit, sItemKey
      
      Set oFruit = Nothing
    Next
  
  
    sText = sText & vbCrLf & "  Keys Collection (Name/Assoc Index: "
    
    sText = sText & WalkKeysCollection()
    
    
    sText = sText & vbCrLf & "  Indexes: " & CStr(.BaseIndex) & " To: " & CStr(.BaseIndex + IIf(.Count, (.Count - 1), .Count))
    sText = sText & vbCrLf & "Count After Add Items: " & .Count & vbCrLf
    

  End With
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)
  
End Sub

Private Sub cmdAddSingleItem_Click()
  
  Dim sName$, sKey$, oFruit As cFruit, sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Add Single Item"
  
  sText = sText & vbCrLf & "  Count Before Add: " & m_oCol.Count & vbCrLf
  
  sName = Trim$(txAddSingleItem(0).Text)
  
  If sName <> vbnullstring Then
    sKey = Trim$(txAddSingleItem(1).Text)
    
    sText = sText & "     Adding: " & sName & "    Key: " & IIf(sKey <> vbnullstring, sKey, "No Key Entered")
    
    Set oFruit = New cFruit
    
    With oFruit
      .Name = sName
      .Key = sKey
    End With
    
    Select Case sKey
      Case vbnullstring: m_oCol.Add oFruit
      Case Else: m_oCol.Add oFruit, sKey
    End Select

    Set oFruit = Nothing
    
  Else
    sText = sText & "  Error!  No item name entered."

    MsgBox "You must enter a name to add to the collection.", vbInformation, MSGBX_TITLE
  End If
  
ExitNow:
    With m_oCol
      sText = sText & vbCrLf & "     Indexes: " & CStr(.BaseIndex) & " To: " & CStr(.BaseIndex + IIf(.Count, (.Count - 1), .Count))
    
      sText = sText & vbCrLf & "  Count After Add: " & .Count
    End With
    
    sText = sText & vbCrLf & "End Add Single Item" & vbCrLf
    
    txDisplay.Text = txDisplay.Text & sText
    txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow
  
End Sub

Private Sub cmdClear_Click()
  
  Dim sText$
  
  sText = vbCrLf & "Clearing Collection" & vbCrLf
  
  With m_oCol
    sText = sText & "Count Before: " & .Count & vbCrLf
    ' clear the collection
    .Clear
    sText = sText & "Count After: " & .Count & vbCrLf
  End With
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)
  
End Sub

Private Sub cmdWalk_Click()

  Dim i&, oFruit As cFruit, sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "For...Each" & vbCrLf & "  Keys: "
  
  
  ' this is an example of what is actually going on behind the scenes when you implement a For...Each loop
  Dim oEnumer As IEnumVARIANTReDef, vVarRet As Variant, nFetched&
  
  Set oEnumer = m_oCol.NewEnum
  
  Do
    oEnumer.Next 1, vVarRet, VarPtr(nFetched)
  
    If nFetched Then
      Set oFruit = vVarRet
      
      If i Then sText = sText & ", "
      sText = sText & IIf(oFruit.Key <> vbnullstring, oFruit.Key, "No Key Entered")
      i = i + 1
      
      Set oFruit = Nothing
      
      ' notice that you could call any of the methods on the IEnumVARIANT interface.
      ' the Skip, Clone and Reset methods are normally unused in VB
      'oEnumer.Skip 2
      'oEnumer.Clone
      'oEnumer.Reset
    End If
    
    Set vVarRet = Nothing
    
  Loop Until nFetched = 0
  
  Set oEnumer = Nothing
  
  ' this is how the above example would look using For...Each syntax except you could not call the Skip, Clone or Reset methods on the IEnumVARIANT interface
'  For Each oFruit In m_oCol
'    If i Then sText = sText & ", "
'    sText = sText & IIf(oFruit.Key <> vbnullstring, oFruit.Key, "No Key Entered")
'    i = i + 1
'  Next
  
  With m_oCol
    sText = sText & vbCrLf & "  Indexes: " & CStr(.BaseIndex) & " To: " & CStr(.BaseIndex + IIf(.Count, (.Count - 1), .Count))
  End With
 
ExitNow:
  sText = sText & vbCrLf & "End For...Each" & vbCrLf
 
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow
  
End Sub

Private Sub cmdNestedWalk_Click()

  Dim i&, oFruit As cFruit, oFruit2 As cFruit, sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Nested For...Each" & vbCrLf
  
  For Each oFruit In m_oCol
    sText = sText & "  Outer Loop Key: " & IIf(oFruit.Key <> vbnullstring, oFruit.Key, "No Key Entered") & vbCrLf & "     Inner Loop" & vbCrLf & "       Keys: "

    For Each oFruit2 In m_oCol
      If i Then sText = sText & ", "
      sText = sText & IIf(oFruit2.Key <> vbnullstring, oFruit2.Key, "No Key Entered")
      i = i + 1
    Next
    
    i = 0
    
    sText = sText & vbCrLf & "     End Inner Loop" & vbCrLf
  Next

  With m_oCol
    sText = sText & "  Indexes: " & CStr(.BaseIndex) & " To: " & CStr(.BaseIndex + IIf(.Count, (.Count - 1), .Count)) & vbCrLf
  End With

ExitNow:
  sText = sText & "End Outer Loop" & vbCrLf
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & "  Error!  " & Err.Description & vbCrLf
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdIterateKeys_Click()

  Dim oKey As cKey, sText$

  On Error GoTo EH
  
  sText = vbCrLf & "Iterate Keys Collection"

  If m_oCol.Keys.Count Then
    For Each oKey In m_oCol.Keys
      With oKey
        sText = sText & vbCrLf & "  Key Name: '" & .Name & "'  Associated Index: " & CStr(.AssociatedIndex)
      End With
    Next
    
  Else
    sText = sText & vbCrLf & "  No keys entered."
  End If
  
ExitNow:
  sText = sText & vbCrLf & "End Iterate Keys Collection" & vbCrLf
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdBaseIndex_Click()
  
  Dim sText$, sBaseIndex$

  On Error GoTo EH
  
  sText = vbCrLf & "Change Base Index" & vbCrLf
  
  sBaseIndex = txBaseIndex.Text
  
  If IsNumeric(sBaseIndex) Then
    With m_oCol
      sText = sText & "  From:  " & CStr(.BaseIndex)
      
      .BaseIndex = CLng(sBaseIndex)
    
      sText = sText & "  To:  " & sBaseIndex
    End With
    
  Else
    sText = sText & "  Error!  Invalid entry."

    MsgBox "Entry must be a number", vbInformation, MSGBX_TITLE
  End If
    
ExitNow:

  sText = sText & vbCrLf & "End Change Base Index" & vbCrLf
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  Select Case Err.Number
    Case AlreadyInitialized
      MsgBox Err.Description & vbCrLf & "Clear the collection and try again.", vbInformation, MSGBX_TITLE
  
    Case Else
      MsgBox Err.Description, vbInformation, MSGBX_TITLE
  End Select
  
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow
  
End Sub

Private Sub cmdMove_Click()
  
  Dim fMoveMethod&, nMoveFrom&, nMoveTo&, sMoveText$, sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Move" & vbCrLf & "  Before Move" & vbCrLf & "   Items/Keys: "
  
  sText = sText & WalkCollection()
  
  
  sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
  
  sText = sText & WalkKeysCollection()


  If IsNumeric(txMove(0).Text) Then
    fMoveMethod = 1
    nMoveFrom = CLng(txMove(0).Text)
  End If
  
  If IsNumeric(txMove(1).Text) Then
    fMoveMethod = fMoveMethod Or 2
    nMoveTo = CLng(txMove(1).Text)
  End If
  
  sText = sText & vbCrLf & "     Move : "
  
  With m_oCol
    Select Case fMoveMethod
      Case 0: .MoveItem txMove(0).Text, txMove(1).Text: sMoveText = txMove(0).Text & " To: " & txMove(1).Text
      Case 1: .MoveItem nMoveFrom, txMove(1).Text: sMoveText = CStr(nMoveFrom) & " To: " & txMove(1).Text
      Case 2: .MoveItem txMove(0).Text, nMoveTo: sMoveText = txMove(0).Text & " To: " & CStr(nMoveTo)
      Case 3: .MoveItem nMoveFrom, nMoveTo: sMoveText = CStr(nMoveFrom) & " To: " & CStr(nMoveTo)
    End Select
  End With
  
  sText = sText & sMoveText
  
  sText = sText & vbCrLf & "  After Move" & vbCrLf & "   Items/Keys: "
  
  sText = sText & WalkCollection()
  
  
  sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
  
  sText = sText & WalkKeysCollection()
  
  
ExitNow:
  sText = sText & vbCrLf & "End Move" & vbCrLf
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdRemove_Click()
  
  Dim sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Remove" & vbCrLf & "  Count Before: " & _
                            m_oCol.Count & vbCrLf & "  Before Remove" & vbCrLf & "   Items/Keys: "
  
  sText = sText & WalkCollection()
  
  
  sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
  
  sText = sText & WalkKeysCollection()
  
  
  sText = sText & vbCrLf & "     Remove : " & txRemove.Text
  
  With m_oCol
    If IsNumeric(txRemove.Text) Then
      .Remove CLng(txRemove.Text)
    Else
      .Remove txRemove.Text
    End If
  End With
    
  sText = sText & vbCrLf & "  After Remove" & vbCrLf & "   Items/Keys: "
  
  sText = sText & WalkCollection()
  
  
  sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
  
  sText = sText & WalkKeysCollection()


  sText = sText & DisplayIndexesAndCount()

ExitNow:
  sText = sText & vbCrLf & "End Remove" & vbCrLf
  
  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdChangeKey_Click()
  
  Dim sText$, sOldKey$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Change Key" & vbCrLf & "  Before Change" & vbCrLf & "   Items/Keys: "
  
  sText = sText & WalkCollection()
  
  
  sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
  
  sText = sText & WalkKeysCollection()

  
  
  sOldKey = txChangeKey(0).Text
  
  sText = sText & vbCrLf & "     Change : " & sOldKey & " To: " & txChangeKey(1).Text
  
  If IsNumeric(sOldKey) Then
    m_oCol.ChangeKeyForIndex CLng(sOldKey), txChangeKey(1).Text
  Else
    m_oCol.ChangeKey sOldKey, txChangeKey(1).Text
  End If

  sText = sText & vbCrLf & "  After Change" & vbCrLf & "   Items/Keys: "
  
  sText = sText & WalkCollection()
  
  
  sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
  
  sText = sText & WalkKeysCollection()


  sText = sText & DisplayIndexesAndCount()

ExitNow:
  sText = sText & vbCrLf & "End Change Key" & vbCrLf

  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdGetKey_Click()

  Dim sKeyText$, sText$

  On Error GoTo EH

  sText = vbCrLf & "Get Key From Index"

  sKeyText = Trim$(txGetKey.Text)
  
  If IsNumeric(sKeyText) Then
    sText = sText & vbCrLf & "  The Key For Index #" & sKeyText & "  Is: " & IIf(m_oCol.Key(CLng(sKeyText)) <> vbnullstring, m_oCol.Key(CLng(sKeyText)), "No Key Entered")
    
  Else
    sText = sText & vbCrLf & "  Error!  Invalid entry."

    MsgBox "Entry must be a number", vbInformation, MSGBX_TITLE
  End If
    
ExitNow:
  sText = sText & vbCrLf & "End Get Key From Index" & vbCrLf

  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdAssign_Click()
  
  Dim sOldKey$, sNewKey$, oFruit As cFruit, sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Assign New Item To Existing Item" & vbCrLf & "  Before Assignment" & vbCrLf & "   Items/Keys: "
  
  If (m_oCol.KeyExists(txAssign(1).Text) <> KEY_NOT_FOUND) And (txAssign(1).Text <> txAssign(0).Text) Then
    MsgBox "An item with the new key already exists in the collection.", vbInformation, MSGBX_TITLE
    
  Else
    sText = sText & WalkCollection()
  
  
    sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
    
    sText = sText & WalkKeysCollection()
    
    
    sOldKey = txAssign(0).Text
    sNewKey = txAssign(1).Text
    
    sText = sText & vbCrLf & "  Assign : " & sNewKey & " To: " & sOldKey
    
    Set oFruit = New cFruit
    
    With oFruit
      .Name = sNewKey
      .Key = sOldKey
    End With
    
    With m_oCol
      Set .Item(sOldKey) = oFruit
  
      ' the following line would change the key to match the new item name
      '.ChangeKey sOldKey, sNewKey
    End With
    
    Set oFruit = Nothing
  
    sText = sText & vbCrLf & "  After Assignment" & vbCrLf & "   Items/Keys: "
    
    sText = sText & WalkCollection()
  
  
    sText = sText & vbCrLf & "   Keys Collection (Name/Assoc Index: "
    
    sText = sText & WalkKeysCollection()
  
  
    sText = sText & DisplayIndexesAndCount()
  End If


ExitNow:
  sText = sText & vbCrLf & "End Assign New Name" & vbCrLf

  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)
  
  Exit Sub
  
EH:
  Select Case Err.Number
    Case 438
      MsgBox Err.Description & vbCrLf & "Check the 'Allow Item Assignments' check box and try again.", vbInformation, MSGBX_TITLE
  
    Case Else
      MsgBox Err.Description, vbInformation, MSGBX_TITLE
  End Select
  
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow

End Sub

Private Sub cmdKeyExists_Click()

  Dim nIndex&, sKey$, sText$
  
  On Error GoTo EH
  
  sText = vbCrLf & "Key Exists" & vbCrLf
  
  sKey = Trim$(txKeyExists.Text)
  
  If sKey <> vbnullstring Then
    nIndex = m_oCol.KeyExists(sKey)
    
    If nIndex <> KEY_NOT_FOUND Then
      sText = sText & "  Key: '" & sKey & "'  Found At Index #" & CStr(nIndex)
    Else
      sText = sText & "  Key: '" & sKey & "'  Not Found"
    End If
    
  Else
    sText = sText & "  Error!  No key entered."
    
    MsgBox "No key entered.", vbInformation, MSGBX_TITLE
  End If
  
ExitNow:
  sText = sText & vbCrLf & "End Key Exists" & vbCrLf

  txDisplay.Text = txDisplay.Text & sText
  txDisplay.SelStart = Len(txDisplay.Text)

  Exit Sub
  
EH:
  MsgBox Err.Description, vbInformation, MSGBX_TITLE
  sText = sText & vbCrLf & "  Error!  " & Err.Description
  Err.Clear
  Resume ExitNow
  
End Sub

Private Sub Form_Load()
  ' instanciate the collection
  Set m_oCol = New CSuperCollection
  
  txDisplay.FontSize = 6
  
  txBaseIndex.Text = CStr(m_oCol.BaseIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' distroy the collection
  Set m_oCol = Nothing
End Sub

Private Function WalkCollection() As String

  Dim i&, sText$, oFruit As cFruit

  For Each oFruit In m_oCol
    If i Then sText = sText & ", "
    sText = sText & IIf(oFruit.Key <> vbnullstring, oFruit.Name & "/" & oFruit.Key, "No Key Entered")
    i = i + 1
  Next

   WalkCollection = sText
End Function

Private Function WalkKeysCollection() As String

  Dim i&, sText$, oKey As cKey

  For Each oKey In m_oCol.Keys
    If i Then sText = sText & ", "
    sText = sText & IIf(oKey.Name <> vbnullstring, oKey.Name & "/" & oKey.AssociatedIndex, "No Key Entered")
    i = i + 1
  Next

   WalkKeysCollection = sText
End Function

Private Function DisplayIndexesAndCount() As String

  Dim sText$
  
  With m_oCol
    sText = vbCrLf & "  Indexes: " & CStr(.BaseIndex) & " To: " & CStr(.BaseIndex + IIf(.Count, (.Count - 1), .Count))
    sText = sText & vbCrLf & "  Count After: " & .Count
  End With

  DisplayIndexesAndCount = sText
  
End Function

Private Sub txBaseIndex_KeyPress(KeyAscii As Integer)
  ' since the textbox is right aligned, prevent the enter key from adding a CrLf
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
  End If
End Sub

Private Sub txGetKey_KeyPress(KeyAscii As Integer)
  ' since the textbox is right aligned, prevent the enter key from adding a CrLf
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
  End If
End Sub

