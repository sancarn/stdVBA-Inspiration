VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
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
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSelOnly 
      Caption         =   "Search in Selected Text Only"
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1185
      Width           =   3180
   End
   Begin VB.CommandButton cmdGone 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2895
      TabIndex        =   8
      Top             =   1530
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   420
      Left            =   1185
      TabIndex        =   7
      Top             =   1530
      Width           =   1695
   End
   Begin VB.OptionButton optDir 
      Caption         =   "UP"
      Height          =   315
      Index           =   2
      Left            =   3690
      TabIndex        =   5
      Top             =   870
      Width           =   690
   End
   Begin VB.OptionButton optDir 
      Caption         =   "Down"
      Height          =   315
      Index           =   1
      Left            =   2865
      TabIndex        =   4
      Top             =   870
      Width           =   885
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "Match Case"
      Height          =   315
      Index           =   1
      Left            =   3105
      TabIndex        =   3
      Top             =   555
      Width           =   1290
   End
   Begin VB.CheckBox chkOpt 
      Caption         =   "Whole Word Only"
      Height          =   315
      Index           =   0
      Left            =   1185
      TabIndex        =   1
      Top             =   555
      Width           =   1935
   End
   Begin VB.ComboBox cboCriteria 
      Height          =   330
      Left            =   1185
      TabIndex        =   0
      Top             =   150
      Width           =   3240
   End
   Begin VB.OptionButton optDir 
      Caption         =   "All"
      Height          =   315
      Index           =   0
      Left            =   2175
      TabIndex        =   2
      Top             =   870
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Direction:"
      Height          =   300
      Index           =   1
      Left            =   1230
      TabIndex        =   10
      Top             =   915
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   300
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   210
      Width           =   1320
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDpiPmAssistant       ' required to receive DPI changes
Dim cDpiAssist As clsDpiPmAssist ' required to react to DPI changes

Private Enum SearchStateEnum
    ssBkwdSearching = 1     ' search up vs down
    ssResetAnchor = 2       ' wrap position changed
    ssCriteriaChg = 4       ' criteria changed/changing
    ssSelectOnly = 8        ' search selected text only
End Enum

Dim WithEvents m_RTB As RichTextBox
Attribute m_RTB.VB_VarHelpID = -1
Dim m_BkwdIdx As Collection
Dim m_Options As FindConstants
Dim m_Flags As SearchStateEnum
Dim m_RangeA&   ' start for fwd searches else m_BkwdIdx index
Dim m_RangeZ&   ' end for fwd searches (-1 = EOF); n/a for bkwd searches
Dim m_Anchor&   ' location to signal end of searches before/after wrapping

Friend Sub SetSource(oRTB As RichTextBox)

    ' called to display this form, set the RTB, and auto-fill criteria as needed

    If m_RTB Is Nothing Then Set m_RTB = oRTB
    If m_RTB.SelLength = 0 Then
        chkSelOnly.Enabled = False
        cmdGo.Enabled = (LenB(cboCriteria.Text) <> 0)
    Else
        Dim lIndex&, sValue$
        sValue = Trim$(m_RTB.SelText)
        lIndex = InStr(sValue, vbNewLine)
        If lIndex <> 0 Then sValue = RTrim$(Left$(sValue, lIndex - 1))
        cboCriteria.Text = sValue
        chkSelOnly.Enabled = True
    End If
    m_Flags = m_Flags Or ssResetAnchor
    Me.Show vbModeless, m_RTB.Parent
    
End Sub

Private Sub pvFindNext()

    ' perform a search attempt
    ' designed to mimic VB's "find" behavior

    Dim lIndex&
    If (m_Flags And ssResetAnchor) = ssResetAnchor Then
        m_Anchor = m_RTB.SelStart           ' reset anchor
        m_RangeA = m_Anchor                 ' start of searches
        m_RangeZ = -1                       ' EOF
        If (m_Flags And ssBkwdSearching) = ssBkwdSearching Then
            m_RangeA = pvGetOffsets(False)  ' get matches from BOF to anchor
        Else
            If (m_Flags And ssSelectOnly) = 0 Then
                m_RangeA = m_RangeA + m_RTB.SelLength
            Else
                m_RangeZ = m_Anchor + m_RTB.SelLength
            End If
            Set m_BkwdIdx = Nothing
        End If
        m_Flags = m_Flags Xor ssResetAnchor
    End If
    
    If (m_Flags And ssBkwdSearching) = 0 Then ' forward searching
        lIndex = m_RTB.Find(cboCriteria.Text, m_RangeA, m_RangeZ, m_Options)
    Else                                    ' backward searching
        m_RangeA = m_RangeA - 1             ' get previous search result
        If m_RangeA < 1 Then
            lIndex = -1
        Else
            lIndex = m_BkwdIdx.Item(m_RangeA)
            If m_RangeZ = m_Anchor And lIndex <= m_Anchor Then
                lIndex = -1                 ' exceeded wrap
            Else
                m_RTB.SelStart = lIndex     ' highlight result
                m_RTB.SelLength = Len(cboCriteria.Text)
            End If
        End If
    End If
    
    If lIndex = -1 Then                     ' no match, wrap if not already done
        If m_RangeZ = m_Anchor Then         ' already wrapped, done
            m_Flags = m_Flags Or ssResetAnchor
        ElseIf (m_Flags And ssSelectOnly) = ssSelectOnly Then
            ' no wrapping if searching selected text only
            m_Flags = m_Flags Or ssResetAnchor
        ElseIf optDir(1).Value = True Then  ' search down only
            If cDpiAssist.MsgBox("End of search scope has been reached. " & _
                "Do you want to continue from the beginning?", vbYesNo + vbQuestion, "Continue") = vbNo Then
                m_Flags = m_Flags Or ssResetAnchor
            End If
        ElseIf (m_Flags And ssBkwdSearching) = ssBkwdSearching Then ' search up only
            If cDpiAssist.MsgBox("End of search scope has been reached. " & _
                "Do you want to continue from the end?", vbYesNo + vbQuestion, "Continue") = vbNo Then
                m_Flags = m_Flags Or ssResetAnchor
            End If
        End If
        If (m_Flags And ssResetAnchor) = ssResetAnchor Then ' done, no more search matches
            cDpiAssist.MsgBox "The specified region has been searched.", vbOKOnly + vbInformation, "Done"
            If (m_Flags And ssSelectOnly) = ssSelectOnly Then
                ' kept last match highlighted until after msgbox closed
                m_Flags = m_Flags Xor ssSelectOnly: m_RTB.SelLength = 0
            End If
        Else
            If (m_Flags And ssBkwdSearching) = ssBkwdSearching Then
                m_RangeA = pvGetOffsets(True) ' get matches from anchor to EOF
            Else
                m_RangeA = 0                ' search from BOF
            End If
            m_RangeZ = m_Anchor             ' wrap and search again
            pvFindNext
        End If
    ElseIf (m_Flags And ssBkwdSearching) = 0 Then ' adjust next start position when fwd searching
        m_RangeA = lIndex + m_RTB.SelLength
    End If

End Sub

Private Function pvGetOffsets(bWrap As Boolean) As Long

    ' RTB can't do backward searches. This is just one way of handling it.
    ' Do a forward search w/o highlighting. As each match is found, cache
    ' its location. Use that cache in place of searches.

    Dim lIndex&, lRange&, lReturn&
    
    ' always return index+1 because caller subtracts one on return
    If bWrap = False Then
        ' find matches from top of document to anchor
        Set m_BkwdIdx = New Collection
        If (m_Flags And ssSelectOnly) = 0 Then
            lIndex = -1: lRange = m_Anchor
        Else
            lIndex = m_Anchor - 1: lRange = m_Anchor + m_RTB.SelLength
            bWrap = True: lReturn = -1
        End If
    Else
        ' find matches from anchor to end of document
        lIndex = m_Anchor - 1: lRange = -1: lReturn = -1
    End If
        
    Do
        lIndex = m_RTB.Find(cboCriteria.Text, lIndex + 1, lRange, m_Options Or rtfNoHighlight)
        If lIndex = -1 Then Exit Do
        m_BkwdIdx.Add lIndex
        If lReturn = 0 Then
            If lIndex >= m_Anchor Then lReturn = m_BkwdIdx.Count
        End If
    Loop
    If lReturn < 1 Then lReturn = m_BkwdIdx.Count + 1
    pvGetOffsets = lReturn
    
End Function

Private Sub cmdGo_Click()

    ' prep for a search attempt
    
    Dim lValue&
    Const CB_FINDSTRINGEXACT = &H158&
    
    If (m_Flags And ssCriteriaChg) = ssCriteriaChg Then
        With cboCriteria
            lValue = SendMessageA(.hWnd, CB_FINDSTRINGEXACT, -1, ByVal .Text)
            ' keep recent selection at top of list, keep case-sensitivity
            .AddItem .Text, 0
            If lValue <> -1 Then .RemoveItem lValue + 1
            If .ListCount = 21 Then .RemoveItem 20
            .ListIndex = 0
        End With
        m_Flags = m_Flags Xor ssCriteriaChg
    End If
    
    ' restrict to selected text if applicable
    If chkSelOnly.Value = vbChecked And chkSelOnly.Enabled = True Then
        If (m_Flags And ssSelectOnly) = 0 Then
            m_Flags = (m_Flags And ssBkwdSearching) Or ssResetAnchor Or ssSelectOnly
        End If
    ElseIf (m_Flags And ssSelectOnly) = ssSelectOnly Then
        ' remove selected text option if applicable
        m_Flags = m_Flags Xor ssSelectOnly Or ssResetAnchor
    End If
    
    ' see if options changed
    If chkOpt(0).Value = vbChecked Then lValue = rtfWholeWord Else lValue = 0
    If chkOpt(1).Value = vbChecked Then lValue = lValue Or rtfMatchCase
    If lValue <> m_Options Then
        m_Options = lValue: Set m_BkwdIdx = Nothing
        m_Flags = m_Flags Or ssResetAnchor
    End If
    pvFindNext
    
End Sub

Private Sub cmdGone_Click()
    Me.Visible = False
    m_RTB.Parent.SetFocus
    Set m_RTB = Nothing     ' stop receiving events
End Sub

Private Sub cboCriteria_Change()
    m_Flags = (m_Flags And 1) Or ssCriteriaChg Or ssResetAnchor
    cmdGo.Enabled = (LenB(cboCriteria.Text) <> 0)
End Sub
Private Sub cboCriteria_Click()
    m_Flags = (m_Flags And 1) Or ssCriteriaChg Or ssResetAnchor
    cmdGo.Enabled = True
End Sub

Private Sub Form_Load()
    cDpiAssist.Activate Me  ' do not move this line to Form_Activate
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode And UnloadMode <> vbFormOwner Then
        Cancel = True
        cmdGone.Value = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_RTB = Nothing
    Set m_BkwdIdx = Nothing
    Set cDpiAssist = Nothing
End Sub

Private Sub m_RTB_Click()
    m_Flags = m_Flags Or ssResetAnchor ' clicked off this form into RTB
End Sub

Private Sub m_RTB_SelChange()
    chkSelOnly.Enabled = (m_RTB.SelLength <> 0)
End Sub

Private Sub optDir_Click(Index As Integer)
    If optDir(2).Value = True Then          ' changing to backward searches
        m_Flags = (m_Flags And ssCriteriaChg) Or ssBkwdSearching
    ElseIf (m_Flags And ssBkwdSearching) = ssBkwdSearching Then
        m_Flags = m_Flags And ssCriteriaChg ' changing from backward searches
    End If
    m_Flags = m_Flags Or ssResetAnchor
End Sub

Private Sub IDpiPmAssistant_Attach(oDpiPmAssist As clsDpiPmAssist)
    '/// this is required else no DPI scaling will occur
    Set cDpiAssist = oDpiPmAssist
End Sub

Private Function IDpiPmAssistant_DpiScalingCycle(ByVal Reason As DpiScaleReason, _
                                ByVal Action As DpiScaleCycleEnum, _
                                ByVal OldDPI As Long, ByVal NewDPI As Long, _
                                ByRef userParams As Variant) As Long
    '/// use this event to prep for rescaling and handle any post-scaling actions you need
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
End Sub

Private Function IDpiPmAssistant_Subclasser(EventValue As Long, ByVal BeforeHwnd As Boolean, _
                        ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '/// respond to this if you have called cDpiAssist.SubclassHwnd
End Function


