VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmValidationHelp 
   Caption         =   "Descriptions of Performed Validations"
   ClientHeight    =   6015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtb 
      Height          =   5910
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   10425
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmValidationHelp.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Save to File"
      Index           =   0
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   2
      Begin VB.Menu mnuEdit 
         Caption         =   "Wrap to &Window"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&No Wrap"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Find"
         Index           =   3
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help Topics"
      Visible         =   0   'False
      Begin VB.Menu mnuTopics 
         Caption         =   "Duplicate Declartions Check"
         Index           =   0
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Duplicate Literals Check"
         Index           =   1
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Empty Code Check"
         Index           =   2
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Malicious Code Check"
         Index           =   3
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Option Explicit Check"
         Index           =   4
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "ReDim Check"
         Index           =   5
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Stop/End Check"
         Index           =   6
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Variant Function Check"
         Index           =   7
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "VarType Check"
         Index           =   8
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "WithEvents Check"
         Index           =   9
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "Zombie Check"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmValidationHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDpiPmAssistant       ' required to receive DPI changes
Dim cDpiAssist As clsDpiPmAssist ' required to react to DPI changes
Attribute cDpiAssist.VB_VarHelpID = -1

Const EM_SETZOOM As Long = (WM_USER + 225)
Dim m_FindFrm As frmFind

Public Sub ShowAdjustedForDPI()     ' called in place of Show

    ' when replacing rtb.TextRTF, not rtb.Text, need to re-zoom if applicable
    ' this is called only when setting rtb.TextRTF outside this form
    ' zooming only applies during DPI changes.
    If cDpiAssist.DpiForForm <> cDpiAssist.DpiForSystem Then
        SendMessageA rtb.hWnd, EM_SETZOOM, cDpiAssist.DpiForForm, ByVal cDpiAssist.DpiForSystem
    End If
    Me.Show

End Sub

Private Sub Form_Load()
    Me.BackColor = rtb.BackColor
    cDpiAssist.Activate Me ' do not move this line to Form_Activate
    Me.KeyPreview = True
End Sub

Private Sub Form_Resize()
    If cDpiAssist.IsScalingCycleActive = False Then pvDoResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cDpiAssist = Nothing
    If Not m_FindFrm Is Nothing Then
        Unload m_FindFrm
        Set m_FindFrm = Nothing
    End If
End Sub

Public Sub SetWordWrap(bWrap As Boolean)
    mnuEdit_Click Abs(bWrap) Xor 1
End Sub

Private Sub pvDoResize()
    If Me.WindowState <> vbMinimized Then
        rtb.Move rtb.Left, rtb.Top, Me.ScaleWidth - rtb.Left * 2, Me.ScaleHeight - rtb.Top * 2
        ' external ocx dimensions should be validated when DPI aware...
        cDpiAssist.SyncOcxToParent rtb
    End If
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Const EM_SETTARGETDEVICE = (WM_USER + 72)
    If Index = 3 Then
        If m_FindFrm Is Nothing Then
            Set m_FindFrm = New frmFind
            Load m_FindFrm
            m_FindFrm.Move Me.Width - m_FindFrm.Width + Me.Left, Me.Top
        End If
        m_FindFrm.SetSource Me.rtb
    Else
        If mnuEdit(Index).Checked = True Then Exit Sub
        If Index = 0 Then
            SendMessageA rtb.hWnd, EM_SETTARGETDEVICE, 0, ByVal 0&
        Else
            SendMessageA rtb.hWnd, EM_SETTARGETDEVICE, 0, ByVal 1&
        End If
        mnuEdit(Index).Checked = True
        mnuEdit(Index Xor 1).Checked = False
    End If
    
End Sub

Private Sub mnuMain_Click(Index As Integer)
    
    If Index <> 0 Then Exit Sub
    
    Dim s$
    s = Replace$(Clipboard.GetText, vbCrLf, " ")
    Clipboard.Clear
    rtb.SelText = s & " Length: " & Len(s)
    Clipboard.SetText s
    Exit Sub
    
    
    Dim cBrowser As CmnDialogEx, hFile As Long
    Dim n As Long, lRead As Long, aData() As Byte
    
    Set cBrowser = New CmnDialogEx
    cBrowser.Filter = "RTF Files|*.rtf|All Files|*.*"
    cBrowser.DefaultExt = "rtf"
    If Len(Me.Tag) > 1 Then cBrowser.FileName = Me.Tag
    If cBrowser.ShowSave(Me.hWnd) = True Then
        hFile = modMain.GetFileHandle(cBrowser.FileName, True)
        If hFile = 0 Or hFile = -1 Then
            cDpiAssist.MsgBox "Access Failure", vbExclamation + vbOKOnly, "No Action"
        Else
            If cBrowser.FilterIndex = 1 Then
                aData() = StrConv(rtb.TextRTF, vbFromUnicode)
            Else
                aData() = StrConv(rtb.Text, vbFromUnicode)
            End If
            n = UBound(aData) + 1
            modMain.WriteFile hFile, aData(0), n, lRead
            modMain.CloseHandle hFile
            If n = lRead Then
                cDpiAssist.MsgBox "Saved to disk", vbInformation + vbOKOnly, "Done"
            Else
                cDpiAssist.MsgBox "Failure. Cound not save to disk", vbExclamation + vbOKOnly, "Error"
            End If
        End If
    End If
    Set cBrowser = Nothing
    
End Sub

Private Sub mnuTopics_Click(Index As Integer)

    Dim sValue$
    Select Case Index
    Case 0: sValue = "Duplicate Declarations Name check (Optional)"
    Case 1: sValue = "Duplicate Literals check (Optional)"
    Case 2: sValue = "Empty Code check"
    Case 3: sValue = "Safety Warnings Checks (Optional)"
    Case 4: sValue = "Option Explicit check"
    Case 5: sValue = "ReDim Statement without Declared Variable check (Optional)"
    Case 6: sValue = "Stop and End check"
    Case 7: sValue = "Variant vs. String Function check (Optional)"
    Case 8: sValue = "VarType check"
    Case 9: sValue = "WithEvents Declaration without Events check"
    Case 10: sValue = "Zombie check"
    End Select
    If rtb.Find(sValue, 0, , rtfWholeWord) = -1 Then
        cDpiAssist.MsgBox "Help topic not found", vbExclamation + vbOKOnly, "Oops"
    End If

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
    If Action = dpiAsst_EndCycleHost Then
      '/// center form if wanted, Me.WindowState = vbNormal & Reason=dpiAsst_InitialLoad
        pvDoResize
    End If
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
    If Action = dpiAsst_EndEvent Then
        SendMessageA rtb.hWnd, EM_SETZOOM, cDpiAssist.DpiForForm, ByVal cDpiAssist.DpiForSystem
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
