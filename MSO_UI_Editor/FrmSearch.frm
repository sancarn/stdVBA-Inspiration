VERSION 5.00
Object = "{07C05129-C2E5-483C-8237-8636C3F11E4E}#1.0#0"; "VBCCR13.OCX"
Begin VB.Form FrmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5745
   Icon            =   "FrmSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   5535
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   5535
      Begin VB.CommandButton CmdReplaceAll 
         Caption         =   "Replace All"
         Height          =   350
         Left            =   4200
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton CmdReplace 
         Caption         =   "Replace"
         Height          =   350
         Left            =   4200
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VBCCR13.CheckBoxW ChkMatchCase 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "FrmSearch.frx":058A
         Transparent     =   -1  'True
      End
      Begin VB.CommandButton CmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   350
         Left            =   4200
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton CmdFindNext 
         Caption         =   "Find Next"
         Default         =   -1  'True
         Height          =   350
         Left            =   4200
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VBCCR13.TextBoxW TxtSearch 
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VBCCR13.TextBoxW TxtReplace 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace with:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find wath:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   735
      End
   End
   Begin VBCCR13.TabStrip TabStrip1 
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InitTabs        =   "FrmSearch.frx":05BE
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lCurPos As Long

Private Sub Form_Load()
    TabStrip1_TabClick TabStrip1.Tabs(1)
End Sub

Private Sub CmdFindNext_Click()
    Dim Str As String
    
    Str = Replace(Form1.RTB1.Text, vbCrLf, Chr(10))
    If lCurPos < 1 Then lCurPos = 1
    
    lCurPos = InStr(lCurPos + 1, Str, TxtSearch.Text, IIf(ChkMatchCase.Value, vbTextCompare, vbBinaryCompare))
    
    If lCurPos Then
        Form1.RtbSel lCurPos - 1, lCurPos + Len(TxtSearch.Text) - 1
    Else
        Form1.RtbSel 0, 0
        Beep
    End If
    
End Sub

Private Sub CmdReplace_Click()
    If StrComp(Form1.RTB1.SelText, TxtSearch.Text, IIf(ChkMatchCase.Value, vbTextCompare, vbBinaryCompare)) = 0 Then
        Form1.RTB1.SelText = TxtReplace.Text
    End If
    CmdFindNext_Click
End Sub

Private Sub CmdReplaceAll_Click()
    Dim lCount As Long, lCurPos As Long
    Dim Str As String
    
    Str = Replace(Form1.RTB1.Text, vbCrLf, Chr(10))
    lCurPos = 1
    Form1.bOn = True
    Do While lCurPos <> 0
        Str = Replace(Form1.RTB1.Text, vbCrLf, Chr(10))
        
        lCurPos = InStr(lCurPos + 1, Str, TxtSearch.Text, IIf(ChkMatchCase.Value, vbTextCompare, vbBinaryCompare))
    
        If lCurPos Then
            Form1.RtbSel lCurPos - 1, lCurPos + Len(TxtSearch.Text) - 1
            Form1.RTB1.SelText = TxtReplace.Text
            lCount = lCount + 1
        Else
            If lCount = 0 Then
                MsgBox "Not found the text search."
            Else
                MsgBox lCount & " substitutions were made."
            End If
            Exit Do
        End If
    Loop
    Form1.bOn = False
    
    
End Sub

Private Sub TabStrip1_TabClick(ByVal TabItem As VBCCR13.TbsTab)
    If TabItem.Index = 1 Then
        Me.Caption = "Find"
        Picture1.Height = 1095
        TabStrip1.Height = 1565
        Me.Height = 2000
        CmdCancel.Top = 650
        ChkMatchCase.Top = 700
    Else
        Me.Caption = "Find and Replace"
        Picture1.Height = 1935
        TabStrip1.Height = 2415
        Me.Height = 2840
        CmdCancel.Top = 1560
        ChkMatchCase.Top = 1560
    End If
     
    Label2.Visible = TabItem.Index = 2
    TxtReplace.Visible = TabItem.Index = 2
    CmdReplace.Visible = TabItem.Index = 2
    CmdReplaceAll.Visible = TabItem.Index = 2
    Picture1.AutoRedraw = True
    TabStrip1.DrawBackground Picture1.HWnd, Picture1.hdc
    Picture1.Refresh
    ChkMatchCase.Refresh
    If Me.Visible Then TxtSearch.SetFocus
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

