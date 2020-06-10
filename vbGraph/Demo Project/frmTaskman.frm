VERSION 5.00
Object = "*\A..\vbGraph.vbp"
Begin VB.Form frmTaskman 
   Caption         =   "Taskman"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin vbGraph.Graph Graph1 
      Height          =   1335
      Left            =   1380
      TabIndex        =   5
      Top             =   180
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   2355
      State           =   "frmTaskman.frx":0000
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   7740
      TabIndex        =   1
      Top             =   1695
      Width           =   7800
      Begin VB.CheckBox chkVisible 
         Caption         =   "Dataset 3 Visible"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   0
         Top             =   600
         Width           =   1515
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Dataset 2 Visible"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   1515
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Dataset 1 Visible"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   60
         Width           =   1515
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   315
         LargeChange     =   50
         Left            =   1890
         Max             =   500
         TabIndex        =   2
         Top             =   300
         Width           =   3675
      End
   End
   Begin VB.Timer tmrChange 
      Left            =   180
      Top             =   0
   End
End
Attribute VB_Name = "frmTaskman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkVisible_Click(Index As Integer)
    Graph1.Datasets.Item(Index + 1).Visible = (chkVisible.Item(Index).Value = vbChecked)
End Sub

Private Sub Form_Load()
    Initialise
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        Graph1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picHolder.Height
    End If
End Sub

Private Sub Initialise()
    SetupGraph
    SetupDatasets
    HScroll1.Value = 400
    chkVisible(0).Value = vbChecked
End Sub

Private Sub SetupGraph()
    With Graph1
        .FadeIn = False
        .MaxValue = 100
        .MinValue = 0
        .YGridInc = 20
        .xGridInc = 1
        .BackColor = RGB(0, 0, 0)
        .GridColor = &H8000&
        .FixedPoints = 80
        .BarWidth = 0.8
        .ShowAxis = False
    End With
End Sub

Private Sub SetupDatasets()
Dim objDataset  As Dataset
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Showpoints = False
        .ShowBars = False
        .ShowLines = True
        .Showcaps = False
        .LineColor = vbGreen
    End With
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Visible = False
        .Showpoints = False
        .ShowBars = False
        .ShowLines = True
        .Showcaps = False
        .LineColor = vbWhite
    End With
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Visible = False
        .Showpoints = False
        .ShowBars = False
        .ShowLines = True
        .Showcaps = False
        .LineColor = vbCyan
    End With
End Sub

Private Sub HScroll1_Change()
    tmrChange.Interval = 501 - HScroll1.Value
    tmrChange.Enabled = Not (tmrChange.Interval = 501)
End Sub

Private Sub tmrChange_Timer()
    ChangeGrid
End Sub

Private Sub ChangeGrid()
Dim lngValue    As Long
Dim lngIndex    As Long
    With Graph1
        .Redraw = False
        For lngIndex = 1 To .Datasets.Count
            If lngIndex = 3 Then
                lngValue = 70 + (Rnd * 10) - 5
            ElseIf lngIndex = 2 Then
                lngValue = 50 + (Rnd * 50) - 25
            Else
                lngValue = (Rnd * (.MaxValue - .MinValue)) + .MinValue
            End If
            .Datasets.Item(lngIndex).Points.Add lngValue
        Next lngIndex
        .Redraw = True
    End With
End Sub

