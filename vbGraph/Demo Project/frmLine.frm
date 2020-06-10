VERSION 5.00
Object = "*\A..\vbGraph.vbp"
Begin VB.Form frmLine 
   Caption         =   "Lines and Bars"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin vbGraph.Graph Graph1 
      Height          =   1395
      Left            =   1650
      TabIndex        =   5
      Top             =   810
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   2461
      State           =   "frmLine.frx":0000
   End
   Begin VB.Timer tmrChange 
      Left            =   180
      Top             =   0
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   2670
      Width           =   7845
      Begin VB.HScrollBar HScroll1 
         Height          =   315
         LargeChange     =   50
         Left            =   1890
         Max             =   500
         TabIndex        =   4
         Top             =   300
         Width           =   3675
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
      Begin VB.CheckBox chkVisible 
         Caption         =   "Dataset 2 Visible"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   330
         Width           =   1515
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Dataset 3 Visible"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmLine"
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
    HScroll1.Value = 250
    chkVisible(0).Value = vbChecked
End Sub

Private Sub SetupGraph()
    With Graph1
        .BarWidth = 0.8
        .FixedPoints = 0
        .MaxValue = 100
        .MinValue = 0
        .xGridInc = 1
        .YGridInc = 10
        .ShowAxis = False
        .ShowGrid = True
        .FadeIn = False
        .BackColor = RGB(255, 255, 255)
        .GridColor = RGB(200, 200, 200)
    End With
End Sub

Private Sub SetupDatasets()
Dim objDataset  As Dataset
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Showpoints = True
        .ShowLines = True
        .ShowBars = False
        .Showcaps = False
        .LineColor = RGB(0, 0, 255)
        .PointColor = RGB(150, 150, 255)
    End With
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Visible = False
        .Showpoints = True
        .ShowLines = True
        .ShowBars = False
        .Showcaps = False
        .LineColor = RGB(255, 0, 0)
        .PointColor = RGB(255, 150, 150)
    End With
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Visible = False
        .Showpoints = True
        .ShowLines = True
        .ShowBars = False
        .Showcaps = False
        .LineColor = RGB(0, 255, 0)
        .PointColor = RGB(150, 255, 150)
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
            lngValue = (Rnd * (.MaxValue - .MinValue)) + .MinValue
            .Datasets.Item(lngIndex).Points.Add lngValue
        Next lngIndex
        .Redraw = True
    End With
End Sub

