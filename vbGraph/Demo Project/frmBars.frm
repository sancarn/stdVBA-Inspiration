VERSION 5.00
Object = "*\A..\vbGraph.vbp"
Begin VB.Form frmBars 
   Caption         =   "Bars"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin vbGraph.Graph Graph1 
      Height          =   1485
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   2619
      State           =   "frmBars.frx":0000
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   7110
      TabIndex        =   0
      Top             =   3165
      Width           =   7170
      Begin VB.HScrollBar HScroll1 
         Height          =   315
         LargeChange     =   50
         Left            =   1890
         Max             =   500
         TabIndex        =   1
         Top             =   300
         Width           =   3675
      End
   End
   Begin VB.Timer tmrChange 
      Left            =   180
      Top             =   0
   End
End
Attribute VB_Name = "frmBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
        .ShowGrid = False
        .FadeIn = False
        .BackColor = RGB(121, 145, 200)
        .GridColor = RGB(110, 135, 190)
    End With
End Sub

Private Sub SetupDatasets()
Dim objDataset  As Dataset
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Showpoints = False
        .ShowLines = False
        .ShowBars = True
        .Showcaps = False
        .LineColor = RGB(255, 255, 255)
        .PointColor = RGB(255, 255, 255)
        .BarColor = &HECD2BF
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
        lngValue = (Rnd * (.MaxValue - .MinValue)) + .MinValue
        .Datasets.Item(1).Points.Add lngValue
        .Redraw = True
    End With
End Sub


