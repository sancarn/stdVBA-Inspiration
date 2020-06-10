VERSION 5.00
Object = "*\A..\vbGraph.vbp"
Begin VB.Form frmProgress 
   Caption         =   "Progress"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   9000
   StartUpPosition =   1  'CenterOwner
   Begin vbGraph.Graph Graph1 
      Height          =   225
      Left            =   1890
      TabIndex        =   3
      Top             =   90
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   397
      State           =   "frmProgress.frx":0000
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
      ScaleWidth      =   8940
      TabIndex        =   0
      Top             =   330
      Width           =   9000
      Begin VB.CheckBox chkStyle 
         Caption         =   "Advanced Style"
         Height          =   225
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
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
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkVisible_Click(Index As Integer)

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
End Sub

Private Sub SetupGraph()
    With Graph1
        .BarWidth = 1
        .FixedPoints = 80
        .MaxValue = 100
        .MinValue = 0
        .xGridInc = 100
        .YGridInc = 1
        .ShowAxis = False
        .ShowGrid = False
        .FadeIn = False
        .BackColor = RGB(255, 255, 255)
        .GridColor = RGB(200, 200, 200)
    End With
End Sub

Private Sub SetupDatasets()
Dim objDataset  As Dataset
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Showpoints = False
        .ShowBars = True
        .ShowLines = False
        .Showcaps = False
        .BarColor = &HDEA68D
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
Dim lngClear    As Long
    With Graph1
        .Redraw = False
        If .Datasets.Item(1).Points.Count = .FixedPoints Then
            .Datasets.Item(1).Points.Clear
        End If
        .Datasets.Item(1).Points.Add .MaxValue
        If chkStyle.Value = vbChecked Then
            lngClear = .Datasets.Item(1).Points.Count - 20
            If lngClear < 1 Then
                lngClear = .FixedPoints + lngClear
            End If
            If lngClear <= .Datasets.Item(1).Points.Count Then
                .Datasets.Item(1).Points.Item(lngClear).Value = 0
            End If
        End If
        .Redraw = True
    End With
End Sub


