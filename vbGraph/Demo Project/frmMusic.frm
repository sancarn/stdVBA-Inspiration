VERSION 5.00
Object = "*\A..\vbGraph.vbp"
Begin VB.Form frmMusic 
   Caption         =   "Music"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin vbGraph.Graph Graph1 
      Height          =   3255
      Left            =   1410
      TabIndex        =   2
      Top             =   600
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5741
      State           =   "frmMusic.frx":0000
   End
   Begin VB.Timer tmrChange 
      Left            =   180
      Top             =   150
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      Height          =   915
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   4620
      Width           =   6150
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
Attribute VB_Name = "frmMusic"
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
    HScroll1.Value = 450
End Sub

Private Sub SetupGraph()
    With Graph1
        .FixedPoints = 40
        .ShowAxis = False
        .ShowGrid = False
        .FadeIn = False
        .MaxValue = 200
        .MinValue = 0
        .YGridInc = 20
        .xGridInc = 1
        .BarWidth = 1
        .BackColor = &H404040
        .GridColor = &H808080
    End With
End Sub

Private Sub SetupDatasets()
Dim objDataset  As Dataset
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Showpoints = False
        .ShowBars = True
        .ShowLines = False
        .Showcaps = True
        .BarColor = &HFF8080
        .CapColor = &H808080
    End With
    Set objDataset = Graph1.Datasets.Add
    With objDataset
        .Showpoints = False
        .ShowBars = False
        .ShowLines = False
        .Showcaps = True
        .CapColor = vbWhite
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
Dim blnFirst    As Boolean
Dim objPoint    As Point
Dim objCap      As Point
Dim lngValue    As Long
Dim lngCap      As Long
Dim lngIndex    As Long
    With Graph1
        .Redraw = False
        blnFirst = (.Datasets.Item(1).Points.Count = 0)
        For lngIndex = 1 To Graph1.FixedPoints
            If blnFirst Then
                lngValue = (Rnd * 80) + 50
                lngValue = .MaxValue * 0.8
                lngCap = lngValue
                Set objPoint = .Datasets.Item(1).Points.Add(lngValue)
                Set objCap = .Datasets.Item(2).Points.Add(lngCap)
            Else
                Set objPoint = .Datasets.Item(1).Points.Item(lngIndex)
                Set objCap = .Datasets.Item(2).Points.Item(lngIndex)
                lngValue = objPoint.Value
                lngCap = objCap.Value
                If lngValue < 30 Then
                    lngValue = .MaxValue * 0.8
                    lngCap = lngValue
                Else
                    lngCap = lngCap - 3
                    lngValue = lngValue - (Rnd * 10)
                    If lngValue >= lngCap Then
                        lngCap = lngValue
                    End If
                End If
            End If
            objPoint.Value = lngValue
            objCap.Value = lngCap
            Set objPoint = Nothing
            Set objCap = Nothing
        Next lngIndex
        .Redraw = True
    End With
End Sub
