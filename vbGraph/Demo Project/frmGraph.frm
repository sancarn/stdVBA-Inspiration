VERSION 5.00
Object = "*\A..\vbGraph.vbp"
Begin VB.Form frmGraph 
   Caption         =   "Graph"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin vbGraph.Graph Graph1 
      Height          =   1395
      Left            =   2400
      TabIndex        =   2
      Top             =   540
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2461
      State           =   "frmGraph.frx":0000
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   3435
      Width           =   8055
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   435
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
    ChangeGrid
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
    ChangeGrid
End Sub

Private Sub SetupGraph()
    With Graph1
        .FadeIn = False
        .FixedPoints = 0
        .MaxValue = 2
        .BarWidth = 0.8
        .MinValue = -2
        .YGridInc = 0.5
        .xGridInc = 90
        .AxisColor = 0
        .BackColor = RGB(255, 255, 255)
        .GridColor = RGB(200, 200, 200)
        .ShowAxis = True
        .ShowGrid = True
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
        .LineColor = RGB(200, 50, 255)
        .PointColor = RGB(255, 100, 100)
    End With
End Sub

Private Sub ChangeGrid()
Dim lngX        As Long
Dim dblValue    As Double
Dim lngType     As Long
    lngType = CLng(Rnd * 3)
    With Graph1
        .Redraw = False
        .Datasets.Item(1).Points.Clear
        For lngX = 1 To 360
            Select Case lngType
                Case 0
                    dblValue = Cos(CDbl(CDbl(3.14) * CDbl(2) * CDbl((lngX / 360))))
                Case 1
                    dblValue = Sin(CDbl(CDbl(3.14) * CDbl(2) * CDbl((lngX / 360))))
                Case 2
                    dblValue = Tan(CDbl(CDbl(3.14) * CDbl(2) * CDbl((lngX / 360))))
                Case 3
                    dblValue = ((((lngX / 36) - 5) ^ 3) - ((lngX / 36) - 5) ^ 2)
            End Select
            .Datasets.Item(1).Points.Add dblValue
        Next lngX
        .Redraw = True
    End With
End Sub
