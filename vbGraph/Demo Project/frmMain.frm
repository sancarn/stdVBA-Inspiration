VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbGraph Control Demo"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Progress Bar"
      Height          =   465
      Index           =   6
      Left            =   420
      TabIndex        =   6
      Top             =   1800
      Width           =   1365
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Line Style"
      Height          =   465
      Index           =   5
      Left            =   1860
      TabIndex        =   5
      Top             =   1290
      Width           =   1365
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Bar Style"
      Height          =   465
      Index           =   4
      Left            =   1860
      TabIndex        =   4
      Top             =   750
      Width           =   1365
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Custom Style"
      Height          =   465
      Index           =   3
      Left            =   1860
      TabIndex        =   3
      Top             =   210
      Width           =   1365
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Graph Style"
      Height          =   465
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   1290
      Width           =   1365
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Taskman Style"
      Height          =   465
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   750
      Width           =   1365
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Music Style"
      Height          =   465
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   210
      Width           =   1365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STYLE_MUSIC As Long = 0
Private Const STYLE_TASKMAN As Long = 1
Private Const STYLE_GRAPH As Long = 2
Private Const STYLE_CUSTOM As Long = 3
Private Const STYLE_BAR As Long = 4
Private Const STYLE_LINE As Long = 5
Private Const STYLE_PROGRESS As Long = 6

Private Sub Form_Load()
    Randomize
End Sub

Private Sub cmdStyle_Click(Index As Integer)
    Select Case Index
        Case STYLE_MUSIC
            frmMusic.Show vbModal, Me
        Case STYLE_TASKMAN
            frmTaskman.Show vbModal, Me
        Case STYLE_PROGRESS
            frmProgress.Show vbModal, Me
        Case STYLE_LINE
            frmLine.Show vbModal, Me
        Case STYLE_BAR
            frmBars.Show vbModal, Me
        Case STYLE_CUSTOM
            frmCustom.Show vbModal, Me
        Case STYLE_GRAPH
            frmGraph.Show vbModal, Me
    End Select
End Sub
