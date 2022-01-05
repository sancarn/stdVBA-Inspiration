VERSION 5.00
Begin VB.Form frmThread 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmThread.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateUnsafe 
      Caption         =   "Create Thread"
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   1620
      Width           =   4455
   End
   Begin VB.Label lblThread 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreateUnsafe_Click()
    CreateThread ByVal 0&, 0, AddressOf ThreadProc, ByVal g_pVbHeader, 0, 0
End Sub

Private Sub Form_Load()
    lblThread.Caption = Hex$(App.ThreadID)
End Sub
