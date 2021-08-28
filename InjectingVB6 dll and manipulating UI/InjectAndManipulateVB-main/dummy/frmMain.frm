VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTest 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   900
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' // Executable should access to a VB.Global property
    Me.Caption = App.EXEName
End Sub
