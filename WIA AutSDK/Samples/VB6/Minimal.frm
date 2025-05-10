VERSION 5.00
Object = "{94A0E92D-43C0-494E-AC29-FD45948A5221}#1.0#0"; "wiaaut.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin WIACtl.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set Picture1.Picture = CommonDialog1.ShowAcquireImage.FileData.Picture
End Sub
