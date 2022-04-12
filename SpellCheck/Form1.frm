VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1800
      List            =   "Form1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "El vurro la tiene larga, la baca la tinee corta"
      Top             =   840
      Width           =   6135
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6800
      _Version        =   393217
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0004
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":008F
      Top             =   1320
      Width           =   8055
   End
   Begin VB.Label Label1 
      Caption         =   "Select language:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   260
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSP1 As ClsSpellCheck
Dim cSP2 As ClsSpellCheck
Dim cSP3 As ClsSpellCheck

Private Sub Combo1_Click()
    cSP1.Language = Combo1.Text
    cSP2.Language = Combo1.Text
    cSP3.Language = Combo1.Text
End Sub

Private Sub Form_Load()
    Dim i As Long

    Set cSP1 = New ClsSpellCheck
    Set cSP2 = New ClsSpellCheck
    Set cSP3 = New ClsSpellCheck
    
    For i = 1 To cSP1.cSupportedLanguages.Count
        Combo1.AddItem cSP1.cSupportedLanguages.Item(i)
    Next

    RichTextBox1.Text = Text2.Text
    
    cSP1.Init Text1.hwnd
    cSP2.Init Text2.hwnd
    cSP3.Init RichTextBox1.hwnd
    
    Me.Caption = "Default Lenguaje: " & cSP1.Language
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cSP1.Terminate
    cSP2.Terminate
    cSP3.Terminate
End Sub

