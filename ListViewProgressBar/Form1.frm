VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   6360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8070
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mostrar Texto"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   5280
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color del Texto"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color de la Barra"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color de Borde"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color de Fondo"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Usar Temas de Windows"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   5880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cLVProgress As cListViewProgress

Private Sub Check1_Click()
    cLVProgress.UseWindowsTheme = Check1
    
    Command1(0).Enabled = Check1 = 0
    Command1(1).Enabled = Check1 = 0
    Command1(2).Enabled = Check1 = 0
End Sub

Private Sub Check2_Click()
    cLVProgress.TextVisible = Check2
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo Cancelar
CommonDialog1.CancelError = True
CommonDialog1.ShowColor

Select Case Index
    Case 0
        cLVProgress.BackColor = CommonDialog1.Color
    Case 1
        cLVProgress.BorderColor = CommonDialog1.Color
    Case 2
        cLVProgress.FillColor = CommonDialog1.Color
    Case 3
        cLVProgress.TextColor = CommonDialog1.Color
End Select

Cancelar:
End Sub






Private Sub Form_Load()

Dim i As Long
Dim Item As ListItem
Set cLVProgress = New cListViewProgress

    With ImageList1
      .ImageHeight = 16
      .ImageWidth = 16
      .ListImages.Add Picture:=Me.Icon
    End With
 

    With ListView1
        .HideSelection = False
        .SmallIcons = ImageList1
        .MultiSelect = True
        .View = lvwReport
        For i = 1 To 4
            .ColumnHeaders.Add Text:="columna " & i
        Next
   
        For i = 0 To 100
            With ListView1.ListItems.Add(, , "item" & i, , 1)
            .SubItems(1) = i
            .SubItems(2) = "EXFCH" & i
            .SubItems(3) = i * 1000
            End With
        Next

            
            cLVProgress.SubItemProgress = 1
            cLVProgress.SubClassListView .hwnd
            
    End With
    Timer1.Interval = 100
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cLVProgress = Nothing
End Sub

Private Sub Timer1_Timer()
Static x As Integer
Dim i As Integer
Dim Item As ListItem

cLVProgress.NoEraseBackGroud = True         'Al poner esta propiedad en True, deshabilita el repintado del fondo y produce menos parpadeo

For i = 0 To ListView1.ListItems.Count
    'Al utilizar la función "SetSubItemText" de la clase "cListViewProgress" para modificar un SubItems es más óptimo ya que no repinta todo el listview sino sólo ese items
    'ese es uno de los tantos problemas que tienen los ocx de los Common Controls de Microsoft. Son muy limitados mientras que las dll que utiliza
    'el sistema de windows tienen muchas más funciones.
    'Otra ventaja de utlizar la función SetSubItemText es que no volverá el ScrollBar del listview a 0 al modificar algún SubItems.
    'Nota si se utliza "SetSubItemText" también debe utilzarce "GetSubItemText" ya que internamente el ocx de los common control obtienen el valor
    'desde alguna colección interna y no mostrará lo que nosotros asignemos via Api.
    If Val(cLVProgress.GetSubItemText(ListView1.hwnd, i, 1)) > 100 Then cLVProgress.SetSubItemText ListView1.hwnd, i, 1, 0
    cLVProgress.SetSubItemText ListView1.hwnd, i, 1, Val(cLVProgress.GetSubItemText(ListView1.hwnd, i, 1)) + (Rnd(2) * 4)
Next

cLVProgress.Refresh                         'Actualizamos todo
cLVProgress.NoEraseBackGroud = False        'importante reponer el repintado del fondo!!

End Sub


