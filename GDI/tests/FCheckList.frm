VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCheckList 
   Caption         =   "FCheckList"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105.001
   OleObjectBlob   =   "FCheckList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Instance de clGdiplus pour le graphisme
Private WithEvents ogdi As clGdiplus
Attribute ogdi.VB_VarHelpID = -1
Private oChecked() As Boolean
Private oTexte() As String
Private oKey() As String

' Sur chargement du formulaire
Private Sub UserForm_Initialize()
Dim lLine As Long
Dim lcpt As Long
' Crée une instance de la classe de graphisme
Set ogdi = New clGdiplus
' Crée une image à la taille du contrôle Image0
ogdi.CreateBitmapForControl Me.Image0
' Charge les valeurs de la feuille TDepartements
With ThisWorkbook.Sheets("Departements")
    ReDim Preserve oChecked(1 To .UsedRange.Row + .UsedRange.Rows.count - 1)
    ReDim Preserve oTexte(1 To .UsedRange.Row + .UsedRange.Rows.count - 1)
    ReDim Preserve oKey(1 To .UsedRange.Row + .UsedRange.Rows.count - 1)
    For lLine = .UsedRange.Row To .UsedRange.Row + .UsedRange.Rows.count - 1
        lcpt = lcpt + 1
        oChecked(lcpt) = False
        oTexte(lcpt) = .Cells(lLine, 3)
        oKey(lcpt) = .Cells(lLine, 2)
    Next
End With
' Initialise les barres de défilement
ogdi.BarNew
ogdi.BarObject = Me.Image0
' Génère l'image
DrawListe
' Dessine l'image dans le contrôle
ogdi.Repaint Me.Image0
End Sub

' Génère l'image de la liste
Private Sub DrawListe()
Dim lcpt As Long
Dim lMaxRight As Long, lMaxBottom As Long
' Remplit l'image de blanc
ogdi.FillColor vbWhite
ogdi.WorldPush
ogdi.WorldTranslate ogdi.BarStartX, ogdi.BarStartY
For lcpt = LBound(oTexte) To UBound(oTexte)
    ogdi.DrawControl 2, 2 + (lcpt - 1) * 20 + 2, 2 + 20, 2 + lcpt * 20 - 2, CtrlButtonCheck, , oChecked(lcpt)
    ogdi.DrawText oTexte(lcpt), 16, "Arial", 2 + 20 + 2, 2 + (lcpt - 1) * 20, _
                            2000, 2 + lcpt * 20, HorzAlignLeft
    ogdi.CreateRegionRect CStr(lcpt), 2, 2 + (lcpt - 1) * 20, ogdi.BarInsideWidth, 2 + lcpt * 20
    If ogdi.LastTextRight > lMaxRight Then lMaxRight = ogdi.LastTextRight
    If ogdi.LastTextBottom > lMaxBottom Then lMaxBottom = ogdi.LastTextBottom
Next
ogdi.WorldPop
ogdi.BarScaleX lMaxRight / 10 + 1, 10
ogdi.BarScaleY CSng(lcpt), 20
ogdi.BarDraw
End Sub

' Raffraichissement du texte contenant le choix
Private Sub RefreshChoiceTexte()
Dim lChoiceText As String
Dim lcpt As Long
' Pour chaque texte
For lcpt = LBound(oTexte) To UBound(oTexte)
    ' Vérifie si la case est cochée
    If oChecked(lcpt) Then
        ' Ajoute la clé et le libellé
        If lChoiceText <> "" Then lChoiceText = lChoiceText & vbCrLf
        lChoiceText = lChoiceText & oKey(lcpt) & " : " & oTexte(lcpt)
    End If
Next
' Affecte le texte à la textbox
Me.txtChoice.value = lChoiceText
Me.txtChoice.SelLength = 0
End Sub

' Souris appuyée
Private Sub ogdi_BarMouseDown(BarName As String, lregion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Si région sous la souris
If Len(lregion) > 0 Then
    ' Inverse la valeur de la case
    oChecked(CLng(lregion)) = Not oChecked(CLng(lregion))
    ' Redessine la liste
    DrawListe
    ogdi.RepaintNoFormRepaint Me.Image0
    RefreshChoiceTexte
End If
End Sub

' Demande de raffraichissement en provenance des barres de défilement
Private Sub ogdi_BarOnRefreshNeeded(BarName As String, MouseUp As Boolean)
DrawListe
ogdi.RepaintNoFormRepaint Me.Image0
End Sub


