VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMap 
   Caption         =   "Carte"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   OleObjectBlob   =   "FMap.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Collection contenant les départements et régions
Dim gSVGNodes As Collection
' Coordonnées minis et maxis
Dim gMinX As Long, gMinY As Long
Dim gMaxX As Long, gMaxY As Long
' Objet gdiplus pour graphisme
Private ogdi As clGdiplus

Private Sub UserForm_Initialize()
Set ogdi = New clGdiplus
ogdi.CreateBitmapForControl Me.Image0
ImportSVG
DrawSVG
ogdi.Repaint Me.Image0
End Sub

Private Sub Image0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Unload Me
End Sub

Private Sub Image0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lIdDep As String
Dim lIdRegion As String
Dim lX As Long, lY As Long
If ogdi Is Nothing Then Exit Sub
' Conversion en coordonnées image
lX = ogdi.CtrlToImgX(X, Me.Image0)
lY = ogdi.CtrlToImgY(Y, Me.Image0)
' Recherche de la région survolée
lIdRegion = ogdi.GetRegionXY(lX, lY)
' Recherche du département survolé
' on exclue la région pour trouver le département "dessous"
lIdDep = ogdi.GetRegionXY(lX, lY, , Array(lIdRegion))
' Remplit la région d'une couleur au hasard
If lIdDep <> "" Then
    gSVGNodes(lIdDep).BackColor = vbWhite * Rnd
    DrawSVG
    ' Redessine l'image (sans redessin du formualaire pour éviter les scintillements)
    ogdi.RepaintNoFormRepaint Me.Image0
End If
End Sub

Private Sub Image0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lIdDep As String
Dim lIdRegion As String
Dim lLibDep As String
Dim lLibRegion As String
Dim lX As Long, lY As Long
Dim l
If ogdi Is Nothing Then Exit Sub
' Conversion en coordonnées image
lX = ogdi.CtrlToImgX(X, Me.Image0)
lY = ogdi.CtrlToImgY(Y, Me.Image0)
' Recherche de la région survolée
lIdRegion = ogdi.GetRegionXY(lX, lY)
' Recherche du département survolé
' on exclue la région pour trouver le département "dessous"
lIdDep = ogdi.GetRegionXY(lX, lY, , Array(lIdRegion))
' Recherche des libellés dans les objets de type clSVGNode
' La région a été créé avec la clé de l'objet comme nom
If lIdDep <> "" Then
    lLibDep = gSVGNodes(lIdDep).liblong
End If
If lIdRegion <> "" Then
    lLibRegion = gSVGNodes(lIdRegion).liblong
End If
' Affichage d'infos
lblInfo.Caption = lX & " : " & lY & " ; " & lLibRegion & ";" & lLibDep
End Sub

'---------------------------------------------------------------------------------------------------------
' Importation des départements et régions
'---------------------------------------------------------------------------------------------------------
Private Function ImportSVG()
Dim loSVGNode As clSVGNode  ' Noeud SVG contenant un département
Dim lLine As Long
Set gSVGNodes = New Collection
' Chargement des départements
With ThisWorkbook.Sheets("departements")
    For lLine = .UsedRange.Row + 1 To .UsedRange.Row + .UsedRange.Rows.count - 1
        Set loSVGNode = New clSVGNode
        loSVGNode.libcourt = .Cells(lLine, 2)
        loSVGNode.liblong = .Cells(lLine, 3)
        loSVGNode.Path = .Cells(lLine, 4)
        loSVGNode.BackColor = -1
        loSVGNode.LineColor = RGB(150, 150, 150)
        loSVGNode.TypeElement = TypeDepartement
        gSVGNodes.Add loSVGNode, loSVGNode.key
    Next
End With
' Chargement des régions
With ThisWorkbook.Sheets("regions")
    For lLine = .UsedRange.Row + 1 To .UsedRange.Row + .UsedRange.Rows.count - 1
        Set loSVGNode = New clSVGNode
        loSVGNode.libcourt = .Cells(lLine, 2)
        loSVGNode.liblong = .Cells(lLine, 3)
        loSVGNode.Path = .Cells(lLine, 4)
        loSVGNode.LineWidth = 3
        loSVGNode.LineColor = vbBlack
        loSVGNode.BackColor = -1
        loSVGNode.TypeElement = TypeRegion
        gSVGNodes.Add loSVGNode, loSVGNode.key
    Next
End With
' Création des régions et recherche des minis/maxis
ogdi.RegionsDelete
For Each loSVGNode In gSVGNodes
    ogdi.CreateRegionSVG loSVGNode.key, loSVGNode.Path
    ogdi.RegionCombine "rgnsvgfull", loSVGNode.key, 2
Next
ogdi.RegionGetRect "rgnsvgfull", gMinX, gMinY, gMaxX, gMaxY
ogdi.RegionDelete "rgnsvgfull"
' Ajout de marges
gMaxX = gMaxX + 5
gMaxY = gMaxY + 5
gMinX = gMinX - 5
gMinY = gMinY - 5
End Function

'---------------------------------------------------------------------------------------------------------
' Dessin des départements et régions
'---------------------------------------------------------------------------------------------------------
Private Function DrawSVG()
Dim lWidth As Long, lHeight As Long
Dim loSVGNode As clSVGNode
lWidth = ogdi.ImageWidth
lHeight = ogdi.ImageHeight
ogdi.FillColor vbWhite
ogdi.RegionsDelete
ogdi.WorldPush
ogdi.WorldScale lWidth / (gMaxX - gMinX + 1), lWidth / (gMaxX - gMinX + 1), True
ogdi.WorldTranslate -gMinX, -gMinY, True
For Each loSVGNode In gSVGNodes
    ogdi.DrawSVG loSVGNode.Path, loSVGNode.BackColor, loSVGNode.LineColor, loSVGNode.LineWidth, , , loSVGNode.key
Next
ogdi.WorldPop
End Function
