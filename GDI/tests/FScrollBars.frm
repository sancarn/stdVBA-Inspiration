VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FScrollBars 
   Caption         =   "ScrollBars"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   OleObjectBlob   =   "FScrollBars.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Instance de clGdiplus pour le graphisme
Private WithEvents ogdi As clGdiplus
Attribute ogdi.VB_VarHelpID = -1
' Facteur de zoom
Private gZoom As Single

Private Sub btnZoomDown_Click()
If gZoom > 0.5 Then
    Me.txtZoom2.value = gZoom - 0.5
    txtZoom2_AfterUpdate
End If
End Sub

Private Sub btnZoomUp_Click()
Me.txtZoom2.value = gZoom + 0.5
txtZoom2_AfterUpdate
End Sub

' Sur chargement du formulaire
Private Sub UserForm_Initialize()
Me.txtZoom2.value = 1
gZoom = 1
' Crée une instance de la classe de graphisme
Set ogdi = New clGdiplus
' Crée une image à la taille du contrôle Image0
ogdi.CreateBitmapForControl Me.Image0
' Charge les deux images
ogdi.ImgNew("fond").LoadControl Me.imgFond1
ogdi.ImgNew("fond2").LoadControl Me.imgFond2
' Initialise l'objet pour les barres de l'image 1
ogdi.BarNew "fond"
ogdi.BarScaleX ogdi.img("fond").ImageWidth / 10, 10, 50, 300, "fond"
ogdi.BarScaleY ogdi.img("fond").ImageHeight / 10, 10, 50, 250, "fond"
ogdi.BarObject = Me.Image0
' Initialise l'objet pour les barres de l'image 2
ogdi.BarNew "fond2"
' Initialise chaque barre
ogdi.BarScaleX ogdi.img("fond2").ImageWidth / 10, 10, 310, 310 + 200, "fond2"
ogdi.BarScaleY ogdi.img("fond2").ImageHeight / 10, 10, 50, 250, "fond2"
ogdi.BarObject = Me.Image0
' Region pour clipping du dessin
ogdi.CreateRegionRect "clip1", 50, 50, 300, 250
ogdi.CreateRegionRect "clip2", 310, 50, 310 + 200, 250
' Dessine les images
DrawImage
' Dessine l'image dans le contrôle
ogdi.Repaint Me.Image0
End Sub

Private Sub DrawImage()
' Remplit l'image de blanc
ogdi.FillColor vbWhite
' Dessine la première image avec limitation à la région clip1
ogdi.DrawClipRegion = "clip1"
ogdi.DrawImg "fond", ogdi.BarLeft("fond") + ogdi.BarStartX("fond"), ogdi.BarTop("fond") + ogdi.BarStartY("fond"), ogdi.BarLeft("fond") + ogdi.BarInsideWidth("fond") - 1, ogdi.BarTop("fond") + ogdi.BarInsideHeight("fond") - 1, , GdipSizeModeClip, GdipAlignTopLeft
ogdi.DrawClipRegion = ""
' Dessine les barres de défilement de la première image
ogdi.BarDraw "fond"
' Dessine la deuxième image avec limitation à la région clip2
ogdi.CreateRegionRect "clip2", ogdi.BarLeft("fond2"), ogdi.BarTop("fond2"), ogdi.BarLeft("fond2") + ogdi.BarInsideWidth("fond2"), ogdi.BarTop("fond2") + ogdi.BarInsideHeight("fond2")
ogdi.DrawClipRegion = "clip2"
ogdi.WorldPush
ogdi.WorldScale gZoom, gZoom
ogdi.WorldTranslate ogdi.BarLeft("fond2") + ogdi.BarStartX("fond2"), ogdi.BarTop("fond2") + ogdi.BarStartY("fond2"), True
ogdi.DrawImg "fond2", 0, 0, , , , GdipSizeModeClip, GdipAlignTopLeft
ogdi.WorldPop
ogdi.DrawClipRegion = ""
' Dessine les barres de défilement de la deuxième image
ogdi.BarDraw "fond2"
' Encadre chaque image avec ses barres
ogdi.RegionFrame "clip1", 0
ogdi.DrawRectangle ogdi.BarLeft("fond2"), ogdi.BarTop("fond2"), ogdi.BarRight("fond2"), ogdi.BarBottom("fond2")
End Sub

' Dessine les images lorsqu'une barre de progression le demande
Private Sub ogdi_BarOnRefreshNeeded(BarName As String, MouseUp As Boolean)
DrawImage
ogdi.RepaintNoFormRepaint Me.Image0
End Sub

Private Sub txtZoom2_AfterUpdate()
Dim lOldZoom As Single
lOldZoom = gZoom
gZoom = txtZoom2.value
If gZoom = 0 Then gZoom = 1
' Initialise chaque barre
ogdi.BarScaleX ogdi.img("fond2").ImageWidth / 10 * gZoom, 10, 310, 310 + 200, "fond2"
ogdi.BarScaleY ogdi.img("fond2").ImageHeight / 10 * gZoom, 10, 50, 250, "fond2"
ogdi.BarStartX("fond2") = -((-ogdi.BarStartX("fond2") + ogdi.BarInsideWidth("fond2") / 2) / lOldZoom * gZoom - ogdi.BarInsideWidth("fond2") / 2)
ogdi.BarStartY("fond2") = -((-ogdi.BarStartY("fond2") + ogdi.BarInsideHeight("fond2") / 2) / lOldZoom * gZoom - ogdi.BarInsideHeight("fond2") / 2)
DrawImage
ogdi.RepaintNoFormRepaint Me.Image0
End Sub


Private Sub UserForm_Terminate()
' Libère les objets
ogdi.BarDelete "fond"
ogdi.BarDelete "fond2"
Set ogdi = Nothing
End Sub
