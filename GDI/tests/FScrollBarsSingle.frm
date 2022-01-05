VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FScrollBarsSingle 
   Caption         =   "ScrollBars"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465.001
   OleObjectBlob   =   "FScrollBarsSingle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FScrollBarssingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Instance de clGdiplus pour le graphisme
Private WithEvents ogdi As clGdiplus
Attribute ogdi.VB_VarHelpID = -1

' Sur chargement du formulaire
Private Sub UserForm_Initialize()
' Crée une instance de la classe de graphisme
Set ogdi = New clGdiplus
' Crée une image à la taille du contrôle Image0
ogdi.CreateBitmapForControl Me.Image0
' Charge l'image
ogdi.ImgNew("fond").LoadControl Me.ImgFond
' Initialise les barres de l'image
ogdi.BarNew
ogdi.BarScaleX ogdi.img("fond").ImageWidth / 10, 10
ogdi.BarScaleY ogdi.img("fond").ImageHeight / 10, 10
ogdi.BarObject = Me.Image0
' Dessine les images
DrawImage
' Dessine l'image dans le contrôle
ogdi.Repaint Me.Image0
End Sub

' Dessine les images lorsqu'une barre de progression le demande
Private Sub ogdi_BarOnRefreshNeeded(BarName As String, MouseUp As Boolean)
DrawImage
' On utilise RepaintControlNoFormRepaint pour réduire d'éventuels scintillements
ogdi.RepaintNoFormRepaint Me.Image0
End Sub

Private Sub DrawImage()
' Remplit l'image de blanc
ogdi.Clear vbWhite
' Dessine l'image avec scroll
ogdi.DrawImg "fond", ogdi.BarLeft + ogdi.BarStartX, ogdi.BarTop + ogdi.BarStartY, , , , GdipSizeModeClip, GdipAlignTopLeft
' Dessine la barre de défilement
ogdi.BarDraw
End Sub

Private Sub UserForm_Terminate()
' Libère les objets
ogdi.BarDelete
Set ogdi = Nothing
End Sub
