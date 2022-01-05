VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FTextures 
   Caption         =   "Textures"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   OleObjectBlob   =   "FTextures.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FTextures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private o As clGdiplus
Private gStop As Boolean ' Flag pour arrêt de la boucle

Private Sub UserForm_Initialize()
Set o = New clGdiplus
o.SmoothingMode = GdipSmoothingAntialias
' Creation bitmap
Call o.CreateBitmap(o.PointsToPixelsX(Me.Image0.Width), o.PointsToPixelsY(Me.Image0.Height))
' Fond blanc
o.FillColor vbWhite
' Crayon a l'interieur
o.PenAlignMode = PenAlignmentInset
' Ajout des textures
o.TextureAddFromControl "filltext1", Me.ImgCiel
o.TextureAddFromControl "filltext2", Me.ImgCitrouille
o.ImgNew("text").CreateBitmap 100, 30
o.img("text").FillColor RGB(255, 255, 200)
o.img("text").DrawText "Texture", 20, , 0, 0, 100, 30, , , vbGreen
o.TextureAddFromImg "filltext3", "text"
o.ImgDelete "text"
' Ecrit un texte
o.FillTexture = ""
o.DrawText "Texture" & vbCrLf & "   Test", 100, , 5, 5, o.ImageWidth + 5, o.ImageHeight + 5, 2, 0, , , , , , , , , True
o.FillTexture = "filltext1"
o.DrawText "Texture" & vbCrLf & "   Test", 100, , 0, 0, o.ImageWidth - 1, o.ImageHeight - 1, 2, 0, vbBlue, , , , , , , , True
' Dessine une ellipse
o.PenTexture = "filltext2"
o.FillTexture = "filltext1"
o.DrawEllipse 10, 10, o.ImageWidth / 2, o.ImageHeight / 2, 0, , vbRed, 5
o.PenTexture = ""
' Dessine un rectangle
o.FillTexture = "filltext2"
o.TextureTranslate "filltext2", o.ImageWidth / 2 + 20, o.ImageHeight / 2 + 20
o.DrawRectangle o.ImageWidth / 2 + 20, o.ImageHeight / 2 + 20, o.ImageWidth / 2 + 200, o.ImageHeight / 2 + 200, , vbBlue, 0
o.TextureTranslate "filltext2", 0, 0
' Dessine un polygone
o.FillTexture = "filltext3"
o.DrawPolygon Array(20, 300, 300, 350, 350, 300, 350, 350, 100, 450, 20, 300), , vbBlue, 2
o.FillTexture = ""
' Affiche l'image
o.Repaint Me.Image0
' Conserve l'image
o.ImageKeep

' Lance le timer pour exécution asynchrone
Application.OnTime DateAdd("s", 0, Now), "RunTextures"
End Sub

Public Sub RunTextures()
Static lcpt As Long
Do
    o.Wait 100, True
    If gStop Then Exit Sub
    lcpt = lcpt + 1
    If lcpt > o.TextureWidth("filltext1") Then lcpt = 0
    o.ImageReset
    ' Dessine une ellipse
    o.PenTexture = "filltext2"
    o.TextureReset "filltext1"
    o.TextureTranslate "filltext1", CSng(lcpt), 0
    o.FillTexture = "filltext1"
    o.DrawEllipse 10, 10, o.ImageWidth / 2, o.ImageHeight / 2, 0, , vbRed, 5
    o.PenTexture = ""
    ' Ecrit un texte
    o.FillTexture = "filltext1"
    o.TextureReset "filltext1"
    o.TextureTranslate "filltext1", -CSng(lcpt) * 2, 0
    o.DrawText "Texture" & vbCrLf & "   Test", 100, , 0, 0, o.ImageWidth - 1, o.ImageHeight - 1, 2, 0, vbBlue
    o.RepaintNoFormRepaint Me.Image0
Loop
End Sub

Private Sub UserForm_Terminate()
gStop = True
End Sub
