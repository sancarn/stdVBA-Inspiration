VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FDragDrop 
   Caption         =   "Drag 'n Drop"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865.001
   OleObjectBlob   =   "FDragDrop.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FDragDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'***************************************************************************************
'*                               Démo Drag&Drop sur une image                          *
'*                               et gestion des collisions entre régions               *
'***************************************************************************************

Private ogdi As clGdiplus    ' Classe image
Private gDragRgn As String    ' Région en cours de déplacement
Private gImages As New Collection    ' Collection d'objets images (type clImage)
Private gWidth As Long, gHeight As Long ' Taille de l'image

'---------------------------------------------------------------------------------------
' Bascule Visualiser les régions
'---------------------------------------------------------------------------------------
Private Sub BVisuRegions_AfterUpdate()
' Affiche ou masque les regions
    ImgPaint
    ogdi.RepaintNoFormRepaint Me.Image0
End Sub

'---------------------------------------------------------------------------------------
' Bascule régions au pixel près
'---------------------------------------------------------------------------------------
Private Sub BCollisionsTransp_AfterUpdate()
' Réaffiche si regions visibles
    If BVisuRegions.value Then
        ImgPaint    ' repaint l'image avec les régions visibles ou invisibles
        ogdi.RepaintNoFormRepaint Me.Image0    ' repaint l'image à l'écran
    End If
End Sub

'---------------------------------------------------------------------------------------
' Click sur image
'---------------------------------------------------------------------------------------
Private Sub Image0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ogdi Is Nothing Then
        ' gDragRgn contient le nom de la region que l'on déplace
        gDragRgn = ogdi.GetRegionXY(ogdi.CtrlToImgX(X, Me.Image0), ogdi.CtrlToImgY(Y, Me.Image0))
    End If
End Sub

'---------------------------------------------------------------------------------------
' Déplacement de la souris sur l'image
'---------------------------------------------------------------------------------------
Private Sub Image0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim lImage As clImage
    Dim lRegionCollision As String
    Dim NewX As Single, NewY As Single
    Static OldX As Single, OldY As Single
    ' Vérifie si la classe est initialisée (elle peut éventuellement être libérée avant
    '    l'événement MouseMove)
    If ogdi Is Nothing Then Exit Sub
    ' Nouvelle position de la souris, en pixel sur l'image
    NewX = ogdi.CtrlToImgX(X, Me.Image0)
    NewY = ogdi.CtrlToImgY(Y, Me.Image0)
    ' Vérifie qu'on a réellement déplacé la souris
    If NewX = OldX And NewY = OldY Then Exit Sub
    ' Si une région est en déplacement
    If gDragRgn <> "" Then
        ' lImage est la région déplacée
        Set lImage = gImages.item(gDragRgn)    ' Objet image déplacé
        lImage.X = lImage.X + CLng(NewX - OldX)    ' Left
        lImage.Y = lImage.Y + CLng(NewY - OldY)     ' Top
        ' Repaint le contôle image avec les nouvelles positions
        ImgPaint
        ' Si gestion des collision
        If Me.BCollisions.value Then
            ' Cherche une région en collision
            Call ogdi.RegionsIntersect(gDragRgn, lRegionCollision)
            ' Si on est entré en collision avec une autre région
            If lRegionCollision <> "" Then
                ' Retour à la position précédente
                lImage.X = lImage.X - CLng(NewX - OldX)   ' Left
                lImage.Y = lImage.Y - CLng(NewY - OldY)     ' Top
                ' Repaint le contôle image
                ImgPaint
                ' Affiche le nom de la région en collision
                LblCollision.Caption = "Collision avec " & gImages(lRegionCollision).name & "!!!"
            Else
                ' Pas de région en collision
                LblCollision.Caption = ""
            End If
        End If
    End If
    ' Redessine l'image rapidement (non persistant)
    ogdi.RepaintFast Me.Image0
    ' Sauvegarde la position de la souris
    OldX = NewX
    OldY = NewY
End Sub

'---------------------------------------------------------------------------------------
' Souris relâchée
'---------------------------------------------------------------------------------------
Private Sub Image0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Pas de région en cours de déplacement
    gDragRgn = ""
    ' Redessine l'image de manière persistante dans le contrôle
    ogdi.RepaintNoFormRepaint Me.Image0
End Sub

'---------------------------------------------------------------------------------------
' Redessine l'image
'---------------------------------------------------------------------------------------
Private Sub ImgPaint()
' Variable pour tableau Image
    Dim lImage As clImage
    ' Remplit le fond de la couleur du formulaire
    ogdi.FillColor Me.BackColor
    ' Dessine chaque image et ajoute une région si la gestion des collisions est activée
    ' Le nom de la région est le pointeur unique de l'objet de type clImage
    For Each lImage In gImages
        ogdi.DrawImg lImage.id, _
                lImage.X, _
                lImage.Y, _
                lImage.X + ogdi.img(lImage.id).ImageWidth - 1, _
                lImage.Y + ogdi.img(lImage.id).ImageHeight - 1, vbWhite, _
                , , , lImage.id, IIf(BCollisionsTransp.value, vbWhite, -1)
        ' Affiche la région si demandé
        If BVisuRegions.value Then ogdi.RegionHatch lImage.id, vbRed, , HatchStyleWideUpwardDiagonal, 100
    Next
End Sub

'---------------------------------------------------------------------------------------
' Sur fermeture du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
' Libère la classe image
If Not ogdi Is Nothing Then Set ogdi = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Sur chargement du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
Dim loCtrl As MSForms.Control
Dim lImage As clImage
    ' Initialisation de l'image
    Set ogdi = New clGdiplus
    ' Détection des collisions
    Me.BCollisions.value = True
    ' Taille de l'image = taille du contrôle
    gWidth = ogdi.PointsToPixelsX(Me.Image0.Width)
    gHeight = ogdi.PointsToPixelsY(Me.Image0.Height)
    ogdi.CreateBitmap gWidth, gHeight
    ' Chargement des images à partir des images du formulaire
    For Each loCtrl In Me.Controls
        ' Les images à ajouter sont dans les contrôles nommés img1,img2,img3
        If loCtrl.name Like "img*" Then
            ' Crée un nouvel objet de type clImage
            Set lImage = New clImage
            ' Ajoute l'image dans la liste d'images de gdi à partir du contrôle
            ' Le nom de l'image est le pointeur unique de l'objet de type clImage
            ogdi.ImgNew(lImage.id).LoadControl loCtrl
            ' Redimensionne les images à 60 pixels de large
            ogdi.img(lImage.id).Resize 60, , True
            ' Retire la bordure de l'image qui peut éventuellement être générée lors du redimensionnement
            ogdi.img(lImage.id).Crop 1, 1, ogdi.img(lImage.id).ImageWidth - 2, ogdi.img(lImage.id).ImageHeight - 2
            ' Position de départ de l'image
            lImage.X = ogdi.PointsToPixelsX(loCtrl.Left)
            lImage.Y = ogdi.PointsToPixelsY(loCtrl.Top)
            ' Nom de l'image
            lImage.name = loCtrl.Tag
            ' Ajout de l'image à la collection
            ' L'id dans la collection est le pointeur unique de l'objet de type clImage
            gImages.Add lImage, lImage.id
        End If
    Next
    ' On dessine les images
    ImgPaint
    ' On affiche à l'écran
    ogdi.Repaint Me.Image0
End Sub






