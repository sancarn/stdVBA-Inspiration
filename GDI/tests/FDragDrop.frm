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
'*                               D�mo Drag&Drop sur une image                          *
'*                               et gestion des collisions entre r�gions               *
'***************************************************************************************

Private ogdi As clGdiplus    ' Classe image
Private gDragRgn As String    ' R�gion en cours de d�placement
Private gImages As New Collection    ' Collection d'objets images (type clImage)
Private gWidth As Long, gHeight As Long ' Taille de l'image

'---------------------------------------------------------------------------------------
' Bascule Visualiser les r�gions
'---------------------------------------------------------------------------------------
Private Sub BVisuRegions_AfterUpdate()
' Affiche ou masque les regions
    ImgPaint
    ogdi.RepaintNoFormRepaint Me.Image0
End Sub

'---------------------------------------------------------------------------------------
' Bascule r�gions au pixel pr�s
'---------------------------------------------------------------------------------------
Private Sub BCollisionsTransp_AfterUpdate()
' R�affiche si regions visibles
    If BVisuRegions.value Then
        ImgPaint    ' repaint l'image avec les r�gions visibles ou invisibles
        ogdi.RepaintNoFormRepaint Me.Image0    ' repaint l'image � l'�cran
    End If
End Sub

'---------------------------------------------------------------------------------------
' Click sur image
'---------------------------------------------------------------------------------------
Private Sub Image0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ogdi Is Nothing Then
        ' gDragRgn contient le nom de la region que l'on d�place
        gDragRgn = ogdi.GetRegionXY(ogdi.CtrlToImgX(X, Me.Image0), ogdi.CtrlToImgY(Y, Me.Image0))
    End If
End Sub

'---------------------------------------------------------------------------------------
' D�placement de la souris sur l'image
'---------------------------------------------------------------------------------------
Private Sub Image0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim lImage As clImage
    Dim lRegionCollision As String
    Dim NewX As Single, NewY As Single
    Static OldX As Single, OldY As Single
    ' V�rifie si la classe est initialis�e (elle peut �ventuellement �tre lib�r�e avant
    '    l'�v�nement MouseMove)
    If ogdi Is Nothing Then Exit Sub
    ' Nouvelle position de la souris, en pixel sur l'image
    NewX = ogdi.CtrlToImgX(X, Me.Image0)
    NewY = ogdi.CtrlToImgY(Y, Me.Image0)
    ' V�rifie qu'on a r�ellement d�plac� la souris
    If NewX = OldX And NewY = OldY Then Exit Sub
    ' Si une r�gion est en d�placement
    If gDragRgn <> "" Then
        ' lImage est la r�gion d�plac�e
        Set lImage = gImages.item(gDragRgn)    ' Objet image d�plac�
        lImage.X = lImage.X + CLng(NewX - OldX)    ' Left
        lImage.Y = lImage.Y + CLng(NewY - OldY)     ' Top
        ' Repaint le cont�le image avec les nouvelles positions
        ImgPaint
        ' Si gestion des collision
        If Me.BCollisions.value Then
            ' Cherche une r�gion en collision
            Call ogdi.RegionsIntersect(gDragRgn, lRegionCollision)
            ' Si on est entr� en collision avec une autre r�gion
            If lRegionCollision <> "" Then
                ' Retour � la position pr�c�dente
                lImage.X = lImage.X - CLng(NewX - OldX)   ' Left
                lImage.Y = lImage.Y - CLng(NewY - OldY)     ' Top
                ' Repaint le cont�le image
                ImgPaint
                ' Affiche le nom de la r�gion en collision
                LblCollision.Caption = "Collision avec " & gImages(lRegionCollision).name & "!!!"
            Else
                ' Pas de r�gion en collision
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
' Souris rel�ch�e
'---------------------------------------------------------------------------------------
Private Sub Image0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Pas de r�gion en cours de d�placement
    gDragRgn = ""
    ' Redessine l'image de mani�re persistante dans le contr�le
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
    ' Dessine chaque image et ajoute une r�gion si la gestion des collisions est activ�e
    ' Le nom de la r�gion est le pointeur unique de l'objet de type clImage
    For Each lImage In gImages
        ogdi.DrawImg lImage.id, _
                lImage.X, _
                lImage.Y, _
                lImage.X + ogdi.img(lImage.id).ImageWidth - 1, _
                lImage.Y + ogdi.img(lImage.id).ImageHeight - 1, vbWhite, _
                , , , lImage.id, IIf(BCollisionsTransp.value, vbWhite, -1)
        ' Affiche la r�gion si demand�
        If BVisuRegions.value Then ogdi.RegionHatch lImage.id, vbRed, , HatchStyleWideUpwardDiagonal, 100
    Next
End Sub

'---------------------------------------------------------------------------------------
' Sur fermeture du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
' Lib�re la classe image
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
    ' D�tection des collisions
    Me.BCollisions.value = True
    ' Taille de l'image = taille du contr�le
    gWidth = ogdi.PointsToPixelsX(Me.Image0.Width)
    gHeight = ogdi.PointsToPixelsY(Me.Image0.Height)
    ogdi.CreateBitmap gWidth, gHeight
    ' Chargement des images � partir des images du formulaire
    For Each loCtrl In Me.Controls
        ' Les images � ajouter sont dans les contr�les nomm�s img1,img2,img3
        If loCtrl.name Like "img*" Then
            ' Cr�e un nouvel objet de type clImage
            Set lImage = New clImage
            ' Ajoute l'image dans la liste d'images de gdi � partir du contr�le
            ' Le nom de l'image est le pointeur unique de l'objet de type clImage
            ogdi.ImgNew(lImage.id).LoadControl loCtrl
            ' Redimensionne les images � 60 pixels de large
            ogdi.img(lImage.id).Resize 60, , True
            ' Retire la bordure de l'image qui peut �ventuellement �tre g�n�r�e lors du redimensionnement
            ogdi.img(lImage.id).Crop 1, 1, ogdi.img(lImage.id).ImageWidth - 2, ogdi.img(lImage.id).ImageHeight - 2
            ' Position de d�part de l'image
            lImage.X = ogdi.PointsToPixelsX(loCtrl.Left)
            lImage.Y = ogdi.PointsToPixelsY(loCtrl.Top)
            ' Nom de l'image
            lImage.name = loCtrl.Tag
            ' Ajout de l'image � la collection
            ' L'id dans la collection est le pointeur unique de l'objet de type clImage
            gImages.Add lImage, lImage.id
        End If
    Next
    ' On dessine les images
    ImgPaint
    ' On affiche � l'�cran
    ogdi.Repaint Me.Image0
End Sub






