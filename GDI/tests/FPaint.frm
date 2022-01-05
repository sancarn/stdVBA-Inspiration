VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FPaint 
   Caption         =   "Dessin"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435.001
   OleObjectBlob   =   "FPaint.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Instance de clGdiplus pour le graphisme
Private ogdi As clGdiplus
' Ancienne position de la souris
Private gX As Single, gY As Single

' Sur chargement du formulaire
Private Sub UserForm_Initialize()
' Crée une instance de la classe de graphisme
Set ogdi = New clGdiplus
' Crée une image à la taille du contrôle Image0
ogdi.CreateBitmapForControl Me.Image0
' Remplit l'image de blanc
ogdi.FillColor vbWhite
' Dessine l'image dans le contrôle
ogdi.Repaint Me.Image0
' Défini le contrôle de référence (évite des appels fréquents à CtrlToImg...)
ogdi.RefControl = Me.Image0
' Ancienne position de la souris par défaut
' Si la position est à (-1,-1), cela signifie qu'on n'a pas de point précédent
gX = -1
gY = -1
End Sub

' Sur souris appuyée sur l'image
Private Sub Image0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Select Case Button
    Case vbKeyLButton ' bouton gauche
        ' Conserve la position du clic pour pouvoir tracer la première ligne
        gX = X
        gY = Y
    Case vbKeyRButton ' bouton droit
        ' Crée une region et la remplit de rouge
        ogdi.CreateRegionFromColor "region", , , True, CLng(X), CLng(Y)
        ogdi.RegionFill "region", vbRed
        ogdi.RepaintNoFormRepaint Me.Image0
End Select
End Sub

' Sur souris déplacée sur l'image
Private Sub Image0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Select Case Button
    Case vbKeyLButton ' bouton gauche
        ' Si on a un point précédent
        If gX <> -1 Or gY <> -1 Then
            ' Trace une ligne du point précédent jusqu'au point courant
            ' Le type de début et fin de ligne est arrondi
            ogdi.LineStart = LineCapRound
            ogdi.LineEnd = LineCapRound
            ' comme on a défini RefControl, les coordonnées sont directement celle prise sur le contrôle
            ogdi.DrawLine gX, gY, X, Y, , 3
            ' Conserve la position du clic pour pouvoir tracer la prochaine ligne
            gX = X
            gY = Y
            ' Dessine l'image sur le contrôle (dessin rapide non persistant)
            ogdi.RepaintFast Me.Image0
        End If
End Select
End Sub

' Sur souris relâchée sur l'image
Private Sub Image0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Select Case Button
    Case vbKeyLButton ' bouton gauche
        ' Ancienne position de la souris par défaut
        ' on relâche la souris, on n'a donc plus de point précédent
        gX = -1
        gY = -1
        ' Dessine l'image sur le contrôle (le dessin est persistant lorqu'on relâche la souris)
        ogdi.RepaintNoFormRepaint Me.Image0
End Select
End Sub


