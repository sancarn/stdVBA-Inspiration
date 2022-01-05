VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FDrawControl 
   Caption         =   "FDrawControl"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045.001
   OleObjectBlob   =   "FDrawControl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FDrawControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Curseur de souris
Private Const IDC_APPSTARTING = 32650
Private Const IDC_HAND = 32649
Private Const IDC_ARROW = 32512
Private Const IDC_CROSS = 32515
Private Const IDC_IBEAM = 32513
Private Const IDC_NO = 32648
Private Const IDC_SIZEALL = 32646
Private Const IDC_SIZENESW = 32643
Private Const IDC_SIZENS = 32645
Private Const IDC_SIZENWSE = 32642
Private Const IDC_SIZEWE = 32644
' Fonction API pour changer le curseur
#If VBA7 Then
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
#Else
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If

' Instance de clGdiplus pour le graphisme
Private ogdi As clGdiplus
' Flag d'état du bouton de test
Private gBtnPushed As Boolean
' Flag d'état du bouton de retournement
Private gBtnRotatePushed As Boolean
' Flags d'état des cases à cocher
Private gChecked(1 To 10) As Boolean
' Flags d'état des flèches de commentaires
Private gArrow(1 To 5) As Boolean
' Angle de rotation
Private gAngle As Long

' Sur chargement du formulaire
Private Sub UserForm_Initialize()
' Crée une instance de la classe de graphisme
Set ogdi = New clGdiplus
' Crée une image à la taille du contrôle Image0
ogdi.CreateBitmapForControl Me.Image0
' Génère l'image
DrawImage
' Dessine l'image dans le contrôle
ogdi.Repaint Me.Image0
End Sub

' Génère l'image
Private Sub DrawImage()
Dim lcpt As Long
Dim lCptLine As Long
Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
' Supprime les regions
ogdi.RegionsDelete
' Remplit l'image de jaune pale
ogdi.FillColor RGB(255, 255, 200)
' Si bouton de retournement poussé
If gAngle <> 0 Then
    ' Conserve les transformations actuelles
    ogdi.WorldPush
    ' Centre l'image en 0,0 avant de la tourner, puis la repositionne
    ogdi.WorldTranslate -ogdi.ImageWidth \ 2, -ogdi.ImageHeight \ 2
    ogdi.WorldRotate gAngle, True
    ogdi.WorldTranslate ogdi.ImageWidth \ 2, ogdi.ImageHeight \ 2, True
End If
' Dessin du cadre principale
ogdi.DrawControl 10, 10, ogdi.ImageWidth - 10, ogdi.ImageHeight - 10, CtrlFrameSunken
' Dessin du bouton
lx1 = 20: ly1 = 20: lx2 = 150: ly2 = 60
ogdi.DrawControl lx1, ly1, lx2, ly2, CtrlButtonPush, , gBtnPushed, , , True, "mybutton"
' Texte du bouton
ogdi.DrawText "Testez moi!", 12, , lx1 + gBtnPushed, ly1 + gBtnPushed, lx2 + gBtnPushed, ly2 + gBtnPushed
' Dessin du deuxième bouton
lx1 = 250: ly1 = 20: lx2 = 355: ly2 = 60
ogdi.DrawControl lx1, ly1, lx2, ly2, CtrlButtonPush, , gBtnRotatePushed, , , True, "mybuttonrotate"
' Texte du deuxième bouton
ogdi.DrawText "Retourner l'image!", 12, , lx1 + gBtnRotatePushed, ly1 + gBtnRotatePushed, lx2 + gBtnRotatePushed, ly2 + gBtnRotatePushed
' Si bouton de test poussé
If gBtnPushed Then
    ' Dessin du cadre secondaire
    lx1 = 20: ly1 = 65: lx2 = ogdi.ImageWidth - 20: ly2 = ogdi.ImageHeight - 20
    ogdi.DrawControl lx1, ly1, lx2, ly2, CtrlFrameRaised, , , , , True
    ogdi.FillColor RGB(230, 230, 255), , , lx1, ly1, lx2, ly2
    ' Texte d'info
    ogdi.DrawText "Testez les cases à chocher et les menus déroulants!", 14, , lx1, ly1, lx2, ly2, HorzAlignCenter, VertAlignTop, , , , , True, True, , , True
    ' Affiche 10 cases
    For lcpt = 1 To 10
        lx1 = 30: ly1 = 90 + (lcpt - 1) * 20: lx2 = 50: ly2 = 90 + lcpt * 20 - 5
        ogdi.DrawControl lx1, ly1, lx2, ly2, CtrlButtonCheck, , gChecked(lcpt), , , , "mycheckbox" & lcpt
        ogdi.DrawText "Test Ligne " & lcpt, 12, , lx2 + 5, ly1, lx2 + 100, ly2, HorzAlignLeft
    Next
    ' Ajoute 5 commentaires
    lcpt = 0
    Do
        lcpt = lcpt + 1
        lCptLine = lCptLine + 1
        If lcpt = 6 Then Exit Do
        ' Dessine une petite flèche
        lx1 = 200: ly1 = 90 + (lCptLine - 1) * 20: lx2 = 220: ly2 = 90 + lCptLine * 20 - 5
        ' Les contrôle CtrlMenu* sont dessinés sur fond blanc
        ' On dessine d'abord la flèche sur une image puis on change sa couleur de fond
        ' Pour de meilleurs performances, on pourrait créer une seule fois l'image dès l'ouverture du formulaire
        With ogdi.ImgNew("arrow")
            .CreateBitmap 15, 15
            .DrawControl 0, 0, 15 - 1, 15 - 1, CtrlMenuArrow
            .ReplaceColor .GetPixel(0, 0), RGB(230, 230, 255)
        End With
        ' Si menu déroulé => retourne la fleche
        If gArrow(lcpt) Then ogdi.img("arrow").Rotate 90
        ogdi.DrawImg "arrow", lx1, ly1, lx2, ly2, , , , , "myarrow" & lcpt
        ' Ecrit le libellé du commentaire
        ogdi.DrawText "Commentaire " & lcpt, 12, , lx2 + 5, ly1, lx2 + 100, ly2, HorzAlignLeft
        ' Ecrit le texte du commentaire s'il est déroulé
        If gArrow(lcpt) Then
            ogdi.DrawText "Contenu du commentaire " & lcpt, 12, , lx2 + 30, ly1 + 20, lx2 + 200, ly2 + 20, HorzAlignLeft, , , , , , True
            lCptLine = lCptLine + 1
        End If
    Loop
End If
' Si bouton de retournement poussé
If gAngle <> 0 Then
    ' Restaure les transformations
    ogdi.WorldPop
End If
End Sub

Private Sub Image0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
Dim lVal As Long
' Région cliquée
lregion = ogdi.GetRegionXY(X, Y, , , Me.Image0)
' Si bouton
If lregion = "mybutton" Then
    ' Inverse la valeur du bouton
    gBtnPushed = Not gBtnPushed
    ' Redessine l'image
    DrawImage
    ogdi.RepaintNoFormRepaint Me.Image0
End If
' Si bouton de retournement
If lregion = "mybuttonrotate" Then
    ' Effet de rotation
    Dim lAngleMin As Long, lAngleMax As Long, lAngleStep As Long
    If gAngle = 0 Then
        lAngleMin = 0
        lAngleMax = 180
        lAngleStep = 5
    ElseIf gAngle = 180 Then
        lAngleMin = 180
        lAngleMax = 0
        lAngleStep = -5
    End If
    For gAngle = lAngleMin To lAngleMax Step lAngleStep
        ' Redessine l'image
        DrawImage
        ogdi.RepaintFast Me.Image0
    Next
    gAngle = lAngleMax
    ' Inverse la valeur du bouton
    gBtnRotatePushed = Not gBtnRotatePushed
    ' Redessine l'image
    DrawImage
    ogdi.RepaintNoFormRepaint Me.Image0
End If
' Si case à cocher
If lregion Like "mycheckbox*" Then
    ' Extrait le numéro de la case
    lVal = Mid(lregion, Len("mycheckbox") + 1)
    ' Inverse la valeur de la case
    gChecked(lVal) = Not gChecked(lVal)
    ' Redessine l'image
    DrawImage
    ogdi.RepaintNoFormRepaint Me.Image0
End If
' Si commentaire
If lregion Like "myarrow*" Then
    ' Extrait le numéro du commentaire
    lVal = Mid(lregion, Len("myarrow") + 1)
    ' Inverse la valeur du menu (la flèche)
    gArrow(lVal) = Not gArrow(lVal)
    ' Redessine l'image
    DrawImage
    ogdi.RepaintNoFormRepaint Me.Image0
End If
End Sub


