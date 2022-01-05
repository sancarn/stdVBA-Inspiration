VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSplashScreen 
   Caption         =   "UserForm3"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585.001
   OleObjectBlob   =   "FSplashScreen.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*                               Démo SplashScreen                                     *
'***************************************************************************************

'***************************************************************************************

Private clGdip As clGdiplus    ' Classe image

'---------------------------------------------------------------------------------------
' Souris appuyée sur image
'---------------------------------------------------------------------------------------
Private Sub ImgFond_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
    ' Région sous le curseur de la souris
    If Not clGdip Is Nothing Then
        lregion = clGdip.GetRegionXY(clGdip.CtrlToImgX(X, Me.ImgFond), clGdip.CtrlToImgY(Y, Me.ImgFond))
    End If
    
    If lregion = "fleche" Then
        ' Si on clique sur la flèche
        ' Dessine la flèche en position enfoncée avec texte rouge
        PaintButton True, True
        ' Redessine l'image
        clGdip.RepaintFast Me.ImgFond
    Else
        ' Sinon on déplace le formulaire
        clGdip.DragForm Me
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Souris déplacée sur l'image
'---------------------------------------------------------------------------------------
Private Sub ImgFond_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
Static sRegion As String
    ' Région sous le curseur de la souris
    If Not clGdip Is Nothing Then
        lregion = clGdip.GetRegionXY(clGdip.CtrlToImgX(X, Me.ImgFond), clGdip.CtrlToImgY(Y, Me.ImgFond))
    End If
    ' Si changement de région survolée
    If sRegion <> lregion Then
        ' Redessine la flèche
        ' En position appuyée si le bouton gauche est appuyé
        ' Avec couleur rouge si on survole la flèche
        PaintButton (Button = vbKeyLButton), (lregion = "fleche")
        ' Redessine l'image
        clGdip.RepaintFast Me.ImgFond
        ' Converse le nom de la dernière région survolée
        sRegion = lregion
    End If
End Sub

'---------------------------------------------------------------------------------------
' Souris relâchée sur l'image
'---------------------------------------------------------------------------------------
Private Sub ImgFond_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
    ' Région sous le curseur de la souris
    If Not clGdip Is Nothing Then
        lregion = clGdip.GetRegionXY(clGdip.CtrlToImgX(X, Me.ImgFond), clGdip.CtrlToImgY(Y, Me.ImgFond))
    End If
    
    If lregion <> "fleche" Then
        ' Si on survole la flèche
        ' On repaint le bouton en position relâchée
        PaintButton False, False
        ' Redessine l'image
        clGdip.RepaintFast Me.ImgFond
    Else
      ' Sinon on ferme le formulaire
        Unload Me
        ' Affiche le menu
        FMenu.Show
    End If
End Sub

'---------------------------------------------------------------------------------------
' Sur fermeture du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
' Libère la classe image
If Not clGdip Is Nothing Then Set clGdip = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Sur chargement du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
 Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
 Dim lDecalage As Long
 Dim lSize As Long
' Initialise la classe
On Error Resume Next
Set clGdip = New clGdiplus
If Err.Number <> 0 Then
    FormGdiplusError.Show
    Exit Sub
End If
On Error GoTo gestion_erreurs
' Charge l'image du contrôle
clGdip.LoadControl Me.ImgFond
' Position du logo
lx1 = clGdip.CtrlToImgX(Me.ImgLogo.Left - Me.ImgFond.Left, Me.ImgFond)
ly1 = clGdip.CtrlToImgY(Me.ImgLogo.Top - Me.ImgFond.Top, Me.ImgFond)
lx2 = clGdip.CtrlToImgX(Me.ImgLogo.Left - Me.ImgFond.Left + Me.ImgLogo.Width, Me.ImgFond)
ly2 = clGdip.CtrlToImgY(Me.ImgLogo.Top - Me.ImgFond.Top + Me.ImgLogo.Height, Me.ImgFond)
' Charge l'image du logo à partir du contrôle image ImgLogo
clGdip.ImgNew("logo").LoadControl Me.ImgLogo
' Dessine le logo à l'emplacement du contrôle ImgLogo
' Le blanc est transparent
' On active l'anti-aliasing
clGdip.DrawImg "logo", lx1, ly1, lx2, ly2, clGdip.img("logo").GetPixel(0, 0), Me.ImgLogo.PictureSizeMode, Me.ImgLogo.PictureAlignment
' Supprime l'image en mémoire
clGdip.ImgDelete "logo"
' Masque l'image qui a servi au positionnement du logo
Me.ImgLogo.Visible = False
' Décalage de 3 pixels pour l'ombre
lDecalage = 3
' Position du texte à l'emplacement du contrôle LblText
lx1 = clGdip.CtrlToImgX(Me.LblText.Left - Me.ImgFond.Left, Me.ImgFond)
ly1 = clGdip.CtrlToImgY(Me.LblText.Top - Me.ImgFond.Top, Me.ImgFond)
lx2 = clGdip.CtrlToImgX(Me.LblText.Left - Me.ImgFond.Left + Me.LblText.Width, Me.ImgFond)
ly2 = clGdip.CtrlToImgY(Me.LblText.Top - Me.ImgFond.Top + Me.LblText.Height, Me.ImgFond)
' Taille du texte en pixels
lSize = clGdip.FontSizeToPixel(Me.LblText.Font.size, Me.ImgFond)
' Affiche le texte décalé en noir et translucide (effet d'ombre)
clGdip.DrawText Me.LblText.Caption, lSize, Me.LblText.Font.name, _
       lDecalage + lx1, lDecalage + ly1, lDecalage + lx2, lDecalage + ly2, _
       , , vbBlack, 70, , , Me.LblText.Font.Italic, Me.LblText.Font.Bold, Me.LblText.Font.Underline, Me.LblText.Font.Strikethrough
' Affiche le texte
clGdip.DrawText Me.LblText.Caption, lSize, Me.LblText.Font.name, _
       lx1, ly1, lx2, ly2, _
       , , Me.LblText.ForeColor, , , , Me.LblText.Font.Italic, Me.LblText.Font.Bold, Me.LblText.Font.Underline, Me.LblText.Font.Strikethrough
' Masque l'étiquette qui a servi au positionnement du texte
Me.LblText.Visible = False
' Charge l'image du bouton à partir du contrôle image ImgButton
clGdip.ImgNew("flechebleu").LoadControl Me.ImgButton
' Clone l'image pour créer une flèche avec texte rouge
clGdip.ImgClone "flechebleu", "flecherouge"
' Change la couleur du texte bleu en rouge
clGdip.img("flecherouge").ReplaceColor vbBlue, vbRed
' Peint le bouton
PaintButton
' Masque l'image qui a servi au positionnement de la flèche
Me.ImgButton.Visible = False
' Crée une région qui contient les points non blanc de l'image
clGdip.CreateRegionFromColor "region", clGdip.GetPixel(1, 1)
' Applique cette région au formulaire
clGdip.SetFormRegion Me, "region", Me.ImgFond, , , -1, , 1.5
' Supprime la région
clGdip.RegionDelete "region"
' Dessine l'image dans le contrôle principal
clGdip.Repaint Me.ImgFond
' Limite le dessin avec FastRepaint à la position du contrôle ImgButton
clGdip.RepaintFastSetClipControl Me.ImgButton, True
' Affiche la date
Me.LblDate.Caption = Format(Date, "dddd dd mmmm yyyy")
Exit Sub
gestion_erreurs:
    MsgBox Err.description
End Sub

'---------------------------------------------------------------------------------------
' Dessin du bouton en forme de flèche
'---------------------------------------------------------------------------------------
Private Sub PaintButton(Optional pDown As Boolean, Optional pOver As Boolean)
 Dim lDecalage As Long
 Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
 Dim lImage As String
' Décalage de 3 pixels pour l'ombre
 lDecalage = 3
' Position du bouton
 lx1 = clGdip.CtrlToImgX(Me.ImgButton.Left - Me.ImgFond.Left, Me.ImgFond)
 ly1 = clGdip.CtrlToImgY(Me.ImgButton.Top - Me.ImgFond.Top, Me.ImgFond)
 lx2 = clGdip.CtrlToImgX(Me.ImgButton.Left - Me.ImgFond.Left + Me.ImgButton.Width, Me.ImgFond)
 ly2 = clGdip.CtrlToImgY(Me.ImgButton.Top - Me.ImgFond.Top + Me.ImgButton.Height, Me.ImgFond)
 ' Rempli un rectangle de la couleur de fond avant de redessiner le bouton
 clGdip.FillColor clGdip.GetPixel(lx1, ly1), , , lx1, ly1, lx2 + lDecalage, ly2 + lDecalage
 ' Texte rouge ou bleu
 If pOver Then
    lImage = "flecherouge"
 Else
    lImage = "flechebleu"
 End If
 ' Si pDown = Vrai, le bouton est enfoncé; on ne dessine que la flèche
 ' Sinon on dessine la flèche (dont une fois décalée et translucide pour un effet d'ombre)
 If Not pDown Then
    clGdip.DrawImg lImage, lx1 + lDecalage, ly1 + lDecalage, lx2 + lDecalage, ly2 + lDecalage, vbWhite, Me.ImgButton.PictureSizeMode, Me.ImgButton.PictureAlignment, 70
    clGdip.DrawImg lImage, lx1, ly1, lx2, ly2, vbWhite, Me.ImgButton.PictureSizeMode, Me.ImgButton.PictureAlignment, , "fleche"
 Else
    clGdip.DrawImg lImage, lx1 + lDecalage, ly1 + lDecalage, lx2 + lDecalage, ly2 + lDecalage, vbWhite, Me.ImgButton.PictureSizeMode, Me.ImgButton.PictureAlignment, , "fleche"
 End If
End Sub
