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
'*                               D�mo SplashScreen                                     *
'***************************************************************************************

'***************************************************************************************

Private clGdip As clGdiplus    ' Classe image

'---------------------------------------------------------------------------------------
' Souris appuy�e sur image
'---------------------------------------------------------------------------------------
Private Sub ImgFond_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
    ' R�gion sous le curseur de la souris
    If Not clGdip Is Nothing Then
        lregion = clGdip.GetRegionXY(clGdip.CtrlToImgX(X, Me.ImgFond), clGdip.CtrlToImgY(Y, Me.ImgFond))
    End If
    
    If lregion = "fleche" Then
        ' Si on clique sur la fl�che
        ' Dessine la fl�che en position enfonc�e avec texte rouge
        PaintButton True, True
        ' Redessine l'image
        clGdip.RepaintFast Me.ImgFond
    Else
        ' Sinon on d�place le formulaire
        clGdip.DragForm Me
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Souris d�plac�e sur l'image
'---------------------------------------------------------------------------------------
Private Sub ImgFond_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
Static sRegion As String
    ' R�gion sous le curseur de la souris
    If Not clGdip Is Nothing Then
        lregion = clGdip.GetRegionXY(clGdip.CtrlToImgX(X, Me.ImgFond), clGdip.CtrlToImgY(Y, Me.ImgFond))
    End If
    ' Si changement de r�gion survol�e
    If sRegion <> lregion Then
        ' Redessine la fl�che
        ' En position appuy�e si le bouton gauche est appuy�
        ' Avec couleur rouge si on survole la fl�che
        PaintButton (Button = vbKeyLButton), (lregion = "fleche")
        ' Redessine l'image
        clGdip.RepaintFast Me.ImgFond
        ' Converse le nom de la derni�re r�gion survol�e
        sRegion = lregion
    End If
End Sub

'---------------------------------------------------------------------------------------
' Souris rel�ch�e sur l'image
'---------------------------------------------------------------------------------------
Private Sub ImgFond_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
    ' R�gion sous le curseur de la souris
    If Not clGdip Is Nothing Then
        lregion = clGdip.GetRegionXY(clGdip.CtrlToImgX(X, Me.ImgFond), clGdip.CtrlToImgY(Y, Me.ImgFond))
    End If
    
    If lregion <> "fleche" Then
        ' Si on survole la fl�che
        ' On repaint le bouton en position rel�ch�e
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
' Lib�re la classe image
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
' Charge l'image du contr�le
clGdip.LoadControl Me.ImgFond
' Position du logo
lx1 = clGdip.CtrlToImgX(Me.ImgLogo.Left - Me.ImgFond.Left, Me.ImgFond)
ly1 = clGdip.CtrlToImgY(Me.ImgLogo.Top - Me.ImgFond.Top, Me.ImgFond)
lx2 = clGdip.CtrlToImgX(Me.ImgLogo.Left - Me.ImgFond.Left + Me.ImgLogo.Width, Me.ImgFond)
ly2 = clGdip.CtrlToImgY(Me.ImgLogo.Top - Me.ImgFond.Top + Me.ImgLogo.Height, Me.ImgFond)
' Charge l'image du logo � partir du contr�le image ImgLogo
clGdip.ImgNew("logo").LoadControl Me.ImgLogo
' Dessine le logo � l'emplacement du contr�le ImgLogo
' Le blanc est transparent
' On active l'anti-aliasing
clGdip.DrawImg "logo", lx1, ly1, lx2, ly2, clGdip.img("logo").GetPixel(0, 0), Me.ImgLogo.PictureSizeMode, Me.ImgLogo.PictureAlignment
' Supprime l'image en m�moire
clGdip.ImgDelete "logo"
' Masque l'image qui a servi au positionnement du logo
Me.ImgLogo.Visible = False
' D�calage de 3 pixels pour l'ombre
lDecalage = 3
' Position du texte � l'emplacement du contr�le LblText
lx1 = clGdip.CtrlToImgX(Me.LblText.Left - Me.ImgFond.Left, Me.ImgFond)
ly1 = clGdip.CtrlToImgY(Me.LblText.Top - Me.ImgFond.Top, Me.ImgFond)
lx2 = clGdip.CtrlToImgX(Me.LblText.Left - Me.ImgFond.Left + Me.LblText.Width, Me.ImgFond)
ly2 = clGdip.CtrlToImgY(Me.LblText.Top - Me.ImgFond.Top + Me.LblText.Height, Me.ImgFond)
' Taille du texte en pixels
lSize = clGdip.FontSizeToPixel(Me.LblText.Font.size, Me.ImgFond)
' Affiche le texte d�cal� en noir et translucide (effet d'ombre)
clGdip.DrawText Me.LblText.Caption, lSize, Me.LblText.Font.name, _
       lDecalage + lx1, lDecalage + ly1, lDecalage + lx2, lDecalage + ly2, _
       , , vbBlack, 70, , , Me.LblText.Font.Italic, Me.LblText.Font.Bold, Me.LblText.Font.Underline, Me.LblText.Font.Strikethrough
' Affiche le texte
clGdip.DrawText Me.LblText.Caption, lSize, Me.LblText.Font.name, _
       lx1, ly1, lx2, ly2, _
       , , Me.LblText.ForeColor, , , , Me.LblText.Font.Italic, Me.LblText.Font.Bold, Me.LblText.Font.Underline, Me.LblText.Font.Strikethrough
' Masque l'�tiquette qui a servi au positionnement du texte
Me.LblText.Visible = False
' Charge l'image du bouton � partir du contr�le image ImgButton
clGdip.ImgNew("flechebleu").LoadControl Me.ImgButton
' Clone l'image pour cr�er une fl�che avec texte rouge
clGdip.ImgClone "flechebleu", "flecherouge"
' Change la couleur du texte bleu en rouge
clGdip.img("flecherouge").ReplaceColor vbBlue, vbRed
' Peint le bouton
PaintButton
' Masque l'image qui a servi au positionnement de la fl�che
Me.ImgButton.Visible = False
' Cr�e une r�gion qui contient les points non blanc de l'image
clGdip.CreateRegionFromColor "region", clGdip.GetPixel(1, 1)
' Applique cette r�gion au formulaire
clGdip.SetFormRegion Me, "region", Me.ImgFond, , , -1, , 1.5
' Supprime la r�gion
clGdip.RegionDelete "region"
' Dessine l'image dans le contr�le principal
clGdip.Repaint Me.ImgFond
' Limite le dessin avec FastRepaint � la position du contr�le ImgButton
clGdip.RepaintFastSetClipControl Me.ImgButton, True
' Affiche la date
Me.LblDate.Caption = Format(Date, "dddd dd mmmm yyyy")
Exit Sub
gestion_erreurs:
    MsgBox Err.description
End Sub

'---------------------------------------------------------------------------------------
' Dessin du bouton en forme de fl�che
'---------------------------------------------------------------------------------------
Private Sub PaintButton(Optional pDown As Boolean, Optional pOver As Boolean)
 Dim lDecalage As Long
 Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
 Dim lImage As String
' D�calage de 3 pixels pour l'ombre
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
 ' Si pDown = Vrai, le bouton est enfonc�; on ne dessine que la fl�che
 ' Sinon on dessine la fl�che (dont une fois d�cal�e et translucide pour un effet d'ombre)
 If Not pDown Then
    clGdip.DrawImg lImage, lx1 + lDecalage, ly1 + lDecalage, lx2 + lDecalage, ly2 + lDecalage, vbWhite, Me.ImgButton.PictureSizeMode, Me.ImgButton.PictureAlignment, 70
    clGdip.DrawImg lImage, lx1, ly1, lx2, ly2, vbWhite, Me.ImgButton.PictureSizeMode, Me.ImgButton.PictureAlignment, , "fleche"
 Else
    clGdip.DrawImg lImage, lx1 + lDecalage, ly1 + lDecalage, lx2 + lDecalage, ly2 + lDecalage, vbWhite, Me.ImgButton.PictureSizeMode, Me.ImgButton.PictureAlignment, , "fleche"
 End If
End Sub
