VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMenu 
   Caption         =   "Menu"
   ClientHeight    =   8670.001
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   OleObjectBlob   =   "FMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************
'*                               Démo Menu avec Images                                 *
'***************************************************************************************
Private WithEvents clGdi As clGdiplus    ' Classe Image
Attribute clGdi.VB_VarHelpID = -1
Private clExpl As ClFormatText ' Classe pour texte formaté

Private Const cEspaceX As Long = 6    ' Espacement X entre les vignettes
Private Const cEspaceY As Long = 15    ' Espacement Y entre les vignettes

Private gTaille As Long   ' Taille de chaque vignette
Private gNbX As Long, gNbY As Long ' Nombre d'images par ligne/colonne
Private gWidth As Long, gHeight As Long ' Taille de l'image

' Texte d'explication par défaut
Private Const cDefaultCaption As String = "<font bold=true color=255 size=12><font backcolor=16768220>Bienvenue dans le formulaire d'exemples de la classe ClGdiplus</font></font>" & vbCrLf & _
        "<font color=16711680 bold=true href=http://arkham46.developpez.com/articles/access/ClGdiplus>Téléchargez la dernière version et la documentation</font>" & vbCrLf & _
        "<font color=16711680 bold=true href=http://www.developpez.net/forums/forumdisplay.php?f=45>Forum Access developpez.com</font>"

'---------------------------------------------------------------------------------------
' Mise à jour taille vignette
'---------------------------------------------------------------------------------------
Private Sub CmbTailleVignettes_Change()
    gTaille = CmbTailleVignettes.value
    ScaleBars
    DisplayMenu
End Sub

'---------------------------------------------------------------------------------------
' Sur fermeture du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    ' On libère les classes
    If Not clGdi Is Nothing Then Set clGdi = Nothing
    If Not clExpl Is Nothing Then Set clExpl = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Sur chargement du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim lFormCadreLeftOld As Long
    Dim lCtrl As Variant
    Dim lcpt As Long
    Dim lLine As Long
    ' Initialise la classe
    On Error GoTo gestion_erreurs
    Set clGdi = New clGdiplus
    ' Initialise l'image
    'gWidth = clGdi.PointsToPixelsX(Me.Image0.Width)
    'gHeight = clGdi.PointsToPixelsY(Me.Image0.Height)
    clGdi.CreateBitmapForControl Me.Image0 ' gWidth, gHeight
    gWidth = clGdi.ImageWidth
    gHeight = clGdi.ImageHeight
    ' Remplit l'image de la couleur du fond
    clGdi.FillColor Me.BackColor
    ' Applique l'image (blanche) dans le contrôle
    clGdi.Repaint Me.Image0
    ' Rempli la liste des tailles pour Excel
    For lcpt = 10 To 250 Step 10
        CmbTailleVignettes.AddItem lcpt
    Next
    ' Taille Vignettes
    gTaille = CmbTailleVignettes.value
    ' Charge les images
    With ThisWorkbook.Sheets("Menu")
    For lLine = .UsedRange.Row To .UsedRange.Row + .UsedRange.Rows.count - 1
        ' Ajoute une image à la liste, de largeur cTaille
        On Error Resume Next
        If Not clGdi.ImgNew(.Cells(lLine, 1).value).LoadControl(Me.Controls(.Cells(lLine, 1).value)) Then
            With clGdi.ImgNew(.Cells(lLine, 1).value)
                .CreateBitmap gTaille, gTaille
                .FillColor vbWhite
                .DrawText "Pas d'image", gTaille / 8, , 0, 0, gTaille, gTaille
            End With
        End If
        On Error GoTo gestion_erreurs
        'clGdi.ImageListResize .Cells(lLine, 1).Value, gTaille, , True
    Next
    End With
    ' Affiche le menu
    ScaleBars
    DisplayMenu
    ' Image pour explications
    Set clExpl = New ClFormatText
    clExpl.BackColor = vbWhite
    clExpl.ActiveURL = True
    clExpl.BackColorGradient = Me.BackColor
    clExpl.Text = cDefaultCaption
    clExpl.Control = Me.ImgExplications
    clExpl.DrawFormattedText

gestion_erreurs:
    If Err.Number <> 0 Then MsgBox Err.description
End Sub

'---------------------------------------------------------------------------------------
' Dimensionne les barres de défilement
'---------------------------------------------------------------------------------------
Private Sub ScaleBars()
    Dim lNbForm As Long
    ' Entrées de menu
    With ThisWorkbook.Sheets("Menu")
        lNbForm = .UsedRange.Rows.count
    End With
    clGdi.BarNew
    gNbX = (clGdi.PointsToPixelsX(Me.Image0.Width) - clGdi.BarSize - cEspaceX) \ (gTaille + cEspaceX)
    gNbY = (lNbForm \ gNbX) - ((lNbForm Mod gNbX) > 0)
    gWidth = clGdi.PointsToPixelsX(Me.Image0.Width)
    gHeight = clGdi.PointsToPixelsX(Me.Image0.Height)  ' cEspaceY + lNbY * (cEspaceY + gTaille) + cEspaceY
    clGdi.BarScaleX 0, 1
    clGdi.BarScaleY cEspaceY + gNbY * (cEspaceY + gTaille), 1
    clGdi.BarObject = Me.Image0
End Sub

'---------------------------------------------------------------------------------------
' Affiche le menu
'---------------------------------------------------------------------------------------
Private Sub DisplayMenu(Optional pFastRepaint As Boolean)
    Dim lLine As Long
    Dim lX As Long
    Dim lY As Long
    On Error GoTo gestion_erreurs
    ' Curseur d'attente (horloge)
    Me.MousePointer = fmMousePointerHourGlass
    'gWidth = Clgdi.PointsToPixelsX(Me.Image0.Width)
    'gHeight = cEspaceY + lNbY * (cEspaceY + gTaille) + cEspaceY
    ' On Error Resume Next
    ' Me.Image0.Height = Clgdi.PixelToPointsY(gHeight)
    ' On Error GoTo gestion_erreurs
    clGdi.CreateBitmapForControl Me.Image0 ' gWidth, gHeight
    ' Rempli l'image de blanc
    ' Remplit l'image de la couleur du fond
    clGdi.FillColor Me.BackColor
    ' On laisse un espace vertical avant de commencer à dessiner
    lY = cEspaceY
    ' On parcourt la table TMenu (ou la feuille Menu pour Excel)
    With ThisWorkbook.Sheets("Menu")
    For lLine = .UsedRange.Row To .UsedRange.Row + .UsedRange.Rows.count - 1
        ' Retour à la ligne si on dépasse l'image à droite
        If lX + cEspaceX + gTaille > gWidth - clGdi.BarSize Then
            lX = 0
            lY = lY + cEspaceY + gTaille
        End If
        lX = lX + cEspaceX
        ' Dessine l'image
        ' et ajoute une region correspondant à l'image avec le nom du formulaire en identifiant
        clGdi.WorldPush
        clGdi.WorldTranslate clGdi.BarStartX, clGdi.BarStartY
        clGdi.DrawImg .Cells(lLine, 1).value, lX, lY, lX + gTaille, lY + gTaille, Me.BackColor, , , , .Cells(lLine, 1).value
        clGdi.MaxTextSize = 16
        clGdi.MinTextSize = 6
        clGdi.DrawText .Cells(lLine, 2).value, 12, , lX, lY + gTaille, lX + gTaille, lY + gTaille + cEspaceY, , , RGB(0, 50, 100)
        clGdi.WorldPop
        ' On avance d'une image vers la droite
        lX = lX + gTaille
    Next
    End With
    'Dessin des barres
    clGdi.BarDraw
    ' Dessin définitif dans le contrôle
    If pFastRepaint Then
        ' Dessin rapide
        clGdi.RepaintFast Me.Image0
    Else
        ' Dessin définitif dans le contrôle
        clGdi.Repaint Me.Image0
    End If
    ' Conserve le menu de base avec les photos en noir et blanc
    clGdi.ImageKeep "Tampon"
gestion_erreurs:
    ' Réinitialisation du curseur
    Me.MousePointer = fmMousePointerDefault
    If Err.Number <> 0 Then MsgBox Err.description
End Sub

'---------------------------------------------------------------------------------------
' Click sur image (MouseDown permet d'avoir les coordonnées à la différence de Click)
'---------------------------------------------------------------------------------------
Private Sub clGdi_BarMouseDown(BarName As String, lregion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo gestion_erreurs
    If Not clGdi Is Nothing Then    ' On vérifie que la classe est initialisée
        If Button = vbKeyLButton Then
            If lregion <> "" Then
                ' Ouvre le formulaire correspondant
                On Error Resume Next
                VBA.UserForms.Add(lregion).Show
                On Error GoTo gestion_erreurs
            End If
        End If
    End If
gestion_erreurs:
    If Err.Number <> 0 Then MsgBox Err.description
End Sub

'---------------------------------------------------------------------------------------
' Sur déplacement de la souris
'---------------------------------------------------------------------------------------
' Modifie le curseur et encadre de rouge l'image survolée par la souris
'---------------------------------------------------------------------------------------
Private Sub clGdi_BarMouseMove(BarName As String, lregion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static OldRegion As String    ' Région lors du précédent appel de cette fonction
    Dim lcpt As Long
    On Error GoTo gestion_erreurs
    If clGdi Is Nothing Then Exit Sub    ' On vérifie que la classe est initialisée
    If OldRegion <> lregion Then    ' Si on a changé de région
        If lregion <> "" Then
            For lcpt = 1 To 1 'gTaille Step 5
            ' Récupère le menu sans encadrement
            clGdi.ImageReset "Tampon"
            ' Dessine un cadre autour de la region
            clGdi.RegionFrame lregion, vbRed, 2
            ' Dessine l'image
            clGdi.RepaintFast Me.Image0
            Next
            ' Applique les modification au contrôle
            clGdi.RepaintFast Me.Image0
            ' Explications
            ' Mise à jour de l'explication
            clExpl.Text = ThisWorkbook.Sheets("Menu").UsedRange.Columns(1).Find(lregion, , , , , xlNext, False).offset(0, 2).value
            clExpl.Text = Replace(clExpl.Text, vbLf, vbCrLf)
            clExpl.DrawFormattedText
        ElseIf lregion = "" Then
            ' Si pas de région sous le curseur on rétablit le menu sans encadrement
            clGdi.ImageReset "Tampon"
            clGdi.RepaintFast Me.Image0
            ' Mise à jour de l'explication
            clExpl.Text = cDefaultCaption
            clExpl.DrawFormattedText
        End If
    End If
    OldRegion = lregion    ' Sauvegarde la valeur de la région survolée
gestion_erreurs:
    If Err.Number <> 0 Then MsgBox Err.description
End Sub

Private Sub clGdi_BarOnRefreshNeeded(BarName As String, MouseUp As Boolean)
DisplayMenu True
End Sub



