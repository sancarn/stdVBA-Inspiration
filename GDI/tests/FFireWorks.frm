VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FFireWorks 
   Caption         =   "Feu d'artifice"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985.001
   OleObjectBlob   =   "FFireWorks.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FFireWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal pMs As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal pMs As Long)
#End If

Private ogdi As clGdiplus  ' Classe pour graphisme
Private gStop As Boolean ' Flag pour arrêt de la boucle
Private gFires As Collection ' Collection de feux
Private gParts As Collection ' Collection de particules
Private gAcceleration As Long ' Acceleration verticale

'--------------------------------------------------------------
' Fermeture du formulaire
'--------------------------------------------------------------
Private Sub UserForm_Terminate()
' Demande l'arrêt de la boucle
gStop = True
End Sub

'--------------------------------------------------------------
' Chargement du formulaire
'--------------------------------------------------------------
Private Sub UserForm_Initialize()
' Nouvel objet pour graphisme
Set ogdi = New clGdiplus
' Initialise l'image
InitializeImage
' Initialise les collections
Set gFires = New Collection
Set gParts = New Collection
' Lance la boucle de manière asynchrone
Application.OnTime DateAdd("s", 1, Now), "RunFire"
End Sub

'--------------------------------------------------------------
' Rendu de l'image
'--------------------------------------------------------------
Private Sub Render()
Dim lFire As clFire
Dim lSubFire As clFire
Dim lPart As clPart
Dim lcpt As Long
Dim lTimer As Double
Dim lSpeed As Long
Dim lRndSubFire As Long
Static sTimerFire As Double
Static sTimer As Double
Static sNb As Long
Static sFPS As Long
On Error GoTo gestion_erreurs
' Fond Noir
ogdi.FillColor vbBlack
' Calcul FPS
lTimer = Timer
If lTimer - sTimer >= 1 Then
    sFPS = CLng((sNb + 1) / (lTimer - sTimer))
    sNb = 0
    sTimer = lTimer
Else
    sNb = sNb + 1
End If
' Affiche le nombre d'images par secondes
ogdi.DrawText "FPS: " & sFPS, 22, , 0, 0, ogdi.ImageWidth, ogdi.ImageHeight, 0, 0, vbWhite
' Affiche le nombre de particules
ogdi.DrawText "Particules: " & gParts.count, 22, , 0, 25, ogdi.ImageWidth, ogdi.ImageHeight, 0, 0, vbWhite
' Envoie un feu aléatoirement toutes les 0.7 à 1.7 secondes
If lTimer - sTimerFire > (0.7 + Rnd) Then
    ' Nouveau feu
    Set lFire = New clFire
    ' Position initiale du feu
    ' Aléatoire en X
    lFire.X0 = Rnd * (ogdi.ImageWidth * 80 / 100) + (ogdi.ImageWidth * 20 / 100)
    ' En bas en Y
    lFire.Y0 = 0
    ' Vitesse initiale horizontale
    ' Aléatoire
    lFire.SpeedX0 = Rnd * 0.3 * ogdi.ImageWidth
    ' Vers la droite ou vers la gauche en fonction de la position sur l'image
    If lFire.X0 > ogdi.ImageWidth / 2 Then lFire.SpeedX0 = -lFire.SpeedX0
    ' Vitesse initiale verticale
    ' Au moins une moitié d'image par seconde + un peu de vitesse aléatoire
    lFire.SpeedY0 = (0.5 * ogdi.ImageHeight) + (Rnd * (0.1 * ogdi.ImageHeight))
    ' Heure de déclenchement de l'explosion
    lFire.TimeUp = lTimer + 1 + Rnd
    ' Heure de lancement
    lFire.Timer = lTimer
    ' Couleur aléatoire
    Select Case Int(Rnd * 3) + 1
        Case 1
            lFire.Color = vbRed
        Case 2
            lFire.Color = vbBlue
        Case 3
            lFire.Color = vbYellow
    End Select
    ' Ajout du feu à la collection
    gFires.Add lFire, lFire.key
    ' Heure de lancement du dernier feu
    sTimerFire = lTimer
End If
' Déplace et affiche les feux
For Each lFire In gFires
    ' Position horizontale
    ' Position initiale + Vitesse Initiale * Temps écoulé + 1/2 * Accélération * (Temps écoulé)²
    lFire.X = lFire.X0 + lFire.SpeedX0 * (lTimer - lFire.Timer)
    ' Position verticale
    lFire.Y = lFire.Y0 + lFire.SpeedY0 * (lTimer - lFire.Timer) + 0.5 * gAcceleration * (lTimer - lFire.Timer) * (lTimer - lFire.Timer)
    ' Si explosion du feu
    If lTimer > lFire.TimeUp Then
        ' Si c'est un sous-feux => explosion en particules
        ' Sinon, calcul un nombre aléatoire => explosion en sous-feux une fois sur quatre
        If lFire.SubFire Then
            lRndSubFire = 1
        Else
            lRndSubFire = Int(Rnd * 4) + 1
        End If
        ' Explosion en sous-feux ou en particules
        Select Case lRndSubFire
            Case Is <= 3
                ' Crée des particules (au moins 100 + nombre aléatoire entre 0 et 200)
                For lcpt = 1 To 100 + Rnd * 200
                    ' Nouvelle particule
                    Set lPart = New clPart
                    
                    ' Position initiale de la particule = la position du feu
                    lPart.X0 = lFire.X
                    lPart.Y0 = lFire.Y
                    ' Vitesse initiale globale de la particule
                    lSpeed = Rnd * ogdi.ImageHeight
                    ' Vitesse initiale horizontale
                    lPart.SpeedX0 = Rnd * lSpeed - lSpeed / 2
                    ' Calcul de la vitesse initiale verticale
                    ' La résultante des deux vitesses doit être égale à la vitesse globale
                    ' Théorème de Pythagore
                    lPart.SpeedY0 = Sqr(lSpeed * lSpeed / 4 - lPart.SpeedX0 * lPart.SpeedX0)
                    ' Vitesse vertical aléatoirement vers le haut ou vers le bas
                    lPart.SpeedY0 = lPart.SpeedY0 * IIf(Rnd * 10 > 5, 1, -1)
                    ' Heure de disparition de la particule
                    lPart.TimeUp = lTimer + 1 + Rnd * 2
                    ' Heure de lancement de la particule
                    lPart.Timer = lTimer
                    ' Couleur de la particule
                    lPart.Color = lFire.Color
                    ' Ajout de la particule à la collection
                    gParts.Add lPart, lPart.key
                Next
            Case 4
                ' Crée des sous-feux qui exploseront en particules
                ' Vitesse initiale globale des sous-feux
                lSpeed = ogdi.ImageHeight / 5 + Rnd * ogdi.ImageHeight / 5
                For lcpt = 1 To 6
                    ' Nouveau feu
                    Set lSubFire = New clFire
                    ' Position initiale du sous-feux = la position du feu
                    lSubFire.X0 = lFire.X
                    lSubFire.Y0 = lFire.Y
                    ' Vitesse initiale horizontale
                    lSubFire.SpeedX0 = lSpeed * Cos(3.14 * 2 / 6 * lcpt)
                    'Vitesse initiale verticale
                    lSubFire.SpeedY0 = lSpeed * Sin(3.14 * 2 / 6 * lcpt)
                    ' Vitesse vertical aléatoirement vers le haut ou vers le bas
                    'lPart.SpeedY0 = lPart.SpeedY0 * IIf(Rnd * 10 > 5, 1, -1)
                    ' Heure d'explosion du sous-feux
                    lSubFire.TimeUp = lTimer + 0.8 + Rnd * 0.2
                    ' lSubFire de lancement du sous-feux
                    lSubFire.Timer = lTimer
                    ' Couleur du sous-feux
                    lSubFire.Color = lFire.Color
                    ' Définit le flag "sous-feux"
                    lSubFire.SubFire = True
                    ' Ajout du sous-feux à la collection
                    gFires.Add lSubFire, lSubFire.key
                Next
        End Select
        ' Supprime le feu
        gFires.Remove lFire.key
    Else
        ' Dessine le feu
        ' Cercle de 2 pixel de rayon
        ogdi.DrawEllipse lFire.X, ogdi.ImageHeight - lFire.Y, 2, 2, 1, lFire.Color, lFire.Color
    End If
Next
' Déplace et affiche les particules
  For Each lPart In gParts
      ' Position horizontale
      lPart.X = lPart.X0 + lPart.SpeedX0 * (lTimer - lPart.Timer)
      ' Position verticale
      lPart.Y = lPart.Y0 + lPart.SpeedY0 * (lTimer - lPart.Timer) + 0.5 * gAcceleration * (lTimer - lPart.Timer) * (lTimer - lPart.Timer)
      
      ' Si disparition de la particule
      If lTimer > lPart.TimeUp Then
          ' Supprime la particule
          gParts.Remove lPart.key
      Else
          ' Dessine la particule
          ' 1 pixel pour une particule
          ogdi.DrawPixel lPart.X, ogdi.ImageHeight - lPart.Y, lPart.Color
          
      End If
  Next
Exit Sub
gestion_erreurs:
gStop = True
End Sub

'--------------------------------------------------------------
' Redimensionnement du formulaire
'--------------------------------------------------------------
Private Sub UserForm_Resize()
' Force le recalcul du repaint
ogdi.RepaintFastResetCalc = True
' Initilise l'image
InitializeImage
End Sub

'--------------------------------------------------------------
' Initialise l'image
'--------------------------------------------------------------
Private Sub InitializeImage()
On Error GoTo gestion_erreurs
' Réduit l'image
Me.Image0.Width = 0
Me.Image0.Height = 0
Me.Image0.Left = 0
Me.Image0.Top = 0
' Redimensionne l'image
Me.Image0.Width = Me.InsideWidth
Me.Image0.Height = Me.InsideHeight
' Crée un nouveau bitmap
ogdi.CreateBitmapForControl Me.Image0
' Remplit de noir
ogdi.FillColor vbBlack
' Dessine dans le contrôle
ogdi.RepaintNoFormRepaint Me.Image0
' Calcul de l'accélération verticale en pixel par secondes au carré
' On défini la hauteur de l'image à 50 mètres
' 9.81 est la gravité
gAcceleration = -(ogdi.ImageHeight / 50 * 9.81)
Exit Sub
gestion_erreurs:
gStop = True
End Sub

'--------------------------------------------------------------
' Affiche l'image
'--------------------------------------------------------------
Private Sub Display()
On Error GoTo gestion_erreurs
' Si demande d'arrêt => on quitte
If gStop Then Exit Sub
' Redessine l'image sur le contrôle
ogdi.RepaintFast Me.Image0
Exit Sub
gestion_erreurs:
gStop = True
End Sub

'--------------------------------------------------------------
' Boucle d'affichage
'--------------------------------------------------------------
Public Sub RunFire()
On Error GoTo gestion_erreurs
Do
    ' Ajouter éventuellement une attente en millisecondes
    'Sleep 5
    ' Exécute la pile d'événements pour ne pas bloquer l'application
    DoEvents
    ' Si demande d'arrêt => on quitte
    If gStop Then Exit Sub
    Render
    Display
Loop
Exit Sub
gestion_erreurs:
gStop = True
End Sub

'--------------------------------------------------------------
' Sur click sur l'image
'--------------------------------------------------------------
Private Sub Image0_Click()
On Error GoTo gestion_erreurs:
If MsgBox("Voulez-vous visitez mon site internet?", vbYesNo Or vbDefaultButton2) = vbYes Then
    Application.FollowHyperlink "http://arkham46.developpez.com"
End If
Exit Sub
gestion_erreurs:
MsgBox "Erreur d'ouverture du lien internet"
End Sub
