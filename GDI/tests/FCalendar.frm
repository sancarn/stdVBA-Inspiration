VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCalendar 
   Caption         =   "FCalendar"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   OleObjectBlob   =   "FCalendar.frx":0000
End
Attribute VB_Name = "FCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Instance de clGdiplus pour le graphisme
Private WithEvents ogdi As clGdiplus
Attribute ogdi.VB_VarHelpID = -1
' Année en cours
Private gYear As Long

' Bouton année - 1
Private Sub btnYearDown_Click()
Me.txtYear.value = gYear - 1
txtYear_AfterUpdate
End Sub

' Bouton année + 1
Private Sub btnYearUp_Click()
Me.txtYear.value = gYear + 1
txtYear_AfterUpdate
End Sub

' Sur chargement du formulaire
Private Sub UserForm_Initialize()
' Formulaire plein écran
Me.Top = Application.Top
Me.Left = Application.Left
Me.Width = Application.Width
Me.Height = Application.Height
' Réduit l'image
Me.Image0.Width = 0
Me.Image0.Height = 0
' Redimensionne l'image
Me.Image0.Width = -Me.Image0.Left + Me.InsideWidth
Me.Image0.Height = -Me.Image0.Top + Me.InsideHeight
' Crée une instance de la classe de graphisme
Set ogdi = New clGdiplus
' Crée une image à la taille du contrôle Image0
ogdi.CreateBitmapForControl Me.Image0
' Chargement de l'icône task
ogdi.ImgNew("task").LoadControl Me.imgTask
' Taille texte maxi et mini
ogdi.MinTextSize = 10
ogdi.MaxTextSize = 14
' Année = année courante par défaut
gYear = Year(Date)
Me.txtYear.value = gYear
' Dessine le calendrier
DrawCalendar
' Dessine l'image dans le contrôle
ogdi.Repaint Me.Image0
End Sub

' Click sur l'image
Private Sub Image0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
' Recherche du nom de la région cliquée
lregion = ogdi.GetRegionXY(X, Y, , , Me.Image0)
' Affichage d'un message si région non vide
If lregion <> "" Then
    ' Le nom de la région est la date en numérique
    ' CDate converti au format Date/Heure
    MsgBox "Click sur " & CDate(lregion)
End If
End Sub

' Déplacement sur l'image
Private Sub Image0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lregion As String
Static sRegion As String
Dim lx1 As Long, lx2 As Long, ly1 As Long, ly2 As Long
Dim lTextRdv As String
' Recherche du nom de la région survolée
lregion = ogdi.GetRegionXY(X, Y, , , Me.Image0)
' Vérifie que la région survolé a changé
' Sinon on ne fait rien
If lregion = sRegion Then
    Exit Sub
Else
    sRegion = lregion
End If
' Rétabli l'image du calendrier
ogdi.ImageReset
' Encadre le jour survolé
ogdi.RegionFrame lregion, vbBlue, 1
' Redessine l'image (non persistante)
ogdi.RepaintFast Me.Image0
End Sub

' Dessin du calendrier
Private Sub DrawCalendar()
Dim lMonth As Long
Dim lDay As Long
Dim lColSize As Long
Dim lRowSize As Long
Dim lDate As Date
Dim lDayText As String
Dim lCalWidth As Long
Dim lCalHeight As Long
' Taille des colonnes et lignes
lColSize = (ogdi.ImageWidth - 13) \ 12 ' 12 Mois - 13 lignes de séparation
lRowSize = (ogdi.ImageHeight - 33) \ (31 + 1) ' 31 jours + 1 en-tête - 33 lignes de séparation
' Taille du calendrier
lCalWidth = lColSize * 12 + 13
lCalHeight = lRowSize * (31 + 1) + 33
' Remplit l'image de la couleur de fond du formulaire
ogdi.FillColor Me.BackColor
' Supprime les regions
ogdi.RegionsDelete
' Entourage du calendrier et le colorie en blanc
ogdi.DrawRectangle 0, 0, lCalWidth - 1, lCalHeight - 1, vbWhite
' Ligne horizontale de séparation de l'en-tête
ogdi.DrawLine 0, lRowSize, lCalWidth - 1, lRowSize
' Pour chaque mois de 1 à 12
For lMonth = 1 To 12
    ' Ligne verticale de séparation
    ogdi.DrawLine lMonth * (lColSize + 1), 0, lMonth * (lColSize + 1), lCalHeight - 1
    ' Libellé du mois
    ogdi.DrawText UCase(MonthName(lMonth)), 6, , (lMonth - 1) * (lColSize + 1), 0, lMonth * (lColSize + 1), lRowSize, , , , , , , , True
    ' Pour chaque jour de 1 à 31
    For lDay = 1 To 31
        ' lDate est au format Date/Heure
        lDate = DateSerial(gYear, lMonth, lDay)
        ' Si le jour du mois est trop élevé (par ex 31 pou février), le jour
        '   ne correspond pas => on ne l'affiche pas
        If Day(lDate) = lDay Then
            ' Crée un région encadrant le jour
            ' Le nom de la région est la date en numérique
            ogdi.CreateRegionRect CLng(lDate), (lMonth - 1) * (lColSize + 1), lDay * (lRowSize + 1), lMonth * (lColSize + 1), (lDay + 1) * (lRowSize + 1)
            ' Si WeekEnd
            If Weekday(lDate, vbMonday) > 5 Then
                ' Colore en rose
                ogdi.FillColor RGB(255, 200, 200), , , 1 + (lMonth - 1) * (lColSize + 1), lDay * (lRowSize + 1), -1 + lMonth * (lColSize + 1), -1 + (lDay + 1) * (lRowSize + 1)
            End If
            ' Si Dimanche
            If Weekday(lDate, vbMonday) = 7 Then
                ' Ligne horizontale de séparation de la semaine
                ogdi.DrawLine (lMonth - 1) * (lColSize + 1), (lDay + 1) * (lRowSize + 1), lMonth * (lColSize + 1), (lDay + 1) * (lRowSize + 1)
            End If
            ' Si date du jour
            If lDate = Date Then
                ' Colore en rouge
                ogdi.FillColor vbRed, , , 1 + (lMonth - 1) * (lColSize + 1), 1 + lDay * (lRowSize + 1), -1 + lMonth * (lColSize + 1), -1 + (lDay + 1) * (lRowSize + 1)
            End If
            ' Texte =  le numéro de jour et la première lettre du nom du jour
            lDayText = lDay & " " & Left(UCase(WeekdayName(Weekday(lDate, vbMonday))), 1)
            ' Dessine le texte (marge de 3px à gauche)
            ogdi.DrawText lDayText, 10, , 3 + 1 + (lMonth - 1) * (lColSize + 1), 1 + lDay * (lRowSize + 1), -1 + lMonth * (lColSize + 1), -1 + (lDay + 1) * (lRowSize + 1), HorzAlignLeft
        End If
    Next
Next
' Conserve l'image générée pour la récupérer rapidement ensuite
ogdi.ImageKeep
End Sub

' Mise à jour de l'année dans la zone de texte
Private Sub txtYear_AfterUpdate()
' Met à jour l'année dans la variable
If IsNull(txtYear.value) Then
    gYear = Year(Now)
Else
    gYear = val(txtYear.value)
End If
' Redessine le calendrier
DrawCalendar
' Affiche dans le contrôle
' RepaintControlNoFormRepaint est un dessin persistant
'         mais il minimise les scintillements
ogdi.RepaintNoFormRepaint Me.Image0
End Sub

