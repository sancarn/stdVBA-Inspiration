VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FHorloge 
   Caption         =   "Horloge"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   OleObjectBlob   =   "FHorloge.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FHorloge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Déclaration de l'instance de classe pour l'horloge
Private WithEvents gClock As ClClock
Attribute gClock.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Sur fermeture du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
' Libération de la classe à la fermeture du formulaire
If Not gClock Is Nothing Then Set gClock = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Sur chargement du formulaire
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
' Création de l'instance de classe qui gère l'horloge
Set gClock = New ClClock
' Initialisation des infos pour l'horloge
gClock.SetClockCtrl Me.Image0, Me.TxtTime
gClock.DisplayDigit = True
gClock.DigitBold = True
gClock.DisplaySecond = True
gClock.ClockDateTime = Now
gClock.MoveWithMouse = True
' Région transparente
gClock.gGdip.CreateRegionFromColor "region", gClock.BackColor, , False
gClock.gGdip.SetFormRegion Me, "region", Me.Image0
gClock.gGdip.RegionDelete "region"
gClock.gGdip.Repaint Me.Image0
' Fond des zones de texte
Me.LblDate.BackColor = gClock.ClockColor
Me.TxtTime.BackColor = gClock.ClockColor
' Ne redessine pas par dessus les zones de texte
gClock.gGdip.RepaintFastSetClipControl Me.LblDate, False, -1, -1, 1, 1
gClock.gGdip.RepaintFastSetClipControl Me.TxtTime, False, -1, -1, 2, 1
End Sub

'---------------------------------------------------------------------------------------
' Date/Heure modifiée
'---------------------------------------------------------------------------------------
Private Sub gClock_DateTimeChanged(pDateTime As Date)
Me.LblDate.Caption = Format(pDateTime, "dddd dd mmmm yyyy")
End Sub

'---------------------------------------------------------------------------------------
' Date/heure en cours de modification (aiguilles en déplacement)
'---------------------------------------------------------------------------------------
Private Sub gClock_DateTimeChanging(pDateTime As Date)
Me.LblDate.Caption = Format(pDateTime, "dddd dd mmmm yyyy")
End Sub

'---------------------------------------------------------------------------------------
' Souris appuyée sur horloge
'---------------------------------------------------------------------------------------
Private Sub gClock_MouseDown(pRegion As String, Button As Integer, Shift As Integer)
' Si bouton gauche et souris n'est pas sur une aiguille => on déplace le formulaire
If Button = vbKeyLButton And pRegion = "" Then gClock.gGdip.DragForm Me
End Sub

'---------------------------------------------------------------------------------------
' Double-click sur l'image
'---------------------------------------------------------------------------------------
Private Sub Image0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Fermeture du formulaire
Unload Me
End Sub
