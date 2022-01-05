VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FShapedForm 
   Caption         =   "Forme personnalisée"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   OleObjectBlob   =   "FShapedForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FShapedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private ogdi As clGdiplus    ' Classe image

'---------------------------------------------------------------------------------------
' Bouton de fermeture du formulaire
'---------------------------------------------------------------------------------------
Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ogdi.DragForm Me
End Sub

Private Sub FrameVisible_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ogdi.DragForm Me
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
Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
    ' Initialisation de la classe
    Set ogdi = New clGdiplus
    ' Création d'une région rectangulaire correspondant au contrôle FrameVisible
    ogdi.GetControlPos Me.FrameVisible, lx1, ly1, lx2, ly2, True
    ogdi.CreateRegionRect "regionvisible", lx1, ly1, lx2 + 1, ly2 + 1
    ' Création d'une région rectangulaire correspondant au contrôle FrameEmpty
    ogdi.GetControlPos Me.FrameEmpty, lx1, ly1, lx2, ly2
    ogdi.CreateRegionRect "regionempty", lx1, ly1, lx2, ly2
    ' Retire FrameEmpty de la region visible
    ogdi.RegionCombine "regionvisible", "regionempty", CombineModeExclude
    ' Applique la région au formulaire pour réduire l'affichage à la région "regionvisible"
    ogdi.SetFormRegion Me, "regionvisible"
    ' Supprime les régions
    ogdi.RegionDelete "regionempty"
    ogdi.RegionDelete "regionvisible"
End Sub



