VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FFormatText 
   Caption         =   "Textes formatés"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985.001
   OleObjectBlob   =   "FFormatText.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FFormatText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'*                           Démo affichage de texte formaté                           *
'***************************************************************************************
Option Explicit
#Const Access = False   ' Mettre à True pour Access, à False pour Excel

#If Access Then
    Option Compare Database
#End If

Private clTexte As ClFormatText

'---------------------------------------------------------------------------------------
' Sur chargement du formulaire
'---------------------------------------------------------------------------------------
#If Access Then
Private Sub Form_Load()
#Else
Private Sub UserForm_Initialize()
#End If
Dim lWidth As Long, lHeight As Long
' Valeur du texte par défaut
txtText.value = "<font backcolor=16768255>Ce formulaire utilise la fonction <font bold=true>DrawFormattedText</font> du module <font bold=true>ClFormatText</font>.</font>" & vbCrLf & _
                "et le module CtrlBars de la librairie LibTGL" & vbCrLf & _
                vbCrLf & _
                "On peut écrire<font bold=true color=255> en couleur</font>," & vbCrLf & _
                "ou <font backcolor=55412>sur fond coloré</font>." & vbCrLf & _
                "Ou alors <font bold=true>en gras</font>, ou même en <font size=24>gros caractères</font>." & vbCrLf & _
                "Autres possibilités : <font underscore=true>texte souligné</font>, <font strikeout=true>texte barré</font>, <font italic=true>texte en italique</font>." & vbCrLf & _
                "<font size=15>Cette ligne est écrite avec la police Comic sans MS</font>" & vbCrLf & _
                "Et on peut aussi mettre un lien <font color=16711680 href=http://Arkham46.developpez.com>vers une page web</font>." & vbCrLf & _
                vbCrLf & _
                "<font size=12>Test de textes <font size=8 vertalign=""up"">en exposant</font> et <font size=8 vertalign=""down"">en indice</font></font>" & vbCrLf & _
                vbCrLf & _
                "<font backcolor=16777180>On peut imbriquer <font bold=true>les <font color=255>balises <font backcolor=14483455>sur</font> autant</font> de niveaux </font>souhaités</font>" & vbCrLf & _
                vbCrLf & _
                "Pour formater du texte, il faut l'inclure dans une balise <font bold=true>font</font> un peu comme du html." & vbCrLf & _
                "Ne pas oublier de refermer la balise!"
    
    ' Initialisation de la classe pour texte formatés
    Set clTexte = New ClFormatText
    #If Access Then
    clTexte.BackColor = Me.Section(Me.Image0.Section).BackColor
    #Else
    clTexte.BackColor = Me.BackColor
    #End If
    clTexte.BackColorGradient = RGB(255, 255, 220)
    clTexte.MarginX = 10
    clTexte.Text = txtText.value
    clTexte.Control = Me.Image0
    clTexte.DrawFormattedText
End Sub

'---------------------------------------------------------------------------------------
' Click sur bouton de mise à jour du texte formaté
'---------------------------------------------------------------------------------------
Private Sub BtnMAJ_Click()
    clTexte.Text = txtText.value
    clTexte.DrawFormattedText
End Sub



