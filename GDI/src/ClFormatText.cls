VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClFormatText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const Access = False   ' Mettre à True pour Access, à False pour Excel
'***************************************************************************************
'*                      CLASSE POUR ECRITURE DE TEXTES FORMATES                        *
'***************************************************************************************

'***************************************************************************************
' Auteur : Thierry GASPERMENT (Arkham46)
' v0.4 (05/10/2010)
' Nécessite la classe VBA ClGdiplus
' Le code est libre pour toute utilisation
'***************************************************************************************
'***************************************************************************************
'*                                      EN-TETE                                        *
'***************************************************************************************
Option Explicit
'***************************************************************************************
'*                                      API                                            *
'***************************************************************************************
#If VBA7 Then
Private Declare PtrSafe Function OleTranslateColor Lib "oleAut32.dll" _
                                           (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
#Else
Private Declare Function OleTranslateColor Lib "olepro32.dll" _
                                           (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If
Private WithEvents ogdi As clGdiplus
Attribute ogdi.VB_VarHelpID = -1
'***************************************************************************************
'*                                    Variables                                        *
'***************************************************************************************
#If Access Then
Private gCtrl As Access.Image
#Else
Private gCtrl As MSForms.Image
#End If
' Texte
Public Text As String
' Couleurs de fond
Private gBackColor As Long
Private gBackColorGradient As Long
Private gGradientVertical As Boolean
Private gLinkColor As Long
' Marge à gauche
Private gMarginX As Long
' Activer les liens URL
Private gActiveURL As Boolean
' Taille maxi du texte
Private gMaxX As Single
Private gMaxY As Single
'***************************************************************************************
'*                                   Constantes                                        *
'***************************************************************************************
Private Const IDC_HAND = 32649              ' Curseur main
'***************************************************************************************
'*                                   Propriétés                                        *
'***************************************************************************************
#If Access Then
Public Property Let Control(pCtrl As Access.Image)
#Else
Public Property Let Control(pCtrl As MSForms.Image)
#End If
Set gCtrl = pCtrl
Set ogdi = New clGdiplus
If Not pCtrl Is Nothing Then
    ogdi.CreateBitmapForControl pCtrl
Else
    ogdi.CloseImage
End If
End Property
#If Access Then
Public Property Get Control() As Access.Image
#Else
Public Property Get Control() As MSForms.Image
#End If
Set Control = gCtrl
End Property
Public Property Get LinkColor() As Long
    LinkColor = gLinkColor
End Property
Public Property Let LinkColor(pColor As Long)
    gLinkColor = GetColor(pColor)
End Property
Public Property Get BackColor() As Long
    BackColor = gBackColor
End Property
Public Property Let BackColor(pColor As Long)
    gBackColor = GetColor(pColor)
End Property
Public Property Get BackColorGradient() As Long
    BackColorGradient = gBackColorGradient
End Property
Public Property Let BackColorGradient(pColor As Long)
    gBackColorGradient = GetColor(pColor)
End Property
Public Property Get GradientVertical() As Boolean
    GradientVertical = gGradientVertical
End Property
Public Property Let GradientVertical(pVertical As Boolean)
    gGradientVertical = pVertical
End Property
Public Property Get ActiveURL() As Boolean
    ActiveURL = gActiveURL
End Property
Public Property Let ActiveURL(pActive As Boolean)
    gActiveURL = pActive
End Property
Public Property Get MarginX() As Long
    MarginX = gMarginX
End Property
Public Property Let MarginX(pPixels As Long)
    gMarginX = pPixels
End Property

'***************************************************************************************
'*                                    Fonctions                                        *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Ecrit le texte formaté
'---------------------------------------------------------------------------------------
' loGdi : Classe image
'---------------------------------------------------------------------------------------
Public Function DrawFormattedText(Optional poGdi As clGdiplus)
    Dim lPosFont1 As Integer
    Dim lPosFont2 As Integer
    Dim lPosFontEnd1 As Integer
    Dim lPosFontEnd2 As Integer
    Dim lPosText1 As Integer
    Dim lPosText2 As Integer
    Dim lPosText2bis As Integer
    Dim lPosCrLf As Integer
    Dim lText As String
    Dim lTextToWrite As String
    Dim lParam As String
    Dim lHeight() As Long
    Dim lWeight() As Boolean
    Dim lItalic() As Boolean
    Dim lUnderScore() As Boolean
    Dim lStrikeOut() As Boolean
    Dim lBackColor() As Long
    Dim lPenColor() As Long
    Dim lLink() As String
    Dim lFontName() As String
    Dim lVertAlign() As String
    Dim lAlign() As String
    Dim lTop As Long
    Dim lLeft As Long
    Dim lTextWidth As Long, lTextHeight As Long
    Dim lColText As New Collection
    Dim lColLine As New Collection
    Dim lColWidth As New Collection
    Dim lColTextItem As Variant
    Dim lTextHeightMax As Long
    Dim lLine As Integer
    Dim lLevel As Integer
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
    Dim lTextAlign As Long
    Dim lDecalage As Long
    Dim loGdi As clGdiplus
    
    If poGdi Is Nothing Then
        Set loGdi = ogdi
    Else
        Set loGdi = poGdi
    End If
    lTextToWrite = Text
    ' Couleur de fond
    loGdi.FillColor gBackColor, _
                         IIf(gBackColor <> gBackColorGradient, gBackColorGradient, -1), _
                         gGradientVertical
    'loGdi.DrawRectangle 0, 0, loGdi.ImageWidth - 1, loGdi.ImageHeight - 1
    'On Error Resume Next
    lLevel = 0
    ' Marge gauche
    lLeft = gMarginX
    ' Découpe les textes entre les balises <font>
    ReDim lWeight(0)
    ReDim lBackColor(0)
    ReDim lPenColor(0)
    ReDim lHeight(0)
    ReDim lLink(0)
    ReDim lFontName(0)
    ReDim lItalic(0)
    ReDim lUnderScore(0)
    ReDim lStrikeOut(0)
    ReDim lVertAlign(0)
    ReDim lAlign(0)
    gMaxX = 0
    gMaxY = 0
    Do
        ' Valeurs par défaut
        lWeight(0) = False
        lBackColor(0) = -1
        lPenColor(0) = 0
        lHeight(0) = 8
        lLink(0) = ""
        lFontName(0) = "Arial"
        lItalic(0) = False
        lUnderScore(0) = False
        lStrikeOut(0) = False
        lVertAlign(0) = ""
        lAlign(0) = ""
        ' redimensionne lColTextItem pour contenir les paramètres
        ReDim lColTextItem(16) As Variant
        ' Position balise font
        lPosFont1 = InStr(1, lTextToWrite, "<font ")
        lPosFontEnd1 = InStr(1, lTextToWrite, "</font>")
        If lPosFont1 = 1 Then    ' Balise font au début
            ' Cherche la fin de la balise font
            lPosFont2 = InStr(lPosFont1, lTextToWrite, ">")
            ' lParam contient tous les paramètres
            lParam = Mid(lTextToWrite, lPosFont1 + 5, lPosFont2 - lPosFont1 - 5)
            ' Le texte commence après la fin de la balise font
            lPosText1 = lPosFont2 + 1
            ' la fin du texte correspond à une balise </font> ou <font
            lPosText2 = InStr(lPosText1, lTextToWrite, "</font>")
            lPosText2bis = InStr(lPosText1, lTextToWrite, "<font ")
            lLevel = lLevel + 1
            ReDim Preserve lWeight(lLevel)
            ReDim Preserve lBackColor(lLevel)
            ReDim Preserve lPenColor(lLevel)
            ReDim Preserve lHeight(lLevel)
            ReDim Preserve lLink(lLevel)
            ReDim Preserve lFontName(lLevel)
            ReDim Preserve lItalic(lLevel)
            ReDim Preserve lUnderScore(lLevel)
            ReDim Preserve lStrikeOut(lLevel)
            ReDim Preserve lVertAlign(lLevel)
            ReDim Preserve lAlign(lLevel)
            lWeight(lLevel) = lWeight(lLevel - 1)
            lBackColor(lLevel) = lBackColor(lLevel - 1)
            lPenColor(lLevel) = lPenColor(lLevel - 1)
            lHeight(lLevel) = lHeight(lLevel - 1)
            lLink(lLevel) = lLink(lLevel - 1)
            lFontName(lLevel) = lFontName(lLevel - 1)
            lItalic(lLevel) = lItalic(lLevel - 1)
            lUnderScore(lLevel) = lUnderScore(lLevel - 1)
            lStrikeOut(lLevel) = lStrikeOut(lLevel - 1)
            lVertAlign(lLevel) = lVertAlign(lLevel - 1)
            lAlign(lLevel) = lAlign(lLevel - 1)
            ' Extractions des divers paramètres
            If ExistsParam(lParam, "size") Then lHeight(lLevel) = CLng(Eval(ExtractParam(lParam, "size")))
            If ExistsParam(lParam, "bold") Then lWeight(lLevel) = CBool(Eval(ExtractParam(lParam, "bold")))
            If ExistsParam(lParam, "italic") Then lItalic(lLevel) = CBool(Eval(ExtractParam(lParam, "italic")))
            If ExistsParam(lParam, "underscore") Then lUnderScore(lLevel) = CBool(Eval(ExtractParam(lParam, "underscore")))
            If ExistsParam(lParam, "strikeout") Then lStrikeOut(lLevel) = CBool(Eval(ExtractParam(lParam, "strikeout")))
            If ExistsParam(lParam, "backcolor") Then lBackColor(lLevel) = CLng(Eval(ExtractParam(lParam, "backcolor")))
            If ExistsParam(lParam, "href") Then
                lLink(lLevel) = ExtractParam(lParam, "href")
                lPenColor(lLevel) = gLinkColor
            End If
            If ExistsParam(lParam, "color") Then lPenColor(lLevel) = CLng(Eval(ExtractParam(lParam, "color")))
            If ExistsParam(lParam, "name") Then lFontName(lLevel) = ExtractParam(lParam, "name")
            If ExistsParam(lParam, "vertalign") Then lVertAlign(lLevel) = ExtractParam(lParam, "vertalign")
            If ExistsParam(lParam, "align") Then lAlign(lLevel) = ExtractParam(lParam, "align")
            ' Si manque la balise de fin, on prend tout le texte restant
            If lPosText2 = 0 And lPosText2bis = 0 Then
                lPosText2 = Len(lTextToWrite)
                lText = Mid(lTextToWrite, lPosText1, lPosText2 - lPosText1 + 1)
            ElseIf lPosText2bis > lPosText2 Or lPosText2bis = 0 Then
                lText = Mid(lTextToWrite, lPosText1, lPosText2 - lPosText1)
                lPosText2 = lPosText2 - 1   ' on va jusqu'à la fin de la balise </font>
            ElseIf lPosText2bis < lPosText2 Then
                lText = Mid(lTextToWrite, lPosText1, lPosText2bis - lPosText1)
                lPosText2 = lPosText2bis - 1   ' on va jusqu'au début de la balise <font >
            End If
        ElseIf lPosFontEnd1 = 1 Then    ' Balise de fin
            lPosText2 = InStr(Len("</font>"), lTextToWrite, "</font>") - 1
            If lPosText2 = -1 Then lPosText2 = Len(lTextToWrite)
            lPosText2bis = InStr(Len("</font>"), lTextToWrite, "<font ") - 1
            If lPosText2bis = -1 Then lPosText2bis = Len(lTextToWrite)
            lPosText2 = fMin(lPosText2, lPosText2bis)
            If lPosText2 = -1 Then lPosText2 = Len(lTextToWrite)
            lText = Mid(lTextToWrite, Len("</font>") + 1, lPosText2 - Len("</font>"))
            If lLevel > 0 Then lLevel = lLevel - 1
            ReDim Preserve lWeight(lLevel)
            ReDim Preserve lBackColor(lLevel)
            ReDim Preserve lPenColor(lLevel)
            ReDim Preserve lHeight(lLevel)
            ReDim Preserve lLink(lLevel)
            ReDim Preserve lFontName(lLevel)
            ReDim Preserve lItalic(lLevel)
            ReDim Preserve lUnderScore(lLevel)
            ReDim Preserve lStrikeOut(lLevel)
            ReDim Preserve lVertAlign(lLevel)
            ReDim Preserve lAlign(lLevel)
        ElseIf lPosFont1 = 0 Then    ' Pas de balise font trouvée
            lPosText1 = 1
            lPosText2 = Len(lTextToWrite)
            ' On affiche tous le texte restant
            lText = lTextToWrite
        Else    ' Balise font trouvée plus loin dans le texte
            lPosText1 = 1
            lPosText2 = lPosFont1 - 1
            ' On affiche le texte jusqu'à la prochaine balise font
            lText = Mid(lTextToWrite, lPosText1, lPosText2 - lPosText1 + 1)
        End If
        ' Boucle tant qu'on trouve des sauts de lignes dans le texte
        Do
            ' Recherche position d'un saut de ligne
            lPosCrLf = InStr(1, lText, vbCrLf)
            If lPosCrLf = 0 Then
                ' Pas de saut de ligne
                lColTextItem(0) = lText
            Else
                ' Saut de ligne trouvé, on s'occupe du texte avant le saut de ligne
                lColTextItem(0) = Left(lText, lPosCrLf - 1)
                lText = Right(lText, Len(lText) - lPosCrLf - 1)
            End If
            ' Alignement du texte
            Select Case lColTextItem(15)
                Case "up"
                    lTextAlign = 0
                Case "down"
                    lTextAlign = 2
                Case "center"
                    lTextAlign = 1
                Case Else
                    lTextAlign = 2
            End Select
            ' Calcul taille du texte
            lx1 = 0
            ly1 = 0
            lx2 = 50000
            ly2 = 5000
            loGdi.DrawText IIf(CStr(lColTextItem(0)) = "", "Û", CStr(lColTextItem(0))), loGdi.FontSizeToPixel(lHeight(lLevel)), lFontName(lLevel), _
                            lx1, ly1, lx2, ly2, 0, lTextAlign, , , , , lItalic(lLevel), lWeight(lLevel), lUnderScore(lLevel), lStrikeOut(lLevel), True
            lTextWidth = IIf(CStr(lColTextItem(0)) = "", 0, lx2 - lx1 + 1)
            lTextHeight = ly2 - ly1 + 1
            ' Si texte vide on remet la largeur à 0
            If lColTextItem(0) = "" Then lTextWidth = 0
            ' Conserve la hauteur maxi de la ligne
            If lTextHeightMax < lTextHeight Then lTextHeightMax = lTextHeight
            ' Stocke les paramètres du texte dans la collection lColText
            lColTextItem(1) = lLeft
            lColTextItem(2) = lTop
            lColTextItem(3) = lLeft + lTextWidth
            lColTextItem(4) = lTop + lTextHeight
            lColTextItem(5) = lBackColor(lLevel)
            lColTextItem(6) = lPenColor(lLevel)
            lColTextItem(7) = lHeight(lLevel)
            lColTextItem(8) = lWeight(lLevel)
            lColTextItem(9) = lLine + 1
            lColTextItem(10) = lLink(lLevel)
            lColTextItem(11) = lFontName(lLevel)
            lColTextItem(12) = lItalic(lLevel)
            lColTextItem(13) = lUnderScore(lLevel)
            lColTextItem(14) = lStrikeOut(lLevel)
            lColTextItem(15) = lVertAlign(lLevel)
            lColTextItem(16) = lAlign(lLevel)
            lColText.Add lColTextItem
            If CLng(lColTextItem(3)) > gMaxX Then gMaxX = CLng(lColTextItem(3))
            'If lColTextItem(2) + lColLine(lColTextItem(9)) > gMaxY Then gMaxY = lColTextItem(2) + lColLine(lColTextItem(9))
            If lColTextItem(2) + lTextHeightMax > gMaxY Then gMaxY = lColTextItem(2) + lTextHeightMax 'lColLine(lColTextItem(9))
            ' Décale la position du prochain texte
            lLeft = lLeft + lTextWidth
            ' Si saut de ligne
            If lPosCrLf <> 0 Then
                ' Passe à la ligne suivante
                lLine = lLine + 1
                ' Ajoute la largeur de la ligne à la collection lColWidth
                lColWidth.Add lLeft
                ' Recommence à gauche
                lLeft = gMarginX
                ' Ajoute la hauteur de la ligne à la position verticale
                lTop = lTop + lTextHeightMax
                ' Ajoute la hauteur de la ligne à la collection lColLine
                lColLine.Add lTextHeightMax
                ' Réinitialise la hauteur de la ligne suivante
                lTextHeightMax = 0
            Else
                ' Si pas de saut de ligne on passe au texte suivant
                Exit Do
            End If
        Loop
        ' Retire le texte déjà écrit
        lTextToWrite = Right(lTextToWrite, Len(lTextToWrite) - lPosText2)
        ' Si plus de texte à écrire
        If Len(lTextToWrite) = 0 Then
            ' Ajoute la hauteur de la ligne à la collection lColLine
            lColLine.Add lTextHeightMax
            ' Ajoute la largeur de la ligne à la collection lColWidth
            lColWidth.Add lLeft
            ' Quitte la boucle
            Exit Do
        End If
    Loop
    If Not gCtrl Is Nothing Then
        ' Barres de défilement
        If Not ogdi.BarExists Then ogdi.BarNew
        ogdi.BarScaleX gMaxX, 1
        ogdi.BarScaleY gMaxY, 1
        ogdi.BarObject = gCtrl
    End If
    ' Affiche tous les textes
    For Each lColTextItem In lColText
        ' Alignement du texte
        Select Case lColTextItem(15)
            Case "up"
                lTextAlign = 0
            Case "down"
                lTextAlign = 2
            Case "center"
                lTextAlign = 1
            Case Else
                lTextAlign = 2
        End Select
        ' Dessine le texte
        Select Case lColTextItem(16)
            Case "center"
                lDecalage = (loGdi.ImageWidth - lColWidth(lColTextItem(9))) / 2
            Case "right"
                lDecalage = loGdi.ImageWidth - lColWidth(lColTextItem(9)) - gMarginX
            Case Else
                lDecalage = 0
        End Select
        ' Dessine le texte
        loGdi.DrawText CStr(lColTextItem(0)), loGdi.FontSizeToPixel(lColTextItem(7)), CStr(lColTextItem(11)), _
                    ogdi.BarStartX + lDecalage + CLng(lColTextItem(1)), ogdi.BarStartY + CLng(lColTextItem(2)), _
                    ogdi.BarStartX + lDecalage + CLng(lColTextItem(3)), ogdi.BarStartY + lColTextItem(2) + lColLine(lColTextItem(9)), _
                    0, lTextAlign, CLng(lColTextItem(6)), , CLng(lColTextItem(5)), , CBool(lColTextItem(12)), CBool(lColTextItem(8)), CBool(lColTextItem(13)), CBool(lColTextItem(14))
        ' Ajoute une région s'il y a un lien href
        If lColTextItem(10) <> "" Then loGdi.CreateRegionRect CStr(lColTextItem(10)), _
                        ogdi.BarStartX + lColTextItem(1), ogdi.BarStartY + lColTextItem(2), _
                        ogdi.BarStartX + lColTextItem(3), ogdi.BarStartY + lColTextItem(2) + lColLine(lColTextItem(9))
    Next
    ' Barres de défilement
    If Not ogdi Is Nothing Then
        ogdi.BarDraw
    End If
    ' Dessin dans le contrôle
    If Not gCtrl Is Nothing Then ogdi.RepaintNoFormRepaint gCtrl
    ' Vide les collections
    Set lColText = Nothing
    Set lColLine = Nothing
    Set lColWidth = Nothing
End Function

'---------------------------------------------------------------------------------------
' Extrait un paramètre
'---------------------------------------------------------------------------------------
' pText : texte contenant les paramètres
' pParam : Nom du pramètre à extraire
'---------------------------------------------------------------------------------------
Private Function ExtractParam(pText As String, pParam As String) As Variant
    Dim lPosParam1 As Integer
    Dim lPosParam2 As Integer
    lPosParam1 = InStr(pText, " " & pParam & "=")
    If lPosParam1 > 0 Then
        If Mid(pText, lPosParam1 + Len(pParam) + 2, 1) = """" Then
            lPosParam1 = lPosParam1 + 1
            lPosParam2 = InStr(lPosParam1 + Len(pParam) + 3, pText, """")
        Else
            lPosParam2 = InStr(lPosParam1 + Len(pParam) + 1, pText, " ")
        End If
        If lPosParam2 = 0 Then lPosParam2 = Len(pText) + 1
        ExtractParam = Mid(pText, lPosParam1 + Len(pParam) + 2, lPosParam2 - lPosParam1 - Len(pParam) - 2)
    End If
End Function
'---------------------------------------------------------------------------------------
' Teste l'existence d'un paramètre
'---------------------------------------------------------------------------------------
' pText : texte contenant les paramètres
' pParam : Nom du pramètre à extraire
'---------------------------------------------------------------------------------------
Private Function ExistsParam(pText As String, pParam As String) As Boolean
    Dim lPosParam1 As Integer
    lPosParam1 = InStr(pText, " " & pParam & "=")
    ExistsParam = (lPosParam1 > 0)
End Function

'---------------------------------------------------------------------------------------
' Valeur Maxi
'---------------------------------------------------------------------------------------
' Renvoie la valeur maxi entre a et b
'---------------------------------------------------------------------------------------
Private Function fMax(A, B) As Variant
    If A > B Then fMax = A Else fMax = B
End Function

'---------------------------------------------------------------------------------------
' Valeur Mini
'---------------------------------------------------------------------------------------
' Renvoie la valeur mini entre a et b
'---------------------------------------------------------------------------------------
Private Function fMin(A, B) As Variant
    If A < B Then fMin = A Else fMin = B
End Function

'---------------------------------------------------------------------------------------
' Initilisation de la classe
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    gLinkColor = RGB(0, 50, 100)
    gBackColor = vbWhite
    gBackColorGradient = vbWhite
    gGradientVertical = False
    gMarginX = 5
    gActiveURL = True
    Set ogdi = New clGdiplus
End Sub

'---------------------------------------------------------------------------------------
' Fermeture de la classe
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
ogdi.BarDelete
Set ogdi = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Remplacement Eval pour Excel
'---------------------------------------------------------------------------------------
#If Access = False Then
Private Function Eval(pValue As String)
#If Win64 Then
    Eval = pValue
#Else
Dim lSC As Object
    Set lSC = CreateObject("ScriptControl")
    lSC.language = "vbscript"
    Eval = lSC.Eval(pValue)
#End If
End Function
#End If

'---------------------------------------------------------------------------------------
' Récupère la couleur système si nécessaire
'---------------------------------------------------------------------------------------
' pColor     : Numéro de la couleur
' Les couleurs définies par défaut sont parfois des couleurs systèmes
' Par exemple la couleur le fond d'un formulaire par défaut est -2147483633
' il faut retrouver la couleur définie dans le système pour cette valeur
' Renvoie la couleur dans un Long
'---------------------------------------------------------------------------------------
Private Function GetColor(ByVal pColor As Long) As Long
    If pColor < 0 Then
        Call OleTranslateColor(pColor, 0, pColor)
    End If
    GetColor = pColor
End Function

'---------------------------------------------------------------------------------------
' Souris appuyée sur image
'---------------------------------------------------------------------------------------
Private Sub ogdi_BarMouseDown(BarName As String, lregion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
If gActiveURL Then
    If lregion <> "" Then
        #If Access Then
        Application.FollowHyperlink lregion
        #Else
        ThisWorkbook.FollowHyperlink lregion
        #End If
    End If
End If
End Sub

'---------------------------------------------------------------------------------------
' Souris déplacée sur image
'---------------------------------------------------------------------------------------
Private Sub oGdi_BarMouseMove(BarName As String, lregion As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
If gActiveURL Then
    If lregion <> "" Then
        SetHandCursor
    End If
End If
End Sub

Private Sub ogdi_BarOnRefreshNeeded(BarName As String, MouseUp As Boolean)
DrawFormattedText
End Sub

'---------------------------------------------------------------------------------------
' Applique un curseur en forme de main
'---------------------------------------------------------------------------------------
Private Sub SetHandCursor()
    #If VBA7 Then
    Dim hCur As LongPtr
    #Else
    Dim hCur As Long
    #End If
    On Error GoTo gestion_erreurs
    hCur = LoadCursor(0, IDC_HAND)
    If (hCur > 0) Then
        SetCursor hCur
    End If
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure SetHandCursor du module ClClock"
End Sub





