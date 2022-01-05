VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
'*                   Classe pour affichage d'une horloge analogique                    *
'***************************************************************************************
' * Nécéssite la classe ClGdiplus
' * Code VB.Net d'origine : http://chgi.developpez.com/dotnet/compo1/vb/
'***************************************************************************************
'***************************************************************************************
'*                                      EN-TETE                                        *
'***************************************************************************************
#Const Access = False ' Mettre à True pour Access, à False pour Excel
Option Explicit
#If Access Then
    Option Compare Database
#End If
'***************************************************************************************
'*                                       API                                           *
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
'***************************************************************************************
'*                                    Constantes                                       *
'***************************************************************************************
Private Const IDC_HAND As Long = 32649
'***************************************************************************************
'*                                    Variables                                        *
'***************************************************************************************
Private gPI As Double                ' Valeur de PI
#If Access Then
Private WithEvents gCtrl As Access.Image   ' Contrôle image
Attribute gCtrl.VB_VarHelpID = -1
#Else
Private WithEvents gCtrl As MSForms.Image   ' Contrôle image
Attribute gCtrl.VB_VarHelpID = -1
#End If
#If Access Then
Private WithEvents gCtrlText As Access.TextBox   ' Contrôle texte pour heure
Attribute gCtrlText.VB_VarHelpID = -1
#Else
Private WithEvents gCtrlText As MSForms.TextBox    ' Contrôle texte pour heure
Attribute gCtrlText.VB_VarHelpID = -1
#End If
Public gGdip As clGdiplus          ' Classe Image
Private gDragRgn As String            ' Nom de l'aiguille en déplacement
Private gClockDateTime As Date          ' Date/Heure de l'horloge
Private gDigitColor As Long        ' Couleur des numéros
Private gBackColor As Long          ' Couleur de fond (autour de l'horloge)
Private gClockColor As Long       ' Couleur de fond de l'horloge
Private gHourColor As Long         ' Couleur de l'aiguille des heures
Private gMinuteColor As Long        ' Couleur de l'aiguille des minutes
Private gCouleurSeconde As Long       ' Couleur de l'aiguille des secondes
Private gDigitSize As Long         ' Taille des numéros
Private gMarkerSize As Long         ' Taille des marques
Private gMarkerColor As Long         ' Couleur des marques
Private gHandWidth As Long  ' Largeur des aiguilles
Private gDisplayDigit As Boolean     ' Affiche les numéros Vrai/Faux
Private gDigitBold As Boolean ' Affiche les numéros en gras
Private gDisplaySecond As Boolean    ' Affiche l'aiguille des secondes Vrai/Faux
Private gMoveWithMouse As Boolean ' Déplace les aiguilles avec la souris
Private gChangedByCode As Boolean ' Flag pour modif de l'heure via code (pour ne pas remettre à jour la zone de texte)
'***************************************************************************************
'*                                  EVENEMENTS                                         *
'***************************************************************************************
Public Event DateTimeChanging(pDateTime As Date)
Public Event DateTimeChanged(pDateTime As Date)
Public Event MouseMove(pRegion As String, Button As Integer, Shift As Integer)
Public Event MouseDown(pRegion As String, Button As Integer, Shift As Integer)
'***************************************************************************************
'*                                  PROPRIETES                                         *
'***************************************************************************************
Public Property Get ClockDateTime() As Date
ClockDateTime = gClockDateTime
End Property
Public Property Let ClockDateTime(pDateTime As Date)
PaintClock pDateTime
RaiseEvent DateTimeChanged(pDateTime)
End Property
Public Property Get BackColor() As Long
BackColor = gBackColor
End Property
Public Property Let BackColor(pColor As Long)
gBackColor = GetColor(pColor)
PaintClockBackGround
End Property
Public Property Get DigitColor() As Long
DigitColor = gDigitColor
End Property
Public Property Let DigitColor(pColor As Long)
gDigitColor = GetColor(pColor)
PaintClockBackGround
End Property
Public Property Get MarkerColor() As Long
MarkerColor = gMarkerColor
End Property
Public Property Let MarkerColor(pColor As Long)
gMarkerColor = GetColor(pColor)
PaintClockBackGround
End Property
Public Property Get ClockColor() As Long
ClockColor = gClockColor
End Property
Public Property Let ClockColor(pColor As Long)
gClockColor = GetColor(pColor)
PaintClockBackGround
End Property
Public Property Get HourColor() As Long
HourColor = gHourColor
End Property
Public Property Let HourColor(pColor As Long)
gHourColor = GetColor(pColor)
PaintClockBackGround
End Property
Public Property Get MinuteColor() As Long
MinuteColor = gMinuteColor
End Property
Public Property Let MinuteColor(pColor As Long)
gMinuteColor = GetColor(pColor)
PaintClockBackGround
End Property
Public Property Get DigitSize() As Long
DigitSize = gDigitSize
End Property
Public Property Let DigitSize(pSize As Long)
gDigitSize = pSize
PaintClockBackGround
End Property
Public Property Get MarkerSize() As Long
MarkerSize = gMarkerSize
End Property
Public Property Let MarkerSize(pSize As Long)
gMarkerSize = pSize
PaintClockBackGround
End Property
Public Property Get HandWidth() As Long
HandWidth = gHandWidth
End Property
Public Property Let HandWidth(pLargeur As Long)
gHandWidth = pLargeur
End Property
Public Property Get DisplayDigit() As Boolean
DisplayDigit = gDisplayDigit
End Property
Public Property Let DisplayDigit(pDisplay As Boolean)
gDisplayDigit = pDisplay
PaintClockBackGround
End Property
Public Property Get DigitBold() As Boolean
DigitBold = gDigitBold
End Property
Public Property Let DigitBold(pBold As Boolean)
gDigitBold = pBold
PaintClockBackGround
End Property
Public Property Get DisplaySecond() As Boolean
DisplaySecond = gDisplaySecond
End Property
Public Property Let DisplaySecond(pDisplay As Boolean)
gDisplaySecond = pDisplay
End Property
Public Property Get MoveWithMouse() As Boolean
MoveWithMouse = gMoveWithMouse
End Property
Public Property Let MoveWithMouse(pMoveWithMouse As Boolean)
gMoveWithMouse = pMoveWithMouse
End Property
'***************************************************************************************
'*                                  FONCTIONS                                          *
'***************************************************************************************
'---------------------------------------------------------------------------------------
' Initialise le contrôle sur lequel l'horloge est dessinée
'---------------------------------------------------------------------------------------
' Ctrl  : Contrôle image
'---------------------------------------------------------------------------------------
#If Access Then
Public Function SetClockCtrl(pCtrl As Access.Image, Optional pCtrlText As Access.TextBox)
#Else
Public Function SetClockCtrl(pCtrl As MSForms.Image, Optional pCtrlText As MSForms.TextBox)
#End If
    On Error GoTo gestion_erreurs
    ' Contrôle image
    Set gCtrl = pCtrl
    Set gCtrlText = pCtrlText
    ' Evenements pour Access
    #If Access Then
    Select Case Application.LanguageSettings.LanguageID(1)
    Case 1036
        If gCtrl.OnMouseMove = "" Then gCtrl.OnMouseMove = "[Procédure événementielle]"
        If gCtrl.OnMouseUp = "" Then gCtrl.OnMouseUp = "[Procédure événementielle]"
        If gCtrl.OnMouseDown = "" Then gCtrl.OnMouseDown = "[Procédure événementielle]"
        If Not pCtrlText Is Nothing Then
            If gCtrlText.OnKeyDown = "" Then gCtrlText.OnKeyDown = "[Procédure événementielle]"
            If gCtrlText.OnChange = "" Then gCtrlText.OnChange = "[Procédure événementielle]"
        End If
    Case Else
        If gCtrl.OnMouseMove = "" Then gCtrl.OnMouseMove = "[Event Procedure]"
        If gCtrl.OnMouseUp = "" Then gCtrl.OnMouseUp = "[Event Procedure]"
        If gCtrl.OnMouseDown = "" Then gCtrl.OnMouseDown = "[Event Procedure]"
        If Not pCtrlText Is Nothing Then
            If gCtrlText.OnKeyDown = "" Then gCtrlText.OnKeyDown = "[Event Procedure]"
            If gCtrlText.OnChange = "" Then gCtrlText.OnChange = "[Event Procedure]"
        End If
    End Select
    #End If
    ' Dessine le fond de l'horloge
    PaintClockBackGround
    ' On dessine à l'écran
    gGdip.RepaintNoFormRepaint gCtrl
    #If Access = False Then
    gCtrl.Parent.Repaint
    #End If
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la fonction SetClockCtrl du module ClClock"
End Function

'---------------------------------------------------------------------------------------
' Dessine le fond de l'horloge
'---------------------------------------------------------------------------------------
Private Sub PaintClockBackGround()
    On Error GoTo gestion_erreurs
    Dim lcpt As Integer
    Dim lAngle As Single
    Dim lX As Long, lY As Long
    If gCtrl Is Nothing Then Exit Sub
    ' On remplit l'image de la couleur de fond
    gGdip.CreateBitmap gGdip.PointsToPixelsX(gCtrl.Width), gGdip.PointsToPixelsY(gCtrl.Height)
    gGdip.FillColor gBackColor
    ' Calcul taille numéros + marques + largeur aiguilles
    If gDigitSize = 0 Then gDigitSize = gGdip.ImageWidth / 15
    If gMarkerSize = 0 Then gMarkerSize = gGdip.ImageWidth / 70
    If gHandWidth = 0 Then gHandWidth = gGdip.ImageWidth / 50
    ' Ellipse principale
    gGdip.DrawEllipse gGdip.ImageWidth / 2, gGdip.ImageHeight / 2, 0.4 * gGdip.ImageWidth, 0.4 * gGdip.ImageHeight, 1, GetColor(gClockColor), 0, gGdip.ImageWidth \ 100
    ' Marques
    For lcpt = 1 To 12
        lAngle = lcpt / 12 * 2 * gPI
        lX = gGdip.ImageWidth * (0.5 + CSng(Math.Sin(lAngle)) * 0.4)
        lY = gGdip.ImageHeight * (0.5 + CSng(-Math.Cos(lAngle)) * 0.4)
        gGdip.DrawEllipse lX, lY, gMarkerSize, gMarkerSize, 1, gMarkerColor, gMarkerColor, 1
    Next lcpt
    ' Numéros
    If gDisplayDigit Then
        For lcpt = 1 To 12
            lAngle = lcpt / 12 * 2 * gPI
            lX = gGdip.ImageWidth * (0.5 + CSng(Math.Sin(lAngle)) * (0.46))
            lY = gGdip.ImageHeight * (0.5 + CSng(-Math.Cos(lAngle)) * (0.46))
            gGdip.DrawText LTrim(CStr(lcpt)), gDigitSize, "Arial", lX, lY, , , , , gDigitColor, , , , , gDigitBold, , , True
        Next lcpt
    End If
    ' On conserve l'horloge vierge sans aiguilles
    gGdip.ImageKeep
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure PaintClockBackGround du module ClClock"
End Sub
'---------------------------------------------------------------------------------------
' Dessine l'horloge avec ses aiguilles
'---------------------------------------------------------------------------------------
' pDate  : Date/heure de l'horloge
'---------------------------------------------------------------------------------------
Private Function PaintClock(pDate As Date, Optional pTemporary As Boolean = False)
    Dim lAngle As Single
    Dim lHandPoints As Variant
    On Error GoTo gestion_erreurs
    ' On récupère l'horloge sans aiguilles
    gGdip.ImageReset
    ' Aiguille des secondes
    If gDisplaySecond Then
        lAngle = DatePart("s", pDate) / 60
        lHandPoints = CalcHandPoints(lAngle, gHandWidth / 2, 0.9)
        gGdip.DrawPolygon lHandPoints, gCouleurSeconde
    End If
    ' Aiguille des Minutes
    lAngle = DatePart("n", pDate) / 60
    lHandPoints = CalcHandPoints(lAngle, gHandWidth, 0.8)
    ' Dessine l'aiguille
    ' Et ajoute une region pour déplacer l'aiguille
    gGdip.DrawPolygon lHandPoints, gMinuteColor, 0, gGdip.ImageWidth \ 100
    gGdip.CreateRegionPolygon "minute", lHandPoints
    ' Aiguille des Heures
    lAngle = DatePart("h", pDate) / 12 + DatePart("n", pDate) / 60 / 12
    lHandPoints = CalcHandPoints(lAngle, gHandWidth, 0.6)
    ' Dessine l'aiguille
    ' Et ajoute une region pour déplacer l'aiguille
    gGdip.DrawPolygon lHandPoints, gHourColor, 0, gGdip.ImageWidth \ 100
    gGdip.CreateRegionPolygon "heure", lHandPoints
    ' Dessin à l'écran
    If pTemporary Then
        gGdip.RepaintFast gCtrl
    Else
        gGdip.RepaintNoFormRepaint gCtrl
        #If Access = False Then
        gCtrl.Parent.Repaint
        #End If
    End If
    ' Mise à jour de la date de l'horloge
    gClockDateTime = pDate
    ' Mise à jour de la zone de texte contenant l'heure
    If Not gChangedByCode Then
        If Not gCtrlText Is Nothing Then
            gCtrlText.value = Format(gClockDateTime, "hh:nn:ss")
        End If
    End If
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    If Err.Number = 2185 Then Resume Next
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la fonction PaintClock du module ClClock"
End Function
'---------------------------------------------------------------------------------------
' Calcul les coordonnées du polygone pour une aiguille
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Private Function CalcHandPoints(pAngle As Single, pTaille As Long, pFact As Single) As Variant
    Dim lAngle As Single, lAngleDroit As Single
    Dim lHandPoints(0 To 9) As Long
    Dim lHandWidth As Long
    On Error GoTo gestion_erreurs
    ' Calcul de la taille des aiguilles
    lHandWidth = 0.5 * 0.9 * IIf(gGdip.ImageWidth < gGdip.ImageHeight, gGdip.ImageWidth, gGdip.ImageHeight)
    ' Angle
    lAngle = pAngle * 2 * gPI
    lAngleDroit = (pAngle + 0.25) * 2 * gPI
    ' Points
    lHandPoints(0) = CSng(Math.Sin(lAngle)) * lHandWidth * pFact + gGdip.ImageWidth / 2
    lHandPoints(1) = CSng(-Math.Cos(lAngle)) * lHandWidth * pFact + gGdip.ImageHeight / 2
    lHandPoints(2) = CSng(-Math.Sin(lAngleDroit)) * pTaille + gGdip.ImageWidth / 2
    lHandPoints(3) = CSng(Math.Cos(lAngleDroit)) * pTaille + gGdip.ImageHeight / 2
    lHandPoints(4) = CSng(-Math.Sin(lAngle)) * lHandWidth * pFact * 0.2 + gGdip.ImageWidth / 2
    lHandPoints(5) = CSng(Math.Cos(lAngle)) * lHandWidth * pFact * 0.2 + gGdip.ImageHeight / 2
    lHandPoints(6) = CSng(Math.Sin(lAngleDroit)) * pTaille + gGdip.ImageWidth / 2
    lHandPoints(7) = CSng(-Math.Cos(lAngleDroit)) * pTaille + gGdip.ImageHeight / 2
    lHandPoints(8) = lHandPoints(0)
    lHandPoints(9) = lHandPoints(1)
    CalcHandPoints = lHandPoints
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la fonction CalcHandPoints du module ClClock"
End Function

'---------------------------------------------------------------------------------------
' Initialisation de la classe horloge
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()
On Error GoTo gestion_erreurs
    ' Classe Image pour dessiner
    Set gGdip = New clGdiplus
    ' Lissage des traits
    gGdip.SmoothingMode = GdipSmoothingAntialias
    ' Valeurs par défaut pour l'horloge
    gDigitColor = vbBlack
    gClockColor = RGB(200, 200, 200)
    gHourColor = vbRed
    gMinuteColor = vbGreen
    gCouleurSeconde = vbBlack
    gBackColor = vbWhite
    gDisplayDigit = False
    ' Calcul de PI
    gPI = 4 * Atn(1)
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure Class_Initialize du module ClClock"
End Sub
'---------------------------------------------------------------------------------------
' Libération de la classe horloge
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo gestion_erreurs
    ' Libère l'objet contrôle
    Set gCtrl = Nothing
    ' On libère la classe Image
    If Not gGdip Is Nothing Then Set gGdip = Nothing
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure Class_Terminate du module ClClock"
End Sub

'---------------------------------------------------------------------------------------
' MouseDown sur l'horloge
'---------------------------------------------------------------------------------------
' X,Y : Position de la souris sur l'image
' DragRgn prend la valeur de la région que l'on sélectionne, donc l'aiguille
'   qui va être déplacée
'---------------------------------------------------------------------------------------
#If Access Then
Private Sub gCtrl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub gCtrl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
Dim lregion As String
On Error GoTo gestion_erreurs
lregion = gGdip.GetRegionXY(gGdip.CtrlToImgX(CLng(X), gCtrl), gGdip.CtrlToImgY(CLng(Y), gCtrl))
If gMoveWithMouse Then
    gDragRgn = lregion
End If
' Lève l'évenement souris appuyée
RaiseEvent MouseDown(lregion, Button, Shift)
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure gCtrl_MouseDown du module ClClock"
End Sub

'---------------------------------------------------------------------------------------
' Déplacement de la souris sur l'horloge
'---------------------------------------------------------------------------------------
' X,Y : Position de la souris sur l'image
'---------------------------------------------------------------------------------------
#If Access Then
Private Sub gCtrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub gCtrl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    Dim lregion As String
    Dim lAngle As Double
    Dim lCalcX As Double
    Dim lCalcY As Double
    Dim lCalcM As Integer
    Dim lMult As Integer
    Dim lTmpDate As Date
    Dim lX As Long, lY As Long
    On Error GoTo gestion_erreurs
    ' Recherche de l'aiguille survolée
    If Not gGdip Is Nothing Then
        lX = gGdip.CtrlToImgX(CLng(X), gCtrl)
        lY = gGdip.CtrlToImgY(CLng(Y), gCtrl)
        lregion = gGdip.GetRegionXY(lX, lY)
    End If
    If gMoveWithMouse Then
        If lregion <> "" Then
            ' Si une aiguille est survolée on change le curseur en main
          SetHandCursor
        End If
        ' Si on déplace une aiguille (DragRgn contient le nom de l'aiguille déplacée)
        If gDragRgn <> "" Then
            ' Calcul de la position de la souris par rapport au centre de l'horloge
            lCalcX = (gGdip.ImageWidth / 2) - lX
            lCalcY = (gGdip.ImageWidth / 2) - lY * (gGdip.ImageWidth / gGdip.ImageHeight)
            ' Calcul de l'angle correspondant à la position de la souris
            If lCalcX = 0 Then
                ' Cas particulier si CalcX = 0
                If lCalcY > 0 Then
                    lAngle = 0
                Else
                    lAngle = 180
                End If
                ' Sinon si CalcX <> 0 on calcul l'angle avec l'Arctangente
            Else
                lAngle = Atn(lCalcY / lCalcX) * 180 / gPI + 90
            End If
            ' Si on est à droite du centre on ajoute 180°
            If lCalcX > 0 Then lAngle = lAngle + 180
            ' Calcul de l'heure correspondant à la position de la souris
            Select Case gDragRgn
            Case "heure"
                ' On stocke l'heure pour les calculs
                lTmpDate = gClockDateTime
                ' Calcul des minutes correspondant à la position de la souris
                ' A-t-on changé de jour?
                ' Si on passé au jour précédent il ne faut pas juste faire un calcul algébrique
                '   entre les anciennes minutes et les nouvelles minutes, il faut décaler d'un journée
                '   et soustraire au lieu d'additionner
                If Abs(Int(lAngle * 720 / 360) - 60 * (DatePart("h", gClockDateTime) Mod 12) - DatePart("n", gClockDateTime)) > 360 Then
                    lCalcM = CInt((360 - lAngle) * 720 / 360)
                    lMult = -1
                Else
                    lCalcM = CInt(lAngle * 720 / 360)
                    lMult = 1
                End If
                ' Ajout des minutes à l'ancienne heure
                lTmpDate = DateAdd("n", lMult * (lCalcM - 60 * (DatePart("h", gClockDateTime) Mod 12) - DatePart("n", gClockDateTime)), gClockDateTime)
                ' repaint l'horloge si l'heure a changé
                If lTmpDate <> gClockDateTime Then
                    PaintClock lTmpDate, True
                    RaiseEvent DateTimeChanging(lTmpDate)
                End If
            Case "minute"
                ' Calcul des minutes correspondant à la position de la souris
                ' A-t-on changé d'heure? (on est passé par ex. de 13h01 à 12h59)
                ' Si on est passé par ex. de 13h01 à 12h59 on doit retirer 2 min au lieu d'en ajouter 59-01 = 58
                lTmpDate = gClockDateTime
                If Abs(Int(lAngle * 60 / 360) - DatePart("n", gClockDateTime)) > 30 Then
                    lCalcM = CInt((360 - lAngle) * 60 / 360)
                    lMult = -1
                Else
                    lCalcM = CInt(lAngle * 60 / 360)
                    lMult = 1
                End If
                ' Ajout des minutes à l'ancienne heure
                lTmpDate = DateAdd("n", lMult * (lCalcM - DatePart("n", gClockDateTime)), gClockDateTime)
                ' repaint l'horloge si l'heure a changé
                If lTmpDate <> gClockDateTime Then
                    PaintClock lTmpDate, True
                    RaiseEvent DateTimeChanging(lTmpDate)
                End If
            End Select
        End If
    End If
    ' Lève l'évenement souris déplacée
    RaiseEvent MouseMove(lregion, Button, Shift)
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure gCtrl_MouseMove du module ClClock"
End Sub

'---------------------------------------------------------------------------------------
' MouseUp sur l'horloge
'---------------------------------------------------------------------------------------
' On arrête le déplacement de l'aiguille
'---------------------------------------------------------------------------------------
#If Access Then
Private Sub gCtrl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
#Else
Private Sub gCtrl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#End If
    On Error GoTo gestion_erreurs
    PaintClock ClockDateTime
    If gMoveWithMouse Then
        gDragRgn = ""
        RaiseEvent DateTimeChanged(gClockDateTime)
    End If
    On Error GoTo 0
    Exit Sub
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la procédure gCtrl_MouseUp du module ClClock"
End Sub


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
    On Error GoTo gestion_erreurs
    If pColor < 0 Then
        Call OleTranslateColor(pColor, 0, pColor)
    End If
    GetColor = pColor
    On Error GoTo 0
    Exit Function
gestion_erreurs:
    MsgBox "Error " & Err.Number & " (" & Err.description & ") dans la fonction GetColor du module ClClock"
End Function

'---------------------------------------------------------------------------------------
' Applique un curseur en forme de main
'---------------------------------------------------------------------------------------
Private Sub SetHandCursor()
    Dim hCur
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

'---------------------------------------------------------------------------------------
' Changement d'heure avec les flèches du clavier
'---------------------------------------------------------------------------------------
#If Access Then
Private Sub gCtrlText_KeyDown(KeyCode As Integer, Shift As Integer)
#Else
Private Sub gCtrlText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
#End If
Dim lPos As Long
Dim lInterval As String
Dim lOffset As Integer
On Error GoTo gestion_erreurs
If KeyCode <> 38 And KeyCode <> 40 Then Exit Sub ' haut ou bas
lOffset = IIf(KeyCode = 38, 1, -1)
lPos = InStr(gCtrlText.Text, ":")
If gCtrlText.SelStart < lPos Then
    lInterval = "h"
Else
    lPos = InStr(lPos + 1, gCtrlText.Text, ":")
    If gCtrlText.SelStart < lPos Then
        lInterval = "n"
    Else
        lInterval = "s"
    End If
End If
lPos = gCtrlText.SelStart
ClockDateTime = DateAdd(lInterval, lOffset, gClockDateTime)
gCtrlText.SelStart = lPos
KeyCode = 0
If gCtrlText.ForeColor <> vbBlack Then gCtrlText.ForeColor = vbBlack
On Error GoTo 0
Exit Sub
gestion_erreurs:
gCtrlText.ForeColor = vbRed
End Sub

'---------------------------------------------------------------------------------------
' Changement d'heure dans la zone de texte
'---------------------------------------------------------------------------------------
Private Sub gCtrlText_Change()
Dim lDate As String
On Error GoTo gestion_erreurs
lDate = DateValue(gClockDateTime) + TimeValue(gCtrlText.Text)
If IsDate(lDate) Then
    gChangedByCode = True
    ClockDateTime = lDate
    gChangedByCode = False
    If gCtrlText.ForeColor <> vbBlack Then gCtrlText.ForeColor = vbBlack
End If
On Error GoTo 0
Exit Sub
gestion_erreurs:
gCtrlText.ForeColor = vbRed
End Sub

