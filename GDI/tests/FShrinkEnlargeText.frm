VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FShrinkEnlargeText 
   Caption         =   "Texte ajusté"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   OleObjectBlob   =   "FShrinkEnlargeText.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FShrinkEnlargeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private ogdi As clGdiplus
Private gX1 As Long, gY1 As Long
Private gX2 As Long, gY2 As Long
Private Const cLongMax = 2147483647

Private Sub UserForm_Initialize()
Set ogdi = New clGdiplus
ogdi.CreateBitmapForControl Me.Image0
ogdi.FillColor vbWhite
ogdi.Repaint Me.Image0
Me.txtMaxTextSize = 40
txtMaxTextSize_AfterUpdate
Me.txtMinTextSize = 8
txtMinTextSize_AfterUpdate
Me.txtText.value = "Test de texte" & vbCrLf & "Sur 2 lignes"
gX1 = -cLongMax
gY1 = -cLongMax
End Sub

Private Sub Image0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = vbKeyLButton Then
    gX1 = ogdi.CtrlToImgX(X, Me.Image0)
    gY1 = ogdi.CtrlToImgY(Y, Me.Image0)
End If
End Sub

Private Sub Image0_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = vbKeyLButton Then
    If gX1 = -cLongMax Or gY2 = -cLongMax Then Exit Sub
    gX2 = ogdi.CtrlToImgX(X, Me.Image0)
    gY2 = ogdi.CtrlToImgY(Y, Me.Image0)
    ogdi.FillColor vbWhite
    ogdi.DrawRectangle gX1, gY1, gX2, gY2
    ogdi.RepaintFast Me.Image0
End If
End Sub

Private Sub Image0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim lText As String
If Button = vbKeyLButton Then
    If gX1 = -cLongMax Or gY2 = -cLongMax Then Exit Sub
    gX2 = ogdi.CtrlToImgX(X, Me.Image0)
    gY2 = ogdi.CtrlToImgY(Y, Me.Image0)
    ogdi.FillColor vbRed, vbWhite, , gX1, gY1, gX2, gY2
    On Error Resume Next
    lText = Nz(Me.txtText.Text, "")
    On Error GoTo 0
    If lText = "" Then lText = Nz(Me.txtText.value, "")
    ogdi.DrawText lText, 12, , gX1, gY1, gX2, gY2
    ogdi.RepaintNoFormRepaint Me.Image0
    Me.lblInfo.Caption = "Taille utilisée : " & ogdi.LastTextSize
End If
End Sub

Private Sub txtMaxTextSize_AfterUpdate()
ogdi.MaxTextSize = val(Nz(Me.txtMaxTextSize.value, 0))
End Sub

Private Sub txtMaxTextSize_Change()
ogdi.MaxTextSize = val(Nz(Me.txtMaxTextSize.Text, 0))
End Sub

Private Sub txtMinTextSize_AfterUpdate()
ogdi.MinTextSize = val(Nz(Me.txtMinTextSize.value, 0))
End Sub

Private Sub txtMinTextSize_Change()
ogdi.MinTextSize = val(Nz(Me.txtMinTextSize.Text, 0))
End Sub

Private Function Nz(pValue As Variant, pValueIfNull As Variant) As Variant
If IsNull(pValue) Then
    Nz = pValueIfNull
Else
    Nz = pValue
End If
End Function
