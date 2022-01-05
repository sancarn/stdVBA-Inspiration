VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
'*              OBJET FEU (AVANT EXPLOSION)
'***************************************************************************************

Option Explicit

Public X0 As Currency
Public Y0 As Currency
Public X As Currency
Public Y As Currency
Public SpeedX0 As Currency
Public SpeedY0 As Currency
Public TimeUp As Double
Public Timer As Double
Public Color As Long
Public SubFire As Boolean

Public Property Get key() As String
key = CStr(ObjPtr(Me))
End Property

