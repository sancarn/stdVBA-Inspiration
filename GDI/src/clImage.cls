VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Objet Image
'---------------------------------------------------------------------------------------

Public X As Long, Y As Long ' Position
Public name As String ' Nom

' Id de cet objet = pointeur de l'objet courant converti en texte
Public Property Get id() As String
id = CStr(ObjPtr(Me))
End Property


