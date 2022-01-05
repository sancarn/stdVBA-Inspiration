VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSVGNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'*   Objet contenant les infos d'un département ou d'une region
'*****************************************************************
Option Explicit

Public Enum ETypeElement
    TypeRegion = 1
    TypeDepartement = 2
End Enum

Public libcourt As String
Public liblong As String
Public Path As String
Public LineWidth As Long
Public BackColor As Long
Public LineColor As Long
Public key As String
Public TypeElement As ETypeElement

Private Sub Class_Initialize()
LineWidth = 1
key = CStr(ObjPtr(Me))
End Sub

