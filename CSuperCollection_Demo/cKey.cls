VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit

  Private m_sName As String
  Private m_sAssociatedIndex As Long
  
Public Property Get Name() As String
  Name = m_sName
End Property

Public Property Get AssociatedIndex() As Long
  AssociatedIndex = m_sAssociatedIndex
End Property

Friend Sub Initialize(ByVal sName$, ByVal nIndex&)
  m_sName = sName
  m_sAssociatedIndex = nIndex
End Sub
