VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFruit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit
  ' single image object.  this object should encapsulate all of the
  ' necessary functionality of an image object.

  ' ###  Begin Required for collection  ###
  Private m_sKey As String
  ' ###  End Required for collection  ###

  Private m_bIsLoading As Boolean
  Private m_bIsDirty As Boolean
  
  Private m_sName As String
  
  Public Enum FRUIT_CATEGORY
    Citrus = 1&
    Pome = 2&
    Berry = 3&
    Other = 4&
  End Enum

  Private m_eCategory As FRUIT_CATEGORY

  Private m_lpSuperCollection As Long
  
' ######  Begin required for collection  ########
Public Property Get Key() As String
  Key = m_sKey
End Property
Public Property Let Key(ByVal sKey$)
  ' this property should *ONLY* be set from the collection object.
  ' there is no way to enforce this, however.
  m_sKey = sKey
End Property
' ######  End Required for collection  ##########

Public Property Get IsLoading() As Boolean
  IsLoading = m_bIsLoading
End Property
Public Property Let IsLoading(ByVal bNewVal As Boolean)
  m_bIsLoading = bNewVal
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = m_bIsDirty
End Property
Public Property Let IsDirty(ByVal bNewVal As Boolean)
  m_bIsDirty = bNewVal
End Property

Public Property Get Name() As String
  Name = m_sName
End Property
Public Property Let Name(ByVal sNewVal$)
  m_sName = sNewVal
  MakeDirty
End Property

Public Property Get Category() As FRUIT_CATEGORY
  Category = m_eCategory
End Property
Public Property Let Category(ByVal eNewVal As FRUIT_CATEGORY)
  m_eCategory = eNewVal
  MakeDirty
End Property

Private Sub MakeDirty()
  ' sub to centrally set the dirty flag for this object.
  ' flag is set only if the object is not being loaded
  If Not m_bIsLoading Then m_bIsDirty = True
End Sub

Friend Property Let CollectionPointer(ByVal nNewVal&)
  If m_lpSuperCollection Then
    ' raise error here
  Else
    m_lpSuperCollection = nNewVal
  End If
End Property

Public Sub FireEvent()

  If MIEnumVARIANT.IsCollectionPointerValid(m_lpSuperCollection) > (-1) Then
    ResolvePointer(m_lpSuperCollection).FireEvent m_sKey, ObjPtr(Me)
  End If
  
End Sub

Private Function ResolvePointer(ByVal lpObj&) As CSuperCollection
  
  ' change the return type and the type for this variable to the class name
  ' of your class that wraps the collection object
  Dim oSC As CSuperCollection
  
  CopyMemory oSC, lpObj, 4
  Set ResolvePointer = oSC
  CopyMemory oSC, 0&, 4

End Function

