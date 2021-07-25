VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit

  Private m_lpObjCollection As Long

  Private m_nCurrentEnumRef As Long

  Implements IEnumVReDef.IEnumVARIANTReDef

Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  ' this property allows you to enumerate the collection with the For...Each syntax
  
  Dim oIev As IEnumVARIANTReDef
  Set oIev = Me
  
  ' replace the vtable entry for the Next function on the IEnumVARIANTReDef interface
  MIEnumVARIANT.ReplaceVtableEntry ObjPtr(oIev), 4, AddressOf MIEnumVARIANT.IEnumVARIANT_Next
    
  Set NewEnum = oIev

  Set oIev = Nothing
  
End Property

Public Property Get Count() As Long
  ' call back into the collection to get the count of the keys array
  
  If MIEnumVARIANT.IsCollectionPointerValid(m_lpObjCollection) > (-1) Then
    Count = MIEnumVARIANT.ResolveCollectionPointer(m_lpObjCollection).GetKeysCount()
  End If
End Property

Private Sub IEnumVARIANTReDef_Clone(lppIEnum As IEnumVReDef.IEnumVARIANTReDef)
  ' from the Win32 SDK:
  ' Using this function, a particular point in the enumeration sequence can
  ' be recorded, and then returned to at a later time. The returned enumerator
  ' is of the same actual interface as the one that is being cloned.
  '
  ' you would be hard pressed to find a way to use this in VB but it wasn't
  ' difficult to implement
    Dim oKeys As CKeys
  Set oKeys = New CKeys

  oKeys.InitializeForEnumeration m_lpObjCollection, m_nCurrentEnumRef
  
  Set lppIEnum = oKeys

  Set oKeys = Nothing

End Sub

Private Sub IEnumVARIANTReDef_GetItems(ByVal cElements As Long, aVariants As Variant, ByVal lpcElementsFetched As Long, lRetVal As Long)
  ' this method is called from the function IEnumVARIANT_Next in the MIEnumVARIANT bas module.
  ' it is an addition to the IEnumVARIANT interface provided in the 'ienumv.tlb' typelib.  this
  ' method provides a backdoor into this class via the 'this' pointer passed to IEnumVARIANT_Next by COM.
  ' in this way, the data in this class can be accessed from the IEnumVARIANT_Next function without jumping
  ' through hoops and creating circular references.  only the aVariants and lRetVal parameters are used in
  ' this implimentation.

  If MIEnumVARIANT.IsCollectionPointerValid(m_lpObjCollection) > (-1) Then
    MIEnumVARIANT.ResolveCollectionPointer(m_lpObjCollection).GetKeysForEnumeration m_nCurrentEnumRef, aVariants, lRetVal

    ' don't forget to increment the counter (passed ByRef from the CKeys object)
    m_nCurrentEnumRef = m_nCurrentEnumRef + 1

  Else
    lRetVal = S_FALSE
  End If

End Sub

Private Sub IEnumVARIANTReDef_Next(ByVal cElements As Long, aVariants As Variant, ByVal lpcElementsFetched As Long)
  ' This function is delegated to the IEnumVARIANT_Next function in our MIEnumVARIANT bas module.
  ' Right before the actual enumeration begins, we alter the vtable of this object so that it
  ' points to the address of the IEnumVARIANT_Next function in the bas module
End Sub

Private Sub IEnumVARIANTReDef_Reset()
  m_nCurrentEnumRef = 0
End Sub

Private Sub IEnumVARIANTReDef_Skip(ByVal cElements As Long)
  ' increment the enum counter the specified number of elements to skip
  m_nCurrentEnumRef = m_nCurrentEnumRef + cElements
End Sub

Friend Function InitializeForEnumeration(ByVal lpObject&, ByVal nEnumPointer&) As Boolean
  ' this function is called when this object is created.
  ' pass in the pointer to the parent collection object and the place to start the enumeration
  
  m_lpObjCollection = lpObject
  
  m_nCurrentEnumRef = nEnumPointer
  
End Function
