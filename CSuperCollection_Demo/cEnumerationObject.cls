VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEnumerationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit

  ' here is where the magic happens.  the IEnumVARIANT interface has been
  ' redefined in the 'ienumv.tlb' type library to make it VB friendly.
  Implements IEnumVReDef.IEnumVARIANTReDef
  
  ' enumerator counter for IEnumVARIANT implementation
  Private m_nCurrentEnumRef As Long
  
  ' change this if you want something other than a one based collection
  Private m_nBaseIndex As Long


  Private m_lpObjCollection As Long

  ' array used to lookup indexes in the m_aItems array in the CSuperCollection class
  Private m_anIndexLookp() As Long

Private Sub Class_Initialize()

  Dim oIev As IEnumVARIANTReDef
  Set oIev = Me
  
  ' replace the vtable entry for the Next function on this interface
  MIEnumVARIANT.ReplaceVtableEntry ObjPtr(oIev), 4, AddressOf MIEnumVARIANT.IEnumVARIANT_Next
  
  Set oIev = Nothing

End Sub

Private Sub Class_Terminate()
  ' free up the memory occupied by the array
  ReDim m_anIndexLookp(0) As Long
  Erase m_anIndexLookp
  
  ' decrement the enum counter in the collection object
  If MIEnumVARIANT.IsCollectionPointerValid(m_lpObjCollection) > (-1) Then
    MIEnumVARIANT.ResolveCollectionPointer(m_lpObjCollection).DecrementEnumCounter
  End If

End Sub

Private Sub IEnumVARIANTReDef_Clone(lppIEnum As IEnumVReDef.IEnumVARIANTReDef)
  ' from the Win32 SDK:
  ' Using this function, a particular point in the enumeration sequence can
  ' be recorded, and then returned to at a later time. The returned enumerator
  ' is of the same actual interface as the one that is being cloned.
  '
  ' you would be hard pressed to find a way to use this in VB but it wasn't
  ' difficult to implement
  Dim oEnumerator As cEnumerationObject
  
  Set oEnumerator = New cEnumerationObject
  
  oEnumerator.InitializeEnumeration m_lpObjCollection, m_nCurrentEnumRef, m_nBaseIndex, UBound(m_anIndexLookp), VarPtr(m_anIndexLookp(m_nBaseIndex))
  
  Set lppIEnum = oEnumerator
  
  Set oEnumerator = Nothing

End Sub

Private Sub IEnumVARIANTReDef_GetItems(ByVal cElements As Long, aVariants As Variant, ByVal lpcElementsFetched As Long, lRetVal As Long)
  ' this method is called from the function IEnumVARIANT_Next in the MIEnumVARIANT bas module.
  ' it is an addition to the IEnumVARIANT interface provided in the 'ienumv.tlb' typelib.  this
  ' method provides a backdoor into this class via the 'this' pointer passed to IEnumVARIANT_Next by COM.
  ' in this way, the data in this class can be accessed from the IEnumVARIANT_Next function without jumping
  ' through hoops and creating circular references.  only the aVariants and lRetVal parameters are used in
  ' this implimentation.

  If MIEnumVARIANT.IsCollectionPointerValid(m_lpObjCollection) > (-1) Then
    MIEnumVARIANT.ResolveCollectionPointer(m_lpObjCollection).GetItemsForIEnumVARIANT m_nCurrentEnumRef, m_anIndexLookp(), aVariants, lRetVal
      
    ' increment the module level counter
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
  ' reset the pointer to the beginning of the collection
  m_nCurrentEnumRef = m_nBaseIndex
End Sub

Private Sub IEnumVARIANTReDef_Skip(ByVal cElements As Long)
  ' increment the enum counter the specified number of elements to skip
  m_nCurrentEnumRef = m_nCurrentEnumRef + cElements
End Sub

Friend Function InitializeEnumeration(ByVal lpObjCollection&, ByVal nCurrentRef&, ByVal nLbound&, _
                                                                      ByVal nUbound&, ByVal lpArray&) As Boolean

  m_lpObjCollection = lpObjCollection

  m_nBaseIndex = nLbound
  
  If nLbound <= nUbound Then
    m_nCurrentEnumRef = nCurrentRef
  
    ReDim m_anIndexLookp(nLbound To nUbound) As Long
    
    CopyMemory ByVal VarPtr(m_anIndexLookp(nLbound)), ByVal lpArray, ((nUbound - nLbound) + 1) * 4&

  Else ' if the lbound is greater than the ubound, set the pointer to one greater than the lbound which will cause the enumerator to return immediately
    m_nCurrentEnumRef = 1
    
    ReDim m_anIndexLookp(0) As Long
  End If
  
End Function

