VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSuperCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit
  
  Public Enum COLLECTION_ADD_OPTIONS
    AddBeforePosition = 1&
    AddAfterPosition = 2&
    BeginningOfCollection = 3&
    EndOfCollection = 4&
  End Enum
  
  Public Enum COLLECTION_ERRORS
    AlreadyInitialized = &H80000500
    KeyNotFound = &H80000501
    MoveNotAllowedWhileInEnum = &H80000502
    InvalidDataType = &H80000503
    NoPositionSupplied = &H80000504
    Undefined = &H80000505
  End Enum
  
  Private Type KEY_LOOKUP
    bInitialized As Boolean
    sKey As String
    nIndex As Long
  End Type
  
  'local array for fast key lookups
  Private m_aKeys() As KEY_LOOKUP
  
  ' cross lookup array used to lookup items from keys as well as
  ' being copied into enumerator objects
  Private m_nHoldKeyArrayIndex As Long
  
  ' holds the current number of enumeration objects in effect for this instance of the collection
  Private m_nCurrentEnumOperations As Long
  
  Public Enum COMPARE_MODE
    [_CompareModeMin] = vbBinaryCompare
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    [_CompareModeMax] = vbTextCompare
  End Enum
  
  ' change this value if you want your keys to be case sensitive
  Private m_eCompareMode As COMPARE_MODE
  
  ' allows the property set for the item property to be activated
  Private m_bAllowItemAsignments As Boolean
  
  ' change this if you want something other than a one based collection
  Private m_nBaseIndex As Long
  
  ' since the collection can start at any number (including negitive numbers)
  ' we need a value that can be returned when a key was not found
  Private Const KEY_NOT_FOUND As Long = (-&HEFFFFFFF)
  
  ' array used to lookup indexes in the m_aItems array
  Private m_anIndexLookp() As Long
  
  Private Type PRIVATE_ITEM_TYPE
    bInitialized As Boolean
    ' set this member to the type of object you are collecting
    oItem As Variant
    sKey As String
  End Type
  
  'local variable to hold collection
  Private m_aItems() As PRIVATE_ITEM_TYPE
  
  Event GenericEvent(ByVal Index&, ByVal Key$)
  
' sets the compare mode to either text or binary for keys in the collection.
' this allows you to have case sensitive or case insensitive keys
Public Property Get CompareMode() As COMPARE_MODE
  CompareMode = m_eCompareMode
End Property
Public Property Let CompareMode(ByVal eNewVal As COMPARE_MODE)
  If Me.Count > 0 Then
    ' if there are already items in the collection we can't change the case sensitivity
    Err.Raise &H80000500, , "Collection already initialized"
    
  Else
    If (eNewVal >= [_CompareModeMin]) And (eNewVal <= [_CompareModeMax]) Then
      m_eCompareMode = eNewVal
    Else
      ' illeagal value
      Err.Raise 9 '<- Subscript out of range
    End If
  End If
End Property

' allows assignments to be made to the Item property
' If this is set, the user is responsible for
Public Property Get AllowItemAsignments() As Boolean
  AllowItemAsignments = m_bAllowItemAsignments
End Property
Public Property Let AllowItemAsignments(ByVal bNewVal As Boolean)
  m_bAllowItemAsignments = bNewVal
End Property

' sets the number at which the collection indexes start.  this allows you
' to have zero or ANY number (including negitive numbers) as the base index
Public Property Get BaseIndex() As Long
  BaseIndex = m_nBaseIndex
End Property
Public Property Let BaseIndex(ByVal nNewVal&)
  If Me.Count > 0 Then
    ' if there are already items in the collection we can't change the base index
    Err.Raise &H80000500, , "Collection already initialized"
    
  Else
    If nNewVal <> m_nBaseIndex Then
      EraseArrays

      m_nBaseIndex = nNewVal
            
      InitializeArrays
    End If
  End If
End Property

Public Property Get Keys() As CKeys
  ' creates a keys collection object that allows the keys to be enumerated

  Dim oKeys As CKeys
  Set oKeys = New CKeys

  oKeys.InitializeForEnumeration ObjPtr(Me), 0

  Set Keys = oKeys

  Set oKeys = Nothing

End Property

Friend Sub GetKeysForEnumeration(ByRef CurrentEnumRef&, aVariants As Variant, ByRef lRetVal&)
  ' callback function used to get the key info from the CKeys class.  we don't do anything fancy
  ' here like we do with enumerating the collection of objects.  instead, we just pass each item
  ' and increment the counter.  if the user deletes an item, it could whack the iteration but we
  ' just aren't going to handle that for the keys.  if keeping the integrity of the keys enumeration
  ' even if deletes are made is something you need, you will need to implement something similar
  ' to the enumeration object class in the CKeys class.

  ' if the current ref is within the array bounds, create, initialize and pass a new cKey object.
  ' otherwise, return S_FALSE to break out of the enumeration
  If CurrentEnumRef <= UBound(m_aKeys) Then
    Dim oKey As cKey
    Set oKey = New cKey

    With m_aKeys(CurrentEnumRef)
      oKey.Initialize .sKey, .nIndex
    End With

    Set aVariants = oKey

  Else
    lRetVal = S_FALSE
  End If

End Sub

Friend Function GetKeysCount() As Long
  ' return the count of the keys array
  If m_aKeys(0).bInitialized Then GetKeysCount = UBound(m_aKeys) + 1
End Function

Public Property Get Key(ByVal nIndex&) As String
  ' returns the Key value associated with the passed index
  
  ' make sure the items are within the bounds of the array
  If (nIndex >= m_nBaseIndex) And (nIndex <= UBound(m_anIndexLookp)) Then
    Key = m_aItems(m_anIndexLookp(nIndex)).sKey
  Else
    Err.Raise 9 '<- Subscript out of range
  End If
  
End Property

Public Property Get KeyExists(ByVal sKey$) As Long
  ' returns the index of an item if the key exists in the collection or KEY_NOT_FOUND
  KeyExists = FindKey(sKey)
End Property

Private Function FindKey(ByVal sKey$, Optional ByRef nKeyLookupIndex&, Optional ByVal bAddIfNotFound As Boolean, _
                                                                                                Optional ByRef nClosestMatch&) As Long
  ' - searches for the desired key in the collection.
  '   returns the array index number if found or KEY_NOT_FOUND (-&HEFFFFFFF) if not found.
  '   we return KEY_NOT_FOUND instead of zero when a key is not found to allow collections to have any base index
  ' - Also returns the index in the lookup key array if nKeyLookupIndex was passed.
  ' - Additionally, if bAddIfNotFound is set, a new lookup key array item is added with the passed key.
  ' - nClosestMatch is an [out] param that holds the index of the closest match that was found in the key lookup array.

  Dim i&, nUbound&, nLow&, nHigh&

  ' assume failure
  nKeyLookupIndex = (-1)
    
  If sKey = vbnullstring Then
    ' short circuit and return value of KEY_NOT_FOUND tells the caller no match was found
    FindKey = KEY_NOT_FOUND
  
  Else
    nUbound = UBound(m_aKeys)
    
    nHigh = nUbound
    
    ' check to see if this is the first item to be added to the collection
    If (i = nHigh) And (m_aKeys(i).bInitialized = False) Then
      ' the only place the bAddIfNotFound flag is set is in the Add method.  if the flag is set
      ' an item is being added to the collection and we need to add the corresponding entry in
      ' the key lookup array.
      If bAddIfNotFound Then
        With m_aKeys(i)
          ' set the flag indicating this member has been initialized
          .bInitialized = True
          
          nUbound = UBound(m_anIndexLookp)
          
          ' an item is being added so we need to add the index of the "about to be added" item to the lookup array
          If m_anIndexLookp(nUbound) <> KEY_NOT_FOUND Then
            .nIndex = nUbound + 1
          Else
            .nIndex = nUbound
          End If
          
          .sKey = sKey
        End With
      End If
      
      ' return value of KEY_NOT_FOUND tells the caller no match was found
      FindKey = KEY_NOT_FOUND
  
    Else
      
      Do
        ' divide and conquer!  Each time we loop devide the difference between the
        ' last items checked and search between the two indexes.  This is MUCH faster
        ' than looping through the entire list when dealing with a sorted array.
        i = nLow + ((nHigh - nLow) / 2)
        
        ' see how sKey relates to the current index....
        Select Case StrComp(sKey, m_aKeys(i).sKey, m_eCompareMode)
          Case 0
            FindKey = m_aKeys(i).nIndex
            
            ' return the lookup index
            nKeyLookupIndex = i
            nClosestMatch = i
            Exit Do
            
          Case (-1): nHigh = i - 1
          Case 1: nLow = i + 1
            
          Case Else
            ' raise error here
        End Select
    
        
        ' if the low search bound has become greater than the high search bound, the
        ' item does not exist in the array.  if the bAddIfNotFound flag is set, a new
        ' item is being added.  otherwise, just return the not found value.
        If nLow > nHigh Then
          If bAddIfNotFound Then
            ' check to see whether or not this item is initialized
            If m_aKeys(nUbound).bInitialized Then
              ReDim Preserve m_aKeys(nUbound + 1) As KEY_LOOKUP
          
              nUbound = UBound(m_aKeys)
            End If
            
            With m_aKeys(nUbound)
              .bInitialized = True
              
              ' bAddIfNotFound is only set from the Add method in this class.
              ' if this flag is set, a new item is about to be added to the m_anIndexLookp array
              ' So, we need to assign the index of the "about to be added" item to the lookup array member
              If m_anIndexLookp(UBound(m_anIndexLookp)) <> KEY_NOT_FOUND Then
                .nIndex = UBound(m_anIndexLookp) + 1
              Else
                .nIndex = UBound(m_anIndexLookp)
              End If
              
              .sKey = sKey
            End With
            
  
            ' see whether we should add this item above or below the item at index 'i'
            Select Case StrComp(sKey, m_aKeys(i).sKey, m_eCompareMode)
              
              ' return the lookup index
              Case (-1): nKeyLookupIndex = i
              Case 1: nKeyLookupIndex = i + 1
                
            End Select
            
            Call MoveKey(nUbound, nKeyLookupIndex)
            
          End If ' bAddIfNotFound
          
          ' return the item key lookup array that was the closest match
          nClosestMatch = i
          
          ' return value of KEY_NOT_FOUND tells the caller no match was found
          FindKey = KEY_NOT_FOUND
          Exit Do
        End If ' nLow > nHigh
    
      Loop
    End If ' (i = nHigh) And (m_aKeys(i).bInitialized = False) ... Else
  End If ' sKey = vbnullstring
  
End Function

Private Function MoveKey(ByVal nMovedKey&, ByVal nTargetPosition&) As Long
  ' moves an item in the key lookup array to a new index
  
  Dim i&, nLbound&, nUbound&, nStructLen&, abytBuffer() As Byte
  
  If nMovedKey = nTargetPosition Then
    ' don't need to move item
    MoveKey = nTargetPosition
  
  Else
    nStructLen = LenB(m_aKeys(nMovedKey))
  
    ReDim abytBuffer(nStructLen - 1) As Byte
  
    nUbound = UBound(m_aKeys)
    
    ' make sure the target position isn't greater than the ubound of the array
    If nTargetPosition > nUbound Then nTargetPosition = nUbound
  
    ' copy the item to be moved
    CopyMemory ByVal VarPtr(abytBuffer(0)), ByVal VarPtr(m_aKeys(nMovedKey)), nStructLen
  
    Select Case True
      Case nMovedKey < nTargetPosition
        ' grab all of the items below the moved one and shift them up
        CopyMemory ByVal VarPtr(m_aKeys(nMovedKey)), ByVal VarPtr(m_aKeys(nMovedKey + 1)), (nTargetPosition - nMovedKey) * nStructLen
    
        ' place the moved item into the position that used to be held by the target
        CopyMemory ByVal VarPtr(m_aKeys(nTargetPosition)), ByVal VarPtr(abytBuffer(0)), nStructLen
        
    
      Case nMovedKey > nTargetPosition
        ' grab all of the items above the moved one and shift them down
        CopyMemory ByVal VarPtr(m_aKeys(nTargetPosition + 1)), ByVal VarPtr(m_aKeys(nTargetPosition)), (nMovedKey - nTargetPosition) * nStructLen
    
        ' place the moved item into the position that used to be held by the target
        CopyMemory ByVal VarPtr(m_aKeys(nTargetPosition)), ByVal VarPtr(abytBuffer(0)), nStructLen
              
    End Select
    
    MoveKey = nTargetPosition
  End If
  
End Function

Public Sub ChangeKey(ByVal sOldKey$, ByVal sNewKey$)
  ' allow for key changes
  
  Dim nItem&, nClosestKeyLookupIndex&, nKeyLookupIndex&, nUbound&
  
  ' first, make sure the new key isn't already in the collection
  If FindKey(sNewKey, , , nClosestKeyLookupIndex) <> KEY_NOT_FOUND Then
    Err.Raise 457 '<- This key is already associated with an element of this collection

  Else
    ' find the index of the item with the given key.
    ' this also returns the index into the lookup item array for the key
    nItem = FindKey(sOldKey, nKeyLookupIndex)
  
    If nItem <> KEY_NOT_FOUND Then
      With m_aItems(m_anIndexLookp(nItem))
        .sKey = sNewKey
        
        ' the one place where we manipulate the key in the child object
        .oItem.Key = sNewKey
      End With
      
      SetChangedKeyPosition nKeyLookupIndex, sNewKey, nClosestKeyLookupIndex, nUbound, nItem
      
    Else
      Err.Raise KeyNotFound, , "Key not found."
    End If ' nItem <> KEY_NOT_FOUND
  End If ' FindKey(sNewKey) <> KEY_NOT_FOUND
  
End Sub

Public Sub ChangeKeyForIndex(ByVal nIndex$, ByVal sNewKey$)
  ' allow for key changes
  
  Dim i&, nClosestKeyLookupIndex&, nKeyLookupIndex&, nUbound&
  
  If (nIndex >= m_nBaseIndex) And (nIndex <= UBound(m_anIndexLookp)) Then
    ' make sure the new key isn't already in the collection
    If FindKey(sNewKey, , , nClosestKeyLookupIndex) <> KEY_NOT_FOUND Then
      Err.Raise 457 '<- This key is already associated with an element of this collection
    
    Else
      With m_aItems(m_anIndexLookp(nIndex))
        .sKey = sNewKey
        
        ' the one place where we manipulate the key in the child object
        .oItem.Key = sNewKey
      End With
      
      nUbound = UBound(m_aKeys)
      
      nKeyLookupIndex = (-1)
      
      ' find the item in the lookup array
      For i = 0 To nUbound
        If m_aKeys(i).nIndex = nIndex Then
          nKeyLookupIndex = i
          Exit For
        End If
      Next
      
      
      SetChangedKeyPosition nKeyLookupIndex, sNewKey, nClosestKeyLookupIndex, nUbound, nIndex
      
      
    End If ' FindKey(sNewKey, , , nClosestKeyLookupIndex) <> KEY_NOT_FOUND
    
  Else ' (nIndex >= m_nBaseIndex) And (nIndex <= UBound(m_anIndexLookp))
    ' raise error here
  End If ' (nIndex >= m_nBaseIndex) And (nIndex <= UBound(m_anIndexLookp))
  
End Sub

Private Sub SetChangedKeyPosition(ByVal nKeyLookupIndex&, ByVal sNewKey$, ByVal nClosestKeyLookupIndex&, ByVal nUbound&, ByVal nIndex&)
  ' handles moving and setting the values for the item in the key lookup array for both ChangKey subs
  
  If (nKeyLookupIndex <> (-1)) And (sNewKey <> vbnullstring) Then
    m_aKeys(nKeyLookupIndex).sKey = sNewKey
    
    ' see whether we should add this item above or below the item at index 'nClosestKeyLookupIndex'
    Select Case StrComp(sNewKey, m_aKeys(nClosestKeyLookupIndex).sKey, m_eCompareMode)
      Case (-1): nClosestKeyLookupIndex = nClosestKeyLookupIndex '- 1
      Case 1: nClosestKeyLookupIndex = nClosestKeyLookupIndex + 1
    End Select
    
    nUbound = UBound(m_aKeys)

    If nClosestKeyLookupIndex > (nUbound + 1) Then nClosestKeyLookupIndex = nUbound + 1
    
    Call MoveKey(nKeyLookupIndex, nClosestKeyLookupIndex)
  
  
  Else
  
    If nKeyLookupIndex <> (-1) Then
      ' now remove the item from the key lookup array....
      Call MoveKey(nKeyLookupIndex, nUbound)
    
      ' if the ubound is greater than zero, preserve the data....
      If nUbound Then
        ReDim Preserve m_aKeys(nUbound - 1) As KEY_LOOKUP
      Else
        ReDim m_aKeys(0) As KEY_LOOKUP
      End If
    End If
    
    
    ' add the new key to the lookup array
    If sNewKey <> vbnullstring Then
      
      ' check to see whether or not this item is initialized
      If m_aKeys(nUbound).bInitialized Then
        ReDim Preserve m_aKeys(nUbound + 1) As KEY_LOOKUP
    
        nUbound = UBound(m_aKeys)
      End If
      
      With m_aKeys(nUbound)
        .bInitialized = True
        .nIndex = nIndex
        .sKey = sNewKey
      End With
      
      ' see whether we should add this item above or below the item at index 'nClosestKeyLookupIndex'
      Select Case StrComp(sNewKey, m_aKeys(nClosestKeyLookupIndex).sKey, m_eCompareMode)
        Case (-1): nKeyLookupIndex = nClosestKeyLookupIndex '- 1
        Case 1: nKeyLookupIndex = nClosestKeyLookupIndex + 1
      End Select
      
      If nClosestKeyLookupIndex > UBound(m_aKeys) Then nClosestKeyLookupIndex = UBound(m_aKeys)
     
      Call MoveKey(nUbound, nKeyLookupIndex)
      
    End If ' sNewKey <> vbnullstring
    
  End If ' (nKeyLookupIndex <> (-1)) And (sNewKey <> vbnullstring)

End Sub

Public Sub Add(vItem As Variant, Optional ByVal sKey$, Optional ByVal eInsert As COLLECTION_ADD_OPTIONS = EndOfCollection, _
                                                                            Optional ByVal nPosition& = KEY_NOT_FOUND)
  
  Dim nLbound&, nUbound&, nKeyArrayIndex&, nFindKey&
  
  nLbound = m_nBaseIndex
  nUbound = UBound(m_anIndexLookp)
  
  ' make sure the specified position is within the bounds of the array
  If nPosition <> KEY_NOT_FOUND Then
    Select Case nPosition
      Case Is < nLbound, Is > nUbound
        Err.Raise 9 '<- Subscript out of range
        
        Exit Sub
        
      Case nLbound
        If eInsert = AddBeforePosition Then eInsert = BeginningOfCollection
        
      Case nUbound
        If eInsert = AddAfterPosition Then eInsert = EndOfCollection
        
    End Select
    
  ElseIf ((eInsert = AddBeforePosition) Or (eInsert = AddAfterPosition)) And (nPosition = KEY_NOT_FOUND) Then
    Err.Raise NoPositionSupplied, , "No position supplied."
    
    Exit Sub
  End If
  
  If sKey <> vbnullstring Then
    nFindKey = FindKey(sKey, nKeyArrayIndex, True)
  Else
    nFindKey = KEY_NOT_FOUND
  End If
  
  ' first thing to do is make sure the key does not already exist in the collection
  If nFindKey <> KEY_NOT_FOUND Then
    Err.Raise 457 '<- This key is already associated with an element of this collection
    
  Else
    ' check to see whether or not this item is initialized
    If m_anIndexLookp(nUbound) <> KEY_NOT_FOUND Then
      ReDim Preserve m_anIndexLookp(m_nBaseIndex To nUbound + 1) As Long
      
      If m_aItems(UBound(m_aItems)).bInitialized Then
        ReDim Preserve m_aItems(UBound(m_aItems) + 1) As PRIVATE_ITEM_TYPE
      End If
      
      nUbound = UBound(m_anIndexLookp)
    End If
    
    m_anIndexLookp(nUbound) = UBound(m_aItems)
    
    ' set the object and the key into the array
    With m_aItems(m_anIndexLookp(nUbound))
      .bInitialized = True
      
      If VarType(vItem) = vbObject Then
        Set .oItem = vItem
      Else
        .oItem = vItem
      End If
      
      .sKey = sKey
    End With
        
    
    ' set module level flag to avoid lookup hit in the MoveItem method
    m_nHoldKeyArrayIndex = nKeyArrayIndex
    
    ' if we were directed to add the item before or after another item, move it to the desired location
    Select Case eInsert
      Case BeginningOfCollection
        Me.MoveItem nUbound, nLbound
        
      Case AddBeforePosition
        Me.MoveItem nUbound, nPosition
      
      Case AddAfterPosition
        Me.MoveItem nUbound, nPosition + 1
    
    End Select
    
  End If

End Sub

Public Property Get Count() As Long
  ' count of the objects in the collection.
  If m_anIndexLookp(UBound(m_anIndexLookp)) <> KEY_NOT_FOUND Then
    Count = (UBound(m_anIndexLookp) + 1) - m_nBaseIndex
  End If
End Property

Public Sub Remove(vIndexKey As Variant)
  ' removes the item at the specified index or associated
  ' with the specified key from the collection
  
  Dim nUbound&, nItem&, nKeyLookupIndex&, bFound As Boolean
  
  ' determine whether this is an index or a string key
  Select Case VarType(vIndexKey)
    Case vbSingle, vbDouble, vbByte, vbLong, vbInteger
      nItem = vIndexKey
      
      If (nItem >= m_nBaseIndex) And (nItem <= UBound(m_anIndexLookp)) Then
        bFound = True
        Call FindKey(m_aItems(m_anIndexLookp(nItem)).sKey, nKeyLookupIndex)
      Else
        Err.Raise 9 '<- Subscript out of range
      End If
      
    ' if it's a string, look for it in the array
    Case vbString
      nItem = FindKey(CStr(vIndexKey), nKeyLookupIndex)
      
      If nItem <> KEY_NOT_FOUND Then bFound = True
  
    Case Else
      Err.Raise InvalidDataType, , "Invalid datatype"
      
  End Select
  
  
  If bFound Then
    ' if we found the item, move it to the end of the array.
    nUbound = MoveItemToEndOfArray(nItem)
    
    ' destroy the reference
    With m_aItems(m_anIndexLookp(nUbound))
      If VarType(.oItem) = vbObject Then
        Set .oItem = Nothing
      Else
        .oItem = 0
      End If
      
      .bInitialized = False
      .sKey = vbnullstring
    End With
    
    ' if the ubound is greater than the collection base number, preserve the data....
    If nUbound > m_nBaseIndex Then
      ReDim Preserve m_anIndexLookp(m_nBaseIndex To nUbound - 1) As Long
    Else
      ReDim m_anIndexLookp(m_nBaseIndex To m_nBaseIndex) As Long
      m_anIndexLookp(m_nBaseIndex) = KEY_NOT_FOUND
    End If
  
    ' now remove the key lookup array item
    If nKeyLookupIndex > (-1) Then
      nUbound = UBound(m_aKeys)
    
      nItem = MoveKey(nKeyLookupIndex, nUbound)
    
      ' if the ubound is greater than zero, preserve the data....
      If nUbound Then
        ReDim Preserve m_aKeys(nUbound - 1) As KEY_LOOKUP
      Else
        ReDim m_aKeys(0) As KEY_LOOKUP
      End If
    End If
  Else
    Err.Raise KeyNotFound, , "Key not found."
  End If
  
End Sub

Public Property Get Item(vIndexKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
  ' retrieves the specified item from the collection.
  ' this is the default property of the collection interface
  ' which allows us to access an item via the following syntax:
  '
  '  Dim oObject As cCollectionChildObject
  '
  '  Set oObject = oCollection('index or Key')
  '
  
  Dim nItem&, bFound As Boolean

  ' determine whether this is an index or a string key
  Select Case VarType(vIndexKey)
    Case vbSingle, vbDouble, vbByte, vbLong, vbInteger
      nItem = vIndexKey
      
      If (nItem >= m_nBaseIndex) And (nItem <= UBound(m_anIndexLookp)) Then
        If m_anIndexLookp(nItem) <> KEY_NOT_FOUND Then bFound = True
      End If
      
    ' if it's a string, look for it in the array
    Case vbString
      nItem = FindKey(CStr(vIndexKey))
      
      If nItem <> KEY_NOT_FOUND Then bFound = True

    Case Else
      Err.Raise InvalidDataType, , "Invalid datatype"

  End Select

  If bFound Then
    With m_aItems(m_anIndexLookp(nItem))
      If VarType(.oItem) = vbObject Then
        Set Item = .oItem
      Else
        Item = .oItem
      End If
    End With
    'Set Item = m_aItems(m_anIndexLookp(nItem)).oItem
  Else
    Err.Raise KeyNotFound, , "Key not found."
  End If
  
End Property
Public Property Set Item(vIndexKey As Variant, vNewItem As Variant)

  If m_bAllowItemAsignments Then
  
    Dim nItem&, bFound As Boolean

    ' determine whether this is an index or a string key
    Select Case VarType(vIndexKey)
      Case vbSingle, vbDouble, vbByte, vbLong, vbInteger
        nItem = vIndexKey

        If (nItem >= m_nBaseIndex) And (nItem <= UBound(m_anIndexLookp)) Then bFound = True

      ' if it's a string, look for it in the array
      Case vbString
        nItem = FindKey(CStr(vIndexKey))

        If nItem <> KEY_NOT_FOUND Then bFound = True

      Case Else
        Err.Raise InvalidDataType, , "Invalid datatype"
    
    End Select

    If bFound Then
'      ' this is the one place where we diverge from a generic encapsulated object model.
'      ' we need to check the key property of the new object to see whether or not we need to
'      ' change our internal key.  if the keys are not the same, change our copy.
'      If Not vNewItem Is Nothing Then
'        If StrComp(vNewItem.Key, m_aItems(nItem).sKey, m_eCompareMode) Then
'          Me.ChangeKeyForIndex nItem, vNewItem.Key
'        End If
'      End If
      
      With m_aItems(m_anIndexLookp(nItem))
        If VarType(.oItem) = vbObject Then
          Set .oItem = vNewItem
        Else
          .oItem = vNewItem
        End If
      End With
      'Set m_aItems(m_anIndexLookp(nItem)).oItem = vNewItem
      
    Else
      Err.Raise KeyNotFound, , "Key not found."
    End If
  
  Else
    ' not allowint item assignments.  raise the same error as the VB collection
    Err.Raise 438
  End If
  
End Property


' #############  Begin Required for enumeration block  ###################
Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
  ' this property allows you to enumerate the collection with the For...Each syntax

  Dim i&, nBase&, nUbound&, oEnumerator As cEnumerationObject
  
  Set oEnumerator = New cEnumerationObject
  
  nUbound = UBound(m_anIndexLookp)
  
  If Me.Count Then
    nBase = m_nBaseIndex
  Else
    nBase = nUbound + 1
  End If
    
  oEnumerator.InitializeEnumeration ObjPtr(Me), nBase, nBase, nUbound, VarPtr(m_anIndexLookp(m_nBaseIndex))
  
  
  m_nCurrentEnumOperations = m_nCurrentEnumOperations + 1
  
  ' return the enumerator object's IEnumVARIANTReDef interface
  Set NewEnum = oEnumerator
  
  ' clean up
  Set oEnumerator = Nothing
  
End Property

Friend Sub GetItemsForIEnumVARIANT(ByRef CurrentEnumRef&, anLookup() As Long, aVariants As Variant, ByRef lRetVal&)
  ' this method is called from the function IEnumVARIANT_Next in the MIEnumVARIANT bas module.
  ' it is an addition to the IEnumVARIANT interface provided in the 'ienumv.tlb' typelib.  this
  ' method provides a backdoor into this class via the 'this' pointer passed to IEnumVARIANT_Next by COM.
  ' in this way, the data in this class can be accessed from the IEnumVARIANT_Next function without jumping
  ' through hoops and creating circular references.  only the aVariants and lRetVal parameters are used in
  ' this implimentation.

  Dim i&, j&, nUbound&, nUboundLookup&, nDifference&, bResetArray As Boolean, bExitDo As Boolean
  
  ' assume no item will be returned
  lRetVal = S_FALSE
  
  nUboundLookup = UBound(anLookup)
  
  ' if the counter is greater than the upper bounds of the
  ' array, return S_FALSE ('1') to break out of the enumeration
  If CurrentEnumRef <= nUboundLookup Then
    
    ' see if there have been any changes to the array in the parent collection while we were gone....
    nUbound = UBound(m_anIndexLookp)
      
    ' what we are concerened with is items that may have been added to or removed from the collection.
    ' if the array upper bound in the enumerator object is less than that of the collection, we need to
    ' investigate further.  if the upper bounds of the array in the enumerator object is greater than
    ' that of the collection, it would indicate that an item had been removed from the collection.  we
    ' handle this in the "Do" loop later in this routine....
    If nUboundLookup < nUbound Then
      ' see if the items pointed to by the array match.  if they do, it would indicate that whatever happened
      ' to the array was above the current index and should thus be reflected in the enumerator object.  we
      ' will redimention the array in the enumerator object and copy the current array from the collection into it
      If m_anIndexLookp(CurrentEnumRef) <> anLookup(CurrentEnumRef) Then
      
        ' an item may have been added or removed below this item.  let's look forward to see if we can find our item.
        ' if we find it, set the pointer to the index at which the item was found set the flag to reset the array.
        ' if we don't find the item, a move or other complex operation could have been preformed in which case, we
        ' will just proceed since we could look forever trying to figure out what had happened.   there is a count of
        ' current enumeration operations held in the m_nCurrentEnumOperations variable.  at this writing, this var is
        ' checked in the MoveItem method of this class and no moves are allowed while enumeration operations are in
        ' progress.
          
          nDifference = nUbound = nUboundLookup
          
          For j = CurrentEnumRef To CurrentEnumRef + nDifference
            If m_anIndexLookp(j) = anLookup(CurrentEnumRef) Then
              CurrentEnumRef = j
              bResetArray = True
              Exit For
            End If
          Next
        
      Else
        ' the indexes match so we will just reset the array
        bResetArray = True
      End If
      
      If bResetArray Then
        ' copy the new array data into the array in the calling object
        ReDim anLookup(m_nBaseIndex To nUbound) As Long
        
        CopyMemory ByVal VarPtr(anLookup(m_nBaseIndex)), ByVal VarPtr(m_anIndexLookp(m_nBaseIndex)), ((nUbound - m_nBaseIndex) + 1) * 4&
      End If
      
    End If
    
    ' here we look to see if the item pointed to by the lookup array in the enumerator object is still
    ' initialized.  if it isn't, we loop through the lookup array until we find an item that is initialized
    ' or we get to the end of the the lookup array.  we also change the CurrentEnumRef pointer as needed
    i = CurrentEnumRef
    
    Do
      With m_aItems(anLookup(i))
        If .bInitialized Then
          If VarType(.oItem) = vbObject Then
            Set aVariants = .oItem
          Else
            aVariants = .oItem
          End If
          
          If i <> CurrentEnumRef Then
            CurrentEnumRef = CurrentEnumRef + (i - CurrentEnumRef)
          End If
          
          lRetVal = S_OK
          bExitDo = True
        End If
      End With
      
      If bExitDo Then Exit Do
      
      i = i + 1
    Loop Until i > nUboundLookup
  End If
    
End Sub

Friend Sub DecrementEnumCounter()
  m_nCurrentEnumOperations = m_nCurrentEnumOperations - 1
End Sub
' #############  End Required for enumeration block  ###################

Private Sub Class_Initialize()
  
  MIEnumVARIANT.AddPointerToLookupList ObjPtr(Me)
  
  ' set the default collection base
  m_nBaseIndex = 1
  
  ' call the sub to initialize the arrays
  InitializeArrays
End Sub

Private Sub Class_Terminate()
  ' terminate the class by looping through the array and setting each
  ' object to Nothing and releasing the memory held by the arrays
  
  MIEnumVARIANT.RemovePointerFromLookupList ObjPtr(Me)
  
  ' call the sub that erases the arrays
  EraseArrays
End Sub

Private Sub InitializeArrays()
  ' initialize the arrays that will hold the collection and the key lookups
  ReDim m_anIndexLookp(m_nBaseIndex To m_nBaseIndex) As Long
  m_anIndexLookp(m_nBaseIndex) = KEY_NOT_FOUND
  
  ReDim m_aItems(0) As PRIVATE_ITEM_TYPE
  ReDim m_aKeys(0) As KEY_LOOKUP
End Sub

Private Sub EraseArrays()
  
  Dim i&
  
  ReDim m_anIndexLookp(0) As Long
  Erase m_anIndexLookp
  
  ' loop through the array data
  ' and set each item to nothing
  For i = 0 To UBound(m_aItems)
    With m_aItems(i)
      If VarType(.oItem) = vbObject Then Set .oItem = Nothing
    End With
    'Set m_aItems(i).oItem = Nothing
  Next
  
  ReDim m_aItems(0) As PRIVATE_ITEM_TYPE
  Erase m_aItems
  
  ReDim m_aKeys(0) As KEY_LOOKUP
  Erase m_aKeys
  
End Sub

Public Function Clear()
  ' clear out the arrays and then reinitialize them
  EraseArrays
  InitializeArrays
End Function

Public Function MoveItem(ByVal vMovedItem As Variant, ByVal vTargetPosition As Variant) As Long
  ' moves an item in the collection to a new index in response to the user rearranging items
  
  Dim nMovedItem&, nTargetPosition&, nLbound&, nUbound&
  Dim i&, nKeyArrayIndex&, nChangedCount&, nChangeNum&, nIndex&
  Dim bNoMove As Boolean, nHoldForMove&
  
  ' only allow moves if we are not currently in any enumeration operations
  If m_nCurrentEnumOperations = 0 Then
    nKeyArrayIndex = KEY_NOT_FOUND
    
    ' determine whether this is an index or a string key for each param....
    Select Case VarType(vMovedItem)
      Case vbSingle, vbDouble, vbByte, vbLong, vbInteger
        nMovedItem = vMovedItem
        m_nHoldKeyArrayIndex = KEY_NOT_FOUND
        
      ' if it's a string, look for it in the array
      Case vbString
        nMovedItem = FindKey(CStr(vMovedItem), nKeyArrayIndex)
        
        If nMovedItem = KEY_NOT_FOUND Then
          Err.Raise KeyNotFound, , "Key not found."
        Else
          ' will only hit here if we are moving an item at the users request.
          ' this is because all of the internal calls to the MoveItem method use the array index.
          m_nHoldKeyArrayIndex = nKeyArrayIndex
        End If
  
      Case Else
        Err.Raise InvalidDataType, , "Invalid datatype"
        
    End Select
    
    Select Case VarType(vTargetPosition)
      Case vbSingle, vbDouble, vbByte, vbLong, vbInteger
        nTargetPosition = vTargetPosition
        
      ' if it's a string, look for it in the array
      Case vbString
        nTargetPosition = FindKey(CStr(vTargetPosition))
        
        If nTargetPosition = KEY_NOT_FOUND Then
          Err.Raise KeyNotFound, , "Key not found."
        End If
  
      Case Else
        Err.Raise InvalidDataType, , "Invalid datatype"
        
    End Select
    ' .....
    
    
    nLbound = m_nBaseIndex
    nUbound = UBound(m_anIndexLookp)
    
    ' make sure the items are within the bounds of the array
    Select Case True
      Case (nMovedItem > nUbound), (nTargetPosition > nUbound), (nMovedItem < nLbound), (nTargetPosition < nLbound)
        bNoMove = True
      
        Err.Raise 9 '<- Subscript out of range
        
    
      Case nMovedItem = nTargetPosition
        ' no move necessary.  return the moved item index
        bNoMove = True
        
        MoveItem = nMovedItem
  
    End Select
  
    If Not bNoMove Then
      nUbound = UBound(m_anIndexLookp)
    
      ' we need to know the index of the item in the keys array.  if a lookup was called
      ' already in the chain of calls leading to this point, the index will be cached in the
      ' module level variable.
      If m_nHoldKeyArrayIndex = KEY_NOT_FOUND Then
        Call FindKey(m_aItems(m_anIndexLookp(nMovedItem)).sKey, nKeyArrayIndex)
        m_nHoldKeyArrayIndex = nKeyArrayIndex
      End If
    
    
      ' copy the item to be moved
      nHoldForMove = m_anIndexLookp(nMovedItem)
    
      Select Case True
        Case nMovedItem < nTargetPosition
          ' grab all of the items below the moved one and shift them up
          CopyMemory ByVal VarPtr(m_anIndexLookp(nMovedItem)), ByVal VarPtr(m_anIndexLookp(nMovedItem + 1)), (nTargetPosition - nMovedItem) * 4&
      
          ' place the moved item into the position that used to be held by the target
          m_anIndexLookp(nTargetPosition) = nHoldForMove
          
          
          ' now that we have moved the items, we need to fix the indexes held in the
          ' keys array.  we try to limit our time looping by keeping track of the
          ' number of indexes we *fix*.  once we have fixed all the indexes that need to
          ' be fixed, we don't need to loop any more.
          nChangeNum = nTargetPosition - nMovedItem
          
          For i = 0 To UBound(m_aKeys)
            nIndex = m_aKeys(i).nIndex
  
            If (nIndex > nMovedItem) And (nIndex <= nTargetPosition) Then
              m_aKeys(i).nIndex = nIndex - 1
              
              nChangedCount = nChangedCount + 1
              
              If nChangedCount >= nChangeNum Then Exit For
            End If
          Next
      
      
        Case nMovedItem > nTargetPosition
          ' grab all of the items above the moved one and shift them down
          CopyMemory ByVal VarPtr(m_anIndexLookp(nTargetPosition + 1)), ByVal VarPtr(m_anIndexLookp(nTargetPosition)), (nMovedItem - nTargetPosition) * 4&
      
          ' place the moved item into the position that used to be held by the target
          m_anIndexLookp(nTargetPosition) = nHoldForMove
          
          ' now that we have moved the items, we need to fix the indexes held in the
          ' keys array.  we try to limit our time looping by keeping track of the
          ' number of indexes we *fix*.  once we have fixed all the indexes that need to
          ' be fixed, we don't need to loop any more.
          nChangeNum = nMovedItem - nTargetPosition
          
          For i = 0 To UBound(m_aKeys)
            nIndex = m_aKeys(i).nIndex
            
            If (nIndex >= nTargetPosition) And (nIndex < nMovedItem) Then
            m_aKeys(i).nIndex = nIndex + 1
            
            nChangedCount = nChangedCount + 1
            
            If nChangedCount = nChangeNum Then Exit For
            End If
          Next
        
        Case Else
          ' raise error here
          Err.Raise Undefined, , "Undifined error"
      End Select
      
      
      ' handle setting the index for the item that we just moved
      If m_nHoldKeyArrayIndex <> KEY_NOT_FOUND Then
        If m_nHoldKeyArrayIndex > (-1) Then m_aKeys(m_nHoldKeyArrayIndex).nIndex = nTargetPosition
      End If
      
      
      MoveItem = nTargetPosition
    End If
    
    ' reset the array index holder
    m_nHoldKeyArrayIndex = KEY_NOT_FOUND
  
  Else
    Err.Raise MoveNotAllowedWhileInEnum, , "Move not allowed while in For...Each loop."
  End If ' m_nCurrentEnumOperations = 0
  
End Function

Private Function MoveItemToEndOfArray(ByVal nItem&) As Long
  ' moves item to the end of the array.  this function is
  ' used to move items to be deleted from the array
  
  Dim nUbound&, nHoldForMove&
  Dim i&, nChangeNum&, nChangedCount&, nIndex&

  nUbound = UBound(m_anIndexLookp)

  If nItem < nUbound Then
    ' copy the item to be moved
    nHoldForMove = m_anIndexLookp(nItem)

    ' grab all of the items above the moved one and shift them up
    CopyMemory ByVal VarPtr(m_anIndexLookp(nItem)), ByVal VarPtr(m_anIndexLookp(nItem + 1)), (nUbound - nItem) * 4&

    ' blast the moved item into the last position
    m_anIndexLookp(nUbound) = nHoldForMove
  End If

        
  ' now that we have moved the items, we need to fix the indexes held in the
  ' keys array.  we try to limit our time looping by keeping track of the
  ' number of indexes we *fix*.  once we have fixed all the indexes that need to
  ' be fixed, we don't need to loop any more.
  nChangeNum = nUbound - nItem
  
  For i = 0 To UBound(m_aKeys)
    nIndex = m_aKeys(i).nIndex
    
    If (nIndex > nItem) And (nIndex <= nUbound) Then
      m_aKeys(i).nIndex = nIndex - 1
      
      nChangedCount = nChangedCount + 1
      
      If nChangedCount = nChangeNum Then Exit For
    End If
  Next


  MoveItemToEndOfArray = nUbound
  
End Function

' this should actually go in the class that wraps the collection object
' but since we didn't wrap the collection for this demo, we stuck it here.
Friend Sub FireEvent(ByVal sKey$, ByVal lpObject&)

  Dim i&, nIndex&
  
  nIndex = FindKey(sKey)
  
  If nIndex = KEY_NOT_FOUND Then
    ' no key was entered for this item when it was added.  we have to loop through the array and match the pointer...
    For i = m_nBaseIndex To UBound(m_anIndexLookp)
      If ObjPtr(m_aItems(m_anIndexLookp(i)).oItem) = lpObject Then
        nIndex = i
        Exit For
      End If
    Next
  End If
  
  If nIndex <> KEY_NOT_FOUND Then RaiseEvent GenericEvent(nIndex, sKey)
  
End Sub
