Attribute VB_Name = "MIEnumVARIANT"
  Option Explicit
  ' this module handles enumerations for all collection classes where cusotm enumerations are implemented.
  ' when a the NewEnum method is called in a collection class, the class creates an enumeration object and
  ' passes a pointer to the IEnumVARIANT interface on that object as the return value to the call.  The
  ' vtable of the the enumeration object is altered so that the Next method of that object is called here
  ' in this BAS module (the IEnumVARIANT_Next function below).  every time the next method is called, we
  ' call back into the referenced instance of the enumeration object using the 'this' pointer passed to
  ' the function by COM.  the enumeration object then calls into the collection object to retrieve the next
  ' item to be enumerated and passes it back to the code in the IEnumVARIANT_Next function to then be
  ' passed back to the caller.
  

  ' IMPORTANT NOTE:  Much of the code in the IEnumVARIANT_Next, ReplaceVtableEntry and MapErr functions
  ' in this module can be found in Bruce McKinney's book "Hardcore Visual Basic".  I have made some
  ' alterations but all in all, there was not much room for improvement.  What I *have* done is eleminated
  ' the need for retreiving a reference to the correct instance of the calling enumeration object by writing
  ' a typelib redefinition of the IEnumVARIANT interface that allows us to use the passed 'this' pointer to
  ' call back into the correct instance of the enumeration object.

  
  Public Const S_OK As Long = &H0&
  Public Const S_FALSE As Long = &H1&
    
  Public Const KEY_NOT_FOUND As Long = (-&HEFFFFFFF)
  
  Public Const ERROR_ALREADY_EXISTS As Long = &H800000B7
  
  Public Const PAGE_EXECUTE_READWRITE As Long = &H40&

  Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress&, ByVal dwSize&, ByVal flNewProtect&, lpflOldProtect&) As Long

  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
  Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDest As Any, ByVal cBytes&)

  ' array to hold pointers to all instances of the collection class.  this allows us to verify
  ' the existence of the collection object before making a callback from an enumerator object.
  Private m_alpCollections() As Long
  
Public Function IEnumVARIANT_Next(ByVal this As IEnumVReDef.IEnumVARIANTReDef, ByVal cElements As Long, _
                                                      avObjects As Variant, ByVal nNumFetched As Long) As Long
                                                      
  ' this - Object pointer to the IEnumVARIANT interface of the collection being enumerated
  ' cElements - Count of variants requested for return
  ' avObjects - Array of variants to hold the requested variants
  ' nNumFetched - Pointer to number of variants actually returned
                      
  Dim i&, lpVariantArray&, lRet&, nFetched&, nDummy&
  Dim vTmp As Variant, vEmpty As Variant

  ' In C++, the first argument of an object method call is the object pointer--known as the
  ' 'this' pointer.  It is normally hidden in VB but since we have altered the collection
  ' class' vtable
  ' so that the 'Next' method call will be envoked here, we must account for it.  In this case, it
  ' points to the IEnumVARIANTReDef interface that we implemented and passed to the NewEnum
  ' property in our collection class. Because a new method was added to the IEnumVARIANT
  ' interface in the 'ienumv.tlb' typelib, we can use the 'this' pointer to callback into
  ' the object and get the items from our implementation of the collection.
  
  On Error Resume Next
  
  ' Get the address of the first variant in array
  lpVariantArray = VarPtr(avObjects)
  
  ' Iterate through each requested variant
  For i = 1 To cElements
    ' Call the method that was added to the IEnumVARIANT interface in the typelib.
    ' nDummy is a space filler since the two params it occupies are not used in this implementation.
    ' lRet is the return value from the GetItems call passed byref
    this.GetItems nDummy, vTmp, nDummy, lRet

    ' If failure or nothing fetched, we're done
    If (Err) Or (lRet = 1) Then Exit For
    ' Copy variant to current array position
    CopyMemory ByVal lpVariantArray, vTmp, 16&
    ' Empty work variant without destroying its object or string
    CopyMemory vTmp, vEmpty, 16&
    ' Count the variant and point to the next one
    nFetched = nFetched + 1
    lpVariantArray = lpVariantArray + 16
  Next
  
  ' If error caused termination, undo what we did
  If Err.Number Then
    ' Iterate back, emptying the invalid fetched variants
    For i = i To 1 Step -1
      ' Copy variant to current array position
      CopyMemory vTmp, ByVal lpVariantArray, 16
      ' Empty work variant, destroying any object or string
      vTmp = Empty
      ' Empty array variant without destroying any object or string
      CopyMemory ByVal lpVariantArray, vEmpty, 16
      ' Point to previous array element
      lpVariantArray = lpVariantArray - 16
    Next
    ' Convert error to COM format
    IEnumVARIANT_Next = MapErr(Err)
    ' Return 0 as the number fetched after error
    If nNumFetched Then CopyMemory ByVal nNumFetched, ByVal 0&, 4&
    
  Else
    ' If nothing fetched, break out of enumeration
    If nFetched = 0 Then IEnumVARIANT_Next = S_FALSE '<-- the value of S_FALSE is &H1&.  Confusing, eh'?
    ' Copy the actual number fetched to the pointer to fetched count
    If nNumFetched Then CopyMemory ByVal nNumFetched, nFetched, 4&
  End If
  
End Function

' Put the function address (callback) directly into the object v-table
Public Function ReplaceVtableEntry(ByVal lpObj As Long, ByVal nEntry As Integer, ByVal lpFunc As Long) As Long
    ' lpObj - Pointer to object whose v-table will be modified
    ' nEntry - Index of v-table entry to be modified
    ' lpFunc - Function pointer of new v-table method
                            
    Dim lpFuncOld&, lpVTableHead&, lpFuncTmp&, nOldProtect&
    
    ' Object pointer contains a pointer to v-table--copy it to temporary
    CopyMemory lpVTableHead, ByVal lpObj, 4&          ' lpVTableHead = *lpObj;
    
    ' Calculate pointer to specified entry
    lpFuncTmp = lpVTableHead + (nEntry - 1) * 4
    
    ' Save address of previous method for return
    CopyMemory lpFuncOld, ByVal lpFuncTmp, 4&       ' lpFuncOld = *lpFuncTmp;
    
    ' Ignore if they're already the same
    If lpFuncOld <> lpFunc Then
        ' Need to change page protection to write to code
        VirtualProtect lpFuncTmp, 4&, PAGE_EXECUTE_READWRITE, nOldProtect
        
        ' Write the new function address into the v-table
        CopyMemory ByVal lpFuncTmp, lpFunc, 4&     ' *lpFuncTmp = lpFunc;
        
        ' Restore the previous page protection
        VirtualProtect lpFuncTmp, 4&, nOldProtect, nOldProtect    'Optional
    End If
    
    ReplaceVtableEntry = lpFuncOld
    
End Function

Private Function MapErr(ByVal ErrNumber As Long) As Long
  If ErrNumber Then
    If (ErrNumber And &H80000000) Or (ErrNumber = 1) Then
      'Error HRESULT already set
      MapErr = ErrNumber
    Else
      'Map back to a basic error number
      MapErr = &H800A0000 Or ErrNumber
    End If
  End If
End Function

' >>>>>>>>>>>>>>>>>>>>  Functions used to deal with unresolved callback pointers  <<<<<<<<<<<<<<<<<<<<
' These functions should only be used by collection objects and their enumerator objects as a means of
' callback communication.  This allows us to callback into a collection object without creating a circular
' reference problem that would keep the collection object from unloading untill all spawned enumeration
' objects were distroyed.
'
Public Sub AddPointerToLookupList(ByVal lpObject&)
  ' pass this task on to the FindPointer routine
  Call FindPointer(lpObject, True)
  
End Sub

Public Sub RemovePointerFromLookupList(ByVal lpObject&)
  ' removes a pointer from the array.  this function is called from the terminate event of the collection class
  
  Dim nItem&, nUbound&
  
  ' first find the location of the pointer in the array
  nItem = IsCollectionPointerValid(lpObject)
  
  ' if we found it, remove it and shift the rest of the items....
  If nItem > (-1) Then
    nUbound = UBound(m_alpCollections)
    
    If nItem < nUbound Then
      ' grab all of the items below the moved one and shift them up
      CopyMemory ByVal VarPtr(m_alpCollections(nItem)), ByVal VarPtr(m_alpCollections(nItem + 1)), (nUbound - nItem) * 4&
    End If
    
    
    If nUbound Then
      ' if there are other items in the array, preserve them
      ReDim Preserve m_alpCollections(nUbound - 1) As Long
    Else
      ' if this is the last item in the array, just redim the array and make sure the value is zero
      ReDim m_alpCollections(0) As Long
      m_alpCollections(0) = 0
    End If
  End If
  
End Sub

Public Function IsCollectionPointerValid(ByVal lpObject&) As Long
  ' checks whether or not the object still exists in the array.  for all intents
  ' and purposes, if the object isn't in the array, it no longer exists.
  IsCollectionPointerValid = FindPointer(lpObject)
End Function

Private Function FindPointer(ByVal lpObject&, Optional ByVal bAddIfNotFound As Boolean) As Long
  ' function to provide fast lookups for pointers in the array
  
  Static bInitialized As Boolean
  
  Dim i&, nLow&, nHigh&, nUbound&
  
  ' make sure the array is initialized
  If bInitialized = False Then
    If bAddIfNotFound Then
      GoTo AddFirsItem
    Else
      FindPointer = (-1)
    End If
  End If
  
    
  nHigh = UBound(m_alpCollections)
  
  ' loop through the array looking for the object pointer.
  ' the array is in numerical order so we can do fast lookups
  Do
    ' divide and conquer!  Each time we loop, devide the difference between the
    ' last items checked and search between the two indexes.  This is MUCH faster
    ' than looping through the entire list when dealing with a sorted array.
    i = nLow + ((nHigh - nLow) / 2)
    
    ' see how sKey relates to the current index....
    Select Case m_alpCollections(i)
      Case Is = lpObject
        FindPointer = i
        Exit Do
        
      Case Is > lpObject: nHigh = i - 1
      Case Is < lpObject: nLow = i + 1
    End Select

    
    ' if the low search bound has become greater than the high search bound, the
    ' item does not exist in the array.  if the bAddIfNotFound flag is set, a new
    ' item is being added.  otherwise, just return the not found value.
    If nLow > nHigh Then
      If bAddIfNotFound Then
      
AddFirsItem:

        ' check to see whether or not this item is initialized
        If Not bInitialized Then
          bInitialized = True
          ReDim m_alpCollections(0) As Long
        Else
          
          If m_alpCollections(0) <> 0 Then
            ReDim Preserve m_alpCollections(UBound(m_alpCollections) + 1) As Long
          End If
        
          nUbound = UBound(m_alpCollections)
          
          ' see whether we should add this item above or below the item at index 'i'
          Select Case m_alpCollections(i)
            Case Is < lpObject: i = i + 1
            Case Is > lpObject: i = i '<- included for self documentation
          End Select
          
          If i > nUbound Then i = nUbound
          
          If i < nUbound Then
            ' grab all of the items above the moved one and shift them down
            CopyMemory ByVal VarPtr(m_alpCollections(i + 1)), ByVal VarPtr(m_alpCollections(i)), (nUbound - i) * 4&
          End If ' i < nUbound
        End If
        
        ' place the new pointer into the position that used to be held by the target
        m_alpCollections(i) = lpObject
        
      End If ' bAddIfNotFound
      
      
      ' return value of KEY_NOT_FOUND tells the caller no match was found
      FindPointer = (-1)
      Exit Do
    End If ' nLow > nHigh
  
  Loop

End Function

Public Function ResolveCollectionPointer(ByVal lpObj&) As CSuperCollection
  ' used by all of the enumerator and key objects to resolve
  ' unreferenced pointers back to the parent collection class
  
  Dim oSC As CSuperCollection
  
  CopyMemory oSC, lpObj, 4
  Set ResolveCollectionPointer = oSC
  CopyMemory oSC, 0&, 4

End Function













