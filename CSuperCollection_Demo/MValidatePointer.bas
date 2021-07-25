Attribute VB_Name = "MValidatePointer"
  Option Explicit
  ' >>>>>>>>>>>>>>>>>>>>  Functions used to deal with unresolved callback pointers  <<<<<<<<<<<<<<<<<<<<
  ' These functions should only be used by collection objects and their enumerator objects as a means of
  ' callback communication.  This allows us to callback into a collection object without creating a circular
  ' reference problem that would keep the collection object from unloading untill all spawned enumeration
  ' objects were distroyed.

  ' array to hold the pointers we are in charge of validating
  Private m_alpCollections() As Long

Public Sub AddToLookupList(ByVal lpObject&)
  ' pass this task on to the FindPointer routine
  Call FindPointer(lpObject, True)
End Sub

Public Sub RemoveFromLookupList(ByVal lpObject&)
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

Public Function IsPointerValid(ByVal lpObject&) As Long
  ' checks whether or not the object still exists in the array.  for all intents
  ' and purposes, if the object isn't in the array, it no longer exists.
  IsPointerValid = FindPointer(lpObject)
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

