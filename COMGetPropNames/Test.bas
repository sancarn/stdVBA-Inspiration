Attribute VB_Name = "Test"
'Example:
' List all Methods and Properties of the excel application Object.
Sub Test()

    Dim oFuncCol As New Collection, i As Long, oObject As Object, sObjName As String

    
    Set oObject = Application '<=== Choose here target object as required.
    Set oFuncCol = GetObjectFunctions(TheObject:=oObject, FuncType:=0)
    
    Cells.CurrentRegion.Offset(1).ClearContents
    For i = 1 To oFuncCol.Count
        Range("A" & i + 1) = Split(oFuncCol.Item(i), vbTab)(0): Range("B" & i + 1) = Split(oFuncCol.Item(i), vbTab)(1)
    Next
    Range("C2") = oFuncCol.Count
    Cells(1).Resize(, 2).EntireColumn.AutoFit
    
    On Error Resume Next
        sObjName = oObject.Name
        If Len(sObjName) Then
            MsgBox "(" & oFuncCol.Count & ")  functions found for:" & vbCrLf & vbCrLf & sObjName
        End If
    On Error GoTo 0
    
End Sub