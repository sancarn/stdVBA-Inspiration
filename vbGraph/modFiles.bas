Attribute VB_Name = "modFiles"
Option Explicit

Public Function GetFile(ByVal pstrFilename As String) As Byte()
Dim bytData()   As Byte
Dim intFile     As Integer
Dim lngLen      As Long
On Error GoTo ErrHandler
    ValidateFilename pstrFilename
    intFile = FreeFile
    Open pstrFilename For Binary Access Read Lock Read Write As #intFile
    lngLen = LOF(intFile)
    ReDim bytData(lngLen - 1) As Byte
    Get #intFile, 1, bytData()
    Close #intFile
    GetFile = bytData
    Exit Function
ErrHandler:
    Close #intFile
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub SaveFile(ByVal pstrFilename As String, ByRef pbytData() As Byte)
Dim intFile         As Integer
On Error GoTo ErrHandler
    intFile = FreeFile
    If Dir(pstrFilename) <> vbNullString And Len(pstrFilename) > 0 Then
        Kill pstrFilename
    End If
    Open pstrFilename For Binary Access Write As #intFile
    Put #intFile, 1, pbytData()
    Close #intFile
    Exit Sub
ErrHandler:
    Close #intFile
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ValidateFilename(ByVal pstrFilename As String)
    If Dir(pstrFilename) = vbNullString Or Len(pstrFilename) = 0 Then
        Err.Raise vbObjectError, App.EXEName, "File not found."
    End If
End Sub
