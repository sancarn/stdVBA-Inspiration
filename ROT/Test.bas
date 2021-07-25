'https://www.vbforums.com/showthread.php?883229-vb6-how-to-getobject(**)-by-rot-List-Rot-GetRunningObjectTable

Option Explicit
 
Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As IUnknown) As Long
Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As IUnknown) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
 
Private m_lCookie As Long
 
Private Sub Form_Load()
    List1.AddItem "test"
    List1.AddItem Now
    m_lCookie = PutObject(Me, "MySpecialProject.Form1")
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    RevokeObject m_lCookie
End Sub
 
Public Function PutObject(oObj As Object, sPathName As String, Optional ByVal Flags As Long) As Long
    Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Long = 1
    Const IDX_REGISTER  As Long = 3
    Dim hResult         As Long
    Dim pROT            As IUnknown
    Dim pMoniker        As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
    If hResult < 0 Then
        Err.Raise hResult, "CreateFileMoniker"
    End If
    DispCallByVtbl pROT, IDX_REGISTER, ROTFLAGS_REGISTRATIONKEEPSALIVE Or Flags, ObjPtr(oObj), ObjPtr(pMoniker), VarPtr(PutObject)
End Function
 
Public Sub RevokeObject(ByVal lCookie As Long)
    Const IDX_REVOKE    As Long = 4
    Dim hResult         As Long
    Dim pROT            As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    DispCallByVtbl pROT, IDX_REVOKE, lCookie
End Sub
 
Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function