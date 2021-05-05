'Created by wqweto
'Src: https://www.vbforums.com/showthread.php?846947-RESOLVED-Ideas-Wanted-ITypeInfo-like-Solution&p=5494379&viewfull=1#post5494379
'Here is my cleanup effort and after I got rid of API UDT's the whole code is reduced to about 70 lines incl. proper error handling (no MsgBox'es).
'It turns out TLBINF32.DLL is *not* present on recent Windows Servers (probably starting since 2012) so COM introspection using TLI is not 
'safe for production and apparently has never been as this DLL is a 3-rd party one that MS shipped as a favor to VB/VBA developers but no more 
'as apparently there is no x64 version of it (and all server editions are x64 only since long ago).
'For this reason the InterfaceInfoFromObject() function below can be used.
'Any bugfixes I make in production will try to backport to the snippet in this post.
'
'@remark
'This function although simpler and improving on error handling loses all structures and thus intent as well as extra useful information, thus is not easy to decipher or extend.



Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long

Public Function InterfaceInfoFromObject(ByVal oObj As Object, Optional ByVal InvokeKind As VbCallType) As Collection
    Const IDX_GetTypeInfo As Long = 4
    Const IDX_GetTypeAttr As Long = 3
    Const IDX_GetFuncDesc As Long = 5
    Const IDX_GetDocumentation As Long = 12
    Const IDX_ReleaseTypeAttr As Long = 19
    Const IDX_ReleaseFuncDesc As Long = 20
    #if Win64 then
      Const PTRSIZE as Long = 8
      Const NULL_PTR as LongLong = 0
    #else
      Const PTRSIZE as Long = 4
      Const NULL_PTR as Long = 0
    #end if
    Dim oCol            As Collection
    Dim pDispatch       As IUnknown
    Dim pTypeInfo       As IUnknown
    Dim lPtr            As Long
    Dim aTypeAttr(0 To 16) As Long
    Dim aFuncDesc(0 To 12) As Long
    Dim lIdx            As Long
    Dim sName           As String

    Set oCol = New Collection
    Call CopyMemory(pDispatch, oObj, PTRSIZE)
    Call CopyMemory(oObj, NULL_PTR, PTRSIZE)
    DispCallByVtbl pDispatch, IDX_GetTypeInfo, NULL_PTR, NULL_PTR, VarPtr(pTypeInfo)
    If pTypeInfo Is Nothing Then
        GoTo QH
    End If
    DispCallByVtbl pTypeInfo, IDX_GetTypeAttr, VarPtr(lPtr)
    If lPtr = 0 Then
        GoTo QH
    End If
    CopyMemory aTypeAttr(0), ByVal lPtr, (UBound(aTypeAttr) + 1) * 4
    DispCallByVtbl pTypeInfo, IDX_ReleaseTypeAttr, lPtr
    For lIdx = 0 To aTypeAttr(11) - 1 '--- [11] = TYPEATTR.cFuncs
        lPtr = 0
        DispCallByVtbl pTypeInfo, IDX_GetFuncDesc, lIdx, VarPtr(lPtr)
        If lPtr <> 0 Then
            CopyMemory aFuncDesc(0), ByVal lPtr, (UBound(aFuncDesc) + 1) * 4
            DispCallByVtbl pTypeInfo, IDX_ReleaseFuncDesc, lPtr
            sName = vbNullString
            DispCallByVtbl pTypeInfo, IDX_GetDocumentation, aFuncDesc(0), VarPtr(sName), NULL_PTR, NULL_PTR, NULL_PTR
            If LenB(sName) <> 0 And aFuncDesc(4) = InvokeKind Or InvokeKind = 0 Then '--- [4] = FUNCDESC.invkind
                oCol.Add Array(sName, aFuncDesc(4))
            End If
        End If
    Next
QH:
    Set InterfaceInfoFromObject = oCol
End Function

Public Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
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





'Currently not integrated but created by Eduardo for integration with InterfaceInfoFromObject
Private Function GetTypeName(ByVal nVarTypeInfo As VarTypeInfo) as string
    Dim iVarType As Long
    
    iVarType = nVarTypeInfo.VarType
    If (iVarType And Not VT_ARRAY) <> 0 Then
        Select Case (iVarType And &HFF&)
            Case VT_BOOL
                GetTypeName = "Boolean"
            Case VT_BSTR, VT_LPSTR, VT_LPWSTR
                GetTypeName = "String"
            Case VT_DATE
                GetTypeName = "Date"
            Case VT_INT
                GetTypeName = "Integer"
            Case VT_VARIANT
                GetTypeName = "Variant"
            Case VT_DECIMAL
                GetTypeName = "Decimal"
            Case VT_I4
                GetTypeName = "Long"
            Case VT_I2
                GetTypeName = "Integer"
            Case VT_I8
                GetTypeName = "Unknown"
            Case VT_SAFEARRAY
                GetTypeName = "SafeArray"
            Case VT_CLSID
                GetTypeName = "CLSID"
            Case VT_UINT
                GetTypeName = "UInt"
            Case VT_UI4
'                GetTypeName = "ULong"
                GetTypeName = "Long"
            Case VT_UNKNOWN
                GetTypeName = "Unknown"
            Case VT_VECTOR
                GetTypeName = "Vector"
            Case VT_R4
                GetTypeName = "Single"
            Case VT_R8
                GetTypeName = "Double"
            Case VT_DISPATCH
                GetTypeName = "Object"
            Case VT_UI1
                GetTypeName = "Byte"
            Case VT_CY
                GetTypeName = "Currency"
            Case VT_HRESULT
                GetTypeName = "HRESULT" ' note if this was a function it should be a sub
            Case VT_VOID
                GetTypeName = "Any"
            Case VT_ERROR
                GetTypeName = "Long"
            Case Else
                GetTypeName = "<Unsupported Variant Type"
                Select Case (iVarType And &HFF&)
                    Case VT_UI1
                        GetTypeName = GetTypeName & "(VT_UI1)"
                    Case VT_UI2
                        GetTypeName = GetTypeName & "(VT_UI2)"
                    Case VT_UI4
                        GetTypeName = GetTypeName & "(VT_UI4)"
                    Case VT_UI8
                        GetTypeName = GetTypeName & "(VT_UI8)"
                    Case VT_USERDEFINED
                        GetTypeName = GetTypeName & "(VT_USERDEFINED)"
                End Select
                GetTypeName = GetTypeName & ">"
        End Select
    Else
        GetTypeName = nVarTypeInfo.TypeInfo.Name
        If Left(GetTypeName, 1) = "_" Then
            GetTypeName = Mid$(GetTypeName, 2)
        End If
    End If
    If (iVarType And VT_ARRAY) = VT_ARRAY Then
        GetTypeName = GetTypeName & "()"
    End If
End Function