Attribute VB_Name = "ITypeInfoHelpers"
'Created by JAAFAR
'Src: https://www.vbforums.com/showthread.php?846947-RESOLVED-Ideas-Wanted-ITypeInfo-like-Solution&p=5449985&viewfull=1#post5449985
'I have managed to make this work for 64 bit (and 32bit as well). For this, I had to define two structures (TYPEATTR and FUNCDESC) from scratch so I can obtain from them the correct cFuncs and INVOKEKIND members (Functions count & Functions type).... Obviously, I also had to define two new arrays (aTYPEATTR and aFUNCDESC) with their respective byte sizes for 32bit and 64bit.
'In case you or anybody want to test this VBA code , here is an [Excel Workbook Sample](https://app.box.com/s/p8kgpwhibzu9zea7sdsua51wbwmi6yz3)
'
'@remark
'This is a pretty good basis. GetObjectFunctions is currently returning a collection of strings, which is quite strange. A better return type would be either an object
'or an array of FUNCDESC 
'@remark
'I have slightly modified the function to help show intent

Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TTYPEDESC
    #If Win64 Then
        pTypeDesc As LongLong
    #Else
        pTypeDesc As Long
    #End If
    vt            As Integer
End Type

Private Type TPARAMDESC
    #If Win64 Then
        pPARAMDESCEX  As LongLong
    #Else
        pPARAMDESCEX  As Long
    #End If
    wParamFlags       As Integer
End Type

Private Type TELEMDESC
    tdesc  As TTYPEDESC
    pdesc  As TPARAMDESC
End Type

Type TYPEATTR
        aGUID As GUID
        LCID As Long
        dwReserved As Long
        memidConstructor As Long
        memidDestructor As Long
        #If Win64 Then
            lpstrSchema As LongLong
        #Else
            lpstrSchema As Long
        #End If
        cbSizeInstance As Integer
        typekind As Long
        cFuncs As Integer
        cVars As Integer
        cImplTypes As Integer
        cbSizeVft As Integer
        cbAlignment As Integer
        wTypeFlags As Integer
        wMajorVerNum As Integer
        wMinorVerNum As Integer
        tdescAlias As Long
        idldescType As Long
End Type


Type FUNCDESC
    memid As Long                  'The function member ID (DispId).
    #If Win64 Then 
        lprgscode As LongLong         'Pointer to status code
        lprgelemdescParam As LongLong 'Pointer to description of the element.
    #Else
        lprgscode As Long             'Pointer to status code
        lprgelemdescParam As Long     'Pointer to description of the element.
    #End If
    funckind As Long                 'virtual, static, or dispatch-only
    INVOKEKIND As Long               'VbMethod / VbGet / VbSet / VbLet
    CallConv As Long                 'typically will be stdecl
    cParams As Integer               'number of parameters
    cParamsOpt As Integer            'number of optional parameters
    oVft As Integer                  'For FUNC_VIRTUAL, specifies the offset in the VTBL.
    cScodes As Integer               'The number of possible return values.
    elemdescFunc As TELEMDESC        'The function return type
    wFuncFlags As Integer            'The function flags. See FUNCFLAGS.
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#End If



'[IUnknown](https://en.wikipedia.org/wiki/IUnknown)
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release () 
Private Enum EIUnknown
  QueryInterface
  AddRef
  Release
End Enum

'[IDispatch](https://en.wikipedia.org/wiki/IDispatch)  extends IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release () 
'3      HRESULT  GetTypeInfoCount(unsigned int * pctinfo)
'4      HRESULT  GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo ** ppTInfo)
'5      HRESULT  GetIDsOfNames(REFIID riid, OLECHAR ** rgszNames, unsigned int cNames, LCID lcid, DISPID * rgDispId)
'6      HRESULT  Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult, EXCEPINFO * pExcepInfo, unsigned int * puArgErr)
Private Enum EIDispatch
    QueryInterface
    AddRef
    Release
    GetTypeInfoCount
    GetTypeInfo
    GetIDsOfNames
    Invoke
End Enum

'[ITypeInfo](https://github.com/tpn/winsdk-10/blob/master/Include/10.0.16299.0/um/OAIdl.h#L2683) extends IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release () 
'3      HRESULT  GetTypeAttr([out] TYPEATTR **ppTypeAttr )
'4      HRESULT  GetTypeComp([out] ITypeComp **ppTComp )
'5      HRESULT  GetFuncDesc([in] UINT index, [out] FUNCDESC **ppFuncDesc) 
'6      HRESULT  GetVarDesc([in] UINT index, [out] VARDESC **ppVarDesc)
'7      HRESULT  GetNames([in] MEMBERID memid, [out] BSTR *rgBstrNames, [in] UINT cMaxNames, [out] UINT *pcNames)
'8      HRESULT  GetRefTypeOfImplType( [in] UINT index, [out] HREFTYPE *pRefType)
'9      HRESULT  GetImplTypeFlags( [in] UINT index, [out] INT *pImplTypeFlags)
'10     HRESULT  GetIDsOfNames( [in] LPOLESTR *rgszNames, [in] UINT cNames, [out] MEMBERID *pMemId)
'11     HRESULT  Invoke( [in] PVOID pvInstance, [in] MEMBERID memid, [in] WORD wFlags, [out][in] DISPPARAMS *pDispParams, [out] VARIANT *pVarResult, [out] EXCEPINFO *pExcepInfo, [out] UINT *puArgErr)
'12     HRESULT  GetDocumentation( [in] MEMBERID memid, [out] BSTR *pBstrName, [out] BSTR *pBstrDocString, [out] DWORD *pdwHelpContext, [out] BSTR *pBstrHelpFile)
'13     HRESULT  GetDllEntry( [in] MEMBERID memid, [in] INVOKEKIND invKind, [out] BSTR *pBstrDllName, [out] BSTR *pBstrName, [out] WORD *pwOrdinal)
'14     HRESULT  GetRefTypeInfo( [in] HREFTYPE hRefType, [out] ITypeInfo **ppTInfo)
'15     HRESULT  AddressOfMember( [in] MEMBERID memid, [in] INVOKEKIND invKind, [out] PVOID *ppv)
'16     HRESULT  CreateInstance( [in] IUnknown *pUnkOuter, [in] REFIID riid, [out] PVOID *ppvObj)
'17     HRESULT  GetMops( [in] MEMBERID memid, [out] BSTR *pBstrMops)
'18     HRESULT  GetContainingTypeLib( [out] ITypeLib **ppTLib, [out] UINT *pIndex)
'19     void     ReleaseTypeAttr( [in] TYPEATTR *pTypeAttr)
'20     void     ReleaseFuncDesc( [in] FUNCDESC *pFuncDesc)
'21     void     ReleaseVarDesc( [in] VARDESC *pVarDesc)
Private Enum EITypeInfo
  QueryInterface
  AddRef
  Release
  GetTypeAttr
  GetTypeComp
  GetFuncDesc
  GetVarDesc
  GetNames
  GetRefTypeOfImplType
  GetImplTypeFlags
  GetIDsOfNames
  Invoke
  GetDocumentation
  GetDllEntry
  GetRefTypeInfo
  AddressOfMember
  CreateInstance
  GetMops
  GetContainingTypeLib
  ReleaseTypeAttr
  ReleaseFuncDesc
  ReleaseVarDesc
End Enum




Function GetObjectFunctions(ByVal TheObject As Object, Optional ByVal FuncType As VbCallType) As Collection
    Dim tTYPEATTR As TYPEATTR
    Dim tFUNCDESC As FUNCDESC

    Dim aGUID(0 To 11) As Long, lFuncsCount As Long
    
    #If Win64 Then
        Const PTRSIZE = 8
        const NULL_PTR as LongLong = 0^ 
        Dim aTYPEATTR() As LongLong, aFUNCDESC() As LongLong, farPtr As LongLong
    #Else
        Const PTRSIZE = 4
        const NULL_PTR as Long = 0& 
        Dim aTYPEATTR() As Long, aFUNCDESC() As Long, farPtr As Long
    #End If
    
    Dim ITypeInfo As IUnknown
    Dim IDispatch As IUnknown
    Dim sName As String, oCol As Collection
    set oCol = new Collection
    
    Const CC_STDCALL As Long = 4

    aGUID(0) = &H20400: aGUID(2) = &HC0&: aGUID(3) = &H46000000
    CallFunction_COM ObjPtr(TheObject), EIUnknown.QueryInterface * PTRSIZE, vbLong, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IDispatch)
    If IDispatch Is Nothing Then MsgBox "error":   Exit Function

    CallFunction_COM ObjPtr(IDispatch), EIDispatch.GetTypeInfo * PTRSIZE, vbLong, CC_STDCALL, 0&, 0&, VarPtr(ITypeInfo)
    If ITypeInfo Is Nothing Then MsgBox "error": Exit Function
    
    CallFunction_COM ObjPtr(ITypeInfo), EITypeInfo.GetTypeAttr * PTRSIZE, vbLong, CC_STDCALL, VarPtr(farPtr)
    If farPtr = 0& Then MsgBox "error": Exit Function

    CopyMemory ByVal VarPtr(tTYPEATTR), ByVal farPtr, LenB(tTYPEATTR)
    ReDim aTYPEATTR(LenB(tTYPEATTR))
    CopyMemory ByVal VarPtr(aTYPEATTR(0)), tTYPEATTR, UBound(aTYPEATTR)
    CallFunction_COM ObjPtr(ITypeInfo), EITypeInfo.ReleaseTypeAttr * PTRSIZE, vbEmpty, CC_STDCALL, farPtr
    
    For lFuncsCount = 0 To tTYPEATTR.cFuncs - 1
        Call CallFunction_COM(ObjPtr(ITypeInfo), EITypeInfo.GetFuncDesc * PTRSIZE, vbLong, CC_STDCALL, lFuncsCount, VarPtr(farPtr))
        If farPtr = 0 Then MsgBox "error": Exit For
        CopyMemory ByVal VarPtr(tFUNCDESC), ByVal farPtr, LenB(tFUNCDESC)
        ReDim aFUNCDESC(LenB(tFUNCDESC))
        CopyMemory ByVal VarPtr(aFUNCDESC(0)), tFUNCDESC, UBound(aFUNCDESC)
        Call CallFunction_COM(ObjPtr(ITypeInfo), EITypeInfo.ReleaseFuncDesc * PTRSIZE, vbEmpty, CC_STDCALL, farPtr)
        Call CallFunction_COM(ObjPtr(ITypeInfo), EITypeInfo.GetDocumentation * PTRSIZE, vbLong, CC_STDCALL, aFUNCDESC(0), VarPtr(sName), NULL_PTR, NULL_PTR, NULL_PTR)

        With tFUNCDESC
            If FuncType Then
                If .INVOKEKIND = FuncType Then
                    'Debug.Print sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                    oCol.Add sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                End If
            Else
                'Debug.Print sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
                oCol.Add sName & vbTab & Switch(.INVOKEKIND = 1, "VbMethod", .INVOKEKIND = 2, "VbGet", .INVOKEKIND = 4, "VbLet", .INVOKEKIND = 8, "VbSet")
            End If
        End With
        sName = vbNullString
    Next
    
    Set GetObjectFunctions = oCol
End Function



#If Win64 Then
    Private Function CallFunction_COM(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As LongLong
#Else
    Private Function CallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As Long
#End If

    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function

    Dim pIndex As Long, pCount As Long
    Dim vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant

    vParams() = FunctionParameters()
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If

    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, _
    vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then
        CallFunction_COM = vRtn
    Else
        SetLastError pIndex
    End If

End Function