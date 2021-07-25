Src: https://www.vbforums.com/showthread.php?796759-VB6-Module-for-working-with-COM-Dll-without-registration
Author: The trick

Hello. I give my module for working with COM-DLL without registration in the registry.
The module has several functions:
GetAllCoclasses - returns to the list of classes and unique identifiers are extracted from a type library.
CreateIDispatch - creates IDispatch implementation by reference to the object and the name of the interface.
CreateObjectEx2 - creates an object by name from a type library.
CreateObjectEx - creates an object by CLSID.
UnloadLibrary - unloads the DLL if it is not used.

Other source: https://github.com/M2000Interpreter/Version9/blob/master/modTrickUnregCOM.bas



```vb
Attribute VB_Name = "modUnregCOM"
' The module modTrickUnregCOM.bas - for working with COM libraries without registration.
' � Krivous Anatolii Anatolevich (The trick), 2015

Option Explicit

Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByVal Destination As Long, ByVal Source As Long) As Long

Private Declare Function ObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long
Private Declare Function ObjSetNoRef Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Type GUID
    data1       As Long
    data2       As Integer
    data3       As Integer
    data4(7)    As Byte
End Type
Private Type ITypeInfo
    iunk                        As IUnknown
    GetTypeAttr                 As Long   ' &h0C
    GetTypeComp                 As Long   ' &h10
    GetFuncDesc                 As Long   ' &h14
    GetVarDesc                  As Long   ' &h18
    GetNames                    As Long   ' &h1C
    GetRefTypeOfImplType        As Long   ' &h20
    GetImplTypeFlags            As Long   '24
    GetIDsOfNames               As Long   '28
    Invoke                      As Long   '2C
    GetDocumentation            As Long   ' &H30
    GetDllEntry                 As Long    '34
    GetRefTypeInfo              As Long    '38
    AddressOfMember             As Long    '3C
    CreateInstance              As Long    '40
    GetMops                     As Long    '44
    GetContainingTypeLib        As Long    '48
    ReleaseTypeAttr             As Long    '4C
    ReleaseFuncDesc             As Long    '50
    ReleaseVarDesc              As Long    '54
End Type
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" ( _
                         ByRef Clsid As GUID, _
                         lpszProgID As Long) As Long
Private Declare Function StringFromCLSID Lib "ole32.dll" ( _
                         ByRef Clsid As GUID, _
                         lpszProgID As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
                            (ByVal hMem As Long)
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal lpszCLSID As Long, _
                         ByRef Clsid As GUID) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function SysFreeString Lib "oleaut32" ( _
                         ByVal lpbstr As Long) As Long
Private Declare Function LoadLibrary Lib "KERNEL32" _
                         Alias "LoadLibraryW" ( _
                         ByVal lpLibFileName As Long) As Long
Private Declare Function GetModuleHandle Lib "KERNEL32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function FreeLibrary Lib "KERNEL32" ( _
                         ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function DispCallFunc Lib "oleaut32" ( _
                         ByVal pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Integer, _
                         ByVal vtReturn As Integer, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
Private Declare Function LoadTypeLibEx Lib "oleaut32" ( _
                         ByVal szFile As Long, _
                         ByVal regkind As Long, _
                         ByRef pptlib As IUnknown) As Long
Private Declare Function memcpy Lib "KERNEL32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function CreateStdDispatch Lib "oleaut32" ( _
                         ByVal punkOuter As IUnknown, _
                         ByVal pvThis As IUnknown, _
                         ByVal ptinfo As IUnknown, _
                         ByRef ppunkStdDisp As IUnknown) As Long
                         
Private Const IID_IClassFactory   As String = "{00000001-0000-0000-C000-000000000046}"
Private Const IID_IUnknown        As String = "{00000000-0000-0000-C000-000000000046}"
Private Const CC_STDCALL          As Long = 4
Private Const REGKIND_NONE        As Long = 2
Private Const TKIND_COCLASS       As Long = 5
Private Const TKIND_DISPATCH      As Long = 4
Private Const TKIND_INTERFACE     As Long = 3

Dim iidClsFctr      As GUID
Dim iidUnk          As GUID
Dim isinit          As Boolean
'' ADDED BY GEORGE
' parameter description by [rm] 2005
Public Type TPARAMDESC
    pPARAMDESCEX            As Long     ' valid if PARAMFLAG_FHASDEFAULT
    
    wParamFlags             As Integer  ' parameter flags (in,out,...)
   dummy1 As Integer
End Type

' extended parameter description
Public Type TPARAMDESCEX
    cBytes                  As Long     ' size of structure
    varDefaultValue         As Variant  ' default value of parameter
End Type
Public Type TTYPEDESC
    pTypeDesc               As Long     ' vt = VT_PTR: points to another TYPEDESC
                                        ' vt = VT_CARRAY: points to another TYPEDESC
                                        ' vt = VT_USERDEFINED: pTypeDesc is a HREFTYPE instead of a pointer
    vt                      As Integer  ' vartype
    dummy1 As Integer
End Type
Public Type TELEMDESC
    tdesc                   As TTYPEDESC    ' type description
    pdesc                   As TPARAMDESC   ' parameter description
End Type

Public Type TELEMDESC1
    pTypeDesc               As Long
    vt                      As Integer
    pPARAMDESCEX            As Long
    wParamFlags             As Integer
End Type


Public Type TYPEATTR
    GUID(15)                As Byte
    tLCID                   As Long
    dwReserved              As Long
    memidConstructor        As Long
    memidDestructor         As Long
    pstrSchema              As Long
    cbSizeInstance          As Long
    typekind                As Long
    cFuncs                  As Integer
    CVars                   As Integer
    cImplTypes              As Integer
    cbSizeVft               As Integer
    cbAlignment             As Integer
    wTypeFlags              As Integer
    wMajorVerNum            As Integer
    wMinorVerNum            As Integer
    tdescAlias              As Long
    idldescType             As Long
End Type

Public Type FUNCDESC
    memid                   As Long
    lprgscode               As Long
    lprgelemdescParam       As Long
    funcking                As Long
    invkind                 As Long
    callconv                As Long
    cParams                 As Integer
    cParamsOpt              As Integer
    oVft                    As Integer
    cScodes                 As Integer
    elemdesc                As TELEMDESC1 ' Contains the return type of the function
    wFuncFlags              As Integer  ' function flags
End Type

' array description
Private Type TARRAYDESC
    tdescElem               As TTYPEDESC    ' type description
    cDims                   As Integer      ' number of dimensions
End Type
Private Type SAFEARRAYBOUND
    cElements               As Long
    lLBound                 As Long
End Type

Public Enum VARKIND
    VAR_PERSISTANCE = 0             '
    VAR_STATIC                      '
    VAR_CONST                       '
    VAR_DISPATCH                    '
End Enum

Private Type VARDESC
    memid                   As Long     ' member ID
    lpstrSchema             As Long     '
    uInstVal                As Long     ' vkind = VAR_PERINSTANCE: offset of this variable within the instance
                                        ' vkind = VAR_CONST: value of it as a variant
    elemdescVar             As TELEMDESC ' variable type
    wVarFlags               As Integer  ' variable flags
    vkind                   As Long     ' variable kind
End Type

' parameter flags
Public Enum PARAMFLAGS
    PARAMFLAG_NONE = &H0            ' ...
    PARAMFLAG_FIN = &H1             ' in
    PARAMFLAG_FOUT = &H2            ' out
    PARAMFLAG_FLCID = &H4           ' lcid
    PARAMFLAG_FRETVAL = &H8         ' return value
    PARAMFLAG_FOPT = &H10           ' optional
    PARAMFLAG_FHASDEFAULT = &H20    ' default value
    PARAMFLAG_FHASCUSTDATA = &H40   ' custom data
End Enum

Public Type fncinf
    Name                    As String
    addr                    As Long
    params                  As Integer
End Type

Public Type enmeinf
    Name                    As String
    invkind                 As invokekind
    params                  As Integer
    
End Type
Public Enum invokekind
    INVOKE_FUNC = &H1
    INVOKE_PROPERTY_GET = &H2
    INVOKE_PROPERTY_PUT = &H4
    INVOKE_PROPERTY_PUTREF = &H8
End Enum
Public Const DISP_E_PARAMNOTFOUND = &H80020004
Public Enum Varenum
    VT_EMPTY = 0&                   '
    VT_NULL = 1&                    ' 0
    VT_I2 = 2&                      ' signed 2 bytes integer
    VT_I4 = 3&                      ' signed 4 bytes integer
    VT_R4 = 4&                      ' 4 bytes float
    VT_R8 = 5&                      ' 8 bytes float
    VT_CY = 6&                      ' currency
    VT_DATE = 7&                    ' date
    VT_BSTR = 8&                    ' BStr
    VT_DISPATCH = 9&                ' IDispatch
    VT_ERROR = 10&                  ' error value
    VT_BOOL = 11&                   ' boolean
    VT_VARIANT = 12&                ' variant
    VT_UNKNOWN = 13&                ' IUnknown
    VT_DECIMAL = 14&                ' decimal
    VT_I1 = 16&                     ' signed byte
    VT_UI1 = 17&                    ' unsigned byte
    VT_UI2 = 18&                    ' unsigned 2 bytes integer
    VT_UI4 = 19&                    ' unsigned 4 bytes integer
    VT_I8 = 20&                     ' signed 8 bytes integer
    VT_UI8 = 21&                    ' unsigned 8 bytes integer
    VT_INT = 22&                    ' integer
    VT_UINT = 23&                   ' unsigned integer
    VT_VOID = 24&                   ' 0
    VT_HRESULT = 25&                ' HRESULT
    VT_PTR = 26&                    ' pointer
    VT_SAFEARRAY = 27&              ' safearray
    VT_CARRAY = 28&                 ' carray
    VT_USERDEFINED = 29&            ' userdefined
    VT_LPSTR = 30&                  ' LPStr
    VT_LPWSTR = 31&                 ' LPWStr
    VT_RECORD = 36&                 ' Record
    VT_FILETIME = 64&               ' File Time
    VT_BLOB = 65&                   ' Blob
    VT_STREAM = 66&                 ' Stream
    VT_STORAGE = 67&                ' Storage
    VT_STREAMED_OBJECT = 68&        ' Streamed Obj
    VT_STORED_OBJECT = 69&          ' Stored Obj
    VT_BLOB_OBJECT = 70&            ' Blob Obj
    VT_CF = 71&                     ' CF
    VT_CLSID = 72&                  ' Class ID
    VT_BSTR_BLOB = &HFFF&           ' BStr Blob
    VT_VECTOR = &H1000&             ' Vector
    VT_ARRAY = &H2000&              ' Array
    VT_BYREF = &H4000&              ' ByRef
    VT_RESERVED = &H8000&           ' Reserved
    VT_ILLEGAL = &HFFFF&            ' illegal
End Enum



Public Function GetGUIDstr(g As GUID) As String
Dim ret As Long, here As Long

ret = StringFromCLSID(g, here)
If ret Then Exit Function
GetGUIDstr = GetBStrFromPtr(here)
CoTaskMemFree here
End Function
Public Function strProgID(g As GUID) As String
Dim ret As Long, here As Long

ret = ProgIDFromCLSID(g, here)
If ret Then Exit Function
strProgID = GetBStrFromPtr(here)
CoTaskMemFree here
End Function

Public Function strProgIDfromSrting(IID_IClassFactory As String) As String
Dim ret As Long, here As Long, iidClsFctr As GUID
CLSIDFromString StrPtr(IID_IClassFactory), iidClsFctr
ret = ProgIDFromCLSID(iidClsFctr, here)
If ret Then Exit Function
strProgIDfromSrting = GetBStrFromPtr(here)
CoTaskMemFree here
End Function

''

' // Get all co-classes described in type library.
Public Function GetAllCoclasses( _
                ByRef Path As String, _
                ByRef listOfClsid() As GUID, _
                ByRef listOfNames() As String, _
                ByRef countCoClass As Long) As Boolean
                
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim count   As Long
    Dim index   As Long
    Dim pAttr   As Long
    Dim tKind   As Long
    
    ret = LoadTypeLibEx(StrPtr(Path), REGKIND_NONE, typeLib)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    count = ITypeLib_GetTypeInfoCount(typeLib)
    countCoClass = 0
    
    If count > 0 Then
    
        ReDim listOfClsid(count - 1)
        ReDim listOfNames(count - 1)
        
        For index = 0 To count - 1
        
            ret = ITypeLib_GetTypeInfo(typeLib, index, typeInf)
                        
            If ret Then
                Err.Raise ret
                Exit Function
            End If
            
            ITypeInfo_GetTypeAttr typeInf, pAttr
            
            GetMem4 ByVal pAttr + &H28, tKind
            
            If tKind = TKIND_COCLASS Then
            
                memcpy listOfClsid(countCoClass), ByVal pAttr, Len(listOfClsid(countCoClass))
                ret = ITypeInfo_GetDocumentation(typeInf, -1, listOfNames(countCoClass), vbNullString, 0, vbNullString)
                
                If ret Then
                    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
                    Err.Raise ret
                    Exit Function
                End If
                
                countCoClass = countCoClass + 1
                
            End If
            
            ITypeInfo_ReleaseTypeAttr typeInf, pAttr
            
            Set typeInf = Nothing
            
        Next
        
    End If
    
    If countCoClass Then
        
        ReDim Preserve listOfClsid(countCoClass - 1)
        ReDim Preserve listOfNames(countCoClass - 1)
    
    Else
    
        Erase listOfClsid()
        Erase listOfNames()
        
    End If
    
    GetAllCoclasses = True
    
End Function
Private Function ResolveObjPtr(ByVal Ptr As Long) As IUnknown
ObjSetAddRef ResolveObjPtr, Ptr
End Function
Public Function ResolveObjPtrNoRef(ByVal Ptr As Long) As IUnknown
ObjSetNoRef ResolveObjPtrNoRef, Ptr
End Function

Public Function GetAllMembers(mList As FastCollection, obj As Object _
               ) As Boolean
        Dim IDsp        As IDispatch.IDispatchM2000
        Dim riid        As IDispatch.IID
        Dim params      As IDispatch.DISPPARAMS
        Dim Excep       As IDispatch.EXCEPINFO
        Dim mAttr As TYPEATTR
        Set mList = New FastCollection
        Dim ppFuncDesc As Long, fncdsc As FUNCDESC, cFuncs As Long
        Dim ppVarDesc As Long, vardsc As VARDESC
        Dim ParamDesc As TPARAMDESC, hlp As Long, pRefType As Long
        Dim TypeDesc As TTYPEDESC, retval$
        Dim ret As Long, pctinfo As Long, ppTInfo As Long, typeInf As IUnknown
        Dim pAttr   As Long
        Dim tKind   As Long
        Set IDsp = obj
        Dim cFncs As Long, CVars As Long, ttt$
        Dim i As Long
        Dim j As Long
        Dim strNames() As String, strName As String, aName As String
        
        Dim acc As Long
        
        Const TYPEFLAG_FDUAL = &H40
        Const TYPEFLAG_FPREDECLID = &H8
        '' may have a GET and a LET for same name
        mList.AllowAnyKey
        
        
        ret = IDsp.GetTypeInfo(ByVal 0, ByVal 0, ppTInfo)
        If ppTInfo = 0 Or ret <> 0 Then
        If Err Then Err.clear
        Exit Function
        
        End If
        Set typeInf = ResolveObjPtrNoRef(ppTInfo)
        ITypeInfo_GetTypeAttr typeInf, pAttr
        If pAttr = 0 Then Set typeInf = Nothing: Exit Function
        memcpy mAttr, ByVal pAttr, Len(mAttr)

         If (mAttr.wTypeFlags And TYPEFLAG_FPREDECLID) = &H8 Then
            ITypeInfo_ReleaseTypeAttr typeInf, pAttr
            Set typeInf = Nothing
            Exit Function
         End If
        If (mAttr.wTypeFlags And TYPEFLAG_FDUAL) Then
            If mAttr.typekind <> TKIND_DISPATCH Then

                ITypeInfo_GetRefTypeOfImplType typeInf, -1, pRefType
                ITypeInfo_ReleaseTypeAttr typeInf, pAttr
                ITypeInfo_GetRefTypeInfo typeInf, pRefType, ppTInfo
                Set typeInf = ResolveObjPtrNoRef(ppTInfo)
                ITypeInfo_GetTypeAttr typeInf, pAttr
                memcpy mAttr, ByVal pAttr, Len(mAttr)
            End If
        End If
   

        If TKIND_DISPATCH = mAttr.typekind Then
        cFuncs = mAttr.cFuncs '' mAttr.cVars
        If cFuncs = 0 And False Then   ' not finished yet
        If mAttr.CVars > 0 Then
        ' HAS ENUM, STRUCT, ETC.
        'For j = 0 To mAttr.CVars
         '   ITypeInfo_GetVarDesc typeInf, j, ppVarDesc
          '  CpyMem vardsc, ByVal ppFuncDesc, Len(vardsc)
           ' ITypeInfo_ReleaseVarDesc typeInf, ppVarDesc
            '''error ReDim strNames(vardsc.cParams + 1) As String
           ' ret = ITypeInfo_GetNames(typeInf, vardsc.memid, strNames(), 1 + vardsc.CVars, CVars)
          'next
       End If
        End If
        
        
        For j = 0 To mAttr.cFuncs - 1
            ITypeInfo_GetFuncDesc typeInf, j, ppFuncDesc
            CpyMem fncdsc, ByVal ppFuncDesc, Len(fncdsc)
            acc = fncdsc.lprgelemdescParam
            ret = ITypeInfo_GetDocumentation(typeInf, fncdsc.memid, strName, vbNullString, 0, vbNullString)
            If True Then
            
            mList.AddKey UCase(strName), ""
            Select Case fncdsc.invkind
            Case INVOKE_FUNC:
                If fncdsc.elemdesc.vt = 24 Then
                strName = "Sub " + strName
                Else
                strName = "Function " + strName
                End If
            Case INVOKE_PROPERTY_GET:
                strName = "Property Get " + strName
            Case INVOKE_PROPERTY_PUT:
                strName = "Property Let " + strName
            Case INVOKE_PROPERTY_PUTREF:
                strName = "Property Set " + strName
            End Select
            mList.ToEnd            ' move to last

          ProcTask2 basestack1
          hlp = fncdsc.cParams
            If hlp > 0 Then
                 cFncs = 0
                ReDim strNames(fncdsc.cParams + 1) As String
                aName = vbNullString
                ret = ITypeInfo_GetDocumentation(typeInf, fncdsc.memid, aName, vbNullString, 0, vbNullString)
                ret = ITypeInfo_GetNames(typeInf, fncdsc.memid, strNames(), fncdsc.cParams + 1, cFncs)
 
                
                       If Not ret Then
                          
                            strName = strName + "("
                            For i = 1 To hlp
                            If IsBadCodePtr(acc) = 0 Then
                                CopyBytes Len(ParamDesc), VarPtr(ParamDesc), ByVal acc + 8
                                CopyBytes Len(TypeDesc), VarPtr(TypeDesc), ByVal acc
                                
                            End If
                                acc = acc + 16
                           ttt$ = ""
                           retval$ = ""
                           If strNames(i) = "" Then strNames(i) = "Value"
                           If (ParamDesc.wParamFlags And PARAMFLAG_FRETVAL) = &H8 Then
                               retval$ = " as " + stringifyTypeDesc(TypeDesc, typeInf)
                           Else
                                If (ParamDesc.wParamFlags And PARAMFLAG_FIN) = &H1 Then ttt$ = "in "
                                If (ParamDesc.wParamFlags And PARAMFLAG_FOUT) = &H2 Then ttt$ = ttt$ + "out "
                                If i > (hlp - fncdsc.cParamsOpt) And fncdsc.cParamsOpt <> 0 Then
                                    strName = strName + "[" + ttt$ + strNames(i) + " " + stringifyTypeDesc(TypeDesc, typeInf) + "]"
                                Else
                                    If fncdsc.cParamsOpt = 0 And (ParamDesc.wParamFlags And PARAMFLAG_FOPT) > 0 Then
                                        strName = strName + "[" + ttt$ + strNames(i) + " " + stringifyTypeDesc(TypeDesc, typeInf) + "]"
                                    Else
                                        strName = strName + ttt$ + strNames(i) + " " + stringifyTypeDesc(TypeDesc, typeInf)
                                    End If
                                End If
                                If i < hlp Then strName = strName + ", "
                            End If
                        Next i
                        strName = strName + ")"
                       End If
                   End If
                   
                If retval$ = "" Then
                     If fncdsc.elemdesc.vt = 24 Then
                     mList.Value = strName
                     Else
                        CopyBytes Len(TypeDesc), VarPtr(TypeDesc), VarPtr(fncdsc.elemdesc.pTypeDesc)
                     mList.Value = strName + " as " + stringifyTypeDesc(TypeDesc, typeInf)
                     
                     End If
                Else
                     mList.Value = strName + retval$
            End If
            
            End If
            ITypeInfo_ReleaseFuncDesc typeInf, ppFuncDesc
        Next j
        ReDim strNames(1) As String
    End If
   
 
 ITypeInfo_ReleaseTypeAttr typeInf, pAttr
 
   Set typeInf = Nothing
    Set IDsp = Nothing
GetAllMembers = True
End Function
Public Function GetAllMembersOld(mList As FastCollection, obj As Object _
               ) As Boolean
 Dim IDsp        As IDispatch.IDispatchM2000
 Dim riid        As IDispatch.IID
Dim params      As IDispatch.DISPPARAMS
Dim Excep       As IDispatch.EXCEPINFO
Dim mAttr As TYPEATTR
Set mList = New FastCollection
Dim ppFuncDesc As Long, fncdsc As FUNCDESC, cFuncs As Long
 'Dim aaa As IDispatch.EXCEPINFO
 Dim ret As Long, pctinfo As Long, ppTInfo As Long, typeInf As IUnknown
     Dim pAttr   As Long
    Dim tKind   As Long
 Set IDsp = obj
 Dim cFncs           As Long
    Dim i As Long


  ret = IDsp.GetTypeInfo(ByVal 0, ByVal 0, ppTInfo)
    Set typeInf = ResolveObjPtrNoRef(ppTInfo)
     ITypeInfo_GetTypeAttr typeInf, pAttr
    memcpy mAttr, ByVal pAttr, Len(mAttr)
   
       If TKIND_DISPATCH = mAttr.typekind Then
         cFuncs = mAttr.cFuncs '' mAttr.cVars

    Dim listOfNames() As String, strNames() As String, strName As String
    ReDim listOfNames(1) As String
    Dim j As Long
    For j = 0 To mAttr.cFuncs - 1
    ITypeInfo_GetFuncDesc typeInf, j, ppFuncDesc
    CpyMem fncdsc, ByVal ppFuncDesc, Len(fncdsc)
    ITypeInfo_ReleaseFuncDesc typeInf, ppFuncDesc
     ret = ITypeInfo_GetDocumentation(typeInf, fncdsc.memid, listOfNames(0), vbNullString, 0, vbNullString)
     strName = listOfNames(0)
     
           mList.AddKey UCase(strName), ""
        Select Case fncdsc.invkind
            Case INVOKE_FUNC:
                strName = "Function " + strName
            Case INVOKE_PROPERTY_GET:
                strName = "Property Get " + strName
            Case INVOKE_PROPERTY_PUT:
                strName = "Property Let " + strName
            Case INVOKE_PROPERTY_PUTREF:
                strName = "Property Set " + strName
        End Select
        mList.ToEnd  ' move to last
    
        If fncdsc.cParams > 0 Then
        cFncs = 0
        ReDim strNames(fncdsc.cParams + 1) As String
        ret = ITypeInfo_GetDocumentation(typeInf, fncdsc.memid, listOfNames(0), vbNullString, 0, vbNullString)
        ret = ITypeInfo_GetNames(typeInf, fncdsc.memid, strNames(), 1 + fncdsc.cParams, cFncs)
        If Not ret Then
         strName = strName + "("
            For i = 1 To fncdsc.cParams
            strName = strName + strNames(i)
            If i < fncdsc.cParams Then strName = strName + ", "
            Next i
        strName = strName + ")"
        End If
        End If
    
        mList.Value = strName
    
    
    Next j
    End If
   
 
 ITypeInfo_ReleaseTypeAttr typeInf, pAttr
 
   Set typeInf = Nothing
    Set IDsp = Nothing

End Function


' // Create IDispach implementation described in type library.
Public Function CreateIDispatch( _
                ByRef obj As IUnknown, _
                ByRef typeLibPath As String, _
                ByRef interfaceName As String) As Object
                
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim retObj  As IUnknown
    Dim pAttr   As Long
    Dim tKind   As Long
    
    ret = LoadTypeLibEx(StrPtr(typeLibPath), REGKIND_NONE, typeLib)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    ret = ITypeLib_FindName(typeLib, interfaceName, 0, typeInf, 0, 1)
    
    If typeInf Is Nothing Then
        Err.Raise &H80004002, , "Interface not found"
        Exit Function
    End If
    
    ITypeInfo_GetTypeAttr typeInf, pAttr
    GetMem4 ByVal pAttr + &H28, tKind
    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
    
    If tKind = TKIND_DISPATCH Then
        Set CreateIDispatch = obj
        Exit Function
    ElseIf tKind <> TKIND_INTERFACE Then
        Err.Raise &H80004002, , "Interface not found"
        Exit Function
    End If
  
    ret = CreateStdDispatch(Nothing, obj, typeInf, retObj)
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    Set CreateIDispatch = retObj

End Function

' // Create object by Name.
Public Function CreateObjectEx2( _
                ByRef pathToDll As String, _
                ByRef pathToTLB As String, _
                ByRef className As String) As IUnknown
                
    Dim typeLib As IUnknown
    Dim typeInf As IUnknown
    Dim ret     As Long
    Dim pAttr   As Long
    Dim tKind   As Long
    Dim Clsid   As GUID
    
    ret = LoadTypeLibEx(StrPtr(pathToTLB), REGKIND_NONE, typeLib)
    
    If ret Then
        Err.Raise ret
        Exit Function
    End If
    
    ret = ITypeLib_FindName(typeLib, className, 0, typeInf, 0, 1)
    
    If typeInf Is Nothing Then
        Err.Raise &H80040111, , "Class not found in type library"
        Exit Function
    End If

    ITypeInfo_GetTypeAttr typeInf, pAttr
    
    GetMem4 ByVal pAttr + &H28, tKind
    
    If tKind = TKIND_COCLASS Then
        memcpy Clsid, ByVal pAttr, Len(Clsid)
    Else
        Err.Raise &H80040111, , "Class not found in type library"
        Exit Function
    End If
    
    ITypeInfo_ReleaseTypeAttr typeInf, pAttr
            
    Set CreateObjectEx2 = CreateObjectEx(pathToDll, Clsid)
    
End Function
                
' // Create object by CLSID and path.
Public Function CreateObjectEx( _
                ByRef Path As String, _
                ByRef Clsid As GUID) As IUnknown
                
    Dim hLib    As Long
    Dim lpAddr  As Long
    Dim isLoad  As Boolean
    
    hLib = GetModuleHandle(StrPtr(Path))
    
    If hLib = 0 Then
    
        hLib = LoadLibrary(StrPtr(Path))
        If hLib = 0 Then
            Err.Raise 53, , error(53) & " " & Chr$(34) & Path & Chr$(34)
            Exit Function
        End If
        
        isLoad = True
        
    End If
    
    lpAddr = GetProcAddress(hLib, "DllGetClassObject")
    
    If lpAddr = 0 Then
        If isLoad Then FreeLibrary hLib
        Err.Raise 453, , "Can't find dll entry point DllGetClasesObject in " & Chr$(34) & Path & Chr$(34)
        Exit Function
    End If

    If Not isinit Then
        CLSIDFromString StrPtr(IID_IClassFactory), iidClsFctr
        CLSIDFromString StrPtr(IID_IUnknown), iidUnk
        isinit = True
    End If
    
    Dim ret     As Long
    Dim out     As IUnknown
    
    ret = DllGetClassObject(lpAddr, Clsid, iidClsFctr, out)
    
    If ret = 0 Then

        ret = IClassFactory_CreateInstance(out, 0, iidUnk, CreateObjectEx)
    
    Else
    
        If isLoad Then FreeLibrary hLib
        Err.Raise ret
        Exit Function
        
    End If
    
    Set out = Nothing
    
    If ret Then
    
        If isLoad Then FreeLibrary hLib
        Err.Raise ret

    End If
    
End Function

' // Unload DLL if not used.
Public Function UnloadLibrary( _
                ByRef Path As String) As Boolean
                
    Dim hLib    As Long
    Dim lpAddr  As Long
    Dim ret     As Long
    
    If Not isinit Then Exit Function
    
    hLib = GetModuleHandle(StrPtr(Path))
    If hLib = 0 Then Exit Function
    
    lpAddr = GetProcAddress(hLib, "DllCanUnloadNow")
    If lpAddr = 0 Then Exit Function
    
    ret = DllCanUnloadNow(lpAddr)
    
    If ret = 0 Then
        FreeLibrary hLib
        UnloadLibrary = True
    End If
    
End Function

' // Call "DllGetClassObject" function using a pointer.
Public Function DllGetClassObject( _
                 ByVal funcAddr As Long, _
                 ByRef Clsid As GUID, _
                 ByRef IID As GUID, _
                 ByRef out As IUnknown) As Long
                 
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim list(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = VarPtr(Clsid)
    params(1) = VarPtr(IID)
    params(2) = VarPtr(out)
    
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 3, types(0), list(0), pReturn)
             
    If resultCall Then Err.Raise 5: Exit Function
    
    DllGetClassObject = pReturn
    
End Function

' // Call "DllCanUnloadNow" function using a pointer.
Private Function DllCanUnloadNow( _
                 ByVal funcAddr As Long) As Long
                 
    Dim resultCall  As Long
    Dim pReturn     As Variant
    
    resultCall = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, pReturn)
             
    If resultCall Then Err.Raise 5: Exit Function
    
    DllCanUnloadNow = pReturn
    
End Function

' // Call "IClassFactory:CreateInstance" method.
Public Function IClassFactory_CreateInstance( _
                 ByVal obj As IUnknown, _
                 ByVal punkOuter As Long, _
                 ByRef riid As GUID, _
                 ByRef out As IUnknown) As Long
    
    Dim params(2)   As Variant
    Dim types(2)    As Integer
    Dim list(2)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = punkOuter
    params(1) = VarPtr(riid)
    params(2) = VarPtr(out)
    
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(obj, &HC, CC_STDCALL, vbLong, 3, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    IClassFactory_CreateInstance = pReturn
    
End Function

' // Call "ITypeLib:GetTypeInfoCount" method.
Private Function ITypeLib_GetTypeInfoCount( _
                 ByVal obj As IUnknown) As Long
    
    Dim resultCall  As Long
    Dim pReturn     As Variant

    resultCall = DispCallFunc(obj, &HC, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeLib_GetTypeInfoCount = pReturn
    
End Function

' // Call "ITypeLib:GetTypeInfo" method.
Public Function ITypeLib_GetTypeInfo( _
                 ByVal obj As IUnknown, _
                 ByVal index As Long, _
                 ByRef ppTInfo As IUnknown) As Long
    
    Dim params(1)   As Variant
    Dim types(1)    As Integer
    Dim list(1)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = index
    params(1) = VarPtr(ppTInfo)
    
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(obj, &H10, CC_STDCALL, vbLong, 2, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeLib_GetTypeInfo = pReturn
    
End Function

' // Call "ITypeLib:FindName" method.
Public Function ITypeLib_FindName( _
                 ByVal obj As IUnknown, _
                 ByRef szNameBuf As String, _
                 ByVal lHashVal As Long, _
                 ByRef ppTInfo As IUnknown, _
                 ByRef rgMemId As Long, _
                 ByRef pcFound As Integer) As Long
    
    Dim params(4)   As Variant
    Dim types(4)    As Integer
    Dim list(4)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = StrPtr(szNameBuf)
    params(1) = lHashVal
    params(2) = VarPtr(ppTInfo)
    params(3) = VarPtr(rgMemId)
    params(4) = VarPtr(pcFound)
    
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(obj, &H2C, CC_STDCALL, vbLong, 5, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeLib_FindName = pReturn
    
End Function
'' "ITypeInfo:GetTypeAttr"
Public Sub ITypeInfo_GetVarDesc( _
            ByVal obj As IUnknown, _
            ByVal index As Long, _
            ByRef ppVarAttr As Long)
    
    Dim resultCall  As Long
    Dim pReturn     As Variant
    Dim params(1)   As Variant
    Dim types(1)    As Integer
    Dim list(1)     As Long
    Dim pIndex      As Long
    params(0) = index
    params(1) = VarPtr(ppVarAttr)
   
       For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(obj, &H18, CC_STDCALL, vbLong, 2, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall
End Sub
Public Sub ITypeInfo_ReleaseVarDesc( _
            ByVal obj As IUnknown, _
            ByVal ppVarAttr As Long)
    
    Dim resultCall  As Long
    
    resultCall = DispCallFunc(obj, &H54, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(ppVarAttr)), 0)
          
    If resultCall Then Err.Raise resultCall

End Sub
' // Call "ITypeInfo:GetTypeAttr" method.
Public Sub ITypeInfo_GetTypeAttr( _
            ByVal obj As IUnknown, _
            ByRef ppTypeAttr As Long)
    
    Dim resultCall  As Long
    Dim pReturn     As Variant
    ppTypeAttr = 0
    pReturn = VarPtr(ppTypeAttr)
    resultCall = DispCallFunc(obj, &HC, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(pReturn), 0)
    If ppTypeAttr = 0 Then Exit Sub
    If resultCall Then Err.Raise resultCall

End Sub
Public Sub ITypeInfo_GetRefTypeOfImplType( _
            ByVal obj As IUnknown, _
            ByVal index As Long, _
            ByRef pRefType As Long)
    Dim resultCall  As Long
    Dim pReturn     As Variant
    Dim params(1)   As Variant
    Dim types(1)    As Integer
    Dim list(1)     As Long
        Dim pIndex      As Long
     params(0) = index
    params(1) = VarPtr(pRefType)
   
       For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(obj, &H14, CC_STDCALL, vbLong, 2, types(0), list(0), pReturn)
    If resultCall Then Err.Raise resultCall

        
End Sub

Public Sub ITypeInfo_GetFuncDesc( _
            ByVal obj As IUnknown, _
            ByVal index As Long, _
            ByRef ppFuncAttr As Long)
    
    Dim resultCall  As Long
    Dim pReturn     As Variant
    Dim params(1)   As Variant
    Dim types(1)    As Integer
    Dim list(1)     As Long
        Dim pIndex      As Long
     params(0) = index
    params(1) = VarPtr(ppFuncAttr)
   
       For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(obj, &H14, CC_STDCALL, vbLong, 2, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall

End Sub

' // Call "ITypeInfo:GetDocumentation" method.
Public Function ITypeInfo_GetDocumentation( _
                 ByVal obj As IUnknown, _
                 ByVal memid As Long, _
                 ByRef pBstrName As String, _
                 ByRef pBstrDocString As String, _
                 ByRef pdwHelpContext As Long, _
                 ByRef pBstrHelpFile As String) As Long
    
    Dim params(4)   As Variant
    Dim types(4)    As Integer
    Dim list(4)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = memid
    params(1) = VarPtr(pBstrName)
    params(2) = VarPtr(pBstrDocString)
    params(3) = VarPtr(pdwHelpContext)
    params(4) = VarPtr(pBstrHelpFile)
    
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    
    resultCall = DispCallFunc(obj, &H30, CC_STDCALL, vbLong, 5, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeInfo_GetDocumentation = pReturn
    
End Function
Public Function ITypeInfo_GetNames( _
                 ByVal obj As IUnknown, _
                 ByVal memid As Long, _
                 pBstrName() As String, _
                 ByVal cMaxNames As Long, _
                 ByRef pcNames As Long) As Long
    
    Dim params(3)   As Variant
    Dim types(3)    As Integer
    Dim list(3)     As Long
    Dim resultCall  As Long
    Dim pIndex      As Long
    Dim pReturn     As Variant
    
    params(0) = memid
    params(1) = VarPtr(pBstrName(0))
    params(2) = cMaxNames
    params(3) = VarPtr(pcNames)
    
    For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next

    resultCall = DispCallFunc(obj, &H1C, CC_STDCALL, vbLong, 4, types(0), list(0), pReturn)
          
    If resultCall Then Err.Raise resultCall: Exit Function
     
    ITypeInfo_GetNames = pReturn
    
End Function
Public Sub ITypeInfo_GetRefTypeInfo( _
            ByVal obj As IUnknown, _
            ByVal hreftype As Long, _
            ByRef ppTInfo As Long)
  
  ' &H38
     Dim resultCall  As Long
    Dim pReturn     As Variant
    Dim params(1)   As Variant
    Dim types(1)    As Integer
    Dim list(1)     As Long
        Dim pIndex      As Long
     params(0) = hreftype
    params(1) = VarPtr(ppTInfo)
   
       For pIndex = 0 To UBound(params)
        list(pIndex) = VarPtr(params(pIndex)):   types(pIndex) = VarType(params(pIndex))
    Next
    resultCall = DispCallFunc(obj, &H38, CC_STDCALL, vbLong, 2, types(0), list(0), pReturn)
    If resultCall Then Err.Raise resultCall
 
  
            End Sub
' // Call "ITypeInfo:ReleaseTypeAttr" method.
Public Sub ITypeInfo_ReleaseTypeAttr( _
            ByVal obj As IUnknown, _
            ByVal ppTypeAttr As Long)
    
    Dim resultCall  As Long
    
    resultCall = DispCallFunc(obj, &H4C, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(ppTypeAttr)), 0)
    If resultCall Then Err.Raise resultCall

End Sub
Public Sub ITypeInfo_ReleaseFuncDesc( _
            ByVal obj As IUnknown, _
            ByVal ppFuncAttr As Long)
    
    Dim resultCall  As Long
    
    resultCall = DispCallFunc(obj, &H50, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(ppFuncAttr)), 0)
    If resultCall Then Err.Raise resultCall

End Sub

Private Function VarTypeName(nVarType As Long) As String
    Select Case nVarType
        Case 0
            VarTypeName = "CustomType"
        Case 2
            VarTypeName = "Integer"
        Case 3, 10
            VarTypeName = "Long"
        Case 4
            VarTypeName = "Single"
        Case 5
            VarTypeName = "Double"
        Case 6
            VarTypeName = "Currency"
        Case 7
            VarTypeName = "Date"
        Case 8
            VarTypeName = "String"
        Case 9, 13
            VarTypeName = "Object"
        Case 11
            VarTypeName = "Boolean"
        Case 12, 1, 36
            VarTypeName = "Variant"
        Case 14
            VarTypeName = "Decimal"
        Case 17
            VarTypeName = "Byte"
        Case 8192
            VarTypeName = "Array"
        Case Else
        VarTypeName = "type" + Trim(Str$(nVarType))
            'Stop
    End Select
End Function
Private Function stringifyCustomType(ByVal hreftype As Long, pTypeInfo As IUnknown) As String

Dim ppTInfo As Long, pCustTypeInfo As IUnknown, bstrType As String, ret As Long

On Error Resume Next
ITypeInfo_GetRefTypeInfo pTypeInfo, hreftype, ppTInfo
If ppTInfo = 0 Or Err Then
If Not ppTInfo = 0 Then
Set pCustTypeInfo = ResolveObjPtrNoRef(ppTInfo)
Set pCustTypeInfo = Nothing
End If
Err.clear: stringifyCustomType = "UnknownCustomType"


Exit Function
End If
Set pCustTypeInfo = ResolveObjPtrNoRef(ppTInfo)
    ret = ITypeInfo_GetDocumentation(pCustTypeInfo, 0, bstrType, vbNullString, 0, vbNullString)
Set pCustTypeInfo = Nothing
If ret Then stringifyCustomType = "UnknownCustomType": Exit Function
stringifyCustomType = bstrType

End Function
Private Function stringifyTypeDesc(TypeDesc As TTYPEDESC, pTypeInfo As IUnknown) As String
Dim out$, td As TTYPEDESC
If IsBadCodePtr(TypeDesc.pTypeDesc) Then
If TypeDesc.vt = VT_PTR Then
stringifyTypeDesc = "LONG"
ElseIf TypeDesc.vt = VT_USERDEFINED Then
stringifyTypeDesc = "USERDEFINED"
Else
GoTo a123
End If
Exit Function
End If
If TypeDesc.vt = VT_PTR Then
    memcpy td, ByVal TypeDesc.pTypeDesc, Len(td)
    stringifyTypeDesc = stringifyTypeDesc(td, pTypeInfo)
    Exit Function
End If
If TypeDesc.vt = VT_SAFEARRAY Then
out$ = "SAFEARRAY("
    memcpy td, ByVal TypeDesc.pTypeDesc, Len(td)
    stringifyTypeDesc = out$ + stringifyTypeDesc(td, pTypeInfo) + ")"
    Exit Function
End If
If TypeDesc.vt = VT_CARRAY Then
    stringifyTypeDesc = "CArray"
    Exit Function
End If
If TypeDesc.vt = VT_USERDEFINED Then
    memcpy td, ByVal TypeDesc.pTypeDesc, Len(td)
    stringifyTypeDesc = stringifyCustomType(td.pTypeDesc, pTypeInfo) ' hreftype=td.pTypeDesc
    Exit Function
End If
a123:
Select Case TypeDesc.vt
Case VT_I2: stringifyTypeDesc = "Integer"
Case VT_I4: stringifyTypeDesc = "Long"
Case VT_R4: stringifyTypeDesc = "Single"
Case VT_R8: stringifyTypeDesc = "Double"
Case VT_CY: stringifyTypeDesc = "Currency"
Case VT_DATE: stringifyTypeDesc = "Date"
Case VT_BSTR: stringifyTypeDesc = "String"
Case VT_DISPATCH: stringifyTypeDesc = "IDispatch*"
Case VT_ERROR: stringifyTypeDesc = "SCODE"
Case VT_BOOL: stringifyTypeDesc = "VARIANT_BOOL"
Case VT_VARIANT: stringifyTypeDesc = "VARIANT"
Case VT_UNKNOWN: stringifyTypeDesc = "IUnknown*"
Case VT_UI1: stringifyTypeDesc = "BYTE"
Case VT_DECIMAL: stringifyTypeDesc = "DECIMAL"
Case VT_I1: stringifyTypeDesc = "char"
Case VT_UI2: stringifyTypeDesc = "USHORT"
Case VT_UI4: stringifyTypeDesc = "ULONG"
Case VT_I8: stringifyTypeDesc = "__int64"
Case VT_UI8: stringifyTypeDesc = "unsigned __int64"
Case VT_INT: stringifyTypeDesc = "int"
Case VT_UINT: stringifyTypeDesc = "UINT"
Case VT_HRESULT: stringifyTypeDesc = "HRESULT"
Case VT_VOID: stringifyTypeDesc = "void"
Case VT_LPSTR: stringifyTypeDesc = "char*"
Case VT_LPWSTR: stringifyTypeDesc = "wchar_t*"
Case Else
stringifyTypeDesc = "BIG ERROR! " + CStr(TypeDesc.vt)
End Select

End Function
```

### Other low level resources:

https://github.com/M2000Interpreter/Version9/blob/master/modTypeInfo.bas
https://github.com/M2000Interpreter/Version9/blob/master/modObjectExtender.bas
