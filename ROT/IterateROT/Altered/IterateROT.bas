Attribute VB_Name = "IterateROT"
'Example of iterating the ROT for Excel instances
'Author: Jaafar Tribak
'Link:   https://www.mrexcel.com/board/threads/how-to-target-instances-of-excel.1118789/page-2#post-5395037

'Example of iterating the ROT for Excel instances
'Author: Jaafar Tribak
'Link:   https://www.mrexcel.com/board/threads/how-to-target-instances-of-excel.1118789/page-2#post-5395037

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetRunningObjectTable Lib "ole32.dll" (ByVal dwReserved As Long, pROT As LongPtr) As Long
    Private Declare PtrSafe Function CreateBindCtx Lib "ole32.dll" (ByVal dwReserved As Long, pBindCtx As LongPtr) As Long
    Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByVal lpiid As LongPtr) As LongPtr
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Function SysReAllocString Lib "oleAut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#Else
    Private Declare Function GetRunningObjectTable Lib "ole32.dll" (ByVal dwReserved As Long, pROT As Long) As Long
    Private Declare Function CreateBindCtx Lib "ole32.dll" (ByVal dwReserved As Long, pBindCtx As Long) As Long
    Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Function SysReAllocString Lib "oleAut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
#End If


#If Win64 Then
    Private Const IUnknownRelease As Long = 2 * 4 * 2
    Private Const vtbl_EnumRunning_Offset As Long = 9 * 4 * 2
    Private Const vtbl_EnumMoniker_Next_Offset As Long = 3 * 4 * 2
    Private Const vtbl_Moniker_GetDisplayName_offset As Long = 20 * 4 * 2
#Else
    Private Const IUnknownRelease As Long = 2 * 4
    Private Const vtbl_EnumRunning_Offset As Long = 9 * 4
    Private Const vtbl_EnumMoniker_Next_Offset As Long = 3 * 4
    Private Const vtbl_Moniker_GetDisplayName_offset As Long = 20 * 4
#End If



Private Const IUnknownQueryInterface As Long = 0
Private Const CC_STDCALL As Long = 4
Private Const S_OK As Long = 0
Private Const ROT_INTERFACE_ID As String = "{00000010-0000-0000-C000-000000000046}"


Sub PrintAllROTNames()
    #If VBA7 Then
        Dim pROT As LongPtr, pRunningObjectTable As LongPtr, pEnumMoniker As LongPtr, pMoniker As LongPtr, pBindCtx As LongPtr, hRes As LongPtr, pName As LongPtr
    #Else
        Dim pROT As Long, pRunningObjectTable As Long, pEnumMoniker As Long, pMoniker As Long, pBindCtx As Long, hRes As Long, pName As Long
    #End If
    
    Dim uGUID(0 To 3) As Long
    Dim sTempArray() As String
    Dim oTempObj As Object
    Dim lRet As Long, nCount As Long, lMatchPos1 As Long, lMatchPos2 As Long
    
    
    If CLng(GetRunningObjectTable(0, pROT)) <> S_OK Then Exit Sub
    If CLng(CreateBindCtx(0, pBindCtx)) <> S_OK Then Exit Sub
    If CLng(IIDFromString(StrPtr(ROT_INTERFACE_ID), VarPtr(uGUID(0)))) <> 0 Then Exit Sub
    If CallFunction_COM(pROT, IUnknownQueryInterface, vbLong, CC_STDCALL, VarPtr(uGUID(0)), (VarPtr(pRunningObjectTable))) <> S_OK Then Exit Sub
    If CallFunction_COM(pRunningObjectTable, vtbl_EnumRunning_Offset, vbLong, CC_STDCALL, (VarPtr(pEnumMoniker))) <> S_OK Then Exit Sub
    
    nCount = nCount + 1
    While CallFunction_COM(pEnumMoniker, vtbl_EnumMoniker_Next_Offset, vbLong, CC_STDCALL, nCount, (VarPtr(pMoniker)), VarPtr(nCount)) = S_OK
        'Get name `sShortPathName`
        Dim sName As String
        If CallFunction_COM(pMoniker, vtbl_Moniker_GetDisplayName_offset, vbLong, CC_STDCALL, VarPtr(pBindCtx), VarPtr(pMoniker), VarPtr(pName)) = S_OK Then _
          sName = GetStrFromPtrW(pName)
        Debug.Print sName
        
        
        
        CallFunction_COM pMoniker, IUnknownRelease, vbLong, CC_STDCALL
    Wend
    
    CallFunction_COM pEnumMoniker, IUnknownRelease, vbLong, CC_STDCALL
    CallFunction_COM pBindCtx, IUnknownRelease, vbLong, CC_STDCALL
    CallFunction_COM pRunningObjectTable, IUnknownRelease, vbLong, CC_STDCALL
    CallFunction_COM pROT, IUnknownRelease, vbLong, CC_STDCALL
End Sub


#If VBA7 Then
    Private Function CallFunction_COM(ByVal InterfacePointer As LongPtr, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    
    Dim vParamPtr() As LongPtr
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
                                                      
    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then
        CallFunction_COM = vRtn
    Else
        SetLastError pIndex
    End If

End Function


#If VBA7 Then
    Private Function GetStrFromPtrW(ByVal ptr As LongPtr) As String
#Else
    Private Function GetStrFromPtrW(ByVal ptr As Long) As String
#End If
    SysReAllocString VarPtr(GetStrFromPtrW), ptr
End Function
