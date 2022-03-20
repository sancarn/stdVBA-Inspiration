'## Original source
'Author: Wqweto
'Link:   https://www.vbforums.com/showthread.php?879529-project-one-workng-with-project2&p=5422507&viewfull=1#post5422507
'## VB7 Upgrade
'Author: Jaafar Tribak
'Link:   https://www.mrexcel.com/board/threads/reference-and-remotely-manipulate-userforms-loaded-in-seperate-workbooks-or-in-seperate-excel-instances-via-file-monikers.1161038/#post-5634620

Option Explicit

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As LongPtr) As Long
    Private Declare PtrSafe Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As LongPtr, pResult As LongPtr) As Long
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Sub PathRemoveExtension Lib "shlwapi.dll" Alias "PathRemoveExtensionA" (ByVal pszPath As String)
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As Long) As Long
    Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As Long) As Long
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
    Private Declare Sub PathRemoveExtension Lib "shlwapi.dll" Alias "PathRemoveExtensionA" (ByVal pszPath As String)
 #End If
 
 
 
Public Sub PutInROT(ByVal UForm As Object, Optional ByVal CallByCaption As Boolean)

    Const GWLP_USERDATA = &HFFFFFFEB
    
    #If Win64 Then
        Dim hwnd As LongLong
    #Else
        Dim hwnd As Long
    #End If

    Dim lCookie As Long
    Dim sWBookName As String, sNameOrCaption As String

    sWBookName = ThisWorkbook.Name
    Call PathRemoveExtension(sWBookName)
    If InStr(sWBookName, vbNullChar) Then
        sWBookName = Left(sWBookName, InStr(sWBookName, vbNullChar) - 1)
    End If
    
    If CallByCaption And Len(UForm.Caption) Then
        sNameOrCaption = UForm.Caption
    Else
        sNameOrCaption = UForm.Name
    End If

    lCookie = RegisterObjectInROT(UForm, sWBookName & "." & sNameOrCaption)
    
    If lCookie Then
        Call IUnknown_GetWindow(UForm, VarPtr(hwnd))
        Call SetWindowLong(hwnd, GWLP_USERDATA, lCookie)
    End If
    
End Sub


Public Sub RemoveFromROT(ByVal UForm As Object)

    Const GWLP_USERDATA = &HFFFFFFEB
    
    #If Win64 Then
        Dim hwnd As LongLong
    #Else
        Dim hwnd As Long
    #End If
    
    Dim lCookie As Long

    Call IUnknown_GetWindow(UForm, VarPtr(hwnd))
    lCookie = CLng(GetWindowLong(hwnd, GWLP_USERDATA))
    If lCookie Then Call RevokeObject(lCookie)

End Sub

 
 
 '___________________________________SUPPORTING ROUTINES________________________________
 
Private Function RegisterObjectInROT(Obj As Object, sPathName As String) As Long

    Const ROTFLAGS_REGISTRATIONKEEPSALIVE = 1
    Const REGISTER_VTBL_OFFSET  As Long = 3
    Const CC_STDCALL = 4
    Const S_OK = 0

    #If Win64 Then
        Const PTR_LEN = 8
        Dim pROT As LongLong
        Dim pMoniker As LongLong
    #Else
        Const PTR_LEN = 4
        Dim pROT As Long
        Dim pMoniker As Long
    #End If

    If GetRunningObjectTable(0, pROT) <> S_OK Then
        MsgBox "GetRunningObjectTable failed !": Exit Function
    End If
    If CreateFileMoniker(StrPtr(sPathName), pMoniker) <> S_OK Then
           MsgBox "CreateFileMoniker failed !": Exit Function
    End If
    
    vtblCall pROT, REGISTER_VTBL_OFFSET * PTR_LEN, vbLong, _
    CC_STDCALL, ROTFLAGS_REGISTRATIONKEEPSALIVE, Obj, pMoniker, VarPtr(RegisterObjectInROT)

End Function
 
Private Sub RevokeObject(ByVal lCookie As Long)

    Const REVOKE_VTBL_OFFSET = 4
    Const CC_STDCALL = 4
    Const S_OK = 0

    #If Win64 Then
        Const PTR_LEN = 8
        Dim pROT As LongLong
    #Else
        Const PTR_LEN = 4
        Dim pROT As Long
    #End If
    
    If GetRunningObjectTable(0, pROT) <> S_OK Then
        MsgBox "GetRunningObjectTable failed !": Exit Sub
    End If
    
    vtblCall pROT, REVOKE_VTBL_OFFSET * PTR_LEN, vbLong, CC_STDCALL, lCookie

End Sub



#If Win64 Then
    Private Function vtblCall(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
    Dim vParamPtr() As LongLong
#Else
    Private Function vtblCall(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant
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
        vtblCall = vRtn
    Else
        SetLastError pIndex
    End If

End Function
