Option Explicit
Private Declare Function JsGetUndefinedValue Lib "Chakra.dll" (ByRef UndefinedValue As Long) As Long

Private Declare Function JsConvertValueToString Lib "Chakra.dll" (ByVal m_JsValue As Long, ByVal VARPTR_RESULstr As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function JsStringToPointer Lib "Chakra.dll" (ByVal m_JsValue As Long, stringValue As Long, stringLength As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function JsCreateRuntime Lib "Chakra.dll" (attributes As Long, threadServiceCallback As Long, ByRef runtime As Long) As Long
Private Declare Function JsCreateContext Lib "Chakra.dll" (runtime As Long, ByRef context As Long) As Long
Private Declare Function JsSetCurrentContext Lib "Chakra.dll" (context As Long) As Long
Private Declare Function JsGetGlobalObject Lib "Chakra.dll" (ByRef globalObject As Long) As Long
Private Declare Function JsGetPropertyIdFromName Lib "Chakra.dll" (propertyName As Long, ByRef propertyId As Long) As Long
Private Declare Function JsGetPropertyIdFromNameW Lib "Chakra.dll" (propertyName As Long, ByRef propertyId As Long) As Long
Private Declare Function JsGetProperty Lib "Chakra.dll" (object As Long, propertyId As Long, ByRef value As Long) As Long
Private Declare Function JsCreateStringUtf16 Lib "Chakra.dll" (ByVal string1 As Long, ByVal length As Long, ByRef value As Long) As Long
Private Declare Function JsPointerToString Lib "Chakra.dll" (ByVal string1 As Long, ByVal length As Long, ByRef value As Long) As Long
Private Declare Function JsCreateArray Lib "Chakra.dll" (ByVal length As Long, ByRef array1 As Long) As Long
Private Declare Function JsSetIndexedProperty Lib "Chakra.dll" (ByVal object As Long, ByVal index As Long, ByVal value As Long) As Long
Private Declare Function JsCallFunction Lib "Chakra.dll" (ByVal functionObject As Long, ByVal arguments As Long, ByVal argumentCount As Long, ByVal result As Long) As Long
Private Declare Function JsRunScript Lib "Chakra.dll" (ByVal scriptW As Long, ByVal sourceContext As Long, ByVal sourceUrlW As Long, ByVal handle As Long) As Long
Private Declare Function JsCreateObject Lib "Chakra.dll" (ByRef object1 As Long) As Long

Public Sub Main()
    Dim script As String
    'script = "function AddString(a,b){return a+b; } "
    script = "function AddString(a,b){return 'aaa'; } "
    CallJavaScriptFunction script, "AddString", "aa", "bb"
End Sub
Private Function CallJavaScriptFunction(JsCode As String, ByVal functionName As String, ParamArray arguments() As Variant) As Long
    Dim runtime As Long
    Dim context As Long
    Dim globalObject As Long
    Dim functionObject As Long
    Dim propertyId As Long
    Dim result As Long
    Dim argumentArray As Long
    Dim i As Long, ret As Long
    
    ' 创建 ChakraCore 运行时 (Create the ChakraCore runtime)
    JsCreateRuntime ByVal 0&, ByVal 0&, runtime
    
    ' 创建 ChakraCore 上下文 (Create ChakraCore context)
    JsCreateContext ByVal runtime, context
    
    ' 设置当前上下文 (Set current context)
  ret = JsSetCurrentContext(ByVal context)
    
ret = JsRunScript(StrPtr(JsCode), 0, StrPtr(""), VarPtr(result))  'good
    
    result = 0
    
    ' 获取全局对象 (Get global object)
    
    JsGetGlobalObject globalObject
    If globalObject = 0 Then
    MsgBox "err: globalObject"
    Exit Function
    End If

    ' 获取函数属性 ID (Get function attribute ID)
    JsGetPropertyIdFromName StrPtr(functionName), propertyId
   
    ' 获取函数对象 (Get function object)
    JsGetProperty ByVal globalObject, ByVal propertyId, functionObject
    
    ' 创建参数数组 (Get function object)
    Dim Args As Long
    Args = UBound(arguments) + 1
    'Args = 3
    JsCreateArray Args, argumentArray
    
    Dim UndefinedValue As Long
     JsGetUndefinedValue UndefinedValue
      JsSetIndexedProperty argumentArray, 0, UndefinedValue
    ' 设置参数 (Setting parameters)
    Dim buffer() As Byte
    For i = LBound(arguments) To UBound(arguments)
        Dim argumentValue As Long
'        JsPointerToString StrPtr(CStr(arguments(i))), Len(CStr(arguments(i))), argumentValue
        buffer = StrConv(arguments(i), vbFromUnicode)
        JsPointerToString VarPtr(buffer(0)), UBound(buffer()), argumentValue
        JsSetIndexedProperty argumentArray, i, argumentValue
        
        'JsSetIndexedProperty argumentArray, i + 1, argumentValue
        ' JsSetIndexedProperty argumentArray, i + 1, ByVal 55&
    Next i
    Dim resultValue As Long
     
    
     JsCreateObject resultValue
   
    ' 调用 JavaScript 函数 (Call JavaScript functions)
    result = 0
    'ret = JsCallFunction(functionObject, argumentArray, UBound(arguments) + 1, resultValue)
    ret = JsCallFunction(functionObject, VarPtr(argumentArray), UBound(arguments) + 1, VarPtr(result))
    If ret <> 0 Then result = ret
    MsgBox "result=" & result
    MsgBox JsValueToSTR(result)
    ' 释放 ChakraCore 运行时 (Release the ChakraCore runtime)
    FreeLibrary runtime
    
    CallJavaScriptFunction = result
End Function

 Private Function JsValueToSTR(m_JsValue As Long) As String
'chakra core.dll unsupport :JsValueToVariant，So Use This Function JsValueToSTR

    Dim JsStringPtr As Long
    Dim VbStringPtr As Long
    Dim StringLen As Long
    Dim ret As Long
    ret = JsConvertValueToString(m_JsValue, VarPtr(JsStringPtr))
    ret = JsStringToPointer(JsStringPtr, VbStringPtr, StringLen)
    JsValueToSTR = GetStrFromPtrw(VbStringPtr)
End Function
Public Function GetStrFromPtrw(ByVal Ptr As Long) As String
    'GOOD（ptr前面是4个字节的长度）  (GOOD (ptr is preceded by a length of 4 bytes))
    SysReAllocString VarPtr(GetStrFromPtrw), Ptr
End Function


Private Function CallJavaScriptFunction2(JsCode As String, ByVal functionName As String, ParamArray arguments() As Variant) As Long

    Dim runtime As Long
    Dim context As Long
    Dim globalObject As Long
    Dim functionObject As Long
    Dim propertyId As Long
    Dim result As Long
    Dim argumentArray As Long
    Dim i As Long, ret As Long
    
    ' 创建 ChakraCore 运行时
    JsCreateRuntime ByVal 0&, ByVal 0&, runtime
    
    ' 创建 ChakraCore 上下文
    JsCreateContext ByVal runtime, context
    
    ' 设置当前上下文
  ret = JsSetCurrentContext(ByVal context)
    
ret = JsRunScript(StrPtr(JsCode), 0, StrPtr(""), VarPtr(result)) 'good
    
    result = 0
    
    ' 获取全局对象
    JsGetGlobalObject globalObject
    
    ' 获取函数属性 ID
    JsGetPropertyIdFromName StrPtr(functionName), propertyId
    
    ' 获取函数对象
    JsGetProperty ByVal globalObject, ByVal propertyId, functionObject
    
    ' 创建参数数组
    JsCreateArray UBound(arguments) + 1, argumentArray
    
    ' 设置参数
    For i = LBound(arguments) To UBound(arguments)
        Dim argumentValue As Long
        JsPointerToString StrPtr(CStr(arguments(i))), Len(CStr(arguments(i))), argumentValue
        JsSetIndexedProperty argumentArray, ByVal i, argumentValue
        ' JsSetIndexedProperty argumentArray, i + 0, ByVal 55&
    Next i
    Dim resultValue As Long
     JsCreateObject resultValue
    JsGetUndefinedValue resultValue
    ' 调用 JavaScript 函数
    JsCallFunction functionObject, argumentArray, ByVal UBound(arguments) + 1, VarPtr(resultValue)
    result = resultValue
    MsgBox "result=" & result
    MsgBox JsValueToSTR(result)
    ' 释放 ChakraCore 运行时
    FreeLibrary runtime
    
    CallJavaScriptFunction2 = result
End Function