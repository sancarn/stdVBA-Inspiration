VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCDECL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'** cCallFunc2.cls - cCallFunc with __fastcall support, call by address and
'**                  additional return types
'**
'** Universal dll function calling class
'**   cdecl/stdcall/__fastcall calling convention
'**   Call functions by ordinal, name or address
'**   Module (.bas) callbacks for cdecl functions.
'**   Object (.cls/.frm/.ctl) callbacks for cdecl/stdcall
'**   Support for multiple callbacks.
'**   Support for multiple cCallFunc2 instances
'**   Support unicode path\module names
'**
'** If you wish to do crazy stuff like CallFunc with callbacks inside a callback
'** then the best solution is to make a copy of the class, eg cCallFunc1.cls, and
'** use an instance of that where needed.
'**
'** Calling conventions:
'**   stdcall:  parameters right to left, called function corrects the stack
'**   cdecl:    parameters right to left, caller corrects the stack
'**   __fastcall: first parameter in the ecx register
'**               second parameter, if present, in the edx register
'**               any other parameters are pushed on to the stack
'**
'**   Note:       __fastcall isn't standardised, differing conventions exist.
'**               This class supports the Microsoft/GCC implementation.
'**
'** paul_caton@hotmail.com
'**
'** 20031029 First cut....................................................... v1.00
'** 20071129 Now using virtual memory to fix a DEP issue..................... v1.01
'** 20071130 Hacked from cCDECL, now supports stdcall and ordinals......... v1.02
'** 20071201 Added support for callback objects.............................. v1.03
'** 20071202 Unicode support for paths\modules where available............... v1.04
'** 20071213 Forked from cCallFunc.cls
'**          Added support for __fastcall calling convention
'**          Added CallPointer
'**          Changed the interface to be more property like.................. v1.05
'** 20080212 Support Byte, Integer, Long, Single and Double return types..... v1.06
'**********************************************************************************

Option Explicit

'API declarations
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetProcByName Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal nOrdinal As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Public Enum eObjType                                                          'Object type for CallbackObj
  objCls = &H1C                                                               'Class object callback
  objFrm = &H6F8                                                              'Form object callback
  objCtl = &H7A4                                                              'UserControl object callback
End Enum                                                                      '
                                                                              '
Public Enum eReturnType                                                       'CallFunc/CallPointer return types
  retByte = 0                                                                 'Return Byte
  retInteger                                                                  'Return Integer
  retLong                                                                     'Return Long
  retInt64                                                                    'Return 64 bit value eg. Currency
  retSingle                                                                   'Return Single
  retDouble                                                                   'Return Double
  retSub                                                                      'No return value
End Enum                                                                      '
                                                                              '
Private Const SRC           As String = "cCallFunc2."                         'Error source
Private Const MAX_ARG       As Long = 16                                      'Maximum number of parameters, this value can altered if needed
                                                                              '
Private Type tParamBlock                                                      'Parameter block type
  ParamCount                As Long                                           'Number of parameters
  Params(0 To MAX_ARG - 1)  As Long                                           'Array of parameters
End Type                                                                      '
                                                                              '
Private m_FastCall          As Boolean                                        'FastCall private property value
Private m_LastError         As Long                                           'LastError private property value
                                                                              
Private bUnicode            As Boolean                                        'Unicode flag                                                                              '
Private vTable              As Long                                           'cCallFunc2 vTable address
Private vCode               As Long                                           'Pointer to the machine-code thunks
Private hModule             As Long                                           'Current/last-used dll handle
Private strLastDLL          As String                                         'Current/last-used dll name
Private strLastFunc         As String                                         'Current/last-used function name
                                                                              '
Public Property Get FastCall() As Boolean                                     'Get FastCall flag
  FastCall = m_FastCall                                                       '
End Property                                                                  '
                                                                              '
Public Property Let FastCall(ByVal bValue As Boolean)                         'Let Fastcall flag
  m_FastCall = bValue                                                         '
                                                                              '
  If m_FastCall Then                                                          '
    PutMem2 vCode + &H11, &H34EB                                              'Patch the code to jump to the Fastcall parameter processor
  Else                                                                        '
    PutMem2 vCode + &H11, &H9090                                              'Patch the code to fall through to the cdecl/stdcall parameter processor
  End If                                                                      '
End Property                                                                  '
                                                                              
Public Property Get LastError() As Long                                       'Get last error
  LastError = m_LastError                                                     '
End Property                                                                  '
                                                                              '
'CallFunc:
'
' strDLL      - Name of the DLL
' ReturnType  - Function return type
' strFunction - Name of the function or it's ordinal value preceded by a '#' eg. "#2"
' ParamLongs  - Any number [or none] of parameters As Long.
'                 To pass the address (ByRef) of a string use StrPtr, eg. StrPtr(strPath)
'                 To pass the address (ByRef) of a variable or UDT use VarPtr, eg. VarPtr(i)
Public Function CallFunc(ByRef strDLL As String, _
                         ByVal ReturnType As eReturnType, _
                         ByRef strFunction As String, _
                    ParamArray ParamLongs() As Variant) As Variant            '
  Dim pb      As tParamBlock                                                  'Parameter block
  Dim bNewDll As Boolean                                                      'New dll flag
                                                                              '
  If StrComp(strDLL, strLastDLL, vbTextCompare) <> 0 Then                     'If the module is new
    Dim hMod As Long                                                          '
                                                                              '
    If bUnicode Then                                                          'If unicode
      hMod = LoadLibraryW(StrPtr(strDLL & vbNullChar))                        'Load the module with the unicode version of LoadLibrary
    Else                                                                      '
      hMod = LoadLibraryA(strDLL)                                             'Load the module with the ascii version of LoadLibrary
    End If                                                                    '
                                                                              '
    If hMod = 0 Then                                                          'If the load failed
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 0, SRC & "CallFunc", "DLL failed load"        'Raise an error if running compiled
    End If                                                                    '
                                                                              '
    If hModule <> 0 Then                                                      'If a module is already loaded
      FreeLibrary hModule                                                     'Free the last module
    End If                                                                    '
                                                                              '
    hModule = hMod                                                            'Save the module handle
    strLastDLL = strDLL                                                       'Save the new module name
    bNewDll = True                                                            'Indicate that it's a new module
  End If                                                                      '
                                                                              '
  If bNewDll Or StrComp(strFunction, strLastFunc, vbBinaryCompare) <> 0 Then  'If the function or module is new
    Dim nFuncAddr As Long                                                     'Function address
                                                                              '
    If Asc(strFunction) = 35 Then                                             'If "#..." eg "#2", ordinal 2
      nFuncAddr = GetProcByOrdinal(hModule, CLng(Mid$(strFunction, 2)))       'Get the address of the function by ordinal
    Else                                                                      '
      nFuncAddr = GetProcByName(hModule, strFunction)                         'Get the address of the function by name
    End If                                                                    '
                                                                              '
    If nFuncAddr = 0 Then                                                     'If the function wasn't found in the module
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 1, SRC & "CallFunc", "Function not found"     'Raise an error if running compiled
    End If                                                                    '
                                                                              '
    PutMem4 vCode + &H19, nFuncAddr - vCode - (&H19 + 4)                      'Patch the code to call the relative address to the target function
    strLastFunc = strFunction                                                 'Save the function name
  End If                                                                      '
                                                                              
  With pb                                                                     '
    Dim I  As Long                                                            'Parameter loop vars
    Dim j  As Long                                                            'Parameter loop vars
                                                                              '
    j = UBound(ParamLongs)                                                    'Get the upper parameter array bound
                                                                              '
    If j >= MAX_ARG Then                                                      'If the user has passed more parameters than we allow for
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 2, SRC & "CallFunc", "Too many parameters"    'Raise error if running compiled
    End If                                                                    '
                                                                              '
    For I = 0 To j                                                            'For each parameter
      .Params(I) = ParamLongs(I)                                              'Store the parameter in the parameter block
    Next I                                                                    '
                                                                              '
    .ParamCount = I                                                           'Store the parameter count (j + 1)
  End With                                                                    '
                                                                              
  CallFunc = CallCommon(ReturnType, VarPtr(pb))                               'Call common code
End Function                                                                  '

'CallPointer: call a function by address
'
' ReturnType  - Function return type
' fnAddress   - Address of the target function
' ParamLongs  - Any number of parameters As Long, or none.
'                 To pass the address (ByRef) of a string use StrPtr, eg. StrPtr(strPath)
'                 To pass the address (ByRef) of a variable or UDT use VarPtr, eg. VarPtr(i)
Public Function CallPointer(ByVal ReturnType As eReturnType, _
                            ByVal fnAddress As Long, _
                       ParamArray ParamLongs() As Variant) As Variant         '
  Dim pb As tParamBlock                                                       'Parameter block
  Dim I  As Long                                                              'Parameter loop vars
  Dim j  As Long                                                              'Parameter loop vars
                                                                              '
  With pb                                                                     '
    j = UBound(ParamLongs)                                                    'Get the upper parameter array bound
                                                                              '
    If j >= MAX_ARG Then                                                      'If the user has passed more parameters than we allow for
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 3, SRC & "CallPointer", "Too many parameters" 'Raise error if running compiled
    End If                                                                    '
                                                                              '
    For I = 0 To j                                                            'For each parameter
      .Params(I) = ParamLongs(I)                                              'Store the parameter in the parameter block
    Next I                                                                    '
                                                                              '
    .ParamCount = I                                                           'Store the parameter count
  End With                                                                    '
                                                                              '
  strLastFunc = vbNullString                                                  '
  PutMem4 vCode + &H19, fnAddress - vCode - (&H19 + 4)                        'Patch the code to call the relative address to the target function
                                                                              '
  CallPointer = CallCommon(ReturnType, VarPtr(pb))                            'Call common code
End Function

'CallCommon: common CallFunc/CallPointer code
'
' ReturnType  - Function return type
' pb          - Address of the parameter block
Private Function CallCommon(ByVal ReturnType As eReturnType, _
                            ByVal pb As Long) As Variant
  SetLastError 0                                                              'Clear the error code
                                                                              '
  Select Case ReturnType                                                      'Select on return type
  Case eReturnType.retByte                                                    '
    PutMem4 vTable + &H44, vCode                                              'Patch the vtable entry for z_CallFunc_i08 to point to vCode
    PutMem1 vCode + &H27, 0                                                   'Patch the return type code
    CallCommon = z_CallFunc_i08(pb)                                           'Call the function, return a Byte
                                                                              '
  Case eReturnType.retInteger                                                 '
    PutMem4 vTable + &H48, vCode                                              'Patch the vtable entry for z_CallFunc_i16 to point to vCode
    PutMem1 vCode + &H27, 4                                                   'Patch the return type code
    CallCommon = z_CallFunc_i16(pb)                                           'Call the function, return a Integer
                                                                              '
  Case eReturnType.retLong                                                    '
    PutMem4 vTable + &H4C, vCode                                              'Patch the vtable entry for z_CallFunc_i32 to point to vCode
    PutMem1 vCode + &H27, 9                                                   'Patch the return type code
    CallCommon = z_CallFunc_i32(pb)                                           'Call the function, return an Long
                                                                              '
  Case eReturnType.retInt64                                                   '
    PutMem4 vTable + &H50, vCode                                              'Patch the vtable entry for z_CallFunc_i64 to point to vCode
    PutMem1 vCode + &H27, &HD                                                 'Patch the return type code
    CallCommon = z_CallFunc_i64(pb)                                           'Call the function, return an int64
                                                                              '
  Case eReturnType.retSingle                                                  '
    PutMem4 vTable + &H54, vCode                                              'Patch the vtable entry for z_CallFunc_Sng to point to vCode
    PutMem1 vCode + &H27, &H14                                                'Patch the return type code
    CallCommon = z_CallFunc_Sng(pb)                                           'Call the function, return a Single
                                                                              '
  Case eReturnType.retDouble                                                  '
    PutMem4 vTable + &H58, vCode                                              'Patch the vtable entry for z_CallFunc_Dbl to point to vCode
    PutMem1 vCode + &H27, &H18                                                'Patch the return type code
    CallCommon = z_CallFunc_Dbl(pb)                                           'Call the function, return a Double
                                                                              '
  Case eReturnType.retSub                                                     '
    PutMem4 vTable + &H5C, vCode                                              'Patch the vtable entry for z_CallFunc_Sub to point to vCode
    PutMem1 vCode + &H27, &H1A                                                'Patch the return type code
    z_CallFunc_Sub pb                                                         'Call the function, no return value
                                                                              '
  Case Else                                                                   'Undefined return type
    Debug.Assert False                                                        'Halt if running under the VB IDE
    Err.Raise vbObjectError + 4, SRC & "CallCommon", "Unknown return type"    'Raise error if running compiled
                                                                              '
  End Select                                                                  '
                                                                              '
  m_LastError = GetLastError()                                                'Get the error code
End Function

'CallbackCdecl: return a wrapper address for a bas module routine to be used as a callback for a cdecl function.
'               Note: stdcall functions don't need a thunk to use a bas module function as a callback, use direct.
'
' nModFuncAddr - The address of the bas module callback function, use AddressOf to get this value
' nParms       - The number of parameters that will be passed to the bas module callback function
' nIndex       - Allow for multiple simultaneous callbacks
Public Function CallbackCdecl(ByVal nModFuncAddr As Long, _
                              ByVal nParams As Long, _
                     Optional ByVal nIndex As Long = 1) As Long
  
  If nIndex < 1 Or nIndex > 60 Or nParams > 60 Then                           'Parameter sanity checks
    Debug.Assert False                                                        'Halt if running under the VB IDE
    Err.Raise vbObjectError + 5, SRC & "CallbackCdecl", "Invalid parameter"   'Raise error if running compiled
  End If                                                                      '
                                                                              '
  CallbackCdecl = vCode + 128 + ((nIndex - 1) * 64)                           'Address of the callback wrapper. Pass this return value as the callback address parameter of the cdecl function
                                                                              '
  PutMem8 CallbackCdecl + 0, 465203369712025.6232@                            'Put the cdecl callback to .bas wrapper into memory
  PutMem8 CallbackCdecl + 8, -140418483381718.8329@                           '
  PutMem8 CallbackCdecl + 16, -4672484613390.9419@                            '
  PutMem4 CallbackCdecl + 24, &HC30672                                        '
  PutMem4 CallbackCdecl + 10, nModFuncAddr - CallbackCdecl - (10 + 4)         'Patch the code buffer to call the vb bas module callback function
  PutMem1 CallbackCdecl + 16, nParams * 4                                     'Patch the code buffer to apply the necessary stack adjustment
End Function                                                                  '
                                                                              '
'CallbackObj: return a wrapper address for an object callback for a cdecl or stdcall function
'
' objType     - Callback object type
' objCallback - The callback object
' nParams     - The number of parameters that will be passed to the object callback function
' nOrdinal    - Callback ordinal. 1 = last private function in the callback object, 2 = second last private function in the callback object, etc
' bCDECL      - specifes whether the callback calling function is cdecl or stdcall
' nIndex      - Allow for multiple simultaneous callbacks
Public Function CallbackObj(ByVal objType As eObjType, _
                            ByRef objCallback As Object, _
                            ByVal nParams As Long, _
                   Optional ByVal nOrdinal As Long = 1, _
                   Optional ByVal bCDECL As Boolean = False, _
                   Optional ByVal nIndex As Long = 1) As Long
  Dim o As Long                                                               'Object pointer
  Dim I As Long                                                               'vTable entry counter
  Dim j As Long                                                               'vTable address
  Dim n As Long                                                               'Method pointer
  Dim b As Byte                                                               'First method byte
  Dim m As Byte                                                               'Known good first method byte
                                                                              '
  If nIndex < 1 Or nIndex > 60 Or nParams > 60 Then                           'Parameter sanity checks
    Debug.Assert False                                                        'Halt if running under the VB IDE
    Err.Raise vbObjectError + 6, SRC & "CallbackObj", "Invalid parameter"     'Raise error if running compiled
  End If                                                                      '
                                                                              '
  o = ObjPtr(objCallback)                                                     'Get the callback object's address
  GetMem4 o, j                                                                'Get the address of the callback object's vTable
  j = j + objType                                                             'Increment to the the first user entry for this callback object type
  GetMem4 j, n                                                                'Get the method pointer
  GetMem1 n, m                                                                'Get the first method byte... &H33 if pseudo-code, &HE9 if native
  j = j + 4                                                                   'Bump to the next vtable entry
                                                                              '
  For I = 1 To 511                                                            'Loop through a 'sane' number of vtable entries
    GetMem4 j, n                                                              'Get the method pointer
                                                                              '
    If IsBadCodePtr(n) Then                                                   'If the method pointer is an invalid code address
      GoTo vTableEnd                                                          'We've reached the end of the vTable, exit the for loop
    End If                                                                    '
                                                                              '
    GetMem1 n, b                                                              'Get the first method byte
                                                                              '
    If b <> m Then                                                            'If the method byte doesn't matche the known good value
      GoTo vTableEnd                                                          'We've reached the end of the vTable, exit the for loop
    End If                                                                    '
                                                                              '
    j = j + 4                                                                 'Bump to the next vTable entry
  Next I                                                                      'Bump counter
  
  Debug.Assert False                                                          'Halt if running under the VB IDE
  Err.Raise vbObjectError + 7, SRC & "CallbackObj", "Ordinal not found"       'Raise error if running compiled
                                                                              '
vTableEnd:                                                                    'We've hit the end of the vTable
  GetMem4 j - (nOrdinal * 4), n                                               'Get the method pointer for the specified ordinal
                                                                              '
  CallbackObj = vCode + 128 + ((nIndex - 1) * 64)                             'Address of the callback wrapper. Pass this return value as the callback address parameter
                                                                              '
  PutMem8 CallbackObj + 0, 648518346342877.6073@                              'Put the callback to object wrapper into memory
  PutMem8 CallbackObj + 8, 9425443564.8131@                                   '
  PutMem8 CallbackObj + 16, -29652486425477.8624@                             '
  PutMem8 CallbackObj + 24, 614907631944580.0296@                             '
  PutMem8 CallbackObj + 32, 648518346342874.0437@                             '
  PutMem8 CallbackObj + 40, 36832747.5851@                                    '
                                                                              '
  PutMem1 CallbackObj + &HD, nParams                                          'Patch the number of params
  PutMem4 CallbackObj + &H19, o                                               'Patch the callback object
  PutMem4 CallbackObj + &H1E, n - CallbackObj - (&H1E + 4)                    'Patch the callback call address
  PutMem1 CallbackObj + &H2C, IIf(bCDECL, 0, nParams * 4)                     'Patch the stack correction
End Function                                                                  '
                                                                              '
'Class_Initialize: initialize the cCallFunc2 instance
Private Sub Class_Initialize()                                                '
  vCode = VirtualAlloc(0, &H1000&, &H1000&, &H40&)                            'Allocate read/write/executable memory
                                                                              '
  PutMem8 vCode + 0, 695618785647368.6248@                                    '
  PutMem8 vCode + 8, -208726556020175.3831@                                   '
  PutMem8 vCode + 16, -29652486425169.6377@                                   '
  PutMem8 vCode + 24, 614902794093417.828@                                    '
  PutMem8 vCode + 32, 193965741455568.6229@                                   '
  PutMem8 vCode + 40, -151277692825560.6392@                                  '
  PutMem8 vCode + 48, -857442152266638.7183@                                  '
  PutMem8 vCode + 56, 21029022751752.3025@                                    '
  PutMem8 vCode + 64, -151319984225536.5667@                                  '
  PutMem8 vCode + 72, 21434680911783.5012@                                    '
  PutMem8 vCode + 80, 59913160095353.8431@                                    '
  PutMem2 vCode + 88, &HBEEB                                                  '
                                                                              '
  GetMem4 ObjPtr(Me), vTable                                                  'Get the address of the class vTable
                                                                              '
  If GetProcByName(LoadLibraryA("user32"), "IsWindowUnicode") Then            'Is IsWindowUnicode present
    bUnicode = IsWindowUnicode(GetDesktopWindow())                            'Determine whether we'll use the unicode version of LoadLibrary
  End If                                                                      '
                                                                              '
  FastCall = False                                                            '
End Sub                                                                       '
                                                                              '
'Class_Terminate: cleanup the cCallFunc2 instance
Private Sub Class_Terminate()                                                 '
  If hModule <> 0 Then                                                        'If a module is loaded
    FreeLibrary hModule                                                       'Free the loaded module
  End If                                                                      '
                                                                              '
  VirtualFree vCode, 0, &H8000&                                               'Free the allocated memory
End Sub
                                                                              
'**********************************************************************************************************
' These following function's vTable method pointers are patched to point to vCode in CallFunc & CallPointer
' Note: these functions must be private and cannot be moved within this source file.
'**********************************************************************************************************

'z_CallFunc_i08: return Byte
Private Function z_CallFunc_i08(ByVal nParmAddr As Long) As Byte              '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Function                                                                  '

'z_CallFunc_i16: return Integer
'
' nParmAddr - address of the parameter block
Private Function z_CallFunc_i16(ByVal nParmAddr As Long) As Integer           '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Function                                                                  '

'z_CallFunc_i32: return Long
'
' nParmAddr - address of the parameter block
Private Function z_CallFunc_i32(ByVal nParmAddr As Long) As Long              '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Function                                                                  '

'z_CallFunc_i64: return int64
'
' nParmAddr - address of the parameter block
Private Function z_CallFunc_i64(ByVal nParmAddr As Long) As Currency          '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Function                                                                  '

'z_CallFunc_Sng: return Single
'
' nParmAddr - address of the parameter block
Private Function z_CallFunc_Sng(ByVal nParmAddr As Long) As Single            '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Function                                                                  '

'z_CallFunc_Dbl: return Double
'
' nParmAddr - address of the parameter block
Private Function z_CallFunc_Dbl(ByVal nParmAddr As Long) As Double            '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Function                                                                  '

'z_CallFunc_Sub: no return value
'
' nParmAddr - address of the parameter block
Private Sub z_CallFunc_Sub(ByVal nParmAddr As Long)                           '
  Debug.Assert False                                                          'Halt if running under the VB IDE
End Sub                                                                       '


