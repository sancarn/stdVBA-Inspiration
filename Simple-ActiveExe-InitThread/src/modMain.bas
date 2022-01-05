Attribute VB_Name = "modMain"
Option Explicit

' // Lazy GUID structure
Private Type tCurGUID
    c1          As Currency
    c2          As Currency
End Type

Public Declare Function CreateThread Lib "kernel32" ( _
                        ByRef lpThreadAttributes As Any, _
                        ByVal dwStackSize As Long, _
                        ByVal lpStartAddress As Long, _
                        ByRef lpParameter As Any, _
                        ByVal dwCreationFlags As Long, _
                        ByRef lpThreadId As Long) As Long
Private Declare Function CreateIExprSrvObj Lib "MSVBVM60.DLL" ( _
                         ByVal pUnk1 As Long, _
                         ByVal lUnk2 As Long, _
                         ByVal pUnk3 As Long) As IUnknown
Private Declare Function VBDllGetClassObject Lib "MSVBVM60.DLL" ( _
                         ByRef phModule As Long, _
                         ByVal lReserved As Long, _
                         ByVal pVbHeader As Long, _
                         ByRef pClsid As Any, _
                         ByRef pIID As Any, _
                         ByRef pObject As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function CoInitialize Lib "ole32" ( _
                         ByRef pvReserved As Any) As Long
Private Declare Sub CoUninitialize Lib "ole32" ()
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long

Public g_hModule    As Long
Public g_pVbHeader  As Long

Sub Main()
    ' // No error checking
    g_hModule = App.hInstance
    g_pVbHeader = GetVBHeader()
    frmThread.Show vbModal  ' // To avoid msgloop
End Sub

Public Function ThreadProc( _
                ByVal pVbHdr As Long) As Long
    Dim cExpSrv As IUnknown
    Dim tClsId  As tCurGUID
    Dim tIID    As tCurGUID
    
    Set cExpSrv = CreateIExprSrvObj(0, 4, 0)
    
    CoInitialize ByVal 0&
    
    tIID.c2 = 504403158265495.5712@
    
    VBDllGetClassObject GetModuleHandle(0), 0, pVbHdr, tClsId, tIID, 0
    
    CoUninitialize
    
End Function

Private Function GetVBHeader() As Long
    Dim ptr     As Long
   
    ' // Get e_lfanew
    GetMem4 ByVal g_hModule + &H3C, ptr
    ' // Get AddressOfEntryPoint
    GetMem4 ByVal ptr + &H28 + g_hModule, ptr
    ' // Get VBHeader
    GetMem4 ByVal ptr + g_hModule + 1, GetVBHeader
    
End Function
