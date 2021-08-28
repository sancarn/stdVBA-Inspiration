VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' //
' // Access to Forms collection of a VB6-EXE
' // By The trick 2021
' //

Option Explicit

Private Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint  As Long
End Type

Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleInformation Lib "psapi" ( _
                         ByVal hProcess As Long, _
                         ByVal hModule As Long, _
                         ByRef lpmodinfo As MODULEINFO, _
                         ByVal cb As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" _
                         Alias "__vbaObjSetAddref" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" ( _
                    ByRef Addr As Any, _
                    ByRef Dst As Any)
        
Public Property Get Forms() As Object
    Dim hModule     As Long
    Dim pVbHdr      As Long
    Dim tModInfo    As MODULEINFO
    Dim pProjData   As Long
    Dim pExtTable   As Long
    Dim lExtCount   As Long
    Dim lIndex      As Long
    Dim lTag        As Long
    Dim pObjDesc    As Long
    Dim pVBGlobal   As Long
    Dim cVBGlobal   As VB.Global
    
    hModule = GetModuleHandle(0)
    
    If GetModuleInformation(GetCurrentProcess(), hModule, tModInfo, Len(tModInfo)) = 0 Then
        MsgBox "GetModuleInformation failed", vbCritical
        Exit Property
    End If
    
    GetMem4 ByVal tModInfo.EntryPoint + 1, pVbHdr
    GetMem4 ByVal pVbHdr + &H30, pProjData
    GetMem4 ByVal pProjData + &H234, pExtTable
    GetMem4 ByVal pProjData + &H238, lExtCount
    
    For lIndex = 0 To lExtCount - 1
        
        GetMem4 ByVal pExtTable + lIndex * 8, lTag
        
        If lTag = 6 Then
            
            GetMem4 ByVal pExtTable + lIndex * 8 + 4, pObjDesc
            GetMem4 ByVal pObjDesc + 4, pVBGlobal
            GetMem4 ByVal pVBGlobal, pVBGlobal
            
            vbaObjSetAddref cVBGlobal, ByVal pVBGlobal
            
            Set Forms = cVBGlobal.Forms
            
        End If
        
    Next
    
End Property

