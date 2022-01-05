' // Get command line of process
' // 64-bit compatible
' // By The trick

Option Explicit

Private Const ProcessBasicInformation           As Long = 0
Private Const SystemProcessInformation          As Long = 5
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const PROCESS_VM_READ                   As Long = &H10
Private Const SE_PRIVILEGE_ENABLED              As Long = 2
Private Const TOKEN_ADJUST_PRIVILEGES           As Long = &H20
Private Const SE_DEBUG_NAME                     As String = "SeDebugPrivilege"

Private Type UNICODE_STRING
    Length      As Integer
    MaxLength   As Integer
    lpBuffer    As Long
End Type

Private Type UNICODE_STRING64
    Length      As Integer
    MaxLength   As Integer
    lPad        As Long
    lpBuffer    As Currency
End Type

Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                      As Long
    PebBaseAddress                  As Long
    AffinityMask                    As Long
    BasePriority                    As Long
    UniqueProcessId                 As Long
    InheritedFromUniqueProcessId    As Long
End Type

Private Type PROCESS_BASIC_INFORMATION64
    ExitStatus                      As Long
    Reserved0                       As Long
    PebBaseAddress                  As Currency
    AffinityMask                    As Currency
    BasePriority                    As Long
    Reserved1                       As Long
    uUniqueProcessId                As Currency
    uInheritedFromUniqueProcessId   As Currency
End Type

Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As OLE_HANDLE
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As OLE_HANDLE, _
                         ByVal lpProcName As String) As Long
Private Declare Function IsWow64Process Lib "kernel32" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByRef lpWow64Process As Long) As Long
Private Declare Function NtQueryInformationProcess Lib "ntdll" ( _
                         ByVal ProcessHandle As OLE_HANDLE, _
                         ByVal InformationClass As Long, _
                         ByRef ProcessInformation As Any, _
                         ByVal ProcessInformationLength As Long, _
                         ByRef ReturnLength As Long) As Long
Private Declare Function NtQuerySystemInformation Lib "ntdll" ( _
                         ByVal SystemInformationClass As Long, _
                         ByRef ProcessInformation As Any, _
                         ByVal ProcessInformationLength As Long, _
                         ByRef ReturnLength As Long) As Long
Private Declare Function NtWow64QueryInformationProcess64 Lib "ntdll" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByVal ProcessInformationClass As Long, _
                         ByRef pProcessInformation As Any, _
                         ByVal ProcessInformationLength As Long, _
                         ByRef puReturnLength As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByVal lpBaseAddress As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nSize As Long, _
                         ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function NtWow64ReadVirtualMemory64 Lib "ntdll" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByVal BaseAddress As Currency, _
                         ByRef Buffer As Any, _
                         ByVal BufferLengthL As Long, _
                         ByVal BufferLengthH As Long, _
                         ByRef ReturnLength As Currency) As Long
Private Declare Function OpenProcess Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwProcessId As Long) As OLE_HANDLE
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As OLE_HANDLE) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" ( _
                         ByRef psz As Any, _
                         ByVal lLen As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
                         Alias "LookupPrivilegeValueW" ( _
                         ByVal lpSystemName As Long, _
                         ByVal lpName As Long, _
                         ByRef lpLuid As Any) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" ( _
                         ByVal TokenHandle As OLE_HANDLE, _
                         ByVal DisableAllPrivileges As Long, _
                         ByRef NewState As Any, _
                         ByVal BufferLength As Long, _
                         ByRef PreviousState As Any, _
                         ByRef ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As OLE_HANDLE
Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
                         ByVal ProcessHandle As OLE_HANDLE, _
                         ByVal DesiredAccess As Long, _
                         ByRef TokenHandle As OLE_HANDLE) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any)
Private Declare Sub PutMem4 Lib "msvbvm60" (Dst As Any, ByVal lVal As Long)

Private Sub Main()
    Dim bProcInfo() As Byte
    Dim lSize       As Long
    Dim lNextOffset As Long
    Dim lCurOffset  As Long
    Dim lPID        As Long
    Dim hProcess    As OLE_HANDLE
    Dim sName       As String
    Dim pStr        As Long
    
    SetPrivilege SE_DEBUG_NAME, True
    
    lSize = 1024
    
    Do
        ReDim bProcInfo(lSize - 1)
    Loop While NtQuerySystemInformation(SystemProcessInformation, bProcInfo(0), lSize, lSize)
    
    Do
        
        lCurOffset = lCurOffset + lNextOffset
        GetMem4 bProcInfo(lCurOffset + &H44), lPID
        
        hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION Or PROCESS_VM_READ, 0, lPID)
        
        If hProcess Then
            
            sName = vbNullString
            GetMem4 bProcInfo(lCurOffset + &H3C), pStr
            PutMem4 ByVal VarPtr(sName), SysAllocStringByteLen(ByVal pStr, bProcInfo(lCurOffset + &H38) Or &H100& * bProcInfo(lCurOffset + &H39))
            
            On Error GoTo close_handle
            
            Debug.Print sName, lPID, GetCommandLineOfProcess(hProcess)
            
close_handle:
            
            CloseHandle hProcess
            
        End If
        
        GetMem4 bProcInfo(lCurOffset), lNextOffset
        
    Loop While lNextOffset
    

End Sub

Public Function GetCommandLineOfProcess( _
                ByVal hProcess As OLE_HANDLE) As String
    Dim tPEB32      As PROCESS_BASIC_INFORMATION
    Dim tPEB64      As PROCESS_BASIC_INFORMATION64
    Dim tCMDLine    As UNICODE_STRING
    Dim tCMDLine64  As UNICODE_STRING64
    Dim pParam      As Long
    Dim pParam64    As Currency
    
    If Is32BitProcess(hProcess) Then
    
        If NtQueryInformationProcess(hProcess, ProcessBasicInformation, tPEB32, LenB(tPEB32), 0) < 0 Then
            Err.Raise 7
        End If
        
        If ReadProcessMemory(hProcess, tPEB32.PebBaseAddress + &H10, pParam, LenB(pParam), 0) = 0 Then
            Err.Raise 7
        End If
        
        If ReadProcessMemory(hProcess, pParam + &H40, tCMDLine, LenB(tCMDLine), 0) = 0 Then
            Err.Raise 7
        End If
        
        If tCMDLine.Length > 0 Then
            
            GetCommandLineOfProcess = Space$(tCMDLine.Length \ 2)
            
            If ReadProcessMemory(hProcess, tCMDLine.lpBuffer, ByVal StrPtr(GetCommandLineOfProcess), tCMDLine.Length, 0) = 0 Then
                Err.Raise 7
            End If
            
        End If
        
    Else
    
        If NtWow64QueryInformationProcess64(hProcess, ProcessBasicInformation, tPEB64, LenB(tPEB64), 0) < 0 Then
            Err.Raise 7
        End If
        
        If NtWow64ReadVirtualMemory64(hProcess, tPEB64.PebBaseAddress + 0.0032@, pParam64, Len(pParam64), 0, 0) < 0 Then
            Err.Raise 7
        End If
        
        If NtWow64ReadVirtualMemory64(hProcess, pParam64 + 0.0112@, tCMDLine64, Len(tCMDLine64), 0, 0) < 0 Then
            Err.Raise 7
        End If
        
        If tCMDLine64.Length > 0 Then
            
            GetCommandLineOfProcess = Space$(tCMDLine64.Length \ 2)
            
            If NtWow64ReadVirtualMemory64(hProcess, tCMDLine64.lpBuffer, ByVal StrPtr(GetCommandLineOfProcess), tCMDLine64.Length, 0, 0) < 0 Then
                Err.Raise 7
            End If
            
        End If
        
    End If
     
End Function

Public Function Is32BitProcess( _
                ByVal hProcess As OLE_HANDLE) As Boolean
    Static s_bHasWow64      As Boolean
    Static s_bInit          As Boolean
    Dim lWow64  As Long
    
    If Not s_bInit Then
        
        s_bHasWow64 = GetProcAddress(GetModuleHandle(StrPtr("kernel32")), "IsWow64Process")
        s_bInit = True
        
    End If
                    
    If s_bHasWow64 Then
    
        If IsWow64Process(hProcess, lWow64) = 0 Then
            Err.Raise 7
        End If
        
        Is32BitProcess = lWow64
        
    Else
        Is32BitProcess = True
    End If
                    
End Function

Public Function SetPrivilege( _
                ByVal sPrivilege As String, _
                ByVal bEnable As Boolean) As Boolean
    Dim bPriv(15)   As Byte
    Dim hToken      As OLE_HANDLE
    
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, hToken) = 0 Then
        Exit Function
    End If
    
    If LookupPrivilegeValue(0, StrPtr(sPrivilege), bPriv(4)) = 0 Then
        CloseHandle hToken
        Exit Function
    End If
    
    bPriv(0) = 1
    
    If bEnable Then
        bPriv(12) = SE_PRIVILEGE_ENABLED
    End If
  
    If AdjustTokenPrivileges(hToken, 0, bPriv(0), 0, ByVal 0&, 0) = 0 Then
        CloseHandle hToken
        Exit Function
    End If
                
    CloseHandle hToken
                
    SetPrivilege = True
    
End Function