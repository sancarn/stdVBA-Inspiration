Attribute VB_Name = "modInject"

' //
' // modInject.bas - module for injecting ActiveX Dlls to process
' // By The trick 2018-2021
' //

Option Explicit

Private Const ThreadBasicInformation      As Long = 0
Private Const PROCESS_CREATE_THREAD       As Long = &H2
Private Const PROCESS_VM_WRITE            As Long = &H20
Private Const PROCESS_VM_OPERATION        As Long = &H8
Private Const PROCESS_QUERY_INFORMATION   As Long = &H400
Private Const PROCESS_VM_READ             As Long = &H10
Private Const THREAD_QUERY_INFORMATION    As Long = &H40
Private Const SYNCHRONIZE                 As Long = &H100000
Private Const MEM_RESERVE                 As Long = &H2000&
Private Const MEM_COMMIT                  As Long = &H1000&
Private Const MEM_RELEASE                 As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE      As Long = &H40&
Private Const PAGE_READWRITE              As Long = 4&
Private Const INFINITE                    As Long = -1&
Private Const WAIT_OBJECT_0               As Long = 0
Private Const DUPLICATE_SAME_ACCESS       As Long = 2
Private Const DONT_RESOLVE_DLL_REFERENCES As Long = 1
Private Const GMEM_MOVEABLE               As Long = &H2
Private Const WM_USER                     As Long = &H400
Private Const WM_CREATEOBJECT             As Long = WM_USER + 1
Private Const WM_DESTROYME                As Long = WM_USER + 2

Private Enum eAPITable

    e_pfnRegisterClassExW
    e_pfnUnregisterClassW
    e_pfnCreateEventW
    e_pfnSetWindowsHookExW
    e_pfnPostThreadMessageW
    e_pfnWaitForSingleObject
    e_pfnCloseHandle
    e_pfnUnhookWindowsHookEx
    e_pfnCallNextHookEx
    e_pfnCreateWindowExW
    e_pfnDestroyWindow
    e_pfnSetEvent
    e_pfnPostMessageW
    e_pfnLoadLibraryW
    e_pfnFreeLibrary
    e_pfnGetProcAddress
    e_pfnSetWindowLongW
    e_pfnGetWindowLongW
    e_pfnHeapAlloc
    e_pfnHeapReAlloc
    e_pfnHeapFree
    e_pfnGetProcessHeap
    e_pfnGlobalFree
    e_pfnGlobalSize
    e_pfnGlobalLock
    e_pfnGlobalUnlock
    e_pfnSetTimer
    e_pfnKillTimer
    
    ' // Filled by shellcode
    e_pfnCoInitialize
    e_pfnCoUninitialize
    e_pfnCoMarshalInterface
    e_pfnCreateStreamOnHGlobal
    e_pfnGetHGlobalFromStream
    
    e_pfn_COUNT
    
End Enum

Private Type UUID
    Data1                   As Long
    Data2                   As Integer
    Data3                   As Integer
    Data4(0 To 7)           As Byte
End Type

Private Type tAPITable
    pfn(e_pfn_COUNT - 1)    As Long
End Type

Private Type tProcessData
    tAPITable               As tAPITable
    dwClassAtom             As Long
    dwNumOfWindows          As Long
    dwTimerID               As Long
End Type

Private Type tThreadData
    dwDestThreadID          As Long
    hEvent                  As Long
    hHook                   As Long
    hWnd                    As Long
    pszDllName              As Long
    clsid                   As UUID
    iid                     As UUID
    hr                      As Long
End Type

Private Type tObjectDesc
    hLibrary                As Long
    pObject                 As Long
    pFactory                As Long
    hMem                    As Long
    dwStmDataSize           As Long
    pStmData                As Long
End Type

Private Type tShellCodeData
    pProcessData            As Long
    pThreadData             As Long
End Type

Private Type tThreadDesc
    lThreadID               As Long
    pShellCode              As Long
    pThreadParams           As Long
    hWnd                    As Long
End Type

Private Type tProcessDesc
    lProcessID              As Long
    hProcess                As Long
    pProcessData            As Long
    lThreadsCount           As Long
    tThreads()              As tThreadDesc
End Type

Private Type tThreadEntry
    lProcessIndex           As Long
    lThreadIndex            As Long
End Type

Private Type CLIENT_ID
    UniqueProcess           As Long
    UniqueThread            As Long
End Type

Private Type THREAD_BASIC_INFORMATION
    ExitStatus              As Long
    TebBaseAddress          As Long
    ClientId                As CLIENT_ID
    AffinityMask            As Long
    Priority                As Long
    BasePriority            As Long
End Type

Private Type ANSI_STRING
    Length                  As Integer
    MaximumLength           As Integer
    Buffer                  As Long
End Type

Private Type MODULEINFO
    lpBaseOfDll             As Long
    SizeOfImage             As Long
    EntryPoint              As Long
End Type

Private Declare Function NtQueryInformationThread Lib "ntdll" ( _
                         ByVal ThreadHandle As Long, _
                         ByVal ThreadInformationClass As Long, _
                         ByRef ThreadInformation As Any, _
                         ByVal ThreadInformationLength As Long, _
                         ByRef ReturnLength As Long) As Long
Private Declare Function LdrGetProcedureAddress Lib "ntdll" ( _
                         ByVal ModuleHandle As Long, _
                         ByRef FunctionName As Any, _
                         ByVal Oridinal As Integer, _
                         ByRef FunctionAddress As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwProcessId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32.dll" ( _
                         ByVal hProcess As Long, _
                         ByRef lpAddress As Any, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32.dll" ( _
                         ByVal hProcess As Long, _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
Private Declare Function DuplicateHandle Lib "kernel32" ( _
                         ByVal hSourceProcessHandle As Long, _
                         ByVal hSourceHandle As Long, _
                         ByVal hTargetProcessHandle As Long, _
                         ByRef lpTargetHandle As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function ReadProcessMemory Lib "kernel32" ( _
                         ByVal hProcess As Long, _
                         ByVal lpBaseAddress As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nSize As Long, _
                         ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" ( _
                         ByVal hProcess As Long, _
                         ByRef lpThreadAttributes As Any, _
                         ByVal dwStackSize As Long, _
                         ByVal lpStartAddress As Long, _
                         ByRef lpParameter As Any, _
                         ByVal dwCreationFlags As Long, _
                         ByRef lpThreadId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function GetProcessId Lib "kernel32" ( _
                         ByVal hProcess As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" ( _
                         ByVal hProcess As Long, _
                         ByVal lpBaseAddress As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nSize As Long, _
                         ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
                         ByVal hHandle As Long, _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenThread Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwThreadId As Long) As Long
Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageW" ( _
                         ByVal hWnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         ByRef lParam As Any) As Long
Private Declare Function Sleep Lib "kernel32" ( _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" _
                         Alias "LoadLibraryExW" ( _
                         ByVal lpLibFileName As Long, _
                         ByVal hFile As Long, _
                         ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
                         ByVal hLibModule As Long) As Long
Private Declare Function GetModuleInformation Lib "psapi" ( _
                         ByVal hProcess As Long, _
                         ByVal hModule As Long, _
                         ByRef lpmodinfo As MODULEINFO, _
                         ByVal cb As Long) As Long
Private Declare Function lstrcmpA Lib "kernel32" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
                         ByVal hGlobal As Long, _
                         ByVal fDeleteOnRelease As Long, _
                         ByRef ppstm As IUnknown) As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
                         ByVal wFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" ( _
                         ByVal hMem As Long) As Long
Private Declare Function CoUnmarshalInterface Lib "ole32" ( _
                         ByVal pStm As IUnknown, _
                         ByRef riid As UUID, _
                         ByRef ppv As Any) As Long

Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" ( _
                    ByRef Addr As Any, _
                    ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" ( _
                    ByRef Addr As Any, _
                    ByVal NewVal As Currency)
Private Declare Sub GetMem4 Lib "msvbvm60" ( _
                    ByRef Addr As Any, _
                    ByRef Dst As Any)

Private m_lProcessCount As Long
Private m_tProcesses()  As tProcessDesc

Public Function InitializeInjectLibrary() As Boolean
    InitializeInjectLibrary = True
End Function

Public Sub UninitializeInjectLibrary()
    Dim lPIndex         As Long
    Dim lTIndex         As Long
    Dim tProcessData    As tProcessData
    
    For lPIndex = 0 To m_lProcessCount - 1
        
        ' // DEstroy all thread windows
        For lTIndex = 0 To m_tProcesses(lPIndex).lThreadsCount - 1
            SendMessage m_tProcesses(lPIndex).tThreads(lTIndex).hWnd, WM_DESTROYME, 0, ByVal 0&
        Next
        
        ' // Destroy process data
        Do
                
            Sleep 15
              
            If ReadProcessMemory(m_tProcesses(lPIndex).hProcess, m_tProcesses(lPIndex).pProcessData, _
                                 tProcessData, Len(tProcessData), 0) Then
                If tProcessData.dwTimerID = 0 Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
                    
        Loop While True
        
        VirtualFreeEx m_tProcesses(lPIndex).hProcess, m_tProcesses(lPIndex).pProcessData, 0, MEM_RELEASE
        
    Next
    
    Sleep 100
    
    ' // Destroy thread data
    For lPIndex = 0 To m_lProcessCount - 1
        For lTIndex = 0 To m_tProcesses(lPIndex).lThreadsCount - 1
            VirtualFreeEx m_tProcesses(lPIndex).hProcess, m_tProcesses(lPIndex).tThreads(lTIndex).pShellCode, 0, MEM_RELEASE
        Next
    Next
    
    m_lProcessCount = 0
    
End Sub

Public Function CreateVBObjectInThread( _
                ByVal lThreadID As Long, _
                ByRef sDllName As String, _
                ByRef sObjectName As String) As Object
    Dim tClsID  As UUID
    
    If Not GetVBClassIDFromName(sDllName, sObjectName, tClsID) Then
       MsgBox "GetVBClassIDFromName failed", vbCritical
       Exit Function
    End If
    
    Set CreateVBObjectInThread = CreateObjectInThread(lThreadID, sDllName, tClsID, IID_IDispatch)
    
End Function

Public Function CreateObjectInThread( _
                ByVal lThreadID As Long, _
                ByRef sDllName As String, _
                ByRef tClsID As UUID, _
                ByRef tIID As UUID) As IUnknown
    Dim hThread         As Long
    Dim tThreadEntry    As tThreadEntry
    Dim tThreadParam    As tThreadData
    Dim hProcess        As Long
    Dim pDllName        As Long
    Dim pObjectDesc     As Long
    Dim tObjectDesc     As tObjectDesc
    Dim hMem            As Long
    Dim pStmData        As Long
    Dim cStmMarshal     As IUnknown
    Dim pResult         As IUnknown

    tThreadEntry = FindThreadEntry(lThreadID)
    
    If tThreadEntry.lThreadIndex = -1 Then
        
        hThread = OpenThread(THREAD_QUERY_INFORMATION, 0, lThreadID)
        
        If hThread = 0 Then
            MsgBox "OpenThread failed", vbCritical
            GoTo CleanUp
        End If
        
        tThreadEntry = InitializeThreadData(hThread)
        
        If tThreadEntry.lThreadIndex = -1 Then
            MsgBox "SetupThread failed", vbCritical
            GoTo CleanUp
        End If
        
    End If
    
    With m_tProcesses(tThreadEntry.lProcessIndex).tThreads(tThreadEntry.lThreadIndex)
        
        hProcess = m_tProcesses(tThreadEntry.lProcessIndex).hProcess
        
        If ReadProcessMemory(hProcess, .pThreadParams, tThreadParam, Len(tThreadParam), 0) = 0 Then
            MsgBox "ReadProcessMemory failed", vbCritical
            GoTo CleanUp
        End If
        
        pDllName = VirtualAllocEx(hProcess, ByVal 0&, LenB(sDllName) + 2, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
        
        If pDllName = 0 Then
            MsgBox "VirtualAllocEx failed", vbCritical
            GoTo CleanUp
        End If
        
        If WriteProcessMemory(hProcess, pDllName, ByVal StrPtr(sDllName), LenB(sDllName) + 2, 0) = 0 Then
            MsgBox "WriteProcessMemory failed", vbCritical
            GoTo CleanUp
        End If
        
        tThreadParam.clsid = tClsID
        tThreadParam.iid = tIID
        tThreadParam.pszDllName = pDllName
        
        If WriteProcessMemory(hProcess, .pThreadParams, tThreadParam, Len(tThreadParam), 0) = 0 Then
            MsgBox "WriteProcessMemory failed", vbCritical
            GoTo CleanUp
        End If
        
        pObjectDesc = SendMessage(.hWnd, WM_CREATEOBJECT, 0, ByVal 0&)
        
        If ReadProcessMemory(hProcess, .pThreadParams, tThreadParam, Len(tThreadParam), 0) = 0 Then
            MsgBox "ReadProcessMemory failed", vbCritical
            GoTo CleanUp
        End If
        
        If pObjectDesc = 0 Then
            MsgBox "Object creation failed 0x" & Hex$(tThreadParam.hr), vbCritical
            GoTo CleanUp
        End If
        
        If ReadProcessMemory(hProcess, pObjectDesc, tObjectDesc, Len(tObjectDesc), 0) = 0 Then
            MsgBox "ReadProcessMemory failed", vbCritical
            GoTo CleanUp
        End If
        
        hMem = GlobalAlloc(GMEM_MOVEABLE, tObjectDesc.dwStmDataSize)
        If hMem = 0 Then
            MsgBox "GlobalAlloc failed", vbCritical
            GoTo CleanUp
        End If
        
        pStmData = GlobalLock(hMem)
        If pStmData = 0 Then
            MsgBox "GlobalLock failed", vbCritical
            GoTo CleanUp
        End If
        
        If ReadProcessMemory(hProcess, tObjectDesc.pStmData, ByVal pStmData, tObjectDesc.dwStmDataSize, 0) = 0 Then
            MsgBox "ReadProcessMemory failed", vbCritical
            GoTo CleanUp
        End If
        
        GlobalUnlock hMem
        
        If CreateStreamOnHGlobal(hMem, 1, cStmMarshal) < 0 Then
            MsgBox "CreateStreamOnHGlobal failed", vbCritical
            GoTo CleanUp
        End If
        
        If CoUnmarshalInterface(cStmMarshal, tIID, pResult) < 0 Then
            MsgBox "CoUnmarshalInterface failed", vbCritical
            GoTo CleanUp
        End If
        
    End With
    
    Set CreateObjectInThread = pResult
    
CleanUp:
    
    If cStmMarshal Is Nothing Then
        If hMem Then
            GlobalFree hMem
        End If
    End If
    
    If pDllName Then
        VirtualFreeEx hProcess, pDllName, 0, MEM_RELEASE
    End If
    
    If hThread Then
        CloseHandle hThread
    End If
    
End Function

Private Function GetVBClassIDFromName( _
                 ByRef sDllName As String, _
                 ByRef sClassName As String, _
                 ByRef tClsID As UUID) As Boolean
    Dim hLib        As Long
    Dim pfn         As Long
    Dim pVbHdr      As Long
    Dim lSignature  As Long
    Dim tModInfo    As MODULEINFO
    Dim pCOMData    As Long
    Dim lOffset     As Long
    Dim lOffsetName As Long
    Dim bNameANSI() As Byte
    
    If Len(sClassName) = 0 Then Exit Function
    
    hLib = LoadLibraryEx(StrPtr(sDllName), 0, DONT_RESOLVE_DLL_REFERENCES)
    
    If hLib = 0 Then
        MsgBox "Unable to load DLL", vbCritical
        GoTo CleanUp
    End If
    
    If GetModuleInformation(GetCurrentProcess(), hLib, tModInfo, Len(tModInfo)) = 0 Then
        MsgBox "GetModuleInformation failed", vbCritical
        GoTo CleanUp
    End If
    
    pfn = GetProcAddress(hLib, "DllGetClassObject")
    
    If pfn = 0 Then
        MsgBox "Invalid module", vbCritical
        GoTo CleanUp
    End If
    
    GetMem4 ByVal pfn + 2, pVbHdr
    
    If pVbHdr < hLib Or pVbHdr >= hLib + tModInfo.SizeOfImage Then
        MsgBox "It isn't VB AX Dll", vbCritical
        GoTo CleanUp
    End If
    
    GetMem4 ByVal pVbHdr, lSignature
    
    If lSignature <> &H21354256 Then
        MsgBox "VB5! signature not found", vbCritical
        GoTo CleanUp
    End If
    
    bNameANSI = StrConv(sClassName & vbNullChar, vbFromUnicode)
    
    GetMem4 ByVal pVbHdr + &H54, pCOMData
    GetMem4 ByVal pCOMData, lOffset
    
    Do While lOffset
        
        GetMem4 ByVal pCOMData + lOffset + 4, lOffsetName
        
        If lstrcmpA(ByVal pCOMData + lOffsetName, bNameANSI(0)) = 0 Then
        
            memcpy tClsID, ByVal pCOMData + lOffset + &H14, Len(tClsID)
            GetVBClassIDFromName = True
            Exit Do
            
        End If
        
        GetMem4 ByVal pCOMData + lOffset, lOffset
        
    Loop
    
CleanUp:
    
    If hLib Then
        FreeLibrary hLib
    End If
    
End Function

Private Function InitializeProcessData( _
                 ByVal hProcess As Long) As Long
    Static s_tAPITable      As tAPITable
    
    Dim tProcessData    As tProcessData
    Dim pProcessData    As Long
    Dim lIndex          As Long
    Dim hDupHandle      As Long
    
    InitializeProcessData = -1
    
    pProcessData = VirtualAllocEx(hProcess, ByVal 0&, Len(tProcessData), MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
                     
    If pProcessData = 0 Then
        MsgBox "VirtualAllocEx failed", vbCritical
        GoTo CleanUp
    End If
    
    If s_tAPITable.pfn(0) = 0 Then
        FillApiTable s_tAPITable
    End If
    
    tProcessData.tAPITable = s_tAPITable
    
    If WriteProcessMemory(hProcess, pProcessData, tProcessData, Len(tProcessData), 0) = 0 Then
        MsgBox "WriteProcessMemory failed", vbCritical
        GoTo CleanUp
    End If
    
    If DuplicateHandle(GetCurrentProcess(), hProcess, GetCurrentProcess(), hDupHandle, 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then
        MsgBox "DuplicateHandle failed", vbCritical
        GoTo CleanUp
    End If
    
    lIndex = m_lProcessCount
    
    If lIndex Then
        If lIndex > UBound(m_tProcesses) Then
            ReDim Preserve m_tProcesses(lIndex + 5)
        End If
    Else
        ReDim m_tProcesses(4)
    End If
    
    With m_tProcesses(lIndex)
        .hProcess = hDupHandle
        .lProcessID = GetProcessId(hProcess)
        .pProcessData = pProcessData
    End With
    
    m_lProcessCount = m_lProcessCount + 1
    
    InitializeProcessData = lIndex
    
CleanUp:
    
    If InitializeProcessData = -1 Then
    
        If hDupHandle Then
            CloseHandle hDupHandle
        End If
        
        If pProcessData Then
            VirtualFreeEx hProcess, ByVal pProcessData, 0, MEM_RELEASE
        End If
        
    End If
    
End Function

Private Function InitializeThreadData( _
                 ByVal hThread As Long) As tThreadEntry
    Static s_cShellcode()   As Currency
    Static s_bShellInit     As Boolean
    
    Dim hProcess        As Long
    Dim lProcessID      As Long
    Dim lProcDescIndex  As Long
    Dim pShellCode      As Long
    Dim lSizeAllData    As Long
    Dim tThreadData     As tThreadData
    Dim bAllData()      As Byte
    Dim lIndex          As Long
    Dim hNewThread      As Long
    
    InitializeThreadData.lProcessIndex = -1
    InitializeThreadData.lThreadIndex = -1
    
    lProcessID = ProcessIDFromThreadHandle(hThread)
    
    hProcess = OpenProcess(PROCESS_CREATE_THREAD Or PROCESS_VM_WRITE Or SYNCHRONIZE Or _
                           PROCESS_VM_OPERATION Or PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcessID)
    
    If hProcess = 0 Then
        MsgBox "OpenProcess failed", vbCritical
        GoTo CleanUp
    End If
    
    lProcDescIndex = FindProcessDesc(lProcessID)
    
    If lProcDescIndex = -1 Then
        
        lProcDescIndex = InitializeProcessData(hProcess)
        
        If lProcDescIndex = -1 Then
            MsgBox "Unable to initialize process data", vbCritical
            GoTo CleanUp
        End If
        
    End If
    
    If Not s_bShellInit Then
        FillShellCode s_cShellcode
        s_bShellInit = True
    End If
    
    With tThreadData
        .dwDestThreadID = ThreadIDFromThreadHandle(hThread)
    End With
    
    lSizeAllData = (UBound(s_cShellcode) + 1) * 8 + Len(tThreadData)
    
    ReDim bAllData(lSizeAllData - 1)
    
    memcpy bAllData(lIndex), s_cShellcode(0), (UBound(s_cShellcode) + 1) * 8
    lIndex = lIndex + (UBound(s_cShellcode) + 1) * 8

    memcpy bAllData(lIndex), tThreadData, Len(tThreadData)
    
    pShellCode = VirtualAllocEx(hProcess, ByVal 0&, lSizeAllData, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    
    If pShellCode = 0 Then
        MsgBox "VirtualAllocEx failed", vbCritical
        GoTo CleanUp
    End If
    
    ' // Update params
    PutMem4 bAllData(0), m_tProcesses(lProcDescIndex).pProcessData
    PutMem4 bAllData(4), pShellCode + (UBound(s_cShellcode) + 1) * 8
    
    If WriteProcessMemory(hProcess, pShellCode, bAllData(0), UBound(bAllData) + 1, 0) = 0 Then
        MsgBox "WriteProcessMemory failed", vbCritical
        GoTo CleanUp
    End If
    
    hNewThread = CreateRemoteThread(hProcess, ByVal 0&, 0, pShellCode + 8, ByVal pShellCode, 0, 0)
    
    If hNewThread = 0 Then
        MsgBox "CreateRemoteThread failed", vbCritical
        GoTo CleanUp
    End If
    
    If WaitForSingleObject(hNewThread, INFINITE) <> WAIT_OBJECT_0 Then
        MsgBox "WaitForSingleObject failed", vbCritical
        GoTo CleanUp
    End If
    
    ' // Check HRESULT
    If ReadProcessMemory(hProcess, pShellCode + (UBound(s_cShellcode) + 1) * 8, tThreadData, Len(tThreadData), 0) = 0 Then
        MsgBox "ReadProcessMemory failed", vbCritical
        GoTo CleanUp
    End If
    
    If tThreadData.hr < 0 Then
        MsgBox "Remote code failed 0x" & Hex$(tThreadData.hr), vbCritical
        GoTo CleanUp
    End If
    
    With m_tProcesses(lProcDescIndex)
        
        lIndex = .lThreadsCount
        
        If lIndex Then
            If lIndex > UBound(.tThreads) Then
                ReDim Preserve .tThreads(lIndex + 5)
            End If
        Else
            ReDim .tThreads(4)
        End If
        
        With .tThreads(lIndex)
            .lThreadID = ThreadIDFromThreadHandle(hThread)
            .pShellCode = pShellCode
            .pThreadParams = pShellCode + (UBound(s_cShellcode) + 1) * 8
            .hWnd = tThreadData.hWnd
        End With
        
        .lThreadsCount = .lThreadsCount + 1
                
        InitializeThreadData.lProcessIndex = lProcDescIndex
        InitializeThreadData.lThreadIndex = lIndex
        
    End With
    
CleanUp:
    
    If InitializeThreadData.lThreadIndex = -1 Then
        
        If pShellCode Then
            VirtualFreeEx hProcess, ByVal pShellCode, 0, MEM_RELEASE
        End If

    End If
    
    If hNewThread Then
        CloseHandle hNewThread
    End If
    
    If hProcess Then
        CloseHandle hProcess
    End If
    
End Function

Private Function FindThreadEntry( _
                 ByVal lThreadID As Long) As tThreadEntry
    Dim lPIndex As Long
    Dim lTIndex As Long
    
    FindThreadEntry.lProcessIndex = -1
    FindThreadEntry.lThreadIndex = -1
    
    For lPIndex = 0 To m_lProcessCount - 1
        For lTIndex = 0 To m_tProcesses(lPIndex).lThreadsCount - 1
            If m_tProcesses(lPIndex).tThreads(lTIndex).lThreadID = lThreadID Then
                
                FindThreadEntry.lProcessIndex = lPIndex
                FindThreadEntry.lThreadIndex = lTIndex
                Exit Function
                
            End If
        Next
    Next
    
End Function

Private Function FindProcessDesc( _
                 ByVal lProcessID As Long) As Long
    Dim lIndex  As Long
    
    FindProcessDesc = -1
    
    For lIndex = 0 To m_lProcessCount - 1
        If m_tProcesses(lIndex).lProcessID = lProcessID Then
            FindProcessDesc = lIndex
            Exit Function
        End If
    Next
    
End Function

Private Sub FillShellCode( _
            ByRef c() As Currency)
    
    ReDim c(203)
        
    c(0) = 0@:                       c(1) = 622149520261890.2869@:    c(2) = -826100909148098.3469@:   c(3) = 1509144460.0957@
    c(4) = 450695363413330.816@:     c(5) = 1419196474.7403@:         c(6) = 367504834842069.7856@:    c(7) = 377176358841117.5424@
    c(8) = 482785880077452.9929@:    c(9) = 780963917754310.8975@:    c(10) = 434822469720352.2153@:   c(11) = 6832418.1897@
    c(12) = 838835718005597.3699@:   c(13) = 598089188157714.4681@:   c(14) = 143452433191143.6287@:   c(15) = 824195393332550.0416@
    c(16) = 731058935675914.8659@:   c(17) = -4869473551191.899@:     c(18) = 644805278153.0199@:      c(19) = 731057518346574.2336@
    c(20) = 794869213241346.9779@:   c(21) = 3050637288880.724@:      c(22) = -169440070504965.7517@:  c(23) = 522090922364633.0901@
    c(24) = 823438813421262.7527@:   c(25) = 788169236512738.0335@:   c(26) = -918522202273144.3456@:  c(27) = 3736526781415.424@
    c(28) = -343912361782306.4064@:  c(29) = 17606650205778.744@:     c(30) = 299139053752320@:        c(31) = -681705095839717.785@
    c(32) = -857253092495510.728@:   c(33) = 975933642.4517@:         c(34) = 735662924748066.3296@:   c(35) = 37263153809.0117@
    c(36) = -525585588559816.064@:   c(37) = 763810497024383.6352@:   c(38) = -7194104402697.1648@:    c(39) = 352407820927480.8407@
    c(40) = 583081162011207.2705@:   c(41) = 47232421437.6073@:       c(42) = 24609929138058.0096@:    c(43) = -172643803741681.4593@
    c(44) = -149849522178306.8401@:  c(45) = 763822188575893.3295@:   c(46) = -457423187639487.7184@:  c(47) = -857485368640457.6651@
    c(48) = -7080085619138.8602@:    c(49) = -224052931805439.0697@:  c(50) = 6620409707783.7824@:     c(51) = 857458700740463.7827@
    c(52) = 232962080165.044@:       c(53) = 835868584980984.6272@:   c(54) = -410043070791765.2211@:  c(55) = 679902806244589.6774@
    c(56) = -140402628079728.8097@:  c(57) = 99783064.5252@:          c(58) = 200111091.584@:          c(59) = 61178917953386.7915@
    c(60) = 850253272048228.48@:     c(61) = 634078725576943.182@:    c(62) = -6967495240308.912@:     c(63) = 3554719.2407@
    c(64) = -53946400711164.4672@:   c(65) = 489406289761936.9216@:   c(66) = -840259100466167.5123@:  c(67) = 59969453250997.0453@
    c(68) = 770943468714446.0394@:   c(69) = 763822151803131.7504@:   c(70) = 634078725698646.272@:    c(71) = 96989297243734.25@
    c(72) = 376684877143501.5912@:   c(73) = -860863175751025.6641@:  c(74) = 634079170949178.9824@:   c(75) = 850253102821130.526@
    c(76) = 764048650612021.6336@:   c(77) = -393660417928593.4336@:  c(78) = 2708144652.6146@:        c(79) = 7205665431673.5232@
    c(80) = 617992388123558.2208@:   c(81) = 88426609189531.7897@:    c(82) = -147679914724268.4184@:  c(83) = 832581380431098.7413@
    c(84) = 75350851266289.5876@:    c(85) = 5655118155140.3448@:     c(86) = 108086391275395.4816@:   c(87) = -7194104580534.3099@
    c(88) = 841287063250868.8501@:   c(89) = -7081430607597.8136@:    c(90) = 625283605927762.5461@:   c(91) = 583638424043836.3787@
    c(92) = -840822004889534.1816@:  c(93) = 576693688593606.2208@:   c(94) = 475872204364.1995@:      c(95) = 795371254986493.1328@
    c(96) = 802045835603741.4485@:   c(97) = -5507234178433.0121@:    c(98) = 396289219719492.5008@:   c(99) = -898606139060359.3729@
    c(100) = 2993131228581.0939@:    c(101) = 657985580966427.6479@:  c(102) = 20596.515@:             c(103) = 15015204751.5647@
    c(104) = 98682693789.0048@:      c(105) = 2983680282160.7424@:    c(106) = 15451375379.0463@:      c(107) = 440336698.0352@
    c(108) = -151320947298611.8283@: c(109) = 60939332464582.5809@:   c(110) = 21030616600887.0399@:   c(111) = 120774906153086.1617@
    c(112) = -170592380311349.9392@: c(113) = -829239206103062.9376@: c(114) = 32075308632480.1539@:   c(115) = 465081104019335.6132@
    c(116) = 576460754987014.8872@:  c(117) = 3730398576692.4113@:    c(118) = -518842824570752.2048@: c(119) = 634078725562748.1228@
    c(120) = 235770.8652@:           c(121) = -441089059639105.9456@: c(122) = 266064650174391.912@:   c(123) = 24219225498.1974@
    c(124) = 96324334.3872@:         c(125) = 382977.8176@:           c(126) = -410728286014667.0137@: c(127) = 502798750898250.8613@
    c(128) = -84505055224817.2304@:  c(129) = 7100841513189.376@:     c(130) = 389231.4112@:           c(131) = -843073848483961.2069@
    c(132) = -6742315674713.8245@:   c(133) = 467700301241654.9975@:  c(134) = -757246382087916.749@:  c(135) = 136258107848890.7777@
    c(136) = 729792093343029.6576@:  c(137) = 708400767826999.794@:   c(138) = 634087476324384.0874@:  c(139) = 10355752611247.0332@
    c(140) = 680127450.3168@:        c(141) = 146551410917797.6205@:  c(142) = 61164619175565.7042@:   c(143) = 145131649.6009@
    c(144) = 3921298878105.3323@:    c(145) = 61164619174302.9503@:   c(146) = 134823728.0905@:        c(147) = -826329131016780.0437@
    c(148) = 3921298447677.1414@:    c(149) = 61164619174276.7359@:   c(150) = 121079832.7433@:        c(151) = 2983680418882.0877@
    c(152) = 61164619175010.9183@:   c(153) = 110771911.2329@:        c(154) = -662565997272.5363@:    c(155) = 6994983024751.4263@
    c(156) = 6527713035111.8461@:    c(157) = 2983680280480.9728@:    c(158) = -4893669178130.6881@:   c(159) = 6995034711994.5845@
    c(160) = 5626993109637.7469@:    c(161) = 665406836552020.7872@:  c(162) = -74447222539.8651@:     c(163) = 19451565090.6623@
    c(164) = -549975735890.5088@:    c(165) = 467700301241656.1239@:  c(166) = -757246382087916.7491@: c(167) = 31833090512492.9536@
    c(168) = -457417470179302.0161@: c(169) = 3005134215755.6084@:    c(170) = 772825210529962.1631@:  c(171) = 2992797969449.5936@
    c(172) = -855336483820743.2705@: c(173) = 684096904056825.2485@:  c(174) = -7194094094680.5925@:   c(175) = 501061560320919.6631@
    c(176) = 376691698178618.316@:   c(177) = 616409519373097.3703@:  c(178) = 580342065090006.4732@:  c(179) = -842839814951943.91@
    c(180) = -112884324657935.6587@: c(181) = 580343824308635.6617@:  c(182) = 770804050455412.968@:   c(183) = 765621611946724.5316@
    c(184) = 463092306978111.4624@:  c(185) = 634080050558483.9248@:  c(186) = -842807768039781.088@:  c(187) = 59925912824269.4213@
    c(188) = 904414444747361.4211@:  c(189) = -549976134883.714@:     c(190) = 68117042499809.0839@:   c(191) = 583638424045304.3595@
    c(192) = 764386201466971.828@:   c(193) = -7190441456162.944@:    c(194) = 3921394593241.9152@:    c(195) = 835892334002702.3615@
    c(196) = -841487914731477.4266@: c(197) = 12716662166047.8533@:   c(198) = -461168601842738.7904@: c(199) = 621665633563154.8416@
    c(200) = 763822591552417.024@:   c(201) = 482798634555096.3968@:  c(202) = 828674975982632.448@:   c(203) = 0@

End Sub

Private Sub FillApiTable( _
            ByRef tTable As tAPITable)
    Dim hKernel32   As Long
    Dim hUser32     As Long
    Dim hLib        As Long
    Dim sFuncName   As String
    Dim lIndex      As Long
    Dim tAnsiString As ANSI_STRING
    Dim bANSI()     As Byte
    
    hKernel32 = GetModuleHandle(StrPtr("kernel32"))
    hUser32 = GetModuleHandle(StrPtr("user32"))
    
    For lIndex = 0 To e_pfn_COUNT - 1
        
        Select Case lIndex
        Case e_pfnRegisterClassExW: hLib = hUser32: sFuncName = "RegisterClassExW"
        Case e_pfnUnregisterClassW: hLib = hUser32: sFuncName = "UnregisterClassW"
        Case e_pfnCreateEventW: hLib = hKernel32: sFuncName = "CreateEventW"
        Case e_pfnSetWindowsHookExW: hLib = hUser32: sFuncName = "SetWindowsHookExW"
        Case e_pfnPostThreadMessageW: hLib = hUser32: sFuncName = "PostThreadMessageW"
        Case e_pfnWaitForSingleObject: hLib = hKernel32: sFuncName = "WaitForSingleObject"
        Case e_pfnCloseHandle: hLib = hKernel32: sFuncName = "CloseHandle"
        Case e_pfnUnhookWindowsHookEx: hLib = hUser32: sFuncName = "UnhookWindowsHookEx"
        Case e_pfnCallNextHookEx: hLib = hUser32: sFuncName = "CallNextHookEx"
        Case e_pfnCreateWindowExW: hLib = hUser32: sFuncName = "CreateWindowExW"
        Case e_pfnDestroyWindow: hLib = hUser32: sFuncName = "DestroyWindow"
        Case e_pfnSetEvent: hLib = hKernel32: sFuncName = "SetEvent"
        Case e_pfnPostMessageW: hLib = hUser32: sFuncName = "PostMessageW"
        Case e_pfnLoadLibraryW: hLib = hKernel32: sFuncName = "LoadLibraryW"
        Case e_pfnFreeLibrary: hLib = hKernel32: sFuncName = "FreeLibrary"
        Case e_pfnGetProcAddress: hLib = hKernel32: sFuncName = "GetProcAddress"
        Case e_pfnSetWindowLongW: hLib = hUser32: sFuncName = "SetWindowLongW"
        Case e_pfnGetWindowLongW: hLib = hUser32: sFuncName = "GetWindowLongW"
        Case e_pfnHeapAlloc: hLib = hKernel32: sFuncName = "HeapAlloc"
        Case e_pfnHeapReAlloc: hLib = hKernel32: sFuncName = "HeapReAlloc"
        Case e_pfnHeapFree: hLib = hKernel32: sFuncName = "HeapFree"
        Case e_pfnGetProcessHeap: hLib = hKernel32: sFuncName = "GetProcessHeap"
        Case e_pfnGlobalFree: hLib = hKernel32: sFuncName = "GlobalFree"
        Case e_pfnGlobalSize: hLib = hKernel32: sFuncName = "GlobalSize"
        Case e_pfnGlobalLock: hLib = hKernel32: sFuncName = "GlobalLock"
        Case e_pfnGlobalUnlock: hLib = hKernel32: sFuncName = "GlobalUnlock"
        Case e_pfnSetTimer: hLib = hUser32: sFuncName = "SetTimer"
        Case e_pfnKillTimer: hLib = hUser32: sFuncName = "KillTimer"
        End Select
        
        bANSI = StrConv(sFuncName & vbNullChar, vbFromUnicode)
        
        tAnsiString.Length = UBound(bANSI)
        tAnsiString.MaximumLength = tAnsiString.Length + 1
        tAnsiString.Buffer = VarPtr(bANSI(0))
        
        ' // Bypass apphelp etc.
        If LdrGetProcedureAddress(hLib, tAnsiString, 0, tTable.pfn(lIndex)) < 0 Then
            Err.Raise 5, , "LdrGetProcedureAddress failed"
        End If
        
    Next

End Sub

' // Get PID from hThread
Private Function ProcessIDFromThreadHandle( _
                 ByVal hThread As Long) As Long
    Dim tTBI    As THREAD_BASIC_INFORMATION
    Dim status  As Long
    
    status = NtQueryInformationThread(hThread, ThreadBasicInformation, tTBI, Len(tTBI), 0)
    
    If status >= 0 Then
        ProcessIDFromThreadHandle = tTBI.ClientId.UniqueProcess
    End If

End Function

' // Get TID from hThread
Private Function ThreadIDFromThreadHandle( _
                 ByVal hThread As Long) As Long
    Dim tTBI    As THREAD_BASIC_INFORMATION
    Dim status  As Long
    
    status = NtQueryInformationThread(hThread, ThreadBasicInformation, tTBI, Len(tTBI), 0)
    
    If status >= 0 Then
        ThreadIDFromThreadHandle = tTBI.ClientId.UniqueThread
    End If

End Function

Private Function IID_IDispatch() As UUID
    PutMem8 IID_IDispatch, 13.2096@
    PutMem8 ByVal VarPtr(IID_IDispatch) + 8, 504403158265495.5712@
End Function
