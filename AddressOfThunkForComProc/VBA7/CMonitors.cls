VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMonitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Private state
' ============================================================================

Private mMonitorCount As Long
Private mPrimaryMonitor As LongPtr


' ============================================================================
' Types / constants
' ============================================================================

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Const MONITORINFOF_PRIMARY As Long = &H1&

Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_RESERVE As Long = &H2000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&


' ============================================================================
' API declarations
' ============================================================================

Private Declare PtrSafe Function EnumDisplayMonitors Lib "user32" ( _
    ByVal hdc As LongPtr, _
    ByVal lprcClip As LongPtr, _
    ByVal lpfnEnum As LongPtr, _
    ByVal dwData As LongPtr) As Long

Private Declare PtrSafe Function GetMonitorInfo Lib "user32.dll" Alias _
    "GetMonitorInfoA" ( _
    ByVal hMonitor As LongPtr, _
    ByRef lpmi As MONITORINFO) As Long

Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias _
    "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByRef Source As Any, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Sub CopyMemoryByAddr Lib "kernel32" Alias _
    "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias _
    "GetModuleHandleA" ( _
    ByVal lpModuleName As String) As LongPtr

Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, _
    ByVal lpProcName As String) As LongPtr

' ============================================================================
' IMPORTANT - CALLBACKS
' ============================================================================
' These two callback methods must remain the first Public members in this class
' because GetComProcAddress(Me, 0) and GetComProcAddress(Me, 1) depend on that.
' Do not move other Public methods above them unless you also adjust the indexes.
' ============================================================================

Public Function zzProtMonitorHandleEnumCallback( _
    ByVal hMonitor As LongPtr, _
    ByVal hdcMonitor As LongPtr, _
    ByVal lprcMonitor As LongPtr, _
    ByVal dwData As LongPtr) As Long

    mMonitorCount = mMonitorCount + 1
    Debug.Print "   "; CStr(hMonitor)

    zzProtMonitorHandleEnumCallback = 1&
End Function

Public Function zzProtPrimaryMonitorHandleEnumCallback( _
    ByVal hMonitor As LongPtr, _
    ByVal hdcMonitor As LongPtr, _
    ByVal lprcMonitor As LongPtr, _
    ByVal dwData As LongPtr) As Long

    Dim mi As MONITORINFO

    mi.cbSize = LenB(mi)

    If GetMonitorInfo(hMonitor, mi) <> 0 Then
        If (mi.dwFlags And MONITORINFOF_PRIMARY) <> 0 Then
            mPrimaryMonitor = hMonitor
            zzProtPrimaryMonitorHandleEnumCallback = 0&
            Exit Function
        End If
    End If

    zzProtPrimaryMonitorHandleEnumCallback = 1&
End Function


'===============================================
' Public methods
'===============================================


Public Function CountMonitors() As Long
    Dim pProc As LongPtr
    Dim pThunk As LongPtr
    Dim thunkSize As LongPtr

    mMonitorCount = 0

    pProc = GetComProcAddress(Me, 0)
    If pProc = 0 Then Err.Raise 5, , "Failed to resolve callback address."

    pThunk = AddressOfThunkForComProc(Me, pProc, 4, thunkSize)
    If pThunk = 0 Then Err.Raise 5, , "Failed to build callback thunk."

    On Error GoTo CleanFail

    Call EnumDisplayMonitors(0, 0, pThunk, 0)

    CountMonitors = mMonitorCount

CleanExit:
    If pThunk <> 0 Then
        FreeTheThunk pThunk
    End If
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Public Function GetPrimaryMonitorHandle() As LongPtr
    Dim pProc As LongPtr
    Dim pThunk As LongPtr
    Dim thunkSize As LongPtr

    mPrimaryMonitor = 0

    pProc = GetComProcAddress(Me, 1)
    If pProc = 0 Then Err.Raise 5, , "Failed to resolve callback address."

    pThunk = AddressOfThunkForComProc(Me, pProc, 4, thunkSize)
    If pThunk = 0 Then Err.Raise 5, , "Failed to build callback thunk."

    On Error GoTo CleanFail

    Call EnumDisplayMonitors(0, 0, pThunk, 0)

    GetPrimaryMonitorHandle = mPrimaryMonitor

CleanExit:
    If pThunk <> 0 Then
        FreeTheThunk pThunk
    End If
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Public Property Get MonitorCountValue() As Long
    MonitorCountValue = mMonitorCount
End Property

Public Property Get PrimaryMonitorHandleValue() As LongPtr
    PrimaryMonitorHandleValue = mPrimaryMonitor
End Property





' ============================================================================
' Resolve a COM method address from the vtable
' ============================================================================

Private Function GetComProcAddress( _
    ByVal o As Object, _
    ByVal publicMemberIndex As Long) As LongPtr

    Dim pObj As LongPtr
    Dim pVTable As LongPtr
    Dim cbPtr As LongPtr
    Dim pEntry As LongPtr

    pObj = ObjPtr(o)
    If pObj = 0 Then Exit Function

    cbPtr = LenB(pObj)

    CopyMemoryByAddr pVTable, pObj, cbPtr

    ' IUnknown = 3 entries
    ' IDispatch = 4 entries
    ' public members start after that
    pEntry = pVTable + ((7 + publicMemberIndex) * cbPtr)

    CopyMemoryByAddr GetComProcAddress, pEntry, cbPtr
End Function


' ============================================================================
' Thunk memory cleanup
' ============================================================================

Private Sub FreeTheThunk(ByVal pThunk As LongPtr)
    If pThunk <> 0 Then
        Call VirtualFree(pThunk, 0, MEM_RELEASE)
    End If
End Sub


' ============================================================================
' Thunk builder
' ============================================================================
' Converts a COM instance method:
'   HRESULT Method(this, arg1, arg2, ..., retval)
'
' into a normal callback:
'   Long Callback(arg1, arg2, ...)
'
' Restriction:
' - 1 to 4 args only
' - integer/pointer-sized args only
' - return value assumed to be Long stored in hidden retval slot
' ============================================================================

Private Function AddressOfThunkForComProc( _
    ByVal o As Object, _
    ByVal ComProcAddress As LongPtr, _
    ByVal ComProcArgCount As Long, _
    ByRef iThunkSize As LongPtr, _
    Optional ByVal AddIDEsafety As Boolean = True) As LongPtr

    ' Resolve EbMode for IDE safety (prevents callback when VBE is in break/design mode).
    ' The Debug.Print 1/0 trick detects the IDE: in the IDE it raises an error, compiled it's stripped.
    Dim lEbMode As LongPtr
    If AddIDEsafety Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            Dim hVBE As LongPtr
            hVBE = GetModuleHandle("vbe7.dll")
            If hVBE <> 0 Then
                lEbMode = GetProcAddress(hVBE, "EbMode")
            End If
        End If
        On Error GoTo 0
    End If

#If Win64 Then
    Dim pThis As LongPtr: pThis = ObjPtr(o)
    If pThis = 0 Or ComProcAddress = 0 Then Exit Function
    If ComProcArgCount < 1 Or ComProcArgCount > 4 Then Exit Function

    Dim p As Long: p = 0
    Dim bb() As Byte: ReDim bb(0 To 159)

    ' Stack layout after "sub rsp, 88":
    '
    ' [rsp+00]..[rsp+1F] = shadow space
    ' [rsp+20]           = stack arg 5
    ' [rsp+28]           = stack arg 6
    ' [rsp+30]           = hidden retval storage
    ' [rsp+38]           = saved arg1
    ' [rsp+40]           = saved arg2
    ' [rsp+48]           = saved arg3
    ' [rsp+50]           = saved arg4

    ' --- Prologue: allocate stack frame and save incoming args ---
    EmitBytes bb, p, &H48, &H83, &HEC, &H58             ' sub rsp, 0x58        ; allocate 88-byte stack frame
    EmitBytes bb, p, &H48, &H89, &H4C, &H24, &H38       ' mov [rsp+38h], rcx   ; save arg1
    EmitBytes bb, p, &H48, &H89, &H54, &H24, &H40       ' mov [rsp+40h], rdx   ; save arg2
    EmitBytes bb, p, &H4C, &H89, &H44, &H24, &H48       ' mov [rsp+48h], r8    ; save arg3
    EmitBytes bb, p, &H4C, &H89, &H4C, &H24, &H50       ' mov [rsp+50h], r9    ; save arg4

    ' --- EbMode IDE safety check ---
    ' If lEbMode is non-zero (IDE), call EbMode() and skip callback unless in run mode.
    ' If lEbMode is zero (compiled or safety disabled), the test/jz skips the check.
    EmitBytes bb, p, &H48, &HB8                          ' mov rax, <imm64>     ; rax = lEbMode
    EmitPtr bb, p, lEbMode
    EmitBytes bb, p, &H48, &H85, &HC0                    ' test rax, rax
    EmitBytes bb, p, &H74, &H7                           ' jz skip_check        ; skip if null (compiled)
    EmitBytes bb, p, &HFF, &HD0                          ' call rax             ; call EbMode()
    EmitBytes bb, p, &H83, &HF8, &H1                     ' cmp eax, 1           ; running?
    Dim jneAt As Long: jneAt = p + 1
    EmitBytes bb, p, &H75, &H0                           ' jne return_zero      ; bail if not running (patched later)

    ' --- Rearrange args for COM call: shift right to insert pThis as rcx, append pRetval ---
    ' COM signature: ComProc(pThis, origArg1, ..., origArgN, pRetval) As HRESULT
    Select Case ComProcArgCount
        Case 1
            ' COM call: ComProc(pThis=rcx, arg1=rdx, pRetval=r8)
            EmitBytes bb, p, &H4C, &H8D, &H44, &H24, &H30   ' lea r8, [rsp+30h]    ; r8  = pRetval (-> retval storage)
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38   ' mov rdx, [rsp+38h]   ; rdx = saved arg1

        Case 2
            ' COM call: ComProc(pThis=rcx, arg1=rdx, arg2=r8, pRetval=r9)
            EmitBytes bb, p, &H4C, &H8D, &H4C, &H24, &H30   ' lea r9, [rsp+30h]    ; r9  = pRetval (-> retval storage)
            EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40   ' mov r8, [rsp+40h]    ; r8  = saved arg2
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38   ' mov rdx, [rsp+38h]   ; rdx = saved arg1

        Case 3
            ' COM call: ComProc(pThis=rcx, arg1=rdx, arg2=r8, arg3=r9, pRetval=[rsp+20h])
            EmitBytes bb, p, &H4C, &H8D, &H54, &H24, &H30   ' lea r10, [rsp+30h]   ; r10 = pRetval (-> retval storage)
            EmitBytes bb, p, &H4C, &H89, &H54, &H24, &H20   ' mov [rsp+20h], r10   ; 5th arg (stack) = pRetval
            EmitBytes bb, p, &H4C, &H8B, &H4C, &H24, &H48   ' mov r9, [rsp+48h]    ; r9  = saved arg3
            EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40   ' mov r8, [rsp+40h]    ; r8  = saved arg2
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38   ' mov rdx, [rsp+38h]   ; rdx = saved arg1

        Case 4
            ' COM call: ComProc(pThis=rcx, arg1=rdx, arg2=r8, arg3=r9, arg4=[rsp+20h], pRetval=[rsp+28h])
            EmitBytes bb, p, &H4C, &H8B, &H54, &H24, &H50   ' mov r10, [rsp+50h]   ; r10 = saved arg4
            EmitBytes bb, p, &H4C, &H89, &H54, &H24, &H20   ' mov [rsp+20h], r10   ; 5th arg (stack) = arg4
            EmitBytes bb, p, &H4C, &H8D, &H54, &H24, &H30   ' lea r10, [rsp+30h]   ; r10 = pRetval (-> retval storage)
            EmitBytes bb, p, &H4C, &H89, &H54, &H24, &H28   ' mov [rsp+28h], r10   ; 6th arg (stack) = pRetval
            EmitBytes bb, p, &H4C, &H8B, &H4C, &H24, &H48   ' mov r9, [rsp+48h]    ; r9  = saved arg3
            EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40   ' mov r8, [rsp+40h]    ; r8  = saved arg2
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38   ' mov rdx, [rsp+38h]   ; rdx = saved arg1
    End Select

    ' --- Load pThis into rcx (1st arg for COM call) and call the COM method ---
    EmitBytes bb, p, &H48, &HB9                          ' mov rcx, <imm64>     ; rcx = pThis
    EmitPtr bb, p, pThis

    EmitBytes bb, p, &H48, &HB8                          ' mov rax, <imm64>     ; rax = ComProcAddress
    EmitPtr bb, p, ComProcAddress

    EmitBytes bb, p, &HFF, &HD0                          ' call rax             ; invoke COM method

    ' --- Epilogue: retrieve return value and tear down stack frame ---
    EmitBytes bb, p, &H8B, &H44, &H24, &H30             ' mov eax, [rsp+30h]   ; load Long retval from hidden storage

    EmitBytes bb, p, &H48, &H83, &HC4, &H58             ' add rsp, 0x58        ; deallocate stack frame
    EmitBytes bb, p, &HC3                                ' ret                  ; return to caller

    ' --- return_zero: EbMode not in run mode, skip callback ---
    Dim retZeroAt As Long: retZeroAt = p
    EmitBytes bb, p, &H31, &HC0                          ' xor eax, eax         ; return 0
    EmitBytes bb, p, &H48, &H83, &HC4, &H58             ' add rsp, 0x58        ; deallocate stack frame
    EmitBytes bb, p, &HC3                                ' ret

    bb(jneAt) = CByte(retZeroAt - (jneAt + 1))           ' patch jne offset to return_zero

    iThunkSize = p
    ReDim Preserve bb(0 To p - 1)

    AddressOfThunkForComProc = VirtualAlloc( _
        0, _
        iThunkSize, _
        MEM_COMMIT Or MEM_RESERVE, _
        PAGE_EXECUTE_READWRITE)

    If AddressOfThunkForComProc = 0 Then Exit Function

    CopyMemoryByPtr AddressOfThunkForComProc, bb(0), iThunkSize

#Else
    If ComProcAddress = 0 Or ObjPtr(o) = 0 Then Exit Function
    If ComProcArgCount < 1 Or ComProcArgCount > 4 Then Exit Function

    Dim p32 As Long
    Dim bb32() As Byte: ReDim bb32(0 To 127)

    ' --- Win32 thunk (cdecl-like frame, stdcall return cleanup) ---
    ' Incoming callback shape:
    '   Function Callback(ByRef Arg1, ..., ByRef ArgN) As Long
    '
    ' COM method shape expected by class instance:
    '   Function ComProc(ByVal pThis As Long, ByRef Arg1, ..., ByRef ArgN, ByRef FnRet As Long) As HRESULT
    '
    ' We push pThis at the front and append FnRet pointer at the end.
    ' The HRESULT return is ignored; FnRet is copied back as callback return.

    EmitBytes bb32, p32, &H55                             ' push ebp                 ; save caller frame pointer
    EmitBytes bb32, p32, &H89, &HE5                       ' mov ebp, esp             ; establish stack frame
    EmitBytes bb32, p32, &H83, &HEC, &H4                  ' sub esp, 4               ; reserve local [ebp-4] for FnRet storage

    ' --- EbMode IDE safety check ---
    EmitBytes bb32, p32, &HB8                             ' mov eax, <imm32>         ; eax = lEbMode
    Dim ebModeAt32 As Long: ebModeAt32 = p32
    EmitBytes bb32, p32, &H0, &H0, &H0, &H0               ' placeholder (patched later)
    EmitBytes bb32, p32, &H85, &HC0                       ' test eax, eax
    EmitBytes bb32, p32, &H74, &H7                         ' jz skip_check            ; skip if null (compiled)
    EmitBytes bb32, p32, &HFF, &HD0                       ' call eax                 ; call EbMode()
    EmitBytes bb32, p32, &H83, &HF8, &H1                  ' cmp eax, 1               ; running?
    Dim jneAt32 As Long: jneAt32 = p32 + 1
    EmitBytes bb32, p32, &H75, &H0                        ' jne return_zero          ; bail if not running (patched later)

    EmitBytes bb32, p32, &H89, &HE8                       ' mov eax, ebp             ; eax = frame base
    EmitBytes bb32, p32, &H83, &HE8, &H4                  ' sub eax, 4               ; eax = &FnRetLocal
    EmitBytes bb32, p32, &H50                             ' push eax                 ; push ByRef FnRet

    ' Push callback args in reverse order for stdcall.
    Dim arg4PushAt As Long: arg4PushAt = p32
    EmitBytes bb32, p32, &HFF, &H75, &H14                 ' push [ebp+20]            ; arg4 (NOP'd out if < 4 args)
    Dim arg3PushAt As Long: arg3PushAt = p32
    EmitBytes bb32, p32, &HFF, &H75, &H10                 ' push [ebp+16]            ; arg3 (NOP'd out if < 3 args)
    Dim arg2PushAt As Long: arg2PushAt = p32
    EmitBytes bb32, p32, &HFF, &H75, &HC                  ' push [ebp+12]            ; arg2 (NOP'd out if < 2 args)
    EmitBytes bb32, p32, &HFF, &H75, &H8                  ' push [ebp+8]             ; arg1 (always present)

    ' Push object pointer (pThis), patched after ObjPtr(o) is read.
    EmitBytes bb32, p32, &H68                             ' push imm32
    Dim objImmAt As Long: objImmAt = p32
    EmitBytes bb32, p32, &H55, &H55, &H55, &H55           ' imm32 placeholder for pThis

    ' Load COM procedure address and call.
    EmitBytes bb32, p32, &HB8                             ' mov eax, imm32
    Dim procImmAt As Long: procImmAt = p32
    EmitBytes bb32, p32, &H66, &H66, &H66, &H66           ' imm32 placeholder for ComProcAddress
    EmitBytes bb32, p32, &HFF, &HD0                       ' call eax                 ; invoke COM method

    ' Return callback result from local FnRet storage.
    EmitBytes bb32, p32, &H8B, &H45, &HFC                 ' mov eax, [ebp-4]         ; callback return value
    EmitBytes bb32, p32, &H89, &HEC                       ' mov esp, ebp             ; tear down local stack frame
    EmitBytes bb32, p32, &H5D                             ' pop ebp                  ; restore caller frame pointer

    EmitBytes bb32, p32, &HC2                             ' ret imm16                ; stdcall cleanup for callback args
    Dim retImmAt As Long: retImmAt = p32
    EmitBytes bb32, p32, &H8, &H0                         ' imm16 placeholder (arg bytes = ArgCount * 4)

    ' --- return_zero: EbMode not in run mode, skip callback ---
    Dim retZeroAt32 As Long: retZeroAt32 = p32
    EmitBytes bb32, p32, &H31, &HC0                       ' xor eax, eax             ; return 0
    EmitBytes bb32, p32, &H89, &HEC                       ' mov esp, ebp
    EmitBytes bb32, p32, &H5D                             ' pop ebp
    EmitBytes bb32, p32, &HC2                             ' ret imm16
    Dim retImmAt2 As Long: retImmAt2 = p32
    EmitBytes bb32, p32, &H8, &H0                         ' placeholder

    bb32(jneAt32) = CByte(retZeroAt32 - (jneAt32 + 1))    ' patch jne offset to return_zero

    iThunkSize = p32
    ReDim Preserve bb32(0 To p32 - 1)

    If ComProcArgCount < 4 Then
        bb32(arg4PushAt) = &H90
        bb32(arg4PushAt + 1) = &H90
        bb32(arg4PushAt + 2) = &H90
    End If

    If ComProcArgCount < 3 Then
        bb32(arg3PushAt) = &H90
        bb32(arg3PushAt + 1) = &H90
        bb32(arg3PushAt + 2) = &H90
    End If

    If ComProcArgCount < 2 Then
        bb32(arg2PushAt) = &H90
        bb32(arg2PushAt + 1) = &H90
        bb32(arg2PushAt + 2) = &H90
    End If

    CopyMemory bb32(ebModeAt32), lEbMode, 4&
    Dim pObj32 As Long: pObj32 = ObjPtr(o)
    CopyMemory bb32(objImmAt), pObj32, 4&
    CopyMemory bb32(procImmAt), ComProcAddress, 4&

    bb32(retImmAt) = ComProcArgCount * 4&
    bb32(retImmAt2) = ComProcArgCount * 4&

    Dim pThunk32 As LongPtr: pThunk32 = VirtualAlloc( _
        0, _
        iThunkSize, _
        MEM_COMMIT Or MEM_RESERVE, _
        PAGE_EXECUTE_READWRITE)

    If pThunk32 = 0 Then Exit Function

    CopyMemoryByPtr pThunk32, bb32(0), iThunkSize
    AddressOfThunkForComProc = pThunk32
#End If
End Function


' ============================================================================
' Bytecode emit helpers for thunk generation
' ============================================================================

Private Sub EmitBytes( _
    ByRef bb() As Byte, _
    ByRef p As Long, _
    ParamArray values() As Variant)

    Dim i As Long

    For i = LBound(values) To UBound(values)
        bb(p) = CByte(CLng(values(i)) And &HFF&)
        p = p + 1
    Next
End Sub

Private Sub EmitPtr( _
    ByRef bb() As Byte, _
    ByRef p As Long, _
    ByVal value As LongPtr)

    Dim cb As Long

    cb = LenB(value)
    CopyMemory bb(p), value, cb
    p = p + cb
End Sub

