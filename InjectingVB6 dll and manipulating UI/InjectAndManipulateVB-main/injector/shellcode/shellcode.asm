
format binary
use32

dummyvar = EntryPoint

include "declarations.inc"

g_tShellData tShellCodeData 0, 0

proc EntryPoint uses edi esi ebx
    locals
	tClass WNDCLASSEX
    endl

    call .get_ip
  .get_ip:
    pop ebx

    lea edi, [tClass]
    xor eax, eax
    mov ecx, (sizeof.WNDCLASSEX + 4) / 4
    rep stosd

    lea esi, [ebx - (.get_ip - g_tShellData)]
    mov edi, [esi + tShellCodeData.pProcessData]
    mov esi, [esi + tShellCodeData.pThreadData]

    push ebx
    call @f
    du "ole32", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnLoadLibraryW
    mov ebx, eax
    call @f
    db "CoInitialize", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCoInitialize], eax
    call @f
    db "CoUninitialize", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCoUninitialize], eax
    call @f
    db "CoMarshalInterface", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCoMarshalInterface], eax
    call @f
    db "CreateStreamOnHGlobal", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCreateStreamOnHGlobal], eax
    call @f
    db "GetHGlobalFromStream", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnGetHGlobalFromStream], eax

    pop ebx

    .if ~[edi + tProcessData.dwClassAtom]

      mov [tClass.cbSize], sizeof.WNDCLASSEX
      lea eax, [ebx + WndProc - .get_ip]
      mov [tClass.lpfnWndProc], eax
      lea eax, [ebx + CLASS_NAME - .get_ip]
      mov [tClass.lpszClassName], eax
      mov eax, [fs:0x30]
      mov eax, [eax + 8] ; Image base
      mov [tClass.hInstance], eax
      mov [tClass.cbWndExtra], 8

      invoke edi + tProcessData.tAPIs.pfnRegisterClassExW, addr tClass

      .if ~ax

	stdcall GetLastErrorHresult
	mov [esi + tThreadData.hr], eax
	jmp .clean_up

      .endif

      movzx eax, ax
      mov [edi + tProcessData.dwClassAtom], eax

    .endif

    invoke edi + tProcessData.tAPIs.pfnCreateEventW, 0, 0, 0, 0

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [esi + tThreadData.hEvent], eax
    lea edx, [ebx + HookProc - .get_ip]

    invoke edi + tProcessData.tAPIs.pfnSetWindowsHookExW, WH_GETMESSAGE, edx, 0, [esi + tThreadData.dwDestThreadID]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [esi + tThreadData.hHook], eax

    invoke edi + tProcessData.tAPIs.pfnPostThreadMessageW, [esi + tThreadData.dwDestThreadID], WM_NULL, 0, esi

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    invoke edi + tProcessData.tAPIs.pfnWaitForSingleObject, [esi + tThreadData.hEvent]

    .if eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

  .clean_up:

    .if [esi + tThreadData.hHook]
      invoke edi + tProcessData.tAPIs.pfnUnhookWindowsHookEx, [esi + tThreadData.hHook]
      mov [esi + tThreadData.hHook], 0
    .endif

    .if [esi + tThreadData.hEvent]
      invoke edi + tProcessData.tAPIs.pfnCloseHandle, [esi + tThreadData.hEvent]
      mov [esi + tThreadData.hEvent], 0
    .endif

    ret
endp

proc HookProc uses esi edi ebx, lCode, wParam, lParam
    locals
      dwRet dd ?
    endl

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    .if [lCode] = HC_ACTION

	invoke edi + tProcessData.tAPIs.pfnCallNextHookEx, 0, [lCode], [wParam], [lParam]
	mov [dwRet], eax

	invoke edi + tProcessData.tAPIs.pfnUnhookWindowsHookEx, [esi + tThreadData.hHook]
	mov [esi + tThreadData.hHook], 0

	invoke edi + tProcessData.tAPIs.pfnCoInitialize, 0

	.if signed eax < 0
	  mov [esi + tThreadData.hr], eax
	  jmp .release_thread
	.endif

	lea ecx, [ebx + CLASS_NAME - g_tShellData]
	mov edx, [fs:0x30]
	mov edx, [edx + 8] ; Image base

	invoke edi + tProcessData.tAPIs.pfnCreateWindowExW, 0, ecx, 0, 0, 0, 0, 0, 0, -3, 0, edx, 0

	mov [esi + tThreadData.hWnd], eax

	.if ~eax

	  stdcall GetLastErrorHresult
	  mov [esi + tThreadData.hr], eax

	  invoke edi + tProcessData.tAPIs.pfnCoUninitialize

	.else
	  inc [edi + tProcessData.dwNumOfWindows]
	.endif

      .release_thread:

	invoke edi + tProcessData.tAPIs.pfnSetEvent, [esi + tThreadData.hEvent]

	mov eax, [dwRet]

    .else
	invoke edi + tProcessData.tAPIs.pfnCallNextHookEx, 0, [lCode], [wParam], [lParam]
    .endif

    ret

endp

proc GetLastErrorHresult

    mov eax, [fs:0x18]
    mov eax, [eax + 0x34]
    and eax, 0xffff
    or eax, 0x80070000

    ret

endp

proc WndProc uses edi esi ebx, hWnd, uMsg, wParam, lParam

    mov eax, [uMsg]

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    .if eax = WM_NCCREATE

      mov eax, 1

    .elseif eax = WM_NCDESTROY

      invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [hWnd], 0

      .if eax

	push esi
	push ebx

	mov ebx, eax

	invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [hWnd], 4
	mov esi, eax

	push esi

	.repeat

	  mov eax, [esi + tObjectDesc.pObject]
	  comcall eax, IUnknown, Release

	  mov eax, [esi + tObjectDesc.pFactory]
	  comcall eax, IClassFactory, LockServer, FALSE

	  mov eax, [esi + tObjectDesc.pFactory]
	  comcall eax, IClassFactory, Release

	  call @f
	  db "DllCanUnloadNow", 0
	  @@:
	  invoke edi + tProcessData.tAPIs.pfnGetProcAddress, [esi + tObjectDesc.hLibrary]

	  stdcall eax

	  .if eax = 0
	    invoke edi + tProcessData.tAPIs.pfnFreeLibrary, [esi + tObjectDesc.hLibrary]
	  .endif

	  add esi, sizeof.tObjectDesc
	  dec ebx

	.until ebx = 0

	pop esi

	invoke edi + tProcessData.tAPIs.pfnHeapFree, <invoke edi + tProcessData.tAPIs.pfnGetProcessHeap>, 0, esi

	pop ebx
	pop esi

      .endif

      mov [esi + tThreadData.hWnd], 0

      invoke edi + tProcessData.tAPIs.pfnCoUninitialize

      dec [edi + tProcessData.dwNumOfWindows]

      .if ZERO?

	lea eax, [ebx + TimerProc - g_tShellData]
	invoke edi + tProcessData.tAPIs.pfnSetTimer, 0, 0, 1, eax
	mov [edi + tProcessData.dwTimerID], eax

      .endif

    .elseif eax = WM_CREATEOBJECT

      stdcall CreateObject

    .elseif eax = WM_DESTROYME

      invoke edi + tProcessData.tAPIs.pfnDestroyWindow, [hWnd]

    .else

      xor eax, eax

    .endif

    ret
endp

proc TimerProc uses esi edi ebx, hWnd, uMsg, idEvent, dwTime

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    mov eax, [fs:0x30]
    mov eax, [eax + 8]
    lea ecx, [ebx + CLASS_NAME - g_tShellData]

    invoke edi + tProcessData.tAPIs.pfnUnregisterClassW, ecx, eax

    mov [edi + tProcessData.dwClassAtom], 0

    invoke edi + tProcessData.tAPIs.pfnKillTimer, 0, [edi + tProcessData.dwTimerID]

    mov [edi + tProcessData.dwTimerID], 0

    ret

endp

proc CreateObject uses esi edi ebx
    locals
      hLib dd 0
      pFactory dd 0
      pObject dd 0
      pStream dd 0
      hMem dd 0
      pStmData dd 0
      dwStmDataSize dd 0
      pList dd ?
      pRet dd 0
    endl

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    invoke edi + tProcessData.tAPIs.pfnLoadLibraryW, [esi + tThreadData.pszDllName]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [hLib], eax
    call @f
    db "DllGetClassObject", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, eax

    .if ~eax
      jmp .clean_up
    .endif

    lea ecx, [ebx + IID_IClassFactory - g_tShellData]

    stdcall eax, addr esi + tThreadData.clsid, ecx, addr pFactory

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    mov eax, [pFactory]
    comcall eax, IClassFactory, LockServer, TRUE

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    ;lea ecx, [ebx + IID_IDispatch - g_tShellData]

    mov eax, [pFactory]
    comcall eax, IClassFactory, CreateInstance, 0, addr esi + tThreadData.iid, addr pObject

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    invoke edi + tProcessData.tAPIs.pfnCreateStreamOnHGlobal, 0, 0, addr pStream

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    invoke edi + tProcessData.tAPIs.pfnGetHGlobalFromStream, [pStream], addr hMem

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    ;lea ecx, [ebx + IID_IDispatch - g_tShellData]

    invoke edi + tProcessData.tAPIs.pfnCoMarshalInterface, [pStream], addr esi + tThreadData.iid, [pObject], 0, 0, 0

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    invoke edi + tProcessData.tAPIs.pfnGlobalSize, [hMem]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [dwStmDataSize], eax

    invoke edi + tProcessData.tAPIs.pfnGlobalLock, [hMem]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [pStmData], eax

    invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [esi + tThreadData.hWnd], 4

    .if eax

      push ebx
      push esi

      mov ebx, eax

      invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [esi + tThreadData.hWnd], 0
      inc eax
      imul eax, sizeof.tObjectDesc
      mov esi, eax

      invoke edi + tProcessData.tAPIs.pfnHeapReAlloc, <invoke edi + tProcessData.tAPIs.pfnGetProcessHeap>, 0, ebx, esi
      mov [pList], eax

      sub esi, sizeof.tObjectDesc
      add eax, esi

      pop esi
      pop ebx

    .else
      invoke edi + tProcessData.tAPIs.pfnHeapAlloc, <invoke edi + tProcessData.tAPIs.pfnGetProcessHeap>, 0, sizeof.tObjectDesc
      mov [pList], eax
    .endif

    .if ~eax
      mov [esi + tThreadData.hr], 0x80070007
      jmp .clean_up
    .endif

    mov edx, [hLib]
    mov [eax + tObjectDesc.hLibrary], edx
    mov edx, [pObject]
    mov [eax + tObjectDesc.pObject], edx
    mov edx, [pFactory]
    mov [eax + tObjectDesc.pFactory], edx
    mov edx, [hMem]
    mov [eax + tObjectDesc.hMem], edx
    mov edx, [pStmData]
    mov [eax + tObjectDesc.pStmData], edx
    mov edx, [dwStmDataSize]
    mov [eax + tObjectDesc.dwStmDataSize], edx

    mov [pRet], eax

    invoke edi + tProcessData.tAPIs.pfnSetWindowLongW, [esi + tThreadData.hWnd], 4, [pList]

    invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [esi + tThreadData.hWnd], 0
    inc eax
    invoke edi + tProcessData.tAPIs.pfnSetWindowLongW, [esi + tThreadData.hWnd], 0, eax

  .clean_up:

    .if [pStream]
      comcall [pStream], IUnknown, Release
    .endif

    .if ~[pRet]

      .if [hMem]
	invoke edi + tProcessData.tAPIs.pfnGlobalFree, [hMem]
      .endif

      .if [pObject]
	comcall [pObject], IUnknown, Release
      .endif

      .if [pFactory]
	comcall [pFactory], IClassFactory, LockServer, FALSE
	comcall [pFactory], IClassFactory, Release
      .endif

      .if [hLib]
	invoke edi + tProcessData.tAPIs.pfnFreeLibrary, [hLib]
      .endif

    .endif

    mov eax, [pRet]

    ret

endp

IID_IClassFactory      GUID 00000001-0000-0000-C000-000000000046
CLASS_NAME: du "VBInjectClass", 0