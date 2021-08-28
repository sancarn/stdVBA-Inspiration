; FASM
; by The trick
; 2018-2021


include "declarations.inc"

entry start

section '.idata' import data readable writeable

library kernel32,'KERNEL32.DLL',\
	user32,'USER32.DLL', \
	ole32,'OLE32.DLL'

import user32,\
       RegisterClassExW, 'RegisterClassExW', \
       UnregisterClassW, 'UnregisterClassW', \
       SetWindowsHookExW, 'SetWindowsHookExW', \
       UnhookWindowsHookEx, 'UnhookWindowsHookEx', \
       PostThreadMessageW, 'PostThreadMessageW', \
       GetMessageW, 'GetMessageW', \
       TranslateMessage, 'TranslateMessage', \
       DispatchMessageW, 'DispatchMessageW' , \
       CallNextHookEx, 'CallNextHookEx', \
       CreateWindowExW, 'CreateWindowExW', \
       DestroyWindow, 'DestroyWindow', \
       PostMessageW, 'PostMessageW', \
       SendMessageW, 'SendMessageW', \
       SetWindowLongW, 'SetWindowLongW', \
       GetWindowLongW, 'GetWindowLongW', \
       SetTimer, 'SetTimer', \
       KillTimer, 'KillTimer'

import kernel32,\
       CreateEventW, 'CreateEventW', \
       CloseHandle, 'CloseHandle', \
       CreateThread, 'CreateThread', \
       WaitForSingleObject, 'WaitForSingleObject', \
       SetEvent, 'SetEvent', \
       LoadLibraryW, 'LoadLibraryW', \
       GetProcAddress, 'GetProcAddress', \
       HeapAlloc, 'HeapAlloc', \
       HeapReAlloc, 'HeapReAlloc', \
       HeapFree, 'HeapFree', \
       GetProcessHeap, 'GetProcessHeap', \
       FreeLibrary, 'FreeLibrary', \
       GlobalSize, 'GlobalSize', \
       GlobalLock, 'GlobalLock', \
       GlobalUnlock, 'GlobalUnlock', \
       GlobalFree, 'GlobalFree'


import ole32, \
       CoInitialize, 'CoInitialize', \
       CoUnmarshalInterface, 'CoUnmarshalInterface', \
       CreateStreamOnHGlobal, 'CreateStreamOnHGlobal'

section '.text' code readable executable writable


g_pszDllName du "C:\Temp\Project1.dll", 0
g_clsid GUID 87F10035-26D5-48B9-85C2-B66D1B120A2C
IID_IDispatch GUID 00020400-0000-0000-C000-000000000046

proc start
    locals
      msg MSG
      pObject dd ?
      pStream dd ?
    endl

    mov [g_tShellData + tShellCodeData.pProcessData], g_tProcessData
    mov [g_tShellData + tShellCodeData.pThreadData], g_tThreadData

    mov esi, [g_tShellData + tShellCodeData.pProcessData]
    mov edi, [g_tShellData + tShellCodeData.pThreadData]

    mov eax, [RegisterClassExW]
    mov [esi + tProcessData.tAPIs.pfnRegisterClassExW], eax
    mov eax, [UnregisterClassW]
    mov [esi + tProcessData.tAPIs.pfnUnregisterClassW], eax
    mov eax, [CreateEventW]
    mov [esi + tProcessData.tAPIs.pfnCreateEventW], eax
    mov eax, [CloseHandle]
    mov [esi + tProcessData.tAPIs.pfnCloseHandle], eax
    mov eax, [SetWindowsHookExW]
    mov [esi + tProcessData.tAPIs.pfnSetWindowsHookExW], eax
    mov eax, [UnhookWindowsHookEx]
    mov [esi + tProcessData.tAPIs.pfnUnhookWindowsHookEx], eax
    mov eax, [PostThreadMessageW]
    mov [esi + tProcessData.tAPIs.pfnPostThreadMessageW], eax
    mov eax, [WaitForSingleObject]
    mov [esi + tProcessData.tAPIs.pfnWaitForSingleObject], eax
    mov eax, [CallNextHookEx]
    mov [esi + tProcessData.tAPIs.pfnCallNextHookEx], eax
    mov eax, [CreateWindowExW]
    mov [esi + tProcessData.tAPIs.pfnCreateWindowExW], eax
    mov eax, [DestroyWindow]
    mov [esi + tProcessData.tAPIs.pfnDestroyWindow], eax
    mov eax, [SetEvent]
    mov [esi + tProcessData.tAPIs.pfnSetEvent], eax
    mov eax, [PostMessageW]
    mov [esi + tProcessData.tAPIs.pfnPostMessageW], eax
    mov eax, [LoadLibraryW]
    mov [esi + tProcessData.tAPIs.pfnLoadLibraryW], eax
    mov eax, [GetProcAddress]
    mov [esi + tProcessData.tAPIs.pfnGetProcAddress], eax
    mov eax, [GetWindowLongW]
    mov [esi + tProcessData.tAPIs.pfnGetWindowLongW], eax
    mov eax, [SetWindowLongW]
    mov [esi + tProcessData.tAPIs.pfnSetWindowLongW], eax
    mov eax, [HeapAlloc]
    mov [esi + tProcessData.tAPIs.pfnHeapAlloc], eax
    mov eax, [HeapReAlloc]
    mov [esi + tProcessData.tAPIs.pfnHeapReAlloc], eax
    mov eax, [HeapFree]
    mov [esi + tProcessData.tAPIs.pfnHeapFree], eax
    mov eax, [GetProcessHeap]
    mov [esi + tProcessData.tAPIs.pfnGetProcessHeap], eax
    mov eax, [FreeLibrary]
    mov [esi + tProcessData.tAPIs.pfnFreeLibrary], eax
    mov eax, [GlobalSize]
    mov [esi + tProcessData.tAPIs.pfnGlobalSize], eax
    mov eax, [GlobalLock]
    mov [esi + tProcessData.tAPIs.pfnGlobalLock], eax
    mov eax, [GlobalUnlock]
    mov [esi + tProcessData.tAPIs.pfnGlobalUnlock], eax
    mov eax, [GlobalFree]
    mov [esi + tProcessData.tAPIs.pfnGlobalFree], eax
    mov eax, [SetTimer]
    mov [esi + tProcessData.tAPIs.pfnSetTimer], eax
    mov eax, [KillTimer]
    mov [esi + tProcessData.tAPIs.pfnKillTimer], eax

    mov eax, [fs : 0x18]
    mov eax, [eax + 0x24]
    mov [edi + tThreadData.dwDestThreadID], eax
    mov [edi + tThreadData.pszDllName], g_pszDllName

    push edi

    lea edi, [edi + tThreadData.clsid]
    lea esi, [g_clsid]
    mov ecx, 4
    rep movsd

    pop edi
    push edi

    lea edi, [edi + tThreadData.iid]
    lea esi, [IID_IDispatch]
    mov ecx, 4
    rep movsd

    pop edi

    invoke CloseHandle, <invoke CreateThread, 0, 0, g_tShellData + 8, addr g_tShellData, 0, 0>

  .msg_loop:

      .if [edi + tThreadData.hWnd]


	invoke SendMessageW, [edi + tThreadData.hWnd], WM_CREATEOBJECT, 0, 0

	invoke CreateStreamOnHGlobal, [eax + tObjectDesc.hMem], 0, addr pStream
	invoke CoUnmarshalInterface, [pStream], IID_IDispatch, addr pObject

	comcall [pObject], IUnknown, Release

	invoke DestroyWindow, [edi + tThreadData.hWnd]

      .endif

      invoke GetMessageW, addr msg, 0, 0, 0

      .if ~eax
	jmp .exit_proc
      .endif

      invoke TranslateMessage, addr msg
      invoke DispatchMessageW, addr msg

    jmp .msg_loop

  .exit_proc:

    ret

endp

g_tProcessData tProcessData ?
g_tThreadData tThreadData ?

g_tShellData:

file "shellcode.bin"