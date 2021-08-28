; FASM
; by The trick 2018-2021

include "win32wx.inc"

WM_CREATEOBJECT = WM_USER + 1
WM_DESTROYME = WM_USER + 2

struc GUID def
 {
   match d1-d2-d3-d4-d5, def
    \{
      .Data1 dd 0x\#d1
      .Data2 dw 0x\#d2
      .Data3 dw 0x\#d3
      .Data4 db 0x\#d4 shr 8,0x\#d4 and 0FFh
      .Data5 db 0x\#d5 shr 40,0x\#d5 shr 32 and 0FFh,0x\#d5 shr 24 and 0FFh,0x\#d5 shr 16 and 0FFh,0x\#d5 shr 8 and 0FFh,0x\#d5 and 0FFh
    \}
 }

struct tAPITable
    pfnRegisterClassExW       dd ?
    pfnUnregisterClassW       dd ?
    pfnCreateEventW	      dd ?
    pfnSetWindowsHookExW      dd ?
    pfnPostThreadMessageW     dd ?
    pfnWaitForSingleObject    dd ?
    pfnCloseHandle	      dd ?
    pfnUnhookWindowsHookEx    dd ?
    pfnCallNextHookEx	      dd ?
    pfnCreateWindowExW	      dd ?
    pfnDestroyWindow	      dd ?
    pfnSetEvent 	      dd ?
    pfnPostMessageW	      dd ?
    pfnLoadLibraryW	      dd ?
    pfnFreeLibrary	      dd ?
    pfnGetProcAddress	      dd ?
    pfnSetWindowLongW	      dd ?
    pfnGetWindowLongW	      dd ?
    pfnHeapAlloc	      dd ?
    pfnHeapReAlloc	      dd ?
    pfnHeapFree 	      dd ?
    pfnGetProcessHeap	      dd ?
    pfnGlobalFree	      dd ?
    pfnGlobalSize	      dd ?
    pfnGlobalLock	      dd ?
    pfnGlobalUnlock	      dd ?
    pfnSetTimer 	      dd ?
    pfnKillTimer	      dd ?

    pfnCoInitialize	      dd ?
    pfnCoUninitialize	      dd ?
    pfnCoMarshalInterface     dd ?
    pfnCreateStreamOnHGlobal  dd ?
    pfnGetHGlobalFromStream   dd ?

ends

struct tObjectDesc
    hLibrary		   dd ?
    pObject		   dd ?
    pFactory		   dd ?
    hMem		   dd ?
    dwStmDataSize	   dd ?
    pStmData		   dd ?
ends

struct tProcessData
    tAPIs		   tAPITable
    dwClassAtom 	   dd ?
    dwNumOfWindows	   dd ?
    dwTimerID		   dd ?
ends

struct tThreadData
    dwDestThreadID	   dd ?
    hEvent		   dd ?
    hHook		   dd ?
    hWnd		   dd ?
    pszDllName		   dd ?
    clsid		   dd 4 dup (?)
    iid 		   dd 4 dup (?)
    hr			   dd ?
ends

struct tShellCodeData
    pProcessData	   dd ?
    pThreadData 	   dd ?
ends

interface IClassFactory,\
	   QueryInterface,\
	   AddRef,\
	   Release,\
	   CreateInstance,\
	   LockServer

interface IUnknown,\
	   QueryInterface,\
	   AddRef,\
	   Release

interface IMoniker,\
	   QueryInterface,\
	   AddRef,\
	   Release,\
	   GetClassID,\
	   IsDirty,\
	   Load,\
	   Save,\
	   GetSizeMax,\
	   BindToObject,\
	   BindToStorage,\
	   Reduce,\
	   ComposeWith,\
	   Enum,\
	   IsEqual,\
	   Hash,\
	   IsRunning,\
	   GetTimeOfLastChange,\
	   Inverse,\
	   CommonPrefixWith,\
	   RelativePathTo,\
	   GetDisplayName,\
	   ParseDisplayName,\
	   IsSystemMoniker
