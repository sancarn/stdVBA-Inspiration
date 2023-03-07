Attribute VB_Name = "modExec"
Option Explicit

'from https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit


Public Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    phReadPipe As LongPtr, _
    phWritePipe As LongPtr, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Public Declare PtrSafe Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, _
    lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, _
    lpProcessInformation As Any) As Long

Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Public Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long


Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Long) As Long
    
Declare PtrSafe Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As LongPtr, ByVal lParam As LongPtr) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type


Public Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

'this is the structure to pass more than 3 fds to a child process

'see https://github.com/libuv/libuv/blob/v1.x/src/win/process-stdio.c
Public Type STDIO_BUFFER
    number_of_fds As Long
    crt_flags(0 To 4) As Byte
    os_handle(0 To 4) As LongPtr
End Type

' the fields crt_flags and os_handle must lie contigously in memory
' i.e. should not be aligned to byte boundaries
' you cannot define a packed struct in VBA
' thats why we need to have a second struct

#If Win64 Then
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 44) As Byte
End Type
#Else
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 24) As Byte
End Type
#End If

Public Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Const STARTF_USESTDHANDLES = &H100&
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESHOWWINDOW As Long = &H1&
'Public Const STARTF_CREATE_NO_WINDOW As Long = &H8000000

' we need to move memory

Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

' Using the following defintions I tried to hide the console windows
' This does not yet work, so it is commented out
'
'Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'
' Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Boolean
'
'Public Function getNameFromHwnd(hWnd As Long) As String
'Dim title As String * 255
'Dim tLen As Long
'tLen = GetWindowTextLength(hWnd)
'GetWindowText hWnd, title, 255
'getNameFromHwnd = Left(title, tLen)
'End Function
'
'
'
'
'Public Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'    Dim Ret As Long, sText As String
'
'    'CloseWindow hwnd ' This is the handle to your process window which you created.
'
'    sText = getNameFromHwnd(hWnd)
'
'    If sText = "" Then
'        Call ShowWindow(hWnd, 0)
'    End If
'
'    EnumThreadWndProc = 1
'
'End Function


