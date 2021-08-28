Attribute VB_Name = "modMain"
' //
' // Access Forms collection of a VB6 executable
' // by The trick 2021
' //

Option Explicit

Private Type STARTUPINFO
    cb              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess        As Long
    hThread         As Long
    dwProcessId     As Long
    dwThreadId      As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
                         Alias "CreateProcessW" ( _
                         ByVal lpApplicationName As Long, _
                         ByVal lpCommandLine As Long, _
                         ByRef lpProcessAttributes As Any, _
                         ByRef lpThreadAttributes As Any, _
                         ByVal bInheritHandles As Long, _
                         ByVal dwCreationFlags As Long, _
                         ByRef lpEnvironment As Any, _
                         ByVal lpCurrentDirectory As Long, _
                         ByRef lpStartupInfo As STARTUPINFO, _
                         ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" ( _
                         ByVal hProcess As Long, _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" _
                    Alias "GetStartupInfoW" ( _
                    ByRef lpStartupInfo As STARTUPINFO)

Sub Main()
    Dim tSI             As STARTUPINFO
    Dim tPI             As PROCESS_INFORMATION
    Dim cVBGetGlobal    As Object
    Dim cForms          As Object
    Dim frmMain         As Object
    
    InitializeInjectLibrary
    
    tSI.cb = Len(tSI)
    
    GetStartupInfo tSI
    
    If CreateProcess(StrPtr(App.Path & "\..\dummy\dummy.exe"), 0, ByVal 0&, ByVal 0&, 0, 0, ByVal 0&, 0, tSI, tPI) = 0 Then
        MsgBox "CreateProcess failed"
        Exit Sub
    End If
    
    WaitForInputIdle tPI.hProcess, -1
    
    CloseHandle tPI.hProcess
    CloseHandle tPI.hThread
    
    Set cVBGetGlobal = CreateVBObjectInThread(tPI.dwThreadId, App.Path & "\..\dll\GetVBGlobal.dll", "CExtractor")
    Set cForms = cVBGetGlobal.Forms
    
    Set frmMain = cForms(0)
    
    ' // Change back color of picturebox
    frmMain.Controls("picTest").BackColor = vbRed
    
    ' // Draw line on picturebox
    frmMain.Controls("picTest").Line (0, 0)-Step(100, 50), vbGreen, BF
    frmMain.Caption = "Test"

    UninitializeInjectLibrary
    
End Sub
