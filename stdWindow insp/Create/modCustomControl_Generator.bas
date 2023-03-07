Attribute VB_Name = "modCustomControl_Generator"
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal lhwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Long
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal lhwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As LongPtr, ByVal lpIconName As Any) As LongPtr
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function ShowCursor Lib "user32" (ByVal fShow As Long) As Long
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal lhwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long

Private Type WNDCLASSEX 'size=48/56 on 32 bit and 80 on 64 bit?
    cbSize As Long  '4
    style As Long   '4
    lpfnWndProc As LongPtr  '4 or 8
    cbClsExtra As Long      '4
    cbWndExtra As Long      '4
    hInstance As LongPtr    '4 or 8
    hIcon As LongPtr        '4 or 8
    hCursor As LongPtr      '4 or 8
    hbrBackground As LongPtr '4 or 8
    lpszMenuName As String   '8
    lpszClassName As String '8
    hIconSm As LongPtr      '4 or 8
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type


Public Type MSG
    lhwnd As LongPtr
    tmessage As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const SW_SHOWNORMAL As Long = 1
Private Const CS_HREDRAW As Long = &H2
Private Const CS_VREDRAW As Long = &H1
Private Const IDI_APPLICATION As Long = 32512&
Private Const IDC_ARROW As Long = 32512&
Private Const IDC_HAND As Long = 32649&
Private Const WHITE_BRUSH As Integer = 0
Private Const BLACK_BRUSH As Integer = 4


Private Const CLASSNAME = "Custom"

Sub main()
    Dim hInst As LongPtr
    Dim hWnd As LongPtr
    Dim a_hWnd As LongPtr
    hInst = Application.HinstancePtr
    
    Dim tmessage As MSG
    
    Dim wc As WNDCLASSEX
    Dim ires As Long
    Dim lres As Long
    
    wc.cbSize = LenB(wc) '80
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnWndProc = FunctionPointer(AddressOf WndProc)
    wc.cbClsExtra = 0&
    wc.cbWndExtra = 0&
    wc.hInstance = hInst
    wc.hIcon = LoadIcon(0&, IDI_APPLICATION)
    wc.hCursor = LoadCursor(0&, IDC_ARROW)
    wc.hbrBackground = GetStockObject(BLACK_BRUSH)
    wc.lpszMenuName = ""
    wc.lpszClassName = "MyClass"
    wc.hIconSm = LoadIcon(0&, IDI_APPLICATION)
    ires = RegisterClassEx(wc)
    lres = GetLastError()
    
    const WS_EX_TOPMOST = &H00000008
    const WS_POPUP = &H80000000
  '  a_hWnd = CreateWindowEx(ByVal 0&, "BUTTON", "Hello !", WS_POPUP, 0, 0, 320, 240, ByVal 0&, ByVal 0&, hInst, ByVal 0&)
    a_hWnd = CreateWindowEx(WS_EX_TOPMOST, "MyClass", "TEST", WS_POPUP, 20, 20, 324, 290, ByVal 0&, ByVal 0&, hInst, ByVal 0&)
  '  a_hWnd = CreateWindowEx(ByVal 0&, "STATIC", "TEST", WS_BORDER Or WS_CAPTION Or WS_POPUP, 20, 20, 324, 290, ByVal 0&, ByVal 0&, hInst, ByVal 0&)
    If a_hWnd = 0 Then
        lres = GetLastError()
        MsgBox "Could not open window - GetLastError reports " + Str(lres)
        GoTo nowindow
    End If
    
    ShowWindow a_hWnd, SW_SHOWNORMAL
    
    Do While 0 <> GetMessage(tmessage, 0&, 0&, 0&)      'Retrieve a message from the calling thread�s message queue
        TranslateMessage tmessage                       'Translate virtual-key messages into character messages (character messages are posted to the calling thread's message queue).
        DispatchMessage tmessage                        'Dispatch message to window procedure (WindowProc)
    Loop


    
    DestroyWindow a_hWnd
nowindow:
    lres = UnregisterClass("MyClass", hInst)
    lres = GetLastError()
End Sub

'Returns the value from the AddressOf unary operator.
Function FunctionPointer(ByVal lPtr As LongPtr) As LongPtr
    FunctionPointer = lPtr
End Function

Private Function WndProc(ByVal hWnd As LongPtr, ByVal message As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    WndProc = DefWindowProc(hWnd, message, wParam, lParam)
End Function
