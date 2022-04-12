Attribute VB_Name = "MdlAnuleKeys"

Option Explicit
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Const WM_KEYDOWN As Long = &H100

Public Const GWL_WNDPROC = (-4)
Dim PrevProc As Long
Public Sub HookForm(ByVal hwnd As Long)
    PrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHookForm(ByVal hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, PrevProc
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If uMsg = WM_KEYDOWN Then
        If ((wParam = vbKeyV) Or (wParam = vbKeyC) Or (wParam = vbKeyX) And GetAsyncKeyState(vbKeyControl)) Then
            WindowProc = 0
            Exit Function
        End If
    End If
    
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    'Debug.Print uMsg, wParam, lParam
End Function

