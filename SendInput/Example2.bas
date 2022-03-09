Private Declare Function SendInput Lib "user32.dll" _
 (ByVal nInputs As Long, ByRef pInputs As Any, _
 ByVal cbSize As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" _
 (ByVal cChar As Byte) As Integer

Private Type KeyboardInput      '   typedef struct tagINPUT {
 dwType As Long                '     DWORD type;
 wVK As Integer                '     union {MOUSEINPUT mi;
 wScan As Integer              '               KEYBDINPUT ki;
 dwFlags As Long               '               HARDWAREINPUT hi;
 dwTime As Long                '              };
 dwExtraInfo As Long           '     }INPUT, *PINPUT;
 dwPadding As Currency         '   8 extra bytes, because mouses take more.
End Type

Private Const INPUT_MOUSE As Long = 0
Private Const INPUT_KEYBOARD As Long = 1
Private Const KEYEVENTF_KEYUP As Long = 2
Private Const VK_LSHIFT = &HA0

Public Sub SendKey(ByVal Data As String)
Dim ki() As KeyboardInput
Dim i As Long
Dim o As Long ' output buffer position
Dim c As String ' character

ReDim ki(1 To Len(Data) * 4) As KeyboardInput
o = 1

For i = 1 To Len(Data)
 c = Mid$(Data, i, 1)
 Select Case c
   Case "A" To "Z": ' upper case
     ki(o).dwType = INPUT_KEYBOARD 'shift down
     ki(o).wVK = VK_LSHIFT
     ki(o + 1) = ki(o) ' key down
     ki(o + 1).wVK = VkKeyScan(Asc(c))
     ki(o + 2) = ki(o + 1) ' key up
     ki(o + 2).dwFlags = KEYEVENTF_KEYUP
     ki(o + 3) = ki(o) ' shift up
     ki(o + 3).dwFlags = KEYEVENTF_KEYUP
     o = o + 4
   Case Else: ' lower case
     ki(o).dwType = INPUT_KEYBOARD
     ki(o).wVK = VkKeyScan(Asc(c))
     ki(o + 1) = ki(o)
     ki(o + 1).dwFlags = KEYEVENTF_KEYUP
     o = o + 2
 End Select
Next i

Debug.Print SendInput(o - 1, ki(1), LenB(ki(1))),
'Debug.Print Err.LastDllError
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Text1.SetFocus
DoEvents
Call SendKey("This Is A Test")
End Sub