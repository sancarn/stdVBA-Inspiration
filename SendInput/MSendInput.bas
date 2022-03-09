Attribute VB_Name = "MSendInput"
' *********************************************************************
'  Copyright �2007 Karl E. Peterson, All Rights Reserved
'  http://vb.mvps.org/
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer

Private Type KeyboardInput       '   typedef struct tagINPUT {
   dwType As Long                '     DWORD type;
   wVK As Integer                '     union {MOUSEINPUT mi;
   wScan As Integer              '            KEYBDINPUT ki;
   dwFlags As Long               '            HARDWAREINPUT hi;
   dwTime As Long                '     };
   dwExtraInfo As Long           '   }INPUT, *PINPUT;
   dwPadding As Currency         '   8 extra bytes, because mouses take more.
End Type

Private Const INPUT_MOUSE As Long = 0
Private Const INPUT_KEYBOARD As Long = 1

Private Const KEYEVENTF_EXTENDEDKEY As Long = 1
Private Const KEYEVENTF_KEYUP As Long = 2

Private m_Data As String
Private m_DatPtr As Long
Private m_Events() As KeyboardInput
Private m_EvtPtr As Long

Private m_NamedKeys As Collection
Private m_ExtendedKeys As Collection
Private m_ShiftFlags As Long

Private Const defBufferSize As Long = 512

'' Toggle used to suck in VB6 functionality in VB5.
'#Const VB6 = False

Public Sub MySendKeys(Data As String, Optional Wait As Boolean)
   Dim i As Long
   
   ' Make sure our collection of named keys has been built.
   If m_NamedKeys Is Nothing Then
      Call BuildCollections
   End If
   
   ' Clear buffer, reset pointers, and cache send data.
   ReDim m_Events(0 To defBufferSize - 1) As KeyboardInput
   m_EvtPtr = 0
   m_DatPtr = 0
   m_Data = Data
   
   ' Loop through entire passed string.
   Do While m_DatPtr < Len(Data)
      ' Process next token in data string.
      Call DoNext
      
      ' Make sure there's still plenty of room in the buffer.
      If m_EvtPtr >= (UBound(m_Events) - 24) Then
         ReDim Preserve m_Events(0 To (UBound(m_Events) + defBufferSize) - 1)
      End If
   Loop
   
   ' Send the processed string to the foreground window!
   If m_EvtPtr > 0 Then
      ' All events are keyboard based.
      For i = 0 To m_EvtPtr - 1
         With m_Events(i)
            .dwType = INPUT_KEYBOARD
            'Debug.Print .wVK, .dwFlags
         End With
      Next i
      ' m_EvtPtr is 0-based, but nInputs is 1-based.
      Call SendInput(m_EvtPtr, m_Events(0), Len(m_Events(0)))
   End If
   
   ' Clean up
   Erase m_Events
End Sub

Private Sub DoNext()
   Dim this As String
   
   ' Advance data pointer, and extract next char.
   m_DatPtr = m_DatPtr + 1
   this = Mid$(m_Data, m_DatPtr, 1)
   
   ' Branch to appropriate helper routine.
   If InStr("+^%", this) Then
      Call ProcessShift(this)
   ElseIf this = "(" Then
      Call ProcessGroup
   ElseIf this = "{" Then
      Call ProcessNamedKey
   Else
      Call ProcessChar(this)
   End If
End Sub

Private Sub ProcessChar(this As String)
   Dim vk As Integer
   Dim capped As Boolean
   ' Add input events for single character, taking capitalization
   ' into account.  HiByte will contain the shift state, and LoByte
   ' will contain the key code.
   vk = VkKeyScan(Asc(this))
   capped = CBool(ByteHi(vk) And 1)
   vk = ByteLo(vk)
   Call StuffBuffer(vk, capped)
End Sub

Private Sub ProcessGroup()
   Dim EndPtr As Long
   Dim this As String
   Dim i As Long
   ' Groups of characters are offered together, surrounded by parenthesis,
   ' in order to all be modified by shift key(s).  We need to dig out the
   ' remainder of the group, and process each in turn.
   EndPtr = InStr(m_DatPtr, m_Data, ")")
   ' No need to do anything if endgroup immediateyl follows beginning.
   If EndPtr > (m_DatPtr + 1) Then
      For i = 1 To (EndPtr - m_DatPtr - 1)
         this = Mid$(m_Data, m_DatPtr + i, 1)
         Call ProcessChar(this)
      Next i
      ' Advance data pointer to closing parenthesis.
      m_DatPtr = EndPtr
   End If
End Sub

Private Sub ProcessNamedKey()
   Dim EndPtr As Long
   Dim this As String
   Dim pieces As Variant  '() As String
   Dim repeat As Long
   Dim vk As Integer
   Dim capped As Boolean
   Dim extend As Boolean
   Dim i As Long
   
   ' Groups of characters are offered together, surrounded by braces,
   ' representing a named keystroke.  We need to dig out the actual
   ' name, and optionally the number of times this keystroke is repeated.
   EndPtr = InStr(m_DatPtr, m_Data, "}")
   ' No need to do anything if endgroup immediately follows beginning.
   If EndPtr > (m_DatPtr + 1) Then
      ' Extract group of characters.
      this = Mid$(m_Data, m_DatPtr + 1, EndPtr - m_DatPtr - 1)
         
      ' Break into pieces, if possible.
      pieces = Split(this, " ")
      
      ' Second element, if avail, is number of times to repeat stroke.
      If UBound(pieces) > 0 Then repeat = Val(pieces(1))
      If repeat < 1 Then repeat = 1
      
      ' Attempt to retrieve named keycode, or else retrieve standard code.
      vk = GetNamedKey(CStr(pieces(0)))
      If vk Then
         ' Is this an extended key?
         extend = IsExtendedKey(this)
      Else
         ' Not a standard named key.
         vk = VkKeyScan(Asc(this))
         capped = CBool(ByteHi(vk) And 1)
         vk = ByteLo(vk)
      End If
      
      ' Stuff buffer as many times as required.
      For i = 1 To repeat
         Call StuffBuffer(vk, capped, extend)
      Next i
      
      ' Advance data pointer to closing parenthesis.
      m_DatPtr = EndPtr
   End If
End Sub

Private Sub ProcessShift(shiftkey As String)
   ' Press appropriate shiftkey.
   With m_Events(m_EvtPtr)
      Select Case shiftkey
         Case "+"
            .wVK = vbKeyShift
            m_ShiftFlags = m_ShiftFlags Or vbShiftMask
         Case "^"
            .wVK = vbKeyControl
            m_ShiftFlags = m_ShiftFlags Or vbCtrlMask
         Case "%"
            .wVK = vbKeyMenu
            m_ShiftFlags = m_ShiftFlags Or vbAltMask
      End Select
   End With
   m_EvtPtr = m_EvtPtr + 1

   ' Process next set of data
   Call DoNext
   
   ' Unpress same shiftkey.
   With m_Events(m_EvtPtr)
      Select Case shiftkey
         Case "+"
            .wVK = vbKeyShift
            m_ShiftFlags = m_ShiftFlags And Not vbShiftMask
         Case "^"
            .wVK = vbKeyControl
            m_ShiftFlags = m_ShiftFlags And Not vbCtrlMask
         Case "%"
            .wVK = vbKeyMenu
            m_ShiftFlags = m_ShiftFlags And Not vbAltMask
      End Select
      .dwFlags = KEYEVENTF_KEYUP
   End With
   m_EvtPtr = m_EvtPtr + 1
End Sub

Private Sub StuffBuffer(ByVal vk As Integer, Optional Shifted As Boolean, Optional Extended As Boolean)
   ' Only mess with Shift key if not already pressed.
   If CBool(m_ShiftFlags And vbShiftMask) = False Then
      If Shifted Then
         With m_Events(m_EvtPtr)
            .wVK = vbKeyShift
         End With
         m_EvtPtr = m_EvtPtr + 1
      End If
   End If
   
   ' Press and release this key.
   With m_Events(m_EvtPtr)
      .wVK = vk
      If Extended Then
         .dwFlags = KEYEVENTF_EXTENDEDKEY
      End If
   End With
   m_EvtPtr = m_EvtPtr + 1
   With m_Events(m_EvtPtr)
      .wVK = vk
      .dwFlags = .dwFlags Or KEYEVENTF_KEYUP
   End With
   m_EvtPtr = m_EvtPtr + 1
   
   ' Only mess with Shift key if not already pressed.
   If CBool(m_ShiftFlags And vbShiftMask) = False Then
      If Shifted Then
         With m_Events(m_EvtPtr)
            .wVK = vbKeyShift
            .dwFlags = KEYEVENTF_KEYUP
         End With
         m_EvtPtr = m_EvtPtr + 1
      End If
   End If
End Sub

Private Function ByteHi(ByVal WordIn As Integer) As Byte
   ' Lop off low byte with divide. If less than
   ' zero, then account for sign bit (adding &h10000
   ' implicitly converts to Long before divide).
   If WordIn < 0 Then
      ByteHi = (WordIn + &H10000) \ &H100
   Else
      ByteHi = WordIn \ &H100
   End If
End Function

Private Function ByteLo(ByVal WordIn As Integer) As Byte
   ' Mask off high byte and return low.
   ByteLo = WordIn And &HFF
End Function

Private Function GetNamedKey(this As String) As Integer
   ' Try retrieving from collection
   On Error Resume Next
      GetNamedKey = m_NamedKeys(UCase$(this))
   On Error Resume Next
End Function

Private Function IsExtendedKey(this As String) As Boolean
   Dim nRet As Integer
   ' Try retrieving from collection
   On Error Resume Next
      nRet = m_ExtendedKeys(UCase$(this))
   On Error Resume Next
   IsExtendedKey = (nRet <> 0)
End Function

Private Sub AddKeyString(ByVal KeyCode As Long, KeyName As String, Optional ByVal Extended As Boolean)
   ' Add to collection(s) of named keycode constants.
   m_NamedKeys.Add KeyCode, KeyName
   If Extended Then
      m_ExtendedKeys.Add KeyCode, KeyName
   End If
End Sub

Private Sub BuildCollections()
   ' Reset both collections of known named keys.
   Set m_NamedKeys = New Collection
   Set m_ExtendedKeys = New Collection
   ' The extended-key flag indicates whether the keystroke message
   ' originated from one of the additional keys on the enhanced
   ' keyboard. The extended keys consist of the ALT and CTRL keys
   ' on the right-hand side of the keyboard; the INS, DEL, HOME, END,
   ' PAGE UP, PAGE DOWN, and arrow keys in the clusters to the left
   ' of the numeric keypad; the NUM LOCK key; the BREAK (CTRL+PAUSE)
   ' key; the PRINT SCRN key; and the divide (/) and ENTER keys in
   ' the numeric keypad. The extended-key flag is set if the key is
   ' an extended key.
   AddKeyString vbKeyBack, "BACKSPACE"
   AddKeyString vbKeyBack, "BS"
   AddKeyString vbKeyBack, "BKSP"
   AddKeyString vbKeyPause, "BREAK", True
   AddKeyString vbKeyCapital, "CAPSLOCK"
   AddKeyString vbKeyDelete, "DELETE", True
   AddKeyString vbKeyDelete, "DEL", True
   AddKeyString vbKeyDown, "DOWN", True
   AddKeyString vbKeyEnd, "END", True
   AddKeyString vbKeyReturn, "ENTER"
   AddKeyString vbKeyReturn, "~"
   AddKeyString vbKeyEscape, "ESC"
   AddKeyString vbKeyHelp, "HELP"
   AddKeyString vbKeyHome, "HOME", True
   AddKeyString vbKeyInsert, "INS", True
   AddKeyString vbKeyInsert, "INSERT", True
   AddKeyString vbKeyLeft, "LEFT", True
   AddKeyString vbKeyNumlock, "NUMLOCK", True
   AddKeyString vbKeyPageDown, "PGDN", True
   AddKeyString vbKeyPageUp, "PGUP", True
   AddKeyString vbKeyPause, "PAUSE"
   AddKeyString vbKeyPrint, "PRINT", True
   AddKeyString vbKeySnapshot, "PRTSC", True
   AddKeyString vbKeySnapshot, "PRTSCN", True
   AddKeyString vbKeySnapshot, "PRINTSCRN", True
   AddKeyString vbKeySnapshot, "PRINTSCREEN", True
   AddKeyString vbKeyRight, "RIGHT", True
   AddKeyString vbKeyScrollLock, "SCROLLLOCK"
   AddKeyString vbKeySelect, "SELECT"
   AddKeyString vbKeyTab, "TAB"
   AddKeyString vbKeyUp, "UP", True
   AddKeyString vbKeyF1, "F1"
   AddKeyString vbKeyF2, "F2"
   AddKeyString vbKeyF3, "F3"
   AddKeyString vbKeyF4, "F4"
   AddKeyString vbKeyF5, "F5"
   AddKeyString vbKeyF6, "F6"
   AddKeyString vbKeyF7, "F7"
   AddKeyString vbKeyF8, "F8"
   AddKeyString vbKeyF9, "F9"
   AddKeyString vbKeyF10, "F10"
   AddKeyString vbKeyF11, "F11"
   AddKeyString vbKeyF12, "F12"
   AddKeyString vbKeyF13, "F13"
   AddKeyString vbKeyF14, "F14"
   AddKeyString vbKeyF15, "F15"
   AddKeyString vbKeyF16, "F16"
End Sub

'#If Not VB6 Then
'Private Function Split(ByVal Expression As String, Optional Delimiter As String = " ", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
'   Dim nCount As Long
'   Dim nPos As Long
'   Dim nDelimLen As Long
'   Dim nStart As Long
'   Dim sRet() As String
'
'   ' Special case #1, Limit=0.
'   If Limit = 0 Then
'      ' Return unbound Variant array.
'      Split = Array()
'      Exit Function
'   End If
'
'   ' Special case #2, no delimiter.
'   nDelimLen = Len(Delimiter)
'   If nDelimLen = 0 Then
'      ' Return expression in single-element Variant array.
'      Split = Array(Expression)
'      Exit Function
'   End If
'
'   ' Always start at beginning of Expression.
'   nStart = 1
'
'   ' Find first delimiter instance.
'   nPos = InStr(nStart, Expression, Delimiter, Compare)
'   Do While nPos
'      ' Extract this element into enlarged array.
'      ReDim Preserve sRet(0 To nCount) As String
'      ' Bail if we hit the limit, or increment
'      ' to next search start position.
'      If nCount + 1 = Limit Then
'         sRet(nCount) = Mid$(Expression, nStart)
'         Exit Do
'      Else
'         sRet(nCount) = Mid$(Expression, nStart, nPos - nStart)
'         nStart = nPos + nDelimLen
'      End If
'      ' Increment element counter
'      nCount = nCount + 1
'      ' Find next delimiter instance.
'      nPos = InStr(nStart, Expression, Delimiter, Compare)
'   Loop
'
'   ' Grab last element.
'   ReDim Preserve sRet(0 To nCount) As String
'   sRet(nCount) = Mid$(Expression, nStart)
'
'   ' Assign results and return.
'   Split = sRet
'End Function
'#End If