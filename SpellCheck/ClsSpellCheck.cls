VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------------------
'Name: ClsSpellCheck
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 05/04/2022
'Gratitude: Fafalone, Frank Schüler, Coco, Jose Liza, wqweto, Dilettante, LaVolpe
'---------------------------
DefObj A-Z
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000
Private Const PTR_SIZE                      As Long = 4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
'-------------------------------
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pstring As Long, ByRef pCLSID As GUID) As Long
Private Declare Function DispCallFunc Lib "Oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByRef pvargResult As Variant) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Long) As Long
Private Declare Function SysReAllocString Lib "Oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
'------------------------------------
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As Size) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InsertMenuW Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
'Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
'Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
'Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
'Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
'Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
'Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadString Lib "user32.dll" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type Size
    cx   As Long
    cy   As Long
End Type

Private Type CWPSTRUCT
    lParam  As Long
    wParam  As Long
    Message As Long
    hwnd    As Long
End Type

Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
End Type

Private Type tSUGGESTIONS
    lStart As Long
    lLen As Long
    sInvalid As String
    'cSug As Collection
    CorrectiveAction As CORRECTIVE_ACTION
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const LF_FACESIZE As Long = 32

Private Type CHARFORMAT
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
End Type

Private Const WM_APP            As Long = &H8000&
Private Const WM_COMMAND        As Long = &H111
Private Const WM_DESTROY        As Long = &H2
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_GETTEXT        As Long = &HD
Private Const WM_GETTEXTLENGTH  As Long = &HE
Private Const WM_GETFONT        As Long = &H31
Private Const WM_CHAR           As Long = &H102
Private Const WM_PASTE          As Long = &H302
Private Const WM_CONTEXTMENU    As Long = &H7B
Private Const WM_CREATE         As Long = &H1
Private Const WM_CUT            As Long = &H300
Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_USER           As Long = &H400&
Private Const WM_INITMENUPOPUP  As Long = &H117
Private Const WM_PAINT          As Long = &HF&


Private Const WH_CALLWNDPROC    As Long = 4

Private Const EN_HSCROLL        As Long = &H601
Private Const EN_VSCROLL        As Long = &H602
Private Const EN_UPDATE         As Long = &H400
Private Const EN_KILLFOCUS      As Long = &H200
Private Const EN_SETFOCUS       As Long = &H100
Private Const EN_CHANGE         As Long = &H300
Private Const EM_GETSEL         As Long = &HB0

Private Const EM_POSFROMCHAR    As Long = &HD6
Private Const EM_CHARFROMPOS    As Long = &HD7
Private Const EM_REPLACESEL     As Long = &HC2
Private Const EM_SETSEL         As Long = &HB1
Private Const EM_UNDO           As Long = &HC7
Private Const EM_LINEFROMCHAR   As Long = &HC9
Private Const EM_LINEINDEX      As Long = &HBB
Private Const EM_GETLINECOUNT   As Long = &HBA

Private Const EM_SETEDITSTYLE   As Long = WM_USER + 204&
Private Const EM_SETLANGOPTIONS As Long = WM_USER + 120&
Private Const EM_GETCHARFORMAT  As Long = WM_USER + 58&

Private Const CFM_SIZE          As Long = &H80000000

Private Const MN_GETHMENU       As Long = &H1E1
Private Const MF_STRING         As Long = &H0&
Private Const MF_SEPARATOR      As Long = &H800&

Private Const RDW_INVALIDATE    As Long = &H1

Private Const IMF_SPELLCHECKING As Long = &H800&
Private Const IMF_TKBPREDICTION As Long = &H1000&
Private Const IMF_TKBAUTOCORRECTION As Long = &H2000&
Private Const SES_USECTF As Long = &H10000
Private Const SES_CTFALLOWEMBED As Long = &H200000
Private Const SES_CTFALLOWSMARTTAG As Long = &H400000
Private Const SES_CTFALLOWPROOFING As Long = &H800000

Private Const LOCALE_SNAME       As Long = &H5C

Private Const S_OK As Long = &H0&
Private Const S_FALSE As Long = &H1&
Private Const CLSCTX_INPROC As Long = &H1&
Private Const IID_Release As Long = &H8&
Private Const CC_STDCALL As Long = &H4&

Private Const CLSID_SpellCheckerFactorys As String = "{7AB36653-1796-484B-BDFA-E74F1DB7C1DC}"
Private Const IID_ISpellCheckerFactory As String = "{8E018A9D-2415-4677-BF08-794EA61F94BB}"

Private Enum CORRECTIVE_ACTION
    CORRECTIVE_ACTION_NONE = 0 ' // None - there's no error
    CORRECTIVE_ACTION_GET_SUGGESTIONS = 1 ' // GetSuggestions - the client should show a list of suggestions (obtained through ISpellChecker::Suggest) to the user
    CORRECTIVE_ACTION_REPLACE = 2 ' // Replace - the client should autocorrect the word to the word obtained from ISpellingError::get_Replacement
    CORRECTIVE_ACTION_DELETE = 3 ' // Delete - the client should delete this word
End Enum


Private Enum vtb_Interfaces
    '---=== ISpellingError ===---
    StartIndex = 3
    Length = 4
    CorrectiveAction = 5
    Replacement = 6

    '---=== IEnumString ===---
    'Next_ = 3
    Skip = 4
    Reset = 5
    Clone = 6
    
    '---=== IEnumSpellingError ===---
    Next_ = 3
    
    '---=== IUnknown ===---
    QueryInterface = 0
    AddRef = 1
    Release = 2
    
    '---=== ISpellChecker ===---
    LanguageTag = 3
    Check = 4
    Suggest = 5
    Add = 6
    Ignore = 7
    AutoCorrect = 8
    GetOptionValue = 9
    ptionIds = 10
    ID = 11
    LocalizedName = 12
    add_SpellCheckerChanged = 13
    remove_SpellCheckerChanged = 14
    GetOptionDescription = 15
    ComprehensiveCheck = 16

    '---=== ISpellCheckerFactory ===---
    SupportedLanguages = 3
    IsSupported = 4
    CreateSpellChecker = 5
End Enum

Private STR_DELETE_WORD As String
Private STR_ADD_TO_DICTIONARY As String
Private STR_OMITE_ALL As String
Private STR_ADD_TO_AUTOCORRECT As String

Private pISpellCheckerFactory As Long
Private pISpellChecker As Long
Private m_pSubclassEdit     As IUnknown
Private m_pSubclassParent   As IUnknown
Private m_pHook             As IUnknown
Private mHwndEdit           As Long
Private mHwndParent         As Long
Private DPIFactor           As Long
Private m_LangTag           As String
Private tSUG()              As tSUGGESTIONS
Private ColSug              As Collection
Private lCount              As Long
Private m_StarPoss          As Long
Private bChange             As Boolean
Private bUpdate            As Boolean
Private m_IsRichTextClass   As Boolean
Private m_FreezeChange      As Boolean
Private cOmittedWords       As Collection
Public cSupportedLanguages As Collection

Public Property Get Language() As String
    Language = m_LangTag
End Property

Public Property Let Language(New_Language As String)
    Dim lResult As Long

    If Invoke(pISpellCheckerFactory, IsSupported, StrPtr(New_Language), VarPtr(lResult)) = S_OK Then
        If CBool(lResult) Then
            m_LangTag = New_Language
        Else
            Err.Raise 2, , "Unsupported language."
        End If
    Else
        Err.Raise 100
    End If
    
    If pISpellChecker Then ReleaseInterface pISpellChecker: pISpellChecker = 0

    If mHwndEdit Then
        SpellCheck
        RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
    End If
End Property


Private Sub DrawErrorMarks()
    Dim i As Long
    Dim nFontHeight As Long
    Dim hFont As Long, OldFont As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim PFC As Long
    Dim hDC As Long
    Dim nLine As Long
    Dim LineIndex As Long
    Dim LineCount As Long
    Dim tCF As CHARFORMAT
    Dim ERect As RECT
 
    If lCount = 0 Then Exit Sub

    hDC = GetDC(mHwndEdit)
    
    GetClientRect mHwndEdit, ERect
    
    hFont = SendMessageW(mHwndEdit, WM_GETFONT, 0&, 0&)
    If hFont Then
        OldFont = SelectObject(hDC, hFont)
        nFontHeight = GetTextHeight(hDC, "Áj")
        SelectObject hDC, OldFont
    End If
    
    LineCount = SendMessageW(mHwndEdit, EM_GETLINECOUNT, 0&, 0&)
    
    For i = 0 To UBound(tSUG)
        With tSUG(i)
            
            PFC = SendMessageW(mHwndEdit, EM_POSFROMCHAR, .lStart, ByVal 0&)
            If PFC > 0 Then
                lLeft = WordLo(PFC)
                If hFont = 0 Then
                    nLine = SendMessageW(mHwndEdit, EM_LINEFROMCHAR, .lStart, 0&)
                    LineIndex = SendMessageW(mHwndEdit, EM_LINEINDEX, nLine + 1, 0&)
                    
                    If nLine < LineCount - 1 Then
                        PFC = SendMessageW(mHwndEdit, EM_POSFROMCHAR, LineIndex, ByVal 0&)
                        lTop = WordHi(PFC)
                    Else
                        tCF.dwMask = CFM_SIZE
                        tCF.cbSize = Len(tCF)
                        Call SendMessageW(mHwndEdit, EM_GETCHARFORMAT, LineIndex + 1, ByVal tCF)
                        lTop = WordHi(PFC) + tCF.yHeight / Screen.TwipsPerPixelY
                    End If
                Else
                    lTop = WordHi(PFC) + nFontHeight
                End If
                If lLeft < ERect.Right And lTop < ERect.Bottom Then
                    PFC = SendMessageW(mHwndEdit, EM_POSFROMCHAR, .lStart + .lLen, ByVal 0&)
                    If PFC <> -1 Then
                        lWidth = WordLo(PFC) - lLeft
                        If .CorrectiveAction = CORRECTIVE_ACTION_GET_SUGGESTIONS Then
                            DrawLine hDC, lLeft, lTop, lWidth, vbRed
                        ElseIf .CorrectiveAction = CORRECTIVE_ACTION_REPLACE Then
                            DrawLine hDC, lLeft, lTop, lWidth, vbGreen
                        ElseIf .CorrectiveAction = CORRECTIVE_ACTION_DELETE Then
                            DrawLine hDC, lLeft, lTop, lWidth, vbBlue
                        End If
                    End If
                End If
            End If
        End With
    Next
    
    ReleaseDC mHwndEdit, hDC
End Sub

Private Sub SpellCheck()
    Dim pIEnumSpellingError As Long
    Dim pISpellingError As Long
    Dim pAction As Long
    Dim lStart As Long
    Dim lLen As Long
    Dim lResult As Long
    Dim StrError As String
    Dim sText As String
    Dim lpsz As Long

    Dim i           As Long
    Dim bOmitted    As Boolean
    
    sText = TextBox_GetText(mHwndEdit)
    lCount = 0
    Erase tSUG

    If LenB(sText) < 3 Then Exit Sub
    
    If pISpellChecker = 0 Then
        If Invoke(pISpellCheckerFactory, IsSupported, StrPtr(m_LangTag), VarPtr(lResult)) = S_OK Then
            If Not CBool(lResult) Then
                MsgBox "Language " & m_LangTag & " no Supported"
                Exit Sub
            End If
            
            If Not Invoke(pISpellCheckerFactory, CreateSpellChecker, StrPtr(m_LangTag), VarPtr(pISpellChecker)) = S_OK Then
                MsgBox "Error on Create Spell Checker"
                Terminate
                Exit Sub
            End If
        End If
    End If

    If Invoke(pISpellChecker, Check, StrPtr(sText), VarPtr(pIEnumSpellingError)) = S_OK Then
        Do While Invoke(pIEnumSpellingError, Next_, VarPtr(pISpellingError)) = S_OK
            If Invoke(pISpellingError, CorrectiveAction, VarPtr(pAction)) = S_OK Then
                If Invoke(pISpellingError, StartIndex, VarPtr(lStart)) = S_OK Then
                    If Invoke(pISpellingError, Length, VarPtr(lLen)) = S_OK Then
                    
                        StrError = Mid$(sText, lStart + 1, lLen)
        
                        If pAction = CORRECTIVE_ACTION_REPLACE Then
                            Dim StrRemplace As String
                            Invoke pISpellingError, Replacement, VarPtr(lpsz)
                            StrRemplace = LPWSTRtoStr(lpsz)
                            m_FreezeChange = True
                            Call SendMessageW(mHwndEdit, EM_SETSEL, lStart, ByVal lStart + lLen)
                            Call SendMessageW(mHwndEdit, EM_REPLACESEL, 1, ByVal StrPtr(StrRemplace))
                            m_FreezeChange = False
                            SpellCheck
                            Exit Sub
                        Else
                            bOmitted = False
                            For i = 1 To cOmittedWords.Count
                               If UCase(cOmittedWords.Item(i)) = UCase(StrError) Then
                                   bOmitted = True
                                   Exit For
                               End If
                            Next
                            
                            If bOmitted = False Then
                                ReDim Preserve tSUG(lCount)
                                With tSUG(lCount)
                                    Set ColSug = New Collection
                                    .lStart = lStart
                                    .lLen = lLen
                                    .sInvalid = StrError
                                    .CorrectiveAction = pAction
                                End With
                                lCount = lCount + 1
                            End If
                        End If
                    End If
                End If
            End If
            ReleaseInterface pISpellingError
         Loop
         
         ReleaseInterface pIEnumSpellingError
    End If

    DrawErrorMarks
End Sub

Private Sub DrawLine(hDC As Long, Left As Long, Top As Long, Width As Long, Optional ByVal Color As Long = -1)
    Dim x As Long
    Dim WaveH As Single, WaveW As Single
    Dim i As Long
     
    Dim hPen As Long
    Dim OldPen As Long
  
    If Color <> -1 Then
        If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
        hPen = CreatePen(0, DPIFactor, Color)
        OldPen = SelectObject(hDC, hPen)
    End If
 
    WaveH = 1 * DPIFactor
    WaveW = 2 * DPIFactor
    For x = Left To Left + Width - WaveW Step WaveW
        MoveToEx hDC, x, Top + IIf(i Mod 2 = 0, WaveH, -WaveH), 0
        LineTo hDC, x + WaveW, Top + IIf(i Mod 2 = 0, -WaveH, WaveH)
        i = i + 1
    Next
    
    If hPen Then DeleteObject SelectObject(hDC, OldPen)
    
End Sub

Private Function IndexFromTxtCurPos(ByVal CurPos As Long) As Long
    Dim i As Long
    IndexFromTxtCurPos = -1
    If lCount = 0 Then Exit Function
    For i = 0 To lCount - 1
        If tSUG(i).lStart < CurPos And CurPos < tSUG(i).lStart + tSUG(i).lLen Then
            IndexFromTxtCurPos = i
            Exit Function
        End If
    Next
End Function

Private Function GetTextHeight(ByVal hDC As Long, ByVal sText As String) As Long
   Dim tSize As Size
   Call GetTextExtentPoint32W(hDC, ByVal StrPtr(sText), Len(sText), tSize)
   GetTextHeight = tSize.cy
End Function

Private Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
    SysReAllocString VarPtr(LPWSTRtoStr), lPtr
    If fFree Then Call CoTaskMemFree(lPtr)
End Function

Private Function TextBox_GetText(hwnd As Long) As String
    Dim sText As String
    Dim lLength As Long
    lLength = SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
    sText = Space$(lLength + 1)
    Call SendMessageW(hwnd, WM_GETTEXT, lLength + 1, ByVal StrPtr(sText))
    TextBox_GetText = Left$(sText, lLength)
End Function

Public Function SubclassEditProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    'Debug.Print "Edit", Hex(uMsg); uMsg
    Dim PT As POINTAPI
    Dim lRet As Long
    Dim Index As Long
    Dim ID As Long
    
    Select Case uMsg
       
        Case WM_CONTEXTMENU, WM_INITMENUPOPUP

            GetCursorPos PT
            ScreenToClient hwnd, PT
            
            If m_IsRichTextClass Then
                lRet = SendMessageW(hwnd, EM_CHARFROMPOS, 0, PT)
            Else
                lRet = SendMessageW(hwnd, EM_CHARFROMPOS, 0, ByVal MakeDWord(PT.x, PT.y))
            End If
            
            m_StarPoss = WordLo(lRet)

            If uMsg = WM_INITMENUPOPUP Then
                Call ModifyMenu(wParam)
                Exit Function
            End If
            
            Set m_pHook = InitHookingThunk(WH_CALLWNDPROC, Me, InitAddressOfMethod(Me, 4).CallWndProcHookProc(0, 0, 0, 0))
            SubclassEditProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
            TerminateHookingThunk m_pHook, Me
            Handled = True

        Case WM_HSCROLL, WM_VSCROLL
            RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
            
        Case WM_KEYDOWN
            Select Case (wParam And &HFF&)
                Case vbKeyLeft, vbKeyUp, vbKeyRight, vbKeyDown
                    If bChange Then
                        RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                        SpellCheck
                        bChange = False
                    End If
            End Select
            
        Case WM_CHAR
            Dim KeyAscii As Long
            KeyAscii = (wParam And &HFFFF&)
            If KeyAscii = vbKeySpace Or KeyAscii = vbKeyReturn Then
                SpellCheck
            End If
            
        Case WM_LBUTTONDOWN
            If bChange Then
                RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                SpellCheck
                bChange = False
            End If

        Case WM_PASTE, WM_CUT, EM_UNDO
            SubclassEditProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
            SpellCheck
            Handled = True
           
        Case WM_PAINT ', &HC04E
            SubclassEditProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
            DrawErrorMarks
            Handled = True
            
        Case WM_SETFOCUS
            Set m_pSubclassParent = InitSubclassingThunk(mHwndParent, Me, InitAddressOfMethod(Me, 5).SubclassParentProc(0, 0, 0, 0, 0))
        
        Case WM_KILLFOCUS
            TerminateSubclassingThunk m_pSubclassParent, Me
        
        Case WM_DESTROY
            Terminate
            
        Case &H2111  ' (WM_USER + 7441) anule this msg to prevent RichTextBox flickers
            'RedrawWindow hwnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE
            'DrawErrorMarks
            Handled = True

        Case Else
            If (uMsg And WM_APP) = WM_APP Then
                ID = uMsg Xor WM_APP
            ElseIf uMsg = WM_COMMAND Then
                ID = WordLo(wParam) Xor WM_APP
            End If

            If ID > 100 And ID < 200 Then
                Index = IndexFromTxtCurPos(m_StarPoss)
                If Index > -1 Then
                    
                    If ID = 102 Then
                        m_FreezeChange = True
                        Call SendMessageW(hwnd, EM_SETSEL, tSUG(Index).lStart, ByVal tSUG(Index).lStart + tSUG(Index).lLen + 1)
                        Call SendMessageW(hwnd, EM_REPLACESEL, 1, ByVal 0)
                        m_FreezeChange = False
                        
                    ElseIf ID = 103 Then
                        If Invoke(pISpellChecker, Add, StrPtr(tSUG(Index).sInvalid)) = S_OK Then
                            m_FreezeChange = False
                            SpellCheck
                            RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                            Exit Function
                        Else
                            MsgBox "Error to add"
                        End If
                    ElseIf ID = 104 Then
                        cOmittedWords.Add tSUG(Index).sInvalid
                        
                    ElseIf ID = 105 Then
                        Dim strReplace As String
                        If ColSug.Count > 0 Then
                            strReplace = ColSug.Item(1)
                        End If
                        strReplace = InputBox(STR_ADD_TO_AUTOCORRECT, , strReplace)
                        If Len(strReplace) Then
                            If Invoke(pISpellChecker, AutoCorrect, StrPtr(tSUG(Index).sInvalid), StrPtr(strReplace)) = S_FALSE Then
                                MsgBox "Error add to AutoCorrect"
                            Else
                                SpellCheck
                            End If
                        End If
                    Else
                        m_FreezeChange = True
                        Call SendMessageW(hwnd, EM_SETSEL, tSUG(Index).lStart, ByVal tSUG(Index).lStart + tSUG(Index).lLen)
                        Call SendMessageW(hwnd, EM_REPLACESEL, 1, ByVal StrPtr(ColSug.Item(ID - 110)))
                        m_FreezeChange = False
                    End If
                    
                    
                    SpellCheck
                    RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                End If
            End If
            
    End Select
    
End Function

'*1
'Modify Menu of Editbox By LaVolpe
Public Function CallWndProcHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    Static hMenuOwner As Long
    
    If nCode = 0& Then                      ' must handle
        Dim CWP As CWPSTRUCT
        CopyMemory CWP, ByVal lParam, 16&   ' get structure
        Select Case CWP.Message             ' test a couple messages
        Case WM_CREATE
            Dim CS As CREATESTRUCT          ' window being created
            CopyMemory CS, ByVal CWP.lParam, Len(CS)
            
            If CS.lpszClass = 32768 Then    ' is it #32768 (32768 is Atom)
                hMenuOwner = CWP.hwnd       ' cache window handle
            End If
        Case MN_GETHMENU                    ' we are sending this; ignore it
        Case Else
            If CWP.hwnd = hMenuOwner Then   ' message for #32768?
                If hMenuOwner <> 0 Then     ' got a menu handle yet?
                    Dim hMenu As Long, puMenu As Long
                    hMenu = SendMessageW(hMenuOwner, MN_GETHMENU, 0&, ByVal 0&)
                    If hMenu <> 0 Then      ' if so, add our menu(s)
                        Call ModifyMenu(hMenu)
                        hMenuOwner = 0&
                        TerminateHookingThunk m_pHook, Me
                    End If
                End If
            End If
        End Select
    End If

End Function

Public Function SubclassParentProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
   ' Debug.Print "Parent", uMsg
    Select Case uMsg
        Case WM_COMMAND
            If lParam = mHwndEdit Then
                Select Case WordHi(wParam)
                    Case EN_CHANGE
                        Dim LenText As Long
                        Dim SelStart As Long, lRet As Long
                        LenText = SendMessageW(mHwndEdit, WM_GETTEXTLENGTH, 0&, 0&)
                        lRet = SendMessageW(mHwndEdit, EM_GETSEL, 0&, 0&)
                        SelStart = WordLo(lRet)
                        
                        If SelStart + 2 < LenText Then
                            If m_FreezeChange = False Then
                                If m_IsRichTextClass Then RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                                SpellCheck
                                
                            End If
                        Else
                            If bUpdate Then
                                DrawErrorMarks
                                bUpdate = False
                            End If
                            bChange = True
                        End If

                    Case EN_HSCROLL, EN_VSCROLL, EN_SETFOCUS, EN_KILLFOCUS
                        RedrawWindow mHwndEdit, ByVal 0&, ByVal 0&, RDW_INVALIDATE
                        
                    Case EN_UPDATE
                        bUpdate = True

                End Select
            End If
        Case WM_DESTROY
            Terminate
    End Select
End Function

Private Sub ModifyMenu(hMenu As Long)
    Dim Index As Long, i As Long
    Dim pIEnumString As Long
    Dim lpsz As Long
    Dim lf As Long
    
    Index = IndexFromTxtCurPos(m_StarPoss)
    If Index > -1 Then

                   
        If tSUG(Index).CorrectiveAction = CORRECTIVE_ACTION_GET_SUGGESTIONS Then
        
            If Invoke(pISpellChecker, Suggest, StrPtr(tSUG(Index).sInvalid), VarPtr(pIEnumString)) = S_OK Then
                Set ColSug = New Collection
                Do While Invoke(pIEnumString, Next_, 1&, VarPtr(lpsz), VarPtr(lf)) = S_OK
                    
                   If i = 0 Then InsertMenuW hMenu, 0, MF_SEPARATOR, 1000, StrPtr("-")
                   i = i + 1
                   InsertMenuW hMenu, 0, MF_STRING, WM_APP Or 110 + i, lpsz
                   ColSug.Add LPWSTRtoStr(lpsz)
                Loop
                ReleaseInterface pIEnumString
            End If
            
        ElseIf tSUG(Index).CorrectiveAction = CORRECTIVE_ACTION_DELETE Then
            InsertMenuW hMenu, 0, MF_SEPARATOR, 1000, StrPtr("-")
            InsertMenuW hMenu, 0, MF_STRING, WM_APP Or 102, StrPtr(STR_DELETE_WORD)
            InsertMenuW hMenu, 0, MF_STRING, WM_APP Or 104, StrPtr(STR_OMITE_ALL)
            SetMenuDefaultItem hMenu, 2, 1
            Exit Sub
        End If

        InsertMenuW hMenu, 0, MF_SEPARATOR, 1000, StrPtr("-")
        InsertMenuW hMenu, 0, MF_STRING, WM_APP Or 103, StrPtr(STR_ADD_TO_DICTIONARY)
        InsertMenuW hMenu, 0, MF_STRING, WM_APP Or 104, StrPtr(STR_OMITE_ALL)
        InsertMenuW hMenu, 0, MF_STRING, WM_APP Or 105, StrPtr(STR_ADD_TO_AUTOCORRECT)
        SetMenuDefaultItem hMenu, 2, 1
    End If

    
End Sub

Public Sub Init(hwnd As Long)
    Dim sClassName As String
    Dim lLen As Long
    
    sClassName = String(255, vbNull)
    lLen = GetClassName(hwnd, sClassName, 255)
    If lLen > 0 Then sClassName = Left$(sClassName, lLen)

    Select Case sClassName
        Case "RICHEDIT50W", "INKTEXTBOX"
            'By dilettante
            SendMessageW hwnd, EM_SETLANGOPTIONS, 0, ByVal IMF_SPELLCHECKING Or IMF_TKBPREDICTION Or IMF_TKBAUTOCORRECTION
            SendMessageW hwnd, EM_SETEDITSTYLE, ByVal SES_USECTF Or SES_CTFALLOWEMBED Or SES_CTFALLOWSMARTTAG _
                     Or SES_CTFALLOWPROOFING, SES_USECTF Or SES_CTFALLOWEMBED Or SES_CTFALLOWSMARTTAG Or SES_CTFALLOWPROOFING
            
        Case "Edit", "RichTextWndClass", "ThunderTextBox", "ThunderRT6TextBox"
            m_IsRichTextClass = sClassName = "RichTextWndClass"
            mHwndEdit = hwnd
            mHwndParent = GetParent(mHwndEdit)
            Set m_pSubclassEdit = InitSubclassingThunk(hwnd, Me, InitAddressOfMethod(Me, 5).SubclassEditProc(0, 0, 0, 0, 0))
            SpellCheck
    End Select
End Sub

Public Sub Terminate()
    If pISpellChecker Then ReleaseInterface pISpellChecker
    If pISpellCheckerFactory Then ReleaseInterface pISpellCheckerFactory
    TerminateSubclassingThunk m_pSubclassEdit, Me
    TerminateSubclassingThunk m_pSubclassParent, Me
End Sub

'SubClass and Hooks by wqweto (Vladimir Vissoultchev)
Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As ClsSpellCheck
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitHookingThunk(ByVal idHook As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEN4AV1aLdCQUg8YIgz4AdCqL+oHHTBLeAIvCBVQR3gCri8IFkBHeAKuLwgWgEd4AqzPAq7kJAAAA86WBwkwS3gBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0oM/0IMgWIM/wAAAI0Eyo0EyI1MiDTHAf80JLiJeQTHQQiJRCQEi8ItTBLeAAXEEd4AUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQ/3QkEGoAUf90JBiLD/9RGIlHDItEJBiJOF5fuIAS3gAtcBDeAAUAFAAAwhAAi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1FIsK/3IM/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVCP9CBItCEIXAdFiLCotBLIXAdCpS/9BaiUIIg/gBd0OLClL/UTBahcB1OIsKUmrw/3Ek/1EoWqkAAAAIdSVSM8BQUI1EJARQjUQkBFD/dRT/dRD/dQz/chD/UhRZWFqFyXUTUosK/3UU/3UQ/3UM/3IM/1EgWlBS6Fr///9YXcIQAJA=" ' 13.5.2020 18:24:28
    Const THUNK_SIZE    As Long = 5648
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitHookingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetWindowsHookExA")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "UnhookWindowsHookEx")
        aParams(6) = GetProcAddress(GetModuleHandle("user32"), "CallNextHookEx")
        '--- for IDE protection
        'Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitHookingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, idHook, App.ThreadID, VarPtr(aParams(0)), VarPtr(InitHookingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function TerminateHookingThunk(pHook As IUnknown, pObj As Object) As IUnknown
    If Not pHook Is Nothing Then
        Debug.Assert ThunkPrivateData(pHook, 2) = ObjPtr(pObj)
        ThunkPrivateData(pHook, 2) = 0
        Set pHook = Nothing
    End If
End Function

Private Function InitSubclassingThunk(ByVal hwnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEBAAV1aLdCQUg8YIgz4AdC+L+oHHKBIQAIvCBQwREACri8IFSBEQAKuLwgVYERAAq4vCBYAREACruQkAAADzpYHCKBIQAFJqHP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlM8Crg+8cagBX/3IM/3cM/1IYi0QkGIk4Xl+4XBIQAC1wEBAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GIsKUv9xDP9yDP9RHItUJASLClL/URQzwMIEAJBVi+yLVRj/QgT/QhiLQhg7QgR0b4tCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFr/ShhQUug4////WF3CGAAPHwA=" ' 9.6.2020 13:56:03
    Const THUNK_SIZE    As Long = 492
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback

    If hThunk = 0 Then
        hThunk = pvThunkGlobalData("InitSubclassingThunk")
    End If

    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)             '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)             '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)             '--- 413 = DefSubclassProc ordinal
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
      
        pvThunkGlobalData("InitSubclassingThunk") = hThunk
       
    End If
    lSize = CallWindowProc(hThunk, hwnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function TerminateSubclassingThunk(pSubclass As IUnknown, pObj As Object) As IUnknown
    If Not pSubclass Is Nothing Then
        Debug.Assert ThunkPrivateData(pSubclass, 2) = ObjPtr(pObj)
        ThunkPrivateData(pSubclass, 2) = 0
        Set pSubclass = Nothing
    End If
End Function

Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, PTR_SIZE)
    End If
End Property

Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, lValue, PTR_SIZE)
    End If
End Property

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property

Private Sub Class_Initialize()
    Dim SL As Long
    Dim lpsz As Long
    Dim lf As Long
    Dim StrLang As String
    Dim lResult As Long
    Dim hMod As Long
    Dim sBuffer As String

    DPIFactor = GetWindowsDPI
    Set cOmittedWords = New Collection
    Set cSupportedLanguages = New Collection

    STR_ADD_TO_DICTIONARY = "Add to dictionary"
    STR_OMITE_ALL = "Omitte all"
    STR_DELETE_WORD = "Delete repeated word"
    STR_ADD_TO_AUTOCORRECT = "Add to autocorrect"
    
    
    hMod = LoadLibrary("mshtml.dll")
    
    If hMod Then
        On Error Resume Next
        sBuffer = String(100, 0)
        lResult = LoadString(hMod, 53296, sBuffer, 100)
        STR_ADD_TO_DICTIONARY = Left$(sBuffer, lResult)
        
        sBuffer = String(100, 0)
        lResult = LoadString(hMod, 53297, sBuffer, 100)
        STR_OMITE_ALL = Left$(sBuffer, lResult)
    
        sBuffer = String(100, 0)
        lResult = LoadString(hMod, 53299, sBuffer, 100)
        STR_DELETE_WORD = Left$(sBuffer, lResult)
        On Error GoTo 0
        
        FreeLibrary hMod
    End If

    m_LangTag = fGetLocaleInfo(LOCALE_SNAME)
    
    If InitInterface(pISpellCheckerFactory, CLSID_SpellCheckerFactorys, IID_ISpellCheckerFactory) = False Then
        Err.Raise 1, , "Failed to create interface ISpellCheckerFactory."
        Exit Sub
    End If
    
    If Invoke(pISpellCheckerFactory, SupportedLanguages, VarPtr(SL)) = S_OK Then
        Do While Invoke(SL, Next_, 1&, VarPtr(lpsz), VarPtr(lf)) = S_OK
            StrLang = LPWSTRtoStr(lpsz)
            cSupportedLanguages.Add StrLang
        Loop
        ReleaseInterface SL
    End If

    If Invoke(pISpellCheckerFactory, IsSupported, StrPtr(m_LangTag), VarPtr(lResult)) = S_OK Then
        If CBool(lResult) = False Then
            m_LangTag = cSupportedLanguages.Item(1)
        End If
    End If
End Sub


Private Function fGetLocaleInfo(Valor As Long) As String
   Dim Simbolo As String
   Dim r1 As Long
   Dim r2 As Long
   Dim p As Integer
   Dim Locale As Long
     
   Locale = GetUserDefaultLCID()
   r1 = GetLocaleInfo(Locale, Valor, vbNullString, 0)
   Simbolo = String$(r1, 0)
   r2 = GetLocaleInfo(Locale, Valor, Simbolo, r1)
   p = InStr(Simbolo, Chr$(0))
   If p > 0 Then fGetLocaleInfo = Left$(Simbolo, p - 1)
End Function

Private Function WordHi(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        WordHi = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        WordHi = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

Private Function WordLo(lngValue As Long) As Long
    WordLo = (lngValue And &HFFFF&)
End Function

Private Function MakeDWord(ByVal wLo As Integer, ByVal wHi As Integer) As Long
    MakeDWord = (wHi * 65536) + (wLo And &HFFFF&)
End Function

Private Function GetWindowsDPI() As Double
    Const LOGPIXELSX = 88
    Dim hDC As Long, LPX  As Double
    hDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub Class_Terminate()
    Me.Terminate
End Sub

Private Function InitInterface(ByRef pInterface As Long, ByVal sCLSID As String, ByVal sIID As String) As Boolean
    Dim bRet As Boolean
    Dim tIID As GUID
    Dim tCLSID As GUID
    Dim psIID As Long
    Dim psCLSID As Long
    Dim ptIID As Long
    Dim ptCLSID As Long
    
    bRet = False
    psCLSID = StrPtr(sCLSID)
    If CLSIDFromString(psCLSID, tCLSID) = S_OK Then
        ptCLSID = VarPtr(tCLSID)
        psIID = StrPtr(sIID)
        If CLSIDFromString(psIID, tIID) = S_OK Then
            ptIID = VarPtr(tIID)
            If CoCreateInstance(ptCLSID, 0&, CLSCTX_INPROC, ptIID, pInterface) = S_OK Then
                bRet = True
            End If
        End If
    End If
    InitInterface = bRet
End Function


Private Function Invoke(ByVal pInterface As Long, ByVal eInterfaceFunction As vtb_Interfaces, ParamArray arrParam() As Variant) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, eInterfaceFunction, arrParam)
    End If
End Function

Private Sub ReleaseInterface(ByRef pInterface As Long)
    Dim varRet As Variant
    If pInterface <> 0& Then
        If DispCallFunc(pInterface, IID_Release, CC_STDCALL, _
            vbLong, 0&, 0&, 0&, varRet) = S_OK Then
            pInterface = 0&
        End If
    End If
End Sub

Private Function OleInvoke(ByVal pInterface As Long, ByVal lngCmd As Long, ParamArray arrParam() As Variant) As Variant
    Dim lngRet As Long
    Dim lngItem As Long
    Dim lngCount As Long
    Dim varRet As Variant
    Dim varParam As Variant
    Dim olePtr(10) As Long
    Dim oleTyp(10) As Integer

    If pInterface <> 0& Then
        If UBound(arrParam) >= 0 Then
            varParam = arrParam
            If IsArray(varParam) Then varParam = varParam(0)
            lngCount = UBound(varParam)
            For lngItem = 0 To lngCount
                oleTyp(lngItem) = VarType(varParam(lngItem))
                olePtr(lngItem) = VarPtr(varParam(lngItem))
            Next
        End If

        lngRet = DispCallFunc(pInterface, lngCmd * 4, CC_STDCALL, vbLong, lngItem, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), varRet)
            
        If lngRet <> S_OK Then
            varRet = S_FALSE
        End If
    End If

    OleInvoke = varRet
End Function


