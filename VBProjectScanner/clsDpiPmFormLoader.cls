VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDpiPmFormLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------------

' This class must be declared in a module and created in Sub Main

' This class creates a Windows hook to catch creation of new forms within your project.
' The hook remains active for the life of the project. If your project will never load
'   another form after Sub Main executes, you can set this class instance to Nothing

' When a new form is created, the hook will catch its creation. It subclasses the new form
' in order to respond to its WM_NCCREATE and WM_CREATE messages.
'   WM_NCCREATE is subclassed to apply EnableNonClientDpiScaling for non-permonitorV2
'   WM_CREATE is subclassed to ensure the form is not aborted before its creation is finished
'       if creation succeeded, then a new clsDpiPmAssist class will be created,
'       and sent to the form if that form implements IDpiPmAssistant
'-----------------------------------------------------------------------------------------------


Private Declare Function GetWindowDpiAwarenessContext Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetThreadDpiAwarenessContext Lib "user32.dll" (ByVal dpiContext As Long) As Long
Private Declare Function AreDpiAwarenessContextsEqual Lib "user32.dll" (ByVal contextA As Long, ByVal contextB As Long) As Long
Private Declare Function EnableNonClientDpiScaling Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDpiForSystem Lib "user32.dll" () As Long
Private Declare Function SetThreadDpiHostingBehavior Lib "user32.dll" (ByVal hContext As Long) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32.dll" (ByRef VersionInformation As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Enum DpiAwarenessContextEnum
    dpiAWARE_INVALID = 0
    dpiUNAWARE = -1
    dpiSYSTEM_AWARE = -2
    dpiPER_MONITOR_AWARE_V1 = -3
    dpiPER_MONITOR_AWARE_V2 = -4
    dpiUNAWARE_GDISCALED = -5
End Enum

Public Enum WinCapsEnum         ' notes, comments
    wcpXPandLower = 0           ' = system aware only
    wcpWinVista                 ' >= then allows system awareness (virtualized unawareness)
    wcpWin8_1                   ' >= then allows per-monitor awareness (virtualized unawareness)
    wcpWin10_Threshold          ' = then Win10, version 1507 or 1511
    wcpWin10_1607               ' >= then O/S allows non-client per-monitor scaling (v1607+)
    wcpWin10_1703               ' >= then O/S allows V2 per-monitor awareness (v1703+)
    wcpWin10_1709
    wcpWin10_1803               ' >= then O/S allows mixed-mode hosting (v1803+)
    wcpWin10_1809               ' >= then O/S allows GDIscaling context dynamically (v1809+)
    wcpWin10_1903
    wcpWin10_1909
    wcpWin10_2004
    wcpMask_Version = &HFFFFFF  ' high byte reserved
    wcpNonClientScaling = wcpWin10_1607     ' system supports this option
    wcpPMcapable = wcpWin8_1                ' system supports this option
    wcpPMv2Capable = wcpWin10_1703          ' system supports this option
    wcpMixedHostingCapable = wcpWin10_1803  ' system supports this option
    wcpGDIscalingCapable = wcpWin10_1809    ' system supports this option @ runtime
End Enum

Const WM_CREATE = &H1&
Const WM_NCCREATE = &H81&
Const WS_CHILD As Long = &H40000000
Const WS_EX_MDICHILD = &H40&

Dim oSubclass As Object         ' used to trap WM_CREATE/WM_NCCREATE messages
Dim oHook As Object             ' used to listen for form creation
Dim m_Caps As WinCapsEnum       ' system version flags
Dim m_ProjSysDPI As Long        ' project's actual DPI

' Thunk-related APIs
Private Declare Function thunkCreateDispTypeInfo Lib "oleaut32.dll" Alias "CreateDispTypeInfo" (ByVal pidata As Long, ByVal lcid As Long, ByRef pptinfo As IUnknown) As Long
Private Declare Function thunkCreateStdDispatch Lib "oleaut32.dll" Alias "CreateStdDispatch" (ByVal ptinfo As Long, ByVal pvThis As Long, ByVal pptinfo As Long, ByVal ppunkStdDisp As Long) As Long
Private Declare Function thunkCoTaskMemAlloc Lib "ole32.dll" Alias "CoTaskMemAlloc" (ByVal cb As Long) As Long
Private Declare Sub thunkCoTaskMemFree Lib "ole32.dll" Alias "CoTaskMemFree" (ByVal pv As Long)
Private Declare Sub thunkCopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function thunkVirtualQuery Lib "kernel32.dll" Alias "VirtualQuery" (ByRef lpAddress As Any, ByRef lpBuffer As Any, ByVal dwLength As Long) As Long
Private Declare Function thunkGetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function thunkGetProcAddress Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByRef lpProcName As Any) As Long
Private Declare Function thunkHeapCreate Lib "kernel32.dll" Alias "HeapCreate" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function thunkHeapAlloc Lib "kernel32.dll" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function thunkHeapDestroy Lib "kernel32.dll" Alias "HeapDestroy" (ByVal hHeap As Long) As Long
Private Declare Function thunkGetClassName Lib "user32.dll" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function thunkGetWindowThreadProcessId Lib "user32.dll" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function thunkGetWindow Lib "user32.dll" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function thunkFindWindow Lib "user32.dll" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function thunkFindWindowEx Lib "user32.dll" Alias "FindWindowExW" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Long, ByVal lpsz2 As Long) As Long
Private Declare Function thunkCreateWindowEx Lib "user32.dll" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function thunkSetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function thunkGetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function thunkRemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function thunkVarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (var() As Any) As Long
Private Declare Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal pdwRefData As Long) As Long
Private Declare Function SetWindowSubclassXP Lib "comctl32.dll" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal pdwRefData As Long) As Long

Public Sub Activate()

    ' Must be called to initialize hook listener

    If Not oHook Is Nothing Then Exit Sub   ' currently activated

    Const WH_CBT = 5&
    Const HCBT_CREATEWND = 3&
    
    ' create a subclasser we will use when new forms are loaded
    Set oSubclass = CreateTasker_Subclass(Me, 1)
    oSubclass.FilterAdd 0, Array(WM_NCCREATE, WM_CREATE) ' only listen for these msgs
    
    ' create a hook to listen for new forms being created & desired message
    Set oHook = CreateTasker_Hook(Me, 2, WH_CBT)
    oHook.FilterAdd 0, 1, HCBT_CREATEWND ' only listen for this event
    
End Sub

Public Property Get WindowCaps() As WinCapsEnum
    ' you can use this in Sub Main, if desired, should you be setting your
    ' project's DPI awareness at runtime via APIs. NOT RECOMMENDED
    
    ' returns the Windows version and DPI capabilities of the O/S, if system >= Win8.1
    ' if property returns zero, then system is less than Win8.1 or not manifested for Win8.1+
    ' (WindowsCaps And wcpMask_Version) = wcpWin8_1 when system is Win8.1
    ' (WindowsCaps And wcpMask_Version) >= wcpWin10_1607 returns whether system is at least Win10,v1607
    ' (WindowsCaps And wcpMask_Version) >= wcpMask_PMv2Capable returns whether system supports per-monitor v2
    
    WindowCaps = m_Caps

End Property

Public Function SetThreadDpiAwareness(ByVal Context As DpiAwarenessContextEnum, _
                             Optional ByVal AllowMixedModeHosting As Boolean = False) As Long
                             
    ' Call this function before you show your form in Sub Main, only if you want to
    ' change the awareness context for the form about to be loaded.
    ' NOT RECOMMENDED
    
    ' If this function is called, call it again after the form is loaded, passing
    '   the value this function returned from the first call, i.e.,
    ' ---------------------------------------------------------
    '   Sub Main()
    '       Dim hContext As Long
    '       Set cFormLoader = New clsDpiPmFormLoader
    '       hContext = cFormLoader.SetThreadDpiAwareness(dpiUNAWARE_GDISCALED)
    '       Form1.Show
    '       cFormLoader.SetThreadDpiAwareness hContext
    '   End Sub
    
    ' NOTE: You should not mix DPI awareness contexts in the same project unless the
    '   operating system is at least Win10,v1703. Otherwise, whenever thread awareness
    '   changes, each displayed form is rescaled to the new context and every new form
    '   will be assigned the new context. Win10,v1703 fixes that by making contexts
    '   truly per-top-level-window, whereas v1607 and lower are per-thread and since
    '   VB is single-threaded; changes apply to all windows/forms.
    
    Const DPI_HOSTING_BEHAVIOR_MIXED = 1&
    Const DPI_HOSTING_BEHAVIOR_DEFAULT = 0&
    
    If Context <> 0 And (m_Caps And wcpMask_Version) >= wcpWin10_1607 Then
        If Context = dpiUNAWARE_GDISCALED Then
            If (m_Caps And wcpMask_Version) >= wcpGDIscalingCapable Then Context = dpiUNAWARE
        ElseIf Context = dpiPER_MONITOR_AWARE_V2 Then
            If (m_Caps And wcpMask_Version) >= wcpPMv2Capable Then Context = dpiPER_MONITOR_AWARE_V1
        End If
        SetThreadDpiAwareness = SetThreadDpiAwarenessContext(Context)
    End If
    If (m_Caps And wcpMask_Version) >= wcpMixedHostingCapable Then
        If AllowMixedModeHosting = True Then
            SetThreadDpiHostingBehavior DPI_HOSTING_BEHAVIOR_MIXED
        Else
            SetThreadDpiHostingBehavior DPI_HOSTING_BEHAVIOR_DEFAULT
        End If
    End If

End Function

Public Property Get DpiForProject() As Long
    ' Returns the DPI that project loaded into
    ' This may be different than: 1440/Screen.TwipsPerPixelX
    DpiForProject = m_ProjSysDPI
End Property

Private Function CreateTasker_Hook(ByRef Host As Object, ByVal HostOrdinal As Long, _
                            ByVal HookType As Long, _
                            Optional ByVal InitializeHook As Boolean = True, _
                            Optional ByVal UseUnicode As Boolean = False, _
                            Optional ByVal AddIDEsafety As Boolean = True) As Object

    ' Params
    '   Host :: VB code object that is hosting/declaring this tasker
    '   HostOrdinal :: Last private method is #1, second to last is #2, etc
    '   HookType :: standard Windows hook type constant
    '   InitializeHook :: starts the hook immediately else you call oTasker.Refresh
    '   UseUnicode :: if true, start hook with SetWindowsHookExW else SetWindowsHookExA
    '   AddIDEsafety :: should always be true & ignored if project is compiled
    
    ' See readme file(s) for full details.
    
    ' Quick summary of tasker methods (* = default method)
    ' ------------------------------------------------------------------------------------
    ' CallNextInChain(nCode, wParam, lParam) :: calls API CallNextHookEx
    '   parameters are per MSDN documentation for specific hook type
    ' FilterAdd(fType, pOrdinal, pItem) :: creates/appends filter messages
    '   fType: 0 = inclusive filter, non-zero = exclusive filter
    '   pOrdinal: 1 = nCode filter, 2 = wParam filter, 3 = lParam filter
    '   pItem: one value then pass that value else pass: Array(value1, value2, etc)
    '   fType & pOrdinal ignored if filter already exists
    ' FilterClear :: removes entire filter
    ' HookType :: hook type used to install the hook
    ' IsFiltered: returns non-zero if events are filtered (see FilterAdd)
    ' Pause :: prevents VB from receiving callback messages; hook remains installed
    ' Refresh :: allows callback messages to continue to VB or reinstalls the hook
    '*State :: returns zero if paused else non-zero
    ' Tag :: read/write property, numeric values only
    ' Unhook :: uninstall the hook
    ' ------------------------------------------------------------------------------------
    
    ' Sample callback function for the Host
    ' ------------------------------------------------------------------------------------
    ' Private Function myHookProc(ByVal TaskerObject As Object, _
    '                             ByRef EatMessage As Boolean, _
    '                             ByVal nCode As Long, ByVal wParam As Long, _
    '                             ByVal lParam As Long) As Long
    '
    '       function return value ignored unless EatMessage is returned as True
    '
    ' End Function
    ' ------------------------------------------------------------------------------------
    
    Dim z_Sc() As Long, arrSafeArray(0 To 5) As Long
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, taskHwnd As Long
    Dim hTLBstructs As Long, lEbMode As Long
    Dim oIUnk As stdole.IUnknown, hMod As VBA.Collection
    Dim oITypeInfo As stdole.IUnknown, pITypeInfo As Long
    Dim hHeapCommon As Long, hFactory As Long
    
    Const FACTORY_PROPNAME = "Factory_Hooks.v1"
    Const THUNK_SIZE As Long = 470              ' size of the thunk array (in longs)
    Const DATA_SIZE As Long = 12                ' Tasker data size (in longs)
    Const THUNK_UnicodeMask As Long = &H10000
    Const THUNK_StateUnhooked As Long = 2

    GoSub InitDataObjects ' Gosub? Just wanted to declutter top half of routine
        
    '/// set up SafeArray overlay onto allocated memory == used so we can update memory directly
    arrSafeArray(0) = 1: arrSafeArray(1) = 4 ' dimensions & bytes per array element
    
    If hThunk = 0 Then                          ' else using existing thunk
        
        ' note to self: heap allocates non-zeroized memory. Ensure no missing z_Sc() elements
        ' below else the memory address will be undefined when it should be zero
        hThunk = pvHeapManage(taskHwnd, hHeapCommon, THUNK_SIZE * 4) ' allocate thunk bytes
        Debug.Assert hThunk <> 0
        If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
        
        'Debug.Print "hook thunk: "; hThunk; " data section: "; hData; " hWnd: "; taskHwnd
                                            
        arrSafeArray(3) = hThunk                ' address to 1st array element data
        arrSafeArray(4) = THUNK_SIZE            ' number elements in array
        thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), VarPtr(arrSafeArray(0)), 4&
    
        z_Sc(0) = &H105BE: z_Sc(1) = &HE800&: z_Sc(2) = &H815E0000: z_Sc(3) = &HAEE&: z_Sc(4) = &HF689C300
        z_Sc(37) = &H8BE58960: z_Sc(38) = &H458B245D: z_Sc(39) = &H74C08528: z_Sc(40) = &H2C7D834B: z_Sc(41) = &H8B457400: z_Sc(42) = &H480B0448: z_Sc(43) = &HC0F98108: z_Sc(44) = &H75000000: z_Sc(45) = &HC78813E: z_Sc(46) = &H46000000: z_Sc(47) = &H88B3575: z_Sc(48) = &HC75C985: z_Sc(49) = &H892C558B: z_Sc(50) = &H443FF1A: z_Sc(51) = &H31EBC031: z_Sc(52) = &H8508438B: z_Sc(53) = &H811C74C0: z_Sc(54) = &H20400F9: z_Sc(55) = &HFF147500: z_Sc(56) = &H75FF2C75: z_Sc(57) = &H8B5028: z_Sc(58) = &H15EB10FF: z_Sc(59) = &H70057B8
        z_Sc(60) = &H8B0EEB80: z_Sc(61) = &HC72C45: z_Sc(62) = &H0&: z_Sc(63) = &H4002B8: z_Sc(64) = &H1C458980: z_Sc(65) = &HCC261: z_Sc(66) = &H245C8B53: z_Sc(67) = &H4438B08: z_Sc(68) = &H4438940: z_Sc(69) = &H4C25B: z_Sc(70) = &H245C8B53: z_Sc(71) = &H4438B08: z_Sc(72) = &H4438948: z_Sc(73) = &H75C0855B: z_Sc(74) = &HE5896025: z_Sc(75) = &H83245D8B: z_Sc(76) = &H74042C7B: z_Sc(77) = &H2C43C718: z_Sc(78) = &H4&: z_Sc(79) = &HFFFEBFE8: z_Sc(80) = &H39E856FF: z_Sc(81) = &HE8000004: z_Sc(82) = &H410&: z_Sc(83) = &H4C2615E: z_Sc(84) = &H768D00: z_Sc(85) = &H245C8B53: z_Sc(86) = &H2C438B08: z_Sc(87) = &H7A9&: z_Sc(88) = &HB8077500: z_Sc(89) = &H1&
        z_Sc(90) = &HE88303EB: z_Sc(91) = &H4C25B02: z_Sc(92) = &H768D00: z_Sc(93) = &H245C8B53: z_Sc(94) = &HC438B08: z_Sc(95) = &H4C25B: z_Sc(96) = &H245C8B53: z_Sc(97) = &H24448B08: z_Sc(98) = &HC43890C: z_Sc(99) = &HC25BC031: z_Sc(100) = &HF6890008: z_Sc(101) = &H8960C031: z_Sc(102) = &H245D8BE5: z_Sc(103) = &HFFFE5FE8: z_Sc(104) = &HD9E856FF: z_Sc(105) = &H5E000003: z_Sc(106) = &H4C261: z_Sc(107) = &H8B53C031: z_Sc(108) = &H8B08245C: z_Sc(109) = &HC2F72C53: z_Sc(110) = &H7&: z_Sc(111) = &HE2810D75: z_Sc(112) = &H10000: z_Sc(113) = &H8901CA83: z_Sc(114) = &HC25B2C53: z_Sc(115) = &HF6890004: z_Sc(116) = &H8960C031: z_Sc(117) = &H245D8BE5: z_Sc(118) = &H3C7E8: z_Sc(119) = &H4C26100
        z_Sc(120) = &H768D00: z_Sc(121) = &H245C8B53: z_Sc(122) = &H20438B08: z_Sc(123) = &H4C25B: z_Sc(124) = &H245C8B53: z_Sc(125) = &H28438B08: z_Sc(126) = &H4C25B: z_Sc(127) = &H70057B8: z_Sc(128) = &HE5896080: z_Sc(129) = &HE8245D8B: z_Sc(130) = &HFFFFFDF4: z_Sc(131) = &H2C458B56: z_Sc(132) = &H840FC085: z_Sc(133) = &HC7&: z_Sc(134) = &HFFFFFCA9: z_Sc(135) = &HBC850FFF: z_Sc(136) = &HFC000000: z_Sc(137) = &HE830558D: z_Sc(138) = &H44C&: z_Sc(139) = &H147D83: z_Sc(140) = &HA9840F: z_Sc(141) = &H7D830000: z_Sc(142) = &H8740314: z_Sc(143) = &H840FC985: z_Sc(144) = &H9B&: z_Sc(145) = &H7B83C031: z_Sc(146) = &H17750028: z_Sc(147) = &HC12C458B: z_Sc(148) = &H8B5104E0: z_Sc(149) = &HC985284D
        z_Sc(150) = &H1B90574: z_Sc(151) = &H41000000: z_Sc(152) = &HC759C809: z_Sc(153) = &H5E1C45: z_Sc(154) = &H45F78007: z_Sc(155) = &H200014: z_Sc(156) = &H89437400: z_Sc(157) = &H8BC289D6: z_Sc(158) = &HE841287B: z_Sc(159) = &H374&: z_Sc(160) = &H1E74C085: z_Sc(161) = &H49287B89: z_Sc(162) = &HC801078B: z_Sc(163) = &H45890789: z_Sc(164) = &H74D2851C: z_Sc(165) = &H4578903: z_Sc(166) = &H7C8DC829: z_Sc(167) = &HF3560887: z_Sc(168) = &H7D815EA5: z_Sc(169) = &H200314: z_Sc(170) = &H89347400: z_Sc(171) = &HE8C931F7: z_Sc(172) = &H340&: z_Sc(173) = &HC68929EB: z_Sc(174) = &H2B9&: z_Sc(175) = &H287B8B00: z_Sc(176) = &H32FE8: z_Sc(177) = &H74C08500: z_Sc(178) = &H287B8916: z_Sc(179) = &H89410F8B
        z_Sc(180) = &H74F6850F: z_Sc(181) = &H4778903: z_Sc(182) = &H48F5489: z_Sc(183) = &H581C4D89: z_Sc(184) = &H1CC261: z_Sc(185) = &H8960C031: z_Sc(186) = &H245D8BE5: z_Sc(187) = &H85287B8B: z_Sc(188) = &H501274FF: z_Sc(189) = &H892873FF: z_Sc(190) = &H1E82843: z_Sc(191) = &HFFFFFFFD: z_Sc(192) = &H7096&: z_Sc(193) = &H4C26100: z_Sc(194) = &H768D00: z_Sc(195) = &H8960C031: z_Sc(196) = &H245D8BE5: z_Sc(197) = &H22C43F7: z_Sc(198) = &H75000000: z_Sc(199) = &H3075FF15: z_Sc(200) = &HFF2C75FF: z_Sc(201) = &HE8512875: z_Sc(202) = &HFFFFFCD4: z_Sc(203) = &H7496FF: z_Sc(204) = &HC2610000: z_Sc(205) = &HF6890010: z_Sc(206) = &H8960C031: z_Sc(207) = &H245D8BE5: z_Sc(208) = &H1C7B83: z_Sc(209) = &HB5E83F75
        z_Sc(210) = &H8BFFFFFC: z_Sc(211) = &H8C86&: z_Sc(212) = &H680C6A00: z_Sc(213) = &H40000: z_Sc(214) = &HFF0870FF: z_Sc(215) = &H5496&: z_Sc(216) = &H74C08500: z_Sc(217) = &H1C438920: z_Sc(218) = &H89BA00C6: z_Sc(219) = &H84BA0158: z_Sc(220) = &H1000004: z_Sc(221) = &H540C6F2: z_Sc(222) = &H65089B8: z_Sc(223) = &HA40C766: z_Sc(224) = &H4589E0FF: z_Sc(225) = &H4C2611C: z_Sc(226) = &H768D00: z_Sc(227) = &H8960C031: z_Sc(228) = &H245D8BE5: z_Sc(229) = &H8928458B: z_Sc(230) = &HC2611043: z_Sc(231) = &HF6890008: z_Sc(232) = &H8960C031: z_Sc(233) = &H2C5D8BE5: z_Sc(234) = &HFFFC53E8: z_Sc(235) = &H75FF53FF: z_Sc(236) = &H6896FF24: z_Sc(237) = &H8B000000: z_Sc(238) = &H9E3284B: z_Sc(239) = &HFF51006A
        z_Sc(240) = &H7096&: z_Sc(241) = &H1C4B8B00: z_Sc(242) = &H685115E3: z_Sc(243) = &H40000: z_Sc(244) = &H8C868B: z_Sc(245) = &H70FF0000: z_Sc(246) = &H5896FF08: z_Sc(247) = &H8B000000: z_Sc(248) = &H28B0853: z_Sc(249) = &H850FF52: z_Sc(250) = &HFF53006A: z_Sc(251) = &H7096&: z_Sc(252) = &H88BE8B00: z_Sc(253) = &H8B000000: z_Sc(254) = &HF89490F: z_Sc(255) = &H7E75C985: z_Sc(256) = &H8B04578B: z_Sc(257) = &H50FF5202: z_Sc(258) = &H8C868B08: z_Sc(259) = &H8D000000: z_Sc(260) = &HFF501047: z_Sc(261) = &H96FF2475: z_Sc(262) = &H60&: z_Sc(263) = &H77FF006A: z_Sc(264) = &H7096FF08: z_Sc(265) = &H6A000000: z_Sc(266) = &H96FF5700: z_Sc(267) = &H70&: z_Sc(268) = &H8B147589: z_Sc(269) = &H8CBE&
        z_Sc(270) = &H8478B00: z_Sc(271) = &H8B1C4589: z_Sc(272) = &HF89490F: z_Sc(273) = &H1374C985: z_Sc(274) = &HC4835961: z_Sc(275) = &H685210: z_Sc(276) = &H50000400: z_Sc(277) = &H58A2FF51: z_Sc(278) = &H8D000000: z_Sc(279) = &HFF500C47: z_Sc(280) = &H96FF2475: z_Sc(281) = &H60&: z_Sc(282) = &HFF57006A: z_Sc(283) = &H7096&: z_Sc(284) = &H83596100: z_Sc(285) = &H515010C4: z_Sc(286) = &H5CA2FF: z_Sc(287) = &HC2610000: z_Sc(288) = &HF6890010: z_Sc(289) = &H8960C031: z_Sc(290) = &HFFD389E5: z_Sc(291) = &H6DE80443: z_Sc(292) = &H56FFFFFB: z_Sc(293) = &H63E8&: z_Sc(294) = &H74F68500: z_Sc(295) = &H4FE830C: z_Sc(296) = &HB5E83775: z_Sc(297) = &HEB000000: z_Sc(298) = &H19AE830: z_Sc(299) = &H458D0000
        z_Sc(300) = &H1475891C: z_Sc(301) = &H2C75FF50: z_Sc(302) = &HFF2875FF: z_Sc(303) = &H458D2475: z_Sc(304) = &H75FF5014: z_Sc(305) = &H1473FF18: z_Sc(306) = &H8B1853FF: z_Sc(307) = &H8B501845: z_Sc(308) = &H850FF00: z_Sc(309) = &H75147539: z_Sc(310) = &H2C75FF18: z_Sc(311) = &HFF2875FF: z_Sc(312) = &H73FF2475: z_Sc(313) = &HFC458B24: z_Sc(314) = &H7490FF: z_Sc(315) = &H45890000: z_Sc(316) = &H53038B1C: z_Sc(317) = &H580850FF: z_Sc(318) = &HCC261: z_Sc(319) = &H832C738B: z_Sc(320) = &HF68507E6: z_Sc(321) = &H458B5475: z_Sc(322) = &H84808BFC: z_Sc(323) = &H85000000: z_Sc(324) = &HFF0E74C0: z_Sc(325) = &H1F883D0: z_Sc(326) = &H1BE0774: z_Sc(327) = &HEB000000: z_Sc(328) = &H287B8B39: z_Sc(329) = &H3274FF85
        z_Sc(330) = &H778B0F8B: z_Sc(331) = &H20C6F704: z_Sc(332) = &H74000000: z_Sc(333) = &H30FE830F: z_Sc(334) = &H458B057D: z_Sc(335) = &H8B08EB28: z_Sc(336) = &H3EB2C45: z_Sc(337) = &HFC24458B: z_Sc(338) = &H8303E683: z_Sc(339) = &HAFF208C7: z_Sc(340) = &HF6830574: z_Sc(341) = &H8303EB02: z_Sc(342) = &H90C301F6: z_Sc(343) = &HB8FC558B: z_Sc(344) = &H3A0&: z_Sc(345) = &H6850D001: z_Sc(346) = &H7D0&: z_Sc(347) = &H8C828B53: z_Sc(348) = &HFF000000: z_Sc(349) = &H92FF0470: z_Sc(350) = &H64&: z_Sc(351) = &H768DC3: z_Sc(352) = &HA92C438B: z_Sc(353) = &H2&: z_Sc(354) = &H251775: z_Sc(355) = &H83000100: z_Sc(356) = &H438902C8: z_Sc(357) = &H2473FF2C: z_Sc(358) = &HFFFC458B: z_Sc(359) = &H8090&
        z_Sc(360) = &HF689C300: z_Sc(361) = &HA92C438B: z_Sc(362) = &H7&: z_Sc(363) = &H1A94474: z_Sc(364) = &H74000000: z_Sc(365) = &H250A&: z_Sc(366) = &H43890001: z_Sc(367) = &HE833742C: z_Sc(368) = &HFFFFFA3C: z_Sc(369) = &H90B6FF: z_Sc(370) = &H6A0000: z_Sc(371) = &HFF1C73FF: z_Sc(372) = &H78BA2073: z_Sc(373) = &HA9000000: z_Sc(374) = &H10000: z_Sc(375) = &HC2830374: z_Sc(376) = &H1614FF04: z_Sc(377) = &HA74C085: z_Sc(378) = &H81244389: z_Sc(379) = &H2C63&: z_Sc(380) = &H90C30001: z_Sc(381) = &H5152C031: z_Sc(382) = &H1375C985: z_Sc(383) = &H4374FF85: z_Sc(384) = &H458B5751: z_Sc(385) = &H7090FFFC: z_Sc(386) = &H31000000: z_Sc(387) = &H4134EBFF: z_Sc(388) = &H1A74FF85: z_Sc(389) = &HD101178B
        z_Sc(390) = &H5102E1C1: z_Sc(391) = &HFC458B57: z_Sc(392) = &H7090FF: z_Sc(393) = &HC0850000: z_Sc(394) = &HC7891974: z_Sc(395) = &HE1C115EB: z_Sc(396) = &H458B5102: z_Sc(397) = &H6C90FFFC: z_Sc(398) = &H85000000: z_Sc(399) = &H890474C0: z_Sc(400) = &H59C78938: z_Sc(401) = &HF689C35A: z_Sc(402) = &H68&: z_Sc(403) = &HC06846: z_Sc(404) = &H6A0000: z_Sc(405) = &H2040068: z_Sc(406) = &H187D8D00: z_Sc(407) = &H8D50078D: z_Sc(408) = &H50042444: z_Sc(409) = &H8510438B: z_Sc(410) = &H8B0375C0: z_Sc(411) = &H8B500843: z_Sc(412) = &H8310FF00: z_Sc(413) = &H90C310C4: z_Sc(414) = &H8FE8&: z_Sc(415) = &H14458900: z_Sc(416) = &H840FC085: z_Sc(417) = &H82&: z_Sc(418) = &H7403F883: z_Sc(419) = &H74D2857D
        z_Sc(420) = &H66D68979: z_Sc(421) = &H8366068B: z_Sc(422) = &H6C7501F8: z_Sc(423) = &H85104E8B: z_Sc(424) = &H816774C9: z_Sc(425) = &HFFF9&: z_Sc(426) = &H8B5D7F00: z_Sc(427) = &H768B0446: z_Sc(428) = &H74F6850C: z_Sc(429) = &H10F88353: z_Sc(430) = &HF2890474: z_Sc(431) = &HFF314CEB: z_Sc(432) = &HFFFF2FE8: z_Sc(433) = &H74C085FF: z_Sc(434) = &HFCD2313F: z_Sc(435) = &H66068B66: z_Sc(436) = &H75FFFCA9: z_Sc(437) = &HF883661E: z_Sc(438) = &H660F7403: z_Sc(439) = &H7502F883: z_Sc(440) = &H468B6612: z_Sc(441) = &HC0BF0F08: z_Sc(442) = &H468B03EB: z_Sc(443) = &H17048908: z_Sc(444) = &H8304C283: z_Sc(445) = &HD4E210C6: z_Sc(446) = &HE9C1D189: z_Sc(447) = &H8904E302: z_Sc(448) = &HE807EBFA: z_Sc(449) = &HFFFFFEEC
        z_Sc(450) = &H90C3C931: z_Sc(451) = &H8B66C031: z_Sc(452) = &H4000A902: z_Sc(453) = &H5750000: z_Sc(454) = &HEB08C283: z_Sc(455) = &H4000350D: z_Sc(456) = &H528B0000: z_Sc(457) = &HCF88308: z_Sc(458) = &HC3DE274: z_Sc(459) = &H74000020: z_Sc(460) = &H20033D22: z_Sc(461) = &H1B740000: z_Sc(462) = &H7403F883: z_Sc(463) = &H2F88316: z_Sc(464) = &HC0310474: z_Sc(465) = &H8B660FEB: z_Sc(466) = &HD2BF0F12: z_Sc(467) = &H3B8&: z_Sc(468) = &H8B02EB00: z_Sc(469) = &HC312

        '/// build Tasker VTable -- same order as sent to pvInitFactory
        z_Sc(5) = ((hThunk Xor &H80000000) + &H94) Xor &H80000000 ' IUnknown:QueryInterface(QI)
        z_Sc(6) = ((hThunk Xor &H80000000) + &H108) Xor &H80000000 ' IUnknown:AddRef
        z_Sc(7) = ((hThunk Xor &H80000000) + &H118) Xor &H80000000 ' IUnknown:Release
        z_Sc(8) = ((hThunk Xor &H80000000) + &H154) Xor &H80000000 ' State Get
        z_Sc(9) = ((hThunk Xor &H80000000) + &H174) Xor &H80000000 ' Tag Get
        z_Sc(10) = ((hThunk Xor &H80000000) + &H180) Xor &H80000000 ' Tag Let
        z_Sc(11) = ((hThunk Xor &H80000000) + &H194) Xor &H80000000 ' Unhook
        z_Sc(12) = ((hThunk Xor &H80000000) + &H1AC) Xor &H80000000 ' Pause
        z_Sc(13) = ((hThunk Xor &H80000000) + &H1D0) Xor &H80000000 ' Refresh
        z_Sc(14) = ((hThunk Xor &H80000000) + &H1E4) Xor &H80000000 ' HookType Get
        z_Sc(15) = ((hThunk Xor &H80000000) + &H1F0) Xor &H80000000 ' IsFiltered
        z_Sc(16) = ((hThunk Xor &H80000000) + &H1FC) Xor &H80000000 ' FilterAdd
        z_Sc(17) = ((hThunk Xor &H80000000) + &H2E4) Xor &H80000000 ' FilterRemove
        z_Sc(18) = ((hThunk Xor &H80000000) + &H30C) Xor &H80000000 ' SendNextInChain
        z_Sc(19) = ((hThunk Xor &H80000000) + &H338) Xor &H80000000 ' CreateStub
        z_Sc(20) = ((hThunk Xor &H80000000) + &H38C) Xor &H80000000 ' WrapperObject Let
        '/// build API list -- order dictated by listing in source script
        z_Sc(21) = thunkGetProcAddress(hMod("k32"), ByVal "HeapAlloc")
        z_Sc(22) = thunkGetProcAddress(hMod("k32"), ByVal "HeapFree")
        z_Sc(23) = thunkGetProcAddress(hMod("k32"), ByVal "HeapDestroy")
        z_Sc(24) = thunkGetProcAddress(hMod("u32"), ByVal "RemovePropA")
        z_Sc(25) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
        z_Sc(26) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
        z_Sc(27) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemAlloc")
        z_Sc(28) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemRealloc")
        z_Sc(29) = thunkGetProcAddress(hMod("u32"), ByVal "CallNextHookEx")
        z_Sc(30) = thunkGetProcAddress(hMod("u32"), ByVal "SetWindowsHookExA")
        z_Sc(31) = thunkGetProcAddress(hMod("u32"), ByVal "SetWindowsHookExW")
        z_Sc(32) = thunkGetProcAddress(hMod("u32"), ByVal "UnhookWindowsHookEx")
        z_Sc(33) = lEbMode
        z_Sc(34) = hFactory
        z_Sc(35) = hHeapCommon
        z_Sc(36) = App.ThreadID
        Set hMod = Nothing
        
        ' /// create the Factory object's data section
        arrSafeArray(3) = hFactory              ' address to 1st array element data
        arrSafeArray(4) = 4                     ' number accessed elements in array
        thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), VarPtr(arrSafeArray(0)), 4&
        z_Sc(0) = 1                             ' reference count
        z_Sc(1) = pITypeInfo                    ' ObjPtr(ITypeInfo) - reusable for new taskers
        z_Sc(2) = hTLBstructs                   ' data used by ITypeInfo
        z_Sc(3) = hThunk                        ' this thunk - reusable for new taskers
        thunkCopyMemory oITypeInfo, 0&, 4&      ' zombie the interface, thunk owns it now
        thunkSetProp taskHwnd, FACTORY_PROPNAME, hFactory
    End If
    
    ' /// create the tasker object's data section
    arrSafeArray(3) = hData                     ' address to 1st array element data
    arrSafeArray(4) = DATA_SIZE                 ' number elements in array
    thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), VarPtr(arrSafeArray(0)), 4&
    z_Sc(0) = ((hThunk Xor &H80000000) + 20) Xor &H80000000
    z_Sc(1) = 1                                 ' ref counter
    z_Sc(2) = 0                                 ' Dispatch::IUnknown
    z_Sc(3) = 0                                 ' Tag
    z_Sc(4) = 0                                 ' Wrapper object
    z_Sc(5) = ObjPtr(Host)                      ' Callback object
    z_Sc(6) = HostOrdinal                       ' Callback function address
    z_Sc(7) = 0                                 ' AddressOf (filled in by tasker later)
    z_Sc(8) = HookType                          ' hook type
    z_Sc(9) = 0                                 ' hook ID
    z_Sc(10) = 0                                ' filter
    z_Sc(11) = THUNK_StateUnhooked              ' state
    If UseUnicode = True Then z_Sc(11) = z_Sc(11) Or THUNK_UnicodeMask
    thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), 0&, 4& ' remove overlay
    thunkCopyMemory oIUnk, hData, 4             ' create Tasker's IUnknown
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, pITypeInfo, VarPtr(oIDispatch)
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                     ' if failure, tasker unwinds on exit
        ' copy the IDispatch pointer to the object data
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 8) Xor &H80000000, oIDispatch, 4
        Set CreateTasker_Hook = oIUnk
        If CreateTasker_Hook.CreateStub = 0 Then
            Set CreateTasker_Hook = Nothing ' failure
        ElseIf InitializeHook = True Then
            CreateTasker_Hook.Refresh
            If CreateTasker_Hook.State = THUNK_StateUnhooked Then Set CreateTasker_Hook = Nothing ' failure
        End If
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
InitDataObjects:
    ' perform sanity checks on passed parameters
    If Host Is Nothing Then Exit Function
    HostOrdinal = pvGetAddressOf(Host, HostOrdinal)
    If HostOrdinal = 0 Then Exit Function
    
    Call pvInitFactory(taskHwnd, -1)
    If taskHwnd = 0 Then Exit Function
    hFactory = thunkGetProp(taskHwnd, FACTORY_PROPNAME)
    If hFactory = 0 Then            ' first time thunk is being created
        
        hFactory = thunkCoTaskMemAlloc(Len(FACTORY_PROPNAME) + 17) ' memory for factory
        Debug.Assert (hFactory <> 0)
        If hFactory = 0 Then Exit Function
        
        Dim bData() As Byte
        ' create array of 4 Longs followed by null terminated ANSI property name
        bData() = StrConv(String$(16, 0) & FACTORY_PROPNAME & vbNullChar, vbFromUnicode)
        thunkCopyMemory ByVal hFactory, bData(0), UBound(bData) + 1
        Erase bData()
        
        ' call local routine to create a temporary typelib for adhoc IDispatch interface
        Set oITypeInfo = pvInitFactory(taskHwnd, hTLBstructs, _
                    2, "State", 2, "Tag", 4, "Tag", "Value", 1, "Unhook", _
                    1, "Pause", 1, "Refresh", 2, "HookType", 2, "IsFiltered", _
                    1, "FilterAdd", "fType", "pOrdinal", "!pItem", _
                    1, "FilterClear", 1, "CallNextInChain", "nCode", "wParam", "lParam", _
                    1, "CreateStub", 4, "WrapperObject", "pObjPtr")
        pITypeInfo = ObjPtr(oITypeInfo)
        Debug.Assert (pITypeInfo <> 0)
        If pITypeInfo = 0 Then GoTo AbortCleanup
        ' modules used by the thunk
        Set hMod = New Collection
        hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
        hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
        hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
        If AddIDEsafety = True Then
            On Error Resume Next
            Debug.Print 1 / 0
            If Err Then
                Err.Clear: On Error GoTo 0
                hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
                ' note: if next line errors, change vba6.dll to vba5.dll
                lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
            End If
            On Error GoTo 0
        End If
    Else                    ' existing thunk - get needed values & increment ref count
        thunkCopyMemory hThunk, ByVal ((hFactory Xor &H80000000) + 12) Xor &H80000000, 4
        thunkCopyMemory pITypeInfo, ByVal ((hFactory Xor &H80000000) + 4) Xor &H80000000, 4
        thunkCopyMemory hData, ByVal hFactory, 4: hData = hData + 1
        thunkCopyMemory ByVal hFactory, hData, 4     ' increment counter
    End If
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData <> 0 Then Return

AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hFactory <> 0 Then thunkCoTaskMemFree hFactory
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If
    If hHeapCommon <> 0 Then pvHeapManage taskHwnd, hHeapCommon, 0
End Function

Friend Function CreateTasker_Subclass(ByRef Host As Object, ByVal HostOrdinal As Long, _
                            Optional ByVal initHwnd As Long = 0, _
                            Optional ByVal AddIDEsafety As Boolean = True) As Object
                            
    ' NOTE: This method should not be called from outside code
    ' Exception: Is called from clsDpiPmAssist so method is exposed as Friend

    ' Params
    '   Host :: VB code object that is hosting/declaring this tasker
    '   HostOrdinal :: Last private method is #1, second to last is #2, etc
    '   initHwnd :: window handle to subclass immediately. Use oTasker.AddItem later as needed
    '   AddIDEsafety :: should always be true & ignored if project is compiled
    
    ' See readme file(s) for full details.
    
    ' Quick summary of tasker methods (* = default method)
    ' ------------------------------------------------------------------------------------
    ' AddItem(pItem) :: append a window for subclassing
    '   pItem: one hWnd then pass that hWnd else pass: Array(hWnd1, hWnd2, etc)
    ' CallDefWindowProcA(hWnd, uMsg, wParam, lParam) :: calls API DefWindowProcA
    ' CallDefWindowProcW(hWnd, uMsg, wParam, lParam) :: calls API DefWindowProcW
    ' CallNextInChain(hWnd, uMsg, wParam, lParam) :: calls API DefSubclassProc
    ' FilterAdd(fType, pItem) :: creates/appends filter messages
    '   fType: 0 = inclusive filter, non-zero = exclusive filter
    '   pItem: one value then pass that value else pass: Array(value1, value2, etc)
    '   fType ignored if filter already exists
    ' FilterClear :: removes entire filter
    ' GetItems(pBuffer, pCount) :: returns list of subclassed hWnds
    '   pBuffer: VarPtr(LongArray(x)). Pass 0 to return count
    '   pCount: number of hWnds to return, ignored if pCount=0
    ' IsFiltered: returns non-zero if events are filtered (see FilterAdd)
    ' IsManaged(pItem) :: returns non-zero if subclassed by tasker
    ' Pause :: prevents VB from receiving callback messages; subclassing remains active
    ' Refresh :: allows callback messages to continue to VB
    ' RemoveItem(pItem) :: unsubclasses a window subclassed by the tasker
    '   pItem: one hWnd then pass that hWnd else pass: Array(hWnd1, hWnd2, etc)
    '         passing -1 will unsubclass all hWnds
    '*State :: returns -1 if paused else number of subclassed hWnds
    ' Tag :: read/write property, numeric values only
    ' TagItem(pItem) :: return custom tag value applied to subclassed hWnd
    ' ------------------------------------------------------------------------------------
    
    ' Sample callback function for the Host
    ' ------------------------------------------------------------------------------------
    ' Private Function mySubclassProc(ByVal TaskerObject As Object, _
    '                             ByRef EatMessage As Boolean, _
    '                             ByVal hWnd As Long, ByVal uMsg As Long, _
    '                             ByVal wParam As Long, ByVal lParam As Long, _
    '                             ByVal hwndTag As Long) As Long
    '
    '       function return value ignored unless EatMessage is returned as True
    '
    ' End Function
    ' ------------------------------------------------------------------------------------
    
    Dim z_Sc() As Long, arrSafeArray(0 To 5) As Long
    Dim hThunk As Long, hData As Long, hTag As Long
    Dim oIDispatch As Long, taskHwnd As Long
    Dim hTLBstructs As Long, lEbMode As Long
    Dim oIUnk As stdole.IUnknown, hMod As VBA.Collection
    Dim oITypeInfo As stdole.IUnknown, pITypeInfo As Long
    Dim hHeapCommon As Long, hFactory As Long
    
    Const FACTORY_PROPNAME = "Factory_Subclasser.v1"
    Const THUNK_SIZE As Long = 584              ' size of the thunk array (in longs)
    Const DATA_SIZE As Long = 11                ' Tasker data size (in longs)
    
    GoSub InitDataObjects ' Gosub? Just wanted to declutter top half of routine
        
    '/// set up SafeArray overlay onto allocated memory == used so we can update memory directly
    arrSafeArray(0) = 1: arrSafeArray(1) = 4 ' dimensions & bytes per array element
    
    If hThunk = 0 Then                          ' else using existing thunk
        
        ' note to self: heap allocates non-zeroized memory. Ensure no missing z_Sc() elements
        ' below else the memory address will be undefined when it should be zero
        hThunk = pvHeapManage(taskHwnd, hHeapCommon, THUNK_SIZE * 4) ' allocate thunk bytes
        Debug.Assert hThunk <> 0
        If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
        
        ' Debug.Print "subclass thunk: "; hThunk; " data section: "; hData; " hWnd: "; taskHwnd
                                            
        arrSafeArray(3) = hThunk                ' address to 1st array element data
        arrSafeArray(4) = THUNK_SIZE            ' number elements in array
        thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), VarPtr(arrSafeArray(0)), 4&
    
        z_Sc(0) = &H106BE: z_Sc(1) = &HE800&: z_Sc(2) = &H815E0000: z_Sc(3) = &HAEE&: z_Sc(4) = &HF689C300
        z_Sc(47) = &H8BE58960: z_Sc(48) = &H458B245D: z_Sc(49) = &H74C08528: z_Sc(50) = &H2C7D834B: z_Sc(51) = &H8B457400: z_Sc(52) = &H480B0448: z_Sc(53) = &HC0F98108: z_Sc(54) = &H75000000: z_Sc(55) = &HC78813E: z_Sc(56) = &H46000000: z_Sc(57) = &H88B3575: z_Sc(58) = &HC75C985: z_Sc(59) = &H892C558B
        z_Sc(60) = &H443FF1A: z_Sc(61) = &H31EBC031: z_Sc(62) = &H8508438B: z_Sc(63) = &H811C74C0: z_Sc(64) = &H20400F9: z_Sc(65) = &HFF147500: z_Sc(66) = &H75FF2C75: z_Sc(67) = &H8B5028: z_Sc(68) = &H15EB10FF: z_Sc(69) = &H70057B8: z_Sc(70) = &H8B0EEB80: z_Sc(71) = &HC72C45: z_Sc(72) = &H0&: z_Sc(73) = &H4002B8: z_Sc(74) = &H1C458980: z_Sc(75) = &HCC261: z_Sc(76) = &H245C8B53: z_Sc(77) = &H443FF08: z_Sc(78) = &H5B04438B: z_Sc(79) = &H900004C2: z_Sc(80) = &H245C8B53: z_Sc(81) = &H44BFF08: z_Sc(82) = &H5B04438B: z_Sc(83) = &H850FC085: z_Sc(84) = &H8D&: z_Sc(85) = &H8BE58960: z_Sc(86) = &H43FF245D: z_Sc(87) = &HFE9EE804: z_Sc(88) = &H43C7FFFF: z_Sc(89) = &H224&
        z_Sc(90) = &H546E800: z_Sc(91) = &H4BFF0000: z_Sc(92) = &H44B8B04: z_Sc(93) = &H851C4D89: z_Sc(94) = &H8B6575C9: z_Sc(95) = &H28B0853: z_Sc(96) = &H850FF52: z_Sc(97) = &HE82873FF: z_Sc(98) = &H564&: z_Sc(99) = &HB0BE8B: z_Sc(100) = &HE8530000: z_Sc(101) = &H558&: z_Sc(102) = &H89490F8B: z_Sc(103) = &H75C9850F: z_Sc(104) = &H4578B3F: z_Sc(105) = &HFF52028B: z_Sc(106) = &H77FF0850: z_Sc(107) = &H53EE808: z_Sc(108) = &H868B0000: z_Sc(109) = &HB4&: z_Sc(110) = &H5204508B: z_Sc(111) = &H5010478D: z_Sc(112) = &H7096FF52: z_Sc(113) = &H57000000: z_Sc(114) = &H523E8: z_Sc(115) = &H5CB85A00: z_Sc(116) = &H1000005: z_Sc(117) = &H326A50F0: z_Sc(118) = &H96FF5256: z_Sc(119) = &H74&
        z_Sc(120) = &H4C261: z_Sc(121) = &H245C8B53: z_Sc(122) = &H2443F708: z_Sc(123) = &H1&: z_Sc(124) = &HFFB80774: z_Sc(125) = &HEBFFFFFF: z_Sc(126) = &H1C438B0C: z_Sc(127) = &H574C085: z_Sc(128) = &H1B8&: z_Sc(129) = &H4C25B00: z_Sc(130) = &H768D00: z_Sc(131) = &H245C8B53: z_Sc(132) = &HC438B08: z_Sc(133) = &H4C25B: z_Sc(134) = &H245C8B53: z_Sc(135) = &H24448B08: z_Sc(136) = &HC43890C: z_Sc(137) = &HC25BC031: z_Sc(138) = &HF6890008: z_Sc(139) = &H8960C031: z_Sc(140) = &H245D8BE5: z_Sc(141) = &HFFFDC7E8: z_Sc(142) = &H8CBE8BFF: z_Sc(143) = &H85000000: z_Sc(144) = &H8D1874FF: z_Sc(145) = &H8B501C45: z_Sc(146) = &H45332845: z_Sc(147) = &HE0B85024: z_Sc(148) = &H1000004: z_Sc(149) = &H75FF50F0
        z_Sc(150) = &H61D7FF28: z_Sc(151) = &H900008C2: z_Sc(152) = &H8960C031: z_Sc(153) = &H245D8BE5: z_Sc(154) = &H8528558B: z_Sc(155) = &HE82874D2: z_Sc(156) = &HFFFFFD8C: z_Sc(157) = &H463E8: z_Sc(158) = &H74C08500: z_Sc(159) = &H2C75FF1A: z_Sc(160) = &HD831D089: z_Sc(161) = &H4E0B850: z_Sc(162) = &HF0010000: z_Sc(163) = &H96FF5250: z_Sc(164) = &H90&: z_Sc(165) = &H611C4589: z_Sc(166) = &H90000CC2: z_Sc(167) = &H245C8B53: z_Sc(168) = &H2443C708: z_Sc(169) = &H1&: z_Sc(170) = &HC25BC031: z_Sc(171) = &HF6890004: z_Sc(172) = &H245C8B53: z_Sc(173) = &H89C03108: z_Sc(174) = &HC25B2443: z_Sc(175) = &HF6890004: z_Sc(176) = &H98BA&: z_Sc(177) = &H3A2E900: z_Sc(178) = &HF6890000: z_Sc(179) = &H84BA&
        z_Sc(180) = &H396E900: z_Sc(181) = &HF6890000: z_Sc(182) = &H88BA&: z_Sc(183) = &H38AE900: z_Sc(184) = &HF6890000: z_Sc(185) = &H8BE58960: z_Sc(186) = &H558B245D: z_Sc(187) = &HFD0EE828: z_Sc(188) = &HE5E8FFFF: z_Sc(189) = &H89000003: z_Sc(190) = &HC2611C45: z_Sc(191) = &HF6890008: z_Sc(192) = &H245C8B53: z_Sc(193) = &H20438B08: z_Sc(194) = &H4C25B: z_Sc(195) = &H8BE58960: z_Sc(196) = &HE9E8245D: z_Sc(197) = &H8DFFFFFC: z_Sc(198) = &H6A2855: z_Sc(199) = &H467E8: z_Sc(200) = &H74D28500: z_Sc(201) = &H850A8B29: z_Sc(202) = &H891774C9: z_Sc(203) = &H8F148BD7: z_Sc(204) = &H3A6E851: z_Sc(205) = &HC0850000: z_Sc(206) = &HFDE80575: z_Sc(207) = &H59000002: z_Sc(208) = &HE857EBE2: z_Sc(209) = &H3A8&
        z_Sc(210) = &H891C4B8B: z_Sc(211) = &HC2611C4D: z_Sc(212) = &HF6890014: z_Sc(213) = &H8BE58960: z_Sc(214) = &HA1E8245D: z_Sc(215) = &H8DFFFFFC: z_Sc(216) = &H6A2855: z_Sc(217) = &H41FE8: z_Sc(218) = &H74D28500: z_Sc(219) = &H850A8B3B: z_Sc(220) = &H892974C9: z_Sc(221) = &H1F983D7: z_Sc(222) = &H7F830D7F: z_Sc(223) = &H775FF04: z_Sc(224) = &H32FE8: z_Sc(225) = &H8B15EB00: z_Sc(226) = &HE8518F14: z_Sc(227) = &H34C&: z_Sc(228) = &H574C085: z_Sc(229) = &H277E8: z_Sc(230) = &HEBE25900: z_Sc(231) = &H34EE857: z_Sc(232) = &H4B8B0000: z_Sc(233) = &H1C4D891C: z_Sc(234) = &H14C261: z_Sc(235) = &H8960C031: z_Sc(236) = &H245D8BE5: z_Sc(237) = &H8520438B: z_Sc(238) = &HE81274C0: z_Sc(239) = &HFFFFFC40
        z_Sc(240) = &H32AE850: z_Sc(241) = &H43C70000: z_Sc(242) = &H20&: z_Sc(243) = &H4C26100: z_Sc(244) = &H768D00: z_Sc(245) = &H8BE58960: z_Sc(246) = &H21E8245D: z_Sc(247) = &H8DFFFFFC: z_Sc(248) = &H16A2C55: z_Sc(249) = &H39FE8: z_Sc(250) = &H147D8300: z_Sc(251) = &HC7747400: z_Sc(252) = &H1C45&: z_Sc(253) = &HA8B0000: z_Sc(254) = &H4D74C985: z_Sc(255) = &H85207B8B: z_Sc(256) = &H8B1575FF: z_Sc(257) = &HC0852845: z_Sc(258) = &H1B80574: z_Sc(259) = &H40000000: z_Sc(260) = &H89044289: z_Sc(261) = &H48EB2053: z_Sc(262) = &H5E1C45C7: z_Sc(263) = &H41800700: z_Sc(264) = &H2DAE852: z_Sc(265) = &H855A0000: z_Sc(266) = &HC71E74C0: z_Sc(267) = &H1C45&: z_Sc(268) = &H7B890000: z_Sc(269) = &H728D5620
        z_Sc(270) = &H49078B08: z_Sc(271) = &HF89C101: z_Sc(272) = &H8877C8D: z_Sc(273) = &HA5F3C129: z_Sc(274) = &HA1E8525E: z_Sc(275) = &H83000002: z_Sc(276) = &H75001C7D: z_Sc(277) = &H8BC9310E: z_Sc(278) = &HFF85207B: z_Sc(279) = &HF8B0274: z_Sc(280) = &H611C4D89: z_Sc(281) = &H900018C2: z_Sc(282) = &H70057B8: z_Sc(283) = &HE5896080: z_Sc(284) = &H8B245D8B: z_Sc(285) = &H7D831C4B: z_Sc(286) = &H5750028: z_Sc(287) = &HEB1C4D89: z_Sc(288) = &H75C98547: z_Sc(289) = &H1C4D8905: z_Sc(290) = &H458B3EEB: z_Sc(291) = &H1F8832C: z_Sc(292) = &HC839367C: z_Sc(293) = &HC889027E: z_Sc(294) = &H6A2875FF: z_Sc(295) = &H6A535000: z_Sc(296) = &H24048D01: z_Sc(297) = &HFB56E850: z_Sc(298) = &H8CB8FFFF: z_Sc(299) = &H1000008
        z_Sc(300) = &HB6FF50F0: z_Sc(301) = &HB8&: z_Sc(302) = &H9C96FF: z_Sc(303) = &H4C8B0000: z_Sc(304) = &H4D890C24: z_Sc(305) = &H14C4831C: z_Sc(306) = &HCC261: z_Sc(307) = &H8960C031: z_Sc(308) = &H245D8BE5: z_Sc(309) = &H8928458B: z_Sc(310) = &HC2611043: z_Sc(311) = &HF6890008: z_Sc(312) = &H8BE58960: z_Sc(313) = &H45333445: z_Sc(314) = &HFFC38924: z_Sc(315) = &HDE80443: z_Sc(316) = &H56FFFFFB: z_Sc(317) = &HBFE8&: z_Sc(318) = &H75F68500: z_Sc(319) = &H182E83A: z_Sc(320) = &H458D0000: z_Sc(321) = &H75FF501C: z_Sc(322) = &H3075FF38: z_Sc(323) = &HFF2C75FF: z_Sc(324) = &H75FF2875: z_Sc(325) = &H14758924: z_Sc(326) = &H5014458D: z_Sc(327) = &HFF1875FF: z_Sc(328) = &H53FF1473: z_Sc(329) = &H18458B18
        z_Sc(330) = &H674C085: z_Sc(331) = &HFF008B50: z_Sc(332) = &H75390850: z_Sc(333) = &HFF187514: z_Sc(334) = &H75FF3075: z_Sc(335) = &H2875FF2C: z_Sc(336) = &H8B2475FF: z_Sc(337) = &H96FFFC75: z_Sc(338) = &H98&: z_Sc(339) = &H5E1C4589: z_Sc(340) = &HFF53038B: z_Sc(341) = &HC2610850: z_Sc(342) = &HF6890018: z_Sc(343) = &HC24548B: z_Sc(344) = &H89E58960: z_Sc(345) = &H75FF52D6: z_Sc(346) = &H7896FF24: z_Sc(347) = &H8B000000: z_Sc(348) = &HB4BE&: z_Sc(349) = &H8478B00: z_Sc(350) = &H8B1C4589: z_Sc(351) = &H15E3490F: z_Sc(352) = &H59610F89: z_Sc(353) = &H5210C483: z_Sc(354) = &H4000068: z_Sc(355) = &HFF515000: z_Sc(356) = &H68A2&: z_Sc(357) = &HC478D00: z_Sc(358) = &H2475FF50: z_Sc(359) = &H7096FF
        z_Sc(360) = &HE8570000: z_Sc(361) = &H148&: z_Sc(362) = &HC4835961: z_Sc(363) = &HFF515010: z_Sc(364) = &H6CA2&: z_Sc(365) = &H768D00: z_Sc(366) = &H24738BFC: z_Sc(367) = &H3975F685: z_Sc(368) = &H8BFC558B: z_Sc(369) = &HAC92&: z_Sc(370) = &H74D28500: z_Sc(371) = &H83D2FF0E: z_Sc(372) = &H77401F8: z_Sc(373) = &H1BE&: z_Sc(374) = &H8B1EEB00: z_Sc(375) = &HFF85207B: z_Sc(376) = &HF8B1774: z_Sc(377) = &H8B04778B: z_Sc(378) = &HC7832845: z_Sc(379) = &H74AFF208: z_Sc(380) = &H2F68305: z_Sc(381) = &HF68303EB: z_Sc(382) = &H287D8301: z_Sc(383) = &H560D7502: z_Sc(384) = &H8BFC758B: z_Sc(385) = &H5E82455: z_Sc(386) = &H5E000000: z_Sc(387) = &H768DC3: z_Sc(388) = &H31D08952: z_Sc(389) = &HE0B850D8
        z_Sc(390) = &H1000004: z_Sc(391) = &HFF5250F0: z_Sc(392) = &H9496&: z_Sc(393) = &HC0855A00: z_Sc(394) = &H4BFF0F74: z_Sc(395) = &H73FF521C: z_Sc(396) = &H96FF5228: z_Sc(397) = &H70&: z_Sc(398) = &HF689C35A: z_Sc(399) = &H89006A52: z_Sc(400) = &H50D831D0: z_Sc(401) = &H4E0B8: z_Sc(402) = &H50F00100: z_Sc(403) = &H9096FF52: z_Sc(404) = &H5A000000: z_Sc(405) = &H1174C085: z_Sc(406) = &H521C43FF: z_Sc(407) = &H73FF016A: z_Sc(408) = &H96FF5228: z_Sc(409) = &HA8&: z_Sc(410) = &HF689C35A: z_Sc(411) = &H24048958: z_Sc(412) = &HF98AE856: z_Sc(413) = &H148BFFFF: z_Sc(414) = &H74D28516: z_Sc(415) = &HE2FF5E03: z_Sc(416) = &H900010C2: z_Sc(417) = &H68&: z_Sc(418) = &HC06846: z_Sc(419) = &H6A0000
        z_Sc(420) = &H2040068: z_Sc(421) = &H187D8D00: z_Sc(422) = &H8D50078D: z_Sc(423) = &H50042444: z_Sc(424) = &H8510438B: z_Sc(425) = &H8B0375C0: z_Sc(426) = &H8B500843: z_Sc(427) = &H8310FF00: z_Sc(428) = &H90C310C4: z_Sc(429) = &H1C7B83: z_Sc(430) = &H6A531E74: z_Sc(431) = &H24048D00: z_Sc(432) = &H88CB850: z_Sc(433) = &HF0010000: z_Sc(434) = &HB8B6FF50: z_Sc(435) = &HFF000000: z_Sc(436) = &H9C96&: z_Sc(437) = &H8C48300: z_Sc(438) = &H768DC3: z_Sc(439) = &HD285C031: z_Sc(440) = &HFF520C74: z_Sc(441) = &HFF522873: z_Sc(442) = &HA496&: z_Sc(443) = &H90C35A00: z_Sc(444) = &H74FF006A: z_Sc(445) = &H96FF0824: z_Sc(446) = &H80&: z_Sc(447) = &H900004C2: z_Sc(448) = &HE351C031: z_Sc(449) = &HFF85412C
        z_Sc(450) = &H78B1574: z_Sc(451) = &HE1C1C101: z_Sc(452) = &HFF575102: z_Sc(453) = &H8096&: z_Sc(454) = &H75C08500: z_Sc(455) = &HC112EB12: z_Sc(456) = &HFF5102E1: z_Sc(457) = &H7C96&: z_Sc(458) = &H74C08500: z_Sc(459) = &H89388904: z_Sc(460) = &H90C359C7: z_Sc(461) = &H8B66C031: z_Sc(462) = &HA96602: z_Sc(463) = &H66107440: z_Sc(464) = &H8B400035: z_Sc(465) = &H83660852: z_Sc(466) = &HE8740CF8: z_Sc(467) = &H8B08EA83: z_Sc(468) = &H89660852: z_Sc(469) = &H10E1C1C1: z_Sc(470) = &H8366C809: z_Sc(471) = &H277403F8: z_Sc(472) = &H200C3D66: z_Sc(473) = &H3D661D74: z_Sc(474) = &H17742003: z_Sc(475) = &H2F88366: z_Sc(476) = &HC0310474: z_Sc(477) = &HBF0F11EB: z_Sc(478) = &H3B866D2: z_Sc(479) = &H3108EB00
        z_Sc(480) = &H6604EBC0: z_Sc(481) = &HC32000B8: z_Sc(482) = &H571C45C7: z_Sc(483) = &HE8800700: z_Sc(484) = &HFFFFFFA0: z_Sc(485) = &H85144589: z_Sc(486) = &HE4840FC0: z_Sc(487) = &HFC000000: z_Sc(488) = &H3F88366: z_Sc(489) = &HB1840F: z_Sc(490) = &HD2850000: z_Sc(491) = &HD4840F: z_Sc(492) = &H8B660000: z_Sc(493) = &HF983660A: z_Sc(494) = &HC7850F01: z_Sc(495) = &H8B000000: z_Sc(496) = &HF981104A: z_Sc(497) = &HFF&: z_Sc(498) = &HB8870F: z_Sc(499) = &H528B0000: z_Sc(500) = &HFD2850C: z_Sc(501) = &HAD84&: z_Sc(502) = &H3FF3100: z_Sc(503) = &HC704244C: z_Sc(504) = &H5E1C45: z_Sc(505) = &HE8528007: z_Sc(506) = &HFFFFFF14: z_Sc(507) = &HFC0855A: z_Sc(508) = &H9184&: z_Sc(509) = &H244C2B00
        z_Sc(510) = &H14458B04: z_Sc(511) = &H6610E8C1: z_Sc(512) = &H7420033D: z_Sc(513) = &H24448B42: z_Sc(514) = &HE0C15704: z_Sc(515) = &H90C70102: z_Sc(516) = &H3C028B66: z_Sc(517) = &H3C047403: z_Sc(518) = &H52197502: z_Sc(519) = &H6608528B: z_Sc(520) = &H744000A9: z_Sc(521) = &H3C128B02: z_Sc(522) = &HF037502: z_Sc(523) = &HC783D2BF: z_Sc(524) = &H5A178904: z_Sc(525) = &HE210C283: z_Sc(526) = &HD7295AD7: z_Sc(527) = &H2B02EFC1: z_Sc(528) = &H8904247C: z_Sc(529) = &H8B3CEB3A: z_Sc(530) = &H56042444: z_Sc(531) = &HFA89D689: z_Sc(532) = &H8D400A89: z_Sc(533) = &HA5F3873C: z_Sc(534) = &HC728EB5E: z_Sc(535) = &H5E1C45: z_Sc(536) = &H1B98007: z_Sc(537) = &H3000000: z_Sc(538) = &H3104244C: z_Sc(539) = &H8DE852FF
        z_Sc(540) = &H5AFFFFFE: z_Sc(541) = &HE74C085: z_Sc(542) = &HC78F1489: z_Sc(543) = &H107&: z_Sc(544) = &HC2FA8900: z_Sc(545) = &HD2310004: z_Sc(546) = &HF689F9EB: z_Sc(547) = &H1B8&: z_Sc(548) = &HE5896000: z_Sc(549) = &HEB287D8B: z_Sc(550) = &H768D15: z_Sc(551) = &H1B8&: z_Sc(552) = &HE5896000: z_Sc(553) = &HC7287D8B: z_Sc(554) = &H2845&: z_Sc(555) = &H5F8B0000: z_Sc(556) = &HF74AE804: z_Sc(557) = &H558BFFFF: z_Sc(558) = &HFE1EE824: z_Sc(559) = &HC085FFFF: z_Sc(560) = &H3F832774: z_Sc(561) = &HE80D7500: z_Sc(562) = &HFFFFFD44: z_Sc(563) = &H1C7B83: z_Sc(564) = &H15EB4274: z_Sc(565) = &H8B0C4F8B: z_Sc(566) = &H14891047: z_Sc(567) = &HC47FF88: z_Sc(568) = &H83084FFF: z_Sc(569) = &H7400087F
        z_Sc(570) = &H287D832B: z_Sc(571) = &HB82C7400: z_Sc(572) = &H89C&: z_Sc(573) = &H75FFF001: z_Sc(574) = &HFF525028: z_Sc(575) = &HA096&: z_Sc(576) = &H3F8300: z_Sc(577) = &H7F830874: z_Sc(578) = &H8740008: z_Sc(579) = &H7B830DEB: z_Sc(580) = &H775001C: z_Sc(581) = &H1C45C7: z_Sc(582) = &H61000000: z_Sc(583) = &H8C2

        '/// build Tasker VTable -- same order as sent to pvInitFactory
        z_Sc(5) = ((hThunk Xor &H80000000) + &HBC) Xor &H80000000 ' IUnknown:QueryInterface(QI)
        z_Sc(6) = ((hThunk Xor &H80000000) + &H130) Xor &H80000000 ' IUnknown:AddRef
        z_Sc(7) = ((hThunk Xor &H80000000) + &H140) Xor &H80000000 ' IUnknown:Release
        z_Sc(8) = ((hThunk Xor &H80000000) + &H1E4) Xor &H80000000 ' State Get
        z_Sc(9) = ((hThunk Xor &H80000000) + &H20C) Xor &H80000000 ' Tag Get
        z_Sc(10) = ((hThunk Xor &H80000000) + &H218) Xor &H80000000 ' Tag Let
        z_Sc(11) = ((hThunk Xor &H80000000) + &H22C) Xor &H80000000 ' TagItem Get
        z_Sc(12) = ((hThunk Xor &H80000000) + &H260) Xor &H80000000 ' TagItem Let
        z_Sc(13) = ((hThunk Xor &H80000000) + &H29C) Xor &H80000000 ' Pause
        z_Sc(14) = ((hThunk Xor &H80000000) + &H2B0) Xor &H80000000 ' Refresh
        z_Sc(15) = ((hThunk Xor &H80000000) + &H2C0) Xor &H80000000 ' CallNextInChain
        z_Sc(16) = ((hThunk Xor &H80000000) + &H2CC) Xor &H80000000 ' DefWindowProcA
        z_Sc(17) = ((hThunk Xor &H80000000) + &H2D8) Xor &H80000000 ' DefWindowProcW
        z_Sc(18) = ((hThunk Xor &H80000000) + &H2E4) Xor &H80000000 ' IsManaged
        z_Sc(19) = ((hThunk Xor &H80000000) + &H300) Xor &H80000000 ' IsFiltered
        z_Sc(20) = ((hThunk Xor &H80000000) + &H30C) Xor &H80000000 ' AddItem
        z_Sc(21) = ((hThunk Xor &H80000000) + &H354) Xor &H80000000 ' RemoveItem
        z_Sc(22) = ((hThunk Xor &H80000000) + &H3AC) Xor &H80000000 ' FilterClear
        z_Sc(23) = ((hThunk Xor &H80000000) + &H3D4) Xor &H80000000 ' FilterAdd
        z_Sc(24) = ((hThunk Xor &H80000000) + &H468) Xor &H80000000 ' GetItems
        z_Sc(25) = ((hThunk Xor &H80000000) + &H4CC) Xor &H80000000 ' WrapperObject Set
        '/// build API list -- order dictated by listing in source script
        z_Sc(26) = thunkGetProcAddress(hMod("k32"), ByVal "HeapFree")
        z_Sc(27) = thunkGetProcAddress(hMod("k32"), ByVal "HeapDestroy")
        z_Sc(28) = thunkGetProcAddress(hMod("u32"), ByVal "RemovePropA")
        z_Sc(29) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
        z_Sc(30) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
        z_Sc(31) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemAlloc")
        z_Sc(32) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemRealloc")
        z_Sc(33) = thunkGetProcAddress(hMod("u32"), ByVal "DefWindowProcA")
        z_Sc(34) = thunkGetProcAddress(hMod("u32"), ByVal "DefWindowProcW")
        z_Sc(35) = hMod("GetWindowSubclass")
        z_Sc(36) = hMod("SetWindowSubclass")
        z_Sc(37) = hMod("RemoveWindowSubclass")
        z_Sc(38) = hMod("DefSubclassProc")
        z_Sc(39) = thunkGetProcAddress(hMod("u32"), ByVal "EnumThreadWindows")
        z_Sc(40) = thunkGetProcAddress(hMod("u32"), ByVal "EnumChildWindows")
        z_Sc(41) = thunkGetProcAddress(hMod("u32"), ByVal "GetPropA")
        z_Sc(42) = thunkGetProcAddress(hMod("u32"), ByVal "SetPropA")
        z_Sc(43) = lEbMode
        z_Sc(44) = hFactory
        z_Sc(45) = hHeapCommon
        z_Sc(46) = App.ThreadID
        
        ' /// create the Factory object's data section
        arrSafeArray(3) = hFactory              ' address to 1st array element data
        arrSafeArray(4) = 4                     ' number accessed elements in array
        thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), VarPtr(arrSafeArray(0)), 4&
        z_Sc(0) = 1                             ' reference count
        z_Sc(1) = pITypeInfo                    ' ObjPtr(ITypeInfo) - reusable for new taskers
        z_Sc(2) = hTLBstructs                   ' data used by ITypeInfo
        z_Sc(3) = hThunk                        ' this thunk - reusable for new taskers
        thunkCopyMemory oITypeInfo, 0&, 4&      ' zombie the interface, thunk owns it now
        thunkSetProp taskHwnd, FACTORY_PROPNAME, hFactory
        Set hMod = Nothing
    End If
    
    ' /// create the tasker object's data section
    arrSafeArray(3) = hData                     ' address to 1st array element data
    arrSafeArray(4) = DATA_SIZE                 ' number elements in array
    thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), VarPtr(arrSafeArray(0)), 4&
    z_Sc(0) = ((hThunk Xor &H80000000) + 20) Xor &H80000000
    z_Sc(1) = 1                                 ' ref counter
    z_Sc(2) = 0                                 ' Dispatch::IUnknown
    z_Sc(3) = 0                                 ' Tag
    z_Sc(4) = 0                                 ' Wrapper object
    z_Sc(5) = ObjPtr(Host)                      ' Callback object
    z_Sc(6) = HostOrdinal                       ' Callback function address
    z_Sc(7) = 0                                 ' client count
    z_Sc(8) = 0                                 ' filter
    z_Sc(9) = 0                                 ' state
    z_Sc(10) = hTag                             ' used for tracking subclassed hWnds
    thunkCopyMemory ByVal thunkVarPtrArray(z_Sc), 0&, 4& ' remove overlay
    thunkCopyMemory oIUnk, hData, 4             ' create Tasker's IUnknown
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, pITypeInfo, VarPtr(oIDispatch)
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                     ' if failure, tasker unwinds on exit
        ' copy the IDispatch pointer to the object data
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 8) Xor &H80000000, oIDispatch, 4
        Set CreateTasker_Subclass = oIUnk
        If initHwnd <> 0 Then
            If CreateTasker_Subclass.AddItem(initHwnd) = 0 Then _
                Set CreateTasker_Subclass = Nothing ' failure
        End If
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
InitDataObjects:
    ' perform sanity checks on passed parameters
    If Host Is Nothing Then Exit Function
    HostOrdinal = pvGetAddressOf(Host, HostOrdinal)
    If HostOrdinal = 0 Then Exit Function
    
    Call pvInitFactory(taskHwnd, -1)
    If taskHwnd = 0 Then Exit Function
    hFactory = thunkGetProp(taskHwnd, FACTORY_PROPNAME)
    If hFactory = 0 Then            ' first time thunk is being created
        
        Dim bData() As Byte
        ' create array of 4 Longs followed by null terminated ANSI property name
        bData() = StrConv(String$(16, 0) & FACTORY_PROPNAME & vbNullChar, vbFromUnicode)
        hFactory = thunkCoTaskMemAlloc(UBound(bData) + 1)
        Debug.Assert (hFactory <> 0)
        If hFactory = 0 Then Exit Function
        thunkCopyMemory ByVal hFactory, bData(0), UBound(bData) + 1
        Erase bData()
        
        ' call local routine to create a temporary typelib for adhoc IDispatch interface
        Set oITypeInfo = pvInitFactory(taskHwnd, hTLBstructs, _
                    2, "State", 2, "Tag", 4, "Tag", "Value", _
                    2, "TagItem", "pItem", 4, "TagItem", "pItem", "Value", _
                    1, "Pause", 1, "Refresh", _
                    1, "CallNextInChain", "hWnd", "uMsg", "wParam", "lParam", _
                    1, "CallDefWindowProcA", "hWnd", "uMsg", "wParam", "lParam", _
                    1, "CallDefWindowProcW", "hWnd", "uMsg", "wParam", "lParam", _
                    1, "IsManaged", "pItem", 1, "IsFiltered", 1, "AddItem", "!pItem", _
                    1, "RemoveItem", "!pItem", 1, "FilterClear", _
                    1, "FilterAdd", "fType", "!pItem", 1, "GetItems", "pBuffer", "pCount", _
                    4, "WrapperObject", "pObjPtr")
        pITypeInfo = ObjPtr(oITypeInfo)
        Debug.Assert (pITypeInfo <> 0)
        If pITypeInfo = 0 Then GoTo AbortCleanup
        ' modules used by the thunk
        Set hMod = New Collection
        hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
        hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
        hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
        
        On Error Resume Next
        If AddIDEsafety = True Then
            Debug.Print 1 / 0
            If Err Then
                Err.Clear: On Error GoTo 0
                hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
                ' note: if next line errors, change vba6.dll to vba5.dll
                lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
                On Error Resume Next
            End If
        End If
        SetWindowSubclass 0, 0, 0, 0  ' bogus call to load library into process, as needed
        If Err Then
            Err.Clear
            SetWindowSubclassXP 0, 0, 0, 0
            If Err Then
                Err.Clear: GoTo AbortCleanup
            End If
        End If
        On Error GoTo 0
        hMod.Add thunkGetModuleHandle("comctl32.dll"), "cc32"
        hData = thunkGetProcAddress(hMod("cc32"), ByVal "SetWindowSubclass")
        If hData = 0 Then          ' XP exported function by ordinal - test for that
            hData = thunkGetProcAddress(hMod("cc32"), ByVal 410&)
            hMod.Add hData, "SetWindowSubclass"
            hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal 412&), "RemoveWindowSubclass"
            hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal 413&), "DefSubclassProc"
            hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal 411&), "GetWindowSubclass"
        Else
            hMod.Add hData, "SetWindowSubclass"
            hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal "RemoveWindowSubclass"), "RemoveWindowSubclass"
            hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal "DefSubclassProc"), "DefSubclassProc"
            hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal "GetWindowSubclass"), "GetWindowSubclass"
        End If
        ' ^^ note to all: GetWindowSubclass is not exported unless project manifested for common controls
        
    Else                    ' existing thunk - get needed values & increment ref count
        thunkCopyMemory hThunk, ByVal ((hFactory Xor &H80000000) + 12) Xor &H80000000, 4
        thunkCopyMemory pITypeInfo, ByVal ((hFactory Xor &H80000000) + 4) Xor &H80000000, 4
        thunkCopyMemory hData, ByVal hFactory, 4: hData = hData + 1
        thunkCopyMemory ByVal hFactory, hData, 4     ' increment counter
    End If
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData <> 0 Then
        bData() = StrConv("SCtag_" & CStr(hData) & vbNullChar, vbFromUnicode)
        hTag = thunkCoTaskMemAlloc(UBound(bData) + 1)
        If hTag <> 0 Then
            thunkCopyMemory ByVal hTag, bData(0), UBound(bData) + 1
            Erase bData()
            Return
        End If
    End If
    
AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hFactory <> 0 Then thunkCoTaskMemFree hFactory
    If hTag <> 0 Then thunkCoTaskMemFree hTag
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If
    If hHeapCommon <> 0 Then pvHeapManage taskHwnd, hHeapCommon, 0
End Function

Private Function pvGetAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long, _
                            Optional MethodCount As Long, Optional MethodOffset As Long) As Long
    
    ' redesigned but based on Paul Caton's zAddrOf method that can find function
    '   pointers within VB forms, classes, etc.
    ' Redesign includes combining 2 routines into 1, including additional known
    '   VB offsets, and use of VirtualQuery over IsBadCodePtr API.
    
    Dim bSub As Byte, bVal As Byte
    Dim nAddr As Long, vOffset As Long
    Dim nAnchor As Long, nMethod As Long, n As Long
    Dim maxAddr As Long, mbi(0 To 6) As Long  ' faux MEMORY_BASIC_INFORMATION structure
    
    Const MAX_METHODS As Long = 700
    ' ^^ limit of expected max methods a host will have. If your host has more than
    '    the limit, you will need to increase that maximum
    
    If nOrdinal < 1 Then Exit Function
    If oCallback Is Nothing Then Exit Function
    
    thunkCopyMemory nAddr, ByVal ObjPtr(oCallback), 4 ' host VTable
    For n = 1 To 4
        Select Case n
        Case 1: vOffset = &H1C                  ' known VB class offset
        Case 2: vOffset = &H6F8                 ' known form offset
        Case 3: vOffset = &H710                 ' known property page offset
        Case 4: vOffset = &H7A4                 ' known usercontrol offset
        End Select
        vOffset = ((vOffset Xor &H80000000) + nAddr) Xor &H80000000 ' start offset
        ' should find a method from offset within a few attempts...
        maxAddr = ((vOffset Xor &H80000000) + 24) Xor &H80000000
        ' walk the VTable looking for first method to be found
        Do While vOffset < maxAddr
            thunkCopyMemory nMethod, ByVal vOffset, 4 ' get function address at VTable entry
            If nMethod <> 0 Then                ' zero = implemented, skip
                thunkCopyMemory bVal, ByVal nMethod, 1 ' get the 1st byte of that method
                If bVal = &H33 Or bVal = &HE9 Then
                    nAnchor = vOffset           ' bingo. got first method, used as anchor
                    bSub = bVal                 ' cache type of code (native vs. pCode)
                    Exit Do                     ' done looking for anchor
                End If
            End If
            vOffset = ((vOffset Xor &H80000000) + 4) Xor &H80000000
        Loop
        If bSub <> 0 Then Exit For
    Next
    If nAnchor = 0 Then Exit Function           ' failure
    
    ' start with next method & walk til end of object...
    MethodOffset = ((nAnchor Xor &H80000000) + 4) Xor &H80000000
    maxAddr = ((MethodOffset Xor &H80000000) + MAX_METHODS * 4) Xor &H80000000
    ' walk the VTable looking for last method to be found
    Do While MethodOffset < maxAddr
        thunkCopyMemory nAddr, ByVal MethodOffset, 4 ' get function pointer for VTable entry
        mbi(5) = 0  ' reset before each call
        thunkVirtualQuery ByVal nAddr, mbi(0), 28 ' query for properties
        If (mbi(5) And &HEE) = 0 Then           ' can't be read - at end of host
            thunkCopyMemory pvGetAddressOf, ByVal MethodOffset - (nOrdinal * 4), 4
            Exit Do
        ElseIf (mbi(5) And &H101) <> 0 Then     ' else is a page guard or has no access?
            thunkCopyMemory pvGetAddressOf, ByVal MethodOffset - (nOrdinal * 4), 4
            Exit Do                             ' at end of host
        End If
        thunkCopyMemory bVal, ByVal nAddr, 1    ' get function's first byte
        If bVal <> bSub Then                    ' anything other than our anchor's byte, done
            thunkCopyMemory pvGetAddressOf, ByVal MethodOffset - (nOrdinal * 4), 4
            Exit Do
        End If                                  ' move to next VTable entry...
        MethodOffset = ((MethodOffset Xor &H80000000) + 4) Xor &H80000000
    Loop
    If maxAddr >= MethodOffset Then
        MethodCount = (MethodOffset - nAnchor) / 4
        If MethodCount < nOrdinal Then pvGetAddressOf = 0
    End If

End Function

Private Function pvInitFactory(taskHwnd As Long, arrMethods As Long, _
                                ParamArray pParams()) As stdole.IUnknown

    ' taskHwnd is zero if hWnd to be created as needed
    ' arrMethods is returned as handle from CoTaskMemAlloc
    ' pParams() array passed in triplets, never empty:
    '   numeric return vartype, string method name, string param name(s) if any
    '   -- notice that a number always precedes list of strings
    '   for simplicity, all param & method vartypes are Longs unless their
    '   names are prefixed with special characters:
    '       ! = Variant vartype (params only)
    '       $ = String (params only)
    '       & = Object vartype (return value only)
    '       * = IUnknown vartype (return value only)
    
    ' This routine creates a JIT ITypeInfo interface so we can use IDispatch with our
    '   VTable-only IUnknown object created in the various thunk-creation routines.
    '   The end result is passing this typelib along with our VTable-Only IUnknown to
    '   CreateStdDispatch API which creates IDispatch and 'hooks' it into our IUnknown.
    '   That makes the IUnknown a VB-compatible object allowing exposed methods/properties.
    ' What is passed here are the methods/properties and their parameter information
    
    ' VTable ordinal for 1st method always starts with #3; after IUnknown methods
    ' - first passed method is always the default method
    ' - DispID always begins with 0 and is incremented as needed
    
    ' the array containing the required structures is serialized like:
    '   header 8 bytes (variable: arrMethods) [INTERFACEDATA structure]
    '   -- only ever one entry
    '       4 bytes: pointer to first method data structure (mOffset)
    '       4 bytes: method structures count
    '   method structures 28 bytes each (variable: mOffset) [METHODDATA structure]
    '   -- one entry per property/sub/function
    '       4 bytes: pointer to method name (sOffset+xxx)
    '       4 bytes: pointer to ParamData structure, if any (pOffset+xxx)
    '       4 bytes: DispatchID (dispID)
    '       4 bytes: method ordinal in VTable, zero-based
    '       4 bytes: method calling convention
    '       4 bytes: number of items in ParamData array
    '       2 bytes: method type (sub, function, property)
    '       2 bytes: method vartype
    '   param structures 8 bytes each (variable: pOffset) [PARAMDATA structure]
    '   -- one entry per parameter used in all methods
    '       4 bytes: pointer to parameter name (sOffset+xxx)
    '       2 bytes: param vartype
    '       2 bytes: padding
    '   unique unicode strings, double null terminated, dword aligned (variable: sOffset)
    
    Dim mOffset As Long, pOffset As Long, strOffset As Long
    Dim mCount As Long, pCount As Long, p As Long, n As Long
    Dim colNames As VBA.Collection, zMethods(0 To 8) As Long
    Dim bIsParam As Boolean, sName As String, sClass As String
    Const GW_HWNDNEXT As Long = 2
    Const WS_CHILD As Long = &H40000000
    Const NAME_taskHwnd = "SuicideTimerMsgs.v1"

    ' if caller wants the tasker hWnd, then find existing or create, as needed
    If taskHwnd = 0 Then
        sClass = "ThunderRT6Main"      ' class name, compiled
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear
            sClass = "ThunderMain"      ' class name, uncompiled
        End If
        On Error GoTo 0                 ' want VB's hidden owner window
        n = thunkFindWindow(StrPtr(sClass), 0)
        If n = 0 Then Exit Function     ' ensure our instance...
        If thunkGetWindowThreadProcessId(n, 0) <> App.ThreadID Then
            sName = sClass
            Do                          ' keep looking at each top level window
                n = thunkGetWindow(n, GW_HWNDNEXT)
                If n = 0 Then Exit Function ' if our thread, test for what we seek
                If thunkGetWindowThreadProcessId(n, 0) = App.ThreadID Then
                    p = thunkGetClassName(n, StrPtr(sClass), Len(sClass) + 1)
                    If p = Len(sClass) Then  ' right nr of characters, is it what we seek?
                        If StrComp(sClass, sName, vbTextCompare) = 0 Then Exit Do
                    End If
                End If
            Loop
        End If  ' find existing child window else create one
        sClass = "Static"
        taskHwnd = thunkFindWindowEx(n, 0&, StrPtr(sClass), StrPtr(NAME_taskHwnd))
        If taskHwnd = 0 Then
            taskHwnd = thunkCreateWindowEx(0, StrPtr(sClass), StrPtr(NAME_taskHwnd), WS_CHILD, 0, 0, 0, 0, n, 0, App.hInstance, ByVal 0&)
            If taskHwnd = 0 Then Exit Function
        End If
        If arrMethods <> 0 Then Exit Function   ' querying for hWnd only
    End If
    
    ' build the ITypeInfo strucuture & return it as a memory address - 2 passes
    '/// pass #1: Tally string sizes, count methods & parameters for array sizing.
    '    On purpose, parameter names are used more than once between different methods,
    '    if it makes sense to do so. By doing this, we can re-use the string, saving memory.
    Set colNames = New Collection
    For p = 0 To UBound(pParams)
        If VarType(pParams(p)) = vbString Then
            sName = pParams(p)
            If Asc(sName) < 65 Then sName = Mid$(sName, 2)
            On Error Resume Next    ' strings prefixed with 3 char hex offet into array
            colNames.Add Right$("00" & Hex$(strOffset), 3) & sName, sName
            If Err Then
                Err.Clear
            Else
                n = (LenB(sName) + 5) And &HFFFFFFFC
                strOffset = strOffset + n
            End If
            On Error GoTo 0
            If bIsParam = True Then pCount = pCount + 1 Else bIsParam = True
        Else
            bIsParam = False: mCount = mCount + 1
        End If
    Next
    '/// calc overall array size, create array & set offsets
    p = strOffset + mCount * 28 + pCount * 8 + 8
    arrMethods = thunkCoTaskMemAlloc(p)
    Debug.Assert arrMethods <> 0
    
    mOffset = ((arrMethods Xor &H80000000) + 8) Xor &H80000000 ' array position where methods start
    pOffset = ((mOffset Xor &H80000000) + mCount * 28) Xor &H80000000 ' array position where parameters start
    strOffset = ((pOffset Xor &H80000000) + pCount * 8) Xor &H80000000 ' array position where strings start
    zMethods(0) = mOffset: zMethods(1) = mCount
    thunkCopyMemory ByVal arrMethods, zMethods(0), 8    ' copy base structure (header)
    
    '/// copy strings to array
    p = strOffset
    For pCount = 1 To colNames.Count
        sName = Mid$(colNames(pCount), 4)
        n = LenB(sName) + 2
        thunkCopyMemory ByVal p, ByVal StrPtr(sName), n
        p = ((p Xor &H80000000) + ((n + 3) And &HFFFFFFFC)) Xor &H80000000
    Next
    
    '/// pass #2: copy methods & param structures to array
    Erase zMethods(): p = 0
    zMethods(3) = 3: zMethods(4) = 4   ' 3=VTable ordinal,4=stdCall calling convention
    For pCount = 0 To UBound(pParams)
        If VarType(pParams(pCount)) = vbString Then
            sName = pParams(pCount)
            If Asc(sName) < 65 Then sName = Mid$(sName, 2)
            n = CLng("&H" & Left$(colNames(sName), 3))  ' relative offset into array
            If bIsParam = True Then
                zMethods(7) = ((strOffset Xor &H80000000) + n) Xor &H80000000 ' param name offset
                Select Case Asc(pParams(pCount))
                Case 33: zMethods(8) = vbVariant            ' ! prefix (variant)
                Case 36: zMethods(8) = vbString             ' $ prefix (string)
                Case Else: zMethods(8) = vbLong             ' no prefix (long)
                End Select
                thunkCopyMemory ByVal pOffset, zMethods(7), 8
                If zMethods(1) = 0 Then zMethods(1) = pOffset ' set param offset
                zMethods(5) = zMethods(5) + 1               ' increment nr params
                pOffset = ((pOffset Xor &H80000000) + 8) Xor &H80000000
            Else
                zMethods(0) = ((strOffset Xor &H80000000) + n) Xor &H80000000 ' method name offset
                If zMethods(0) <> p Then
                    If p <> 0 Then zMethods(2) = zMethods(2) + 1 ' increment DispID
                    p = zMethods(0)
                End If
                Select Case Asc(pParams(pCount))
                Case 38: zMethods(6) = &H90000 Or zMethods(6) ' ampersand prefix, return ObjPtr
                Case 42: zMethods(6) = &HD0000 Or zMethods(6) ' asterisk prefix, return IUnknown
                Case Else: zMethods(6) = &H30000 Or zMethods(6) ' return Long
                End Select
                bIsParam = True
            End If
        Else
            If pCount <> 0 Then
                thunkCopyMemory ByVal mOffset, zMethods(0), 28
                zMethods(1) = 0: zMethods(5) = 0 ' reset param offset & param count
                zMethods(3) = zMethods(3) + 1
                mOffset = ((mOffset Xor &H80000000) + 28) Xor &H80000000
            End If
            zMethods(6) = pParams(pCount)
            bIsParam = False
        End If
    Next
    thunkCopyMemory ByVal mOffset, zMethods(0), 28      ' copy final method
    Set colNames = Nothing
    If thunkCreateDispTypeInfo(arrMethods, &H800, pvInitFactory) = 0 Then Exit Function
    
ExitRoutine:
    If arrMethods <> 0 Then thunkCoTaskMemFree arrMethods: arrMethods = 0

End Function

Private Function pvHeapManage(taskHwnd As Long, arrCommon As Long, allocSize As Long) As Long

    ' support routine for creating/releasing private heap & thunk allocations
    '---------------------------------------------------------------------------
    ' Params
    '   taskHwnd :: window that contains/will contain the common heap array address
    '   arrCommon :: address to common heap data or zero, nothing else
    '   allocSize :: thunk allocation in bytes, zero passed to recycle allocation
    '---------------------------------------------------------------------------
    '   arrCommon    allocSize  Result
    '---------------------------------------------------------------------------
    '       0        non-zero   return allocation & address to common heap data in arrCommon
    '   non-zero        0       deref and destroy private heap if ref count = 0
    '   non-zero     non-zero   return allocation
    '       0           0       never passed
    '---------------------------------------------------------------------------
    
    Const COMMONHEAP_PROPNAME = "CommonHeapAddr.v1"
    Const HEAP_ATTRS = &H40000  ' Vista+ uses low-fragmentation heaps with this flag
    Dim bData() As Byte
    Dim hHandle As Long, lCount As Long
    
    If arrCommon = 0 Then       ' create/return private heap
        arrCommon = thunkGetProp(taskHwnd, COMMONHEAP_PROPNAME)
        If arrCommon = 0 Then                           ' create private heap
            hHandle = thunkHeapCreate(HEAP_ATTRS, 0, 0)
            If hHandle = 0 Then Exit Function           ' prep contents of common heap data
            ' create array of 3 Longs followed by null terminated ANSI property name
            bData() = StrConv(String$(12, 0) & COMMONHEAP_PROPNAME & vbNullChar, vbFromUnicode)
            arrCommon = thunkCoTaskMemAlloc(UBound(bData) + 1)
            If arrCommon = 0 Then                       ' ensure common data array created
                thunkHeapDestroy hHandle
                Exit Function
            End If                                      ' fill in the array & set window prop
            bData(0) = 1                                ' reference count queried by this class & thunks
            thunkCopyMemory bData(4), taskHwnd, 4       ' queried by the thunks
            thunkCopyMemory bData(8), hHandle, 4        ' queried by the thunks
            thunkCopyMemory ByVal arrCommon, bData(0), UBound(bData) + 1
            thunkSetProp taskHwnd, COMMONHEAP_PROPNAME, arrCommon
            Erase bData()
        Else
            thunkCopyMemory lCount, ByVal arrCommon, 4  ' increment ref count on heap
            lCount = lCount + 1
            thunkCopyMemory ByVal arrCommon, lCount, 4
        End If
    End If                                              ' get heap handle if not done above
    If hHandle = 0 Then _
        thunkCopyMemory hHandle, ByVal ((arrCommon Xor &H80000000) + 8) Xor &H80000000, 4
    If allocSize = 0 Then                               ' releasing allocation or heap
        thunkCopyMemory lCount, ByVal arrCommon, 4      ' get ref count
        If lCount = 1 Then                              ' last reference?
            thunkHeapDestroy hHandle                    ' destroy private heap
            thunkCoTaskMemFree arrCommon                ' destroy common heap data array
            thunkRemoveProp taskHwnd, COMMONHEAP_PROPNAME   ' remove window prop
        Else
            lCount = lCount - 1
            thunkCopyMemory ByVal arrCommon, lCount, 4  ' decrement ref count
        End If
    Else                                                ' allocate for new thunk/stub
        pvHeapManage = thunkHeapAlloc(hHandle, HEAP_ATTRS, allocSize)
    End If
    
End Function

Private Function pvGetForm(hWnd As Long) As VB.Form
    ' Routine checks if passed hWnd is a form & if so, returns that form
    Dim frm As VB.Form
    For Each frm In VB.Global.Forms
        If frm.hWnd = hWnd Then
            Set pvGetForm = frm
            Set frm = Nothing
            Exit For
        End If
    Next
End Function

Private Sub Class_Initialize()

'     If project not manifested to support Win8.1 or better,
'       then version returned by system is: Win8.0 or lower

'   https://docs.microsoft.com/en-us/windows/release-information/
'   Version              Build          Alias
'   ------------------------------------------------------------------------------
'    Windows 10 (1909)   10.0.18363     Nov 2019 Update, 19H2
'    Windows 10 (1903)   10.0.18362     May 2019 Update, 19H1
'    Windows 10 (1809)   10.0.17763     October 2018 Update, Redstone 5
'       ^^ runtime GDI-scaling DPI context supported
'    Windows 10 (1803)   10.0.17134     April 2018 Update, Redstone 4
'       ^^ mixed-mode DPI hosting support
'    Windows 10 (1709)   10.0.16299     Fall Creators Update, Redstone 3
'    Windows 10 (1703)   10.0.15063     Creators Update, Redstone 2
'       ^^ from here up: per-monitor v2 w/non-client support
'    Windows 10 (1607)   10.0.14393     Anniversary Update, Redstone 1
'       ^^ per-monitor awareness with EnableNonClientDpiScaling support
'    Windows 10 (1507, 1511)            Threshold 1 & 2
'    Windows 8.1          6.3.9200+
'       ^^ per-monitor awareness v1, without non-client scaling support
'    Windows Vista to Win8.0
'       ^^ system awareness

    Dim aData(0 To 36) As Long ' Faux OSVERSIONINFO Structure
    Const LOGPIXELSX As Long = 88
    
    aData(0) = 148&         ' 37*4
    GetVersionExA aData(0)  ' get O/S
    If aData(1) = 10 Then   ' will not report as Win10 unless is Win10 & manifested
        ' set version-related flags
        Select Case aData(3)
        Case Is >= 19037:   m_Caps = wcpWin10_2004
        Case Is >= 18363:   m_Caps = wcpWin10_1909
        Case Is >= 18362:   m_Caps = wcpWin10_1903
        Case Is >= 17763:   m_Caps = wcpWin10_1809    ' also WinServer 2019
        Case Is >= 17134:   m_Caps = wcpWin10_1803
        Case Is >= 16299:   m_Caps = wcpWin10_1709
        Case Is >= 15063:   m_Caps = wcpWin10_1703
        Case Is >= 14393:   m_Caps = wcpWin10_1607    ' also WinServer 2016
        Case Else:          m_Caps = wcpWin10_Threshold ' less than Win10,v1607
        End Select
    ElseIf aData(1) = 6 Then    ' Vista and above
        If aData(2) = 3 Then    ' Win8.1? If so, also applies to WinServer 2012/2102R2
            If aData(3) >= 9200 Then m_Caps = wcpWin8_1
        End If
        If m_Caps = 0 Then m_Caps = wcpWinVista ' system is at least Vista
    End If
    
    If (m_Caps And wcpMask_Version) >= wcpWin10_1607 Then
        m_ProjSysDPI = GetDpiForSystem
    Else
        aData(0) = GetDC(0)
        m_ProjSysDPI = GetDeviceCaps(aData(0), LOGPIXELSX)
        ReleaseDC 0, aData(0)
    End If
    
End Sub

Private Sub Class_Terminate()
    Set oHook = Nothing
    Set oSubclass = Nothing
End Sub

'/////////////////////// DO NOT PLACE ANY NEW CODE AFTER THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function myHookProc(ByVal TaskerObject As Object, _
                            ByRef EatMessage As Boolean, _
                            ByVal nCode As Long, ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
    
    ' listen for creation of forms, including MDI child forms
    ' we can't know for sure whether this will be a form, verified after subclassed
    ' note: if a form, an entry in the global Forms collection exists, but w/o an hWnd
    If lParam <> 0 Then
        thunkCopyMemory lParam, ByVal lParam, 4 ' pointer to a CREATEWND structure's 1st member
        If lParam <> 0 Then                     ' pointer to a CREATESTRUCT, get 'style' member
            thunkCopyMemory nCode, ByVal ((lParam Xor &H80000000) + 32) Xor &H80000000, 4
            If (nCode And WS_CHILD) <> 0 Then   ' get 'extended style' member
                thunkCopyMemory nCode, ByVal ((lParam Xor &H80000000) + 44) Xor &H80000000, 4
                If (nCode And WS_EX_MDICHILD) = 0 Then Exit Function
            End If
            oSubclass.AddItem wParam ' listen for WM_NCCREATE & WM_CREATE
        End If
    End If
    
End Function    ' ordinal #2

Private Function mySubclassProc(ByVal TaskerObject As Object, _
                            ByRef EatMessage As Boolean, _
                            ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long, _
                            ByVal hwndTag As Long) As Long

    Select Case uMsg
    Case WM_NCCREATE        ' following not needed for permonitorV2
        If pvGetForm(hWnd) Is Nothing Then
            TaskerObject.RemoveItem hWnd
        ElseIf (m_Caps And wcpMask_Version) = wcpWin10_1607 Then
            ' if Win10, v1607 then EnableNonClientDpiScaling must be called here
            If AreDpiAwarenessContextsEqual(GetWindowDpiAwarenessContext(hWnd), dpiPER_MONITOR_AWARE_V1) <> 0 Then
                EnableNonClientDpiScaling hWnd
            End If
        End If
        
    Case WM_CREATE
        Dim oIAssistant As IDpiPmAssistant, oAssistant As clsDpiPmAssist
        
        TaskerObject.RemoveItem hWnd    ' done subclassing for our purposes
        mySubclassProc = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
        If mySubclassProc <> -1 Then    ' user/system prevented creation, nothing to do
            On Error Resume Next
            Set oIAssistant = pvGetForm(hWnd)
            If Err Then Err.Clear ' we have a form, but it doesn't implement IDpiPmAssistant
            On Error GoTo 0
            If Not oIAssistant Is Nothing Then   ' set up new clsDpiPmAssist class for new form
                Set oAssistant = New clsDpiPmAssist ' call clsDpiPmAssist private method @ known offset
                CallWindowProc pvGetAddressOf(oAssistant, 1), ObjPtr(oAssistant), ObjPtr(oIAssistant), m_ProjSysDPI, 0
                Set oIAssistant = Nothing
            End If
        End If
        EatMessage = True               ' message handled herein
    End Select

End Function    ' ordinal #1
'/////////////////////// DO NOT PLACE ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\

