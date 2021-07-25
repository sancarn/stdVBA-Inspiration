VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDpiPmAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------------

' This class must be declared in each form, MDI, or MDI child for DPI changes to be effective.
' This class is always declared in your form's declarations section and assigned via
'   the implemented IDpiPmAssistant_Attach method.
' Your form must call this class' Activate method from Form_Load only

' This class has three primary purposes:
'   1. Subclass forms, as needed, to listen for and react to DPI changes reported by Windows
'   2. Properly scale form when initially loaded in other than the system DPI
'   3. Provide user with functions to help scale their form, controls, fonts.

' Each form in your project may be subclassed if it implements IDpiPmAssistant and:
'   - the form loaded into a context of per-monitor DPI awareness
'   - if user directed, only if form is loaded into non-integral DPI
' FYI: Subclass procedure is the second to last method in this class.

' DPI can change in these scenarios and possibly others in the future:
'   1. Window is dragged onto another monitor with a different DPI than current monitor
'   2. User changes DPI via the Settings applet
'   3. User docks/undocks laptop
'   4. Connecting via Remote Desktop to a device differing in DPI

'-----------------------------------------------------------------------------------------------
' CODER NOTES
'-----------------------------------------------------------------------------------------------
' Twips (typographical value): 1/1440 of an inch, 1/567 of a centimeter
'   Twips are DPI-independent and are relative to the DPI that initialized the value.
'   Values should not change due to DPI changes. Rather a scaler is used to convert
'   DPI-independence to DPI-dependence. This scaler is Screen.TwipsPerPixelX/Y (TPP).
'   TPP should be calculated as: 1440 / DPI. Since VB is at best system aware only,
'   once a project loads, TPP is calculated but is never recalculated when DPI changes.
'       Tip: VB's TPP is actually calculated as: (1440 \ DPI), rounding down to whole numbers
'   So @ 100% DPI, (1440/96) is 15 TPP and @ 150% DPI, (1440/144) is 10 TPP.
'   100 pixels are 1500 twips and @ 100% DPI, 1500/15 = 100. At 150% DPI, 1500/10 = 150.
'   Perfect. But that all breaks down when DPI changes for DPI-aware applications. Because
'   VB's TPP is locked when the project loads, it does not change when DPI changes and the
'   scaler should change. We cannot change TPP, but we can change twip values. This means
'   we need to tweak the DPI-independent twips to force them to scale to the TPP. When this
'   happens, we have just reversed the twips and scaler dependency. Now the scaler is
'   independent to DPI changes, and the twips are dependent on the DPI via the TPP. Within
'   the form, this is not a problem. But when importing/exporting twips, their values can
'   be incorrect relative to the form's TPP. When importing we must assume the twips are
'   independent and if form's DPI is not the system DPI, those twips need to be scaled to
'   the form's DPI/TPP. Likewise, if twips are exported, they should be independent values.
'   But if the form's DPI is not system DPI, then those twips need to be scaled to independent
'   values. To address this, the class has a TwipsPerRealTwip function that will convert
'   dependent twip values to/from independent values.

' Himetric (metric value): 1/2540 of an inch, 1/1000 of a centimeter
'   These are also DPI independent, similar to twips. stdPicture images use himetrics
'   for their dimensions and its Render function. The scaler for himetrics is the DPI.
'   Similar to twips, himetric values should not change when DPI changes.
'   When VB loads an image into a StdPicture object, the current system DPI is used
'   For example, a 128x128 image would have himetric value of 3387 @ 100% DPI.
'       Conversion to pixels should be this formula: Pixels = HiMetricValue / 2540 * DPI
'       @ 96 DPI (100%), this would be: 3387 / 2540 * 96 = 128 after rounding
'       and @ 168 DPI (175%), this would be: 3387 / 2540 * 168 = 224 after rounding
'   However, himetric image dimensions are scaled to DPI to return actual size vs scaled size
'       Converting actual size to himetrics: HiMetrics = ImageSize / SystemDPI * 2540
'       @ 96 DPI (100%), this would be: 128 / 96 * 2540 = 3387 after rounding
'       and @ 168 DPI (175%), this would be: 128 / 168 * 2540 = 1935 after rounding
'   So, himetric in VB is DPI-dependent.
'   The class provides the GetStdPictureSize method to return real image dimensions,
'   regardless of the relationship between actual system DPI and form's DPI awareness.
'   Also, note that VB for whatever reason seems to use TwipsPerPixel (or similar logic)
'       when returning image dimensions via ScaleX/ScaleY. As noted above, a 128x128
'       image @ 175% DPI will have a himetric value of 1935. A proper conversion to
'       pixels would be: 1935/2540*168. But VB's ScaleX uses something like:
'       1935/2540*(1440\Screen.TwipsPerPixeX) and @ 175%, TPP is 8. That formula
'       returns size as 137 instead of 128. Class GetStdPictureSize method fixes that.
'   Note: In a project where a form is loaded into a different DPI awareness context
'       than what the project loaded into, then stdPictures may not render at the
'       correct scale. In these cases, the stdPicture's Render method should be
'       subclassed and rescaled during rendering. A class property can tell you whether
'       this applies and also provides the scale to use: MixedModeDpiRatio
'   Note: Because OCXs use himetrics when it talks to the parent object regarding
'       coordinates/size relative to each other, this has an effect in non-intrinsic
'       DPIs where an OCX may draw to a smaller size than expected. A class property
'       can tell you whether that applies: SyncOcxRequired. And the class has a
'       method to fix the problem: call SyncOcxToParent after the ocx is resized.

' Public Events. To receive these events, this class must be declared WithEvents
' 1. ResyncAlignedOcx. That event will trigger only if all of these scenarios are true:
'       - SyncOcxRequired property returns true
'       - Form is not minimized
'       - Form resized in any way, including due to DPI scaling unless in an active scaling cycle
'       If you have non-VB intrinsic controls (OCXs, usercontrols) with its Align property
'       set to other than vbAlignNone, then when this event triggers, you should call
'       SyncOcxToParent for each non-VB intrinsic control with Align property
'       set to non-zero. This event should only be used for aligned controls
' 2. RestoreFromMaximized. Event triggers only if these scenarios are true:
'       - form was maximized at time DPI changed
'       - form restored to normal at some point afterwards
'       In that scenario, if the DPI increased, it is possible the form can now be large
'       enough to be rendered partially off screen after DPI scaling. When this event
'       is received, you may want to check that & move the form so it isn't off screen.
'       Note. This event is not sent when a user drags & releases a maximized form
'           that results in the form no longer being maximized. In that case the user
'           is positioning the unmaximized form where they want, so we should respect
'           that decision and not override it by repositioning it.
'-----------------------------------------------------------------------------------------------

Public Event ResyncAlignedOcx() ' see comments above regarding this event
Public Event RestoreFromMaximized() ' see comments above regarding this event

Private Declare Function GetDpiForSystem Lib "user32.dll" () As Long
Private Declare Function GetDpiForWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function AreDpiAwarenessContextsEqual Lib "user32.dll" (ByVal contextA As Long, ByVal contextB As Long) As Long
Private Declare Function SetThreadDpiAwarenessContext Lib "user32.dll" (ByVal dpiContext As Long) As Long
Private Declare Function GetThreadDpiAwarenessContext Lib "user32.dll" () As Long
Private Declare Function SetThreadDpiHostingBehavior Lib "user32.dll" (ByVal hContext As Long) As Long
Private Declare Function GetDpiForMonitor Lib "Shcore.dll" (ByVal hMonitor As Long, ByVal dpiType As Long, ByRef dpiX As Long, ByRef dpiY As Long) As Long
Private Declare Function GetProcessDpiAwareness Lib "Shcore.dll" (ByVal hProcess As Long, ByRef Awareness As Long) As Long
Private Declare Function IsProcessDPIAware Lib "user32.dll" () As Long

Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Function GetObjectW Lib "gdi32.dll" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function CreateFontIndirectW Lib "gdi32.dll" (ByRef lpLogFont As Any) As Long

Private Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hWnd As Long, ByRef lpwndpl As Any) As Long
Private Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hWnd As Long, ByRef lpwndpl As Any) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal lpiid As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef pRect As Any) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef pRect As Any) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As Any) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As Any) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As Any) As Long

' used for Win8.1 compatibility of Me.DpiForSystem & Me.DpiForWindow
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef MonInfo As Any) As Long

Private Enum StateEnum
    flgActivated = &H1          ' Activate() method called
    flgMoveSizeLoop = &H2       ' Form is being dragged for sizing/moving (DPI can change while dragged)
    flgDpiMaxChange = &H4       ' DPI changed while form is maximized
    flgMdiMinimized = &H8       ' DPI changed while MDI child form minimized
    flgShown = &H10             ' set when form's first show occurs
    flgScalingCycle = &H20      ' scaling cycle is in effect
    flgHostSizeRqst = &H40      ' WM_GETDPISCALEDSIZE called
    flgNoAdjMaxRestore = &H80   ' prevent scaling of restored dimensions when DPI changes while maximized
    flgIsPmAware = &H100        ' form's context is per-monitor (v1 or v2)
    flgNoSubclassing = &H200    ' ignore subclassed messages (set/released internally)
    flgSyncOcx = &H400          ' used for the SyncOcxRequired property
    flgShowDpiWarnings = &H800  ' ShowIDEdpiWarnings property value
    flgMDIchild = &H40000000    ' form is a MDI child
    flgMDIparent = &H80000000   ' form is a MDI parent
End Enum
Private Enum ControlTypesEnum
    ctlNone = 0&
    ctlStandard             ' VB intrinsic controls with font properties
    ctlSimpleCombo          ' height change may be enabled (IntegralHeight=False)
    ctlLabel                ' specially handled to honor AutoSize property
    ctlObsolete             ' controls that really should not be used with DPI-awareness
    ctlStdNoFont            ' intrinsic controls without font properties
    ctlLine                 ' specially handled properties
    ctlOcx                  ' not a VB-intrinsic control
    ctlTypeMask = &HF&
    ctlNoAutoSize = &H10&   ' control has no width/height property, user must scale this
    ctlAligned = &H20&      ' control with non-null Align property
    ctlSStab = &H40&        ' SSTab control (only ocx completly handled by the class)
    ctlIntegralHt = &H100&  ' VB control with IntegralHeight set to True
    ctlFixedHt = &H200&     ' VB control where Height property is read-only
    ctlHeightMask = &H300&  ' Controls where Height is restricted, includes IntegralHeight
    ctlNagInIDE = &H400&    ' nag user about controls that can fail scaling on initial load
    ctlImage = &H800        ' image control that will not be scaled
    ctlPictureBox = &H1000  ' picturebox control whose image won't auto-scale with AutoSize=True
    ' intrinsic controls with known font properties
    ctlHasFontMin = ctlStandard
    ctlHasFontMax = ctlObsolete
End Enum

Const WM_DPICHANGED = &H2E0&
Const WM_GETDPISCALEDSIZE = &H2E4&
Const WM_DPICHANGED_AFTERPARENT = &H2E3&
Const WM_ENTERMOVESIZE = &H231&
Const WM_EXITMOVESIZE = &H232&
Const WM_DESTROY = &H2&
Const WM_SIZE = &H5&
Const WM_PAINT = &HF&
Const WS_MINIMIZE = &H20000000
Const WS_MAXIMIZE = &H1000000
Const SWP_NOMOVE = &H2&
Const SWP_NOZORDER_REPOSITION = &H204&
Const SWP_NOSENDCHANGING = &H400&
Const SWP_NOACTIVATE = &H10&
Const GWL_STYLE = -16&
Const MONITOR_DEFAULTTONEAREST As Long = &H2
Const HIMETRIC_CONST = 2540&
Const TWIP_CONST = 1440&
Const BASEDPI_CONST = 96&
Const ClassOnlyMsg = 4&         ' used in SubclassHwnd & zzClientSubclassProc

Dim m_Subclassers As VBA.Collection ' used when host calls SubclassHwnd
Dim m_SubclassDPI As Object     ' subclasser for DPI related messages only, as needed
Dim m_BaseDPIs As Long          ' DPI that project loaded into & form loaded into
Dim m_DPIform As Long           ' DPI that form is currently scaled to
Dim m_Host As Long              ' IDpiPmAssistant soft/weak reference
Dim m_Context As DpiAwarenessContextEnum ' context that form was loaded into
Dim m_Caps As WinCapsEnum       ' version flags, O/S capabilities
Dim m_State As StateEnum        ' various flags used for this class instance
Dim m_CycleParams As Variant    ' userParams parameter passed during scaling cycle events
Dim m_SplashCtrl As Control     ' optional splash control during scaling cycle

Public Sub Activate(Host As Object, _
                    Optional ByVal DesignDPI As Long = 0, _
                    Optional ByVal WantRealignOCXEvents As Boolean = True)
    
    ' This routine must be called from within the form's Load event, after any
    '   dynamically loaded controls are created, and BEFORE the form's Show command.
    ' Reason why we want this in Form_Load is so that initial scaling can be done
    '   without form being Minimized/Maximized, especially if the scale needs to be
    '   changed for first showing. If called after Show or in Form_Activate, form is
    '   already minimized or maximized if designed to show that way at startup.
    ' The routine ensures form is scaled to current DPI depending on that DPI and
    '   the thread's DPI awareness.
    
    ' This class is created by clsDpiPmFormLoader and is passed to your form via
    '   the IDpiPmAssistant_Attach event. You must cache the class at that point then
    '   call this Activate method in your Form_Load routine before the form's Show method
    '   is called, if it is called.
    
    ' Creating this class manually and trying to Activate it within your form is not allowed.
    ' No subclassing or initial form scaling will occur if the above scenario is true.
    ' Most class methods will return unexpected results since key variables will not
    '   have valid values, i.e., m_DPIform, m_BaseDPIs, m_Caps, m_Context, m_State
    
    ' DesignDPI. If passed, the scaling cycle as a result of initial loading,
    ' will not occur when the following is true:
    '   - current DPI and system DPI is same as passed DesignDPI
    
    ' Host is the form that received this class during its IDpiPmAssistant_Attach event.
    ' WantRealignOCXEvents parameter should be True when non-integral DPI is in effect
    '   (SyncOcxRequire=True) and when you are using any ocxs/usercontrols on the form.
    '   This Parameter enables ResyncAlignedOcx events to be sent, but requires form subclassing
    '   When form is per-monitor aware, subclassing is always performed & this parameter
    '   is ignored, treated as True.
    
    ' Subclassing the host (form) will occur in these scenarios:
    '   Form's DPI Context      WantRealignOCXEvents*   SyncOcxRequired*    Subclassed
    '   -----------------------------------------------------------------------------------
    '   dpiPER_MONITOR_AWARE_V1         n/a                  n/a            Yes
    '   dpiPER_MONITOR_AWARE_V2         n/a                  n/a            Yes
    '   all other contexts              True                 True           Yes
    '   any other combination of the first 3 columns above:                 No
    '   (*) WantRealignOCXEvents is a parameter of this method
    '   (*) SyncOcxRequired is a property of this class
    
    If (m_State And flgActivated) = flgActivated Then Exit Sub ' already activated
    If Host Is Nothing Then Exit Sub            ' invalid parameter
    If m_Host = 0 Then Exit Sub                 ' Attach event not responded to
    If TypeOf Host Is VB.Form Then              ' not passed a form
        If pvGetHost() Is Host Then             ' not same form that Attach event was sent to
            m_State = m_State Or flgActivated   ' set flags
            If (TypeOf Host Is MDIForm) = True Then
                m_State = m_State Or flgMDIparent ' type: MDI parent
            ElseIf Host.MDIChild = True Then
                m_State = m_State Or flgMDIchild  ' type: MDI child else stand-alone form
            End If
            If Not m_SubclassDPI Is Nothing Then  ' if applies, created in zzRemoteInit
                If (m_State And flgIsPmAware) = flgIsPmAware Then
                    If (m_State And flgMDIchild) = 0 Then ' trap these for all forms except MDI children
                        m_SubclassDPI.FilterAdd 0, Array(WM_SIZE, WM_GETDPISCALEDSIZE, WM_DPICHANGED, WM_ENTERMOVESIZE, WM_EXITMOVESIZE)
                    Else                                ' trap these for MDI children
                        m_SubclassDPI.FilterAdd 0, Array(WM_SIZE, WM_DPICHANGED_AFTERPARENT, WM_ENTERMOVESIZE, WM_EXITMOVESIZE)
                    End If
                ElseIf WantRealignOCXEvents = True Then
                    m_SubclassDPI.FilterAdd 0, WM_SIZE
                Else
                    Set m_SubclassDPI = Nothing
                End If
            End If
            If m_SubclassDPI Is Nothing Then
                m_State = m_State Or flgShown ' check not needed, set flag
            Else
                If (m_State And flgSyncOcx) = 0 Then
                    m_State = m_State Or flgShown ' check not needed, set flag
                Else
                    m_SubclassDPI.FilterAdd 0, WM_PAINT ' test for first showing
                End If
                m_SubclassDPI.AddItem Host.hWnd
            End If
            ' kick off initial load
            ' note: WindowState should never be minimized if this method was called
            ' from the Form_Load event. If form startup state is minimized/maximized,
            ' then that window state changes after Form_Load or upon Form.Show command.
            Dim lDPI As Long: lDPI = Me.DpiForWindow(Host.hWnd)
            If DesignDPI = lDPI Then
                If Me.DpiForProject = DesignDPI Then Exit Sub
            End If
            pvStartScaleCycle dpiAsst_InitialLoad, lDPI, 0&
        End If
    End If
    
End Sub

Public Property Let AdjustMaximizedRestoreSize(ByVal Value As Boolean)
    ' You should set this immediately before or after your call to Activate
    ' Once you set this, you should not call this again.
    ' This property is True by default
    
    ' This property adjusts the "restore" dimensions of maximized forms when the
    '   DPI changes while the form is maximized. If this is not done, then when
    '   the form is restored to normal, the form's content is scaled, but the
    '   form's dimensions are unchanged from the last time it was at a normal state.
    ' Should future versions of Windows do this adjustment, turn this option off.
    
    m_State = (m_State Or flgNoAdjMaxRestore) Xor flgNoAdjMaxRestore * Abs(Value)
End Property
Public Property Get AdjustMaximizedRestoreSize() As Boolean
    AdjustMaximizedRestoreSize = CBool((m_State And flgNoAdjMaxRestore) = 0)
End Property

Public Property Get IsScalingCycleActive() As Boolean
    ' Courtesy method, not used in this class
    ' this should be called from any routine that resizes/scales the form
    ' or its controls, i.e., Form_Resize, Picture1_Resize, etc.
    ' If this property returns true, you should not do that scaling
    '   else if you do the scaling while this property is true, then
    '   you should not allow the control to be scaled when the
    '   implemented IDpiPmAssistant control-scaling event is sent.
    ' In other words, scale the just once, not twice.
    IsScalingCycleActive = CBool(m_State And flgScalingCycle)
End Property

Public Property Get WindowsCaps() As WinCapsEnum
    ' Courtesy method, not used in this class
    ' returns the Windows version and DPI capabilities of the O/S
    ' if property returns zero, then system is less than Vista
    ' sample usage, see WinCapsEnum Enum definition for more
    ' (WindowsCaps And wcpMask_Version) = wcpWin8_1 when system is Win8.1
    ' (WindowsCaps And wcpMask_Version) >= wcpWin10_1607 when system is at least Win10,v1607
    ' (WindowsCaps And wcpMask_Version) >= wcpMask_PMv2Capable when system supports per-monitor v2
    WindowsCaps = m_Caps
End Property

Public Property Let ShowIDEdpiWarnings(ByVal bShow As Boolean)
    ' Property enables/disables Debug.Print statements regarding known DPI issues
    ' related to your form and/or controls properties. If this is set to True,
    ' you must call it before you call the class Activate method. If project is
    ' compiled, this property is ignored and treated as False.
    If Not bShow = Me.ShowIDEdpiWarnings Then m_State = m_State Xor flgShowDpiWarnings
End Property
Public Property Get ShowIDEdpiWarnings() As Boolean
    ShowIDEdpiWarnings = CBool(m_State And flgShowDpiWarnings)
End Property

Public Property Get DpiContextForForm() As DpiAwarenessContextEnum
    ' Courtesy method, not used in this class
    ' Returns DPI awareness that the form was created in (dependent on manifest)
    DpiContextForForm = m_Context
End Property

Public Property Get DpiForForm() As Long
    ' Courtesy method, not used in this class
    ' Returns DPI that form is currently scaled in.
    ' note: DPI scaling doesn't occur when form is minimized
    ' This property value is only valid after class Activate method is called
    If m_Host = 0 Then
        ' can only be zero if this class created from other than clsDpiPmFormLoader
        DpiForForm = TWIP_CONST \ Screen.TwipsPerPixelX
    Else
        DpiForForm = m_DPIform
    End If
End Property

Public Property Get DpiForProject() As Long
    ' Returns the real system DPI that project loaded into
    ' This may be different than: 1440/Screen.TwipsPerPixelX
    ' project DPI is set just once, when clsDpiPmFormLoader is 1st created in Sub Main()
    If m_Host = 0 Then
        ' can only be zero if this class created from other than clsDpiPmFormLoader
        DpiForProject = Me.DpiForSystem
    Else
        DpiForProject = (m_BaseDPIs And &HFFFF0000) \ &H10000
    End If
End Property

Public Property Get DpiForWindow(hWnd As Long) As Long
    ' Returns DPI that hWnd is running in now, Win8.1 compatible
    If hWnd = 0 Then Exit Property
    If (m_Caps And wcpMask_Version) < wcpWin10_1607 Then ' less than Win10,v1607
        If (m_Caps And wcpMask_Version) < wcpWin8_1 Then ' less than Win8.1
            DpiForWindow = Me.DpiForSystem
        Else
            Dim hMonitor As Long, yDPI As Long
            hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
            GetDpiForMonitor hMonitor, 0, DpiForWindow, yDPI
        End If
    Else
        DpiForWindow = GetDpiForWindow(hWnd)
    End If
    
End Property

Public Property Get DpiForSystem() As Long
    ' Returns system DPI, Win8.1 compatible, in context of thread/process
    If (m_Caps And wcpMask_Version) < wcpWin10_1607 Then ' less than Win10,v1607
        Dim hDC As Long
        Const LOGPIXELSX As Long = 88
        hDC = GetDC(0)
        DpiForSystem = GetDeviceCaps(hDC, LOGPIXELSX)
        ReleaseDC 0, hDC
    Else
        DpiForSystem = GetDpiForSystem
    End If
End Property

Public Property Get MixedModeDpiRatio(Optional ByVal ScaleToProjectDPI As Boolean = False) As Single

    ' Property returns the scaling ratio between the project's DPI and the form's
    '   system DPI. Unless the form was loaded into a different DPI awareness
    '   context than the project, this value is always 1!
    ' This property's value is static.
    
    If m_Host = 0 Then
        ' can only be zero if this class created from other than clsDpiPmFormLoader
        MixedModeDpiRatio = 1!
    ElseIf ScaleToProjectDPI = False Then
        MixedModeDpiRatio = Me.DpiForProject / (m_BaseDPIs And &HFFFF&)
    Else
        MixedModeDpiRatio = (m_BaseDPIs And &HFFFF&) / Me.DpiForProject
    End If
    
End Property

Public Sub GetScreenSize(ScreenWidthPixels As Long, ScreenHeightPixels As Long, _
                        Optional ByVal WorkAreaOnly As Boolean = False, _
                        Optional ByVal asTwips As Boolean = False)
                            
    ' Courtesy method, not used by this class
    ' Returns actual screen size, reported by APIs, in pixels or twips
    ' - WorkAreaOnly excludes space used by desktop taskbars
    ' - the monitor tested is the one the form is currently on
    
    Dim hMonitor As Long, frm As VB.Form
    Dim aData(0 To 9) As Long ' faux MONITORINFO 40-byte structure
    
    If m_Host = 0 Then
        ScreenWidthPixels = Screen.Width / Screen.TwipsPerPixelX
        ScreenHeightPixels = Screen.Height / Screen.TwipsPerPixelY
    Else
        Call pvGetHost(frm)
        hMonitor = MonitorFromWindow(frm.hWnd, MONITOR_DEFAULTTONEAREST)
        Set frm = Nothing: aData(0) = 40
        GetMonitorInfo hMonitor, aData(0)
        If WorkAreaOnly = True Then
            ScreenWidthPixels = (aData(7) - aData(5))
            ScreenHeightPixels = (aData(8) - aData(6))
        Else
            ScreenWidthPixels = (aData(3) - aData(1))
            ScreenHeightPixels = (aData(4) - aData(2))
        End If
        If asTwips = True Then
            ScreenWidthPixels = ScreenWidthPixels * Screen.TwipsPerPixelX
            ScreenHeightPixels = ScreenHeightPixels * Screen.TwipsPerPixelY
        End If
    End If

End Sub

Public Function GetMaxEffectiveDPI(ByVal WidthPixels As Long, ByVal HeightPixels As Long, _
                                Optional ByVal WorkAreaOnly As Boolean = True, _
                                Optional Want25PctMultiple As Boolean = False) As Long
                                
    ' Courtesy method, not used by this class
    ' Returns the largest DPI that can be used to fully display the passed
    '   pixel dimensions. Dividing this value by 0.96 results in a DPI percentage.
    ' Passed dimensions must be values at 100% DPI
    ' - WorkAreaOnly excludes space used by desktop taskbars
    ' - Want25PctMultiple=True then return value is multiple of 24 (1/4 of 100% DPI)
    ' - the monitor tested is the one the form is currently on
    
    Dim Cx As Single, Cy As Single
    Dim X As Long, Y As Long
    Const StdDpi25Pct As Long = 24
    
    Me.GetScreenSize X, Y, WorkAreaOnly
    Cx = WidthPixels / X: Cy = HeightPixels / Y
    If Cy < Cx Then Cx = Cy
    If Want25PctMultiple = False Then
        GetMaxEffectiveDPI = Cx * BASEDPI_CONST
    Else
        GetMaxEffectiveDPI = CLng(Cx * BASEDPI_CONST / StdDpi25Pct) * StdDpi25Pct
    End If

End Function

Public Sub GetMaxEffectiveRez(ScreenWidthPixels As Long, ScreenHeightPixels As Long, _
                            Optional ByVal WorkAreaOnly As Boolean = True)
                            
    ' Courtesy method, not used by this class
    ' Returns the effective resolution of the screen, in pixels, relative to current DPI
    ' - WorkAreaOnly excludes space used by desktop taskbars
    ' - the monitor tested is the one the form is currently on
    
    ' Effective resolution is basically how much screen space you have at
    '   the current DPI relative to 100% DPI.

    Me.GetScreenSize ScreenWidthPixels, ScreenHeightPixels, WorkAreaOnly
    ScreenWidthPixels = ScreenWidthPixels / (m_DPIform / BASEDPI_CONST)
    ScreenHeightPixels = ScreenHeightPixels / (m_DPIform / BASEDPI_CONST)

End Sub

Public Property Get NonIntegralDpiRatio(Optional ByVal ScaleToIntegral As Boolean = False) As Single

    ' This is other than 1! when System DPI is non-integral DPI: 175%, 200%, etc
    '    Example when Me.DpiForSystem = 168 (175% DPI), TwipsPerPixel<>(1440/SystemDPI)
    '    then VB's TPP is 8 which is <> to (1440 / 168) = 8.57
    ' This property may be useful in scaling some OCX properties. Value is static.
    
    ' ScaleToIntegral
    '   False: use return value to scale from non-integral to integral
    '   True: use return value to scale from integral to non-integral
    
    ' Non-integral DPI affects scaling of OCXs. See SyncOcxRequired for more.
    ' At 25% increments, non-integral DPI occurs in majority of DPIs above 150%
    
    If (m_State And flgSyncOcx) = 0 Then
        NonIntegralDpiRatio = 1!
    ElseIf ScaleToIntegral = False Then
        NonIntegralDpiRatio = Screen.TwipsPerPixelX / (TWIP_CONST / Me.DpiForProject)
    Else
        NonIntegralDpiRatio = (TWIP_CONST / Me.DpiForProject / Screen.TwipsPerPixelX)
    End If
    
End Property

Public Sub GetWindowSizePosition(ByVal hWnd As Long, ByRef X As Single, ByRef Y As Single, _
                                ByRef Cx As Single, ByRef Cy As Single, ByVal asTwips As Boolean)
    ' courtesy function returns a window's size and position, in client coordinates
    ' Could be useful to retrieve API window info or possibly info of controls
    '   when the control's dimensions and size are not updated due to external
    '   API calls used to reposition or resize that control.
    
    Dim aData(0 To 3) As Long   ' faux RECT 16-byte structure
    If GetWindowRect(hWnd, aData(0)) = 0 Then
        Cx = 0: Cy = 0: X = 0: Y = 0
    Else
        Cx = aData(2) - aData(0): Cy = aData(3) - aData(1)
        ScreenToClient GetParent(hWnd), aData(0)
        X = aData(0): Y = aData(1)
        If asTwips = True Then
            Cx = Cx * Screen.TwipsPerPixelX: X = X * Screen.TwipsPerPixelX
            Cy = Cy * Screen.TwipsPerPixelY: Y = Y * Screen.TwipsPerPixelY
        End If
    End If
    
End Sub

Public Function TranslateAwarenessContext(ByVal Context As Long) As DpiAwarenessContextEnum
    ' Translates return value of Me.ThreadAwareness or
    '   apis GethThreadDpiAwarenessContext, GetWindowDpiAwarenessContext
    ' to a simple user-friendly DPI awareness value
    ' Requires at least Win10,v1607
    Dim hContext As Long
    If (m_Caps And wcpMask_Version) >= wcpWin10_1607 Then
        For hContext = dpiUNAWARE_GDISCALED To dpiUNAWARE
            If AreDpiAwarenessContextsEqual(Context, hContext) <> 0 Then
                TranslateAwarenessContext = hContext: Exit For
            End If
        Next
    End If
    If hContext = 0 Then TranslateAwarenessContext = Context
End Function

Public Property Get ThreadAwareness() As Long
    ' Returns the thread's current context. Can be different than form's awareness
    ' This value is not user-friendly in Win10,v1607+
    ' To convert it to a better value, call TranslateAwarenessContext(ThreadAwareness)
    Select Case (m_Caps And wcpMask_Version)
    Case Is >= wcpWin10_1607: ThreadAwareness = GetThreadDpiAwarenessContext
    Case Is >= wcpWin8_1
        Dim hContext As Long
        GetProcessDpiAwareness 0, hContext ' api is Win8.1 compatible
        Select Case hContext
        Case 1: ThreadAwareness = dpiSYSTEM_AWARE
        Case 2: ThreadAwareness = dpiPER_MONITOR_AWARE_V1
        Case Else: ThreadAwareness = dpiUNAWARE
        End Select
    Case Is >= wcpWinVista
        If IsProcessDPIAware() = 0 Then
            ThreadAwareness = dpiUNAWARE
        Else
            ThreadAwareness = dpiSYSTEM_AWARE
        End If
    Case Else: ThreadAwareness = dpiSYSTEM_AWARE ' no virtualization applies
    End Select
End Property
Public Property Let ThreadAwareness(ByVal Context As Long)
    ' Sets the thread's current context, applies to Win10,v1607 or greater only
    ' You can pass a DpiAwarenessContextEnum value or value returned by APIs or Me.ThreadAwareness
    
    ' NOTE: You should not mix DPI awareness contexts in the same project unless the
    '   operating system is at least Win10,v1703. Otherwise, whenever thread awareness
    '   changes, each displayed form is rescaled to the new context and every new form
    '   will be assigned the new context. Win10,v1703 fixes that by making contexts
    '   truly per-top-level-window, whereas v1607 and lower are per-thread and since
    '   VB is single-threaded; changes apply to all windows/forms.
    
    If (m_Caps And wcpMask_Version) >= wcpWin10_1607 Then
        Select Case Me.TranslateAwarenessContext(Context)
        Case dpiPER_MONITOR_AWARE_V2
            If (m_Caps And wcpMask_Version) < wcpPMv2Capable Then Context = dpiPER_MONITOR_AWARE_V1
        Case dpiUNAWARE_GDISCALED
            If (m_Caps And wcpMask_Version) < wcpGDIscalingCapable Then Context = dpiUNAWARE
        End Select
        SetThreadDpiAwarenessContext Context
    End If
End Property

Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                    Optional ByVal Title As String, Optional HelpFile As Variant, _
                    Optional Context As Variant) As VbMsgBoxResult

    ' Courtesy method, not used in this class
    ' Forces message boxes to system dpi aware so they are scaled to current DPI
    ' Applies to Win10,v1703 or better else called as-is
    
    ' If Help context is not passed, MessageBoxW is used.
    ' Otherwise, VB's MsgBox function is used
    
    ' The only change in your code to call this function is to preface your
    '   MsgBox() statement with this class, i.e., thisClassVariable.MsgBox(...)

    Dim hContext As Long, frm As VB.Form
    If IsMissing(Context) = True Then
        pvGetHost frm
        MsgBox = MessageBoxW(frm.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons)
        Set frm = Nothing
    Else
        If (m_Caps And wcpMask_Version) >= wcpWin10_1703 Then
            hContext = SetThreadDpiAwarenessContext(dpiSYSTEM_AWARE)
        End If
        MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
        If hContext <> 0 Then SetThreadDpiAwarenessContext hContext
    End If

End Function

Public Function InputBox(ByVal Prompt As String, Optional ByVal Title As String, _
                        Optional ByVal Default As String, _
                        Optional xPos As Variant, Optional yPos As Variant, _
                        Optional HelpFile As Variant, Optional Context As Variant) As String

    ' Courtesy method, not used in this class
    ' Forces input boxes to system dpi aware so they are scaled to current DPI
    ' Applies to Win10,v1703 or better and in a specific case else called as-is
    
    ' The only change in your code to call this function is to preface your
    '   InputBox() statement with this class, i.e., thisClassVariable.InputBox(...)

    Dim hContext As Long
    If (m_Caps And wcpMask_Version) >= wcpWin10_1703 Then
        ' note: Microsoft claims to have made dialogs per-monitor DPI aware, with
        ' some exceptions. The VB InputBox is still broken in one scenario: If the
        ' dialog is activated when the current DPI is not the system DPI, then the
        ' "Prompt" string is rendered at system DPI. To prevent that, we set to fuzzy
        ' system aware in that case.
        If m_DPIform <> Me.DpiForProject Then _
            hContext = SetThreadDpiAwarenessContext(dpiSYSTEM_AWARE)
    End If
    InputBox = VBA.InputBox(Prompt, Title, Default, xPos, yPos, HelpFile, Context)
    If hContext <> 0 Then SetThreadDpiAwarenessContext hContext
    
End Function

Public Sub GetStdPictureSize(thePicture As StdPicture, Width As Long, Height As Long)
    ' Courtesy method, not used in this class
    ' Returned values do not change with DPI
    
    ' Depending on the DPI that VB loaded into, using
    '   ScaleX(thePicture.Width, vbHimetric, vbPixels) can return
    '   the wrong value. This method fixes that.
    If thePicture Is Nothing Then
        Width = 0: Height = 0
    Else
        Dim dScale As Single
        dScale = HIMETRIC_CONST * Me.DpiForProject
        Width = thePicture.Width / dScale
        Height = thePicture.Height / dScale
    End If
End Sub

Public Function TwipsPerRealTwip() As Single
    ' Courtesy method, not used in this class
    ' This property is dynamic and its value changes each time form's DPI changes
    ' note: DPI scaling doesn't occur when form is minimized
    
    ' When form is loaded in non-integral DPI (SyncOcxRequired=True) or a form
    '   is loaded into a different DPI awareness than what the project started
    '   in (MixedModeDpiRatio<>1!), then twips can be scaled incorrectly. This
    '   should not cause you many problems. However, if you need to pass a twips
    '   value outside of your application, or retrieve an outside twips value,
    '   then that value will technically be incorrect when scaled by the form.
    
    ' When exporting twips, multiply twips by this property value
    ' When importing twips, divide twips by this property value
    ' This class performs this conversion, if applicable, during inital form load,
    '   but you may need this function to import/export twips during runtime.
    
    ' Examples.
    ' VB TwipsPerPixel (TPP) values at various DPIs
    ' DPI     VB TPP*    System's TPP*
    ' -------------------------------------------------
    ' 100%      15          15
    ' 125%      12          12
    ' 150%      10          10
    ' 175%       8          8.57 (non-integral DPI)
    ' (*) VB TPP = 1440\SystemDPI
    ' System TPP = 1440/SystemDPI
    
    ' 1500 twips @ 100% DPI is 100 pixels. Here is 1500 twips when scaled to DPI by VB
    ' DPI, TPP   1500/VB TPP   1500/System's TPP
    ' -------------------------------------------------
    ' 100%, 15      100         100     pixels
    ' 125%, 12      125         125     pixels
    ' 150%, 10      150         150     pixels
    ' 175%, 8       188*        175     pixels
    ' (*) 188 pixels is wrong, since 100 pixels scaled by 175% should be 175
    ' When 1500/TwipsPerRealTwip is executed, 1400 twips is returned
    '   and 1400 twips / 8 TPP = 175 pixels which is now correct
    
    ' likewise, if converting 1400 twips @ 175% DPI to real twips for saving elsewhere ...
    ' 1400 twips * TwipsPerRealTwip = 1500 twips
    
    ' one more example. In this scenario, VB was loaded at 125% DPI, TPP is 12.
    ' But a new form was loaded as unaware. That means it is scaled to 100% DPI from 125%.
    ' So, 100 pixels should be 1500 twips and TPP should be 15. However, VB TPP is 12 not 15.
    ' That would mean that 1500/12 would be 125 pixels, so twips are scaled down to 1200 twips
    ' when 1500*TwipsPerRealTwip is returned. 1200 twips / 12 TPP = 100 pixels.
    
    TwipsPerRealTwip = (TWIP_CONST / m_DPIform) / Screen.TwipsPerPixelX
   
End Function

Public Function ScaleStdFont(theFont As IFont, ByVal toDPI As Long) As StdFont
    ' Scales a stdFont object to passed DPI and returns the same object as its return value
    ' note: You may have to reset the font if it is
    '   assigned to a control: Set Text1.Font = ScaleStdFont(Text1.Font, 144)
    If Not theFont Is Nothing Then
        theFont.SetRatio toDPI, HIMETRIC_CONST
        Set ScaleStdFont = theFont
    End If
End Function

Public Function ScaleHFontPointSize(ByVal hFont As Long, ByVal toDPI As Long, ByVal PointSizeAt100Pct As Single) As Long
    ' Courtesy method, not used in this class
    ' Function returns a value used for LOGFONT.lfHeight member or a new font handle

    ' If hFont is passed as non-zero, then a new font will be created
    '   Scales a font handle, from 100% DPI, to a specific DPI based on its point size & returns new font.
    '   You are responsible for destruction of both the provided & returned fonts
    
    ' Otherwise, the passed point size is scaled to the passed DPI and that value
    '   is returned, no font created since no hFont was passed

    ' PointSizeAt100Pct is the font's Size property @ 100% DPI
    ' Formula used: -(pointSize * toDPI) / 72
    
    Dim cbLen As Long, aData() As Long  ' faux LOGFONTW variable-byte structure
    Const PointsPerInch As Long = 72    ' points per inch @ 100% DPI

    If hFont = 0 Then
        ScaleHFontPointSize = -(PointSizeAt100Pct * toDPI) / PointsPerInch
    Else
        cbLen = GetObjectW(hFont, 0&, ByVal 0&)
        If cbLen <> 0 Then
            ReDim aData(0 To cbLen \ 4)
            GetObjectW hFont, cbLen, aData(0)
            aData(0) = -(PointSizeAt100Pct * toDPI) / PointsPerInch
            ScaleHFontPointSize = CreateFontIndirectW(aData(0))
        End If
    End If
    
End Function

Public Function ScaleHFontHeight(ByVal hFont As Long, ByVal ScaleRatio As Single) As Long
    ' Courtesy method, not used in this class
    ' scales a font handle, from its current Height by a scale ratio and returns new font
    ' You are responsible for destruction of both the provided & returned fonts
    
    Dim cbLen As Long, aData() As Long  ' faux LOGFONTW variable-byte structure
    
    If hFont = 0 Or ScaleRatio <= 0! Then Exit Function
    cbLen = GetObjectW(hFont, 0&, ByVal 0&)
    If cbLen <> 0 Then
        ReDim aData(0 To cbLen \ 4)
        GetObjectW hFont, cbLen, aData(0)
        aData(0) = aData(0) * ScaleRatio
        ScaleHFontHeight = CreateFontIndirectW(aData(0))
    End If

End Function

Public Function ScaleValueToDPI(inValue As Variant, ByVal fromDPI As Long, ByVal toDPI As Long) As Variant
    ' Courtesy method, not used in this class
    ' simple DPI conversion routine: MulDiv(value, toDPI, fromDPI)

    On Error Resume Next
    ScaleValueToDPI = (inValue * toDPI) / fromDPI
    If Err Then
        Err.Clear
        ScaleValueToDPI = inValue
    End If
    On Error GoTo 0

End Function

Public Sub ScaleApiWindow(ByVal hWnd As Long, ByVal ScaleRatio As Single, _
                        Optional ByVal newX As Long, Optional ByVal newY As Long, _
                        Optional ByVal newCx As Long, Optional ByVal newCy As Long, _
                        Optional ByVal hFont As Long)
    ' Courtesy method, not used in this class
    ' You can use this method to scale any API child windows you have created and
    ' exist on your form. The ideal place to call this routine is within the
    ' IDpiPmAssistant_DpiScalingCycle when that method's Action parameter
    ' is dpiAsst_EndCycleControls.
    
    ' hWnd. The API child window handle
    ' ScaleRatio
    '   If passed as zero, then newX,newY,newCx,newCy are used as is
    '   If passed as non-zero, then newX,newY,newCx,newCy are ignored and
    '       window's current position/size are scaled by ScaleRatio
    ' newX,newY. Position in Pixel values if ScaleRatio is zero
    ' newCx,newCy. Dimensions in Pixel values if ScaleRatio is zero
    ' hFont. If passed, a SendMessage WM_SETFONT will be called
    
    Dim aData(0 To 3) As Long   ' faux RECT 16-byte structure
    Const WM_SETFONT As Long = &H30
    
    If hWnd = 0 Then Exit Sub
    If ScaleRatio < 0! Then
        Exit Sub
    ElseIf ScaleRatio = 0! Then
        aData(0) = newX: aData(1) = newY
        aData(2) = newCx: aData(3) = newCy
    Else
        If GetWindowRect(hWnd, aData(0)) = 0 Then Exit Sub
        On Error Resume Next            ' should a very large ScaleRatio be given
        aData(2) = (aData(2) - aData(0)) * ScaleRatio
        aData(3) = (aData(3) - aData(1)) * ScaleRatio
        ScreenToClient GetParent(hWnd), aData(0)
        aData(0) = aData(0) * ScaleRatio
        aData(1) = aData(1) * ScaleRatio
        If Err Then
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    End If
    If hFont <> 0 Then SendMessageA hWnd, WM_SETFONT, hFont, 0
    SetWindowPos hWnd, 0, aData(0), aData(1), aData(2), aData(3), SWP_NOACTIVATE Or SWP_NOZORDER_REPOSITION

End Sub

Public Function SubclassHwnd(ByVal hWnd As Long, ByVal When As MsgTiming, ParamArray Messages() As Variant) As Boolean

    ' Since we are already subclassing (likely), might as well offer the client
    ' easy access to subclassing.
    ' DO NOT subclass any hWnd that was created on another form. If you do, you
    '   must ensure that other form unloads before the host of this class unloads.
    ' If you call this more than once for the same hWnd, then
    '   the passed parameters overwrite the previous parameters
    
    ' hWnd: The window to subclass
    '   If this is the form's hWnd, overriding any messages subclassed for DPI changes
    '   can have negative effects
    ' When:
    '   BeforeWindow = get message before hWnd gets it (can prevent hWnd from getting it)
    '   AfterWindow = get message after hWnd processes it
    '   BeforeAfter = both of the above
    ' Messages: if not passed, all messages are trapped. Else provide only the ones
    '   you want to receive. This filter can significantly increase efficiency.
    
    ' Subclassed messages are sent to form's IDpiPmAssistant_Subclasser event.
    ' In that event, to prevent the hWnd from getting the message, 2 actions needed:
    '   1) Set IDpiPmAssistant_Subclasser return value to True
    '   2) Set IDpiPmAssistant_Subclasser EventValue parameter to return to Windows
    ' Above only applies when Subclasser event's BeforeHwnd parameter is true
    ' See IDpiPmAssistant class Subclasser method for more details

    If m_Host = 0 Then Exit Function    ' sanity checks
    If hWnd = 0 Then Exit Function
    If When < BeforeAfter Or When > AfterWindow Then Exit Function
    
    Dim oObj As Object, n As Long, v As Variant
    Dim oLoader As clsDpiPmFormLoader
    
    If m_Subclassers Is Nothing Then Set m_Subclassers = New Collection
    On Error Resume Next                ' test for already subclassed
    Set oObj = m_Subclassers(CStr(hWnd))
    If Err Then
        Err.Clear                       ' not yet subclassed
        Set oLoader = New clsDpiPmFormLoader
        Set oObj = oLoader.CreateTasker_Subclass(Me, 3, hWnd)
        Set oLoader = Nothing           ' abort if subclasser failed creation
        If oObj Is Nothing Then Exit Function
        m_Subclassers.Add oObj, CStr(hWnd)
    Else
        oObj.FilterClear                ' clear existing filter, if any
    End If
    On Error GoTo 0
    
    If IsMissing(Messages) = False Then ' ensure WM_Destroy is trapped
        v = Messages()                  ' so we can remove subclasser from
        oObj.FilterAdd 0, v: v = Empty  ' our m_Subclassers collection
        When = When Or ClassOnlyMsg     ' prevent sending to client if they don't want it
        For n = LBound(Messages) To UBound(Messages)
            Select Case VarType(Messages(n))
            Case vbLong, vbInteger
                If Messages(n) = WM_DESTROY Then
                    When = When Xor ClassOnlyMsg ' client wants it, remove flag
                    Exit For
                End If
            End Select
        Next                            ' add WM_Destroy if necessary
        If (When And ClassOnlyMsg) = ClassOnlyMsg Then oObj.FilterAdd 0, WM_DESTROY
    End If
    oObj.TagItem(hWnd) = When           ' set hWnd tag value for subclassing event
    
End Function

Public Function UnsubclassHwnd(ByVal hWnd As Long) As Boolean

    ' call to unsubclass an hWnd previously passed to SubclassHwnd.
    ' method returns True if the hWnd was subclassed by this class

    On Error Resume Next
    m_Subclassers.Remove CStr(hWnd)
    If Err Then Err.Clear Else UnsubclassHwnd = True
    On Error GoTo 0

End Function

Public Property Get SyncOcxRequired() As Boolean
    ' Returns whether OCX controls should call SyncOcxToParent whenever and after
    ' the OCX is resized. See comments in SyncOcxToParent for more details.
    ' This method always returns true when the following comparison is True:
    '   (1440 / ProjectDPI) <> Screen.TwipsPerPixelX
    ' At 25% increments, non-integral DPI occurs in majority of DPIs above 150%
    
    SyncOcxRequired = CBool(m_State And flgSyncOcx)
    
End Property

Public Sub SyncOcxToParent(theControl As Control, _
                        Optional ByVal OnlyIfRequired As Boolean = True, _
                        Optional InternalAutoSizing As Boolean = False)

    ' Borrowed idea from Krool's Common Controls Replacement project
    
    ' OnlyIfRequired is True, then sub aborts if Me.SyncOcxRequired = False
    ' InternalAutoSizing is True, only when you know that the control is auto-sized
    '   after it has been scaled and auto-sizing occurs externally and incorrectly.
    
    ' InternalAutoSizing = False (typical for OCXs in non-integral system DPI)
    '   Control's Width,Height is set. Then control's window is actually sized
    '   smaller than Width,Height w/scale ratio of: TwipsPerPixelX/(1440/SystemDPI)
    '   Example @ 175% DPI: theControl.Width = 500 pixels. Actual width: 467 pixels
    ' InternalAutoSizing = True
    '   Control's Width,Height is set. Then control's window is actually sized
    '   smaller, but not due to its dimensions being set, due to internal calculations
    '   by the control. Correcting the issue with InternalAutoSizing=False has no
    '   effect because the control keeps auto-sizing to the same wrong size.
    '   Currently, only known OCX that applies is the Common Controls 3 CoolBar

    ' This can be called in any DPI, but for for non-intrinsic controls, should always
    ' be called if SyncOcxRequired=True. The call will force the control to resize to
    ' the dimensions reported by VB, synchronizing the API dimensions with the control.
    
    ' When controls are scaled during a scaling cycle, this method is called from
    ' within the scaling cycle, by this class, unless scaling is aborted by the user.

    ' The routine finds the control's IOleInPlaceSite to retrieve control's
    ' reported dimensions. Then routine finds the control's IOleInPlaceObject to
    ' reapply those dimensions
    
    If OnlyIfRequired = True Then
        If (m_State And flgSyncOcx) = 0 Then Exit Sub
    End If

    Dim oIUnk As stdole.IUnknown, oIOCS As stdole.IUnknown, oIOIPO As stdole.IUnknown
    Dim oIOFrame As stdole.IUnknown, oIOIPWindow As stdole.IUnknown
    Dim oObj As Object, n As Long
    Dim vParamPtr(0 To 4) As Long, vParamType(0 To 4) As Integer
    Dim vRtn As Variant, vParams(0 To 4) As Variant
    Dim aData(0 To 12) As Long  ' parameter values and/or RECT structures
    Const IID_IOleObject = "{00000112-0000-0000-c000-000000000046}"
    Const IID_IOleInPlaceObject = "{00000113-0000-0000-c000-000000000046}"
    Const IID_IOleInPlaceSite = "{00000119-0000-0000-c000-000000000046}"
    Const IOO_GetClientSite As Long = 16&
    Const IOIPS_GetWndContext As Long = 32&
    Const IOIP_SetRects As Long = 28&

    On Error Resume Next
    Set oObj = theControl.object    ' get the object (not its extender)
    If Err Then                     ' VB intrinsic controls don't apply here
        Err.Clear
    ElseIf InternalAutoSizing = True Then
        GetWindowRect theControl.hWnd, aData(0) ' err would be no hWnd proprty
        If Err Then Err.Clear: Set oObj = Nothing
    End If
    On Error GoTo 0
    If oObj Is Nothing Then Exit Sub
    
    If InternalAutoSizing = False Then
        For n = 0 To 4  ' set static DipsCallFunc parameter values
            vParamPtr(n) = VarPtr(vParams(n))
            vParamType(n) = vbLong
        Next
        
        ' set up call to DispCallFunc to get IOleInPlaceObject
        IIDFromString StrPtr(IID_IOleInPlaceObject), VarPtr(aData(0))
        vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIOIPO)
        ' does object implement IID_IOleInPlaceObject?
        DispCallFunc ObjPtr(oObj), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
        If oIOIPO Is Nothing Then Exit Sub
        
        ' set up call to DispCallFunc to get IOleObject
        ' from that, we'll look for IOleClientSite and then IOleInPlaceObjecSite
        IIDFromString StrPtr(IID_IOleObject), VarPtr(aData(0))
        vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIUnk)
        DispCallFunc ObjPtr(oObj), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
        If oIUnk Is Nothing Then Exit Sub
        
        vParams(0) = VarPtr(oIOCS) ' get IOleClientSite from the IOleObject
        DispCallFunc ObjPtr(oIUnk), IOO_GetClientSite, 4&, vbLong, 1, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
        Set oIUnk = Nothing
        If oIOCS Is Nothing Then Exit Sub
        
        ' now get IOleInPlaceSite and retrieve its 2 rectangles for control dimensions
        IIDFromString StrPtr(IID_IOleInPlaceSite), VarPtr(aData(0))
        vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIUnk)
        DispCallFunc ObjPtr(oIOCS), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
        Set oIOCS = Nothing
        If oIUnk Is Nothing Then Exit Sub
        
        ' set the 5 parameters for the call
        vParams(0) = VarPtr(oIOFrame): vParams(1) = VarPtr(oIOIPWindow)
        vParams(2) = VarPtr(aData(0)): vParams(3) = VarPtr(aData(4)) ' 2 rectangles
        vParams(4) = VarPtr(aData(8))                                ' OLEINPLACEFRAMEINFO structure
        DispCallFunc ObjPtr(oIUnk), IOIPS_GetWndContext, 4&, vbLong, 5, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
        Set oIUnk = Nothing: Set oIOFrame = Nothing: Set oIOIPWindow = Nothing
    
        ' reapply the dimensions
        vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(aData(4))
        DispCallFunc ObjPtr(oIOIPO), IOIP_SetRects, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
        Set oObj = Nothing: Set oIOIPO = Nothing
    Else
        On Error Resume Next                ' get pixel size of the control
        n = vbTwips
        Set oObj = theControl.Container
        If Err Then
            Err.Clear
        ElseIf TypeOf oObj Is VB.Form Then
            n = oObj.ScaleMode
        ElseIf TypeOf oObj Is VB.PictureBox Then
            n = oObj.ScaleMode
        End If
        On Error GoTo 0
        If n = vbTwips Then
            aData(4) = theControl.Width / Screen.TwipsPerPixelX
            aData(5) = theControl.Height / Screen.TwipsPerPixelY
        ElseIf n = vbPixels Then
            aData(4) = theControl.Width: aData(5) = theControl.Height
        Else
            aData(4) = oObj.ScaleX(theControl.Width, n, vbPixels)
            aData(5) = oObj.ScaleY(theControl.Height, n, vbPixels)
        End If
        Set oObj = Nothing
        ' use APIs to set the control's size
        If Not (aData(4) = aData(2) - aData(0) And aData(5) = aData(3) - aData(1)) Then
            SetWindowPos theControl.hWnd, 0, 0, 0, aData(4), aData(5), SWP_NOZORDER_REPOSITION Or SWP_NOSENDCHANGING Or SWP_NOMOVE
            theControl.Refresh
         End If
    End If
    
End Sub

Public Sub SyncParentToOcx(theControl As Control, _
                Optional ByVal SyncToParentOnlyIfRequired As Boolean = True)

    ' Similar to SyncOcxToParent. The difference is that this will synchronize
    '   the host dimensions to the control's actual dimensions vs the other
    '   way around. The passed control must have an hWnd.
    
    ' SyncToParentOnlyIfRequired when True will first sync the control to the
    '   the parent to ensure control is sized per the host. Then the host will
    '   be sized per the control after it auto-sizes.
    
    ' Controls that apply would be those that auto-size themselves without their
    ' actual new sizes being known by the host.
    
    Dim oIUnk As stdole.IUnknown, oIOCS As stdole.IUnknown, oIOIPO As stdole.IUnknown
    Dim oIOFrame As stdole.IUnknown, oIOIPWindow As stdole.IUnknown
    Dim oObj As Object, n As Long
    Dim vParamPtr(0 To 4) As Long, vParamType(0 To 4) As Integer
    Dim vRtn As Variant, vParams(0 To 4) As Variant
    Dim aData(0 To 12) As Long  ' parameter values and/or RECT structures
    Const IID_IOleObject = "{00000112-0000-0000-c000-000000000046}"
    Const IID_IOleInPlaceObject = "{00000113-0000-0000-c000-000000000046}"
    Const IID_IOleInPlaceSite = "{00000119-0000-0000-c000-000000000046}"
    Const IOO_GetClientSite As Long = 16&
    Const IOIPS_GetWndContext As Long = 32&
    Const IOIPS_OnPosRectChange As Long = 56&

    On Error Resume Next
    Set oObj = theControl.object    ' get the object (not its extender)
    If Err Then                     ' VB intrinsic controls don't apply here
        Err.Clear
    Else
        n = theControl.hWnd
        If Err Then Err.Clear: Set oObj = Nothing
    End If
    On Error GoTo 0
    If oObj Is Nothing Then Exit Sub
    
    If SyncToParentOnlyIfRequired = True Then SyncOcxToParent theControl
    
    For n = 0 To 4  ' set static DipsCallFunc parameter values
        vParamPtr(n) = VarPtr(vParams(n))
        vParamType(n) = vbLong
    Next
    
    ' set up call to DispCallFunc to get IOleInPlaceObject
    IIDFromString StrPtr(IID_IOleInPlaceObject), VarPtr(aData(0))
    vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIOIPO)
    ' does object implement IID_IOleInPlaceObject?
    DispCallFunc ObjPtr(oObj), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    If oIOIPO Is Nothing Then Exit Sub
    
    ' set up call to DispCallFunc to get IOleObject
    ' from that, we'll look for IOleClientSite and then IOleInPlaceObjecSite
    IIDFromString StrPtr(IID_IOleObject), VarPtr(aData(0))
    vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIUnk)
    DispCallFunc ObjPtr(oObj), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    If oIUnk Is Nothing Then Exit Sub
    
    vParams(0) = VarPtr(oIOCS) ' get IOleClientSite from the IOleObject
    DispCallFunc ObjPtr(oIUnk), IOO_GetClientSite, 4&, vbLong, 1, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    Set oIUnk = Nothing
    If oIOCS Is Nothing Then Exit Sub
    
    ' now get IOleInPlaceSite and retrieve its 2 rectangles for control dimensions
    IIDFromString StrPtr(IID_IOleInPlaceSite), VarPtr(aData(0))
    vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIUnk)
    DispCallFunc ObjPtr(oIOCS), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    Set oIOCS = Nothing
    If oIUnk Is Nothing Then Exit Sub
    
    ' set the 5 parameters for the call
    vParams(0) = VarPtr(oIOFrame): vParams(1) = VarPtr(oIOIPWindow)
    vParams(2) = VarPtr(aData(0)): vParams(3) = VarPtr(aData(4)) ' 2 rectangles
    vParams(4) = VarPtr(aData(8))                                ' OLEINPLACEFRAMEINFO structure
    DispCallFunc ObjPtr(oIUnk), IOIPS_GetWndContext, 4&, vbLong, 5, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    Set oIOFrame = Nothing: Set oIOIPWindow = Nothing
    
    GetWindowRect theControl.hWnd, aData(8)
    If ((aData(2) - aData(0)) <> (aData(10) - aData(8))) Or _
        ((aData(3) - aData(1)) <> (aData(11) - aData(9))) Then
        aData(2) = aData(0) + (aData(10) - aData(8))
        aData(3) = aData(1) + (aData(11) - aData(9))
        vParams(0) = VarPtr(aData(0))
        DispCallFunc ObjPtr(oIUnk), IOIPS_OnPosRectChange, 4&, vbLong, 1, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    End If
    Set oIUnk = Nothing
    Set oObj = Nothing: Set oIOIPO = Nothing
    
End Sub

Public Function LoadFormToContext(ByVal FormName As String, ByVal Context As DpiAwarenessContextEnum, _
                                Optional ByVal EnableMixedModeHosting As Boolean = False) As VB.Form
    
    ' This function can be used to load a new form into a different DPI awareness context.
    ' The context must be compatible with the operating system capabilities
    ' EnableMixedModeHosting should only be passed as true if verson is Win10,1803+ and
    '   you plan on hosting a top level window of another DPI context via SetParent API.
    
    ' NOTE: This is not recommended.
    '   You should not mix DPI awareness contexts in the same project unless the
    '   operating system is at least Win10,v1703. See notes in: ThreadAwareness
    
    Dim hContext As Long, hHosting As Long
    Const DPI_HOSTING_BEHAVIOR_MIXED = 1&
    Const DPI_HOSTING_BEHAVIOR_INVALID = -1&

    On Error GoTo errHandler
    hHosting = DPI_HOSTING_BEHAVIOR_INVALID
    If (m_Caps And wcpMask_Version) >= wcpWin10_1607 Then
        If Me.TranslateAwarenessContext(Context) <> 0 Then
            If EnableMixedModeHosting = True Then
                If (m_Caps And wcpMask_Version) >= wcpMixedHostingCapable Then EnableMixedModeHosting = False
            End If
            If EnableMixedModeHosting Then hHosting = SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR_MIXED)
            hContext = GetThreadDpiAwarenessContext
            If AreDpiAwarenessContextsEqual(hContext, Context) = 0 Then
                Me.ThreadAwareness = Context
                Set LoadFormToContext = Forms.Add(FormName)
                SetThreadDpiAwarenessContext hContext
            Else
                hContext = 0
            End If
        End If
    End If
    If hContext = 0 Then Set LoadFormToContext = Forms.Add(FormName)
    
errHandler:
    If hHosting <> DPI_HOSTING_BEHAVIOR_INVALID Then SetThreadDpiHostingBehavior hHosting
    If hContext <> 0 Then SetThreadDpiAwarenessContext hContext
    If Err Then Err.Clear
    On Error GoTo 0
    
End Function

Private Sub pvResizeControls(frm As VB.Form, dScale As Single, lReason As DpiScaleReason, AddedControls As VBA.Collection)

    ' Key routine to resize controls based on DPI change
    ' Images/Pictures (all formats) are not scaled
    ' Fonts for only VB intrinsic controls are scaled automatically
    ' Other controls added from the toolbox require user to specify fonts to be scaled
    '   For fonts, the scaling event sent to the host has a fontProperties parameter where
    '   user can include font property names to be scaled, as a comma-delimited list.

    ' special considerations:
    '   PictureBox with AutoSize=True will not auto-scale picture if image is non-metafile
    '   Image controls with Stretch=False, dimensions are not scaled if they contain non-metafile images
    '   Control's BorderWidth is not scaled (Line/Shape). User must scale it if wanted.
    '   Form/Control DrawWidth is not scaled. User must scale property if wanted.
    '   Controls that have no Visible property are not scaled, nor passed to you.
    '   Labels. It is important to note that font scaling and dimension scaling by
    '       the same scale ratio does not guarantee caption text will not be clipped
    '       after scaling. Therefore, you should set AutoSize to True for proper
    '       dimension scaling. If a Label caption is multi-line, its WordWrap
    '       property should also be True.
    '   For Data controls that have a negative Left property value, the passed
    '       newX & newY parameters must be corrected if that control is not a child
    '       of the SSTab control. When a negative Left property exists, this routine
    '       assumes Data control is in a SSTab control. These scaling classes track
    '       children by Container property. The Data control has no Container nor hWnd
    '       property, so it can't easily be checked that it is not a SSTab child.
    '   SSTab control is the only ocx that is handled automatically if desired
    '       The scaling event is passed the font in the fontProperties parameter.
    '       Return fontProperties as null string, if font is not to be scaled.
    '   If scaling event aborts scaling for a control, the control must be scaled
    '       by the user. No auto-scaling will occur, no font will be scaled
    '   If the control has no Left,Top,Width or Height property, then the scaling
    '       event is passed new dimensions as zero. User must scale control manually.
    '       Changing the parameter values has no effect since the routine relies on
    '       the Move method for scaling. Such controls won't have a Move method.
    '       If fontProperties parameter is returned, scaling of font(s) occurs.
    '   VB Menu & Timer controls will always be skipped during scaling cycles
    
    Dim c As Control, cc As Control, lType As ControlTypesEnum
    Dim oAssist As IDpiPmAssistant, colCtrls As VBA.Collection
    Dim bNag As Integer, bOk As Boolean
    Dim newX As Single, newCy As Single, newCx As Single, newY As Single
    Dim ctlX As Single, ctlY As Single, ctlCx As Single, ctlCy As Single
    Dim lValue As Long, sFonts As String
    Dim userParams As Variant, bHasSSTab As Boolean
    Const SSTabOffset = -75000!
    
    ' get list of form's controls
    pvGetControlHierarchy frm, colCtrls, AddedControls
    If colCtrls.Count = 0 Then Exit Sub
    
    ' setup flag to report known scaling issues with this form
    If Me.ShowIDEdpiWarnings = True And lReason = dpiAsst_InitialLoad Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear
            bNag = True
        End If
        On Error GoTo 0
    End If
   
    Set oAssist = frm
    For Each c In colCtrls
    
    '/// Step 1. identify the type of control and any special considerations

        Select Case TypeName(c)
        Case "CommandButton", "OptionButton", "TextBox", "CheckBox", "Frame"
            lType = ctlStandard
        Case "Label": lType = ctlLabel
            If c.AutoSize = True Then lType = lType Or ctlFixedHt
        Case "ComboBox"
            If c.Style = vbComboSimple Then ' style honors integral height property
                lType = ctlSimpleCombo
                If c.IntegralHeight = True Then
                    lType = lType Or ctlIntegralHt Or ctlNagInIDE
                End If
            Else ' integral height property ignored, height is read-only
                lType = ctlStandard Or ctlFixedHt
            End If
        Case "ListBox": lType = ctlStandard
            If c.IntegralHeight = True Then
                lType = lType Or ctlIntegralHt Or ctlNagInIDE
            End If
        Case "PictureBox"
            lType = ctlStandard
            If c.Align <> vbAlignNone Then  ' Align takes precedence over AutoSize
                lType = lType Or ctlAligned
            ElseIf c.AutoSize = True Then
                ' images won't scale with control, except metafiles
                Select Case c.Picture.Type
                Case vbPicTypeNone, vbPicTypeEMetafile, vbPicTypeMetafile
                Case Else: lType = lType Or ctlPictureBox Or ctlNagInIDE
                End Select
            End If
        Case "Line": lType = ctlLine
            ' if container scalemode is pixels, line won't scale in other DPIs @ startup
            If bNag <> 0 Then
                lValue = vbTwips
                On Error Resume Next
                lValue = c.Container.ScaleMode
                If Err Then Err.Clear
                On Error GoTo 0
                If lValue = vbPixels Then lType = lType Or ctlNagInIDE
            End If
        Case "Image"
            lType = ctlStdNoFont
            If c.Stretch = False Then
                ' images won't scale with control, except metafiles
                Select Case c.Picture.Type
                Case vbPicTypeNone, vbPicTypeEMetafile, vbPicTypeMetafile
                Case Else: lType = lType Or ctlImage Or ctlNagInIDE
                End Select
            End If
        Case "Shape": lType = ctlStdNoFont
        Case "Data": lType = ctlStandard
            If c.Align <> vbAlignNone Then lType = lType Or ctlAligned
        Case "HScrollBar", "VScrollBar", "OLE": lType = ctlStdNoFont
        Case "DriveListBox": lType = ctlObsolete Or ctlFixedHt
        Case "FileListBox", "DirListBox"
            lType = ctlObsolete Or ctlIntegralHt Or ctlNagInIDE
        Case "Timer", "Menu": lType = ctlNone
        Case "SSTab": lType = ctlOcx Or ctlSStab: bHasSSTab = True
        Case Else
            On Error Resume Next
            lValue = c.Visible ' test if OCX is invisible at runtime - no scaling required
            If Err Then        ' i.e., ImageList controls
                Err.Clear: lType = ctlNone
            Else
                lValue = (c.Width Or c.Height Or c.Left Or c.Top)
                If Err Then     ' expected property doesn't exist, user must scale this
                    Err.Clear: lType = ctlNoAutoSize Or ctlOcx
                Else
                    lType = ctlOcx: lValue = c.Align
                    If Err Then ' control has no Align property
                        Err.Clear
                    ElseIf lValue <> vbAlignNone Then
                        lType = lType Or ctlAligned
                    End If
                End If
            End If
            On Error GoTo 0
        End Select
        
    '/// Step 2. calculate new position and dimensions
        
        If (lType And ctlTypeMask) <> ctlNone Then
            If (lType And ctlTypeMask) = ctlLine Then
                ' handle Line controls individually
                newX = c.X1 * dScale: newY = c.Y1 * dScale
                newCx = c.X2 * dScale: newCy = c.Y2 * dScale
                
            ElseIf (lType And ctlImage) = ctlImage Then
                ' Image control with non-metafile picture & Stretch=False
                newX = c.Left: newY = c.Top ' only scale position, not dimensions
                newCx = c.Width: newCy = c.Height
                
            ' for all other controls with Left,Top,Width,Height properties...
            ElseIf (lType And ctlNoAutoSize) = 0 Then
                If (lType And ctlAligned) = 0 Then
                    newX = c.Left * dScale: newY = c.Top * dScale
                    newCx = c.Width * dScale: newCy = c.Height * dScale
                Else
                    Select Case c.Align
                    Case vbAlignBottom, vbAlignTop  ' scale Top,Height properties only
                        newX = c.Left: newCx = c.Width
                        newY = c.Top: newCy = c.Height * dScale
                        If c.Align = vbAlignBottom Then newY = c.Top + c.Height - newCy
                    Case vbAlignLeft, vbAlignRight  ' scale Left,Width properties only
                        newY = c.Top: newCy = c.Height
                        newX = c.Left: newCx = c.Width * dScale
                        If c.Align = vbAlignRight Then newX = c.Left + c.Width - newCx
                    End Select
                End If
            Else
                newCx = 0!: newCy = 0!: newX = 0!: newY = 0!
            End If
            
            If newX < 0! And bHasSSTab = True Then
                ' specifically address SSTab controls
                On Error Resume Next
                lValue = CBool(TypeName(c.Container) = "SSTab")
                If Err Then
                    ' error? VB Data control has no .Container property
                    ' if an error occurs here, we are going to assume that control
                    ' is contained within an SSTab control.
                    ' User must override passed Left property, if needed
                    Err.Clear: lValue = True
                End If
                On Error GoTo 0
                If lValue = True Then
                    ' SSTab uses a hardcoded value of -5000*15 for offsetting
                    If (lType And ctlTypeMask) = ctlLine Then
                        newCx = (c.X2 - SSTabOffset) * dScale + SSTabOffset
                        newX = (c.X1 - SSTabOffset) * dScale + SSTabOffset
                    Else
                        newX = (c.Left - SSTabOffset) * dScale + SSTabOffset
                    End If
                End If
            End If
        
            If (lType And ctlIntegralHt) = ctlIntegralHt Then
                lValue = 0: pvPreserveIntegralHt c, lValue
            End If
            
    '/// Step 3. Call the host to get scaling options
    
            Set cc = c
            If dScale = 1! Then
                ctlX = newX: ctlY = newY: ctlCx = newCx: ctlCy = newCy
            End If
            If (lType And ctlTypeMask) = ctlOcx Then
                If (lType And ctlSStab) = 0 Then sFonts = vbNullString Else sFonts = "Font"
                bOk = (oAssist.ScaleControlOCX(cc, lReason, dpiAsst_BeginEvent, dScale, newX, newY, newCx, newCy, sFonts, userParams) = 0)
            Else
                bOk = (oAssist.ScaleControlVB(cc, lReason, dpiAsst_BeginEvent, dScale, newX, newY, newCx, newCy, userParams) = 0)
            End If
            Set cc = Nothing
            If bNag <> 0 Then pvShowDpiWarnings bNag, c, lType, sFonts, frm
        
    '/// Step 4. Scale the control and optionally its font(s)
    
            ' if host didn't decline auto-scaling and control not marked as no auto-scale...
            If (lType And ctlNoAutoSize) = ctlNoAutoSize Then
                If bOk = True Then pvScaleStdFontByName c, sFonts
                bOk = False
            ElseIf bOk = True Then
                If dScale = 1! Then
                    If newCx = ctlCx And newCy = ctlCy Then
                        If newX = ctlX And newY = ctlY Then bOk = False
                    End If
                End If
            End If
            
            If bOk = True Then
                If (lType And ctlTypeMask) = ctlSimpleCombo Then
                    pvScaleStdFontByName c, "Font"
                    c.Move newX, newY, newCx    ' can't include height else error
                    c.Height = newCy            ' must resize height separtely
                    If (lType And ctlIntegralHt) = ctlIntegralHt Then pvPreserveIntegralHt c, lValue
            
                ' handle controls where height property is auto-sized or integral...
                ElseIf (lType And ctlHeightMask) <> 0 Then
                    If (lType And ctlLabel) = 0 Then
                        pvScaleStdFontByName c, "Font"
                        If (lType And ctlFixedHt) = 0 Then
                            c.Move newX, newY, newCx, newCy
                            If (lType And ctlIntegralHt) = ctlIntegralHt Then pvPreserveIntegralHt c, lValue
                        Else
                            c.Move newX, newY, newCx
                        End If
                    Else
                        c.Move newX, newY, newCx
                        pvScaleStdFontByName c, "Font"
                    End If
                ' else all other controls...
                Else
                    Select Case (lType And ctlTypeMask)
                    Case ctlOcx
                        ' scale the non-VB intrinsic control...
                        If (lType And ctlSStab) = ctlSStab Then
                            If (m_State And flgIsPmAware) = flgIsPmAware Then
                                c.TabHeight = c.TabHeight * dScale
                                c.TabMaxWidth = c.TabMaxWidth * dScale
                            End If
                        End If
                        pvScaleStdFontByName c, sFonts ' scale passed fonts, if any
                        On Error Resume Next
                        c.Move newX, newY, newCx, newCy
                        If Err Then     ' likely height is read-only
                            Err.Clear
                            c.Move newX, newY, newCx, c.Height
                            If Err Then ' likely width is read-only
                                Err.Clear
                                c.Move newX, newY, c.Width, newCy
                                If Err Then ' done playing with it, send it back to user to scale
                                    Err.Clear
                                    newCx = 0!: newCy = 0!: newX = 0!: newY = 0!
                                    Set cc = c
                                    bOk = (oAssist.ScaleControlOCX(cc, lReason, dpiAsst_BeginEvent, dScale, newX, newY, newCx, newCy, sFonts, userParams) = 0)
                                End If
                            End If
                        End If
                        On Error GoTo 0
                        Me.SyncOcxToParent c
                    Case ctlLine            ' Line controls
                        c.X1 = newX: c.X2 = newCx: c.Y1 = newY: c.Y2 = newCy
                    Case Else               ' VB controls, scale font & then the control
                        Select Case (lType And ctlTypeMask)
                        Case ctlHasFontMin To ctlHasFontMax
                            Set c.Font = ScaleStdFont(c.Font, m_DPIform)
                        End Select
                        c.Move newX, newY, newCx, newCy
                    End Select
                End If
            End If
            
    '/// Step 5. Inform host that control has been scaled
    
            If bOk = True Then
                If (lType And ctlTypeMask) = ctlOcx Then
                    oAssist.ScaleControlOCX c, lReason, dpiAsst_EndEvent, dScale, newX, newY, newCx, newCy, sFonts, userParams
                Else
                    oAssist.ScaleControlVB c, lReason, dpiAsst_EndEvent, dScale, newX, newY, newCx, newCy, userParams
                End If
            End If
            userParams = Empty
        End If
    Next
    
    If bNag <> 0 Then           ' tidy up scaling issues as needed
        If (m_State And flgMDIparent) = 0 Then pvShowDpiWarnings bNag, Nothing, 0, vbNullString, frm
    End If
    If bNag = 1 Then
        Debug.Print String$(80, "-")
        Me.MsgBox "See the immediate/debug window for potential DPI issues", vbInformation + vbOKOnly, "DPI Warnings"
    End If
    
End Sub

Private Sub pvGetControlHierarchy(frm As VB.Form, colCtrls As VBA.Collection, AddedControls As VBA.Collection)

    ' Routine organizes the form's controls (based on its Controls collection)
    '   so that containers of child controls are listed before its child controls
    '   and that controls moved to another form (via SetParent API) are excluded.
    ' If container is on another form then the chldren are also on that other form.
    
    ' The main purpose for this routine is to ensure a container is first scaled
    '   before any of its children are scaled. Typically, children may be scaled
    '   relative to its container's dimensions, not the other way around.
    
    ' Exceptions: Controls that have no Container property may not be listed
    '   after their container if container is other than the form. Such controls
    '   include Data control and any control added to a common-controls Coolbar.
    '   Heirarchy is dependent on Container properties & form's Controls collection.
    ' Note: a better option would likely be to walk the heirarchy by APIs and build
    '   from that. Will revisit this if this method causes major problems. The API
    '   method will take more time to cross-reference form child hWnds to its controls.
    
    ' AddedControls originally passed via pvResizeControls. Cleared before exiting.

    Dim c As Control, cContainers As VBA.Collection
    Dim oPtr As Long, sKey As String, oContainer As Object
    Dim srcList As Object, lLooper As Long, Index As Long
    
    Set colCtrls = New Collection
    Set cContainers = New Collection
    cContainers.Add 1&, CStr(ObjPtr(frm)) ' collection item: non-zero=this form, 0=not this form
    
    On Error Resume Next
    Set srcList = frm.Controls
    For lLooper = 0 To 1 ' 0=frm.Controls, 1=AddedControls
        For Each c In srcList
            Set oContainer = c.Container
            If Err Then    ' no container property: data control or coolbar child?
                Err.Clear  ' add to control collection
            Else
                ' see if container is already in controls & container collections
                sKey = CStr(ObjPtr(oContainer))
                oPtr = cContainers(sKey)
                If Err Then                     ' container seen yet?
                    Err.Clear                   ' and already in our control collection?
                    oPtr = ObjPtr(colCtrls(sKey))
                    If Err Then
                        Err.Clear               ' container not seen, add if on form
                        Do                      ' verify container is not on another form
                            If oContainer Is frm Then Exit Do
                            Set oContainer = oContainer.Container
                            If Err Then
                                Err.Clear
                                Set c = Nothing: Exit Do
                            End If
                        Loop                    ' loop iterations will be small
                        If c Is Nothing Then
                            oPtr = 0            ' not this form
                        Else
                            oPtr = 1: colCtrls.Add c.Container, sKey ' this form
                        End If
                    End If
                    cContainers.Add oPtr, sKey  ' track containers
                ElseIf oPtr = 0 Then
                    Set c = Nothing             ' container is another form
                Else
                    oPtr = ObjPtr(colCtrls(CStr(ObjPtr(c)))) ' ensure not already added from above
                    If Err Then Err.Clear Else Set c = Nothing
                End If
            End If
            If Not c Is Nothing Then colCtrls.Add c, CStr(ObjPtr(c))
        Next
        If lLooper = 0 Then
            If AddedControls Is Nothing Then Exit For
            For Index = AddedControls.Count To 1 Step -1
                Set c = AddedControls(Index)
                If Err Then     ' valid control?
                    Err.Clear: AddedControls.Remove Index
                Else            '  already in the collection?
                    oPtr = ObjPtr(colCtrls(CStr(ObjPtr(c))))
                    If Err Then Err.Clear Else AddedControls.Remove Index
                End If
            Next
            Set srcList = AddedControls
            Set AddedControls = Nothing
        End If
    Next
    Set oContainer = Nothing: Set cContainers = Nothing
    If Not m_SplashCtrl Is Nothing Then colCtrls.Remove CStr(ObjPtr(m_SplashCtrl))
    Set c = Nothing: Set srcList = Nothing
    On Error GoTo 0

End Sub

Private Sub pvShowDpiWarnings(nagValue As Integer, c As Control, cType As ControlTypesEnum, ByVal fontProps As String, frm As VB.Form)

    ' Routine will check passed Font property for being TrueType and
    '  also the passed cType for known DPI issues of intrinsic controls

    Dim aData(0 To 12) As Long ' faux TEXTMETRIC 52-byte structure
    Dim sFonts() As String, n As Long
    Dim hFont As Long, tFont As IFont
    Dim colIndex As VBA.Collection, hDC As Long
    Const TMPF_TRUETYPE = &H4000000
    
    Select Case (cType And ctlTypeMask)
    Case ctlHasFontMin To ctlHasFontMax, 0
        If c Is Nothing Then Set tFont = frm.Font Else Set tFont = c.Font
        If Not tFont Is Nothing Then
            hDC = frm.hDC
            hFont = SelectObject(hDC, tFont.hFont)
            GetTextMetrics hDC, aData(0)
            SelectObject hDC, hFont
            Set tFont = Nothing
            If (aData(12) And TMPF_TRUETYPE) = 0 Then
                Set colIndex = New Collection
                colIndex.Add -1
            End If
        End If
    Case ctlOcx
        If LenB(fontProps) = 0 Then
            sFonts = Split("Font", ",")
        Else
            sFonts = Split(fontProps, ",")
        End If
        For n = 0 To UBound(sFonts)
            fontProps = Trim$(sFonts(n))
            If LenB(fontProps) Then
                On Error Resume Next
                Set tFont = CallByName(c, fontProps, VbGet)
                If Err Then
                    Err.Clear
                Else
                    hDC = frm.hDC
                    hFont = SelectObject(hDC, tFont.hFont)
                    GetTextMetrics hDC, aData(0)
                    SelectObject hDC, hFont
                    Set tFont = Nothing
                    If (aData(12) And TMPF_TRUETYPE) = 0 Then
                        If colIndex Is Nothing Then Set colIndex = New Collection
                        colIndex.Add n
                    End If
                End If
                On Error GoTo 0
            End If
        Next
    End Select
    If (cType And ctlNagInIDE) <> 0 Or (Not colIndex Is Nothing) Then
        If nagValue = -1 Then
            nagValue = 1
            Debug.Print String$(80, "-")
            Debug.Print "Objects that may not scale well at start-up in some DPIs:"
            Debug.Print String$(80, "-")
        End If
        If Not colIndex Is Nothing Then
            For n = 1 To colIndex.Count
                If c Is Nothing Then
                    Debug.Print vbTab; "Form - ";
                Else
                    Debug.Print vbTab; c.Name; " - ";
                End If
                If colIndex(n) = -1 Then
                    Debug.Print "Font";
                Else
                    Debug.Print sFonts(colIndex(n));
                End If
                Debug.Print " is not TrueType"
            Next
        End If
        If (cType And ctlNagInIDE) <> 0 Then
            Debug.Print vbTab; c.Name; " - warning: ";
            If (cType And ctlTypeMask) = ctlLine Then
                Debug.Print "container's scalemode is pixels"
            ElseIf (cType And ctlPictureBox) = ctlPictureBox Then
                Debug.Print "AutoSize property is True, non-metafile picture won't rescale with control."
            ElseIf (cType And ctlImage) = ctlImage Then
                Debug.Print "Stretch property is False, control won't resize with non-metafile picture"
            Else ' integral height items
                Debug.Print "visible item count may be less in some DPIs due to IntegralHeight enforcement"
            End If
        End If
    End If
    
End Sub

Private Function pvGetHost(Optional frm As Variant) As IDpiPmAssistant
    ' helper routine to get IDpiPmAssistant implementation of this class' host
    Dim oIUnk As Object
    CopyMemory oIUnk, m_Host, 4
    Set pvGetHost = oIUnk
    If IsMissing(frm) = False Then Set frm = oIUnk
    CopyMemory oIUnk, 0&, 4
End Function

Private Sub pvScaleStdFontByName(oObj As Object, propNames As String)
    ' helper function for scaling stdFont objects.
    ' propNames is a comma-delimited list of font property names
    
    If LenB(propNames) = 0 Then Exit Sub

    Dim oFont As IFont, sFonts() As String, n As Long
    sFonts() = Split(propNames, ",")
    On Error Resume Next
    For n = 0 To UBound(sFonts)
        sFonts(n) = Trim$(sFonts(n))
        If LenB(sFonts(n)) <> 0 Then
            Set oFont = CallByName(oObj, sFonts(n), VbGet)
            If Err Then
                Err.Clear
            Else
                oFont.SetRatio m_DPIform, HIMETRIC_CONST
                CallByName oObj, sFonts(n), VbSet, oFont
            End If
        End If
    Next
    On Error GoTo 0
    Erase sFonts(): Set oFont = Nothing
    
End Sub

Private Sub pvPreserveIntegralHt(c As Control, lstCount As Long)

    ' routine attempts to keep integral height relative to number of list items
    ' before and after font scaling due to DPI change.
    ' Do note that when form is initalized via VB, height may have already
    '   been changed to a smaller value than desired. The value we use in
    '   this routine is whatever it is after VB first got hold of it.
    
    ' Applies only to intrinsic controls else don't use APIs for sizing OCXs
    
    ' Routine called twice
    ' first run: before font scaling, pass lstCount as zero
    '   calculates & returns number of visible list items relative to scaleHeight
    ' second run: after font scaling, pass lstCount from first run
    '   adjusts scaled height to passed lstCount

    Dim lValue As Long, uMsg As Long, aRect(0 To 7) As Long ' faux RECT 16-byte structures
    Const LB_GETITEMHEIGHT As Long = &H1A1
    Const CB_GETITEMHEIGHT As Long = &H154
    
    If TypeOf c Is ComboBox Then
        uMsg = CB_GETITEMHEIGHT
    Else
        uMsg = LB_GETITEMHEIGHT
    End If
    lValue = SendMessageA(c.hWnd, uMsg, 0&, 0&)
    If lValue < 1 Then Exit Sub
    GetClientRect c.hWnd, aRect(0)
    aRect(3) = aRect(3) - aRect(1)      ' client height
    If lstCount = 0 Then
        lstCount = aRect(3) \ lValue    ' items per height
    Else ' calculate minimal client height needed
        lValue = (aRect(3) - (aRect(3) \ lValue) * lValue) + lValue * lstCount
        If aRect(3) < lValue Then
            GetWindowRect c.hWnd, aRect(4)
            aRect(7) = aRect(7) + (lValue - aRect(3)) ' calculate increase in height & apply
            SetWindowPos c.hWnd, 0, 0, 0, aRect(6) - aRect(4), aRect(7) - aRect(5), SWP_NOMOVE Or SWP_NOZORDER_REPOSITION Or SWP_NOACTIVATE
        End If
    End If
    
End Sub

Private Sub pvStartScaleCycle(lReason As DpiScaleReason, _
                            Optional wParam As Long, Optional lParam As Long)
                            
    ' function starts and ends a scaling cycle due to DPI-related events
    ' this function is called from multiple places:
    '   Subclassed messages, Activate, pvInitDpiChangeMdiChild

    ' wParam is passed as the new DPI
    ' lParam only passed when lReason is dpiAsst_DpiChanged
    
    Dim dScale As Single, oAssist As IDpiPmAssistant, frm As VB.Form
    Dim newCx As Single, newCy As Single, initCx As Single, initCy As Single
    Dim aData() As Long, bOk As Boolean, vControls As VBA.Collection
    
    ReDim aData(0 To 7) ' faux RECT 16-byte structures
    Set oAssist = pvGetHost(frm)
    
'/// Step 1: set scale factor and reason for scaling cycle
    dScale = (wParam And &HFFFF&) / m_DPIform   ' set scale based on old/new DPI
    If lReason = dpiAsst_DpiChanged Then        ' called from WM_DPICHANGED
        If (m_State And flgMoveSizeLoop) = flgMoveSizeLoop Then _
            lReason = dpiAsst_DpiChangeWhileDragging
    End If
    
    If (m_State And flgScalingCycle) = 0 Then
        oAssist.DpiScalingCycle lReason, dpiAsst_BeginCycleHost, m_DPIform, (wParam And &HFFFF&), m_CycleParams
        m_State = m_State Or flgScalingCycle
    End If

'/// Step 2: set form's new dimensions and offer host chance to change size; scale form
    If lReason = dpiAsst_InitialLoad Then   ' called from class methods, set dimensions/position
        aData(0) = frm.Left / Screen.TwipsPerPixelX
        aData(1) = frm.Top / Screen.TwipsPerPixelY
        aData(2) = (frm.Width / Screen.TwipsPerPixelX) * dScale + aData(0)
        aData(3) = (frm.Height / Screen.TwipsPerPixelY) * dScale + aData(1)
    Else
        CopyMemory aData(0), ByVal lParam, 16   ' get suggested dimensions/position
    End If
    newCx = (aData(2) - aData(0)) * Screen.TwipsPerPixelX
    newCy = (aData(3) - aData(1)) * Screen.TwipsPerPixelY
        
    If (m_State And flgHostSizeRqst) = 0 Then ' offer host abiltiy to change form's size
        If frm.WindowState = vbNormal Then
            If lReason = dpiAsst_InitialLoad And Me.MixedModeDpiRatio <> 1! Then
                ' Form's are always loaded in project's DPI by VB.
                ' In cases where the project's system DPI is different than this
                ' form's system DPI, the form is initially created by VB using a
                ' different scale factor that can affect the overall form size &
                ' client area size being slightly too small or too large. Here we
                ' snapshot the client area size before scaling. And after the form
                ' is scaled, we will adjust the form dimensions to ensure client
                ' area is properly scaled, not too small/large. When project's
                ' system DPI <> form's system DPI, then that can only happen when
                ' one of these are true @ other than 96 DPI:
                '   Project is unaware and form is not unaware
                '   Form is unaware and project is not unaware
                ' Remember: when not unaware, project DPI is always used for initial
                '   form scaling by VB, regardless of the current DPI value. And
                '   when unaware, system DPI is always 96 as reported to VB.
                GetClientRect frm.hWnd, aData(4)
                aData(6) = (aData(6) - aData(4)) * (newCx / frm.Width)
                aData(7) = (aData(7) - aData(5)) * (newCy / frm.Height)
                initCx = newCx: initCy = newCy  ' initial load scale
            End If
            oAssist.ScaleHost lReason, newCx, newCy, m_DPIform, (wParam And &HFFFF&), bOk
            If aData(7) <> 0 Then
                aData(2) = (newCx / Screen.TwipsPerPixelX) + aData(0)
                aData(3) = (newCy / Screen.TwipsPerPixelY) + aData(1)
            End If
        Else
            oAssist.ScaleHost lReason, (newCx), (newCy), m_DPIform, (wParam And &HFFFF&), bOk
        End If
        If bOk = True Then pvCreateSplashCtrl frm, dScale
    Else
        m_State = m_State Xor flgHostSizeRqst
    End If
        
    If frm.WindowState = vbMaximized Then
        If (m_State And flgNoAdjMaxRestore) = 0 Then
            ReDim aData(0 To 14): aData(0) = 60
            GetWindowPlacement frm.hWnd, aData(0) ' current restore size, calculate new window size
            aData(9) = aData(7) + (aData(9) - aData(7)) * dScale
            aData(10) = aData(8) + (aData(10) - aData(8)) * dScale
            SetWindowPlacement frm.hWnd, aData(0) ' restored position/dimensions
            m_State = m_State Or flgDpiMaxChange
        End If
    ElseIf (newCx <> frm.Width Or newCy <> frm.Height) Then
        m_State = m_State Or flgNoSubclassing
        SetWindowPos frm.hWnd, 0, aData(0), aData(1), newCx / Screen.TwipsPerPixelX, newCy / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_NOSENDCHANGING Or SWP_NOZORDER_REPOSITION
        If aData(7) <> 0 Then   ' non-zero only if we snapshot client area size
            ' ensure client area is scaled properly after form loaded
            ' unless host chose non-default initial form size
            If initCx = newCx And initCy = newCy Then
                GetClientRect frm.hWnd, aData(0)
                aData(2) = aData(2) - aData(0): aData(3) = aData(3) - aData(1)
                If aData(2) <> aData(6) Or aData(3) <> aData(7) Then
                    newCx = newCx / Screen.TwipsPerPixelX + (aData(6) - aData(2))
                    newCy = newCy / Screen.TwipsPerPixelY + (aData(7) - aData(3))
                    SetWindowPos frm.hWnd, 0, 0, 0, newCx, newCy, SWP_NOACTIVATE Or SWP_NOSENDCHANGING Or SWP_NOZORDER_REPOSITION Or SWP_NOMOVE
                End If
            End If
        End If
        m_State = m_State Xor flgNoSubclassing
    End If
    
    m_DPIform = wParam And &HFFFF& ' set form's DPI & scale its font
    pvScaleStdFontByName frm, "Font"
    
    If (m_State And flgMDIchild) = flgMDIchild Then ' maximized MDI child?
        If frm.WindowState = vbMaximized Then pvAlignMdiChild frm.hWnd
    End If

'/// Step 3: offer option to scale controls
    If oAssist.DpiScalingCycle(lReason, dpiAsst_BeginCycleControls, (m_DPIform / dScale), m_DPIform, m_CycleParams) = 0 Then
        Set vControls = New VBA.Collection
        oAssist.IncludeSetParentControls vControls, lReason, dScale, m_CycleParams
        pvResizeControls frm, dScale, lReason, vControls
    End If
    oAssist.DpiScalingCycle lReason, dpiAsst_EndCycleControls, (m_DPIform / dScale), m_DPIform, m_CycleParams
    If Not m_SplashCtrl Is Nothing Then
        frm.Controls.Remove m_SplashCtrl
        Set m_SplashCtrl = Nothing
    End If
    
'/// Step 4: trigger end of scaling cycle
    m_State = m_State Xor flgScalingCycle
    oAssist.DpiScalingCycle lReason, dpiAsst_EndCycleHost, (m_DPIform / dScale), m_DPIform, m_CycleParams
    m_CycleParams = Empty
    Set oAssist = Nothing
    If (m_State And flgSyncOcx) = flgSyncOcx Then
        If (m_State And flgShown) = flgShown Then RaiseEvent ResyncAlignedOcx
    End If
    
'/// Step 5: if MDI parent, send child forms the change event
    If m_State < 0 Then
        ' When a MDI parent gets this, its 'child' forms do not get this messsage.
        ' In PMv2, the children will get WM_DPICHANGED_AFTERPARENT/WM_DPICHANGED_BEOREPARENT.
        ' So when not PMv2, we will fake WM_DPICHANGED_AFTERPARENT messages
        ' Exception: we always send this on initial form load
        If (lReason = dpiAsst_InitialLoad) = True Or (m_Context = dpiPER_MONITOR_AWARE_V2) = False Then
            aData(1) = FindWindowEx(frm.hWnd, 0, "MDICLIENT", vbNullString)
            aData(0) = FindWindowEx(aData(1), 0, vbNullString, vbNullString)
            Do Until aData(0) = 0
                ' must use SendMessage, not PostMessage else ERROR_MESSAGE_SYNC_ONLY
                SendMessageA aData(0), WM_DPICHANGED_AFTERPARENT, 0, 0
                aData(0) = FindWindowEx(aData(1), aData(0), vbNullString, vbNullString)
            Loop
        End If
    End If
    Set frm = Nothing

End Sub

Private Sub pvInitDpiChangeMdiChild(hWnd As Long)

    ' Routine handles DPI changes in MDI child forms
    ' note: to emulate top-level forms, this is not called if MDI child is minimized,
    ' when DPI changed, until the MDI child is restored or maximized
    
    ' Child windows do not get WM_DPICHANGED messages, but they can receive other
    '   DPI-related messages, depending on per-monitor DPI awareness context/version.
    ' Here we are faking a WM_DPICHANGED message for the MDI child form

    Dim dScale As Single, lDPI As Long, aRect(0 To 3) As Long ' faux RECT 16-byte structure
    
    GetWindowRect hWnd, aRect(0)
    lDPI = Me.DpiForWindow(hWnd)
    dScale = lDPI / m_DPIform
    aRect(2) = (aRect(2) - aRect(0)) * dScale
    aRect(3) = (aRect(3) - aRect(1)) * dScale
    ScreenToClient GetParent(hWnd), aRect(0)
    aRect(2) = aRect(2) + aRect(0): aRect(3) = aRect(3) + aRect(1)
    
    pvStartScaleCycle dpiAsst_DpiChanged, lDPI, VarPtr(aRect(0))

End Sub

Private Sub pvAlignMdiChild(hWnd As Long)

    ' Routine handles positioning for maximized MDI child forms
    ' VB seems to cache sizes of a MDI child's non-client parts and fails to properly
    '   offset maximized children when the DPI is non-system DPI. This routine
    '   attempts to overcome that.

    Dim hMdiClient As Long
    Dim aRect(0 To 7) As Long ' faux RECT 16-byte structures
    ' 0-3   MDI Client & Form client left/top coordinates, offsets
    ' 4-7   Form overall, adjusted size
    
    If Me.DpiForWindow(hWnd) <> Me.DpiForProject Then
        hMdiClient = GetParent(hWnd)
        ' get MDI Client client rect and translate to screen coordinates
        GetClientRect hMdiClient, aRect(0): ClientToScreen hMdiClient, aRect(0)
        ' get form's client rect and translate to screen coordinates
        GetClientRect hWnd, aRect(2): ClientToScreen hWnd, aRect(2)
        ' calculate the difference in left & top coordinates
        aRect(0) = aRect(2) - aRect(0): aRect(1) = aRect(3) - aRect(1)
        If aRect(0) = 0 And aRect(1) = 0 Then Exit Sub ' no change, nothing to do
        
        ' get form's screen size/coordinates
        GetWindowRect hWnd, aRect(4)
        ' calculate form's adjusted width/height
        aRect(6) = (aRect(6) - aRect(4)) + aRect(0) ' width
        aRect(7) = (aRect(7) - aRect(5)) + aRect(1) ' height
        ' calculate form's new left,top
        aRect(4) = aRect(4) - aRect(0): aRect(5) = aRect(5) - aRect(1)
        ' translate new rect to the MDI client & reposition
        ScreenToClient hMdiClient, aRect(4)
        SetWindowPos hWnd, 0, aRect(4), aRect(5), aRect(6), aRect(7), SWP_NOACTIVATE Or SWP_NOSENDCHANGING Or SWP_NOZORDER_REPOSITION
    End If

End Sub

Private Function pvGetHostDpiSize(wParam As Long, lParam As Long) As Long

    ' helper function to handle WM_GETDPISCALEDSIZE message

    Dim lReason As DpiScaleReason, aData(0 To 1) As Long ' faux SIZE 8-byte structure
    Dim newCx As Single, newCy As Single
    Dim dScale As Single, bOk As Boolean
    Dim frm As VB.Form, oAssistant As IDpiPmAssistant
    
    If (m_State And flgMoveSizeLoop) = flgMoveSizeLoop Then
        lReason = dpiAsst_DpiChangeWhileDragging
    Else
        lReason = dpiAsst_DpiChanged
    End If
    
    Set oAssistant = pvGetHost(frm)
    dScale = wParam / m_DPIform
    CopyMemory aData(0), ByVal lParam, 8        ' system recommended size
    newCx = aData(0) * Screen.TwipsPerPixelX    ' set to twips
    newCy = aData(1) * Screen.TwipsPerPixelY    ' and trigger scaling cycle start
    oAssistant.DpiScalingCycle lReason, dpiAsst_BeginCycleHost, m_DPIform, wParam, m_CycleParams
    
    If frm.WindowState = vbMaximized Then       ' size is not changeable
        oAssistant.ScaleHost lReason, newCx, newCy, m_DPIform, wParam, bOk
    Else                                        ' offer option to change size
        oAssistant.ScaleHost lReason, newCx, newCy, m_DPIform, wParam, bOk
        If Not (newCx = aData(0) * Screen.TwipsPerPixelX And newCy = aData(1) * Screen.TwipsPerPixelY) Then
            aData(0) = newCx / Screen.TwipsPerPixelX    ' altered size
            aData(1) = newCy / Screen.TwipsPerPixelY    ' need to feed it back
            CopyMemory ByVal lParam, aData(0), 8
            pvGetHostDpiSize = 1                        ' result for caller
        End If
    End If                                              ' set flags for new cycle
    m_State = m_State Or flgScalingCycle Or flgHostSizeRqst
    If bOk = True Then pvCreateSplashCtrl frm, dScale   ' create splash control
    Set frm = Nothing: Set oAssistant = Nothing

End Function

Private Sub pvCreateSplashCtrl(frm As VB.Form, ScaleRatio As Single)
    ' created during sclaing cycle per user request
    ' does not apply to MDI parents
    
    If (m_State And flgMDIparent) = flgMDIparent Then Exit Sub
    
    Set m_SplashCtrl = frm.Controls.Add("VB.PictureBox", "lvsplsh" & ObjPtr(Me) & "ctrl")
    With m_SplashCtrl
        If ScaleRatio < 1! Then
            .Move 0, 0, frm.ScaleWidth, frm.ScaleHeight
        Else
            .Move 0, 0, frm.ScaleWidth * ScaleRatio, frm.ScaleHeight * ScaleRatio
        End If
        .BorderStyle = 0: .Appearance = 0
        .Enabled = False: .ZOrder vbBringToFront
        .BackColor = vbWindowBackground
        .Visible = True: .Refresh
    End With
End Sub

Private Sub Class_Terminate()
    Set m_SubclassDPI = Nothing
    Set m_Subclassers = Nothing
End Sub

'/////////////////////// DO NOT PLACE ANY NEW CODE AFTER THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function zzClientSubclassProc(ByVal TaskerObject As Object, _
                            ByRef EatMessage As Boolean, _
                            ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long, _
                            ByVal When As MsgTiming) As Long
                            
    ' This is triggered if client called SubclassHwnd method
    ' Messages here are forwarded to the host for review/action
    ' WM_Destroy always received here and cannot be overridden
    
    Dim lRtn As Long
    Const WHEN_MASK = 3&
    
    If uMsg = WM_DESTROY Then
        Me.UnsubclassHwnd hWnd
        ' if this class added the message, then done
        If (When And ClassOnlyMsg) = ClassOnlyMsg Then Exit Function
    End If
    If (When And WHEN_MASK) <> AfterWindow Then  ' else user wants only "After" messages
        ' offer chance to override, unless WM_Destroy
        EatMessage = pvGetHost().Subclasser(lRtn, BeforeWindow, hWnd, uMsg, wParam, lParam)
        If uMsg = WM_DESTROY Then EatMessage = False
    End If
    If (When And WHEN_MASK) <> BeforeWindow Then ' else user wants only "Before" messages
        If EatMessage = False Then  ' if not overridden, send as "After" message
            lRtn = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
            pvGetHost().Subclasser (lRtn), AfterWindow, hWnd, uMsg, wParam, lParam
        End If
    End If
    zzClientSubclassProc = lRtn
                            
End Function    ' ordinal #3

Private Function zzSubclassProc(ByVal TaskerObject As Object, _
                            ByRef EatMessage As Boolean, _
                            ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long, _
                            ByVal hwndTag As Long) As Long
                            
    ' This subclasser does not prevent messages from being passed to
    '   subclassers further down the chain, nor to the default class procedure.
    '   Messages herein are subclassed for action after they've been processed
    '   by lower subclassers and/or the system.
    ' Exceptions:
    '   WM_DPICHANGED. Handled here & not forwarded
    '   WM_DPICHANGED_AFTERPARENT. For MDI child only, handled here & not forwarded
    '   WM_GETDPISCALEDSIZE. Handled here & not fowarded
    '   VB does not process any of those messages.
    
    If (m_State And flgNoSubclassing) = flgNoSubclassing Then Exit Function
    
    Select Case uMsg
    Case WM_DPICHANGED
        ' only called with per-monitor awareness and only for top-level (non-child) windows
        ' WM_DPICHANGED_BEFOREPARENT already called for each child window, if applicable
        ' WM_DPICHANGED_AFTERPARENT called afterwards for each child window, if applicable
        If (m_State And flgIsPmAware) = flgIsPmAware Then
            zzSubclassProc = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
            pvStartScaleCycle dpiAsst_DpiChanged, wParam, lParam
        End If
        EatMessage = True
        
    Case WM_SIZE
        zzSubclassProc = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
        ' here we are checking a couple of things
        ' 1. Form coming out of Maximize state after DPI changed (if not in move/size loop)
        ' 2. MDI child coming out of Minimize state after DPI changed
        ' 3. MDI child moving into Maximize state
        If (m_State And flgMDIchild) = flgMDIchild Then
            ' handle MDI children restore/maximize
            If (m_State And flgMdiMinimized) = flgMdiMinimized Then ' coming out of minimize after DPI change
                If wParam = vbNormal Or wParam = vbMaximized Then
                    m_State = m_State Xor flgMdiMinimized
                    If (m_State And flgIsPmAware) = flgIsPmAware Then
                        uMsg = 0 ' prevent next IF below from triggering; handled here
                        pvInitDpiChangeMdiChild hWnd
                    End If
                End If
            End If
            If wParam = vbMaximized And uMsg = WM_SIZE Then pvAlignMdiChild hWnd
        End If
        If wParam = vbNormal Then
            If (m_State And flgMoveSizeLoop) = 0 Then
                If (m_State And flgDpiMaxChange) = flgDpiMaxChange Then
                    m_State = m_State Xor flgDpiMaxChange
                    ' form coming out of Maximized & DPI changed while maximized
                    ' send event to host to inform of that scenario
                    RaiseEvent RestoreFromMaximized
                End If
            End If
        End If
        If wParam <> vbMinimized Then
            If (m_State And (flgScalingCycle Or flgSyncOcx)) = flgSyncOcx Then
                If (m_State And flgShown) = flgShown Then RaiseEvent ResyncAlignedOcx
            End If
        End If
        EatMessage = True
        
    Case WM_ENTERMOVESIZE               ' form is entering a drag loop
        ' if DPI changes while in loop, a new scaling cycle will start & not end til loop exits
        If (m_State And flgMDIchild) = 0 Then m_State = m_State Or flgMoveSizeLoop
        
    Case WM_EXITMOVESIZE                ' form is exiting drag loop
        zzSubclassProc = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
        m_State = m_State And Not flgMoveSizeLoop ' don't use Xor; flag may not be set
        If (m_State And flgDpiMaxChange) = flgDpiMaxChange Then
            ' if dragged to normal state, then remove flag to inform host
            ' that repositioning needed due to DPI change while maximized
            If (GetWindowLong(hWnd, GWL_STYLE) And (WS_MINIMIZE Or WS_MAXIMIZE)) = 0 Then _
                m_State = m_State Xor flgDpiMaxChange
        End If
        EatMessage = True
        
    Case WM_GETDPISCALEDSIZE    ' permonitor v2 only message
        ' if maximized, system defines window size & this is never sent
        ' if minimized, this will be sent when window restored to normal
        zzSubclassProc = pvGetHostDpiSize(wParam, lParam)
        EatMessage = True
    
    Case WM_DPICHANGED_AFTERPARENT      ' permonitor v2 only message
        ' subclassed only for forms having property: MDIChild=True
        If (m_State And flgIsPmAware) = flgIsPmAware Then
            wParam = GetWindowLong(hWnd, GWL_STYLE) And (WS_MINIMIZE Or WS_MAXIMIZE)
            If wParam = WS_MINIMIZE Then
                ' delay DPI scaling until comes out of minimize
                m_State = m_State Or flgMdiMinimized
            Else
                pvInitDpiChangeMdiChild hWnd ' scale MDI child
            End If
        End If
        EatMessage = True
        
    Case WM_PAINT
        ' looking for first show of the form, only trapped if SyncOcxRequired=True
        If (m_State And flgShown) = 0 Then
            m_State = m_State Or flgShown
            zzSubclassProc = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
            RaiseEvent ResyncAlignedOcx
            EatMessage = True
        End If
    End Select
    
End Function    ' ordinal #2

Private Sub zzRemoteInit(ByVal ptrImplementor As Long, ByVal lDpiProject As Long, ByVal lParam As Long)
    
    ' routine called by clsDpiPmFormLoader when a form implementing IDpiPmAssistant
    ' is loaded. This must be the last private method in this class because it is
    ' called from clsDpiPmFormLoader by the address relative to end of this class.
    
    ' ptrImplementor: ObjPtr to IDpiPmAssistant implementation of loaded form
    ' lDpiProject: project's DPI
    ' lParam: currently unused but scalable for future upgrades
    
    ' We set initial class variables to prepare for form load...
    '   m_Host, m_Caps, m_Context, m_BaseDPIs, m_DPIform, m_State, m_SubclassDPI
    
    Dim oIAssist As IDpiPmAssistant     ' the newly loaded form
    Dim oLoader As clsDpiPmFormLoader   ' used to create subclassing as needed
    
    m_Host = ptrImplementor
    Set oLoader = New clsDpiPmFormLoader
    m_Caps = oLoader.WindowCaps
    m_BaseDPIs = lDpiProject * &H10000  ' place into hiword
    m_DPIform = TWIP_CONST \ Screen.TwipsPerPixelX ' VB's DPI
    
    Set oIAssist = pvGetHost()
    ' determine what awareness context this form is loaded into
    m_Context = Me.TranslateAwarenessContext(Me.ThreadAwareness)
    Select Case m_Context
    Case dpiUNAWARE, dpiUNAWARE_GDISCALED
        m_BaseDPIs = m_BaseDPIs Or BASEDPI_CONST ' loword = system DPI
    Case Else
        m_BaseDPIs = m_BaseDPIs Or Me.DpiForSystem ' loword = system DPI
        If m_Context <> dpiSYSTEM_AWARE Then _
            m_State = m_State Or flgIsPmAware ' flag = pm aware, either v1 or v2
    End Select
    
    If TWIP_CONST / lDpiProject <> Screen.TwipsPerPixelX Then m_State = m_State Or flgSyncOcx
    If (m_State And (flgIsPmAware Or flgSyncOcx)) <> 0 Then
        Set m_SubclassDPI = oLoader.CreateTasker_Subclass(Me, 2) ' create but don't start yet
    End If
    Set oLoader = Nothing
    
    ' execute the impelmented Attach method, passing this class
    oIAssist.Attach Me
    Set oIAssist = Nothing
    
End Sub         ' ordinal #1
'/////////////////////// DO NOT PLACE ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\

