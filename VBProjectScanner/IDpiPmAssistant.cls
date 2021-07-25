VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDpiPmAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------------

' You must implement this class and declare a clsDpiPmAssistant class in each form
'   in order for that form to receive and adjust to DPI changes. The
'   clsDpiPmAssistant class is a companion class to this class.

' All your forms, including MDIs and MDI children will include these common lines of code
' _________________________________________________________________________________________

'   Implements IDpiPmAssistant
'   Dim WithEvents cDpiPmAssist As clsDpiPmAssist ' change variable name as desired

'   Private Sub IDpiPmAssistant_Attach(oDpiPmAssist As clsDpiPmAssist)
'       Set cDpiPmAssist = oDpiPmAssist ' variable matches what you declared in Declarations
'   End Sub

'   Private Sub Form_Load()
'       ... custom pre-initializing code as needed
'       cDpiPmAssist.Activate Me  ' variable matches what you declared in Declarations
'       ... custom post-initializing code as needed
'   End Sub
' _________________________________________________________________________________________

' Of the implemented methods in this class, only one is absolutely required to be
'   coded. The others depend on how you want to deal with scaling due to DPI changes.
' The only method required to be coded is Attach, as shown above.

' In addition, the project must be started with Sub Main() not with an existing project form.
' Why? Your project may be started in a DPI other than the system DPI. The form may need to
'   be scaled to that DPI and depending on the O/S version, one or more actions may need to
'   be executed before your form is fully created/displayed. Also, these classes need to
'   know the DPI context that the project was loaded in. By creating this class in
'   Sub Main before any changes to DPI contexts are made ensures the project's DPI is known.
' To set startup to Sub Main, use "Project | [ProjectName] Properties" IDE menu

' Your Sub Main() will obviously be in a module and a typical startup module looks like:
' _________________________________________________________________________________________

'   Dim oFormLoader As clsDpiPmFormLoader

'   Private Sub Main()      ' can be Public if desired
'       Set oFormLoader = New clsDpiPmFormLoader
'       oFormLoader.Activate
'       Form1.Show          ' or whatever the project's initial form/procedure name is
'   End Sub
' _________________________________________________________________________________________

'-----------------------------------------------------------------------------------------------

Public Enum DpiActionEnum
    dpiAsst_BeginEvent = 0          ' scaling event starts
    dpiAsst_EndEvent                ' scaling event ends
End Enum
Public Enum DpiScaleCycleEnum
    dpiAsst_BeginCycleHost = 0      ' host (form) scaling begins
    dpiAsst_BeginCycleControls      ' host controls scaling begins
    dpiAsst_EndCycleControls        ' host controls scaling ends, all controls
    dpiAsst_EndCycleHost            ' host (form) scaling ends
End Enum
Public Enum DpiScaleReason
    dpiAsst_DpiChanged = 0          ' DPI changed due to system settings
    dpiAsst_DpiChangeWhileDragging  ' DPI changed due to dragging to another monitor
    dpiAsst_InitialLoad             ' form's initial load, during cDpiPmAssist.Activate
End Enum
Public Enum MsgTiming               ' See Subclasser event
    BeforeAfter = 0                 ' before & after subclassed hWnd gets a message
    BeforeWindow                    ' only before the subclassed hWnd gets a message
    AfterWindow                     ' only after the subclassed hWnd gets a message
End Enum

Public Sub Attach(ByRef oDpiPmAssist As clsDpiPmAssist)
    ' Sent to the form that implements this class when that form is first loaded.
    ' When event is received, cache the passed class otherwise no DPI tracking is
    '   performed and no other events from this class will be sent.
End Sub

Public Function DpiScalingCycle(ByVal Reason As DpiScaleReason, ByVal Action As DpiScaleCycleEnum, _
                                ByVal OldDPI As Long, ByVal NewDPI As Long, _
                                ByRef userParams As Variant) As Long
    ' Triggered when a scaling cycle begins. This event will occur 4 times
    '   per DPI change, each with a different Action parameter.
    ' Return value only applies when Action = dpiAsst_BeginCycleControls
    
    ' A standard scaling cycle looks a bit like this:
    '   DpiScalingCycle sent with Action = dpiAsst_BeginCycleHost
    '       next event is: ScaleHost
    '   DpiScalingCycle sent with Action = dpiAsst_BeginCycleControls
    '       return non-zero to abort; if not aborted, next events are:
    '       - IncludeSetParentControls: occurs only once
    '       - ScaleControlVB: occurs for each VB intrinsic control
    '       - ScaleControlOCX: occurs for each non-intrinsic control
    '   DpiScalingCycle sent with Action = dpiAsst_EndCycleControls
    '       manually scale any API windows as needed
    '   DpiScalingCycle sent with Action = dpiAsst_EndCycleHost
    
    ' Reason: does not change between Actions
    '   dpiAsst_DpiChanged: DPI changed for other reasons than dragging/resizing the form
    '   dpiAsst_DpiChangeWhileDragging: DPI changed while form is being moved/sized by dragging
    '   dpiAsst_InitialLoad: form is currently loading for the first time
    
    ' Action:
    '   dpiAsst_BeginCycleHost: scaling cycle begins
    '       perform any custom actions for the scaling cycle, if needed
    '       ScaleHost event follows and allows changing the scaled host dimensions
    '   dpiAsst_BeginCycleControls: controls are about to be scaled, host is scaled
    '       set flags to prevent code execution that would change control dimensions/positions
    '       return non-zero to prevent scaling of any controls else following are sent:
    '       - IncludeSetParentControls event will immediately follow
    '       - ScaleControlVB or ScaleControlOCX events follow for each control
    '       Note: if any controls are per-monitor aware, they can receive DPI
    '       change messages from Windows. So, the controls may partially scale anyway
    '       even if this event is aborted. The only way to prevent this is to subclass
    '       such controls and prevent DPI change messages from getting to them.
    '   dpiAsst_EndCycleControls: all controls have been scaled, if applicable
    '       release flags that prevented code execution that would change control dimensions/positions
    '       scale any API child windows you may have created
    '   dpiAsst_EndCycleHost: scaling cycle ended
    '       clean up any custom actions you performed, if needed
    '       reposition the form as needed, based on the Reason parameter & min/max state
    '       ------------------------------------------------------------------------------
    '       ...
    '       If Reason <> dpiAsst_DpiChangeWhileDragging Then
    '           If Me.WindowState = vbNormal Then
    '               ... reposition if needed
    '           End If
    '       End If
    '       ------------------------------------------------------------------------------
    
    ' OldDPI: The DPI previously applied to the form
    ' NewDPI: The DPI being applied to the form
    '   - scale ratio to new DPI is: NewDPI / OldDPI
                           
    ' userParams: user-defined values
    '   set this during dpiAsst_BeginCycleHost, passed to each subsequent event
    '   if more than one value needs to be passed, use: Array(value1, value2, etc)
    '   can set this parameter to a class, collection, and more.
End Function

Public Function ScaleControlVB(ByRef theControl As Control, ByVal Reason As DpiScaleReason, _
                            ByVal Action As DpiActionEnum, ByVal ScaleRatio As Single, _
                            ByRef newX As Single, ByRef newY As Single, _
                            ByRef newCx As Single, ByRef newCy As Single, _
                            ByRef userParams As Variant) As Long
    ' Called during scaling cycle to allow control resizing/scaling
    ' DO NOT allow any code to execute that changes the control's position/size
    '   during this event, regardless of the Action value. If you scale them anyway,
    '   then do not allow them to be scaled here
    ' This event is only sent for VB intrinsic controls
    
    ' theControl: an instance of the control to be scaled
    
    ' Reason: Same as sent to DpiScalingCycle (see that event for comments)
    
    ' Action:
    '   dpiAsst_BeginEvent: control is about to be scaled
    '       you can change the position and/or dimensions
    '       return non-zero to prevent scaling the control and font if applies
    '   dpiAsst_EndEvent: control is scaled, if applicable
    
    ' ScaleRatio: The change in scale from previous DPI to current DPI
    '   ratio can be 1! when Reason is dpiAsst_InitialLoad
    
    ' newX, newY: control position at new DPI. Can be changed
    ' newCx, newCy: control dimensions at new DPI. Can be changed
    '   position/dimensions are in control's container's scale mode
    '   fonts are scaled, images are not
    
    ' userParams: Same as sent to DpiScalingCycle (see that event for comments)
End Function

Public Function ScaleControlOCX(ByRef theControl As Control, ByVal Reason As DpiScaleReason, _
                            ByVal Action As DpiActionEnum, ByVal ScaleRatio As Single, _
                            ByRef newX As Single, ByRef newY As Single, _
                            ByRef newCx As Single, ByRef newCy As Single, _
                            ByRef fontProperties As String, ByRef userParams As Variant) As Long
    ' Called during scaling cycle to allow control resizing/scaling
    ' DO NOT allow any code to execute that changes the control's position/size
    '   during this event, regardless of the Action value. If you scale them anyway,
    '   then do not allow them to be scaled here
    ' This event is sent for all controls except VB intrinsic controls
    ' The only OCX that is completely handled is the SSTab control
    
    ' theControl: an instance of the control to be scaled
    
    ' Reason: Same as sent to DpiScalingCycle (see that event for comments)
    
    ' Action:
    '   dpiAsst_BeginEvent: control is about to be scaled
    '       you can change the position and/or dimensions
    '       return non-zero to prevent scaling the control
    '       Note: if any controls are per-monitor aware, they will receive DPI
    '       change messages from Windows. So, the controls may scale anyway. The
    '       only way to prevent this is to subclass such controls and prevent DPI
    '       change messages from getting to them.
    '   dpiAsst_EndEvent: control is scaled, if applicable
    
    ' ScaleRatio: The change in scale from previous DPI to new DPI
    '   ratio can be 1! when Reason is dpiAsst_InitialLoad
    
    ' newX, newY: control position at new DPI. Can be changed
    ' newCx, newCy: control dimensions at new DPI. Can be changed
    '   position/dimensions are in control's container's scale mode
    '   if newCx or newCy are zero, then all 4 parameters are zero and
    '       you must position and scale the control yourself. Those parameters
    '       will only be zero if the control is missing one or more of these
    '       properties: Left, Top, Width and/or Height
    
    ' fontProperties: a comma-delimited list of public font property names to be scaled
    '   If this list is not provided, no font properties of the control will be scaled
    '   The SSTab control will have this parameter filled in, all others will not.
    
    ' userParams: Same as sent to DpiScalingCycle (see that event for comments)
End Function

Public Sub IncludeSetParentControls(ByRef theControls As VBA.Collection, ByVal Reason As DpiScaleReason, _
                                    ByVal ScaleRatio As Single, ByRef userParams As Variant)
    ' Called once during scaling cycles, immediately after DpiScalingCycle is called
    ' when its Action parameter set to: dpiAsst_BeginCycleControls
    
    ' This event allows you to include controls (not API windows) that exist on
    '   your form that have been added from a different form via SetParent. When
    '   scaling cycles occur, the form's Controls collection dictates which controls
    '   are listed for scaling. That collection does not include any other form's
    '   controls that you have added to your form @ runtime via SetParent.

    ' theControls: assign like so. theControls.Add Form2.Text1
    '   you should list containers before child controls of that container
    '   the passed collection is initialized with zero items
    
    ' other parameters: Same as sent to DpiScalingCycle (see that event for comments)
End Sub

Public Sub ScaleHost(ByVal Reason As DpiScaleReason, _
                          ByRef TwipsWidth As Single, ByRef TwipsHeight As Single, _
                          ByVal OldDPI As Long, ByVal NewDPI As Long, _
                          ByRef IncludeSplashControl As Boolean)
    ' Called once during scaling cycle
    
    ' If the host is not to be scaled, set the passed TwipsWidth,TwipsHeight
    '   parameters to the host's current dimensions. Do note that if NewDPI is
    '   not same as OldDPI, the system will scale the non-client area of the host
    
    ' The host Font property is always scaled. If you don't want this scaled, you
    '   need to unscale it at some point during the scaling cycle.
    
    ' TwipsWidth,TwipsHeight are the precalculated dimensions at the new DPI. If
    '   these are changed, ensure they are relative to the new DPI and in twips.
    '   - scale ratio is: NewDPI / OldDPI
    
    ' IncludeSplashControl: False (default) does not create a splash control,
    '   otherwise a splash control will be created. The splash control is
    '   an empty PictureBox covering entire client area of the form as displayed
    '   during scaling and removed afterwards. Its back color is vbWindowBackground
    '   if desired, sample usage: IncludeSplashControl = Me.Visible
    
    ' other parameters: Same as sent to DpiScalingCycle (see that event for comments)
    
End Sub

Public Function Subclasser(ByRef EventValue As Long, ByVal BeforeHwnd As Boolean, _
                            ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) As Boolean
    ' Called only if clsDpiPmAssist.SubclassHwnd method is called
    
    ' Function return value: Set to True to prevent hWnd from getting the message
    '   If set to True, you must ensure the EventValue parameter is also set.
    
    ' EventValue: Ignored unless function returns True
    '   Set this to the value to return to Windows for the subclassed message
    '   only if wanting to prevent the event from being passed to hWnd
    ' BeforeHwnd:
    '   If True, then hWnd has not yet seen this message
    '       hWnd can be prevented from getting this message
    '       if hWnd prevented from getting message, no "after" event is triggered
    '   If False, hWnd has processed the message, this is an "after" event
    ' Other parameters: self-explanatory
    
    ' To send a message directly to the hWnd's base class, use DefWindowProc API
    '   EventValue = DefWindowProc(hWnd, uMsg, wParam, lParam)
    '   IDpiPmAssistant_Subclasser = True ' prevent further processing
    
    ' All subclassed hWnds come to this event. You may want to organize this event
    ' like one of the two scenarios below:
    
    ' Scenario 1, typically used when only hWnds of one window class are subclassed
    '   If BeforeWindow = True Then
    '       Select Case uMsg
    '       ...
    '       End Select
    '   Else    ' "after" event
    '       Select Case uMsg
    '       ...
    '       End Select
    '   End If
    
    ' Scenario 2 applies when multiple window classes are subclassed
    '   Select Case hWnd
    '       Case ListView1.hWnd, ListView2.Hwnd
    '           IDpiPmAssistant_Subclasser = myLstViewSubclass(EventValue, BeforeHwnd, hWnd, uMsg, wParam, lParam)
    '       Case Text1.hWnd
    '           IDpiPmAssistant_Subclasser = myTxtBoxSubclass(EventValue, BeforeHwnd, hWnd, uMsg, wParam, lParam)
    '       Case Me.hWnd
    '           IDpiPmAssistant_Subclasser = myFormSubclass(EventValue, BeforeHwnd, hWnd, uMsg, wParam, lParam)
    '   End Select
End Function


