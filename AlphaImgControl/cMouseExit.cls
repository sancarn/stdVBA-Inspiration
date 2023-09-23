VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMouseExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Help track Mouse Enter/Exit ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' Tracking MouseExit on a windowless control isn't always an easy task.
' You don't have a window handle to pass to APIs to monitor for you
'   and you really don't want to subclass the container window
'   nor set a mouse hook if a less intrusive means is available.
' This is that less intrusive means: Timer + SendMessage

' After several different logic approaches used, this one seems to be ideal.
' -- This class is only created once & made Public in a module
' -- The controls declare this class WithEvents and respond to the events
' 1) When control gets mouse move event, it calls this class' Iniitate event if not previously done
'       and triggers its own MouseEnter event as needed
' 2) It then sets a reference to the public copy of this class & doesn't call Initiate again
' 3) The control that had previous reference & received the result of ReleaseCapture, if any,
'       destroys its reference to the public copy of this class & triggers its own MouseExit event
' Timer works like this.
' 1) Get current cursor position
' 2) See if cursor is over the control's container & if not triggers a ReleaseCapture call
' 3) If it is over the container, then this 3 part procedure is used...
'       a) Send control an event telling it that this class will be sending a bogus message
'       b) Control sets a flag & this class sends a fake mouse_move event
'       c) This class then asks control if it got the event & if not, triggers a ReleaseCapture call
' 4) When ReleaseCapture call is made, control releases reference to this class & resets any flags

' Note: Other control instances will know that some other control is tracking the mouse via this
'   class' Owner property. This is because a single instance of this class is Public to the project.
' If one of those other controls gets a mouse move event before the tracking-control gets a
' MouseExit event, then the other control will trigger the tracking-control to fire a MouseExit
' event. This has the advantage of no 2 controls overlapping events. One control will always
' fire a MouseExit event before another conrol fires a MouseEnter event, regardless whether or
' not the timer actually detected the MouseExit.

Public Event MouseExited()
Public Event AnticpateMessage(SetValidation As Boolean, returnValidation As Boolean)

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_MOUSEMOVE As Long = &H200

Private c_Hwnd As Long
Private c_Owner As Long

Public Property Get Owner() As Long
    Owner = c_Owner
End Property

Public Sub Iniitate(ContainerHwnd As Long, ptrOwner As Long)

    If c_Owner <> ptrOwner Then Me.ReleaseMouseCapture True, c_Owner
    c_Owner = ptrOwner                          ' and start a timer (5x per second)
    Me.TimerhWnd = ContainerHwnd
    
End Sub

Public Property Get TimerhWnd() As Long
    TimerhWnd = c_Hwnd
End Property
Public Property Let TimerhWnd(newHwnd As Long)
    If c_Hwnd Then KillTimer c_Hwnd, ObjPtr(Me)
    c_Hwnd = newHwnd    ' cache hWnd & owner reference
    If c_Hwnd Then SetTimer c_Hwnd, ObjPtr(Me), 200&, AddressOf modCommon.CommonTimerProc
End Property

Public Sub ReleaseMouseCapture(CreateEvent As Boolean, Owner As Long)

    Me.TimerhWnd = 0&
    If CreateEvent = True And c_Owner <> 0& Then
        If c_Owner = Owner Or Owner = 0& Then RaiseEvent MouseExited
    End If
    c_Owner = 0&

End Sub

Public Sub MouseInControl()

    ' This event can be called in 3 different scenarios
    ' 1) Another control gets MouseMove, calls this routine to have it fire MouseExit for the client
    ' 2) During control's HitTest event if HitTest is invalid (can be when using HitTest other than the entire control)
    ' 3) From the modCommon.CommonTimerProc as result of a timer event

    Dim mPt As POINTAPI, bValidate As Boolean
    
    If c_Hwnd Then                              ' called via timer
        GetCursorPos mPt                        ' if not, ask control for a hit test result
        RaiseEvent AnticpateMessage(True, bValidate)
        If bValidate Then
            If (WindowFromPoint(mPt.X, mPt.Y) <> c_Hwnd) Then
                Me.ReleaseMouseCapture True, c_Owner
            Else
                ScreenToClient c_Hwnd, mPt
                SendMessage c_Hwnd, WM_MOUSEMOVE, 0&, ByVal mPt.X Or mPt.Y * &H10000
                RaiseEvent AnticpateMessage(False, bValidate)
                If bValidate = False Then Me.ReleaseMouseCapture True, c_Owner
            End If
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Me.ReleaseMouseCapture False, 0&
End Sub
