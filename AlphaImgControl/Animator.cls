VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Animator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'  -----======== PURPOSE: Frame navigation manager for animation ========-----
' Class Instancing = Public (not creatable by users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' Class add animation control to animated image formats
' Used by both the usercontrol and the custom property page

' This class is tied to the image's container window's visibility
' When the window is not visible, MoveForward, MoveBackward do not function
' 1. StartAnimation, PauseAnimation & ResumeAnimation are not processed  while
'       window is invisible. When window becomes visible, then the any previous call
'       to StartAnimation or ResumeAnimation will activate automatically
' 2. If wanting to navigate frames while window is invisible, use the MoveTo method
' 3. The IsSuspendedAnimation method will let you know if animation is halted due to invisibility

' About frame durations. Many animated formats fail to provide a duration for the individual frames
' When this occurs, the class will use a default duration shown as the DEFAULTDURATION constant below.
' This is how the class manages frame durations based on original setting and user-defined settings:

' If original duration is less than 10 then
'   user-defined minimum (if set) is used else default minimum used
' else if duration is less than user-defined minimum
'   then user-defined minimum is used
' else if rate is greater than user-defined maximum, if set, then
'   user-defined maximum is used
' else if duration is greater than 32767
'   then 32767 is used
' else actual duration is used
' To set all frames to a constant duration; set both the minimum & maximum default durations to same value


Public Event AnimationFinished()
Public Event Looped(Count As Long)
Public Event FrameChanged(Index As Long)

Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long

Private m_Image As GDIpImage
Private m_Loops As Long
    ' LoWord: 0=stopped looping else remaining loops. If infinite looping then value is always 65535
    ' HiWord: set loop count by user/default, max of &H7FFF& or &HFFFF&=infinite
Private m_Flags ' 0-2 is the state, 4=end loop on 1st frame, 256=internal flag
Private m_TimerWnd As Long
Private m_MinMaxDuration As Long
Private Const DEFAULTDURATION As Long = 75&

Public Sub MoveForward()

    ' function forwards the animation one frame. Inactive if window is invisible
    ' If the image's container is hidden, animation is suspended until the container becomes visible
    
    If (m_Flags And &H100&) Then                ' checking for window visibility
        If IsWindowVisible(m_TimerWnd) = 0& Then Exit Sub
        
        KillTimer m_TimerWnd, ObjPtr(Me)        ' stop checking
        m_Flags = m_Flags Xor &H100&            ' remove flag
        If (m_Flags And lvicAniCmdStart) Then Me.ResumeAnimation
        
    ElseIf IsWindowVisible(m_TimerWnd) = 0& Then ' did window go invisible?
        m_Flags = m_Flags Or &H100&             ' add flag & start visibility checking
        SetTimer m_TimerWnd, ObjPtr(Me), 100&, AddressOf modCommon.CommonTimerProc
    
    ElseIf (m_Flags And lvicAniCmdStart) = 0& Then ' not animating; just move + 1
        m_Image.ImageIndex = m_Image.ImageIndex + 1&
    
    ElseIf m_Image.ImageCount < 2& Then         ' while animating a new image was assigned
        Me.StopAnimation
        m_Flags = 0&
    
    Else    ' move to next frame or wrap to first frame
        Me.PauseAnimation
        If m_Image.ImageIndex = m_Image.ImageCount Then ' looped
            If m_Loops = -1& Then                   ' infinite looping; move to next frame
                m_Image.ImageIndex = m_Image.ImageIndex + 1&
                RaiseEvent FrameChanged((m_Image.ImageIndex))
            ElseIf (m_Loops And &HFFFF&) - 1& < 1& Then ' loops exceeded?
                Me.StopAnimation
                m_Loops = (m_Loops And &HFFFF0000)
                If (m_Flags And &H4&) Then
                    m_Image.ImageIndex = 1&             ' stop on 1st frame
                    RaiseEvent FrameChanged((m_Image.ImageIndex))
                End If
            Else                                    ' decrement loops
                m_Loops = (m_Loops And &HFFFF0000) Or ((m_Loops And &H7FFF&) - 1&)
                m_Image.ImageIndex = m_Image.ImageIndex + 1&
                RaiseEvent FrameChanged((m_Image.ImageIndex))
            End If
            RaiseEvent Looped((m_Loops And &HFFFF&))
        Else
            m_Image.ImageIndex = m_Image.ImageIndex + 1&
            RaiseEvent FrameChanged((m_Image.ImageIndex))
        End If
        If (m_Flags And lvicAniCmdPause) Then
            Me.ResumeAnimation
        Else
            RaiseEvent AnimationFinished
        End If
    End If
    

End Sub

Public Sub MoveBackward()   ' simply move back a frame or wrap to last frame
    ' never called by the usercontrol or property page. Inactive if window is invisible
    If (m_Flags And &H103&) = lvicAniCmdStart Then
        Me.PauseAnimation
        m_Image.ImageIndex = m_Image.ImageIndex - 1&
        Me.ResumeAnimation
    ElseIf (m_Flags And &H100&) = 0& Then
        m_Image.ImageIndex = m_Image.ImageIndex - 1&
    End If
End Sub

Public Sub MoveTo(ByVal Index As Long) ' moves to a specific frame
    ' never called by the usercontrol or property page
    If (m_Flags And &H103&) = lvicAniCmdStart Then
        Me.PauseAnimation
        m_Image.ImageIndex = Index
        Me.ResumeAnimation
    Else
        m_Image.ImageIndex = Index
    End If
End Sub

Public Sub StartAnimation() ' starts animation after reseting frame to beginning
    
    ' if window is invisible, animation will start when visibility returns
    If m_Image.ImageCount Then
        Me.StopAnimation
        m_Image.ImageIndex = 1&
        If m_Loops > -1& Then m_Loops = (m_Loops And &HFFFF0000) Or ((m_Loops And &H7FFF0000) \ &H10000)
        m_Flags = m_Flags Or lvicAniCmdStart
        If IsWindowVisible(m_TimerWnd) Then
            SetTimer m_TimerWnd, ObjPtr(Me), pvGetDuration, AddressOf modCommon.CommonTimerProc
        Else
            m_Flags = m_Flags Or &H100&
            SetTimer m_TimerWnd, ObjPtr(Me), 100&, AddressOf modCommon.CommonTimerProc
        End If
    End If
End Sub

Public Sub StopAnimation()
    If (m_Flags And (&H100& Or lvicAniCmdStart)) Then ' currently animating or checking window visibility
        KillTimer m_TimerWnd, ObjPtr(Me)
    End If
    m_Flags = (m_Flags And Not &H103&)
End Sub

Public Sub ResumeAnimation() ' resumes animation from current frame

    ' if window is invisible, animation will resume when visibility returns
    If (m_Flags And &H3&) = lvicAniCmdPause Then
        If m_Image.ImageCount > 1& Then
            m_Flags = (m_Flags And Not &H3&) Or lvicAniCmdStart
            If (m_Flags And &H100&) = 0& Then ' else container is not visible
                SetTimer m_TimerWnd, ObjPtr(Me), pvGetDuration, AddressOf modCommon.CommonTimerProc
            End If
        Else
            Me.StopAnimation
        End If
    ElseIf (m_Flags And &H3&) = lvicAniCmdStop Then
        Me.StartAnimation
    End If
End Sub

Public Sub PauseAnimation() ' pauses any current animation
    
    If (m_Flags And (&H100& Or lvicAniCmdStart)) Then ' currently animating or checking window visibility
        KillTimer m_TimerWnd, ObjPtr(Me)
        m_Flags = (m_Flags And Not &H103&) Or lvicAniCmdPause
    ElseIf m_Image.ImageCount Then
        m_Flags = (m_Flags And Not &H3&) Or lvicAniCmdPause
    End If
    
End Sub

Public Property Get IsSuspendedAnimation() As Boolean
    ' animation is automatically suspended whenever the image's container window becomes invisible
    ' This method will return whether that window is invisible or not
    IsSuspendedAnimation = CBool(m_Flags And &H100&)
End Property

Public Property Get AnimationState() As AnimationActionEnum
    ' return the current animation state: stopped, playing, paused
    AnimationState = (m_Flags And &H3&)
End Property

Public Property Get CurrentFrame() As Long
    ' returns the current frame
    CurrentFrame = m_Image.ImageIndex
End Property

Public Property Get FrameCount() As Long
    ' returns the number of frames being animated
    FrameCount = m_Image.ImageCount
End Property

Public Property Get FrameDuration(Index As Long) As Long
    ' returns the frame duration of any frame. Passing zero uses the current frame
    If Not (Index < 0& Or Index > m_Image.ImageCount) Then FrameDuration = m_Image.FrameDuration(Index)
End Property

Public Property Get LoopCount() As Long
    ' returns the number of loops used in this animation. If infinite looping return value is 0
    If m_Loops > -1& Then LoopCount = (m_Loops And &H7FFF0000) \ &H10000
End Property
Public Property Let LoopCount(newValue As Long)
    ' sets the number of loops used for animation. Pass zero to make looping infinite
    If newValue < 1& Or newValue > &H7FFF& Then
        m_Loops = -1&
    ElseIf newValue < (m_Loops And &H7FFF&) Then
        m_Loops = (newValue * &H10000 Or 1&)
    Else
        m_Loops = (newValue * &H10000) Or (m_Loops And &HFFFF&)
    End If
End Property

Public Sub RestoreOriginalLoopCount()
    ' returns the original loop count from the image's resource vs. what may have been set by you
    ' via the LoopCount property
    If m_Image.AnimationLoops < 1& Then
        m_Loops = -1&
    Else
        m_Loops = (m_Image.AnimationLoops And &H7FFF&) * &H10000 Or (m_Image.AnimationLoops And &H7FFF&)
    End If
End Sub

Public Property Get LoopsRemaining() As Long
    ' returns the number of loops remaining. If infinite looping return value is 65535
    LoopsRemaining = (m_Loops And &HFFFF&)
End Property

Public Property Get DefaultMinimumDuration() As Long
    ' returns the absolute minimal frame duration used during animation
    DefaultMinimumDuration = (m_MinMaxDuration And &H7FFF&)
End Property
Public Property Let DefaultMinimumDuration(newValue As Long)
    ' sets/returns the user-defined absolute minimal frame duration used during animation
    If newValue > 9& Then
        m_MinMaxDuration = (m_MinMaxDuration And &H7FFF0000) Or (newValue And &H7FFF&)
    Else
        m_MinMaxDuration = (m_MinMaxDuration And &H7FFF0000)
    End If
End Property

Public Property Get DefaultMaximumDuration() As Long
    ' sets/returns the user-defined absolute maximum frame duration used during animation
    DefaultMaximumDuration = (m_MinMaxDuration And &H7FFF0000) \ &H10000
End Property
Public Property Let DefaultMaximumDuration(newValue As Long)
    If newValue > 9& Then
        m_MinMaxDuration = (m_MinMaxDuration And &HFFFF&) Or ((newValue And &H7FFF&) * &H10000)
    Else
        m_MinMaxDuration = (m_MinMaxDuration And &HFFFF&)
    End If
End Property

Public Property Get EndLoopOnFirstFrame() As Boolean
    ' default behavior is to have last frame displayed when animation terminates due to loop count met
    ' If desired to display 1st frame instead of last frame, set this property to true
    EndLoopOnFirstFrame = CBool(m_Flags And &H4&)
End Property
Public Property Let EndLoopOnFirstFrame(newValue As Boolean)
    If newValue <> Me.EndLoopOnFirstFrame Then m_Flags = m_Flags Xor &H4
End Property

Private Function pvGetDuration() As Long
    
    ' helper function to calculate a frame's duration. See notes at top of this class
    
    Dim lMinMax As Long
    
    pvGetDuration = m_Image.FrameDuration(0&)
    lMinMax = (m_MinMaxDuration And &HFFFF&)
    If pvGetDuration < 10& Then
        If lMinMax = 0& Then pvGetDuration = DEFAULTDURATION Else pvGetDuration = lMinMax
    ElseIf pvGetDuration < lMinMax Then
        pvGetDuration = lMinMax
    Else
        lMinMax = (m_MinMaxDuration And &H7FFF0000) \ &H10000
        If lMinMax Then
            If pvGetDuration > lMinMax Then pvGetDuration = lMinMax
        ElseIf lMinMax > &H7FFF& Then
            pvGetDuration = &H7FFF&
        End If
    End If

End Function

Private Sub Class_Initialize()
    If Not g_NewImageData Is Nothing Then
        g_NewImageData.CacheSourceInfo m_Image, m_TimerWnd, 0&, False, True
        Set g_NewImageData = Nothing
        Me.RestoreOriginalLoopCount
    End If
End Sub

Private Sub Class_Terminate()
    Me.StopAnimation
    Set m_Image = Nothing
End Sub
