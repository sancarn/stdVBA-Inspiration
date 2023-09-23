VERSION 5.00
Begin VB.UserControl AlphaImgCtl 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ClipBehavior    =   0  'None
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   MaskColor       =   &H00000000&
   PropertyPages   =   "AlphaImgCtl.ctx":0000
   ScaleHeight     =   1455
   ScaleWidth      =   1455
   ToolboxBitmap   =   "AlphaImgCtl.ctx":001D
   Windowless      =   -1  'True
End
Attribute VB_Name = "AlphaImgCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ACKNOWLEDGEMENTS:
' /////////////////////////////////////////////////////////////////////////////////////
' Adaptive Palette Color Quantization - http://www.microsoft.com/msj/archive/s3f1a.htm
' Path Containing Round Rectangle - http://www.eggheadcafe.com/software/aspnet/29469337/bob-powells-how-to-draw-a-rounded-rectangle.aspx
' Point In Polygon Algorithm - http://www.visibone.com/inpoly/
' Low Level COM Calls Via DispCallFunc - http://msdn2.microsoft.com/en-us/library/ms688421.aspx
' Base64 Extraction Algorithm - http://www.vbforums.com/showthread.php?t=498548
' MetaFile Header Information - http://wvware.sourceforge.net/caolan/ora-wmf.html
' GDI+ Image Format GUIDs - http://com.it-berater.org/gdiplus/noframes/GdiPlus_constants.htm
' /////////////////////////////////////////////////////////////////////////////////////

' Abbreviated documentation for this control is contained in the code.
' All routines/properties that are not prefixed with the 3 letters "spt" are public
' The usercontrol is mostly a GUI wrapper around the GDIpImage class

' Jump to very end for a complete change history
' ///////////////////////////////////////////////////////////////////////////////////
' Source code maintained at: http://www.vbforums.com/showthread.php?t=630193

' You must think of this control differently than VB's image control.
' This control is more of a container for an image. The image can be positioned anywhere
'   inside the control via use of the AlignCenter and XOffset/YOffset properties
' Additionally, special events allow you to draw behind the image and on top of the image
'   without actually modifying the image itself
' ///////////////////////////////////////////////////////////////////////////////////

' The usercontrol has procedure attributes applied for every public property, method, event
' You will not see them in this code while in IDE, but you will see them if viewed in Notepad.
' Copying any code to another window will not transfer the procedure attributes. Rearranging
' the code within this page can/will result in the attributes being deleted when project is saved.

Public Event AsyncDownloadDone(Success As Boolean, ErrorCode As Long)
Attribute AsyncDownloadDone.VB_Description = "Event that occurs once async download either succedes or fails"
' ^^ raised when downloading image to control and image finishes or fails
Public Event AsyncDownloadDoneBkgImg(Success As Boolean, ErrorCode As Long)
Attribute AsyncDownloadDoneBkgImg.VB_Description = "Event that occurs once async download either succedes or fails for background image only"
' ^^ same as AsyncDownloadDone but applies only to the BkgImage property

Public Event UpdateDataboundImage(theImage As GDIpImage)
' ^^ occurs just before a databound control's image updates a database with changes
' -- can change/modify what will be saved by setting theImage to another image (i.e., SavePictureGDIplus call)

' animation-specific events: GIF, Cursors, Segmented images
Public Event AnimationLoopsFinished()
Attribute AnimationLoopsFinished.VB_Description = "Occurs when animation terminated due to loops/cycles completing"
    '^^ Only called when animation stops as a result of the number of sequence loops met
Public Event AnimationFrameChanged(Frame As Long)
Attribute AnimationFrameChanged.VB_Description = "Event that indicates animation progress"
    '^^ Called when animating only and frame changes
Public Event Changed()
Attribute Changed.VB_Description = "Occurs when a new .Picture property is set"
    '^^ Called when the .Picture property changes due to drag/drop, copy/paste, setting .Picture property
Public Event PrePaint(hDC As Long, Left As Long, Top As Long, Width As Long, Height As Long, HitTestRgn As Long, Cancel As Boolean)
    '^^ Raised before image painted if WantPrePpostEvents=True. Cancel when set to True prevents image from being drawn
    '   can select anything into passed DC for custom drawing; must remove those selected objects too
    '   Exception: If adding a clipping region, can leave it in the DC; it will be cleaned up after PostPaint event
Public Event PostPaint(hDC As Long, Left As Long, Top As Long, Width As Long, Height As Long, HitTestRgn As Long)
    '^^ Raised after image is painted but before it is transfered to the control, if WantPrePpostEvents=True
    '   can select anything into passed DC for custom drawing; must remove those selected objects too
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, Cancel As Boolean)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual"
    '^^ Raised whenever something is dropped onto the control (OLEDropMode<>None). Cancel prevents it from being used
    '^^ When files are dropped, can modify the DataObject to remove all files except the one to be used for the control/image
' Following OLE events are only raised if OLEDragMode and/or OLEDropMode are set to Manual
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual.\n"
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled"
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed.\n"
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event"
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically.\n"
' Following are typical mouse events & forwarded only if WantMouseEvents=True
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus"
Public Event MouseEnter()   ' Determined by HitTest property settings
Attribute MouseEnter.VB_Description = "Occurs when mouse first enters the hit test area"
Public Event MouseExit()    ' MouseExit will not be received if no MouseEnter event occurred
Attribute MouseExit.VB_Description = "Occurs when mouse leaves the hit test area"
' Following are always forwarded
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object.\n"
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object.\n"
Attribute Click.VB_MemberFlags = "200"
' ^^ This event should be the default user gets when viewing code for the control
' To verify it is default action
'   1. IDE Menu: Tools | Procedure Attributes
'   2. Find Click in the "Name" combobox
'   3. Click the "Advanced" button
'   4. Check the box: "User Interface Default"
'   5. Click the "Apply" button
'////////////////////////////////////////////////////////////////////////////////////////


Public Enum HitTestModeEnum
    lvicPerimeter = 0               ' entire control is the hit test
    lvicEntirelImage = 1            ' just the image boundaries
    lvicTrimmedImage = 2            ' tight rectangle including all non-transparent pixels
    lvicRunTimeRegion = 3           ' can pass a valid region during runtime only
End Enum
Public Enum AutoSizeEnum
    lvicNoAutoSize = 0              ' no Auto-Sizing
    lvicSingleAngle = 1             ' control sized to current rendered image at current angle
    lvicMultiAngle = 2              ' control sized to allow rendering at all angles
End Enum
Public Enum ScalingRatioEnum
    lvicActualSize = 0              ' image is drawn full scale
    lvicStretch = 1                 ' image is stretched to control's dimensions
    lvicScaled = 2                  ' image is scaled to control's dimensions
    lvicScaleDownOnly = 3           ' image is scaled down if needed, never a ratio > 1:1
    lvicFixedSize = 4               ' image size is fixed, will not be resized until Aspect parameter changes
    lvicFixedSizeStretched = 5
End Enum
Public Enum TransparentColorModeEnum
    lvicNoTransparentColor = 0      ' no color specified as transparent
    lvicTransparentTopLeft = 1      ' use the color at top/left corner
    lvicTransparentTopRight = 2     ' use the color at top/right corner
    lvicTransparentBottomRight = 3  ' use the color at bottom/right corner
    lvicTransparentBottomLeft = 4   ' use the color at bottom/left corner
    lvicUseTransparentColor = 5     ' use the value in the TransparentColor property
End Enum
Public Enum BorderShapeEnum
    lvicNoBorder = 0                ' forces Border property=False
    lvicRectangular = 1             ' rectangular border & Border property=True
    lvicRoundedCorners = 2          ' blended rounded corner border & Border property=True
    lvicRoundedCornersRough = 3     ' non-blended rounded corner border & Border property=True
    lvicUserDefinedBorder = 4       ' reserved space for user-drawn border & Border property=True
End Enum
Public Enum GradientStyleEnum       ' to reverse gradient direction, swap out BackColor & GradientForeColor
    lvicNoGradient = 0
    lvicGradientHorizontal = 1      ' gradient from left to right
    lvicGradientVertical = 2        ' gradient from top to bottom
    lvicGradientDiagonalDown = 3    ' gradient from top/left to bottom/right
    lvicGradientDiagonalUp = 4      ' gradient from bottom/left to top/right
End Enum

' ///// Following require destruction \\\\\\\\
Private m_DC As cDeviceContext                  ' used on demand: FastRedraw=True, WantPrePpostEvents=True, or SetRedraw=False
Private WithEvents m_Image As GDIpImage         ' our image
Attribute m_Image.VB_VarHelpID = -1
Private WithEvents m_Animator As Animator       ' an animator class to help automate animation
Attribute m_Animator.VB_VarHelpID = -1
Private WithEvents m_MouseTracker As cMouseExit ' used to track MouseEnter & MouseExit events
Attribute m_MouseTracker.VB_VarHelpID = -1
Private WithEvents m_Effects As GDIpEffects     ' used to store/create image attributes and v1.1 effects
Attribute m_Effects.VB_VarHelpID = -1
Private WithEvents m_BkgImage As GDIpImage      ' optional background image
Attribute m_BkgImage.VB_VarHelpID = -1
' /////////////////////\\\\\\\\\\\\\\\\\\\\\\\\

Private m_bAffects() As Byte                    ' ensures set effects do not get erased if uncompiled project moved to non-v1.1 GDI+ system
Private m_Flags As AttributeFlagsEnum           ' non-specific attributes/property values
Private m_RenderFlags As RenderFlagsEnum        ' rendering specific attributes/property values
Private m_CntFlags As ContainerFlagsEnum        ' container-specific attributes/properties
Private m_HitRegion As Long                     ' optional user-defined hit test region
Private m_HitTestPts() As POINTAPI              ' rectangle points, rotated or not, used for hit testing
Private m_Offsets As POINTAPI                   ' optional X,Y rendering offsets
Private m_ClipRect As RECTI                     ' optional cropping rectangle (See SetClipRect)
Private m_Attributes As Long                    ' 1st byte=effects, 2=transmode, 3=mirror, 4=interpolation
Private m_Angle As Single                       ' positive angles are clockwise, negative are counterclockwise
Private m_Size As ScalerStruct                  ' cached scaled dimensions
Private m_DragDrop As DragDropStruct            ' used for drag/drop events

'//// The deault property for this control
Public Property Get Picture() As GDIpImage
Attribute Picture.VB_Description = "Returns/sets a graphic object to be displayed in a control."
Attribute Picture.VB_UserMemId = 0
    ' must be default property. Do not move/rearrange this property from within this code
    ' To verify/set to default
    '   1. IDE Menu: Tools | Procedure Attributes
    '   2. Find Picture in the "Name" combobox
    '   3. Click the "Advanced" button
    '   4. Select "(Default)" in the "Procedure ID" combobox
    '   5. Click the "Apply" button
    Set Picture = m_Image
End Property
Public Property Let Picture(newImage As GDIpImage)
    Set Me.Picture = newImage
End Property
Public Property Set Picture(newImage As GDIpImage)
    ' To share images between controls: Set AlphaImgCtl2.Picture = AlphaImageCtl1.Picture
    ' To destroy control's image: Set AlphaImgCtl1.Picture = Nothing
    
    ' About sharing images....
    ' When Picture property later set, any other controls sharing the superseded image are not affected
    ' Animating or changing Image/Group indexes effects all controls sharing that images
    ' Segmenting a shared image effects all controls sharing that image
    
    Dim lValue As Long
    If Not newImage Is m_Image Then
        Set m_Animator = Nothing                        ' stop animation
        If Not m_MouseTracker Is Nothing Then           ' stop MouseExit event for now
            If m_MouseTracker.Owner = ObjPtr(Me) Then m_MouseTracker.TimerhWnd = 0&
        End If                                          ' assign new image & set render flags
        If newImage Is Nothing Then
            Set m_Image = New GDIpImage
        Else
            Set m_Image = newImage
            m_CntFlags = m_CntFlags Or cnt_InitLoad
            Me.TransparentColorMode = lvicNoTransparentColor
            m_CntFlags = m_CntFlags Xor cnt_InitLoad
        End If
        m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest Or render_DoResize
        m_Size.Width = 0&
        sptGetScaledSizes Me.SetRedraw ' resize control/image as needed & start animation if flag set
        If Me.SetRedraw = False Then Me.SetRedraw = True
        If (m_Flags And attr_AutoAnimate) Then Me.Animate lvicAniCmdStart
        If Not m_MouseTracker Is Nothing Then           ' resume tracking MouseExit event
            If m_MouseTracker.Owner = ObjPtr(Me) Then m_MouseTracker.TimerhWnd = UserControl.ContainerHwnd
        End If
        PropertyChanged "ImageIndex"
        ' added to support databound images
        If (m_CntFlags And cnt_DatabaseImage) Then          ' image just loaded from SinkImageData LET
            m_CntFlags = m_CntFlags Xor cnt_DatabaseImage   ' remove flag & continue with no change
        ElseIf CanPropertyChange("SinkImageData") Then      ' if not async download, update database
            If m_Image.AsyncDownloadURL = vbNullString Then PropertyChanged "SinkImageData"
        End If
        RaiseEvent Changed
    End If
    
End Property

'/// A little credit for myself
Public Sub About()
Attribute About.VB_Description = "Author acknowledgement"
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    MsgBox "Created by LaVolpe" & vbCrLf & "Version " & _
        CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), _
        vbInformation + vbOKOnly, "Alpha Image Control"
End Sub

'//// Determines if image is top/left or center aligned in control
Public Property Get AlignCenter() As Boolean
Attribute AlignCenter.VB_Description = "Returns/sets if image is drawn from center or left edge of control"
    AlignCenter = (m_CntFlags And cnt_AlignCenter)
End Property
Public Property Let AlignCenter(newValue As Boolean)
    If Not newValue = Me.AlignCenter Then
        m_CntFlags = m_CntFlags Xor cnt_AlignCenter     ' alignment effects hit test; fix it
        m_RenderFlags = m_RenderFlags Or render_DoHitTest
        If Me.SetRedraw = True Then UserControl.Refresh ' refresh control
        PropertyChanged "AlignCenter"
    End If
End Property

Public Property Get Animate2() As Animator
Attribute Animate2.VB_Description = "Controls animation of the image frames. Class-based method."

    ' Class bound animation control. Supersedes Animate function
    
    ' IMPORTANT: All instances sharing the Image property are affected
    If g_TokenClass.Token = 0! Then Exit Property
    If m_Animator Is Nothing Then
        Set g_NewImageData = New cGDIpMultiImage
        g_NewImageData.CacheSourceInfo m_Image, UserControl.ContainerHwnd, 0&, False, False
        Set m_Animator = New Animator
    End If
    Set Animate2 = m_Animator

End Property

Public Function Animate(ByVal Action As AnimationActionEnum, Optional ByVal DurationLoopCount As Long) As Long
Attribute Animate.VB_Description = "Controls animation of the image frames"

'   IMPORTANT: All instances sharing the Image property are affected

' OBSOLETE. REMAINS FOR BACKWARD COMPATIBILITY ONLY
' Replaced by Animate2 property above

    ' return value dependent on Action parameter and animation state. See below
    ' DurationLoopCount parameter only valid when
    '   Action = lvicAniCmdSetMinDuration or lvicAniCmdSetMaxDuration or lvicAniCmdSetLoopCount
    ' Note: To force animation at a constant speed, all frames, set both min/max durations to same value
    '   Animate lvicAniCmdSetMinDuration, Speed
    '   Animate lvicAniCmdSetMaxDuration, Speed
    ' To reset minimum and maximum defaults, set each to zero
    
    Select Case Action
    Case lvicAniCmdResume, lvicAniCmdStart
        If (m_CntFlags And cnt_Runtime) = 0& Then
            Call sptSetUserMode
            If (m_CntFlags And cnt_Runtime) = 0& Then Exit Function
        End If
    End Select

    Select Case Action
        Case lvicAniCmdStop                         ' returns 0 if not currently animating/paused,
            Me.Animate2.StopAnimation
        Case lvicAniCmdPause                        ' returns 0 if not currently animating/paused,
            Me.Animate2.PauseAnimation
        Case lvicAniCmdResume                       ' if not currently animating, starts animation same as lvicAniCmdStart
            Me.Animate2.ResumeAnimation
        Case lvicAniCmdStart
            Me.Animate2.StartAnimation
        Case lvicAniCmdGetMinDuration               ' if animating/paused, returns minimal duration set during AnimationInitialize event
            Animate = Me.Animate2.DefaultMinimumDuration
        Case lvicAniCmdGetFrameIndex                ' if animating/paused, returns current frame index
            Animate = Me.Animate2.CurrentFrame
        Case lvicAniCmdGetState                     ' if animating/paused, returns current state: lvicAniCmdStart, ani_ActionPause, lvicAniCmdStop
            Animate = Me.Animate2.AnimationState
        Case lvicAniCmdGetLoopCount                 ' if animating/paused, returns loop count (0=infinite)
            Animate = Me.Animate2.LoopCount
        Case lvicAniCmdSetLoopCount                 ' if animating/paused, sets loop count (0=infinite)
            Me.Animate2.LoopCount = DurationLoopCount
            PropertyChanged "Picture"
        Case lvicAniCmdSetMinDuration
            Me.Animate2.DefaultMinimumDuration = DurationLoopCount
            PropertyChanged "Picture"
        Case lvicAniCmdGetMaxDuration
            Animate = Me.Animate2.DefaultMaximumDuration
        Case lvicAniCmdSetMaxDuration
            Me.Animate2.DefaultMaximumDuration = DurationLoopCount
            PropertyChanged "Picture"
    End Select
    
End Function

'//// Determines if a multi-image format begins animation as soon as loaded
'   IMPORTANT: All instances sharing the Image property are affected
Public Property Get AnimateOnLoad() As Boolean
Attribute AnimateOnLoad.VB_Description = "Returns/sets whether image frames will animate when control is loaded"
    AnimateOnLoad = (m_Flags And attr_AutoAnimate)
End Property
Public Property Let AnimateOnLoad(newValue As Boolean)
    If newValue <> Me.AnimateOnLoad Then
        m_Flags = m_Flags Xor attr_AutoAnimate
        If newValue Then Me.Animate lvicAniCmdStart
        PropertyChanged "AnimateOnLoad"
    End If
End Property

'//// Determines image scaling and may affect control size based on AutoSize settings
Public Property Get Aspect() As ScalingRatioEnum
Attribute Aspect.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control.\n"
    Aspect = (m_Flags And attr_StretchMask)
End Property
Public Property Let Aspect(newValue As ScalingRatioEnum)
    If Not (newValue < lvicActualSize Or newValue > lvicFixedSizeStretched) Then
        If Not newValue = Me.Aspect Then
            If newValue < lvicFixedSize Then
                m_Flags = (m_Flags And Not attr_StretchMask) Or newValue
                m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest Or render_DoResize
                If Me.AutoSize Then
                    Call sptGetScaledSizes(Me.SetRedraw)
                Else
                    If Me.SetRedraw Then UserControl.Refresh
                End If
            ElseIf Me.Border Then
                Call SetFixedSizeAspect(UserControl.ScaleWidth - 2&, UserControl.ScaleHeight - 2&, (newValue = lvicFixedSizeStretched))
            Else
                Call SetFixedSizeAspect(UserControl.ScaleWidth, UserControl.ScaleHeight, (newValue = lvicFixedSizeStretched))
            End If
            PropertyChanged "Aspect"
        End If
    End If
End Property
    
'//// Determines whether the control will size itself to the boundaries of the rendered image
'   When image is stretched, no Auto-sizing will be performed
Public Property Get AutoSize() As AutoSizeEnum
Attribute AutoSize.VB_Description = "Returns/sets whether control will be sized to size/angle of rendered image"
    AutoSize = (m_CntFlags And cnt_AutoSizeMask)
End Property
Public Property Let AutoSize(newValue As AutoSizeEnum)
    If Not (newValue < lvicNoAutoSize Or newValue > lvicMultiAngle) Then
        If Not newValue = Me.AutoSize Then
            m_CntFlags = (m_CntFlags And Not cnt_AutoSizeMask) Or newValue
            m_RenderFlags = m_RenderFlags Or render_DoResize Or render_RedoAutoSize Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest
            Call sptGetScaledSizes(Me.SetRedraw) ' hit test area recalculated in that routine
            PropertyChanged "AutoSize"
        End If
    End If
End Property

'//// Determines the backcolor used if BackStyleOpaque is True
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the backcolor used for the control if BackStyleOpaque is True"
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(newValue As OLE_COLOR)
    If Not newValue = UserControl.BackColor Then
        UserControl.BackColor = newValue
        PropertyChanged "BackColor"
        If Me.BackStyleOpaque Then
            If Me.SetRedraw Then UserControl.Refresh
        End If
    End If
End Property

'//// Determines if solid backcolor will be rendered in the control behind the image
Public Property Get BackStyleOpaque() As Boolean
Attribute BackStyleOpaque.VB_Description = "Returns/sets whether the control is filled with color or transparent"
    BackStyleOpaque = (m_CntFlags And cnt_Opaque)
End Property
Public Property Let BackStyleOpaque(newValue As Boolean)
    If Not newValue = Me.BackStyleOpaque Then
        m_CntFlags = m_CntFlags Xor cnt_Opaque
        If Me.SetRedraw Then UserControl.Refresh
        PropertyChanged "BackStyleOpaque"
    End If
End Property

'//// Determines which color is blended into the rendered image. See BlendPct
Public Property Get BlendColor() As OLE_COLOR
Attribute BlendColor.VB_Description = "Returns/sets the color blended into the image if BlendPct is non-zero"
    BlendColor = m_Effects.BlendColor
End Property
Public Property Let BlendColor(newValue As OLE_COLOR)
    m_Effects.BlendColor = newValue
End Property

'//// Determines how much of BlendColor is blended into the rendered image
Public Property Get BlendPct() As Long
Attribute BlendPct.VB_Description = "Returns/sets the amount of BlendColor blended into the image"
    BlendPct = m_Effects.BlendPct
End Property
Public Property Let BlendPct(newValue As Long)
    m_Effects.BlendPct = newValue
End Property

'//// Sets/returns the background image assigned to the control
Public Property Get BkgImage() As GDIpImage
Attribute BkgImage.VB_Description = "Sets/returns the background image for the control"
    If m_BkgImage Is Nothing Then Set m_BkgImage = New GDIpImage
    Set BkgImage = m_BkgImage
End Property
Public Property Let BkgImage(newImage As GDIpImage)
    Set Me.BkgImage = newImage
End Property
Public Property Set BkgImage(newImage As GDIpImage)
    If Not newImage Is m_BkgImage Then
        Set m_BkgImage = newImage
        PropertyChanged "BkgImage"
        If Me.SetRedraw Then UserControl.Refresh
    End If
End Property

'//// Sets/returns whether background image will be stretched or clipped (default)
' The background image can be segmented at runtime by accessing its GDIpImage class directly
' The background image cannot be animated thru this control, but can have its frames changed at
'   runtime by accessing its GDIpImage class directly
Public Property Get BkgImageStretch() As Boolean
Attribute BkgImageStretch.VB_Description = "Sets/returns whether background image is stretched or clipped"
    BkgImageStretch = CBool(m_CntFlags And cnt_BkgStretch)
End Property
Public Property Let BkgImageStretch(newValue As Boolean)
    If Not Me.BkgImageStretch = newValue Then
        m_CntFlags = m_CntFlags Xor cnt_BkgStretch
        PropertyChanged "BkgImageStretch"
        If Not m_BkgImage Is Nothing Then
            If Me.SetRedraw Then UserControl.Refresh
        End If
    End If
End Property

'//// Adds a one pixel border to the control, resizing control as needed. See BorderColor
Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Returns/sets whether a border will be drawn around the control"
    Border = (m_CntFlags And cnt_Border)
End Property
Public Property Let Border(newValue As Boolean)
    If newValue <> Me.Border Then
        m_CntFlags = m_CntFlags Xor cnt_Border
        If newValue = False Then m_CntFlags = m_CntFlags And Not cnt_BorderMask
        m_RenderFlags = (m_RenderFlags And Not render_Shown)
        If newValue Then
            UserControl.Size ScaleX(UserControl.ScaleWidth + 2&, vbPixels, vbTwips), ScaleY(UserControl.ScaleHeight + 2&, vbPixels, vbTwips)
        Else
            UserControl.Size ScaleX(UserControl.ScaleWidth - 2&, vbPixels, vbTwips), ScaleY(UserControl.ScaleHeight - 2&, vbPixels, vbTwips)
        End If
        m_RenderFlags = m_RenderFlags Xor render_Shown Or render_DoHitTest
        If Me.SetRedraw = True Then UserControl.Refresh
        PropertyChanged "Border"
    End If
End Property

'//// Draws border with either sharp or rounded corners
' Note this property and the Border property are mutually inclusive
Public Property Get BorderShape() As BorderShapeEnum
Attribute BorderShape.VB_Description = "Returns/sets whether border has sharp or rounded corners"
    If Me.Border = True Then
        BorderShape = (m_CntFlags And cnt_BorderMask) \ cnt_RoundBorder + 1&
    End If
End Property
Public Property Let BorderShape(newValue As BorderShapeEnum)
    ' note about the lvicUserDefinedBorder setting
    ' That setting will ensure control is sized to allow a 1 pixel user-defined/drawn border
    '   It will also offset the image 1 pixel in both directions to prevent the border from
    '   drawing over much of the image edges.
    ' To draw your own border, recommend these steps be followed
    ' 1. WantPrePostEvents = True
    ' 2. During the control's PrePaint event, select a clipping region into the passed DC &
    '       either let the control paint the image, or paint it yourself then set Cancel parameter to True
    ' 3. During the control's PostPaint event, draw your custom border after removing any clipping region
    '   To ensure no clipping region: SelectClipRgn [hDC], 0&
    If Not newValue = Me.BorderShape Then
        If newValue >= lvicNoBorder And newValue <= lvicUserDefinedBorder Then
            If newValue = lvicNoBorder Then
                Me.Border = False
            Else
                Dim bRedraw As Boolean
                bRedraw = Me.SetRedraw
                Me.SetRedraw = False
                Me.Border = True
                m_CntFlags = (m_CntFlags And Not cnt_BorderMask) Or (newValue - 1&) * cnt_RoundBorder
                Me.SetRedraw = bRedraw
                PropertyChanged "BorderShape"
            End If
        End If
    End If
End Property

'//// Color used for the control's border when Border property is True
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the border color"
    BorderColor = UserControl.ForeColor
End Property
Public Property Let BorderColor(newValue As OLE_COLOR)
    If newValue <> UserControl.ForeColor Then
        UserControl.ForeColor = newValue
        If Me.Border Then
            If Me.SetRedraw = True Then UserControl.Refresh
        End If
        PropertyChanged "BorderColor"
    End If
End Property

'//// Determines which effect will be applied.
' Effects require GDI+ v1.1 (Vista and above possibly, Office 2003+ I believe)
' Either way, effects can be set during runtime or via the property page if correct GDI+ version exists
Public Property Get Effect() As EffectsEnum
Attribute Effect.VB_Description = "Returns/sets an index to a created effect to be used for rendering"
    Effect = (m_Attributes And &HF&)
End Property
Public Property Let Effect(newValue As EffectsEnum)
    If newValue <> Me.Effect Then
        m_Attributes = (m_Attributes And &HFFFFFFF0) Or newValue
        If g_TokenClass.Version > 1! Then
            If Me.SetRedraw Then UserControl.Refresh
        End If
    End If
End Property

'//// Exposes the GDIpEffects class so you can create/destroy effects during runtime
Public Property Get Effects() As GDIpEffects
Attribute Effects.VB_Description = "Returns the GDIpEffects class bound to the control"
    Set Effects = m_Effects
End Property

'//// Determines whether control will respond to system mouse events
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events.\n"
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(newValue As Boolean)
    If newValue <> UserControl.Enabled Then
        UserControl.Enabled = newValue
        m_RenderFlags = m_RenderFlags Or render_DoHitTest
        PropertyChanged "Enabled"
    End If
End Property

'//// Determines whether rendered snapshot of image will be maintained
'   Requires additional resources: a 32bpp bitmap will be used
'   Not recommended for small images less than or equal to 256x256
Public Property Get FastRedraw() As Boolean
Attribute FastRedraw.VB_Description = "Returns/sets whether a cached copy of rendered image is maintained"
    FastRedraw = (m_RenderFlags And render_FastRedraw)
End Property
Public Property Let FastRedraw(newValue As Boolean)
    If Not newValue = Me.FastRedraw Then
        If newValue Then        ' call routine that creates the FastRedraw bitmap
            m_RenderFlags = m_RenderFlags Or render_FastRedraw
            Call sptCreateFastRedrawImage
        Else                    ' clean up bitmap & destroy class unless being used
            m_DC.ResizeBitmap 0&, 0&, 0&, False
            If m_DC.hBitmap(True) = 0& Then Set m_DC = Nothing
            m_RenderFlags = m_RenderFlags Xor render_FastRedraw
        End If
        PropertyChanged "FastRedraw"
    End If
End Property

'//// Color used to start gradient & ends at BackColor
' can pass standard RGB or VB system colors
Public Property Get GradientForeColor() As OLE_COLOR
Attribute GradientForeColor.VB_Description = "Returns/Sets color that starts gradient. Gradient ends at Backcolor"
    GradientForeColor = UserControl.FillColor
End Property
Public Property Let GradientForeColor(newValue As OLE_COLOR)
    If newValue <> UserControl.FillColor Then
        UserControl.FillColor = newValue
        If Me.BackStyleOpaque Then
            If Me.SetRedraw Then UserControl.Refresh
        End If
        PropertyChanged "GradientForeColor"
    End If
End Property
'//// Gradient style. To reverse gradients, swap GradientForeColor & BackColor property values
'   Horizontal: From left (GradientForeColor) to right (BackColor)
'   Vertical: From top (GradientForeColor) to bottom (BackColor)
'   DiagonalDown: From top/left (GradientForeColor) to bottom/right (BackColor)
'   DiagonalUp: From bottom/left (GradientForeColor) to top/right (BackColor)
Public Property Get GradientStyle() As GradientStyleEnum
Attribute GradientStyle.VB_Description = "Returns/Sets gradient direction. To reverse direction swap GradientForeColor and Backcolor"
    GradientStyle = (m_CntFlags And cnt_GradientMask) \ cnt_GradHorizontal
End Property
Public Property Let GradientStyle(newValue As GradientStyleEnum)
    If newValue <> Me.GradientStyle Then
        If newValue >= lvicNoGradient And newValue <= lvicGradientDiagonalUp Then
            m_CntFlags = (m_CntFlags And Not cnt_GradientMask) Or newValue * cnt_GradHorizontal
            If Me.BackStyleOpaque Then
                If Me.SetRedraw Then UserControl.Refresh
            End If
            PropertyChanged "GradientStyle"
        End If
    End If
End Property

'//// Determines whether image is rendered with grayscale or not
Public Property Get GrayScale() As GrayScaleRatioEnum
Attribute GrayScale.VB_Description = "Returns/sets the gray scale formula used to render the image"
    GrayScale = m_Effects.GrayScale
End Property
Public Property Let GrayScale(newValue As GrayScaleRatioEnum)
    m_Effects.GrayScale = newValue
End Property

'//// Determines what part(s) of the control will react to system mouse events
Public Property Get HitTest() As HitTestModeEnum
Attribute HitTest.VB_Description = "Returns/sets the area of the control that respoinds to mouse events"
    If m_HitRegion Then
        HitTest = lvicRunTimeRegion
    Else
        HitTest = (m_Flags And attr_HitTestMask) \ attr_HitTestShift
    End If
End Property
Public Property Let HitTest(ByVal newValue As HitTestModeEnum)
    Dim bOK As Boolean
    If (newValue < lvicPerimeter Or newValue > lvicTrimmedImage) Then
        ' During runtime you can pass a valid region for a hit test.
        ' If you do, you must not destroy the region, this control will own & manage it
        ' Note: The region is not modified at all and is valid until a new HitTest property value is set
        If GetRegionData(newValue, 0&, ByVal 0&) Then
            m_Flags = (m_Flags And Not attr_HitTestMask) Or lvicRunTimeRegion * attr_HitTestShift
            bOK = True
        End If
    
    ElseIf newValue <> Me.HitTest Then
        m_Flags = (m_Flags And Not attr_HitTestMask) Or newValue * attr_HitTestShift
        newValue = 0&
        bOK = True
    End If
    If bOK Then
        m_RenderFlags = m_RenderFlags Or render_DoHitTest
        Call sptCreateHitTestPoints(newValue)
        PropertyChanged "HitTest"
    End If
    
End Property

'//// Determines which image in a multi-image format will be displayed
'   Multi-image formats include TIFF, GIF, ICO, CUR, ANI. Test ImageCount property
'   IMPORTANT: All instances sharing the Image property are affected
Public Property Get ImageCount() As Long
Attribute ImageCount.VB_Description = "Returns the number of frames contained in the loaded image"
    ImageCount = m_Image.ImageCount
End Property
Public Property Let ImageCount(newValue As Long)
    ' dummy LET property to allow count displayed on property sheet during design time
End Property

'//// Win7 has animated cursors that can contain multiple groups of cursors
' This property returns/sets the current group of animated cursors
'   IMPORTANT: All instances sharing the Image property are affected
Public Property Get ImageGroup() As Long
Attribute ImageGroup.VB_Description = "Animated cursors may contain more than one group of images. Returns/sets the current group"
    ImageGroup = m_Image.ImageGroup
End Property
Public Property Let ImageGroup(newValue As Long)
    m_Image.ImageGroup = newValue
    PropertyChanged "ImageGroup"
End Property
'//// Returns the number of animated cursor groups
Public Property Get ImageGroupCount() As Long
Attribute ImageGroupCount.VB_Description = "Animated cursors may contain more than one group of images. Returns number of groups"
    ImageGroupCount = m_Image.ImageGroups
End Property

'//// Determines which image in a multi-image format will be displayed'
'   Multi-image formats include TIFF, GIF, ICO, CUR, ANI, Segmented images
'   Test ImageCount property
'   IMPORTANT: All instances sharing the Image property are affected
Public Property Get ImageIndex() As Long
Attribute ImageIndex.VB_Description = "Returns/sets the frame index to be displayed"
    ImageIndex = m_Image.ImageIndex
End Property
Public Property Let ImageIndex(newValue As Long)
    m_Image.ImageIndex = newValue       ' triggers an event from the image class
    PropertyChanged "ImageIndex"        ' rendering flags are set there
End Property

'//// Determines rendering quality when scaled or rotated
' See AICGlobals RenderInterpolation enumeration for more details
Public Property Get Interpolation() As RenderInterpolation
Attribute Interpolation.VB_Description = "Returns/Sets rendering quality. Nearest-Neighbor is lowest quality"
    Interpolation = (m_Attributes And &HF000000) \ &H1000000
End Property
Public Property Let Interpolation(newValue As RenderInterpolation)
    If newValue >= lvicAutoInterpolate And newValue <= lvicHighQualityBicubic Then
        If newValue <> Me.Interpolation Then
            m_Attributes = (m_Attributes And &HFFFFFF) Or newValue * &H1000000
            Me.Refresh
            PropertyChanged "Interpolation"
        End If
    End If
End Property

'//// Returns whether or not rendering is inverted (color negative)
Public Property Get Inverted() As Boolean
Attribute Inverted.VB_Description = "Returns/sets whether or not rendering is inverted (color negative)"
    Inverted = m_Effects.Invert
End Property
Public Property Let Inverted(newValue As Boolean)
    m_Effects.Invert = newValue
End Property

'//// Returns whether or not an image is segmented. See SegmentImage
Public Property Get IsSegmented() As Boolean
Attribute IsSegmented.VB_Description = "Returns whether the image has been segmented into multiple images"
    IsSegmented = m_Image.Segmented
End Property

'//// Determines additional lightness rendered with the image
'   Negative values remove lightness while positive values add lightness
Public Property Get LightnessPct() As Long
Attribute LightnessPct.VB_Description = "Returns/sets the percentage of lightness added to control. Negative values remove lightness"
    LightnessPct = m_Effects.LightnessPct
End Property
Public Property Let LightnessPct(ByVal newValue As Long)
    m_Effects.LightnessPct = newValue
End Property

'//// Mirrors the image
Public Property Get Mirror() As MirroredEnum
Attribute Mirror.VB_Description = "Returns/sets whether image is rendered mirrored"
     Mirror = (m_Attributes And &HF0000) \ &H10000
End Property
Public Property Let Mirror(newValue As MirroredEnum)
    If Not (newValue < lvicMirrorNone Or newValue > lvicMirrorBoth) Then
        If Me.Mirror <> newValue Then
            m_Attributes = (m_Attributes And &HFF00FFFF) Or newValue * &H10000
            m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoHitTest
            If Me.SetRedraw = True Then UserControl.Refresh
            PropertyChanged "Mirror"
        End If
    End If
End Property

'//// Returns/Sets the cursor for the control. Enabled must be true to see icon
Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon"
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Let MouseIcon(ByVal newValue As StdPicture)
    Set MouseIcon = newValue
End Property
Public Property Set MouseIcon(ByVal newValue As StdPicture)
    Set UserControl.MouseIcon = newValue
    PropertyChanged "MouseIcon"
End Property

'//// Returns/Sets the cursor style for the control. Enabled must be true to see cursor
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object"
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal newValue As MousePointerConstants)
    If Not newValue = UserControl.MousePointer Then
        UserControl.MousePointer = newValue
        PropertyChanged "MousePointer"
    End If
End Property

'//// Enables positioning adjustments of rendered image within the control
Public Property Get XOffset() As Long
Attribute XOffset.VB_Description = "Returns/sets left offset adjustment the image will be rendered from"
    XOffset = m_Offsets.X
End Property
Public Property Let XOffset(newValue As Long)
    If Not newValue = m_Offsets.X Then
        If newValue > &HFFF& Then
            m_Offsets.X = &HFFF&
        ElseIf newValue < &HFFFFF001 Then
            m_Offsets.X = &HFFFFF001
        Else
            m_Offsets.X = newValue
        End If
        m_RenderFlags = m_RenderFlags Or render_DoHitTest
        If Me.SetRedraw = True Then UserControl.Refresh
        PropertyChanged "XOffset"
    End If
End Property
Public Property Get YOffset() As Long
Attribute YOffset.VB_Description = "Returns/sets top offset adjustment the image will be rendered from"
    YOffset = m_Offsets.Y
End Property
Public Property Let YOffset(newValue As Long)
    If Not newValue = m_Offsets.Y Then
        If newValue > &HFFF& Then
            m_Offsets.Y = &HFFF&
        ElseIf newValue < &HFFFFF001 Then
            m_Offsets.Y = &HFFFFF001
        Else
            m_Offsets.Y = newValue
        End If
        m_RenderFlags = m_RenderFlags Or render_DoHitTest
        If Me.SetRedraw = True Then UserControl.Refresh
        PropertyChanged "YOffset"
    End If
End Property

'//// Initiates an OLE Drag event
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source.\n"
    m_DragDrop.Originator = True            ' identify we are the draggers
    UserControl.OLEDrag                     ' start the drag & reset the drop mode
    m_DragDrop.AutoDragPts.X = 0&: m_DragDrop.AutoDragPts.Y = 0&
    m_DragDrop.Originator = False
    If (m_CntFlags And &HF000000) Then
        If (m_Flags And attr_MouseEvents) Then
            RaiseEvent MouseUp((m_CntFlags And &HF000000) \ cnt_LastButtonShift, 0, -1!, -1!)
        End If
        m_CntFlags = (m_CntFlags And &HFFFFFF) And Not cnt_MouseDown
    End If
End Sub

'//// Determines whether image dragging is manual or automatic
Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target.\n"
    OLEDragMode = (m_Flags And attr_OLEDragModeShift) \ attr_OLEDragModeShift
End Property
Public Property Let OLEDragMode(newValue As OLEDragConstants)
    If Not (newValue < vbOLEDragManual Or newValue > vbOLEDragAutomatic) Then
        If newValue <> Me.OLEDragMode Then
            m_Flags = (m_Flags And Not attr_OLEDragModeShift) Or newValue * attr_OLEDragModeShift
            PropertyChanged "OLEDragMode"
        End If
    End If
End Property

'//// Determines whether image dropping is manual or automatic or not used
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target.\n"
    OLEDropMode = (m_Flags And attr_OLEDropMask) \ attr_OLEDropShift
End Property
Public Property Let OLEDropMode(newValue As OLEDropConstants)
    If Not (newValue < vbOLEDropNone Or newValue > vbOLEDropAutomatic) Then
        If newValue <> Me.OLEDropMode Then
            If newValue = vbOLEDropNone Then
                UserControl.OLEDropMode = newValue
            Else    ' can't set usercontrol's drop mode to Automatic; so we use manual and do Automatic via code
                UserControl.OLEDropMode = vbOLEDropManual
            End If
            m_Flags = (m_Flags And Not attr_OLEDropMask) Or newValue * attr_OLEDropShift
            PropertyChanged "OLEDropMode"
        End If
    End If
End Property

'//// Refreshes the control and resets SetRedraw property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    m_RenderFlags = (m_RenderFlags And Not render_NoRedraw)
    Call sptGetScaledSizes(False)
    UserControl.Refresh
End Sub

'//// Saves the control, as drawn, to a GDIpImage class as a PNG format
Public Function SaveControlAsDrawnToGDIpImage(ByVal IncludeContainerBkg As Boolean) As GDIpImage
Attribute SaveControlAsDrawnToGDIpImage.VB_Description = "Saves control appearance to a PNG format"
' The optional parameter if True will include control's container background & border if used
' Control must be visible to capture its contents

    If (m_RenderFlags And render_Shown) = 0& Then Exit Function ' control not yet shown or Hidden event called

    ' saves the control, same size, as drawn
    Dim newHandle As Long, SS As SAVESTRUCT
    
    m_DragDrop.AutoDragPts.X = 0&
    m_DragDrop.AutoDragPts.Y = 0&

    If sptRemoteRender(True, newHandle, IncludeContainerBkg) Then
        
        Set g_NewImageData = New cGDIpMultiImage
        g_NewImageData.CacheSourceInfo Empty, newHandle, lvicPicTypeBitmap, True, False
        Set SaveControlAsDrawnToGDIpImage = New GDIpImage
        Set g_NewImageData = Nothing
        
        On Error Resume Next
        SS.ColorDepth = lvicDefaultReduction
        modCommon.SaveImage SaveControlAsDrawnToGDIpImage, SaveControlAsDrawnToGDIpImage, lvicSaveAsPNG, SS
        
    End If
    
End Function

'//// Saves the image, as drawn, to a GDIpImage class as a PNG format. Image will not be clipped
Public Function SaveImageAsDrawnToGDIpImage() As GDIpImage
Attribute SaveImageAsDrawnToGDIpImage.VB_Description = "Saves control's image as drawn to a PNG format"
' Control need not be visible
    Dim newHandle As Long, SS As SAVESTRUCT
    Dim Cx As Long, Cy As Long, hGraphics As Long
    Dim Width As Long, Height As Long
    
    If m_Image.Handle = 0& Then Exit Function
    
    modCommon.GetScaledCanvasSize m_Size.Width, m_Size.Height, Cx, Cy, m_Angle
    If GdipCreateBitmapFromScan0(Cx, Cy, 0&, lvicColor32bppAlpha, ByVal 0&, newHandle) Then Exit Function
    If GdipGetImageGraphicsContext(newHandle, hGraphics) Then
        GdipDisposeImage newHandle: Exit Function
    End If
    Width = m_Image.Width: Height = m_Image.Height
    If (Me.Mirror And lvicMirrorHorizontal) Then Width = -Width
    If (Me.Mirror And lvicMirrorVertical) Then Height = -Height
    m_Image.Render 0&, (Cx - m_Size.Width) \ 2, (Cy - m_Size.Height) \ 2, m_Size.Width, m_Size.Height, , , Width, Height, m_Angle, m_Effects.AttributesHandle, hGraphics, m_Effects.EffectsHandle(Me.Effect), Me.Interpolation
    GdipDeleteGraphics hGraphics
    
    Set g_NewImageData = New cGDIpMultiImage
    g_NewImageData.CacheSourceInfo Empty, newHandle, lvicPicTypeBitmap, True, False
    Set SaveImageAsDrawnToGDIpImage = New GDIpImage
    Set g_NewImageData = Nothing
    
    On Error Resume Next
    SS.ColorDepth = lvicDefaultReduction
    modCommon.SaveImage SaveImageAsDrawnToGDIpImage, SaveImageAsDrawnToGDIpImage, lvicSaveAsPNG, SS
        
End Function

'//// Renders the image, as drawn, to any DC. The image will not be clipped
Public Function PaintImageAsDrawnToHDC(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Boolean
Attribute PaintImageAsDrawnToHDC.VB_Description = "Renders control's image to a device context"
    If hDC = 0& Then Exit Function
    If m_Image.Handle = 0& Then Exit Function
    
    Dim Cx As Long, Cy As Long
    Dim Width As Long, Height As Long
    
    modCommon.GetScaledCanvasSize m_Size.Width, m_Size.Height, Cx, Cy, m_Angle
    Width = m_Image.Width: Height = m_Image.Height
    If (Me.Mirror And lvicMirrorHorizontal) Then Width = -Width
    If (Me.Mirror And lvicMirrorVertical) Then Height = -Height
    PaintImageAsDrawnToHDC = m_Image.Render(hDC, X + (Cx - m_Size.Width) \ 2, Y + (Cy - m_Size.Height) \ 2, m_Size.Width, m_Size.Height, , , Width, Height, m_Angle, m_Effects.AttributesHandle, , m_Effects.EffectsHandle(Me.Effect), Me.Interpolation)

End Function

'//// Renders the control, as drawn, to any DC. Control must be visible and viewable
Public Function PaintControlAsDrawnToHDC(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal IncludeContainerBkg As Boolean) As Boolean
Attribute PaintControlAsDrawnToHDC.VB_Description = "Renders control to a device context"
    If (m_RenderFlags And render_Shown) = 0& Then Exit Function ' control not yet shown or Hidden event called
    If hDC = 0& Then Exit Function
    ' paints the control as drawn to another DC
    m_DragDrop.AutoDragPts.X = X
    m_DragDrop.AutoDragPts.Y = Y
    PaintControlAsDrawnToHDC = sptRemoteRender(False, hDC, IncludeContainerBkg)

End Function

'//// Returns the controls inside dimensions in pixels, less any border pixels
Public Property Get ScaleLeft() As Long
Attribute ScaleLeft.VB_Description = "Left edge of the control, dependent on border settings"
    ScaleLeft = Abs(Me.Border)
End Property
Public Property Get ScaleTop() As Long
Attribute ScaleTop.VB_Description = "Top edge of the control, dependent on border settings"
    ScaleTop = Abs(Me.Border)
End Property
Public Property Get ScaleHeight() As Long
Attribute ScaleHeight.VB_Description = "Actual control height in pixels less border width"
    ScaleHeight = UserControl.ScaleHeight + Me.Border * 2&
End Property
Public Property Get ScaleWidth() As Long
Attribute ScaleWidth.VB_Description = "Actual control width in pixels less border width"
    ScaleWidth = UserControl.ScaleWidth + Me.Border * 2&
End Property

'//// Allows an image to be interpretted as several images. The actual image is not modified
'  If your image is like a film strip where each "frame" is treated as an individual image,
'   this routine allows you to refernce your image as several images. See GDIpImage class for more
'   IMPORTANT: All instances sharing the Image property are affected
Public Sub SegmentImage(ByVal Rows As Long, ByVal Columns As Long, _
                        Optional ByVal TilesUsed As Long, Optional Duration As Long)
Attribute SegmentImage.VB_Description = "Creates a tiled image or removes the tiles"
    m_Image.SegmentImage Columns, Rows, TilesUsed, Duration
End Sub

'//// Determines the angle the image will be rendered at
'   Negative angles result in counter-clockwise rotation & positive angles are clockwise
Public Property Get Rotation() As Single
Attribute Rotation.VB_Description = "Returns/sets the angle of rendering. Negative is counter-clockwise rotation"
        Rotation = m_Angle
End Property
Public Property Let Rotation(ByVal newValue As Single)
    If Not Me.Rotation = newValue Then
        m_Angle = newValue
        m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest
        If Me.AutoSize = lvicSingleAngle Then m_RenderFlags = m_RenderFlags
        Call sptGetScaledSizes(Me.SetRedraw)
        PropertyChanged "Rotation"
    End If
End Property

'/// sets optional clipping on the rendered image. Border usage is exempt from clipping
Public Sub SetClipRect(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Refresh As Boolean = True)
Attribute SetClipRect.VB_Description = "Method clips the control's graphic output"
    ' pass width and/or height as zero to remove clipping rectangle
    If (Width Or Height) = 0 Then
        SetRect m_ClipRect, 0&, 0&, 0&, 0&
        If Refresh Then UserControl.Refresh
    ElseIf (Width > 0& And Height > 0&) Then
        SetRect m_ClipRect, X, Y, Width + X, Height + X
        If Refresh Then UserControl.Refresh
    End If

End Sub

'/// Enables assigning fixed aspect ration during runtime
Public Sub SetFixedSizeAspect(ByVal Width As Long, ByVal Height As Long, ByVal Stretched As Boolean)
Attribute SetFixedSizeAspect.VB_Description = "Runtime setting for aspect property of FixedSize or FixedSizeStretched"

    If (Width > 0& And Height > 0&) Then
        m_Flags = (m_Flags And Not attr_StretchMask) Or (Abs(Stretched) + lvicFixedSize)
        m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest Or render_DoResize
        m_Size.FixedCx = Width
        m_Size.FixedCy = Height
        If Me.AutoSize Then
            Call sptGetScaledSizes(Me.SetRedraw)
        Else
            If Me.SetRedraw Then UserControl.Refresh
        End If
    End If
    
End Sub

'//// sets the gradient settings in one call
Public Sub SetGradientOptions(ByVal ForeColor As Long, ByVal BackColor As Long, _
                        ByVal Style As GradientStyleEnum, Optional ByVal MakeOpaqueBackStyle As Boolean = True)
Attribute SetGradientOptions.VB_Description = "Sets gradient options in one call"
    On Error Resume Next
    If Style >= lvicNoGradient And Style <= lvicGradientDiagonalUp Then
        m_CntFlags = (m_CntFlags And Not cnt_GradientMask) Or Style * cnt_GradHorizontal
    End If
    UserControl.FillColor = ForeColor
    UserControl.BackColor = BackColor
    If MakeOpaqueBackStyle = True Then
        m_CntFlags = m_CntFlags Or cnt_Opaque
    Else
        m_CntFlags = m_CntFlags And Not cnt_Opaque
    End If
    If Me.SetRedraw Then UserControl.Refresh
End Sub

'//// Sets both the XOffset & YOffset properties
Public Sub SetOffsets(ByVal newXoffset As Long, ByVal newYoffset As Long)
Attribute SetOffsets.VB_Description = "Sets optional X,Y rendering offsets"

    Dim bResetRedraw As Boolean
    bResetRedraw = Me.SetRedraw
    m_RenderFlags = m_RenderFlags Or render_NoRedraw
    Me.XOffset = newXoffset: Me.YOffset = newYoffset
    m_RenderFlags = m_RenderFlags Or render_DoHitTest
    If bResetRedraw Then m_RenderFlags = m_RenderFlags Xor render_NoRedraw
    If Me.SetRedraw = True Then UserControl.Refresh

End Sub

'//// Determines whether or not changes to properties are immediately applied
'   Setting this to False will pause all updates from being displayed until
'   the property is set to True, Refresh is called
Public Property Get SetRedraw() As Boolean
Attribute SetRedraw.VB_Description = "Returns/sets whether control will apply changes"
Attribute SetRedraw.VB_MemberFlags = "400"
    ' This property should be hidden from the property sheet, it is a runtime only setting
    ' To verify hidden from user
    '   1. IDE Menu: Tools | Procedure Attributes
    '   2. Find SetRedraw in the "Name" combobox
    '   3. Click the "Advanced" button
    '   4. Check the box: "Don't show in property browser"
    '   5. Click the "Apply" button
    SetRedraw = (m_RenderFlags And render_NoRedraw) = 0&
End Property
Public Property Let SetRedraw(ByVal newValue As Boolean)
    ' When setting multiple properties or loading an image and you don't want to update the
    ' image right away, setting this property to False will prevent any control updates until
    ' this property is set back to True or Refresh is called.
    ' This option gives you the opportunity to load an image, check its size/bit depth,
    '   reload a different ImageIndex if needed, resize control, set attributes and anything else
    '   without the user seeing any of this. Useful.
    If newValue Then
        Me.Refresh
    Else
        m_RenderFlags = m_RenderFlags Or render_NoRedraw
    End If
End Property

'///// Databound image property. Is NOT compatible with the DAO data control (Data1)
' Compatible with the ADO data control (Adodc1) and ADO recordsets
' To verify it is databound and hidden
'   1. IDE Menu: Tools | Procedure Attributes
'   2. Find "SinkImageData" in the "Name" combobox
'   3. Click the "Advanced" button
'   4. Check the box: "Property is data bound"
'   5. Check the box: "This property binds to DataField"
'   6. Check the box: "Property will call CanPropertyChange before changing"
'   7. Check the box: Hide this member
'   8. Click the "Apply" button
Public Property Get SinkImageData() As Variant
Attribute SinkImageData.VB_Description = "Databound picture property"
Attribute SinkImageData.VB_MemberFlags = "6c"
    Dim bData() As Byte, tImg As GDIpImage
    
'    Debug.Print "updating databbound control"
    
    Set tImg = m_Image
    RaiseEvent UpdateDataboundImage(tImg)
    If tImg Is Nothing Then
        SinkImageData = Null
    ElseIf tImg.ExtractImageData(bData()) = True Then
        modCommon.MoveArrayToVariant SinkImageData, bData(), True
    Else
        SinkImageData = Null
    End If
End Property
Public Property Let SinkImageData(newValue As Variant)
    m_CntFlags = m_CntFlags Or cnt_DatabaseImage
    Set Me.Picture = modCommon.LoadImage(newValue, , , True)
    m_CntFlags = (m_CntFlags And Not cnt_DatabaseImage)
End Property

'//// Determines amount of transparency/translucency for the rendered image
Public Property Get TransparencyPct() As Long
Attribute TransparencyPct.VB_Description = "Returns/sets level of transparency the image is rendered with"
    TransparencyPct = m_Effects.GlobalTransparencyPct
End Property
Public Property Let TransparencyPct(ByVal newValue As Long)
    m_Effects.GlobalTransparencyPct = newValue
End Property

'//// Determines whether a single color will be made transparent throughout the image
'   Setting enables or ignores the TransparentColor property
'   This property is reset each time a new image is selected into the control
Public Property Let TransparentColorMode(newValue As TransparentColorModeEnum)
Attribute TransparentColorMode.VB_Description = "Returns/sets whether TransparentColor property valid and/or where to get that color from"
    If newValue <> Me.TransparentColorMode Then
        ' negative values in range of lvicTransparentTopLeft thru lvicTransparentBottomRight
        ' are used to refresh this property on multi-frame images, from within this control
        ' So, if an AVI is being played and the top left color is suppose to be transparent,
        ' each time a new frame is about to be displayed, this property is called again with
        ' negative values to ensure the current frame's top left color is used.
        If newValue > -lvicUseTransparentColor And newValue <= lvicUseTransparentColor Then
            If newValue = lvicNoTransparentColor Then
                m_Effects.TransparentColorUsed = False
            Else
                If newValue < lvicUseTransparentColor Then
                    Dim lColor As Long
                    If g_TokenClass.Token Then
                        Select Case Abs(newValue)
                        Case lvicTransparentTopLeft: Call GdipBitmapGetPixel(m_Image.Handle, 0&, 0&, lColor)
                        Case lvicTransparentTopRight: Call GdipBitmapGetPixel(m_Image.Handle, m_Image.Width - 1&, 0&, lColor)
                        Case lvicTransparentBottomLeft: Call GdipBitmapGetPixel(m_Image.Handle, 0&, m_Image.Height - 1&, lColor)
                        Case lvicTransparentBottomRight: Call GdipBitmapGetPixel(m_Image.Handle, m_Image.Width - 1&, m_Image.Height - 1&, lColor)
                        End Select
                        lColor = (lColor And &HFF00&) Or (lColor And &HFF0000) \ &H10000 Or (lColor And &HFF&) * &H10000
                        If newValue < lvicNoTransparentColor Then
                            If lColor = m_Effects.TransparentColor Then Exit Property
                            m_CntFlags = m_CntFlags Or cnt_InitLoad
                            m_Effects.TransparentColor = lColor
                            m_CntFlags = m_CntFlags Xor cnt_InitLoad
                        Else
                            m_Effects.TransparentColor = lColor
                        End If
                    End If
                End If
                m_Effects.TransparentColorUsed = True
            End If
            m_Attributes = (m_Attributes And &HFFFF00FF) Or (Abs(newValue) * &H100&)
            PropertyChanged "TransparentColorMode"
        End If
    End If
End Property
Public Property Get TransparentColorMode() As TransparentColorModeEnum
    TransparentColorMode = (m_Attributes And &HF00&) \ &H100&
End Property

'//// Determines the color to be made transparent throughout the entire image
'   Depends on the TransparentColorMode property
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = m_Effects.TransparentColor
End Property
Public Property Let TransparentColor(newValue As OLE_COLOR)
Attribute TransparentColor.VB_Description = "Returns/Sets color to be made transparent throughout the image"
    m_Effects.TransparentColor = newValue
    m_Effects.TransparentColorUsed = True
End Property

'//// Mouse events are forwarded by default, set to False & following won't be forwarded
'     MouseUp, MouseDown, MouseMove, MouseEnter, MouseExit
' MouseExit & MouseEnter use additional resources for tracking events
Public Property Get WantMouseEvents() As Boolean
Attribute WantMouseEvents.VB_Description = "Returns/sets whether mouse events will be sent to the user"
    WantMouseEvents = (m_Flags And attr_MouseEvents)
End Property
Public Property Let WantMouseEvents(newValue As Boolean)
    If newValue <> Me.WantMouseEvents Then
        m_Flags = m_Flags Xor attr_MouseEvents
        m_CntFlags = (m_CntFlags And Not cnt_MouseValidate)
        If newValue = False Then
            If Not m_MouseTracker Is Nothing Then Call m_MouseTracker.ReleaseMouseCapture(False, ObjPtr(Me))
            Set m_MouseTracker = Nothing
        End If
        If m_HitRegion = 0& Then Call sptCreateHitTestPoints(m_HitRegion)
        PropertyChanged "WantMouseEvents"
    End If
End Property

'//// Determines whether or not Pre & Post Render events will be sent
' Property uses additional resources: a 24bpp or 32bpp bitmap as a backbuffer
' Note: If you request 32bpp, you must take into account the alpha channel and
'   the rendered image must always be premultiplied alpha. If using 32bpp,
'   strongly recommend using GDI+ since it is alpha-channel aware
Public Property Get WantPrePostEvents() As Boolean
Attribute WantPrePostEvents.VB_Description = "Retruns/sets whether Pre- and Post-Paint events will be sent to the user"
    WantPrePostEvents = (m_RenderFlags And render_PrePost)
End Property
Public Property Let WantPrePostEvents(newValue As Boolean)
    If Not newValue = Me.WantPrePostEvents Then
        m_RenderFlags = m_RenderFlags Xor render_PrePost
        If newValue Then                                        ' want events; need 24bpp bitmap to assist
            If m_DC Is Nothing Then Set m_DC = New cDeviceContext
            If m_DC.ResizeBitmap(UserControl.ScaleWidth, UserControl.ScaleHeight, 24&, True) = False Then
                m_RenderFlags = (m_RenderFlags And Not render_PrePost) ' if system failed to create bmp, we reset property
            End If
        Else                                        ' remove 24bpp bitmap unless being used
            m_DC.ResizeBitmap 0&, 0&, 0&, True
        End If                                      ' clear class if no bitmaps being used
        If m_DC.hBitmap(False) = 0& And m_DC.hBitmap(True) = 0& Then Set m_DC = Nothing
        If (m_CntFlags And cnt_InitLoad) = 0& Then PropertyChanged "WantPrePostEvents"
    End If
End Property

'//// helper routine that creates an Auto-Redaw image
Private Sub sptCreateFastRedrawImage()

    ' Function creates the FastRedraw image and/or the snapshot image used during SetRedraw=False

    ' Note: Following properties require FastRedraw to be refreshed
    ' Changes in: Control Size and the following properties
    '       AlignCenter, Aspect, AutoSize, Angle,
    '       GrayScale, BlendColor & Pct, Lightness Pct, Mirroring, ImageIndex

    Dim lTrans As Long, X As Long, Y As Long
    
    If (m_RenderFlags And render_FastRedraw) = 0& Then
        m_RenderFlags = (m_RenderFlags And Not render_DoFastRedraw)
        Exit Sub
    ElseIf m_Image.Handle = 0& Then
        Exit Sub
    End If
        
    If Me.SetRedraw = False Then
        ' postpone FastRedraw until SetRedraw=True
        m_RenderFlags = m_RenderFlags Or render_DoFastRedraw
        
    Else
        If m_DC Is Nothing Then Set m_DC = New cDeviceContext
        modCommon.GetScaledCanvasSize m_Size.Width, m_Size.Height, m_Size.FRdrWidth, m_Size.FRdrHeight, m_Angle
        m_DC.ResizeBitmap m_Size.FRdrWidth, m_Size.FRdrHeight, 32&, False
        If m_DC.hBitmap(False) Then
            X = (m_Size.FRdrWidth - m_Size.Width) \ 2
            Y = (m_Size.FRdrHeight - m_Size.Height) \ 2
            Call m_DC.UpdateDC(False)                           ' select bitmap into DC
            If (Me.Mirror And lvicMirrorHorizontal) Then m_Size.Width = -m_Size.Width
            If (Me.Mirror And lvicMirrorVertical) Then m_Size.Height = -m_Size.Height
            
            ' we don't FastRedraw transparency to prevent unnecessary re-creation of the image
            lTrans = m_Effects.GlobalTransparencyPct
            If lTrans Then m_Effects.GlobalTransparencyPct = 0&
            m_Image.Render m_DC.DC, X, Y, m_Size.Width, m_Size.Height, , , , , m_Angle, m_Effects.AttributesHandle, , m_Effects.EffectsHandle(m_Attributes And &HFF&), Me.Interpolation
            m_Effects.TransparentColor = lTrans
            Call m_DC.UpdateDC(False)                           ' remove bitmap from DC
            m_RenderFlags = (m_RenderFlags And Not render_DoFastRedraw)
        Else
            m_RenderFlags = (m_RenderFlags And Not (render_FastRedraw Or render_DoFastRedraw))
            If m_DC.hBitmap(True) = 0& Then Set m_DC = Nothing
        End If
    End If
    
End Sub

'//// helper routine that creates the hit test coordinates
Private Sub sptCreateHitTestPoints(UserDefinedRegion As Long)

    ' called whenever image changes or controls size changes that would effect
    ' the positioning of the image
    ' Note: Do not call if a user-defined region is in use. Exception is HitTest property event

    If (m_RenderFlags And render_DoHitTest) = 0& Then Exit Sub

    Dim cRect As RECTI
    Dim ctrSrc As POINTAPI, ctrDest As POINTAPI
    Dim d2r As Double, a As Double
    Dim sinT As Double, cosT As Double
    
    Erase m_HitTestPts()
    If m_HitRegion Then DeleteObject m_HitRegion: m_HitRegion = 0&
    m_RenderFlags = m_RenderFlags Xor render_DoHitTest
    
    If UserDefinedRegion Then
        m_HitRegion = UserDefinedRegion
        Exit Sub
    ElseIf UserControl.Enabled = False Or m_Size.Width = 0& Then
        Exit Sub
    ElseIf (m_Flags And attr_MouseEvents) = 0& Then
        Exit Sub
    End If
    
    Select Case Me.HitTest
        Case lvicPerimeter                                 ' do nothing
            
        Case lvicEntirelImage, lvicTrimmedImage
            ' ok, now we are going to have some fun with trig. Bet there's a short algo to get the points....
            ' Two type of image hit tests: full image, trimmed image
        
            If Me.HitTest = lvicEntirelImage Then           ' full image; size calculated already
                SetRect cRect, 0&, 0&, m_Size.Width, m_Size.Height
            Else
                If m_Effects.TransparentColorUsed = False Then
                    If modCommon.HasTransparency(m_Image.Handle) Then
                        modCommon.TrimImage m_Image.Handle, cRect, 0&
                    Else
                        SetRect cRect, 0&, 0&, m_Image.Width, m_Image.Height
                    End If
                Else
                    modCommon.TrimImage m_Image.Handle, cRect, &HFF000000 Or (m_Effects.TransparentColor And &HFF00&) _
                            Or (m_Effects.TransparentColor And &HFF&) * &H10000 Or (m_Effects.TransparentColor And &HFF0000) \ &H10000
                End If
                
                If (Me.Mirror And lvicMirrorHorizontal) Then cRect.nLeft = m_Image.Width - (cRect.nWidth + cRect.nLeft)
                If (Me.Mirror And lvicMirrorVertical) Then cRect.nTop = m_Image.Height - (cRect.nHeight + cRect.nTop)
                d2r = m_Size.Width / (m_Image.Width)  ' get horizontal scaling ratio
                cRect.nWidth = cRect.nWidth * d2r           ' calculate scaled size
                cRect.nLeft = cRect.nLeft * d2r
                d2r = m_Size.Height / (m_Image.Height)  ' get vertical scaling ratio
                cRect.nHeight = cRect.nHeight * d2r         ' calculate scaled size
                cRect.nTop = cRect.nTop * d2r
            End If
            
            ReDim m_HitTestPts(0 To 3)
                
            a = (Int(m_Angle) Mod 360!) + (m_Angle - Int(m_Angle))
            If a = 0# Then
                
                If (m_CntFlags And cnt_AlignCenter) Then
                    m_HitTestPts(0).X = (UserControl.ScaleWidth - cRect.nWidth) \ 2& + m_Offsets.X
                    m_HitTestPts(0).Y = (UserControl.ScaleHeight - cRect.nHeight) \ 2& + m_Offsets.Y
                ElseIf (m_CntFlags And cnt_Border) Then
                    m_HitTestPts(0).X = 1& + m_Offsets.X + cRect.nLeft
                    m_HitTestPts(0).Y = 1& + m_Offsets.Y + cRect.nTop
                End If
                m_HitTestPts(1).X = m_HitTestPts(0).X + cRect.nWidth
                m_HitTestPts(1).Y = m_HitTestPts(0).Y
                m_HitTestPts(2).X = m_HitTestPts(1).X
                m_HitTestPts(2).Y = m_HitTestPts(1).Y + cRect.nHeight
                m_HitTestPts(3).X = m_HitTestPts(0).X
                m_HitTestPts(3).Y = m_HitTestPts(2).Y
                            
            Else                                            ' all other angles
                
                cRect.nWidth = (cRect.nWidth + cRect.nLeft)
                cRect.nHeight = (cRect.nHeight + cRect.nTop)
                ctrSrc.X = cRect.nWidth \ 2                 ' cache center of image from 0,0
                ctrSrc.Y = cRect.nHeight \ 2
                If (m_CntFlags And cnt_AlignCenter) Then    ' centering on control....
                    ctrDest.X = UserControl.ScaleWidth \ 2& + m_Offsets.X ' cache center of control
                    ctrDest.Y = UserControl.ScaleHeight \ 2& + m_Offsets.Y
                Else
                    If (m_CntFlags And cnt_Border) Then         ' borders? cache center of image+border size
                        ctrDest.X = ctrSrc.X + 1& + m_Offsets.X: ctrDest.Y = ctrSrc.Y + 1& + m_Offsets.Y
                    Else                                        ' else just cache center of image
                        ctrDest.X = ctrSrc.X + m_Offsets.X: ctrDest.Y = ctrSrc.Y + m_Offsets.Y
                    End If
                End If
        
                d2r = (4& * Atn(1)) / 180#                      ' degree to radian conversion factor ( PI/180 )
                sinT = Sin(a * d2r)                             ' cache SIN & COS of our angle
                cosT = Cos(a * d2r)
                                                                ' and calculate the rotated rectangle's coords
                m_HitTestPts(0).X = (-ctrSrc.X * cosT) - (-ctrSrc.Y * sinT) + ctrDest.X
                m_HitTestPts(0).Y = (-ctrSrc.X * sinT) + (-ctrSrc.Y * cosT) + ctrDest.Y
    
                m_HitTestPts(1).X = (cRect.nWidth - ctrSrc.X) * cosT - (-ctrSrc.Y * sinT) + ctrDest.X
                m_HitTestPts(1).Y = (cRect.nWidth - ctrSrc.X) * sinT + (-ctrSrc.Y * cosT) + ctrDest.Y
    
                m_HitTestPts(2).X = (cRect.nWidth - ctrSrc.X) * cosT - (cRect.nHeight - ctrSrc.Y) * sinT + ctrDest.X
                m_HitTestPts(2).Y = (cRect.nWidth - ctrSrc.X) * sinT + (cRect.nHeight - ctrSrc.Y) * cosT + ctrDest.Y
    
                m_HitTestPts(3).X = (-ctrSrc.X * cosT) - (cRect.nHeight - ctrSrc.Y) * sinT + ctrDest.X
                m_HitTestPts(3).Y = (-ctrSrc.X * sinT) + (cRect.nHeight - ctrSrc.Y) * cosT + ctrDest.Y
            
            End If
    End Select
    
End Sub

'//// Adds a rounded rectangle to a path. If the path does not exist, it is created
Private Sub sptDrawRoundRect(tDC As Long)
' ACKNOWLEDGEMENT: http://www.eggheadcafe.com/software/aspnet/29469337/bob-powells-how-to-draw-a-rounded-rectangle.aspx
    
    Dim hPath As Long, hPen As Long, hGraphics As Long
    Dim Cx As Single
    
    If (UserControl.ForeColor And &H80000000) Then
        hPen = GetSysColor(UserControl.ForeColor And &HFF&)
    Else
        hPen = UserControl.ForeColor
    End If
    If UserControl.ScaleWidth < UserControl.ScaleHeight Then Cx = UserControl.ScaleWidth / 4! Else Cx = UserControl.ScaleHeight / 4!
    
    If g_TokenClass.Token Then
    
        Call GdipCreatePen1(Color_RGBtoARGB(hPen, 255&), 1!, UnitPixel, hPen)
        If hPen Then
        
            GdipCreatePath hPath, hPath
            If hPath Then
                ' Top Left
                If GdipAddPathArc(hPath, 0!, 0!, Cx, Cx, 180!, 90!) = 0& Then
                    ' Top Right
                    If GdipAddPathArc(hPath, UserControl.ScaleWidth - Cx - 1!, 0!, Cx, Cx, 270!, 90!) = 0& Then
                        ' Bottom Right
                        If GdipAddPathArc(hPath, UserControl.ScaleWidth - Cx - 1!, UserControl.ScaleHeight - Cx - 1!, Cx, Cx, 0!, 90!) = 0& Then
                            'Bottom Left
                            GdipAddPathArc hPath, 0!, UserControl.ScaleHeight - Cx - 1!, Cx, Cx, 90!, 90!
                            ' Close will automatically join the path parts
                            GdipClosePathFigure hPath
                            If GdipCreateFromHDC(tDC, hGraphics) = 0& Then
                                If (m_CntFlags And cnt_BorderMask) = cnt_RoundBorder Then
                                    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
                                End If
                                GdipDrawPath hGraphics, hPen, hPath
                                GdipDeleteGraphics hGraphics
                            End If
                        End If
                    End If
                End If
                GdipDeletePath hPath
            End If
            GdipDeletePen hPen
        End If
        
    Else    ' GDI+ not installed
    
        hPen = CreateSolidBrush(hPen)
        hPath = CreateRoundRectRgn(0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, CLng(Cx), CLng(Cx))
        FrameRgn tDC, hPath, hPen, 1&, 1&
        DeleteObject hPath
        DeleteObject hPen
        
    End If

End Sub

'//// helper routine that determines if mouse is within the control. See HitTest property
Private Function sptGetHitTest(X As Long, Y As Long) As Long

  ' About the point-in-polygon algorithm below:
  '     - ACKNOWLEDGEMENT: converted to VB & borrowed from: http://www.visibone.com/inpoly/
  '     - requirement: closed polygon, all points in clockwise or counterclockwise order
  '     - reliability: not 100% accurate for hits on edges of polygon; but very fast
  '     - The points for images are calculated in sptCreateHitTestPoints

    Dim pOld As Long, pNew As Long, lResult As Long
    Dim i As Long, X1 As Long, X2 As Long
    Dim cRect As RECTI

    If UserControl.Enabled Then
        SetRect cRect, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
        If PtInRect(cRect, X, Y) Then
            If m_HitRegion Then
                If PtInRegion(m_HitRegion, X, Y) Then lResult = vbHitResultHit
            ElseIf m_Size.Width = 0& Then
                lResult = vbHitResultHit
            ElseIf (m_Flags And attr_MouseEvents) = 0& Then
                lResult = vbHitResultHit
            ElseIf Me.HitTest = lvicPerimeter Then
                lResult = vbHitResultHit
            Else
                pOld = 3&
                For i = 0& To pOld
                    pNew = i
                    If m_HitTestPts(pNew).X > m_HitTestPts(pOld).X Then
                        X1 = pOld: X2 = pNew
                    Else
                        X1 = pNew: X2 = pOld
                    End If
                    If (m_HitTestPts(pNew).X < X) = (X <= m_HitTestPts(pOld).X) Then
                        If ((Y - m_HitTestPts(X1).Y) * (m_HitTestPts(X2).X - m_HitTestPts(X1).X)) < ((m_HitTestPts(X2).Y - m_HitTestPts(X1).Y) * (X - m_HitTestPts(X1).X)) Then
                            lResult = lResult Xor vbHitResultHit
                        End If
                    End If
                    pOld = pNew
                Next
            End If
        End If
    End If

    sptGetHitTest = lResult

End Function

'//// helper function to calculate rendering area in relation to the control & the source image
Private Function sptGetRepaintArea(hDC As Long, fillArea As RECTI, imageArea As RECTI, targetArea As RECTI) As Boolean

    Dim Offsets As POINTAPI
    
    GetClipBox hDC, targetArea
    If targetArea.nWidth <= targetArea.nLeft Or targetArea.nHeight <= targetArea.nTop Then Exit Function
    fillArea = targetArea
    If m_Image.Handle = 0& Then Exit Function
    
    If (m_CntFlags And cnt_Border) Then         ' area to be filled/painted, excluding borders
        If (m_CntFlags And cnt_AlignCenter) = 0& Then Offsets.X = 1&: Offsets.Y = 1&
    End If
    
    Offsets.X = Offsets.X + m_Offsets.X         ' offsets defining image area adjustments
    Offsets.Y = Offsets.Y + m_Offsets.Y
    
    If (m_RenderFlags And render_FastRedraw) Then
        ' fast redraw has scaled/rotated image already redrawn, determine imageArea to be repainted
        If (m_CntFlags And cnt_AlignCenter) Then
            Offsets.X = Offsets.X + (UserControl.ScaleWidth - m_Size.FRdrWidth) \ 2
            Offsets.Y = Offsets.Y + (UserControl.ScaleHeight - m_Size.FRdrHeight) \ 2
        End If
        ' calculate position of fastredraw image within control
        SetRect imageArea, Offsets.X, Offsets.Y, m_Size.FRdrWidth + Offsets.X, m_Size.FRdrHeight + Offsets.Y
        If IntersectRect(imageArea, imageArea, targetArea) Then
            targetArea = imageArea
            OffsetRect imageArea, -Offsets.X, -Offsets.Y
        End If
    Else
        If (m_CntFlags And cnt_AlignCenter) Then
            Offsets.X = Offsets.X + (UserControl.ScaleWidth - m_Size.Width) \ 2
            Offsets.Y = Offsets.Y + (UserControl.ScaleHeight - m_Size.Height) \ 2
        End If
        ' calculate position of full size image within control
        SetRect imageArea, Offsets.X, Offsets.Y, m_Size.Width + Offsets.X, m_Size.Height + Offsets.Y
        If m_Size.One2One = True Then
            If m_Effects.EffectsHandle(Me.Effect) = 0& Then
                If IntersectRect(imageArea, imageArea, targetArea) Then
                    targetArea = imageArea
                    OffsetRect imageArea, -Offsets.X, -Offsets.Y
                End If
            End If
        End If
    End If
    
    imageArea.nWidth = imageArea.nWidth - imageArea.nLeft
    imageArea.nHeight = imageArea.nHeight - imageArea.nTop
    
    targetArea.nWidth = targetArea.nWidth - targetArea.nLeft
    targetArea.nHeight = targetArea.nHeight - targetArea.nTop
    
    sptGetRepaintArea = (imageArea.nHeight > 0& And imageArea.nWidth > 0&)

End Function

'//// helper routine that determines image scaling and/or control dimensions
Private Sub sptGetScaledSizes(bRefresh As Boolean)

    ' Routine handles caching of calculated image sizes and/or resizing the control as needed
    
    If (m_RenderFlags And render_AutoSizing) Then Exit Sub
    
    If (m_RenderFlags And (render_DoReScale Or render_DoResize Or render_RedoAutoSize)) = 0& Then Exit Sub
    
    Dim ucCx As Long, ucCy As Long, theAngle As Long, a As Single, lFlags As Long
    
    If m_Image.Handle = 0& Then
        m_Size.Width = 0&
        If (m_RenderFlags And render_PrePost) Then
            If Not m_DC Is Nothing Then
                ' patch to ensure cached bmp resized when no image & WantPrePostEvents=True
                If Not (m_DC.Width(True) = UserControl.ScaleWidth And m_DC.Height(True) = UserControl.ScaleHeight) Then
                    m_Size.Width = -1&
                    ucCx = m_DC.Width(True)
                    ucCy = m_DC.Height(True)
                End If
            End If
        End If
        
    Else
        
        ucCx = UserControl.ScaleWidth       ' set default control size
        ucCy = UserControl.ScaleHeight
        
        ' patch to prevent some outrageously large image to resize control beyond its max-OS-defined size
        ' We will use the screen's size as the delimiting factor
        If (m_RenderFlags And render_DoResize) Then
            If ucCx > Screen.Width \ Screen.TwipsPerPixelX Then ucCx = Screen.Width \ Screen.TwipsPerPixelX
            If ucCy > Screen.Height \ Screen.TwipsPerPixelY Then ucCy = Screen.Height \ Screen.TwipsPerPixelY
        End If
        
        If (m_CntFlags And cnt_Border) Then ucCx = ucCx - 2&: ucCy = ucCy - 2&
        
        Select Case Me.Aspect
        
        Case lvicStretch                        ' no AutoSizing when scaling is Stretched
            m_Size.Width = ucCx
            m_Size.Height = ucCy
            a = Abs(Int(m_Angle) Mod 360!) + (m_Angle - Int(m_Angle))
            If Not (a = 180! Or a = 0!) Then modCommon.GetScaledImageSizes ucCx, ucCy, ucCx, ucCy, m_Size.Width, m_Size.Height, m_Angle, True, False
        
        Case lvicActualSize                     ' no stretching or scaling
            m_Size.Width = m_Image.Width
            m_Size.Height = m_Image.Height
            Select Case Me.AutoSize
                Case lvicNoAutoSize             ' nothing extra to do, image is displayed as is, clipped if needed
                Case lvicSingleAngle            ' get new control sizes based on rotated image size
                    modCommon.GetScaledCanvasSize m_Image.Width, m_Image.Height, ucCx, ucCy, m_Angle
                    ' patch to prevent some outrageously large image to resize control beyond its max-OS-defined size
                    ' We will use the container's size as the delimiting factor
                    If ucCx > Screen.Width \ Screen.TwipsPerPixelX Then ucCx = Screen.Width \ Screen.TwipsPerPixelX
                    If ucCy > Screen.Height \ Screen.TwipsPerPixelY Then ucCy = Screen.Height \ Screen.TwipsPerPixelY
                Case Else                       ' set new control size to size for all angles
                    ucCy = Sqr(m_Size.Width * m_Size.Width + m_Size.Height * m_Size.Height)
                    ucCx = ucCy
            End Select
            
        Case lvicFixedSize, lvicFixedSizeStretched
            If Me.Aspect = lvicFixedSizeStretched Then
                m_Size.Width = m_Size.FixedCx
                m_Size.Height = m_Size.FixedCy
            Else
                modCommon.GetScaledImageSizes m_Image.Width, m_Image.Height, m_Size.FixedCx, m_Size.FixedCy, m_Size.Width, m_Size.Height, , False, False
            End If
            Select Case Me.AutoSize
                Case lvicSingleAngle
                    ucCx = m_Size.Width: ucCy = m_Size.Height
                Case lvicMultiAngle
                    ucCy = Sqr(m_Size.Width * m_Size.Width + m_Size.Height * m_Size.Height)
                    ucCx = ucCy
            End Select
        
        Case lvicScaled, lvicScaleDownOnly          ' scaling, a bit more intense
            ' get scaled image size that can fit in current dimensions
            
            Select Case Me.AutoSize
                Case lvicMultiAngle
                    If (m_RenderFlags And render_RedoAutoSize) Then ' changing autosize to MultiAngle
                        If m_Size.Width = 0& Then
                            modCommon.GetScaledImageSizes m_Image.Width, m_Image.Height, ucCx, ucCy, m_Size.Width, m_Size.Height, 45, (Me.Aspect = lvicScaled), False
                        Else
                            m_Size.Width = m_Image.Width
                            m_Size.Height = m_Image.Height
                        End If
                    Else
                        modCommon.GetScaledImageSizes m_Image.Width, m_Image.Height, ucCx, ucCy, m_Size.Width, m_Size.Height, 45, (Me.Aspect = lvicScaled), False
                    End If
                    modCommon.GetScaledCanvasSize m_Size.Width, m_Size.Height, ucCx, ucCy, 45
                Case lvicSingleAngle
                    If (m_RenderFlags And render_DoResize) Or m_Size.Width = 0& Then ' resizing control
                        modCommon.GetScaledImageSizes m_Image.Width, m_Image.Height, ucCx, ucCy, m_Size.Width, m_Size.Height, m_Angle, (Me.Aspect = lvicScaled), False
                    End If
                    modCommon.GetScaledCanvasSize m_Size.Width, m_Size.Height, ucCx, ucCy, m_Angle
                Case Else
                    modCommon.GetScaledImageSizes m_Image.Width, m_Image.Height, ucCx, ucCy, m_Size.Width, m_Size.Height, m_Angle, (Me.Aspect = lvicScaled), False
            End Select
        
        End Select
        
        ' is this a one to one ratio. Faster drawing algos used if it is
        a = Abs(Int(m_Angle) Mod 360!) + (m_Angle - Int(m_Angle))
        m_Size.One2One = ((m_Size.Width = m_Image.Width) And (m_Size.Height = m_Image.Height) And (a = 0!))
        
        If (m_CntFlags And cnt_Border) Then ucCx = ucCx + 2&: ucCy = ucCy + 2&
    End If
    
    If Not (ucCx < 1& And ucCy < 1&) Then
        If (UserControl.ScaleWidth = ucCx And UserControl.ScaleHeight = ucCy) Then
            If m_Size.Width = -1& Then m_Size.Width = 0&
        Else
            If m_Size.Width = -1& Then  ' patch to ensure cached bmp resized when no image & WantPrePostEvents=True
                ucCx = UserControl.ScaleWidth
                ucCy = UserControl.ScaleHeight
                m_Size.Width = 0&
            Else
                ' set flags to prevent Paint event from running its code & this routine from recursing
                lFlags = m_RenderFlags
                m_RenderFlags = (m_RenderFlags And Not render_Shown) Or render_AutoSizing
                ' resize the control & reset the flag we just set
                UserControl.Size ScaleX(ucCx, vbPixels, vbTwips), ScaleY(ucCy, vbPixels, vbTwips)
                m_RenderFlags = lFlags
            End If
        End If
        ' if WantPrePostEvents=True, then resize bitmap used for pre-post events as needed
        If (m_RenderFlags And render_PrePost) Then
            If m_DC.ResizeBitmap(ucCx, ucCy, 24&, True) = False Then
                m_RenderFlags = (m_RenderFlags And Not render_PrePost)
                If m_DC.hBitmap(False) = 0& Then Set m_DC = Nothing
            End If
        End If
    End If
    
    m_RenderFlags = (m_RenderFlags And Not (render_DoReScale Or render_DoResize Or render_RedoAutoSize))
    If Me.SetRedraw = True And (m_RenderFlags And render_Shown) = render_Shown Then
        If m_HitRegion = 0& Then Call sptCreateHitTestPoints(m_HitRegion)
        If bRefresh Then UserControl.Refresh
    End If

End Sub

'//// uses GDI+ to create gradient
Private Function sptFillGradient(hDC As Long) As Boolean
    
    If g_TokenClass.Token = 0! Then Exit Function
    
    Dim dwColor1 As Long, dwColor2 As Long
    Dim hGraphics As Long, hBrush As Long
    Dim pRect As RECTI
    Const BrushWrapModeTileFlipy As Long = &H2
    
    If Me.GradientStyle = lvicGradientDiagonalUp Then
        dwColor2 = UserControl.FillColor: dwColor1 = UserControl.BackColor
    Else
        dwColor1 = UserControl.FillColor: dwColor2 = UserControl.BackColor
    End If
    If GdipCreateFromHDC(hDC, hGraphics) = 0& Then
        If Me.Border Then
            SetRect pRect, 1&, 1&, UserControl.ScaleWidth - 2&, UserControl.ScaleHeight - 2&
        Else
            SetRect pRect, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
        If GdipCreateLineBrushFromRectI(pRect, Color_RGBtoARGB(dwColor1, 255), Color_RGBtoARGB(dwColor2, 255), Me.GradientStyle - 1&, BrushWrapModeTileFlipy, hBrush) = 0& Then
            GdipFillRectangleI hGraphics, hBrush, pRect.nLeft, pRect.nTop, pRect.nWidth, pRect.nHeight
            GdipDeleteBrush hBrush
            sptFillGradient = True
        End If
        GdipDeleteGraphics hGraphics
    End If

End Function

'//// helper function to draw the control to an outside DC
Private Function sptRemoteRender(toHandle As Boolean, Destination As Long, IncludeContainerBkg As Boolean) As Boolean

    If Me.SetRedraw = False Then
        Me.SetRedraw = True
        Exit Function
    End If
    
    Dim hHandle As Long, gPic As GDIpImage
    
    Dim sControlName As String, tSizeI As RECTI
    Dim meControl As Control, myForm As Object
    Dim X As Long, Index As Integer
    
    Const RDW_ERASE As Long = &H4
    Const RDW_INVALIDATE As Long = &H1
    Const RDW_NOCHILDREN As Long = &H40
    Const RDW_ERASENOW As Long = &H200
    Const RDW_UPDATENOW As Long = &H100
    
    On Error Resume Next
    Set myForm = ParentControls(0)                  ' set instance of form/mdi
    X = myForm.Controls.Count
    If Err Then
        Err.Clear
        Exit Function
    End If
    
    sControlName = Ambient.DisplayName              ' get our control's name
    X = InStr(sControlName, "(")                    ' indexed?
    If X Then                                       ' if so, get the index
        Index = Val(Mid$(sControlName, X + 1))      ' adjust control name & assign
        sControlName = Left$(sControlName, X - 1)
        Set meControl = myForm.Controls(sControlName)(Index)
    Else                                            ' assign
        Set meControl = myForm.Controls(sControlName)
    End If
    If Err Then Exit Function
    
    Set myForm = Nothing                            ' done with this
    X = meControl.Container.ScaleMode
    If Err Then Err.Clear: X = vbTwips
    On Error GoTo 0
    
    tSizeI.nLeft = ScaleX(meControl.Left, X, vbPixels)
    tSizeI.nTop = ScaleY(meControl.Top, X, vbPixels)
    tSizeI.nWidth = tSizeI.nLeft + UserControl.ScaleWidth
    tSizeI.nHeight = tSizeI.nTop + UserControl.ScaleHeight
    
    If toHandle Then
        If g_TokenClass.Token = 0! Then Exit Function
        If Me.Border = True And IncludeContainerBkg = False Then X = -2& Else X = 0&
        If GdipCreateBitmapFromScan0(UserControl.ScaleWidth + X, UserControl.ScaleHeight + X, 0&, lvicColor32bppAlpha, ByVal 0&, hHandle) Then Exit Function
        If GdipGetImageGraphicsContext(hHandle, m_Size.DestDC) Then
            GdipDisposeImage hHandle: Exit Function
        End If
        m_RenderFlags = m_RenderFlags Or render_RemoteToHGraphics
    Else
        m_Size.DestDC = Destination
    End If
    
    If IncludeContainerBkg Then m_RenderFlags = m_RenderFlags Or render_RemoteDrawBkg
    RedrawWindow UserControl.ContainerHwnd, tSizeI, 0&, RDW_ERASE Or RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_NOCHILDREN
    If IncludeContainerBkg Then m_RenderFlags = m_RenderFlags And Not render_RemoteDrawBkg

    If toHandle Then
        GdipDeleteGraphics m_Size.DestDC
        Destination = hHandle
        m_RenderFlags = m_RenderFlags Xor render_RemoteToHGraphics
    End If
    m_Size.DestDC = 0&
    sptRemoteRender = True

End Function

'//// helper function for Paint routine. Creates a custom clipping region
Private Function sptSetClippingTarget(hDC As Long, includeRgn As Long) As Long

    Dim cRgn As Long, hRgn As Long
    Dim rgnRect As RECTI
    Const RGN_AND As Long = 1

    hRgn = CreateRectRgn(0, 0, 0, 0)   ' copy clipping region VB assigned to usercontrol
    If GetClipRgn(hDC, hRgn) = 0& Then
        DeleteObject hRgn
    Else
        GetRgnBox hRgn, rgnRect             ' get its bounds & adjust our region accordingly
        sptSetClippingTarget = hRgn
    End If
    
    If m_ClipRect.nHeight Then
        hRgn = CreateRectRgn(m_ClipRect.nLeft, m_ClipRect.nTop, m_ClipRect.nWidth, m_ClipRect.nHeight)
        If includeRgn Then
            CombineRgn hRgn, hRgn, includeRgn, RGN_AND
            DeleteObject includeRgn
        End If
    Else
        hRgn = includeRgn
    End If
    If hRgn Then
        If sptSetClippingTarget <> 0& Then
            OffsetRgn hRgn, rgnRect.nLeft, rgnRect.nTop
            CombineRgn hRgn, hRgn, sptSetClippingTarget, RGN_AND
        End If
        SelectClipRgn hDC, hRgn           ' replace clip region
        DeleteObject hRgn                 ' destroy; no longer needed
    ElseIf sptSetClippingTarget <> 0& Then
        DeleteObject sptSetClippingTarget
        sptSetClippingTarget = 0&
    End If

End Function

'//// helper routine to determine if control is in runtime or design-time
Private Sub sptSetUserMode()
    On Error Resume Next
    Dim bMode As Boolean
    bMode = UserControl.Ambient.UserMode
    If Err Then bMode = True ' assume runtime
    If bMode Then m_CntFlags = m_CntFlags Or cnt_Runtime
    On Error GoTo 0
End Sub

'//// callback event when background image async download is done
Private Sub m_BkgImage_AsyncDownloadDone(Success As Boolean, ErrorCode As AsyncDownloadStatusEnum)
    RaiseEvent AsyncDownloadDoneBkgImg((Success), ErrorCode)
    If Me.SetRedraw = True And Success Then UserControl.Refresh
End Sub

'//// callback event when background image changes
Private Sub m_BkgImage_ImageChanged(FrameIndex As Long)
    If Me.SetRedraw = True Then UserControl.Refresh
End Sub

'//// callback event when GDIpEffects updates an attribute/effect
Private Sub m_Effects_Changed(PropertyName As String)

    If (m_RenderFlags And render_Shown) = 0& Then Exit Sub
    If (m_CntFlags And cnt_InitLoad) Then Exit Sub

    Select Case PropertyName
        Case "GrayScale"
            m_RenderFlags = m_RenderFlags Or render_DoFastRedraw
            PropertyChanged "GrayScale"
            
        Case "BlendColor", "BlendPct"
            m_RenderFlags = m_RenderFlags Or render_DoFastRedraw
            PropertyChanged "BlendColor"
            
        Case "GlobalTransparency"
            If Me.HitTest = lvicTrimmedImage Then m_RenderFlags = m_RenderFlags Or render_DoHitTest
            PropertyChanged "TransparencyPct"
            
        Case "LightnessPct"
            m_RenderFlags = m_RenderFlags Or render_DoFastRedraw
            PropertyChanged "LightnessPct"
            
        Case "TransparentColor", "TransparentColorUsed"
            PropertyChanged "TransparentColorMode"
            If PropertyName = "TransparentColorUsed" Then
                If m_Effects.TransparentColorUsed Then
                    m_Attributes = (m_Attributes And &HFFFF00FF) Or lvicUseTransparentColor * &H100&
                Else
                    m_Attributes = (m_Attributes And &HFFFF00FF)
                End If
            ElseIf m_Effects.TransparentColorUsed = False Then
                Exit Sub
            End If
            m_RenderFlags = m_RenderFlags Or render_DoFastRedraw
            If Me.HitTest = lvicTrimmedImage Then m_RenderFlags = m_RenderFlags Or render_DoHitTest
            
        Case "Invert"
            m_RenderFlags = m_RenderFlags Or render_DoFastRedraw
            PropertyChanged "Inverted"
            
        Case Else   ' effects changed
            PropertyChanged "Effects"
            If Me.Effect <> Val(Right$(PropertyName, 1)) Then Exit Sub ' no need to refresh
            
    End Select
    If Me.SetRedraw = True Then UserControl.Refresh
    
End Sub

'//// callback event when control's source image async download is done
Private Sub m_Image_AsyncDownloadDone(Success As Boolean, ErrorCode As Long)
    RaiseEvent AsyncDownloadDone(Success, ErrorCode)
    If CanPropertyChange("SinkImageData") Then PropertyChanged "SinkImageData"
'    If Success = False Then Debug.Print ObjPtr(Me); " failed "; ErrorCode Else Debug.Print ObjPtr(Me); " image downloaded"
End Sub

'//// standard click event
Private Sub UserControl_Click()
    ' Click event not forwarded thru this event.
    ' A standard click event looks like this for buttons:    mouseDown, Click, mouseUp
    '   But VB sends usercontrols the click order as:        mouseDown, mouseUp, Click
    ' Likewise, a usercontrol double click event looks like: mouseDown, mouseUp, Click, DblClick, mouseUp
    ' Well that can get a bit troublesome if using this control like a button because you may be
    '   offsetting image during a MouseDown and resetting it during MouseUp. Notice that you
    '   do not get 2 mouse down events during a double click.. So we control 3 of the 4 events
    '   Click events will be raised in this order, only if done with left mouse button: mouseDown, Click, mouseUp
    '   and Double Click events in this order, only if done with left button:  mouseDown, Click, mouseUp, DblClick
    '   Click & Double Click events done with Middle/Right button are not raised, just the mouseUp & mouseDown are
End Sub

'//// standard double click event & triggers for any mouse button's double click
Private Sub UserControl_DblClick()
    ' see Usercontrol_Click for more info
End Sub

'//// event occurs when control is about to be destroyed
Private Sub UserControl_Hide()
    If Not m_Animator Is Nothing Then m_Animator.PauseAnimation
    If Not m_MouseTracker Is Nothing Then
        m_MouseTracker.ReleaseMouseCapture True, ObjPtr(Me)
        Set m_MouseTracker = Nothing
    End If
    m_RenderFlags = (m_RenderFlags And Not render_Shown)
End Sub

'//// event requested by container to see if control want's to accept the pending mouse event(s)
Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If (m_CntFlags And cnt_Runtime) = 0& Then
        HitResult = vbHitResultHit
    ElseIf (m_CntFlags And cnt_MouseValidate) Then
        HitResult = vbHitResultHit
    Else
        If (m_CntFlags And cnt_MouseDown) Then
            HitResult = vbHitResultHit
        Else
            HitResult = sptGetHitTest(CLng(X), CLng(Y))
            If Not m_MouseTracker Is Nothing Then           ' if tracking for MouseExit, fire event if not valid hitTest
                If HitResult = vbHitResultOutside Then
                    Call m_MouseTracker.ReleaseMouseCapture(True, ObjPtr(Me))
                End If
            End If
        End If
    End If
End Sub

'//// event occurs every time control is about to be created
Private Sub UserControl_Initialize()
    Set m_Image = New GDIpImage
    Set m_Effects = New GDIpEffects
    UserControl.ScaleMode = vbPixels
End Sub

'//// event occurs only once, first time control is created
Private Sub UserControl_InitProperties()
    m_Flags = attr_MouseEvents
    m_CntFlags = cnt_AlignCenter ' autosize at actualsize
    UserControl.FillColor = vbWhite
End Sub

'//// standard mouse-down event
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_CntFlags = (m_CntFlags And &HFFFFFF) Or Button * cnt_LastButtonShift
    If (m_Flags And attr_MouseEvents) Then RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'//// standard mouse-move event but also tracks MouseEnter & auto-OLEDrag events
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (m_CntFlags And cnt_MouseValidate) Then
        m_CntFlags = m_CntFlags Xor cnt_MouseValidate           ' test message; abort
        Exit Sub
    End If
    ' before raising a MouseMove and possible MouseEnter event, force any currently tracking
    '   instance to fire a MouseExit event if needed
    ' Note: This routine is only called if UserControl_HitTest is non-zero
    
    If m_MouseTracker Is Nothing Then                           ' we aren't tracking, is another control?
        If (m_Flags And attr_MouseEvents) Then
            g_MouseExitClass.Iniitate UserControl.ContainerHwnd, ObjPtr(Me)
            RaiseEvent MouseEnter                               ' raise MouseEnter
            RaiseEvent MouseMove(Button, Shift, X, Y)           ' raise MouseMove
        Else
            g_MouseExitClass.Iniitate 0&, ObjPtr(Me)            ' setup a null tracker
        End If
        Set m_MouseTracker = g_MouseExitClass                   ' set reference
    ElseIf (m_Flags And attr_MouseEvents) Then
        RaiseEvent MouseMove(Button, Shift, X, Y)               ' raise MouseMove
    End If
    
    If Button = vbLeftButton Then
        m_CntFlags = m_CntFlags Or cnt_MouseDown
        If Me.OLEDragMode = vbOLEDragAutomatic Then
            If m_DragDrop.AutoDragPts.X = 0& And m_DragDrop.AutoDragPts.Y = 0& Then
                m_DragDrop.AutoDragPts.X = X: m_DragDrop.AutoDragPts.Y = Y
            ElseIf Abs(m_DragDrop.AutoDragPts.X - X) > 3& Then
                Me.OLEDrag
            ElseIf Abs(m_DragDrop.AutoDragPts.Y - Y) > 3& Then
                Me.OLEDrag
            End If
        End If
    End If
End Sub

'//// standard mouse-up event
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then m_CntFlags = (m_CntFlags And Not cnt_MouseDown)
    
    If (m_CntFlags And cnt_DblClicked) Then
        m_CntFlags = m_CntFlags Xor cnt_DblClicked
        RaiseEvent DblClick
    Else
        If Button = vbLeftButton Then
            If (m_CntFlags And &HF000000) \ cnt_LastButtonShift = vbLeftButton Then
                If sptGetHitTest(CLng(X), CLng(Y)) Then
                    m_CntFlags = m_CntFlags Or cnt_DblClicked
                    RaiseEvent Click
                End If
            End If
        End If
        If (m_Flags And attr_MouseEvents) Then
            If (m_CntFlags And &HF000000) Then
                RaiseEvent MouseUp(Button, Shift, X, Y)
            End If
        End If
        m_CntFlags = (m_CntFlags And &HF0FFFFFF)
    End If
End Sub

'//// standard OLE event
Private Sub UserControl_OLECompleteDrag(Effect As Long)
    If Me.OLEDragMode = vbOLEDragManual Then RaiseEvent OLECompleteDrag(Effect)
End Sub

'//// standard OLE event
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Note. Unlike VB when DropMode is Automatic, this control will give you the option to abort
    ' When DropMode is automatic and you wish to decide to accept the Data object or not, you
    '   must monitor and reply to this control's OLEDragDrop event. If you wish that the Data
    '   object not be processed, pass the Cancel parameter as True
    ' When Data object contains dropped files, the file names will contain unicode characters if applicable
    If m_DragDrop.Originator Then
        Effect = vbDropEffectNone
    Else
        Dim bCancel As Boolean, tPic As GDIpImage
        ' when files are dropped, populate Data object with unicode filenames, as needed,
        '   before sending the OLEDragDrop event to the user. Only need to do this if
        '   Automatic drop mode; else already done in OLEDDragOver
        If Me.OLEDropMode = vbOLEDropAutomatic Then
            If Data.GetFormat(vbCFFiles) Then Call modCommon.GetDroppedFileNames(Data, vbCFFiles)
        End If
        RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y, bCancel)
        If Not bCancel Then
            Set tPic = modCommon.LoadImage(Data, True, True)
            If Not tPic Is Nothing Then Me.Picture = tPic
        End If
    End If
    
End Sub

'//// standard OLE event
Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    If UserControl.OLEDropMode = vbOLEDropNone Or m_DragDrop.Originator = True Then
        Effect = vbDropEffectNone
    
    ElseIf Me.OLEDropMode = vbOLEDropAutomatic Then
        If State = vbEnter Then         ' only check this once per drag
            Dim iFormat As Integer, lFormat As Long
            m_DragDrop.Effect = vbDropEffectNone
            For iFormat = 1& To 9&
                Select Case iFormat
                    Case 1&: lFormat = vbCFFiles
                    Case 2&: lFormat = g_ClipboardFormat
                    Case 3&: lFormat = vbCFBitmap
                    Case 4&: lFormat = vbCFDIB
                    Case 5&: lFormat = vbCFMetafile
                    Case 6&: lFormat = vbCFEMetafile
                    Case 7&: lFormat = vbCFText
                    Case 8&: lFormat = CF_UNICODE
                    Case 9&: lFormat = CF_FILEGROUPDESCRIPTORW
                End Select
                If lFormat Then
                    If Data.GetFormat(lFormat) Then
                        m_DragDrop.Effect = vbDropEffectCopy
                        Exit For
                    End If
                End If
            Next
        End If
        Effect = m_DragDrop.Effect
    Else
        ' when files are dropped, populate Data object with unicode filenames, as needed,
        '   before sending the OLEDragDrop event to the user
        If State = vbEnter Then
            If Data.GetFormat(vbCFFiles) Then Call modCommon.GetDroppedFileNames(Data, vbCFFiles)
        End If
        ' you must set the Effect parameter to one of VB's OLEDropEffectConstants values
        RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
    End If

End Sub

'//// standard OLE event
Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    If UserControl.OLEDropMode = vbOLEDropNone Or m_DragDrop.Originator = True Then
        Effect = vbDropEffectNone
    Else
        If Me.OLEDropMode = vbOLEDropManual Then RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
    End If
End Sub

'//// standard OLE event
Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    If Me.OLEDragMode = vbOLEDragManual Then RaiseEvent OLESetData(Data, DataFormat)
End Sub

'//// standard OLE event
Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    
    Data.Clear
    If Me.OLEDragMode = vbOLEDragAutomatic Then
        If modCommon.SaveImage(m_Image, Data, lvicSaveAsCurrentFormat) Then
            AllowedEffects = vbDropEffectCopy
        Else
            AllowedEffects = vbDropEffectNone
        End If
    Else
        RaiseEvent OLEStartDrag(Data, AllowedEffects)
        ' if user doesn't set AllowedEffects, VB will cancel drag event
    End If
    
End Sub

'//// standard paint event. When called, control is partially or completely erased
Private Sub UserControl_Paint()

    If (m_RenderFlags And render_ShowEvent) = 0& Then Exit Sub ' control not yet shown or Hidden event called
    If (m_RenderFlags And render_Shown) = 0& Then
        m_RenderFlags = m_RenderFlags Or render_Shown Or render_InitAnimation
    End If
    
    Dim X As Long, Y As Long, tDC As Long, hBrush As Long, hPen As Long
    Dim srcRect As RECTI, dstRect As RECTI, bkgRect As RECTI
    Dim bCancel As Boolean, bDirtyImage As Boolean, rDC As Long, hGraphics As Long
    ' following used for round rectangle borders only
    Dim hRgn As Long, hOldRgn As Long, rgnRect As RECTI
    
    '////////// SetRedraw = False. When set to False, a snapshot of the current image was created
    If Me.SetRedraw Then
        If (m_RenderFlags And (render_DoReScale Or render_DoResize Or render_RedoAutoSize)) Then
            Call sptGetScaledSizes(True)
            If Me.SetRedraw Then Exit Sub
        End If
        If (m_RenderFlags And render_DoFastRedraw) Then Call sptCreateFastRedrawImage   ' update FastRedraw as needed
        If (m_RenderFlags And render_DoHitTest) Then
            If m_HitRegion = 0& Then Call sptCreateHitTestPoints(m_HitRegion)
        End If
    End If
    
    ' get the update rect of the control
    bDirtyImage = sptGetRepaintArea(UserControl.hDC, bkgRect, srcRect, dstRect)
    
    '/////////  Setup image offsets in relation to control's dimensions
    If (m_CntFlags And cnt_Border) Then     ' border used, create pen for the border
        If (m_CntFlags And cnt_BorderMask) = 0& Then ' rectangular border
            If (UserControl.ForeColor And &H80000000) Then
                hPen = CreatePen(0&, 1&, GetSysColor(UserControl.ForeColor And &HFF&))
            Else
                hPen = CreatePen(0&, 1&, UserControl.ForeColor)
            End If
        ElseIf (m_CntFlags And cnt_BorderMask) <> cnt_BorderMask Then  ' else user-defined border
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then X = UserControl.ScaleWidth / 4! Else X = UserControl.ScaleHeight / 4!
            ' create a round rectangle region if border type requested
            hRgn = CreateRoundRectRgn(0&, 0&, UserControl.ScaleWidth + 1&, UserControl.ScaleHeight + 1&, X, X)
        End If
    ElseIf m_Image.Handle = 0& Then         ' no image? When in design view, draw dotted border
        If (m_CntFlags And cnt_Runtime) = 0& Then hPen = CreatePen(4&, 1&, vbBlack)
    End If
    
    '////////// Pre-Paint
    If (m_RenderFlags And render_PrePost) Then      ' allow user to paint between background & image
        m_DC.UpdateDC True                          ' put 24bpp bitmap into DC & usercontrol to that DC
        tDC = m_DC.DC                               ' DC we use to paint on
        If (m_CntFlags And cnt_Opaque) = 0& And m_DC.BitDepth(True) = 32& Then
            m_DC.EraseBitmap True
        Else
            If (m_CntFlags And cnt_Opaque) = cnt_Opaque Then
                If (m_CntFlags And cnt_GradientMask) And UserControl.BackColor <> UserControl.FillColor Then
                    If sptFillGradient(tDC) = False Then
                        m_CntFlags = (m_CntFlags And Not cnt_GradientMask)
                        m_DC.FillBitmap True, UserControl.BackColor
                    End If
                Else
                    m_DC.FillBitmap True, UserControl.BackColor
                End If
            Else
                With bkgRect
                    BitBlt tDC, .nLeft, .nTop, .nWidth, .nHeight, UserControl.hDC, .nLeft, .nTop, vbSrcCopy
                End With
            End If
        End If
        If (hRgn Or m_ClipRect.nHeight) Then DeleteObject sptSetClippingTarget(tDC, hRgn)
        With bkgRect
            RaiseEvent PrePaint((tDC), (.nLeft), (.nTop), (.nWidth), (.nHeight), (m_HitRegion), bCancel)
        End With
    Else
        tDC = UserControl.hDC
        If (hRgn Or m_ClipRect.nHeight) Then hOldRgn = sptSetClippingTarget(tDC, hRgn)
        If (m_CntFlags And cnt_Opaque) = cnt_Opaque Then    ' create background brush as needed
            If (m_CntFlags And cnt_GradientMask) And UserControl.BackColor <> UserControl.FillColor Then
                hBrush = sptFillGradient(tDC)
                If hBrush = 0& Then m_CntFlags = (m_CntFlags And Not cnt_GradientMask)
            End If
            If hBrush = 0& Then
                If (UserControl.BackColor And &H80000000) Then
                    hBrush = GetSysColorBrush(UserControl.BackColor And &HFF&)
                Else
                    hBrush = CreateSolidBrush(UserControl.BackColor)
                End If
                FillRect tDC, bkgRect, hBrush ' fill bacground color as needed
                If (UserControl.BackColor And &H80000000) = 0& Then DeleteObject hBrush
            End If
        End If
    End If
    If Not m_BkgImage Is Nothing Then
        If m_BkgImage.Handle Then
            If Me.Border Then X = 1& Else X = 0&
            If Me.BkgImageStretch Then
                m_BkgImage.Render tDC, X, X, UserControl.ScaleWidth - X * 2&, UserControl.ScaleHeight - X * 2&
            Else
                m_BkgImage.Render tDC, X, X, m_BkgImage.Width, m_BkgImage.Height
            End If
            X = 0&
        End If
    End If
    
    If bCancel Then
        bDirtyImage = False
        If m_Size.DestDC <> 0& Then m_RenderFlags = m_RenderFlags Or render_RemoteDrawBkg
    End If
    
    '////////// Paint
    If bDirtyImage = True Then
        If m_Size.DestDC <> 0& And (m_RenderFlags And render_RemoteDrawBkg) = 0& Then
            If (m_RenderFlags And render_RemoteToHGraphics) Then
                hGraphics = m_Size.DestDC
            Else
                rDC = m_Size.DestDC
            End If
        Else
            rDC = tDC
        End If
        Do
            If (m_RenderFlags And render_FastRedraw) = render_FastRedraw And rDC = tDC Then
                ' FastRedraw image is always 32bpp pre-multiplied; so we use AlphaBlend vs GDI+ (faster overall)
                With dstRect
                    If rDC = m_DC.DC Then                       ' render from cached bitmap to cached bitmap
                        m_DC.TransferFastRedrawToPrePost .nLeft, .nTop, srcRect.nWidth, srcRect.nHeight, srcRect.nLeft, srcRect.nTop, Me.TransparencyPct
                    Else                                        ' render from cached bitmap to uc dc
                        m_DC.UpdateDC False
                        AlphaBlend rDC, .nLeft, .nTop, _
                            srcRect.nWidth, srcRect.nHeight, m_DC.DC, srcRect.nLeft, srcRect.nTop, srcRect.nWidth, srcRect.nHeight, &H1000000 Or ((((100& - Me.TransparencyPct) * 255&) \ 100&) * &H10000)
                        m_DC.UpdateDC False
                    End If
                End With
            Else                                                ' render directly to uc dc
                X = 1&: Y = X
                If (Me.Mirror And lvicMirrorHorizontal) Then X = -X
                If (Me.Mirror And lvicMirrorVertical) Then Y = -Y
                If m_Size.One2One = True And m_Effects.EffectsHandle(Me.Effect) = 0& Then ' faster repaints, repaint just dirty area
                    If rDC <> tDC Then
                        If Me.Border = True Then OffsetRect dstRect, -1, -1
                    End If
                    With dstRect
                        m_Image.Render rDC, .nLeft, .nTop, .nWidth, .nHeight, _
                            srcRect.nLeft, srcRect.nTop, srcRect.nWidth * X, srcRect.nHeight * Y, m_Angle, _
                            m_Effects.AttributesHandle, hGraphics, m_Effects.EffectsHandle(m_Attributes And &HFF&), Me.Interpolation
                    End With
                Else                                        ' slower repaints; rotated or stretched so repaint all
                    dstRect = srcRect
                    If rDC <> tDC Then
                        If Me.Border = True Then OffsetRect dstRect, -1, -1
                    End If
                    With dstRect
                        m_Image.Render rDC, .nLeft, .nTop, m_Size.Width * X, m_Size.Height * Y _
                            , , , , , m_Angle, m_Effects.AttributesHandle, hGraphics, m_Effects.EffectsHandle(m_Attributes And &HFF&), Me.Interpolation
                    End With
                End If
            End If
            If rDC = tDC Then Exit Do
            rDC = tDC: hGraphics = 0&: X = 0&
            If Me.Border = True Then OffsetRect dstRect, 1, 1
        Loop
    End If
    
    '////////// Post-Paint
    If (m_RenderFlags And render_PrePost) Then              ' allow user to paint over result (ideal for watermarking)
        With bkgRect                                        ' transfer result to usercontrol
            .nHeight = Abs(.nHeight)
            .nWidth = Abs(.nWidth)
            RaiseEvent PostPaint((tDC), (.nLeft), (.nTop), (.nWidth), (.nHeight), (m_HitRegion))
        End With
    End If
    
    If hOldRgn Then
        SelectClipRgn UserControl.hDC, hOldRgn
        DeleteObject hOldRgn
    Else
        SelectClipRgn tDC, 0&
    End If
    If hPen Then                                            ' drawing a border
        hBrush = SelectObject(tDC, GetStockObject(NULL_BRUSH))
        hPen = SelectObject(tDC, hPen)
        Rectangle tDC, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
        DeleteObject SelectObject(tDC, hPen)
        SelectObject tDC, hBrush
    ElseIf hRgn Then   ' create roundrect path & draw it
        ' if hRgn was created, it was already destroyed after selecting into the target DC previously
         sptDrawRoundRect tDC
    End If
    
    If (m_RenderFlags And render_PrePost) Then              ' allow user to paint over result (ideal for watermarking)
        With bkgRect
            BitBlt UserControl.hDC, .nLeft, .nTop, .nWidth, .nHeight, tDC, .nLeft, .nTop, vbSrcCopy
        End With
        m_DC.UpdateDC True                                  ' remove the bitmap from the DC
    End If
    
    If m_Size.DestDC <> 0& And (m_RenderFlags And render_RemoteDrawBkg) = render_RemoteDrawBkg Then
        If (m_RenderFlags And render_RemoteToHGraphics) Then
            GdipGetDC m_Size.DestDC, rDC
        Else
            rDC = m_Size.DestDC
        End If
        BitBlt rDC, m_DragDrop.AutoDragPts.X, m_DragDrop.AutoDragPts.Y, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hDC, 0&, 0&, vbSrcCopy
        If (m_RenderFlags And render_RemoteToHGraphics) Then GdipReleaseDC m_Size.DestDC, rDC
    End If
    
    If (m_RenderFlags And render_InitAnimation) Then
        m_RenderFlags = m_RenderFlags Xor render_InitAnimation
        If (m_Flags And attr_AutoAnimate) Then
            If (m_CntFlags And cnt_Runtime) Then Me.Animate lvicAniCmdStart
        End If
    End If
    
End Sub

'//// event occurs whenever control is about to be created, except for the very fist time
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' cached m_Attributes contains following
    ' &H......7F > lightness percent (&H......8. > set if lightness is negative value)
    ' &H....FF.. > global transparency percent
    ' &H...F.... > mirror value
    ' &H..F..... > effect value
    ' &H7F...... > grayscale
    ' &H8....... > set if Invert = true
    ' after reading cached value, m_Attributes contains following
    ' &H......FF > effect setting
    ' &H....FF.. > transparency color mode
    ' &H7F...... > interpolation
    ' &H..FF.... > mirror setting

    Dim bData() As Byte, propValue As Long
    With PropBag
        m_Angle = .ReadProperty("Angle", 0!)                    ' get rotation value
        m_Attributes = .ReadProperty("Settings", 0&)            ' get lightness/global trans/mirror-effect/grayscale
            m_Effects.GrayScale = (m_Attributes And &H7F000000) \ &H1000000
            m_Effects.LightnessPct = (m_Attributes And &H7F)
            If (m_Attributes And &H80&) Then m_Effects.LightnessPct = -m_Effects.LightnessPct
            m_Effects.GlobalTransparencyPct = (m_Attributes And &HFF00&) \ &H100&
            m_Effects.Invert = CBool(m_Attributes And &H80000000)
            ' place effect in 1st byte, keep mirror in 3rd byte
            m_Attributes = (m_Attributes And &HF0000) Or (m_Attributes And &HF00000) \ &H100000
        propValue = .ReadProperty("Blend", 0&)      ' get blend color
            m_Effects.BlendColor = (propValue And &H80FFFFFF)
            m_Effects.BlendPct = (propValue And &H7F000000) \ &H1000000
        UserControl.ForeColor = .ReadProperty("Border", vbBlack) ' get border color
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace) ' get opaque back color
        UserControl.FillColor = .ReadProperty("Gradient", vbWhite)
        m_RenderFlags = .ReadProperty("Render", 0&)             ' get FastRedraw/Pre-Post events
        m_CntFlags = .ReadProperty("Frame", cnt_AlignCenter)    ' get AutoSize/Align/Opaque/Border values
        m_Flags = .ReadProperty("Attr", attr_MouseEvents)       ' get key control properties (several)
        ' most recent addition. The high 2 bytes will be the interpolation setting
        m_Attributes = m_Attributes Or (m_Flags And &H7F000000)
        m_Flags = (m_Flags And &HFFFFFF)
        propValue = .ReadProperty("Trans", 0&)
            m_Effects.TransparentColor = (propValue And &H80FFFFFF) ' and place mode in 2nd byte
            m_Attributes = m_Attributes Or ((propValue And &HF000000) \ &H1000000) * &H100&
            m_Effects.TransparentColorUsed = (Me.TransparentColorMode <> lvicNoTransparentColor)
        propValue = .ReadProperty("Offsets", 0&)                ' get XYOffsets
        m_Offsets.X = (propValue And &HFFF&)
        m_Offsets.Y = ((propValue And &HFFF0000) \ &H10000)
        If (propValue And &H8000&) Then m_Offsets.X = -m_Offsets.X
        If (propValue And &H80000000) Then m_Offsets.Y = -m_Offsets.Y
        
        bData() = .ReadProperty("BkgImage", bData())               ' get our background image
        Set m_BkgImage = modCommon.LoadImage(bData, True, , True)  ' load image
        If m_BkgImage.Handle = 0& Then
            Set m_BkgImage = Nothing
        Else
            propValue = PropBag.ReadProperty("BkgIndex", 65537)
            If propValue > &HFFFF& Then m_BkgImage.ImageGroup = (propValue \ &H10000) Else m_BkgImage.ImageGroup = 1&
            m_BkgImage.ImageIndex = (propValue And &HFFFF&)
        End If
        Erase bData()
        
        propValue = .ReadProperty("Tiles", 0&)                  ' get segmented tile sizes
        bData() = .ReadProperty("Image", bData())               ' get our image
        If Me.Aspect = lvicFixedSize Then                      ' get fixed size, if needed
            m_Size.FixedCx = .ReadProperty("FixedCx", 32&)
            m_Size.FixedCy = .ReadProperty("FixedCy", 32&)
        End If
        Set UserControl.MouseIcon = .ReadProperty("MouseIco", UserControl.MouseIcon)
        UserControl.MousePointer = .ReadProperty("MousePtr", vbDefault)
    End With
    If m_CntFlags < 0& Then                                     ' enable/disable control
        UserControl.Enabled = False
        m_CntFlags = (m_CntFlags Xor &H80000000)
    End If                                                      ' set OLEDrop ability
    If Me.OLEDropMode <> vbOLEDropNone Then UserControl.OLEDropMode = vbOLEDropManual
    
    m_CntFlags = m_CntFlags Or cnt_InitLoad
    Set m_Image = modCommon.LoadImage(bData, True, , True)                ' load image, segement & set index
    If propValue Then Me.SegmentImage (propValue And &HFF), (propValue And &HFF0000) \ &H10000
    propValue = PropBag.ReadProperty("Index", 65537)
    If propValue > &HFFFF& Then m_Image.ImageGroup = (propValue \ &H10000) Else m_Image.ImageGroup = 1&
    m_Image.ImageIndex = (propValue And &HFFFF&)
    m_bAffects() = PropBag.ReadProperty("Effects", m_bAffects)
    If g_TokenClass.Version > 1! Then
        m_Effects.ImportEffectsParameters m_bAffects
        Erase m_bAffects
    End If
    If (m_RenderFlags And render_PrePost) Then
        m_RenderFlags = m_RenderFlags Xor render_PrePost
        Me.WantPrePostEvents = True
    End If
    m_CntFlags = (m_CntFlags And Not cnt_InitLoad)
    
    propValue = PropBag.ReadProperty("FrameDur", 0&) ' get min animated frame duration
    Me.Animate2.DefaultMinimumDuration = (propValue And &H7FFF&)
    Me.Animate2.DefaultMaximumDuration = (propValue And &H7FFF0000) \ &H10000
    m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest
        
End Sub

'//// event occurs each time control is resized by user or via code
Private Sub UserControl_Resize()
    m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoResize Or render_DoReScale Or render_DoHitTest
    Call sptGetScaledSizes(False)
End Sub

'//// event occurs whenever the control is shown
Private Sub UserControl_Show()

    If (m_CntFlags And cnt_Runtime) = 0& Then Call sptSetUserMode
    m_RenderFlags = m_RenderFlags Or render_ShowEvent Or render_DoReScale
    Call sptGetScaledSizes(False)
    If Not m_Animator Is Nothing Then
        If m_Animator.AnimationState = lvicAniCmdPause Then m_Animator.ResumeAnimation
    End If

End Sub

'//// event occurs when control is about to be destroyed & a PropertyChanged event occured
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Dim bData() As Byte, propValue As Long, Size As RECTF
    
    ' before writinhg value, m_Attributes contains following
    ' &H......FF > effect setting
    ' &H....FF.. > transparency color mode
    ' &H..FF.... > mirror setting
    ' &H7F...... > interpolation quality
    ' after writing, cached value contains
    ' &H......7F > lightness percent (&H......8. > set if lightness is negative value)
    ' &H....FF.. > global transparency percent
    ' &H...F.... > mirror value
    ' &H..F..... > effect value
    ' &H7F...... > grayscale setting
    ' &H8....... > set if Invert = true
    
    If UserControl.Enabled = False Then m_CntFlags = m_CntFlags Or &H80000000
    With PropBag
        If m_Image.ExtractImageData(bData) Then .WriteProperty "Image", bData
        .WriteProperty "Angle", m_Angle, 0!
        .WriteProperty "Trans", m_Effects.TransparentColor Or Me.TransparentColorMode * &H1000000, 0& ' add mode to color
        .WriteProperty "Blend", m_Effects.BlendColor Or m_Effects.BlendPct * &H1000000, 0& ' write combined blend color/pct
        propValue = (m_Attributes And &HF0000) Or (m_Attributes And &HF) * &H100000    ' place effects/mirror in 3rd byte
        If m_Effects.LightnessPct < 0& Then propValue = propValue Or &H80               ' place lightness iun 1st byte
        propValue = propValue Or Abs(m_Effects.LightnessPct)
        propValue = propValue Or m_Effects.GlobalTransparencyPct * &H100&               ' place global trans in 2nd byte
        propValue = propValue Or m_Effects.GrayScale * &H1000000                        ' place grayscale in 4th byte
        If m_Effects.Invert Then propValue = propValue Or &H80000000                    ' set high bit if Invert=True
        .WriteProperty "Settings", propValue, 0&
        
        .WriteProperty "Render", (m_RenderFlags And (render_Shown - 1&)), 0&
        .WriteProperty "Frame", (m_CntFlags And cnt_KeyProps), cnt_AlignCenter
        .WriteProperty "Border", UserControl.ForeColor, vbBlack
        .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
        .WriteProperty "Gradient", UserControl.FillColor, vbWhite
        ' most recent addition. Add Interpolation to this setting; don't want to create a new propbag value & break compatibility
        .WriteProperty "Attr", ((m_Flags And attr_KeyProps) Or (m_Attributes And &H7F000000)), attr_MouseEvents
        If m_Animator Is Nothing Then
            propValue = 0&
        Else
            propValue = Me.Animate2.DefaultMinimumDuration Or (Me.Animate2.DefaultMaximumDuration) * &H10000
        End If
        .WriteProperty "FrameDur", propValue, 0&
        If Not UserControl.MouseIcon Is Nothing Then
            If UserControl.MouseIcon <> 0& Then .WriteProperty "MouseIco", UserControl.MouseIcon
        End If
        .WriteProperty "MousePtr", UserControl.MousePointer, vbDefault
        If (Me.Aspect = lvicFixedSize) Then
            .WriteProperty "FixedCx", m_Size.FixedCx, 32&
            .WriteProperty "FixedCy", m_Size.FixedCy, 32&
        End If
        propValue = Abs(m_Offsets.X) Or Abs(m_Offsets.Y) * &H10000
        If m_Offsets.X < 0& Then propValue = propValue Or &H8000&
        If m_Offsets.Y < 0& Then propValue = propValue Or &H80000000
        .WriteProperty "Offsets", propValue, 0&
        .WriteProperty "Index", m_Image.ImageIndex Or m_Image.ImageGroup * &H10000, 65537
        If m_Image.Segmented Then
            propValue = GdipGetImageBounds(m_Image.Handle, Size, UnitPixel)
            propValue = CLng(Size.nHeight) \ m_Image.Height Or (CLng(Size.nWidth) \ m_Image.Width) * &H10000
            .WriteProperty "Tiles", propValue, 0&
        End If
        If g_TokenClass.Token > 1! Then
            m_Effects.ExportEffectsParameters m_bAffects()
            .WriteProperty "Effects", m_bAffects()
            Erase m_bAffects
        Else
            .WriteProperty "Effects", m_bAffects()
        End If
        If Not m_BkgImage Is Nothing Then
            If m_BkgImage.ExtractImageData(bData) Then
                .WriteProperty "BkgImage", bData
                .WriteProperty "BkgIndex", m_BkgImage.ImageIndex Or m_BkgImage.ImageGroup * &H10000, 65537
            End If
        End If
    End With
    If UserControl.Enabled = False Then m_CntFlags = m_CntFlags Xor &H80000000
    
End Sub

'//// event occurs when control is just about to be destroyed
Private Sub UserControl_Terminate()
    Set m_Animator = Nothing
    If Not m_MouseTracker Is Nothing Then
        m_MouseTracker.ReleaseMouseCapture False, ObjPtr(Me)
        Set m_MouseTracker = Nothing
    End If
    Set m_Effects = Nothing
    Set m_BkgImage = Nothing
    Set m_Image = Nothing
    Set m_DC = Nothing
    If m_HitRegion Then DeleteObject m_HitRegion
End Sub


'/////////////// INTERFACE CALLBACKS \\\\\\\\\\\\\\\\

'//// event occurs as a result of a timer advancing a multi-image format to a new index
Private Sub m_Animator_AnimationFinished()
    ' callback event from the cAnimator class
    RaiseEvent AnimationLoopsFinished
End Sub
Private Sub m_Animator_FrameChanged(Index As Long)
    RaiseEvent AnimationFrameChanged(Index)
End Sub
Private Sub m_Animator_Looped(Count As Long)
    If Count = 0& Then RaiseEvent AnimationLoopsFinished
End Sub

'//// event occurs when multi-image format index changes
Private Sub m_Image_ImageChanged(FrameIndex As Long)
    ' callback event from the GDIpImage class
    Dim lValue As Long
    m_RenderFlags = m_RenderFlags Or render_DoFastRedraw Or render_DoReScale Or render_DoHitTest
    m_Size.Width = 0&
    Call sptGetScaledSizes(False)
    Me.TransparentColorMode = -Me.TransparentColorMode
    If Me.SetRedraw = True Then UserControl.Refresh
End Sub

'//// event occurs when timer triggers tracking mechanism to test for mouse leaving the control
'  See HitTest property & MouseEnter/MouseExit events
Private Sub m_MouseTracker_AnticpateMessage(SetValidation As Boolean, Validation As Boolean)

    If SetValidation Then                           ' allow event to continue?
        If (m_CntFlags And cnt_MouseDown) Then      ' not if mouse button down; no point
            Validation = False
        Else                                        ' otherwise; set local flag for upcoming event
            m_CntFlags = m_CntFlags Or cnt_MouseValidate
            Validation = True                       ' flag is processed in Mouse_Move
        End If
    Else                                            ' event came & passed; did we get it?
        Validation = ((m_CntFlags And cnt_MouseValidate) = 0&)  ' return true or false
        m_CntFlags = m_CntFlags And Not cnt_MouseValidate       ' reset flags
    End If

End Sub

'//// event occurs when timer-driven event indicates mouse left control.
'  See HitTest property & MouseEnter/MouseExit events
Private Sub m_MouseTracker_MouseExited()
    ' MouseExit tracker indicates no longer over control/image
    Set m_MouseTracker = Nothing        ' release tracker & raise event
    m_CntFlags = (m_CntFlags And Not (cnt_DblClicked Or cnt_MouseValidate))
    If (m_Flags And attr_MouseEvents) Then RaiseEvent MouseExit
End Sub


' /////////////////////////////////////////////////////////////////////////////
'                            CHANGE HISTORY
' /////////////////////////////////////////////////////////////////////////////
' 15 Jan 2012, v2.1.32 :: Last update barring major bugs
' - UnicodeBrowseFolders now can be called from within your project
' - Property page could return some 32x32 & 48x48 icons in PNG format & shouldn't. Fixed
' 8 Jan 2012, v2.1.31
' - property page patched to have dragged/pasted files related as associated icons, as appropriate
' 7 Jan 2012, v2.1.30
' - can load icons associated with file system objects
' - modified property page for associated icon selection
' - added ASSOCIATEDICON structure, SHILIconSizeEnum & AssocIconTypeEnum enumerations
' - added UnicodeBrowseFolders global class
' - updated LoadPictureGdip & SavePictureGDIp
' - fixes a logic error where control may not resize while it is hidden
' 26 Nov 2011, v2.1.29
' - calculation typo converting icon/cursor handle could foul up black/white sources. Fixed
' 16 Nov 2011, v2.1.28
' - Saving unmodified AVI to other destinations (i.e., file, array, etc) now provided
' - Found several minor bugs
'   -- cFunctionsICO.HICONtoArray custom header changes prevented loading icons by handle
'   -- cFunctionsPNM.LoadPNMResource invalid flag setting could invert PBM formats
'   -- cFunctionsPNM.LoadPNMResource saved invalid token when writing 32 bpp alpha PAM formats
'   -- cFunctionsTGA.SaveAsTGA could save image upside down
'   -- cFunctionsPCX.SaveAsPCX could save image upside down
'   -- property page did not offer AVI as an image source
' 7 Nov 2011, v2.1.27
' - test code from a few versions ago remained by mistake. Could crash project. Fixed
' 6 Nov 2011, v2.1.26
' - If reading 30 byte mp3 tag or smaller, array referencing was incorrect. Fixed
' 3 Nov 2011, v2.1.25
' - added WMA files as an image source
' - enabled control to be bound to a database table/field. See LoadPictureGDIp.RTF for more
' - added UpdateDataboundImage event to allow changing what will be saved before image written to database
' - rewrote the MP3 parsing logic; more robust & should properly handle v2,3,4 ID3 tags + unsynchronized tags
' 25 Oct 2011, v2.1.24
' - SaveImageAsDrawnToGDIpImage & PaintImageAsDrawnToHDC could fail if source control is hidden or not viewable. Fixed
' - cFunctionMP3 had error that could corrupt created TIFF file. Fixed (want to rewrite that class for proper parsing)
' - Added RegionFromImage function in GDIpImage class
' - Added version information in the About Box
' 21 Oct 2011, v2.1.23
' - Added reading/writing of pbm, pgm, ppm, pam image formats
' - Added routines for reading AVIs, no save to AVI functionality
' - 32bpp bitmaps extracted from binaries didn't honor alpha channel. Fixed
' - PNG icon format from binaries not processed correctly at times. Fixed
' - Added following classes: cFunctionsDLL, cFunctionsPNM, cFunctionsAVI
' 5 Oct 2011, v2.1.22
' - Added BkgImage and BkgImageStretch properties
' - Added lvicFixedSizeStretched to the ScalingRatioEnum enumeration
' - Added AsyncDownloadDoneBkgImg event to support the new BkgImage property
' - Added SetFixedSizeAspect method to allow run-time adjustment of FixedSize aspect
' - Modified property page to allow design-time assignment of background image
' - Rewrote the sptFillGradient method to use GDI+ vs. msimg32.dll
' 1 Oct 2011, v2.1.21
' - Added gradient background options
' - Added SetClipRect function
' - Added option to choose blended or non-blended rounded corner borders
' 23 Sep 2011, v2.1.20
' - Added new BorderShape property to allow borders with rounded corners
' - Added ability to drag/drop, copy/paste from zip files accessed in Explorer as compressed folders
' - Tweaked dropped/pasted files to process all files if multiple files dropped/pasted, as needed, until image successfully returned
' 21 Sep 2011, v2.1.19
' - Last update broke TransparencyPct functionality. Fixed
' 20 Sep 2011, v2.1.18
' - When inverting colors while blending or adding/subtracting lightness, improper colors returned. Fixed
' 19 Sep 2011, v2.1.17
'- Fixes rounding error that could result in metafile rendered clipped 1 pixel in width or height
'- Added Black and White as a grayscale option
' 9 Sep 2011, v2.1.16
'- Calculation error when auto-sizing on fractions of degrees. Fixed
'- Mirrored vertical/horizontal & rotation could result in failed rendering. Fixed I believe
'- Global async support functions added: AsyncAbortDownloads, AsyncGetDownloadStates
'- Added AsyncDownloadURL property to the GDIpImage class
' 6 Sep 2011, v2.1.15
'- Tweaking async routines for more reliability. Don't expect this to be last async-related tweak.
' 5 Sep 2011, v2.1.14
'- Added async download capability. See LoadPictureGDIp.rtf & LoadPictureGDIplus function for more info
'- Tweaked icon parsing routines; some black & white icons were not parsed correctly
'- Added 2 new classes: cAsyncController & cAsyncClient
'- Added new global property: AsyncDownloadsEnabled
' 19 Aug 2011, v2.1.13
'- Tweaked rendering routines to prevent nearest-neighbor stretching from appearing offset from destination X,Y coords
'- SaveControlAsDrawnToGDIpImage tweaked to handle cases where control is user-drawn via PrePost Paint events
' 14 Aug 2011, v2.1.12
'- Creates animated PNG
'- Class-based Animate2 method supersedes depricated Animate method
'- LoadPictureGDIplus(Screen) now creates screen capture
' 13 Jul 2011, v2.1.11
'- Found array indexing error in pvConvert24bppIcon of cFunctionsICO that could corrupt icon mask. Fixed
' 12 Jul 2011, v2.1.10
'- Fixed couple minor errors in ICO & GIF creation routines
'- Added AlphaMask & RenderSkewed functions in GDIpImage class
'- Added ability to load images from Base64 encoded strings
'- prjTheBasics sample project updated for Base64 loading example
'- prjZoomPan sample project updated for AlphaMask & RenderSkewed examples
' 12 Jun 2011, v2.1.9
'- Interpolation property added for scaled/rotated images. Default is highest quality
' 2 Jun 2011, v2.1.8
'- Extracted icons/cursors from executables may not have been provided within the property page. Fixed
'- Now uses header parsing to determine if file is executable vs. the file extension
' 1 Jun 2011, v2.1.7
'- New MULTIIMAGESAVESTRUCT structure added to support saving in animated GIF, multi-image icon/cursor & multi-page TIFF formats
'- Added ScaleWidth & ScaleHeight properties. Values returned in pixels less border width, if used
'- Tweaked some minor stuff for optimization/correctness
' 14 Apr 2011, v2.1.6
'- Segmenting an  image in property page generated error. Fixed
'- After segmenting image in property page, effect not shown in design-time control. Fixed
' 11 Mar 2011, v2.1.5
'- Animation re-worked
' 4 Mar 2011, v2.1.4
'- vbCFText & CF_UNICODE supported when loading from Data object & Clipboard
'- Unicode OpenFile dialog class had error which could return extra characters in file/path name. Fixed
'- Control now entering maintenance mode; don't foresee any additional enhancements at this time.
' 27 Feb 2011, v2.1.3
'- Icons were not always saved correctly when saving to a stdPicture. Fixed
'- SavePictureGDIplus can now create blank GDI+ bitmaps within a GDIpImage class/object
' 21 Feb 2011, v2.1.0
'- Converting to B&W in some formats did not really work. Fixed
'- Setting animation commands in form_load may be ignored. Fixed
'- Couple of settings not written to & read from the propertybag correctly. Fixed
'- Added ImageGroupFormat property to GDIpImage class in support of images loaded from binaries
'- Added 4 new control methods to save/print either the entire control or its image to a GDIpImage or any hDC:
':: SaveControlAsDrawnToGDIpImage , SaveImageAsDrawnToGDIpImage
':: PaintControlAsDrawnToHDC , PaintImageAsDrawnToHDC
'- Can now load/extract bitmaps, icons, cursors from exe, dll or ocx
' 12 Feb 2011, v2.0.0. [B]Binary compatibility lost[/B]
'- Massive rewrite to support cross-converting image formats
'- Added new cColorReduction class to support color depths relative to specific image formats
'- Added ability to load images from URL
'- Added new sample project: prjSaveAs. Updated all other sample projects
'- Tweaked all routines to increase speed/efficiency
' 2 Jan 2011
'- Added PCX format reading and writing. New class added: cFunctionsPCX
'- Added WMF/EMF writing
'- PCX routines wouldn't save to 8bpp if GDI+ reported transparency in use. Fixed
' 11 Dec 2010
'- Added MousePointer and MouseIcon properties identical to most VB controls
' 30 Nov 2010
'- Now shares responsibility for rendering GIFs with GDI+ in order to avoid GDI+ bugs related to animatd GIFs
'- Now supports APNG (animated PNG) formats
'- Added 2 new classes to support the above: cFunctionsGIF & cFunctionsPNG
' 27 Nov 2010
'- MouseExit event did not trigger in specific cases when mouse entered an overlapped windowless-control placed above this control's ZOrder. Fixed
'- If unloading form in control's Click event, error could be generated in some cases. Fixed I believe
' 7 Nov 2010
'- Restricted control size if extremely large images loaded. Max size for the control will be approximately screen size
'- Property name changes: Stretch is now Aspect.  AutoAnimate is now AnimateOnLoad
' 30 Oct 2010
'- Added global TilePictureGDIplus routine
'- Added UsesTransparency property to GDIpImage class
' 27 Oct 2010
'- Added inverted color (photo negative) option. Resulted in several methods modified
'- Added Sepia as a grayscale constant. I know it is not grayscale, more like brown-scale
' 24 Oct 2010
'- WantMouseEvents when set to False now allows click & double click events
'- Added optional Sequence parameter to GDIpImage.SegmentImage routine
' 20 Oct 2010
'- Fixed AutoSize effect where rotated images can get progessively smaller.
'- Applied workaround to GDI+ effects not painting correctly in all cases
'- Added 2 new sample projects: The Basics and Animation showing segmented images
' 17 Oct 2010
'- Added new Global Class: GDIpEffects which enables v1.1 Effects like Blur, Color Corrections, and more.
':: Class also exposes the image attributes properties used in the control (GrayScale, BlendColor, etc)
'- Modified all rendering/saving routines to accept a v1.1 Effects handle
'- Added property page to enable design-time creation of v1.1 Effects
'- Added Global function: GDIplusTokenVersion
' 14 Oct 2010
'- Many icon related functions revamped to enable grouped animated cursors found with Win7
'- Now extracts default "Jiffies" if animated cursor format provides it. More accurate animation of cursors
'- Added ImageGroup and ImageGroupCount properties to control for Win7+ grouped animated cursor support
' 12 Oct 2010: BETA. Do expect some bugs here & there
' /////////////////////////////////////////////////////////////////////////////


