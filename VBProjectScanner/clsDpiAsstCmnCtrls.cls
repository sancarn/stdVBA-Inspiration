VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDpiAsstCmnCtrls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' //// COMMON CONTROLS DPI ADJUSTMENTS
' All known common controls (cc) are handled in this class
'   -- cc v5, cc v6, cc-2 v5, cc-2 v6, cc-3

' This class must be declared in the form's declaration section
' If you are using ImageList controls then you also should Implement this class
'   Dim CmnCtrlsDpiAssist As clsDpiAsstCmnCtrls
'   Implements clsDpiAsstCmnCtrls ' only if image lists are used

' Coder notes: Because this class can be implemented, only implemented
'   methods are declared PUBLIC and are prefixed with keyword: Callback.
'   Other exposed methods are declared FRIEND

' ------------------------------------------
' Toolbars...
' ------------------------------------------
' The font on v5 of the toolbar can be per-monitor aware and only in
'   Windows 10 version 1703 or better. If you cannot guarantee that your
'   project will run on those operating systems, your toolbar font will not
'   scale with a per-monitor aware application. Recommend not using captions.

' If a toolbar is used that contains other controls, then the following should
'   be done to help ensure those controls are positioned correctly after
'   the toolbar scales due to DPI changes.
'   1. Always create a placeholder button for any child control you are using
'   -- the width of the placeholder should be the same as the control's width
'   2. In Form_Load, always align the control with its placeholder, i.e.,
'       Combo1.Move Toolbar1.Buttons("cboPlaceHolder").Left, Toolbar1.Buttons("cboPlaceHolder").Top
'   3. During the IDpiPmAssistant_ScaleControlVB event for the control on the
'       toolbar, adjust the newX parameter similar as you did in Form_Load:
'           newX = Toolbar1.Buttons("cboPlaceHolder").Left
'           newY = Toolbar1.Buttons("cboPlaceHolder").Top
'       FYI: The toolbar is scaled before any of its child controls
'       Note: The toolbar button positions may be wrong when returned directly
'       from the control. You should use this class method: GetToolbarBtnPos

' ------------------------------------------
' Coolbars... are a pain
' ------------------------------------------
' The font for Coolbars can only be per-monitor aware in Windows 10 version 1703
'   or better. If you cannot guarantee that your project will run on those operating
'   systems, your coolbar font will not scale with a per-monitor aware application.
'   Recommend not using captions.
' If you are adding controls to any of the Bands in the Coolbar, ensure you set
'   the Child property of that band during design-time. This is the only way to
'   help guarantee that the control will be positioned correctly after the
'   Coolbar rescales.
' If these contain image lists, you must reassign the property during
'   Form_Load before that form's clsDpiPmAssist Activate method is called,
'   i.e., Set Coolbar1.ImageList = ImageList1
'   Otherwise, the ImageList property can return Nothing until Form_Load exits.

' Coolbar resizes its child controls via APIs. When a child control is an OCX, then
'   VB will not be aware of its actual size. You should not base any calculations
'   on dimension properties of Coolbar child OCX controls. Instead, as needed, get
'   their dimensions from APIs like GetWindowRect.
' When a Coolbar has its dimensions changed, afterwards you should call clsDpiPmAssist class
'   method SyncOcxToParent passing its last param as True. That method is called
'   herein during scaling due to DPI changes only.

' ------------------------------------------
' ListViews...
' ------------------------------------------
' v6 of the common controls must not use the .Bold property of any ListItem or
'   ListSubItem. If .Bold = True, then the font becomes static throughout subsequent
'   scaling due to DPI changes -- font size will not change.
' Checkbox size is static, drawn at system DPI only. Recommend not using ths style

' ------------------------------------------
' TreeViews...
' ------------------------------------------
' Checkbox size is static, drawn at system DPI only. Recommend not using ths style

' ------------------------------------------
' DatePickers...
' ------------------------------------------
' The entire non-client area is custom drawn by OCX and never changes scale
' If calendar is displayed at time of DPI change, it does not get scaled

' ------------------------------------------
' Sliders...
' ------------------------------------------
' v6 of the common controls has its thumb and font scaled to system DPI and they
'   do not change size. Both versions auto-size themselves if their maximum dimension
'   is exceeded. If the host must report the correct dimensions after scaling, call
'   the following method after all scaling: clsDpiPmAssist.SyncParentToOcx.

' ------------------------------------------
' ImageCombo...
' ------------------------------------------
' The starting width of the ImageCombo should always be set in Form_Load before
'   any scaling actions occur. When this control is created and non-integral DPI
'   applies, the control will be smaller than expected otherwise.
' The Indentation properties are not scaled. Those are Integer values. Since
'   integers are rounded when scaled, they should not be used for future scaling.
'   To ensure those properties are scaled correctly, if used, you should hardcode
'   the Indentation properties and scale from that vs. the actual property value.

' ----------------------------------------------------------------
' ImageLists... These cannot be scaled directly within this class
' ----------------------------------------------------------------
' Only share imagelist between controls where both either scale or both do not scale.
' When DPI changes, this class will send an event to scale the image list.
'   but you need to implement this class in order to receive the event(s).
' When CallbackScaleImageList is sent to your form, you provide the new image list
'   to use for the current DPI. See that method, at bottom of class, for sample usage

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef pRect As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindowExW Lib "user32.dll" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Long, ByVal lpsz2 As Long) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemMetricsForDpi Lib "user32.dll" (ByVal nIndex As Long, ByVal lDPI As Long) As Long

Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (clsid As Any, progid As Any) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Any) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Enum CCTLEnum ' see pvIsCommonControl
    cctlNone = 0
    cctlToolbar5
    cctlToolbar6
    cctlSlider5
    cctlSlider6
    cctlListView5
    cctlListView6
    cctlTreeView5
    cctlTreeview6
    cctlStatusbar5
    cctlStatusbar6
    cctlProgressbar5
    cctlProgressbar6
    cctlTabStrip5
    cctlTabStrip6
    cctlImageCombo
    cctlUpDown5
    cctlUpDown6
    cctlAnimation5
    cctlAnimation6
    cctlDtPicker
    cctlFlatScroll
    cctlMonthView
    cctlCoolbar
End Enum

Const SWP_NOMOVE = &H2&
Const SWP_NOZORDER_REPOSITION = &H204&

Dim oSubclasser As Object           ' used only to help scale ImageCombo in one scenario
Dim m_CanScale As Boolean           ' DpiScalingCycle called
Dim m_Host As Long                  ' form soft reference (ObjPtr), -1, or 0
Dim m_Assistant As clsDpiPmAssist   ' form's clsDpiPmAssist instance
Dim m_Cbars As VBA.Collection           ' coolbars bar references for post-scale sizing
    ' Coolbar bands are added to this collection
    ' After all band child controls are scaled, the band is resized. This is in
    '   opposite order of normal scaling which is container first, children next.

Dim m_ImageLists As VBA.Collection      ' created/destroyed per scaling cycle
    ' each entry is Array(SrcImgLst, ReplacementImgLst)
Dim m_DynamicLists As VBA.Collection    ' destroyed on class termination
    ' each entry is Array(clonedImgList, CloneSrcName)

Friend Sub Attach(Host As Object, oDpiPmAssist As clsDpiPmAssist)
    ' The form calling this method must also implement IDpiPmAssistant
    
    ' Add this class in the declarations section and implement this class
    '   Implements clsDpiAsstCmnCtrls ' only if ImageLists are used
    '   Dim CmnCtrlsDpiAssist As clsDpiAsstCmnCtrls ' always include this
    
    ' in the implemented callback from IDpiPmAssistant, call this method:
    ' Private Sub IDpiPmAssistant_Attach(oDpiPmAssist As clsDpiPmAssist)
    '   ...
    '   Set CmnCtrlsDpiAssist = New clsDpiAsstCmnCtrls
    '   CmnCtrlsDpiAssist.Attach Me, oDpiPmAssist
    ' End Sub
    
    Dim oAssist As IDpiPmAssistant
    If m_Host = 0 Then              ' else already called
        If Not Host Is Nothing Then ' sanity checks
            If Not oDpiPmAssist Is Nothing Then
                If TypeOf Host Is VB.Form Then
                    On Error Resume Next
                    Set oAssist = Host ' host must implement that too
                    If Err Then
                        Err.Clear
                    ElseIf Not oAssist Is Nothing Then
                        If TypeOf Host Is clsDpiAsstCmnCtrls Then
                            m_Host = ObjPtr(Host)
                        Else
                            m_Host = -1 ' host does not implement this class (not required)
                        End If
                        Set m_Assistant = oDpiPmAssist
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    End If
End Sub

Friend Sub DpiScalingCycle(ByVal Reason As DpiScaleReason, _
                           ByVal Action As DpiScaleCycleEnum, _
                           ByVal OldDPI As Long, ByVal NewDPI As Long)
    
    ' You must call this method, otherwise, no scaling will be done
    ' In the implemented callback from IDpiPmAssistant, call this method, i.e.
    
    ' Private Function IDpiPmAssistant_DpiScalingCycle(ByVal Reason As DpiScaleReason, ByVal Action As DpiScaleCycleEnum, ByVal OldDPI As Long, ByVal NewDPI As Long, userParams As Variant) As Long
    '
    '   thisClass.DpiScalingCycle Reason, Action, OldDPI, NewDPI
    '   ...
    '
    ' End Function
    
    If m_Host = 0 Then Exit Sub     ' this class Attach method not called
    If Action = dpiAsst_BeginCycleControls Then
        m_CanScale = True           ' else no scaling occurs
        
    ElseIf Action = dpiAsst_EndCycleControls Then
        If m_CanScale = False Then Exit Sub
        
        Dim frm As VB.Form, Index As Long, c As Control
        
        m_CanScale = False          ' reset
        If Not m_Cbars Is Nothing Then
            For Each c In m_Cbars
                pvDoCoolbar c, NewDPI / OldDPI, 0!, 0!, 0!, 0!, True
            Next
            Set m_Cbars = Nothing
            Set c = Nothing
        End If
        ' remove unused, dyanmically created imagelists, as needed
        If Not m_ImageLists Is Nothing Then
            If Not m_DynamicLists Is Nothing Then
                On Error Resume Next
                Set frm = pvGetHost()
                For Index = 1 To m_ImageLists.Count
                    Set c = m_ImageLists(Index)(0)
                    Set c = m_DynamicLists(CStr(ObjPtr(c)))
                    If Err Then
                        Err.Clear   ' Imagelist isn't one we created or still in use?
                    Else
                        frm.Controls.Remove c
                        If Err Then
                            Err.Clear ' still in use, client didn't pass all controls
                        Else          ' to this class that use that imagelist. Shame!
                            m_DynamicLists.Remove CStr(ObjPtr(c))
                            If m_DynamicLists.Count = 0 Then
                                Set m_DynamicLists = Nothing
                                Exit For
                            End If
                        End If
                    End If
                Next
                Set frm = Nothing: Set c = Nothing
                On Error GoTo 0
            End If
            Set m_ImageLists = Nothing  ' reset
        End If
    End If
End Sub

Friend Function ScaleControlOCX(theControl As Control, ByVal Reason As DpiScaleReason, _
                            ByVal Action As DpiActionEnum, ByVal ScaleRatio As Single, _
                            newX As Single, newY As Single, newCx As Single, newCy As Single) As Long
                            
    ' method returns non-zero if control was scaled by this class
   
    ' in the implemented callback from IDpiPmAssistant, call this method, i.e.
    
    ' Private Function IDpiPmAssistant_ScaleControlOCX(theControl As Control, ByVal Reason As DpiScaleReason, ByVal Action As DpiActionEnum, ByVal ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single, fontProperties As String, userParams As Variant) As Long
    '
    '   If Action = dpiAsst_BeginEvent Then
    '        IDpiPmAssistant_ScaleControlOCX = thisClass.ScaleControlOCX(theControl, Reason, Action, ScaleRatio, newX, newY, newCx, newCy)
    '       ...
    '   Else      ' dpiAsst_EndEvent
    '       ...
    '   End If
    '
    ' End Function
    
    ' Routine handles passed control in one of 2 ways:
    '   1. is not a common controls object
    '       control is not scaled, routine returns zero to flag as not scaled
    '   2. is a common controls object
    '      return value is non-zero (1) and control is scaled
    
    
    If Action = dpiAsst_BeginEvent Then
        If m_CanScale = False Then Exit Function ' DpiScalingCycle method not called
        
        Dim lType As CCTLEnum
        lType = pvIsCommonControl(theControl)
        
        Select Case lType
        Case cctlNone: Exit Function            ' not a common control object
        Case cctlToolbar5, cctlToolbar6         ' tool bars
            pvDoToolbar lType, theControl, ScaleRatio, newX, newY, newCx, newCy, Reason = dpiAsst_InitialLoad
        Case cctlListView5, cctlListView6       ' list views
            pvDoListView lType, theControl, ScaleRatio, newX, newY, newCx, newCy
        Case cctlTreeView5, cctlTreeview6       ' tree views
            pvDoTreeView lType, theControl, ScaleRatio, newX, newY, newCx, newCy
        Case cctlStatusbar5, cctlStatusbar6     ' status bars
            pvDoStatusBar lType, theControl, ScaleRatio, newX, newY, newCx, newCy
        Case cctlTabStrip5, cctlTabStrip6       ' tab strips
            pvDoTabStrip lType, theControl, ScaleRatio, newX, newY, newCx, newCy
        Case cctlImageCombo                     ' image combobox
            pvDoImageCombo theControl, ScaleRatio, newX, newY, newCx, newCy, Reason = dpiAsst_InitialLoad
        Case cctlCoolbar                        ' cool bars
            pvDoCoolbar theControl, ScaleRatio, newX, newY, newCx, newCy, False
        Case cctlMonthView, cctlDtPicker ' controls needing only size/fonts to be scaled
            Set theControl.Font = m_Assistant.ScaleStdFont(theControl.Font, m_Assistant.DpiForForm)
            theControl.Move newX, newY, newCx, newCy
            m_Assistant.SyncOcxToParent theControl
        Case Else ' controls w/o fonts & no special scaling actions needed
            ' ProgressBar, UpDown, Animation, FlatScrollBar
            theControl.Move newX, newY, newCx, newCy
            m_Assistant.SyncOcxToParent theControl
        End Select
        ScaleControlOCX = 1
    End If
    
End Function

Friend Function CloneImageList(SrcImageList As Control, ByVal Width As Integer, ByVal Height As Integer) As Control

    If SrcImageList Is Nothing Then Exit Function
    If Width = 0 Or Height = 0 Then Exit Function
    If m_Host = -1 Or m_Host = 0 Then Exit Function
    If TypeName(SrcImageList) <> "ImageList" Then Exit Function
    
    ' routine clones SrcImageList using passed Width/Height params and
    '   returns a dynamically loaded ImageList as that clone.
    ' If Width and/or Height are passed as negative values, then a dynamically loaded
    '   ImageList control will be returned with all properties cloned except pictures/keys.
    '   In this case, you are asking for an empty ImageList control that you
    '   will populate the pictures/keys. The ImageWidth & ImageHeight property values will
    '   be the absolute values of passed parameters.
    
    ' This method can be called from this class' implemented CallbackScaleImageList
    ' method. See that method at bottom of class for sample usage.
    
    ' If this routine is simply copying the pictures from the passed SrcImageList,
    '   then this produces least desirable results. The images that will be
    '   contained in the clone are not scaled beforehand. This means that
    '   the same image used in SrcImageList are simply stretched to the width
    '   and height that the clone will be set to when rendered. Scaling is lowest quality.
    ' A better solution is to supply your own scaled images and not use this routine
    '   to do so. This requires your application to have multiple sizes of each
    '   picture used in an ImageList or use higher quality scaling techniques.
    
    Dim Index As Long, c As Control
    Dim sName As String, frm As VB.Form
    
    sName = pvGetProgID(SrcImageList)       ' get ProgID for passed ImageList version
    If LenB(sName) = 0 Then Exit Function
    
    On Error Resume Next                    ' dynamically load new ImageList
    Set frm = pvGetHost()
    Set c = frm.Controls.Add(sName, "IL" & CLng(Timer) & ObjPtr(SrcImageList))
    If Err Then Exit Function
    
    On Error GoTo errHandler
    With SrcImageList                       ' set clone properties
        c.MaskColor = .MaskColor
        c.BackColor = .BackColor
        c.UseMaskColor = .UseMaskColor
        c.Tag = .Tag                        ' and copy source imagelist pictures
        If Width < 1 Or Height < 1 Then
            c.ImageWidth = Abs(Width): c.ImageHeight = Abs(Height)
        Else
            c.ImageWidth = Width: c.ImageHeight = Height
            For Index = 1 To .ListImages.Count
                With .ListImages(Index)
                    c.ListImages.Add(, .Key, .Picture).Tag = .Tag
                End With
            Next
        End If
    End With
    On Error GoTo 0                         ' track created clone
    If m_DynamicLists Is Nothing Then Set m_DynamicLists = New Collection
    m_DynamicLists.Add Array(c, SrcImageList.Name), CStr(ObjPtr(c))
    Set CloneImageList = c
    Set c = Nothing
    Exit Function

errHandler:
    frm.Controls.Remove c
    Set c = Nothing
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Friend Function GetToolbarBtnPos(cToolbar As Control, ByVal ButtonIndex As Long, _
                            Left As Single, Top As Single, ByVal asTwips As Boolean) As Boolean
    ' courtesy function, not used in this class
    
    ' When controls are placed inside a toolbar, to help ensure they are aligned
    '   properly, you should be aligning them to the top/left of a toolbar placeholder
    '   button. That button's top/left position can be incorrectly reported by VB
    '   when the toolbar is scaled due to DPI.
    
    ' method returns the actual Left & Top position values of the passed button
    ' index as either pixel or twip values
    
    ' cToolbar. A common controls toolbar control instance
    ' ButtonIndex. The one-based button index to retrieve position
    ' Left, Top. Position is returned in these parameters
    ' asTwips. If true, Left/Top are twips else pixels
    
    If ButtonIndex < 1 Then Exit Function
    If cToolbar Is Nothing Then Exit Function
    
    Dim cWnd As Long, aData(0 To 4) As Long ' faux TBBUTTON structure (20 bytes)
    Const WND_CLASS5 = "ToolbarWindow32"
    Const WND_CLASS6 = "msvb_lib_toolbar"
    Const WM_USER As Long = &H400
    Const TB_GETRECT As Long = WM_USER + 51
    Const TB_GETBUTTON As Long = WM_USER + 23
    
    On Error Resume Next
    If pvIsCommonControl(cToolbar) = cctlToolbar5 Then
        cWnd = FindWindowExW(cToolbar.hWnd, 0, StrPtr(WND_CLASS5), 0)
    Else
        cWnd = FindWindowExW(cToolbar.hWnd, 0, StrPtr(WND_CLASS6), 0)
    End If
    If Err Then
        Err.Clear           ' error only possible if v5Toolbar is a windowless control
        On Error GoTo 0
    ElseIf cWnd <> 0 Then   ' can be zero only if passed an invalid control
        On Error GoTo 0     ' if any SendMessage fails, then passed an invalid control
        If SendMessageA(cWnd, TB_GETBUTTON, ButtonIndex - 1, VarPtr(aData(0))) <> 0 Then
            If SendMessageA(cWnd, TB_GETRECT, aData(1), VarPtr(aData(0))) <> 0 Then
                If asTwips = True Then
                    Left = aData(0) * Screen.TwipsPerPixelX
                    Top = aData(1) * Screen.TwipsPerPixelY
                Else
                    Left = aData(0): Top = aData(1)
                End If
                GetToolbarBtnPos = True
            End If
        End If
    End If

End Function

Private Sub pvDoToolbar(lType As CCTLEnum, c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single, initLoad As Boolean)
                
    ' v6 of the toolbar font is not scalable
    ' v5 of the toolbar font is scalable as of Win10.1703

    Dim Index As Long, oIL As Control
    Dim lImages() As Variant
    Const tbrPlaceholder As Long = 4
    
    Set oIL = pvTrackImageList(c.ImageList, ScaleRatio)
    If Not oIL Is Nothing Then
        With c.Buttons  ' track assigned images
            If .Count <> 0 Then
                ReDim lImages(1 To .Count)
                For Index = 1 To .Count
                    lImages(Index) = .Item(Index).Image
                Next
            End If
        End With
        Set c.ImageList = oIL
    End If
    c.Move newX, newY, newCx, newCy
    m_Assistant.SyncOcxToParent c
    c.ButtonWidth = c.ButtonWidth * ScaleRatio
    For Index = 1 To c.Buttons.Count
        With c.Buttons(Index) ' scale Placeholders
            If .Style = tbrPlaceholder Then
                .Width = .Width * ScaleRatio
                If initLoad = True Then .Width = .Width * m_Assistant.NonIntegralDpiRatio(True)
            End If
            If Not oIL Is Nothing Then ' reassign images
                If IsEmpty(lImages(Index)) = False Then .Image = lImages(Index)
            End If
        End With
    Next
    Set oIL = Nothing: Erase lImages()
    
End Sub

Private Sub pvDoCoolbar(c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single, PostPos As Boolean)

    ' note: when swapping out image lists, the coolbar does not
    ' erase the image index assigned to the band, as long as the
    ' new image list has at least the same number of images
    
    ' coolbar font is not scalable in O/S less than Win10.1703
    ' note: coolbar can use v5 or v6 of the common control imagelist
    
    Dim Index As Long, oIL As Control

    If PostPos = False Then
        Set oIL = pvTrackImageList(c.ImageList, ScaleRatio)
        If Not oIL Is Nothing Then Set c.ImageList = oIL
        Set oIL = Nothing
        c.Move newX, newY, newCx, newCy
        If m_Cbars Is Nothing Then Set m_Cbars = New Collection
        m_Cbars.Add c
        For Index = 1 To c.Bands.Count
            If c.Bands(Index).Child Is Nothing Then
                ' when child is scaled later, it scales the band
                c.Bands(Index).Width = c.Bands(Index).Width * ScaleRatio
            End If
        Next
    Else
        For Index = 1 To c.Bands.Count
            With c.Bands(Index)
                .MinWidth = .MinWidth * ScaleRatio
                .MinHeight = .MinHeight * ScaleRatio
            End With
        Next
        m_Assistant.SyncOcxToParent c, , True
    End If
    
End Sub

Private Sub pvDoListView(lType As CCTLEnum, c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single)

    ' v6 of Common Controls isn't as DPI aware as we would like when using
    '   views other than Report. When the Arrange property is not AutoTop/Left,
    '   adding new items to the listview after DPI changes doesn't place that
    '   next item where you would have hoped.
    
    ' Do not set the .Bold property to True for any listview item or its subitems.
    '   Any item where .Bold = True, results in that item's font remaining static
    '   during DPI changes. The bolded font will not scale. This only affects the
    '   listitem/subitem .Bold property, it does not affect the
    '   ListView1.Font.Bold property
    
    ' Checkbox rendering is always at system DPI, it does not scale

    Dim lImages() As Variant
    Dim lIndex As Long, sIndex As Long
    Dim oIL(0 To 2) As Control
    Const lvwReport As Long = 3
    
    Set oIL(0) = pvTrackImageList(c.Icons, ScaleRatio)
    Set oIL(1) = pvTrackImageList(c.SmallIcons, ScaleRatio)
    If lType = cctlListView6 Then
        Set oIL(2) = pvTrackImageList(c.ColumnHeaderIcons, ScaleRatio)
    End If
    If Not (oIL(0) Is Nothing And oIL(1) Is Nothing And oIL(2) Is Nothing) Then
        With c.ListItems
            If lType = cctlListView5 Then
                ReDim lImages(-1 To 0, 1 To .Count)
            Else    ' v6 column headers can have images, track them
                ReDim lImages(-1 To c.ColumnHeaders.Count, 0 To .Count)
                If Not oIL(2) Is Nothing Then
                    For lIndex = 1 To c.ColumnHeaders.Count
                        lImages(lIndex, 0) = c.ColumnHeaders(lIndex).Icon
                    Next
                    Set c.ColumnHeaderIcons = oIL(2)
                End If
            End If
            If Not (oIL(0) Is Nothing And oIL(1) Is Nothing) Then
                For lIndex = 1 To .Count ' track assigned images
                    With .Item(lIndex)
                        lImages(-1, lIndex) = .SmallIcon
                        lImages(0, lIndex) = .Icon
                        If lType = cctlListView6 Then
                            For sIndex = 1 To .ListSubItems.Count
                                lImages(sIndex, lIndex) = .ListSubItems(sIndex).ReportIcon
                            Next
                        End If
                    End With
                Next
                If Not oIL(0) Is Nothing Then Set c.Icons = oIL(0)
                If Not oIL(1) Is Nothing Then Set c.SmallIcons = oIL(1)
            End If
        End With
    End If
    
    Set c.Font = m_Assistant.ScaleStdFont(c.Font, m_Assistant.DpiForForm)
    c.Move newX, newY, newCx, newCy
    m_Assistant.SyncOcxToParent c
    For lIndex = 1 To c.ColumnHeaders.Count
        With c.ColumnHeaders(lIndex)
            .Width = .Width * ScaleRatio
            If Not oIL(2) Is Nothing Then ' reassign images
                If IsEmpty(lImages(lIndex, 0)) = False Then c.ColumnHeaders(lIndex).Icon = lImages(lIndex, 0)
            End If
        End With
    Next
    With c.ListItems
        For lIndex = 1 To .Count
            With .Item(lIndex) ' reassign images
                If Not oIL(0) Is Nothing Then
                    If IsEmpty(lImages(0, lIndex)) = False Then .Icon = lImages(0, lIndex)
                End If
                If Not oIL(1) Is Nothing Then
                    If IsEmpty(lImages(-1, lIndex)) = False Then .SmallIcon = lImages(-1, lIndex)
                    If lType = cctlListView6 Then
                        For sIndex = 1 To .ListSubItems.Count
                            If IsEmpty(lImages(sIndex, lIndex)) = False Then .ListSubItems(sIndex).ReportIcon = lImages(sIndex, lIndex)
                        Next
                    End If
                End If
                If c.View <> lvwReport Then
                    .Left = .Left * ScaleRatio: .Top = .Top * ScaleRatio
                End If
            End With
        Next
    End With
    Erase oIL(): Erase lImages()
        
End Sub

Private Sub pvDoTreeView(lType As CCTLEnum, c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single)

    ' treeviews seem to not have any scaling quirks except:
    '   if checkbox style applied
    '   - changing the imagelist also resets all check values
    '   - checkbox rendering is always at system DPI, it does not scale

    Dim lImages() As Variant, lIndex As Long, oIL As Control
    
    Set oIL = pvTrackImageList(c.ImageList, ScaleRatio)
    If Not oIL Is Nothing Then
        With c.Nodes
            If lType = cctlTreeView5 Then
                ReDim lImages(0 To 1, 1 To .Count)
            Else
                ReDim lImages(0 To 2, 1 To .Count)
            End If
            If .Count <> 0 Then ' track assigned images/checked properties
                For lIndex = 1 To .Count
                    With .Item(lIndex)
                        lImages(0, lIndex) = .Image
                        lImages(1, lIndex) = .SelectedImage
                        If lType = cctlTreeview6 Then lImages(2, lIndex) = .Checked
                    End With
                Next
            End If
        End With
        Set c.ImageList = oIL
    End If
    Set c.Font = m_Assistant.ScaleStdFont(c.Font, m_Assistant.DpiForForm)
    c.Move newX, newY, newCx, newCy
    m_Assistant.SyncOcxToParent c
    c.Indentation = c.Indentation * ScaleRatio
    If Not oIL Is Nothing Then
        With c.Nodes
            For lIndex = 1 To .Count
                With .Item(lIndex) ' reassign images/checked properties
                    If IsEmpty(lImages(0, lIndex)) = False Then .Image = lImages(0, lIndex)
                    If IsEmpty(lImages(1, lIndex)) = False Then .SelectedImage = lImages(1, lIndex)
                    If lType = cctlTreeview6 Then .Checked = lImages(2, lIndex)
                End With
            Next
        End With
    End If
    Set oIL = Nothing: Erase lImages()
    
End Sub

Private Sub pvDoStatusBar(lType As CCTLEnum, c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single)

    ' status bar has no imagelists

    Dim Index As Long
    
    Set c.Font = m_Assistant.ScaleStdFont(c.Font, m_Assistant.DpiForForm)
    c.Move newX, newY, newCx, newCy
    m_Assistant.SyncOcxToParent c
    For Index = 1 To c.Panels.Count
        c.Panels(Index).MinWidth = c.Panels(Index).MinWidth * ScaleRatio
    Next
    
End Sub

Private Sub pvDoTabStrip(lType As CCTLEnum, c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single)

    ' TabStrips appear to be straightforward

    Dim Index As Long, oIL As Control
    Dim lImages() As Variant

    Set oIL = pvTrackImageList(c.ImageList, ScaleRatio)
    If Not oIL Is Nothing Then
        With c.Tabs
            If .Count <> 0 Then
                ReDim lImages(1 To .Count)
                For Index = 1 To .Count ' track assigned images
                    lImages(Index) = .Item(Index).Image
                Next
            End If
        End With
        Set c.ImageList = oIL
    End If
    Set c.Font = m_Assistant.ScaleStdFont(c.Font, m_Assistant.DpiForForm)
    c.Move newX, newY, newCx, newCy
    m_Assistant.SyncOcxToParent c
    c.TabFixedHeight = c.TabFixedHeight * ScaleRatio
    c.TabFixedWidth = c.TabFixedWidth * ScaleRatio
    If Not oIL Is Nothing Then
        With c.Tabs
            If .Count <> 0 Then
                For Index = 1 To .Count ' reassign images
                    If IsEmpty(lImages(Index)) = False Then .Item(Index).Image = lImages(Index)
                Next
            End If
        End With
        Set oIL = Nothing: Erase lImages()
    End If
    
End Sub

Private Sub pvDoImageCombo(c As Control, ScaleRatio As Single, newX As Single, newY As Single, newCx As Single, newCy As Single, initLoad As Boolean)

    ' There are 3 special scaling items with this control
    ' 1) Indentation properties. Those are Integer values. This class does not scale
    '       those properties. It is up to the user to scale them if they are used.
    '       If scaled, you should hardcode the indentation property and scale that
    '       from design DPI to current DPI. Since the value is Integer, rounding
    '       occurs. Scaling a value that forces rounding should not be used for
    '       subsequent scaling.
    ' 2) You should set the control's width in Form_Load to the same value it is
    '       in design view. If the form is loaded into non-integral DPI, then the
    '       control will be scaled by VB and internally, the control will scale
    '       smaller, before Form_Load is triggered. Re-setting the control width in
    '       Form_Load should fix that issue.
    ' 3) The "Edit" box of the combobox is not scaled correctly in only one scenario.
    '       This applies for Win10.1703 or later which will scale the non-client.
    '       First, it appears the width of the dropdown arrow is cached at system DPI
    '       by the control. However, it also appears the "Edit" box is not sized until
    '       the control is first shown. Now, when the control is first loaded into a
    '       DPI that is not the system DPI, it is using the arrow width at system DPI
    '       to size the "Edit" box. The arrow width is different at different DPIs.
    '       Adjusting the "Edit" box width in this routine is overridden. To fix this
    '       problem, this routine will subclass the "Edit" box and adjust its width;
    '       only when first shown, and only when the form is loaded into a DPI other
    '       than system DPI. Subclass is released immediately after the "Edit" box
    '       width is adjusted.

    Dim Index As Long, oIL As Control, lImages() As Variant
    Dim oLoader As clsDpiPmFormLoader
    Dim aData() As Long, cHwnd As Long
    Const SM_CXHSCROLL As Long = 21
    Const WM_PAINT = &HF&
    
    Set oIL = pvTrackImageList(c.ImageList, ScaleRatio)
    If Not oIL Is Nothing Then
        With c.ComboItems
            If .Count <> 0 Then
                ReDim lImages(0 To 1, 1 To .Count)
                For Index = 1 To .Count ' track assigned images
                    lImages(0, Index) = .Item(Index).Image
                    lImages(1, Index) = .Item(Index).SelImage
                Next
            End If
            Set c.ImageList = oIL
            If .Count <> 0 Then
                For Index = 1 To c.ComboItems.Count ' reassign images
                    If IsEmpty(lImages(0, Index)) = False Then .Item(Index).Image = lImages(0, Index)
                    If IsEmpty(lImages(1, Index)) = False Then .Item(Index).SelImage = lImages(1, Index)
                Next
            End If
        End With
        Set oIL = Nothing: Erase lImages()
    End If
    Set c.Font = m_Assistant.ScaleStdFont(c.Font, m_Assistant.DpiForForm)
    c.Move newX, newY, newCx, newCy
    m_Assistant.SyncOcxToParent c
    If (m_Assistant.WindowsCaps And wcpMask_Version) > wcpWin10_1607 Then
        ' adjust the combobox "Edit" control based on the dropdown arrow width
        ' as defined by this DPI
        ReDim aData(0 To 5)
        aData(4) = GetSystemMetricsForDpi(SM_CXHSCROLL, m_Assistant.DpiForForm)
        aData(5) = GetSystemMetricsForDpi(SM_CXHSCROLL, m_Assistant.DpiForSystem)
        If aData(4) <> aData(5) Then
            cHwnd = FindWindowExW(c.hWnd, 0, StrPtr("ComboBox"), 0)
            If cHwnd <> 0 Then
                cHwnd = FindWindowExW(cHwnd, 0, StrPtr("Edit"), 0)
                If cHwnd <> 0 Then
                    If initLoad = True Then ' not shown yet
                        If oSubclasser Is Nothing Then  ' create subclasser as needed
                            Set oLoader = New clsDpiPmFormLoader
                            Set oSubclasser = oLoader.CreateTasker_Subclass(Me, 1)
                            Set oLoader = Nothing
                            oSubclasser.FilterAdd 0, WM_PAINT ' only message we care about
                        End If
                        oSubclasser.AddItem cHwnd       ' subclass & set arrow width change in Tag
                        oSubclasser.TagItem(cHwnd) = aData(4) - aData(5)
                    Else
                        GetWindowRect cHwnd, aData(0)   ' adjust Edit box width
                        aData(2) = aData(2) - aData(0): aData(3) = aData(3) - aData(1)
                        SetWindowPos cHwnd, 0, 0, 0, aData(2) - (aData(4) - aData(5)), aData(3), SWP_NOZORDER_REPOSITION Or SWP_NOMOVE
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Private Function pvTrackImageList(c As Control, ScaleRatio As Single) As Control

    ' routine keeps track of which image lists have been scaled
    ' parameter is an imagelist

    If m_Host = -1 Then Exit Function   ' this class is not implemented
    If c Is Nothing Then Exit Function  ' very possible
    If c.ListImages.Count = 0 Then Exit Function ' unlikely, but possible

    Dim oIL As Control, sName As String
    
    On Error Resume Next
    ' already been here for this imagelist? If so, return the scaled imagelist
    Set pvTrackImageList = m_ImageLists(CStr(ObjPtr(c)))(1)
    If Err Then
        Err.Clear   ' else fire event to have it scaled
        sName = m_DynamicLists(CStr(ObjPtr(c)))(1)
        If Err Then
            Err.Clear
            sName = c.Name
        End If
        pvGetHost().CallbackScaleImageList sName, oIL, ScaleRatio
        If Not oIL Is Nothing Then
            If Not oIL Is c Then
                If TypeName(oIL) = "ImageList" Then Set pvTrackImageList = oIL
            End If
            Set oIL = Nothing
        End If
        If m_ImageLists Is Nothing Then Set m_ImageLists = New Collection
        m_ImageLists.Add Array(c, pvTrackImageList), CStr(ObjPtr(c))
    End If
    On Error GoTo 0

End Function

Private Function pvGetProgID(c As Control) As String

    ' this class cannot be hardcoded with a specific version of imagelist controls
    ' because two versions exist. it is also possible that both versions are being
    ' used simultaneously in a project.

    ' helper function to get the ProgID of passed ImageList control.
    ' The ProgID is needed to guarantee we can create a run-time
    '   ImageList of the same version as the passed control.

    Dim oIPersist As stdole.IUnknown, pPtr As Long, lLen As Long
    Dim vParamPtr(0 To 1) As Long, vParamType(0 To 1) As Integer
    Dim vRtn As Variant, vParams(0 To 2) As Variant
    Dim aData(0 To 3) As Long ' parameter values
    Const IID_IPersist = "{0000010c-0000-0000-c000-000000000046}"
    Const IP_GetCLSID = 12&
    
    vParamPtr(0) = VarPtr(vParams(0)): vParamType(0) = vbLong
    vParamPtr(1) = VarPtr(vParams(1)): vParamType(1) = vbLong
    
    On Error GoTo errHandler    ' need the IPersist interface of the passed control
    IIDFromString StrPtr(IID_IPersist), VarPtr(aData(0))
    vParams(0) = VarPtr(aData(0)): vParams(1) = VarPtr(oIPersist)
    DispCallFunc ObjPtr(c.object), 0&, 4&, vbLong, 2, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    If oIPersist Is Nothing Then Exit Function  ' sanity check; unexpected if aborts
    
    vParams(0) = VarPtr(aData(0)) ' get the CLSID of the passed control
    DispCallFunc ObjPtr(oIPersist), IP_GetCLSID, 4&, vbLong, 1, VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vRtn
    Set oIPersist = Nothing
    
    ' get ProgID from CLSID, returned as a pointer to unicode string
    If ProgIDFromCLSID(aData(0), pPtr) = 0 Then
        lLen = lstrlenW(ByVal pPtr)
        pvGetProgID = Space$(lLen)
        CopyMemory ByVal StrPtr(pvGetProgID), ByVal pPtr, lLen * 2
        CoTaskMemFree pPtr          ' clean up
    End If
    
errHandler:
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Private Function pvIsCommonControl(c As Control) As CCTLEnum

    ' routine checks passed control as a common controls object
    ' The window class of the control is used for validation and
    '   distinguishing between v5 or v6 is also performed
    
    Dim sClass As String, n As Long
    n = 30: sClass = Space$(n)
    On Error Resume Next
    n = GetClassName(c.hWnd, StrPtr(sClass), n + 1)
    If Err Then
        Err.Clear       ' err? control exposes no hWnd, i.e., ADO Data control
        Exit Function
    End If
    On Error GoTo 0
    
    Select Case Left$(sClass, n)
    
    ' version 6 common controls (mscomctl.ocx)
    Case "ListView20WndClass": pvIsCommonControl = cctlListView6
    Case "TreeView20WndClass": pvIsCommonControl = cctlTreeview6
    Case "ProgressBar20WndClass": pvIsCommonControl = cctlProgressbar6
    Case "StatusBar20WndClass": pvIsCommonControl = cctlStatusbar6
    Case "ImageCombo20WndClass": pvIsCommonControl = cctlImageCombo
    Case "Slider20WndClass": pvIsCommonControl = cctlSlider6
    Case "Toolbar20WndClass": pvIsCommonControl = cctlToolbar6
    Case "TabStrip20WndClass": pvIsCommonControl = cctlTabStrip6
    
    ' version 5 common controls (comctl32.ocx)
    Case "ListViewWndClass": pvIsCommonControl = cctlListView5
    Case "TreeViewWndClass": pvIsCommonControl = cctlTreeView5
    Case "ProgressBarWndClass": pvIsCommonControl = cctlProgressbar5
    Case "StatusBarWndClass": pvIsCommonControl = cctlStatusbar5
    Case "SliderWndClass": pvIsCommonControl = cctlSlider5
    Case "ToolbarWndClass": pvIsCommonControl = cctlToolbar5
    Case "TabStripWndClass": pvIsCommonControl = cctlTabStrip5
    
    ' version 6 common controls-2 (mscomct2.ocx)
    Case "DTPicker20WndClass": pvIsCommonControl = cctlDtPicker
    Case "MonthView20WndClass": pvIsCommonControl = cctlMonthView
    Case "FlatSB20WndClass": pvIsCommonControl = cctlFlatScroll
    Case "UpDown20WndClass": pvIsCommonControl = cctlUpDown6
    Case "Animation20WndClass": pvIsCommonControl = cctlAnimation6
    
    ' version 5 common controls-2 (comct232.ocx)
    Case "UpDownWndClass": pvIsCommonControl = cctlUpDown5
    Case "AnimationWndClass": pvIsCommonControl = cctlAnimation5
    
    ' version 6 common controls-3 (comct332.ocx)
    Case "ThunderRT6UserControlDC": pvIsCommonControl = cctlCoolbar
    End Select
    
End Function

Private Function pvGetHost() As clsDpiAsstCmnCtrls
    ' helper routine to get implementation of this class' host
    Dim oIUnk As Object
    CopyMemory oIUnk, m_Host, 4
    Set pvGetHost = oIUnk
    CopyMemory oIUnk, 0&, 4
End Function

Private Sub Class_Terminate()
    Set oSubclasser = Nothing
    Set m_DynamicLists = Nothing
    Set m_Assistant = Nothing
End Sub

Public Sub CallbackScaleImageList(ByVal ImageListName As String, ByRef ImgLstReplacement As Control, ByVal ScaleRatio As Single)
    ' this is sent via Implementation. Do not call this from your code -- does nothing
    ' it will be sent if an ImageList needs scaling and you declared this class
    '   with the keyword: Implements
    
    ' When this is sent to you, then
    '   ImageListName is the name of the control to be scaled
    '   ImgLstReplacement must be returned as either Nothing or a valid ImageList
    '       If ImgLstReplacement is not Nothing, then all controls that
    '       reference passed ImageList will be changed to reference ImgLstReplacement
    
    ' About the easist way to scale ImageLists via this class is...
    ' 1. Ensure your form implements this class
    ' 2. Create your primary imagelist(s), recommend this simple rule:
    '       If more than one ImageList is created for same images at different
    '       scales, then create them as a control array; i.e., set Index property.
    ' 3. In this implemented callback, use code similar to the following
    '    notes:
    '    DpiAssist below is the form's variable for its clsDpiPmAssist class instance
    '    CmnCtrlsAssist below is the form's variable for this class instance
    '    Don't include CASE statements that do not apply, add any others that apply
    
    ' Private Sub CmnCtrlsAssist_CallbackScaleImageList(ByVal ImageListName As String, ImgLstReplacement As Control)
    '
    '    Dim lSize As Long, sngScale As Single
    '    sngScale = DpiAssist.DpiForForm / 96
    '    Select Case ImageListName
    '
    '    ' example: 2 base lists (indexed), 1 for each: 100%, 200% DPI, i.e., 16x16 & 32x32
    '    Case "ilTreeView"
    '        lSize = 16 * sngScale ' scale from smallest of the imagelists
    '        If sngScale <= 1! Then
    '            Set ImgLstReplacement = CmnCtrlsAssist.CloneImageList(ilTreeView(0), lSize, lSize)
    '        Else
    '            Set ImgLstReplacement = CmnCtrlsAssist.CloneImageList(ilTreeView(1), lSize, lSize)
    '        End If
    '
    '    ' example: 3 base lists (indexed), 1 for each: 100%, 150%, 200% DPI, i.e., 16x16, 24x24, 32x32
    '    Case "ilListView"
    '        lSize = 16 * sngScale ' scale from smallest of the imagelists
    '        If sngScale <= 1! Then
    '            Set ImgLstReplacement = CmnCtrlsAssist.CloneImageList(ilListView(0), lSize, lSize)
    '        ElseIf sngScale <= 1.5! then
    '            Set ImgLstReplacement = CmnCtrlsAssist.CloneImageList(ilListView(1), lSize, lSize)
    '        Else
    '            Set ImgLstReplacement = CmnCtrlsAssist.CloneImageList(ilListView(2), lSize, lSize)
    '        End If
    '
    '    ' example: 1 base list for 96 DPI, where 32x32 is the smallest size
    '    Case Else
    '        lSize = 32 * sngScale
    '        Set ImgLstReplacement = CmnCtrlsAssist.CloneImageList(ImageList1, lSize, lSize)
    '    End Select
    ' End Sub
    
    ' You can request an empty cloned ImageList by passing CloneImageList negative Width,Height
    ' values. When the clone is returned, populate the images and keys (if applies) yourself.
    
End Sub

'/////////////////////// DO NOT PLACE ANY NEW CODE AFTER THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function zzSubclassProc(ByVal TaskerObject As Object, _
                            ByRef EatMessage As Boolean, _
                            ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long, _
                            ByVal hwndTag As Long) As Long
                            
    ' see pvDoImageCombo
    Dim aData(0 To 3) As Long
    
    zzSubclassProc = TaskerObject.CallNextInChain(hWnd, uMsg, wParam, lParam)
    If TaskerObject.RemoveItem(hWnd) = 0 Then Set oSubclasser = Nothing
    GetWindowRect hWnd, aData(0)
    aData(2) = aData(2) - aData(0): aData(3) = aData(3) - aData(1)
    SetWindowPos hWnd, 0, 0, 0, aData(2) - hwndTag, aData(3), SWP_NOZORDER_REPOSITION Or SWP_NOMOVE
    EatMessage = True
    
End Function
'/////////////////////// DO NOT PLACE ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\



