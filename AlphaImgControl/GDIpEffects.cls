VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIpEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'  -----======== PURPOSE: Manage GDI+ ImageAttributes Object ========-----
' Class Instancing = MultiUse (can be created by users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


Public Event Changed(PropertyName As String)

' Global class that can be created on demand

' This class assists in creating GDI+ v1.1 graphics effects
' This class assists in creating GDI+ image attributes handles using custom matrices
' a list of some nice color matrix effects...
'   http://www.universalthread.com/ViewPageArticle.aspx?Session=71624655657A38726E4F6B3D206245742F4D50367A7859536659496D50352F72456E413D3D
'   http://www.universalthread.com/ViewPageArticle.aspx?Session=64776F5432455165304C593D20426F4449336E364A366C485864734A72335A4C4C77413D3D

' To create an image attributes handle, simply set the public properties that apply
'   BlendColor :: the RGB or System color value that is to be blended into the rendered image
'   BlendPct :: the amount of BlendColor to be blended. Set to zero for no blending
'   GlobalTransparencyPct :: amount of translucency applied to entire rendered image. Set to zero for none
'   GrayScale :: the grayscale formula to use for grayscaling rendered image
'   Invert :: the result is a color negative of the image
'   LightnessPct :: the percentage of lightness to add or remove from rendered image. Negative values remove lightness
'   TransparentColor :: RGB or System color to make transparent throughout the image
'   TransparentColorUsed :: triggers if TransparentColor is used or not
' That's all. To get the image attributes handle for rendering, call the AttributesHandle property
' DestroyAttributesHandle method releases the attributes handle as does ResetAllAttributes
' ResetAllAttributes resets grayscale, global transparency, lightness, transparency color & blend colors to unused/default

' To create various effects, call one of the following Createxxxx functions
' Version 1.1 or greater of GDI+ is required
'   CreateBlurEffect
'   CreateBrightnessContrastEffect
'   CreateColorBalanceEffect
'   CreateColorCurveEffect
'   CreateHueSaturationLightnessEffect
'   CreateRedEyeCorrectionEffect
'   CreateSharpenEffect
'   CreateTintEffect
'   CreateToneLevelEffect
' To get the handle of the effects, call the EffectsHandle passing the effect type you created
'   DestroyEffectsHandle clears all or any of the effects you have created
' Note: All 9 effects can be created, however, only one can be applied per rendering
' If any attributes are applied, they can be rendered with/without any effects & vice versa

' v1.1 functions
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef Effect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal Effect As Long, ByRef params As Any, ByVal Size As Long) As Long
Private Declare Function GdipGetEffectParameters Lib "gdiplus" (ByVal Effect As Long, ByRef Size As Long, ByRef params As Any) As Long
Private Declare Function GdipGetEffectParameterSize Lib "gdiplus" (ByVal Effect As Long, ByRef Size As Long) As Long
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal Effect As Long) As Long

Private Type BlurParams
  Radius As Single
  ExpandEdge As Long
End Type
Private Type FXParamsTriLong
    Param1 As Long
    Param2 As Long
    Param3 As Long
End Type
Private Type FXParamsLong
    Param1 As Long
    Param2 As Long
End Type
Private Type FXParamsSingle
    Param1 As Single
    Param2 As Single
End Type

Private m_Token As cGDIpToken
Private m_hEffect(lvicBlurFX To lvicRedEyeCorrectionFX)  As Long ' effect handles
Private m_hAttributes As Long   ' image attributes handle

' m_AttrFlag masking
' &H8....... > set if no attributes created; also used to prevent internally for other purposes
' &H.7...... > Inverted colors
' &H...F.... > Gray scale value
' &H....7F.. > Global transparency value
' &H......7F > Lightness percent
' &H......8. > Negative lightness percent
' &H78F.8... > reserved for future use

' m_TransColor masking
' &H8.FFFFFF > transparent color value (can be system color)
' &H.F...... > transparent color used or not

' m_BlendColor masking
' &H8.FFFFFF > blend color vlaue (can be system color)
' &H7F...... > blend percent

Private m_AttrFlag As Long      ' contains lightness, transparency pct, grayscale, usage flag
Private m_TransColor As Long    ' transparent color and usage
Private m_BlendColor As Long    ' blend color and percentage

Public Property Get EffectsHandle(FXtype As EffectsEnum) As Long
    If Not (FXtype < lvicBlurFX Or FXtype > lvicRedEyeCorrectionFX) Then EffectsHandle = m_hEffect(FXtype)
End Property

Public Property Get AttributesHandle() As Long
    If m_hAttributes = 0& Then
        If (m_AttrFlag And &H80000000) = 0& Then                     ' set? If so, we know nothing to do
            If m_Token.Token Then
                Dim lTrans As Long, lBlend As Long
                ' colors are required to be in BGRA format
                If (m_TransColor And &HF000000) Then                ' color used
                    If (m_TransColor And &H80000000) Then           ' system color vs RGB
                        lTrans = GetSysColor(m_TransColor And &HFF&) Or &HFF000000
                    Else
                        lTrans = (m_TransColor And &HFFFFFF) Or &HFF000000
                    End If
                End If
                If (m_BlendColor And &H7F000000) Then               ' color used
                    If (m_BlendColor And &H80000000) Then           ' system color vs RGB
                        lBlend = GetSysColor(m_BlendColor And &HFF&) Or &H7F000000
                    Else
                        lBlend = m_BlendColor
                    End If
                End If                                              ' create the attributes handle
                modCommon.CreateGDIpAttributeHandle m_hAttributes, Me.GrayScale, Me.LightnessPct, Me.GlobalTransparencyPct, lTrans, lBlend, Me.Invert
                If m_hAttributes = 0& Then m_AttrFlag = m_AttrFlag Or &H80000000 ' so we don't keep doing this all the time
                ' Note: If any attributes do change, their properties reset the high byte
            End If
        End If
    End If
    AttributesHandle = m_hAttributes
End Property

'/// Grayscale maintained in 3rd byte of m_AttrFlag
Public Property Get GrayScale() As GrayScaleRatioEnum
    GrayScale = (m_AttrFlag And &HF0000) \ &H10000
End Property
Public Property Let GrayScale(newValue As GrayScaleRatioEnum)
    If Not (newValue < lvicNoGrayScale Or newValue > lvicBlackWhite) Then
        If Not newValue = Me.GrayScale Then
            m_AttrFlag = (m_AttrFlag And &H7FF0FFFF) Or newValue * &H10000 Or &H80000000
            DestroyAttributesHandle
            RaiseEvent Changed("GrayScale")
        End If
    End If
End Property

'/// GlobalTransparencyPct maintained in 2nd byte of m_AttrFlag
Public Property Get GlobalTransparencyPct() As Long
    GlobalTransparencyPct = (m_AttrFlag And &H7F00) \ &H100&
End Property
Public Property Let GlobalTransparencyPct(newValue As Long)
    If Not (newValue < 0& Or newValue > 100&) Then
        If Not newValue = Me.GlobalTransparencyPct Then
            m_AttrFlag = (m_AttrFlag And &H7FFF80FF) Or newValue * &H100& Or &H80000000
            DestroyAttributesHandle
            RaiseEvent Changed("GlobalTransparency")
        End If
    End If
End Property

'/// LightnessPct maintained in 1st byte of m_AttrFlag
' 8th bit set if LightnessPct is negative value
Public Property Get LightnessPct() As Long
    If (m_AttrFlag And &H80) Then
        LightnessPct = -(m_AttrFlag And &H7F&)
    Else
        LightnessPct = (m_AttrFlag And &H7F&)
    End If
End Property
Public Property Let LightnessPct(newValue As Long)
    If Not (newValue < -100& Or newValue > 100&) Then
        If Not newValue = Me.LightnessPct Then
            m_AttrFlag = (m_AttrFlag And &H7FFFFF00) Or Abs(newValue) Or &H80000000
            If newValue < 0& Then m_AttrFlag = m_AttrFlag Or &H80&
            DestroyAttributesHandle
            RaiseEvent Changed("LightnessPct")
        End If
    End If
End Property

'//// Mask is &H80FFFFFF. 7th nibble is set if transparent color used
Public Property Get TransparentColor() As Long
    ' Returned color will be standard VB RGB or system color
    TransparentColor = (m_TransColor And &H80FFFFFF)
End Property
Public Property Let TransparentColor(newValue As Long)
    ' Passed color must be standard VB RGB or system color
    If Not newValue = Me.TransparentColor Then
        m_TransColor = (m_TransColor And &HF000000) Or (newValue And &H80FFFFFF)
        m_AttrFlag = (m_AttrFlag And &H7FFFFFFF) Or &H80000000
        DestroyAttributesHandle
        RaiseEvent Changed("TransparentColor")
    End If
End Property

'//// Inverted colors setting maintained in 7th nibble
Public Property Get Invert() As Boolean
    Invert = (m_AttrFlag And &H7000000)
End Property
Public Property Let Invert(newValue As Boolean)
    If Not newValue = Me.Invert Then
        m_AttrFlag = (m_AttrFlag And &H78FFFFFF) Or &H7000000 * Abs(newValue) Or &H80000000
        DestroyAttributesHandle
        RaiseEvent Changed("Invert")
    End If
End Property

'//// Mask is &H80FFFFFF. 7th nibble is set if transparent color used
Public Property Get TransparentColorUsed() As Boolean
    TransparentColorUsed = (m_TransColor And &HF000000)
End Property
Public Property Let TransparentColorUsed(newValue As Boolean)
    If Not newValue = Me.TransparentColorUsed Then
        m_TransColor = (m_TransColor And &H80FFFFFF)
        If newValue Then m_TransColor = m_TransColor Or &HF000000
        m_AttrFlag = (m_AttrFlag And &H7FFFFFFF) Or &H80000000
        DestroyAttributesHandle
        RaiseEvent Changed("TransparentColorUsed")
    End If
End Property

'//// Mask is &H80FFFFFF. Percentage Mask is &H7F000000
Public Property Get BlendColor() As Long
    ' Returned color will be standard VB RGB or system color
    BlendColor = (m_BlendColor And &H80FFFFFF)
End Property
Public Property Let BlendColor(newValue As Long)
    ' Passed color must be standard VB RGB or system color
    If Not Me.BlendColor = newValue Then
        m_BlendColor = (m_BlendColor And &H7F000000) Or (newValue And &H80FFFFFF)
        m_AttrFlag = (m_AttrFlag And &H7FFFFFFF) Or &H80000000
        DestroyAttributesHandle
        RaiseEvent Changed("BlendColor")
    End If
End Property

'//// Mask is &H80FFFFFF. Percentage Mask is &H7F000000
Public Property Get BlendPct() As Long
    BlendPct = (m_BlendColor And &H7F000000) \ &H1000000
End Property
Public Property Let BlendPct(newValue As Long)
    If Not (newValue < 0& Or newValue > 100&) Then
        If Not newValue = Me.BlendPct Then
            m_BlendColor = (m_BlendColor And &H80FFFFFF) Or (newValue * &H1000000)
            m_AttrFlag = (m_AttrFlag And &H7FFFFFFF) Or &H80000000
            DestroyAttributesHandle
            RaiseEvent Changed("BlendPct")
        End If
    End If
End Property

Public Function CreateBlurEffect(ByVal Radius As Byte, ByVal ExpandEdge As Boolean) As Boolean

    ' Radius valid values are: 0 to 255
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534057(v=VS.85).aspx
    
    Dim bp As BlurParams, sEffect As String
    If pvCreateNewEffect(lvicBlurFX, sEffect) Then
        bp.Radius = CSng(Radius)
        bp.ExpandEdge = Abs(ExpandEdge)
        If GdipSetEffectParameters(m_hEffect(lvicBlurFX), bp, Len(bp)) = 0& Then
            CreateBlurEffect = True
            RaiseEvent Changed(sEffect)
        End If
    End If
End Function

Public Function CreateBrightnessContrastEffect(ByVal BrightnessPct As Long, ByVal ContrastPct As Long) As Boolean

    ' BrightnessPct valid values are: -255 to 255
    ' ContrastPct valid values are: -100 to 100
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534058(v=VS.85).aspx

    Dim bp As FXParamsLong, sEffect As String
    If Not (BrightnessPct = 0& And ContrastPct = 0&) Then
        If pvCreateNewEffect(lvicBrightnessContrastFX, sEffect) Then
            If BrightnessPct < -255& Then
                BrightnessPct = -255&
            ElseIf BrightnessPct > 255& Then
                BrightnessPct = 255&
            End If
            If ContrastPct < -100& Then
                ContrastPct = -100&
            ElseIf ContrastPct > 100& Then
                ContrastPct = 100&
            End If
            bp.Param1 = BrightnessPct
            bp.Param2 = ContrastPct
            If GdipSetEffectParameters(m_hEffect(lvicBrightnessContrastFX), bp, Len(bp)) = 0& Then
                CreateBrightnessContrastEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If
End Function

Public Function CreateSharpenEffect(ByVal Radius As Byte, ByVal Percent As Long) As Boolean

    ' radius valid values are: 0 to 255
    ' Percent valid values are: 0 to 100
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534073(v=VS.85).aspx

    Dim bp As FXParamsSingle, sEffect As String
    
    If Percent > 100& Then
        Percent = 100&
    ElseIf Percent < 1& Then
        Percent = 0&
    End If
    
    If Not (Radius = 0 Or Percent = 0&) Then
        If pvCreateNewEffect(lvicSharpenFX, sEffect) Then
            bp.Param1 = CSng(Radius)
            bp.Param2 = Percent
            If GdipSetEffectParameters(m_hEffect(lvicSharpenFX), bp, Len(bp)) = 0& Then
                CreateSharpenEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If

End Function

Public Function CreateTintEffect(ByVal Hue As Long, ByVal Percent As Long) As Boolean

    ' Hue valid values are: -180 to 180
    ' Percent valid values are: -100 to 100
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534074(v=VS.85).aspx

    Dim bp As FXParamsLong, sEffect As String
    
    If Not (Hue = 0& And Percent = 0&) Then
        If pvCreateNewEffect(lvicTintFX, sEffect) Then
            If Percent > 100& Then
                Percent = 100&
            ElseIf Percent < -100& Then
                Percent = -100&
            End If
            If Hue > 180& Then
                Hue = 180&
            ElseIf Hue < -180& Then
                Hue = -180&
            End If
            bp.Param1 = Hue
            bp.Param2 = Percent
            If GdipSetEffectParameters(m_hEffect(lvicTintFX), bp, Len(bp)) = 0& Then
                CreateTintEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If

End Function

Public Function CreateHueSaturationLightnessEffect(ByVal Hue As Long, ByVal Saturation As Long, ByVal Lightness As Long) As Boolean

    ' Hue valid values are: -180 to 180
    ' Saturation & Lightness valid values are: -100 to 100
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534069(v=VS.85).aspx
    
    Dim bp As FXParamsTriLong, sEffect As String
    If (Hue Or Saturation Or Lightness) Then
        If pvCreateNewEffect(lvicHueSaturationLightnessFX, sEffect) Then
            If Hue < -180& Then
                Hue = -180&
            ElseIf Hue > 180& Then
                Hue = 180&
            End If
            If Saturation < -100& Then
                Saturation = -100&
            ElseIf Saturation > 100& Then
                Saturation = 100&
            End If
            If Lightness < -100& Then
                Lightness = -100&
            ElseIf Lightness > 100& Then
                Lightness = 100&
            End If
            bp.Param1 = Hue: bp.Param2 = Saturation: bp.Param3 = Lightness
            If GdipSetEffectParameters(m_hEffect(lvicHueSaturationLightnessFX), bp, Len(bp)) = 0& Then
                CreateHueSaturationLightnessEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If

End Function

Public Function CreateRedEyeCorrectionEffect(AdjustmentRECTs() As REDEYECORRECTIONAREA) As Boolean
    
    ' AdjustmentRECTs define the area(s) to where red-eye correction will take place
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534072(v=VS.85).aspx
    
    Dim rc() As REDEYECORRECTIONAREA, lCount As Long, X As Long
    Dim bp As FXParamsLong, sEffect As String
    
    If Not (Not AdjustmentRECTs()) Then             ' else passed null array
        
        If pvCreateNewEffect(lvicRedEyeCorrectionFX, sEffect) Then
            On Error GoTo ExitRoutine                   ' error if passing multiple dimensioned array
            For X = LBound(AdjustmentRECTs) To UBound(AdjustmentRECTs)
                If AdjustmentRECTs(X).Right > AdjustmentRECTs(X).Left Then
                    If AdjustmentRECTs(X).Bottom > AdjustmentRECTs(X).Top Then lCount = lCount + 1&
                End If
            Next
            On Error GoTo 0
            If lCount = 0& Then GoTo ExitRoutine
            
            ReDim rc(0 To lCount - 1&)
            For X = LBound(AdjustmentRECTs) To UBound(AdjustmentRECTs)
                If AdjustmentRECTs(X).Right > AdjustmentRECTs(X).Left Then
                    If AdjustmentRECTs(X).Bottom > AdjustmentRECTs(X).Top Then
                        rc(bp.Param1) = AdjustmentRECTs(X)
                        bp.Param1 = bp.Param1 + 1&
                    End If
                End If
            Next
            bp.Param2 = VarPtr(rc(0))
            If GdipSetEffectParameters(m_hEffect(lvicRedEyeCorrectionFX), bp, Len(bp) + lCount * 16&) = 0& Then
                CreateRedEyeCorrectionEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If
    
ExitRoutine:
    Debug.Assert App.hInstance
End Function

Public Function CreateColorBalanceEffect(ByVal CyanPct As Long, ByVal MagentaPct As Long, ByVal YellowPct As Long) As Boolean
    
    ' All 3 param valid values are: -100 to 100
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534059(v=VS.85).aspx

    Dim bp As FXParamsTriLong, sEffect As String
    If (CyanPct Or MagentaPct Or YellowPct) Then
        If pvCreateNewEffect(lvicColorBalanceFX, sEffect) Then
            If CyanPct < -100& Then
                CyanPct = -100&
            ElseIf CyanPct > 100& Then
                CyanPct = 100&
            End If
            If MagentaPct < -100& Then
                MagentaPct = -100&
            ElseIf MagentaPct > 100& Then
                MagentaPct = 100&
            End If
            If YellowPct < -100& Then
                YellowPct = -100&
            ElseIf YellowPct > 100& Then
                YellowPct = 100&
            End If
            bp.Param1 = CyanPct: bp.Param2 = MagentaPct: bp.Param3 = YellowPct
            If GdipSetEffectParameters(m_hEffect(lvicColorBalanceFX), bp, Len(bp)) = 0& Then
                CreateColorBalanceEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If

End Function

Public Function CreateToneLevelEffect(ByVal HighlightPct As Long, ByVal MidTonePct As Long, ByVal ShadowPct As Long) As Boolean
    
    ' HighlightPct & ShadowPct param valid values are: 0 to 100
    ' MidTonePct valid values are -100 to 100
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534070(v=VS.85).aspx

    Dim bp As FXParamsTriLong, sEffect As String
    If HighlightPct < 0& Then
        HighlightPct = 0&
    ElseIf HighlightPct > 100& Then
        HighlightPct = 100&
    End If
    If ShadowPct < 0& Then
        ShadowPct = 0&
    ElseIf ShadowPct > 100& Then
        ShadowPct = 100&
    End If
    
    If (HighlightPct Or MidTonePct Or ShadowPct) Then
        If pvCreateNewEffect(lvicToneLevelsFX, sEffect) Then
            If MidTonePct < -100& Then
                MidTonePct = -100&
            ElseIf MidTonePct > 100& Then
                MidTonePct = 100&
            End If
            bp.Param1 = HighlightPct: bp.Param2 = MidTonePct: bp.Param3 = ShadowPct
            If GdipSetEffectParameters(m_hEffect(lvicToneLevelsFX), bp, Len(bp)) = 0& Then
                CreateToneLevelEffect = True
                RaiseEvent Changed(sEffect)
            End If
        End If
    End If

End Function

Public Function CreateColorCurveEffect(ByVal AdjustmentItem As CurveAdjustmentsEnum, ByVal Channel As CurveChannel, ByVal AdjustmentValue As Long) As Boolean

    ' AdjustmentItem: CurveAdjustment enumeration value
    ' Channel: CurveChannel enumeration value. Passing lvicCurveNoChannels destroys the Effect
    ' Amount: valid values depends on AdjustmentItem & are...
    '       AdjustExposure: -255 to 255
    '       AdjustDensity: -255 to 255
    '       AdjustContrast: -100 to 100
    '       AdjustHighlight: -100 to 100
    '       AdjustShadow: -100 to 100
    '       AdjustMidtone: -100 to 100
    '       AdjustWhiteSaturation: 0 to 255
    '       AdjustBlackSaturation: 0 to 255
    ' Ref: http://msdn.microsoft.com/en-us/library/ms534060(v=VS.85).aspx
    
    Dim bp As FXParamsTriLong, sEffect As String
    If AdjustmentItem < lvicAdjustExposure Or AdjustmentItem > lvicAdjustBlackSaturation Then Exit Function
    If Channel < lvicCurveChannelAll Or Channel > lvicCurveChannelBlue Then Exit Function
    
    If pvCreateNewEffect(lvicColorCurveFX, sEffect) Then
        Select Case AdjustmentItem
            Case lvicAdjustExposure, lvicAdjustDensity
                If AdjustmentValue < -255& Then
                    AdjustmentValue = -255&
                ElseIf AdjustmentValue > 255& Then
                    AdjustmentValue = 255&
                End If
            Case lvicAdjustBlackSaturation, lvicAdjustWhiteSaturation
                If AdjustmentValue < 0& Then
                    AdjustmentValue = 0&
                ElseIf AdjustmentValue > 255& Then
                    AdjustmentValue = 255&
                End If
            Case Else
                If AdjustmentValue < -100& Then
                    AdjustmentValue = -100&
                ElseIf AdjustmentValue > 100& Then
                    AdjustmentValue = 100&
                End If
        End Select
        bp.Param1 = AdjustmentItem: bp.Param2 = Channel: bp.Param3 = AdjustmentValue
        If GdipSetEffectParameters(m_hEffect(lvicColorCurveFX), bp, Len(bp)) = 0& Then
            CreateColorCurveEffect = True
            RaiseEvent Changed(sEffect)
        End If
    End If
    
End Function

Public Function GetEffectParameterValues(ByVal FXtype As EffectsEnum, Param1 As Variant, Param2 As Variant, Optional Param3 As Variant) As Long

    If m_Token.Version < 1.1! Then Exit Function
    
    ' returns the data used to create an effect
    
    Dim paramData() As Byte, paramSize As Long
    Dim tSingle As Single, tLong As Long
    Dim tRC() As REDEYECORRECTIONAREA
    
    If Not (FXtype < lvicBlurFX Or FXtype > lvicRedEyeCorrectionFX) Then
        If m_hEffect(FXtype) Then
            GdipGetEffectParameterSize m_hEffect(FXtype), paramSize
            If paramSize Then
                ReDim paramData(0 To paramSize - 1&)
                If GdipGetEffectParameters(m_hEffect(FXtype), paramSize, paramData(0)) = 0& Then
                    Select Case FXtype
                    Case lvicBlurFX
                        If paramSize = 8& Then
                            CopyMemory tSingle, paramData(0), 4&: Param1 = tSingle
                            CopyMemory tLong, paramData(4), 4&: Param2 = tLong
                            GetEffectParameterValues = 2&
                        End If
                    Case lvicBrightnessContrastFX, lvicTintFX
                        If paramSize = 8& Then
                            CopyMemory tLong, paramData(0), 4&: Param1 = tLong
                            CopyMemory tLong, paramData(4), 4&: Param2 = tLong
                            GetEffectParameterValues = 2&
                        End If
                    Case lvicSharpenFX
                        If paramSize = 8& Then
                            CopyMemory tSingle, paramData(0), 4&: Param1 = tSingle
                            CopyMemory tSingle, paramData(4), 4&: Param2 = tSingle
                            GetEffectParameterValues = 2&
                        End If
                    Case lvicHueSaturationLightnessFX, lvicColorBalanceFX, lvicToneLevelsFX, lvicColorCurveFX
                        If paramSize = 12& Then
                            CopyMemory tLong, paramData(0), 4&: Param1 = tLong
                            CopyMemory tLong, paramData(4), 4&: Param2 = tLong
                            CopyMemory tLong, paramData(8), 4&: Param3 = tLong
                            GetEffectParameterValues = 3&
                        End If
                    Case lvicRedEyeCorrectionFX
                        CopyMemory tLong, paramData(0), 4&: Param1 = tLong
                        paramSize = (paramSize - 8&) \ 16&
                        If paramSize Then
                            ReDim tRC(0 To paramSize - 1&)
                            CopyMemory ByVal VarPtr(tRC(0)), paramData(8), paramSize * 16&
                            Param2 = tRC()
                        End If
                        GetEffectParameterValues = 3&
                    End Select
                End If
            End If
        End If
    End If
    
        

End Function

Public Function ExportEffectsParameters(outArray() As Byte) As Boolean
    
    ' Routine exports effects parameters to a byte array
    ' The usercontrol and propertypage use this function
    ' You can also use it to store favorite parameter settings and upload
    '   into this class via the ImportEffectsParameters function
    
    Dim paramSize As Long, aPtr As Long
    Dim e As Long, aSize As Long, aCount As Long
    
    ' exported data follows this format
    ' 1st byte: number of effects
    ' for each effect
    ' 4 bytes: last byte is effect type, 1st 3 bytes is param size
    ' next n bytes is the param data
    
    For e = lvicBlurFX To lvicRedEyeCorrectionFX
        If m_hEffect(e) Then
            GdipGetEffectParameterSize m_hEffect(e), paramSize
            If paramSize > 0& And paramSize < &H1000000 Then
                aSize = aSize + paramSize
                aCount = aCount + 1&
            End If
        End If
    Next
    If aCount Then
        ReDim outArray(0 To aSize + aCount * 4&)
        outArray(0) = aCount: aPtr = 1&
        For e = lvicBlurFX To lvicRedEyeCorrectionFX
            If m_hEffect(e) Then
                GdipGetEffectParameterSize m_hEffect(e), paramSize
                If paramSize > 0& And paramSize < &H1000000 Then
                    GdipGetEffectParameters m_hEffect(e), paramSize, outArray(aPtr + 4&)
                    aSize = paramSize Or e * &H1000000
                    CopyMemory outArray(aPtr), aSize, 4&
                    aPtr = aPtr + 4& + paramSize
                End If
            End If
        Next
        ExportEffectsParameters = True
    End If
    
End Function

Public Function ImportEffectsParameters(inArray() As Byte) As Boolean

    If m_Token.Version > 1! Then
        
        Dim paramSize As Long, aPtr As Long, sEffect As String
        Dim e As Long, aSize As Long, aCount As Long
        Dim eData() As Byte, lMaxSize As Long
        
        ' imported data follows this format
        ' 1st byte: number of effects
        ' for each effect
        ' 4 bytes: last byte is effect type, 1st 3 bytes is param size
        ' next n bytes is the param data
        
        If modCommon.NormalizeArray(inArray(), eData()) = False Then Exit Function
        
        ' run a validation check first to ensure garbage wasn't passed
        lMaxSize = UBound(eData) + 1&
        aCount = eData(0)                                   ' number of effects
        If (aCount = 0& Or aCount > lvicRedEyeCorrectionFX) Then Exit Function
        If aCount * 12& > lMaxSize Then Exit Function       ' not enough bytes
        aPtr = 1&
        For e = 0& To aCount - 1&
            CopyMemory paramSize, eData(aPtr), 4&
            aSize = (paramSize And &HFFFFFF)
            If (aSize < 1& Or aSize > &HFFFFFF) Then Exit Function ' invalid size
            aPtr = aPtr + aSize + 4&
            If aPtr > lMaxSize Then Exit Function           ' not enough bytes
            If (paramSize \ &H1000000) > lvicRedEyeCorrectionFX Then Exit Function ' invalid FX type
        Next
        
        aPtr = 1&                                           ' ok, validation done, let's import
        For aCount = 0& To aCount - 1&
            CopyMemory aSize, eData(aPtr), 4&
            paramSize = (aSize And &HFFFFFF)
            e = (aSize \ &H1000000)
            If pvCreateNewEffect(e, sEffect) Then
                If GdipSetEffectParameters(m_hEffect(e), eData(aPtr + 4&), paramSize) = 0& Then RaiseEvent Changed(sEffect)
            End If
            aPtr = aPtr + (aSize And &HFFFFFF) + 4&
        Next
        ImportEffectsParameters = True
    
    End If

End Function

Public Sub DestroyEffectsHandle(Optional ByVal FXtype As EffectsEnum = -1&)
    
    ' simply destroy any created effects
    If FXtype = -1& Then
        For FXtype = lvicBlurFX To lvicRedEyeCorrectionFX
            If m_hEffect(FXtype) Then
                GdipDeleteEffect m_hEffect(FXtype)
                If Not (m_AttrFlag = &HFFFFFFFF) Then RaiseEvent Changed("Effect_" & CStr(FXtype))
            End If
        Next
        Erase m_hEffect()
    
    ' destroy by type
    ElseIf Not (FXtype < lvicBlurFX Or FXtype > lvicRedEyeCorrectionFX) Then
        If m_hEffect(FXtype) Then
            GdipDeleteEffect m_hEffect(FXtype)
            m_hEffect(FXtype) = 0&
        End If
    End If
End Sub

Public Sub ResetAllAttributes()
    ' resets the attributes containd in this class, does not effect the Effects
    Me.BlendPct = 0&
    Me.GlobalTransparencyPct = 0&
    Me.GrayScale = lvicNoGrayScale
    Me.LightnessPct = 0&
    Me.TransparentColorUsed = False
    Me.Invert = False
    m_TransColor = 0&: m_BlendColor = 0&
End Sub

Public Sub DestroyAttributesHandle()
    ' destroy the attributes handle
    If m_hAttributes Then
        GdipDisposeImageAttributes m_hAttributes
        m_hAttributes = 0&
        If (m_AttrFlag And &H80000000) = 0& Then Call Me.ResetAllAttributes
    End If
    m_AttrFlag = (m_AttrFlag And Not &H80000000)
End Sub

Private Function pvCreateNewEffect(FXtype As EffectsEnum, EffectName As String) As Boolean

    ' helper function used by each the Create_xxxx routines & ImportEffectsParameters
    If m_hEffect(FXtype) Then
        pvCreateNewEffect = True
        
    ElseIf m_Token.Token > 1! Then
        Dim GUID(0 To 3) As Long
        Select Case FXtype
            Case lvicBlurFX
                GUID(0) = &H633C80A4: GUID(1) = &H482B1843: GUID(2) = &H28BEF29E: GUID(3) = &HD4FDC534
            Case lvicSharpenFX
                GUID(0) = &H63CBF3EE: GUID(1) = &H402CC526: GUID(2) = &HC562718F: GUID(3) = &H4251BF40
            Case lvicTintFX
                GUID(0) = &H1077AF00: GUID(1) = &H44412848: GUID(2) = &HAD448994: GUID(3) = &H2C7A2D4C
            Case lvicBrightnessContrastFX
                GUID(0) = &HD3A1DBE1: GUID(1) = &H4C178EC4: GUID(2) = &H97EA4C9F: GUID(3) = &H3D341CAD
            Case lvicHueSaturationLightnessFX
                GUID(0) = &H8B2DD6C3: GUID(1) = &H4D87EB07: GUID(2) = &H871F0A5: GUID(3) = &H5F9C6AE2
            Case lvicColorBalanceFX
                GUID(0) = &H537E597D: GUID(1) = &H48DA251E: GUID(2) = &HCA296496: GUID(3) = &HF8706B49
            Case lvicToneLevelsFX
                GUID(0) = &H99C354EC: GUID(1) = &H4F3A2A31: GUID(2) = &HA817348C: GUID(3) = &H253AB303
            Case lvicColorCurveFX
                GUID(0) = &HDD6A0022: GUID(1) = &H4A6758E4: GUID(2) = &H8ED49B9D: GUID(3) = &H3DA581B8
            Case lvicRedEyeCorrectionFX
                GUID(0) = &H74D29D05: GUID(1) = &H426669A4: GUID(2) = &HC53C4995: GUID(3) = &H32B63628
            Case Else
                Exit Function
        End Select
        pvCreateNewEffect = (GdipCreateEffect(GUID(0), GUID(1), GUID(2), GUID(3), m_hEffect(FXtype)) = 0&)
    End If
    EffectName = "Effect_" & CStr(FXtype)

End Function

Private Sub Class_Initialize()
    If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
    Set m_Token = g_TokenClass
    m_Token.AddRef
End Sub

Private Sub Class_Terminate()
    m_AttrFlag = m_AttrFlag Or &H80000000   ' prevent firing change events
    DestroyAttributesHandle
    m_AttrFlag = &HFFFFFFFF                 ' prevent firing change events
    DestroyEffectsHandle
    m_Token.RemoveRef
    Set m_Token = Nothing
End Sub
