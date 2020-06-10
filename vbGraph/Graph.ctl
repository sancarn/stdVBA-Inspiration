VERSION 5.00
Begin VB.UserControl Graph 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picDraw 'Note VBA doesn't have a PictureBox
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

Private WithEvents mobjSets      As Datasets
Attribute mobjSets.VB_VarHelpID = -1

Private mudtControlProps    As gtypControlProps
Private mudtGraphProps      As gtypGraphProps

Private mblnDesignMode  As Boolean

Public Enum eBorderStyle
   egrNone = 0
   egrFixedSingle = 1
End Enum

Public Enum eAppearance
   egrFlat = 0
   egr3D = 1
End Enum

Private Type mtypPOINT
    X   As Long
    Y   As Long
End Type

Private Type mtypRECT
    Left    As Long
    Right   As Long
    Top     As Long
    Bottom  As Long
End Type

Private Sub UserControl_Initialize()
    picDraw.FillStyle = vbFSSolid
    Set mobjSets = New Datasets
End Sub

Private Sub UserControl_InitProperties()
    InitProperties
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
    PropertyChanged PB_STATE
End Sub

Private Sub UserControl_Paint()
    DrawGraph
End Sub

Private Sub UserControl_Terminate()
    Set mobjSets = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    State = PropBag.ReadProperty(PB_STATE, State)
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PB_STATE, State
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    With UserControl
        picDraw.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
    Refresh
End Sub

Private Sub mobjSets_Changed()
Static blnWorking As Boolean
    If Not blnWorking Then
        blnWorking = True
        RemovePoints
        If Not mblnDesignMode Then
            DrawGraph
        End If
        blnWorking = False
    End If
End Sub

Private Property Let GraphState(ByRef Value() As Byte)
Dim udtData     As gtypGraphData
    udtData.Data = Value
    LSet mudtGraphProps = udtData
End Property

Private Property Get GraphState() As Byte()
Dim udtData     As gtypGraphData
    LSet udtData = mudtGraphProps
    GraphState = udtData.Data
End Property

Friend Property Let ControlState(ByRef Value() As Byte)
Dim udtData     As gtypControlData
    udtData.Data = Value
    LSet mudtControlProps = udtData
End Property

Friend Property Get ControlState() As Byte()
Dim udtData     As gtypControlData
    LSet udtData = mudtControlProps
    ControlState = udtData.Data
End Property

Private Property Let State(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        ControlState = .ReadProperty(PB_CONTROL)
        GraphState = .ReadProperty(PB_GRAPH)
    End With
    Set objPB = Nothing
End Property

Private Property Get State() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_CONTROL, ControlState
        .WriteProperty PB_GRAPH, GraphState
        State = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let SuperState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        State = .ReadProperty(PB_STATE, State)
        mobjSets.SuperState = .ReadProperty(PB_DATASETS, mobjSets.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_STATE, State
        .WriteProperty PB_DATASETS, mobjSets.SuperState
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let FileState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        GraphState = .ReadProperty(PB_GRAPH, GraphState)
        mobjSets.SuperState = .ReadProperty(PB_POINTS, mobjSets.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get FileState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_GRAPH, GraphState
        .WriteProperty PB_POINTS, mobjSets.SuperState
        FileState = .Contents
    End With
    Set objPB = Nothing
End Property


Private Sub InitProperties()
    With mudtGraphProps
        .BackColor = RGB(255, 255, 255)
        .AxisColor = RGB(0, 0, 0)
        .GridColor = RGB(223, 223, 223)
        .FixedPoints = 20
        .XGridInc = 1
        .YGridInc = 10
        .MaxValue = 100
        .MinValue = 0
        .FadeIn = False
        .ShowGrid = True
        .ShowAxis = False
        .BarWidth = 0.8
    End With
    With mudtControlProps
        .Redraw = True
        .BorderStyle = eBorderStyle.egrFixedSingle
        .Appearance = eAppearance.egr3D
    End With
End Sub

Public Property Get Datasets() As Datasets
    Set Datasets = mobjSets
End Property

Public Property Let Redraw(ByVal Value As Boolean)
    mudtControlProps.Redraw = Value
    If Value Then
        Refresh
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mudtControlProps.Redraw
End Property

Public Property Let Appearance(ByVal Value As eAppearance)
    mudtControlProps.Appearance = Value
    UserControl.Appearance = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get Appearance() As eAppearance
    Appearance = mudtControlProps.Appearance
End Property

Public Property Let BorderStyle(ByVal Value As eBorderStyle)
    mudtControlProps.BorderStyle = Value
    UserControl.BorderStyle = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = mudtControlProps.BorderStyle
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.BackColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mudtGraphProps.BackColor
End Property

Public Property Let AxisColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.AxisColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get AxisColor() As OLE_COLOR
    AxisColor = mudtGraphProps.AxisColor
End Property

Public Property Let GridColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.GridColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = mudtGraphProps.GridColor
End Property

Public Property Let FixedPoints(ByVal Value As Long)
    mudtGraphProps.FixedPoints = Value
    RemovePoints
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FixedPoints() As Long
    FixedPoints = mudtGraphProps.FixedPoints
End Property

Public Property Let XGridInc(ByVal Value As Long)
    mudtGraphProps.XGridInc = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get XGridInc() As Long
    XGridInc = mudtGraphProps.XGridInc
End Property

Public Property Let YGridInc(ByVal Value As Double)
    mudtGraphProps.YGridInc = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get YGridInc() As Double
    YGridInc = mudtGraphProps.YGridInc
End Property

Public Property Let MaxValue(ByVal Value As Double)
    mudtGraphProps.MaxValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MaxValue() As Double
    MaxValue = mudtGraphProps.MaxValue
End Property

Public Property Let MinValue(ByVal Value As Double)
    mudtGraphProps.MinValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MinValue() As Double
    MinValue = mudtGraphProps.MinValue
End Property

Public Property Let ShowGrid(ByVal Value As Boolean)
    mudtGraphProps.ShowGrid = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowGrid() As Boolean
    ShowGrid = mudtGraphProps.ShowGrid
End Property

Public Property Let ShowAxis(ByVal Value As Boolean)
    mudtGraphProps.ShowAxis = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowAxis() As Boolean
    ShowAxis = mudtGraphProps.ShowAxis
End Property

Public Property Let FadeIn(ByVal Value As Boolean)
    mudtGraphProps.FadeIn = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FadeIn() As Boolean
    FadeIn = mudtGraphProps.FadeIn
End Property

Public Property Let BarWidth(ByVal Value As Single)
    mudtGraphProps.BarWidth = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BarWidth() As Single
    BarWidth = mudtGraphProps.BarWidth
End Property

Private Sub AddDefaultPoints()
    mobjSets.Clear
    mobjSets.Add
    AddDefaultPoint 80
    AddDefaultPoint 10
    AddDefaultPoint 70
    AddDefaultPoint 25
    AddDefaultPoint 50
    AddDefaultPoint 45
    AddDefaultPoint 15
    AddDefaultPoint 85
    AddDefaultPoint 5
    AddDefaultPoint 75
    AddDefaultPoint 65
End Sub

Private Sub AddDefaultPoint(ByVal plngPercent As Long)
    mobjSets.Item(1).Points.Add (plngPercent / 100) * (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) + mudtGraphProps.MinValue
End Sub

Private Sub RemovePoints()
Dim objDataset  As Dataset
    If mudtGraphProps.FixedPoints > 0 Then
        For Each objDataset In mobjSets
            Do While objDataset.Points.Count > mudtGraphProps.FixedPoints
                objDataset.Points.Remove 1
            Loop
        Next objDataset
    End If
End Sub

Public Sub Refresh()
    DrawGraph
End Sub

Public Sub DrawControl()
    With UserControl
        .Appearance = mudtControlProps.Appearance
        .BorderStyle = mudtControlProps.BorderStyle
    End With
End Sub

Private Sub DrawGraph()
Dim lngX        As Long
Dim lngY        As Long

Dim lngStepX    As Long
Dim lngStepY    As Long
Dim lngWidth    As Long
Dim lngHeight   As Long
Dim lngIndex    As Long
Dim udtPoints() As mtypPOINT
Dim lngYAxis    As Long
Dim lngBarWidth As Long
Dim lngFixedCount   As Long
Dim udtBar      As mtypRECT
Dim udtGrid     As mtypRECT

Dim objDataset  As Dataset
Dim objPoint    As Point
Dim objPoints   As Points
Dim blnDraw     As Boolean

    If UserControl.Height > 0 And UserControl.Width > 0 Then
    If mudtControlProps.Redraw Or mblnDesignMode Then
        If mblnDesignMode Then
            AddDefaultPoints
        End If
        With picDraw
            .Cls
            .BackColor = mudtGraphProps.BackColor

            lngWidth = .ScaleWidth - 15
            lngHeight = .ScaleHeight - 15

            'draw grid
            
            lngFixedCount = GetMaxPointCount
            
            
            

            With udtGrid
                .Left = 0
                .Top = 0
                .Right = lngWidth
                .Bottom = lngHeight
            End With
            
            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
                lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
            End If
            
            DrawGrid udtGrid, mudtGraphProps.GridColor, lngBarWidth
            
            For Each objDataset In mobjSets
            
                Set objPoints = objDataset.Points
                
                If objPoints.Count > 0 And objDataset.Visible Then
                    
                    If lngFixedCount > 0 Then
                        If objDataset.ShowBars Then
                            If objPoints.Count > lngFixedCount Or Not mudtGraphProps.FadeIn Then
                                lngBarWidth = CLng((lngWidth / lngFixedCount) * mudtGraphProps.BarWidth)
                            Else
                                lngBarWidth = CLng((lngWidth / objPoints.Count) * mudtGraphProps.BarWidth)
                            End If
                        End If
                    End If
                    
                    
                    udtPoints = GetPoints(objPoints, udtGrid, lngBarWidth)
        
                    
        
                    
        
                    'drawlines and bars
                    If objDataset.ShowLines Or objDataset.ShowBars Or objDataset.ShowCaps Then
                        For lngIndex = 1 To UBound(udtPoints)
                                udtBar.Left = udtPoints(lngIndex).X - (lngBarWidth / 2)
                                udtBar.Right = udtPoints(lngIndex).X + (lngBarWidth / 2)
                                udtBar.Top = udtPoints(lngIndex).Y
                                udtBar.Bottom = lngYAxis
                            If objDataset.ShowBars Then
                                DrawBar udtBar, objDataset.BarColor
                            End If
                            If objDataset.ShowLines And lngIndex > 1 Then
                                DrawLine udtPoints(lngIndex - 1), udtPoints(lngIndex), objDataset.LineColor
                            End If
                            If objDataset.ShowCaps Then
                
                                lngY = (lngYAxis - objDataset.Points.Item(lngIndex).Value * ((lngHeight) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)))
                                
                                udtBar.Bottom = udtBar.Top
                                udtBar.Top = lngY
                                udtBar.Bottom = udtBar.Top + 30
                                DrawBar udtBar, objDataset.CapColor
                           End If
                        Next lngIndex
                    End If
        
                    'draw axis
                    If mudtGraphProps.ShowAxis Then
                        picDraw.Line (0, 0)-(0, lngHeight), mudtGraphProps.AxisColor
                        If mudtGraphProps.MaxValue <= 0 Then
                            picDraw.Line (0, 0)-(lngWidth, 0), mudtGraphProps.AxisColor
                        ElseIf mudtGraphProps.MinValue < 0 Then
                            picDraw.Line (0, (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue)-(lngWidth, (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue), mudtGraphProps.AxisColor
                        Else
                            picDraw.Line (0, lngHeight)-(lngWidth, lngHeight), mudtGraphProps.AxisColor
                        End If
                    End If
        
                    'draw points
                    If objDataset.ShowPoints Then
                        For lngIndex = 1 To UBound(udtPoints)
                            If lngIndex = 1 Then
                                blnDraw = True
                            Else
                                If mudtGraphProps.XGridInc = 0 Then
                                    blnDraw = True
                                Else
                                    blnDraw = (lngIndex Mod mudtGraphProps.XGridInc = 0)
                                End If
                            End If
                            If blnDraw Then
                                DrawPoint udtPoints(lngIndex), objDataset.PointColor
                            End If
                        Next lngIndex
                    End If
                End If
                Set objPoints = Nothing
            Next objDataset
            'copy picture to usercontrol
            BitBlt UserControl.hDC, 0, 0, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, SRCCOPY
        End With
    End If
    End If
End Sub

Private Function GetPoints(ByRef pobjPoints As Points, ByRef pudtGrid As mtypRECT, ByVal plngBarWidth As Long) As mtypPOINT()
Dim udtPoints() As mtypPOINT
Dim lngCount    As Long
Dim lngIndex    As Long
Dim objPoint    As Point
Dim lngX        As Long
Dim lngPtCount  As Long
Dim lngFixedCount   As Long
Dim lngYAxis    As Long
    lngCount = pobjPoints.Count
    If lngCount > 0 Then
        If mudtGraphProps.FixedPoints = 0 Or lngCount < mudtGraphProps.FixedPoints Then
            lngPtCount = lngCount
            If mudtGraphProps.FixedPoints > 0 Then
                lngFixedCount = mudtGraphProps.FixedPoints
            Else
                lngFixedCount = lngCount
            End If
        Else
            lngPtCount = mudtGraphProps.FixedPoints
            lngFixedCount = mudtGraphProps.FixedPoints
        End If
        ReDim udtPoints(lngPtCount) As mtypPOINT

        For Each objPoint In pobjPoints
            lngIndex = lngIndex + 1
            If mudtGraphProps.FixedPoints > 0 And lngIndex > mudtGraphProps.FixedPoints Then
                Set objPoint = Nothing
                Exit For
            End If

            If lngIndex = 1 Then
                If lngFixedCount = 1 Then
                    lngX = pudtGrid.Left + (((pudtGrid.Right - pudtGrid.Left)) / 2)
                Else
                    lngX = pudtGrid.Left + (plngBarWidth / 2)
                End If
            ElseIf lngIndex = lngFixedCount Then
                lngX = pudtGrid.Right - (plngBarWidth / 2)
            Else
                lngX = (lngIndex - 1) * (((pudtGrid.Right - pudtGrid.Left) - plngBarWidth) / (lngFixedCount - 1)) + (plngBarWidth / 2)
            End If

            udtPoints(lngIndex).X = lngX
            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) <> 0 Then
                lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
                udtPoints(lngIndex).Y = lngYAxis - objPoint.Value * ((pudtGrid.Bottom - pudtGrid.Top) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue))
            End If
        Next objPoint
    End If
    GetPoints = udtPoints
End Function

Private Sub DrawLine(ByRef pudtPt1 As mtypPOINT, ByRef pudtPt2 As mtypPOINT, ByVal plngColor As String)
    picDraw.Line (pudtPt1.X, pudtPt1.Y)-(pudtPt2.X, pudtPt2.Y), plngColor
End Sub

Private Sub DrawPoint(ByRef pudtPt As mtypPOINT, ByVal plngColor As Long)
    picDraw.FillColor = plngColor
    picDraw.Circle (pudtPt.X, pudtPt.Y), 40, 0
End Sub

Private Sub DrawBar(ByRef pudtRect As mtypRECT, ByVal plngColor As Long)
    picDraw.FillColor = plngColor
    With pudtRect
        picDraw.Line (.Left, .Top)-(.Right, .Bottom), 0, B
    End With
End Sub

Private Sub DrawGrid(ByRef pudtRect As mtypRECT, ByVal plngColor As Long, ByVal plngBarWidth As Long)
Dim lngCount    As Long
Dim lngIndex    As Long
Dim lngX        As Long
Dim lngY        As Long
Dim lngFixedCount   As Long
Dim lngYAxis    As Long
Dim lngStepY    As Long
Dim lngHeight   As Long
    lngCount = GetMaxPointCount
    lngFixedCount = lngCount
    If lngCount > 0 And mudtGraphProps.ShowGrid Then

        lngHeight = picDraw.ScaleHeight - 15

        If mudtGraphProps.XGridInc > 0 Then
            For lngIndex = 1 To lngFixedCount
                If lngIndex Mod mudtGraphProps.XGridInc = 0 Then
                    If lngIndex = 1 Then
                        If lngFixedCount = 1 Then
                            lngX = pudtRect.Left + (((pudtRect.Right - pudtRect.Left)) / 2)
                        Else
                            lngX = pudtRect.Left + (plngBarWidth / 2)
                        End If
                    ElseIf lngIndex = lngFixedCount Then
                        lngX = pudtRect.Right - (plngBarWidth / 2)
                    Else
                        lngX = (lngIndex - 1) * (((pudtRect.Right - pudtRect.Left) - plngBarWidth) / (lngFixedCount - 1)) + (plngBarWidth / 2)
                    End If
                    picDraw.Line (lngX, 0)-(lngX, lngHeight), plngColor
                End If
            Next lngIndex
        End If

        'draw horizontal lines
        If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
            lngYAxis = ((picDraw.ScaleHeight - 15) / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.MaxValue
            lngStepY = (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.YGridInc
            If lngStepY > 0 Then
                For lngY = lngYAxis To 0 Step -lngStepY
                    picDraw.Line (0, lngHeight - lngY)-(picDraw.ScaleWidth - 15, lngHeight - lngY), mudtGraphProps.GridColor
                Next lngY
                
                For lngY = lngYAxis To lngHeight Step lngStepY
                    picDraw.Line (0, lngHeight - lngY)-(picDraw.ScaleWidth - 15, lngHeight - lngY), mudtGraphProps.GridColor
                Next lngY
            End If
        End If
    End If
End Sub

Private Function GetMaxPointCount() As Long
Dim objDataset  As Dataset
Dim lngMaxCount    As Long
Dim lngPtCount  As Long
    If mudtGraphProps.FixedPoints = 0 Then
        For Each objDataset In mobjSets
            lngPtCount = objDataset.Points.Count
            If lngPtCount > lngMaxCount Then
                lngMaxCount = lngPtCount
            End If
        Next objDataset
    Else
        lngMaxCount = mudtGraphProps.FixedPoints
    End If
    GetMaxPointCount = lngMaxCount
End Function

Public Sub SaveSettings(ByVal Filename As String)
    If Len(Filename) > 0 Then
        If Dir(Filename) <> vbNullString Then
            Kill Filename
        End If
    End If
    SaveFile Filename, FileState
End Sub

Public Sub LoadSettings(ByVal Filename As String)
    FileState = GetFile(Filename)
    Refresh
End Sub

