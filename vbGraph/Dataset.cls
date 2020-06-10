VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mobjPoints   As Points
Attribute mobjPoints.VB_VarHelpID = -1

Private mudtProps       As gtypDatasetProps

Private mlngParentPtr   As Long

Private Sub Class_Initialize()
    Set mobjPoints = New Points
    InitProperties
End Sub

Private Sub Class_Terminate()
    Set mobjPoints = Nothing
End Sub

Friend Sub Initialise(ByRef pobjParent As Datasets)
    mlngParentPtr = ObjPtr(pobjParent)
End Sub

Private Sub mobjPoints_Changed()
    StateChanged
End Sub

Friend Property Get Datasets() As Datasets
    Set Datasets = PtrObj(mlngParentPtr)
End Property

Public Property Get Points() As Points
    Set Points = mobjPoints
End Property

Private Property Let State(ByVal Value As String)
Dim udtData     As gtypDatasetData
    udtData.Data = Value
    LSet mudtProps = udtData
End Property

Private Property Get State() As String
Dim udtData     As gtypDatasetData
    LSet udtData = mudtProps
    State = udtData.Data
End Property

Friend Property Let SuperState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        State = .ReadProperty(PB_STATE)
        mobjPoints.SuperState = .ReadProperty(PB_POINTS)
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_STATE, State
        .WriteProperty PB_POINTS, mobjPoints.SuperState
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Private Sub InitProperties()
    With mudtProps
        .Visible = True
        .LineColor = RGB(84, 131, 169)
        .BarColor = &HDEA68D
        .PointColor = &HFFFFC0
        .CapColor = RGB(0, 0, 255)
        .ShowLines = True
        .ShowPoints = True
        .ShowBars = True
        .ShowCaps = False
    End With
End Sub

Friend Sub StateChanged()
    Datasets.StateChanged
End Sub

Public Property Let Visible(ByVal Value As Boolean)
    If Value = Not mudtProps.Visible Then
        mudtProps.Visible = Value
        StateChanged
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = mudtProps.Visible
End Property

Public Property Let LineColor(ByVal Value As OLE_COLOR)
    If Value <> mudtProps.LineColor Then
        mudtProps.LineColor = Value
        StateChanged
    End If
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = mudtProps.LineColor
End Property

Public Property Let BarColor(ByVal Value As OLE_COLOR)
    If Value <> mudtProps.BarColor Then
        mudtProps.BarColor = Value
        StateChanged
    End If
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = mudtProps.BarColor
End Property

Public Property Let PointColor(ByVal Value As OLE_COLOR)
    If Value <> mudtProps.PointColor Then
        mudtProps.PointColor = Value
        StateChanged
    End If
End Property

Public Property Get PointColor() As OLE_COLOR
    PointColor = mudtProps.PointColor
End Property

Public Property Let CapColor(ByVal Value As OLE_COLOR)
    If Value <> mudtProps.CapColor Then
        mudtProps.CapColor = Value
        StateChanged
    End If
End Property

Public Property Get CapColor() As OLE_COLOR
    CapColor = mudtProps.CapColor
End Property

Public Property Let ShowLines(ByVal Value As Boolean)
    If Value = Not mudtProps.ShowLines Then
        mudtProps.ShowLines = Value
        StateChanged
    End If
End Property

Public Property Get ShowLines() As Boolean
    ShowLines = mudtProps.ShowLines
End Property

Public Property Let ShowBars(ByVal Value As Boolean)
    If Value = Not mudtProps.ShowBars Then
        mudtProps.ShowBars = Value
        StateChanged
    End If
End Property

Public Property Get ShowBars() As Boolean
    ShowBars = mudtProps.ShowBars
End Property

Public Property Let ShowPoints(ByVal Value As Boolean)
    If Value = Not mudtProps.ShowPoints Then
        mudtProps.ShowPoints = Value
        StateChanged
    End If
End Property

Public Property Get ShowPoints() As Boolean
    ShowPoints = mudtProps.ShowPoints
End Property

Public Property Let ShowCaps(ByVal Value As Boolean)
    If Value = Not mudtProps.ShowCaps Then
        mudtProps.ShowCaps = Value
        StateChanged
    End If
End Property

Public Property Get ShowCaps() As Boolean
    ShowCaps = mudtProps.ShowCaps
End Property
