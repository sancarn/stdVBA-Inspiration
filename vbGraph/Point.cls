VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps   As gtypPointProps

Private mlngParentPtr   As Long

Public Property Let Value(ByVal Value As Double)
    If Value <> mudtProps.Value Then
        mudtProps.Value = Value
        StateChanged
    End If
End Property

Public Property Get Value() As Double
    Value = mudtProps.Value
End Property

Friend Sub Initialise(ByRef pobjParent As Points)
    mlngParentPtr = ObjPtr(pobjParent)
End Sub

Friend Property Get Points() As Points
    Set Points = PtrObj(mlngParentPtr)
End Property

Private Property Let State(ByVal Value As String)
Dim udtData     As gtypPointData
    udtData.Data = Value
    LSet mudtProps = udtData
End Property

Private Property Get State() As String
Dim udtData     As gtypPointData
    LSet udtData = mudtProps
    State = udtData.Data
End Property

Friend Property Let SuperState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        State = .ReadProperty(PB_STATE)
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_STATE, State
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Sub StateChanged()
    Points.StateChanged
End Sub
