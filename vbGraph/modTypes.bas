Attribute VB_Name = "modTypes"
Option Explicit

Public Type gtypPointProps
    Value       As Double
End Type

Public Type gtypPointData
    Data        As String * 4
End Type

Public Type gtypGraphProps
    BackColor   As Long
    AxisColor   As Long
    GridColor   As Long
    FixedPoints As Long
    XGridInc    As Long
    YGridInc    As Double
    MaxValue    As Double
    MinValue    As Double
    FadeIn      As Boolean
    ShowGrid    As Boolean
    ShowAxis    As Boolean
    BarWidth    As Single
End Type

Public Type gtypGraphData
    Data    As String * 28
End Type

Public Type gtypControlProps
    Redraw      As Boolean
    Appearance  As Integer
    BorderStyle As Integer
End Type

Public Type gtypControlData
    Data        As String * 3
End Type

Public Type gtypDatasetProps
    Visible     As Boolean
    LineColor   As Long
    BarColor    As Long
    PointColor  As Long
    CapColor    As Long
    ShowLines   As Boolean
    ShowBars    As Boolean
    ShowPoints  As Boolean
    ShowCaps    As Boolean
End Type

Public Type gtypDatasetData
    Data        As String * 14
End Type
