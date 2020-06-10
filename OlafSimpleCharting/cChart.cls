VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ForeColor As Long, LineGraphColor As Long
Public AxisAdjustmentFacX As Double, FormatStrX As String
Public AxisAdjustmentFacY As Double, FormatStrY As String
Public ShowMinIndicators As Boolean, MinIndicatorColor As Long
Public ShowMaxIndicators As Boolean, MaxIndicatorColor As Long
Public OffsL As Double, OffsR As Double, OffsT As Double, OffsB As Double
Public BSplineInterpolate As SplineInterpolationEnm

Private Sub Class_Initialize()
  OffsL = 50: OffsR = 50
  OffsT = 25: OffsB = 60
  AxisAdjustmentFacX = 1: AxisAdjustmentFacY = 1
  LineGraphColor = vbGreen
  ShowMinIndicators = True: MinIndicatorColor = vbMagenta
  ShowMaxIndicators = True: MaxIndicatorColor = vbRed
End Sub

Public Sub Render(CC As cCairoContext, Rs As cRecordset, Optional BackGround As cCairoPattern)
Dim x#, dx#, sx#, MinAxisX#, MaxAxisX#, StepWidthX#
Dim y#, dy#, sy#, MinAxisY#, MaxAxisY#, StepWidthY#

  If Not BackGround Is Nothing Then CC.Paint 1, BackGround 'draw a background onto the charting-area
  
  CC.SelectFont "Arial", 10, ForeColor
  
  With Rs.ActiveConnection.MemDB
    .Cnn.CreateTableFromRsContent Rs.Content, "Tmp"
    GetAxisScales .GetMin("Tmp", "AvgT"), .GetMax("Tmp", "AvgT"), MinAxisX, MaxAxisX, StepWidthX, AxisAdjustmentFacX
    GetAxisScales .GetMin("Tmp", IIf(ShowMinIndicators, "MinP", "AvgP")), _
                  .GetMax("Tmp", IIf(ShowMaxIndicators, "MaxP", "AvgP")), MinAxisY, MaxAxisY, StepWidthY, AxisAdjustmentFacY
  End With
  
  dx = CC.Surface.Width - OffsL - OffsR:  sx = dx / (MaxAxisX - MinAxisX)
  dy = CC.Surface.Height - OffsT - OffsB: sy = -dy / (MaxAxisY - MinAxisY)
  
  CC.TranslateDrawings Int(OffsL - MinAxisX * sx) + 0.5, Int(CC.Surface.Height - MinAxisY * sy - OffsB) + 0.5
 
  For x = MinAxisX To MaxAxisX + StepWidthX / 2 Step StepWidthX 'x-Axis-Grid (and -TextLabels)
    CC.DrawLine Int(x * sx), Int(MinAxisY * sy + 5), Int(x * sx), Int(MaxAxisY * sy), , 0.2, ForeColor, 0.5
    CC.DrawText (x - StepWidthX / 2) * sx, MinAxisY * sy, StepWidthX * sx, OffsB, Format$(x, FormatStrX), , vbCenter, 3, 1
  Next
  
  For y = MinAxisY To MaxAxisY + StepWidthY / 2 Step StepWidthY 'y-Axis-Grid (and -TextLabels)
    CC.DrawLine Int(MinAxisX * sx - 4), Int(y * sy), Int(MaxAxisX * sx), Int(y * sy), , 0.2, ForeColor, 0.5
    CC.DrawText MinAxisX * sx - OffsL - 8, y * sy - 15, OffsL, 30, Format$(y, FormatStrY), , vbRightJustify, 0, 1
  Next
  
  Dim i As Long, PMin() As Double, PAvg() As Double, PMax() As Double
  'fill three Poly-Point-Array with Double-Values from the Rs (one for the Min, Avg and Max within a group)
  FillPolyArrFromRs Rs, PMin, Rs.Fields("AvgT"), Rs.Fields("MinP"), sx, sy
  FillPolyArrFromRs Rs, PAvg, Rs.Fields("AvgT"), Rs.Fields("AvgP"), sx, sy
  FillPolyArrFromRs Rs, PMax, Rs.Fields("AvgT"), Rs.Fields("MaxP"), sx, sy
  
  'Draw the Min- and Max-Points as regular Polygons (the small stars)
  CC.SetLineWidth 1
  For i = 0 To UBound(PMin) Step 2
    If ShowMinIndicators Then CC.DrawRegularPolygon PMin(i), PMin(i + 1), 4, 5, splNone, 2
    If ShowMinIndicators Then CC.Stroke , Cairo.CreateSolidPatternLng(MinIndicatorColor)
    If ShowMaxIndicators Then CC.DrawRegularPolygon PMax(i), PMax(i + 1), 4, 4, splNone, 2
    If ShowMaxIndicators Then CC.Stroke , Cairo.CreateSolidPatternLng(MaxIndicatorColor)
    CC.Save 'draw dotted lines between the two Min-Max-Stars
      CC.SetDashes 0.5, 1, 3
      If ShowMinIndicators Then CC.DrawLine PAvg(i), PAvg(i + 1), PMin(i), PMin(i + 1), , 1, MinIndicatorColor, 0.3
      If ShowMaxIndicators Then CC.DrawLine PAvg(i), PAvg(i + 1), PMax(i), PMax(i + 1), , 1, MaxIndicatorColor, 0.3
    CC.Restore
  Next i
 
  'finally draw the Average-Values (of each group) as a BSpline-interpolated Curve
  CC.SetLineWidth 2
  CC.Polygon PAvg, False, BSplineInterpolate, True, True
  CC.Stroke , Cairo.CreateSolidPatternLng(LineGraphColor)
End Sub

Private Sub GetAxisScales(ByVal MinVal#, ByVal MaxVal#, MinAxis, MaxAxis, StepWidth, Optional Fac = 1)
  Const Eps# = 0.00000001
  StepWidth = 10 ^ Int(Log(Abs(MaxVal - MinVal + Eps)) / Log(10)) * Sgn(MaxVal - MinVal + Eps) / Fac
  If StepWidth < Eps Then StepWidth = Eps: MinAxis = MinVal - Eps / 2: MaxAxis = MaxVal + Eps / 2: Exit Sub

  CalcAxisBounds MinVal, MaxVal, MinAxis, MaxAxis, StepWidth
  If Int((MaxAxis - MinAxis) / StepWidth) > 10 Then StepWidth = CalcAxisBounds(MinVal, MaxVal, MinAxis, MaxAxis, StepWidth, 2)
  If Int((MaxAxis - MinAxis) / StepWidth) < 5 Then StepWidth = CalcAxisBounds(MinVal, MaxVal, MinAxis, MaxAxis, StepWidth, 0.5)
End Sub

Private Function CalcAxisBounds(ByVal MinVal#, ByVal MaxVal#, MinAxis, MaxAxis, StepWidth, Optional StepFac = 1)
  CalcAxisBounds = StepWidth * StepFac
  MinAxis = (Int(MinVal / CalcAxisBounds)) * CalcAxisBounds
  MaxAxis = (Int(MaxVal / CalcAxisBounds) + 1) * CalcAxisBounds
End Function

Private Sub FillPolyArrFromRs(Rs As cRecordset, P() As Double, FieldX As cField, FieldY As cField, sx, sy)
Dim i As Long: ReDim P(0 To Rs.RecordCount * 2 - 1)
  Rs.MoveFirst
  Do Until Rs.EOF
    P(i + 0) = Int(FieldX.Value * sx)
    P(i + 1) = Int(FieldY.Value * sy)
    Rs.MoveNext: i = i + 2
  Loop
End Sub
