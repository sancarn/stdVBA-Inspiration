VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Status(Msg As String)
    ' sent by clsProject & clsSourceFile during parsing
End Sub
Public Function ParsedBegin(ByVal bGroupProject As Boolean) As Boolean
    ' sent by clsProject when project is about to begin parsing
End Function
Public Sub ParseComplete()
    ' sent by clsProject when project has completed parsing
End Sub
Public Sub ParseError(ErrCode As ValidationConstants)
    ' sent by caller when an error occurs
End Sub
Public Sub ValidationBegin()

End Sub
Public Sub ValidationComplete(cValidation As clsValidation, lOptions As ValidationTypeEnum)

End Sub
Public Sub ReportComplete(ReportType As Long, lParam As Long)

End Sub
