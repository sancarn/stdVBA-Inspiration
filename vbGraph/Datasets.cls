VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Datasets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Changed()

Private mcolItems   As Collection

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Set mcolItems = Nothing
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Property Get Item(ByVal Index As Long) As Dataset
    Set Item = mcolItems.Item(Index)
End Property

Public Property Get Count() As Long
    Count = mcolItems.Count
End Property

Public Function Add() As Dataset
Dim objNew  As Dataset
    Set objNew = New Dataset
    With objNew
        .Initialise Me
    End With
    mcolItems.Add objNew
    Set Add = objNew
    Set objNew = Nothing
    RaiseEvent Changed
End Function

Public Sub Remove(ByVal Index As Long)
    mcolItems.Remove Index
    RaiseEvent Changed
End Sub

Public Sub Clear()
    Set mcolItems = New Collection
    RaiseEvent Changed
End Sub

Friend Property Let SuperState(ByRef Value() As Byte)
Dim lngCount    As Long
Dim lngIndex    As Long
Dim objNew      As Dataset
Dim objPB       As PropertyBag
    Clear
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        lngCount = .ReadProperty(PB_COUNT)
        For lngIndex = 1 To lngCount
            Set objNew = New Dataset
            objNew.Initialise Me
            objNew.SuperState = .ReadProperty(PB_ITEM & CStr(lngIndex))
            mcolItems.Add objNew
            Set objNew = Nothing
        Next lngIndex
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim lngCount    As Long
Dim objNew      As Dataset
Dim objPB       As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        For Each objNew In mcolItems
            lngCount = lngCount + 1
            .WriteProperty PB_ITEM + CStr(lngCount), objNew.SuperState
        Next objNew
        .WriteProperty PB_COUNT, lngCount
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Public Sub Save(ByVal Filename As String)
    If Len(Filename) > 0 Then
        If Dir(Filename) <> vbNullString Then
            Kill Filename
        End If
    End If
    SaveFile Filename, SuperState
End Sub

Public Sub Load(ByVal Filename As String)
    SuperState = GetFile(Filename)
End Sub

Friend Sub StateChanged()
    RaiseEvent Changed
End Sub
