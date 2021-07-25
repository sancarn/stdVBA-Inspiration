VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWithTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Names$(), m_Count&, m_Level&

Public Sub BeginWithStatement(Name As String)
    
    ' called as each new With Event statement is encountered
    
    ' ensure m_Names sized as needed
    If m_Count = 0 Then
        ReDim m_Names(0 To 2)
    ElseIf m_Count > UBound(m_Names) Then
        ReDim Preserve m_Names(0 To m_Count + 2)
    End If
    ' m_Count increases, but m_Level does not, not yet
    ' it is possible to have nested With Statements (see notes in EndWithStatement)
    m_Names(m_Count) = Name: m_Count = m_Count + 1
    
End Sub

Public Sub EndWithStatement()

    ' called after a With Statement is completely processed
    ' it is possible to have nested With statements before statement
    '   is completely processed.
    
    ' For example, in the following code, .Item(x) belongs to myCollection.
    ' When "With ListView1..." is encountered, can't make it the current With
    ' statement simply based on the start of that statement.
    ' Need to wait until it is fully parsed.
    '    ------------------------------------------------------
    '    With myCollection
    '        For x = 1 To .Count
    '            With ListView1.ListItems(.Item(x))
    '            ...
    '            End With
    '        Next
    '    End With
    '    ------------------------------------------------------
    
    m_Level = m_Level + 1   ' at this point, m_Level & m_Count should be equal
    
End Sub

Public Sub EndWithBlock()

    ' called when End With statement is encountered

    m_Count = m_Count - 1           ' reduce count
    If m_Count = 0 Then             ' when zero, all nested With blocks are done
        m_Count = 0: m_Level = 0
        Erase m_Names()
    Else
        m_Level = m_Level - 1       ' else move to previous With statement
    End If
End Sub

Public Property Get WithName() As String

    ' returns the active With name
    On Error Resume Next ' can error when source has syntax errors
    WithName = m_Names(m_Level - 1)
    If Err Then
        Err.Clear
        WithName = "{ERR}"  ' return something
    End If
    On Error GoTo 0
    
End Property

Public Sub Reset()
    m_Count = 0: m_Level = 0
End Sub
