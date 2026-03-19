Attribute VB_Name = "VBAStack_Extended"
Option Explicit

#If VBA7 = False Then
    Private Enum LongPtr
        [_]
    End Enum
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As LongPtr, ByVal cbCopy As Long)
#Else
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As LongPtr, ByVal cbCopy As Long)
#End If

#If Win64 Then
    Const PtrSize As Integer = 8
#Else
    Const PtrSize As Integer = 4
#End If

'=====================================================
' Structures
'=====================================================

Public Type StackFrame
    ProjectName As String
    ObjectName As String
    ProcedureName As String
    FrameNumber As Integer
    RealFrameNumber As Integer
    ModuleType As String
    flags As Long
    Errored As Boolean
End Type

'=====================================================
' Core API
'=====================================================

Public Function FrameCount() As Integer
On Error GoTo fail

    FrameCount = -1

    Dim errObj As LongPtr
    errObj = ObjPtr(VBA.Err)

    Dim g_ebThread As LongPtr
    CopyMemory g_ebThread, (errObj + PtrSize * 6), PtrSize
    If g_ebThread = 0 Then GoTo fail

    Dim g_ExFrameTOS As LongPtr
    #If Win64 Then
        g_ExFrameTOS = g_ebThread + &H10
    #Else
        g_ExFrameTOS = g_ebThread + &HC
    #End If
    If g_ExFrameTOS = 0 Then GoTo fail

    Dim pExFrame As LongPtr
    CopyMemory pExFrame, g_ExFrameTOS, PtrSize
    If pExFrame = 0 Then GoTo fail

    Do
        CopyMemory pExFrame, pExFrame, PtrSize
        FrameCount = FrameCount + 1
        If pExFrame = 0 Then Exit Do
    Loop

    Exit Function
fail:
End Function

Public Function GetCurrentProcedure() As StackFrame
    GetCurrentProcedure = GetStackFrame(2)
End Function

Public Function GetCallstack() As StackFrame()
On Error GoTo fail

    Dim count As Integer
    count = FrameCount
    If count < 2 Then GoTo fail

    Dim arr() As StackFrame
    ReDim arr(count - 2)

    Dim idx As Integer
    For idx = 1 To count - 1
        arr(idx - 1) = GetStackFrame(idx + 1)
    Next
    GetCallstack = arr
    Exit Function

fail:
    Dim emptyArr(0 To 0) As StackFrame
    emptyArr(0).Errored = True
    GetCallstack = emptyArr
End Function

Public Function GetStackFrame(Optional ByVal FrameNumber As Integer = 1) As StackFrame
On Error GoTo fail

    Dim f As StackFrame
    f.RealFrameNumber = FrameNumber
    f.FrameNumber = FrameNumber - 1

    Dim errObj As LongPtr
    errObj = ObjPtr(VBA.Err)

    Dim g_ebThread As LongPtr
    CopyMemory g_ebThread, (errObj + PtrSize * 6), PtrSize
    If g_ebThread = 0 Then GoTo fail

    Dim g_ExFrameTOS As LongPtr
    #If Win64 Then
        g_ExFrameTOS = g_ebThread + &H10
    #Else
        g_ExFrameTOS = g_ebThread + &HC
    #End If

    Dim pExFrame As LongPtr
    CopyMemory pExFrame, g_ExFrameTOS, PtrSize
    If pExFrame = 0 Then GoTo fail

    Do
        CopyMemory pExFrame, pExFrame, PtrSize
        If pExFrame = 0 Then GoTo fail
        FrameNumber = FrameNumber - 1
    Loop Until FrameNumber = 0

    Dim pRTMI As LongPtr
    CopyMemory pRTMI, (pExFrame + PtrSize * 3), PtrSize
    If pRTMI = 0 Then GoTo fail

    Dim pObjectInfo As LongPtr
    CopyMemory pObjectInfo, pRTMI, PtrSize
    If pObjectInfo = 0 Then GoTo fail

    Dim pPublicObject As LongPtr
    CopyMemory pPublicObject, (pObjectInfo + PtrSize * 6), PtrSize
    If pPublicObject = 0 Then GoTo fail

    '----- Object Name -----
    Dim pObjName As LongPtr
    CopyMemory pObjName, (pPublicObject + PtrSize * 6), PtrSize
    f.ObjectName = ReadCString(pObjName)

    '----- Method enumeration -----
    Dim pMethodsArr As LongPtr
    CopyMemory pMethodsArr, (pObjectInfo + PtrSize * 9), PtrSize

    Dim methodCount As Long
    CopyMemory methodCount, (pPublicObject + PtrSize * 7), 4

    Dim pMethodRTMI As LongPtr
    Dim methodIdx As Integer: methodIdx = -1
    Dim i As Integer
    For i = methodCount - 1 To 0 Step -1
        CopyMemory pMethodRTMI, (pMethodsArr + PtrSize * i), PtrSize
        If pMethodRTMI = pRTMI Then
            methodIdx = i
            Exit For
        End If
    Next
    If methodIdx = -1 Then GoTo fail

    Dim pMethodNames As LongPtr
    CopyMemory pMethodNames, (pPublicObject + PtrSize * 8), PtrSize

    Dim pMethodName As LongPtr
    CopyMemory pMethodName, (pMethodNames + PtrSize * methodIdx), PtrSize
    f.ProcedureName = ReadCString(pMethodName)

    '----- Project -----
    Dim pObjectTable As LongPtr
    CopyMemory pObjectTable, (pObjectInfo + PtrSize * 1), PtrSize

    Dim pProjName As LongPtr
    #If Win64 Then
        CopyMemory pProjName, (pObjectTable + &H68), PtrSize
    #Else
        CopyMemory pProjName, (pObjectTable + &H40), PtrSize
    #End If
    f.ProjectName = ReadCString(pProjName)

    '----- Module type (rough heuristic) -----
    Dim flags As Long
    CopyMemory flags, (pObjectInfo + PtrSize * 7), 4
    f.flags = flags
    f.ModuleType = ResolveModuleType(flags)

    GetStackFrame = f
    Exit Function

fail:
    f.Errored = True
    GetStackFrame = f
End Function

'=====================================================
' Helpers
'=====================================================

Private Function ReadCString(ByVal pStr As LongPtr) As String
On Error GoTo fail
    Dim ch As Byte, s As String
    Do
        CopyMemory ch, pStr, 1
        pStr = pStr + 1
        If ch = 0 Then Exit Do
        s = s & Chr$(ch)
    Loop
    ReadCString = s
    Exit Function
fail:
End Function

Private Function ResolveModuleType(ByVal flags As Long) As String
    Dim lowByte As Long: lowByte = flags And &HFF&
    Select Case lowByte
        Case &H50: ResolveModuleType = "UserForm"
        Case &H20: ResolveModuleType = "Workbook"
        Case &HF0: ResolveModuleType = "Worksheet"
        Case &H38: ResolveModuleType = "StdModule"
        Case &HB8: ResolveModuleType = "Class"
        Case Else: ResolveModuleType = "Unknown(" & Hex$(lowByte) & ")"
    End Select
End Function

'=====================================================
' Utility output
'=====================================================

Public Sub DumpCallStack(Optional ByVal ToImmediate As Boolean = True)
    Dim frames() As StackFrame
    frames = GetCallstack()

    Dim i As Long
    Dim f As StackFrame

    If ToImmediate Then Debug.Print "==== CALL STACK ===="
    For i = LBound(frames) To UBound(frames)
        f = frames(i)
        Dim msg As String
        msg = f.FrameNumber & ": " & f.ProjectName & "::" _
            & f.ObjectName & "::" & f.ProcedureName
        If f.ModuleType <> "" Then msg = msg & " [" & f.ModuleType & "]"
        If f.Errored Then msg = msg & " (Errored)"
        If ToImmediate Then
            Debug.Print msg
        Else
            MsgBox msg
        End If
    Next i
End Sub



'=====================================================
' Testing
'=====================================================

Public Sub TestVBAStack()
    'Entry point for validation
    Debug.Print ">>> Starting VBAStack test sequence"
    Call LevelOne
    Debug.Print ">>> End of test sequence"
End Sub

Private Sub LevelOne()
    LevelTwo
End Sub

Private Sub LevelTwo()
    LevelThree
End Sub

Private Sub LevelThree()
    Dim f As StackFrame
    f = GetCurrentProcedure()
    Debug.Print "Current procedure: " & f.ModuleType & "|" & f.ObjectName & "::" & f.ProcedureName
    
    DumpCallStack True
End Sub

Private Sub tt()
  Dim xx As Class1
  Set xx = New Class1
  xx.t
End Sub
Private Sub t0()
  Debug.Print GetStackFrame().ModuleType
End Sub
  
