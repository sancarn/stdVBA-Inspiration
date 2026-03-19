Attribute VB_Name = "VBAStack"
Option Explicit On

'Tested on x86 Access 2003, x86 Access 2013, x86 Access 365, x64 Access 2013, and x64 Access 365.

'Example use:

'    Private Sub Example()
'
'    Dim StackFrames() As VBAStack.StackFrame
'    StackFrames = VBAStack.GetCallstack()
'
'    Dim str As String
'    Dim i As Integer
'
'    For i = 0 To UBound(StackFrames)
'
'        str = str & StackFrames(i).FrameNumber & ", " & StackFrames(i).ProjectName & "::" & StackFrames(i).ObjectName & "::" & StackFrames(i).ProcedureName & vbCrLf
'
'    Next
'    MsgBox (str)
'
'    'Above outputs this:
'    ' 1, MyMod::Example
'    ' 2, MyMod::Sub2
'    ' 3, Form_Form1::Command0_Click
'
'    Dim frame As VBAStack.StackFrame
'    frame = VBAStack.GetCurrentProcedure
'
'    MsgBox (frame.ObjectName & "::" & frame.ProcedureName)
'    'Outputs this:
'    ' MyMod::Example
'
'    End Sub


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

Public Type StackFrame
    ProjectName As String
    ObjectName As String
    ProcedureName As String
    realFrameNumber As Integer
    FrameNumber As Integer
    Errored As Boolean
End Type

Public Function FrameCount() As Integer

On Error GoTo ErrorOccurred

    FrameCount = -1

    'Get ptr to VBA.Err
    Dim errObj As LongPtr
    errObj = ObjPtr(VBA.Err)
    
    
    'Get g_ebThread
    Dim g_ebThread As LongPtr
    CopyMemory g_ebThread, (errObj + PtrSize * 6), PtrSize
    If g_ebThread = 0 Then GoTo ErrorOccurred
    
    'Get g_ExFrameTOS
    Dim g_ExFrameTOS As LongPtr
    #If Win64 Then
        g_ExFrameTOS = g_ebThread + (&H10)
    #Else
        g_ExFrameTOS = g_ebThread + (&HC)
    #End If
    If g_ExFrameTOS = 0 Then GoTo ErrorOccurred
    
    
    'Get top ExFrame
    Dim pTopExFrame As LongPtr
    CopyMemory pTopExFrame, g_ExFrameTOS, PtrSize
    If pTopExFrame = 0 Then GoTo ErrorOccurred

    
    'Loop over frames to count
    Dim pExFrame As LongPtr: pExFrame = pTopExFrame
    Do
        CopyMemory pExFrame, pExFrame, PtrSize
        FrameCount = FrameCount + 1
        If pExFrame = 0 Then Exit Do
    Loop

Exit Function

ErrorOccurred:
    
End Function

Public Function GetCurrentProcedure() As StackFrame
    GetCurrentProcedure = VBAStack.GetStackFrame(2)
End Function

Public Function GetCallstack() As StackFrame()
    Dim stackCount As Integer: stackCount = VBAStack.FrameCount
    Dim index As Integer: index = 1
    Dim FrameArray() As StackFrame
    ReDim FrameArray(stackCount - 2)
    
    Do Until index = stackCount
                
        FrameArray(index - 1) = VBAStack.GetStackFrame(index + 1)
        index = index + 1
        
    Loop
    
    GetCallstack = FrameArray
End Function

Public Function GetStackFrame(Optional ByVal FrameNumber As Integer = 1) As StackFrame

On Error GoTo ErrorOccurred
    
    If FrameNumber < 1 Then GoTo ErrorOccurred
    
    Dim retVal As StackFrame
    retVal.realFrameNumber = FrameNumber
    retVal.FrameNumber = FrameNumber - 1
    
    'Get ptr to VBA.Err
    Dim errObj As LongPtr
    errObj = ObjPtr(VBA.Err)
    
    
    'Get g_ebThread
    Dim g_ebThread As LongPtr
    CopyMemory g_ebThread, (errObj + PtrSize * 6), PtrSize
    If g_ebThread = 0 Then GoTo ErrorOccurred
    
    
    'Get g_ExFrameTOS
    Dim g_ExFrameTOS As LongPtr
    #If Win64 Then
        g_ExFrameTOS = g_ebThread + (&H10)
    #Else
        g_ExFrameTOS = g_ebThread + (&HC)
    #End If
    If g_ExFrameTOS = 0 Then GoTo ErrorOccurred
    
    
    'Get top ExFrame
    Dim pTopExFrame As LongPtr
    CopyMemory pTopExFrame, g_ExFrameTOS, PtrSize
    If pTopExFrame = 0 Then GoTo ErrorOccurred


    'Get next ExFrame (do this minimum once, since top frame is this procedure)
    Dim pExFrame As LongPtr: pExFrame = pTopExFrame
    Do
        CopyMemory pExFrame, pExFrame, PtrSize
        If pExFrame = 0 Then GoTo ErrorOccurred
        FrameNumber = FrameNumber - 1
    Loop Until FrameNumber = 0
    
    
    'Get RTMI
    Dim pRTMI As LongPtr
    CopyMemory pRTMI, (pExFrame + PtrSize * 3), PtrSize
    If pRTMI = 0 Then GoTo ErrorOccurred
    
    
    'Get ObjectInfo
    Dim pObjectInfo As LongPtr
    CopyMemory pObjectInfo, pRTMI, PtrSize
    If pObjectInfo = 0 Then GoTo ErrorOccurred
    
    
    'Get Public Object Descriptor
    Dim pPublicObject As LongPtr
    CopyMemory pPublicObject, (pObjectInfo + PtrSize * 6), PtrSize
    If pPublicObject = 0 Then GoTo ErrorOccurred
    
    
    'Get pointer to module name string from Public Object Descriptor
    Dim pObjectName As LongPtr
    CopyMemory pObjectName, (pPublicObject + PtrSize * 6), PtrSize
    If pObjectName = 0 Then GoTo ErrorOccurred
    
    
    'Read the object name string
    Dim objName As String
    Dim readByteObjName As Byte
    Do
        CopyMemory readByteObjName, pObjectName, 1
        pObjectName = pObjectName + 1
        If readByteObjName = 0 Then Exit Do 'Read null char - end loop
        objName = objName & Chr(readByteObjName)
    Loop
    retVal.ObjectName = objName
    
    'Get pointer to methods array from ObjectInfo
    Dim pMethodsArr As LongPtr
    CopyMemory pMethodsArr, (pObjectInfo + PtrSize * 9), PtrSize
    If pMethodsArr = 0 Then GoTo ErrorOccurred
    
    
    'Get count of methods from Public Object Descriptor
    Dim methodCount As Long
    CopyMemory methodCount, (pPublicObject + PtrSize * 7), 4
    If methodCount = 0 Then GoTo ErrorOccurred
    
    
    'Search the method array to find our RTMI
    Dim methodIndex As Integer: methodIndex = -1
    Dim i As Integer
    Dim pMethodRTMI As LongPtr
    For i = methodCount - 1 To 0 Step -1
        CopyMemory pMethodRTMI, (pMethodsArr + PtrSize * i), PtrSize
        If pMethodRTMI = 0 Then GoTo ErrorOccurred
        If pMethodRTMI = pRTMI Then
            methodIndex = i
            Exit For
        End If
    Next
    
    If methodIndex = -1 Then GoTo ErrorOccurred
    
    
    'Get array of method names from Public Object Descriptor
    Dim pMethodNamesArr As LongPtr
    CopyMemory pMethodNamesArr, (pPublicObject + PtrSize * 8), PtrSize
    If pMethodNamesArr = 0 Then GoTo ErrorOccurred
    
    
    'Get pointer to our method name
    Dim pMethodName As LongPtr
    CopyMemory pMethodName, (pMethodNamesArr + PtrSize * methodIndex), PtrSize
    If pMethodName = 0 Then GoTo ErrorOccurred
    
    
    'Read the method name string
    Dim procName As String
    Dim readByteProcName As Byte
    Do
        CopyMemory readByteProcName, pMethodName, 1
        pMethodName = pMethodName + 1
        If readByteProcName = 0 Then Exit Do 'Read null char - end loop
        procName = procName & Chr(readByteProcName)
    Loop
    retVal.ProcedureName = procName
    
    
    'Get ObjectTable
    Dim pObjectTable As LongPtr
    CopyMemory pObjectTable, (pObjectInfo + PtrSize * 1), PtrSize
    If pObjectTable = 0 Then GoTo ErrorOccurred
    
    
    'Get project name from ObjectTable
    Dim pProjName As LongPtr
    #If Win64 Then
        CopyMemory pProjName, (pObjectTable + &H68), PtrSize
    #Else
        CopyMemory pProjName, (pObjectTable + &H40), PtrSize
    #End If
    If pProjName = 0 Then GoTo ErrorOccurred
    
    
    'Read the project name string
    Dim projName As String
    Dim readByteProjName As Byte
    Do
        CopyMemory readByteProjName, pProjName, 1
        pProjName = pProjName + 1
        If readByteProjName = 0 Then Exit Do 'Read null char - end loop
        projName = projName & Chr(readByteProjName)
    Loop
    
    
    retVal.ProjectName = projName
    GetStackFrame = retVal
    
Exit Function

ErrorOccurred:
    retVal.Errored = True
    GetStackFrame = retVal
End Function