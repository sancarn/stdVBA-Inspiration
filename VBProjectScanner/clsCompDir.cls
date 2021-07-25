VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Purpose: Perform tests on compilation directives to determine
' if contained blocks of code are to be parsed/ignored

' Reset method must be called before scanning each source file

Const BlkEvalAsTrue = 1
Const BlkDecidedNoProcess = 2
Const BlkDecidedProcess = 3

Dim m_Eval() As Byte
    ' bit#1     1 if current statement is True, 0 = False
    ' bit#2     1 if block is decided, 0 = undecided
Dim m_Levels&               ' size of m_Eval() array
Dim m_LocalConst$           ' module-level #Constants
Dim m_GlobalConst$          ' global #Constants
Dim m_Tester As ScriptControl ' script control for testing IF statements
Attribute m_Tester.VB_VarHelpID = -1

Public Function IsBlockValid(s As String, cType As StatementParseFlags) As StatementParseFlags

    ' Processes compiler directives #If, #Else, #ElseIf, #End If

    ' Return value is either 0 (valid/true) or fCompIgnore (invalid/false)
    ' Passing s parameter required if cType is fCompIf,fCompElseIf
    
    ' to determine whether or not code following compiler directive can be processed,
    ' two bits make the decision...
    ' - once bit 2 is set, it never changes.
    '   any #Else,#ElseIf that follow have its bit 1 set to 0 automatically.
    ' - bit 1 changes depending on the evaluation of statement's #If or #ElseIf
    '   statement's current bit 2 value and/or parent-level bit 1 value.
    
    ' Condition                 Bit 1   Bit 2   Result
    ' ----------------------------------------------------------------------
    ' #If False = True Then       0       0     following code not processed
    '   ....
    ' #Else                       1       1     following code is processed
    '   ....
    ' #End If                                   parent-level resumes
    ' ----------------------------------------------------------------------
    ' #If True = True Then        1       1     following code is processed
    '   ....
    ' #ElseIf False = True Then   0       1     not processed
    '   ....
    ' #Else                       0       1     not processed
    '   ....
    ' #End If                                   parent-level resumes
    
    Dim lResult As StatementParseFlags: lResult = fCompIgnore
    
    cType = cType And fCompMask
    If cType = fCompIf Then             ' #If
        m_Levels = m_Levels + 1         ' increment block count & resize array
        If m_Levels > UBound(m_Eval) Then ReDim Preserve m_Eval(0 To m_Levels + 3)
        
        ' when #If is encountered then test it only when parent level is True
        If (m_Eval(m_Levels - 1) And BlkEvalAsTrue) = BlkEvalAsTrue Then
            m_Eval(m_Levels) = pvTestStatement(s, False)
            If m_Eval(m_Levels) = BlkDecidedProcess Then lResult = 0
        Else
            m_Eval(m_Levels) = BlkDecidedNoProcess ' decided, don't process
        End If
        
    ElseIf cType = fCompEnd Then        ' #End If
        m_Levels = m_Levels - 1         ' parent-level decides evaluation
        If (m_Eval(m_Levels) And BlkEvalAsTrue) = BlkEvalAsTrue Then lResult = 0
    
    Else                                ' #Else,#ElseIf
        ' when #ElseIf is encountered then test only when current level is undecided
        ' when #Else is encountered, statement is automatically True if level is undecided
        If m_Eval(m_Levels) = 0 Then    ' undecided
            If cType = fCompElseIf Then ' #ElseIf
                m_Eval(m_Levels) = pvTestStatement(s, True)
                If m_Eval(m_Levels) = BlkDecidedProcess Then lResult = 0
            Else                        ' #Else
                m_Eval(m_Levels) = BlkDecidedProcess: lResult = 0
            End If
        Else
            m_Eval(m_Levels) = BlkDecidedNoProcess ' decided, don't process
        End If
    End If
    
    IsBlockValid = lResult

End Function

Public Sub AddLocalConstant(s As String)

    ' cache compiler directives #Const parsed from non-vbp files

    m_LocalConst = vbCrLf & m_LocalConst & Mid$(s, 8)

End Sub

Public Property Let GlobalConstants(s As String)
    ' cache global #Const statements parsed from vbp files
    If LenB(s) <> 0 Then
        If AscW(s) = vbKeyQuote Then
            m_GlobalConst = Mid$(s, 2, Len(s) - 2)
        Else
            m_GlobalConst = s
        End If
    End If

End Property

Public Sub Reset()

    ' called in modMain.ScanSource for each project module,class

    ReDim m_Eval(0 To 5)    ' sized for five #If/#EndIf nested blocks
    m_Eval(0) = BlkDecidedProcess
    m_LocalConst = vbNullString
    m_Levels = 0
    
End Sub

Private Function pvTestStatement(s As String, isElseIf As Boolean) As Long

    Dim n&

    On Error Resume Next
    ' create the script control used to execute IF statement
    If m_Tester Is Nothing Then
        Set m_Tester = New ScriptControl
        m_Tester.Language = "VBScript"
    Else
        m_Tester.Reset ' remove previous function from control
    End If
    
    ' build our little function, looks like this:
        '   Function pvTest()
        '        ... list of compiler constants & their values, without #Const prefix
        '       If [passed #If/#ElseIf statement] = True Then
        '            pvTest = True      ' passed statement tests True
        '       Else
        '            pvTest = False     ' passed statement tests False
        '       End If
        '   End Function
    ' find "If" from #If or #ElseIf, then build function
    If isElseIf = False Then n = 2 Else n = 6
    m_Tester.AddCode _
        "Function pvTest()" & vbCrLf & _
            m_GlobalConst & m_LocalConst & vbCrLf & _
            Mid$(s, n) & " pvTest = True Else pvTest = False" & vbCrLf & _
        "End Function"
    
    ' execute the function & return result
    If m_Tester.Run("pvTest") = False Then
        If Err Then Err.Clear
    Else
        pvTestStatement = BlkDecidedProcess
    End If
    On Error GoTo 0

End Function

Private Sub Class_Terminate()
    Set m_Tester = Nothing
    Erase m_Eval()
    m_LocalConst = vbNullString
    m_GlobalConst = vbNullString
End Sub

