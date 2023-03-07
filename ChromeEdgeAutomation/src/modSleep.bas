Attribute VB_Name = "modSleep"
Option Explicit

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)



'Custom sleep function
'change sleep period if processing is not robust
Public Const cnlngSleepPeriod As Long = 1000

Public Sub Sleep(Optional dblFrac As Double = 1)
    DoEvents
    Call sleep2(cnlngSleepPeriod * dblFrac)
    DoEvents
End Sub
    
