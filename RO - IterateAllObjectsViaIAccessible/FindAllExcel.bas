'Author: Jaafar Tribak
'Link:   https://www.mrexcel.com/board/threads/how-to-target-instances-of-excel.1118789/page-2#post-5395037

Option Explicit

Private Type GUID
    lData1 As Long
    iData2 As Integer
    iData3 As Integer
    aBData4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Sub AccessibleObjectFromWindow Lib "OLEACC.DLL" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As GUID, ppvObject As Any)
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
#Else
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Sub AccessibleObjectFromWindow Lib "OLEACC.DLL" (ByVal hwnd As Long, ByVal dwId As Long, riid As GUID, ppvObject As Any)
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
#End If

Private Const OBJID_NATIVEOM = &HFFFFFFF0

Sub Test()
    Call ListAllBooks
End Sub

Public Sub ListAllBooks()

    #If VBA7 Then
        Dim lXLhwnd As LongPtr, lWBhwnd As LongPtr
    #Else
        Dim lXLhwnd As Long, lWBhwnd As Long
   #End If
  
    Dim IDispatch As GUID
    Dim i As Long, lPID As Long, lPrevPID As Long
    Dim owb As Object

    With IDispatch
        .lData1 = &H20400
        .iData2 = &H0
        .iData3 = &H0
        .aBData4(0) = &HC0
        .aBData4(1) = &H0
        .aBData4(2) = &H0
        .aBData4(3) = &H0
        .aBData4(4) = &H0
        .aBData4(5) = &H0
        .aBData4(6) = &H0
        .aBData4(7) = &H46
    End With

    Do
        lXLhwnd = FindWindowEx(0, lXLhwnd, "XLMAIN", vbNullString)
        If lXLhwnd = 0 Then
            Exit Do
        Else
            lWBhwnd = FindWindowEx(FindWindowEx(lXLhwnd, 0&, "XLDESK", vbNullString), 0&, "EXCEL7", vbNullString)
            If lWBhwnd Then
                Call AccessibleObjectFromWindow(lWBhwnd, OBJID_NATIVEOM, IDispatch, owb)
                    GetWindowThreadProcessId owb.Application.hwnd, lPID
                    If lPID <> lPrevPID Then
                        Debug.Print IIf(lXLhwnd <> Application.hwnd, "Remote Application PID", "Current Application PID") & " : "; lPID
                        Debug.Print "************************************"
                        Debug.Print Space(4) & "Workbooks count : "; owb.Application.Workbooks.Count
                        For i = 1 To owb.Application.Workbooks.Count
                            Debug.Print Space(4) & "Workbook(" & i & "): "; owb.Application.Workbooks(i).FullName
                        Next i
                        Debug.Print
                    End If
                    lPrevPID = lPID
            End If
        End If
    Loop
    Set owb = Nothing

End Sub
