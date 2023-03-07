VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hListView1 As LongPtr
Private Const WC_LISTVIEW As String = "SysListView32"
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal IAcessible As Object, ByRef hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal uCmd As Long) As LongPtr
Private Declare PtrSafe Function InitCommonControlsEx Lib "COMCTL32" (ByRef pInitCtrls As Any) As Long
Private Declare PtrSafe Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Const GW_CHILD                  As Long = &O5&
Private Const ICC_LISTVIEW_CLASSES      As Long = &H1&
Private Const WS_EX_CLIENTEDGE          As Long = &H200&
Private Const WS_TABSTOP                As Long = &H10000
Private Const WS_BORDER                 As Long = &H800000
Private Const WS_CLIPSIBLINGS           As Long = &H4000000
Private Const WS_CHILD                  As Long = &H40000000
Private Const WS_VISIBLE                As Long = &H10000000
Private Const LVS_REPORT                As Long = &H1&
Private Const LVS_SINGLESEL             As Long = &H4&
Private Const LVS_SHOWSELALWAYS         As Long = &H8&
Private Const LVCFMT_LEFT               As Long = &H0
Private Const LVCFMT_RIGHT              As Long = &H1
Private Const LVCFMT_CENTER             As Long = &H2
Private Const LVCFMT_JUSTIFYMASK        As Long = &H3
Private Const LVCFMT_IMAGE              As Long = &H800
Private Const LVCFMT_BITMAP_ON_RIGHT    As Long = &H1000
Private Const LVCFMT_COL_HAS_IMAGES     As Long = &H8000
Private Const LVCF_FMT                  As Long = &H1&
Private Const LVCF_WIDTH                As Long = &H2&
Private Const LVCF_TEXT                 As Long = &H4&
Private Const LVCF_SUBITEM              As Long = &H8&
Private Const LVCF_IMAGE                As Long = &H10&
Private Const LVCF_ORDER                As Long = &H20&
Private Const LVIF_TEXT                 As Long = &H1&
Private Const LVIF_IMAGE                As Long = &H2&
Private Const LVIF_PARAM                As Long = &H4&
Private Const LVIF_STATE                As Long = &H8&
Private Const LVIF_INDENT               As Long = &H10&
Private Const LVIF_NORECOMPUTE          As Long = &H800&
Private Const LVIF_GROUPID              As Long = &H100&
Private Const LVIF_COLUMNS              As Long = &H200&
Private Const LVIF_COLFMT               As Long = &H10000
Private Type LVCOLUMNA
    mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
    iCxMin As Long
    iCxDefault As Long
    iCxIdeal As Long
End Type

Private Type LVITEMA
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As LongPtr
    iIndent As Long
    iGroupId As Long
    cColumns As Long
    pUColumns As LongPtr
    piColFmt As LongPtr
    iGroup As Long
End Type

Private Property Get hwnd() As LongPtr
    WindowFromAccessibleObject Me, hwnd
End Property

Private Sub InitCommonControlsEx_VBA(ByVal dwICC As Long)
    Dim iccE As LongPtr 'LongLong
    iccE = dwICC * &H10000000 + CLngPtr(LenB(iccE))
    Call InitCommonControlsEx(iccE)
End Sub

Private Function ListView_InsertColumnA_VBA(ByVal hwnd As LongPtr, ByVal iCol As Long, ByVal fmt As Long, ByVal Width As Long, ByVal strText As String, ByVal SubItem As Long, Optional iImage As Variant, Optional iOrder As Variant) As LongPtr
    Dim pCol As LVCOLUMNA
    pCol.mask = LVCF_FMT Or LVCF_WIDTH Or LVCF_TEXT Or LVCF_SUBITEM
    pCol.fmt = fmt
    pCol.cx = Width
    pCol.pszText = strText & vbNullChar
    pCol.iSubItem = SubItem
    If Not IsMissing(iImage) Then
        pCol.iImage = CLng(iImage)
        pCol.mask = pCol.mask Or LVCF_IMAGE
    End If
    If Not IsMissing(iOrder) Then
        pCol.iOrder = CLng(iOrder)
        pCol.mask = pCol.mask Or LVCF_ORDER
    End If
    
    Const LVM_INSERTCOLUMN As Long = &H101B&
    ListView_InsertColumnA_VBA = SendMessageA(hwnd, LVM_INSERTCOLUMN, CLngPtr(iCol), pCol)
End Function

Private Function ListView_InsertItemA_VBA(ByVal hwnd As LongPtr, ByVal iItem As Long, ByVal strText As String, Optional iImage As Variant) As LongPtr
    Dim pItem As LVITEMA
    If IsMissing(iImage) Then
        pItem.mask = LVIF_TEXT
    Else
        pItem.mask = LVIF_TEXT Or LVIF_IMAGE
        pItem.iImage = iImage
    End If
    pItem.iItem = iItem
    pItem.iSubItem = 0&
    pItem.pszText = strText
    Const LVM_INSERTITEM As Long = &H1007&
    ListView_InsertItemA_VBA = SendMessageA(hwnd, LVM_INSERTITEM, 0, pItem)
End Function

Private Function ListView_SetItemA_VBA(ByVal hwnd As LongPtr, ByVal iItem As Long, ByVal iSubItem As String, ByVal strText As String) As LongPtr
    Dim pItem As LVITEMA
    pItem.mask = LVIF_TEXT
    pItem.iItem = iItem
    pItem.iSubItem = iSubItem
    pItem.pszText = strText
    Const LVM_SETITEM As Long = &H1006&
    ListView_SetItemA_VBA = SendMessageA(hwnd, LVM_SETITEM, 0, pItem)
End Function

Private Sub UserForm_Initialize()
    Me.Width = 420 * (72 / 96) + 2
    Me.Height = 300 * (72 / 96)
    
    InitCommonControlsEx_VBA ICC_LISTVIEW_CLASSES

    Dim hClient As LongPtr
    hClient = GetWindow(hwnd, GW_CHILD)
    hListView1 = CreateWindowExA(WS_EX_CLIENTEDGE, WC_LISTVIEW, "ListView1", WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS, 4&, 4&, 400&, 250&, hClient, 0, Application.HinstancePtr, ByVal 0)

    ListView_InsertColumnA_VBA hListView1, 0&, LVCFMT_LEFT, 140, "Name", 0&
    ListView_InsertColumnA_VBA hListView1, 1&, LVCFMT_LEFT, 130, "Type", 2&
    ListView_InsertColumnA_VBA hListView1, 2&, LVCFMT_RIGHT, 120, "Size", 1&

    ListView_InsertItemA_VBA hListView1, 0&, "Program Files"
    ListView_InsertItemA_VBA hListView1, 1&, "Program Files (x86)"
    ListView_InsertItemA_VBA hListView1, 2&, "Windows"
    ListView_InsertItemA_VBA hListView1, 3&, "File"

    ListView_SetItemA_VBA hListView1, 0&, 1&, "Folder"
    ListView_SetItemA_VBA hListView1, 1&, 1&, "Folder"
    ListView_SetItemA_VBA hListView1, 2&, 1&, "Folder"
    ListView_SetItemA_VBA hListView1, 3&, 1&, "System Folder"
    ListView_SetItemA_VBA hListView1, 3&, 2&, "6,194,720 KB"

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    DestroyWindow hListView1
    hListView1 = 0
End Sub

Private Sub UserForm_Terminate()
    DestroyWindow hListView1
    hListView1 = 0
End Sub
