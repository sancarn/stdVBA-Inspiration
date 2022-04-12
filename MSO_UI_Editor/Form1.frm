VERSION 5.00
Object = "{07C05129-C2E5-483C-8237-8636C3F11E4E}#1.0#0"; "VBCCR13.OCX"
Begin VB.Form Form1 
   Caption         =   "MSO Custom UI"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VBCCR13.Slider Slider1 
      Height          =   255
      Left            =   9600
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Min             =   10
      Max             =   500
      Value           =   100
      TickFrequency   =   100
      SmallChange     =   10
      LargeChange     =   20
      TickStyle       =   3
      ShowTip         =   0   'False
      SelStart        =   100
      Transparent     =   -1  'True
   End
   Begin VBCCR13.ImageList ImageList2 
      Left            =   3720
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      InitListImages  =   "Form1.frx":0000
   End
   Begin VBCCR13.ImageList ImageList1 
      Left            =   3000
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   32
      InitListImages  =   "Form1.frx":1EE0
   End
   Begin VBCCR13.ListBoxW List1 
      Height          =   3780
      Left            =   6120
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   6668
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      IntegralHeight  =   0   'False
      DrawMode        =   2
   End
   Begin VBCCR13.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   5640
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   476
      Step            =   10
   End
   Begin VBCCR13.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      Top             =   6075
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InitPanels      =   "Form1.frx":47C0
   End
   Begin VBCCR13.RichTextBox RTB1 
      Height          =   3735
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6588
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      HideSelection   =   0   'False
      MultiLine       =   -1  'True
      MaxLength       =   350000
      ScrollBars      =   3
      WantReturn      =   -1  'True
      AutoURLDetect   =   0   'False
      SelectionBar    =   -1  'True
      UndoLimit       =   0
      TextRTF         =   "Form1.frx":4B44
   End
   Begin VBCCR13.TreeView TreeView1 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7011
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageList       =   "ImageList1"
      LineStyle       =   1
      HideSelection   =   0   'False
   End
   Begin VBCCR13.ToolBar ToolBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageList       =   "ImageList1"
      Style           =   1
      ShowTips        =   -1  'True
      ButtonWidth     =   23
      InitButtons     =   "Form1.frx":4CB8
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   2520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5775
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   360
      Width           =   135
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu MnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "Close"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu MnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MnuInsert 
      Caption         =   "Insert"
      Begin VB.Menu MnuOffice2010 
         Caption         =   "Office 2010 Custom UI Part"
      End
      Begin VB.Menu MnuOffice2007 
         Caption         =   "Office 2007 Custom UI Part"
      End
      Begin VB.Menu MnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIcons 
         Caption         =   "Icons"
      End
      Begin VB.Menu MnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSamples 
         Caption         =   "Samples"
         Begin VB.Menu SubMnuSamples 
            Caption         =   ""
            Index           =   0
         End
      End
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "MnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu MnuChangeID 
         Caption         =   "Change ID"
      End
      Begin VB.Menu MnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAddImage 
         Caption         =   "Add Image"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetDCBrushColor Lib "gdi32" (ByVal hdc As Long, ByVal colorref As Long) As Long
Private Declare Function SetDCPenColor Lib "gdi32" (ByVal hdc As Long, ByVal colorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, lpStr As Any, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetCaretPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const ILD_TRANSPARENT = &H1
Private Const TRANSPARENT As Long = 1
Private Const COLOR_WINDOW As Long = 5
Private Const COLOR_WINDOWTEXT As Long = 8
Private Const COLOR_HIGHLIGHT As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const ODS_SELECTED As Long = &H1
Private Const DC_PEN As Long = 19
Private Const DC_BRUSH As Long = 18
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const LB_GETITEMHEIGHT      As Long = &H1A1
Private Const LB_SETITEMHEIGHT      As Long = &H1A0
Private Const WM_USER As Long = &H400
Private Const EM_GETSCROLLPOS As Long = (WM_USER + 221)
Private Const EM_SETSCROLLPOS As Long = (WM_USER + 222)
Private Const EM_EXSETSEL As Long = (WM_USER + 55)
Private Const EM_LINEINDEX As Long = &HBB
Private Const WM_SETICON As Long = &H80
Private Const ICON_BIG As Long = 1
Private Const ICON_SMALL As Long = 0

Private Type Callback
    Control As String
    CallBackName As String
    SignaturesSub As String
End Type

Private Enum HighlightColors
    HC_NODE = &H404080
    HC_STRING = vbBlue
    HC_ATTRIBUTE = vbRed
    HC_COMMENT = &H158719
    HC_INNERTEXT = vbBlack
End Enum

Public bOn As Boolean
Private bTextChange As Boolean
Private SourceFile As String
Private sTempDir As String
Private sZipPath  As String
Private sFullName As String
Private sWorkPath As String
Private sCustomUI14 As String
Private sCustomUI As String
Private PosStart As Long
Private ChrStart As Integer
Private sCurControl As String
Private tCB() As Callback
Private m_Tv_SelectedItemIndex As Long
Private hModShell32 As Long
Private sCallbacks As String
Private mLine As Long
Private cHistory As Collection
Private lHistoryPos As Long
Private LastSelStart As Long
Private cMenuImage As clsMenuImage
Private bDocumentChange As Boolean

Private Sub InizializeHistory()
    Set cHistory = New Collection
    Dim cRtbHistory As ClsRtbHistory
    Set cRtbHistory = New ClsRtbHistory
    cRtbHistory.SelStart = RTB1.SelStart
    cRtbHistory.Text = RTB1.Text
    cHistory.Add cRtbHistory
    lHistoryPos = cHistory.Count
End Sub

Private Sub AddHistory()
    Dim cRtbHistory As ClsRtbHistory
    Dim i As Long
    
    If lHistoryPos < cHistory.Count Then
        For i = cHistory.Count To lHistoryPos + 1 Step -1
            cHistory.Remove i
        Next
    End If
    
    Set cRtbHistory = New ClsRtbHistory
    cRtbHistory.SelStart = RTB1.SelStart
    cRtbHistory.Text = RTB1.Text
    cHistory.Add cRtbHistory
    lHistoryPos = cHistory.Count
End Sub

Private Sub UNDO()
    If bTextChange Then AddHistory
    
    If lHistoryPos > 1 Then
        bOn = True
        lHistoryPos = lHistoryPos - 1
        With cHistory(lHistoryPos)
           RTB1.Text = .Text
           RTB1.SelStart = .SelStart
        End With

        ColorearTexto
    Else
        Beep
    End If
    
    RTB1.SetFocus
End Sub

Private Sub REDO()
    If lHistoryPos < cHistory.Count Then
        bOn = True
        If lHistoryPos = 0 Then lHistoryPos = 1
        lHistoryPos = lHistoryPos + 1
        With cHistory(lHistoryPos)
           RTB1.Text = .Text
           RTB1.SelStart = .SelStart
        End With
        ColorearTexto
    Else
        Beep
    End If
    RTB1.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bDocumentChange Then
        Select Case MsgBox("Do you want to save changes?", vbYesNoCancel, Me.Caption)
            Case vbYes
                Call SaveChanges
            Case vbCancel
                Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHookForm RTB1.HWnd
End Sub

Private Sub List1_Click()
    RTB1.SetFocus
End Sub

Private Sub List1_ItemDraw(ByVal Item As Long, ByVal itemAction As Long, ByVal ItemState As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim obr As Long, opn As Long, L As Long, s As String, sBuff As String * 255
    Dim tRECT As RECT
    
    obr = SelectObject(hdc, GetStockObject(DC_BRUSH))
    opn = SelectObject(hdc, GetStockObject(DC_PEN))
    
    If (ItemState And ODS_SELECTED) Then
        SetDCBrushColor hdc, GetSysColor(COLOR_HIGHLIGHT)
        SetDCPenColor hdc, GetSysColor(COLOR_HIGHLIGHT)
        Rectangle hdc, Left, Top, Right, Bottom
        SetDCPenColor hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
        SetTextColor hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
    Else
        SetDCBrushColor hdc, GetSysColor(COLOR_WINDOW)
        SetDCPenColor hdc, GetSysColor(COLOR_WINDOW)
        Rectangle hdc, Left, Top, Right, Bottom
        SetDCPenColor hdc, GetSysColor(COLOR_WINDOWTEXT)
        SetTextColor hdc, GetSysColor(COLOR_WINDOWTEXT)
    End If
    SetBkMode hdc, TRANSPARENT
    ImageList2.ListImages(List1.ItemData(Item) + 1).Draw hdc, 2, Top + 1, ImlDrawTransparent
    
    With tRECT
        .Left = Left + 20
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    
    DrawText hdc, ByVal List1.List(Item), Len(List1.List(Item)), tRECT, DT_VCENTER Or DT_SINGLELINE
    SelectObject hdc, obr
    SelectObject hdc, opn
End Sub

Public Sub RtbSel(Min As Long, Max As Long)
    Dim RECR As POINTAPI
    RECR.x = Min
    RECR.y = Max
    SendMessage RTB1.HWnd, EM_EXSETSEL, 0, ByVal VarPtr(RECR)
End Sub

Public Sub SyntaxHighlightXML(RTB As RichTextBox)
    Dim k As Long
    Dim Str As String
    Dim st As Long, en As Long, st2 As Long
    Dim xStar As Long, xSelLength As Long
    Dim p As POINTAPI
    
    RTB.HideSelection = True
    RTB.Visible = False
    SendMessage RTB.HWnd, EM_GETSCROLLPOS, 0, ByVal VarPtr(p)
    
    'LockWindowUpdate RTB.HWnd
    xStar = RTB.SelStart
    xSelLength = RTB.SelLength
    Str = Replace(RTB.Text, Chr(13), vbNullString)

    k = 1

    RtbSel 0, Len(Str)
    RTB.SelColor = HighlightColors.HC_INNERTEXT
    RTB.SelBkColor = &H80000005
   
    Do While k > 0
        st = InStr(k, Str, "<")

        If st Then
            en = InStr(st, Str, ">")
            If en Then
                RtbSel st - 1, en
                RTB.SelColor = HighlightColors.HC_STRING
            Else
                Exit Do
            End If
            
            k = st + 1
            st = en
            st2 = InStr(k, Str, Space$(1))
            If st > st2 Then
                If st2 > 0 Then
                    en = st2
                Else
                    en = st
                End If
            Else
                If st > 0 Then
                    en = st
                Else
                    en = st2
                End If
            End If

            RtbSel k - 1, en - 1
            RTB.SelColor = HighlightColors.HC_NODE
            
            k = en
        Else
            Exit Do
        End If
    Loop
'GoTo caca
    k = 1
    Do
        en = InStr(k, Str, "=")

        If en Then
            st = InStrRev(Str, " ", en)
            RtbSel st, en - 1
            RTB.SelColor = HighlightColors.HC_ATTRIBUTE
            k = en + 1
        Else
            Exit Do
        End If
    Loop
    
    k = 1
    Do
        st = InStr(k, Str, "<!--")

        If st Then
            en = InStr(st, Str, "-->")
            If en Then
                RtbSel st - 1, en + 2
                RTB.SelColor = HighlightColors.HC_COMMENT
                k = en + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    k = 1
    
    Do While k > 0
        st = InStr(k, Str, ">")

        If st Then
            en = InStr(st, Str, "<")
            If en Then
                If en - st > 3 Then
                    RtbSel st + 1, en - 1
                    RTB.SelColor = HighlightColors.HC_INNERTEXT
                End If
            Else
                Exit Do
            End If
            k = en
        Else
            Exit Do
        End If
    Loop

    RTB.SelStart = xStar
    RTB.SelLength = xSelLength
    SendMessage RTB.HWnd, EM_SETSCROLLPOS, 0, ByVal VarPtr(p)
    'LockWindowUpdate 0&
    RTB.HideSelection = False
    RTB.Visible = True
    RTB.SetFocus
    'RTB.Refresh
    
End Sub

Private Function ValidateAsXmlFile() As Boolean
    Dim xmlDoc As Object
    Dim strResult As String
    Dim oSC As Object

    Set oSC = CreateObject("MSXML2.XMLSchemaCache.6.0")
    
    If RTB1.Tag = "customUI" Then
        oSC.Add "http://schemas.microsoft.com/office/2006/01/customui", App.Path & "\CustomUI.xsd"
    Else
        oSC.Add "http://schemas.microsoft.com/office/2009/07/customui", App.Path & "\CustomUI14.xsd"
    End If

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    With xmlDoc
        .SetProperty "ProhibitDTD", False
        .schemas = oSC
        .validateOnParse = True
        .async = False
        .loadXML RTB1.Text
    End With
    
    Select Case xmlDoc.parseError.errorCode
       Case 0
            ValidateAsXmlFile = True
       Case Else
            strResult = vbCrLf & "ERROR! Failed to validate " & _
                        vbCrLf & xmlDoc.parseError.reason & vbCr & _
                        "Error code: " & xmlDoc.parseError.errorCode & ", Line: " & _
                        xmlDoc.parseError.Line & ", Character: " & _
                        xmlDoc.parseError.linepos & ", Source: " & _
                        Chr(34) & xmlDoc.parseError.srcText & Chr(34)
                          
            Dim sLine As String
            Dim lPos As Long, lpos1 As Long, lPos2 As Long
            
            ColorearTexto
            bOn = True
            
            sLine = RTB1.GetLine(xmlDoc.parseError.Line)
            lPos = InStr(Replace(RTB1.Text, Chr(10), ""), sLine)
            
            RtbSel lPos - 1, lPos + Len(sLine)
            RTB1.SelBkColor = vbYellow
            RTB1.SelStart = lPos - 1
            
            bOn = False
            bTextChange = False
            
            MsgBox strResult, vbCritical
            
            RTB1.SetFocus
    End Select
    

    Set xmlDoc = Nothing
End Function

Private Sub Form_Initialize()
    hModShell32 = LoadLibrary("Shell32.dll")
    Call InitCommonControls
End Sub

Private Sub Form_Terminate()
    FreeLibrary hModShell32
End Sub

Private Sub Form_Load()
    Dim LargeIcon As Long
    Dim SmallIcon As Long
    Dim sDir As String
    Dim iPos As Integer
  
    ExtractIconEx App.Path & "\" & App.EXEName & ".exe", 0, LargeIcon, SmallIcon, 1
    If LargeIcon Then Call SendMessage(Me.HWnd, WM_SETICON, ICON_BIG, ByVal LargeIcon)
    If SmallIcon Then Call SendMessage(Me.HWnd, WM_SETICON, ICON_SMALL, ByVal SmallIcon)

    mLine = 1

    SendMessage List1.HWnd, LB_SETITEMHEIGHT, 0, ByVal CLng(18)

    Call ReadCalbackTable
    
    Call InizializeHistory
     
    If App.LogMode Then
       HookForm RTB1.HWnd
       AddIconsInMenu
    End If
    
    sDir = Dir$(App.Path & "\Samples\")
    
    Do While Len(sDir)
       If iPos = 0 Then
           If App.LogMode Then
            cMenuImage.AddIconFromHandle GetFileIcon(App.Path & "\Samples\" & sDir)
           End If
       Else
           Load SubMnuSamples(iPos)
       End If
       SubMnuSamples(iPos).Caption = GetFileName(sDir)
       SubMnuSamples(iPos).Tag = App.Path & "\Samples\" & sDir
       If App.LogMode Then cMenuImage.PutImageToVBMenu 12, iPos, 2, 5
       iPos = iPos + 1
       sDir = Dir$()
    Loop
    
    If App.LogMode Then
        If Not cMenuImage.IsWindowVistaOrLater Then
            cMenuImage.RemoveMenuCheckVB 2, 5
        End If
    End If
    
    If Len(Command) Then
        Me.Show
        OpenFile Replace(Command, Chr$(34), vbNullString)
    End If
End Sub

Private Sub ReadCalbackTable()
    Dim sDatos As String, slines() As String, sParse() As String
    Dim i As Long
    
    sDatos = ReadFileText(App.Path & "\Callback.txt")
    
    slines = Split(sDatos, vbCrLf)
    
    ReDim tCB(UBound(slines) - 1)
    
    For i = 0 To UBound(slines) - 1
        sParse = Split(slines(i), vbTab)
        tCB(i).Control = sParse(0)
        tCB(i).CallBackName = sParse(1)
        tCB(i).SignaturesSub = sParse(2)
    Next
End Sub

Private Sub AddIconsInMenu()
    Dim i As Long
    Set cMenuImage = New clsMenuImage

    With cMenuImage
        
        .Init Me.HWnd, 16, 16
    
        For i = 1 To 12
           .AddImageFromStream LoadResData("PNG_" & i, "PNG")
        Next
        
        '---------'
        '  FILE   '
        '---------'
        .PutImageToVBMenu 0, 0, 0       ' OPEN
        .PutImageToVBMenu 1, 2, 0       ' SAVE
        .PutImageToVBMenu 2, 3, 0       ' SAVE AS
        .PutImageToVBMenu 3, 5, 0       ' CLOSE

        '---------'
        '   EDIT  '
        '---------'
        .PutImageToVBMenu 4, 0, 1       ' UNDO
        .PutImageToVBMenu 5, 1, 1       ' REDO
        .PutImageToVBMenu 6, 3, 1       ' CUT
        .PutImageToVBMenu 7, 4, 1       ' COPY
        .PutImageToVBMenu 8, 5, 1       ' PASTE
        .PutImageToVBMenu 9, 7, 1       ' SLELECT ALL
                
        '---------'
        ' INSERT  '
        '---------'
        .PutImageToVBMenu 10, 0, 2      ' CUSTOM 2010
        .PutImageToVBMenu 10, 1, 2      ' CUSTOM 2007
        .PutImageToVBMenu 11, 3, 2      ' ICONS
        .PutImageToVBMenu 10, 5, 2      ' SAMPLES
        
        ' En Windows XP queda mejor si remobemos el style check, ya que éste agrega un margen adicional para las marcas de verificación.
        ' En Windows Vista o Windows 7 esto no es necesario, ya que lo remarca debajo de la imágen.
        If Not .IsWindowVistaOrLater Then
            .RemoveMenuCheckVB 0
            .RemoveMenuCheckVB 1
            .RemoveMenuCheckVB 2
        End If
        
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        RTB1.Move TreeView1.Width + 50, ToolBar1.Height, Me.ScaleWidth - TreeView1.Width - 50, Me.ScaleHeight - ToolBar1.Height - StatusBar1.Height
        TreeView1.Move 0, ToolBar1.Height, TreeView1.Width, Me.ScaleHeight - ToolBar1.Height - StatusBar1.Height
        Picture1.Left = TreeView1.Width
        ProgressBar1.Move StatusBar1.Panels(2).Left - ProgressBar1.Width - 50, Me.ScaleHeight - (StatusBar1.Height / 2) - (ProgressBar1.Height / 2)
        Slider1.Move Me.ScaleWidth - Slider1.Width - 800, Me.ScaleHeight - Slider1.Height - 20
    End If
End Sub

Private Sub List1_DblClick()
    EnterItemToText
    RTB1.SetFocus
End Sub

Private Sub List1_LostFocus()
    If ActiveControl.Name <> RTB1.Name Then
        List1.Visible = False
        PosStart = 0
        ChrStart = 0
    End If
End Sub

Private Sub MnuAddImage_Click()
    Call AddImage
End Sub

Private Sub MnuChangeID_Click()
    TreeView1.Nodes(m_Tv_SelectedItemIndex).Selected = True
    TreeView1.StartLabelEdit
End Sub

Private Sub MnuClose_Click()
    Dim FSO As Object
    Dim i As Long
    
    If bDocumentChange Then
        Select Case MsgBox("Do you want to save changes?", vbYesNoCancel, Me.Caption)
            Case vbYes
                Call SaveChanges
            Case vbCancel
                Exit Sub
            'Case vbNo
        End Select
    End If
    
    
    Set FSO = CreateObject("scripting.filesystemobject")
 
    If Len(sTempDir) Then
        If FSO.FolderExists(sTempDir) Then
            FSO.DeleteFolder sTempDir, True
        End If
        
        sCustomUI14 = vbNullString
        sCustomUI = vbNullString
        TreeView1.Nodes.Clear
        RTB1.Text = vbNullString
        RTB1.Tag = vbNullString
        sTempDir = vbNullString
        
        For i = ImageList1.ListImages.Count To 9 Step -1
            ImageList1.ListImages.Remove i
        Next
        ToolBar1.Buttons(2).Enabled = False
        EnabledControls False
        Me.Caption = App.Title
    End If
  
End Sub

Private Sub MnuCopy_Click()
    RTB1.Copy
End Sub

Private Sub MnuCut_Click()
    RTB1.Cut
End Sub

Private Sub MnuDelete_Click()
    Dim sPath As String
    Dim xLevel As Integer
    Dim i As Integer
    Dim sKey As String
    
    xLevel = TreeView1.Nodes(m_Tv_SelectedItemIndex).Level

    If xLevel = 1 Then
    
        For i = TreeView1.Nodes.Count To 2 Step -1
            If TreeView1.Nodes(i).Parent.Text = TreeView1.Nodes(m_Tv_SelectedItemIndex).Text Then
                sPath = sWorkPath & "customUI\" & TreeView1.Nodes(i).Tag
                If FileExists(sPath) Then
                    Kill sPath
                End If
                TreeView1.Nodes.Remove i
            End If
        Next
        
        sKey = TreeView1.Nodes(m_Tv_SelectedItemIndex).Key
        If sKey = "customUI" Then sCustomUI = vbNullString
        If sKey = "customUI14" Then sCustomUI14 = vbNullString
        If sKey = RTB1.Tag Then
            RTB1.Text = vbNullString
            RTB1.Tag = vbNullString
        End If
    
        sPath = sWorkPath & "customUI\" & TreeView1.Nodes(m_Tv_SelectedItemIndex).Text
    Else
        sPath = sWorkPath & "customUI\" & TreeView1.Nodes(m_Tv_SelectedItemIndex).Tag
    End If
    
    If FileExists(sPath) Then
        Kill sPath
    End If

    TreeView1.Nodes.Remove m_Tv_SelectedItemIndex
    
    If TreeView1.Nodes.Count < 2 Then
        EnabledControls False
    End If
End Sub

Private Sub MnuExport_Click()
    Dim SourceFile As String, sDestPath  As String
    Dim xLevel As Integer
    Dim cCDialog As CommonDialog
    Dim sExtention As String

    xLevel = TreeView1.Nodes(m_Tv_SelectedItemIndex).Level

    If xLevel = 1 Then
        SourceFile = sWorkPath & "customUI\" & TreeView1.Nodes(m_Tv_SelectedItemIndex).Text
        Dim sData As String, sKey As String
        Dim FF As Integer
        
        Set cCDialog = New CommonDialog
        
        With cCDialog
        
            .Filter = "Documento XML (.xml)"
            .FileName = TreeView1.Nodes(m_Tv_SelectedItemIndex).Text
            .DefaultExt = "XML"
            .Flags = CdlOFNOverwritePrompt
            If .ShowSave Then
                sKey = TreeView1.Nodes(m_Tv_SelectedItemIndex).Key
                If sKey = "customUI" Then sData = sCustomUI
                If sKey = "customUI14" Then sData = sCustomUI14
                If sKey = RTB1.Tag Then
                    sData = RTB1.Text
                End If
                
                If FileExists(.FileName) Then Kill .FileName

                WriteFile sData, .FileName
            End If
            
        End With
        
    Else
        SourceFile = sWorkPath & "customUI\" & TreeView1.Nodes(m_Tv_SelectedItemIndex).Tag

        If FileExists(SourceFile) Then
            Set cCDialog = New CommonDialog
            
            With cCDialog
                sExtention = GetFileExtention(SourceFile)
                .Filter = GetFileDescription(SourceFile)
                .FileName = SourceFile
                .DefaultExt = sExtention
                .Flags = CdlOFNOverwritePrompt
                
                If .ShowSave Then
                    If GetFileExtention(.FileName) <> sExtention Then
                        sDestPath = GetFileFolder(.FileName) & GetFileName(.FileName) & "." & sExtention
                    Else
                        sDestPath = .FileName
                    End If
                    
                    If FileExists(sDestPath) Then Kill sDestPath
                    FileCopy SourceFile, sDestPath
                End If
            End With
        End If
    End If

End Sub

Private Sub MnuIcons_Click()
    Call AddImage
End Sub

Private Sub MnuInsert_Click()
    Dim i As Long
    MnuOffice2010.Enabled = Len(sTempDir)
    MnuOffice2007.Enabled = Len(sTempDir)
    MnuIcons.Enabled = Len(sTempDir)
    MnuSamples.Enabled = Len(sTempDir)
    
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Text = "customUI14.xml" Then
            MnuOffice2010.Enabled = False
        End If
        
        If TreeView1.Nodes(i).Text = "customUI.xml" Then
            MnuOffice2007.Enabled = False
        End If
    Next
End Sub

Private Sub MnuOffice2007_Click()
    TreeView1.Nodes.Add(1, TvwNodeRelationshipChild, "customUI", "customUI.xml", 2).Selected = True
    RTB1.Tag = "customUI"
    RTB1.Text = ""
    EnabledControls True
    bDocumentChange = True
End Sub

Private Sub MnuOffice2010_Click()
    TreeView1.Nodes.Add(1, TvwNodeRelationshipChild, "customUI14", "customUI14.xml", 2).Selected = True
    RTB1.Tag = "customUI14"
    RTB1.Text = ""
    EnabledControls True
    bDocumentChange = True
End Sub

Private Sub EnabledControls(bEnabled As Boolean)
    RTB1.Locked = Not bEnabled
    With ToolBar1
        .Buttons(4).Enabled = bEnabled
        .Buttons(5).Enabled = bEnabled
        .Buttons(6).Enabled = bEnabled
        .Buttons(8).Enabled = bEnabled
    End With
End Sub

Private Sub MnuFile_Click()
    If Len(sTempDir) Then
        MnuSave.Enabled = True
        MnuSaveAs.Enabled = True
        MnuClose.Enabled = True
    Else
        MnuSave.Enabled = False
        MnuSaveAs.Enabled = False
        MnuClose.Enabled = False
    End If
    
End Sub

Private Sub MnuOpen_Click()
    Call OpenFile
End Sub

Private Sub MnuPaste_Click()
    RTB1.Paste
End Sub

Private Sub MnuRedo_Click()
    REDO
End Sub

Private Sub MnuSave_Click()
    Call SaveChanges
End Sub

Private Sub MnuSaveAs_Click()
    Dim cCDialog As CommonDialog
    Dim sExtention As String
    If Len(SourceFile) = 0 Then Exit Sub
    
    Set cCDialog = New CommonDialog
    
    With cCDialog
        sExtention = GetFileExtention(SourceFile)
        .Filter = GetFileDescription(SourceFile)
        .FileName = SourceFile
        .DefaultExt = sExtention
        
        If .ShowSave Then
            If GetFileExtention(.FileName) <> sExtention Then
                SourceFile = GetFileFolder(.FileName) & GetFileName(.FileName) & "." & sExtention
            Else
                SourceFile = .FileName
            End If
            Call SaveChanges
        End If
    End With

End Sub

Private Sub MnuSelectAll_Click()
    RTB1.SelStart = 0
    RTB1.SelLength = Len(RTB1.Text) '- 1
    RTB1.SetFocus
End Sub

Private Sub MnuEdit_Click()
    With RTB1
        MnuUndo.Enabled = lHistoryPos > 1 Or bTextChange = True  '.CanUndo
        MnuRedo.Enabled = lHistoryPos < cHistory.Count '.CanRedo
        MnuPaste.Enabled = .CanPaste
        MnuCopy.Enabled = .SelLength > 0
        MnuCut.Enabled = .SelLength > 0
    End With
End Sub

Private Sub MnuUndo_Click()
    UNDO
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If x > -TreeView1.Width And x < Me.ScaleWidth - 5000 Then
            TreeView1.Width = Picture1.Left + x
            RTB1.Move TreeView1.Width + 50, ToolBar1.Height, Me.ScaleWidth - TreeView1.Width - 50, Me.ScaleHeight - ToolBar1.Height
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_Resize
End Sub

Private Sub RTB1_Change()
    Dim sWord As String
    Dim i As Long
    
    bTextChange = True
    bDocumentChange = True
    
    If List1.Visible Then
        Dim sText As String
        Dim sItem As String
        If (RTB1.SelStart) - PosStart > 0 Then
            sText = Trim(RTB1.Text)
            sText = Replace(sText, Chr(13), vbNullString)

            sWord = UCase(Mid$(sText, PosStart + 1, RTB1.SelStart - PosStart))
            sWord = Replace(sWord, Chr$(34), vbNullString)
            
            List1.ListIndex = -1
            For i = 0 To List1.ListCount - 1
                
                sItem = UCase(Left$(List1.List(i), Len(sWord)))
                
                If (sItem = sWord) Then
                    List1.ListIndex = i
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub LoadImgMsoTable()
    Dim sDatos As String
    Dim sElments() As String
    Dim i As Long
    
    sDatos = ReadFileText(App.Path & "\ImgMsoTable.txt")
    sElments = Split(sDatos, vbCrLf)
    List1.Clear
    For i = 0 To UBound(sElments) - 1
        List1.AddItem Trim(Split(sElments(i), vbTab)(0))
        List1.ItemData(List1.NewIndex) = 3
    Next
End Sub

Private Sub LoadImgageDocument()
    Dim i As Long
    List1.Clear
    For i = 2 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Parent.Key = RTB1.Tag Then
            List1.AddItem TreeView1.Nodes(i).Text
            List1.ItemData(List1.NewIndex) = 3
        End If
    Next
End Sub

Private Sub LoadElments()
    Dim sDatos As String
    Dim sElments() As String
    Dim i As Long

    sDatos = ReadFileText(App.Path & "\UI.txt")
    sElments = Split(sDatos, vbCrLf)
    List1.Clear
    For i = 0 To UBound(sElments) - 1
        List1.AddItem Trim(Split(sElments(i), vbTab)(0))
    Next

End Sub

Private Sub RTB1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    
        Case vbKeyBack
            AddHistory
            
        Case vbKeyDown
            If List1.Visible Then
                If Not List1.ListIndex >= List1.ListCount - 1 Then
                    List1.ListIndex = List1.ListIndex + 1
                End If
                KeyCode = 0
            End If
        
        Case vbKeyUp
            If List1.Visible Then
                If Not List1.ListIndex <= 0 Then
                    List1.ListIndex = List1.ListIndex - 1
                End If
                KeyCode = 0
            End If

        Case vbKeyEscape
             If List1.Visible Then
                List1.Visible = False
                PosStart = 0
             End If
        
        Case vbKeySpace, vbKeyReturn
            If List1.Visible Then
                If EnterItemToText Then KeyCode = 0
            End If

        Case vbKeyControl
            Slider1.Value = RTB1.ZoomFactor * 100
            StatusBar1.Panels(3).Text = Slider1.Value & "% "
    End Select
 
End Sub

Private Function EnterItemToText() As Boolean
    If RTB1.SelStart - PosStart + 1 > 0 Then
        If List1.ListIndex > -1 Then
            Dim lEnd As Long
            lEnd = RTB1.SelStart
            AddHistory
            RtbSel PosStart, lEnd
            If ChrStart = 61 Then
                RTB1.SelText = Chr$(34) & List1.Text & Chr$(34)
            Else
                RTB1.SelText = List1.Text
            End If
            
            List1.Visible = False
            PosStart = 0
            
            EnterItemToText = True
        Else
            List1.Visible = False
            PosStart = 0
        End If
    End If
End Function

Private Sub ListBoxShowItems(NroItems As Integer)
    Dim itemHeight As Long, MaxWidth As Long, i As Long
    If NroItems > 10 Then NroItems = 10
    If NroItems = 0 Then NroItems = 1
    itemHeight = SendMessage(List1.HWnd, LB_GETITEMHEIGHT, 0, ByVal 0)
    List1.Height = ((itemHeight * NroItems + 6) * Screen.TwipsPerPixelY)
End Sub

Private Sub RTB1_KeyPress(KeyChar As Integer)
    Dim PT As POINTAPI
    Dim x As Long, y As Long
    Dim lPosIni As Long, lPosEnd As Long
    Dim sText As String
    Dim sControl As String

    Select Case KeyChar
        Case vbKeyBack
            If RTB1.SelStart < PosStart Then
                List1.Visible = False
                ChrStart = 0
                PosStart = 0
            End If
        Case 47, 60 '/,<
                                                                                   
            Call LoadElments
            Call ListBoxShowItems(List1.ListCount)
            ChrStart = KeyChar
            PosStart = RTB1.SelStart + 1
            If (RTB1.SelStart > 0) Then
                If (KeyChar = 47) And Mid(Replace(RTB1.Text, Chr(13), ""), RTB1.SelStart, 1) <> "<" Then
                    ChrStart = 0
                    PosStart = 0
                End If
            End If
        Case 61 '=
            ChrStart = KeyChar
            PosStart = RTB1.SelStart + 1
            sText = Left$(Replace(RTB1.Text, Chr(13), vbNullString), PosStart - 1) ' & Space(1)
            lPosIni = InStrRev(sText, " ")
            sControl = UCase(Trim(Mid(sText, lPosIni + 1, PosStart - lPosIni - 1)))
            
            Select Case sControl
                Case "IMAGEMSO"
                    Call LoadImgMsoTable
                    Call ListBoxShowItems(List1.ListCount)
                Case "IMAGE"
                    Call LoadImgageDocument
                    Call ListBoxShowItems(List1.ListCount)
                    If List1.ListCount = 0 Then
                        ChrStart = 0
                        PosStart = 0
                    End If
                Case "SIZE", "ITEMSIZE"
                    With List1
                        .Clear
                        .AddItem "large"
                        .ItemData(.NewIndex) = 6
                        .AddItem "normal"
                        .ItemData(.NewIndex) = 7
                    End With
                    Call ListBoxShowItems(2)
                Case "VISIBLE", "SHOWIMAGE", "SHOWLABEL", "ENABLED"
                    With List1
                        .Clear
                        .AddItem "false"
                        .ItemData(.NewIndex) = 4
                        .AddItem "true"
                        .ItemData(.NewIndex) = 5
                    End With
                    Call ListBoxShowItems(2)
                Case "BOXSTYLE"
                    With List1
                        .Clear
                        .AddItem "horizontal"
                        .ItemData(.NewIndex) = 6
                        .AddItem "vertical"
                        .ItemData(.NewIndex) = 7
                    End With
                    Call ListBoxShowItems(2)
                Case "XMLNS"
                    With List1
                        .Clear
                        If RTB1.Tag = "customUI" Then
                            .AddItem "http://schemas.microsoft.com/office/2006/01/customui"
                        Else
                            .AddItem "http://schemas.microsoft.com/office/2009/07/customui"
                        End If
                        .ItemData(.NewIndex) = 1
                    End With
                    Call ListBoxShowItems(1)
                Case Else
                    PosStart = 0
                    List1.Visible = False
            End Select
            
            
        Case 62 '>
            List1.Visible = False
            PosStart = 0
        Case 32 'Space

            ChrStart = KeyChar
            PosStart = RTB1.SelStart + 1

            sText = Left$(Replace(RTB1.Text, Chr(13), vbNullString), PosStart - 1) & Space(1)
            
            lPosIni = InStrRev(sText, "<")
            If lPosIni > InStrRev(sText, ">") Then
                lPosEnd = InStr(lPosIni, sText, " ")
                If lPosEnd > lPosIni Then
                    sControl = UCase(Trim(Mid(sText, lPosIni + 1, lPosEnd - lPosIni - 1)))
                    If Len(sControl) Then sCurControl = sControl
                    If sControl = "!--" Then
                        PosStart = 0
                    Else
                        Dim sDatos As String
                        Dim sParts() As String
                        Dim sProperty() As String
                        Dim i As Long, j As Long, n As Long
                        
                        sDatos = ReadFileText(App.Path & "\UI.txt")
                        sParts = Split(sDatos, vbCrLf)
                        For i = 0 To UBound(sParts) - 1
                            If UCase(Trim(Split(sParts(i), vbTab)(0))) = sControl Then
                                List1.Clear
                                sProperty = Split(sParts(i), vbTab)
                                For j = 1 To UBound(sProperty) - 1
                                    If Len(Trim(sProperty(j))) Then
                                        List1.AddItem Trim(sProperty(j))
                                        List1.ItemData(List1.NewIndex) = 1
                                        For n = 0 To UBound(tCB)
                                            If tCB(n).CallBackName = sProperty(j) Then
                                                List1.ItemData(List1.NewIndex) = 2
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Next
                                Call ListBoxShowItems(List1.ListCount)
                                Exit For
                            End If
                        Next
                    End If
                    
                    If List1.ListCount = 0 Then
                        ChrStart = 0
                        PosStart = 0
                    End If
                    
                End If
                
            Else
               PosStart = 0
            End If

            List1.Visible = False
    End Select

    If PosStart > 0 And List1.Visible = False Then
        GetCaretPos PT
        x = ((PT.x + RTB1.LeftMargin) * Screen.TwipsPerPixelX) + RTB1.Left
        y = ((PT.y + (RTB1.SelFontSize * 2)) * Screen.TwipsPerPixelY) + RTB1.Top
        If y + List1.Height > Me.ScaleHeight Then
            y = (PT.y * Screen.TwipsPerPixelY) + RTB1.Top - List1.Height
        End If
        If x + List1.Width > Me.ScaleWidth Then
            x = Me.ScaleWidth - List1.Width
        End If
        List1.Move x, y
        List1.Visible = True
    End If

End Sub

Private Sub RTB1_LostFocus()
    If ActiveControl.Name <> List1.Name Then
        List1.Visible = False
        PosStart = 0
        ChrStart = 0
    End If
End Sub

Private Sub RTB1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu MnuEdit
End Sub

Private Sub RTB1_PreviewKeyDown(ByVal KeyCode As Integer, IsInputKey As Boolean)
    On Error Resume Next
    Dim lPosStart As Long

    If (KeyCode = vbKeyTab) Then
    
        If RTB1.Locked Then Exit Sub

        If InStr(RTB1.SelText, Chr(13)) Then
            TabuleSelection
        Else
            If GetKeyState(vbKeyShift) < 0 Then
            
                lPosStart = RTB1.SelStart
                RTB1.SelStart = lPosStart - 1
                RTB1.SelLength = 1
                If RTB1.SelText = vbTab Then
                    RTB1.SelText = vbNullString
                    RTB1.SelLength = 0
                Else
                    
                    RTB1.SelStart = lPosStart
                    RTB1.SelLength = 1
                    If RTB1.SelText = vbTab Then
                        RTB1.SelText = vbNullString
                        RTB1.SelLength = 0
                    Else
                        RTB1.SelStart = lPosStart
                    End If
                End If
                
            Else
                RTB1.SelText = vbTab
            End If
        End If
        KeyCode = 0
        IsInputKey = True
    End If
End Sub

Private Sub TabuleSelection()
    On Error Resume Next
    Dim sLine As String
    Dim lPos As Long, lPosEnd As Long
    Dim sText As String
    Dim slines() As String
    Dim i As Long

    sLine = RTB1.GetLine(RTB1.GetLineFromChar(RTB1.SelStart))
    lPos = InStr(Replace(RTB1.Text, Chr(10), ""), sLine)
    
    sLine = RTB1.GetLine(RTB1.GetLineFromChar(RTB1.SelStart + RTB1.SelLength - 1))
    lPosEnd = InStr(Replace(RTB1.Text, Chr(10), ""), sLine) + Len(sLine) - 1

    
    RtbSel lPos - 1, lPosEnd

    sText = RTB1.SelText
    slines = Split(sText, Chr(13))
    
    
    sText = vbNullString
    If GetKeyState(vbKeyShift) < 0 Then
        For i = 0 To UBound(slines)
            If Left(slines(i), 1) = vbTab Then
                sText = sText & Mid$(slines(i), 2) & Chr(13)
            Else
                sText = sText & slines(i) & Chr(13)
            End If
            
        Next
    Else
        For i = 0 To UBound(slines)
            sText = sText & vbTab & slines(i) & Chr(13)
        Next
    End If
    lPosEnd = lPos - 1 + Len(sText) - 1
    RTB1.SelText = Left(sText, Len(sText) - 1)
    RtbSel lPos - 1, lPosEnd
End Sub


Private Sub RTB1_SelChange(ByVal SelType As Integer, ByVal SelStart As Long, ByVal SelEnd As Long)

    Dim nLine As Long

    If bOn Then Exit Sub

    nLine = RTB1.GetLineFromChar(SelStart)
  
    StatusBar1.Panels(2).Text = "Ln: " & nLine & Space(5) & "Col: " & GetColPos(RTB1)
  
    If bTextChange = False Then
        mLine = RTB1.GetLineFromChar(SelStart)
        Exit Sub
    End If

    If RTB1.Tag = vbNullString Then Exit Sub

    If mLine <> nLine Then
        mLine = nLine
        AddHistory
        bOn = True
        List1.Visible = False
        PosStart = 0
        ChrStart = 0
        RTB1.Refresh
        SyntaxHighlightXML RTB1
        bOn = False
        bTextChange = False
    End If

End Sub

Private Sub Slider1_Scroll()
   RTB1.ZoomFactor = (Slider1.Value / 100)
   StatusBar1.Panels(3).Text = Slider1.Value & "% "
End Sub

Private Sub SubMnuSamples_Click(Index As Integer)
    RTB1.Text = ReadFileText(SubMnuSamples(Index).Tag)
    ColorearTexto
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As TbrButton)
    Select Case Button.Index
        Case 1
            Call OpenFile
        Case 2
            Call SaveChanges
        Case 4
            Call AddImage
        Case 5
            If ValidateAsXmlFile Then MsgBox RTB1.Tag & " XML is well formed!", vbInformation
        Case 6
            GenerateCallbacks
        Case 8
            FrmSearch.Show , Me
    End Select
End Sub

Private Sub GenerateCallbacks()
    Dim XDoc As Object 'MSXML2.DOMDocument
    Dim listNode As Object 'MSXML2.IXMLDOMNode
    Dim Elemnt As Object 'MSXML2.IXMLDOMElement
    Dim i As Long
    
    If RTB1.Tag = vbNullString Then Exit Sub
    If ValidateAsXmlFile = False Then Exit Sub

    For i = 1 To TreeView1.Nodes.Count
        With TreeView1.Nodes(i)
        If .ForeColor = vbBlue Then .ForeColor = TreeView1.ForeColor
        End With
    Next
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False

    If RTB1.Tag = "customUI" Then
        sCustomUI = RTB1.Text
        XDoc.loadXML sCustomUI
    End If
    
    MnuInsert.Enabled = False
    TreeView1.Nodes(1).Selected = True
    ToolBar1.Buttons(4).Enabled = False
    ToolBar1.Buttons(5).Enabled = False
    
    If RTB1.Tag = "customUI14" Then
        sCustomUI14 = RTB1.Text
        XDoc.loadXML sCustomUI14
    End If
    
    RTB1.Tag = vbNullString
   
    If Len(XDoc.xml) = 0 Then Exit Sub
    
    sCallbacks = vbNullString
    
    RecursiveXLS XDoc.childNodes

    RTB1.Text = sCallbacks
    
    bOn = True
    SyntaxHighlightVB RTB1
    
    RTB1.Locked = True
    InizializeHistory
End Sub

Private Sub RecursiveXLS(ParentNode As Object)
    Dim listNode As Object  ' As MSXML2.IXMLDOMNode
    Dim objAttribute ' As MSXML2.IXMLDOMAttribute
    Dim i As Long
    
    For Each listNode In ParentNode
        If listNode.nodeName <> "#comment" Then
            For Each objAttribute In listNode.Attributes
                For i = 0 To UBound(tCB)
                    If listNode.nodeName = tCB(i).Control And objAttribute.Name = tCB(i).CallBackName Then
                        sCallbacks = sCallbacks & "'Callback for " & listNode.nodeName & " " & objAttribute.Name & vbCrLf
                        sCallbacks = sCallbacks & "Sub " & objAttribute.Text & tCB(i).SignaturesSub & vbCrLf
                        sCallbacks = sCallbacks & "End Sub" & vbCrLf & vbCrLf
                    End If
                Next
            Next
        End If
        If listNode.childNodes.Length > 0 Then RecursiveXLS listNode.childNodes
    Next
    
    
End Sub

Private Sub SyntaxHighlightVB(RTB As RichTextBox)
    Dim lPosStart As Long, lPosEnd As Long, i As Long
    Dim Str As String, sWord() As String
    
    Str = Replace(RTB.Text, Chr(13), vbNullString)

    RtbSel 0, Len(Str)
    RTB.SelColor = HighlightColors.HC_INNERTEXT
    RTB.SelBkColor = &H80000005
   
    Do
        lPosStart = InStr(lPosStart + 1, Str, "'")
        If lPosStart Then
            lPosEnd = InStr(lPosStart, Str, Chr(10))
            RtbSel lPosStart - 1, lPosEnd
            RTB.SelColor = HighlightColors.HC_COMMENT
        Else
            Exit Do
        End If
    Loop
    
    Const sWords = "Sub ,End Sub, As , ByRef , String, Integer"
    sWord = Split(sWords, ",")
    
    For i = 0 To UBound(sWord)
        lPosStart = 0
        Do
            lPosStart = InStr(lPosStart + 1, Str, sWord(i))
            If lPosStart Then
                lPosEnd = lPosStart + Len(sWord(i)) - 1
                RtbSel lPosStart - 1, lPosEnd
                RTB.SelColor = HighlightColors.HC_STRING
            Else
                Exit Do
            End If
        Loop
    Next
    
    
    RtbSel 0, 0
End Sub

Private Sub SaveChanges()
    On Error GoTo ErrHandler

    Dim lhLib As Long
    Dim iFiles As ZIPnames
    Dim PrevPath As String
    
    If Len(SourceFile) = 0 Then Exit Sub
    ProgressBar1.Max = 7
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    Screen.MousePointer = vbHourglass
    StatusBar1.Panels(1).Text = "Saving file " & GetFileTitle(SourceFile)
    
    If RTB1.Tag = "customUI" Then sCustomUI = RTB1.Text
    If RTB1.Tag = "customUI14" Then sCustomUI14 = RTB1.Text
    If Len(sCustomUI) Then SaveXML sCustomUI, "customUI/customUI.xml", "2006"
    If Len(sCustomUI14) Then SaveXML sCustomUI14, "customUI/customUI14.xml", "2007"
    
    If FolderExists(sWorkPath & "customUI\images") Then
        WriteContentTypes
        ProgressBar1.Value = 1: DoEvents
        WriteRelationship "customUI"
        ProgressBar1.Value = 2: DoEvents
        WriteRelationship "customUI14"
        ProgressBar1.Value = 3: DoEvents
    End If
    
    If FileExists(sZipPath) Then Kill sZipPath
    ProgressBar1.Value = 4: DoEvents
    '--------------
    lhLib = LoadLibrary(App.Path & "\zip32.dll")
    PrevPath = CurDir
    ChDir sWorkPath
    iFiles.s(0) = "*"
    VBZip 1, sZipPath, iFiles, 0, 1, 0, 0, sWorkPath
    FreeLibrary lhLib
    ChDir PrevPath
    '--------------
    
    ProgressBar1.Value = 5: DoEvents

    If Not FileExists(sZipPath) Then
        Screen.MousePointer = vbDefault
        ProgressBar1.Visible = False
        MsgBox msOutput, vbCritical, "Error"
        Exit Sub
    End If

    If FileExists(SourceFile) Then Kill SourceFile
    ProgressBar1.Value = 6: DoEvents

    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    oFSO.MoveFile sZipPath, SourceFile
    Set oFSO = Nothing

    ProgressBar1.Value = 7: DoEvents
    
    StatusBar1.Panels(1).Text = "Ready"
    ProgressBar1.Visible = False
    Screen.MousePointer = vbDefault
    bDocumentChange = False
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
    Screen.MousePointer = vbDefault
    ProgressBar1.Visible = False
End Sub

Private Sub AddImage()
    Dim cCDialog As CommonDialog
    Dim xLevel As Long
    Dim ParentNode As Long
    Dim sFileName As String
    Dim sImgDirectory As String
    
    Set cCDialog = New CommonDialog
    
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    xLevel = TreeView1.SelectedItem.Level
    
    If xLevel > 0 Then
        With cCDialog
            .Filter = "Archivos de Imágen|*.jpg;*.Jpeg;*.bmp;*.gif;*.tif;*.ico;*.png|Todos los Archivos|*.*"
            .ShowOpen
            If Len(.FileName) Then
                ParentNode = IIf(xLevel > 1, TreeView1.SelectedItem.Parent.Index, TreeView1.SelectedItem.Index)
                
                AddImageFromFile .FileName, ImageList1
                
                sImgDirectory = sWorkPath & "customUI\images"
                 
                If Not FolderExists(sImgDirectory) Then
                    MkDir sImgDirectory
                End If
                
                sFileName = GetFreeName(EncodeURL(GetFileTitle(.FileName)), sImgDirectory)
                
                With TreeView1.Nodes.Add(ParentNode, TvwNodeRelationshipChild, , GetFreeTvName(GetFileName(.FileName)), ImageList1.ListImages.Count)
                    .Tag = "images/" & sFileName
                    .Selected = True
                End With
                FileCopy .FileName, sImgDirectory & "\" & sFileName
                bDocumentChange = True
            End If
        End With
    End If
End Sub

Private Function GetFreeTvName(ByVal sName As String) As String
    Dim i As Long, n As Long
    Dim isOk As Boolean

    Do While isOk = False
        isOk = True
        For i = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(i).Text = sName Then
                n = n + 1
                sName = "rId" & n
                isOk = False
                Exit For
            End If
        Next
    Loop
    GetFreeTvName = sName
End Function

Private Sub ReadRelationship(ByVal sFile As String, ParentNode As Long)
    Dim XDoc As Object 'MSXML2.DOMDocument
    Dim listNode As Object 'MSXML2.IXMLDOMElement
    Dim sID As String
    Dim sPathImage As String

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load sFile
    
    For Each listNode In XDoc.documentElement.childNodes
        sID = listNode.Attributes.getNamedItem("Id").Text
        sPathImage = listNode.Attributes.getNamedItem("Target").Text
        
        AddImageFromFile sWorkPath & "customUI\" & sPathImage, ImageList1
        
        TreeView1.Nodes.Add(ParentNode, TvwNodeRelationshipChild, , sID, ImageList1.ListImages.Count).Tag = sPathImage
    Next

    Set XDoc = Nothing
End Sub

Public Sub OpenFile(Optional sFilePath As String)
    Dim FSO As Object
    Dim sName As String
    Dim sFile As String
    Dim FF As Integer
    Dim lFileHead As Long
    Dim i As Long
    Dim SmallIcon As Long
    Dim cCDialog As CommonDialog
    Dim lhLib As Long

    On Error Resume Next
    
    If bDocumentChange Then
        Select Case MsgBox("Do you want to save changes?", vbYesNoCancel, Me.Caption)
            Case vbYes
                Call SaveChanges
            Case vbCancel
                Exit Sub
            'Case vbNo
        End Select
    End If
    
    If Len(sFilePath) Then
        sFile = sFilePath
    Else
        Set cCDialog = New CommonDialog
        cCDialog.Filter = "OOXML Document (*.???x;*.???m)|*.???x;*.???m|Microsoft Word Documents(*.do?x;*.do?m)|*.do?x;*.do?m|Microsoft Excel Workbook(*.xl?x;*.xl?m)|*.xl?x;*.xl?m|Microsoft PowerPoint Presentations(*.pp?x;*.pp?m;*.potx;*.potm;*.thmx)|*.pp?x;*.pp?m;*.potx;*.potm;*.thmx|All Files|*.*"
    
        cCDialog.ShowOpen
        If Len(cCDialog.FileName) = 0 Then Exit Sub
        
        sFile = cCDialog.FileName
    End If
    
    ProgressBar1.Max = 10
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    sCustomUI14 = vbNullString
    sCustomUI = vbNullString
    TreeView1.Nodes.Clear
    RTB1.Text = vbNullString
    RTB1.Tag = vbNullString
    
    For i = ImageList1.ListImages.Count To 9 Step -1
        ImageList1.ListImages.Remove i
    Next
    ProgressBar1.Value = 1
    If Len(sFile) = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    FF = FreeFile
    Open sFile For Binary Access Read Lock Read Write As #FF
        Get #FF, 1, lFileHead
    Close #FF
       
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbInformation, "Error " & Err.Number
        Screen.MousePointer = vbDefault
        StatusBar1.Panels(1).Text = ""
        ProgressBar1.Visible = False
        Exit Sub
    End If
    
    If lFileHead <> &H4034B50 Then 'zip head
        MsgBox "Unsupported file type", vbInformation
        Screen.MousePointer = vbDefault
        StatusBar1.Panels(1).Text = ""
        ProgressBar1.Visible = False
        Exit Sub
    End If
    
    SmallIcon = GetFileIcon(sFile)
    ImageList_ReplaceIcon ImageList1.hImageList, 0, SmallIcon
    
    ProgressBar1.Value = 2
    SourceFile = sFile
    
    Set FSO = CreateObject("scripting.filesystemobject")
    
    If Len(sTempDir) Then
        If FSO.FolderExists(sTempDir) Then
            FSO.DeleteFolder sTempDir, True
        End If
    End If
    On Error GoTo 0
    
    DoEvents
    ProgressBar1.Value = 3
    sFullName = GetFileTitle(SourceFile)
    sName = GetFileName(SourceFile)

    StatusBar1.Panels(1).Text = "Open " & GetFileTitle(SourceFile)
 
    DoEvents

    sTempDir = Environ("Temp") & "\customUI-" & sName
    
    On Error Resume Next
    
    If FolderExists(sTempDir) Then
        FSO.DeleteFolder sTempDir, True
        Do While Err.Number <> 0
            Err.Clear
            FSO.DeleteFolder sTempDir
        Loop
    End If
    
    ProgressBar1.Value = 4
    
    FSO.CreateFolder sTempDir
    Do While Err.Number <> 0
        Err.Clear
        FSO.CreateFolder sTempDir
    Loop
    
    On Error GoTo 0
    ProgressBar1.Value = 5
    DoEvents
    
    sZipPath = sTempDir & "\" & sName & ".zip"
    
    FSO.CopyFile SourceFile, sZipPath, True
    ProgressBar1.Value = 6
    sWorkPath = sTempDir & "\" & sName & "\"
    
    FSO.CreateFolder sWorkPath
      
    '-------------
    lhLib = LoadLibrary(App.Path & "\Unzip32.dll")
    VBUnzip sZipPath, sWorkPath, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0
    FreeLibrary lhLib

    ProgressBar1.Value = 7
    TreeView1.Nodes.Clear
    TreeView1.Nodes.Add , , , sFullName, 1
    
    If FSO.FolderExists(sWorkPath & "customUI") Then
        If FSO.FileExists(sWorkPath & "customUI\customUI14.xml") Then
            TreeView1.Nodes.Add 1, TvwNodeRelationshipChild, "customUI14", "customUI14.xml", 2

            sCustomUI14 = ReadFileXML(sWorkPath & "customUI\customUI14.xml")
             
            If FileExists(sWorkPath & "customUI\_rels\customUI14.xml.rels") Then
                ReadRelationship sWorkPath & "customUI\_rels\customUI14.xml.rels", 2
            
            End If
        End If
        ProgressBar1.Value = 8
        If FSO.FileExists(sWorkPath & "customUI\customUI.xml") Then
            TreeView1.Nodes.Add 1, TvwNodeRelationshipChild, "customUI", "customUI.xml", 2
            sCustomUI = ReadFileXML(sWorkPath & "customUI\customUI.xml")
            
            If FSO.FileExists(sWorkPath & "customUI\_rels\customUI.xml.rels") Then
                ReadRelationship sWorkPath & "customUI\_rels\customUI.xml.rels", TreeView1.Nodes.Count
            End If
        End If
        ProgressBar1.Value = 9
    Else
        MkDir sWorkPath & "customUI"
    End If
    
    TreeView1.Nodes(1).Expanded = True
    If TreeView1.Nodes.Count > 1 Then
        TreeView1.Nodes(2).Selected = True
        TreeView1_NodeClick TreeView1.Nodes(2), 1
        EnabledControls True
    End If
    
    ToolBar1.Buttons(2).Enabled = True
    ProgressBar1.Value = 10
    
    Screen.MousePointer = vbDefault
    StatusBar1.Panels(1).Text = "Ready"
    Me.Caption = GetFileTitle(SourceFile) & " MSO Custom UI"
    ProgressBar1.Visible = False
    bDocumentChange = False
    
End Sub

Private Sub ColorearTexto()
    bOn = True
    SyntaxHighlightXML RTB1
    bOn = False
    bTextChange = False
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Boolean)
    If TreeView1.SelectedItem.Level < 2 Then Cancel = True
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As TvwNode, ByVal Button As Integer)
    
    If Node.Level > 0 Then
        ToolBar1.Buttons(4).Enabled = True
        ToolBar1.Buttons(5).Enabled = True
        RTB1.Locked = False
        MnuInsert.Enabled = True
    End If
    
    If Button = 1 Then
        If Node.Level = 1 Then

            If Len(RTB1.Text) Then
                Select Case TreeView1.SelectedItem.Text
                    Case "customUI.xml"
                        sCustomUI = RTB1.Text
                    Case "customUI14.xml"
                        sCustomUI14 = RTB1.Text
                End Select
            End If
            If Node.Text = "customUI.xml" Then
                RTB1.Text = Replace(sCustomUI, Chr$(13), vbNullString)
                RTB1.Tag = "customUI"
            End If
            
            If Node.Text = "customUI14.xml" Then
                RTB1.Text = Replace(sCustomUI14, Chr$(13), vbNullString)
                RTB1.Tag = "customUI14"
            End If
            
            ColorearTexto
            InizializeHistory
           
            Dim i As Long
            For i = 1 To TreeView1.Nodes.Count
                With TreeView1.Nodes(i)
                If .ForeColor = vbBlue Then .ForeColor = TreeView1.ForeColor
                End With
            Next
            Node.ForeColor = vbBlue
        End If
    End If
    
    If Button = 2 Then
        Dim xLevel As Integer
        m_Tv_SelectedItemIndex = Node.Index
    
        Select Case Node.Level
            Case 0: Exit Sub
            Case 1: MnuChangeID.Visible = False
            Case 2: MnuChangeID.Visible = True
        End Select
        
        PopupMenu MnuPopUp
    End If
        
End Sub

Private Function SaveXML(sXML As String, sDestFile As String, Version As String)
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    With xmlDoc
        .SetProperty "ProhibitDTD", False
        .validateOnParse = True
        .async = False
        .loadXML sXML
        .Save sWorkPath & sDestFile
    End With
    AddRelationship sDestFile, Version
End Function

Private Sub WriteRelationship(tvKey As String)
    Dim XDoc As Object 'MSXML2.DOMDocument
    Dim listNode As Object 'MSXML2.IXMLDOMElement
    Dim Elemnt As Object 'MSXML2.IXMLDOMElement
    Dim i As Long
    Dim n As TvwNode
    Dim Node, Attr, Root

    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Key = tvKey Then
            Set n = TreeView1.Nodes(i)
        End If
    Next
    
    If Not n Is Nothing Then
        If n.Children = 0 Then Exit Sub
        Set XDoc = CreateObject("MSXML2.DOMDocument")
        With XDoc
            .async = False
            .validateOnParse = False
            .preserveWhiteSpace = True
        End With
    
        Set Node = XDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8' standalone='yes'")
        XDoc.appendChild Node
        
    
        Set Root = XDoc.createElement("Relationships")
    
        Set Attr = XDoc.createAttribute("xmlns")
        Attr.Value = "http://schemas.openxmlformats.org/package/2006/relationships"
        Root.setAttributeNode Attr
        Set Attr = Nothing

        XDoc.appendChild Root
    
        For i = n.Child.Index To n.Child.Index + n.Children - 1
            Root.appendChild XDoc.createTextNode(vbNewLine + vbTab)
            Set Elemnt = XDoc.createElement("Relationship")
            Elemnt.setAttribute "Target", TreeView1.Nodes(i).Tag
            Elemnt.setAttribute "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Elemnt.setAttribute "Id", TreeView1.Nodes(i).Text
            XDoc.documentElement.appendChild Elemnt
        Next
    
        Root.appendChild XDoc.createTextNode(vbNewLine)
        
        XDoc.loadXML Replace(XDoc.xml, " xmlns=" & Chr(34) & Chr(34), "")
        If Not FolderExists(sWorkPath & "customUI\_rels\") Then MkDir sWorkPath & "customUI\_rels\"
        XDoc.Save sWorkPath & "customUI\_rels\" & tvKey & ".xml.rels"
    End If
    
    Set XDoc = Nothing

End Sub

Private Sub WriteContentTypes()
    Dim XDoc As Object 'MSXML2.DOMDocument
    Dim listNode As Object 'MSXML2.IXMLDOMElement
    Dim Elemnt As Object 'MSXML2.IXMLDOMElement
    Dim i As Long, j As Long
    Dim cExtension As Collection
    Dim sFileExt As String
    Dim bExist As Boolean
    Dim sDir As String
    Set cExtension = New Collection
    
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (sWorkPath & "[Content_Types].xml")
  
    For Each listNode In XDoc.documentElement.childNodes
        If listNode.nodeName = "Default" Then
            For i = 0 To listNode.Attributes.Length - 1
                If listNode.Attributes(i).nodeName = "Extension" Then
                    cExtension.Add listNode.Attributes(i).Text
                End If
            Next
        End If
    Next

    sDir = Dir$(sWorkPath & "customUI\images\")
    
    Do While Len(sDir)
        sFileExt = GetFileExtention(sDir)
        bExist = False
        For i = 1 To cExtension.Count
            If UCase(sFileExt) = UCase(cExtension(i)) Then
                bExist = True
            End If
        Next
        
        If Not bExist Then
            cExtension.Add sFileExt
            Set Elemnt = XDoc.createElement("Default")
            Elemnt.setAttribute "ContentType", "image/." & sFileExt 'GetMimeType(sFileExt)
            Elemnt.setAttribute "Extension", LCase(sFileExt)
            XDoc.documentElement.appendChild Elemnt
        End If
        sDir = Dir$()
    Loop
    
    XDoc.loadXML Replace(XDoc.xml, " xmlns=" & Chr(34) & Chr(34), "")
    XDoc.Save sWorkPath & "[Content_Types].xml"
        
    Set XDoc = Nothing

End Sub

Private Sub AddRelationship(ByVal sFileName As String, Version As String)
    Dim XDoc As Object 'MSXML2.DOMDocument
    Dim listNode As Object 'MSXML2.IXMLDOMElement
    Dim Elemnt As Object 'MSXML2.IXMLDOMElement
    Dim cId As Collection
    Dim bExist As Boolean
    Dim i As Long
    Dim ID As Integer


    Set cId = New Collection

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (sWorkPath & "_rels\.rels")

  
    For Each listNode In XDoc.documentElement.childNodes
        For i = 0 To listNode.Attributes.Length - 1
            If listNode.Attributes(i).nodeName = "Target" Then
                If listNode.Attributes(i).nodeValue = sFileName Then
                    Exit Sub
                End If
            End If
            
            If listNode.Attributes(i).nodeName = "Id" Then
                cId.Add listNode.Attributes(i).Text
            End If
        Next
    Next
        
    Do
       ID = ID + 1
       bExist = False
       For i = 1 To cId.Count
           If cId(i) = "rId" & ID Then
               bExist = True
               Exit For
           End If
       Next
       If bExist = False Then Exit Do
    Loop
    

    Set Elemnt = XDoc.createElement("Relationship")
    Elemnt.setAttribute "Target", sFileName
    Elemnt.setAttribute "Type", "http://schemas.microsoft.com/office/" & Version & "/relationships/ui/extensibility"
    Elemnt.setAttribute "Id", "rId" & ID
    
    XDoc.documentElement.appendChild Elemnt

    XDoc.loadXML Replace(XDoc.xml, " xmlns=" & Chr(34) & Chr(34), "")
    XDoc.Save sWorkPath & "_rels\.rels"
    
    Set XDoc = Nothing
End Sub

 Public Function GetColPos(tBox As Object) As Long
   GetColPos = tBox.SelStart - SendMessage(tBox.HWnd, EM_LINEINDEX, -1&, 0&)
 End Function
