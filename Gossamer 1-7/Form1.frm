VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Gossamer HTTP Server Demonstration"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin GossDemo1.Gossamer Gossamer 
      Left            =   7260
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      VDir            =   "site"
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3060
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   1620
      TabIndex        =   3
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   360
      Left            =   660
      TabIndex        =   2
      Text            =   "8080"
      Top             =   240
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   8599
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":1642
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'GossDemo1
'=========
'
'A simple application the demonstrates the Gossamer control.
'
'Here simple event logging is done and a simple static page with
'an image is available as well as a GET form and a POST form.
'
'Changes
'-------
'
'Version: 1.1
'
' o Added " processing to EntityEncode.
'
'Version: 1.2
'
' o Use VDirPath property of Gossamer in dynamic requests.
'
'Version: 1.3
'
' no change
'
'Version: 1.4
'
' o Removed EntityEncode function.  Moved to Gossamer as a method.
'
'Version: 1.5
'
' no change
'
'Version: 1.6
'
' o Gossamer_LogEvent updated to handle EventType = getWSSoftError
'   and to handle EventType = getWSError more gracefully so the
'   Start/Stop button enabled state doesn't get messed up.
'
'Version: 1.7
'
' o Declared some color const values in Gossamer_LogEvent().
'

Private Function FormatNumWidth(ByVal Value As Long, ByVal NumWidth As Integer)
    FormatNumWidth = CStr(Value)
    If Len(FormatNumWidth) < NumWidth Then
        FormatNumWidth = Right$(Space$(NumWidth - 1) & FormatNumWidth, NumWidth)
    End If
End Function

Private Sub Log(Optional ByVal Text As String, Optional ByVal Color As ColorConstants = vbBlack)
    With rtbLog
        .SelStart = Len(.Text)
        .SelColor = Color
        .SelText = Text & vbNewLine
    End With
End Sub

Private Sub cmdStart_Click()
    Gossamer.StartListening CLng(txtPort.Text)
    cmdStart.Enabled = False
    txtPort.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Gossamer.StopListening
    cmdStop.Enabled = False
    txtPort.Enabled = True
    cmdStart.Enabled = True
End Sub

Private Sub Form_Load()
    Show
    Log "Gossamer Demo 1 ready"
    Log
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        With rtbLog
            .Height = ScaleHeight - .Top
            .Width = ScaleWidth
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Gossamer.State = sckListening Then
        Gossamer.StopListening
    End If
End Sub

Private Sub Gossamer_DynamicRequest(ByVal Method As String, _
                                    ByVal URI As String, _
                                    ByVal Params As String, _
                                    ByVal ReqHeaders As Collection, _
                                    RespStatus As Single, _
                                    RespStatusText As String, _
                                    RespMIME As String, _
                                    RespExtraHeaders As String, _
                                    RespBody() As Byte, _
                                    ByVal ClientIndex As Integer)
    Dim intFile As Integer
    Dim intParam As Integer
    Dim strDyn As String
    Dim strParams() As String
    Dim strParts() As String
    Dim strRespBody As String
    
    Select Case Method
        Case "GET"
            If LCase$(URI) = "\formget.htm" Then
                intFile = Gossamer.GetFreeFile()
                Open Gossamer.VDirPath & URI For Binary Access Read As #intFile
                strRespBody = Input$(LOF(intFile), #intFile)
                Close #intFile
                
                strParams = Split(Params, "&")
                For intParam = 0 To UBound(strParams)
                    strParts = Split(strParams(intParam), "=")
                    strDyn = strDyn & Gossamer.URLDecode(strParts(0))
                    If UBound(strParts) > 0 Then
                        strDyn = strDyn _
                               & " = " _
                               & Gossamer.EntityEncode(Gossamer.URLDecode(strParts(1))) _
                               & "<BR>"
                    End If
                Next
                strRespBody = Replace$(strRespBody, "<!-- INSERT -->", "<P>" & strDyn & "</P>")
                RespStatus = 200
                RespStatusText = "Ok"
                RespMIME = "text/html"
                RespBody = StrConv(strRespBody, vbFromUnicode)
            End If
        
        Case "POST"
            If LCase$(URI) = "\formpost.htm" Then
                intFile = Gossamer.GetFreeFile()
                Open Gossamer.VDirPath & URI For Binary Access Read As #intFile
                strRespBody = Input$(LOF(intFile), #intFile)
                Close #intFile
                
                strParams = Split(Params, "&")
                For intParam = 0 To UBound(strParams)
                    strParts = Split(strParams(intParam), "=")
                    strDyn = strDyn & Gossamer.URLDecode(strParts(0))
                    If UBound(strParts) > 0 Then
                        strDyn = strDyn _
                               & " = " _
                               & Gossamer.EntityEncode(Gossamer.URLDecode(strParts(1))) _
                               & "<BR>"
                    End If
                Next
                strRespBody = Replace$(strRespBody, "<!-- INSERT -->", "<P>" & strDyn & "</P>")
                RespStatus = 200
                RespStatusText = "Ok"
                RespMIME = Gossamer.ExtensionToMIME("htm")
                RespBody = StrConv(strRespBody, vbFromUnicode)
            End If
        End Select
End Sub

Private Sub Gossamer_LogEvent(ByVal GossEvent As GossEvent, ByVal ClientIndex As Integer)
    Const DARK_GREEN As Long = &H8000&
    Const PURPLE As Long = &H800080
    
    With GossEvent
        Log Format$(.Timestamp, "HH:NN:SS") _
          & FormatNumWidth(ClientIndex, 4) _
          & FormatNumWidth(.EventType, 3) _
          & FormatNumWidth(.EventSubtype, 3) _
          & " " & .IP _
          & " " & .Method _
          & " " & Left$(.Text, 100), _
            Choose(.EventType, vbRed, vbBlue, DARK_GREEN, PURPLE)
        If .EventType = getWSError And ClientIndex = -1 Then
            cmdStop_Click
        End If
    End With
End Sub
