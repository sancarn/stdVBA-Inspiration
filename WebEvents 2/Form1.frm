VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'how to work with COM-EventDelegates in an elevated IE-Control

Private wb, wbExt As VBControlExtender, WithEvents DocEvents As cDocEvents, cnv1 As Object
Attribute wb.VB_VarHelpID = -1
Attribute DocEvents.VB_VarHelpID = -1

Private Sub Form_Load()
  With New cIEFeatures
    .FEATURE_BROWSER_EMULATION = Int(Val(.InstalledVersion)) 'elevate the Browser-Version from its default-version 7
  End With
  
  Set wbExt = Controls.Add("Shell.Explorer.2", "wb") 'only after the above went through, are we allowed to create a BrowserControl
      wbExt.Visible = True
      
  Set wb = wbExt.object
      wb.navigate2 "about:blank"
  Do: DoEvents: Loop Until wb.readyState = 4 '<- READYSTATE_COMPLETE
  
  'no need, to load anything from disk, we can directly apply the template from a String...
  wb.Document.write "<!DOCTYPE HTML><html><head>"
  wb.Document.write "<meta http-equiv='msThemeCompatible' content='yes'><meta charset='UTF-8'></head>"
  wb.Document.write "<body id='bdy1' style='font-family:Arial;font-size:10.5pt' >"
  wb.Document.write "  <input id='txt1' value='Text1'>"
  wb.Document.write "  <button id='btn1'>Click Me</button><br><br>Draw on me with the Mouse:<br>"
  wb.Document.write "  <canvas id='cnv1' width='200px' height='200px' style='border:1px solid red;'></canvas>"
  wb.Document.write "  <br><button id='btn2'>Fill the Canvas with a gradient</button>"
  wb.Document.write "</body></html>"
  
  Set DocEvents = New cDocEvents
      DocEvents.InitOn wb.Document

      DocEvents.AddListenerFor "bdy1", "onmousemove"

      DocEvents.AddListenerFor "btn1", "onclick"
      DocEvents.AddListenerFor "btn2", "onclick"

      DocEvents.AddListenerFor "txt1", "onkeypress"
      DocEvents.AddListenerFor "txt1", "oncut"
      DocEvents.AddListenerFor "txt1", "oncopy"
      DocEvents.AddListenerFor "txt1", "onpaste"
  
  On Error Resume Next 'let's try to instantiate the above defined canvas in a VB-variable (should work from IE9 onwards)
    Set cnv1 = wb.Document.getElementById("cnv1").getContext("2d")
  On Error GoTo 0
  
  If Not cnv1 Is Nothing Then 'let's capture the Mouse-Events on the Canvas, to draw simple lines
    DocEvents.AddListenerFor "cnv1", "onmousedown"
    DocEvents.AddListenerFor "cnv1", "onmousemove"
    DocEvents.AddListenerFor "cnv1", "onmouseup"
  End If
End Sub

'the EventHandler-Routine below will receive the Events from all the Controls on the Page, you priorily connected to
Private Sub DocEvents_DocEvent(Element, ID, EventName, E, AllowFurtherProcessing As Boolean)
Static MouseButton As Long
  Select Case EventName
    Case "onmousedown"
      MouseButton = E.which 'store the current MouseButton
      If ID = "cnv1" Then
        cnv1.beginPath
        cnv1.moveTo E.offsetX, E.offsetY
      End If
    Case "onmouseup"
      MouseButton = 0
      If ID = "cnv1" Then
        cnv1.closePath
      End If
    Case "onmousemove"
      Caption = "MouseX :" & E.offsetX & ", MouseY: " & E.offsetY
      If ID = "cnv1" And MouseButton <> 0 Then  'we are over the canvas
         cnv1.strokeStyle = "blue"
         cnv1.lineTo E.offsetX, E.offsetY
         cnv1.stroke
      End If
    Case "onclick":
      If ID = "btn1" Then MsgBox "Hello from " & ID
      If ID = "btn2" Then 'fill the canvas with a vertical gradient
        Dim Pat As Object
        Set Pat = cnv1.createLinearGradient(0, 0, 0, 200)
            Pat.addColorStop 0, "#AAAAAA"
            Pat.addColorStop 1, "#FFFFFF"
        Set cnv1.FillStyle = Pat
        cnv1.fillRect 0, 0, 200, 200
      End If
    Case "onkeypress"
      Debug.Print "Key: "; Chr(E.keyCode), "KeyCode: "; E.keyCode
      If InStr("0123456789", Chr(E.keyCode)) Then 'disallow numeric digits
        AllowFurtherProcessing = False
      End If
    Case Else
      Debug.Print EventName
  End Select
End Sub

Private Sub Form_Resize()
  If Not wbExt Is Nothing Then wbExt.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
