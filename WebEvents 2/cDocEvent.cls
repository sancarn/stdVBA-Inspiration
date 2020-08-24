VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDocEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent, Elmt, ID, EventName

Friend Sub InitWith(Parent, Elmt, ID, EventName)
  On Error Resume Next
  Set Me.Parent = Parent
  Set Me.Elmt = Elmt
      Me.ID = ID
      Me.EventName = EventName
  Connect
End Sub

'needs the "Default"-> extended Procedure-Attribute
Public Function EventHandler() As Boolean
Attribute EventHandler.VB_UserMemId = 0
Dim AllowFurtherProcessing As Boolean
    AllowFurtherProcessing = True 'set the default to further up-bubbling
    Parent.Delegate Me, AllowFurtherProcessing 'the Parent then raises a VB-Event
    EventHandler = AllowFurtherProcessing
End Function

Private Sub Connect() 'since CallByName apparently doesn't work in non-elevated Browser, here the workaround
  On Error Resume Next
  Select Case EventName
    Case "onactivate": Set Elmt.onactivate = Me
    Case "onafterupdate": Set Elmt.onafterupdate = Me
    Case "onbeforeactivate": Set Elmt.onbeforeactivate = Me
    Case "onbeforecopy": Set Elmt.onbeforecopy = Me
    Case "onbeforecut": Set Elmt.onbeforecut = Me
    Case "onbeforedeactivate": Set Elmt.onbeforedeactivate = Me
    Case "onbeforeeditfocus": Set Elmt.onbeforeeditfocus = Me
    Case "onbeforepaste": Set Elmt.onbeforepaste = Me
    Case "onbeforeupdate": Set Elmt.onbeforeupdate = Me
    Case "onblur": Set Elmt.onblur = Me
    Case "oncellchange": Set Elmt.oncellchange = Me
    Case "onclick": Set Elmt.onclick = Me
    Case "oncontextmenu": Set Elmt.oncontextmenu = Me
    Case "oncontrolselect": Set Elmt.oncontrolselect = Me
    Case "oncopy": Set Elmt.oncopy = Me
    Case "oncut": Set Elmt.oncut = Me
    Case "ondataavailable": Set Elmt.ondataavailable = Me
    Case "ondatasetchanged": Set Elmt.ondatasetchanged = Me
    Case "ondatasetcomplete": Set Elmt.ondatasetcomplete = Me
    Case "ondblclick": Set Elmt.ondblclick = Me
    Case "ondeactivate": Set Elmt.ondeactivate = Me
    Case "ondrag": Set Elmt.ondrag = Me
    Case "ondragend": Set Elmt.ondragend = Me
    Case "ondragenter": Set Elmt.ondragenter = Me
    Case "ondragleave": Set Elmt.ondragleave = Me
    Case "ondragover": Set Elmt.ondragover = Me
    Case "ondragstart": Set Elmt.ondragstart = Me
    Case "ondrop": Set Elmt.ondrop = Me
    Case "onerrorupdate": Set Elmt.onerrorupdate = Me
    Case "onfilterchange": Set Elmt.onfilterchange = Me
    Case "onfocus": Set Elmt.onfocus = Me
    Case "onfocusin": Set Elmt.onfocusin = Me
    Case "onfocusout": Set Elmt.onfocusout = Me
    Case "onhelp": Set Elmt.onhelp = Me
    Case "onkeydown": Set Elmt.onkeydown = Me
    Case "onkeypress": Set Elmt.onkeypress = Me
    Case "onkeyup": Set Elmt.onkeyup = Me
    Case "onlayoutcomplete": Set Elmt.onlayoutcomplete = Me
    Case "onlosecapture": Set Elmt.onlosecapture = Me
    Case "onmousedown": Set Elmt.onmousedown = Me
    Case "onmouseenter": Set Elmt.onmouseenter = Me
    Case "onmouseleave": Set Elmt.onmouseleave = Me
    Case "onmousemove": Set Elmt.onmousemove = Me
    Case "onmouseout": Set Elmt.onmouseout = Me
    Case "onmouseover": Set Elmt.onmouseover = Me
    Case "onmouseup": Set Elmt.onmouseup = Me
    Case "onmousewheel": Set Elmt.onmousewheel = Me
    Case "onmove": Set Elmt.onmove = Me
    Case "onmoveend": Set Elmt.onmoveend = Me
    Case "onmovestart": Set Elmt.onmovestart = Me
    Case "onpage": Set Elmt.onpage = Me
    Case "onpaste": Set Elmt.onpaste = Me
    Case "onpropertychange": Set Elmt.onpropertychange = Me
    Case "onreadystatechange": Set Elmt.onreadystatechange = Me
    Case "onresize": Set Elmt.onresize = Me
    Case "onresizeend": Set Elmt.onresizeend = Me
    Case "onresizestart": Set Elmt.onresizestart = Me
    Case "onrowenter": Set Elmt.onrowenter = Me
    Case "onrowexit": Set Elmt.onrowexit = Me
    Case "onrowsdelete": Set Elmt.onrowsdelete = Me
    Case "onrowsinserted": Set Elmt.onrowsinserted = Me
    Case "onscroll": Set Elmt.onscroll = Me
    Case "onselectstart": Set Elmt.onselectstart = Me
  End Select
End Sub
