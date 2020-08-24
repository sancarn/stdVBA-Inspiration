VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDocEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DocEvent(Element, ID, EventName, E, AllowFurtherProcessing As Boolean)

Private mDoc

Friend Sub InitOn(Document)
  Set mDoc = Document
End Sub

Public Sub AddListenerFor(ID As String, EventName As String)
  Dim NewEvent As New cDocEvent
      NewEvent.InitWith Me, mDoc.getElementById(ID), ID, EventName
End Sub

Public Sub Delegate(Evt As cDocEvent, AllowFurtherProcessing As Boolean)
  RaiseEvent DocEvent(Evt.Elmt, Evt.ID, Evt.EventName, mDoc.parentWindow.event, AllowFurtherProcessing)
End Sub
 
