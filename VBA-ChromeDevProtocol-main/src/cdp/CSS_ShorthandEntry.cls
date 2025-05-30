VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCSSShorthandEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ShorthandEntry"
' CSS.ShorthandEntry
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Shorthand name.
Public name AS string

' value : string
' Shorthand value.
Public value AS string

' important : boolean
' Optional
' Whether the property has "!important" annotation (implies `false` if absent).
Public important AS Variant ' boolean

Public Function init(ByVal obj as Dictionary) As cdpCSSShorthandEntry
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let value = CStr(obj.Item("value"))
    If obj.Exists("important") Then Let important = CBool(obj.Item("important"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("value") = value
    if Not IsEmpty(important) Then dict("important") = important

    Set asDictionary = dict
End Function

