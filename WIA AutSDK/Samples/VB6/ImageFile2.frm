VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   Caption         =   "Properties"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public img As ImageFile

Private Function PropType(id As WiaImagePropertyType) As String
    Select Case id
    Case ByteImagePropertyType
        PropType = "Byte"
    Case LongImagePropertyType
        PropType = "Long"
    Case RationalImagePropertyType
        PropType = "Rational"
    Case StringImagePropertyType
        PropType = "String"
    Case UnsignedIntegerImagePropertyType
        PropType = "Unsigned Integer"
    Case UnsignedLongImagePropertyType
        PropType = "Unsigned Long"
    Case UnsignedRationalImagePropertyType
        PropType = "Unsigned Rational"
    Case VectorOfBytesImagePropertyType
        PropType = "Vector Of Bytes"
    Case VectorOfLongsImagePropertyType
        PropType = "Vector Of Longs"
    Case VectorOfRationalsImagePropertyType
        PropType = "Vector Of Rationals"
    Case VectorOfStringsImagePropertyType
        PropType = "Vector Of Strings"
    Case VectorOfUndefinedImagePropertyType
        PropType = "Vector Of Undefined"
    Case VectorOfUnsignedIntegersImagePropertyType
        PropType = "Vector Of Unsigned Integers"
    Case VectorOfUnsignedLongsImagePropertyType
        PropType = "Vector Of Unsigned Longs"
    Case VectorOfUnsignedRationalsImagePropertyType
        PropType = "Vector Of Unsigned Rationals"
    Case Else
        PropType = "Undefined"
    End Select
End Function

Private Sub EnumProperties(ByRef props As Properties)
    Dim p As Property
    Dim i As Integer
    Dim s As String
    Dim li As ListItem
    Dim v As Vector
    Dim r As Rational

    ListView1.ListItems.Clear
    
    If props.Count > 0 Then
        For Each p In props
            Set li = ListView1.ListItems.Add(, , p.Name)
            Set li.Tag = p
            li.SubItems(1) = p.PropertyID
            li.SubItems(2) = PropType(p.Type)
            If Not p.IsVector Then
                If TypeOf p.Value Is Rational Then
                    Set r = p.Value
                    li.SubItems(3) = r.Numerator & "/" & r.Denominator
                Else
                    li.SubItems(3) = p.Value
                End If
            Else
                s = ""
                Set v = p.Value
                For i = 1 To v.Count
                    If TypeOf v(i) Is Rational Then
                        Set r = v(i)
                        s = s & r.Numerator & "/" & r.Denominator
                    Else
                        s = s & CStr(v(i))
                    End If
                    If i < 50 Then
                        If i <> v.Count Then
                            s = s & ", "
                        End If
                    Else
                        s = s & "... (Size = " & v.Count & ")"
                        Exit For
                    End If
                Next
                li.SubItems(3) = s
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    ListView1.ColumnHeaders.Add , , "Name"
    ListView1.ColumnHeaders.Add , , "Id"
    ListView1.ColumnHeaders.Add , , "Type"
    ListView1.ColumnHeaders.Add , , "Value"
    ListView1.View = lvwReport
End Sub

Private Sub Form_Resize()
    Dim NewWidth As Integer
    Dim NewHeith As Integer
    
    NewWidth = Me.Width - ListView1.Left - 175
    NewHeight = Me.Height - ListView1.Top - 575
    
    If NewWidth < 1000 Then NewWidth = 1000
    If NewHeight < 1000 Then NewHeight = 1000
    ListView1.Width = NewWidth
    ListView1.Height = NewHeight

    ListView1.ColumnHeaders(1).Width = ListView1.Width / 10
    ListView1.ColumnHeaders(2).Width = ListView1.Width / 10
    ListView1.ColumnHeaders(3).Width = ListView1.Width / 10
    ListView1.ColumnHeaders(4).Width = (ListView1.Width * 6.8) / 10
    
    If Not img Is Nothing Then
        If ListView1.ListItems.Count = 0 Then
            EnumProperties img.Properties
        End If
    End If
End Sub
