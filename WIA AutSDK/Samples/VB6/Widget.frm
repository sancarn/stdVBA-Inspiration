VERSION 5.00
Object = "{94A0E92D-43C0-494E-AC29-FD45948A5221}#1.0#0"; "wiaaut.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Properties"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1508
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin WIACtl.VideoPreview VideoPreview1 
      Height          =   2895
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2415
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4260
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   2745
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2895
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin WIACtl.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   2640
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   8655
   End
   Begin WIACtl.DeviceManager wia 
      Left            =   5640
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentItem As Item
Dim CurrentDevice As Device
Dim MinHeight As Long
Dim MinWidth As Long
Dim BottomMargin As Long
Dim RightMargin As Long
Dim MiddleMargin As Long

Private Function PropType(id As WiaPropertyType) As String
    Select Case id
    Case BooleanPropertyType
        PropType = "Boolean"
    Case BytePropertyType
        PropType = "Byte"
    Case ClassIDPropertyType
        PropType = "Class ID"
    Case CurrencyPropertyType
        PropType = "Currency"
    Case DatePropertyType
        PropType = "Date"
    Case DoublePropertyType
        PropType = "Double"
    Case ErrorCodePropertyType
        PropType = "Error Code"
    Case FileTimePropertyType
        PropType = "File Time"
    Case HandlePropertyType
        PropType = "Handle"
    Case IntegerPropertyType
        PropType = "Integer"
    Case LargeIntegerPropertyType
        PropType = "Large Integer"
    Case LongPropertyType
        PropType = "Long"
    Case ObjectPropertyType
        PropType = "Object"
    Case SinglePropertyType
        PropType = "Single"
    Case StringPropertyType
        PropType = "String"
    Case UnsignedIntegerPropertyType
        PropType = "Unsigned Integer"
    Case UnsignedLargeIntegerPropertyType
        PropType = "Unsigned Large Integer"
    Case UnsignedLongPropertyType
        PropType = "Unsigned Long"
    Case VariantPropertyType
        PropType = "Variant"
    Case VectorOfBooleansPropertyType
        PropType = "Vector Of Booleans"
    Case VectorOfBytesPropertyType
        PropType = "Vector Of Bytes"
    Case VectorOfClassIDsPropertyType
        PropType = "Vector Of Class IDs"
    Case VectorOfCurrenciesPropertyType
        PropType = "Vector Of Currencies"
    Case VectorOfDatesPropertyType
        PropType = "Vector Of Dates"
    Case VectorOfDoublesPropertyType
        PropType = "Vector Of Doubles"
    Case VectorOfErrorCodesPropertyType
        PropType = "Vector Of Error Codes"
    Case VectorOfFileTimesPropertyType
        PropType = "Vector Of File Times"
    Case VectorOfIntegersPropertyType
        PropType = "Vector Of Integers"
    Case VectorOfLargeIntegersPropertyType
        PropType = "Vector Of Large Integers"
    Case VectorOfLongsPropertyType
        PropType = "Vector Of Longs"
    Case VectorOfSinglesPropertyType
        PropType = "Vector Of Singles"
    Case VectorOfStringsPropertyType
        PropType = "Vector Of Strings"
    Case VectorOfUnsignedIntegersPropertyType
        PropType = "Vector Of Unsigned Integers"
    Case VectorOfUnsignedLargeIntegersPropertyType
        PropType = "Vector Of Unsigned Large Integers"
    Case VectorOfUnsignedLongsPropertyType
        PropType = "Vector Of Unsigned Longs"
    Case VectorOfVariantsPropertyType
        PropType = "Vector Of Variants"
    Case Else
        PropType = "Unsupported"
    End Select
End Function

Private Sub EnumCommands(ByRef cmds As DeviceCommands)
    Dim i As Integer
    Dim cmd As DeviceCommand
    Dim li As ListItem
    
    ListView3.ListItems.Clear
    
    If cmds.Count > 0 Then
        For Each cmd In cmds
            Set li = ListView3.ListItems.Add(, , cmd.CommandID)
            li.SubItems(1) = cmd.Name
            li.SubItems(2) = cmd.Description
        Next
    End If
End Sub

Private Sub EnumEvents(ByRef evts As DeviceEvents)
    Dim i As Integer
    Dim evt As DeviceEvent
    Dim li As ListItem
    
    ListView4.ListItems.Clear
    
    If evts.Count > 0 Then
        For Each evt In evts
            Set li = ListView4.ListItems.Add(, , evt.EventID)
            li.SubItems(1) = evt.Name
            li.SubItems(2) = evt.Description
        Next
    End If
End Sub

Private Sub EnumFormats(ByRef fmts As Formats)
    Dim v As Variant
    Dim i As Integer
    
    ListView2.ListItems.Clear
    
    If fmts.Count > 0 Then
        For Each v In fmts
            ListView2.ListItems.Add , , v
        Next
    End If
End Sub

Private Sub EnumProperties(ByRef props As Properties, ByRef ThumbnailData As Vector, ByRef ThumbnailWidth As Long, ByRef ThumbnailHeight As Long)
    Dim p As Property
    Dim i As Long
    Dim s As String
    Dim li As ListItem
    Dim v As Vector
    Dim pc As Integer

    On Error GoTo BadDriver
    
    ListView1.ListItems.Clear
    
    pc = 0
    
    If props.Count > 0 Then
        For Each p In props
            Set li = ListView1.ListItems.Add(, , p.Name)
            Set li.Tag = p
            li.SubItems(1) = p.PropertyID
            li.SubItems(2) = PropType(p.Type)
            If Not p.IsVector Then
                li.SubItems(3) = p.Value
            Else
                s = ""
                Set v = p.Value
                For i = 1 To v.Count
                    s = s & CStr(v(i))
                    If i < 50 Then
                        If i <> v.Count Then
                            s = s & ", "
                        End If
                    Else
                        s = s & ", ... (Size = " & v.Count & ")"
                        Exit For
                    End If
                Next
                li.SubItems(3) = s
            End If
            If p.Name = "Thumbnail Data" Then Set ThumbnailData = p.Value
            If p.Name = "Thumbnail Width" Then ThumbnailWidth = p.Value
            If p.Name = "Thumbnail Height" Then ThumbnailHeight = p.Value
            Select Case p.SubType
            Case ListSubType
                li.SubItems(4) = "List"
                li.SubItems(5) = p.SubTypeDefault
                s = ""
                Set v = p.SubTypeValues
                For i = 1 To v.Count
                    s = s & CStr(v(i))
                    If i < 50 Then
                        If i <> v.Count Then
                            s = s & ", "
                        End If
                    Else
                        s = s & ", ... (Size = " & v.Count & ")"
                        Exit For
                    End If
                Next
                li.SubItems(6) = s
            Case FlagSubType
                li.SubItems(4) = "Flags"
                li.SubItems(5) = p.SubTypeDefault
                s = ""
                Set v = p.SubTypeValues
                For i = 1 To v.Count
                    s = s & CStr(v(i))
                    If i <> v.Count Then
                        s = s & ", "
                    End If
                Next
                li.SubItems(6) = s
            Case RangeSubType
                li.SubItems(4) = "Range"
                li.SubItems(5) = p.SubTypeDefault
                li.SubItems(6) = "Min = " & p.SubTypeMin & ", Max = " & p.SubTypeMax & ", Step = " & p.SubTypeStep
            Case Else
                li.SubItems(4) = "Unspecified"
            End Select
BadDriverResume:
            pc = pc + 1
        Next
    End If
    
    Exit Sub
BadDriver:
    If Err.Number = -2145320836 Then
        MsgBox Err.Description, vbOKOnly, p.Name & " Property"
        Err.Clear
        Resume BadDriverResume
    Else
        MsgBox Err.Description
        Resume
    End If
End Sub

Private Sub EnumChildren(ByRef itms As Items, ByRef nde As Node)
    Dim itm As Item
    Dim newNode As Node
    
    For Each itm In itms
        Set newNode = TreeView1.Nodes.Add(nde.Index, tvwChild, , itm.Properties("Item Name").Value)
        Set newNode.Tag = itm
        
        If itm.Items.Count > 0 Then EnumChildren itm.Items, newNode
    Next
End Sub

Private Sub BuildTree()

    Dim di As DeviceInfo
    Dim dev As Device
    Dim nde As Node
    
    ListView1.ListItems.Clear
    TreeView1.Nodes.Clear
        
    For Each di In wia.DeviceInfos
        Set dev = di.Connect
        If Not dev Is Nothing Then
            Set nde = TreeView1.Nodes.Add(, , , di.Properties("Name").Value)
            Set nde.Tag = dev
            
            EnumChildren dev.Items, nde
        End If
    Next
        
End Sub

Private Sub Command1_Click()
    If Not CurrentItem Is Nothing Then
        Dim img As ImageFile
        Set img = CommonDialog1.ShowTransfer(CurrentItem)
        If Not img Is Nothing Then
            Set Image1.Picture = img.FileData.Picture
            ListView1.Visible = False
        End If
    End If
End Sub

Private Sub ListView3_DblClick()
    Dim li As ListItem
    Set li = ListView3.SelectedItem
    Dim itm As Item
    Dim img As ImageFile
    If Not CurrentDevice Is Nothing Then
        Set itm = CurrentDevice.ExecuteCommand(li.Text)
        If Not itm Is Nothing Then
            Set img = CommonDialog1.ShowTransfer(itm)
            If Not img Is Nothing Then
                Set Image1.Picture = img.FileData.Picture
                ListView1.Visible = False
            End If
        End If
    End If
    If Not CurrentItem Is Nothing Then
        Set itm = CurrentItem.ExecuteCommand(li.Text)
        If Not itm Is Nothing Then
            Set img = CommonDialog1.ShowTransfer(itm)
            If Not img Is Nothing Then
                Set Image1.Picture = img.FileData.Picture
                ListView1.Visible = False
            End If
        End If
    End If
End Sub

Private Sub ListView4_DblClick()
    Dim li As ListItem
    Set li = ListView4.SelectedItem
    If Not CurrentDevice Is Nothing Then
        On Error Resume Next
        wia.RegisterEvent li.Text, CurrentDevice.DeviceID
        If Err.Number <> 0 Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub Image1_Click()
    ListView1.Visible = True
End Sub

Private Sub Form_Resize()
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight
    
    ListView1.Width = Me.Width - ListView1.Left - RightMargin
    ListView1.Height = Me.Height - ListView1.Top - BottomMargin
    Image1.Width = ListView1.Width
    Image1.Height = ListView1.Height
    
    ListView2.Left = Me.Width - ListView2.Width - RightMargin
    VideoPreview1.Left = ListView2.Left
    
    TreeView1.Width = ListView2.Left - TreeView1.Left - MiddleMargin
    Command2.Width = TreeView1.Width
End Sub

Private Sub Form_Load()
    MinHeight = Me.Height
    MinWidth = Me.Width
    BottomMargin = Me.Height - ListView1.Top - ListView1.Height
    RightMargin = Me.Width - ListView1.Left - ListView1.Width
    MiddleMargin = ListView2.Left - TreeView1.Left - TreeView1.Width
    
    ListView1.ColumnHeaders.Add , , "Name", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Id", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Type", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Value", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "SubType", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Default", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "Values", ListView1.Width / 7
    ListView1.View = lvwReport
    
    ListView2.ColumnHeaders.Add , , "Format", ListView2.Width
    ListView2.View = lvwReport
    
    ListView3.ColumnHeaders.Add , , "Id", ListView3.Width / 3
    ListView3.ColumnHeaders.Add , , "Name", ListView3.Width / 3
    ListView3.ColumnHeaders.Add , , "Description", ListView3.Width / 3
    ListView3.View = lvwReport
    
    ListView4.ColumnHeaders.Add , , "Id", ListView3.Width / 3
    ListView4.ColumnHeaders.Add , , "Name", ListView3.Width / 3
    ListView4.ColumnHeaders.Add , , "Description", ListView3.Width / 3
    ListView4.View = lvwReport
    
    BuildTree
    
    wia.RegisterEvent wiaEventDeviceConnected
    wia.RegisterEvent wiaEventDeviceDisconnected

End Sub

Private Sub Command2_Click()
    If Not CurrentItem Is Nothing Then
        CommonDialog1.ShowItemProperties CurrentItem
    End If
    
    If Not CurrentDevice Is Nothing Then
        CommonDialog1.ShowDeviceProperties CurrentDevice
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim itm As Item
    Dim dev As Device
    Dim ThumbnailData As Vector
    Dim ThumbnailWidth As Long
    Dim ThumbnailHeight As Long
    
    Set ThumbnailData = Nothing
    ThumbnailWidth = 0
    ThumbnailHeight = 0
    
    If TypeOf Node.Tag Is Item Then
        Set itm = Node.Tag
        Set CurrentDevice = Nothing
        Set CurrentItem = itm
        'Set Picture1.Picture = itm.Thumbnail
        EnumProperties itm.Properties, ThumbnailData, ThumbnailWidth, ThumbnailHeight
        If Not ThumbnailData Is Nothing Then
            If ThumbnailWidth <> 0 Then
                If ThumbnailHeight <> 0 Then
                    Set Picture1.Picture = ThumbnailData.Picture(ThumbnailWidth, ThumbnailHeight)
                End If
            End If
        End If
        EnumFormats itm.Formats
        EnumCommands itm.Commands
        ListView2.ZOrder
        Picture1.ZOrder
    Else
        Set dev = Node.Tag
        Set CurrentItem = Nothing
        Set CurrentDevice = dev
        If dev.Type = VideoDeviceType Then
            Set VideoPreview1.Device = dev
        End If
        EnumProperties dev.Properties, ThumbnailData, ThumbnailWidth, ThumbnailHeight
        EnumCommands dev.Commands
        EnumEvents dev.Events
        ListView4.ZOrder
        VideoPreview1.ZOrder
    End If
End Sub

Private Sub wia_OnEvent(ByVal EventID As String, ByVal DeviceID As String, ByVal ItemID As String)
    If EventID = wiaEventDeviceConnected Then
        BuildTree
    ElseIf EventID = wiaEventDeviceDisconnected Then
        BuildTree
    Else
        MsgBox "RECEIVED EVENT!" & vbCrLf & vbCrLf & _
        "EventID = """ & EventID & """" & vbCrLf & _
        "DeviceID = """ & DeviceID & """" & vbCrLf & _
        "ItemID = """ & ItemID & """"
    End If
End Sub
