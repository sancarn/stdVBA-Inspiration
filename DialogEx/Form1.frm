VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Dialogs"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEvents 
      Height          =   2790
      Left            =   255
      TabIndex        =   6
      Top             =   3660
      Width           =   8250
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Show Sample"
      Height          =   375
      Left            =   4530
      TabIndex        =   5
      Top             =   2865
      Width           =   3930
   End
   Begin VB.TextBox txtInfo 
      Height          =   2595
      Left            =   4530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Form1.frx":0000
      Top             =   150
      Width           =   3915
   End
   Begin VB.CheckBox chkMultiSelect 
      Caption         =   "Allow Multiple File Selection"
      Height          =   315
      Left            =   315
      TabIndex        =   4
      Top             =   3210
      Width           =   2625
   End
   Begin VB.OptionButton optEvents 
      Caption         =   "Control Events"
      Height          =   240
      Index           =   2
      Left            =   2955
      TabIndex        =   3
      Top             =   2925
      Width           =   1455
   End
   Begin VB.OptionButton optEvents 
      Caption         =   "Dialog Events"
      Height          =   240
      Index           =   1
      Left            =   1545
      TabIndex        =   2
      Top             =   2925
      Width           =   1335
   End
   Begin VB.OptionButton optEvents 
      Caption         =   "No Events"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   2925
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form1.frx":0006
      Left            =   270
      List            =   "Form1.frx":0031
      TabIndex        =   0
      Top             =   150
      Width           =   4155
   End
   Begin VB.Label Label1 
      Caption         =   "Events: populated from bottom-up"
      Height          =   300
      Left            =   4560
      TabIndex        =   8
      Top             =   3420
      Width           =   3885
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuXP 
         Caption         =   "Show File Dialog"
         Index           =   0
      End
      Begin VB.Menu mnuXP 
         Caption         =   "Show Folder Browser"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' notes:
' - This project uses no TLBs
' - If this project were to use TLBs that define interfaces like IShellItem, IFileDialog, etc then
'   you could simply assign the CmnDialogEx class returned interfaces to the appropriate TLB interface.
' For example, let's say you have a TLB that defines the IShellItem interface. To use your TLB's
'   IShellItem interface, you could simply assign the class' passed interface, i.e,
'   ---------------------------------------------------------
'       Private Sub cBrowser_DialogOnFolderChanging(oShellItem As stdole.IUnknown, Cancel As Boolean)
'           Dim myTlbIShellItem As IShellItem
'           Set myTlbIShellItem = oShellItem ' oShellItem is passed to many of the events
'           ' now call your myTlbIShellItem.DisplayName method
'           ' instead of the class' IShellItem_GetDisplayName method
'       End Sub
'   ---------------------------------------------------------
' Likewise, the class' ShowOpen and ShowSave can return an optional interface containing the selected
'   items from the dialog. That interface will either be an IShellItem or IShellItemArray, depending
'   on number of items selected from the dialog. If you have a TLB that defines both the IShellItem
'   and IShellItemArray interfaces, you can simply assign the return value to your TLB interfaces
'   and call its methods, i.e.,
'   ---------------------------------------------------------
'   Dim pUnk As stdole.IUnknown
'   Dim myTlbShellItem As IShellItem
'   Dim myTlbShellItemArray As IShellItemArray
'   ...
'       If cBrowser.ShowOpen(Me.hWnd, , VarPtr(pUnk)) = True Then
'           If InStr(cBrowser.FileName, vbNullChar) Then    ' multi-selected items else single selection
'               Set myTlbShellItemArray = pUnk
'           Else
'               Set myTlbShellItem = pUnk
'           End If
'       End If
'   ---------------------------------------------------------
' Bottom line. Took into consideration that the CmnDialogEx class could be used in projects that
'   include a TLB with typical Shell interfaces defined. I wanted the class to be compatible with
'   such projects and feel it is. Wherever the class may expect an input as stdole.IUnknown, ensure
'   you do not pass your TLB's IUnknown interface, so declare IUnknown passed to this class
'   with the stdole prefix: stdole.IUnknown

' Should you want to display a MsgBox within an event, recommend this:
'   1. Use unicode-aware notifications, i.e., do not recommend VB's MsgBox function
'   2. Microsoft says to use the hWnd of the dialog for MessageBox or similar APIs
'       For XP/Win2K, the hWnd is provided in the event
'       For Vista+, call the class' IFileDialog_GetHwnd function

Private WithEvents cBrowser As CmnDialogEx   ' must use WithEvents keyword if events are desired
Attribute cBrowser.VB_VarHelpID = -1

Private Sub Form_Load()
    Set cBrowser = New CmnDialogEx
    If cBrowser.Version < dvXP_Win2K Then
        Me.Show
        DoEvents
        ' note: this could actually support any earlier system, but lack of unicode support needs to be
        '   accounted for. The class expects an operating system that supports full unicode.
        '   Class will need modifications to stop preventing O/S less than Win2k/XP since no ANSI APIs used
        MsgBox "This sample project requires Win2k/XP or higher operating system", vbExclamation + vbOKOnly
        Unload Me
        Exit Sub
    End If
    If cBrowser.Version = dvXP_Win2K Then
        List1.Enabled = False
        optEvents(2).Enabled = False
        txtInfo.Text = "This class is very similar to any API version of the common dialog for selecting files." & vbCrLf & vbCrLf & _
            "Since you are on a system less than Vista, only the controls below the list box are enabled, not the list itself."
        Me.Show
        DoEvents
        MsgBox "The examples in the listbox are for Vista and above." & vbCrLf & _
            "They have been disabled since you are running on a lower operating system.", vbInformation + vbOKOnly
    Else
        List1.ListIndex = 0
    End If
End Sub

Private Sub cmdGo_Click()
    
    If cBrowser.Version < dvVista Then
        PopupMenu mnuPopup
    Else
        pvDoSamples List1.ListIndex
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cBrowser = Nothing
End Sub

Private Sub pvDoSamples(Index As Long)

    Dim sItems() As String
    Dim lItem As Long, bReturn As Boolean
    Dim eEvents As EventTypeEnum
    
    ' flags used below to return additional info when dialog closes
    Dim bSplitBtnUsed As Boolean, bReadOnlyUsed As Boolean
    
    cBrowser.Clear
    ' note: if the .FlagsDialog property is not set, the class will use default values based on Open or Save intent

    Select Case Index
    
    ' user-defined control/container IDs can range from 1 to 268,435,455 & cannot be duplicated
    
    Case 0 ' notepad-like example adds combobox next to Open button
        With cBrowser
            .Filter = "Text Documents (*.txt)|*.txt|All Files (*.*)|*.*"
            ' create a container, then add combobox to the container
            .Controls_Add 500, 0, ctlType_Container, "Encoding"
            ' 500 is our container ID; add combobox to the container & 100 is our combobox control ID
            .Controls_Add 500, 100, ctlType_ComboBox, "ANSI", "Unicode", "Unicode big endian", "UTF-8"
        End With
    
    Case 1 ' MS Word-like example adds menu next to open button
        ReDim sItems(0 To 6)
        sItems(0) = "Open": sItems(1) = "Open Read Only"
        sItems(2) = "Open as Copy": sItems(3) = "Open in Browser"
        sItems(4) = "Open with Transform": sItems(5) = "Open in protected view"
        sItems(6) = "Open and Repair"
        With cBrowser
            .Filter = "All Word Documents|*.txt|All Files (*.*)|*.*"
            ' create the menu control
            ' this statement uses the ParamArray parameter to pass multiple sub-items
            .Controls_Add 0, 100, ctlType_Menu, "Tools", "Map Network Drive"
            ' create the split button
            ' this statement uses a string array to pass multiple sub-items
            .Controls_Add 0, 200, ctlType_DefaultButton, "", sItems()
            ' in Word, depending on what type of office document being looked at,
            '   sets enabled state of individual menu items. In this example
            '   we will disable all but the 1st and last
            For lItem = 1 To UBound(sItems) - 1
                .Controls_PropertySet 200, ppEnabled, False, lItem
            Next
        End With
        bSplitBtnUsed = True
    
    Case 2 ' add read-only option as checkbox
        cBrowser.Controls_AddReadOnlyOption 100, True
        bReadOnlyUsed = True
        
    Case 3 ' add read-only option as a split-button on the default Open button
        cBrowser.Controls_AddReadOnlyOption 200, False
        bSplitBtnUsed = True
        bReadOnlyUsed = True
    
    Case 4 ' allow navigation into Zip files (key flag: DLGex_AllNonStorageItems)
        cBrowser.FlagsDialog = DLG_FileMustExist Or DLGex_AllNonStorageItems
        
    Case 5 ' allow navigation into Zip files & enable selection of zip folder
        cBrowser.Controls_SetCustomMode cm_CompressedFolderPlusFiles
        
    Case 6 ' standard open dialog + hide filter dropdown
        cBrowser.FlagsDialog = DLG__BaseOpenDialogFlags
        cBrowser.Filter = vbNullChar ' flag to hide the filter
        
    Case 7 ' standard browse for folders
        cBrowser.FlagsDialog = DLG__BaseOpenDialogFlags Or DLGex_PickFolders
        
    Case 8 ' combination of browsing for folder (showing files) but folder-selection only
        cBrowser.Controls_SetCustomMode cm_BrowseFoldersShowFiles
        
    Case 9 ' combination of browsing for folder and/or files, select either
        cBrowser.Controls_SetCustomMode cm_PickFilesOrFolders
        
    Case 10 ' basket mode. select files across various folders. there are 5 basket mode options
        cBrowser.Controls_SetCustomMode cm_BasketModeFilesOnly
        
    Case 11 ' standard "Save As" dialog
        ' use all default settings
        With cBrowser
            .Filter = "Text Files|*.txt|Data Files|*.dat|Log Files|*.log|All Files|*.*"
            .DefaultExt = "txt"
            .Tag = "SaveDialog" ' flag used below
        End With
        
    Case 12 ' old-style browse for folder dialog
        cBrowser.Tag = "FolderBrowser"
        cBrowser.InitDir = App.Path
    
    Case Else
    
    End Select

    lstEvents.Clear
    Select Case True                            ' request events if applicable
        Case optEvents(0): eEvents = EventTypeEnum.evtNoEvents
        Case optEvents(1): eEvents = EventTypeEnum.evtDialogEvents
        Case optEvents(2): eEvents = EventTypeEnum.evtAllEvents
    End Select
    
    If cBrowser.Tag = "FolderBrowser" Then
        
        ' Passing hWnd to function not absolutely required. But if you don't, then
        ' this form is not disabled while the dialog is displayed. In any case,
        ' when multiple forms are displayed, only the hWnd passed can be disabled;
        ' the other displayed forms will not. If that is unacceptable, then you
        ' should disable those other forms, show the dialog, then re-enable them.
        
        bReturn = cBrowser.ShowBrowseForFolder(Me.hWnd, _
            "Place your instructions here" & vbCrLf & "-- can contain a few lines" & _
            vbCrLf & "Very long lines of instructions will not wrap. Be warned", , eEvents)
    
    Else
        If chkMultiSelect.Value = vbChecked Then    ' add multi-select option (ignored if ShowSave called)
            If cBrowser.FlagsDialog = 0 Then
                If cBrowser.Tag = "SaveDialog" Then      ' Save As
                    ' multi-select prevented in SaveAs dialogs by Windows, however, allowing it to be passed
                    ' prevents our options being accepted. The class will automatically remove this
                    ' flag, if set by you, whenever the SaveAs Dialog is displayed
                    cBrowser.FlagsDialog = DLG__BaseSaveDialogFlags Or DLG_AllowMultiSelect
                Else
                    cBrowser.FlagsDialog = DLG__BaseOpenDialogFlags Or DLG_AllowMultiSelect
                End If
            Else
                cBrowser.FlagsDialog = cBrowser.FlagsDialog Or DLG_AllowMultiSelect
            End If
        End If
        
        If cBrowser.Tag = "SaveDialog" Then ' Save dialog
            bReturn = cBrowser.ShowSave(Me.hWnd, eEvents)
        Else
            bReturn = cBrowser.ShowOpen(Me.hWnd, eEvents)
        End If
    
    End If
    
    If bReturn Then ' process dialog selected item(s)
    
        If bReadOnlyUsed Then
            lstEvents.AddItem "Dialog closing with Read-Only option set to " & CBool(cBrowser.FlagsDialog And DLG_ReadOnly), 0
        ElseIf bSplitBtnUsed Then
            lstEvents.AddItem "Split-Button closing with menu item selected: " & cBrowser.Controls_PropertyGet(200, ppSelValue), 0
        End If
        
        lstEvents.AddItem "Dialog closed with item(s) selected:", 0
        
        sItems() = Split(cBrowser.FileName, vbNullChar)
        For lItem = 0 To UBound(sItems)
            lstEvents.AddItem "    Selected: " & sItems(lItem), lItem + 1&
        Next
        
    ElseIf Err.LastDllError = CDERR_CANCELED Then
        lstEvents.AddItem "Dialog Canceled", 0
    Else
        lstEvents.AddItem "Dialog failed with error: " & Err.LastDllError, 0
    End If

End Sub

Private Function pvControlTypeFromID(ControlID As Long) As String
    
    Select Case cBrowser.Controls_TypeOf(ControlID)
        Case ctlType_CancelButton: pvControlTypeFromID = "Cancel Button"
        Case ctlType_CheckBox: pvControlTypeFromID = "CheckBox"
        Case ctlType_ComboBox: pvControlTypeFromID = "ComboBox"
        Case ctlType_CommandButton: pvControlTypeFromID = "Command Button"
        Case ctlType_DefaultButton: pvControlTypeFromID = "Default Button"
        Case ctlType_Menu: pvControlTypeFromID = "Menu"
        Case ctlType_OptionBtnGroup: pvControlTypeFromID = "Option Button Group"
        Case ctlType_Separator: pvControlTypeFromID = "Separator"
        Case ctlType_Label: pvControlTypeFromID = "Label"
        Case ctlType_TextBox: pvControlTypeFromID = "TextBox"
        Case ctlType_Container: pvControlTypeFromID = "Container"
        ' other's are not true controls, i.e., ctlType_SelectionBoxCaption
    End Select
    
End Function

Private Sub cBrowser_DialogOnFileOk(Cancel As Boolean)
    ' can prevent dialog from closing by passing back Cancel as True
End Sub

Private Sub cBrowser_DialogOnFileTypeChange()
    lstEvents.AddItem "Dialog: Changed filter selection: " & cBrowser.FilterIndex, 0
    ' example of setting the default extension relative to the filter item selected (items are 1-bound)
    If cBrowser.Tag = "SaveDialog" Then
        If cBrowser.FilterIndex = 4 Then cBrowser.DefaultExt = "" ' All Files(*.*) filter chosen
    End If
End Sub

Private Sub cBrowser_DialogOnFolderChanged()
    lstEvents.AddItem "Dialog: Navigated to selected folder", 0
End Sub

Private Sub cBrowser_DialogOnFolderChanging(oShellItem As stdole.IUnknown, Cancel As Boolean)
    ' can prevent navigating to folder by returning Cancel as True
    lstEvents.AddItem "Dialog: Navigating to different folder: " & _
        cBrowser.IShellItem_GetDisplayName(oShellItem, dfn_DesktopAbsoluteEditing, True), 0
End Sub

Private Sub cBrowser_DialogOnInit(ByVal DialogHwnd As Long)
    ' DialogHwnd provided should you need to subclass
    lstEvents.AddItem "Dialog initialized. hWnd is " & DialogHwnd & _
                 " (0x" & Right$("0000" & Hex(DialogHwnd), 8) & ")", 0
End Sub

Private Sub cBrowser_DialogOnOverwrite(oShellItem As stdole.IUnknown, lResult As DialogResponseEnum)
    lstEvents.AddItem "Dialog: Option to provide own overwrite prompt for: " & _
        cBrowser.IShellItem_GetDisplayName(oShellItem, dfn_FileSysPath, True, True), 0
    ' not changing the lResult parameter will display a system-supplied message
    ' unless warning is turned off: not including DLG_OverwritePrompt in class .FlagsDialog property
End Sub

Private Sub cBrowser_DialogOnSelectionChanged()
    ' class has a function to return the currently selected item: IDialog_GetCurrentSelection
    Dim oSelected As stdole.IUnknown
    
    Set oSelected = cBrowser.IFileDialog_GetCurrentSelection()
    If oSelected Is Nothing Then
        ' no current item in listview selected
        lstEvents.AddItem "Dialog: Selection changed", 0
    Else
        lstEvents.AddItem "Dialog: Selected folder/file in current view: " & _
            cBrowser.IShellItem_GetDisplayName(oSelected, dfn_FileSysPath, True, True), 0
    End If
End Sub

Private Sub cBrowser_DialogOnShareViolation(oShellItem As stdole.IUnknown, lResult As DialogResponseEnum)
    lstEvents.AddItem "Dialog: Option to respond to sharing violation: " & _
        cBrowser.IShellItem_GetDisplayName(oShellItem, dfn_FileSysPath, True, True), 0
End Sub

Private Sub cBrowser_DialogAddBasketItem(ByVal oShellItemArray As stdole.IUnknown, Reject As Boolean)
    lstEvents.AddItem cBrowser.IShellItemArray_Action(oShellItemArray, siGetArrayItemCount) & _
        " item(s) requested to be added to the basket", 0&
End Sub

Private Sub cBrowser_DialogButtonClicked(ByVal ControlID As Long)
    lstEvents.AddItem pvControlTypeFromID(ControlID) & ": clicked", 0
End Sub

Private Sub cBrowser_DialogCheckBoxChanged(ByVal ControlID As Long, ByVal Value As CheckBoxConstants)
    lstEvents.AddItem "CheckBox: value set to " & Value, 0
End Sub

Private Sub cBrowser_DialogMenuToBeDisplayed(ByVal ControlID As Long)
    lstEvents.AddItem pvControlTypeFromID(ControlID) & ": Menu items being activated", 0
End Sub

Private Sub cBrowser_DialogSubItemClicked(ByVal ControlID As Long, ByVal SubItemIndex As Long)
    lstEvents.AddItem pvControlTypeFromID(ControlID) & ": SubItem selected index is " & SubItemIndex, 0
End Sub

Private Sub cBrowser_DialogXPInit(ByVal DialogHwnd As Long)
    lstEvents.AddItem "XP-Style dialog initialized. hWnd is " & DialogHwnd & _
                 " (0x" & Right$("0000" & Hex(DialogHwnd), 8) & ")", 0
End Sub

Private Sub cBrowser_DialogXPEvent(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, EatMessage As Boolean)
    lstEvents.AddItem "XP Event message: " & Msg, 0
End Sub

Private Sub cBrowser_PreInitMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, EatMessage As Boolean)
    lstEvents.AddItem "XP Pre-Initialization message: " & Msg, 0
End Sub

Private Sub cBrowser_DialogBrowseFoldersEvent(ByVal hWnd As Long, ByVal Msg As Long, ByVal lParam As Long)
    lstEvents.AddItem "BrowseFolders Event message: " & Msg, 0
End Sub

Private Sub cBrowser_DialogBrowseFoldersInit(ByVal DialogHwnd As Long)
    lstEvents.AddItem "BrowseFolders dialog initialized. hWnd is " & DialogHwnd & _
                 " (0x" & Right$("0000" & Hex(DialogHwnd), 8) & ")", 0
End Sub

Private Sub cBrowser_DialogBrowseFoldersOnSelectionChanged(oShellItem As stdole.IUnknown)
    lstEvents.AddItem "BrowseFolders Dialog: Navigating to different folder: " & _
        cBrowser.IShellItem_GetDisplayName(oShellItem, dfn_DesktopAbsoluteEditing, True), 0
End Sub

Private Sub cBrowser_DialogBrowseFoldersValidateFailed(ByVal InvalidName As String, CloseDialog As Boolean)
    lstEvents.AddItem "BrowseFolders Dialog: Invlalid path provided by user: " & InvalidName, 0
End Sub

Private Sub cBrowser_DialogBrowseFoldersPreInitMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal lParam As Long)
    lstEvents.AddItem "BrowseFolders Dialog Pre-Initialization message: " & Msg, 0
End Sub

Private Sub List1_Click()
    Select Case List1.ListIndex
    Case -1: txtInfo.Text = vbNullString
    Case 0
        txtInfo.Text = "NotePad adds a combobox near the 'Open' button. " & _
            "That combobox simply offers users to select the encoding of the text file being selected." & vbCrLf & vbCrLf & _
            "Note: The list index of that combobox can be retrieved before, during & after the dialog displays."
    Case 1
        txtInfo.Text = "Office, depending on what office product owns the dialog, offers various options " & _
            "when selecting a file. Those options are presented as split-button submenu items. " & _
            "Depending on the office product, several of those menu items can be disabled." & _
            "There is also a 'Tools' menu offered that activates drive mapping wizard." & vbCrLf & vbCrLf & _
            "This sample simply mimics those options, the tools menu does nothing. Note: While the dialog is active and if you " & _
            "opted to receive events, you can enable/disable those menu items, as needed."
    Case 2
        txtInfo.Text = "Previous dialog versions always displayed a checkbox to allow the user to specify whether the " & _
            "file should be opened as read-only or not. This option mimics that. Also, the caption for the checkbox " & _
            "is locale-aware regarding unicode. The value of the checkbox can be tested before, during & after dialog displays."
            
    Case 3
        txtInfo.Text = "The current dialog, when shown via GetOpenFileName API vs. the IFileDialog classes, offers a read-only " & _
            "option also. However, that option is not a checkbox, but rather submenu items of the 'Open' button as " & _
            "a split-button. This example mimics that and also presents submenu item captions as locale-aware regarding unicode."
    Case 4
        txtInfo.Text = "The new IFileDialog classes allow you to navigate into zip files " & _
            "when a specific dialog initialization flag is set. However, when that flag is set, you can no longer select " & _
            "the zip file because the 'Open' button treats it as a folder."
    Case 5
        txtInfo.Text = "The new IFileDialog classes allow you to navigate into zip files " & _
            "when a specific dialog initialization flag is set. However, when that flag is set, you can no longer select " & _
            "the zip file because the 'Open' button treats it as a folder." & vbCrLf & vbCrLf & _
            "This is a custom workaround option that enables selecting the zip file itself when that flag is used."
    Case 6
        txtInfo.Text = "You can hide the Filter dropdown box in the Open dialog. This is a nice alternative if your only filter " & _
            "is going to be 'All Files'. Rhetorical: Why show the filter if it has no real purpose?"
    Case 7
        txtInfo.Text = "The new IFileDialog classes can also be used to Browse for Folders just by " & _
            "setting one of the various intialization flags."
    Case 8
        txtInfo.Text = "The Browse for Folder option of the dialog does not offer a way of also showing files." & vbCrLf & vbCrLf & _
            "This class has a custom option to do that and restrict the user to selecting only folders."
    Case 9
        txtInfo.Text = "The Browse for Folder option of the dialog does not offer a way of also showing files." & vbCrLf & vbCrLf & _
            "This class has a custom option to do that and allow user to select either just folders or just files."
    Case 10
        txtInfo.Text = "You can opt to select files from different folders. The dialog will remain open until you click the 'OK' button." & vbCrLf & vbCrLf & _
            "Whether events were requested or not, this option always forwards an event asking if you want to allow the user selection(s)." & vbCrLf & vbCrLf & _
            "In real life, you may want to provide a confirmation for the selection. In any case, you should always perform a confirmation " & _
            "of the selected items before you process them."
    
    Case 12
        txtInfo.Text = "Just a typical Save As dialog. Little to no change since Win2000/XP. You can request events if desired."
    End Select
End Sub

Private Sub mnuXP_Click(Index As Integer)
    If Index = 0& Then
        pvDoSamples 2&
    Else
        pvDoSamples List1.ListCount - 1&
    End If
End Sub
