VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeBrowseFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//////////////////// APIs for displaying Browse for Folder dialog \\\\\\\\\\\\\\\\\\\\\\\
Private Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidList As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As Long) As Long

Private Declare Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHBrowseForFolderW Lib "shell32.dll" (ByVal lpBrowseInfo As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pbc As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long
Private Declare Function SHILCreateFromPath Lib "shell32.dll" (ByVal pszPath As Long, ByRef ppidl As Long, ByRef rgflnOut As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Long) As Long

Private Const MAX_PATH As Long = 260&
Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Enum BrowseForFolderDialogFlagsEnum
    ' full descriptions: http://msdn.microsoft.com/en-us/library/windows/desktop/bb773205%28v=vs.85%29.aspx
    BIF_RETURNONLYFSDIRS = &H1
        'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
    BIF_DONTGOBELOWDOMAIN = &H2
        'Do not include network folders below the domain level in the dialog box's tree view control.
    BIF_STATUSTEXT = &H4
        'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box
    BIF_RETURNFSANCESTORS = &H8
        'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy
    BIF_EDITBOX = &H10
        'Include an edit control in the browse dialog box that allows the user to type the name of an item.
    BIF_VALIDATE = &H20
        'If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc
    BIF_NEWDIALOGSTYLE = &H40
        'Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized.
    BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
        'Use the new user interface, including an edit box
    BIF_BROWSEINCLUDEURLS = &H80
        'The browse dialog box can display URLs
    BIF_UAHINT = &H100
        'When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box.
    BIF_NONEWFOLDERBUTTON = &H200
        'Do not include the New Folder button in the browse dialog box
    BIF_NOTRANSLATETARGETS = &H400
        'When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
    BIF_BROWSEFORCOMPUTER = &H1000
        'Only return computers. If the user selects anything other than a computer, the OK button is grayed
    BIF_BROWSEFORPRINTER = &H2000
        'Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed.
    BIF_BROWSEINCLUDEFILES = &H4000
        'The browse dialog box displays files as well as folders.
    BIF_SHAREABLE = &H8000&
        'The browse dialog box can display shareable resources on remote systems
    BIF_BROWSEFILEJUNCTIONS = &H10000
        'Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
End Enum

' the following can be used in the hook procedure. Here just for conveience/FYI

' http://msdn.microsoft.com/en-us/library/aa452756.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/bb762598%28v=vs.85%29.aspx

'Private Const WM_USER = &H400
'Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
'Private Const BFFM_SETSTATUSTEXTW = WM_USER + 104
'Private Const BFFM_SETSELECTION = (WM_USER + 102)
'Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
'Private Const BFFM_ENABLEOK = WM_USER + 101
'Private Const BFFM_INITIALIZED = 1
'Private Const BFFM_SELCHANGED = 2
'Private Const BFFM_VALIDATEFAILEDA = 3  '// lParam:szPath ret:1(continue),0(EndDialog)
'Private Const BFFM_VALIDATEFAILEDW = 4  '// lParam:wzPath ret:1(continue),0(EndDialog)
'Private Const BFFM_IUNKNOWN = 5         '// provides IUnknown to client. lParam: IUnknown*

Private obif As BrowseInfo
Private m_CancelErr As Boolean
Private m_PIDL As Long
Private m_File As String
Private m_Unicode As Boolean

Public Function ShowBrowseForFolder(ByVal ownerHwnd As Long, Optional ReleasePIDL As Boolean = True) As Boolean
    
    ' depending on what flags you set, you may get a returned path and/or file name or nothing
    ' If the user makes a selection to a virtual path/object, you will not get a path returned.
    ' However, in all cases where the function does not fail, a PIDL will be returned, this PIDL
    '   may be to a virtual path/object like the Control Panel for example. The PIDL by default
    '   will be destroyed after the browser closes unless the optional parameter above is passed as False.
    '   Should you pass that parameter as False, it is your responsibility to destroy the PIDL. You
    '   can destroy it manually or call this class' DestroyPIDL method.
    ' The PIDL if not destroyed herein, can be retrieved via the PIDL_ReturnValue property
    
    Dim lLen As Long, hLib As Long
   
    Call InitStructure(ownerHwnd)
    
    If m_Unicode Then
        m_PIDL = SHBrowseForFolderW(VarPtr(obif))
    Else
        m_PIDL = SHBrowseForFolderA(obif)
    End If

    If m_PIDL = 0& Then 'failure
        If m_CancelErr Then Err.Raise 18 + vbObjectError, "ShowBrowseForFolder", "Canceled by user"
    Else
        m_File = String$(MAX_PATH + 2&, vbNullChar)
        If m_Unicode Then
            lLen = SHGetPathFromIDListW(m_PIDL, StrPtr(m_File))
        Else
            lLen = SHGetPathFromIDListA(m_PIDL, m_File)
        End If
        If lLen Then
            lLen = lstrlenW(StrPtr(m_File))
            m_File = Left$(m_File, lLen)
        Else
            m_File = vbNullString       ' virtual path? PIDL valid, but no path, i.e., Control Panel
        End If
         If ReleasePIDL Then            ' Free memory allocated for PIDL
            DestroyPIDL m_PIDL
            m_PIDL = 0&
        End If
        ShowBrowseForFolder = True
    End If
    
End Function

' forces an error to occur if user cancels/closes dialog without selecting anything
Public Property Get CancelError() As Boolean
    CancelError = m_CancelErr
End Property
Public Property Let CancelError(ByVal cErr As Boolean)
    m_CancelErr = cErr
End Property

Public Sub Clear()
    ' clears all properties and resets to default
    Dim blankOFN As BrowseInfo
    Me.SetRootFolder vbNullString
    obif = blankOFN
    m_CancelErr = False
End Sub

' returns/sets custom value sent to hook procedure
Public Property Get CustomHookData() As Long
    CustomHookData = obif.lParam
End Property
Public Property Let CustomHookData(ByVal nPointer As Long)
    ' If you want to set an initial directory for the browse window, you can set
    '   this value to a pointer of the path. You could also use use a PIDL if desired
    ' Note that you must also set the CustomHookProc property and use a BrowseCallbackProc function
    ' A sample BrowseCallbackProc can be found at very end of this class
    
    ' If on a non-unicode system (pretty much Win98 and below), set the value to this:
    ' 1. Place the initial directory path in a string variable, i.e., strInitDirectory
    ' 2. Set this property like so:
    '   strInitDirectory = StrConv(strInitDirectory & vbNullChar, vbFromUnicode)
    '   myFolderBrowser.CustomHookData = StrPtr(strInitDirectory)
    
    ' if on a unicode system, set value to this:
    ' 1. Place the initial directory path in a string variable, i.e., strInitDirectory
    ' 2. Set this property like so: myFolderBrowser.CustomHookData = StrPtr(strInitDirectory)
    
    obif.lParam = nPointer
End Property

' returns/sets the hook procedure address: pass result of AddressOf()
' ref: msdn.microsoft.com/en-us/library/windows/desktop/bb762598(v=vs.85).aspx
Public Property Get CustomHookProc() As Long
    CustomHookProc = obif.lpfnCallback
End Property
Public Property Let CustomHookProc(ByVal nProc As Long)
    obif.lpfnCallback = nProc
    
    ' how to get the AddressOf() into a Long value?
    ' use a helper function like so:
    
'       ' following would be placed in module along with your BrowseCallbackProc function
'       ' the BrowseCallbackProc function must also be Public for this usage
'        Public Function GetFunctionAddress(inAddr As Long) As Long
'            GetFunctionAddress = inAddr
'        End Function
'
'       ' now a sample call might look like:
'       myFolderBrowser.CustomHookProc = GetFunctionAddress(AddressOf BrowseCallbackProc)

'   Note that you will most likely want to or need to set the CustomHookData property also

End Property

' returns/sets the dialog window caption/title
Public Property Get DialogTitle() As String
    DialogTitle = obif.lpszTitle
End Property
Public Property Let DialogTitle(ByVal dTitle As String)
    obif.lpszTitle = dTitle
End Property

' returns the folder/filename the dialog returns
' This may return blank if user selected a virutual object: Control Panel for example
Public Property Get SelectedItem() As String
    SelectedItem = m_File
End Property

' reutrns the display name of the item selected
' this only returns blank if the ShowBrowseForFolder failed
Public Property Get DisplayName() As String
    Dim lLen As Long
    lLen = lstrlenW(StrPtr(obif.pszDisplayName))
    If lLen Then DisplayName = Left$(obif.pszDisplayName, lLen)
End Property

' returns/sets the flags used for the dialog
Public Property Get flags() As BrowseForFolderDialogFlagsEnum
    flags = obif.ulFlags
End Property
Public Property Let flags(ByVal dFlags As BrowseForFolderDialogFlagsEnum)
    obif.ulFlags = dFlags
End Property

' Returns the PIDL selected by the user after a ShowBrowseForFolder
'   call returns and the optional DestroyPIDL parameter to that function was False
' PIDL is only valid until next ShowBrowseForFolder call
' YOU ARE RESPONSIBLE FOR DESTROYING THE PIDL if it is non-zero
Public Property Get PIDL_ReturnValue() As Long
    PIDL_ReturnValue = m_PIDL
End Property

' Note that this property seems to always return zero when the
' BIF_NEWDIALOGSTYLE is used
Public Property Get SysImageListIconIndex() As Long
    SysImageListIconIndex = obif.iImage
End Property

Public Sub DestroyPIDL(PIDL As Long)
    If PIDL Then CoTaskMemFree PIDL
    If m_PIDL = PIDL Then m_PIDL = 0&
End Sub

Public Function CreatePIDLfromFolder(thePath As String) As Long
    
    ' can return a PIDL that could be useful for a BrowseCallbackProc function
    ' YOU ARE RESPONSIBLE FOR DESTROYING THE PIDL
    Call SHILCreateFromPath(StrPtr(thePath), CreatePIDLfromFolder, ByVal 0&)
    
End Function

' sets the top level folder where browsing will begin from
Public Function SetRootFolder(thePath As String) As Boolean
    If obif.pIDLRoot Then
        Call CoTaskMemFree(obif.pIDLRoot)
        obif.pIDLRoot = 0&
    End If
    If Len(thePath) Then
        Call SHILCreateFromPath(StrPtr(thePath), obif.pIDLRoot, ByVal 0&)
        SetRootFolder = Not (obif.pIDLRoot = 0&)
    Else
        SetRootFolder = True
    End If
    
End Function

Private Sub InitStructure(hWnd As Long)

   With obif
        ' add/remove any flags that are in conflict with the dialog mode (save/open)
        If (.ulFlags And BIF_BROWSEINCLUDEURLS) Then
            ' The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set
            .ulFlags = .ulFlags Or BIF_USENEWUI Or BIF_BROWSEINCLUDEFILES
        End If
        If (.ulFlags And BIF_SHAREABLE) Then
            ' The BIF_NEWDIALOGSTYLE flag must also be set.
            .ulFlags = .ulFlags Or BIF_NEWDIALOGSTYLE
        End If
        
        .pszDisplayName = String$(MAX_PATH + 2&, vbNullChar)
        .hWndOwner = hWnd
        .iImage = 0&
    End With
    m_File = vbNullString
        
End Sub

Private Sub Class_Initialize()
    m_Unicode = Not (IsWindowUnicode(GetDesktopWindow) = 0&)
End Sub

Private Sub Class_Terminate()
    Me.SetRootFolder vbNullString   ' destroy the PIDL we may have created
End Sub



' sample of a callback procedure to set the initial directory
' See CustomHookData for ways of setting the initial directory for the lParam parameter

' api declarations you'll want. Constants are provided at top of this class (rem'd out)

' ////////////////////////////////////////////////////////////////
' if lParam will be set on a unicode system, you can use this
' Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

' ////////////////////////////////////////////////////////////////
' if lParam will be set on a non-unicode system or is a PIDL value, you can use this
' Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

'Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal lParam As Long) As Long
'    Dim sBuffer As String
'    Select Case uMsg
'       Case BFFM_INITIALIZED
'            ' http://msdn.microsoft.com/en-us/library/aa452875.aspx
'            ' the 3rd parameter is either zero or non-zero
'            '   zero :: lParam is a PIDL
'            '   non-zero :: lParam is StrPtr(yourDesiredInitialDirectoryPath)
'
'            ' /// lParam is result of StrPtr() to a full path
'                ' non-unicode system only:
'                Call SendMessage(hWnd, BFFM_SETSELECTION, 1&, ByVal lParam)
'                ' or on unicode systems: Call SendMessageW(hWnd, BFFM_SETSELECTIONW, 1&, ByVal lParam)
'
'            ' /// lParam is a PIDL
'                ' non-unicode, but should work fine on unicode systems:
'                Call SendMessage(hWnd, BFFM_SETSELECTION, 0&, ByVal lParam)
'                ' or on unicode systems: Call SendMessageW(hWnd, BFFM_SETSELECTIONW, 0&, ByVal lParam)
'
'        ' other Case BFFM_xxxx messages as desired
'
'   End Select
'   BrowseCallbackProc = 0
'End Function


