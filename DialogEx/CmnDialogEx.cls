VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmnDialogEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' recent changes
' added old-style Browse For Folder option, capable with callbacks (also added new public events)
' when browsing for folders, via the FileDialog, while showing files (not basket mode), behavior changed:
'   If only file(s) selected and the "Select Folder" button is clicked, the returned
'   item will be the file(s) parent folder
' added option to assign GUIDs in the ShowOpen and ShowSave functions
'   also included methods: ResetClientData, CreateClientDataGUID, ClientDataGUID property
' pvIFileDialog_OnItemSelected was returning menu/combo/listbox selected item index off by 1. fixed
' DefaultExt property can now be updated during dialog events, useful for Save dialogs
' Excluding DLG_OverwritePrompt from Flags for SaveAs dialog failed to be honored if the
'   DLG_AllowMultiSelect was included. Fixed. pvInitStructure now prevents that flag for Save dialogs
' added auto-sizing for multi-select buffer for the pre-Vista dialog
' some properties coded for runtime get/set: FlagsDialog, DefaultExt, DialogTitle, FilterIndex
' remove a couple of properties: FileExtension, FilePath, MaxTitleSize
' renamed Flags/FlagsEx to FlagsDialog/FlagsDialogEx
' added FlagsFolderBrowsing for the folder browser
' added all known CSIDL and Vista+ KnownFolderIDs as courtesy
' added Tag property for the class

'  -----======== PURPOSE: A unicode aware/compatible open/save dialog class ========-----
' Supports XP/Win2K and above. Displays dialog version realtive to user's O/S
' Class Instancing = MultiUse (can be created by users)

' Supports 3 dialogs (minimum O/S is Win2K/XP), with callbacks capability
'   Browse for Folders (old-style dialog). Applicable properties include
'       FlagsFolderBrowsing property
'       DialogTitle property
'       InitDir property
'       FileName property (selected folder's complete path, if any)
'       FileTille property (select folder, if any)
'       Tag property
'       SetRootFolder property
'       CancelError property
'       CloseDialog method when events are requested
'       CustomData property
'       Clear method
'       ShowBrowseForFolder function
'   Open File Dialog
'       All properties/methods apply unless noted in the comments for that function/property
'   Save File Dialog
'       All properties/methods apply unless noted in the comments for that function/property

' Coder notes:
'   1. Always declare IUnknown as stdole.IUnknown in this class and when passing IUnknown to this class.
'       Should a TLB be used in your project, that TLB probably has an IUnknown defined. By declaring
'       this class' IUnknown as stdole, guaranteed to ensure the expected interface definition used
'   2. Any class method that returns stdole.IUnknown or has such parameters are compatible with
'       a TLB IUnknown definition also.
'   3. Any stdole.IUnknown passed to one of the class' events or returned from these methods can be
'       coherced to a TLB specific interface, if applicable. For example, IShellItem interfaces are passed
'       to some events as stdole.IUnknown. If a TLB is in use, the user can call IShellItem methods
'       from the passed interface by setting that interface to the user's TLB IShellItem interface.
'       Example and assumption is user has TLB defining IShellItem:
'           Ths class passes stdole.IUnknown in the DialogOnOverwrite event as: oShellItem
'           Dim myItem As IShellItem  ' << TLB IShellItem
'           Set myItem = oShellItem   ' cohercing stdole.IUnknown to TLB IShellItem
'           ' call TLB IShellItem methods from myItem
'   4. This class is TLB-free. It is written to not conflict with any TLBs in use within a project
'       that includes this class. Because it is TLB-free, it uses low-level functions to talk to
'       non-VB interfaces. Calls to interfaces are wrapped in the class pvCallInterface method

' Changing or Adding dialog controls and their properties. Vista and higher only
'   Related methods regarding controls:
'       Controls_Add
'       Controls_AddReadOnlyOption
'       Controls_PropertyGet, Controls_PropertySet
'       Controls_TypeOf
'       Controls_SetCustomMode (see that routine for more info)
'   Controls can be added solo or "grouped" in a container
'   Controls are displayed on the dialog in the order they were added
'   Containerized controls are displayed in the container in the order they were added
'   Exception for placement: if any control is made "prominent", it is displayed next to
'       Open/Save button. This includes a container, only if the container contains just
'       one control, and that control is one of a select few. See Controls_PropertySet
'   To create a control or container for other controls: Call Controls_Add
'   To change default state (enabled/visible, prominent): Call Controls_PropertySet
'   To get their state: Call Controls_PropertyGet
'   To query a control's type based on its user-defined ID, call Controls_TypeOf

' There are several controls that can be added via those properties
'   Comboboxes, Option Buttons, Labels, Checkboxes, Command buttons, menus and Separator
'   A special case checkbox/split-button Read-Only option that makes captions locale-aware
' The two dialog buttons can be modified and also the dialog textbox caption
'   Converting the default Open/Save button to a split button with menu items
'   Changing the caption of the dialog Cancel button (requires Win7+)
'   Changing the caption alongside the dialog's selection textbox

' Win7+ added options/methods
'   SetNavFolders. Adds folders to the left Navigation pane (Places)
'   SetRootFolder. Replaces drives & some other items from the Nagivation pane with this folder

' XP/Win2K restrictions
'   Due to conflicting flags between the dialog versions, 6 pre-Vista flags not supported.
'   For details see FileDialogFlagsEnum below

' Following events can be raised during the lifecycle of the dialog
' To receive these events, declare this class using WithEvents keyword
Public Event DialogOnInit(ByVal DialogHwnd As Long)
'   dialog initialized, about to be displayed; hWnd provided
Public Event DialogOnFileOk(Cancel As Boolean)
'   dialog's default button (Open/Save) being clicked; can abort
Public Event DialogOnFolderChanging(oShellItem As stdole.IUnknown, Cancel As Boolean)
'   user is moving to another path/folder; can abort
Public Event DialogOnShareViolation(oShellItem As stdole.IUnknown, Response As DialogResponseEnum)
'   enables response to sharing violations that arise from Open or Save operations
Public Event DialogOnOverwrite(oShellItem As stdole.IUnknown, Response As DialogResponseEnum)
'   user chooses to overwrite an existing file (Save Dialog)
Public Event DialogOnFileTypeChange()   ' user changed the filter selection
'   good time to change the default extension property if this is a (Save Dialog)
Public Event DialogOnFolderChanged()    ' user moved to a different path/folder
Public Event DialogOnSelectionChanged() ' user changing selection in current folder

Public Event DialogAddBasketItem(ByVal oShellItemArray As stdole.IUnknown, Reject As Boolean)
' custom event when Controls_SetCustomMode is called with a basket mode option
' oShellItemArray will be an IShellItemArray interface with 1+ items the user is adding to the basket
' this event is raised whether events were requested or not

' Following control events can be raised if the dialog is customized at all
Public Event DialogSubItemClicked(ByVal ControlID As Long, ByVal SubItemIndex As Long)
'   added option button clicked or combobox's list item changed
Public Event DialogButtonClicked(ByVal ControlID As Long)
'   added command button clicked
Public Event DialogCheckBoxChanged(ByVal ControlID As Long, ByVal Value As CheckBoxConstants)
'   added checkbox checked/unchecked
Public Event DialogMenuToBeDisplayed(ByVal ControlID As Long)
'   menu's submenu item or split-button menu item being displayed

' Following dialog events are sent only from XP/Win2K style dialog
Public Event DialogXPInit(ByVal DialogHwnd As Long)
'   dialog initialized, about to be displayed; hWnd provided
Public Event DialogXPEvent(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, EatMessage As Boolean)
'   note: EatMessage parameter prevents dialog from receiving the message
Public Event PreInitMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, EatMessage As Boolean)
'   messages (if any) sent before DialogXPInit is triggered
'   note: EatMessage parameter prevents dialog from receiving the message

' Following dialog events are sent only from old-style Browse For Folder dialog
Public Event DialogBrowseFoldersInit(ByVal DialogHwnd As Long)
'   dialog initialized, about to be displayed; hWnd provided
Public Event DialogBrowseFoldersEvent(ByVal hWnd As Long, ByVal Msg As Long, ByVal lParam As Long)
'   note: EatMessage parameter prevents dialog from receiving the message
Public Event DialogBrowseFoldersValidateFailed(ByVal InvalidName As String, CloseDialog As Boolean)
'   note: EatMessage parameter prevents dialog from receiving the message
Public Event DialogBrowseFoldersOnSelectionChanged(oShellItem As stdole.IUnknown)
' user changing selection in current folder
Public Event DialogBrowseFoldersPreInitMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal lParam As Long)
'   messages (if any) sent before DialogXPInit is triggered


'//////////////////// APIs for displaying Browse for Folder dialog \\\\\\\\\\\\\\\\\\\\\\\
Private Declare Function ILCloneFull Lib "Shell32.dll" (ByVal pIDL As Long) As Long
Private Declare Sub ILFree Lib "Shell32.dll" (ByVal pIDL As Long)
Private Declare Function SHBrowseForFolderW Lib "Shell32.dll" (ByVal lpBrowseInfo As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long) As Long
' ref: https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/ns-shlobj_core-_browseinfoa
Private Type BROWSEINFO
   hwndOwner As Long        ' passed in ShowBrowseForFolder function
   pidlRoot As Long         ' SetRootFolder property
   pszDisplayName As String ' FileName property after dialog closes
   lpszTitle As String      ' DialogTitle property
   ulFlags As Long          ' FlagsFolderBrowsing property
   lpfnCallback As Long     ' internally set
   lpCallbackData As Long   ' CustomData property
   iImage As Long           ' identifies which icon associated with the selected folder: SysImageListIconIndex property
End Type                    ' returned in ShowBrowseForFolder function

' used to get locale-aware caption for the read-only checkbox option & other captions
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExW" (ByVal lpLibFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadString Lib "user32.dll" Alias "LoadStringW" (ByVal hInstance As Long, ByVal wID As Long, lpBuffer As Any, ByVal nBufferMax As Long) As Long

' used for creating events callback thunks
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cb As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32.dll" (ByVal lpfn As Long) As Long

' Pre-Vista dialogs
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
' ref: https://docs.microsoft.com/en-us/windows/desktop/api/commdlg/ns-commdlg-tagofna
Private Type OPENFILENAME
   lStructSize          As Long     'Length of structure, in bytes
   hwndOwner            As Long     'Window that owns the dialog, or NULL
   hInstance            As Long     'Handle of mem object containing template (not used)
   lpstrFilter          As String   'File types/descriptions, delimited with vbnullchar, ends with 2xvbnullchar
   lpstrCustomFilter    As String   'Filters typed in by user
   nMaxCustFilter       As Long     'Length of CustomFilter, min 40x chars
   nFilterIndex         As Long     'Filter Index to use (1,2,etc) or 0 for custom
   lpstrFile            As String   'Initial file/returned file(s), delimited with vbnullchar for multi files
   nMaxFile             As Long     'Size of Initial File long  , min 256
   lpstrFileTitle       As String   'File.ext excluding path
   nMaxFileTitle        As Long     'Length of FileTitle
   lpstrInitialDir      As String   'Initial file dir, null for current dir
   lpstrTitle           As String   'Title bar of dialog
   Flags                As Long     'See DLG_Flags
   nFileOffset          As Integer  'Offset to file name in full path, 0-based
   nFileExtension       As Integer  'Offset to file ext in full path, 0-based (excl '.')
   lpstrDefExt          As String   'Default ext appended, excl '.', max 3 chars
   lCustData            As Long     'Appl defined data for lpfnHook
   lpfnHook             As Long     'Pointer to hook procedure
   lpTemplateName       As String   'Template Name
   pvReserved           As Long     'new Win2000 / WinXP members
   dwReserved           As Long     'new Win2000 / WinXP members
   FlagsEx              As Long     'new Win2000 / WinXP members
End Type
'///////// IFileOpenDialog & IFileSaveDialog related APIs (Vista and later)
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare Function CoCreateInstance Lib "ole32.dll" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Any) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal pclsid As Long) As Long
Private Declare Function SHCreateItemFromParsingName Lib "Shell32.dll" (ByVal pszPath As Long, ByVal pBC As Long, ByVal riid As Long, ByVal pUnk As Long) As Long
Private Declare Function SHILCreateFromPath Lib "Shell32.dll" (ByVal pszPath As Long, ByRef pPIDL As Long, ByVal rgflnOut As Long) As Long
Private Declare Function SHCreateItemFromIDList Lib "Shell32.dll" (ByVal pIDL As Long, ByVal riid As Long, ByVal pUnk As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Sub ReleaseStgMedium Lib "ole32.dll" (ByVal lpstgmedium As Long)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHCreateShellItemArrayFromIDLists Lib "Shell32.dll" (ByVal cidl As Long, ByVal rgpidl As Long, ppsiItemArray As stdole.IUnknown) As Long
Private Declare Function SHGetIDListFromObject Lib "Shell32.dll" (ByVal pItem As Long, ByRef pIDL As Long) As Long
Private Declare Function SHSetTemporaryPropertyForItem Lib "Shell32.dll" (ByVal psi As Long, ByVal propkey As Long, ByVal propvar As Long) As Long
Private Declare Function ILIsEqual Lib "Shell32.dll" (ByVal pidlA As Long, ByVal pidlB As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExW" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Long, ByVal lpsz2 As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32.dll" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function SHGetFolderPath Lib "Shell32.dll" Alias "SHGetFolderPathW" (ByVal hWnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long
Private Declare Function SHGetKnownFolderPath Lib "Shell32.dll" (ByVal rfid As Long, ByVal dwFlags As Long, ByVal hToken As Long, ByRef ppszPath As Long) As Long

Private Enum pvInterfaceMethodOffsets ' in vtable ordinal-order & notes for me
    ifUnknown_QueryInterface = 0                    ' pass 2 params
    ifUnknown_Release = 2                           ' no params
    ifOleWindow_GetWindow = 3                       ' pass 1 param
    ifModalWindow_Show = 3                          ' pass 1 param
    ifFileDialog_SetFileTypes = 4                   ' pass 2 params
    ifFileDialog_SetFileTypeIndex = 5               ' pass 1 param
    ifFileDialog_GetFileTypeIndex = 6               ' pass 1 param
    ifFileDialog_Advise = 7                         ' pass 2 params
    ifFileDialog_Unadvise = 8                       ' pass 1 param
    ifFileDialog_SetOptions = 9                     ' pass 1 param
    ifFileDialog_GetOptions = 10                    ' pass 1 param
    ifFileDialog_SetDefaultFolder = 11              ' pass 1 param
    ifFileDialog_SetFolder = 12                     ' pass 1 param
    ifFileDialog_GetFolder = 13                     ' pass 1 param
    ifFileDialog_GetCurrentSelection = 14           ' pass 1 param
    ifFileDialog_SetFileName = 15                   ' pass 1 param
    ifFileDialog_SetTitle = 17                      ' pass 1 param
    ifFileDialog_SetOkButtonLabel = 18              ' pass 1 param
    ifFileDialog_SetFileNameLabel = 19              ' pass 1 param
    ifFileDialog_GetResult = 20                     ' pass 1 param
    ifFileDialog_AddPlace = 21                      ' pass 2 params
    ifFileDialog_SetDefaultExtension = 22           ' pass 1 param
    ifFileDialog_Close = 23                         ' pass 1 param
    ifFileDialog_SetClientGUID = 24                 ' pass 1 param
    ifFileDialog_ClearClientData = 25               ' no params
    ifFileDialog_SetFilter = 26                     ' pass 1 param
    ifFileDialog2_SetCancelButtonLabel = 27         ' pass 1 param
    ifFileDialog2_SetNavigationRoot = 28            ' pass 1 param
    ifFileDialogOpen_GetResults = 27                ' pass 1 param
    ifFileDialogOpen_GetSelectedItems = 28          ' pass 1 param
    ifFileDlgCustomize_EnableOpenDropDown = 3       ' pass 1 param
    ifFileDlgCustomize_AddMenu = 4                  ' pass 2 params
    ifFileDlgCustomize_AddPushButton = 5            ' pass 2 params
    ifFileDlgCustomize_AddComboBox = 6              ' pass 1 param
    ifFileDlgCustomize_AddRadioButtonList = 7       ' pass 1 param
    ifFileDlgCustomize_AddCheckButton = 8           ' pass 3 params
    ifFileDlgCustomize_AddEditBox = 9               ' pass 2 params
    ifFileDlgCustomize_AddSeparator = 10            ' pass 1 param
    ifFileDlgCustomize_AddText = 11                 ' pass 1 param
    ifFileDlgCustomize_SetControlLabel = 12         ' pass 2 params
    ifFileDlgCustomize_GetControlState = 13         ' pass 2 params
    ifFileDlgCustomize_SetControlState = 14         ' pass 2 params
    ifFileDlgCustomize_GetEditBoxText = 15          ' pass 2 params
    ifFileDlgCustomize_SetEditBoxText = 16          ' pass 2 params
    ifFileDlgCustomize_GetCheckButtonState = 17     ' pass 2 params
    ifFileDlgCustomize_SetCheckButtonState = 18     ' pass 2 params
    ifFileDlgCustomize_AddControlItem = 19          ' pass 3 params
    ifFileDlgCustomize_GetControlItemState = 22     ' pass 3 params
    ifFileDlgCustomize_SetControlItemState = 23     ' pass 3 params
    ifFileDlgCustomize_GetSelectedControlItem = 24  ' pass 2 params
    ifFileDlgCustomize_SetSelectedControlItem = 25  ' pass 2 params
    ifFileDlgCustomize_StartVisualGroup = 26        ' pass 2 params
    ifFileDlgCustomize_EndVisualGroup = 27          ' no params
    ifFileDlgCustomize_MakeProminent = 28           ' pass 1 param
    ifShellItem_BindToHandler = 3                   ' pass 4 params
    ifShellItem_GetParent = 4                       ' pass 1 param
    ifShellItem_GetDisplayName = 5                  ' pass 2 params
    ifShellItem_GetAttributes = 6                   ' pass 2 params
    ifShellItem2_GetString = 17                     ' pass 2 params
    ifShellItemArray_GetCount = 7                   ' pass 1 param
    ifShellItemArray_GetItemAt = 8                  ' pass 2 params
End Enum

Private Enum BrowseForFolderCallBackEnum    ' browse for folder callback messages
    BFFM_INITIALIZED = 1        ' dialog initialized, hWnd is available
    BFFM_SELCHANGED = 2         ' folder selection changed
    BFFM_VALIDATEFAILEDW = 4    ' user entered invalid path in textbox
    BFFM_IUNKNOWN = 5           ' message is passed an IUknown interface
End Enum

Public Enum DisplayFormatEnum   ' used in IShellItem_GetDisplayName
    dfn_NormalDisplay = &H0&
    dfn_ParentRelativeParsing = &H80018001    ' API-parsable for creating IShell/PIDL objects
    dfn_DesktopAbsoluteParsing = &H80028000   ' API-parsable for creating IShell/PIDL objects
    dfn_ParentRelativeEditing = &H80031001
    dfn_DesktopAbsoluteEditing = &H8004C000
    dfn_FileSysPath = &H80058000
    dfn_URL = &H80068000
    dfn_ParentRelativeForAddressBar = &H8007C001
    dfn_ParentRelative = &H80080001
    dfn_ParentRelativeForUI = &H80094001
End Enum

Public Enum ShellItemAttrsEnum ' (used in IFileDialog_GetCurrentSelection)
    ' nice table of how to use some attributes to determine type of file/folder selected
    ' https://blogs.msdn.microsoft.com/oldnewthing/20171101-00/?p=97325
    siAttr_CANCOPY = &H1
    siAttr_CANMOVE = &H2
    siAttr_CANLINK = &H4
    siAttr_STORAGE = &H8
    siAttr_CANRENAME = &H10
    siAttr_CANDELETE = &H20
    siAttr_HASPROPSHEET = &H40
    siAttr_DROPTARGET = &H100
    siAttr_CAPABILITYMASK = &H177
    siAttr_SYSTEM = &H1000
    siAttr_ENCRYPTED = &H2000
    siAttr_ISSLOW = &H4000
    siAttr_GHOSTED = &H8000&
    siAttr_LINK = &H10000
    siAttr_SHARE = &H20000
    siAttr_READONLY = &H40000
    siAttr_HIDDEN = &H80000
    siAttr_DISPLAYATTRMASK_ZIP = &HFC000
    siAttr_NONENUMERATED = &H100000
    siAttr_NEWCONTENT = &H200000
    siAttr_STREAM = &H400000
    siAttr_STORAGEANCESTOR = &H800000
    siAttr_VALIDATE = &H1000000
    siAttr_REMOVABLE = &H2000000
    siAttr_COMPRESSED = &H4000000
    siAttr_BROWSABLE = &H8000000
    siAttr_FILESYSANCESTOR = &H10000000
    siAttr_FOLDER = &H20000000
    siAttr_FILESYSTEM = &H40000000
    siAttr_STORAGECAPMASK = &H70C50008
    siAttr_HASSUBFOLDER = &H80000000
    siAttr_CONTENTSMASK = &H80000000
    siAttr_PKEYSFGAOMASK = &H81044000
End Enum

Public Enum DialogResponseEnum      ' used as parameters for some events
    rsp_Default = 0                 ' allow dialog to make decision
    rsp_Accept = 1                  ' accept the user's decision
    resp_Refuse = 2                 ' refuse the user's decision
End Enum

' ref: https://docs.microsoft.com/en-us/windows/desktop/api/commdlg/nf-commdlg-commdlgextendederror
Public Enum CommonDialogErrorsEnum ' (used in pvDisplayError)
    CDERR_DIALOGFAILURE = &HFFFF&
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_INITIALIZATION = &H2
    CDERR_LOADSTRFAILURE = &H5
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHINSTANCE = &H4
    CDERR_NOHOOK = &H8
    CDERR_NOTEMPLATE = &H3
    CDERR_REGISTERMSGFAIL = &HC
    CDERR_STRUCTSIZE = &H1
    FNERR_BUFFERTOOSMALL = &H3003&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_SUBCLASSFAILURE = &H3001&
    CDERR_CANCELED = &H800704C7
End Enum

Public Enum EventTypeEnum ' (used in ShowOpen & ShowSave)
    evtNoEvents = 0
    evtDialogEvents = 1
    evtAllEvents = 2        ' dialog events and control events. Control events require dialog events
End Enum

Public Enum DialogOSVersionEnum ' (used in Version property, enum members used throughout)
    dvUnsupported = 0   ' O/S less than Win2K and XP
    dvXP_Win2K = 5      ' Win2K or XP :: uses GetOpenFileName/GetSaveFileName APIs
    dvVista = 6         ' Vista       :: else uses IFileDialog/IFileOpenDialog/IFileSaveDialog
    dvWin7 = 7          ' Win7
    dvWin8 = 8          ' Win8
    dvWin8pt1 = 9       ' Win8.1
    dvWin10 = 10        ' Win10
    ' any values > 10 will be future O/S versions
End Enum

Public Enum PropertyTypeEnum ' (used in Controls_PropertyGet & Controls_PropertySet)
    ppProminent = 0     ' prominent state -- only 1 control is prominent (1st come, 1st served)
    ppEnabled = 1       ' enabled state
    ppVisible = 2       ' visible state
    ppText = 3          ' text/caption
    ppSelValue = 4      ' list index, menu index, checkbox value
    ppTag = 5           ' user-defined value
    ppChildControlID = 6 ' used to query container child controls relative to their zero-bound position in container
End Enum

Public Enum CommonDialogControlsEnum ' (used in Controls_[xxx] methods, enum members used throughout)
    ctlType_Container = 0       ' container/grouping of controls; has its own Caption
    ctlType_DefaultButton = 1   ' "Open/Save" button; can contain menu items as a split-button
    ctlType_ComboBox = 2        ' standard combobox
    ctlType_OptionBtnGroup = 3  ' group of individual option buttons
    ctlType_CheckBox = 4        ' standard checkbox
    ctlType_Label = 5           ' standard label
    ctlType_Separator = 6       ' separator (line-like control)
    ctlType_TextBox = 7         ' standard single-line textbox
    ctlType_Menu = 8            ' menu-like control; if used should respond to its click event
    ctlType_CommandButton = 9   ' standard command button; if used should respond to its click event
    ctlType_CancelButton = 10   ' Win7+ option to customize caption
    ctlType_SelectionBoxLabel = 11 ' label next to the dialog's selection textbox
End Enum

Public Enum IShellViewModeEnum  ' used in IShellViewModeEnum function
  svm_Auto = -1
  svm_Icon = 1
  svm_SmallIcon = 2
  svm_List = 3
  svm_Details = 4
  svm_Thumbnail = 5
  svm_Tile = 6
  svm_Thumbstrip = 7
  svm_Content = 8               ' Win7+ option
End Enum

Public Enum IShellItemActionEnum    ' used in IShellItem_Action function
    siGetAttributes = 0
    siGetParentIShellItem = 1
    siGetIStream = 2
    siGetIStorage = 3
    siGetIDataObject = 4
End Enum

Public Enum IShellItemArrayActionEnum ' used in IShellItemArrayActionEnum function
    siGetArrayItemCount = 0
    siGetIShellItemAt = 1
End Enum

Public Enum CSIDLenum           ' can be passed to InitDir, SetRootFolder, SetNavFolders, SetNavFoldersEx
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1        'Internet Explorer (icon on desktop)
    CSIDL_PROGRAMS = &H2        'Start Menu\Programs
    CSIDL_CONTROLS = &H3        'My Computer\Control Panel
    CSIDL_PRINTERS = &H4        'My Computer\Printers
    CSIDL_PERSONAL = &H5        'My Documents
    CSIDL_FAVORITES = &H6       '<user name>\Favourites
    CSIDL_STARTUP = &H7         'Start Menu\Programs\Startup
    CSIDL_RECENT = &H8          '<user name>\Recent
    CSIDL_SENDTO = &H9          '<user name>\SendTo
    CSIDL_BITBUCKET = &HA       '<desktop>\Recycle Bin
    CSIDL_STARTMENU = &HB       '<user name>\Start Menu
    CSIDL_MYDOCUMENTS = &HC     'logical "My Documents" desktop icon
    CSIDL_MYMUSIC = &HD         '"My Music" folder
    CSIDL_MYVIDEO = &HE         '"My Videos" folder
    CSIDL_DESKTOPDIRECTORY = &H10 '<user name>\Desktop
    CSIDL_DRIVES = &H11         'My Computer
    CSIDL_NETWORK = &H12        'Network Neighborhood (My Network Places)
    CSIDL_NETHOOD = &H13        '<user name>\nethood
    CSIDL_FONTS = &H14          'windows\fonts
    CSIDL_TEMPLATES = &H15      'templates
    CSIDL_COMMON_STARTMENU = &H16 'All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17 'All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18 'All Users\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
    CSIDL_APPDATA = &H1A        '<user name>\Application Data
    CSIDL_PRINTHOOD = &H1B      '<user name>\PrintHood
    CSIDL_LOCAL_APPDATA = &H1C  '<user name>\Local Settings\Application Data (non roaming)
    CSIDL_ALTSTARTUP = &H1D     'non localized startup
    CSIDL_COMMON_ALTSTARTUP = &H1E 'non localized common startup
    CSIDL_COMMON_FAVORITES = &H1F 'non localized common startup
    CSIDL_INTERNET_CACHE = &H20 'non localized common startup
    CSIDL_COOKIES = &H21        'non localized common startup
    CSIDL_HISTORY = &H22        'non localized common startup
    CSIDL_COMMON_APPDATA = &H23 'All Users\Application Data
    CSIDL_WINDOWS = &H24        'GetWindowsDirectory()
    CSIDL_SYSTEM = &H25         'GetSystemDirectory()
    CSIDL_PROGRAM_FILES = &H26  'C:\Program Files
    CSIDL_MYPICTURES = &H27     'C:\Program Files\My Pictures
    CSIDL_PROFILE = &H28        'USERPROFILE
    CSIDL_SYSTEMX86 = &H29      'x86 system directory on RISC
    CSIDL_PROGRAM_FILESX86 = &H2A 'x86 C:\Program Files on RISC
    CSIDL_PROGRAM_FILES_COMMON = &H2B 'C:\Program Files\Common
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
    CSIDL_COMMON_TEMPLATES = &H2D 'All Users\Templates
    CSIDL_COMMON_DOCUMENTS = &H2E 'All Users\Documents
    CSIDL_COMMON_ADMINTOOLS = &H2F 'All Users\Start Menu\Programs\Administrative Tools
    CSIDL_ADMINTOOLS = &H30     '<user name>\Start Menu\Programs\Administrative Tools
    CSIDL_CONNECTIONS = &H31    'Network and Dial-up Connections
    CSIDL_COMMON_MUSIC = &H35   'All Users\My Music
    CSIDL_COMMON_PICTURES = &H36 'All Users\My Pictures
    CSIDL_COMMON_VIDEO = &H37   'All Users\My Video
    CSIDL_RESOURCES = &H38      'Resource Directory
    CSIDL_RESOURCES_LOCALIZED = &H39 'Localized Resource Directory
    CSIDL_COMMON_OEM_LINKS = &H3A 'Links to All Users OEM specific apps
    CSIDL_CDBURN_AREA = &H3B    'USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
    CSIDL_COMPUTERSNEARME = &H3D 'Computers Near Me (computered from Workgroup membership)
End Enum
Private Const CSIDL_FLAG_NO_ALIAS As Long = &H1000

' parsed from: https://docs.microsoft.com/en-us/windows/desktop/shell/knownfolderid
Public Enum KnownFolderIDenum   ' current as of Aug 5th, 2018
    ' Vista+: can be passed to InitDir, SetRootFolder, SetNavFolders, SetNavFoldersEx
    ' Values below are arbritrary. Get actual GUID via class function: GetKnownFolderGUID(OneOfTheseEnumValues)
    eFOLDERID_AccountPictures = 0&          ' no equivalent csidl
    eFOLDERID_AddNewPrograms                ' no equivalent csidl
    eFOLDERID_AdminTools                    ' pre-Vista equivalent: CSIDL_ADMINTOOLS
    eFOLDERID_AppDataDesktop                ' no equivalent csidl
    eFOLDERID_AppDataDocuments              ' no equivalent csidl
    eFOLDERID_AppDataFavorites              ' no equivalent csidl
    eFOLDERID_AppDataProgramData            ' no equivalent csidl
    eFOLDERID_ApplicationShortcuts          ' no equivalent csidl
    eFOLDERID_AppsFolder                    ' no equivalent csidl
    eFOLDERID_AppUpdates                    ' no equivalent csidl
    eFOLDERID_CameraRoll                    ' no equivalent csidl
    eFOLDERID_CDBurning                     ' pre-Vista equivalent: CSIDL_CDBURN_AREA
    eFOLDERID_ChangeRemovePrograms          ' no equivalent csidl
    eFOLDERID_CommonAdminTools              ' pre-Vista equivalent: CSIDL_COMMON_ADMINTOOLS
    eFOLDERID_CommonOEMLinks                ' pre-Vista equivalent: CSIDL_COMMON_OEM_LINKS
    eFOLDERID_CommonPrograms                ' pre-Vista equivalent: CSIDL_COMMON_PROGRAMS
    eFOLDERID_CommonStartMenu               ' pre-Vista equivalent: CSIDL_COMMON_STARTMENU
    eFOLDERID_CommonStartup                 ' pre-Vista equivalent: CSIDL_COMMON_STARTUP, CSIDL_COMMON_ALTSTARTUP
    eFOLDERID_CommonTemplates               ' pre-Vista equivalent: CSIDL_COMMON_TEMPLATES
    eFOLDERID_ComputerFolder                ' pre-Vista equivalent: CSIDL_DRIVES
    eFOLDERID_ConflictFolder                ' no equivalent csidl
    eFOLDERID_ConnectionsFolder             ' pre-Vista equivalent: CSIDL_CONNECTIONS
    eFOLDERID_Contacts                      ' no equivalent csidl
    eFOLDERID_ControlPanelFolder            ' pre-Vista equivalent: CSIDL_CONTROLS
    eFOLDERID_Cookies                       ' pre-Vista equivalent: CSIDL_COOKIES
    eFOLDERID_Desktop                       ' pre-Vista equivalent: CSIDL_DESKTOP, CSIDL_DESKTOPDIRECTORY
    eFOLDERID_DeviceMetadataStore           ' no equivalent csidl
    eFOLDERID_Documents                     ' pre-Vista equivalent: CSIDL_MYDOCUMENTS, CSIDL_PERSONAL
    eFOLDERID_DocumentsLibrary              ' no equivalent csidl
    eFOLDERID_Downloads                     ' no equivalent csidl
    eFOLDERID_Favorites                     ' pre-Vista equivalent: CSIDL_FAVORITES, CSIDL_COMMON_FAVORITES
    eFOLDERID_Fonts                         ' pre-Vista equivalent: CSIDL_FONTS
    eFOLDERID_Games                         ' no equivalent csidl
    eFOLDERID_GameTasks                     ' no equivalent csidl
    eFOLDERID_History                       ' pre-Vista equivalent: CSIDL_HISTORY
    eFOLDERID_HomeGroup                     ' no equivalent csidl
    eFOLDERID_HomeGroupCurrentUser          ' no equivalent csidl
    eFOLDERID_ImplicitAppShortcuts          ' no equivalent csidl
    eFOLDERID_InternetCache                 ' pre-Vista equivalent: CSIDL_INTERNET_CACHE
    eFOLDERID_InternetFolder                ' pre-Vista equivalent: CSIDL_INTERNET
    eFOLDERID_Libraries                     ' no equivalent csidl
    eFOLDERID_Links                         ' no equivalent csidl
    eFOLDERID_LocalAppData                  ' pre-Vista equivalent: CSIDL_LOCAL_APPDATA
    eFOLDERID_LocalAppDataLow               ' no equivalent csidl
    eFOLDERID_LocalizedResourcesDir         ' pre-Vista equivalent: CSIDL_RESOURCES_LOCALIZED
    eFOLDERID_Music                         ' pre-Vista equivalent: CSIDL_MYMUSIC
    eFOLDERID_MusicLibrary                  ' no equivalent csidl
    eFOLDERID_NetHood                       ' pre-Vista equivalent: CSIDL_NETHOOD
    eFOLDERID_NetworkFolder                 ' pre-Vista equivalent: CSIDL_NETWORK, CSIDL_COMPUTERSNEARME
    eFOLDERID_Objects3D                     ' no equivalent csidl
    eFOLDERID_OriginalImages                ' no equivalent csidl
    eFOLDERID_PhotoAlbums                   ' no equivalent csidl
    eFOLDERID_PicturesLibrary               ' no equivalent csidl
    eFOLDERID_Pictures                      ' pre-Vista equivalent: CSIDL_MYPICTURES
    eFOLDERID_Playlists                     ' no equivalent csidl
    eFOLDERID_PrintersFolder                ' pre-Vista equivalent: CSIDL_PRINTERS
    eFOLDERID_PrintHood                     ' pre-Vista equivalent: CSIDL_PRINTHOOD
    eFOLDERID_Profile                       ' pre-Vista equivalent: CSIDL_PROFILE
    eFOLDERID_ProgramData                   ' pre-Vista equivalent: CSIDL_COMMON_APPDATA
    eFOLDERID_ProgramFiles                  ' pre-Vista equivalent: CSIDL_PROGRAM_FILES
    eFOLDERID_ProgramFilesX64               ' no equivalent csidl
    eFOLDERID_ProgramFilesX86               ' pre-Vista equivalent: CSIDL_PROGRAM_FILESX86
    eFOLDERID_ProgramFilesCommon            ' pre-Vista equivalent: CSIDL_PROGRAM_FILES_COMMON
    eFOLDERID_ProgramFilesCommonX64         ' no equivalent csidl
    eFOLDERID_ProgramFilesCommonX86         ' pre-Vista equivalent: CSIDL_PROGRAM_FILES_COMMONX86
    eFOLDERID_Programs                      ' pre-Vista equivalent: CSIDL_PROGRAMS
    eFOLDERID_Public                        ' no equivalent csidl
    eFOLDERID_PublicDesktop                 ' pre-Vista equivalent: CSIDL_COMMON_DESKTOPDIRECTORY
    eFOLDERID_PublicDocuments               ' pre-Vista equivalent: CSIDL_COMMON_DOCUMENTS
    eFOLDERID_PublicDownloads               ' no equivalent csidl
    eFOLDERID_PublicGameTasks               ' no equivalent csidl
    eFOLDERID_PublicLibraries               ' no equivalent csidl
    eFOLDERID_PublicMusic                   ' pre-Vista equivalent: CSIDL_COMMON_MUSIC
    eFOLDERID_PublicPictures                ' pre-Vista equivalent: CSIDL_COMMON_PICTURES
    eFOLDERID_PublicRingtones               ' no equivalent csidl
    eFOLDERID_PublicUserTiles               ' no equivalent csidl
    eFOLDERID_PublicVideos                  ' pre-Vista equivalent: CSIDL_COMMON_VIDEO
    eFOLDERID_QuickLaunch                   ' no equivalent csidl
    eFOLDERID_Recent                        ' pre-Vista equivalent: CSIDL_RECENT
    eFOLDERID_RecordedTV                    ' no equivalent csidl. vista-only, n/a for Win7+
    eFOLDERID_RecordedTVLibrary             ' no equivalent csidl
    eFOLDERID_RecycleBinFolder              ' pre-Vista equivalent: CSIDL_BITBUCKET
    eFOLDERID_ResourceDir                   ' pre-Vista equivalent: CSIDL_RESOURCES
    eFOLDERID_Ringtones                     ' no equivalent csidl
    eFOLDERID_RoamingAppData                ' pre-Vista equivalent: CSIDL_APPDATA
    eFOLDERID_RoamedTileImages              ' no equivalent csidl
    eFOLDERID_RoamingTiles                  ' no equivalent csidl
    eFOLDERID_SampleMusic                   ' no equivalent csidl
    eFOLDERID_SamplePictures                ' no equivalent csidl
    eFOLDERID_SamplePlaylists               ' no equivalent csidl
    eFOLDERID_SampleVideos                  ' no equivalent csidl
    eFOLDERID_SavedGames                    ' no equivalent csidl
    eFOLDERID_SavedPictures                 ' no equivalent csidl
    eFOLDERID_SavedPicturesLibrary          ' no equivalent csidl
    eFOLDERID_SavedSearches                 ' no equivalent csidl
    eFOLDERID_Screenshots                   ' no equivalent csidl
    eFOLDERID_SEARCH_CSC                    ' no equivalent csidl
    eFOLDERID_SearchHistory                 ' no equivalent csidl
    eFOLDERID_SearchHome                    ' no equivalent csidl
    eFOLDERID_SEARCH_MAPI                   ' no equivalent csidl
    eFOLDERID_SearchTemplates               ' no equivalent csidl
    eFOLDERID_SendTo                        ' pre-Vista equivalent: CSIDL_SENDTO
    eFOLDERID_SidebarDefaultParts           ' no equivalent csidl
    eFOLDERID_SidebarParts                  ' no equivalent csidl
    eFOLDERID_SkyDrive                      ' no equivalent csidl
    eFOLDERID_SkyDriveCameraRoll            ' no equivalent csidl
    eFOLDERID_SkyDriveDocuments             ' no equivalent csidl
    eFOLDERID_SkyDrivePictures              ' no equivalent csidl
    eFOLDERID_StartMenu                     ' pre-Vista equivalent: CSIDL_STARTMENU
    eFOLDERID_Startup                       ' pre-Vista equivalent: CSIDL_STARTUP, CSIDL_ALTSTARTUP
    eFOLDERID_SyncManagerFolder             ' no equivalent csidl
    eFOLDERID_SyncResultsFolder             ' no equivalent csidl
    eFOLDERID_SyncSetupFolder               ' no equivalent csidl
    eFOLDERID_System                        ' pre-Vista equivalent: CSIDL_SYSTEM
    eFOLDERID_SystemX86                     ' pre-Vista equivalent: CSIDL_SYSTEMX86
    eFOLDERID_Templates                     ' pre-Vista equivalent: CSIDL_TEMPLATES
    eFOLDERID_TreeProperties                ' no equivalent csidl. vista-only, n/a for Win7+
    eFOLDERID_UserPinned                    ' no equivalent csidl
    eFOLDERID_UserProfiles                  ' no equivalent csidl
    eFOLDERID_UserProgramFiles              ' no equivalent csidl
    eFOLDERID_UserProgramFilesCommon        ' no equivalent csidl
    eFOLDERID_UsersFiles                    ' no equivalent csidl
    eFOLDERID_UsersLibraries                ' no equivalent csidl
    eFOLDERID_Videos                        ' pre-Vista equivalent: CSIDL_MYVIDEO
    eFOLDERID_VideosLibrary                 ' no equivalent csidl
    eFOLDERID_Windows                       ' pre-Vista equivalent: CSIDL_WINDOWS
End Enum

' ref pre-Vista: https://docs.microsoft.com/en-us/windows/desktop/api/commdlg/ns-commdlg-tagofna
' ref Vista+: https://docs.microsoft.com/en-us/windows/desktop/api/shobjidl_core/ne-shobjidl_core-_fileopendialogoptions#constants
Public Enum FileDialogFlagsEnum
    DLG_ReadOnly = &H1                       ' Read/Write on pre-Vista. ReadOnly on Vista+ (applies after dialog closes (see SetReadOnlyCheckBox))
    DLG_OverwritePrompt = &H2                ' common between versions
    DLG_NoChangeDir = &H8                    ' common between versions
    DLG_NoValidate = &H100                   ' common between versions
    DLG_AllowMultiSelect = &H200             ' common between versions
    DLG_PathMustExist = &H800                ' common between versions
    DLG_FileMustExist = &H1000               ' common between versions
    DLG_CreatePrompt = &H2000                ' common between versions
    DLG_ShareAware = &H4000                  ' common between versions
    DLG_NoReadOnlyReturn = &H8000&           ' common between versions
    DLG_NoTestFileCreate = &H10000           ' common between versions
    DLG_Explorer = &H80000                   ' always applied all versions, but unsupported flag in Vista+
    DLG_NoDereferenceLinks = &H100000        ' common between versions
    DLG_EnableSizing = &H800000              ' always aplies all versions, but unsupported flag in Vista+
    DLG_DontAddToRecent = &H2000000          ' common between versions
    DLG_ForceShowHidden = &H10000000         ' common between versions
    DLG__BaseOpenDialogFlags = DLG_PathMustExist Or DLG_FileMustExist Or DLG_NoChangeDir
    DLG__BaseSaveDialogFlags = DLG_OverwritePrompt Or DLG_NoReadOnlyReturn Or DLG_PathMustExist Or DLG_NoChangeDir
    ' following applies to Vista and higher, not pre-Vista --------------------------------------------
    ' note: if these are applied to a pre-Vista system, they will be ignored unless noted below
    DLGex_StrictFileTypes = &H4             ' treated as DLGXPonly_HideReadOnly in pre-Vista
    DLGex_PickFolders = &H20                '
    DLGex_ForceFileSystem = &H40            ' treated as DLGXPonly_EnableTemplate in pre-Vista
    DLGex_AllNonStorageItems = &H80         ' treated as DLGXPonly_EnableTemplateHandle in pre-Vista
    DLGex_HideMRUplaces = &H20000           ' only Vista, ingored on Win7+ and pre-Vista
    DLGex_HidePinnedPlaces = &H40000        ' treated as DLGXPonly_NoNetworkButton in pre-Vista
    DLGex_DefaultNoMinMode = &H20000000     ' only Vista, ignored on Win7+ and pre-Vista
    DLGex_ForcePreviewPaneOn = &H40000000
    DLGex_SupportStreamableItems = &H80000000
    ' following applies only to pre-Vista -------------------------------------------------------------
    ' note: if these are applied to a Vista+ system, they will be ignored unless noted below
    DLGXPonly_HideReadOnly = &H4            ' treated as DLGex_StrictFileTypes in Vista+
    DLGXPonly_ExtensionDifferent = &H400    ' ignored by Vista+. XP usage must test this on dialog closure
    DLGXPonly_NoNetworkButton = &H20000     ' treated as DLGex_HideMRUplaces in Vista, ignored on Win7+
    DLGXPonly_EnableIncludeNotify = &H400000 ' filtering message available when trapping events, ignored in Vista+
    
    ' following pre-Vista flags are permanently ignored. Masked out in FlagsDialog property
    ' OFN_SHOWHELP(&H10)        not supported past XP (requires extra effort to apply)
    ' OFN_ENABLEHOOK(&H20)      conflicts with FOS_PICKFOLDERS; ShowOpen/Save method offers workaround via WantEvents
    ' OFN_NOLONGNAMES(&H40000)  applies to old-style dialog & conflicts with FOS_HIDEPINNEDPLACES
    ' OFN_LONGNAMES(&H200000)   applies to old-style dialog & not supported past XP
    ' OFN_ENABLETEMPLATE(&H40)  applies to resource file entries (compatible only when compiled)
    ' OFN_ENABLETEMPLATEHANDLE(&H80) applies to resource file entries (compatible only when compiled)
    ' The above 6 values (mask: &H2400F0) will not be allowed on pre-Vista systems
    ' should you want to include support for templates, then do this
    '   1. add one or both of the restricted flags above to the enumeration so they can be selected
    '       also remove their values from the mask filter in the FlagsDialog property
    '   2. create a new property for each flag you want to support:
    '   for the OFN_ENABLETEMPLATE flag
    '       add a new property allowing you to set the resource file template name
    '       -- make that a string property and pass numeric resource ids with # prefix, i.e., "#101"
    '   for the OFN_ENABLETEMPLATEHANDLE flag
    '       the structure's ofn.hInstance value contains the template location, add a property to set that
    '   3. tweak pvInitStructure function to sanity-check the new property values & apply it to only pre-Vista dialogs
End Enum
Public Enum FileDialogFlagsExtendedEnum
    DLGXPonly_NoPlacesBar = &H1         ' hides the side bar on pre-Vista only
End Enum

' ref: https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/ns-shlobj_core-_browseinfoa#members
Public Enum FolderDialogFlagsEnum
    BIF_RETURNONLYFSDIRS = &H1  'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
    BIF_DONTGOBELOWDOMAIN = &H2 'Do not include network folders below the domain level in the dialog box's tree view control.
    BIF_STATUSTEXT = &H4        'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box
    BIF_RETURNFSANCESTORS = &H8 'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy
    BIF_EDITBOX = &H10          'Include an edit control in the browse dialog box that allows the user to type the name of an item.
    BIF_VALIDATE = &H20         'If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc
    BIF_NEWDIALOGSTYLE = &H40   'Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized.
    BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX) 'Use the new user interface, including an edit box
    BIF_BROWSEINCLUDEURLS = &H80 'The browse dialog box can display URLs
    BIF_UAHINT = &H100          'When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box.
    BIF_NONEWFOLDERBUTTON = &H200 'Do not include the New Folder button in the browse dialog box
    BIF_NOTRANSLATETARGETS = &H400 'When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
    BIF_BROWSEFORCOMPUTER = &H1000 'Only return computers. If the user selects anything other than a computer, the OK button is grayed
    BIF_BROWSEFORPRINTER = &H2000 'Only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed.
    BIF_BROWSEINCLUDEFILES = &H4000 'The browse dialog box displays files as well as folders.
    BIF_SHAREABLE = &H8000&     'The browse dialog box can display shareable resources on remote systems
    BIF_BROWSEFILEJUNCTIONS = &H10000 'Windows 7 and later. Allow folder junctions such as a library or a compressed file with a .zip file name extension to be browsed.
End Enum

Public Enum CustomModeEnum ' (used in Controls_SetCustomMode)
    cm_BasketModeFilesOnly = &H1000&        ' select files from different folders
    cm_BasketModeFoldersOnly = &H2000&
    cm_BasketModeFoldersOnlyShowFiles = &H3000&
    cm_BasketModeFilesOrFolders = &H4000&
    cm_BrowseFoldersShowFiles = &H8000&     ' select folders, but show files also
    cm_CompressedFolderPlusFiles = &H9000&  ' browse into zips & option to select zip itself
    cm_PickFilesOrFolders = &HA000&         ' select mix of folders and/or files
End Enum

Private Type ThunkVTableStruct
    IID(0 To 3) As Long         ' IFileDialogEvents or IFileDialogControlEvents
    VTable As Long              ' Interface's vTable
End Type

Private Type ControlStruct      ' custom UDT used for managing added/edited controls
    ID As Long                  ' container/control ID (negative); positive in m_ControlKeys
    ' ^^ note: internally, all IDs are negative. IDs with high bit (&H80000000) reserved for Controls_SetCustomMode
    ' ^^ Valid user IDs range from 1 to 268435455
    Caption As String           ' textbox,label,container,command button
    SubItems() As String        ' option button captions, combobox items, submenus (menu & split-button)
    ' ^^ note: for containers, this is list of contained control IDs
    ' ^^ in any case, max number of items supported: 32767
    SubItemState() As Byte      ' individual subitem enabled/visible states; populated on demand, as needed
    Flags As Long               ' ^^ default is all controls/items are enabled and visible
    ' 0x0000000F    type of control. CommonDialogControlsEnum value
    ' 0x000000F0    control's state 1=enabled,2=visible,4=prominent,8=SubItemState() array sized
    ' 0x007FFF00    depends on control type
    '               ComboBox,OptionButtonGroup,DefaultButton (if not read-only option): selected item
    '               CheckBox,DefaultButton (if read-only option): vbChecked or vbUnchecked
    '               containers: size of SubItems() array, number of contained controls
    '               any other control: ignored and/or reserved
    ' 0x01000000    default button is split-button, no read-only option
    ' 0x02000000    default button is split-button read-only option (see Controls_AddReadOnlyOption)
    ' 0x03000000    split-button usage mask
    ' 0x10000000    no initial selected item for comboboxes/option button groups
    ' 0x80000000    part of a container vs. stand-alone control when a control vs container
    ' 0x6C800000    reserved/unused
    Tag As Variant              ' optional user-defined value
End Type

Private Const CLSID_FileOpenDialog  As String = "{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}"
Private Const IID_IFileOpenDialog   As String = "{D57C7288-D4AD-4768-BE02-9D969532D960}"
Private Const IID_IFileDialog       As String = "{42F85136-DB7E-439C-85F1-E4075D135FC8}"
Private Const IID_IFileDialog2      As String = "{61744FC7-85B5-4791-A9B0-272276309B13}"
Private Const IID_IShellItem        As String = "{43826D1E-E718-42EE-BC55-A1E261C37BFE}"
Private Const IID_IShellItemArray   As String = "{B63EA76D-1F85-456F-A19C-48159EFA858B}"
Private Const CLSCTX_INPROC_SERVER As Long = 1&  ' used when creating IFileDialog
Private Const E_NOINTERFACE As Long = &H80004002 ' possible return value of QueryInterface
Private Const E_NOTIMPL As Long = &H80004001     ' possible return value of any COM method
Private Const E_POINTER As Long = &H80004003     ' possible return value of any COM method
Private Const WM_SETTEXT As Long = &HC&          ' used to change folder browser dialog title
Private Const WM_INITDIALOG As Long = &H110      ' XP dialog constant
Private Const WM_COMMAND As Long = &H111         ' subclassing message
Private Const BM_SETSTYLE As Long = &HF4&        ' subclassing message
Private Const WM_NOTIFY As Long = &H4E           ' autosizing multi-select buffer, pre-Vista
Private Const WM_USER As Long = &H400&           ' used in browse for folder events
Private Const BFFM_SETSELECTIONW = WM_USER + 103 ' used in browse for folder events
Private Const CDN_SELCHANGE As Long = -602&      ' really: CDN_FIRST - 1, autosizing multi-select buffer, pre-Vista
Private Const CDM_GETSPEC = WM_USER + 100&       ' really: CDM_FIRST + 0, autosizing multi-select buffer, pre-Vista
Private Const ID_OK As Long = &H1                ' dialog's "Open/Save" button's Windows control ID
Private Const S_FALSE As Long = &H1              ' possible return value in events
Private Const MAX_PATH As Long = 260             ' XP dialog default string lengths
Private Const SIGN_BIT As Long = &H80000000      ' common-use value throughout class

Private Const emskAllEvents = &H30000           ' mask often used
Private Const emskCtrlEvents = &H20000          ' mask often used
Private Const emskCustMode = &HF000&            ' mask often used

Private m_VTable As Long                   ' memory location of VTable & thunks
Private m_Vsize As Long                    ' size of thunks in memory (used for cleanup)
Private m_EventsCookie As Long             ' needed to receive dialog events
Private m_SubclassFnc As Long              ' function pointer for pvDialog_Subclass function at end of class
Private m_Events() As ThunkVTableStruct    ' IFileFileDialogEvents,IFileDialogControlEvents faux-interfaces
Private m_Dialog As stdole.IUnknown        ' Vista+ dialog
Private m_DlgMod As stdole.IUnknown        ' Vista+ dialog customization
Private m_Dlg2 As stdole.IUnknown          ' Vista+ IFileDialog2 reference (on demand, Win7+)
Private m_AddPlaces As Collection          ' Vista+ list of added places to navigation bar
Private m_ControlKeys As Collection        ' XRef between m_Controls() array and Control ID(s)
Private m_Controls() As ControlStruct      ' Vista+ added controls onto the dialog
'   ^^ note: array item 0 is reserved for Controls_SetCustomMode
Private m_NavRestrictor As Variant         ' Vista+ root folder option or folder browser root folder
Private m_ModeResults As stdole.IUnknown   ' IShellItemArray of selected items when in custom mode
Private m_BasketItems As Collection        ' used only with custom basket mode options
Private m_StateFlags As Long               ' running/state flags
'   0x0000007F  Version property value
'   0x00000080  CancelError property value
'   0x00000100  Add Places option set to top else bottom
'   0x0000F000  custom mode options
'   0x00030000  want events mask: 1 = dialog only, 2 = control events, 3 = all events
'   0x01000000  Old-Style BrowseForFolder dialog mode
'   0x02000000  Old-Style BrowseForFolder running/initialized
'   0x04000000  Old-Style BrowseForFolder needs initial folder set
'   0x10000000  Dialog is subclassed (used for some custom modes)
'   0x20000000  XP Dialog running, but in process of closing via CloseDialog
'   0x40000000  XP Dialog running
'   0x80000000  Vista Dialog running
'   0xE0000000  Dialog running mask (any O/S versions)
'   0x0C000000  reservied for old-style browse for folder
'   0x08FC0E00  reserved/unused
Private m_Guid(0 To 3) As Long      ' used so often, we'll reuse this vs Dim/ReDim in the routines
Private ofn As OPENFILENAME         ' the API's UDT
Private obif As BROWSEINFO          ' old-style browse for folder structure
Private m_minBufferSize As Long     ' for XP-style Multi-Select option. See MinimumBufferSizeRequired property
Private m_ClientGUID As String      ' current GUID assigned to dialog (set in ShowOpen/ShowSave; get in ClientDataGUID property)
Private m_Tag As String             ' user-defined Tag value for this class instance

Public Function ShowOpen(ByVal ownerHwnd As Long, _
                         Optional ByVal WantEvents As EventTypeEnum = evtNoEvents, _
                         Optional ByVal pRtnIUnknownVarPtr As Long = 0&, _
                         Optional ByVal ClientDataGUID As String = vbNullString) As Boolean
    
    ' returns TRUE if user selected something and no error
    ' if an error occurs or user cancels then
    '   if CancelError is True, an error is raised and sent to calling routine; you should use On Error statements
    '   otherwise Err.LastDllError returns the reason/error code, i.e., CDERR_CANCELED, CDERR_DIALOGFAILURE, etc
    
    ' Optional WantEvents parameter. Allows you to receive events during the lifetime of the dialog
    '   To receive events, you must declare this class using the WithEvents keyword
    '   Events should be requested in at least these scenarios:
    '   1. You are adding an enabled textbox to the dialog and need the user-entered text, if any.
    '       This is because the textbox content must be queried before the dialog closes, not after.
    '       This class will auto-query any textbox for its content if WantEvents is not evtNoEvents.
    '       If events are requested:
    '           you do not have to respond to the events if you do not need to
    '           you can get the content during/after the dialog closes via Controls_PropertyGet
    '           Additionally, you can query the textbox at any time, during any event.
    '   2. You are adding any control that performs no self-contained action. These types of controls
    '       include command buttons and menu controls. Request evtAllEvents. In other words, there
    '       is no action taken by the dialog when that control is accessed.
    '   3. You need to prevent the user from selecting specific files or folders. Since Win7+ has
    '       no per-file filter ability, you can prevent folders from being navigated to and prevent
    '       the Open/Save button from closing the dialog, for whatever reasons, during events.
    '   4. Running on XP/Win2K and you want to "hook" the dialog. Requesting evtDiaogEvents is
    '       similar to the older method of including the OFN_ENABLEHOOK flag. The main difference is
    '       that hooked messages are sent to this class and forwarded to you via class public events.
    
    ' Optional pRtnIUnknownVarPtr parameter (Vista+). If wanted, then:
    '   Not applicable for pre-Vista dialogs
    '   - Declare a variable of type: IUnknown, i.e., Dim IUnkReturn As stdole.IUnknown
    '   - Pass its pointer as: VarPtr(IUnkReturn)
    '   If user selected anything, your IUnkReturn will be either an IShellItem interface
    '       or, if more than one item selected, an IShellItemArray interface
    '   If user didn't select anything or user canceled dialog, then IUnkReturn will be: Nothing
    '   Your IUnkReturn can be passed to IFileOperation for copying, moving, etc.
    '   Otherwise, that IUnkReturn is of use to you for other reasons, else you wouldn't want it returned
    
    ' Optional ClientDataGUID parameter must be passed as a valid GUID if passed. Vista+
    '   When passed, some global dialog settings will be stored against this GUID. So make it unique
    '   per dialog scenario. Consider this GUID like a configuration file for the last visited
    '   folder (if no startup folder/file selected), the position, and size of the dialog. You can
    '   have as many configurations/GUIDs as you like.
    '   Note: GUIDs are not case-sensitive and are not required to have enclosing curly brackets { }
    '       however, if brackets are included, ensure they are the curly type, not [ ] nor ( ) nor < >
    '   Note: You may want to consider having unique GUIDs per user, per dialog. Since all GUIDs are
    '       associated with your application, UserA runs your app and uses a dialog GUID you provide.
    '       Then when UserB runs the app with same GUID, they get settings that UserA had & now UserA
    '       settings may have just been changed by UserB. Make these user-unique if used.
    '   Tip: See class' CreateClientDataGUID function. Can be used to generate reusable GUIDs for you
    '   Use ResetClientData to clear any persisted data related to GUIDs you have set/created
    
    ' Tip: If dialog wasn't canceled, you can test the FileName property to determine if multiple items
    '   were selected: If InStr(cBrowser.FileName, vbNullChar) > 0 Then ' more than one item selected
    '   These options allow multiple selections
    '   1. Including DLG_AllowMultiSelect in the Flags property
    '   2. Including a basket mode option in the Controls_SetCustomMode function
    
    Dim lReturn As Long, lError As Long
    
    If Me.Version = dvUnsupported Then
        lError = CDERR_DIALOGFAILURE
        pvDisplayError lError, True
        Exit Function
    
    ElseIf pvInitStructure(ownerHwnd, True, WantEvents, ClientDataGUID) Then
        If Me.Version = dvXP_Win2K Then
            lReturn = GetOpenFileNameW(VarPtr(ofn))
            ' note: cannot trust the return value. I have repeatedly seen API return non-zero (success)
            ' when an error occurred, like buffer too small in multi-select cases. So..., always gonna
            ' ask the dialog for a second opinion
            If lReturn = 0& Then ' failure
                If (m_StateFlags And &H20000000) Then   ' canceled via Me.CloseDialog method
                    lError = ofn.pvReserved: ofn.pvReserved = 0&
                Else
                    lError = CommDlgExtendedError()
                End If
                If lError = 0& Then lError = CDERR_CANCELED
            ElseIf (ofn.Flags And DLG_AllowMultiSelect) Then
                lError = CommDlgExtendedError()
            End If
            If lError = 0& Then
                pvSetFileName
                ShowOpen = True
            End If
        ElseIf m_Dialog Is Nothing Then
            lError = CDERR_DIALOGFAILURE
            
        Else
            lReturn = pvCallInterface(ObjPtr(m_Dialog), ifModalWindow_Show, ownerHwnd)
            pvSubclassDialog False
            pvGetResults lReturn, pRtnIUnknownVarPtr, True
            If lReturn Then
                If lReturn = CDERR_CANCELED Then lError = lReturn Else lError = Err.LastDllError
            Else
                ShowOpen = True
            End If
        End If
    End If
    m_StateFlags = (m_StateFlags And Not &HFF030000)    ' remove flags: dialog running, events
    If lError Then
        If Me.CancelError Then pvDisplayError lError, 1 Else SetLastError lError
    End If
    
End Function

Public Function ShowSave(ByVal ownerHwnd As Long, _
                         Optional ByVal WantEvents As EventTypeEnum = evtNoEvents, _
                         Optional ByVal pRtnIUnknownVarPtr As Long = 0&, _
                         Optional ByVal ClientDataGUID As String = vbNullString) As Boolean

    ' see commnents in ShowOpen method; they are identical to what would be added here
    ' Exception: Multi-select and basket mode does not apply to the Save Dialog
    
    Dim lReturn As Long, lError As Long
   
    If Me.Version = dvUnsupported Then
        lError = CDERR_DIALOGFAILURE
        pvDisplayError lError, False
        Exit Function
        
    ElseIf pvInitStructure(ownerHwnd, False, WantEvents, ClientDataGUID) Then
        If Me.Version = dvXP_Win2K Then
            If GetSaveFileNameW(VarPtr(ofn)) = 0& Then  ' failure
                If (m_StateFlags And &H20000000) Then   ' canceled via Me.CloseDialog method
                    lError = ofn.pvReserved: ofn.pvReserved = 0&
                Else
                    lError = CommDlgExtendedError()
                End If
                If lError = 0& Then lError = CDERR_CANCELED
            Else
                pvSetFileName
                ShowSave = True
            End If
        ElseIf m_Dialog Is Nothing Then
            lError = CDERR_DIALOGFAILURE
            
        Else
            lReturn = pvCallInterface(ObjPtr(m_Dialog), ifModalWindow_Show, ownerHwnd)
            pvGetResults lReturn, pRtnIUnknownVarPtr, False
            If lReturn Then
                If lReturn = CDERR_CANCELED Then lError = lReturn Else lError = Err.LastDllError
            Else
                ShowSave = True
            End If
        End If
    End If
    m_StateFlags = (m_StateFlags And Not &HFF030000)    ' remove flags: dialog running, events
    If lError Then
        If Me.CancelError Then pvDisplayError lError, 0 Else SetLastError lError
    End If
    
End Function

Public Function ShowBrowseForFolder(ByVal ParentHwnd As Long, _
                                    Optional ByVal BrowserHint As String, _
                                    Optional ByVal ReturnFriendlyPath As Boolean = False, _
                                    Optional ByVal WantEvents As EventTypeEnum = evtNoEvents, _
                                    Optional ByRef RtnFolderIconIndex As Long, _
                                    Optional ByVal pRtnIUnknownVarPtr As Long = 0&) As Boolean
    
    ' returns TRUE if user selected something and no error
    ' if an error occurs or user cancels then
    '   if CancelError is True, an error is raised and sent to calling routine; you should use On Error statements
    '   otherwise Err.LastDllError returns the reason/error code, i.e., CDERR_CANCELED, CDERR_DIALOGFAILURE, etc
    
    ' ParentHwnd is required if you want dialog to be modal to the calling form, otherwise form can
    '   be accessed while dialog is open. Dialog does not disable other active forms in your project
    
    ' Optional BrowserHint parameter. This is text that is displayed between the title bar and folder list
    
    ' Optional ReturnFriendlyPath parameter. If user selects control panel folders or virtual folders,
    '   the default format returned is GUID, i.e., ::{26EE0668-A00A-44D7-9371-BEB064C98683}\3
    '   the friendly path of that GUID would be: Control Panel\Network and Internet
    
    ' Optional WantEvents parameter. Allows you to receive events during the lifetime of the dialog
    '   To receive events, you must declare this class using the WithEvents keyword
    
    ' Optional RtnFolderIconIndex parameter. The dialog may provide the icon index associated with the
    '   selected item. If you want this icon index, then pass a the parameter. The index is relative
    '   to a system image list.
    
    ' Optional pRtnIUnknownVarPtr parameter (Vista+). If wanted, then:
    '   - Declare a variable of type: IUnknown, i.e., Dim IUnkReturn As stdole.IUnknown
    '   - Pass its pointer as: VarPtr(IUnkReturn)
    '   If user selected anything, your IUnkReturn will be an IShellItem interface
    '   If user didn't select anything or user canceled dialog, then IUnkReturn will be: Nothing
    '   Your IUnkReturn can be passed to IFileOperation for copying, moving, etc.
    '   Otherwise, that IUnkReturn is of use to you for other reasons, else you wouldn't want it returned
    
    Dim lValue As Long, lError As Long, oItem As stdole.IUnknown
    
    If Me.Version = dvUnsupported Then
        lError = CDERR_DIALOGFAILURE
        pvDisplayError lError, True
        Exit Function
    End If
    
    If m_VTable Then        ' may need to recreate thunk
        If obif.lpfnCallback = 0 Then pzDestroyThunks
    End If
    m_StateFlags = m_StateFlags Or &H1000000    ' flag for folder-browser mode

    ' add/remove any flags that are in conflict with the dialog
    lValue = Me.FlagsFolderBrowsing
    If (lValue And BIF_SHAREABLE) > 0& Or lValue = 0& Then lValue = lValue Or BIF_NEWDIALOGSTYLE
    If (lValue And BIF_BROWSEINCLUDEURLS) Then lValue = lValue Or BIF_USENEWUI Or BIF_BROWSEINCLUDEFILES
    
    If WantEvents Then m_StateFlags = m_StateFlags Or (WantEvents And &H3) * &H10000
    With obif
        .ulFlags = lValue
        .lpfnCallback = 0&
        .lpCallbackData = 0&
        .hwndOwner = ParentHwnd
        .iImage = 0&
        .lpszTitle = BrowserHint
        .pszDisplayName = String$(MAX_PATH, vbNullChar)
        If Len(ofn.lpstrInitialDir) Then m_StateFlags = m_StateFlags Or &H4000000
        If .pidlRoot Then ILFree .pidlRoot: .pidlRoot = 0&
        pvFreeOrCreateRootPIDL True, False          ' create new PIDL root as needed
        If VarType(m_NavRestrictor) = vbLong Then   ' if created, transfer to structure
            .pidlRoot = m_NavRestrictor
            m_NavRestrictor = Empty                 ' remove pidl from m_NavRestrictor
        End If
     End With
     
    If ((m_StateFlags And &H4030000) Or Len(ofn.lpstrTitle)) > 0 Then
        ' create callback if events are wanted or init dir to be set or custom dialog title to be set
        ' init dir and custom dialog title need to be set within events due to dialog being modal
        pzCreateThunks evtDialogEvents
    End If
     
    lValue = SHBrowseForFolderW(VarPtr(obif))   ' returns a PIDL if successful
    If lValue Then
        Set oItem = pvGetIShellItem(lValue)     ' convert PIDL to shell item & get path from that
        If pRtnIUnknownVarPtr Then pvGetIShellItem lValue, pRtnIUnknownVarPtr
        CoTaskMemFree lValue
        If ReturnFriendlyPath Then lValue = dfn_DesktopAbsoluteEditing Else lValue = dfn_DesktopAbsoluteParsing
        ofn.lpstrFile = pvGetDisplayName(ObjPtr(oItem), lValue)
        Set oItem = Nothing
        ofn.nFileOffset = Len(ofn.lpstrFile)    ' ensure non-zero if path returned
        RtnFolderIconIndex = obif.iImage        ' return this to caller
        ShowBrowseForFolder = True
    Else
        ofn.lpstrFile = vbNullString
        RtnFolderIconIndex = 0&: ofn.nFileOffset = 0&
        lError = CDERR_CANCELED
    End If
    m_StateFlags = (m_StateFlags And Not &HFF030000)    ' remove flags: dialog running, events
    obif.pszDisplayName = vbNullString
    If lError Then
        If Me.CancelError = True Then pvDisplayError lError, 2 Else SetLastError lError
    End If
    
End Function

Public Function CloseDialog(Optional ByVal XPDialogHwnd As Long, Optional ByVal ExitCode As Long = CDERR_CANCELED) As Boolean
    
    ' Can only be called from within the passed events for the dialog; else ignored
    ' For XP: pass XPDialogHwnd as the hWnd passed to your event else parameter ignored on Vista+
    
    Const IDAbort As Long = 2&
    Const BFFM_ENABLEOK As Long = (WM_USER + 101)
    If m_StateFlags < 0& Then   ' else Vista+ dialog not active
        pvSubclassDialog False
        pvGetTextBoxValue
        CloseDialog = (pvCallInterface(ObjPtr(m_Dialog), ifFileDialog_Close, ExitCode) = 0&)
    ElseIf (m_StateFlags And &H2000000) Then
        If obif.lpCallbackData Then _
            SendMessage obif.lpCallbackData, BFFM_ENABLEOK, 0&, ByVal 0&
    ElseIf (m_StateFlags And &HE0000000) = &H40000000 Then  ' else XP dialog not active or this already called
        If XPDialogHwnd Then
            ofn.pvReserved = ExitCode
            m_StateFlags = m_StateFlags Or &H20000000   ' flag to stop forwarding messages to user
            CloseDialog = Not (PostMessage(GetParent(XPDialogHwnd), WM_COMMAND, IDAbort, ExitCode) = 0&)
        End If
    End If
End Function

Public Sub ResetClientData(ClientDataGUID As String)

    ' function allows you to reset ClientData without calling ShowOpen or ShowSave
    ' for more information on client data, see ShowOpen & CreateClientDataGUID comments
    
    ' If ClientDataGUID is empty/blank then the persisted data is cleared when no ClientDataGUID is used
    ' otherwise, this parameter simply resets the persisted data associated with the ClientDataGUID.

    If Me.Version > dvXP_Win2K Then
        Dim aGuid(0 To 7) As Long, tDialog As stdole.IUnknown
        
        CLSIDFromString StrPtr(CLSID_FileOpenDialog), VarPtr(aGuid(0))
        IIDFromString StrPtr(IID_IFileOpenDialog), VarPtr(aGuid(4))
        If CoCreateInstance(VarPtr(aGuid(0)), 0&, CLSCTX_INPROC_SERVER, VarPtr(aGuid(4)), tDialog) = 0& Then
            If ClientDataGUID = vbNullString Then
                pvCallInterface ObjPtr(tDialog), ifFileDialog_ClearClientData
            ElseIf IIDFromString(StrPtr(ClientDataGUID), VarPtr(aGuid(0))) = 0& Then
                pvCallInterface ObjPtr(tDialog), ifFileDialog_SetClientGUID, VarPtr(aGuid(0))
                pvCallInterface ObjPtr(tDialog), ifFileDialog_ClearClientData
            End If
            Set tDialog = Nothing
        End If
    End If
End Sub

' returns/sets the string Tag property for this class instance
Public Property Get Tag() As String
    Tag = m_Tag
End Property
Public Property Let Tag(Value As String)
    m_Tag = Value
End Property

Public Property Get Version() As DialogOSVersionEnum
    ' returns O/S in integer format. May depend on your application's Manifest
    Version = (m_StateFlags And &H7F)
End Property

Public Sub Clear()
    If (m_StateFlags And &HF2000000) Then Exit Sub   ' not allowing reset if dialog is running
    ' clears all properties and resets to default
    m_StateFlags = Me.Version       ' only keep O/S-related flags
    m_minBufferSize = 0&
    Set m_DlgMod = Nothing          ' reset these interfaces
    Set m_Dialog = Nothing
    Set m_Dlg2 = Nothing
    Set m_ModeResults = Nothing
    Erase m_Controls()              ' erase user-defined controls & properties
    Set m_AddPlaces = Nothing       ' reset these collections
    Set m_ControlKeys = Nothing
    Set m_BasketItems = Nothing
    m_Tag = vbNullString
    m_ClientGUID = vbNullString
    pzDestroyThunks                 ' release all thunks
    pvFreeOrCreateRootPIDL False, True
    With ofn                        ' clear structures
        .Flags = 0&: .FlagsEx = 0&
        .hInstance = 0&: .hwndOwner = 0&
        .lpstrDefExt = vbNullString: .lpstrFile = vbNullString
        .lpstrFileTitle = vbNullString: .lpstrFilter = vbNullString
        .lpstrInitialDir = vbNullString: .lpstrTitle = vbNullString
        .nMaxFile = 0&: .nFileOffset = 0: .nFilterIndex = 0&: .lCustData = 0&
    End With
    With obif                       ' free PIDL if needed
        If .pidlRoot Then ILFree .pidlRoot: .pidlRoot = 0&
        .hwndOwner = 0&: .iImage = 0&
        .lpfnCallback = 0&: .lpszTitle = vbNullString
        .pszDisplayName = vbNullString: .ulFlags = 0&
    End With
End Sub

' if true then forces an error to occur if user cancels/closes dialog without selecting anything
Public Property Get CancelError() As Boolean
    CancelError = CBool(m_StateFlags And &H80&)
End Property
Public Property Let CancelError(ByVal Value As Boolean)
    If Not Me.CancelError = Value Then m_StateFlags = m_StateFlags Xor &H80&
End Property

Public Property Get ClientDataGUID() As String  ' Read-Only, Vista+ only
    ' returns the client data GUID assigned in the ShowOpen/ShowSave method
    ClientDataGUID = m_ClientGUID
End Property

' returns/sets the filename the dialog starts with or returns
Public Property Get FileName() As String
    ' This property differs fromt typical file dialogs. Reason: to support multiple file
    '   selections across multiple paths, as is possible with basket mode option &
    '   libraries (Vista+).
    ' XP/Win2K does not allow selection from multiple paths; however for consistency,
    '   multiple selected files are returned the same as Vista+
    
    ' Multiple Files Option         Returned Format
    ' -------------------------------------------------------------------------------------
    '           No                  Path\FileName of selected file, if any
    ' Yes, only 1 item selected     Path\FileName of selected file
    ' Yes, 2+ items selected        Path\FileName1 vbNullChar Path\FileName2 vbNullChar etc
    ' note: if vbNullChar exists in returned string, multiple files selected
    ' -------------------------------------------------------------------------------------
    
    FileName = ofn.lpstrFile    ' note: pre-Vista return processed in pvSetFileName

End Property
Public Property Let FileName(ByVal Value As String)

    ' This property can accept the following formats
    ' 1. Filename only, i.e., Scores.txt
    ' 2. Path/FileName, i.e., C:\My Documents\Scores.txt
    '    note: if this is passed, the InitDir folder is overridden with the file's path
    '   Passing C:\My Documents\Scores.txt to this property has same effect as separately:
    '       Passing C:\My Documents to InitDir and passing Scores.txt to this property

    Dim lLen As Long    ' buffer also used to return selection; min size: 260 chars
    ofn.lpstrFile = Value: ofn.nFileOffset = 0
    If Me.Version > dvXP_Win2K Then
        If Not m_Dialog Is Nothing Then _
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileName, StrPtr(ofn.lpstrFile)
    ElseIf Len(ofn.lpstrFile) Then
        If InStr(ofn.lpstrFile, "/") Then ofn.lpstrFile = Replace$(ofn.lpstrFile, "/", "\")
    End If
End Property

' reutrns the file title dialog returns
' note: if multiple items selected, title returned is for the first item only
Public Property Get FileTitle() As String

    Dim n As Long, sValue As String
    
    If ofn.nFileOffset Then
        sValue = ofn.lpstrFile
        If Len(sValue) Then
            n = InStr(sValue, vbNullChar)
            If n > 0& Then sValue = Left$(sValue, n - 1&)
            n = InStrRev(sValue, "\")
            If n = 0& Then
                FileTitle = sValue
            Else
                FileTitle = Mid$(sValue, n + 1&)
            End If
        End If
    End If
End Property

' returns/sets the initial folder the dialog should open too. Available to all dialogs
Public Property Get InitDir() As String
    InitDir = ofn.lpstrInitialDir
End Property
Public Property Let InitDir(ByVal Value As String)

    ' This property can accept the following formats
    ' 1. Full path (with/without trailing slash), i.e., C:\Windows
    ' 2. A PIDL, you are responsible for destroying the PIDL any time after this method returns
    ' 3. A numeric value which is assumed to be a CSIDL. see CSIDLenum for a list
    '    example for CSIDL_COMMON_PROGRAMS (value=23), simply pass: CSIDLenum.CSIDL_COMMON_PROGRAMS or just 23
    ' 4. Known folder GUID (Vista+), representing a system/virtual folder, enclosed curly brackets optional
    '    example using class function, pass: GetKnownFolderGUID(eFOLDERID_AppDataProgramData)

    If Len(Value) = 0& Then Exit Property
    If (m_StateFlags And &HF2000000) Then Exit Property ' dialog running
    
    Dim n As Long, oItem As stdole.IUnknown
    
    ofn.lpstrInitialDir = vbNullString
    If InStr(Value, ":") Then
        ofn.lpstrInitialDir = Value
        If InStr(ofn.lpstrInitialDir, "/") Then ofn.lpstrInitialDir = Replace$(ofn.lpstrInitialDir, "/", "\")
    ElseIf IsNumeric(Value) Then        ' assume is CSIDL or PIDL
        On Error Resume Next
        n = CLng(Value)
        If Err Then
            Err.Clear
        ElseIf (n And &HFF) = n Then ' CSIDL
            ofn.lpstrInitialDir = String$(MAX_PATH, vbNullChar)
            SHGetFolderPath 0&, n Or CSIDL_FLAG_NO_ALIAS, 0&, 0&, StrPtr(ofn.lpstrInitialDir)
            n = InStr(ofn.lpstrInitialDir, vbNullChar)
            If n > 1 Then
                ofn.lpstrInitialDir = Left$(ofn.lpstrInitialDir, n - 1)
            Else
                ofn.lpstrInitialDir = vbNullString
            End If
        Else                        ' PIDL
            Set oItem = pvGetIShellItem(n)
            If Not oItem Is Nothing Then
                ofn.lpstrInitialDir = pvGetDisplayName(ObjPtr(oItem), dfn_DesktopAbsoluteParsing)
                Set oItem = Nothing
            End If
        End If
        On Error GoTo 0
                
    ElseIf Me.Version > dvXP_Win2K Then ' assume GUID
        If IIDFromString(StrPtr(Value), VarPtr(m_Guid(0))) = 0& Then
                If SHGetKnownFolderPath(VarPtr(m_Guid(0)), 0&, 0&, n) = 0& Then _
                    ofn.lpstrInitialDir = pvComStrPtrToVBstring(n, True)
        End If
    End If
        
End Property

' returns/sets the dialog window caption/title
Public Property Get DialogTitle() As String
    DialogTitle = ofn.lpstrTitle
End Property
Public Property Let DialogTitle(ByVal Value As String)
    ofn.lpstrTitle = Value
    If m_StateFlags < 0& Then                   ' Vista+ dialog running
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetTitle, StrPtr(ofn.lpstrTitle)
    ElseIf (m_StateFlags And &H3000000) Then    ' Folder Browser dialog running
        If obif.lpCallbackData Then             ' we have an hWnd else no action
            SendMessage obif.lpCallbackData, WM_SETTEXT, 0&, ByVal StrPtr(ofn.lpstrTitle)
        End If
    End If
End Property

' return/sets the filter(s)
Public Property Get Filter() As String
    If Len(ofn.lpstrFilter) Then
        If ofn.lpstrFilter = vbNullChar Then
            Filter = ofn.lpstrFilter
        Else
            Dim n As Long
            n = InStrRev(ofn.lpstrFilter, vbNullChar & vbNullChar)
            Filter = Replace$(Left$(ofn.lpstrFilter, n - 1&), vbNullChar, "|")
        End If
    End If
End Property
Public Property Let Filter(ByVal Value As String)
    ' Pass filter same as you've always done: FilterName1|FilterPattern1|FilterName2|FilterPattern2 etc
    ' Optional property value
    '   vbNullString results in default "All Files", locale-aware, being used
    '   Vista+ only: vbNullChar results in hiding the filter dropdown (Open dialog only, not Save Dialog)
    If Len(Trim$(Value)) = 0& Then
        ofn.lpstrFilter = vbNullString
    ElseIf Value = vbNullChar Then
        ofn.lpstrFilter = Value
    Else
        ofn.lpstrFilter = Replace$(Value, "|", vbNullChar) & vbNullChar
    End If
End Property

' returns/sets the flags used for the old-style browse for folder dialog
Public Property Get FlagsFolderBrowsing() As FolderDialogFlagsEnum
    FlagsFolderBrowsing = obif.ulFlags
End Property
Public Property Let FlagsFolderBrowsing(ByVal Value As FolderDialogFlagsEnum)
    obif.ulFlags = Value
End Property

' returns/sets the flags used for the dialog
Public Property Get FlagsDialog() As FileDialogFlagsEnum
    If m_StateFlags < 0& Then
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetOptions, VarPtr(FlagsDialog)
    Else
        FlagsDialog = ofn.Flags
    End If
End Property
Public Property Let FlagsDialog(ByVal Value As FileDialogFlagsEnum)
    If Me.Version < dvVista Then
        ' prevent these flags on pre-Vista:
        ' OFN_SHOWHELP(&H10),OFN_ENABLEHOOK(&H20),OFN_NOLONGNAMES(&H40000),OFN_LONGNAMES(&H200000)
        ' OFN_ENABLETEMPLATE(&H40),OFN_ENABLETEMPLATEHANDLE(&H80)
        Value = Value And Not &H2400F0
    Else ' prevent these flags on Vista+, see FlagsDialogEx for overriding this behavior
        ' OFN_SHOWHELP(&H10),OFN_LONGNAMES(&H200000),OFN_ExtensionDifferent(&H400),OFN_EnableIncludeNotify(&H400000)
        Value = Value And Not &H600410
    End If
    ofn.Flags = Value
    If m_StateFlags < 0& Then _
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetOptions, ofn.Flags
End Property

' returns/sets extended Win2K+ flags for the dialog; or overriding Vista+ unsupported flags
Public Property Get FlagsDialogEx() As FileDialogFlagsExtendedEnum
    FlagsDialogEx = ofn.FlagsEx
End Property
Public Property Let FlagsDialogEx(ByVal Value As FileDialogFlagsExtendedEnum)
    
    If Me.Version > dvXP_Win2K Then Value = Value And Not DLGXPonly_NoPlacesBar
    ofn.FlagsEx = Value
    ' for pre-Vista, should only contain FileDialogFlagsExtendedEnum values
    
    ' For Vista+, this is a backdoor to force an excluded Flag to be used.
    ' whether the dialog will honor the flag or not, we have no control over that.
    ' For excluded/unsupported flags, see FlagsDialog property.
    ' Only reason you should use this is if a newer/updated O/S later adds a
    '   flag that was previously unsupported and no interface methods are required
    '   to configure the dialog for that flag.
    If m_StateFlags < 0& Then _
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetOptions, ofn.Flags Or ofn.FlagsEx

End Property

' returns/sets custom data for the dialog/folder browser. This can be set/returned at any point during callbacks
Public Property Get CustomData() As Long
    CustomData = ofn.lCustData
End Property
Public Property Let CustomData(Value As Long)
    ofn.lCustData = Value
End Property

' returns/sets maximum buffer size for returned files; applies only to pre-Vista dialogs
Public Property Get MaxFileSize() As Long
    If ofn.nMaxFile = 0& Then
        MaxFileSize = MAX_PATH + MAX_PATH
    Else
        MaxFileSize = ofn.nMaxFile
    End If
End Property
Public Property Let MaxFileSize(ByVal Value As Long)
        
    ' setting size to zero, resets to default
    If (m_StateFlags And &H6000000) Then Exit Property ' dialog running
    If Value < 1& Then
        ofn.nMaxFile = 0&
    Else
        ofn.nMaxFile = Value
    End If
End Property

' returns/sets the Filter index dialog starts with or returns
Public Property Get FilterIndex() As Long
    If m_StateFlags < 0& Then
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetFileTypeIndex, VarPtr(FilterIndex)
    Else
        FilterIndex = ofn.nFilterIndex
    End If
End Property
Public Property Let FilterIndex(ByVal Value As Long)
    If Value < 1& Then Value = 1&
    If m_StateFlags < 0& Then
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileTypeIndex, Value
    Else
        ofn.nFilterIndex = Value
    End If
End Property

' returns/sets the default extension applied to files in ShowSave
Public Property Get DefaultExt() As String
    DefaultExt = ofn.lpstrDefExt
End Property
Public Property Let DefaultExt(ByVal Value As String)
    If Len(Value) = 0 Then Value = vbNullString
    ofn.lpstrDefExt = Value
    If m_StateFlags < 0& Then _
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetDefaultExtension, StrPtr(ofn.lpstrDefExt)
End Property

' Contains minimum buffer size required if error FNERR_BUFFERTOOSMALL is returned for multi-selected files
' A value of -1 indicates unknown and should never happen.
' Is always reset to zero before dialog is displayed; applies only to pre-Vista dialogs
Public Property Get MinimumBufferSizeRequired() As Long
    MinimumBufferSizeRequired = m_minBufferSize
End Property

Public Function CreateClientDataGUID(DialogID As String, _
                                Optional ByVal AppID As String, _
                                Optional ByVal ClientID As String = "All Clients") As String
                                
    ' This does not guarantee a truly system unique ID

    ' Generates a GUID for your use to pass to ShowOpen and/or ShowSave methods.
    ' You can use your own GUIDs. However, GUIDs should be system unique, not just
    ' application unique. If using your own, consider the GUIDGEN app that VB provides.
    ' If using your own GUIDs, pass your GUIDs directly to ShowOpen/ShowSave
    
    ' This is just a courtesy function & allows you to return the same GUID for the passed
    '   IDs each time it is called; no risk of different GUID per same IDs. Also prevents
    '   saving this info in some user's INI file or DB record since it can be called on demand.
    ' Passed IDs are case-sensitive
    
    ' DialogID is a nickname for your dialog. If you are using GUIDs, then you
    '   should uniquely identify dialogs with a nickname. Always pass this; it's required
    
    ' AppID is a string that uniquely identifies your application. This ID simply
    '   helps make the GUID more unique. Can be app.Title, app.EXEName, app.FileDescription,
    '   a combination of those, or anything you want. Another option is to actually create
    '   a globally unique GUID (via GUIDGEN or other tool) and pass that as your AppID. If
    '   that's done, ensure you make that global GUID a constant in your application so it
    '   is stored and pass the constant as the AppID parameter.
    '   If not provided, then App.Title will be used
    
    ' ClientID is something that uniquely identifies a user of your app
    '   Typically this could be a UserName or UserID if it applies. You may want to
    '   use the API GetUserName. If you don't want a GUID that includes the
    '   user's information, then simply don't pass the parameter.
    
    If Len(DialogID) = 0& Then Exit Function ' Dialog ID required & not provided
    
    Dim i As Long, n As Long, aPtr As Long
    Dim aData() As Byte, l As Long
    Dim aCRC(0 To 255) As Long
    Const CRCpolynomial As Long = &HEDB88320
    
    For n = 0& To 255&  ' create the CRC32 lookup table
        i = n
        For l = 8& To 1& Step -1&
            If (i And 1&) Then
                i = (((i And &HFFFFFFFE) \ 2&) And &H7FFFFFFF) Xor CRCpolynomial
            Else
                i = ((i And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next
        aCRC(n) = i
    Next
    
    If Len(AppID) = 0& Then AppID = App.Title
    For l = 1& To 3&
        Select Case l
            Case 1: aData() = ClientID
            Case 2: aData() = AppID
            Case 3: aData() = DialogID
        End Select
        n = -1&
        For i = 0& To UBound(aData)
            aPtr = (n And &HFF) Xor aData(i)
            n = (((n And &HFFFFFF00) \ &H100&) And &HFFFFFF) Xor aCRC(aPtr)
        Next
        m_Guid(l) = n
    Next
    m_Guid(0) = &H9F9A6617   ' filler, class-specific CRC value
    CreateClientDataGUID = Space$(38)
    StringFromGUID2 VarPtr(m_Guid(0)), StrPtr(CreateClientDataGUID), 39&

End Function

Public Function Controls_SetCustomMode(ByVal Mode As CustomModeEnum, _
                                Optional ByVal ButtonCaption As String, _
                                Optional ByVal OkCaption As String, _
                                Optional ByVal IncludeSelCount As Boolean = True, _
                                Optional ByVal SelCountCaptionFormat As String) As Boolean

    ' Do not call this if intending to show a Save dialog. Aything here will be ignored
    '   in that case when the Save dialog is displayed

    ' function fails in these scenarios:
    '   - if this method called multiple times or dialog has been created
    '   - Open button has been modified
    '   - passing invalid Mode value
    
    ' FYI: Subclassing is introduced (different than events) in all cases
    '   except for these two Mode values: cm_BasketModeFilesOnly, cm_BasketModeFoldersOnly
    
    ' Mode. these custom options circumvent normal dialog behavior
    '   cm_CompressedFolderPlusFiles - browse into zip files and/or select the zip itself
    '       note: by default, DLGex_AllNonStorageItems browses into zips but won't allow selecting zip itself
    '   cm_BrowseFoldersShowFiles - browse for folders only but show files too; can pick only folders
    '       note: by default, DLGex_PickFolders does not show files
    '   cm_PickFilesOrFolders - browse for folders/files, can pick files and/or folders
    '   cm_BasketModeFilesOnly - can select only files across multiple folders
    '   cm_BasketModeFoldersOnly - can select only folders, no files shown
    '   cm_BasketModeFoldersOnlyShowFiles - can select folders only, but files are shown
    '   cm_BasketModeFilesOrFolders - can select both files and/or folders
    '   Basket mode has these dialog behaviors...
    '   - the dialog stays open until the user closes the dialog via a "Ok" button or dialog canceled
    '       whether ShowOpen requests events or not, adding to the basket triggers DialogAddBasketItem events
    '   - the option adds an "OK" button in the prominent position to finish basket mode
    '       if ShowOpen method requests events, the "OK" button triggers a DialogOnFileOk event
    
    ' ButtonCaption. Depends on the Mode value
    '   cm_CompressedFolderPlusFiles. Default caption is: "&Open"   (locale-aware)
    '   cm_BrowseFoldersShowFiles. Default is "Select Folder"       (locale-aware)
    '   cm_PickFilesOrFolders. Default is "Select"                  (locale-aware)
    '   cm_BasketMode[xxx]. Default is "&Add..."                    (locale-aware)
    
    ' OkCaption. applies only if cm_BasketMode[xxxx] is included. Default is "OK" (locale-aware)
    ' IncludeSelCount. applies if cm_BasketMode[xxxx] is included
    '   Adds a label near the "OK" button indicating how many items added to the basket
    
    ' SelCountCaptionFormat. applies if cm_BasketMode[xxxx] is included and IncludeSelCount = True
    '   The caption for the selection count label. Default is "%s items selected" (locale-aware)
    '   If you provide your own caption, you must include %s as a format flag to be replaced by count
    '   Examples:   "Select Count: %s"     "Basket has %s items"

    If Not m_Dialog Is Nothing Then Exit Function           ' dialog created, called this too late
    If (m_StateFlags And emskCustMode) Then Exit Function   ' already called this
    If Me.Version < dvVista Then Exit Function              ' not applicable for pre-Vista
    If pvKeyToIndex("def", False) > -1& Then Exit Function  ' default Open button already modified
    
    Dim hMod As Long, lValue As Long
    Dim sCaption As String, n As Long
    Const ID_Select As Long = 49927         ' String ID in "ExplorerFrame.dll" for "Select"
    Const ID_SelectFolder As Long = 439     ' String ID in comdlg32.dll for "Select Folder" caption
    Const ID_FolderCaption As Long = 438    ' String ID in comdlg32.dll for "Folder:"
    Const ID_OKCaption As Long = 800        ' String ID in user32.dll for "OK" caption
    Const ID_AddCaption As Long = 38313     ' String ID in shell32.dll for "&Add..." caption
    Const ID_ItemsCaption As Long = 38194   ' String ID in shell32.dll for "%s items selected" caption
    
    Select Case Mode
    Case cm_BrowseFoldersShowFiles
        If Len(Trim$(ButtonCaption)) = 0& Then
            ButtonCaption = pvLoadString(hMod, "comdlg32.dll", ID_SelectFolder, 100&)
            If Len(ButtonCaption) = 0& Then ButtonCaption = "Select Folder"
        End If
        IncludeSelCount = False
    
    Case cm_PickFilesOrFolders
        If Len(Trim$(ButtonCaption)) = 0& Then
            ButtonCaption = pvLoadString(hMod, "ExplorerFrame.dll", ID_Select, 100&)
            If Len(ButtonCaption) = 0& Then ButtonCaption = "Select"
        End If
        IncludeSelCount = False
    
    Case cm_CompressedFolderPlusFiles
        ' note: null captions are not updated to buttons; so default will be used in the end
        ' and the default is what we want if no caption provided
        ButtonCaption = Trim$(ButtonCaption)
        IncludeSelCount = False
        
    Case cm_BasketModeFilesOnly, cm_BasketModeFilesOrFolders, cm_BasketModeFoldersOnly, cm_BasketModeFoldersOnlyShowFiles
        ' do this whether captions are provided or not. Thinking vast majority of time, defaults will apply
        If Len(Trim$(OkCaption)) = 0& Then
            OkCaption = pvLoadString(hMod, "user32.dll", ID_OKCaption, 100&)
            If Len(OkCaption) = 0& Then OkCaption = "OK"
            If hMod Then FreeLibrary hMod: hMod = 0&
        End If
        If Len(Trim$(ButtonCaption)) = 0& Then
            ButtonCaption = pvLoadString(hMod, "shell32.dll", ID_AddCaption, 100&)
            If Len(ButtonCaption) = 0& Then ButtonCaption = "&Add..."
        End If
        If IncludeSelCount Then
            If Len(Trim$(SelCountCaptionFormat)) = 0& Or InStr(1, SelCountCaptionFormat, "%s", vbTextCompare) = 0 Then
                SelCountCaptionFormat = pvLoadString(hMod, "shell32.dll", ID_ItemsCaption, 100&)
                If Len(SelCountCaptionFormat) = 0& Then SelCountCaptionFormat = "%s items selected"
            End If
        End If
        Set m_BasketItems = New Collection
    Case Else: Exit Function
    End Select
            
    If hMod Then FreeLibrary hMod: hMod = 0&
    
    pvKeyToIndex SIGN_BIT, True
    If Not m_BasketItems Is Nothing Then
        If IncludeSelCount Then lValue = 2& Else lValue = 1&
        ReDim Preserve m_Controls(0 To UBound(m_Controls) + lValue)
    Else
        lValue = 0&
    End If
    With m_Controls(lValue)                 ' update "Open" button caption
        .Caption = ButtonCaption
        .ID = SIGN_BIT + 2&
        .Flags = &H30 Or ctlType_DefaultButton
        m_ControlKeys.Add lValue, "def"     ' flag to indicate Open button is modified
    End With
    If lValue > 0& Then
        With m_Controls(0)                  ' create "OK" button
            .Caption = OkCaption
            .ID = SIGN_BIT
            .Flags = &H30 Or ctlType_CommandButton
            If IncludeSelCount Then .Flags = .Flags Or SIGN_BIT
        End With
        If IncludeSelCount Then             ' create container with label
            With m_Controls(lValue - 1&)
                .Caption = Replace$(SelCountCaptionFormat, "%S", "%s")
                .ID = SIGN_BIT + 1&
                .Flags = &H130              ' container has 1 item & enabled/visible
                ReDim .SubItems(0 To 0)
                .SubItems(0) = SIGN_BIT
                m_ControlKeys.Add lValue - 1&, CStr(.ID)
            End With
        End If
    End If
    
    ' if selecting only folders. set the SelectionBoxLabel; user can override this
    If Mode = cm_BasketModeFoldersOnly Or Mode = cm_BrowseFoldersShowFiles Then
        ReDim Preserve m_Controls(0 To UBound(m_Controls) + 1&)
        With m_Controls(UBound(m_Controls))
            .Caption = pvLoadString(hMod, "comdlg32.dll", ID_FolderCaption)
            If hMod Then FreeLibrary hMod
            If Len(.Caption) = 0& Then .Caption = "Folder: "
            .Flags = ctlType_SelectionBoxLabel Or &H30
            .ID = SIGN_BIT + 3&
            m_ControlKeys.Add UBound(m_Controls), CStr(.ID)
        End With
    End If
    m_StateFlags = (m_StateFlags And Not emskCustMode) Or Mode
    
End Function

Public Function Controls_AddReadOnlyOption(ByVal ControlID As Long, _
                                           Optional ByVal AsCheckBox As Boolean = True, _
                                           Optional ByVal isChecked As Boolean = False, _
                                           Optional ByVal Caption1 As String, _
                                           Optional ByVal Caption2 As String) As Boolean

    ' creates a read-only checkbox or split-button menu items with locale-aware
    '   caption(s) if optional Caption parameters are not provided.
    ' Additionally, this method ties the user-selection of the read-only option
    '   to the class' Flags and DLG_ReadOnly enumeration value. Whether this option is
    '   a checkbox or split-button menu items, the result of the user's selection can
    '   be queried simply: If (thisClass.Flags And DLG_ReadOnly) Then ' read-only used
    ' By default, checkbox style has the prominence property set. See exceptions below.
    
    ' function fails in these scenarios:
    '   ControlID < 1 or > &HFFFFFFF or dialog already created
    '   AsCheckBox=False and default Open/Save button already modified
    '   calling this multiple times, system < Vista, or dialog already created
    
    ' Parameters
    '   ControlID   AsCheckBox   isChecked   Caption1*  Caption2*
    '   -----------------------------------------------------------------------------
    '       > 0       True          T/F      Optional   n/a; ignored
    '       > 0       False         T/F      Optional   Required if Caption1 provided
    '   -----------------------------------------------------------------------------
    ' (*) Caption1 and Caption2 ignored if an empty string
    '     Caption1 is checkbox caption or split-button's first submenu item
    '     Caption2 is split-button's second submenu item
    '     Default captions (in system's current language) when optional captions not provided:
    '       Checkbox: "Open as &read-only"
    '       Split-Button menu items
    '           if isChecked=True then: "&Open" and "Open for &write"
    '           if isChecked=False then: "&Open" and "Open for &read-only"
    '   -----------------------------------------------------------------------------
    ' Note: This method will create a checkbox in the prominent position, if checkbox is requested.
    '   There are three exceptions regarding prominent positioning of the checkbox
    '   1) If you also call Controls_SetCustomMode with one of the basket mode constants, then the
    '       "OK" button created in that method will always be in the prominent position.
    '   2) If you have called Controls_PropertySet to assign prominence to another control that
    '       was added via Controls_Add before this routine called, then that other control will
    '       take the prominent position, if applicable. Typically, call this routine first.
    '   3) You call Controls_PropertySet later to remove the prominence property for this checkbox
    
    If Me.Version < dvVista Then Exit Function            ' not applicable for pre-Vista
    If Not m_Dialog Is Nothing Then Exit Function         ' dialog created, called this too late
    If ControlID < 1& Or ControlID > &HFFFFFFF Then Exit Function  ' invalid
    If pvKeyToIndex("ro", False) > -1& Then Exit Function ' already set up
    If pvKeyToIndex(ControlID, False) > -1& Then Exit Function ' control ID already used
    
    Dim lIndex As Long, hMod As Long
    Const ID_Open As Long = 370&        ' String ID in comdlg32.dll for "&Open" caption
    Const ID_ReadOnly As Long = 427&    ' String ID in comdlg32.dll for "Open as &read-only" caption
    Const ID_OpenWrite As Long = 428&   ' String ID in comdlg32.dll for "Open for &write" caption

    If AsCheckBox = False Then _
        If pvKeyToIndex("def", False) > -1& Then Exit Function ' default Open button modified already
    
    lIndex = pvKeyToIndex(ControlID, True)
    m_ControlKeys.Add ControlID, "ro"   ' x-ref for retrieving value when dialog closes
    If AsCheckBox = True Then
        With m_Controls(lIndex)
            .ID = -ControlID
            .Flags = &H70& Or ctlType_CheckBox Or Abs(isChecked) * &H100&
            If Len(Trim$(Caption1)) = 0& Then
                .Caption = pvLoadString(hMod, "comdlg32.dll", ID_ReadOnly, 100&)
                If Len(.Caption) = 0& Then .Caption = "Open as &read-only"
            Else
                .Caption = Caption1
            End If
        End With
    Else
        m_ControlKeys.Add lIndex, "def"                 ' xref
        With m_Controls(lIndex)
            .ID = -ControlID
            .Flags = &H2000030 Or ctlType_DefaultButton Or Abs(isChecked) * &H100& Or Abs(isChecked) * &H1000000
            ReDim .SubItems(0 To 1)
            If Len(Trim$(Caption1)) = 0& Then
                .SubItems(0) = pvLoadString(hMod, "comdlg32.dll", ID_Open, 100&)
                If Len(.SubItems(0)) = 0& Then .SubItems(0) = "Open"
            Else
                .SubItems(0) = Caption1
            End If
            If Len(Trim$(Caption2)) = 0& Then
                If isChecked = True Then lIndex = ID_OpenWrite Else lIndex = ID_ReadOnly
                .SubItems(1) = pvLoadString(hMod, "comdlg32.dll", lIndex, 100&)
                If Len(.SubItems(1)) = 0& Then
                    If isChecked = True Then .SubItems(1) = "Open for &write" Else .SubItems(1) = "Open as &read-only"
                End If
            Else
                .SubItems(1) = Caption2
            End If
        End With
    End If
    If hMod Then FreeLibrary hMod
    Controls_AddReadOnlyOption = True

End Function

Public Function Controls_Add(ByVal ContainerID As Long, ByVal ControlID As Long, _
                            ByVal ControlType As CommonDialogControlsEnum, _
                            ByVal CaptionText As String, ParamArray SubItemCaptions() As Variant) As Boolean
    
    ' Method adds extra controls and containers to the dialog
    ' Container: like a borderless frame with optional caption to the left of the container.
    '   Controls can be grouped/placed into a container. Container must be created first
    '   if a container is made enabled/invisible, all controls inside container are affected also
    '   Containers can be used to add labels to controls without labels, i.e., to add a label to a textbox:
    '       Controls_Add 100, 0, ctlType_Container, yourLabel       ' create container & its caption is the label
    '       Controls_Add 100, 101, ctlType_TextBox, textboxContent  ' add a textbox to the container
    
    ' Parameters:
    '   ContainerID, ControlID    ControlType               what is being created
    '   --------------------------------------------------------------------------------------------------------
    '       0          > 0        not ctlType_Container     containerless control
    '      > 0         > 0        not ctlType_Container     *containerized control. Container must exist
    '      > 0          0         ctlType_Container         container
    '      > 0         > 0        ctlType_Container         error, can't create a container in a container
    '      < 1         < 1        any                       error, invalid IDs
    '   --------------------------------------------------------------------------------------------------------
    '   (*) ControlType that cannot be containerized will ignore a valid ContainerID value
    '       Ignores ContainerID: ctlType_SelectionBoxLabel, ctlType_DefaultButton, ctlType_CancelButton
    
    '   CaptionText: applies to the following and required unless noted otherwise, else ignored
    '       ctlType_CheckBox,ctlType_Label,ctlType_CommandButton,ctlType_Menu: control's caption
    '       ctlType_TextBox: textbox content (not required, can be blank)
    '       ctlType_SelectionBoxLabel: caption of label next to dialog's "current selection" textbox
    '       ctlType_DefaultButton: Open/Save button's caption if  button not used as a split-button
    '       ctlType_CancelButton: Cancel button's caption; requires Win7+
    
    '   SubItemCaptions: applies to the following and required, else ignored. Do not pass null captions
    '       ctlType_DefaultButton: menu items only; converts the default Open/Save to a split-button
    '       ctlType_ComboBox: list items
    '       ctlType_OptionBtnGroup: each SubItemCaptions entry creates a button with the entry as button's caption
    '       ctlType_Menu: menu subitems
    
    ' Some notes:
    '   All added items are defaulted as not prominent, are enabled/visible, including any SubItems
    '       call Controls_PropertySet, later, to change property values as needed
    '   Controls do not need to be in a container, they can be added directly to the dialog
    '   Containers without any contained controls will be ignored when the dialog is created
    '   Controls within a container are added to the container in order of their creation
    '   Controls w/o containers and containers are added to the dialog in order of their creation
    
    If ControlID < 1& And ContainerID < 1& Then Exit Function   ' invalid IDs?
    If ControlID > &HFFFFFFF Or ContainerID > &HFFFFFFF Then Exit Function
    If Me.Version < dvVista Then Exit Function                  ' not applicable for pre-Vista
    If Not m_Dialog Is Nothing Then Exit Function               ' dialog created, called this too late
    
    Dim nrItems As Long, n As Long, lIndex As Long
    
    If ControlID > 0& Then
        If ContainerID > 0& Then
            ContainerID = pvKeyToIndex(ContainerID, False)
            If ContainerID = -1& Then Exit Function                 ' ContainerID invalid
            If (m_Controls(ContainerID).Flags And &HF) > 0& Then Exit Function ' ID is not for a container
        End If
        If pvKeyToIndex(ControlID, False) > 0& Then Exit Function   ' ControlID already exists
        nrItems = -1&
        Select Case ControlType
            
            Case ctlType_DefaultButton
                If pvKeyToIndex("def", False) > -1& Then Exit Function ' called this already
                nrItems = UBound(SubItemCaptions)
                If nrItems = -1& Then _
                    If Len(Trim$(CaptionText)) = 0& Then Exit Function ' caption or SubItemCaptions required
                ContainerID = 0&
                
            Case ctlType_ComboBox, ctlType_OptionBtnGroup, ctlType_Menu
                nrItems = UBound(SubItemCaptions)
                If nrItems = -1& Then Exit Function                      ' SubItemCaptions required
                If ControlType = ctlType_Menu Then _
                    If Len(Trim$(CaptionText)) = 0& Then Exit Function ' caption & SubItemCaptions required
            
            Case ctlType_CheckBox, ctlType_Label, ctlType_SelectionBoxLabel, ctlType_CommandButton
                If Len(Trim$(CaptionText)) = 0& Then Exit Function ' caption required
                If ControlType = ctlType_SelectionBoxLabel Then ContainerID = 0&
            
            Case ctlType_CancelButton                                   ' caption & Win7+ required
                If Me.Version < dvWin7 Or Len(Trim$(CaptionText)) = 0& Then Exit Function
                If ContainerID > 0& Then Exit Function
                If pvKeyToIndex("cnx", False) > -1& Then Exit Function  ' already modified
                ContainerID = 0&
            
            Case ctlType_TextBox
                If StrPtr(CaptionText) = 0& Then CaptionText = "" ' vbNullString prevents textbox from being added; fix this
            
            Case ctlType_Separator
            
            Case Else: Exit Function
        End Select
    ElseIf Not ControlType = ctlType_Container Then
        Exit Function                                           ' invalid parameter value(s)
    ElseIf pvKeyToIndex(ContainerID, False) > -1& Then
        Exit Function                                           ' ContainerID already exists
    Else
        lIndex = pvKeyToIndex(ContainerID, True)
        With m_Controls(lIndex)
            .ID = -ContainerID
            .Flags = &H30
            .Caption = CaptionText
            If StrPtr(.Caption) = 0& Then .Caption = ""         ' vbNullString not allowed
        End With
        Controls_Add = True
        Exit Function
    End If
    
    lIndex = pvKeyToIndex(ControlID, True)
    With m_Controls(lIndex)
        If nrItems > -1& Then
            On Error Resume Next       ' in case user is passing bogus paramArray items
            If VarType(SubItemCaptions(0)) = (vbArray Or vbString) Then     ' string array
                n = LBound(SubItemCaptions(0))
                nrItems = Abs(UBound(SubItemCaptions(0)) - n)
                If nrItems > &H7FFE& Then nrItems = &H7FFE& ' max 32767 items
                If n = 0& Then
                    .SubItems() = SubItemCaptions(0)
                    If nrItems < UBound(.SubItems) Then ReDim Preserve .SubItems(0 To nrItems)
                Else
                    For nrItems = 0& To nrItems
                        .SubItems(nrItems) = SubItemCaptions(0)(n)
                        If Err Then Exit For
                        n = n + 1&
                    Next
                End If
            Else
                If nrItems > &H7FFE& Then nrItems = &H7FFE&
                ReDim .SubItems(0 To nrItems)
                For nrItems = 0& To nrItems
                    .SubItems(nrItems) = SubItemCaptions(nrItems)
                    If Err Then Exit For
                Next
            End If
            If Err Then
                Err.Clear: nrItems = SIGN_BIT
            ElseIf ControlType = ctlType_DefaultButton Then
                .Flags = &H1000000 ' flag indicating split-button
            End If
            On Error GoTo 0
        End If
        .ID = -ControlID
        .Caption = CaptionText
        .Flags = .Flags Or ControlType Or &H30& ' default: enabled,visible,1st item selected (if any)
        If ContainerID > 0& Then .Flags = .Flags Or SIGN_BIT ' identifies this as containerized control
    End With
    If nrItems = SIGN_BIT Then  ' roll back this control addition; subItems is invalid
        ReDim Preserve m_Controls(0 To UBound(m_Controls) - 1)
        m_ControlKeys.Remove m_ControlKeys.Count
    Else
        If ContainerID > 0& Then            ' containerized control
            With m_Controls(ContainerID)
                nrItems = (.Flags And &H7FFF00) \ &H100&
                If nrItems < &H7FFF& Then
                    ReDim Preserve .SubItems(0 To nrItems)
                    .SubItems(nrItems) = ControlID
                    .Flags = (.Flags And &HFF0000FF) Or (nrItems + 1&) * &H100&
                End If
            End With
        End If
        If ControlType = ctlType_DefaultButton Then
            m_ControlKeys.Add UBound(m_Controls), "def"
        ElseIf ControlType = ctlType_CancelButton Then
            m_ControlKeys.Add UBound(m_Controls), "cnx"
        End If
        Controls_Add = True
    End If

End Function

Public Function Controls_TypeOf(ByVal ControlID As Long, Optional SubItemCount As Long) As CommonDialogControlsEnum
    
    ' Returns the type of control for passed ControlID; ControlID can be a container ID.
    '   to return the number of selected items when basket mode is used, pass ControlID as 0
    '   note: Controls_SetCustomMode must have been called with the IncludeSelCount parameter as True
    '       otherwise, internal counting of basket items isn't performed & SubItemCount will contain
    '       either 0 meaning no items added to basket or 1 meaning at least one item added
    '       Also note that the count may change when the user closes the dialog. This can happen
    '       if the user selected an item more than once. Duplicate items are removed from the
    '       basket when the dialog closes and that can reduce the count.
            
    ' Optionally returns the number of subitems within the control
    '   ctlType_ComboBox: number of list items
    '   ctlType_OptionBtnGroup: number of option buttons in group
    '   ctlType_Menu: number of menu items
    '   ctlType_DefaultButton: number of menu items, if button is a split-button
    '   ctlType_Container: number of contained controls
    '   all others: returns zero
    ' Return value of -1 indicates invalid ControlID passed or pre-Vista O/S

    Dim lReturn As Long
    SubItemCount = 0&
    If ControlID < 0& Or Me.Version < dvVista Then
        lReturn = -1&
    ElseIf ControlID = 0& Then  ' return basket count (see remarks above)
        lReturn = pvKeyToIndex(SIGN_BIT + 1&, False)
        If lReturn > -1& Then
            SubItemCount = (m_Controls(lReturn).Flags And &HFFFF000) \ &H1000&
            lReturn = ctlType_Label
        ElseIf Not m_BasketItems Is Nothing Then
            SubItemCount = Abs(m_BasketItems.Count > 0&)
            lReturn = ctlType_Label
        End If
    Else
        lReturn = pvKeyToIndex(ControlID, False)
        If lReturn > 0& Then
            With m_Controls(lReturn)
                lReturn = .Flags And &HF
                Select Case lReturn
                Case ctlType_ComboBox, ctlType_OptionBtnGroup, ctlType_Menu
                    SubItemCount = UBound(.SubItems) + 1&
                Case ctlType_DefaultButton  ' check flag indicating split-button
                    If (.Flags And &H3000000) Then SubItemCount = UBound(.SubItems) + 1&
                Case ctlType_Container
                    SubItemCount = (.Flags And &H7FFF00) \ &H100&
                End Select
            End With
        End If
    End If
    Controls_TypeOf = lReturn

End Function

Public Sub Controls_PropertySet(ByVal ControlID As Long, _
                            ByVal propType As PropertyTypeEnum, Value As Variant, _
                            Optional ByVal SubItemlIndex As Long = -1&)
    
    ' ControlID must exist, ID can be for a control or container
    '   note: Containers can have their properties set also
    '   note: most other controls can only have their SelIndex, Enabled, Visible set when dialog is shown
    ' propType
    '   ppEnabled, ppVisible, ppProminent: Value must be either True or False
    '   ppText: Captions can be set before dialog displayed and some while displayed. Value is string
    '   ppSelValue: If control has subitems, Value is selected subitem, zero-bound
    '   ppTag: Sets the user-defined Tag property associated with the control
    '   ppChildControlID: does not apply to this routine
    '   if property doesn't apply, it is ignored
    ' Optional SubItemlIndex
    '   allows changing the enabled or visible state of subItems (zero-bound)
    
    ' Property types and expected Value parameter
    ' ControlType            ppEnabled  ppVisible  ppText   ppProminent  ppSelValue
    ' ------------------------------------------------------------------------------
    ' ctlType_CheckBox          T/F       T/F       n/a         T/F     vbChecked/vbUnchecked
    ' ctlType_ComboBox#         T/F       T/F       n/a         T/F     -1 or 0-bound list index(^)
    ' ctlType_OptionBtnGroup#   T/F       T/F       n/a         n/a     -1 or 0-bound index(^)
    ' ctlType_Menu#             T/F       T/F       n/a         T/F     n/a
    ' ctlType_CommandButton     T/F       T/F       Caption     T/F     n/a
    ' ctlType_TextBox           T/F       T/F       Text value  n/a     n/a
    ' ctlType_Separator         n/a       n/a       n/a         n/a     n/a
    ' ctlType_Label             T/F       T/F       Caption     n/a     n/a
    ' ctlType_DefaultButton*#   n/a       n/a       Caption     n/a     n/a
    ' ctlType_CancelButton**    n/a       n/a       Caption     n/a     n/a
    ' ctlType_SelectionBoxLabel n/a       n/a       Caption     n/a     n/a
    ' ctlType_Container         T/F       T/F       Caption     n/a     n/a
    ' ------------------------------------------------------------------------------
    ' (*) If button set up as split-button, Caption can't be changed after dialog created
    ' (**) Requires Win7 or higher
    ' (#) Can set enabled/visible state for individual subitems, if applicable
    ' (n/a) Property doesn't exist or cannot be changed after dialog is created
    ' (^) -1 is only valid until the dialog is created
    '  Prominent property value can't be changed after dialog is created
    
    If Me.Version < dvVista Then Exit Sub       ' not applicable for pre-Vista
    
    Dim lIndex As Long, lValue As Long
    
    lIndex = pvKeyToIndex(ControlID, False)
    If lIndex < 1& Then Exit Sub
    
    On Error GoTo ExitRoutine       ' in case user passed something suspect
    If (m_Controls(lIndex).Flags And &HF) = ctlType_Container And SubItemlIndex > -1& Then
        If Not propType = ppChildControlID Then
            Controls_PropertySet CLng(m_Controls(lIndex).SubItems(SubItemlIndex)), propType, Value
        End If
        On Error GoTo 0
        Exit Sub
    End If
    
    With m_Controls(lIndex)
        Select Case propType
        Case ppEnabled, ppVisible
            If SubItemlIndex < 0& Then  ' set control vs subitem state
                Select Case (.Flags And &HF)
                Case ctlType_DefaultButton, ctlType_SelectionBoxLabel ' can't toggle them
                Case Else
                    If propType = ppEnabled Then
                        .Flags = (.Flags And &HFFFFFFEF) Or Abs(CBool(Value)) * &H10&
                    Else
                        .Flags = (.Flags And &HFFFFFFDF) Or Abs(CBool(Value)) * &H20&
                    End If
                    If Not m_DlgMod Is Nothing Then _
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlState, .ID, (.Flags And &H30) \ &H10
                End Select
            Else                                ' set subitem state, if applicable
                Select Case (.Flags And &HF)
                Case ctlType_ComboBox, ctlType_OptionBtnGroup, ctlType_Menu
                    ' above have subitems
                Case ctlType_DefaultButton ' split-buttons have submenus
                    If (.Flags And &H3000000) = 0& Then Exit Sub
                Case Else: Exit Sub
                End Select
                If (.Flags And &H80) = 0& Then ' individual attributes not yet set
                    ReDim .SubItemState(0 To UBound(.SubItems))
                    For lValue = 0& To UBound(.SubItemState): .SubItemState(lValue) = 3&: Next
                    .Flags = .Flags Or &H80
                End If
                If propType = ppEnabled Then
                    lValue = (.SubItemState(SubItemlIndex) And &H2&) Or Abs(CBool(Value))
                Else
                    lValue = (.SubItemState(SubItemlIndex) And &H1&) Or Abs(CBool(Value)) * 2&
                End If
                .SubItemState(SubItemlIndex) = lValue
                If Not m_DlgMod Is Nothing Then _
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlItemState, .ID, SubItemlIndex, lValue
            End If
        Case ppText
            Select Case (.Flags And &HF)
            Case ctlType_TextBox
                .Caption = Value
                If StrPtr(.Caption) = 0& Then .Caption = "" ' null not allowed, use this instead
                If Not m_DlgMod Is Nothing Then _
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetEditBoxText, .ID, StrPtr(.Caption)
            Case ctlType_CommandButton, ctlType_Label, ctlType_Container
                If Len(Value) = 0& Then _
                    If (.Flags And &HF) > ctlType_Container Then GoTo ExitRoutine ' don't allow null captions
                .Caption = Value
                If (.Flags And &HF) = ctlType_Container Then _
                    If StrPtr(.Caption) = 0& Then .Caption = "" ' don't allow null captions
                If Not m_DlgMod Is Nothing Then _
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlLabel, .ID, StrPtr(.Caption)
            Case ctlType_DefaultButton
                If (.Flags And &H3000000) = 0& Then ' not split-button
                    If Len(Value) = 0& Then GoTo ExitRoutine
                    .Caption = Value
                    If Not m_DlgMod Is Nothing Then _
                        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetOkButtonLabel, StrPtr(.Caption)
                End If
            Case ctlType_CancelButton
                If Len(Value) = 0& Then GoTo ExitRoutine
                .Caption = Value
                If Not m_Dlg2 Is Nothing Then _
                    pvCallInterface ObjPtr(m_Dlg2), ifFileDialog2_SetCancelButtonLabel, StrPtr(.Caption)
            Case ctlType_SelectionBoxLabel
                If Len(Value) = 0& Then GoTo ExitRoutine
                .Caption = Value
                If Not m_DlgMod Is Nothing Then _
                    pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileNameLabel, StrPtr(.Caption)
            End Select
        Case ppProminent
            If m_StateFlags < 0& Then GoTo ExitRoutine  ' not applicable if dialog displayed
            Select Case (.Flags And &HF)
            Case ctlType_CheckBox, ctlType_CommandButton, ctlType_ComboBox, ctlType_Menu
                .Flags = (.Flags And Not &H40) Or Abs(Value) * &H40
            End Select
        Case ppSelValue
            Select Case (.Flags And &HF)
            Case ctlType_OptionBtnGroup, ctlType_ComboBox
                If Value = -1& Then                     ' applicable only if dialog not yet active
                    If m_StateFlags > -1& Then .Flags = .Flags Or &H10000000
                ElseIf Not (Value < 0& Or Value > UBound(.SubItems())) Then
                    .Flags = (.Flags And &HF00000FF) Or Value * &H100&
                    If Not m_DlgMod Is Nothing Then _
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetSelectedControlItem, .ID, CLng(Value)
                End If
            Case ctlType_CheckBox
                .Flags = (.Flags And &HF00000FF) Or Abs(CBool(Value))
                If Not m_DlgMod Is Nothing Then _
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetCheckButtonState, .ID, CLng(Value)
            End Select
        Case ppTag
            If IsObject(Value) Then Set .Tag = Value Else .Tag = Value
        End Select
    End With

ExitRoutine:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Public Function Controls_PropertyGet(ByVal ControlID As Long, _
                                    ByVal propType As PropertyTypeEnum, _
                                    Optional ByVal SubItemlIndex As Long = -1&) As Variant
    
    ' ControlID must exist, ID can be for a control or container
    '   note: Containers can have their properties returned also
    ' propType
    '   ppEnabled, ppVisible, ppProminent: function returns True or False
    '   ppText: If control has caption or text property, it is returned as string
    '   ppSelValue: If control has subitems, selected subitem passed, zero-bound
    '   ppTag: Returns the user-defined Tag property associated with the control
    '   ppChildControlID: applies only if ControlID is for a container and SubItemIndex is passed
    '       -- returns ControlID for the container's child relative to SubItemIndex
    '       -- allows navigating containerized controls based on an ordinal index
    '   a value of -1 is used internally to update the m_Controls() collection with current values
    ' Optional SubItemlIndex
    '   allows returning the enabled/visible state or caption of subItems (zero-bound)
    
    ' Property values returned
    ' ControlType           ppEnabled  ppVisible  ppProminent   ppText      ppSelValue
    ' -----------------------------------------------------------------------------------------------
    ' ctlType_CheckBox          T/F       T/F       T/F         Caption     vbChecked/vbUnchecked
    ' ctlType_ComboBox#         T/F       T/F       T/F         ""          -1 or 0-bound list index
    ' ctlType_OptionBtnGroup#   T/F       T/F       False       ""          -1 or 0-bound index
    ' ctlType_Menu#             T/F       T/F       T/F         Caption     n/a
    ' ctlType_TextBox**         T/F       T/F       False       Text value  n/a
    ' ctlType_CommandButton     T/F       T/F       T/F         Caption     n/a
    ' ctlType_Separator         T/F       T/F       False       ""          n/a
    ' ctlType_Label             T/F       T/F       False       Caption     n/a
    ' ctlType_DefaultButton#*   True      True      False       Caption     n/a
    ' ctlType_CancelButton      T/F       T/F       False       Caption     n/a
    ' ctlType_SelectionBoxLabel True      True      False       Caption     n/a
    ' ctlType_Container         T/F       T/F       False       Caption     n/a
    ' -----------------------------------------------------------------------------------------------
    ' (#) If SubItemIndex is passed, Caption is the subitem caption
    ' (*) If button set up as split-button & SubItemIndex not passed, caption is 1st subitem caption
    ' (**) When dialog displayed, must be retrieved before dialog closes. Note that this class
    '       automatically does that only if the dialog is opened with WantEvents option

    If Me.Version < dvVista Then Exit Function       ' not applicable for pre-Vista
    
    Dim lIndex As Long, lPtr As Long
    lIndex = pvKeyToIndex(ControlID, False)
    If lIndex = -1& Then Exit Function  ' invalid ControlID
    
    If (m_Controls(lIndex).Flags And &HF) = ctlType_Container And SubItemlIndex > -1& Then
        On Error GoTo ExitRoutine
        If propType = ppChildControlID Then
            Controls_PropertyGet = CLng(m_Controls(lIndex).SubItems(SubItemlIndex))
        Else
            Controls_PropertyGet = Controls_PropertyGet(CLng(m_Controls(lIndex).SubItems(SubItemlIndex)), propType)
        End If
        On Error GoTo 0
        Exit Function
    End If
    
    With m_Controls(lIndex)
        If m_StateFlags < 0& And (propType = ppSelValue Or propType = ppText Or propType = -1&) And SubItemlIndex < 0& Then
            ' get real-time values as appropriate
            If propType = ppSelValue Or propType = -1& Then
                Select Case (.Flags And &HF)
                Case ctlType_ComboBox, ctlType_Menu, ctlType_OptionBtnGroup
                    If pvCallInterface(ObjPtr(m_DlgMod), ifFileDlgCustomize_GetSelectedControlItem, .ID, VarPtr(lPtr)) = 0& Then
                        .Flags = (.Flags And &HEF0000FF) Or (lPtr * &H100&) ' this removes the &H10000000 flag
                        Controls_PropertyGet = lPtr
                    Else
                        Controls_PropertyGet = -1&
                    End If
                Case ctlType_DefaultButton
                    If (.Flags And &H3000000) Then
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_GetSelectedControlItem, .ID, VarPtr(lPtr)
                        If (.Flags And &H2000000) Then lPtr = lPtr - 1&  ' treat as a checkbox value
                        .Flags = (.Flags And &HFF0000FF) Or lPtr * &H100&
                        Controls_PropertyGet = lPtr
                    End If
                Case ctlType_CheckBox
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_GetCheckButtonState, .ID, VarPtr(lPtr)
                    .Flags = (.Flags And &HFF0000FF) Or lPtr * &H100&
                    Controls_PropertyGet = lPtr
                Case Else
                    Controls_PropertyGet = True
                End Select
            End If
            If propType = ppText Or propType = -1& Then
                Select Case (.Flags And &HF)
                Case ctlType_TextBox
                    If m_StateFlags < 0& Then ' dialog running
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_GetEditBoxText, .ID, VarPtr(lPtr)
                        .Caption = pvComStrPtrToVBstring(lPtr, True)
                    End If
                    If Not propType = -1& Then Controls_PropertyGet = .Caption
                Case ctlType_CheckBox, ctlType_CommandButton, ctlType_SelectionBoxLabel, _
                        ctlType_Label, ctlType_CancelButton, ctlType_Menu, ctlType_Container
                    If propType = ppText Then Controls_PropertyGet = .Caption
                Case ctlType_DefaultButton
                    If propType = ppText Then
                        If (.Flags And &H3000000) = 0& Then
                            Controls_PropertyGet = .Caption
                        Else
                            Controls_PropertyGet = .SubItems(0)
                        End If
                    End If
                Case Else
                    Controls_PropertyGet = vbNullString
                End Select
            End If
        
        ElseIf propType = ppEnabled Or propType = ppVisible Then
            If propType = ppEnabled Then lPtr = &H10 Else lPtr = &H20
            If SubItemlIndex < 0& Then
                Controls_PropertyGet = CBool(.Flags And lPtr)
            ElseIf (.Flags And &H80) Then
                Controls_PropertyGet = CBool(.SubItemState(SubItemlIndex Or (lPtr \ &H10)))
            Else
                Controls_PropertyGet = True
            End If
        ElseIf propType = ppText Then
            If SubItemlIndex < 0& Then
                Controls_PropertyGet = .Caption
                If (.Flags And &H3000000) Then Controls_PropertyGet = .SubItems(0)
            Else
                On Error Resume Next
                Select Case (.Flags And &HF)
                Case ctlType_ComboBox, ctlType_Menu
                    Controls_PropertyGet = .SubItems(SubItemlIndex)
                Case ctlType_DefaultButton
                    If (.Flags And &H3000000) Then
                        Controls_PropertyGet = .SubItems(SubItemlIndex)
                    Else
                        Controls_PropertyGet = vbNullString
                    End If
                Case Else
                    Controls_PropertyGet = vbNullString
                End Select
                If Err Then
                    Err.Clear
                    Controls_PropertyGet = vbNullString
                End If
                On Error GoTo 0
            End If
        ElseIf propType = ppProminent Then
            Controls_PropertyGet = CBool(.Flags And &H40)
        ElseIf propType = ppTag Then
            If IsObject(.Tag) Then Set Controls_PropertyGet = .Tag Else Controls_PropertyGet = .Tag
            
        ElseIf propType = ppSelValue Then
            Select Case (.Flags And &HF)
            Case ctlType_CheckBox
                Controls_PropertyGet = (.Flags And &H7FFF00) \ &H100&
            Case ctlType_ComboBox, ctlType_OptionBtnGroup
                If (.Flags And &H10000000) Then
                    Controls_PropertyGet = -1&
                Else
                    Controls_PropertyGet = ((.Flags And &H7FFF00) \ &H100&)
                End If
            Case ctlType_DefaultButton  ' if used as Read-Only or split button
                If (.Flags And &H3000000) Then Controls_PropertyGet = (.Flags And &H7FFF00) \ &H100&
            End Select
        End If
    End With

ExitRoutine:
    If Err Then Err.Clear
    On Error GoTo 0
    If IsEmpty(Controls_PropertyGet) Then Controls_PropertyGet = 0&
End Function

Public Property Get IFileDialog() As stdole.IUnknown ' Vista+ only
    ' returns a reference to the active dialog, implementing IFileOpenDialog or IFileSaveDialog
    ' If IFileOpenDialog or IFileSaveDialog is needed instead, simply assign it on the return
    '   if you have a TLB: Dim IFOD As IFileOpenDialog: Set IFOD = cBrowser.IFileDialog
    '   if no TLB, then you'll be using CallDispFunc API, query returned IUnknown for IFileOpenDialog
    Set IFileDialog = m_Dialog
End Property

Public Property Get IFileDialog_IFolderView2() As stdole.IUnknown ' Vista+ only
    ' returns a reference to the active dialog's current IFolderView2 interface
    ' If IFolderView vs IFolderView2 is needed, simply assign it on the return
    '   if you have a TLB: Dim IFV As IFolderView: Set IFV = cBrowser.IFileDialog_IFolderView2
    '   if no TLB, then you'll be using CallDispFunc API, query returned IUnknown for IFolderView
    ' FYI: IFolderView2 inherits from IFolderView
    
    If m_StateFlags < 0& Then                   ' else dialog not running or pre-Vista
        Set IFileDialog_IFolderView2 = pvGetIFolderView2
    End If
    
End Property

Public Function IFileDialog_GetViewMode(Optional IconSizePixels As Long) As IShellViewModeEnum ' Vista+ only
    ' returns the current folder's view mode: list, details, small icon, etc and if icon mode: icon size
    ' cannot be called before dialog is running. If you want to control this, you need to
    '   start the dialog by asking for Events and respond to the DialogOnFolderChanged event

    ' IconSizePixels parameter set to -1 when size is "default"
    
    If m_StateFlags < 0& Then                   ' else dialog not running or pre-Vista
        Dim IFolderView2 As stdole.IUnknown
        Const IfIFolderView2_GetVwModeIcoSize As Long = 36
        
        Set IFolderView2 = pvGetIFolderView2
        If Not IFolderView2 Is Nothing Then _
            pvCallInterface ObjPtr(IFolderView2), IfIFolderView2_GetVwModeIcoSize, VarPtr(IFileDialog_GetViewMode), VarPtr(IconSizePixels)
    End If
    
End Function

Public Function IFileDialog_SetViewMode(ByVal Mode As IShellViewModeEnum, Optional ByVal IconSizePixels As Long = -1&) As Boolean ' Vista+ only
    ' sets the current folder's view mode: list, details, small icon, etc and if icon mode: icon size
    ' cannot be called before dialog is running. If you want to control this, you need to
    '   start the dialog by asking for Events and respond to the DialogOnFolderChanged event
    
    ' set IconSizePixels parameter to -1 for "default" size
    
    If m_StateFlags < 0& Then                   ' else dialog not running or pre-Vista
        Dim IFolderView2 As stdole.IUnknown
        Const IfIFolderView2_SetVwModeIcoSize As Long = 35
        
        If IconSizePixels < -1& Then Exit Function
        If Mode < svm_Auto Or Mode = 0& Then Exit Function
        If Mode = svm_Content And Me.Version < dvWin7 Then Mode = svm_Details ' svm_Content requires Win7+
        
        Set IFolderView2 = pvGetIFolderView2
        If Not IFolderView2 Is Nothing Then
            pvCallInterface ObjPtr(IFolderView2), IfIFolderView2_SetVwModeIcoSize, (Mode And &HF&), IconSizePixels
            IFileDialog_SetViewMode = True
        End If
    End If
    
End Function

Public Function IFileDialog_GetHwnd() As Long ' Vista+ only

    ' returns the dialog hWnd for Vista+ dialog only
    ' For XP/Win2K, this function returns zero
    '   If you are receiving XP/Win2K events, all events include that dialog's hWnd
    '   If you need the hWnd, then call ShowOpen/ShowSave with the optional WantEvents parameter
    
    If Not m_Dialog Is Nothing Then
        Dim oPtr As Long
        Const IID_OleWindow As String = "{00000114-0000-0000-C000-000000000046}"
        
        IIDFromString StrPtr(IID_OleWindow), VarPtr(m_Guid(0))
        pvCallInterface ObjPtr(m_Dialog), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(oPtr)
        If oPtr Then
            pvCallInterface oPtr, ifOleWindow_GetWindow, VarPtr(IFileDialog_GetHwnd)
            pvCallInterface oPtr, ifUnknown_Release
        End If
    End If
    
End Function

Public Function IFileDialog_GetCurrentFolder() As stdole.IUnknown ' Vista+ only
    ' returns IShellItem of folder currently selected in the dialog
    ' if dialog not yet displayed, returns folder that dialog will initially display
    If Not m_Dialog Is Nothing Then _
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetFolder, VarPtr(IFileDialog_GetCurrentFolder)
End Function

Public Function IFileDialog_GetCurrentSelection(Optional ByVal WantArrayOfItems As Boolean = False, _
                                                Optional Count As Long) As stdole.IUnknown ' Vista+ only
    
    ' WantArrayOfItems if true will return IShellItemArray else first selected IShellItem, if any
    ' Count will return the number of items contained by the returned IShellItem/IShellItemArray
    
    ' notes: When multi-selection is in play, and user selects multiple folders when files are expected,
    '   then the dialog may only return the 1st selected folder. To get the true selection, regardless
    '   of dialog behavior, see pvProcessCustomMode.
    
    If m_StateFlags > -1& Then Exit Function    ' dialog not active
    
    Dim pUnk As Long, oItem As stdole.IUnknown
    Dim n As Long
    
    Count = 0&
    If WantArrayOfItems Then
        ' want item array vs single shell item
        ' will fail if dialog is called to Save vs Open. Failure handled later
        IIDFromString StrPtr(IID_IFileOpenDialog), VarPtr(m_Guid(0))
        pvCallInterface ObjPtr(m_Dialog), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(pUnk)
        If pUnk Then                                ' get IShellItemArray
            pvCallInterface pUnk, ifFileDialogOpen_GetSelectedItems, VarPtr(oItem)
            pvCallInterface pUnk, ifUnknown_Release ' release IFileDialog
            If Not oItem Is Nothing Then _
                pvCallInterface ObjPtr(oItem), ifShellItemArray_GetCount, VarPtr(Count)
        End If
    End If
    If pUnk = 0& Then
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetCurrentSelection, VarPtr(oItem)
        If Not oItem Is Nothing Then
            Count = 1&
            If WantArrayOfItems Then    ' convert from IShellItem to IShellItemArray
                ' note: we could use this API instead: SHCreateShellItemArrayFromShellItem
                ' but this scenario is highly unlikely; don't want to add an API not likely to be used
                SHGetIDListFromObject ObjPtr(oItem), pUnk
                If pUnk Then
                    Set oItem = Nothing
                    SHCreateShellItemArrayFromIDLists Count, VarPtr(pUnk), oItem
                    CoTaskMemFree pUnk
                End If
            End If
        End If
    End If
    Set IFileDialog_GetCurrentSelection = oItem

End Function

Public Function IShellItem_Action(oShellItem As stdole.IUnknown, ByVal Action As IShellItemActionEnum, _
                            Optional ByVal pInterfaceVarPtr As Long) As Long

    ' courtesy function to return common IShellItem property values
    
    ' oShellItem is an interface returned/passed from this class or a valid TLB IShellItem
    
    ' Return value depends on Action
    '   siGetAttributes: returns shell attributes for the passed IShellItem
    '   siGetParentIShellItem: returns non-zero if successful
    '       pass pInterfaceVarPtr as VarPtr() of a declared stdole.IUnknown or TLB IShellItem
    '       to return parent IShellItem interface of passed oShellItem, if any
    '   siGetIStream: returns non-zero if successful
    '       pass pInterfaceVarPtr as VarPtr() of a declared stdole.IUnknown or TLB IStream
    '       to return an IStream interface from passed oShellItem, if any
    '   siGetIStorage: returns non-zero if successful
    '       pass pInterfaceVarPtr as VarPtr() of a declared stdole.IUnknown or TLB IStorage
    '       to return an IStorage interface from passed oShellItem, if any
    '   siGetIDataObject: returns non-zero if successful
    '       pass pInterfaceVarPtr as VarPtr() of a declared stdole.IUnknown or TLB IDataObject
    '       to return an IDataObject interface from passed oShellItem, if any
    
    If oShellItem Is Nothing Then Exit Function
    If Not Action = siGetAttributes Then
        If pInterfaceVarPtr = 0& Then Exit Function
    End If
    
    Dim aGuid(0 To 7) As Long, oItem As stdole.IUnknown
    Dim pIDL As Long, oPAI As stdole.IUnknown
    Const IID_IStorage          As String = "{0000000b-0000-0000-c000-000000000046}"
    Const IID_IStream           As String = "{0000000C-0000-0000-C000-000000000046}"
    Const IID_IDataObject       As String = "{0000010E-0000-0000-C000-000000000046}"
    Const IID_IParentAndItem    As String = "{B3A4B685-B685-4805-99D9-5DEAD2873236}"
    Const IID_BHID_ISTREAM      As String = "{1CEBB3AB-7C10-499a-A417-92CA16C4CB83}"
    Const IID_BHID_IDATAOBJECT  As String = "{B8C0BD9F-ED24-455C-83E6-D5390C4FE8C4}"
    Const ifShellItem_GetAttributes As Long = 6
    Const ifShellItem_GetParent As Long = 4
    Const ifShellItem_BindToHandler As Long = 3
    Const ifPAI_GetParentAndItem As Long = 4
    Const ifShellFolder_BindToStorage As Long = 8

    IIDFromString StrPtr(IID_IShellItem), VarPtr(aGuid(0))
    pvCallInterface ObjPtr(oShellItem), ifUnknown_QueryInterface, VarPtr(aGuid(0)), VarPtr(oItem)
    If Not oItem Is Nothing Then
        Select Case Action
        
        Case siGetAttributes
            pvCallInterface ObjPtr(oItem), ifShellItem_GetAttributes, -1&, VarPtr(IShellItem_Action)
        
        Case siGetParentIShellItem
            IShellItem_Action = (pvCallInterface(ObjPtr(oItem), ifShellItem_GetParent, pInterfaceVarPtr) = 0&)
            
        Case siGetIStream
            IIDFromString StrPtr(IID_BHID_ISTREAM), VarPtr(aGuid(0))
            IIDFromString StrPtr(IID_IStream), VarPtr(aGuid(4))
            IShellItem_Action = (pvCallInterface(ObjPtr(oItem), ifShellItem_BindToHandler, 0&, VarPtr(aGuid(0)), VarPtr(aGuid(4)), pInterfaceVarPtr) = 0&)
                
        Case siGetIDataObject
            IIDFromString StrPtr(IID_BHID_IDATAOBJECT), VarPtr(aGuid(0))
            IIDFromString StrPtr(IID_IDataObject), VarPtr(aGuid(4))
            IShellItem_Action = (pvCallInterface(ObjPtr(oItem), ifShellItem_BindToHandler, 0&, VarPtr(aGuid(0)), VarPtr(aGuid(4)), pInterfaceVarPtr) = 0&)
            
        Case siGetIStorage
            ' need item's parent as IShellFolder and item as PIDL (relative to parent, not absolute)
            IIDFromString StrPtr(IID_IParentAndItem), VarPtr(aGuid(0))
            pvCallInterface ObjPtr(oItem), ifUnknown_QueryInterface, VarPtr(aGuid(0)), VarPtr(oPAI)
            If oPAI Is Nothing Then Exit Function ' shouldn't get no interface
            Set oItem = Nothing
            pvCallInterface ObjPtr(oPAI), ifPAI_GetParentAndItem, 0&, VarPtr(oItem), VarPtr(pIDL)
            If Not (pIDL = 0& Or oItem Is Nothing) Then
                IIDFromString StrPtr(IID_IStorage), VarPtr(aGuid(0))
                IShellItem_Action = (pvCallInterface(ObjPtr(oItem), ifShellFolder_BindToStorage, pIDL, 0&, VarPtr(aGuid(0)), pInterfaceVarPtr) = 0&)
            End If
            If pIDL Then CoTaskMemFree pIDL                         ' free pIDL as needed
        End Select
    End If

End Function

Public Function IShellItem_GetDisplayName(oShellItem As stdole.IUnknown, _
                                        ByVal DisplayFormat As DisplayFormatEnum, _
                                        Optional ByVal FallBackToAbsoluteParsing As Boolean, _
                                        Optional ByVal PreferSysPathOverAbsolute As Boolean) As String
    
    ' courtesy function to return IShellItem display name
    
    ' oShellItem is an interface returned/passed from this class or a valid TLB IShellItem
    
    ' FallBackToAbsoluteParsing when true will requery the display name using dfn_DesktopAbsoluteParsing
    '   if the requested display name returns null string. Otherwise, return value can be null string.
    
    ' PreferSysPathOverAbsolute attempts to return a system path, i.e., disk file path, if it can be
    '   retrieved from a virtual/protected folder and applies only if DisplayFormat = dfn_FileSysPath.
    '   For example: fonts within the C:\Windows\Fonts folder on newer operating systems can fail to
    '       return a display name when dfn_FileSysPath is used. This typically applies when a font family
    '       contains multiple font styles. In any case, PreferSysPathOverAbsolute will return a true
    '       disk file path whereas FallBackToAbsoluteParsing will not. But FallBackToAbsoluteParsing
    '       should be guaranteed to return something.
    '   If both FallBackToAbsoluteParsing and PreferSysPathOverAbsolute are true,
    '       PreferSysPathOverAbsolute takes precedence and FallBackToAbsoluteParsing performed on failure
    
    
    If oShellItem Is Nothing Then Exit Function
    
    Dim oItem As stdole.IUnknown, lPtr As Long, sName As String
    
    IIDFromString StrPtr(IID_IShellItem), VarPtr(m_Guid(0))
    pvCallInterface ObjPtr(oShellItem), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(oItem)
    If Not oItem Is Nothing Then
        If PreferSysPathOverAbsolute = True And DisplayFormat = dfn_FileSysPath Then
            ' already doing this internally, call internal routine
            sName = pvGetDisplayName(ObjPtr(oItem), DisplayFormat)
        Else
            pvCallInterface ObjPtr(oItem), ifShellItem_GetDisplayName, DisplayFormat, VarPtr(lPtr)
            If lPtr Then sName = pvComStrPtrToVBstring(lPtr): lPtr = 0&
        End If
        If Len(sName) = 0& And FallBackToAbsoluteParsing = True Then
            pvCallInterface ObjPtr(oItem), ifShellItem_GetDisplayName, dfn_DesktopAbsoluteParsing, VarPtr(lPtr)
            If lPtr Then sName = pvComStrPtrToVBstring(lPtr)
        End If
    End If
    IShellItem_GetDisplayName = sName
                                        
End Function

Public Function IShellItemArray_Action(oShellItemArray As stdole.IUnknown, _
                            ByVal Action As IShellItemArrayActionEnum, _
                            Optional ByVal ArrayIndex As Long, _
                            Optional ByVal pInterfaceVarPtr As Long) As Long
                            
    ' courtesy function to return common IShellItemArray property values
    
    ' oShellItemArray is an interface returned/passed from this class or a valid TLB IShellItemArray
    
    ' Return value depends on Action
    '   siGetArrayItemCount: returns the number of IShellItem interfaces contained in passed array
    '       other parameters are ignored
    '   siGetIShellItemAt: returns non-zero if successful
    '       pass pInterfaceVarPtr as VarPtr() of a declared stdole.IUnknown or TLB IShellItem
    '       to return the IShellItem at ArrayIndex
    
    If oShellItemArray Is Nothing Then Exit Function
    
    Dim oItem As stdole.IUnknown
    
    IIDFromString StrPtr(IID_IShellItemArray), VarPtr(m_Guid(0)) ' ensure IShellItemArray passed
    pvCallInterface ObjPtr(oShellItemArray), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(oItem)
    If Not oItem Is Nothing Then
        Select Case Action
        Case siGetArrayItemCount
            pvCallInterface ObjPtr(oItem), ifShellItemArray_GetCount, VarPtr(IShellItemArray_Action)
        Case siGetIShellItemAt
            If pInterfaceVarPtr = 0& Then Exit Function
            If ArrayIndex > -1& Then
                pvCallInterface ObjPtr(oItem), ifShellItemArray_GetCount, VarPtr(IShellItemArray_Action)
                If ArrayIndex < IShellItemArray_Action Then _
                    IShellItemArray_Action = (pvCallInterface(ObjPtr(oItem), ifShellItemArray_GetItemAt, ArrayIndex, pInterfaceVarPtr) = 0&)
            End If
        End Select
    End If

End Function

Public Function SetRootFolder(RootFolder As Variant) As Boolean
    ' For Folder Browser, must should a PIDL, anything else will be attempted to be converted to PIDL
    ' For the FileDialog (Win7+ only)
    '   can pass one of these sources but must be a folder reference, not file reference
    '   String path, i.e., C:\Windows
    '   PIDL, return value from various PIDL-related APIs
    '    example for CSIDL_COMMON_PROGRAMS (value=23), simply pass: CSIDLenum.CSIDL_COMMON_PROGRAMS or just 23
    '   String GUID, i.e. control panels GUID: ::{26EE0668-A00A-44D7-9371-BEB064C98683}
    '       control panel GUIDs must be prefixed with double colons else may not be usable
    '    another example using class function, pass: GetKnownFolderGUID(eFOLDERID_AppDataProgramData)
    '   IShellItem as an stdole.IUnknown
    ' You are responsible for destroying any passed PIDLs, any time after this routine returns
    
    ' coder note: m_NavRestrictor holds the value until the dialog is created
    ' if a FileDialog, m_NavRestrictor will eventually contain a IShellItem reference
    ' if a FolderBrowser, m_NavRestrictor will be emptied and browser structure is given the PIDL
    
    Dim n As Long, sValue As String
    
    If (m_StateFlags And &H8F000000) Then Exit Function  ' dialog running
    
    pvFreeOrCreateRootPIDL False, True
    If VarType(RootFolder) = vbDataObject Then
        Set m_NavRestrictor = RootFolder
        
    ElseIf IsNumeric(RootFolder) Then
        On Error Resume Next
        n = CLng(RootFolder)
        If Err Then
            Err.Clear
        ElseIf (n And &HFF) = n Then    ' CSIDL
            sValue = String$(MAX_PATH, vbNullChar)
            SHGetFolderPath 0&, n Or CSIDL_FLAG_NO_ALIAS, 0&, 0&, StrPtr(sValue)
            n = InStr(sValue, vbNullChar)
            If n > 1 Then m_NavRestrictor = Left$(sValue, n - 1)
        Else                            ' PIDL
            m_NavRestrictor = ILCloneFull(RootFolder)
            If m_NavRestrictor = 0& Then m_NavRestrictor = Empty
        End If
        On Error GoTo 0
        
    ElseIf VarType(RootFolder) = vbString Then
        If InStr(RootFolder, "/") Then
            m_NavRestrictor = Replace(RootFolder, "/", "\")
        ElseIf InStr(RootFolder, "\") Then
            m_NavRestrictor = RootFolder
        ElseIf InStr(RootFolder, "{") Then
            If Right$(RootFolder, 1) = "}" Then
                If Left$(RootFolder, 3) = "::{" Then
                    If Len(RootFolder) = 40 Then m_NavRestrictor = RootFolder
                ElseIf Len(RootFolder) = 38 Then
                    If Me.Version > dvXP_Win2K Then
                        If IIDFromString(StrPtr(CStr(RootFolder)), VarPtr(m_Guid(0))) = 0& Then
                            If SHGetKnownFolderPath(VarPtr(m_Guid(0)), 0&, 0&, n) = 0& Then _
                                m_NavRestrictor = pvComStrPtrToVBstring(n, True)
                        End If
                    End If
                    If IsEmpty(m_NavRestrictor) Then _
                        m_NavRestrictor = "::" & RootFolder ' gotta assume it's a control panel folder
                End If
            End If
        End If
    End If
    SetRootFolder = Not IsEmpty(m_NavRestrictor)
    
End Function

Public Function SetNavFolders(ByVal AtPanelTop As Boolean, ParamArray Folders() As Variant) As Boolean
    ' Vista+ only
    ' can pass multiple Folders (only) in the paramArray. Each must be one of these, mixing is ok.
    '  Valid string path , i.e., C:\Program Files
    '  PIDL, return value from various PIDL-related APIs
    '  A numeric value which is assumed to be a CSIDL. see CSIDLenum for a list
    '    example for CSIDL_COMMON_PROGRAMS (value=23), simply pass: CSIDLenum.CSIDL_COMMON_PROGRAMS or just 23
    '  String GUID, i.e. control panels GUID: ::{26EE0668-A00A-44D7-9371-BEB064C98683}
    '       control panel GUIDs must be prefixed with double colons else may not be usable
    '    another example using class function, pass: GetKnownFolderGUID(eFOLDERID_AppDataProgramData)
    '  IShellItem as an TLB defined item or IUnknown
    ' you are responsible for destroying any passed PIDLs, any time after this routine returns
    
    ' Notes you should be aware of:
    ' 1. These are added to the navigation pane
    ' 2. They will be under a grouping name that depends on the following
    '   When in the IDE, the name is: "Visual Basic"
    '   When compiled then depends on App.FileDescription (project properties, Make tab, File Description)
    '       App.FileDescription is not blank: App.FileDescription value
    '       App.FileDescription is blank: "Application Links"
    
    If m_StateFlags < 0& Then Exit Function     ' Vista+ dialog running
    If UBound(Folders) = -1& Then Exit Function
    
    If Me.Version > dvXP_Win2K Then
        Dim n As Long, lAttr As Long, oItem As stdole.IUnknown
        Dim sValue As String, lPtr As Long
        
        Set m_AddPlaces = New Collection
        For n = 0& To UBound(Folders)
            If IsMissing(Folders(n)) = False Then
                If VarType(Folders(n)) = vbDataObject Then
                    Set oItem = Folders(n)
                ElseIf VarType(Folders(n)) = vbString Then
                    If Left$(Folders(n), 3) = "::{" Then
                        Set oItem = pvGetIShellItem(Folders(n))
                    ElseIf Left$(Folders(n), 1) = "{" Then
                        If Me.Version < dvXP_Win2K Then ' gotta assume control panel folder
                            Set oItem = pvGetIShellItem("::" & Folders(n))
                        ElseIf IIDFromString(StrPtr(CStr(Folders(n))), VarPtr(m_Guid(0))) = 0& Then
                            If SHGetKnownFolderPath(VarPtr(m_Guid(0)), 0&, 0&, lPtr) = 0& Then
                                Set oItem = pvGetIShellItem(pvComStrPtrToVBstring(lPtr, True))
                            Else ' assume control panel GUID
                                Set oItem = pvGetIShellItem("::" & Folders(n))
                            End If
                        End If
                    End If
                ElseIf IsNumeric(Folders(n)) Then
                    On Error Resume Next
                    lPtr = CLng(Folders(n))
                    If Err Then
                        Err.Clear
                    ElseIf (lPtr And &HFF) = lPtr Then  ' CSIDL
                        sValue = String$(MAX_PATH, vbNullChar)
                        SHGetFolderPath 0&, lPtr Or CSIDL_FLAG_NO_ALIAS, 0&, 0&, StrPtr(sValue)
                        lPtr = InStr(sValue, vbNullChar)
                        If n > 1 Then sValue = Left$(sValue, lPtr - 1)
                        Set oItem = pvGetIShellItem(sValue): sValue = vbNullString
                    Else                            ' PIDL
                        Set oItem = pvGetIShellItem(Folders(n))
                    End If
                    On Error GoTo 0
                End If
                If Not oItem Is Nothing Then
                    pvCallInterface ObjPtr(oItem), ifShellItem_GetAttributes, siAttr_FOLDER, VarPtr(lAttr)
                    If lAttr Then m_AddPlaces.Add oItem
                    Set oItem = Nothing
                End If
            End If
        Next
        If m_AddPlaces.Count = 0& Then
            Set m_AddPlaces = Nothing
        Else
            m_StateFlags = (m_StateFlags And &HFFFFFDFF) Or Abs(AtPanelTop) * &H100&
        End If
    End If
    SetNavFolders = Not (m_AddPlaces Is Nothing)
    
End Function

Public Function SetNavFoldersEx(ByVal AtPanelTop As Boolean, ParamArray FoldersAndAliases() As Variant) As Boolean
    ' Vista+ only
    ' this is nearly identical to SetNavFolders except it allows aliasing/nicknaming the folders
    ' the paramarray is passed in comma-delimited pairs: Folder, Alias
    ' when displayed in the dialog's navigational pane, the Folder will be displayed as Alias
    ' Alias portion of the paramarray can be missing if no Alias is desired:
    ' Example shows aliasing the second folder, but not the first (has missing Alias member)
    '   SetNavFoldersEx True, "C:\Windows\Branding", , "C:\Windows\System32", "System Folder"
    
    If m_StateFlags < 0& Then Exit Function     ' Vista+ dialog running
    If UBound(FoldersAndAliases) = -1& Then Exit Function
    
    Dim tCol As Collection
    Dim n As Long, oItem As stdole.IUnknown
    Dim v As Variant, lAttr As Long
    Dim pProp(0 To 4) As Long ' faux PROPERTYKEY structure
    Const PKEY_ItemNameDisplay As String = "{B725F130-47EF-101A-A5F1-02608C9EEBAC}"
    
    If Me.Version > dvXP_Win2K Then
        Set tCol = New Collection
        IIDFromString StrPtr(PKEY_ItemNameDisplay), VarPtr(pProp(0)): pProp(4) = 10
        
        On Error Resume Next    ' in case user passed something bogus
        For n = 0& To UBound(FoldersAndAliases) - (UBound(FoldersAndAliases) And 1) Step 2
            If IsMissing(FoldersAndAliases(n)) = False Then
                If Me.SetNavFolders(AtPanelTop, FoldersAndAliases(n)) = True Then
                    Set oItem = m_AddPlaces.Item(1)
                    If n < UBound(FoldersAndAliases) Then
                        If IsMissing(FoldersAndAliases(n + 1&)) = False Then
                            v = CStr(FoldersAndAliases(n + 1&))
                            If Err Then
                                Err.Clear
                            Else
                                SHSetTemporaryPropertyForItem ObjPtr(oItem), VarPtr(pProp(0)), VarPtr(v)
                            End If
                        End If
                    End If
                    tCol.Add oItem: Set oItem = Nothing
                End If
            End If
        Next
        If Err Then Err.Clear
        On Error GoTo 0
        If tCol.Count = 0& Then
            Set m_AddPlaces = Nothing
        Else
            Set m_AddPlaces = tCol: Set tCol = Nothing
        End If
    End If
    SetNavFoldersEx = Not (m_AddPlaces Is Nothing)
    
End Function

Public Function GetKnownFolderGUID(ByVal FolderID As KnownFolderIDenum) As String

    ' courtesy function to return known Vista+ folder GUIDs
    ' current as of Aug 5th, 2018
    ' ref: https://docs.microsoft.com/en-us/windows/desktop/shell/knownfolderid

    Select Case FolderID
        Case eFOLDERID_AccountPictures: GetKnownFolderGUID = "{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}"
        Case eFOLDERID_AddNewPrograms: GetKnownFolderGUID = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
        Case eFOLDERID_AdminTools: GetKnownFolderGUID = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
        Case eFOLDERID_AppDataDesktop: GetKnownFolderGUID = "{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}"
        Case eFOLDERID_AppDataDocuments: GetKnownFolderGUID = "{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}"
        Case eFOLDERID_AppDataFavorites: GetKnownFolderGUID = "{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}"
        Case eFOLDERID_AppDataProgramData: GetKnownFolderGUID = "{559D40A3-A036-40FA-AF61-84CB430A4D34}"
        Case eFOLDERID_ApplicationShortcuts: GetKnownFolderGUID = "{A3918781-E5F2-4890-B3D9-A7E54332328C}"
        Case eFOLDERID_AppsFolder: GetKnownFolderGUID = "{1e87508d-89c2-42f0-8a7e-645a0f50ca58}"
        Case eFOLDERID_AppUpdates: GetKnownFolderGUID = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
        Case eFOLDERID_CameraRoll: GetKnownFolderGUID = "{AB5FB87B-7CE2-4F83-915D-550846C9537B}"
        Case eFOLDERID_CDBurning: GetKnownFolderGUID = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
        Case eFOLDERID_ChangeRemovePrograms: GetKnownFolderGUID = "{df7266ac-9274-4867-8d55-3bd661de872d}"
        Case eFOLDERID_CommonAdminTools: GetKnownFolderGUID = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
        Case eFOLDERID_CommonOEMLinks: GetKnownFolderGUID = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
        Case eFOLDERID_CommonPrograms: GetKnownFolderGUID = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
        Case eFOLDERID_CommonStartMenu: GetKnownFolderGUID = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
        Case eFOLDERID_CommonStartup: GetKnownFolderGUID = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
        Case eFOLDERID_CommonTemplates: GetKnownFolderGUID = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
        Case eFOLDERID_ComputerFolder: GetKnownFolderGUID = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
        Case eFOLDERID_ConflictFolder: GetKnownFolderGUID = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
        Case eFOLDERID_ConnectionsFolder: GetKnownFolderGUID = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
        Case eFOLDERID_Contacts: GetKnownFolderGUID = "{56784854-C6CB-462b-8169-88E350ACB882}"
        Case eFOLDERID_ControlPanelFolder: GetKnownFolderGUID = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
        Case eFOLDERID_Cookies: GetKnownFolderGUID = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
        Case eFOLDERID_Desktop: GetKnownFolderGUID = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
        Case eFOLDERID_DeviceMetadataStore: GetKnownFolderGUID = "{5CE4A5E9-E4EB-479D-B89F-130C02886155}"
        Case eFOLDERID_Documents: GetKnownFolderGUID = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
        Case eFOLDERID_DocumentsLibrary: GetKnownFolderGUID = "{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}"
        Case eFOLDERID_Downloads: GetKnownFolderGUID = "{374DE290-123F-4565-9164-39C4925E467B}"
        Case eFOLDERID_Favorites: GetKnownFolderGUID = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
        Case eFOLDERID_Fonts: GetKnownFolderGUID = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
        Case eFOLDERID_Games: GetKnownFolderGUID = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
        Case eFOLDERID_GameTasks: GetKnownFolderGUID = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
        Case eFOLDERID_History: GetKnownFolderGUID = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
        Case eFOLDERID_HomeGroup: GetKnownFolderGUID = "{52528A6B-B9E3-4ADD-B60D-588C2DBA842D}"
        Case eFOLDERID_HomeGroupCurrentUser: GetKnownFolderGUID = "{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}"
        Case eFOLDERID_ImplicitAppShortcuts: GetKnownFolderGUID = "{BCB5256F-79F6-4CEE-B725-DC34E402FD46}"
        Case eFOLDERID_InternetCache: GetKnownFolderGUID = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
        Case eFOLDERID_InternetFolder: GetKnownFolderGUID = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
        Case eFOLDERID_Libraries: GetKnownFolderGUID = "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}"
        Case eFOLDERID_Links: GetKnownFolderGUID = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
        Case eFOLDERID_LocalAppData: GetKnownFolderGUID = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
        Case eFOLDERID_LocalAppDataLow: GetKnownFolderGUID = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
        Case eFOLDERID_LocalizedResourcesDir: GetKnownFolderGUID = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
        Case eFOLDERID_Music: GetKnownFolderGUID = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
        Case eFOLDERID_MusicLibrary: GetKnownFolderGUID = "{2112AB0A-C86A-4FFE-A368-0DE96E47012E}"
        Case eFOLDERID_NetHood: GetKnownFolderGUID = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
        Case eFOLDERID_NetworkFolder: GetKnownFolderGUID = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
        Case eFOLDERID_Objects3D: GetKnownFolderGUID = "{31C0DD25-9439-4F12-BF41-7FF4EDA38722}"
        Case eFOLDERID_OriginalImages: GetKnownFolderGUID = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
        Case eFOLDERID_PhotoAlbums: GetKnownFolderGUID = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
        Case eFOLDERID_Pictures: GetKnownFolderGUID = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
        Case eFOLDERID_PicturesLibrary: GetKnownFolderGUID = "{A990AE9F-A03B-4E80-94BC-9912D7504104}"
        Case eFOLDERID_Playlists: GetKnownFolderGUID = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
        Case eFOLDERID_PrintersFolder: GetKnownFolderGUID = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
        Case eFOLDERID_PrintHood: GetKnownFolderGUID = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
        Case eFOLDERID_Profile: GetKnownFolderGUID = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
        Case eFOLDERID_ProgramData: GetKnownFolderGUID = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
        Case eFOLDERID_ProgramFiles: GetKnownFolderGUID = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
        Case eFOLDERID_ProgramFilesCommon: GetKnownFolderGUID = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
        Case eFOLDERID_ProgramFilesCommonX64: GetKnownFolderGUID = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
        Case eFOLDERID_ProgramFilesCommonX86: GetKnownFolderGUID = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
        Case eFOLDERID_ProgramFilesX64: GetKnownFolderGUID = "{6D809377-6AF0-444b-8957-A3773F02200E}"
        Case eFOLDERID_ProgramFilesX86: GetKnownFolderGUID = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
        Case eFOLDERID_Programs: GetKnownFolderGUID = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
        Case eFOLDERID_Public: GetKnownFolderGUID = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
        Case eFOLDERID_PublicDesktop: GetKnownFolderGUID = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
        Case eFOLDERID_PublicDocuments: GetKnownFolderGUID = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
        Case eFOLDERID_PublicDownloads: GetKnownFolderGUID = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
        Case eFOLDERID_PublicGameTasks: GetKnownFolderGUID = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
        Case eFOLDERID_PublicLibraries: GetKnownFolderGUID = "{48DAF80B-E6CF-4F4E-B800-0E69D84EE384}"
        Case eFOLDERID_PublicMusic: GetKnownFolderGUID = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
        Case eFOLDERID_PublicPictures: GetKnownFolderGUID = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
        Case eFOLDERID_PublicRingtones: GetKnownFolderGUID = "{E555AB60-153B-4D17-9F04-A5FE99FC15EC}"
        Case eFOLDERID_PublicUserTiles: GetKnownFolderGUID = "{0482af6c-08f1-4c34-8c90-e17ec98b1e17}"
        Case eFOLDERID_PublicVideos: GetKnownFolderGUID = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
        Case eFOLDERID_QuickLaunch: GetKnownFolderGUID = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
        Case eFOLDERID_Recent: GetKnownFolderGUID = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
        Case eFOLDERID_RecordedTV: GetKnownFolderGUID = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
        Case eFOLDERID_RecordedTVLibrary: GetKnownFolderGUID = "{1A6FDBA2-F42D-4358-A798-B74D745926C5}"
        Case eFOLDERID_RecycleBinFolder: GetKnownFolderGUID = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
        Case eFOLDERID_ResourceDir: GetKnownFolderGUID = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
        Case eFOLDERID_Ringtones: GetKnownFolderGUID = "{C870044B-F49E-4126-A9C3-B52A1FF411E8}"
        Case eFOLDERID_RoamedTileImages: GetKnownFolderGUID = "{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}"
        Case eFOLDERID_RoamingAppData: GetKnownFolderGUID = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
        Case eFOLDERID_RoamingTiles: GetKnownFolderGUID = "{00BCFC5A-ED94-4e48-96A1-3F6217F21990}"
        Case eFOLDERID_SampleMusic: GetKnownFolderGUID = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
        Case eFOLDERID_SamplePictures: GetKnownFolderGUID = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
        Case eFOLDERID_SamplePlaylists: GetKnownFolderGUID = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
        Case eFOLDERID_SampleVideos: GetKnownFolderGUID = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
        Case eFOLDERID_SavedGames: GetKnownFolderGUID = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
        Case eFOLDERID_SavedPictures: GetKnownFolderGUID = "{3B193882-D3AD-4eab-965A-69829D1FB59F}"
        Case eFOLDERID_SavedPicturesLibrary: GetKnownFolderGUID = "{E25B5812-BE88-4bd9-94B0-29233477B6C3}"
        Case eFOLDERID_SavedSearches: GetKnownFolderGUID = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
        Case eFOLDERID_Screenshots: GetKnownFolderGUID = "{b7bede81-df94-4682-a7d8-57a52620b86f}"
        Case eFOLDERID_SEARCH_CSC: GetKnownFolderGUID = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
        Case eFOLDERID_SEARCH_MAPI: GetKnownFolderGUID = "{98ec0e18-2098-4d44-8644-66979315a281}"
        Case eFOLDERID_SearchHistory: GetKnownFolderGUID = "{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}"
        Case eFOLDERID_SearchHome: GetKnownFolderGUID = "{190337d1-b8ca-4121-a639-6d472d16972a}"
        Case eFOLDERID_SearchTemplates: GetKnownFolderGUID = "{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}"
        Case eFOLDERID_SendTo: GetKnownFolderGUID = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
        Case eFOLDERID_SidebarDefaultParts: GetKnownFolderGUID = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
        Case eFOLDERID_SidebarParts: GetKnownFolderGUID = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
        Case eFOLDERID_SkyDrive: GetKnownFolderGUID = "{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
        Case eFOLDERID_SkyDriveCameraRoll: GetKnownFolderGUID = "{767E6811-49CB-4273-87C2-20F355E1085B}"
        Case eFOLDERID_SkyDriveDocuments: GetKnownFolderGUID = "{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}"
        Case eFOLDERID_SkyDrivePictures: GetKnownFolderGUID = "{339719B5-8C47-4894-94C2-D8F77ADD44A6}"
        Case eFOLDERID_StartMenu: GetKnownFolderGUID = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
        Case eFOLDERID_Startup: GetKnownFolderGUID = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
        Case eFOLDERID_SyncManagerFolder: GetKnownFolderGUID = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
        Case eFOLDERID_SyncResultsFolder: GetKnownFolderGUID = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
        Case eFOLDERID_SyncSetupFolder: GetKnownFolderGUID = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
        Case eFOLDERID_System: GetKnownFolderGUID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
        Case eFOLDERID_SystemX86: GetKnownFolderGUID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
        Case eFOLDERID_Templates: GetKnownFolderGUID = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
        Case eFOLDERID_TreeProperties: GetKnownFolderGUID = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
        Case eFOLDERID_UserPinned: GetKnownFolderGUID = "{9E3995AB-1F9C-4F13-B827-48B24B6C7174}"
        Case eFOLDERID_UserProfiles: GetKnownFolderGUID = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
        Case eFOLDERID_UserProgramFiles: GetKnownFolderGUID = "{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}"
        Case eFOLDERID_UserProgramFilesCommon: GetKnownFolderGUID = "{BCBD3057-CA5C-4622-B42D-BC56DB0AE516}"
        Case eFOLDERID_UsersFiles: GetKnownFolderGUID = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
        Case eFOLDERID_UsersLibraries: GetKnownFolderGUID = "{A302545D-DEFF-464b-ABE8-61C8648D939B}"
        Case eFOLDERID_Videos: GetKnownFolderGUID = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
        Case eFOLDERID_VideosLibrary: GetKnownFolderGUID = "{491E922F-5643-4AF4-A7EB-4E7A138D8174}"
        Case eFOLDERID_Windows: GetKnownFolderGUID = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
    End Select

End Function

Private Sub pvGetResults(hResult As Long, ptrIShellReturn As Long, bOpenMode As Boolean)

    ' routine process the selected file(s) when dialog closed

    Dim oReturn As stdole.IUnknown, oItem As stdole.IUnknown
    Dim sValue As String, sFile As String, n As Long
    
    If m_EventsCookie Then  ' free events cookie
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_Unadvise, m_EventsCookie
        m_EventsCookie = 0&
    End If
    
    If bOpenMode Then
        If Not m_BasketItems Is Nothing Then
            If m_BasketItems.Count = 0& Then
                hResult = CDERR_CANCELED
            ElseIf hResult = 0& Then
                Set m_ModeResults = pvUpdateBasket(-1&)
            End If
            n = pvKeyToIndex(SIGN_BIT + 1&, False)
            If n > -1& Then ' reset count in case dialog called again without Clearing
                m_Controls(n).Flags = m_Controls(n).Flags And &H80000FFF
            End If
        End If
        Set oReturn = m_ModeResults
    ElseIf Not m_ModeResults Is Nothing Then    ' get 1st array item (count should never be > 1)
        ' note to self: m_ModeResults would never be populated, as of this date, when ShowSave is called
        '   m_ModeResults used only for custom mode, which is not applicable with Save dialog
        ' leaving code as-is, should I decide to change this in the future
        pvCallInterface ObjPtr(m_ModeResults), ifShellItemArray_GetItemAt, 0&, VarPtr(oReturn)
    End If
    Set m_ModeResults = Nothing
    
    If hResult = 0& Then            ' else dialog was canceled/failed to load
    
        ' update which filter index was selected by user
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetFileTypeIndex, VarPtr(ofn.nFilterIndex)
        ofn.nFileOffset = 0&
        
        ' process selected file(s)
        If ((ofn.Flags And DLG_AllowMultiSelect) Or Not (oReturn Is Nothing)) And bOpenMode = True Then
            If oReturn Is Nothing Then
                If pvCallInterface(ObjPtr(m_Dialog), ifFileDialogOpen_GetResults, VarPtr(oReturn)) Then GoTo ExitRoutine
            End If
            If pvCallInterface(ObjPtr(oReturn), ifShellItemArray_GetCount, VarPtr(n)) = 0& Then
                ' if only 1 item was selected, use IShellItem vs IShellItemArray
                If n = 1& Then
                    pvCallInterface ObjPtr(oReturn), ifShellItemArray_GetItemAt, 0&, VarPtr(oItem)
                    If Not oItem Is Nothing Then Set oReturn = oItem
                Else    ' walk thru each selected file and build ofn.lpstrFile and/or passed collection
                    ' multi-select items are listed as complete path/filenames separated by vbNullChar
                    If Not ptrIShellReturn = 0& Then
                        IIDFromString StrPtr(IID_IShellItemArray), VarPtr(m_Guid(0))
                        pvCallInterface ObjPtr(oReturn), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), ptrIShellReturn
                    End If
                    For n = 0& To n - 1&
                        If pvCallInterface(ObjPtr(oReturn), ifShellItemArray_GetItemAt, n, VarPtr(oItem)) = 0& Then
                            sValue = pvGetDisplayName(ObjPtr(oItem), dfn_FileSysPath)
                            If n = 0& Then
                                ofn.lpstrFile = sValue: ofn.nFileOffset = InStrRev(sValue, "\")
                                If ofn.nFileOffset = 0 Then ofn.nFileOffset = 1
                            Else
                                ofn.lpstrFile = ofn.lpstrFile & vbNullChar & sValue
                            End If
                            Set oItem = Nothing
                        End If
                    Next
                    Set oReturn = Nothing
                End If
            Else        ' no items selected
                Set oReturn = Nothing
            End If
            
        Else
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetResult, VarPtr(oReturn)
        End If
        If Not oReturn Is Nothing Then     ' IShellItem pointer
            ofn.lpstrFile = pvGetDisplayName(ObjPtr(oReturn), dfn_FileSysPath)
            ofn.nFileOffset = InStrRev(ofn.lpstrFile, "\")
            If ofn.nFileOffset = 0 Then ofn.nFileOffset = 1
            If Not ptrIShellReturn = 0& Then
                IIDFromString StrPtr(IID_IShellItem), VarPtr(m_Guid(0))
                pvCallInterface ObjPtr(oReturn), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), ptrIShellReturn
            End If
        End If
    Else
        ofn.lpstrFile = vbNullString
        ofn.nFileOffset = 0
    End If
    
ExitRoutine:
    pvSetGetValues False, CBool(hResult)     ' get values of added controls
    Set m_Dlg2 = Nothing
    Set m_DlgMod = Nothing
    Set m_Dialog = Nothing
End Sub

Private Sub pvGetTextBoxValue()
    
    ' when a textbox control has been added to the dialog, its content must
    '   be read before the dialog closes. If user wants this value, then
    '   dialog events must be requested. This is called from within the
    '   dialog events just before the dialog closes or from the CloseDialog method
    Dim n As Long
    If Not m_DlgMod Is Nothing Then
        For n = 1& To UBound(m_Controls)
            If (m_Controls(n).Flags And &HF) = ctlType_TextBox Then
                Me.Controls_PropertyGet -m_Controls(n).ID, ppText
            End If
        Next
    End If
    
End Sub

Private Function pvGetDisplayName(ByVal oIShellItem As Long, ByVal DisplayFormat As DisplayFormatEnum, _
                                    Optional Attrs As Long) As String
    
    ' returns the path/file name of passed IShellItem
    ' BHID guids
    ' http://read.pudn.com/downloads84/sourcecode/delphi_control/323829/JwaShlGuid.pas__.htm
    
    Dim aGuid(0 To 7) As Long       ' faux GUID structures
    Dim afmtEtcDF(0 To 4) As Long   ' faux FORMATETC & DROPFILES structure
    Dim fMedium(0 To 2) As Long     ' faux STGMEDIUM structure
    Dim lPtr As Long, lSize As Long
    Dim sFile As String, bANSI() As Byte
    Const IID_BHID_IDATAOBJECT As String = "{B8C0BD9F-ED24-455C-83E6-D5390C4FE8C4}"
    Const IID_IDataObject As String = "{0000010E-0000-0000-C000-000000000046}"
    Const ifIDataObject_GetData As Long = 3&
    
    pvCallInterface oIShellItem, ifShellItem_GetAttributes, -1&, VarPtr(Attrs)
    pvCallInterface oIShellItem, ifShellItem_GetDisplayName, DisplayFormat, VarPtr(lPtr)
    If lPtr Then
        sFile = pvComStrPtrToVBstring(lPtr, True)
        
    ElseIf (Attrs And siAttr_STREAM) = 0& And DisplayFormat = dfn_FileSysPath Then ' else streamable
        IIDFromString StrPtr(IID_BHID_IDATAOBJECT), VarPtr(aGuid(0))
        IIDFromString StrPtr(IID_IDataObject), VarPtr(aGuid(4))
        pvCallInterface oIShellItem, ifShellItem_BindToHandler, 0&, VarPtr(aGuid(0)), VarPtr(aGuid(4)), VarPtr(lPtr)
        If lPtr Then
            afmtEtcDF(0) = vbCFFiles: afmtEtcDF(2) = 1&: afmtEtcDF(3) = -1&: afmtEtcDF(4) = 1&
            pvCallInterface lPtr, ifIDataObject_GetData, VarPtr(afmtEtcDF(0)), VarPtr(fMedium(0))
            pvCallInterface lPtr, ifUnknown_Release: lPtr = 0&
            If fMedium(1) Then
                CopyMemory afmtEtcDF(0), ByVal fMedium(1), 20&
                lPtr = pvSafeOffset32(fMedium(1), afmtEtcDF(0))
                If afmtEtcDF(4) Then    ' unicode vs ANSI
                    sFile = pvComStrPtrToVBstring(lPtr, False)
                Else    ' shouldn't get here, but we'll process ANSI anyway
                    lSize = lstrlen(lPtr)
                    ReDim bANSI(0 To lSize - 1&)
                    CopyMemory ByVal bANSI(0), ByVal lPtr, lSize ' populate the buffer
                    sFile = StrConv(bANSI(), vbUnicode)
                End If
                ReleaseStgMedium VarPtr(fMedium(0))
                Attrs = Attrs Or siAttr_SYSTEM ' add FileSysPath as attribute
            End If
        End If
    End If
    If Len(sFile) = 0& Then
        If lPtr = 0& Then pvCallInterface oIShellItem, ifShellItem_GetDisplayName, dfn_DesktopAbsoluteParsing, VarPtr(lPtr)
        pvGetDisplayName = pvComStrPtrToVBstring(lPtr, True)
    Else
        pvGetDisplayName = sFile
    End If
    
End Function

Private Function pvComStrPtrToVBstring(pString As Long, Optional FreePointer As Boolean = True) As String
    ' IShellI[xxxx] interfaces pass string pointers that are to be destroyed by us
    Dim lLen As Long
    If Not pString = 0& Then
        lLen = lstrlenW(pString)
        If lLen > 0& Then
            pvComStrPtrToVBstring = String$(lLen, vbNullChar)
            CopyMemory ByVal StrPtr(pvComStrPtrToVBstring), ByVal pString, lLen * 2&
        End If
        If FreePointer Then CoTaskMemFree pString
    End If
End Function

Private Sub pvDisplayError(ByVal errCode As Long, bOpenMode As Integer)

    Dim sMode As String
    
    Select Case bOpenMode
        Case 0: sMode = "Open"
        Case 1: sMode = "Save"
        Case 2: sMode = "BrowseForFolder"
    End Select
    pzDestroyThunks
    ' release thunks, if any, since we will be aborting the class by raising an error
    ' don't want user to close the project in IDE, due to the error, and lose track of the thunk memory

    ' NOTE: If your VB is stopping in this routine, it is due to your error trapping settings
    ' Select Tools|Options from the menu above
    ' Select the "General Tab"
    ' Select option button: "Break on Unhandled Errors"
    Select Case errCode
    Case 0&, CDERR_CANCELED
        Err.Raise CDERR_CANCELED, "Common Dialog " & sMode, "Cancel was selected"
    Case CDERR_DIALOGFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "The dialog box could not be created"
    Case CDERR_FINDRESFAILURE, CDERR_LOADRESFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "The common dialog box function failed to find a specified resource."
    Case CDERR_INITIALIZATION
        Err.Raise errCode, "Common Dialog " & sMode, "The common dialog box function failed during initialization."
    Case CDERR_LOADSTRFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "The common dialog box function failed to load a specified string."
    Case CDERR_LOCKRESFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "The common dialog box function failed to lock a specified resource."
    Case CDERR_MEMALLOCFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "The common dialog box function was unable to allocate memory for internal structures."
    Case CDERR_MEMLOCKFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "The common dialog box function was unable to lock the memory associated with a handle."
    Case CDERR_NOHINSTANCE
        Err.Raise errCode, "Common Dialog " & sMode, "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
    Case CDERR_NOHOOK
        Err.Raise errCode, "Common Dialog " & sMode, "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
    Case CDERR_NOTEMPLATE
        Err.Raise errCode, "Common Dialog " & sMode, "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template."
    Case CDERR_REGISTERMSGFAIL
        Err.Raise errCode, "Common Dialog " & sMode, "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
    Case CDERR_STRUCTSIZE
        Err.Raise errCode, "Common Dialog " & sMode, "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
    Case FNERR_BUFFERTOOSMALL
        If Len(ofn.lpstrFile) > 4 Then
            CopyMemory m_minBufferSize, ByVal StrPtr(ofn.lpstrFile), 2&
        Else
            m_minBufferSize = -1
        End If
        Err.Raise errCode, "Common Dialog " & sMode, "The PathFileBufferSize is too small for the file name(s) specified by the user."
    Case FNERR_INVALIDFILENAME
        Err.Raise errCode, "Common Dialog " & sMode, "A file name is invalid."
    Case FNERR_SUBCLASSFAILURE
        Err.Raise errCode, "Common Dialog " & sMode, "An attempt to subclass a list box failed because sufficient memory was not available."
    Case Else
        Err.Raise errCode, "Common Dialog " & sMode, "Unknown error code of " & errCode
    End Select

End Sub

Private Function pvInitStructure(hWnd As Long, bOpenMode As Boolean, _
                                bEvents As EventTypeEnum, ClientGUID As String) As Boolean

    ' Dialog intialization for file dialogs
    If obif.lpfnCallback Then pzDestroyThunks   ' release old-style browse for folder callback

    Dim oPtr As Long, oItem As stdole.IUnknown
    Dim lFlags As Long, lEvents As Long, sCaption As String
    Dim sItems() As String, aGuid(0 To 7) As Long, n As Long
    Const ID_SelectFolder As Long = 439     ' String ID in comdlg32.dll for "Select Folder" caption
    Const ID_AllFiles As Long = 34193       ' String ID in shell32.dll for "All Files" caption
    Const CLSID_FileSaveDialog As String = "{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}"
    Const IID_IFileSaveDialog As String = "{84bccd23-5fde-4cdb-aea4-af64b83d78ab}"
    
    m_StateFlags = (m_StateFlags And &HFFF0FFFF) Or (bEvents And &H3&) * &H10000
    m_ClientGUID = vbNullString
    
    If Me.Version = dvXP_Win2K Then         ' Pre-Vista
        ofn.hwndOwner = hWnd                ' set hWnd; disables that window til dialog closed
        ofn.lStructSize = Len(ofn)
        If ofn.Flags = 0& Then
            If bOpenMode Then ofn.Flags = DLG__BaseOpenDialogFlags Else ofn.Flags = DLG__BaseSaveDialogFlags
        Else
            ' fix up common conflicts
            If bOpenMode Then
                ofn.Flags = ofn.Flags And Not DLG_OverwritePrompt
            Else
                ofn.Flags = ofn.Flags And Not DLG_AllowMultiSelect
            End If
        End If                              ' remove Vista+ related flags
        ofn.Flags = ofn.Flags Or DLG_Explorer Or DLG_EnableSizing
        
        m_minBufferSize = 0&                ' reset before each call
        ' ensure the FileName property is minimally sized
        n = Me.MaxFileSize
        If n <= Len(ofn.lpstrFile) Then
            ofn.nMaxFile = Len(ofn.lpstrFile)
        Else
            ofn.lpstrFile = ofn.lpstrFile & String$(n - Len(ofn.lpstrFile), vbNullChar)
            If ofn.nMaxFile = 0& Then ofn.nMaxFile = n
        End If
        
        If Len(ofn.lpstrFilter) = 0& Then
            n = 0&: ofn.lpstrFilter = pvLoadString(n, "shell32.dll", ID_AllFiles) & vbNullChar & "*.*" & vbNullChar
            If n Then FreeLibrary n
            If Len(ofn.lpstrFilter) = 5 Then _
                ofn.lpstrFilter = "All Files" & vbNullChar & "*.*" & vbNullChar
        End If
        If bEvents Or (ofn.Flags And DLG_AllowMultiSelect) Then
            pzCreateThunks emskAllEvents
            ofn.Flags = ofn.Flags Or DLGex_PickFolders ' (also=OFN_ENABLEHOOK)
            ofn.lpfnHook = m_SubclassFnc
        Else
            ofn.lpfnHook = 0&
            m_StateFlags = m_StateFlags Or &H40000000 ' flag indicating dialog activated else set in callback
        End If
        If bEvents = 0& Then ofn.Flags = ofn.Flags And Not DLGXPonly_EnableIncludeNotify ' n/a if not asking for events
        pvInitStructure = True
        Exit Function
    End If
    
    If bOpenMode Then                   ' create the dialog
        CLSIDFromString StrPtr(CLSID_FileOpenDialog), VarPtr(aGuid(0))
        IIDFromString StrPtr(IID_IFileOpenDialog), VarPtr(aGuid(4))
    Else
        CLSIDFromString StrPtr(CLSID_FileSaveDialog), VarPtr(aGuid(0))
        IIDFromString StrPtr(IID_IFileSaveDialog), VarPtr(aGuid(4))
        m_StateFlags = m_StateFlags And Not emskCustMode
        Set m_BasketItems = Nothing
    End If
    If CoCreateInstance(VarPtr(aGuid(0)), 0&, CLSCTX_INPROC_SERVER, VarPtr(aGuid(4)), m_Dialog) Then Exit Function
    
    ' assign client GUID
    If Not ClientGUID = m_ClientGUID Then
        If Len(ClientGUID) = 36 Then
            m_ClientGUID = "{" & ClientGUID & "}"
        ElseIf Len(ClientGUID) = 38 Then
            If Left$(ClientGUID, 1) = "{" And Right$(ClientGUID, 1) = "}" Then m_ClientGUID = vbNullString
        End If
    End If
    If Len(m_ClientGUID) Then
        If IIDFromString(StrPtr(m_ClientGUID), VarPtr(aGuid(0))) = 0& Then _
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetClientGUID, VarPtr(aGuid(0))
    End If
    
    If ofn.Flags Then
        ' fix up common conflicts
        If bOpenMode Then
            lFlags = ofn.Flags And Not DLG_OverwritePrompt
        Else
            lFlags = ofn.Flags And Not (DLGex_PickFolders Or DLG_AllowMultiSelect)
        End If
        If (lFlags And DLGex_AllNonStorageItems) Then lFlags = lFlags And Not DLGex_ForceFileSystem
        lFlags = lFlags And Not (DLG_EnableSizing Or DLG_Explorer)
    Else
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetOptions, VarPtr(lFlags)
    End If
    lFlags = lFlags Or ofn.FlagsEx              ' add user-opted forced flags
    
    lEvents = (m_StateFlags And emskAllEvents)
    
    If (m_StateFlags And emskCustMode) > 0& Then    ' tweak flags depending on type of custom mode
        lFlags = lFlags And Not DLGex_PickFolders   ' not applicable except for 1 custom mode
        If (m_StateFlags And emskCustMode) = cm_CompressedFolderPlusFiles Then
            lFlags = (lFlags Or DLGex_AllNonStorageItems) And Not DLGex_ForceFileSystem
        ElseIf (m_StateFlags And emskCustMode) = cm_BasketModeFoldersOnly Then
            lFlags = lFlags Or DLGex_PickFolders
        End If
        If m_BasketItems Is Nothing Then
            lEvents = lEvents Or evtDialogEvents
        Else
            lEvents = lEvents Or emskAllEvents
        End If
    End If
 
     ' apply the Flags
    pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetOptions, lFlags

    ' apply the filter if any, default to All Files
    If Not (ofn.lpstrFilter = vbNullChar And bOpenMode = True) Then
        If Len(ofn.lpstrFilter) = 0& Or ofn.lpstrFilter = vbNullChar Then
            n = 0&: sItems() = Split(pvLoadString(n, "shell32.dll", ID_AllFiles) & " (*.*)" & vbNullChar & "*.*", vbNullChar)
            If n Then FreeLibrary n
            If sItems(0) = " (*.*)" Then _
                sItems() = Split("All Files (*.*)" & vbNullChar & "*.*", vbNullChar)
        Else
            n = InStrRev(ofn.lpstrFilter, vbNullChar) ' exclude trailing null
            sItems() = Split(Left$(ofn.lpstrFilter, n - 1&), vbNullChar)
        End If
        If UBound(sItems) > 0& Then
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileTypes, CLng((UBound(sItems) + 1) \ 2&), VarPtr(sItems(0))
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileTypeIndex, Me.FilterIndex
        End If
        Erase sItems()
    End If
    
    ' apply option navigational pane additions
    If Not m_AddPlaces Is Nothing Then
        IIDFromString StrPtr(IID_IShellItem), VarPtr(m_Guid(0))
        For n = 1& To m_AddPlaces.Count
            Set oItem = m_AddPlaces.Item(n)
            If Not oItem Is Nothing Then
                oPtr = 0&: pvCallInterface ObjPtr(oItem), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(oPtr)
                If oPtr Then
                    pvCallInterface ObjPtr(m_Dialog), ifFileDialog_AddPlace, oPtr, (m_StateFlags And &H100&) \ &H100&
                    pvCallInterface oPtr, ifUnknown_Release
                End If
                Set oItem = Nothing
            End If
        Next
    End If
    
    ' apply the FileBox entry
    If Len(ofn.lpstrFile) Then
        ' treat a full path/filename as both: init dir + select file
        n = InStrRev(ofn.lpstrFile, "\")
        If n Then Set oItem = pvGetIShellItem(Left$(ofn.lpstrFile, n))
        If oItem Is Nothing Then
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileName, StrPtr(ofn.lpstrFile)
        Else
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileName, StrPtr(Mid$(ofn.lpstrFile, n + 1&))
        End If
    End If
    
    ' apply the startup folder
    If oItem Is Nothing And Len(Me.InitDir) > 0& Then Set oItem = pvGetIShellItem(Me.InitDir)
    If Not oItem Is Nothing Then
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFolder, ObjPtr(oItem)
        Set oItem = Nothing
    End If
    
    ' apply dialog title
    If Len(Me.DialogTitle) = 0& Then
        ' if selecting only folders and no dialog title provided, default to standard: Select Folder
        If (m_StateFlags And emskCustMode) = cm_BrowseFoldersShowFiles Or (m_StateFlags And emskCustMode) = cm_BasketModeFoldersOnly Then
            n = 0&: sCaption = pvLoadString(n, "comdlg32.dll", ID_SelectFolder)
            If n Then FreeLibrary n
            If Len(sCaption) Then _
                pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetTitle, StrPtr(sCaption)
        End If
    Else
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetTitle, StrPtr(Me.DialogTitle)
    End If
    
    ' apply default extension
    If Len(Me.DefaultExt) Then _
        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetDefaultExtension, StrPtr(Me.DefaultExt)
    
    ' add optional root folder & create a IFileDialog2 interface as needed
    If Not IsEmpty(m_NavRestrictor) Then
        Set oItem = pvGetIShellItem(m_NavRestrictor)
        pvFreeOrCreateRootPIDL False, True
    End If
                
    If (Not oItem Is Nothing) Or pvKeyToIndex("cnx", False) > 0& Then
        If m_Dlg2 Is Nothing Then
            If Me.Version > dvVista Then
                IIDFromString StrPtr(IID_IFileDialog2), VarPtr(aGuid(0))
                pvCallInterface ObjPtr(m_Dialog), ifUnknown_QueryInterface, VarPtr(aGuid(0)), VarPtr(m_Dlg2)
                If (Not m_Dlg2 Is Nothing) And (Not oItem Is Nothing) Then
                    pvCallInterface ObjPtr(m_Dlg2), ifFileDialog2_SetNavigationRoot, ObjPtr(oItem)
                    Set m_NavRestrictor = oItem
                End If
            End If
        End If
        Set oItem = Nothing
    End If
    
    ' finally, add any controls and/or change captions on existing controls
    If pvSetGetValues(True, False) = True Then
        If lEvents Then
            pzCreateThunks lEvents
            pvCallInterface ObjPtr(m_Dialog), ifFileDialog_Advise, m_Events(0).VTable, VarPtr(m_EventsCookie)
        Else
            m_StateFlags = m_StateFlags Or SIGN_BIT ' flag = dialog being shown. Else set within events
        End If
        pvInitStructure = True
    Else
        Set m_BasketItems = Nothing
    End If
    
End Function

Private Function pvSetGetValues(bCreate As Boolean, dlgCancel As Boolean) As Boolean

    ' sets control values before dialog displayed and retrieves them after it closes
    ' Containerized controls are created when the container is processed
    '   each containerized control is created in order it was added to the container
    ' Other controls are in order they were created via Controls_Add

    Dim itemNr As Long, nrItems As Long
    Dim lIndex As Long, c As Long, t As Long
    Dim bIsCnt As Boolean, bProminent As Boolean
    Const IID_IFileDialogCustomize As String = "{8016b7b3-3d49-4504-a0aa-2a37494e606f}"
    
    If m_ControlKeys Is Nothing Then
        pvSetGetValues = True
        Exit Function
    ElseIf bCreate = False Then ' retreving control properties, unless canceled or no controls added
        If Not (dlgCancel = True Or m_DlgMod Is Nothing) Then
            lIndex = -pvKeyToIndex("ro", False)
            For itemNr = 1& To UBound(m_Controls)
                If (m_Controls(itemNr).Flags And &HF) > 0& Then
                    Me.Controls_PropertyGet -m_Controls(itemNr).ID, -1&
                    If lIndex < 0& Then          ' for checkbox read-only option, update flags
                        With m_Controls(itemNr)
                            If .ID = lIndex Then
                                ofn.Flags = (ofn.Flags And &HFFFFFFFE) Or ((.Flags And &H100&) \ &H100&)
                                lIndex = 0&          ' no need to keep checking now
                            End If
                        End With
                    End If
                End If
            Next
        End If
        pvSetGetValues = True
        Exit Function
    End If
        
    ' create the dialog customizer interface
    IIDFromString StrPtr(IID_IFileDialogCustomize), VarPtr(m_Guid(0))
    If pvCallInterface(ObjPtr(m_Dialog), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(m_DlgMod)) Then
        Set m_ControlKeys = Nothing: Erase m_Controls()         ' aborting
        m_StateFlags = m_StateFlags And Not emskAllEvents
        Exit Function ' should never happen; sanity check only
    End If
    
    For itemNr = 0& To UBound(m_Controls)
        bIsCnt = ((m_Controls(itemNr).Flags And &HF) = ctlType_Container)
        nrItems = 0&
        If bIsCnt Then
            With m_Controls(itemNr)
                If (.Flags And &H7FFF00) = 0& Then ' number of contained controls
                    nrItems = -1&
                ElseIf (.ID And &H70000000) = 0& Then  ' custom mode container
                    If (m_StateFlags And emskCustMode) = 0& Then nrItems = -1&
                End If
                If nrItems = 0& Then                ' create the container & update its ID
                    If (.ID And &H70000000) = 0& Then
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_StartVisualGroup, .ID, StrPtr(Replace$(.Caption, "%s", "0000"))
                    Else
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_StartVisualGroup, .ID, StrPtr(.Caption)
                    End If
                    nrItems = UBound(.SubItems())
                Else
                    bIsCnt = False                  ' not a container, not a control, won't be added
                End If
            End With
        Else
            If (m_Controls(itemNr).ID And &H70000000) = 0& Then ' custom mode control
                If (m_StateFlags And emskCustMode) = 0& Then nrItems = -1&
            End If
            If nrItems = 0& Then
                ' if containerized control, skip it as they are processed from the container above
                If m_Controls(itemNr).Flags < 0& Then nrItems = -1& Else lIndex = itemNr
            End If
        End If
    
        For c = 0& To nrItems
            If bIsCnt Then lIndex = pvKeyToIndex(m_Controls(itemNr).SubItems(c), False)
            With m_Controls(lIndex)
                Select Case (.Flags And &HF)
                '//// Open/Save button
                Case ctlType_DefaultButton
                    If (.Flags And &H3000000) Then
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_EnableOpenDropDown, .ID
                        If (.Flags And &H2000000) Then  ' used as read-only option
                            t = ((.Flags And &H1000000) \ &H1000000)
                            pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddControlItem, .ID, (t Xor 1&), StrPtr(.SubItems(0))
                            pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddControlItem, .ID, t, StrPtr(.SubItems(1))
                            ofn.Flags = (ofn.Flags And &HFFFFFFFE) Or (t Xor 1&)
                        Else
                            For t = 0& To UBound(.SubItems)
                                pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddControlItem, .ID, t, StrPtr(.SubItems(t))
                            Next
                        End If
                    ElseIf Len(.Caption) Then
                        pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetOkButtonLabel, StrPtr(.Caption)
                    End If
                '//// Checkbox
                Case ctlType_CheckBox
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddCheckButton, .ID, StrPtr(.Caption), (.Flags And &HF00&) \ &H100&
                    If pvKeyToIndex("ro", False) = -.ID Then ofn.Flags = (ofn.Flags And &HFFFFFFFE) Or (.Flags And &H100&) \ &H100&
                '//// ComboBox, OptionButtons, Menu
                Case ctlType_ComboBox, ctlType_OptionBtnGroup, ctlType_Menu ' controls with subitems
                    If (.Flags And &HF) = ctlType_ComboBox Then
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddComboBox, .ID
                    ElseIf (.Flags And &HF) = ctlType_OptionBtnGroup Then
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddRadioButtonList, .ID
                    Else
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddMenu, .ID, StrPtr(.Caption)
                    End If
                    For t = 0& To UBound(.SubItems)
                        pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddControlItem, .ID, t, StrPtr(.SubItems(t))
                    Next
                    If (.Flags And &H10000000) = 0& Then    ' else flagged as no initial subitem selected
                        If Not (.Flags And &HF) = ctlType_Menu Then _
                            pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetSelectedControlItem, .ID, (.Flags And &H7FFF00) \ &H100&
                    End If
                '//// Command Button
                Case ctlType_CommandButton
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddPushButton, .ID, StrPtr(.Caption)
                '//// TextBox
                Case ctlType_TextBox
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddEditBox, .ID, StrPtr(.Caption)
                '//// Label
                Case ctlType_Label
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddText, .ID, StrPtr(.Caption)
                '//// Separator
                Case ctlType_Separator
                    pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_AddSeparator, .ID
                '//// Dialog selection textbox label
                Case ctlType_SelectionBoxLabel
                    pvCallInterface ObjPtr(m_Dialog), ifFileDialog_SetFileNameLabel, StrPtr(.Caption)
                '//// Cancel button
                Case ctlType_CancelButton           ' IFileDialog2 method, set Cancel button caption
                    If Not m_Dlg2 Is Nothing Then _
                        pvCallInterface ObjPtr(m_Dlg2), ifFileDialog2_SetCancelButtonLabel, StrPtr(.Caption)
                End Select
                ' if control is not to be enabled/visible, change its property now
                If (.Flags And &H30&) < &H30& Then pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlState, .ID, (.Flags And &H30&) \ &H10&
                If (.Flags And &H80) Then   ' if individual subitem enabled/visible states set, do that now
                    For t = 0& To UBound(.SubItemState)
                        If .SubItemState(t) < &H3& Then ' &H3 = ppEnabled Or ppVisible
                            pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlItemState, .ID, t, CLng(.SubItemState(t))
                        End If
                    Next
                End If
                If bProminent = False Then              ' test for prominent setting
                    If (.Flags And &H40) Then           ' prominent requested
                        Select Case (.Flags And &HF)
                        Case ctlType_CheckBox, ctlType_ComboBox, ctlType_CommandButton, ctlType_Menu
                            If bIsCnt = True Then       ' make container is prominent, not control
                                If (m_Controls(itemNr).Flags And &H7FFF00) = &H100& Then _
                                    bProminent = True: lIndex = m_Controls(itemNr).ID
                            Else                        ' make control prominent
                                bProminent = True: lIndex = .ID
                            End If
                            If bProminent Then _
                                pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_MakeProminent, lIndex
                        End Select
                    End If
                End If
            End With
        Next
        ' close container as needed
        If bIsCnt = True Then
            pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_EndVisualGroup
            If (m_Controls(itemNr).Flags And &H30) < &H30 Then _
                pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlState, m_Controls(itemNr).ID, (m_Controls(itemNr).Flags And &H30&) \ &H10&
        End If
    Next
    pvSetGetValues = True
    
ExitRoutine:
End Function

Private Sub pvFreeOrCreateRootPIDL(bCreate As Boolean, bEmpty As Boolean)

    ' called from a couple different routines
    ' if m_NavRestrictor has a long value in it, then that is a PIDL
    '   the PIDL must be released to prevent memory leak

    If bCreate Then
        Dim lPidl As Long, oItem As stdole.IUnknown
        If VarType(m_NavRestrictor) = vbDataObject Then ' IUnknown
            Set oItem = m_NavRestrictor
        ElseIf VarType(m_NavRestrictor) = vbString Then
            Set oItem = pvGetIShellItem(m_NavRestrictor)
        End If
        If Not oItem Is Nothing Then
            SHGetIDListFromObject ObjPtr(oItem), lPidl
            If lPidl Then m_NavRestrictor = lPidl
            Set oItem = Nothing
        End If
    ElseIf VarType(m_NavRestrictor) = vbLong Then
        ILFree CLng(m_NavRestrictor)
        m_NavRestrictor = Empty
    ElseIf bEmpty Then
        m_NavRestrictor = Empty
    End If

End Sub

Private Sub pvSetFileName()

    ' force pre-Vista FileName property behavior to Vista+ behavior

    Dim lLen As Long, lPosA As Long, lPosZ As Long
    Dim sPath As String, sFile As String
    
    If ofn.nFileOffset Then ' set by the file dialog when something was selected
        If Mid$(ofn.lpstrFile, ofn.nFileOffset, 1) = vbNullChar Then  ' multiple files chosen
            lLen = InStr(ofn.nFileOffset, ofn.lpstrFile, vbNullChar & vbNullChar)
            If lLen = 0& Then lLen = Len(ofn.lpstrFile) + 1&
            lPosA = ofn.nFileOffset
            sPath = vbNullChar & Left$(ofn.lpstrFile, lPosA - 1&) & "\"
            On Error Resume Next ' potential that appending path to tons of selections, string exceeds limit
            Do
                ' move thru the multi-selected items and build return string, nullchar-delimited
                lPosZ = InStr(lPosA + 1&, ofn.lpstrFile, vbNullChar)
                If lPosZ = 0& Then lPosZ = lLen
                sFile = sFile & sPath & Mid$(ofn.lpstrFile, lPosA + 1&, lPosZ - lPosA - 1&)
                If Err Then
                    Err.Clear
                    Exit Do
                End If
                If lPosZ >= lLen Then Exit Do
                lPosA = lPosZ
            Loop
            ofn.lpstrFile = Mid$(sFile, 2&)
            On Error GoTo 0
        Else    ' one file selected
            lLen = InStr(ofn.lpstrFile, vbNullChar)
            If lLen Then ofn.lpstrFile = Left$(ofn.lpstrFile, lLen - 1&)
        End If
    Else
        ofn.lpstrFile = vbNullString
    End If

End Sub

Private Function pvKeyToIndex(ByVal Key As String, ByVal bAppend As Boolean) As Long
    ' helper function to xref passed key to m_Controls() item
    ' optionally; new item can be created & key assigned to it

    Dim lIndex As Long
    If m_ControlKeys Is Nothing Then                ' m_Controls() not initialized yet
        If bAppend = False Then pvKeyToIndex = -1&
    Else
        On Error Resume Next
        lIndex = m_ControlKeys(Key)                 ' get array index
        If Err Then
            Err.Clear                               ' Key not used yet
            If Key = "-2147483648" Then ' &H80000000    ' will always be m_Controls(0)
                bAppend = False
            ElseIf bAppend = False Then
                pvKeyToIndex = -1&
            End If
        Else
            pvKeyToIndex = lIndex: bAppend = False  ' Key exists
        End If
        On Error GoTo 0
    End If
    
    If bAppend Then                                 ' create new entry in m_Controls()
        If m_ControlKeys Is Nothing Then
            Set m_ControlKeys = New Collection      ' initialize m_Controls()
            If Not Key = "-2147483648" Then lIndex = 1&
            ReDim m_Controls(0 To lIndex)
        Else
            lIndex = UBound(m_Controls) + 1&        ' add new m_Controls() entry
            ReDim Preserve m_Controls(0 To lIndex)
        End If
        If lIndex Then
            m_ControlKeys.Add lIndex, Key           ' update cross-reference
            pvKeyToIndex = lIndex                   ' return m_Controls() index
        End If
    End If
    
End Function

Private Function pvGetIShellItem(Source As Variant, Optional ptrShellItem As Long) As stdole.IUnknown
    ' Attempts to create IShellItem from passed source
    ' Source can contain a path in 1 of 4 formats:
    ' 1) plain text path
    ' 2) PIDL to the path
    ' 3) an IShellItem stdole.IUnknown object
    ' 4) a string GUID representing a virtual folder; can be prefixed with double colons
    
    ' ptrShellItem is only passed, as desired, from ShowBrowseForFolder. Source will be PIDL
    
    If IsEmpty(Source) Then Exit Function
    On Error GoTo ExitRoutine
    IIDFromString StrPtr(IID_IShellItem), VarPtr(m_Guid(0))
    If VarType(Source) = vbDataObject Then  ' IUnknown
        pvCallInterface ObjPtr(Source), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(pvGetIShellItem)
    ElseIf VarType(Source) = vbString Then
        If Left$(Source, 1) = "{" And Right$(Source, 1) = "}" And Len(Source) = 38& Then
            SHCreateItemFromParsingName StrPtr("::" & Source), 0&, VarPtr(m_Guid(0)), VarPtr(pvGetIShellItem)
        ElseIf Not Len(Source) = 0& Then
            SHCreateItemFromParsingName StrPtr(CStr(Source)), 0&, VarPtr(m_Guid(0)), VarPtr(pvGetIShellItem)
        End If
    ElseIf VarType(Source) = vbLong Then
        If ptrShellItem Then
            SHCreateItemFromIDList Source, VarPtr(m_Guid(0)), ptrShellItem
        Else
            SHCreateItemFromIDList Source, VarPtr(m_Guid(0)), VarPtr(pvGetIShellItem)
        End If
    End If

ExitRoutine:
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Private Function pvCallInterface(ByVal InterfacePointer As Long, ByVal VTableOffset As pvInterfaceMethodOffsets, _
                            ParamArray FunctionParameters() As Variant) As Variant
                            
' Coded to call ActiveX or COM objects, not standard dlls
' If the function fails, the return value is: Empty and Err.LastDllError is reason

' Parameters:
'   InterfacePointer. A pointer to an object/class, i.e., ObjPtr(IPicture)
'       Passing invalid pointers likely to result in crashes
'   VTableOffset. The zero-bound ordinal offset from the passed InterfacePointer where the virtual function exists.
'       Example: call IUnknown::Release (3rd interface method (ordinal #2), no additional parameters):
'            pvCallInterface InterfacePointer, 2

' FYI: SetLastError used in this routine so that you can call Err.LastDllError to see the reason for failure
'   Optionally. This can be changed to raise an error instead.

    Dim pIndex As Long, pCount As Long
    Dim pTypes As Long, pValues As Long
    Dim vParamPtr() As Long, vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant
    Const CallConvention As Long = 4&                   ' STDCALL
    
    If VTableOffset < 0& Or InterfacePointer = 0& Then
        pIndex = 5
    Else
        pCount = UBound(FunctionParameters) + 1&
        If pCount Then                                  ' else no return value (sub vs function)
            vParams() = FunctionParameters()            ' copy passed parameters, if any
            ReDim vParamPtr(0 To pCount - 1&)           ' need matching array of parameter types
            ReDim vParamType(0 To pCount - 1&)          ' and pointers to the parameters
            For pIndex = 0& To pCount - 1&
                vParamPtr(pIndex) = VarPtr(vParams(pIndex))
                vParamType(pIndex) = VarType(vParams(pIndex))
            Next
            pTypes = VarPtr(vParamType(0)): pValues = VarPtr(vParamPtr(0))
        End If
        pIndex = DispCallFunc(InterfacePointer, VTableOffset * 4&, CallConvention, _
                              vbLong, pCount, pTypes, pValues, vRtn)
    End If
    If pIndex = 0& Then                                 ' 0 = S_OK
        pvCallInterface = vRtn                          ' return result
    Else
        SetLastError pIndex
    End If

End Function

Private Sub pzCreateThunks(bEvents As Long)

    ' Vista+... routine creates 1 or 2 run-time interfaces that require interaction with the class
    '           and possibly a subclassing thunk if Controls_SetCustomMode() was called
    ' XP/Win2K... routine creates a callback thunk
    
    If (m_StateFlags And &H1000000) Then    ' old-style browse for folder
        If obif.lpfnCallback = 0& Then      ' callback created?
            If m_VTable Then pzDestroyThunks ' remove file dialog callbacks as needed
        Else
            Exit Sub
        End If
    ElseIf m_VTable Then
        If obif.lpfnCallback Then           ' old-style browse for folder callback exists?
            pzDestroyThunks
            If bEvents = 0& Then Exit Sub
        Else
            If Me.Version = dvXP_Win2K Then Exit Sub ' thunk already created
            If (m_StateFlags And emskCustMode) > &H2000& Then
                ' if needing subclassing thunk and not yet created, recreate VTable+thunks
                If m_SubclassFnc = 0& Then pzDestroyThunks
            ElseIf (bEvents And emskCtrlEvents) Then    ' want control events
                ' if needing 2nd interface and just 1 is created, recreate VTable+thunks
                If UBound(m_Events) = 0& Then pzDestroyThunks
            End If
            If m_VTable Then Exit Sub               ' thunks exist, done
        End If
    ElseIf bEvents = 0& Then
        Exit Sub                                ' no events requested
    End If
    
    Const CB_PAGE_RWX    As Long = &H40         ' Allocate executable memory
    Const CB_MEM_COMMIT  As Long = &H1000       ' Commit allocated memory
    Const FUNK_LEN As Long = 48&                ' size of thunk, per method
    Const IID_IFileDialogEvents As String = "{973510DB-7D7F-452B-8975-74A85828D354}"
    Const IID_IFileDialogControlEvents As String = "{36116642-D713-4B97-9B83-7484A9D00433}"
    
    Dim lTotalMethods As Long, p As Long, m As Long
    Dim offsetCode As Long, offsetVTable As Long, offsetThunk As Long
    Dim nMethods As Long, nrInterfaces As Long
    Dim cb() As Long
    
    ReDim cb(0 To FUNK_LEN \ 4& - 1&)           ' basic thunk structure 48 bytes
    cb(0) = ObjPtr(Me)                          ' thunks call back to this class
    cb(2) = &HBB60E089: cb(4) = &HE883C589: cb(5) = &HB9509004: cb(7) = &H74FF06E3
    cb(8) = &HFAE2008D: cb(9) = &H53FF33FF: cb(10) = &HC2906104
'   cb(1)=callback mem addr: cb(3)=thunk mem addr: cb(6)=nr params: cb(11)=LenB(all params)
    
    offsetCode = pvFindLastMethod
    
    If (Me.Version = dvXP_Win2K Or ((m_StateFlags And &H1000000) > 0&)) Then
        ' create simple callback thunk for pre-Vista or folder browser
        m_Vsize = FUNK_LEN
        m_VTable = VirtualAlloc(0&, m_Vsize, CB_MEM_COMMIT, CB_PAGE_RWX)
    
        If (m_StateFlags And &H1000000) Then    ' folder browser mode
            obif.lpfnCallback = pvSafeOffset32(m_VTable, 8)
            p = 4
        Else                                    ' pre-Vista dialog
            p = 8
            m_SubclassFnc = pvSafeOffset32(m_VTable, 8)
        End If
        CopyMemory cb(1), ByVal pvSafeOffset32(offsetCode, -p), 4& ' address of callback
        cb(3) = m_VTable
        cb(6) = 4&: cb(11) = cb(6) * 4&
        CopyMemory ByVal m_VTable, cb(0), FUNK_LEN
        
    Else
    
        ' logic used below to create a VTable that implements multiple interfaces:
        ' for each interface we create a stack of 4-byte pointers: 3 for IUnknown + 1 for each interface method
        '   IUnknown.QueryInterface     redirects to pvVtable_IUnknownQueryInterface
        '   IUnknown.AddRef             redirects to pvVtable_IUnknownAddRef
        '   IUnknonw.Release            redirects to pvVtable_IUnknownRelease
        '   for each interface method   redirect to pvIFileDialog_xxxx methods
        ' note: IFileDialogControlEvents requires IFileDialogEvents
        ' after the stack of pointers, then we add a 48-byte thunk per interface method, for all interfaces
        ' if needed, we finish with a 48-byte thunk for subclassing the dialog
    
        lTotalMethods = 7& + 4&  ' tally of all event methods (used or not) in this class, except IUnknown methods
        ' ^^ 7 for IFileDialogEvents, 4 for IFileDialogControlEvents
        
        If Not m_DlgMod Is Nothing Then             ' only create if customizing dialog
            If (bEvents And emskCtrlEvents) Then    ' want control events
                ReDim m_Events(0 To 1)
                IIDFromString StrPtr(IID_IFileDialogControlEvents), VarPtr(m_Events(1).IID(0))
                nrInterfaces = 2&   ' IFileDialogControlEvents requires IFileDialogEvents
                nMethods = lTotalMethods
            End If
        End If
        If nrInterfaces = 0& Then
            ReDim m_Events(0 To nrInterfaces)
            nrInterfaces = 1&
            nMethods = 7&
        End If
        IIDFromString StrPtr(IID_IFileDialogEvents), VarPtr(m_Events(0).IID(0))

        ' the 1st +3 below is for the 3 IUnknown function addresses
        ' the 2nd +3 below is to include the pvDialog_XPHookProc & pvDialog_Subclass & old-style FolderBrowser function addresses
        ' locate the address for pvVtable_IUnknownQueryInterface
        offsetCode = pvSafeOffset32(offsetCode, (lTotalMethods + 3& + 3&) * -4&)
        offsetThunk = (nMethods + 3&) * 4& + nrInterfaces * 12& ' where in our stack the first thunk exists
        m_Vsize = offsetThunk + (nMethods + 3&) * FUNK_LEN      ' calc size of stack (VTables + Thunks)
        If (m_StateFlags And emskCustMode) > &H2000& Then m_Vsize = m_Vsize + FUNK_LEN
        m_VTable = VirtualAlloc(0&, m_Vsize, CB_MEM_COMMIT, CB_PAGE_RWX)
        offsetThunk = pvSafeOffset32(m_VTable, offsetThunk)
        offsetVTable = m_VTable
        
        ' let's build the 3 IUnknown thunks first & add them to the VTable. All active interfaces share these
        For p = 0& To 2&
            CopyMemory cb(1), ByVal offsetCode, 4&  ' location of callback function in this class
            cb(3) = offsetThunk                     ' location of thunk code
            cb(6) = IIf(p = 0&, 3&, 1&)             ' number parameters for this method
            cb(11) = cb(6) * 4&                     ' Bytes to release on return
            CopyMemory ByVal pvSafeOffset32(offsetVTable, p * 4&), pvSafeOffset32(offsetThunk, 8), 4&
            CopyMemory ByVal offsetThunk, cb(0), FUNK_LEN ' copy this thunk
            offsetThunk = pvSafeOffset32(offsetThunk, FUNK_LEN)
            offsetCode = pvSafeOffset32(offsetCode, 4)
       Next
        
        ' now let's build the methods for active interfaces
        For p = 0& To nrInterfaces - 1&
            m = CoTaskMemAlloc(8&)                  ' create soft Object for each interface
            CopyMemory ByVal m, offsetVTable, 4&    ' set value at ObjPtr()
            CopyMemory ByVal pvSafeOffset32(m, 4), 0&, 4& ' Ref count
            m_Events(p).VTable = m                  ' cache VTable locally
    
            ' copy IUnknown entries for 2nd and subsequent interfaces
            If Not offsetVTable = m_VTable Then CopyMemory ByVal offsetVTable, ByVal m_VTable, 12&
            offsetVTable = pvSafeOffset32(offsetVTable, 12)
    
            ' setup the current interface's methods
            If p = 0& Then m = 7& Else m = 4&       ' nr methods for current interface
            For m = 1& To m
                CopyMemory cb(1), ByVal offsetCode, 4&  ' location of callback function in this class
                cb(3) = offsetThunk                     ' location of thunk code
                If p = 0& Then                      ' IFileDialogEvents 7 events
                    Select Case m                   ' set number of parameters + 1 for pUnk
                        Case 2: cb(6) = 3&
                        Case 5, 7: cb(6) = 4&
                        Case Else: cb(6) = 2&
                    End Select
                Else                                ' IFileDialogControlEvents 4 events
                    Select Case m                   ' set number of parameters + 1 for pUnk
                        Case 1, 3: cb(6) = 4&
                        Case Else: cb(6) = 3&
                    End Select
                End If
                cb(11) = cb(6) * 4&                 ' Bytes to release on return
                CopyMemory ByVal offsetVTable, pvSafeOffset32(offsetThunk, 8), 4&
                CopyMemory ByVal offsetThunk, cb(0), FUNK_LEN      ' copy this thunk
    
                offsetThunk = pvSafeOffset32(offsetThunk, FUNK_LEN) ' prep for next interface
                offsetCode = pvSafeOffset32(offsetCode, 4)
                offsetVTable = pvSafeOffset32(offsetVTable, 4)
            Next
        Next
        
        ' create the subclass thunk, as needed
        If (m_StateFlags And emskCustMode) > &H2000& Then
            ' move to the correct function pointer, if all interfaces were not created
            If Not (nMethods = lTotalMethods) Then offsetCode = pvSafeOffset32(offsetCode, (lTotalMethods - nMethods) * 4&)
            CopyMemory cb(1), ByVal offsetCode, 4& ' address of pvDialog_Subclass
            cb(3) = offsetThunk
            cb(6) = 6&: cb(11) = cb(6) * 4&
            CopyMemory ByVal offsetThunk, cb(0), FUNK_LEN
            m_SubclassFnc = pvSafeOffset32(offsetThunk, 8)
        End If
    End If
    
End Sub

Private Sub pzDestroyThunks()
    ' destroy any thunks we created
    If m_VTable Then
        Const CB_MEM_DECOMMIT  As Long = &H4000&    ' Decommit allocated memory flag
        Const CB_MEM_RELEASE   As Long = &H8000&    ' Release allocated memory flag
        If Not VirtualFree(m_VTable, m_Vsize, CB_MEM_DECOMMIT) = 0& Then _
            VirtualFree m_VTable, 0&, CB_MEM_RELEASE
        If Me.Version > dvXP_Win2K Then
            On Error Resume Next    ' m_Events() may not be initialized
            For m_Vsize = LBound(m_Events) To UBound(m_Events)
                If Err Then
                    Err.Clear
                    Exit For
                End If
                If m_Events(m_Vsize).VTable Then CoTaskMemFree m_Events(m_Vsize).VTable
            Next
            Erase m_Events()
        End If
        m_VTable = 0&: m_Vsize = m_VTable: m_SubclassFnc = m_VTable
        obif.lpfnCallback = 0&
    End If
End Sub

Private Function pvFindLastMethod() As Long
    '/// This method must not be modified.
    ' Return the address of the last private method
  Dim bSub  As Byte                         ' Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                         ' Address of the vTable
  Dim i     As Long                         ' Loop index
  Dim j     As Long                         ' Loop limit
  Dim tblLoc As Long
    CopyMemory nAddr, ByVal ObjPtr(Me), 4&  ' Get the address of this object instance
    If pvProbeClassMethods(pvSafeOffset32(nAddr, &H1C), i, bSub) = 0 Then Exit Function  ' Probe for a Class method
    i = pvSafeOffset32(i, 4)                 ' Bump to the next entry
    j = pvSafeOffset32(i, 1024)              ' Set a reasonable limit, scan 256 vTable entries
    Do Until i = j
      CopyMemory nAddr, ByVal i, 4&         ' Get the address stored in this vTable entry
      If IsBadCodePtr(nAddr) Then           ' Is the entry an invalid code address?
        tblLoc = i                          ' Cache the vTable end-point
        GoTo Found                          ' Bad method signature, quit loop
      End If
      CopyMemory bVal, ByVal nAddr, 1&      ' Get the byte pointed to by the vTable entry
      If bVal <> bSub Then                  ' If the byte doesn't match the expected value...
        tblLoc = i                          ' Cache the vTable end-point
        GoTo Found                          ' Bad method signature, quit loop
      End If
      i = pvSafeOffset32(i, 4)               ' Next vTable entry
    Loop
    Exit Function                           ' Final method not found
    
Found:                                      ' Return the specified vTable entry address
  pvFindLastMethod = tblLoc
End Function

Private Function pvProbeClassMethods(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  '/// This method must not be modified.
  'Probe at the specified start address for a method signature
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  nAddr = nStart                            ' Start address
  nLimit = pvSafeOffset32(nAddr, 32)         ' Probe eight entries
  Do Until nAddr = nLimit                   ' While we've not reached our probe depth
    CopyMemory nEntry, ByVal nAddr, 4&      ' Get the vTable entry
    If nEntry <> 0 Then                     ' If not an implemented interface
      CopyMemory bVal, ByVal nEntry, 1&     ' Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then    ' Check for a native or pcode method signature
        nMethod = nAddr                     ' Store the vTable entry
        bSub = bVal                         ' Store the found method signature
        pvProbeClassMethods = True          ' Indicate success
        Exit Function                       ' Return
      End If
    End If
    nAddr = pvSafeOffset32(nAddr, 4)         ' Next vTable entry
  Loop
End Function

Private Sub pvSubclassDialog(bSet As Boolean)
    
    ' helper function subclass/unsubclass the dialog's "Open" button
    '   subclassing can occur when Controls_SetCustomMode is called
    '   See pvDialog_Subclass() for more
    
    If (m_StateFlags And &H10000000) Then   ' abort if action already taken
        If bSet Then Exit Sub
    ElseIf bSet = False Then
        Exit Sub
    End If
    
    Dim cWnd As Long, hWnd As Long
    Const GWL_WNDPROC As Long = -4
    Const GWL_ID As Long = -12
    
    hWnd = Me.IFileDialog_GetHwnd()         ' locate the "Open" button, ID will be 1: ID_OK
    If hWnd = 0& Then Exit Sub
    Do
        cWnd = FindWindowEx(hWnd, cWnd, StrPtr("Button"), 0&)
        If cWnd = 0& Then Exit Do
        If GetWindowLong(cWnd, GWL_ID) = ID_OK Then
            If bSet Then
                If SetWindowSubclass(cWnd, m_SubclassFnc, cWnd, 0&) Then
                    m_StateFlags = m_StateFlags Or &H10000000
                    SetWindowSubclass hWnd, m_SubclassFnc, hWnd, hWnd
                End If
            Else
                m_StateFlags = m_StateFlags And Not &H10000000
                RemoveWindowSubclass cWnd, m_SubclassFnc, cWnd
                RemoveWindowSubclass hWnd, m_SubclassFnc, hWnd
            End If
            Exit Do
        End If
    Loop
    
End Sub

Private Function pvSafeOffset32(ByVal VBpointer As Long, ByVal AdjAmount As Double) As Long
    ' helper function to ensure adding/subtracting from 32bit pointer wraps around
    ' the high bit as needed. Since class can be used in 64 bit O/S with /LARGEADDRESSAWARE
    ' option during compilation; helps ensure pointer math across 2GB boundary.
    
    Const UINT_MaxPlus1 As Double = 4294967296#
    Const INT_MaxPlus1 As Double = 2147483648#
    
    Dim dPointer As Double
    If Abs(AdjAmount) >= UINT_MaxPlus1 Then Err.Raise 6
    
    If VBpointer > -1& Then
        dPointer = VBpointer + AdjAmount
    Else
        dPointer = VBpointer + AdjAmount + UINT_MaxPlus1
    End If
    If dPointer < INT_MaxPlus1 Then
        If dPointer < 0# Then Err.Raise 6
        pvSafeOffset32 = dPointer
    Else
        pvSafeOffset32 = (dPointer - INT_MaxPlus1) Or SIGN_BIT
    End If

End Function

Private Function pvLoadString(hModule As Long, sDLL As String, sID As Long, _
                            Optional bufferSize As Long = MAX_PATH) As String

    ' helper function for loading locale-aware strings from DLLs
    ' caller must FreeLibrary on hModule

    Dim sValue As String, n As Long
    Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
    
    If hModule = 0& Then
        hModule = LoadLibraryEx(StrPtr(sDLL), 0&, LOAD_LIBRARY_AS_DATAFILE)
        If hModule = 0& Then Exit Function
    End If
    sValue = String$(bufferSize, vbNullChar)
    If LoadString(hModule, sID, ByVal StrPtr(sValue), bufferSize) Then
        n = InStr(sValue, vbNullChar)
        If n > 1& Then pvLoadString = Left$(sValue, n - 1&)
    End If
    
End Function

Private Function pvUpdateBasket(lCount As Long) As stdole.IUnknown

    ' if lCount < 0 then return basket As stdole.IUnknown (IShellItemArray)
    '   else add m_ModeResults to basket and update item count on dialog, if applicable

    Dim oItem As stdole.IUnknown, oArray As stdole.IUnknown
    Dim aPIDLs() As Long, aCount As Long
    Dim n As Long, c As Long, aPtr As Long
    Dim sCaption As String, p As Long

    If lCount < 0& Then
        IIDFromString StrPtr(IID_IShellItemArray), VarPtr(m_Guid(0))
        For c = 1& To m_BasketItems.Count
            Set oItem = m_BasketItems.Item(c)
            If pvCallInterface(ObjPtr(oItem), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(oArray)) = 0& Then
                If pvCallInterface(ObjPtr(oArray), ifShellItemArray_GetCount, VarPtr(n)) = 0& Then
                    aCount = aCount + n: Set oItem = Nothing
                    ReDim Preserve aPIDLs(0 To aCount - 1&)
                    For n = 0& To n - 1&
                        If pvCallInterface(ObjPtr(oArray), ifShellItemArray_GetItemAt, n, VarPtr(oItem)) = 0& Then
                            SHGetIDListFromObject ObjPtr(oItem), aPIDLs(aPtr)
                            For p = aPtr - 1& To 0 Step -1&
                                If ILIsEqual(aPIDLs(p), aPIDLs(aPtr)) Then Exit For
                            Next
                            If p < 0& Then aPtr = aPtr + 1&
                            Set oItem = Nothing
                        End If
                    Next
                End If
                Set oArray = Nothing
            End If
        Next
        If aPtr > 0& Then
            SHCreateShellItemArrayFromIDLists aPtr, VarPtr(aPIDLs(0)), pvUpdateBasket
            For aPtr = 0& To aPtr - 1&
                CoTaskMemFree aPIDLs(aPtr)
            Next
        End If
    Else
        If lCount > 0& Then m_BasketItems.Add m_ModeResults
        n = pvKeyToIndex(SIGN_BIT + 1&, False)
        If n > -1& Then
            With m_Controls(n)
                lCount = ((.Flags And &HFFFF000) \ &H1000& + lCount) And &H7FFF&
                .Flags = (.Flags And &H80000FFF) Or (lCount * &H1000&)
                pvCallInterface ObjPtr(m_DlgMod), ifFileDlgCustomize_SetControlLabel, .ID, StrPtr(Replace$(.Caption, "%s", CStr(lCount)))
            End With
        End If
    End If
    
End Function

Private Function pvProcessCustomMode() As Long

    ' this routine processes the subclassed "Open" button. User wants to select something
    ' applies only when Controls_SetCustomMode called and a subclass was created
    
    ' overriding default behavior is a bit complex
    
    Dim bCancel As Boolean, nrItems As Long
    Dim lAttrs As Long, oPtr As Long, aPtr As Long
    Dim lOptions As Long, n As Long, aData() As Long
    Dim IFolderView2 As stdole.IUnknown
    Const IfIFolderView2_GetSelection As Long = 32
    
    pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetOptions, VarPtr(lOptions)
    If (lOptions And DLG_AllowMultiSelect) Then
        ' see IFileDialog_GetCurrentSelection method for comments why we ask for multi-select items this way
        Set IFolderView2 = pvGetIFolderView2
        If Not IFolderView2 Is Nothing Then
            pvCallInterface ObjPtr(IFolderView2), IfIFolderView2_GetSelection, 0&, VarPtr(m_ModeResults)
            Set IFolderView2 = Nothing
        End If
    End If
    If m_ModeResults Is Nothing Then
        Set m_ModeResults = Me.IFileDialog_GetCurrentSelection(True, nrItems)
        If m_ModeResults Is Nothing Then Exit Function      ' no items selected
    Else
        pvCallInterface ObjPtr(m_ModeResults), ifShellItemArray_GetCount, VarPtr(nrItems)
    End If
    
    Select Case (m_StateFlags And emskCustMode)
     
     Case cm_BasketModeFilesOrFolders, cm_PickFilesOrFolders ' can select anything
         If (m_StateFlags And emskCustMode) = cm_BasketModeFilesOrFolders Then
             RaiseEvent DialogAddBasketItem(m_ModeResults, bCancel)
         ElseIf (m_StateFlags And emskAllEvents) Then
             RaiseEvent DialogOnFileOk(bCancel)
         End If
         If bCancel Then
             Set m_ModeResults = Nothing ' function returns 0 to keep dialog open
         ElseIf (m_StateFlags And emskCustMode) = cm_PickFilesOrFolders Then
             pvProcessCustomMode = True   ' closes dialog
         Else
             pvUpdateBasket nrItems      ' dialog remains open
             Set m_ModeResults = Nothing
         End If
         
     Case cm_BrowseFoldersShowFiles, cm_BasketModeFoldersOnlyShowFiles
         For n = 0& To nrItems - 1&                ' test attributes of each selected item
             pvCallInterface ObjPtr(m_ModeResults), ifShellItemArray_GetItemAt, n, VarPtr(oPtr)
             If oPtr Then
                 pvCallInterface oPtr, ifShellItem_GetAttributes, -1&, VarPtr(lAttrs)
                 pvCallInterface oPtr, ifUnknown_Release: oPtr = 0&
                 If (lAttrs And siAttr_FOLDER) = 0& Then Exit For    ' not a folder
                 If (lOptions And DLGex_AllNonStorageItems) = 0 Then ' only system folders (not zip content)?
                     If Not (lAttrs And (siAttr_FILESYSTEM Or siAttr_STREAM)) = siAttr_FILESYSTEM Then Exit For
                 End If
             End If
         Next
         If n < nrItems Then                         ' at least one non-folder selected
             If nrItems = 1& Then                    ' else multi-selection
                 nrItems = 0&
             Else
                 ReDim aData(0 To nrItems - 1&)          ' loop again, keeping only valid items
                 For n = 0& To nrItems - 1&
                     pvCallInterface ObjPtr(m_ModeResults), ifShellItemArray_GetItemAt, n, VarPtr(oPtr)
                     If oPtr Then
                         pvCallInterface oPtr, ifShellItem_GetAttributes, -1&, VarPtr(lAttrs)
                         If (lAttrs And siAttr_FOLDER) Then
                             SHGetIDListFromObject oPtr, aData(aPtr): aPtr = aPtr + 1&
                             If (lOptions And DLGex_AllNonStorageItems) = 0 Then ' only system folders (not zip content)?
                                 If Not (lAttrs And (siAttr_FILESYSTEM Or siAttr_STREAM)) = siAttr_FILESYSTEM Then _
                                     aPtr = aPtr - 1&
                             End If
                         End If
                         pvCallInterface oPtr, ifUnknown_Release: oPtr = 0&
                     End If
                 Next
                 Set m_ModeResults = Nothing: nrItems = aPtr
                 If nrItems > 0& Then                      ' rebuild results, skipping invalid selections
                    SHCreateShellItemArrayFromIDLists nrItems, VarPtr(aData(0)), m_ModeResults
                    For nrItems = 0& To nrItems - 1&
                        CoTaskMemFree aData(nrItems)
                    Next
                End If
            End If
            If (m_StateFlags And emskCustMode) = cm_BrowseFoldersShowFiles And nrItems = 0& Then
                ' only files were selected, choose the files folder & place into IShellItemArray
                pvCallInterface ObjPtr(m_Dialog), ifFileDialog_GetFolder, VarPtr(m_ModeResults)
                If Not m_ModeResults Is Nothing Then
                    SHGetIDListFromObject ObjPtr(m_ModeResults), aPtr
                    Set m_ModeResults = Nothing
                    If aPtr Then
                        SHCreateShellItemArrayFromIDLists 1&, VarPtr(aPtr), m_ModeResults
                        CoTaskMemFree aPtr: nrItems = 1&
                    End If
                End If
            End If
         End If
         If nrItems < 1& Then
            bCancel = True
         ElseIf (m_StateFlags And emskCustMode) = cm_BasketModeFoldersOnlyShowFiles Then
             RaiseEvent DialogAddBasketItem(m_ModeResults, bCancel)
         ElseIf (m_StateFlags And emskAllEvents) Then
             RaiseEvent DialogOnFileOk(bCancel)
         End If
         If bCancel Then
             Set m_ModeResults = Nothing     ' keep dialog open
         ElseIf (m_StateFlags And emskCustMode) = cm_BasketModeFoldersOnlyShowFiles Then
             pvUpdateBasket nrItems          ' keep dialog open
             Set m_ModeResults = Nothing
         Else
             pvProcessCustomMode = True       ' close dialog
         End If
     
     Case cm_CompressedFolderPlusFiles, cm_BasketModeFilesOnly
         For n = 0& To nrItems - 1&
             pvCallInterface ObjPtr(m_ModeResults), ifShellItemArray_GetItemAt, n, VarPtr(oPtr)
             If oPtr Then
                 pvCallInterface oPtr, ifShellItem_GetAttributes, -1&, VarPtr(lAttrs)
                 pvCallInterface oPtr, ifUnknown_Release: oPtr = 0&
                 If (lAttrs And siAttr_FOLDER) Then
                     If Not (lAttrs And (siAttr_FILESYSTEM Or siAttr_STREAM)) = _
                         (siAttr_FILESYSTEM Or siAttr_STREAM) Then Exit For
                 End If
             End If
         Next
         If n < nrItems Then
             Set m_ModeResults = Nothing
         ElseIf (m_StateFlags And emskCustMode) = cm_BasketModeFilesOnly Then
             pvUpdateBasket nrItems
             Set m_ModeResults = Nothing
         Else
             If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogOnFileOk(bCancel)
             If bCancel Then Set m_ModeResults = Nothing Else pvProcessCustomMode = True
         End If
         
    'Case cm_PickFilesOrFolders, cm_BasketModeFoldersOnly  ' these are not subclassed
         
    End Select

End Function

Private Function pvGetIFolderView2() As stdole.IUnknown

    ' helper function to return IFolderView2 interface from the dialog, roundabout method

    Dim aGuid(0 To 7) As Long, oPtr As Long
    Const IID_IServiceProvider As String = "{6D5140C1-7436-11CE-8034-00AA006009FA}"
    Const IID_IFolderView2 As String = "{1AF3A467-214F-4298-908E-06B03E0B39F9}"
    Const SID_FolderView As String = "{CDE725B0-CCC9-4519-917E-325D72FAB4CE}"
    Const IfISP_QueryService As Long = 3

    ' this routine basically replicates the IUnknown_QueryService API
    IIDFromString StrPtr(IID_IServiceProvider), VarPtr(m_Guid(0))
    pvCallInterface ObjPtr(m_Dialog), ifUnknown_QueryInterface, VarPtr(m_Guid(0)), VarPtr(oPtr)
    If oPtr Then
        ReDim aData(0 To 7) ' need to use 2 GUIDs (service GUID + IFolderView2 GUID)
        IIDFromString StrPtr(SID_FolderView), VarPtr(aGuid(0))
        IIDFromString StrPtr(IID_IFolderView2), VarPtr(aGuid(4))
        pvCallInterface oPtr, IfISP_QueryService, VarPtr(aGuid(0)), VarPtr(aGuid(4)), VarPtr(pvGetIFolderView2)
        pvCallInterface oPtr, ifUnknown_Release
    End If
    
End Function

Private Sub Class_Initialize()
    Dim osv As OSVERSIONINFO
    osv.dwOSVersionInfoSize = Len(osv)
    GetVersionEx osv
    Select Case osv.dwMajorVersion
    Case Is < 5     ' less than Win2K/XP
    Case 5          ' Win2K or XP
        m_StateFlags = 5
    Case Else       ' Vista or better
        m_StateFlags = osv.dwMinorVersion + osv.dwMajorVersion
    End Select
    
'     m_StateFlags = 5  ' <<< un-rem to pretend system is pre-Vista

End Sub

Private Sub Class_Terminate()
    Me.Clear
End Sub

' ////////////////////////////////////////////////////////////////////////////////////////////////////////

'   NO NON-INTERFACE METHODS/PROCEDURES CAN BE ADDED AFTER THIS POINT. EVERYTHING FROM THIS
'   POINT TO END OF CLASS ARE INTERFACE METHODS OR SUBCLASS PROCEDURES IN A SPECIFIC ORDER.

' ////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Function pvVtable_IUnknownQueryInterface(ByVal pUnk As Long, ByVal riid As Long, ByRef pOut As Long) As Long
  '/// This method must not be modified. Each active interface has its own IID
  ' If supporting multiple interfaces, pUnk will be one of those that were created or IUnknown
    
    pOut = 0&
    If riid = 0& Then
        pvVtable_IUnknownQueryInterface = E_POINTER ' required if null pointer received
    Else
        Dim n As Long
        CopyMemory m_Guid(0), ByVal riid, 16&
        If m_Guid(0) = 0& Then   ' querying for IUnknown? {00000000-0000-0000-C000-000000000046}
            If m_Guid(1) = 0& And m_Guid(2) = 192& And m_Guid(3) = 1174405120 Then _
                pOut = m_Events(0).VTable
        Else
            For n = 0& To UBound(m_Events) ' querying for specific interface
                If m_Guid(0) = m_Events(n).IID(0) And m_Guid(1) = m_Events(n).IID(1) Then
                    If m_Guid(2) = m_Events(n).IID(2) And m_Guid(3) = m_Events(n).IID(3) Then _
                        pOut = m_Events(n).VTable: Exit For
                End If
            Next
        End If
        If pOut = 0& Then
            pvVtable_IUnknownQueryInterface = E_NOINTERFACE
        Else
            Call pvVtable_IUnknownAddRef(pOut) ' must AddRef to anything we return
        End If
    End If
End Function
Private Function pvVtable_IUnknownAddRef(ByVal pUnk As Long) As Long
  '/// This method must not be modified. Each active interface has its own ref counter
  ' If supporting multiple interfaces, pUnk will be one of those that were created
    Dim lPtr As Long
    lPtr = pvSafeOffset32(pUnk, 4)
    CopyMemory pvVtable_IUnknownAddRef, ByVal lPtr, 4&
    pvVtable_IUnknownAddRef = pvVtable_IUnknownAddRef + 1&
    CopyMemory ByVal lPtr, pvVtable_IUnknownAddRef, 4&
End Function
Private Function pvVtable_IUnknownRelease(ByVal pUnk As Long) As Long
  '/// This method must not be modified. Each active interface has its own ref counter
  ' If supporting multiple interfaces, pUnk will be one of those that were created
    Dim lPtr As Long
    lPtr = pvSafeOffset32(pUnk, 4)
    CopyMemory pvVtable_IUnknownRelease, ByVal lPtr, 4&
    pvVtable_IUnknownRelease = pvVtable_IUnknownRelease - 1&
    CopyMemory ByVal lPtr, pvVtable_IUnknownRelease, 4&
End Function

Private Function pvIFileDialog_OnFileOk(ByVal pUnk As Long, ByVal pFD As Long) As Long
    ' Open/Save button clicked
    
    Dim bCancel As Boolean, nrItems As Long
    
    Select Case (m_StateFlags And emskCustMode)
    Case cm_BrowseFoldersShowFiles, cm_BasketModeFoldersOnlyShowFiles
        ' abort: option is to select folders only, file selected via DblClick else this event not triggered
        pvIFileDialog_OnFileOk = S_FALSE
    
    Case cm_BasketModeFilesOrFolders, cm_BasketModeFilesOnly, cm_BasketModeFoldersOnly ' add to basket
        Set m_ModeResults = Me.IFileDialog_GetCurrentSelection(True, nrItems)
        If Not m_ModeResults Is Nothing Then
            RaiseEvent DialogAddBasketItem(m_ModeResults, bCancel)
            If bCancel = False Then pvUpdateBasket nrItems
            Set m_ModeResults = Nothing
        End If
        pvIFileDialog_OnFileOk = S_FALSE
        
   'Case cm_CompressedFolderPlusFiles, cm_PickFilesOrFolders
        ' these are handled in the subclassing procedure: pvProcessCustomMode
        
    Case Else
        If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogOnFileOk(bCancel)
        If bCancel Then pvIFileDialog_OnFileOk = S_FALSE Else pvGetTextBoxValue
    End Select
    
End Function
Private Function pvIFileDialog_OnFolderChanging(ByVal pUnk As Long, ByVal pFD As Long, ByVal psiFolder As stdole.IUnknown) As Long
    If (m_StateFlags And emskAllEvents) Then
        Dim bCancel As Boolean
        RaiseEvent DialogOnFolderChanging(psiFolder, bCancel)
        If bCancel Then pvIFileDialog_OnFolderChanging = S_FALSE
    End If
End Function
Private Function pvIFileDialog_OnFolderChange(ByVal pUnk As Long, ByVal pFD As Long) As Long
    If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogOnFolderChanged
End Function
Private Function pvIFileDialog_OnSelectionChanged(ByVal pUnk As Long, ByVal pFD As Long) As Long
    If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogOnSelectionChanged
End Function
Private Function pvIFileDialog_OnShareViolation(ByVal pUnk As Long, ByVal pFD As Long, ByVal psItem As stdole.IUnknown, ByRef pResponse As Long) As Long
    If (m_StateFlags And emskAllEvents) Then
        Dim lResult As DialogResponseEnum
        RaiseEvent DialogOnShareViolation(psItem, lResult)
    End If
End Function
Private Function pvIFileDialog_OnTypeChange(ByVal pUnk As Long, ByVal pFD As Long) As Long
    If m_StateFlags > -1& Then
        m_StateFlags = m_StateFlags Or SIGN_BIT
        If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogOnInit(Me.IFileDialog_GetHwnd)
        If (m_StateFlags And emskCustMode) > &H2000& Then pvSubclassDialog True
    ElseIf (m_StateFlags And emskAllEvents) Then
        RaiseEvent DialogOnFileTypeChange
    End If
End Function
Private Function pvIFileDialog_OnOverwrite(ByVal pUnk As Long, ByVal pFD As Long, ByVal psItem As stdole.IUnknown, ByRef pResponse As Long) As Long
    If (m_StateFlags And emskAllEvents) Then
        Dim lResult As DialogResponseEnum
        RaiseEvent DialogOnOverwrite(psItem, lResult)
    End If
End Function

Private Function pvIFileDialog_OnItemSelected(ByVal pUnk As Long, ByVal pFDC As Long, ByVal dwIDCtl As Long, ByVal dwIDItem As Long) As Long
    ' combobox (dwIDItem=listindex), menu (dwIDItem=menuItem) or optionbutton (index) selection
    If (m_StateFlags And emskCtrlEvents) Then RaiseEvent DialogSubItemClicked(-dwIDCtl, dwIDItem)
End Function
Private Function pvIFileDialog_OnButtonClicked(ByVal pUnk As Long, ByVal pFDC As Long, ByVal dwIDCtl As Long) As Long

    If dwIDCtl = SIGN_BIT Then
        '   "Ok" button when custom basket mode used
        Dim bCancel As Boolean
        If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogOnFileOk(bCancel)
        If bCancel = False Then Me.CloseDialog 0&, 0&
        
    ElseIf (m_StateFlags And emskCtrlEvents) Then
        RaiseEvent DialogButtonClicked(-dwIDCtl)
    End If
    
End Function
Private Function pvIFileDialog_OnCheckButtonToggled(ByVal pUnk As Long, ByVal pFDC As Long, ByVal dwIDCtl As Long, ByVal isChecked As Long) As Long
    If (m_StateFlags And emskCtrlEvents) Then RaiseEvent DialogCheckBoxChanged(-dwIDCtl, isChecked)
End Function
Private Function pvIFileDialog_OnControlActivating(ByVal pUnk As Long, ByVal pFDC As Long, ByVal dwIDCtl As Long) As Long
    If (m_StateFlags And emskCtrlEvents) Then RaiseEvent DialogMenuToBeDisplayed(-dwIDCtl)
End Function

Private Function pvDialog_Subclass(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                            ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    
    ' subclassing the Vista+ dialog's "Open" button, as needed
    ' Subclassing used to prevent adding extra controls to the dialog and have the
    '   browse for folder option, while showing files, work as if no files were shown
    ' Basically, we prevent the button click from working. Instead
    '   1) If folder is selected, we process the click internally and close the dialog with success
    '   2) Otherwise, we simply ignore the click
    
    ' WARNING     WARNING     WARNING     WARNING     WARNING     WARNING     WARNING     WARNING
    ' --------------------------------------------------------------------------------------------
    ' DO NOT place a stop or breakpoint in this routine
    ' This function is not truly safe to walk thru. The thunk uses safe subclassing techniques
    '   but wasn't designed to be IDE safe. Walking this function is at your own risk.
    '   It is safer to walk from the raised event, but still not 100% IDE-safe.
    '   pvProcessCustomMode routine may raise events
    
    If hWnd = dwRefData Then    ' dialog else "open" button
    
        If uMsg = WM_COMMAND Then
            If wParam = ID_OK Then                          ' BM_Click message?
                If pvProcessCustomMode = True Then
                    Me.CloseDialog 0&, 0&                   ' unsubclass, get added textbox value, close dialog
                    PostMessage hWnd, uMsg, ID_OK, 0&       ' ensure dialog closes (may fail/delay while subclassed)
                End If
                Exit Function
            End If
        End If
        
    ElseIf uMsg = BM_SETSTYLE Then                          ' adding the split-button style?
        If lParam = ID_OK And wParam = &HD Then Exit Function
    End If
    
    pvDialog_Subclass = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    
End Function

Private Function pvDialog_XPHookProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' ref: https://msdn.microsoft.com/library/ms646931(v=VS.85).aspx
    ' minimal effort used here; XP is kinda obsolete, but at least user has
    ' a means of receiving events.
    ' Note: in the event, you can call this class' CustomData property to get/set that value at any time

    Dim bEatIt As Boolean
    If (m_StateFlags And &H20000000) = 0& Then ' else Me.CloseDialog was called
        If uMsg = WM_INITDIALOG Then
            m_StateFlags = m_StateFlags Or &H40000000 ' flag indicating dialog activated
            If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogXPInit(hWnd)
            
        ElseIf (m_StateFlags And &H40000000) Then
            If uMsg = WM_NOTIFY Then
                If (ofn.Flags And DLG_AllowMultiSelect) > 0& Then
                    ' silently adjust the buffer to accept all selected files (if possible)
                    CopyMemory m_Guid(0), ByVal lParam, 12& ' get 12-byte NMHDR structure
                    If m_Guid(2) = CDN_SELCHANGE Then       ' notification node & get character count
                        m_Guid(0) = SendMessage(GetParent(hWnd), CDM_GETSPEC, 0&, ByVal 0&)
                        If m_Guid(0) > Len(ofn.lpstrFile) Then
                            On Error Resume Next            ' adjust the buffer as needed
                            ofn.lpstrFile = ofn.lpstrFile & String$(MAX_PATH, vbNullChar)
                            If Err Then
                                Err.Clear
                            Else                            ' adjust the buffer count
                                ofn.nMaxFile = ofn.nMaxFile + MAX_PATH
                            End If
                            On Error GoTo 0
                        End If
                    End If
                End If
            End If
            If (m_StateFlags And emskAllEvents) Then
                RaiseEvent DialogXPEvent(hWnd, uMsg, wParam, lParam, bEatIt)
                pvDialog_XPHookProc = bEatIt
            End If
        ElseIf (m_StateFlags And emskAllEvents) Then
            RaiseEvent PreInitMessage(hWnd, uMsg, wParam, lParam, bEatIt)
            pvDialog_XPHookProc = bEatIt
        End If
    End If

End Function

Private Function pvDialogBrowseForFolderProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    
    ' ref: https://msdn.microsoft.com/en-us/e9109f38-34c7-46c0-aa0c-a6b4570f1c3a
    ' Note: in the event, you can call this class' CustomData property to get/set that value at any time
    
    If uMsg = BFFM_INITIALIZED Then
        If Len(ofn.lpstrTitle) Then SendMessage hWnd, WM_SETTEXT, 0&, ByVal StrPtr(ofn.lpstrTitle)
        If (m_StateFlags And emskAllEvents) Then RaiseEvent DialogBrowseFoldersInit(hWnd)
        obif.lpCallbackData = hWnd
        m_StateFlags = m_StateFlags Or &H2000000
        If (m_StateFlags And &H4000000) Then    ' initialize with startup folder
            SendMessage obif.lpCallbackData, BFFM_SETSELECTIONW, 1&, ByVal StrPtr(ofn.lpstrInitialDir)
        End If
        
    Else
    
        If (m_StateFlags And &HE000000) = &H6000000 Then  ' initialized, waiting on 2nd init folder action
            ' tweak to enure selected item is visible in tree (scrolled to)
            If uMsg = BFFM_SELCHANGED Then
                m_StateFlags = m_StateFlags Xor &H4000000
                SendMessage obif.lpCallbackData, BFFM_SETSELECTIONW, 1&, ByVal StrPtr(ofn.lpstrInitialDir)
            End If
        End If
        
        If (m_StateFlags And emskAllEvents) Then
            If (m_StateFlags And &H2000000) Then    ' initialized
                If uMsg = BFFM_SELCHANGED Then
                    RaiseEvent DialogBrowseFoldersOnSelectionChanged(pvGetIShellItem(lParam))
                ElseIf uMsg = BFFM_VALIDATEFAILEDW Then ' only sent when flags include BIF_VALIDATE
                    Dim bClose As Boolean
                    RaiseEvent DialogBrowseFoldersValidateFailed(pvComStrPtrToVBstring(lParam, False), bClose)
                    If bClose = False Then pvDialogBrowseForFolderProc = 1&
                Else
                    ' note: a typical unhandled message is uMsge=5. lParam is a pointer to an IUnknown interface
                    RaiseEvent DialogBrowseFoldersEvent(hWnd, uMsg, lParam)
                End If
            Else
                RaiseEvent DialogBrowseFoldersPreInitMessage(hWnd, uMsg, lParam)
            End If
        End If
    End If
    
End Function
' ////////////////////////////////////////////////////////////////////////////////////////////////////////
'   NO CODE WHATSOEVER CAN EXIST PAST THIS POINT. COMMENTS ARE OK, BUT NOTHING ELSE
' ////////////////////////////////////////////////////////////////////////////////////////////////////////


