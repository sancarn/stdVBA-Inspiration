VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'  -----======== PURPOSE: A unicode aware/compatible open/save dialog class ========-----
' Class Instancing = MultiUse (can be created by users)

Private Declare Function GetSaveFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function GetOpenFileNameA Lib "comdlg32.dll" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (ByVal pOpenfilename As Long) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

' ref: http://msdn.microsoft.com/en-us/library/ms646839(VS.85).aspx
Private Type OPENFILENAME
   lStructSize          As Long     'Length of structure, in bytes
   hWndOwner            As Long     'Window that owns the dialog, or NULL
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
   Flags                As Long     'See OFN_Flags
   nFileOffset          As Integer  'Offset to file name in full path, 0-based
   nFileExtension       As Integer  'Offset to file ext in full path, 0-based (excl '.')
   lpstrDefExt          As String   'Default ext appended, excl '.', max 3 chars
   lCustData            As Long     'Appl defined data for lpfnHook
   lpfnHook             As Long     'Pointer to hook procedure
   lpTemplateName       As Long     'Template Name
   ' 76 bytes for O/S less than Win2K
   pvReserved           As Long     'new Win2000 / WinXP members
   dwReserved           As Long     'new Win2000 / WinXP members
   FlagsEx              As Long     'new Win2000 / WinXP members
End Type


' ref: http://msdn.microsoft.com/en-us/library/ms646916(VS.85).aspx
Public Enum CommonDialogErrorsEnum ' see DisplayError for descriptions
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
    CDERR_CANCELED = vbObjectError
End Enum

' ref: http://msdn.microsoft.com/en-us/library/ms646839(VS.85).aspx
Public Enum CommonDialogFlagsEnum
    OFN_READONLY = &H1                  ' shows ReadOnly checkbox and on return, Flags property contains &H1, box was checked when file selected else it wasn't
    OFN_OVERWRITEPROMPT = &H2           ' prompts user for confirmation if overwriting file during ShowSave
    OFN_HIDEREADONLY = &H4              ' Hides read-only checkbox
    OFN_NOCHANGEDIR = &H8               ' Prevents system directory change during browsing; NT4/2K/XP: does not work with ShowOpen
    OFN_SHOWHELP = &H10                 ' shows help button on dialog, sends CDN_HELP to hook procedure
    OFN_ENABLEHOOK = &H20               ' Makes the dialog hookable. CustomHookProc property must be set
    'OFN_ENABLETEMPLATE = &H40          ' determines how Templates are used; N/A for this class
    'OFN_ENABLETEMPLATEHANDLE = &H80    ' determines how hInstance is used pertaining to templates; N/A for this class
    OFN_NOVALIDATE = &H100              ' allows invalid characters in the file name
    OFN_ALLOWMULTISELECT = &H200        ' allows multiselection of files; OFN_Explorer=True then null terminated list else space terminated
    OFN_EXTENSIONDIFFERENT = &H400      ' if DefaultExt property <> "", set on return if user selects filename different than default extension
    OFN_PATHMUSTEXIST = &H800           ' forces user to open/save files in existing paths only
    OFN_FILEMUSTEXIST = &H1000          ' ShowOpen only: forces user to select existing files only; must also use OFN_PATHMUSTEXIST
    OFN_CREATEPROMPT = &H2000           ' prompts user to create non-existant file during ShowSave
    OFN_SHAREAWARE = &H4000             ' sends CDN_SHAREVIOLATION messages to hook procedure
    OFN_NOREADONLYRETURN = &H8000&      ' Selected file selected without ReadOnly checkbox checked and is not in write-protected folder
    OFN_NOTESTFILECREATE = &H10000      ' refers to create-nonmodify network shares. See link
    OFN_NONETWORKBUTTON = &H20000       ' Hides/disables the Network button
    OFN_NOLONGNAMES = &H40000           ' Forces Win3.1 style boxes to use short file names; OFN_Explore ignores this flag
    OFN_EXPLORER = &H80000              ' shows Explorer-type window else old Win3.1 type window. See link if wanting 3.1 window for single-select dialogs
    OFN_NODEREFERENCELINKS = &H100000   ' Returns LNK file; else select LNK files return what they link to
    OFN_LONGNAMES = &H200000            ' Forces long file names if Win3.1 style box is shown
    OFN_ENABLEINCLUDENOTIFY = &H400000  ' Win2K+: sends CDN_INCLUDEITEM messages to hook procedure
    OFN_ENABLESIZING = &H800000         ' Non-NT4: allows user to size when hook procedure is used. If not hooked, user can size anyway
    'OFN_USEMONIKERS = &H1000000
    OFN_DONTADDTORECENT = &H2000000     ' Win2K+: don't include file in MRU
    '&H4000000 & &H8000000 not defined
    OFN_FORCESHOWHIDDEN = &H10000000    ' Win2K+: Show files with hidden attributes. If hidden+system, then not shown
   
    ' the following are sent to hook procedures only?
    'OFN_SHAREFALLTHROUGH = 2&           ' if sharing violation occurs, filename is accepted: ShowOpen
    'OFN_SHARENOWARN = 1&                ' if sharing violation occurs, don't display error: ShowOpen
End Enum
Public Enum CommonDialogFlagsExEnum
    OFN_EX_NOPLACESBAR = &H1             ' Win2K+: hides the side bar
End Enum

Private Enum CommonDialogModeEnum
    modeOpen = 0&
    modeSave = 1&
End Enum

Private Const MAX_PATH As Long = 260&

Private ofn As OPENFILENAME     ' the API's UDT
Private m_Flags As Long         ' see Flags property & InitStructure method
Private m_CancelErr As Boolean  ' see CancelError property
Private m_minBufferSize As Long ' see MinimumBufferSizeRequired property

Public Function ShowOpen(ByVal ownerHwnd As Long) As Boolean
    ' returns true if user selected a file or files and no error
    Dim lReturn As Long
   
    InitStructure ownerHwnd, modeOpen
    If IsWindowUnicode(GetDesktopWindow) Then
        lReturn = GetOpenFileNameW(VarPtr(ofn))
    Else
        lReturn = GetOpenFileNameA(ofn)
    End If
    If lReturn = 0& Then 'failure
        lReturn = CommDlgExtendedError()
        If Not (lReturn = 0& And m_CancelErr = False) Then DisplayError lReturn, modeOpen
    Else
        ShowOpen = True
    End If

End Function

Public Function ShowSave(ByVal ownerHwnd As Long) As Boolean
    ' returns true if user did not cancel & no error
    Dim lReturn As Long
   
    InitStructure ownerHwnd, modeSave
    If IsWindowUnicode(GetDesktopWindow) Then
        lReturn = GetSaveFileNameW(VarPtr(ofn))
    Else
        lReturn = GetSaveFileNameA(ofn)
    End If
    If lReturn = 0& Then 'failure
        lReturn = CommDlgExtendedError()
        If Not (lReturn = 0& And m_CancelErr = False) Then DisplayError lReturn, modeSave
    Else
        ShowSave = True
    End If

End Function

Public Sub Clear()
    ' clears all properties and resets to default
    Dim blankOFN As OPENFILENAME
    ofn = blankOFN
    m_Flags = 0&
    m_minBufferSize = m_Flags
    m_CancelErr = False
End Sub

' forces an error to occur if user cancels/closes dialog without selecting anything
Public Property Get CancelError() As Boolean
    CancelError = m_CancelErr
End Property
Public Property Let CancelError(ByVal cErr As Boolean)
    m_CancelErr = cErr
End Property

' returns/sets the filename the dialog starts with or returns
Public Property Get FileName() As String
    Dim lLen As Long
    If ofn.nFileOffset Then
        lLen = InStr(ofn.nFileOffset, ofn.lpstrFile, vbNullChar) - 1&
        If lLen < 1& Then lLen = Len(ofn.lpstrFile)
    Else
        lLen = lstrlenW(StrPtr(ofn.lpstrFile))
    End If
    If lLen Then FileName = Left$(ofn.lpstrFile, lLen)
End Property
Public Property Let FileName(ByVal fTitle As String)
    If ofn.nMaxFile = 0 Or Len(fTitle) > ofn.nMaxFile Then
        If ofn.nMaxFile = 0 Then
            ofn.nMaxFile = MAX_PATH * 2&
        Else
            ofn.nMaxFile = Len(fTitle)
        End If
        ofn.lpstrFile = String$(ofn.nMaxFile, vbNullChar)
    End If
    If fTitle = vbNullString Then
        Mid$(ofn.lpstrFile, 1, 1) = vbNullChar
    ElseIf ofn.nMaxFile = Len(fTitle) Then
        ofn.lpstrFile = fTitle
    Else
        Mid$(ofn.lpstrFile, 1, Len(fTitle) + 1) = fTitle & vbNullChar
    End If
End Property

' reutrns/sets the file title dialog starts with or returns
Public Property Get FileTitle() As String
    Dim lLen As Long
    lLen = lstrlenW(StrPtr(ofn.lpstrFileTitle))
    If lLen Then FileTitle = Left$(ofn.lpstrFileTitle, lLen)
End Property
Public Property Let FileTitle(ByVal fTitle As String)
    If ofn.nMaxFileTitle = 0 Or Len(fTitle) > ofn.nMaxFileTitle Then
        If ofn.nMaxFileTitle = 0 Then
            ofn.nMaxFileTitle = MAX_PATH
        Else
            ofn.nMaxFileTitle = Len(fTitle)
        End If
        ofn.lpstrFileTitle = String$(ofn.nMaxFileTitle, vbNullChar)
    End If
    If fTitle = vbNullString Then
        Mid$(ofn.lpstrFileTitle, 1, 1) = vbNullChar
    ElseIf ofn.nMaxFileTitle = Len(fTitle) Then
        ofn.lpstrFileTitle = fTitle
    Else
        Mid$(ofn.lpstrFileTitle, 1, Len(fTitle) + 1) = fTitle & vbNullChar
    End If
End Property

' returns/sets the initial dialog; FileName property overrides this setting
Public Property Get InitDir() As String
    InitDir = ofn.lpstrInitialDir
End Property
Public Property Let InitDir(ByVal fPath As String)
    If fPath <> vbNullString Then
        If Right$(fPath, 1) <> "\" Then fPath = fPath & "\"
    End If
    ofn.lpstrInitialDir = fPath
End Property

' returns/sets the file path the dialog starts with or returns; FileName overrides this
Public Property Get FilePath() As String
    If ofn.nFileOffset Then
        FilePath = Left$(ofn.lpstrFile, ofn.nFileOffset - 1)
    End If
End Property
Public Property Let FilePath(ByVal fPath As String)
    InitDir = fPath
End Property

' returns the file extension of the selected file
Public Property Get FileExtension() As String
    If ofn.nFileExtension Then
        Dim lLen As Long
        FileExtension = Mid$(ofn.lpstrFile, ofn.nFileExtension + 1)
        lLen = lstrlenW(StrPtr(FileExtension))
        If lLen Then FileExtension = Left$(FileExtension, lLen)
    End If
End Property

' returns/sets the dialog window caption/title
Public Property Get DialogTitle() As String
    DialogTitle = ofn.lpstrTitle
End Property
Public Property Let DialogTitle(ByVal dTitle As String)
    ofn.lpstrTitle = dTitle
End Property

' return/sets the filter(s)
Public Property Get Filter() As String
    Dim Index As Long, lLen As String
    lLen = InStr(ofn.lpstrFilter, vbNullChar & vbNullChar) - 1&
    If lLen < 1& Then lLen = Len(ofn.lpstrFilter)
    Filter = Replace$(Left$(ofn.lpstrFilter, lLen), vbNullChar, "|")
End Property
Public Property Let Filter(ByVal fFilter As String)
    If fFilter = vbNullString Then
        ofn.lpstrFilter = vbNullString
    Else
        ofn.lpstrFilter = Replace$(fFilter, "|", vbNullChar) & vbNullChar
    End If
End Property

' returns/sets the flags used for the dialog
Public Property Get Flags() As CommonDialogFlagsEnum
    Flags = m_Flags
End Property
Public Property Let Flags(ByVal dFlags As CommonDialogFlagsEnum)
    If dFlags >= 0& Then m_Flags = (dFlags And Not OFN_EXTENSIONDIFFERENT)
End Property

' returns/sets extended Win2K+ flags for the dialog
Public Property Get FlagsEx() As CommonDialogFlagsExEnum
    FlagsEx = ofn.FlagsEx ' valid on Win2000 and above only
End Property
Public Property Let FlagsEx(ByVal dFlags As CommonDialogFlagsExEnum)
    ofn.FlagsEx = dFlags ' valid on Win2000 and above only
End Property

' returns/sets maximum buffer size for returned files
Public Property Get MaxFileSize() As Long
    If ofn.nMaxFile = 0& Then
        MaxFileSize = MAX_PATH * 2&
    Else
        MaxFileSize = ofn.nMaxFile
    End If
End Property
Public Property Let MaxFileSize(ByVal nSize As Long)
    If nSize < 1& Then
        ofn.nMaxFile = 0&
    Else
        ofn.nMaxFile = nSize
        ofn.lpstrFile = String$(nSize, vbNullChar)
    End If
End Property

' returns/sets maximum buffer size for file title
Public Property Get MaxFileTitleSize() As Long
    If ofn.nMaxFileTitle = 0& Then
        MaxFileTitleSize = MAX_PATH
    Else
        MaxFileTitleSize = ofn.nMaxFileTitle
    End If
End Property
Public Property Let MaxFileTitleSize(ByVal nSize As Long)
    If nSize < 1& Then
        ofn.nMaxFileTitle = 0&
    Else
        ofn.nMaxFileTitle = nSize
        ofn.lpstrFileTitle = String$(nSize, vbNullChar)
    End If
End Property

' returns/sets the Filter index dialog starts with or returns
Public Property Get FilterIndex() As Long
    FilterIndex = ofn.nFilterIndex
End Property
Public Property Let FilterIndex(ByVal nIndex As Long)
    If nIndex < 0 Then nIndex = 0  'set to zero to use custom filter if CustomFilterSize>0
    ofn.nFilterIndex = nIndex
End Property

' returns/sets the default extension applied to files in ShowSave
Public Property Get DefaultExt() As String
    DefaultExt = ofn.lpstrDefExt
End Property
Public Property Let DefaultExt(ByVal fExt As String)
    ofn.lpstrDefExt = fExt
End Property

' returns/sets custom value sent to hook procedure
Public Property Get CustomHookData() As Long
    CustomHookData = ofn.lCustData
End Property
Public Property Let CustomHookData(ByVal nPointer As Long)
    ofn.lCustData = nPointer
End Property

' returns/sets the hook procedure address: pass result of AddressOf()
' the Flags property must include: OFN_ENABLEHOOK
' ref: http://msdn.microsoft.com/en-us/library/ms646931(VS.85).aspx
Public Property Get CustomHookProc() As Long
    CustomHookProc = ofn.lpfnHook
End Property
Public Property Let CustomHookProc(ByVal nProc As Long)
    ofn.lpfnHook = nProc
End Property

' returns/sets buffer size for user-entered filters
Public Property Get CustomFilterSize() As Long
    CustomFilterSize = ofn.nMaxCustFilter
End Property
Public Property Let CustomFilterSize(ByVal nSize As Long)
    If nSize < 1 Then
        ofn.nMaxCustFilter = 0&
        ofn.lpstrCustomFilter = vbNullString
    Else
        ofn.nMaxCustFilter = nSize
        ofn.lpstrCustomFilter = String$(nSize, vbNullChar)
    End If
End Property

' Contains minimum buffer size required if error FNERR_BUFFERTOOSMALL is returned
' A value of -1 indicates unknown and should never happen
' Reset to zero before dialog is displayed
Public Property Get MinimumBufferSizeRequired() As Long
    MinimumBufferSizeRequired = m_minBufferSize
End Property

Private Sub DisplayError(ByVal errCode As Long, Mode As CommonDialogModeEnum)

    Dim sMode As String
    If Mode = modeOpen Then sMode = "Open" Else sMode = "Save"

' NOTE: If your VB is stopping in this routine, it is due to your error trapping settings
' Select Tools|Options from the menu above
' Select the "General Tab"
' Select option button: "Break on Unhandled Errors"
    
    Select Case errCode
    Case 0&
        errCode = CDERR_CANCELED
        Err.Raise errCode Or vbObjectError, "Common Dialog " & sMode, "Cancel was seleted"
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

Private Sub InitStructure(hWnd As Long, Mode As CommonDialogModeEnum)

    Const VER_PLATFORM_WIN32_NT As Long = 2
    Const OPENFILENAME_SIZE_VERSION_400 = 76
  
    ' determine required API UDT structure size
    If ofn.lStructSize = 0& Then
        Dim osv As OSVERSIONINFO
        osv.dwOSVersionInfoSize = Len(osv)
        If GetVersionEx(osv) Then
            If osv.dwPlatformId = VER_PLATFORM_WIN32_NT Then
                If osv.dwMajorVersion < 5& Then osv.dwPlatformId = 0&
            End If
        End If
        If osv.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            ofn.lStructSize = Len(ofn)
        Else
            ofn.lStructSize = OPENFILENAME_SIZE_VERSION_400
        End If
    End If
   
    ofn.hWndOwner = hWnd    ' set hWnd; disables that window til dialog closed
    m_minBufferSize = 0&    ' reset before each call
   
    ' ensure the FileName property is minimally sized
    If ofn.nMaxFile = m_minBufferSize Then
        ofn.nMaxFile = MAX_PATH * 2
        ofn.lpstrFile = String$(ofn.nMaxFile, vbNullChar)
    End If
    ' ensure the FileTitle property is minimally sized
    If ofn.nMaxFileTitle = m_minBufferSize Then
        ofn.nMaxFileTitle = MAX_PATH
        ofn.lpstrFileTitle = String$(ofn.nMaxFileTitle, vbNullChar)
    End If
   
    ' remove any flags that are in conflict with the dialog mode (save/open)
    If Mode = modeSave Then
        ofn.Flags = (m_Flags And Not OFN_FILEMUSTEXIST)
    ElseIf (m_Flags And OFN_FILEMUSTEXIST) Then
        ofn.Flags = ((m_Flags Or OFN_PATHMUSTEXIST) And Not OFN_OVERWRITEPROMPT)
    Else
        ofn.Flags = (m_Flags And Not OFN_OVERWRITEPROMPT)
    End If
   
End Sub


