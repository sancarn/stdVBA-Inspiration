VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIEFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'this module has built-in IE-feature-control according to:  http://msdn.microsoft.com/en-us/library/ee330733%28v=vs.85%29.aspx
Const FeaturePageSetup As String = "HKCU\Software\Microsoft\Internet Explorer\PageSetup\"
Const FeatureBaseKey As String = "HKCU\Software\Microsoft\Internet Explorer\Main\FeatureControl\"
Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
 
Private Sh As Object, mExeName As String, mInstalledVersion As String

Private Sub Class_Initialize()
  Set Sh = CreateObject("WScript.Shell")
 
  mExeName = Space$(260)
  mExeName = Left$(mExeName, GetModuleFileNameA(0, mExeName, Len(mExeName)))
  mExeName = Mid$(mExeName, InStrRev(mExeName, "\") + 1)
  
  If FEATURE_BROWSER_EMULATION = 7000 Then  'in this case it is not elevated yet - and we ensure new elevated Defaults
     FEATURE_ENABLE_WEB_CONTROL_VISUALS = 1
     FEATURE_RESTRICT_ACTIVEXINSTALL = 1
     FEATURE_ADDON_MANAGEMENT = 1
     FEATURE_AJAX_CONNECTIONEVENTS = 1
     FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION = 1
     FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS = 0
     FEATURE_MIME_HANDLING = 1
     FEATURE_DOMSTORAGE = 1
     FEATURE_GPU_RENDERING = 0      ' testen
     FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI = 0 ' testen
     FEATURE_NINPUT_LEGACYMODE = 0
     FEATURE_DISABLE_LEGACY_COMPRESSION = 1
     FEATURE_LOCALMACHINE_LOCKDOWN = 0
     FEATURE_BLOCK_LMZ_OBJECT = 0
     FEATURE_BLOCK_LMZ_SCRIPT = 0
     FEATURE_DISABLE_NAVIGATION_SOUNDS = 1
     FEATURE_SCRIPTURL_MITIGATION = 1
     FEATURE_SPELLCHECKING = 1
     FEATURE_STATUS_BAR_THROTTLING = 1
     FEATURE_TABBED_BROWSING = 1
     FEATURE_VALIDATE_NAVIGATE_URL = 1
     FEATURE_WEBOC_DOCUMENT_ZOOM = 1
     FEATURE_WEBOC_POPUPMANAGEMENT = 0
     FEATURE_WEBOC_MOVESIZECHILD = 1
     FEATURE_WEBSOCKET = 1
     FEATURE_WINDOW_RESTRICTIONS = 0
     FEATURE_XMLHTTP = 1
  End If
End Sub

Public Property Get ExeName() As String
  ExeName = mExeName
End Property
 
Public Property Get InstalledVersion() As String
  On Error Resume Next
    InstalledVersion = Sh.RegRead("HKLM\Software\Microsoft\Internet Explorer\svcVersion")
    If Len(InstalledVersion) = 0 Then mInstalledVersion = Sh.RegRead("HKLM\Software\Microsoft\Internet Explorer\Version")
    If Len(InstalledVersion) = 0 Then mInstalledVersion = 7 'let's assume that as a fallback, since it's the default-emulation for a Browser-Control
  If Err Then Err.Clear
End Property
 
Public Sub WriteFeature(FeatureKey As String, ByVal Value, Optional ValueType As String = "REG_DWORD")
  Sh.RegWrite FeatureBaseKey & FeatureKey & "\" & mExeName, Value, ValueType
End Sub
Public Function ReadFeature(FeatureKey As String, Optional ByVal Default)
  On Error Resume Next
    ReadFeature = Sh.RegRead(FeatureBaseKey & FeatureKey & "\" & mExeName)
    If Err Then
      If Not IsEmpty(Default) And FeatureKey <> "FEATURE_BROWSER_EMULATION" Then WriteFeature FeatureKey, Default
      ReadFeature = Default
    End If
  On Error GoTo 0
End Function

Public Property Get FEATURE_BROWSER_EMULATION() As Long
  FEATURE_BROWSER_EMULATION = ReadFeature("FEATURE_BROWSER_EMULATION", 7000)
End Property
Public Property Let FEATURE_BROWSER_EMULATION(ByVal Version As Long)
  If Version < 7 Then Version = 7
  If Version < 1000 Then Version = Version * 1000
  WriteFeature "FEATURE_BROWSER_EMULATION", Version
End Property
 
Public Property Get FEATURE_ENABLE_WEB_CONTROL_VISUALS()
  FEATURE_ENABLE_WEB_CONTROL_VISUALS = ReadFeature("FEATURE_ENABLE_WEB_CONTROL_VISUALS", 1)
End Property
Public Property Let FEATURE_ENABLE_WEB_CONTROL_VISUALS(ByVal Value)
  WriteFeature "FEATURE_ENABLE_WEB_CONTROL_VISUALS", Value
End Property
 
Public Property Get FEATURE_RESTRICT_ACTIVEXINSTALL()
  FEATURE_RESTRICT_ACTIVEXINSTALL = ReadFeature("FEATURE_RESTRICT_ACTIVEXINSTALL", 1)
End Property
Public Property Let FEATURE_RESTRICT_ACTIVEXINSTALL(ByVal Value)
  WriteFeature "FEATURE_RESTRICT_ACTIVEXINSTALL", Value
End Property

Public Property Get FEATURE_ADDON_MANAGEMENT()
  FEATURE_ADDON_MANAGEMENT = ReadFeature("FEATURE_ADDON_MANAGEMENT", 0)
End Property
Public Property Let FEATURE_ADDON_MANAGEMENT(ByVal Value)
  WriteFeature "FEATURE_ADDON_MANAGEMENT", Value
End Property

Public Property Get FEATURE_AJAX_CONNECTIONEVENTS()
  FEATURE_AJAX_CONNECTIONEVENTS = ReadFeature("FEATURE_AJAX_CONNECTIONEVENTS", 1)
End Property
Public Property Let FEATURE_AJAX_CONNECTIONEVENTS(ByVal Value)
  WriteFeature "FEATURE_AJAX_CONNECTIONEVENTS", Value
End Property

Public Property Get FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION()
  FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION = ReadFeature("FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION", 1)
End Property
Public Property Let FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION(ByVal Value)
  WriteFeature "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION", Value
End Property

Public Property Get FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS()
  FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS = ReadFeature("FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS", 1)
End Property
Public Property Let FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS(ByVal Value)
  WriteFeature "FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS", Value
End Property

Public Property Get FEATURE_MIME_HANDLING()
  FEATURE_MIME_HANDLING = ReadFeature("FEATURE_MIME_HANDLING", 1)
End Property
Public Property Let FEATURE_MIME_HANDLING(ByVal Value)
  WriteFeature "FEATURE_MIME_HANDLING", Value
End Property

Public Property Get FEATURE_DOMSTORAGE()
  FEATURE_DOMSTORAGE = ReadFeature("FEATURE_DOMSTORAGE", 1)
End Property
Public Property Let FEATURE_DOMSTORAGE(ByVal Value)
  WriteFeature "FEATURE_DOMSTORAGE", Value
End Property

Public Property Get FEATURE_GPU_RENDERING()
  FEATURE_GPU_RENDERING = ReadFeature("FEATURE_GPU_RENDERING", 1)
End Property
Public Property Let FEATURE_GPU_RENDERING(ByVal Value)
  WriteFeature "FEATURE_GPU_RENDERING", Value
End Property

Public Property Get FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI()
  FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI = ReadFeature("FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI", 0)
End Property
Public Property Let FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI(ByVal Value)
  WriteFeature "FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI", Value
End Property

Public Property Get FEATURE_NINPUT_LEGACYMODE()
  FEATURE_NINPUT_LEGACYMODE = ReadFeature("FEATURE_NINPUT_LEGACYMODE", 0)
End Property
Public Property Let FEATURE_NINPUT_LEGACYMODE(ByVal Value)
  WriteFeature "FEATURE_NINPUT_LEGACYMODE", Value
End Property

Public Property Get FEATURE_DISABLE_LEGACY_COMPRESSION()
  FEATURE_DISABLE_LEGACY_COMPRESSION = ReadFeature("FEATURE_DISABLE_LEGACY_COMPRESSION", 1)
End Property
Public Property Let FEATURE_DISABLE_LEGACY_COMPRESSION(ByVal Value)
  WriteFeature "FEATURE_DISABLE_LEGACY_COMPRESSION", Value
End Property

Public Property Get FEATURE_LOCALMACHINE_LOCKDOWN()
  FEATURE_LOCALMACHINE_LOCKDOWN = ReadFeature("FEATURE_LOCALMACHINE_LOCKDOWN", 0)
End Property
Public Property Let FEATURE_LOCALMACHINE_LOCKDOWN(ByVal Value)
  WriteFeature "FEATURE_LOCALMACHINE_LOCKDOWN", Value
End Property

Public Property Get FEATURE_BLOCK_LMZ_OBJECT()
  FEATURE_BLOCK_LMZ_OBJECT = ReadFeature("FEATURE_BLOCK_LMZ_OBJECT", 0)
End Property
Public Property Let FEATURE_BLOCK_LMZ_OBJECT(ByVal Value)
  WriteFeature "FEATURE_BLOCK_LMZ_OBJECT", Value
End Property

Public Property Get FEATURE_BLOCK_LMZ_SCRIPT()
  FEATURE_BLOCK_LMZ_SCRIPT = ReadFeature("FEATURE_BLOCK_LMZ_SCRIPT", 0)
End Property
Public Property Let FEATURE_BLOCK_LMZ_SCRIPT(ByVal Value)
  WriteFeature "FEATURE_BLOCK_LMZ_SCRIPT", Value
End Property

Public Property Get FEATURE_DISABLE_NAVIGATION_SOUNDS()
  FEATURE_DISABLE_NAVIGATION_SOUNDS = ReadFeature("FEATURE_DISABLE_NAVIGATION_SOUNDS", 1)
End Property
Public Property Let FEATURE_DISABLE_NAVIGATION_SOUNDS(ByVal Value)
  WriteFeature "FEATURE_DISABLE_NAVIGATION_SOUNDS", Value
End Property

Public Property Get FEATURE_SCRIPTURL_MITIGATION()
  FEATURE_SCRIPTURL_MITIGATION = ReadFeature("FEATURE_SCRIPTURL_MITIGATION", 1)
End Property
Public Property Let FEATURE_SCRIPTURL_MITIGATION(ByVal Value)
  WriteFeature "FEATURE_SCRIPTURL_MITIGATION", Value
End Property

Public Property Get FEATURE_SPELLCHECKING()
  FEATURE_SPELLCHECKING = ReadFeature("FEATURE_SPELLCHECKING", 1)
End Property
Public Property Let FEATURE_SPELLCHECKING(ByVal Value)
  WriteFeature "FEATURE_SPELLCHECKING", Value
End Property

Public Property Get FEATURE_STATUS_BAR_THROTTLING()
  FEATURE_STATUS_BAR_THROTTLING = ReadFeature("FEATURE_STATUS_BAR_THROTTLING", 1)
End Property
Public Property Let FEATURE_STATUS_BAR_THROTTLING(ByVal Value)
  WriteFeature "FEATURE_STATUS_BAR_THROTTLING", Value
End Property

Public Property Get FEATURE_TABBED_BROWSING()
  FEATURE_TABBED_BROWSING = ReadFeature("FEATURE_TABBED_BROWSING", 1)
End Property
Public Property Let FEATURE_TABBED_BROWSING(ByVal Value)
  WriteFeature "FEATURE_TABBED_BROWSING", Value
End Property

Public Property Get FEATURE_VALIDATE_NAVIGATE_URL()
  FEATURE_VALIDATE_NAVIGATE_URL = ReadFeature("FEATURE_VALIDATE_NAVIGATE_URL", 1)
End Property
Public Property Let FEATURE_VALIDATE_NAVIGATE_URL(ByVal Value)
  WriteFeature "FEATURE_VALIDATE_NAVIGATE_URL", Value
End Property

Public Property Get FEATURE_WEBOC_DOCUMENT_ZOOM()
  FEATURE_WEBOC_DOCUMENT_ZOOM = ReadFeature("FEATURE_WEBOC_DOCUMENT_ZOOM", 1)
End Property
Public Property Let FEATURE_WEBOC_DOCUMENT_ZOOM(ByVal Value)
  WriteFeature "FEATURE_WEBOC_DOCUMENT_ZOOM", Value
End Property

Public Property Get FEATURE_WEBOC_POPUPMANAGEMENT()
  FEATURE_WEBOC_POPUPMANAGEMENT = ReadFeature("FEATURE_WEBOC_POPUPMANAGEMENT", 1)
End Property
Public Property Let FEATURE_WEBOC_POPUPMANAGEMENT(ByVal Value)
  WriteFeature "FEATURE_WEBOC_POPUPMANAGEMENT", Value
End Property

Public Property Get FEATURE_WEBOC_MOVESIZECHILD()
  FEATURE_WEBOC_MOVESIZECHILD = ReadFeature("FEATURE_WEBOC_MOVESIZECHILD", 0)
End Property
Public Property Let FEATURE_WEBOC_MOVESIZECHILD(ByVal Value)
  WriteFeature "FEATURE_WEBOC_MOVESIZECHILD", Value
End Property

Public Property Get FEATURE_WEBSOCKET()
  FEATURE_WEBSOCKET = ReadFeature("FEATURE_WEBSOCKET", 1)
End Property
Public Property Let FEATURE_WEBSOCKET(ByVal Value)
  WriteFeature "FEATURE_WEBSOCKET", Value
End Property

Public Property Get FEATURE_WINDOW_RESTRICTIONS()
  FEATURE_WINDOW_RESTRICTIONS = ReadFeature("FEATURE_WINDOW_RESTRICTIONS", 1)
End Property
Public Property Let FEATURE_WINDOW_RESTRICTIONS(ByVal Value)
  WriteFeature "FEATURE_WINDOW_RESTRICTIONS", Value
End Property

Public Property Get FEATURE_XMLHTTP()
  FEATURE_XMLHTTP = ReadFeature("FEATURE_XMLHTTP", 1)
End Property
Public Property Let FEATURE_XMLHTTP(ByVal Value)
  WriteFeature "FEATURE_XMLHTTP", Value
End Property

Public Sub WritePageSetup(PageSetupKey As String, ByVal Value, Optional ValueType As String = "REG_SZ")
' sets e.g. "header" to a new string (attention: these are general settings for all IE-instances (their Print-Jobs)
' possible values for PageSetupKey are:
' "header", "footer", "margin_top","margin_left","margin_right""margin_bottom", "Shrink_To_Fit","Print_Background"
  Sh.RegWrite FeaturePageSetup & PageSetupKey, Value, ValueType
End Sub

Public Function ReadPageSetup(PageSetupKey As String, Optional ByVal Default)
  On Error Resume Next
    ReadPageSetup = Sh.RegRead(FeaturePageSetup & PageSetupKey)
    If Err Then
      ReadPageSetup = Default
      If Not IsEmpty(Default) Then WritePageSetup PageSetupKey, Default
    End If
  On Error GoTo 0
End Function
'----------------------- End of IE-Features-Interface --------------------------


