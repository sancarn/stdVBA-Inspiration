Attribute VB_Name = "MdlRutinas"
Private Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "Shell32" Alias "SHGetFileInfoA" _
    (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long



Public Function GetFileDescription(ByVal sPath As String) As String
    Const SHGFI_TYPENAME = &H400
    Dim FInfo As SHFILEINFO
    SHGetFileInfo sPath, 0, FInfo, Len(FInfo), SHGFI_TYPENAME
    GetFileDescription = Left$(FInfo.szTypeName, InStr(FInfo.szTypeName & vbNullChar, vbNullChar) - 1)
End Function

Public Function GetFileIcon(ByVal sPath As String) As Long
    Const SHGFI_SMALLICON = &H1
    Const SHGFI_ICON = &H100
    Dim FInfo As SHFILEINFO
    SHGetFileInfo sPath, 0, FInfo, Len(FInfo), SHGFI_ICON Or SHGFI_SMALLICON
    GetFileIcon = FInfo.hIcon
End Function

Public Function EncodeURL(ByVal sURL As String) As String
    Dim i           As Long
    Dim sChar       As String * 1

    For i = 1 To Len(sURL)
        sChar = Mid$(sURL, i, 1)
        Select Case sChar
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_", ".", "~"
                EncodeURL = EncodeURL & sChar
            Case Else
                EncodeURL = EncodeURL & "H" & Right$("0" & Hex(Asc(sChar)), 2)
        End Select
    Next
End Function

Public Function GetFreeName(ByVal sFileName As String, ByVal sDestFolder As String)
    Dim sName As String, sExt As String, lNro As Long
    If Right$(sDestFolder, 1) <> "\" Then sDestFolder = sDestFolder & "\"
    sFileName = GetFileTitle(sFileName)
    If FileExists(sDestFolder & sFileName) Then
        sName = GetFileName(sFileName)
        sExt = GetFileExtention(sFileName, True)
        lNro = 1
        Do While FileExists(sDestFolder & sName & lNro & sExt)
            lNro = lNro + 1
        Loop
        GetFreeName = sName & lNro & sExt
    Else
        GetFreeName = sFileName
    End If
    
End Function

Public Function GetFileFolder(ByVal sPath As String) As String
    GetFileFolder = Left$(sPath, InStrRev(sPath, "\"))
End Function


Public Function GetFileName(ByVal sPath As String) As String
    Dim lPos As Long, sName As String
    lPos = InStrRev(sPath, ".")
    If lPos Then
        sName = Left$(sPath, lPos - 1)
        lPos = InStrRev(sName, "\")
        If lPos Then
           sName = Mid$(sName, lPos + 1)
        End If
        GetFileName = sName
    Else
        lPos = InStrRev(sName, "\")
        If lPos Then
           GetFileName = Mid$(sName, lPos + 1)
        End If
    End If
End Function

Public Function GetFileTitle(ByVal sPath As String) As String
    Dim lPos As Long
    lPos = InStrRev(sPath, "\")
    If lPos Then
       GetFileTitle = Mid$(sPath, lPos + 1)
    Else
        GetFileTitle = sPath
    End If
End Function

Public Function GetFileExtention(ByVal sPath As String, Optional InlcudePoint As Boolean) As String
    Dim lPos As Long
    lPos = InStrRev(sPath, ".")
    If lPos Then
       GetFileExtention = Mid$(sPath, lPos + IIf(InlcudePoint, 0, 1))
    End If
End Function

Public Function FileExists(ByVal sPath As String) As Boolean
    FileExists = Len(Dir(sPath)) > 0
End Function
    
Public Function FolderExists(ByVal sPath As String) As Boolean
    FolderExists = Len(Dir(sPath, vbDirectory)) > 0
End Function

Public Function ReadFileXML(sFile As String) As String
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    With xmlDoc
        .SetProperty "ProhibitDTD", False
        '.SetProperty "ResolveExternals", True
        .validateOnParse = True
        .async = False
        .Load (sFile)
        ReadFileXML = .xml
    End With
    Set xmlDoc = Nothing
End Function

Public Function WriteFile(sData As String, sDestFile As String) As Boolean
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    With xmlDoc
        .SetProperty "ProhibitDTD", False
        .validateOnParse = True
        .async = False
        .loadXML sData
        .Save sDestFile
    End With
    Set xmlDoc = Nothing
End Function

Public Function ReadFileText(sFile As String) As String
    Dim FF As Integer
    FF = FreeFile
    Open sFile For Input As FF
    ReadFileText = Input$(LOF(FF), #FF)
    Close FF
End Function
