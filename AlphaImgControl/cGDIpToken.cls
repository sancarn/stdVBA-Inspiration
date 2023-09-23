VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'  -----======== PURPOSE: Manage GDI+ token & related properties ========-----
' Class Instancing = Private (not shared to users)
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.


' HOW TO USE THIS CLASS
' This class must be created first before any other GDI+ classes call their functions

' Only ONE instance of this class needs to be created and should exist until you close your project
' Recommend making this public in your main form or in a module

' Note: When to destroy the GDI+ token it created.
' VB doesn't always unload classes right away. I have noticed crashes during
' design time when this class unloads later than "normal", whatever normal is.
' I have also seen in some cases (usually group projects) where this class will call its
'   terminate event before other classes referencing it call their terminate event. Makes it
'   really impossible to assume that when this class terminates all other classes referencing
'   it terminated first. Therefore, the Terminate event cannot destroy the token.
' So to prevent these problems, a 2-step solution was employed.

' Create an maintain a reference to the class during your Load/Initialize events & keep reference count
'       Public g_TokenClass As cGDIpToken   ' < in a module
'
'       Private m_Token As cGDIpToken       ' < in your form/class/uc declarations section
'                                           ' < in your load/initialize event
'       If g_TokenClass Is Nothing Then Set g_TokenClass = New cGDIpToken
'       Set m_Token = g_TokenClass: m_Token.AddRef
'                                           ' < in your terminate/unload event
'       m_Token.RemoveRef
'       Set m_Token = Nothing
' This class will only destroy its token when the ref count reaches zero
' I also suspect this is an uncompiled issue only. When compiled, expect no such problems.

Option Explicit

Private Declare Sub AVIFileInit Lib "avifil32.dll" ()
Private Declare Sub AVIFileExit Lib "avifil32.dll" ()

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private m_Token As Long
Private m_Instances As Long
Private m_Version As Single
Private m_AVIinit As Long
Private m_DPIscale As Single

Public Property Get Token() As Long
    If Not m_Token = -1& Then Token = m_Token
End Property

Public Property Get Version() As Single
    Version = m_Version
End Property

Public Sub AddRef()         ' should be called when you ref the class
    If m_Token = 0& Then CreateToken
    m_Instances = m_Instances + 1&
End Sub

Public Sub RemoveRef()      ' should be called just before you unreference the class
    m_Instances = m_Instances - 1&
    If m_Instances = 0& Then DestroyToken
End Sub

Public Function AVIClassesInitiated() As Boolean
    If m_AVIinit = 0& Then
        On Error Resume Next
        Call AVIFileInit
        If Err Then
            Err.Clear
            m_AVIinit = -1&
        ElseIf Err.LastDllError Then
            m_AVIinit = -1&
        Else
            m_AVIinit = 1&
        End If
    End If
    AVIClassesInitiated = (m_AVIinit = 1&)
    
End Function

Public Property Get DPIscale() As Single
    DPIscale = m_DPIscale
End Property

Private Sub CreateToken()

    Dim GSI As GdiplusStartupInput
    Dim pa As Long, hMod As Long
    
    On Error Resume Next
    GSI.GdiplusVersion = 1&
    Call GdiplusStartup(m_Token, GSI)
    If m_Token Then                                         ' is v1.1 GDI+?
        hMod = LoadLibrary("gdiplus.dll")                   ' test it
        If hMod Then
            pa = GetProcAddress(hMod, "GdipDrawImageFX")    ' if it has this function, then at least GDI+ v1.1
            If pa Then m_Version = 1.1!
            FreeLibrary hMod
        End If
        If m_Version = 0! Then m_Version = 1!
    Else
        m_Token = -1&                                       ' prevent calling this time & time again if failure
    End If
    Set g_MouseExitClass = New cMouseExit
    modCommon.CreateCustomClipboardFormat

End Sub

Private Sub DestroyToken()
    If m_AVIinit = 1& Then AVIFileExit
    Set g_AsyncController = Nothing
    If Me.Token Then GdiplusShutdown m_Token
    m_Token = 0&
    m_Version = 0!
    m_AVIinit = 0&
End Sub

Private Sub Class_Initialize()
    CreateToken
    m_DPIscale = (1440! / Screen.TwipsPerPixelX) / 96
End Sub

