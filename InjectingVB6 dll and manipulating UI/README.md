# [Injecting a VB6 dll and manipulating the UI](https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI&p=5526939&viewfull=1#post5526939)

Hi guys.

I am new here. I'll try to keep it simple.

Mission: Manipulate a legacy VB6 app UI.

The strategy:
1. Create a VB6 ActiveXDll
2. inject it into the process using SetWindowHookex
3. set a timer for handling execution outside the handler itself
4. Do as I please with the UI

I have reached step 4 after much effort, Yet I find no way to access the forms.
I have read that the 'Forms' collection is not accessible via activex dll. as a matter of fact - the entire VB.Global is not accessible.

So my friends - any VB6 guru here can show me the way?
Given that I have an dll that is running - is it possible for me to gain access to the open forms?

(I thought about obtaining the IAccessible pointer and then try to manipulate it into a form "pointer" yet I find no clues for that, nor do I know if the form class itself is COM accessible, or what is it's COM GUID if any.)

Thank you!

## [Posted by The Trick](https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI&p=5527040&viewfull=1#post5527040)

> how can I grab hold of any pointer to an open form, or even all the open forms.

The simpliest way is to access the `VB.Global.Forms` collection of a project and then get all the forms. To get `VB.Global` object you can inspect `VBHeader`.`lpProjectData->lpExternalTable` items and search for items with the tag (at 0 offset) equals to `6`. The item value (at 4 offset) is a pointer to an object descriptor which consists of a pointer to the object CLSID (at zero offset) and pointer to the Global object itself (at 4 offset).

So you can extract this pointer and call AddRef method then you can cast this pointer to a `VB.Global` object variable. From this moment, the project methods and properties (like `App`/`Screen`/`Forms` etc.) are available to you.


## [Research by dz32](https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI&p=5527217&viewfull=1#post5527217)

so I played around a bit with this based on the tips from trick. I will describe the complete process.

first I made an exe which set a textbox to `hex(objptr(vb.global))` so I knew what value I should be searching for.
then I opened that exe up in vbdec to look at the vb headers and offsets.
then I attached ollydbg to the process and started probing those offsets

I ended up finding the path trick mentioned to the `vb.global` pointer. Set a memory write breakpoint on that offset and reloaded.
its set in `MSVBVM60.6601802F 8908 MOV DWORD PTR DS:[EAX],ECX` which is `_TipRegAppObject`

(vbruntime: `EEBEB73979D0AD3C74B248EBF1B6E770` with symbols: http://sandsprite.com/vb-reversing/files/msvbvm60.zip )


Then I loaded up a complex project set a breakpoint at my runtime offset (6601802F) to catch the value it will write
for me it was (These offsets are relative to my exe but included for clarity)

offset to write to: 0047D11C (eax)
value (ecx): 02BF1294 -> 660130D0 MSVBVM60.??_7CVBApplication@@6BVBGlobal@@@

Next I examine the structures in vbdec to get offsets to walk in olly

```
VB Header
VA          rel offset    name                  value
4048A8   0x30        aProjectInfo           40540C

VB Header.aProjectInfo @40540C
405640   0x234   aExternalTable         404FF8
405644   0x238   ExternalCount          3E

at 404FF8 I find data such as 

00404FF8  00000007
00404FFC  00429470  PDFStrea.00429470
00405000  00000007
00405004  0042942C  PDFStrea.0042942C
00405008  00000007
0040500C  004293E8  PDFStrea.004293E8
... many more entries ....
004051A8  00000006                 <--- flag we are looking for
004051AC  0041CC7C  PDFStrea.0041CC7C     <---

Looking at 0041CC7C

	0041CC7C  0041C95C  PDFStrea.0041C95C  -> offset + 0 points to clsid/IID
	0041CC80  0047D11C  PDFStrea.0047D11C  <-- offset+ 4 value we saw written to in _TipRegAppObject 

looking at 47D11C
	0047D11C  = 02BF1294 -> 660130D0  MSVBVM60.??_7CVBApplication@@6BVBGlobal@@@

Then from there you can work with the vb.global object (sorry didnt switch to relative offsets before closing process)

660130D0 >660E2074  MSVBVM60.?QueryInterface@CVBApplication@@UAGJABU_GUID@@PAPAX@Z
660130D4  6601808A  MSVBVM60.?AddRef@CVBApplication@@UAGKXZ
660130D8  66028553  MSVBVM60.?Release@CVBApplication@@UAGKXZ
660130DC  6605CE6B  MSVBVM60.?Load@CVBApplication@@UAGJPAUIDispatch@@@Z
660130E0  6605CEE2  MSVBVM60.?Unload@CVBApplication@@UAGJPAUIDispatch@@@Z
660130E4  66026F9D  MSVBVM60.?get_App@CVBApplication@@UAGJPAPAV_AppObject@@@Z
660130E8  6603D750  MSVBVM60.?get_Screen@CVBApplication@@UAGJPAPAV_ScreenObject@@@Z
660130EC  660489C7  MSVBVM60.?get_Clipboard@CVBApplication@@UAGJPAPAV_ClipboardObject@@@Z
660130F0  660E1EC6  MSVBVM60.?get_Printer@CVBApplication@@UAGJPAPAV_PrinterObject@@@Z
660130F4  660E1EF8  MSVBVM60.?putref_Printer@CVBApplication@@UAGJPAV_PrinterObject@@@Z
660130F8  66048EAE  MSVBVM60.?get_Forms@CVBApplication@@UAGJPAPAUIDispatch@@@Z
660130FC  660E1EE0  MSVBVM60.?get_Printers@CVBApplication@@UAGJPAPAUIDispatch@@@Z
66013100  660E1FA0  MSVBVM60.?LoadResStringOld@CVBApplication@@UAGJFPAPAG@Z
66013104  6604A7C3  MSVBVM60.?LoadResPicture@CVBApplication@@UAGJUtagVARIANT@@FPAPAUIPictureDisp@@@Z
66013108  6604AB06  MSVBVM60.?LoadResData@CVBApplication@@UAGJUtagVARIANT@@0PAU2@@Z
6601310C  660E1FD9  MSVBVM60.?LoadPictureOld@CVBApplication@@UAGJUtagVARIANT@@PAPAUIPictureDisp@@@Z
66013110  6604776E  MSVBVM60.?SavePicture@CVBApplication@@UAGJPAUIPictureDisp@@PAG@Z
66013114  6602C9D8  MSVBVM60.?LoadPicture@CVBApplication@@UAGJUtagVARIANT@@0000PAPAUIPictureDisp@@@Z
66013118  6604A683  MSVBVM60.?LoadResString@CVBApplication@@UAGJJPAPAG@Z
6601311C  660E1F71  MSVBVM60.?get_Licenses@CVBApplication@@UAGJPAPAULicenses@@@Z

0041C95C  23 3D FB FC FA A0 68 10 A7 38 08 00 2B 33 71 B5  = {FCFB3D23-A0FA-1068-A738-08002B3371B5} = VBEGlobal Clsid
0041C96C  22 3D FB FC FA A0 68 10 A7 38 08 00 2B 33 71 B5 = {FCFB3D22-A0FA-1068-A738-08002B3371B5} = VBGlobal IID
```

Misc notes:
Since you are working with a static exe that wont be recompiled, your VB Header.aProjectInfo.aExternalTable offset and count will be static

VbHeader.aExternalComponentTable was = VB Header.aProjectInfo.aExternalTable for a small native exe, but not for a large pcode exe
so I would stick with the latter

In the large exe my runtime breakpoint hit twice, once for the exe, and once again for another vb6 ocx control it loaded. values written were different

in my test exe I could not set tmp as object = vb.global got type mismatch

## [Example by dz32](https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI&p=5527276&viewfull=1#post5527276)

add a listbox and 2 command buttons to your form
fully working

```vb
Option Explicit

Dim WithEvents btn As CommandButton

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub GetMem4 Lib "msvbvm60.dll" (ByVal lAddress As Long, var As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

Private Type extEntry
    flag As Long
    offset As Long
End Type

Private Type extEntryTarget
    clsidIIDStructOffset As Long
    lpValue As Long
End Type

Dim readyToReturn As Boolean

Private Sub Command1_Click()
    readyToReturn = True
End Sub

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()

'    Dim v  As VB.Global
'    MsgBox v.App.EXEName 'variable not set..this is good, sanity check
'    End
    
    If IsIde Then
        MsgBox "Only run compiled"
        End
    Else
        Me.Caption = "It worked!"
        Command1.Caption = "Continue"
        Command2.Caption = "End"
        List1.Move 0, 0, 8000, 8000
        Me.Width = 8400
        Me.Height = 9100
        Command1.Move Me.Width - 3000, 8025
        Command2.Move Me.Width - 1500, 8025
        DoTheTrickStuff
    End If
    
End Sub

Sub xx(msg As String, v As Long, Optional wait As Boolean = False)
    List1.AddItem msg & ": " & Hex(v) & IIf(wait, "  --- Waiting for continue click...", Empty)
    List1.Refresh
    DoEvents
    If wait Then
        readyToReturn = False
        While Not readyToReturn
            DoEvents
        Wend
    End If
End Sub


Private Sub DoTheTrickStuff()
   
    Dim vbHeader As Long, lpProjectData As Long, projectDataPointerValue As Long
    Dim lpExternalTable As Long, externalTablePointerValue As Long
    Dim externalCount As Long, lpExternalCount As Long
    Dim i As Integer, e As extEntry, myCaption As String
    Dim globalPtr As Long, lpGlobalPtr As Long, globalClone As VB.Global

    Me.Visible = True
    xx "The pointer I am after", ObjPtr(VB.Global)
     
    vbHeader = GetVBHeader()
    xx "Header", vbHeader
    
    lpProjectData = vbHeader + &H30
    GetMem4 ByVal lpProjectData, projectDataPointerValue
    xx "lpProject", projectDataPointerValue

    lpExternalCount = projectDataPointerValue + &H238
    xx "lpExternalCount", lpExternalCount
    
    GetMem4 ByVal lpExternalCount, externalCount
    xx "ExtCount", externalCount
    
    lpExternalTable = projectDataPointerValue + &H234
    xx "lpExtTable", lpExternalTable
    
    GetMem4 ByVal lpExternalTable, externalTablePointerValue
    xx "ExtTable", externalTablePointerValue
    
    For i = 0 To externalCount
        CopyMemory ByVal VarPtr(e), ByVal externalTablePointerValue, 8
        
        Debug.Print i & ") " & Hex(externalTablePointerValue) & " " & e.flag & " " & Hex(e.offset)
        
        If e.flag = 6 Then 'Tag of global object
            xx "Found flag", e.flag
        
            GetMem4 ByVal e.offset + 4, lpGlobalPtr
            xx "lpGlobalPointer", lpGlobalPtr
            
            GetMem4 ByVal lpGlobalPtr, globalPtr
            xx "GlobalPointer", globalPtr
            
            If globalPtr = ObjPtr(VB.Global) Then
                List1.AddItem "Success!!"
            Else
                List1.AddItem "FAIL!!"
            End If
            
            Set globalClone = GlobalFromPointer(ByVal globalPtr)
            xx "ObjPtr(globalClone)", ObjPtr(globalClone) ', True

            myCaption = globalClone.Forms(0).Caption
            List1.AddItem "globalClone.Forms(0).Caption = " & myCaption
            globalClone.Forms(0).Caption = "Are you sure?!"
            
            Exit For
        End If
        
        externalTablePointerValue = externalTablePointerValue + 8
    Next
    
    List1.AddItem "Complete"
    
End Sub

Private Function GlobalFromPointer(ByVal ptr&) As VB.Global
  ' this function returns a reference to our form class
  
  ' dimension an object variable that we can copy the
  ' passed pointer into
  Dim globalClone As VB.Global
    
  ' use the CopyMemory API function to copy the
  ' long pointer into the object variable.
  CopyMemory globalClone, ptr, 4&
  Set GlobalFromPointer = globalClone
  CopyMemory ptr&, 0&, 4&
  
End Function

Public Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    CopyMemory obj, pObj, LenB(obj) 'not size 4!
    Set ObjFromPtr = obj
    CopyMemory obj, 0&, LenB(obj) 'not size 4!
End Function

' // Get VBHeader structure
Private Function GetVBHeader() As Long
    Dim ptr     As Long
    Dim hModule As Long
    hModule = GetModuleHandle(ByVal "Project1.exe")
    ' Get e_lfanew
    GetMem4 ByVal hModule + &H3C, ptr
    ' Get AddressOfEntryPoint
    GetMem4 ByVal ptr + &H28 + hModule, ptr
    ' Get VBHeader
    GetMem4 ByVal ptr + hModule + 1, GetVBHeader
    
End Function

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width
End Sub
```


## [Example by The Trick](https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI&p=5527341&viewfull=1#post5527341)
Posted by:  The trick
Forum link: https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI

I've made special example where you can access to Forms collection (and controls) from your code like that:

```vb
' //
' // Access Forms collection of a VB6 executable
' // by The trick 2021
' //

Option Explicit

Private Type STARTUPINFO
    cb              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess        As Long
    hThread         As Long
    dwProcessId     As Long
    dwThreadId      As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" _
                         Alias "CreateProcessW" ( _
                         ByVal lpApplicationName As Long, _
                         ByVal lpCommandLine As Long, _
                         ByRef lpProcessAttributes As Any, _
                         ByRef lpThreadAttributes As Any, _
                         ByVal bInheritHandles As Long, _
                         ByVal dwCreationFlags As Long, _
                         ByRef lpEnvironment As Any, _
                         ByVal lpCurrentDirectory As Long, _
                         ByRef lpStartupInfo As STARTUPINFO, _
                         ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32" ( _
                         ByVal hProcess As Long, _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" _
                    Alias "GetStartupInfoW" ( _
                    ByRef lpStartupInfo As STARTUPINFO)

Sub Main()
    Dim tSI             As STARTUPINFO
    Dim tPI             As PROCESS_INFORMATION
    Dim cVBGetGlobal    As Object
    Dim cForms          As Object
    Dim frmMain         As Object
    
    InitializeInjectLibrary
    
    tSI.cb = Len(tSI)
    
    GetStartupInfo tSI
    
    If CreateProcess(StrPtr(App.Path & "\..\dummy\dummy.exe"), 0, ByVal 0&, ByVal 0&, 0, 0, ByVal 0&, 0, tSI, tPI) = 0 Then
        MsgBox "CreateProcess failed"
        Exit Sub
    End If
    
    WaitForInputIdle tPI.hProcess, -1
    
    CloseHandle tPI.hProcess
    CloseHandle tPI.hThread
    
    Set cVBGetGlobal = CreateVBObjectInThread(tPI.dwThreadId, App.Path & "\..\dll\GetVBGlobal.dll", "CExtractor")
    Set cForms = cVBGetGlobal.Forms
    
    Set frmMain = cForms(0)
    
    ' // Change back color of picturebox
    frmMain.Controls("picTest").BackColor = vbRed
    
    ' // Draw line on picturebox
    frmMain.Controls("picTest").Line (0, 0)-Step(100, 50), vbGreen, BF
    
    UninitializeInjectLibrary
    
End Sub
```

The executable contains a PictureBox (picTest) and you can acccess to it. You can access to a VB.Global class from a AX-Dll without any restrictions (because no marshaling is needed).

https://github.com/thetrik/InjectAndManipulateVB


## [2nd Example by The Trick](https://www.vbforums.com/showthread.php?892441-Injecting-a-VB6-dll-and-manipulating-the-UI&p=5527342&viewfull=1#post5527342)

> The pointer I get back is not equal to the value I get from ObjPtr(VB.Global)
> no wonder it crashes...

The proper code:

```vb
' //
' // Access to Forms collection of a VB6-EXE
' // By The trick 2021
' //

Option Explicit

Private Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint  As Long
End Type

Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleInformation Lib "psapi" ( _
                         ByVal hProcess As Long, _
                         ByVal hModule As Long, _
                         ByRef lpmodinfo As MODULEINFO, _
                         ByVal cb As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" _
                         Alias "__vbaObjSetAddref" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" ( _
                    ByRef Addr As Any, _
                    ByRef Dst As Any)
        
Public Property Get Forms() As Object
    Dim hModule     As Long
    Dim pVbHdr      As Long
    Dim tModInfo    As MODULEINFO
    Dim pProjData   As Long
    Dim pExtTable   As Long
    Dim lExtCount   As Long
    Dim lIndex      As Long
    Dim lTag        As Long
    Dim pObjDesc    As Long
    Dim pVBGlobal   As Long
    Dim cVBGlobal   As VB.Global
    
    hModule = GetModuleHandle(0)
    
    If GetModuleInformation(GetCurrentProcess(), hModule, tModInfo, Len(tModInfo)) = 0 Then
        MsgBox "GetModuleInformation failed", vbCritical
        Exit Property
    End If
    
    GetMem4 ByVal tModInfo.EntryPoint + 1, pVbHdr
    GetMem4 ByVal pVbHdr + &H30, pProjData
    GetMem4 ByVal pProjData + &H234, pExtTable
    GetMem4 ByVal pProjData + &H238, lExtCount
    
    For lIndex = 0 To lExtCount - 1
        
        GetMem4 ByVal pExtTable + lIndex * 8, lTag
        
        If lTag = 6 Then
            
            GetMem4 ByVal pExtTable + lIndex * 8 + 4, pObjDesc
            GetMem4 ByVal pObjDesc + 4, pVBGlobal
            GetMem4 ByVal pVBGlobal, pVBGlobal
            
            vbaObjSetAddref cVBGlobal, ByVal pVBGlobal
            
            Set Forms = cVBGlobal.Forms
            
        End If
        
    Next
    
End Property
```