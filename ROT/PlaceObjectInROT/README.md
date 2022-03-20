# Register Forms to ROT

Hi, this is a continuation of some code I posted in the past see Here ... This new addition uses file monikers to register objects in the Running Object Table and hence makes it possible to access UserForms from remote processes via the standard `GetObject` vba function.

The advantage of using this method over the one showed in the other thread is that it can be applied to multiple userforms at once.

You can reference the remote UserForm via its Name Property or its Caption - You decide by setting the Optional Boolean CallByCaption parameter in the PutInROT SUB.

Here is an example of how to reference a remote userform :
Set oRemoteUserForm = GetObject("MonikerTest.UserForm1")

MonikerTest [name of the workbook (Without the file extension) that contains the userForm] + "." + UserForm1 [ name or caption of the UserForm].

##  Code authorage
### Original thread

Author: Wqweto
Link:   https://www.vbforums.com/showthread.php?879529-project-one-workng-with-project2&p=5422507&viewfull=1#post5422507

### VB7 (and moniker) upgrade

Author: Jaafar Tribak
Link:   https://www.mrexcel.com/board/threads/reference-and-remotely-manipulate-userforms-loaded-in-seperate-workbooks-or-in-seperate-excel-instances-via-file-monikers.1161038/#post-5634620


## Missing from example

It turns out that by using `CLSIDFromProgID()` one can register new ProgIDs without needing to register a DLL. As such, the following code:

```vb
Option Explicit

Private Type GUIDs
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' Declares needed to register object in the ROT
Private Const ACTIVEOBJECT_STRONG = 0
Private Const ACTIVEOBJECT_WEAK = 1
Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As LongPtr, ByRef rclsid As GUIDs) As Long
Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, ByRef pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32.dll" (ByVal pUnk As IUnknown, ByRef rclsid As GUIDs, ByVal dwFlags As Long, ByRef pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32.dll" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long

Private OLEInstance As Long

Private Sub Class_Initialize()
    Dim typGUID As GUIDs
    Dim lp As Long

    OLEInstance = 0
    ' This code is responsible for creating the entry in the rot
    lp = CLSIDFromProgID(StrPtr("Test.MyClass"), typGUID)
    If lp = 0 Then
        lp = RegisterActiveObject(Me, typGUID, ACTIVEOBJECT_WEAK, OLEInstance)
    End If
End Sub

Private Sub Class_Terminate()
    ' Once we are done with the main program, lets clean up the rot
    ' by removing the entry for our ActiveX Server
    If OLEInstance <> 0 Then
        RevokeActiveObject OLEInstance, 0
    End If
    CoDisconnectObject Me, 0
End Sub
```

Will produce a scenario where you can use:

```vb
Dim obj as object: set obj = GetObject("Test.MyClass")
```

