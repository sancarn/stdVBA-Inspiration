# API CallBacks using an Object's Procedure

* Author: Elroy
* Attributions: Huge thanks goes out to The Trick, Wqweto, Dz32, Fafalone, and LaVolpe. They all provided insights, inspiration, and code toward what you see in this thread.
* Thread: https://www.vbforums.com/showthread.php?898136-API-CallBacks-Using-an-Object-s-Procedure

**Prerequisite:** It's assumed you know what an API CallBack is. If not, you really don't have to worry about any of this.

----

This has been done before, and this is just my version of it. Also, I've documented it extensively so that anyone who's interested should be able to "think through" everything I've done.

Just to say it, it's a lot of code for a sort of small thing. I mean, we can always put our CallBacks in BAS modules, and they work just fine. But this is something I've wanted for years, so that I can create object modules (Classes, Forms, UserControls, PropertyPages, & DataReports) that use CallBacks that are all wrapped into their respective code areas.

This can also be used for subclassing, but that's not what it is. If you think about it, all subclassing does is make a "CallBack" with everything that's coming through the message pump for any particular window.

Also, just to try and eliminate confusion, I'm talking about a specific VB6 project and it's associated source code (and those source code objects). I'm **not** talking about referenced DLLs (ActiveX or otherwise). Some of this code could be used for those things, but that's not what this thread is about.

Downsides to the way I've done it:

* There's no IDE protection. But that's only a problem if you try to trace through the CallBack procedure and/or put breakpoints in it. Otherwise, everything should be fine regarding the rest of your code. Again, this isn't subclassing, but it could be used for subclassing. When used with subclassing, those other uses could provide IDE protection. That's up to the subclassing, and not these routines.
* I've limited the CallBack function to being a "Function" (no Subs, nor Properties), and it must have 1, 2, 3, or 4 arguments. Furthermore, those arguments must be either ByRef or 4-byte ByVal arguments. And the return must also be 4-bytes (preferably a Long). But this covers the vast majority of Microsoft API CallBacks.
* I haven't tested an object that has another object "Implemented". As soon as I get that tested, I'll report back on that one.

Upsides to the way I've done it:

* There's no modification to the actual code of the object. Basically, all of this is just a "wrapper" that turns a call for a standard BAS procedure into a call for a COM (object) procedure. But that's much easier said than done.
* It's (hopefully) extremely well documented so that anyone who wants to use it can see precisely what's going on.
* It works equally well when running in the IDE or compiled as either p-code or machine code.
* There's no special order in which the CallBack procedures need to appear in your code.

The challenges. There were many, but I'll go through them one-by-one, providing solutions for each as we go. And later on, I'll provide a few projects with complete examples that you could use to "mold" into your own uses and applications.

## Challenge #1: Are we dealing with a VB6 object that has code, or some other kind of object?

In many of the procedures herein, an object is passed in. I didn't want to have any problems if someone tried to pass in a non-code object (such as a `StdFont`, `StdPicture`, or `Collection` object, etc., etc.). I also wanted to assure that the object was instantiated.

This problem is solved with the following function and API call: (Specific thanks to The Trick & Wqweto for this.)

```vb
Private Declare Function vbaCheckType Lib "msvbvm60" Alias "__vbaCheckType" (ByVal pObj As Any, ByRef pIID As Any) As Boolean

Public Function ObjectIsVb6ComCodeModule(ByRef o As IUnknown) As Boolean
    ' If it's an instantiated Class, Form, UC, PropPage, DataReport, returns TRUE, else FALSE.
    If ObjPtr(o) = 0& Then Exit Function                    ' Make sure "something" is instantiated.
    Dim aGUID(1&) As Currency                               ' Just to get 16 easily accessible bytes.
    aGUID(0&) = 128347367577987.1845@                       ' Const AreYouABasicInstance As String = "{0B6C9465-D082-11CF-8B4F-00A0C90F2704}"
    aGUID(1&) = 29922525889064.5387@                        ' turned into two numbers stuffed into our Currency array.
    ObjectIsVb6ComCodeModule = vbaCheckType(o, aGUID(0))    ' Check and see if we are this "TYPE" (Class, Form, UC, PropPage, or DataRep).
End Function
```

## Challenge #2: The AddressOf operator doesn't work on object procedures

Therefore, we've got to figure out how to get the address of procedures in our object's code. This immediately takes us to a discussion of vTables. A vTable is nothing but a list of addresses. These are all the memory addresses of all the procedures in our object's code. All VB6 objects contain "interface" procedures for the `IUnknown` and `IDispatch` procedures (3 for `IUnknown` and 4 for `IDispatch`), but we really don't need to worry about those. All but CLS based objects also have "hidden" procedures that are placed in by VB6 to make the object "work". Fortunately, we don't have to worry about those either. Following those, there are two more groups of procedures in an object's vTable: 1) Public procedures (including `Get`/`Let`/`Set` procedures built for Public variables) that we've coded, and 2) `Private`/`Friend` procedures that we've coded.

(Just to note, `Private` variables aren't in the vTable, and the compiler makes no distinction between `Private` or `Friend` procedures with respect to the vTable. Also, in terms of the vTable, there's no distinction between events and all other procedures. If you code up an event, it's in the vTable.)

So, our challenge #2 transforms into figuring out which address in the vTable is pointing toward our desired CallBack procedure. This basically breaks down into two parts: 1) getting the address of the vTable, and 2) getting an offset into the vTable for our CallBack procedure. Getting the address of any vTable is easy:

```vb
Dim pVtable  As Long:   GetMem4 ByVal ObjPtr(o), pVtable        ' Get pointer to start of vTable.
```

Now, we just need to get our offset for the vTable's address that we actually want. I tried several approaches (some of which I'll discuss later on), but I finally settled on one that I particularly like. Supply it the object, and the procedure name, and it returns the vTable offset as well as the number of arguments passed into the procedure. Here it is: (Thanks to The Trick & Dz32 on this one.)

```vb
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByRef lpString As Any) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByRef psz As Any, ByVal lSize As Long) As String

Public Function VtableOffsetForVb6ComMethod(ByVal o As Object, ByVal sMethodName As String, Optional ByRef lArgCount As Long) As Long
    ' Searches PUBLIC methods.  It "could" find Private & Friend, but only in the IDE, not compiled.
    ' Does NOT search properties (i.e., Public variables or Get/Let/Set procedures).
    ' Returns an OFFSET address ready to be added to the vTable address.
    ' Optional bbArgCount return is a count of passed in arguments.
    '       ByRef/ByVal doesn't matter, and TYPE doesn't matter.
    '       If it's a Function the return isn't counted.
    '
    ' If it can't be found, ZERO is returned.
    '
    If Not ObjectIsVb6ComCodeModule(o) Then Exit Function                           ' Make sure we're dealing with a VB6 COM-code object.
    sMethodName = UCase$(sMethodName)
    '
    Dim pVTbl       As Long:    GetMem4 ByVal ObjPtr(o), pVTbl                      ' Pointer to vTable.
    Dim pObjInfo    As Long:    GetMem4 ByVal pVTbl - 4&, pObjInfo                  ' Pointer to tObjectInfo structure.
    Dim pPubDesc    As Long:    GetMem4 ByVal pObjInfo + &H18&, pPubDesc            ' tObjectInfo.aObject which points to tObject structure.
    Dim pPrivDesc   As Long:    GetMem4 ByVal pObjInfo + &HC&, pPrivDesc            ' tObjectInfo.lpPrivateObject which points to tPrivateObj structure.
    '
    If pPrivDesc = 0& Then Exit Function                                            ' Just a double-check.
    '
    Dim lIndex      As Long
    Dim pName       As Long
    '
    ' Search the procedures within the module.
    Dim pMethDesc   As Long
    Dim iMethOffset As Integer
    Dim bbArgs      As Byte
    Dim lMethodsCnt As Long:    GetMem2 ByVal pPubDesc + &H1C&, lMethodsCnt         ' tObject.ProcCount value.
    Dim pNames      As Long:    GetMem4 ByVal pPubDesc + &H20&, pNames              ' tObject.aProcNamesArray which points to an array of name pointers.
    Dim pMethodsPtr As Long:    GetMem4 ByVal pPrivDesc + &H18&, pMethodsPtr        ' tPrivateObj.lpFuncTypeInfo which points to an array of pointers.
    '
    ' Loop through methods and see if we can find the one we want.
    For lIndex = 0& To lMethodsCnt - 1&
        GetMem4 ByVal pMethodsPtr + lIndex * 4&, pMethDesc                          ' From the array, getting a pointer to a method structure (tMethInfo).
        If pMethDesc Then                                                           ' Not sure if this ever returns zero, maybe for "Private" methods?
            GetMem2 ByVal pMethDesc + 2&, iMethOffset                               ' Out of tMethInfo structure.
            GetMem1 ByVal pMethDesc, bbArgs                                         ' First two bits of bbArgs are: set=3, get=1, let=2, method=0 (Sub or Fn).
            If (bbArgs And CByte(3)) = CByte(0) Then                                ' Make sure it's a method.
                If iMethOffset And 1 Then                                           ' First bit, 1=Public.
                    GetMem4 ByVal pNames + lIndex * 4&, pName                       ' Dig pointer to method name from array of name pointers.
                    If sMethodName = UCase$(SysAllocStringByteLen(ByVal pName, lstrlenA(ByVal pName))) Then
                        VtableOffsetForVb6ComMethod = CLng(iMethOffset And &HFFFC)  ' First two bits are something else (first is Public=1,Private=0).
                        Dim bbFlags As Byte: GetMem1 ByVal pMethDesc + 1&, bbFlags  ' Both bbArgs & bbFlags out of tMethInfo structure.
                        bbFlags = bbFlags And CByte(1)                              ' 0 (no return), 1 (return).
                        lArgCount = CLng(bbArgs \ CByte(4) - bbFlags)               ' Calculate arguments, excluding any return argument.  Tested for vbGet, vbLet, vbSet, vbMethod (both Function & Sub).
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    '
    ' Return zero if not found.
End Function
```

This `VtableOffsetForVb6ComMethod` uses the above `ObjectIsVb6ComCodeModule` call, which is absolutely essential for what this `VtableOffsetForVb6ComMethod` does. I've documented this as best I could. It's essentially digging into compiled structures of any VB6 COM object. I've tried to identify these structures as I dig through them as best I can, and there is some (limited) documentation out on the web regarding these. Basically, the code is figuring out whether the procedures are Public vs Private, what type they are (`Sub`, `Function`, `Property`), what their name is, and how many arguments into it there are. But the main thing we're after is the vTable offset, and that's the return value (if the procedure is found).

Note that this only works for `Public` procedures. Below, I'll discuss an alternative approach that can be taken if we'd like to keep our CallBack procedures `Private`.

----

So, from where we are now, we can write the following function:

```vb
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long

Public Function Vb6ComCodeObjectAddressOf(ByVal o As Object, ByVal sMethodName As String, Optional ByRef lArgCount As Long) As Long
    ' lArgCount is an optional RETURN.  ByRef/ByVal doesn't matter, and TYPE doesn't matter.  If it's a Function the return isn't counted.
    ' The caller is responsible for knowing how to use the returned address, or crash may result.
    ' See more notes within VtableOffsetForVb6ComMethod procedure.
    '
    ' Returns ZERO if not found, or it's not a PUBLIC method.
    '
    lArgCount = 0&                                                  ' Zero out anything there to start.
    Dim iVoffset As Long:   iVoffset = VtableOffsetForVb6ComMethod(o, sMethodName, lArgCount)
    If iVoffset = 0& Then Exit Function                             ' This checks to make sure the object is okay.
    Dim pVtable  As Long:   GetMem4 ByVal ObjPtr(o), pVtable        ' Get pointer to start of vTable.
    pVtable = pVtable + iVoffset                                    ' Address to our method in the vTable.
    GetMem4 ByVal pVtable, Vb6ComCodeObjectAddressOf                ' Pointer into actual code (the method).
End Function
```

And this is basically the same thing as the AddressOf operator, but it works for any of the VB6 COM/code objects. One might think that we're home free after this, but we're actually far from it. VB6 COM/code procedures have "under the hood" stuff that BAS procedures don't have. And if we attempt to use this `Vb6ComCodeObjectAddressOf` for a CallBack address, we'll just crash.

## Challenge #3: Making a procedure in a VB6 object "look like" a regular BAS procedure.

In a BAS procedure, what you see is what you get. However, that's not true for a COM procedure. Let's take an example to illustrate. Here's a typical BAS procedure that might be used as a CallBack:

```vb
Public Function EnumResNameProc(ByVal hModule As Long, _
                                ByVal lpszType As Long, _
                                ByVal lpszName As Long, _
                                ByVal lParam As Long) As Long
```

We might think that we can copy-paste that Function declaration and place it into an object, and we certainly can. However, when we do that, it's no longer "what you see is what you get". Putting it into a COM object adds an argument onto the left, and also rearranges the return. When placed into a COM object, here's what it looks like "under the hood":

```vb
Public Function EnumResNameProc(ByVal ObjPtrValue As Long, _
                                ByVal hModule As Long, _
                                ByVal lpszType As Long, _
                                ByVal lpszName As Long, _
                                ByVal lParam As Long, _
                                ByRef OurReturn As Long) As Long ' Returns API HRESULT value.
```

So, we can begin to see why we'd crash if we called this as if it were a regular BAS procedure. Furthermore, that ObjPtrValue argument **MUST** be a valid instantiated pointer, or we'll crash. Basically, when calling an object, it always needs to know where to find its instantiated module-level variables and any Static variables. And that's true even if we don't use them. That's what the `ObjPtrValue` is used for.

So, there's just no way around this without writing a bit of assembly language code (a "thunk"), and that's what I've done. I wrote it in such a way that it's a "wrapper" for any specific COM procedure, requiring that COM procedure's address and how many arguments are being passed. It always assumes it's a Function. The actual assembly code gets patched up for three things: 1) our ObjPtr value, 2) our procedure's address (from `Vb6ComCodeObjectAddressOf`), and 3) the number of arguments in the COM procedure. It also rearranges the stack to accept what would normally be on a BAS call's stack and makes it look like what should be on a COM call's stack. Here it is: (LaVolpe's & Wqweto's code were my inspiration for getting this sorted.)

```vb
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

Public Function AddressOfThunkForComProc(ByVal o As Object, _
                                         ByVal ComProcAddress As Long, _
                                         ByVal ComProcArgCount As Long, _
                                         ByRef iThunkSize As Long) As Long
    ' This makes a thunk and places it into a Byte array.
    ' This thunk is specifically for allowing COM procedures to be called as standard BAS procedures.
    '
    ' ComProcAddress:   The actual address of the COM procedure, typically dug out of the vTable.
    '
    ' ComObjPtr:        Usually comes out of ObjPtr(TheObject) for an instantiated VB6 COM/Code object (Form, Class, UC, PropertyPage, or DataReport).
    '
    ' ComProcArgCount:  The number of arguments "seen" in the code of the COM procedure.
    '                   It's MANDATORY that this be a number from 1 to 4 !!!!
    '                   It is assumed all the arguments are either ByRef or 4-byte ByVal arguments.
    '                   Anything else will fail and probably cause a crash.
    '
    ' iThunkSize:       Returned, and it must be used with VirtualFree to free this thunk memory.
    '                   The FreeTheThunk procedure is setup for this.
    '
    If ComProcAddress = 0& Or ObjPtr(o) = 0& Then Exit Function
    If ComProcArgCount < 1& Or ComProcArgCount > 4& Then Exit Function
    Dim bb() As Byte
    '
    ' Auto generated by Elroy's thunk maker.
    ReDim bb(44)
    ' ;
    ' ; The idea here is to convert a call that thinks it's a regular BAS call
    ' ; into a call that can go into a COM object procedure.  It is assumed that
    ' ; all arguments are either ByRef or 4-byte ByVal parameters.  It is
    ' ; further assumed that it's a Function that returns a Long.  This should
    ' ; cover the vast majority of API callbacks as well as subclassing.
    ' ;
    ' ; This thunk can handle one, two, three, or four incoming arguments.
    ' ; As an example, we'll assume two incoming arguments.  In a BAS module,
    ' ; such a function would look like the following:
    ' ;
    ' ;   Function OurCallBack(ByRef Arg1 As Long, ByRef Arg2 As Long) As Long
    ' ;
    ' ; When in a COM object, under the hood, this would be transformed as follows:
    ' ;
    ' ;   Function OurCallBack(ByVal OurObjPtr As Long, ByRef Arg1 As Long, _
    ' ;                        ByRef Arg2 As Long, ByRef FnRet As Long) As HRESULT
    ' ;
    ' ; So, to treat it as a BAS module call, we've got to add the OurObjPtr and
    ' ; deal with the return as an argument.  We just discard the HRESULT return.
    ' ;
    bb(0) = &H55                                                                        ' push    ebp                 ; Save base pointer, always done.
    bb(1) = &H89: bb(2) = &HE5                                                          ' mov     ebp, esp            ; Save stack pointer in ebp.
    bb(3) = &H83: bb(4) = &HEC: bb(5) = &H4                                             ' sub     esp, 4              ; Allocate 4 bytes of storage for local variables.
    ' ;
    ' ; We now start setting up for the COM procedure call.
    ' ;
    bb(6) = &H89: bb(7) = &HE8                                                          ' mov     eax, ebp            ; Base pointer into eax.
    bb(8) = &H83: bb(9) = &HE8: bb(10) = &H4                                            ' sub     eax, 4              ; Address for COM proc's FnRet to return.
    bb(11) = &H50                                                                       ' push    eax                 ;   and pushed on the stack for ByRef return.
    bb(12) = &HFF: bb(13) = &H75: bb(14) = &H14                                         ' push    [ebp + 20]          ; Arg4 onto stack.  Possibly patch with NOP.
    bb(15) = &HFF: bb(16) = &H75: bb(17) = &H10                                         ' push    [ebp + 16]          ; Arg3 onto stack.  Possibly patch with NOP.
    bb(18) = &HFF: bb(19) = &H75: bb(20) = &HC                                          ' push    [ebp + 12]          ; Arg2 onto stack.  Possibly patch with NOP.
    bb(21) = &HFF: bb(22) = &H75: bb(23) = &H8                                          ' push    [ebp + 8]           ; Arg1 onto stack.  At least one arg is required.
    bb(24) = &H68: bb(25) = &H55: bb(26) = &H55: bb(27) = &H55: bb(28) = &H55           ' push    0x55555555          ; We'll patch this up with the OurObjPtr address.
    ' ;
    bb(29) = &HB8: bb(30) = &H66: bb(31) = &H66: bb(32) = &H66: bb(33) = &H66           ' mov     eax, 0x66666666     ; We'll patch this up with the address to the COM procedure.
    bb(34) = &HFF: bb(35) = &HD0                                                        ' call    eax                 ; Call the COM procedure.
    ' ;
    ' ; We're back, so patch up the return from COM proc, and return.
    ' ; The stack will take care of itself, as ByVal are discarded,
    ' ; and ByRef were passed through with the same address.
    ' ;
    bb(36) = &H8B: bb(37) = &H45: bb(38) = &HFC                                         ' mov     eax, [ebp - 4]      ; Return the last argument's value as our return.
    bb(39) = &H89: bb(40) = &HEC                                                        ' mov     esp, ebp            ; Restore stack pointer from base pointer.
    bb(41) = &H5D                                                                       ' pop     ebp                 ; Restore base pointer.
    bb(42) = &HC2: bb(43) = &H8: bb(44) = &H0                                           ' ret     8                   ; Reset stack (for passed args) and return.  Patch up 8 for exact number of args (x 4).
    '
    ' We will need the size.
    iThunkSize = UBound(bb) - LBound(bb) + 1&
    '
    ' If not four arguments, blank out unused.
    If ComProcArgCount < 4& Then bb(12) = &H90: bb(13) = &H90: bb(14) = &H90    ' &H90 = NOP (no operation)
    If ComProcArgCount < 3& Then bb(15) = &H90: bb(16) = &H90: bb(17) = &H90
    If ComProcArgCount < 2& Then bb(18) = &H90: bb(19) = &H90: bb(20) = &H90
    ' At least one argument is required, so bb(21) thru bb(23) don't change.
    '
    ' Patch up our two supplied addresses.
    CopyMemory bb(25), ObjPtr(o), 4&            ' The ObjPtr() that's needed as an argument.
    CopyMemory bb(30), ComProcAddress, 4&       ' Where the actual call to the COM procedure is being made.
    '
    ' Patch up the return for how much of the stack to reset.
    bb(43) = ComProcArgCount * 4&
    '
    ' Get some executable memory.  Make sure we release it when we're done.
    Const MEM_COMMIT                    As Long = &H1000&
    Const PAGE_EXECUTE_READWRITE        As Long = &H40&
    AddressOfThunkForComProc = VirtualAlloc(0&, iThunkSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    '
    ' Move thunk into executable memory, and return our virtual memory's address.
    CopyMemory ByVal AddressOfThunkForComProc, bb(0), iThunkSize
End Function
```

The actual assembly code is shown far out to the right of the machine code being put into the Byte array. This is actually not terribly difficult to follow. When actually returning, the actual HRESULT return is discarded (as it is when we call any procedure in a COM object), and replaced with that last Long argument that was added.

Again, so basically, what we have is a piece of machine code that turns a BAS call into a COM call ... and then, upon return, it discards HRESULT and makes it look like a typical BAS function return.

One last piece ... when we made our thunk, we allocated some memory which allows executable code. When we're done, we should deallocate this memory:

```vb
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Public Sub FreeTheThunk(ByVal pVirtualMem As Long, ByVal iThunkSize As Long)
    Const MEM_RELEASE As Long = &H8000&
    VirtualFree pVirtualMem, iThunkSize, MEM_RELEASE
End Sub
```

If we don't do this last piece, we'll have a bit of a memory leak. However, those do clear up when an application exits.

And voilà, we've got all the pieces for using a COM object's procedure as an API CallBack.

## Example 1

Just to illustrate that it can work, I'll setup to run an API timer, using only that code in the `CallingVb6Objects.bas` file (attached in post above), and a CallBack procedure in a `Form1`:

```vb
Option Explicit
'
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'
Dim TimerID     As Long     ' Save this so we can kill the timer.
'
Dim ThunkSize   As Long     ' \
Dim pVirtualMem As Long     '  Save these two so we can release the thunk memory.
'


Private Sub Form_Load()
    Me.AutoRedraw = True
    Me.Font.Name = "Segoe UI Semibold"
    Me.Font.Size = 12
End Sub


Private Sub Form_Initialize()


    ' Get our procedure address and argument count.
    Dim iArgCount       As Long
    Dim pProcAddress    As Long
    pProcAddress = Vb6ComCodeObjectAddressOf(Me, "TimerCallBackProc", iArgCount)

    ' Make THUNK into executable memory.
    pVirtualMem = AddressOfThunkForComProc(Me, pProcAddress, iArgCount, ThunkSize)

    ' Create an API timer with our thunk.
    ' Callsback every 5000 milliseconds (5 seconds).
    TimerID = SetTimer(0&, 0&, 5000&, pVirtualMem)
End Sub


Public Function TimerCallBackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
    Me.Print "Time is: "; Format$(Time, "hh:nn:ss AM/PM")
End Function


Private Sub Form_Terminate()
    KillTimer 0&, TimerID                   ' Kill our API timer.
    FreeTheThunk pVirtualMem, ThunkSize     ' Give our thunk memory back.
End Sub
```

Let's look at the individual procedures in this example:

* `Form_Load`: All I'm doing here is setting up the form a bit.
* `Form_Initialize`: Here's where I'm setting up to use the TimerCallBackProc procedure as a CallBack. Notice that it's the pVirtualMem thunk address that's passed to the SetTimer API, so that the BAS-like CallBack from Windows can be rearranged for the COM procedure.
* `TimerCallBackProc`: This is the CallBack that Windows calls every 5 seconds.
* `Form_Terminate`: This is just cleanup: 1) kill the timer, 2) release our thunk memory.

## Example 2

Part of the fascination with all of this is the ability to keep everything wrapped into a single object. Therefore, in this next example, I've just copied all the code from the `CallingVb6Objects.bas` module, and pasted it into the Form1's code. I also made all the support procedures Private.

So, here's code for a `Form1`, with no other code needed, that performs a Timer API CallBack:

```vb
Option Explicit
'
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function vbaCheckType Lib "msvbvm60" Alias "__vbaCheckType" (ByVal pObj As Any, ByRef pIID As Any) As Boolean
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByRef lpString As Any) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByRef psz As Any, ByVal lSize As Long) As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'
Dim TimerID     As Long     ' Save this so we can kill the timer.
'
Dim ThunkSize   As Long     ' \
Dim pVirtualMem As Long     '  Save these two so we can release the thunk memory.
'


Private Sub Form_Load()
    Me.AutoRedraw = True
    Me.Font.Name = "Segoe UI Semibold"
    Me.Font.Size = 12
End Sub


Private Sub Form_Initialize()


    ' Get our procedure address and argument count.
    Dim iArgCount       As Long
    Dim pProcAddress    As Long
    pProcAddress = Vb6ComCodeObjectAddressOf(Me, "TimerCallBackProc", iArgCount)

    ' Make THUNK into executable memory.
    pVirtualMem = AddressOfThunkForComProc(Me, pProcAddress, iArgCount, ThunkSize)

    ' Create an API timer with our thunk.
    ' Callsback every 5000 milliseconds (5 seconds).
    TimerID = SetTimer(0&, 0&, 5000&, pVirtualMem)
End Sub


Public Function TimerCallBackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
    Me.Print "Time is: "; Format$(Time, "hh:nn:ss AM/PM")
End Function


Private Sub Form_Terminate()
    KillTimer 0&, TimerID                   ' Kill our API timer.
    FreeTheThunk pVirtualMem, ThunkSize     ' Give our thunk memory back.
End Sub



' *************************************************
' API CallBack to object ... support procedures.
' *************************************************

Private Function Vb6ComCodeObjectAddressOf(ByVal o As Object, ByVal sMethodName As String, Optional ByRef lArgCount As Long) As Long
    ' lArgCount is an optional RETURN.  ByRef/ByVal doesn't matter, and TYPE doesn't matter.  If it's a Function the return isn't counted.
    ' The caller is responsible for knowing how to use the returned address, or crash may result.
    ' See more notes within VtableOffsetForVb6ComMethod procedure.
    '
    ' Returns ZERO if not found, or it's not a PUBLIC method.
    '
    lArgCount = 0&                                                  ' Zero out anything there to start.
    Dim iVoffset As Long:   iVoffset = VtableOffsetForVb6ComMethod(o, sMethodName, lArgCount)
    If iVoffset = 0& Then Exit Function                             ' This checks to make sure the object is okay.
    Dim pVtable  As Long:   GetMem4 ByVal ObjPtr(o), pVtable        ' Get pointer to start of vTable.
    pVtable = pVtable + iVoffset                                    ' Address to our method in the vTable.
    GetMem4 ByVal pVtable, Vb6ComCodeObjectAddressOf                ' Pointer into actual code (the method).
End Function

Private Function VtableOffsetForVb6ComMethod(ByVal o As Object, ByVal sMethodName As String, Optional ByRef lArgCount As Long) As Long
    ' Searches PUBLIC methods.  It "could" find Private & Friend, but only in the IDE, not compiled.
    ' Does NOT search properties (i.e., Public variables or Get/Let/Set procedures).
    ' Returns an OFFSET address ready to be added to the vTable address.
    ' Optional bbArgCount return is a count of passed in arguments.
    '       ByRef/ByVal doesn't matter, and TYPE doesn't matter.
    '       If it's a Function the return isn't counted.
    '
    ' If it can't be found, ZERO is returned.
    '
    If Not ObjectIsVb6ComCodeModule(o) Then Exit Function                           ' Make sure we're dealing with a VB6 COM-code object.
    sMethodName = UCase$(sMethodName)
    '
    Dim pVTbl       As Long:    GetMem4 ByVal ObjPtr(o), pVTbl                      ' Pointer to vTable.
    Dim pObjInfo    As Long:    GetMem4 ByVal pVTbl - 4&, pObjInfo                  ' Pointer to tObjectInfo structure.
    Dim pPubDesc    As Long:    GetMem4 ByVal pObjInfo + &H18&, pPubDesc            ' tObjectInfo.aObject which points to tObject structure.
    Dim pPrivDesc   As Long:    GetMem4 ByVal pObjInfo + &HC&, pPrivDesc            ' tObjectInfo.lpPrivateObject which points to tPrivateObj structure.
    '
    If pPrivDesc = 0& Then Exit Function                                            ' Just a double-check.
    '
    Dim lIndex      As Long
    Dim pName       As Long
    '
    ' Search the procedures within the module.
    Dim pMethDesc   As Long
    Dim iMethOffset As Integer
    Dim bbArgs      As Byte
    Dim lMethodsCnt As Long:    GetMem2 ByVal pPubDesc + &H1C&, lMethodsCnt         ' tObject.ProcCount value.
    Dim pNames      As Long:    GetMem4 ByVal pPubDesc + &H20&, pNames              ' tObject.aProcNamesArray which points to an array of name pointers.
    Dim pMethodsPtr As Long:    GetMem4 ByVal pPrivDesc + &H18&, pMethodsPtr        ' tPrivateObj.lpFuncTypeInfo which points to an array of pointers.
    '
    ' Loop through methods and see if we can find the one we want.
    For lIndex = 0& To lMethodsCnt - 1&
        GetMem4 ByVal pMethodsPtr + lIndex * 4&, pMethDesc                          ' From the array, getting a pointer to a method structure (tMethInfo).
        If pMethDesc Then                                                           ' Not sure if this ever returns zero, maybe for "Private" methods?
            GetMem2 ByVal pMethDesc + 2&, iMethOffset                               ' Out of tMethInfo structure.
            GetMem1 ByVal pMethDesc, bbArgs                                         ' First two bits of bbArgs are: set=3, get=1, let=2, method=0 (Sub or Fn).
            If (bbArgs And CByte(3)) = CByte(0) Then                                ' Make sure it's a method.
                If iMethOffset And 1 Then                                           ' First bit, 1=Public.
                    GetMem4 ByVal pNames + lIndex * 4&, pName                       ' Dig pointer to method name from array of name pointers.
                    If sMethodName = UCase$(SysAllocStringByteLen(ByVal pName, lstrlenA(ByVal pName))) Then
                        VtableOffsetForVb6ComMethod = CLng(iMethOffset And &HFFFC)  ' First two bits are something else (first is Public=1,Private=0).
                        Dim bbFlags As Byte: GetMem1 ByVal pMethDesc + 1&, bbFlags  ' Both bbArgs & bbFlags out of tMethInfo structure.
                        bbFlags = bbFlags And CByte(1)                              ' 0 (no return), 1 (return).
                        lArgCount = CLng(bbArgs \ CByte(4) - bbFlags)               ' Calculate arguments, excluding any return argument.  Tested for vbGet, vbLet, vbSet, vbMethod (both Function & Sub).
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    '
    ' Return zero if not found.
End Function

Private Sub FreeTheThunk(ByVal pVirtualMem As Long, ByVal iThunkSize As Long)
    Const MEM_RELEASE As Long = &H8000&
    VirtualFree pVirtualMem, iThunkSize, MEM_RELEASE
End Sub

Private Function AddressOfThunkForComProc(ByVal o As Object, _
                                         ByVal ComProcAddress As Long, _
                                         ByVal ComProcArgCount As Long, _
                                         ByRef iThunkSize As Long) As Long
    ' This makes a thunk and places it into a Byte array.
    ' This thunk is specifically for allowing COM procedures to be called as standard BAS procedures.
    '
    ' ComProcAddress:   The actual address of the COM procedure, typically dug out of the vTable.
    '
    ' ComObjPtr:        Usually comes out of ObjPtr(TheObject) for an instantiated VB6 COM/Code object (Form, Class, UC, PropertyPage, or DataReport).
    '
    ' ComProcArgCount:  The number of arguments "seen" in the code of the COM procedure.
    '                   It's MANDATORY that this be a number from 1 to 4 !!!!
    '                   It is assumed all the arguments are either ByRef or 4-byte ByVal arguments.
    '                   Anything else will fail and probably cause a crash.
    '
    ' iThunkSize:       Returned, and it must be used with VirtualFree to free this thunk memory.
    '                   The FreeTheThunk procedure is setup for this.
    '
    If ComProcAddress = 0& Or ObjPtr(o) = 0& Then Exit Function
    If ComProcArgCount < 1& Or ComProcArgCount > 4& Then Exit Function
    Dim bb() As Byte
    '
    ' Auto generated by Elroy's thunk maker.
    ReDim bb(44)
    ' ;
    ' ; The idea here is to convert a call that thinks it's a regular BAS call
    ' ; into a call that can go into a COM object procedure.  It is assumed that
    ' ; all arguments are either ByRef or 4-byte ByVal parameters.  It is
    ' ; further assumed that it's a Function that returns a Long.  This should
    ' ; cover the vast majority of API callbacks as well as subclassing.
    ' ;
    ' ; This thunk can handle one, two, three, or four incoming arguments.
    ' ; As an example, we'll assume two incoming arguments.  In a BAS module,
    ' ; such a function would look like the following:
    ' ;
    ' ;   Function OurCallBack(ByRef Arg1 As Long, ByRef Arg2 As Long) As Long
    ' ;
    ' ; When in a COM object, under the hood, this would be transformed as follows:
    ' ;
    ' ;   Function OurCallBack(ByVal OurObjPtr As Long, ByRef Arg1 As Long, _
    ' ;                        ByRef Arg2 As Long, ByRef FnRet As Long) As HRESULT
    ' ;
    ' ; So, to treat it as a BAS module call, we've got to add the OurObjPtr and
    ' ; deal with the return as an argument.  We just discard the HRESULT return.
    ' ;
    bb(0) = &H55                                                                        ' push    ebp                 ; Save base pointer, always done.
    bb(1) = &H89: bb(2) = &HE5                                                          ' mov     ebp, esp            ; Save stack pointer in ebp.
    bb(3) = &H83: bb(4) = &HEC: bb(5) = &H4                                             ' sub     esp, 4              ; Allocate 4 bytes of storage for local variables.
    ' ;
    ' ; We now start setting up for the COM procedure call.
    ' ;
    bb(6) = &H89: bb(7) = &HE8                                                          ' mov     eax, ebp            ; Base pointer into eax.
    bb(8) = &H83: bb(9) = &HE8: bb(10) = &H4                                            ' sub     eax, 4              ; Address for COM proc's FnRet to return.
    bb(11) = &H50                                                                       ' push    eax                 ;   and pushed on the stack for ByRef return.
    bb(12) = &HFF: bb(13) = &H75: bb(14) = &H14                                         ' push    [ebp + 20]          ; Arg4 onto stack.  Possibly patch with NOP.
    bb(15) = &HFF: bb(16) = &H75: bb(17) = &H10                                         ' push    [ebp + 16]          ; Arg3 onto stack.  Possibly patch with NOP.
    bb(18) = &HFF: bb(19) = &H75: bb(20) = &HC                                          ' push    [ebp + 12]          ; Arg2 onto stack.  Possibly patch with NOP.
    bb(21) = &HFF: bb(22) = &H75: bb(23) = &H8                                          ' push    [ebp + 8]           ; Arg1 onto stack.  At least one arg is required.
    bb(24) = &H68: bb(25) = &H55: bb(26) = &H55: bb(27) = &H55: bb(28) = &H55           ' push    0x55555555          ; We'll patch this up with the OurObjPtr address.
    ' ;
    bb(29) = &HB8: bb(30) = &H66: bb(31) = &H66: bb(32) = &H66: bb(33) = &H66           ' mov     eax, 0x66666666     ; We'll patch this up with the address to the COM procedure.
    bb(34) = &HFF: bb(35) = &HD0                                                        ' call    eax                 ; Call the COM procedure.
    ' ;
    ' ; We're back, so patch up the return from COM proc, and return.
    ' ; The stack will take care of itself, as ByVal are discarded,
    ' ; and ByRef were passed through with the same address.
    ' ;
    bb(36) = &H8B: bb(37) = &H45: bb(38) = &HFC                                         ' mov     eax, [ebp - 4]      ; Return the last argument's value as our return.
    bb(39) = &H89: bb(40) = &HEC                                                        ' mov     esp, ebp            ; Restore stack pointer from base pointer.
    bb(41) = &H5D                                                                       ' pop     ebp                 ; Restore base pointer.
    bb(42) = &HC2: bb(43) = &H8: bb(44) = &H0                                           ' ret     8                   ; Reset stack (for passed args) and return.  Patch up 8 for exact number of args (x 4).
    '
    ' We will need the size.
    iThunkSize = UBound(bb) - LBound(bb) + 1&
    '
    ' If not four arguments, blank out unused.
    If ComProcArgCount < 4& Then bb(12) = &H90: bb(13) = &H90: bb(14) = &H90    ' &H90 = NOP (no operation)
    If ComProcArgCount < 3& Then bb(15) = &H90: bb(16) = &H90: bb(17) = &H90
    If ComProcArgCount < 2& Then bb(18) = &H90: bb(19) = &H90: bb(20) = &H90
    ' At least one argument is required, so bb(21) thru bb(23) don't change.
    '
    ' Patch up our two supplied addresses.
    CopyMemory bb(25), ObjPtr(o), 4&            ' The ObjPtr() that's needed as an argument.
    CopyMemory bb(30), ComProcAddress, 4&       ' Where the actual call to the COM procedure is being made.
    '
    ' Patch up the return for how much of the stack to reset.
    bb(43) = ComProcArgCount * 4&
    '
    ' Get some executable memory.  Make sure we release it when we're done.
    Const MEM_COMMIT                    As Long = &H1000&
    Const PAGE_EXECUTE_READWRITE        As Long = &H40&
    AddressOfThunkForComProc = VirtualAlloc(0&, iThunkSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    '
    ' Move thunk into executable memory, and return our virtual memory's address.
    CopyMemory ByVal AddressOfThunkForComProc, bb(0), iThunkSize
End Function

Private Function ObjectIsVb6ComCodeModule(ByRef o As IUnknown) As Boolean
    ' If it's an instantiated Class, Form, UC, PropPage, DataReport, returns TRUE, else FALSE.
    If ObjPtr(o) = 0& Then Exit Function                    ' Make sure "something" is instantiated.
    Dim aGUID(1&) As Currency                               ' Just to get 16 easily accessible bytes.
    aGUID(0&) = 128347367577987.1845@                       ' Const AreYouABasicInstance As String = "{0B6C9465-D082-11CF-8B4F-00A0C90F2704}"
    aGUID(1&) = 29922525889064.5387@                        ' turned into two numbers stuffed into our Currency array.
    ObjectIsVb6ComCodeModule = vbaCheckType(o, aGUID(0))    ' Check and see if we are this "TYPE" (Class, Form, UC, PropPage, or DataRep).
End Function
```

## Example 3

And, for grins, I'll do it to list our monitor handles and then to find the primary monitor's handle. I'll use two separate CallBacks in this one, but I'll still keep it all in a single module (Form1). And here it is:

```vb
Option Explicit
'
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function vbaCheckType Lib "msvbvm60" Alias "__vbaCheckType" (ByVal pObj As Any, ByRef pIID As Any) As Boolean
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByRef lpString As Any) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByRef psz As Any, ByVal lSize As Long) As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As Any) As Long
'



Public Function MonitorHandleEnumCallback(ByVal hMonitor As Long, ByVal hdcMonitor As Long, uRect As Variant, dwData As Long) As Long
    ' uRect is a Variant because it's not used and it's 16 bytes.
    dwData = dwData + 1&            ' The actual count.
    Me.Print "   "; CStr(hMonitor)
    MonitorHandleEnumCallback = 1&  ' Count them all.
End Function

Public Function PrimaryMonitorHandleEnumCallback(ByVal hMonitor As Long, ByVal hdcMonitor As Long, uRect As Variant, dwData As Long) As Long
    ' uRect is a Variant because it's not used and it's 16 bytes.
    Dim MonInfo(9&) As Long ' cbSize As Long, rcMonitor As RECT, rcWork As RECT, dwFlags As Long
    MonInfo(0&) = 40&
    GetMonitorInfo hMonitor, MonInfo(0&)
    If MonInfo(9&) = &H1& Then
        dwData = hMonitor
        PrimaryMonitorHandleEnumCallback = 0& ' Found it.
    Else
        PrimaryMonitorHandleEnumCallback = 1& ' Keep looking.
    End If
End Function


Private Sub Form_Load()

    ' Form setup.
    Me.AutoRedraw = True
    Me.Font.Name = "Segoe UI Semibold"
    Me.Font.Size = 12


    ' Some variables we'll use.
    Dim ThunkSize   As Long
    Dim pVirtualMem As Long
    Dim iArgCount       As Long
    Dim pProcAddress    As Long

    ' First, just list our monitors.

    ' Get our procedure address and argument count.
    pProcAddress = Vb6ComCodeObjectAddressOf(Me, "MonitorHandleEnumCallback", iArgCount)

    ' Make THUNK into executable memory.
    pVirtualMem = AddressOfThunkForComProc(Me, pProcAddress, iArgCount, ThunkSize)

    ' Count our monitors ... their handles are listed in the callback.
    Me.Print "Monitor handle(s):"
    Dim iMonitorCount As Long
    EnumDisplayMonitors 0&, ByVal 0&, pVirtualMem, iMonitorCount
    Me.Print "Total monitor count: "; CStr(iMonitorCount)

    ' Give our thunk memory back.
    FreeTheThunk pVirtualMem, ThunkSize



    ' Now let's use a callback to find our primary monitor.

    ' Get our procedure address and argument count.
    pProcAddress = Vb6ComCodeObjectAddressOf(Me, "PrimaryMonitorHandleEnumCallback", iArgCount)

    ' Make THUNK into executable memory.
    pVirtualMem = AddressOfThunkForComProc(Me, pProcAddress, iArgCount, ThunkSize)

    ' Get primary monitor handle using callback.
    Dim hPrimaryMonitor As Long
    EnumDisplayMonitors 0&, ByVal 0&, pVirtualMem, hPrimaryMonitor
    Me.Print "Handle of primary monitor: "; CStr(hPrimaryMonitor)

    ' Give our thunk memory back.
    FreeTheThunk pVirtualMem, ThunkSize


End Sub




' *************************************************
' API CallBack to object ... support procedures.
' *************************************************

Private Function Vb6ComCodeObjectAddressOf(ByVal o As Object, ByVal sMethodName As String, Optional ByRef lArgCount As Long) As Long
    ' lArgCount is an optional RETURN.  ByRef/ByVal doesn't matter, and TYPE doesn't matter.  If it's a Function the return isn't counted.
    ' The caller is responsible for knowing how to use the returned address, or crash may result.
    ' See more notes within VtableOffsetForVb6ComMethod procedure.
    '
    ' Returns ZERO if not found, or it's not a PUBLIC method.
    '
    lArgCount = 0&                                                  ' Zero out anything there to start.
    Dim iVoffset As Long:   iVoffset = VtableOffsetForVb6ComMethod(o, sMethodName, lArgCount)
    If iVoffset = 0& Then Exit Function                             ' This checks to make sure the object is okay.
    Dim pVtable  As Long:   GetMem4 ByVal ObjPtr(o), pVtable        ' Get pointer to start of vTable.
    pVtable = pVtable + iVoffset                                    ' Address to our method in the vTable.
    GetMem4 ByVal pVtable, Vb6ComCodeObjectAddressOf                ' Pointer into actual code (the method).
End Function

Private Function VtableOffsetForVb6ComMethod(ByVal o As Object, ByVal sMethodName As String, Optional ByRef lArgCount As Long) As Long
    ' Searches PUBLIC methods.  It "could" find Private & Friend, but only in the IDE, not compiled.
    ' Does NOT search properties (i.e., Public variables or Get/Let/Set procedures).
    ' Returns an OFFSET address ready to be added to the vTable address.
    ' Optional bbArgCount return is a count of passed in arguments.
    '       ByRef/ByVal doesn't matter, and TYPE doesn't matter.
    '       If it's a Function the return isn't counted.
    '
    ' If it can't be found, ZERO is returned.
    '
    If Not ObjectIsVb6ComCodeModule(o) Then Exit Function                           ' Make sure we're dealing with a VB6 COM-code object.
    sMethodName = UCase$(sMethodName)
    '
    Dim pVTbl       As Long:    GetMem4 ByVal ObjPtr(o), pVTbl                      ' Pointer to vTable.
    Dim pObjInfo    As Long:    GetMem4 ByVal pVTbl - 4&, pObjInfo                  ' Pointer to tObjectInfo structure.
    Dim pPubDesc    As Long:    GetMem4 ByVal pObjInfo + &H18&, pPubDesc            ' tObjectInfo.aObject which points to tObject structure.
    Dim pPrivDesc   As Long:    GetMem4 ByVal pObjInfo + &HC&, pPrivDesc            ' tObjectInfo.lpPrivateObject which points to tPrivateObj structure.
    '
    If pPrivDesc = 0& Then Exit Function                                            ' Just a double-check.
    '
    Dim lIndex      As Long
    Dim pName       As Long
    '
    ' Search the procedures within the module.
    Dim pMethDesc   As Long
    Dim iMethOffset As Integer
    Dim bbArgs      As Byte
    Dim lMethodsCnt As Long:    GetMem2 ByVal pPubDesc + &H1C&, lMethodsCnt         ' tObject.ProcCount value.
    Dim pNames      As Long:    GetMem4 ByVal pPubDesc + &H20&, pNames              ' tObject.aProcNamesArray which points to an array of name pointers.
    Dim pMethodsPtr As Long:    GetMem4 ByVal pPrivDesc + &H18&, pMethodsPtr        ' tPrivateObj.lpFuncTypeInfo which points to an array of pointers.
    '
    ' Loop through methods and see if we can find the one we want.
    For lIndex = 0& To lMethodsCnt - 1&
        GetMem4 ByVal pMethodsPtr + lIndex * 4&, pMethDesc                          ' From the array, getting a pointer to a method structure (tMethInfo).
        If pMethDesc Then                                                           ' Not sure if this ever returns zero, maybe for "Private" methods?
            GetMem2 ByVal pMethDesc + 2&, iMethOffset                               ' Out of tMethInfo structure.
            GetMem1 ByVal pMethDesc, bbArgs                                         ' First two bits of bbArgs are: set=3, get=1, let=2, method=0 (Sub or Fn).
            If (bbArgs And CByte(3)) = CByte(0) Then                                ' Make sure it's a method.
                If iMethOffset And 1 Then                                           ' First bit, 1=Public.
                    GetMem4 ByVal pNames + lIndex * 4&, pName                       ' Dig pointer to method name from array of name pointers.
                    If sMethodName = UCase$(SysAllocStringByteLen(ByVal pName, lstrlenA(ByVal pName))) Then
                        VtableOffsetForVb6ComMethod = CLng(iMethOffset And &HFFFC)  ' First two bits are something else (first is Public=1,Private=0).
                        Dim bbFlags As Byte: GetMem1 ByVal pMethDesc + 1&, bbFlags  ' Both bbArgs & bbFlags out of tMethInfo structure.
                        bbFlags = bbFlags And CByte(1)                              ' 0 (no return), 1 (return).
                        lArgCount = CLng(bbArgs \ CByte(4) - bbFlags)               ' Calculate arguments, excluding any return argument.  Tested for vbGet, vbLet, vbSet, vbMethod (both Function & Sub).
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    '
    ' Return zero if not found.
End Function

Private Sub FreeTheThunk(ByVal pVirtualMem As Long, ByVal iThunkSize As Long)
    Const MEM_RELEASE As Long = &H8000&
    VirtualFree pVirtualMem, iThunkSize, MEM_RELEASE
End Sub

Private Function AddressOfThunkForComProc(ByVal o As Object, _
                                         ByVal ComProcAddress As Long, _
                                         ByVal ComProcArgCount As Long, _
                                         ByRef iThunkSize As Long) As Long
    ' This makes a thunk and places it into a Byte array.
    ' This thunk is specifically for allowing COM procedures to be called as standard BAS procedures.
    '
    ' ComProcAddress:   The actual address of the COM procedure, typically dug out of the vTable.
    '
    ' ComObjPtr:        Usually comes out of ObjPtr(TheObject) for an instantiated VB6 COM/Code object (Form, Class, UC, PropertyPage, or DataReport).
    '
    ' ComProcArgCount:  The number of arguments "seen" in the code of the COM procedure.
    '                   It's MANDATORY that this be a number from 1 to 4 !!!!
    '                   It is assumed all the arguments are either ByRef or 4-byte ByVal arguments.
    '                   Anything else will fail and probably cause a crash.
    '
    ' iThunkSize:       Returned, and it must be used with VirtualFree to free this thunk memory.
    '                   The FreeTheThunk procedure is setup for this.
    '
    If ComProcAddress = 0& Or ObjPtr(o) = 0& Then Exit Function
    If ComProcArgCount < 1& Or ComProcArgCount > 4& Then Exit Function
    Dim bb() As Byte
    '
    ' Auto generated by Elroy's thunk maker.
    ReDim bb(44)
    ' ;
    ' ; The idea here is to convert a call that thinks it's a regular BAS call
    ' ; into a call that can go into a COM object procedure.  It is assumed that
    ' ; all arguments are either ByRef or 4-byte ByVal parameters.  It is
    ' ; further assumed that it's a Function that returns a Long.  This should
    ' ; cover the vast majority of API callbacks as well as subclassing.
    ' ;
    ' ; This thunk can handle one, two, three, or four incoming arguments.
    ' ; As an example, we'll assume two incoming arguments.  In a BAS module,
    ' ; such a function would look like the following:
    ' ;
    ' ;   Function OurCallBack(ByRef Arg1 As Long, ByRef Arg2 As Long) As Long
    ' ;
    ' ; When in a COM object, under the hood, this would be transformed as follows:
    ' ;
    ' ;   Function OurCallBack(ByVal OurObjPtr As Long, ByRef Arg1 As Long, _
    ' ;                        ByRef Arg2 As Long, ByRef FnRet As Long) As HRESULT
    ' ;
    ' ; So, to treat it as a BAS module call, we've got to add the OurObjPtr and
    ' ; deal with the return as an argument.  We just discard the HRESULT return.
    ' ;
    bb(0) = &H55                                                                        ' push    ebp                 ; Save base pointer, always done.
    bb(1) = &H89: bb(2) = &HE5                                                          ' mov     ebp, esp            ; Save stack pointer in ebp.
    bb(3) = &H83: bb(4) = &HEC: bb(5) = &H4                                             ' sub     esp, 4              ; Allocate 4 bytes of storage for local variables.
    ' ;
    ' ; We now start setting up for the COM procedure call.
    ' ;
    bb(6) = &H89: bb(7) = &HE8                                                          ' mov     eax, ebp            ; Base pointer into eax.
    bb(8) = &H83: bb(9) = &HE8: bb(10) = &H4                                            ' sub     eax, 4              ; Address for COM proc's FnRet to return.
    bb(11) = &H50                                                                       ' push    eax                 ;   and pushed on the stack for ByRef return.
    bb(12) = &HFF: bb(13) = &H75: bb(14) = &H14                                         ' push    [ebp + 20]          ; Arg4 onto stack.  Possibly patch with NOP.
    bb(15) = &HFF: bb(16) = &H75: bb(17) = &H10                                         ' push    [ebp + 16]          ; Arg3 onto stack.  Possibly patch with NOP.
    bb(18) = &HFF: bb(19) = &H75: bb(20) = &HC                                          ' push    [ebp + 12]          ; Arg2 onto stack.  Possibly patch with NOP.
    bb(21) = &HFF: bb(22) = &H75: bb(23) = &H8                                          ' push    [ebp + 8]           ; Arg1 onto stack.  At least one arg is required.
    bb(24) = &H68: bb(25) = &H55: bb(26) = &H55: bb(27) = &H55: bb(28) = &H55           ' push    0x55555555          ; We'll patch this up with the OurObjPtr address.
    ' ;
    bb(29) = &HB8: bb(30) = &H66: bb(31) = &H66: bb(32) = &H66: bb(33) = &H66           ' mov     eax, 0x66666666     ; We'll patch this up with the address to the COM procedure.
    bb(34) = &HFF: bb(35) = &HD0                                                        ' call    eax                 ; Call the COM procedure.
    ' ;
    ' ; We're back, so patch up the return from COM proc, and return.
    ' ; The stack will take care of itself, as ByVal are discarded,
    ' ; and ByRef were passed through with the same address.
    ' ;
    bb(36) = &H8B: bb(37) = &H45: bb(38) = &HFC                                         ' mov     eax, [ebp - 4]      ; Return the last argument's value as our return.
    bb(39) = &H89: bb(40) = &HEC                                                        ' mov     esp, ebp            ; Restore stack pointer from base pointer.
    bb(41) = &H5D                                                                       ' pop     ebp                 ; Restore base pointer.
    bb(42) = &HC2: bb(43) = &H8: bb(44) = &H0                                           ' ret     8                   ; Reset stack (for passed args) and return.  Patch up 8 for exact number of args (x 4).
    '
    ' We will need the size.
    iThunkSize = UBound(bb) - LBound(bb) + 1&
    '
    ' If not four arguments, blank out unused.
    If ComProcArgCount < 4& Then bb(12) = &H90: bb(13) = &H90: bb(14) = &H90    ' &H90 = NOP (no operation)
    If ComProcArgCount < 3& Then bb(15) = &H90: bb(16) = &H90: bb(17) = &H90
    If ComProcArgCount < 2& Then bb(18) = &H90: bb(19) = &H90: bb(20) = &H90
    ' At least one argument is required, so bb(21) thru bb(23) don't change.
    '
    ' Patch up our two supplied addresses.
    CopyMemory bb(25), ObjPtr(o), 4&            ' The ObjPtr() that's needed as an argument.
    CopyMemory bb(30), ComProcAddress, 4&       ' Where the actual call to the COM procedure is being made.
    '
    ' Patch up the return for how much of the stack to reset.
    bb(43) = ComProcArgCount * 4&
    '
    ' Get some executable memory.  Make sure we release it when we're done.
    Const MEM_COMMIT                    As Long = &H1000&
    Const PAGE_EXECUTE_READWRITE        As Long = &H40&
    AddressOfThunkForComProc = VirtualAlloc(0&, iThunkSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    '
    ' Move thunk into executable memory, and return our virtual memory's address.
    CopyMemory ByVal AddressOfThunkForComProc, bb(0&), iThunkSize
End Function

Private Function ObjectIsVb6ComCodeModule(ByRef o As IUnknown) As Boolean
    ' If it's an instantiated Class, Form, UC, PropPage, DataReport, returns TRUE, else FALSE.
    If ObjPtr(o) = 0& Then Exit Function                    ' Make sure "something" is instantiated.
    Dim aGUID(1&) As Currency                               ' Just to get 16 easily accessible bytes.
    aGUID(0&) = 128347367577987.1845@                       ' Const AreYouABasicInstance As String = "{0B6C9465-D082-11CF-8B4F-00A0C90F2704}"
    aGUID(1&) = 29922525889064.5387@                        ' turned into two numbers stuffed into our Currency array.
    ObjectIsVb6ComCodeModule = vbaCheckType(o, aGUID(0&))   ' Check and see if we are this "TYPE" (Class, Form, UC, PropPage, or DataRep).
End Function
```

This time, both CallBacks were "one-shot" callbacks, so I created and deleted the thunks all in the same section of code. When it runs, here's the output:

```
|----------------------------------|
| Form1                            |
|----------------------------------|
| Monitor handle(s):               |
|  65539                           |
|  65537                           |
| Total monitor count: 2           |
| Handle of primary monitor: 65537 |
```

Ok, what if we'd like to keep our CallBack procedures `Private`? You can do this when they're in a BAS module, even with subclassing. And doing this provides some advantages:

* Being able to declare Private UDTs in the module and have them used in the CallBack arguments.
* Making sure no other code outside the module calls the CallBack module.

However, we stated above that the `VtableOffsetForVb6ComMethod` procedure only finds `Public` procedures within our object's code. So, we have to resort to other means, and to do this, we need to return to ideas surrounding the vTable.

## Challenge #4: How many entries, and what categories of entries, are in the vTable? 

We will need this information to proceed with making a `Private` CallBack procedure within a COM object.

(As a note, Paul Caton, LaVolpe, and others have used "magic numbers" to figure this out. However, I was never fond of these "magic numbers". The following approach goes back into the structures associated with each COM object and digs out the information.)

```vb
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long

Public Function Vb6ComCodeObjectVtableEntries(o As Object, Optional iInterfaceCount As Long, Optional iIntrinsicCount As Long, Optional iCoderCodedCount As Long) As Long
    ' Return is the TOTAL, with Optional arguments returned as individual pieces.
    '
    If Not ObjectIsVb6ComCodeModule(o) Then Exit Function
    '
    Dim pVtable As Long:        GetMem4 ByVal ObjPtr(o), pVtable
    Dim ptObjectInfo As Long:   GetMem4 ByVal pVtable - 4&, ptObjectInfo
    iInterfaceCount = 7&                                    ' IUnknown and IDispatch.
    GetMem2 ByVal ptObjectInfo + &H62&, iIntrinsicCount     ' Out of the tObjectInfo structure (wPCodeCount).
    GetMem2 ByVal ptObjectInfo + &H60&, iCoderCodedCount    ' Out of the tObjectInfo structure (wMethodLinkCount).
    Vb6ComCodeObjectVtableEntries = iInterfaceCount + iIntrinsicCount + iCoderCodedCount
End Function
```

Notice that we're still dependent upon it being either a `Form`, `Class`, `UserControl`, `PropPage`, or `DataReport`, by checking with `ObjectIsVb6ComCodeModule`.

This `ObjectIsVb6ComCodeModule` function returns the total of the vTable entries. However, more important to us are the numbers in each category, the optional `iInterfaceCount`, `iIntrinsicCount`, and `iCoderCodedCount`. Above, it was stated that we're not concerned about the interface entries nor the intrinsic entries ... and we're still not. However, we do need to know their counts, and we've now got them.

But we do need to focus on the `iCoderCodedCount`, and we need to learn how to actually count them. This count is all of our `Public` variables, `Public` procedures, and `Private`/`Friend` procedures that we've coded up in our COM code module, and they appear in the vTable **in that order**. `Private` variables (even at the module level) do not appear in the vTable.

## Challenge #5: Learning how to count vTable entries we've created from our code

(I mostly followed LaVolpe's lead on this one.)

`Public` variables are actually compiled as a Property Get and Property Let (or Property Set for `Public` object variables). For these, the Get always appears first. So, Public variables actually have two vTable entries.

After those, the `Public` procedures (including any event procedures you've declared as `Public`) appear. This includes `Public Sub`, `Public Function`, `Public Property Get`, `Public Property Let`, and `Public Property Set` procedures. And these always appear in the vTable in the same order they appear in the code.

Once all the `Public` variables and `Public` procedures are dealt with, there is an entry for each `Private`/`Friend` procedure in the code (including event procedures). Again, after the `Public` procedures, these `Private`/`Friend` procedures all appear in the vTable in the order in which they appear in the code.

Skipping over the `iInterfaceCount` and `iIntrinsicCount`, let's look at an example of counting vTable entries. We'll stick with the idea that it's all in Form1, although these concepts hold for all COM module code.

```vb
Option Explicit
   
Private var1 As Long                    ' Not counted.
Public var2 As Long                     ' vTable entry #1 (Get) & #2 (Let).
Private var3 As Long                    ' Not counted.

Private Sub Test_Load(): End Sub                       ' vTable entry #5 (an event, but doesn't matter)
Public Sub Test_Sub(arg1 As Byte): End Sub             ' vTable entry #3 <--- it's PUBLIC
Friend Property Let Test_Prop(i As Long): End Property ' vTable entry #6
Friend Property Get Test_Prop() As Long: End Property  ' vTable entry #7
Private Sub Test_Click(): End Sub                      ' vTable entry #8 (an event, but doesn't matter)
Public Function Test_Fn() As Variant: End Function     ' vTable entry #4 <--- it's PUBLIC
Friend Function Test_AnotherFn() As Long: End Function ' vTable entry #9

' Again, Private vs Friend doesn't matter.
```

If you carefully study the vTable counting rules and that `Form1` code, you will (hopefully) understand.

So, to use a `Private` CallBack procedure, we must be able to count which vTable procedure it is (starting from 1). Then, we can use that number to get our procedure address out of the vTable. Here's the code that does that for us (with our vTable number being `iCoderProcNumber`): (LaVolpe's code & tips from The Trick helped me to get this sorted.)

```vb
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long

Public Function Vb6ComCodeObjectAddressOfByNumber(ByVal o As Object, ByVal iCoderProcNumber As Long) As Long
    ' The caller is responsible for knowing how to use the returned address, or crash may result.
    ' COM/code objects have one (or two) hidden arguments, one for the ObjPtr address, and,[
    ' if it's a function (or property get), another for that return.  ObjPtr is first and return value is second.
    '
    ' The iCoderProcNumber starts at 1 and can be no larger than the iCoderCodedCount returned by Vb6ComCodeObjectVtableEntries.
    ' The procedure number is a CODER CODED procedure, including any events.
    '
    Dim iInterfaceCount As Long, iIntrinsicCount As Long, iCoderCodedCount As Long
    Call Vb6ComCodeObjectVtableEntries(o, iInterfaceCount, iIntrinsicCount, iCoderCodedCount)
    If iCoderProcNumber < 1& Or iCoderProcNumber > iCoderCodedCount Then Exit Function    ' This will catch bad objects.
    '
    ' Return the address in the vTable.
    Dim pVtable As Long:    GetMem4 ByVal ObjPtr(o), pVtable        ' Get pointer to start of vTable.
    pVtable = pVtable + (iInterfaceCount + iIntrinsicCount) * 4&    ' Jump over the interface and intrinsic entries. 4& is 4 bytes per pointer (32-bit).
    pVtable = pVtable + (iCoderProcNumber - 1&) * 4&                ' And now go up to our specified procedure in the vTable.
    GetMem4 ByVal pVtable, Vb6ComCodeObjectAddressOfByNumber        ' Pointer into actual code (the member).
End Function
```

Notice that this `Vb6ComCodeObjectAddressOfByNumber` procedure uses the prior listed `Vb6ComCodeObjectVtableEntries` procedure to get its work done. It needs this so it can jump over the interface and intrinsic vTable entries to get to the one of interest to us (our CallBack procedure).

Now, we've finally arrived back at a point where we can use this address along with our thunk to make CallBack calls into our COM object's procedures. And this time, we can keep them Private.

## Example #4: I'll just rework the above "monitor handle" example to use this approach

Again, I'm just going to put all the code into a `Form1` module.

Also, with this approach, there was nothing to tell us how many arguments the CallBack procedure had, which was needed in the call to `AddressOfThunkForComProc`. So, we just have to count them and hard-code this value, which is what you'll see I've done. I've also changed the CallBack procedures around to actually use the `RECT` type, since we can now do that, and I report the monitor dimensions in one of the callbacks.

```vb
Option Explicit
'
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Function vbaCheckType Lib "msvbvm60" Alias "__vbaCheckType" (ByVal pObj As Any, ByRef pIID As Any) As Boolean
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByRef lpString As Any) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByRef psz As Any, ByVal lSize As Long) As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long ' This is +1 (right - left = width)
    Bottom As Long ' This is +1 (bottom - top = height)
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As Any) As Long
'


' vTable entry #1 (of the coder procs), as it's the FIRST coder vTable item in this module ... and there are 4 arguments.
' Notice that there are no Public variables nor Public procedures in the module, or this vTable count would change.
Private Function MonitorHandleEnumCallback(ByVal hMonitor As Long, ByVal hdcMonitor As Long, uRect As RECT, dwData As Long) As Long
    dwData = dwData + 1&            ' The actual count.
    Me.Print "   "; CStr(hMonitor); " Left: "; Format$(uRect.Left, "0000"); _
                                   " Width: "; Format$(uRect.Right - uRect.Left, "0000"); _
                                     " Top: "; Format$(uRect.Top, "0000"); _
                                  " Height: "; Format$(uRect.Bottom - uRect.Top, "0000")
    MonitorHandleEnumCallback = 1&  ' Count them all.
End Function

' vTable entry #2 (of the coder procs), as it's the SECOND coder vTable item in this module ... and there are 4 arguments.
Private Function PrimaryMonitorHandleEnumCallback(ByVal hMonitor As Long, ByVal hdcMonitor As Long, uRect As RECT, dwData As Long) As Long
    Dim MonInfo As MONITORINFO
    MonInfo.cbSize = LenB(MonInfo)
    GetMonitorInfo hMonitor, MonInfo
    If MonInfo.dwFlags = &H1& Then
        dwData = hMonitor
        PrimaryMonitorHandleEnumCallback = 0& ' Found it.
    Else
        PrimaryMonitorHandleEnumCallback = 1& ' Keep looking.
    End If
End Function


Private Sub Form_Load()

    ' Form setup.
    Me.AutoRedraw = True
    Me.Font.Name = "Segoe UI Semibold"
    Me.Font.Size = 12


    ' Some variables we'll use.
    Dim ThunkSize   As Long
    Dim pVirtualMem As Long
    Dim pProcAddress    As Long

    ' First, just list our monitors.

    ' Get our procedure address and argument count.
    pProcAddress = Vb6ComCodeObjectAddressOfByNumber(Me, 1&)                ' <--- 1& is the vTable number.

    ' Make THUNK into executable memory.
    pVirtualMem = AddressOfThunkForComProc(Me, pProcAddress, 4&, ThunkSize) ' <--- 4& is the argument count.

    ' Count our monitors ... their handles are listed in the callback.
    Me.Print "Monitor handle(s):"
    Dim iMonitorCount As Long
    EnumDisplayMonitors 0&, ByVal 0&, pVirtualMem, iMonitorCount
    Me.Print "Total monitor count: "; CStr(iMonitorCount)

    ' Give our thunk memory back.
    FreeTheThunk pVirtualMem, ThunkSize



    ' Now let's use a callback to find our primary monitor.
    
    ' Get our procedure address and argument count.
    pProcAddress = Vb6ComCodeObjectAddressOfByNumber(Me, 2&)                ' <--- 2& is the vTable number.

    ' Make THUNK into executable memory.
    pVirtualMem = AddressOfThunkForComProc(Me, pProcAddress, 4&, ThunkSize) ' <--- 4& is the argument count.

    ' Get primary monitor handle using callback.
    Dim hPrimaryMonitor As Long
    EnumDisplayMonitors 0&, ByVal 0&, pVirtualMem, hPrimaryMonitor
    Me.Print "Handle of primary monitor: "; CStr(hPrimaryMonitor)

    ' Give our thunk memory back.
    FreeTheThunk pVirtualMem, ThunkSize


End Sub




' *************************************************
' API CallBack to object ... support procedures.
' *************************************************

Private Function Vb6ComCodeObjectAddressOfByNumber(ByVal o As Object, ByVal iCoderProcNumber As Long) As Long
    ' The caller is responsible for knowing how to use the returned address, or crash may result.
    ' COM/code objects have one (or two) hidden arguments, one for the ObjPtr address, and,
    ' if it's a function (or property get), another for that return.  ObjPtr is first and return value is second.
    '
    ' The iCoderProcNumber starts at 1 and can be no larger than the iCoderCodedCount returned by Vb6ComCodeObjectVtableEntries.
    ' The procedure number is a CODER CODED procedure, including any events.
    '
    Dim iInterfaceCount As Long, iIntrinsicCount As Long, iCoderCodedCount As Long
    Call Vb6ComCodeObjectVtableEntries(o, iInterfaceCount, iIntrinsicCount, iCoderCodedCount)
    If iCoderProcNumber < 1& Or iCoderProcNumber > iCoderCodedCount Then Exit Function    ' This will catch bad objects.
    '
    ' Return the address in the vTable.
    Dim pVtable As Long:    GetMem4 ByVal ObjPtr(o), pVtable        ' Get pointer to start of vTable.
    pVtable = pVtable + (iInterfaceCount + iIntrinsicCount) * 4&    ' Jump over the interface and intrinsic entries. 4& is 4 bytes per pointer (32-bit).
    pVtable = pVtable + (iCoderProcNumber - 1&) * 4&                ' And now go up to our specified procedure in the vTable.
    GetMem4 ByVal pVtable, Vb6ComCodeObjectAddressOfByNumber        ' Pointer into actual code (the member).
End Function

Private Function Vb6ComCodeObjectVtableEntries(o As Object, Optional iInterfaceCount As Long, Optional iIntrinsicCount As Long, Optional iCoderCodedCount As Long) As Long
    ' Return is the TOTAL, with Optional arguments returned as individual pieces.
    '
    If Not ObjectIsVb6ComCodeModule(o) Then Exit Function
    '
    Dim pVtable As Long:        GetMem4 ByVal ObjPtr(o), pVtable
    Dim ptObjectInfo As Long:   GetMem4 ByVal pVtable - 4&, ptObjectInfo
    iInterfaceCount = 7&                                    ' IUnknown and IDispatch.
    GetMem2 ByVal ptObjectInfo + &H62&, iIntrinsicCount     ' Out of the tObjectInfo structure (wPCodeCount).
    GetMem2 ByVal ptObjectInfo + &H60&, iCoderCodedCount    ' Out of the tObjectInfo structure (wMethodLinkCount).
    Vb6ComCodeObjectVtableEntries = iInterfaceCount + iIntrinsicCount + iCoderCodedCount
End Function

Private Sub FreeTheThunk(ByVal pVirtualMem As Long, ByVal iThunkSize As Long)
    Const MEM_RELEASE As Long = &H8000&
    VirtualFree pVirtualMem, iThunkSize, MEM_RELEASE
End Sub

Private Function AddressOfThunkForComProc(ByVal o As Object, _
                                         ByVal ComProcAddress As Long, _
                                         ByVal ComProcArgCount As Long, _
                                         ByRef iThunkSize As Long) As Long
    ' This makes a thunk and places it into a Byte array.
    ' This thunk is specifically for allowing COM procedures to be called as standard BAS procedures.
    '
    ' ComProcAddress:   The actual address of the COM procedure, typically dug out of the vTable.
    '
    ' ComObjPtr:        Usually comes out of ObjPtr(TheObject) for an instantiated VB6 COM/Code object (Form, Class, UC, PropertyPage, or DataReport).
    '
    ' ComProcArgCount:  The number of arguments "seen" in the code of the COM procedure.
    '                   It's MANDATORY that this be a number from 1 to 4 !!!!
    '                   It is assumed all the arguments are either ByRef or 4-byte ByVal arguments.
    '                   Anything else will fail and probably cause a crash.
    '
    ' iThunkSize:       Returned, and it must be used with VirtualFree to free this thunk memory.
    '                   The FreeTheThunk procedure is setup for this.
    '
    If ComProcAddress = 0& Or ObjPtr(o) = 0& Then Exit Function
    If ComProcArgCount < 1& Or ComProcArgCount > 4& Then Exit Function
    Dim bb() As Byte
    '
    ' Auto generated by Elroy's thunk maker.
    ReDim bb(44)
    ' ;
    ' ; The idea here is to convert a call that thinks it's a regular BAS call
    ' ; into a call that can go into a COM object procedure.  It is assumed that
    ' ; all arguments are either ByRef or 4-byte ByVal parameters.  It is
    ' ; further assumed that it's a Function that returns a Long.  This should
    ' ; cover the vast majority of API callbacks as well as subclassing.
    ' ;
    ' ; This thunk can handle one, two, three, or four incoming arguments.
    ' ; As an example, we'll assume two incoming arguments.  In a BAS module,
    ' ; such a function would look like the following:
    ' ;
    ' ;   Function OurCallBack(ByRef Arg1 As Long, ByRef Arg2 As Long) As Long
    ' ;
    ' ; When in a COM object, under the hood, this would be transformed as follows:
    ' ;
    ' ;   Function OurCallBack(ByVal OurObjPtr As Long, ByRef Arg1 As Long, _
    ' ;                        ByRef Arg2 As Long, ByRef FnRet As Long) As HRESULT
    ' ;
    ' ; So, to treat it as a BAS module call, we've got to add the OurObjPtr and
    ' ; deal with the return as an argument.  We just discard the HRESULT return.
    ' ;
    bb(0) = &H55                                                                        ' push    ebp                 ; Save base pointer, always done.
    bb(1) = &H89: bb(2) = &HE5                                                          ' mov     ebp, esp            ; Save stack pointer in ebp.
    bb(3) = &H83: bb(4) = &HEC: bb(5) = &H4                                             ' sub     esp, 4              ; Allocate 4 bytes of storage for local variables.
    ' ;
    ' ; We now start setting up for the COM procedure call.
    ' ;
    bb(6) = &H89: bb(7) = &HE8                                                          ' mov     eax, ebp            ; Base pointer into eax.
    bb(8) = &H83: bb(9) = &HE8: bb(10) = &H4                                            ' sub     eax, 4              ; Address for COM proc's FnRet to return.
    bb(11) = &H50                                                                       ' push    eax                 ;   and pushed on the stack for ByRef return.
    bb(12) = &HFF: bb(13) = &H75: bb(14) = &H14                                         ' push    [ebp + 20]          ; Arg4 onto stack.  Possibly patch with NOP.
    bb(15) = &HFF: bb(16) = &H75: bb(17) = &H10                                         ' push    [ebp + 16]          ; Arg3 onto stack.  Possibly patch with NOP.
    bb(18) = &HFF: bb(19) = &H75: bb(20) = &HC                                          ' push    [ebp + 12]          ; Arg2 onto stack.  Possibly patch with NOP.
    bb(21) = &HFF: bb(22) = &H75: bb(23) = &H8                                          ' push    [ebp + 8]           ; Arg1 onto stack.  At least one arg is required.
    bb(24) = &H68: bb(25) = &H55: bb(26) = &H55: bb(27) = &H55: bb(28) = &H55           ' push    0x55555555          ; We'll patch this up with the OurObjPtr address.
    ' ;
    bb(29) = &HB8: bb(30) = &H66: bb(31) = &H66: bb(32) = &H66: bb(33) = &H66           ' mov     eax, 0x66666666     ; We'll patch this up with the address to the COM procedure.
    bb(34) = &HFF: bb(35) = &HD0                                                        ' call    eax                 ; Call the COM procedure.
    ' ;
    ' ; We're back, so patch up the return from COM proc, and return.
    ' ; The stack will take care of itself, as ByVal are discarded,
    ' ; and ByRef were passed through with the same address.
    ' ;
    bb(36) = &H8B: bb(37) = &H45: bb(38) = &HFC                                         ' mov     eax, [ebp - 4]      ; Return the last argument's value as our return.
    bb(39) = &H89: bb(40) = &HEC                                                        ' mov     esp, ebp            ; Restore stack pointer from base pointer.
    bb(41) = &H5D                                                                       ' pop     ebp                 ; Restore base pointer.
    bb(42) = &HC2: bb(43) = &H8: bb(44) = &H0                                           ' ret     8                   ; Reset stack (for passed args) and return.  Patch up 8 for exact number of args (x 4).
    '
    ' We will need the size.
    iThunkSize = UBound(bb) - LBound(bb) + 1&
    '
    ' If not four arguments, blank out unused.
    If ComProcArgCount < 4& Then bb(12) = &H90: bb(13) = &H90: bb(14) = &H90    ' &H90 = NOP (no operation)
    If ComProcArgCount < 3& Then bb(15) = &H90: bb(16) = &H90: bb(17) = &H90
    If ComProcArgCount < 2& Then bb(18) = &H90: bb(19) = &H90: bb(20) = &H90
    ' At least one argument is required, so bb(21) thru bb(23) don't change.
    '
    ' Patch up our two supplied addresses.
    CopyMemory bb(25), ObjPtr(o), 4&            ' The ObjPtr() that's needed as an argument.
    CopyMemory bb(30), ComProcAddress, 4&       ' Where the actual call to the COM procedure is being made.
    '
    ' Patch up the return for how much of the stack to reset.
    bb(43) = ComProcArgCount * 4&
    '
    ' Get some executable memory.  Make sure we release it when we're done.
    Const MEM_COMMIT                    As Long = &H1000&
    Const PAGE_EXECUTE_READWRITE        As Long = &H40&
    AddressOfThunkForComProc = VirtualAlloc(0&, iThunkSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    '
    ' Move thunk into executable memory, and return our virtual memory's address.
    CopyMemory ByVal AddressOfThunkForComProc, bb(0&), iThunkSize
End Function

Private Function ObjectIsVb6ComCodeModule(ByRef o As IUnknown) As Boolean
    ' If it's an instantiated Class, Form, UC, PropPage, DataReport, returns TRUE, else FALSE.
    If ObjPtr(o) = 0& Then Exit Function                    ' Make sure "something" is instantiated.
    Dim aGUID(1&) As Currency                               ' Just to get 16 easily accessible bytes.
    aGUID(0&) = 128347367577987.1845@                       ' Const AreYouABasicInstance As String = "{0B6C9465-D082-11CF-8B4F-00A0C90F2704}"
    aGUID(1&) = 29922525889064.5387@                        ' turned into two numbers stuffed into our Currency array.
    ObjectIsVb6ComCodeModule = vbaCheckType(o, aGUID(0&))   ' Check and see if we are this "TYPE" (Class, Form, UC, PropPage, or DataRep).
End Function
```

Be sure to notice the use of `Vb6ComCodeObjectAddressOfByNumber` rather than `Vb6ComCodeObjectAddressOf`.

And here's what the output looks like:

```
| Form1                                                 |
|-------------------------------------------------------|
| Monitor handle(s):                                    |
|  65539 Left: 3440 Width: 1920 Top: 0000 Height: 1080  |
|  65537 Left: 0000 Width: 3440 Top: 0000 Height: 1440  |
| Total monitor count: 2                                |
| Handle of primary monitor: 65537                      |
```

## Review:

Ok, that's about it. Let's review a bit and I'll also give you a ZIP file with all the support procedures we've discussed.

Support Procedures We've Discussed:
* `ObjectIsVb6ComCodeModule`
* `VtableOffsetForVb6ComMethod`
* `Vb6ComCodeObjectAddressOf`
* `AddressOfThunkForComProc`
* `FreeTheThunk`
* `Vb6ComCodeObjectVtableEntries`
* `Vb6ComCodeObjectAddressOfByNumber`

And here's a zipped BAS file with all of these, including their API support: Attachment 186075 (CallingVb6ObjectsWithByNum.BAS)

These are in a BAS, and they're all declared as `Public`. But don't forget that they can be pulled into any VB6 code object (`Form`, `Class`, `UserControl`, `PropPage`, or `DataReport`) and declare as `Private`, therefore wrapping everything up into a single module ... which is really the whole objective of this.

And just to review the steps to using it all:

If you're willing to declare your CallBack procedure(s) as `Public`:

1. Use `Vb6ComCodeObjectAddressOf` to get your COM procedure address and argument count.
2. Use `AddressOfThunkForComProc` to make a thunk, getting its address (and size).
3. Make your API callback call, using the thunk address for the callback address.
4. When done making callbacks, use `FreeTheThunk` to free the thunk's memory.

If you must declare your CallBack procedure(s) as `Private`:

1. Use the outlined rules (above) to figure out the coder's vTable entry number for your callback procedure.
2. Use `Vb6ComCodeObjectAddressOfByNumber` to get your COM procedure address.
3. Count the argument(s) in your callback procedure.
4. Use `AddressOfThunkForComProc` to make a thunk, getting its address (and size).
5. Make your API callback call, using the thunk address for the callback address.
6. When done making callbacks, use `FreeTheThunk` to free the thunk's memory.

## Some Friendly Advice:

If/when I want to use this stuff, I'll always be "debugging" my CallBack testing by using a regular BAS module for the CallBack procedure. Then, once I've got all that running the way I'd like, then I'll move it all into my COM (Class, Form, etc.) module. That way, if something doesn't work, you'll know whether it was your coding of the CallBack, or your implementation of these CallBack-in-COM procedures.

Enjoy,
Elroy

----

## Notes from Sancarn

* Elroy's attachments are included in VB6 folder.
* An LLM Translated approaches for VBA7 are included in the VBA7 folder