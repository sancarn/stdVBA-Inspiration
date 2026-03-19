Attribute VB_Name = "modCallingVb6Objects"
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

Public Sub FreeTheThunk(ByVal pVirtualMem As Long, ByVal iThunkSize As Long)
    Const MEM_RELEASE As Long = &H8000&
    VirtualFree pVirtualMem, iThunkSize, MEM_RELEASE
End Sub

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
    CopyMemory ByVal AddressOfThunkForComProc, bb(0&), iThunkSize
End Function

Public Function ObjectIsVb6ComCodeModule(ByRef o As IUnknown) As Boolean
    ' If it's an instantiated Class, Form, UC, PropPage, DataReport, returns TRUE, else FALSE.
    If ObjPtr(o) = 0& Then Exit Function                    ' Make sure "something" is instantiated.
    Dim aGUID(1&) As Currency                               ' Just to get 16 easily accessible bytes.
    aGUID(0&) = 128347367577987.1845@                       ' Const AreYouABasicInstance As String = "{0B6C9465-D082-11CF-8B4F-00A0C90F2704}"
    aGUID(1&) = 29922525889064.5387@                        ' turned into two numbers stuffed into our Currency array.
    ObjectIsVb6ComCodeModule = vbaCheckType(o, aGUID(0&))   ' Check and see if we are this "TYPE" (Class, Form, UC, PropPage, or DataRep).
End Function


' **********************************************
' If we want callback procedure to stay Private,
' here are procedures for getting vTable address
' by vTable procedure count.
' **********************************************


Public Function Vb6ComCodeObjectAddressOfByNumber(ByVal o As Object, ByVal iCoderProcNumber As Long) As Long
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

