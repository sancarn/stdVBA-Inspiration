# Test 1 - Different variables and callstack

## Inputs

```vb
Sub t0()
  Debug.Print "Sub 0     -> " & getPtr(AddressOf t0) & " - " & Hex(getPtr(AddressOf t0))
  Debug.Print "Sub 1     -> " & getPtr(AddressOf t1) & " - " & Hex(getPtr(AddressOf t1))
  Debug.Print "Sub 2     -> " & getPtr(AddressOf t2) & " - " & Hex(getPtr(AddressOf t2))

  Debug.Print

  Debug.Print "===t0==="
  Dim nByte       As Byte
  Dim nInteger    As Integer
  Dim nLong       As Long
  Dim nLongLong   As LongLong    '64-bit only
  Dim nLongPtr    As LongPtr     'Pointer-sized (32/64-bit)
  Dim nSingle     As Single
  Dim nDouble     As Double
  Dim nCurrency   As Currency
  Dim bFlag       As Boolean
  Dim dtWhen      As Date
  Dim sText       As String
  Dim vAnything   As Variant
  Dim oObject     As Object
  Dim colItems    As New Collection

  Debug.Print "nByte     -> " & VarPtr(nByte) & " - " & Hex(VarPtr(nByte))
  Debug.Print "nInteger  -> " & VarPtr(nInteger) & " - " & Hex(VarPtr(nInteger))
  Debug.Print "nLong     -> " & VarPtr(nLong) & " - " & Hex(VarPtr(nLong))
  Debug.Print "nLongLong -> " & VarPtr(nLongLong) & " - " & Hex(VarPtr(nLongLong))
  Debug.Print "nLongPtr  -> " & VarPtr(nLongPtr) & " - " & Hex(VarPtr(nLongPtr))
  Debug.Print "nSingle   -> " & VarPtr(nSingle) & " - " & Hex(VarPtr(nSingle))
  Debug.Print "nDouble   -> " & VarPtr(nDouble) & " - " & Hex(VarPtr(nDouble))
  Debug.Print "nCurrency -> " & VarPtr(nCurrency) & " - " & Hex(VarPtr(nCurrency))
  Debug.Print "bFlag     -> " & VarPtr(bFlag) & " - " & Hex(VarPtr(bFlag))
  Debug.Print "dtWhen    -> " & VarPtr(dtWhen) & " - " & Hex(VarPtr(dtWhen))
  Debug.Print "sText     -> " & VarPtr(sText) & " - " & Hex(VarPtr(sText))
  Debug.Print "vAnything -> " & VarPtr(vAnything) & " - " & Hex(VarPtr(vAnything))
  Debug.Print "oObject   -> " & VarPtr(oObject) & " - " & Hex(VarPtr(oObject))
  Debug.Print "colItems  -> " & VarPtr(colItems) & " - " & Hex(VarPtr(colItems))

  Call t1
End Sub

Private Sub t1()
  Debug.Print "===t1==="
  Dim nByte       As Byte
  Dim nInteger    As Integer
  Dim nLong       As Long
  Dim nLongLong   As LongLong    '64-bit only
  Dim nLongPtr    As LongPtr     'Pointer-sized (32/64-bit)
  Dim nSingle     As Single
  Dim nDouble     As Double
  Dim nCurrency   As Currency
  Dim bFlag       As Boolean
  Dim dtWhen      As Date
  Dim sText       As String
  Dim vAnything   As Variant
  Dim oObject     As Object
  Dim colItems    As New Collection

  Debug.Print "nByte     -> " & VarPtr(nByte) & " - " & Hex(VarPtr(nByte))
  Debug.Print "nInteger  -> " & VarPtr(nInteger) & " - " & Hex(VarPtr(nInteger))
  Debug.Print "nLong     -> " & VarPtr(nLong) & " - " & Hex(VarPtr(nLong))
  Debug.Print "nLongLong -> " & VarPtr(nLongLong) & " - " & Hex(VarPtr(nLongLong))
  Debug.Print "nLongPtr  -> " & VarPtr(nLongPtr) & " - " & Hex(VarPtr(nLongPtr))
  Debug.Print "nSingle   -> " & VarPtr(nSingle) & " - " & Hex(VarPtr(nSingle))
  Debug.Print "nDouble   -> " & VarPtr(nDouble) & " - " & Hex(VarPtr(nDouble))
  Debug.Print "nCurrency -> " & VarPtr(nCurrency) & " - " & Hex(VarPtr(nCurrency))
  Debug.Print "bFlag     -> " & VarPtr(bFlag) & " - " & Hex(VarPtr(bFlag))
  Debug.Print "dtWhen    -> " & VarPtr(dtWhen) & " - " & Hex(VarPtr(dtWhen))
  Debug.Print "sText     -> " & VarPtr(sText) & " - " & Hex(VarPtr(sText))
  Debug.Print "vAnything -> " & VarPtr(vAnything) & " - " & Hex(VarPtr(vAnything))
  Debug.Print "oObject   -> " & VarPtr(oObject) & " - " & Hex(VarPtr(oObject))
  Debug.Print "colItems  -> " & VarPtr(colItems) & " - " & Hex(VarPtr(colItems))
  Call t2
End Sub
Private Sub t2()
  Debug.Print "===t2==="
  Dim nByte       As Byte
  Dim nInteger    As Integer
  Dim nLong       As Long
  Dim nLongLong   As LongLong    '64-bit only
  Dim nLongPtr    As LongPtr     'Pointer-sized (32/64-bit)
  Dim nSingle     As Single
  Dim nDouble     As Double
  Dim nCurrency   As Currency
  Dim bFlag       As Boolean
  Dim dtWhen      As Date
  Dim sText       As String
  Dim vAnything   As Variant
  Dim oObject     As Object
  Dim colItems    As New Collection

  Debug.Print "nByte     -> " & VarPtr(nByte) & " - " & Hex(VarPtr(nByte))
  Debug.Print "nInteger  -> " & VarPtr(nInteger) & " - " & Hex(VarPtr(nInteger))
  Debug.Print "nLong     -> " & VarPtr(nLong) & " - " & Hex(VarPtr(nLong))
  Debug.Print "nLongLong -> " & VarPtr(nLongLong) & " - " & Hex(VarPtr(nLongLong))
  Debug.Print "nLongPtr  -> " & VarPtr(nLongPtr) & " - " & Hex(VarPtr(nLongPtr))
  Debug.Print "nSingle   -> " & VarPtr(nSingle) & " - " & Hex(VarPtr(nSingle))
  Debug.Print "nDouble   -> " & VarPtr(nDouble) & " - " & Hex(VarPtr(nDouble))
  Debug.Print "nCurrency -> " & VarPtr(nCurrency) & " - " & Hex(VarPtr(nCurrency))
  Debug.Print "bFlag     -> " & VarPtr(bFlag) & " - " & Hex(VarPtr(bFlag))
  Debug.Print "dtWhen    -> " & VarPtr(dtWhen) & " - " & Hex(VarPtr(dtWhen))
  Debug.Print "sText     -> " & VarPtr(sText) & " - " & Hex(VarPtr(sText))
  Debug.Print "vAnything -> " & VarPtr(vAnything) & " - " & Hex(VarPtr(vAnything))
  Debug.Print "oObject   -> " & VarPtr(oObject) & " - " & Hex(VarPtr(oObject))
  Debug.Print "colItems  -> " & VarPtr(colItems) & " - " & Hex(VarPtr(colItems))
End Sub

Private Function getPtr(ByVal ptr As LongPtr) As LongPtr
  getPtr = ptr
End Function
```

## Outputs

```
Sub 0     -> 2353850278484 - 2240C68E254
Sub 1     -> 2353850276692 - 2240C68DB54
Sub 2     -> 2353850277588 - 2240C68DED4

===t0===
nByte     -> 2353875672352 - 2240DEC5D20
nInteger  -> 2353875672344 - 2240DEC5D18
nLong     -> 2353875672336 - 2240DEC5D10
nLongLong -> 2353875672328 - 2240DEC5D08
nLongPtr  -> 2353875672320 - 2240DEC5D00
nSingle   -> 2353875672312 - 2240DEC5CF8
nDouble   -> 2353875672304 - 2240DEC5CF0
nCurrency -> 2353875672296 - 2240DEC5CE8
bFlag     -> 2353875672288 - 2240DEC5CE0
dtWhen    -> 2353875672280 - 2240DEC5CD8
sText     -> 2353875672272 - 2240DEC5CD0
vAnything -> 2353875672248 - 2240DEC5CB8
oObject   -> 2353875672240 - 2240DEC5CB0
colItems  -> 2353875672232 - 2240DEC5CA8
===t1===
nByte     -> 2353875672216 - 2240DEC5C98
nInteger  -> 2353875672208 - 2240DEC5C90
nLong     -> 2353875672200 - 2240DEC5C88
nLongLong -> 2353875672192 - 2240DEC5C80
nLongPtr  -> 2353875672184 - 2240DEC5C78
nSingle   -> 2353875672176 - 2240DEC5C70
nDouble   -> 2353875672168 - 2240DEC5C68
nCurrency -> 2353875672160 - 2240DEC5C60
bFlag     -> 2353875672152 - 2240DEC5C58
dtWhen    -> 2353875672144 - 2240DEC5C50
sText     -> 2353875672136 - 2240DEC5C48
vAnything -> 2353875672112 - 2240DEC5C30
oObject   -> 2353875672104 - 2240DEC5C28
colItems  -> 2353875672096 - 2240DEC5C20
===t2===
nByte     -> 2353875671936 - 2240DEC5B80
nInteger  -> 2353875671928 - 2240DEC5B78
nLong     -> 2353875671920 - 2240DEC5B70
nLongLong -> 2353875671912 - 2240DEC5B68
nLongPtr  -> 2353875671904 - 2240DEC5B60
nSingle   -> 2353875671896 - 2240DEC5B58
nDouble   -> 2353875671888 - 2240DEC5B50
nCurrency -> 2353875671880 - 2240DEC5B48
bFlag     -> 2353875671872 - 2240DEC5B40
dtWhen    -> 2353875671864 - 2240DEC5B38
sText     -> 2353875671856 - 2240DEC5B30
vAnything -> 2353875671832 - 2240DEC5B18
oObject   -> 2353875671824 - 2240DEC5B10
colItems  -> 2353875671816 - 2240DEC5B08
```

## Conclusion

It seems that 2240DEC##### refers to the callstack and variables on the stack? We could probably take a snapshot of VBA to see if we find instances of these values...

# Test 2 - First variable across multiple methods

## Inputs

Simply.

```vb
Sub t0()
  Dim v
  Debug.Print Hex(VarPtr(v))
End Sub

Sub t1()
  Dim v
  Debug.Print Hex(VarPtr(v))
End Sub

Sub t2()
  Dim v
  Debug.Print Hex(VarPtr(v))
End Sub
```

## Outputs

No matter which sub you call, the output is always the same!

```
2240DEC5DA0
2240DEC5DA0
2240DEC5DA0
```

This must mean that the callstack starts at the same address every time...

This means that we could procedurally search to find the callstack from any given position in memory.

## Note:

Looks like, at least, across instances, the ptr is not the same. Another instance of excel gave the following pointers:

```
208364A3300
```

# Test 3 - query block of memory

## Inputs

```vb
Sub t0()
  Debug.Print "Sub 0     -> " & getPtr(AddressOf t0) & " - " & Hex(getPtr(AddressOf t0))
  Debug.Print "SubMem:      " & QueryBlock(AddressOf t0, -16) & "|" & QueryBlock(AddressOf t0, 16)
  Debug.Print "Sub 1     -> " & getPtr(AddressOf t1) & " - " & Hex(getPtr(AddressOf t1))
  Debug.Print "SubMem:      " & QueryBlock(AddressOf t1, -16) & "|" & QueryBlock(AddressOf t1, 16)
  Debug.Print "Sub 2     -> " & getPtr(AddressOf t2) & " - " & Hex(getPtr(AddressOf t2))
  Debug.Print "SubMem:      " & QueryBlock(AddressOf t2, -16) & "|" & QueryBlock(AddressOf t2, 16)



  Debug.Print

  Debug.Print "===t0==="
  Dim nByte       As Byte
  Dim nInteger    As Integer
  Dim nLong       As Long
  Dim nLongLong   As LongLong    '64-bit only
  Dim nLongPtr    As LongPtr     'Pointer-sized (32/64-bit)
  Dim nSingle     As Single
  Dim nDouble     As Double
  Dim nCurrency   As Currency
  Dim bFlag       As Boolean
  Dim dtWhen      As Date
  Dim sText       As String
  Dim vAnything   As Variant
  Dim oObject     As Object
  Dim colItems    As New Collection

  Debug.Print "nByte     -> " & VarPtr(nByte) & " - " & Hex(VarPtr(nByte))
  Debug.Print "nInteger  -> " & VarPtr(nInteger) & " - " & Hex(VarPtr(nInteger))
  Debug.Print "nLong     -> " & VarPtr(nLong) & " - " & Hex(VarPtr(nLong))
  Debug.Print "nLongLong -> " & VarPtr(nLongLong) & " - " & Hex(VarPtr(nLongLong))
  Debug.Print "nLongPtr  -> " & VarPtr(nLongPtr) & " - " & Hex(VarPtr(nLongPtr))
  Debug.Print "nSingle   -> " & VarPtr(nSingle) & " - " & Hex(VarPtr(nSingle))
  Debug.Print "nDouble   -> " & VarPtr(nDouble) & " - " & Hex(VarPtr(nDouble))
  Debug.Print "nCurrency -> " & VarPtr(nCurrency) & " - " & Hex(VarPtr(nCurrency))
  Debug.Print "bFlag     -> " & VarPtr(bFlag) & " - " & Hex(VarPtr(bFlag))
  Debug.Print "dtWhen    -> " & VarPtr(dtWhen) & " - " & Hex(VarPtr(dtWhen))
  Debug.Print "sText     -> " & VarPtr(sText) & " - " & Hex(VarPtr(sText))
  Debug.Print "vAnything -> " & VarPtr(vAnything) & " - " & Hex(VarPtr(vAnything))
  Debug.Print "oObject   -> " & VarPtr(oObject) & " - " & Hex(VarPtr(oObject))
  Debug.Print "colItems  -> " & VarPtr(colItems) & " - " & Hex(VarPtr(colItems))

  Call t1
End Sub

Private Sub t1()
  Debug.Print "===t1==="
  Dim nByte       As Byte
  Dim nInteger    As Integer
  Dim nLong       As Long
  Dim nLongLong   As LongLong    '64-bit only
  Dim nLongPtr    As LongPtr     'Pointer-sized (32/64-bit)
  Dim nSingle     As Single
  Dim nDouble     As Double
  Dim nCurrency   As Currency
  Dim bFlag       As Boolean
  Dim dtWhen      As Date
  Dim sText       As String
  Dim vAnything   As Variant
  Dim oObject     As Object
  Dim colItems    As New Collection

  Debug.Print "nByte     -> " & VarPtr(nByte) & " - " & Hex(VarPtr(nByte))
  Debug.Print "nInteger  -> " & VarPtr(nInteger) & " - " & Hex(VarPtr(nInteger))
  Debug.Print "nLong     -> " & VarPtr(nLong) & " - " & Hex(VarPtr(nLong))
  Debug.Print "nLongLong -> " & VarPtr(nLongLong) & " - " & Hex(VarPtr(nLongLong))
  Debug.Print "nLongPtr  -> " & VarPtr(nLongPtr) & " - " & Hex(VarPtr(nLongPtr))
  Debug.Print "nSingle   -> " & VarPtr(nSingle) & " - " & Hex(VarPtr(nSingle))
  Debug.Print "nDouble   -> " & VarPtr(nDouble) & " - " & Hex(VarPtr(nDouble))
  Debug.Print "nCurrency -> " & VarPtr(nCurrency) & " - " & Hex(VarPtr(nCurrency))
  Debug.Print "bFlag     -> " & VarPtr(bFlag) & " - " & Hex(VarPtr(bFlag))
  Debug.Print "dtWhen    -> " & VarPtr(dtWhen) & " - " & Hex(VarPtr(dtWhen))
  Debug.Print "sText     -> " & VarPtr(sText) & " - " & Hex(VarPtr(sText))
  Debug.Print "vAnything -> " & VarPtr(vAnything) & " - " & Hex(VarPtr(vAnything))
  Debug.Print "oObject   -> " & VarPtr(oObject) & " - " & Hex(VarPtr(oObject))
  Debug.Print "colItems  -> " & VarPtr(colItems) & " - " & Hex(VarPtr(colItems))
  Call t2
End Sub
Private Sub t2()
  Debug.Print "===t2==="
  Dim nByte       As Byte
  Dim nInteger    As Integer
  Dim nLong       As Long
  Dim nLongLong   As LongLong    '64-bit only
  Dim nLongPtr    As LongPtr     'Pointer-sized (32/64-bit)
  Dim nSingle     As Single
  Dim nDouble     As Double
  Dim nCurrency   As Currency
  Dim bFlag       As Boolean
  Dim dtWhen      As Date
  Dim sText       As String
  Dim vAnything   As Variant
  Dim oObject     As Object
  Dim colItems    As New Collection

  Debug.Print "nByte     -> " & VarPtr(nByte) & " - " & Hex(VarPtr(nByte))
  Debug.Print "nInteger  -> " & VarPtr(nInteger) & " - " & Hex(VarPtr(nInteger))
  Debug.Print "nLong     -> " & VarPtr(nLong) & " - " & Hex(VarPtr(nLong))
  Debug.Print "nLongLong -> " & VarPtr(nLongLong) & " - " & Hex(VarPtr(nLongLong))
  Debug.Print "nLongPtr  -> " & VarPtr(nLongPtr) & " - " & Hex(VarPtr(nLongPtr))
  Debug.Print "nSingle   -> " & VarPtr(nSingle) & " - " & Hex(VarPtr(nSingle))
  Debug.Print "nDouble   -> " & VarPtr(nDouble) & " - " & Hex(VarPtr(nDouble))
  Debug.Print "nCurrency -> " & VarPtr(nCurrency) & " - " & Hex(VarPtr(nCurrency))
  Debug.Print "bFlag     -> " & VarPtr(bFlag) & " - " & Hex(VarPtr(bFlag))
  Debug.Print "dtWhen    -> " & VarPtr(dtWhen) & " - " & Hex(VarPtr(dtWhen))
  Debug.Print "sText     -> " & VarPtr(sText) & " - " & Hex(VarPtr(sText))
  Debug.Print "vAnything -> " & VarPtr(vAnything) & " - " & Hex(VarPtr(vAnything))
  Debug.Print "oObject   -> " & VarPtr(oObject) & " - " & Hex(VarPtr(oObject))
  Debug.Print "colItems  -> " & VarPtr(colItems) & " - " & Hex(VarPtr(colItems))
End Sub

Private Function getPtr(ByVal ptr As LongPtr) As LongPtr
  getPtr = ptr
End Function

'Queries a block of memory and returns the hex representation of this section of memory
'@param ptr    - pointer to query
'@param length - length to query, if negative this returns the block prior to the ptr
Private Function QueryBlock(ByVal ptr As LongPtr, ByVal length As Long) As String
  If length < 1 Then
    ptr = ptr - length
    length = length * -1
  End If
  Dim b() As Byte
  ReDim b(1 To length)
  Call CopyMemory(VarPtr(b(1)), ptr, length)

  Dim ret As String: ret = ""
  Dim i As Long, h As String
  For i = 1 To length
    h = Hex(b(i))
    If Len(h) = 1 Then h = "0" & h
    ret = ret & h
  Next

  QueryBlock = ret
End Function
```

## Outputs

```
Sub 0     -> 2353850276804 - 2240C68DBC4
SubMem:      894C242048B8C0DB680C24020000480B|48894C240848895424104C894424184C
Sub 1     -> 2353850276916 - 2240C68DC34
SubMem:      894C242048B830DC680C24020000480B|48894C240848895424104C894424184C
Sub 2     -> 2353850277140 - 2240C68DD14
SubMem:      894C242048B810DD680C24020000480B|48894C240848895424104C894424184C

===t0===
nByte     -> 2353875672312 - 2240DEC5CF8
nInteger  -> 2353875672304 - 2240DEC5CF0
nLong     -> 2353875672296 - 2240DEC5CE8
nLongLong -> 2353875672288 - 2240DEC5CE0
nLongPtr  -> 2353875672280 - 2240DEC5CD8
nSingle   -> 2353875672272 - 2240DEC5CD0
nDouble   -> 2353875672264 - 2240DEC5CC8
nCurrency -> 2353875672256 - 2240DEC5CC0
bFlag     -> 2353875672248 - 2240DEC5CB8
dtWhen    -> 2353875672240 - 2240DEC5CB0
sText     -> 2353875672232 - 2240DEC5CA8
vAnything -> 2353875672208 - 2240DEC5C90
oObject   -> 2353875672200 - 2240DEC5C88
colItems  -> 2353875672192 - 2240DEC5C80
===t1===
nByte     -> 2353875672176 - 2240DEC5C70
nInteger  -> 2353875672168 - 2240DEC5C68
nLong     -> 2353875672160 - 2240DEC5C60
nLongLong -> 2353875672152 - 2240DEC5C58
nLongPtr  -> 2353875672144 - 2240DEC5C50
nSingle   -> 2353875672136 - 2240DEC5C48
nDouble   -> 2353875672128 - 2240DEC5C40
nCurrency -> 2353875672120 - 2240DEC5C38
bFlag     -> 2353875672112 - 2240DEC5C30
dtWhen    -> 2353875672104 - 2240DEC5C28
sText     -> 2353875672096 - 2240DEC5C20
vAnything -> 2353875672072 - 2240DEC5C08
oObject   -> 2353875672064 - 2240DEC5C00
colItems  -> 2353875672056 - 2240DEC5BF8
===t2===
nByte     -> 2353875671896 - 2240DEC5B58
nInteger  -> 2353875671888 - 2240DEC5B50
nLong     -> 2353875671880 - 2240DEC5B48
nLongLong -> 2353875671872 - 2240DEC5B40
nLongPtr  -> 2353875671864 - 2240DEC5B38
nSingle   -> 2353875671856 - 2240DEC5B30
nDouble   -> 2353875671848 - 2240DEC5B28
nCurrency -> 2353875671840 - 2240DEC5B20
bFlag     -> 2353875671832 - 2240DEC5B18
dtWhen    -> 2353875671824 - 2240DEC5B10
sText     -> 2353875671816 - 2240DEC5B08
vAnything -> 2353875671792 - 2240DEC5AF0
oObject   -> 2353875671784 - 2240DEC5AE8
colItems  -> 2353875671776 - 2240DEC5AE0
```

## Conclusion

So despite being defined in this order:

```
nByte     -> 2353875672312 - 2240DEC5CF8
nInteger  -> 2353875672304 - 2240DEC5CF0
nLong     -> 2353875672296 - 2240DEC5CE8
nLongLong -> 2353875672288 - 2240DEC5CE0
nLongPtr  -> 2353875672280 - 2240DEC5CD8
nSingle   -> 2353875672272 - 2240DEC5CD0
nDouble   -> 2353875672264 - 2240DEC5CC8
nCurrency -> 2353875672256 - 2240DEC5CC0
bFlag     -> 2353875672248 - 2240DEC5CB8
dtWhen    -> 2353875672240 - 2240DEC5CB0
sText     -> 2353875672232 - 2240DEC5CA8
vAnything -> 2353875672208 - 2240DEC5C90
oObject   -> 2353875672200 - 2240DEC5C88
colItems  -> 2353875672192 - 2240DEC5C80
```

The memory itself looks like this:

```
2240DEC5C80 (colItems)
2240DEC5C88 (oObject)
2240DEC5C90 (vAnything)
2240DEC5CA8 (sText)
2240DEC5CB0 (dtWhen)
2240DEC5CB8 (bFlag)
2240DEC5CC0 (nCurrency)
2240DEC5CC8 (nDouble)
2240DEC5CD0 (nSingle)
2240DEC5CD8 (nLongPtr)
2240DEC5CE0 (nLongLong)
2240DEC5CE8 (nLong)
2240DEC5CF0 (nInteger)
2240DEC5CF8 (nByte)
```


So we can see that the callstack is at the top of the memory, and the variables are at the bottom.


# Test 4 - Query memory above the callstack

## Inputs

```vb
Sub t0()
  Dim v
  Debug.Print "Var:" & Hex(VarPtr(v))
  Debug.Print "Sub:" & Hex(getPtr(AddressOf t0))
  
  Debug.Print QueryBlockAsString(VarPtr(v), -256)
  Debug.Print QueryBlockAsHex(VarPtr(v), -256))
End Sub










Private Function getPtr(ByVal ptr As LongPtr) As LongPtr
  getPtr = ptr
End Function

Private Function QueryBlockAsString(ByVal ptr As LongPtr, ByVal length As Long) As String
  Dim b: b = QueryBlock(ptr, length)
  QueryBlockAsString = b
End Function

'Queries a block of memory and returns the hex representation of this section of memory
'@param ptr    - pointer to query
'@param length - length to query, if negative this returns the block prior to the ptr
Private Function QueryBlockAsHex(ByVal ptr As LongPtr, ByVal length As Long) As String
  Dim b: b = QueryBlock(ptr, length)
  
  Dim ret As String: ret = ""
  Dim i As Long, h As String
  For i = 1 To length
    h = Hex(b(i))
    If Len(h) = 1 Then h = "0" & h
    ret = ret & h
  Next
  
  QueryBlockAsHex = ret
End Function

'Queries a block of memory and returns the bytes representation of this section of memory
'@param ptr    - pointer to query
'@param length - length to query, if negative this returns the block prior to the ptr
Private Function QueryBlock(ByVal ptr As LongPtr, ByVal length As Long) As Byte()
  If length < 1 Then
    ptr = ptr - length
    length = length * -1
  End If
  Dim b() As Byte
  ReDim b(1 To length)
  Call CopyMemory(VarPtr(b(1)), ptr, length)
  QueryBlock = b
End Function
```

## Outputs

```
Var:2240DEC5DA0
Sub:2240D1B0984
                    ??      ??O?                                                                  ??????? ???   ??? ???     
00000000000000000000000000000000000000000000000000000000000000000000000000000000FEFFFFFF000000000000000000000000A64BBABA4C01009000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000400000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000F00000000000000BD4BA3BA3502009028549A98FF7F0000B093D17E240200000300000002000000D050FF6E24020000B096D17E240200000000000000000000


```
