# SendInput 
## Description

For all the complaints that Vista broke this, and Vista broke that, Vista really didn't break Classic VB all that badly. That said, one VB statement did truly get hammered. This sample provides a drop-in replacement for the standard SendKeys statement, and should work just fine in all the environments (VB5, VB6 IDE, VBA other than Office 2007) where this broke. I believe it's written to exactly emulate SendKeys, without exception \<g\>.

Well, all but one "feature" of SendKeys, at any rate -- the Wait parameter. Can anyone tell me what that's all about? My routine accepts, but ignores, this parameter.

As you'll see, using MySendKeys is absolutely identical to native SendKeys:

```vb
Private Sub Command1_Click()
   Text1.SetFocus
   Call MySendKeys(Text2.Text)
   DoEvents
   Command1.SetFocus
End Sub

Private Sub Command2_Click()
   Text1.SetFocus
   Call VBA.SendKeys(Text2.Text)
   DoEvents
   Command2.SetFocus
End Sub

Private Sub Form_Load()
   Text2.Text = "{home}+{end}Testing123+(123)"
End Sub
```

One change you may want to think about making would be to rename the MySendKeys subroutine to SendKeys. That will avoid having to change anything else in your code, as it will override the SendKeys statement in the VBA object library from that point onward. In order to use the original SendKeys, you'll then have to prefix it with "VBA." as shown above.

## Unicode Support

Following several requests from readers, I decided to give a shot at supporting SendKeys like functionality with Unicode characters. Amazingly, this didn't seem to require extraordinary modification of the routines in MSendInput at all. The trickiest part, I'll admit, was trying to understand how y'all might actually want to use this functionality, and offering a convenient means to do so. I am, unfortunately, a prisoner of geography. Unicode simply isn't much of an issue here. So the best I can do is try to envision what might be useful, and ask a lot of questions. Hope this succeeds.

The basic idea is that you use MySendKeys just as before, embedding Unicode characters in your send string as you wish. The only catch is, embarrassingly, how to identify an "actual" Unicode character within the BSTR. We had a lot of discussion about this online, and no consensus was ever really reached. I went with my gut, and decided to treat anything outside the AscW range of 0-255 as Unicode. The actual test looks like this:

```vb
Private Sub ProcessChar(this As String)
   Dim code As Integer
   Dim vk As Integer
   Dim capped As Boolean

   ' Determine whether we need to treat as Unicode.
   code = AscW(this)
   If code >= 0 And code < 256 Then 'ascii
      ' Add input events for single character, taking capitalization
      ' into account.  HiByte will contain the shift state, and LoByte
      ' will contain the key code.
      vk = VkKeyScan(Asc(this))
      capped = CBool(ByteHi(vk) And 1)
      vk = ByteLo(vk)
      Call StuffBuffer(vk, capped)
   Else 'unicode
      Call StuffBufferW(code)
   End If
End Sub
```

I think that snippet calls out the other initial concern I had with this approach. SendInput doesn't offer any way to control the state of shift keys for Unicode characters. As you can see, the StuffBufferW function offers no parameter for that. I hope that won't be a problem. Anyway, that's pretty much it! I'll be interested in hearing how/if this works for you.

Update: Well, it turns out messing with Unicode is a bit trickier than I originally thought. There's now an exception build into the routine above for cases where the character doesn't map to anything on the keyboard. In those cases, we'll automatically call StuffBufferW now. Also needed to add a special case for the tilde character, which is an alias for vbKeyReturn in the native SendKeys. And then, I was reminded of AltGr behavior on some (a lot of?) European keyboard layouts. So that got wrapped in as well. Here's the updated routine:

```vb
Private Sub ProcessChar(this As String)
   Dim code As Integer
   Dim vk As Integer
   Dim capped As Boolean
   Dim AltGr As Boolean

   ' Determine whether we need to treat as Unicode.
   code = AscW(this)
   If code >= 0 And code < 256 Then 'ascii
      ' MODIFIED 16-Dec-2009:
      ' Special case for tilde character!
      If this = "~" Then
         vk = vbKeyReturn
      Else
         vk = VkKeyScan(Asc(this))
      End If

      ' Not all chars (in particular 128-255) will have direct keyboard
      ' translations, so treat those as Unicode if need be.
      If vk = -1 Then
         ' ADDED 16-Dec-2009
         Call StuffBufferW(code)
      Else
         ' Add input events for single character, taking capitalization
         ' into account.  HiByte will contain the shift state, and LoByte
         ' will contain the key code.
         capped = CBool(ByteHi(vk) And 1)
         ' ADDED 21-Dec-2009
         ' If CAPSLOCK is toggled on, the hibyte will be the inverse of
         ' what it ought to be to properly recreate the input string,
         ' as the SHIFT key would need to be depressed to compensate.
         If CapsLock() Then
            Select Case this
               Case "A" To "Z", "a" To "z"
                  capped = Not capped
            End Select
         End If
         ' ADDED 02-Apr-2010
         ' Some keyboard layouts have an AltGr key for special characters
         ' which comes through as CTRL+ALT here. Check for that.
         AltGr = CBool(ByteHi(vk) And 6)
         ' Proceed to stuff the keycode and capitalization into buffer.
         vk = ByteLo(vk)
         Call StuffBuffer(vk, capped, , AltGr)
      End If
   Else 'unicode
      Call StuffBufferW(code)
   End If
End Sub
```

And, of course, the StuffBuffer routine was modified for AltGr, as well. It now depresses vbKeyControl and vbKeyMenu (ALT) before such characters, and releases them afterwards. The point is, be sure you have the most recent code, if you're experiencing problems with non-ASCII characters.

## Note on Conditional Compilation

The MSendInput.bas module contains a conditional compilation constant that determines whether the native language implementation of Split() is compiled in or not. If this constant is defined as False, as it is in the download, the module may be used in VB5 as well as VB6 and VBA.

Additional VBA support is offered with a block of conditionally defined constants that are intrinsic in VB5/6 but missing in VBA.

```
Module	       | Library  |  Function
---------------|----------|-------------------|
MSendInput.bas | kernel32 |  GetVersionEx
               | user32   |  GetKeyboardLayout
               |          |  GetKeyState
               |          |  MapVirtualKeyEx
               |          |  SendInput
               |          |  VkKeyScan
```
