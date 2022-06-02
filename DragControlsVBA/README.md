SOURCE: https://www.vbforums.com/showthread.php?888843-Load-image-into-STATIC-control-Win32&p=5496575&viewfull=1#post5496575

@Steve Grant @sancarn

File Demo

This is what I ended up with:

- Allows moving and\or copying the images.
- The static control is semi-transparent and confined within the bounderies of the parent form.
- A colored border frame is drawn around the static control.
- The cursor changes dynamically depending on moving the images, copying them (Holding CTRL key down) or when -the image is being moved outside the parent form. (custom cursor not showing on the Gif below but works as expected in the file demo above)
- Right-click context menu for deleting the images.
- A label control can optionally be integrated into the class for displaying the current activity.

This is the Only Class Method that hooks the images :
Public Sub HookControl(ByVal ThisClassInstance As cls_DraggableControl, ByVal Ctrl As Control, Optional ByVal UILabel As Control)

![img](https://www.vbforums.com/images/ieimages/2020/10/5.gif)

There remains one issue that I haven't been able to solve and that is the transparency of the layered static control doesn't work if the machine's [Desktop Composition](https://docs.microsoft.com/en-us/windows/win32/api/_dwm/) has been disabled.

I have made a few attempts to resolve this problem by changing the Static Control owner, by using the UpdateLayeredWindow API to update the control ... etc but with no success.

Does anyone have an idea if\how this remaining issue could be addressed ?

Thanks.
