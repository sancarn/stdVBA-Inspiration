sub DumpInfo(file)
Dim img

    set img = WScript.CreateObject("WIA.ImageFile")
    
    img.LoadFile file

    wscript.echo "============================================================"
    wscript.echo "Image Info for " & file
    wscript.echo "------------------------------------------------------------"
    wscript.echo "Width = " & img.Width
    wscript.echo "Height = " & img.Height
    wscript.echo "Depth = " & img.PixelDepth
    wscript.echo "HorizontalResolution = " & img.HorizontalResolution
    wscript.echo "VerticalResolution = " & img.VerticalResolution
    wscript.echo "FrameCount = " & img.FrameCount
    wscript.echo "------------------------------------------------------------"
    if img.IsIndexedPixelFormat then
        wscript.echo "Pixel data contains palette indexes"
    end if

    if img.IsAlphaPixelFormat then
        wscript.echo "Pixel data has alpha information"
    end if
    
    if img.IsExtendedPixelFormat then
        wscript.echo "Pixel data has extended color information (16 bit/channel)"
    end if

    if img.IsAnimated then
        wscript.echo "Image is animated"
    end if

    if img.Properties.Exists("40091") then
        wscript.echo "Title = " & img.Properties("40091").Value.String
    end if

    if img.Properties.Exists("40092") then
        wscript.echo "Comment = " & img.Properties("40092").Value.String
    end if

    if img.Properties.Exists("40093") then
        wscript.echo "Author = " & img.Properties("40093").Value.String
    end if

    if img.Properties.Exists("40094") then
        wscript.echo "Keywords = " & img.Properties("40094").Value.String
    end if

    if img.Properties.Exists("40095") then
        wscript.echo "Subject = " & img.Properties("40095").Value.String
    end if
end sub



Dim args, i

set args= WScript.Arguments
if args.Count > 0 then
    for i = 0 to args.Count - 1
        dumpinfo args(i)
    next
    wscript.echo "============================================================"
else
    wscript.echo "Please specify one or more image files"
end if  

