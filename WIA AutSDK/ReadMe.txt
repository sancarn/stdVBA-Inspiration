
Getting Started
---------------

	To install the Windows Image Acquisition Library v2.0, 
	copy the contents of this compressed file to a directory on your hard drive.

	Copy the wiaaut.chm and wiaaut.chi files to your Help directory (usually located at C:\Windows\Help)

	Copy the wiaaut.dll file to your System32 directory (usually located at C:\Windows\System32)

	From a Command Prompt in the System32 directory run the following command: 
	
		RegSvr32 WIAAut.dll 

Errata
------

When using the VideoPreview control, some webcam drivers have a bug that will cause the ExecuteCommand call in the sample below to hang while the VideoPreview is paused.

    Dim Itm 'As Item
    
    VideoPreview1.Pause = True
    Set Itm = VideoPreview1.Device.ExecuteCommand(wiaCommandTakePicture)
	

The Windows Image Acquisition Library v2.0 is only designed to support the PNG, BMP, JPG, GIF and TIFF image formats.  It should not be relied upon to support other formats, though they may appear to be supported depending on system configuration.