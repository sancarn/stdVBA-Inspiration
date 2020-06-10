VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "c32bppDIB (Best Compiled)"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAngle 
      Height          =   315
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Use mousewheel or up/down arrows to cycle thru"
      Top             =   855
      Width           =   2130
   End
   Begin VB.TextBox txtOpacity 
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "100"
      Text            =   "100"
      ToolTipText     =   "Valid values are 0 to 100"
      Top             =   540
      Width           =   630
   End
   Begin VB.OptionButton optSource 
      Caption         =   "From Resource"
      Height          =   240
      Index           =   1
      Left            =   2505
      TabIndex        =   2
      ToolTipText     =   "Load from Resource example"
      Top             =   585
      Width           =   1440
   End
   Begin VB.OptionButton optSource 
      Caption         =   "From File"
      Height          =   240
      Index           =   0
      Left            =   2505
      TabIndex        =   1
      ToolTipText     =   "Load from file Example"
      Top             =   285
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   300
      List            =   "frmTest.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Sample Images"
      Top             =   195
      Width           =   2130
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE7FB&
      Height          =   3900
      Left            =   195
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Drag & Drop or Copy & Paste here as needed"
      Top             =   1200
      Width           =   3900
   End
   Begin VB.CheckBox chkBiLinear 
      Caption         =   "Quality Sizing"
      Height          =   240
      Left            =   270
      TabIndex        =   4
      ToolTipText     =   "Stretch Quality Option"
      Top             =   915
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Drag && Drop, Copy && Paste too.  Unicode Compatible"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "To Paste: Click on display box and press Ctrl+V"
      Top             =   6060
      Width           =   3840
   End
   Begin VB.Label lblType 
      Caption         =   "Label2"
      Height          =   360
      Left            =   270
      TabIndex        =   10
      ToolTipText     =   "Basic Image Info"
      Top             =   5145
      Width           =   3780
   End
   Begin VB.Label Label1 
      Caption         =   "* pARGB pixel format - ARGB but pre-multiplied RGB"
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   9
      Top             =   5685
      Width           =   3840
   End
   Begin VB.Label Label1 
      Caption         =   "* ARGB pixel format where alpha channel is included"
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   8
      Top             =   5475
      Width           =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Opacity (0 - 100)"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   7
      ToolTipText     =   "Valid values are 0 to 100"
      Top             =   600
      Width           =   1500
   End
   Begin VB.Menu mnuGrayScale 
      Caption         =   "Gray Scale/Shadows"
      Begin VB.Menu mnuGray 
         Caption         =   "NTSC-PAL"
         Index           =   0
      End
      Begin VB.Menu mnuGray 
         Caption         =   "CIRC 702"
         Index           =   1
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Simple Average"
         Index           =   2
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Red Mask"
         Index           =   3
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Green Mask"
         Index           =   4
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Blue Mask"
         Index           =   5
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Red-Green Mask"
         Index           =   6
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Blue-Green Mask"
         Index           =   7
      End
      Begin VB.Menu mnuGray 
         Caption         =   "No Gray Scaling"
         Index           =   8
      End
      Begin VB.Menu mnuGray 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuGray 
         Caption         =   "Shadow"
         Index           =   10
         Begin VB.Menu mnuShadow 
            Caption         =   "Use Shadow"
            Index           =   0
         End
         Begin VB.Menu mnuShadow 
            Caption         =   "Shadow Color"
            Index           =   1
            Begin VB.Menu mnuShadowColor 
               Caption         =   "Black"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuShadowColor 
               Caption         =   "Red"
               Index           =   1
            End
            Begin VB.Menu mnuShadowColor 
               Caption         =   "Green"
               Index           =   2
            End
            Begin VB.Menu mnuShadowColor 
               Caption         =   "Blue"
               Index           =   3
            End
         End
         Begin VB.Menu mnuShadow 
            Caption         =   "Blur Depth"
            Index           =   2
            Begin VB.Menu mnuShadowDepth 
               Caption         =   "Light"
               Index           =   0
            End
            Begin VB.Menu mnuShadowDepth 
               Caption         =   "Medium"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu mnuShadowDepth 
               Caption         =   "Heavy"
               Index           =   2
            End
         End
      End
   End
   Begin VB.Menu mnuOtherOpts 
      Caption         =   "Other Options"
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Don't Use GDI+"
         Index           =   0
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Use GDI+"
         Index           =   1
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Mirroring"
         Index           =   3
         Begin VB.Menu mnuMirror 
            Caption         =   "Mirror Horizontally"
            Index           =   0
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "Mirror Vertically"
            Index           =   1
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "Mirror Both Directions"
            Index           =   2
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "No Mirroring"
            Checked         =   -1  'True
            Index           =   4
         End
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Use Negative Rotation Angles"
         Index           =   5
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Save As"
         Index           =   7
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As PNG (Using GDI+)"
            Index           =   0
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As PNG (Using zLIB)"
            Index           =   1
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Default Filter"
               Index           =   0
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use No Filters (Fastest)"
               Index           =   1
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Left Filter"
               Index           =   2
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Top Filter"
               Index           =   3
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Average Filter"
               Index           =   4
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Paeth Filter"
               Index           =   5
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adaptive Filtering (Slowest)"
               Index           =   6
            End
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As JPG"
            Index           =   2
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As TGA"
            Index           =   3
            Begin VB.Menu mnuTGA 
               Caption         =   "Compressed"
               Index           =   0
            End
            Begin VB.Menu mnuTGA 
               Caption         =   "Uncompressed"
               Index           =   1
            End
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As GIF"
            Index           =   4
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As BMP (Red Bkg)"
            Index           =   5
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As Rendered Example (GDI+ required)"
            Index           =   7
         End
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Make Image Inverse (Invert Colors)"
         Index           =   9
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Make Transparent Example"
         Index           =   10
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Light Adjustment"
         Index           =   12
         Begin VB.Menu mnuLight 
            Caption         =   "No Light Adjustment"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuLight 
            Caption         =   "Lighter by 10%"
            Index           =   1
         End
         Begin VB.Menu mnuLight 
            Caption         =   "Lighter by 50%"
            Index           =   2
         End
         Begin VB.Menu mnuLight 
            Caption         =   "Darker by 10%"
            Index           =   3
         End
         Begin VB.Menu mnuLight 
            Caption         =   "Darker by 50%"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Blend To Color (25% Blend)"
         Index           =   14
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Simple Text Example"
         Index           =   16
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuSubOpts 
         Caption         =   "Tile Examples"
         Index           =   18
         Begin VB.Menu mnuTiles 
            Caption         =   "Basic Tiles"
            Index           =   0
         End
         Begin VB.Menu mnuTiles 
            Caption         =   "Staggered Tiles"
            Index           =   1
         End
         Begin VB.Menu mnuTiles 
            Caption         =   "Multiple Overlapping Tiles"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuScale 
      Caption         =   "Scale"
      Begin VB.Menu mnuScalePop 
         Caption         =   "Scale Down As Needed"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuScalePop 
         Caption         =   "Scale Up/Down To Checkerboard"
         Index           =   1
      End
      Begin VB.Menu mnuScalePop 
         Caption         =   "Reduce by 50%"
         Index           =   2
      End
      Begin VB.Menu mnuScalePop 
         Caption         =   "Enlarge by 50%"
         Index           =   3
      End
      Begin VB.Menu mnuScalePop 
         Caption         =   "Actual Size"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPos 
      Caption         =   "Position"
      Begin VB.Menu mnuPosSub 
         Caption         =   "Centered"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPosSub 
         Caption         =   "Top Left"
         Index           =   1
      End
      Begin VB.Menu mnuPosSub 
         Caption         =   "Top Right"
         Index           =   2
      End
      Begin VB.Menu mnuPosSub 
         Caption         =   "Bottom Left"
         Index           =   3
      End
      Begin VB.Menu mnuPosSub 
         Caption         =   "Bottom Right"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' SAMPLE FORM ONLY, used to expose many of the c32bppDIB class options/capabilities

' Unicode-aware Open/Save Dialog box
' ////////////////////////////////////////////////////////////////
Private Type OPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     Flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (lpString As Any) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Const OFN_DONTADDTORECENT As Long = &H2000000
Private Const OFN_ENABLESIZING As Long = &H800000
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_NOCHANGEDIR As Long = &H8
' ////////////////////////////////////////////////////////////////

Private cImage As c32bppDIB
Private cShadow As c32bppDIB

' Note: If GDI+ is available, it is more efficient for you to
' create the token then pass the token to each class.  Not required,
' but if you don't do this, then the classes will create and destroy
' a token everytime GDI+ is used to render or modify an image.
' Passing the token can result in up to 3x faster processing overall
Private m_GDItoken As Long


Private Sub cboAngle_Click()
    
    If cboType.ListIndex = -1 Then Exit Sub
    
    ' show our message regarding compiled vs uncompiled speed issues
    ' if the message hasn't already been shown
    If cImage.isGDIplusEnabled = False Then
        If chkBiLinear.Tag = "" Then
            If Not (cboAngle.ListIndex = 0 Or cboAngle.ListIndex = cboAngle.ListCount - 1) Then
                Call chkBiLinear_Click
                Exit Sub
            End If
        End If
    End If
    
    RefreshImage
    
End Sub

Private Sub cboType_Click()
    ShowImage
End Sub

Private Sub chkBiLinear_Click()

    If chkBiLinear.Value = 1 Then
        If chkBiLinear.Tag = "" Then
            If cImage.isGDIplusEnabled = False Then
                If Not (cboAngle.ListIndex = 0 Or cboAngle.ListIndex = cboAngle.ListCount - 1) Then
                    chkBiLinear.Tag = "noMsg"
                    On Error Resume Next
                    Debug.Print 1 / 0
                    If Err Then ' uncompiled
                        Err.Clear
                        MsgBox "Non-GDI+ rotation with bilinear interpolation is painfully slow in IDE." & vbCrLf & _
                            "But is acceptable when the routines are compiled", vbInformation + vbOKOnly
                    End If
                End If
            End If
        End If
    End If
    
    cImage.HighQualityInterpolation = chkBiLinear.Value
    RefreshImage

End Sub

Private Sub Form_Load()

    ExtractSampleImages ' extracts up to 11 images from the resource file
    
    Dim i As Integer
    For i = 0 To 360 Step 15
        cboAngle.AddItem "Rotate " & i & " degrees"
    Next
    
    ' create our checkboard pattern
    Picture1.AutoRedraw = True
    Set cImage = New c32bppDIB
    cImage.InitializeDIB ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, vbPixels), ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, vbPixels)
    cImage.CreateCheckerBoard 32, vbWhite, Picture1.BackColor
    cImage.Render Picture1.hDC
    cImage.DestroyDIB
    Picture1.Picture = Picture1.Image
    If cImage.isGDIplusEnabled = True Then
        chkBiLinear.Value = 1           ' when GDI+ available, default stretch quality is BiCubic
        Call mnuSubOpts_Click(1)        ' show GDI+ is being used and create a shared token too
    Else ' when GDI+ not installed, then disable GDI+ related options
        mnuSubOpts(0).Checked = True    ' show we are not using GDI+ on the menu
        mnuSubOpts(0).Enabled = False   ' disable option to use GDI+
        mnuSubOpts(1).Enabled = False   ' option to use/not use GDI+
        mnuSaveAs(0).Enabled = False    ' option to save to PNG using GDI+
        mnuSaveAs(2).Enabled = False    ' option to save to JPG using GDI+
        mnuSaveAs(5).Enabled = False    ' option to save as rendered to PNG using GDI+
    End If
    mnuSaveAs(1).Enabled = cImage.isZlibEnabled
    mnuGray(mnuGray.UBound - 2).Checked = True
    mnuGrayScale.Tag = -1
    
    Show
    cboAngle.ListIndex = 0 ' start with zero degree rotation
    cboType.ListIndex = 8 ' set starting point, arbritrary
    
    Me.KeyPreview = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' when you create a token to be shared, you must
    ' destroy it in the Unload or Terminate event
    ' and also reset gdiToken property for each existing class
    If m_GDItoken Then
        If Not cShadow Is Nothing Then cShadow.gdiToken = 0&
        If Not cImage Is Nothing Then
            cImage.gdiToken = 0&
            cImage.DestroyGDIplusToken m_GDItoken
        End If
    End If
    
End Sub

Private Sub mnuBlend_Click()
    
    Dim Index As Long
    
    If mnuSubOpts(14).Tag = vbNullString Then
        MsgBox "There are some functions that cannot be performed on the fly." & vbCrLf & _
            "This function will permanently change the image. In order to revert back " & vbCrLf & _
            "to the original image, it must be re-loaded.", vbInformation + vbOKOnly
    End If
    
    ' in this example, we will replicate the image 3 times, using different colors
    ' But since the original is permanently changed, we will make a copy of it,
    ' change the copy, then render the copy onto the composite image.
    
    ' This is also a good example on how to render to multiple DIB classes
    
    Dim cpyImage As c32bppDIB, cmpImage As c32bppDIB
    
    ' create our composite image, blank but sized appropriately
    Set cmpImage = New c32bppDIB
    cmpImage.InitializeDIB cImage.Width * 2, cImage.Height * 2
    
    For Index = 0 To 3
        cImage.CopyImageTo cpyImage
        Select Case Index
        Case 0: ' do nothing, original will be in upper left
        Case 1: cpyImage.BlendToColor vbBlue, 25
        Case 2: cpyImage.BlendToColor vbRed, 25
        Case 3: cpyImage.BlendToColor vbGreen, 25
        End Select
        ' When rendering class to class, pass the destination DIB class as one of the parameters >>>>>>
        cpyImage.Render 0, (Index And 1) * cImage.Width, (Index \ 2) * cImage.Height, , , , , , , , , , cmpImage
    Next
    Set cpyImage = Nothing  ' don't need the copy any longer
    ' optional step here:
    cmpImage.ImageType = imgBmpPARGB
    ' make our composite the current image
    Set cImage = cmpImage
    ShowImage True, True    ' display it & show msgbox explaining what you are looking at
    
    If mnuSubOpts(14).Tag = vbNullString Then
        mnuSubOpts(14).Tag = "msg shown"
        MsgBox "Top Left: Original image" & vbCrLf & _
            "Top Right: Blue blend" & vbCrLf & _
            "Bottom Left: Red blend" & vbCrLf & _
            "Bottom Right: Green blend", vbInformation + vbOKOnly, "Blend/Tint Sample"
    End If

    
End Sub

Private Sub mnuGray_Click(Index As Integer)

    If Index < mnuGray.UBound Then
        Dim i As Integer
        For i = mnuGray.LBound To mnuGray.UBound - 2
            If mnuGray(i).Checked = True Then
                mnuGray(i).Checked = False
                Exit For
            End If
        Next
        mnuGray(Index).Checked = True
        If Index = mnuGray.UBound - 2 Then Index = -1
        mnuGrayScale.Tag = Index
        RefreshImage
    End If
    
End Sub

Private Sub mnuLight_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 4
        If mnuLight(i).Checked = True Then
            mnuLight(i).Checked = False
            If Index = i Then Index = 0 ' no mirroring mnu option
            Exit For
        End If
    Next
    mnuLight(Index).Checked = True
    Select Case Index
    Case 0: mnuLight(0).Tag = 0
    Case 1: mnuLight(0).Tag = 10
    Case 2: mnuLight(0).Tag = 50
    Case 3: mnuLight(0).Tag = -10
    Case 4: mnuLight(0).Tag = -50
    End Select
    RefreshImage
    
End Sub

Private Sub mnuMirror_Click(Index As Integer)

    Dim i As Integer
    For i = 0 To 2
        If mnuMirror(i).Checked = True Then
            mnuMirror(i).Checked = False
            If Index = i Then Index = 4 ' no mirroring mnu option
            Exit For
        End If
    Next
    mnuMirror(Index).Checked = True
    If Not Index = 4 Then mnuMirror(4).Checked = False ' mirroring is in use
    RefreshImage

End Sub

Private Sub mnuPosSub_Click(Index As Integer)
    If mnuPosSub(Index).Checked = True Then Exit Sub
    Dim i As Integer
    For i = mnuPosSub.LBound To mnuScalePop.UBound
        If mnuPosSub(i).Checked = True Then
            mnuPosSub(i).Checked = False
            Exit For
        End If
    Next
    mnuPosSub(Index).Checked = True
    RefreshImage
End Sub

Private Sub mnuSaveAs_Click(Index As Integer)

    Dim sFile As String
    Select Case Index
    Case 0: ' save as PNG using GDI+
        sFile = OpenSaveFileDialog(True, "Save As", "png", True)
        If Not sFile = vbNullString Then
            ' to force use of GDI+, we can't have any optional PNG properties
            cImage.PngPropertySet pngProp_ClearProps
            If cImage.SaveToFile_PNG(sFile, False) = True Then
                If MsgBox("PNG successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
                    If cImage.LoadPicture_File(sFile) = False Then
                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
                    Else
                        ShowImage True, True
                    End If
                End If
            Else
                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
            End If
        End If
    Case 1 ' save using zLIB
    
    Case 2 ' save as jpg
        sFile = OpenSaveFileDialog(True, "Save As", "jpg", True)
        If Not sFile = vbNullString Then
            If cImage.SaveToFile_JPG(sFile, , vbWhite, False) = True Then
                If MsgBox("JPG successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
                    If cImage.LoadPicture_File(sFile) = False Then
                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
                    Else
                        ShowImage True, True
                    End If
                End If
            Else
                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
            End If
        End If
    Case 3 ' save as tga options
    
    Case 4 ' save as GIF
        sFile = OpenSaveFileDialog(True, "Save As", "gif", True)
        If Not sFile = vbNullString Then
            If cImage.SaveToFile_GIF(sFile, True, 200, False) = True Then
                If MsgBox("GIF successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
                    If cImage.LoadPicture_File(sFile) = False Then
                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
                    Else
                        ShowImage True, True
                    End If
                End If
            Else
                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
            End If
        End If
    Case 5 ' save as BMP using red solid bkg (only applies to images with transparency)
        sFile = OpenSaveFileDialog(True, "Save As", "bmp", True)
        If Not sFile = vbNullString Then
            If cImage.SaveToFile_Bitmap(sFile, , vbRed, False) = True Then
                If MsgBox("BMP successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
                    If cImage.LoadPicture_File(sFile) = False Then
                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
                    Else
                        ShowImage True, True
                    End If
                End If
            Else
                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
            End If
        End If
    Case 7 ' save as rendered example
    
        ' There are at least 2 ways to save an image as rendered
        ' 1. Modify the image directly by calling routines that permanently change the image
        '   such as MakeLighterDarker, MakeGrayScale, etc. Then simply call the appropriate Save function
        ' 2. This way, render the image to another DIB class and save that one
        
        ' This is actually much simpler than it looks, but we will try to accomodate all the
        ' user-selected options on the form & render this to another DIB then save that DIB
        ' The only additional steps would be when the image is rotated. If that is the case,
        ' then you would need to resize the new DIB appropriately. Example does this...
        
        ' The other important thing to remember when rendering DIB to DIB is to pass the
        ' optional destDibHost parameter in the render function. Example does this...
        
        ' Note: This sample does not honor the shadow if it is applied.  If you want to also
        ' include a shadow, simply adjust your image size to accomodate the shadow and then
        ' render it first, then render the image over the shadow.
        
        Dim rImage As c32bppDIB
        Dim newWidth As Long, newHeight As Long
        Dim mirrorOffsetX As Long, mirrorOffsetY As Long
        Dim negAngleOffset As Long
        Dim LightAdjustment As Single
        
        sFile = OpenSaveFileDialog(True, "Save As", "png", True)
        If Not sFile = vbNullString Then
        
            mirrorOffsetX = 1
            mirrorOffsetY = 1
            Select Case True
                Case mnuMirror(0).Checked   ' horizontal mirroring
                    mirrorOffsetX = -1
                Case mnuMirror(1).Checked   ' vertical mirroring
                    mirrorOffsetY = -1
                Case mnuMirror(2).Checked   ' both directions mirrored
                    mirrorOffsetX = -1
                    mirrorOffsetY = -1
            End Select
            If mnuSubOpts(5).Checked = True Then negAngleOffset = -1 Else negAngleOffset = 1
            LightAdjustment = CSng(Val(mnuLight(0).Tag))
            Select Case True    ' scaling options from menu
                Case mnuScalePop(0).Checked ' only scale down as needed
                    cImage.ScaleImage Picture1.ScaleWidth, Picture1.ScaleHeight, newWidth, newHeight, scaleDownAsNeeded
                Case mnuScalePop(1).Checked ' scale up and/or down
                    cImage.ScaleImage Picture1.ScaleWidth, Picture1.ScaleHeight, newWidth, newHeight, ScaleToSize
                Case mnuScalePop(2).Checked ' reduce by 1/2
                    cImage.ScaleImage cImage.Width \ 2, cImage.Height \ 2, newWidth, newHeight, ScaleToSize
                Case mnuScalePop(3).Checked ' enlarge by 1/2
                    cImage.ScaleImage cImage.Width * 1.5, cImage.Height * 1.5, newWidth, newHeight, ScaleToSize
                Case mnuScalePop(4).Checked ' actual size
                    newWidth = cImage.Width: newHeight = cImage.Height
            End Select
            
            ' the cboAngle entries are at 15 degree intervals, so we simply multiply ListIndex by 15
            If (cboAngle.ListIndex * 15) Mod 360 Then ' rotated
                ' rotation: size the dib to the maximum size needed to handle all rotation angles
                newWidth = Sqr(newWidth * newWidth + newHeight * newHeight)
                newHeight = newWidth
            End If
            
            ' create a new DIB & size it
            Set rImage = New c32bppDIB
            rImage.InitializeDIB newWidth, newHeight
            
            ' rendering to the center (last parameter) as shown below is optional but if rendering
            ' rotated then always render to the center of the target area
            
            ' To correctly render DIB to DIB, always pass the target DIB as the optional destDibHost parameter
            ' When rendering DIB to DIB, the hDC is ignored and that is why we pass zero.
            cImage.Render 0, newWidth \ 2, newHeight \ 2, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
                Val(txtOpacity.Text), , , rImage, Val(mnuGrayScale.Tag), LightAdjustment, (cboAngle.ListIndex * 15) * negAngleOffset, True
        
            rImage.TrimImage True, trimAll
            ' ^^ if the image was rotated, you should call this to remove any transparent "borders"
        
            If rImage.SaveToFile_PNG(sFile, False) Then
                MsgBox "PNG successfully created", vbInformation + vbOKOnly, "Success"
            Else
                MsgBox "PNG failed to be created", vbExclamation + vbOKOnly, "Failure"
            End If
    
        End If
    
    End Select

End Sub

Private Sub mnuScalePop_Click(Index As Integer)

    If mnuScalePop(Index).Checked = True Then Exit Sub
    Dim i As Integer
    For i = mnuScalePop.LBound To mnuScalePop.UBound
        If mnuScalePop(i).Checked = True Then
            mnuScalePop(i).Checked = False
            Exit For
        End If
    Next
    mnuScalePop(Index).Checked = True
    RefreshImage
End Sub

Private Sub mnuShadow_Click(Index As Integer)
    If Index = 0 Then
        If mnuShadow(Index).Checked = True Then
            Set cShadow = Nothing
            RefreshImage
        Else
            CreateNewShadowClass
        End If
        mnuShadow(Index).Checked = Not (cShadow Is Nothing)
    End If
End Sub

Private Sub mnuShadowColor_Click(Index As Integer)
    Dim i As Integer
    For i = mnuShadowColor.LBound To mnuShadowColor.UBound
        If mnuShadowColor(i).Checked = True Then
            If Index = i Then Exit Sub
            mnuShadowColor(i).Checked = False
            Exit For
        End If
    Next
    mnuShadowColor(Index).Checked = True
    mnuShadow(0).Checked = True
    CreateNewShadowClass
End Sub

Private Sub mnuShadowDepth_Click(Index As Integer)
    Dim i As Integer
    For i = mnuShadowDepth.LBound To mnuShadowDepth.UBound
        If mnuShadowDepth(i).Checked = True Then
            If Index = i Then Exit Sub
            mnuShadowDepth(i).Checked = False
            Exit For
        End If
    Next
    mnuShadow(0).Checked = True
    mnuShadowDepth(Index).Checked = True
    CreateNewShadowClass
End Sub

Private Sub mnuSubOpts_Click(Index As Integer)

    ' The 1st two options will be disabled if you do not have GDI+ installed
    
    Select Case Index
    Case 0: ' do not use GDI+
        If mnuSubOpts(Index).Checked = True Then Exit Sub
        cImage.isGDIplusEnabled = False
        mnuSubOpts(0).Checked = Not mnuSubOpts(0).Checked
        mnuSubOpts(1).Checked = False
        
        If m_GDItoken Then  ' when using token, we'll clean up here
            cImage.DestroyGDIplusToken m_GDItoken
            m_GDItoken = 0&
            cImage.gdiToken = m_GDItoken ' reset the token
            If Not cShadow Is Nothing Then cShadow.gdiToken = m_GDItoken
        End If
        
        RefreshImage
    
    Case 1: ' always usge GDI+.
        If mnuSubOpts(Index).Checked = True Then Exit Sub
        mnuSubOpts(0).Checked = False ' remove checkmark on "Don't Use GDI+"
        mnuSubOpts(1).Checked = True  ' show using GDI+
        cImage.isGDIplusEnabled = True
        ' verify it enabled correct and get a token to share
        If cImage.isGDIplusEnabled Then
            m_GDItoken = cImage.CreateGDIplusToken()
            cImage.gdiToken = m_GDItoken
            If Not cShadow Is Nothing Then cShadow.gdiToken = m_GDItoken
        End If
        ' tell GDI+ that we want high quality interpolation
        If chkBiLinear.Value = 0 Then chkBiLinear.Value = 1 Else RefreshImage
        
    Case 3: ' mirroring sub menus
    
    Case 5: ' negative vs positive angles
        mnuSubOpts(Index).Checked = Not mnuSubOpts(Index).Checked
        RefreshImage
        
    Case 7: ' save as
        
    Case 9 ' make inverted colors
        cImage.MakeImageInverse
        RefreshImage
        
    Case 10 ' make transparent example
    
        cImage.LoadPicture_Resource "MTPsample", vbResBitmap, VB.Global
        If mnuGrayScale.Tag <> -1 Then
            Call mnuGray_Click(mnuGray.UBound)
        Else
            ShowImage True, True
        End If
        MsgBox "The Fox has a yellow background." & vbNewLine & " Clicking Ok will tell the c32bppDIB class " & _
            "to use the top left pixel and make that color transparent." & vbCrLf & vbCrLf & _
            "You may need to move this message window to see the fox image.", vbInformation + vbOKOnly, "Make Transparent Example"
            
        ' note: here I am using the class' GetPixel function, but since I know the color I want
        ' to make transparent, I could have just easily used: cImage.MakeTransparent vbYellow
        cImage.MakeTransparent cImage.GetPixel(0, 0)
        If Not cShadow Is Nothing Then
            CreateNewShadowClass
        Else
            ShowImage True, True
        End If
        
    Case 12 ' adding/subtracting lightness from image menus
    Case 14 ' blending color into image menus
        Call mnuBlend_Click
    Case 16 ' text overlay example
        ' show a larger sample image
        If cboType.ListIndex = 8 Then Call ShowImage(True) Else cboType.ListIndex = 8
        
        Me.FontName = "Times New Roman"
        Me.Font.Size = 16
        Me.FontBold = True
        Me.Font.Weight = 400
        
        ' Example on how to create a separate text DIB and place it where you want
        Dim textDIB As c32bppDIB, glowShadow As c32bppDIB
        Set textDIB = New c32bppDIB
        textDIB.DrawText_stdFont Me.Font, "TEXT OVERLAY" & vbNewLine & "Example #1", TA_CENTER, , , , , vbBlue, RGB(192, 192, 192)
        textDIB.Render cImage.LoadDIBinDC(True), (cImage.Width - textDIB.Width) / 2, 6, , , , , , , 55
        cImage.LoadDIBinDC False
        Set textDIB = Nothing

        ' Example of rendering text directly on the same DIB
        Me.FontName = "Tahoma"
        Me.Font.Weight = 800
        Me.Font.Size = 16
        cImage.DrawText_stdFont Me.Font, "TEXT OVERLAY" & vbNewLine & "Example #3", TA_CENTER, 39, cImage.Height \ 2 + 1, , , vbBlack, , , , -90, True
        cImage.DrawText_stdFont Me.Font, "TEXT OVERLAY" & vbNewLine & "Example #3", TA_CENTER, 40, cImage.Height \ 2, , , vbCyan, , , , -90, True
        
        ' Example of using multiple DIBs and overlaying to produce different effects
        Me.FontName = "Comic Sans MS"
        Me.Font.Weight = 800
        Me.Font.Size = 14
        Set textDIB = New c32bppDIB
        ' create separate text dib
        textDIB.DrawText_stdFont Me.Font, "TEXT OVERLAY" & vbNewLine & "Example #2", TA_CENTER, , , , , vbBlue
        ' create a blurred shadow
        Set glowShadow = textDIB.CreateDropShadow(10, vbBlue)
        ' render the 2 over the main image
        glowShadow.Render cImage.LoadDIBinDC(True), cImage.Width * 0.75, cImage.Height * 0.75, , , , , , , 80, , , , , , -45, True
        textDIB.Render cImage.LoadDIBinDC(True), cImage.Width * 0.75, cImage.Height * 0.75, , , , , , , 90, , , , , , -45, True
        cImage.LoadDIBinDC False
        Set textDIB = Nothing
        Set glowShadow = Nothing
        
        RefreshImage
        
    Case 18 ' tiling examples
        
    End Select
ExitRoutine:
End Sub

Private Sub mnuTGA_Click(Index As Integer)
    
    Dim sFile As String
    sFile = OpenSaveFileDialog(True, "Save As", "tga", True)
    If Not sFile = vbNullString Then
        If cImage.SaveToFile_TGA(sFile, (Index = 0), False, True, False) = True Then
            If MsgBox("TGA successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
                If cImage.LoadPicture_File(sFile) = False Then
                    MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
                Else
                    ShowImage True, True
                End If
            End If
        Else
            MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
        End If
    End If
    
End Sub

Private Sub mnuTiles_Click(Index As Integer)

    If mnuTiles(0).Tag = vbNullString Then
        mnuTiles(0).Tag = "nomsg"
        MsgBox "The tiling examples are only temporary. Selecting any other menu options or " & vbNewLine & _
            "clicking on any other form controls will erase the sample.", vbInformation + vbOKOnly, "FYI"
    End If
    
    ' FYI: the tiling algorithm I use can tile a full screen in only a few milliseconds
    
    Dim tImage As c32bppDIB
    Set tImage = New c32bppDIB
    
    tImage.LoadPicture_Resource "TILESAMPLE", vbResBitmap, VB.Global
    
    Select Case Index
    Case 0 ' simple tile
        tImage.TileImage Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Case 1 ' staggered tile
        tImage.TileImage Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, , , , , True
    Case 2 ' overlapping tiles
    
        ' having a little fun ....
        Dim dcWidth As Long, dcHeight As Long
        Dim oImage As c32bppDIB
        Set oImage = New c32bppDIB
        ' extract my fox head bitmap & make it transparent. I know it isn't by default
        oImage.LoadPicture_Resource "MTPSAMPLE", vbResBitmap, VB.Global
        oImage.MakeTransparent oImage.GetPixel(0, 0)
    
        ' calculate width/height want to tile, adding gapping into the calcs
        dcWidth = (oImage.Width \ 2) * 4 + 60 '(3x 20 pixel gap)
        dcHeight = (oImage.Height \ 2) * 3 + 60 ' (2x 30 pixel gap)
        
        ' tile the background
        tImage.TileImage Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        
        ' tile the fox head over it, each fox head is 1/2 width/height
        oImage.TileImage Picture1.hDC, (Picture1.ScaleWidth - dcWidth) \ 2, (Picture1.ScaleHeight - dcHeight) \ 2, _
            dcWidth, dcHeight, oImage.Width \ 2, oImage.Height \ 2, 20, 30, , , True, 80
    End Select
    
    Picture1.Refresh
    
End Sub

Private Sub mnuZlibPng_Click(Index As Integer)
    
    Dim sFile As String
    ' by setting optional parameters, class will use zLIB over GDI+
    ' to the contrary, if no parameters are set, class uses GDI+ over zLIB
    Select Case Index
    Case 0: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterDefault
    Case 1: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterNone
    Case 2: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdjLeft
    Case 3: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdjTop
    Case 4: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdjAvg
    Case 5: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterPaeth
    Case 6: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdaptive
    End Select
    
    sFile = OpenSaveFileDialog(True, "Save As", "png", True)
    If Not sFile = vbNullString Then
        If cImage.SaveToFile_PNG(sFile, False) = True Then
            If MsgBox("PNG successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
                If cImage.LoadPicture_File(sFile) = False Then
                    MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
                Else
                    ShowImage True, True
                End If
            End If
        Else
            MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
        End If
    End If
    
ExitRoutine:
End Sub

Private Sub optSource_Click(Index As Integer)
    ShowImage
    cboType.SetFocus
End Sub

Private Sub RefreshImage()

    Dim newWidth As Long, newHeight As Long
    Dim mirrorOffsetX As Long, mirrorOffsetY As Long
    Dim negAngleOffset As Long
    Dim X As Long, Y As Long
    Dim ShadowOffset As Long
    Dim LightAdjustment As Single
    
    ' This one routine handles all the options of the sample form
    
    ShadowOffset = Val(mnuShadowDepth(0).Tag) + 2   ' set shadow's blur depth as needed
    
    mirrorOffsetX = 1
    mirrorOffsetY = 1
    Select Case True
        Case mnuMirror(0).Checked   ' horizontal mirroring
            mirrorOffsetX = -1
        Case mnuMirror(1).Checked   ' vertical mirroring
            mirrorOffsetY = -1
        Case mnuMirror(2).Checked   ' both directions mirrored
            mirrorOffsetX = -1
            mirrorOffsetY = -1
    End Select
    
    If mnuSubOpts(5).Checked = True Then negAngleOffset = -1 Else negAngleOffset = 1
    LightAdjustment = CSng(Val(mnuLight(0).Tag))
    
    Select Case True    ' scaling options from menu
        Case mnuScalePop(0).Checked ' only scale down as needed
            cImage.ScaleImage Picture1.ScaleWidth, Picture1.ScaleHeight, newWidth, newHeight, scaleDownAsNeeded
        Case mnuScalePop(1).Checked ' scale up and/or down
            cImage.ScaleImage Picture1.ScaleWidth, Picture1.ScaleHeight, newWidth, newHeight, ScaleToSize
        Case mnuScalePop(2).Checked ' reduce by 1/2
            cImage.ScaleImage cImage.Width \ 2, cImage.Height \ 2, newWidth, newHeight, ScaleToSize
        Case mnuScalePop(3).Checked ' enlarge by 1/2
            cImage.ScaleImage cImage.Width * 1.5, cImage.Height * 1.5, newWidth, newHeight, ScaleToSize
        Case mnuScalePop(4).Checked ' actual size
            newWidth = cImage.Width: newHeight = cImage.Height
    End Select
    
    ' in this sample form, to make it easier to calculate rendering X,Y coordinates,
    ' we will always pass the X,Y of where the center of the image should appear.
    ' This way, whether rotating or not, we can use the same Render call without
    ' modifying the destination X,Y and CenterOnDestXY paramters based on rotating or not
    Select Case True
        Case mnuPosSub(0).Checked   ' centered on canvas
            X = (Picture1.ScaleWidth - newWidth) \ 2
            Y = (Picture1.ScaleHeight - newHeight) \ 2
        Case mnuPosSub(2).Checked   ' top right
            X = Picture1.ScaleWidth - newWidth
        Case mnuPosSub(3).Checked   ' bottom left
            Y = Picture1.ScaleHeight - newHeight
        Case mnuPosSub(4).Checked   ' bottom right
            X = Picture1.ScaleWidth - newWidth
            Y = Picture1.ScaleHeight - newHeight
        Case mnuPosSub(1).Checked   ' top left
    End Select
    
    Picture1.Cls
    If Not cShadow Is Nothing Then
        Picture1.CurrentX = 20
        Picture1.CurrentY = 5
        Picture1.Print "See c32bppDIB.CreateDropShadow for more ": Picture1.CurrentX = 20
        Picture1.Print "Color, Opacity, Blur Effect, ": Picture1.CurrentX = 20
        Picture1.Print "  and X,Y Position are adjustable"
    End If
    
    
    ' Generally, when rotating and/or resizing, it is easier to calculate the center of where you want the image rotated vs
    '   calculating the top/left coordinate of the resized, rotated image.  The last parameter of the Render call (CenterOnDestXY)
    '   will render around that center point if that paremeter is set.  So, what about when an image is not rotated? The Render
    '   function will still draw around that center point if the parameter is true. Or render, starting at the passed
    '   DestX,DestY coordinates if that parameter is false.
    
    ' The Render call only has one required parameter.  All others are optional and defaulted as follows
        ' srcX, srcY, destX, destY defaults are zero
        ' srcWidth, destWidth defaults are the image's width
        ' srcHeight, destHeight defaults are the image's height
        ' Opacity (Global Alpha) default is 100% opaque, pixel LigthAdjustmnet default is zero (no additional adjustment)
        ' GrayScale default is not grayscaled
        ' Rotation angle is at zero degrees
        ' Rendering image around a center point is false
    
    ' the cboAngle entries are at 15 degree intervals, so we simply multiply ListIndex by 15
    
    If Not cShadow Is Nothing Then
        ' the 55 below is the shadow's opacity; hardcoded here but can be modified to your heart's delight
        cShadow.Render Picture1.hDC, X + newWidth \ 2 + ShadowOffset, Y + newHeight \ 2 + ShadowOffset, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
            55, , , , , LightAdjustment, (cboAngle.ListIndex * 15) * negAngleOffset, True
    End If
    
    cImage.Render Picture1.hDC, X + newWidth \ 2, Y + newHeight \ 2, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
        Val(txtOpacity.Text), , , , Val(mnuGrayScale.Tag), LightAdjustment, (cboAngle.ListIndex * 15) * negAngleOffset, True
    
    Picture1.Refresh

End Sub

Private Sub ShowImage(Optional bRefresh As Boolean = True, Optional DragDropCutPast As Boolean)

    Dim sSource As Variant
    Dim cX As Long, cY As Long

    If Not DragDropCutPast Then
        If optSource(0).Enabled = True Then
            If optSource(0) = True Then ' from file
            
                Select Case cboType.ListIndex
                Case 0: sSource = "Forest.bmp"
                Case 1: sSource = "Alpha-ARGB.bmp"
                Case 2: sSource = "Alpha-pARGB.bmp"
                Case 3: sSource = "Knight.gif"
                Case 4: sSource = "Desktop.ico"
                Case 5: sSource = "XP-Alpha.ico"
                Case 6: sSource = "Vista-PNG.ico"
                Case 7: sSource = "Risk.jpg"
                Case 8: sSource = "Spider.png"
                Case 9: sSource = "Lion.wmf"
                Case 10: sSource = "Hand.cur"
                Case 11: sSource = "GlobalSearch.tga"
                Case 12: sSource = OpenSaveFileDialog(False, "Select Image")
                End Select
                If cboType.ListIndex < cboType.ListCount - 1 Then
                    If Right$(App.Path, 1) = "\" Then
                        sSource = App.Path & sSource
                    Else
                        sSource = App.Path & "\" & sSource
                    End If
                End If
                cImage.LoadPicture_File sSource, 256, 256, (cboType.ListIndex = 12)
                ' ^^ the end parameters: 256,256 is just telling the class that
                ' we want that size icon if one exists in the passed resource. If not,
                ' then give us the one closest to it & best quality too.
                ' The final parameter is telling the class to cache the image bytes once it is loaded.
                ' I will use those bytes to reload the image as needed vs having the user re-select
                ' the image from the browser. Look in this sample form for
                ' cimage.LoadPicture_FromOrignalFormat to see how those bytes are used
                
            Else    ' from resource
                Select Case cboType.ListIndex
                Case 0 ' bitmap
                    sSource = vbResBitmap
                Case 4 'icon
                    sSource = vbResIcon
                Case 10 ' cursor
                    sSource = vbResCursor
                Case 12 ' browse for file, n/a
                    optSource(0) = True ' change source option & browser pop up
                    Exit Sub
                Case Else ' pARGB bmp, ARGB bmp, GIF, alpha icon, png icon, jpg, png, wmf, tga
                    sSource = "Custom"
                End Select
                cImage.LoadPicture_Resource (cboType.ListIndex + 101) & "LaVolpe", sSource, VB.Global, 256, 256, , , 32
                ' ^^ the last two parameters: 256,256 is just telling the class that
                ' we want that size icon if one exists in the passed resource. If not,
                ' then give us the one closest to it & best quality too.
            End If
        End If
    End If

    Select Case cImage.ImageType ' want to know source image's format?
        Case imgNone, imgError:     lblType.Caption = "Image was not loaded"
        Case imgBitmap:             lblType.Caption = "Format: Standard Bitmap or JPG"
        Case imgEMF:                lblType.Caption = "Format: Extended Windows Metafile"
        Case imgWMF:                lblType.Caption = "Format: Standard Windows Metafile"
        Case imgIcon:               lblType.Caption = "Format: Standard Icon"
        Case imgBmpARGB:            lblType.Caption = "Format: 32bpp Bitmap with ARGB"
        Case imgBmpPARGB:           lblType.Caption = "Format: 32bpp Bitmap with pARGB"
        Case imgCursor:             lblType.Caption = "Format: Standard Cursor"
        Case imgCursorARGB:         lblType.Caption = "Format: Alpha Cursor"
        Case imgIconARGB:           lblType.Caption = "Format: Alpha Icon"
        Case imgPNG:                lblType.Caption = "Format: PNG"
        Case imgPNGicon:            lblType.Caption = "Format: PNG in Vista Icon"
        Case imgGIF
            If cImage.Alpha > AlphaNone Then
                                    lblType.Caption = "Format: Transparent GIF"
            Else
                                    lblType.Caption = "Format: GIF"
            End If
        Case imgTGA
            If cImage.Alpha > AlphaNone Then
                                    lblType.Caption = "Format: Transparent TGA"
            Else
                                    lblType.Caption = "Format: TGA (Targa)"
            End If
        Case Else:                  lblType.Caption = "..."
    End Select
    
    If cImage.ImageType > imgNone Then
        lblType.Caption = lblType.Caption & " {" & cImage.Width & " x " & cImage.Height & "}"
    End If
    
    If Not cShadow Is Nothing Then
        CreateNewShadowClass
    Else
        If bRefresh Then RefreshImage
    End If
    
    If Me.Tag = "" Then
        If optSource(1) = True And cboType.ListIndex = 10 Then
            On Error Resume Next    ' only show this message in IDE
            Debug.Print 1 / 0
            If Err Then
                MsgBox "Notice this is black and white." & vbCrLf & _
                    "VB, while in IDE, forces 2 color cursors to be black & white, even though they may not be." & vbCrLf & _
                    "When the cursor is loaded from a resource file when the project is compiled, the cursor magically shows its colors", vbInformation + vbOKOnly
            End If
            Me.Tag = "Message Shown"    ' only show message once
        End If
    End If
    
End Sub

Private Sub ExtractSampleImages()

    Dim sPath As String
    Dim sFile As String
    Dim sResSection As Variant
    Dim X As Long, fnr As Integer
    Dim imgArray() As Byte, tPic As StdPicture
    
    On Error GoTo eh
    sPath = App.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    For X = 1 To 12
        Select Case X
        Case 1
            sFile = "Forest.bmp"     ' standard bitmap (never has transparency)
            sResSection = vbResBitmap
        Case 2
            sFile = "Alpha-ARGB.bmp" ' 32bpp without pre-multiplied RGB values
            sResSection = "Custom"
        Case 3
            sFile = "Alpha-pARGB.bmp" ' 32bpp with premultiplied RGB values
            sResSection = "Custom"
        Case 4
            sFile = "Knight.gif"     ' transaprent GIF frame
            sResSection = "Custom"
        Case 5
            sFile = "Desktop.ico"    ' standard icon
            sResSection = vbResIcon
        Case 6
            sFile = "XP-Alpha.ico"   ' 32bpp alpha blended icon
            sResSection = "Custom"
        Case 7
            sFile = "Vista-PNG.ico"  ' PNG file encoded into icon file
            sResSection = "Custom"
        Case 8
            sFile = "Risk.jpg"       ' standard jpg (never has transparency)
            sResSection = "Custom"
        Case 9
            sFile = "Spider.png"    ' a typical PNG file
            sResSection = "Custom"
        Case 10
            sFile = "Lion.wmf"      ' windows meta file (may have transparency)
            sResSection = "Custom"
        Case 11
            sFile = "Hand.cur"      ' colored cursor
            sResSection = "Custom"
        Case 12
            sFile = "GlobalSearch.tga" ' alpha blended TGA
            sResSection = "Custom"
        End Select
        
        sFile = sPath & sFile
        If Len(Dir(sFile, vbArchive Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then
           Select Case sResSection
           Case vbResBitmap, vbResIcon, vbResCursor
                Set tPic = LoadResPicture((X + 100) & "LaVolpe", sResSection)
                SavePicture tPic, sFile
            Case "Custom"
                imgArray = LoadResData((X + 100) & "LaVolpe", sResSection)
                fnr = FreeFile()
                Open sFile For Binary As #fnr
                Put #fnr, , imgArray()
                Close #fnr
            End Select
        End If
    Next
eh:
    If Err Then
        MsgBox Err.Description, vbInformation + vbOKOnly, "Oops...."
        Err.Clear
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Simple example of pasting file names
    If KeyCode = vbKeyV Then
        
        If (Shift And vbCtrlMask) = vbCtrlMask Then
            ' use class to load 1st file that was pasted, if any & if more than one
            ' Unicode filenames supported
            If cImage.LoadPicture_PastedFiles(1, 256, 256) = False Then
                ' couldn't load anything from the files, maybe image itself was pasted
                If cImage.LoadPicture_ClipBoard = False Then
                    MsgBox "Failed to load whatever was placed in the clipboard", vbInformation + vbOKOnly
                    Exit Sub
                End If
            End If
            
            If Not cShadow Is Nothing Then
                CreateNewShadowClass
            Else
                RefreshImage
            End If
            ShowImage False, True
        
        End If
    End If
End Sub

Private Sub Picture1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' simmple OLE drag/drop example
    
    ' use class to load 1st file that was dropped, if more than one. Unicode compatible
    If cImage.LoadPicture_DropedFiles(data, 1, 256, 256) Then
        If Not cShadow Is Nothing Then
            CreateNewShadowClass
        Else
            RefreshImage
        End If
        ShowImage False, True
    End If

End Sub

Private Sub txtOpacity_Validate(Cancel As Boolean)
    If txtOpacity.Tag <> txtOpacity.Text Then
        txtOpacity.Tag = txtOpacity.Text
        RefreshImage
    End If
End Sub

Private Function OpenSaveFileDialog(bSave As Boolean, DialogTitle As String, Optional DefaultExt As String, Optional SingleFilter As Boolean) As String

    ' using API version vs commondialog enables Unicode filenames to be passed to c32bppDIB classes
    Dim ofn As OPENFILENAME
    Dim rtn As Long
    Dim bUnicode As Boolean
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd
        .hInstance = App.hInstance
        If SingleFilter Then
            Select Case DefaultExt
            Case "png"
                .lpstrFilter = "PNG" & vbNullChar & "*.png" & vbNullChar
            Case "jpg"
                .lpstrFilter = "JPG" & vbNullChar & "*.jpg" & vbNullChar
            Case "tga"
                .lpstrFilter = "TGA (Targa)" & vbNullChar & "*.tga" & vbNullChar
            Case "gif"
                .lpstrFilter = "GIF" & vbNullChar & "*.gif" & vbNullChar
            Case "bmp"
                .lpstrFilter = "Bitmap" & vbNullChar & "*.bmp" & vbNullChar
            End Select
        Else
            .lpstrFilter = "Image Files" & vbNullChar & "*gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.wmf;*.emf;*.png;*.tga"
            If cImage.isGDIplusEnabled Then
                .lpstrFilter = .lpstrFilter & ";*.tiff"
            End If
            .lpstrFilter = .lpstrFilter & vbNullChar & "Bitmaps" & vbNullChar & "*.bmp" & vbNullChar & "GIFs" & vbNullChar & "*.gif" & vbNullChar & _
                            "Icons/Cursors" & vbNullChar & "*.ico;*.cur" & vbNullChar & "JPGs" & vbNullChar & "*.jpg;*.jpeg" & vbNullChar & _
                            "Meta Files" & vbNullChar & "*.wmf;*.emf" & vbNullChar & "PNGs" & vbNullChar & "*.png" & vbNullChar & "TGAs (Targa)" & vbNullChar & "*.tga" & vbNullChar
            If cImage.isGDIplusEnabled Then
                .lpstrFilter = .lpstrFilter & "TIFFs" & vbNullChar & "*.tiff" & vbNullChar
            End If
            .lpstrFilter = .lpstrFilter & "All Files" & vbNullChar & "*.*" & vbNullChar
        End If
        .lpstrDefExt = DefaultExt
        .lpstrFile = String$(256, 0)
        .nMaxFile = 256
        .nMaxFileTitle = 256
        .lpstrTitle = DialogTitle
        .Flags = OFN_LONGNAMES Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_DONTADDTORECENT _
                Or OFN_NOCHANGEDIR
        ' ^^ don't want to change paths otherwise VB IDE locks folder until IDE is closed
        If bSave Then
            .Flags = .Flags Or OFN_CREATEPROMPT Or OFN_OVERWRITEPROMPT
        Else
            .Flags = .Flags Or OFN_FILEMUSTEXIST
        End If
    
        bUnicode = Not (IsWindowUnicode(GetDesktopWindow) = 0&)
        If bUnicode Then
            .lpstrInitialDir = StrConv(.lpstrInitialDir, vbUnicode)
            .lpstrFile = StrConv(.lpstrFile, vbUnicode)
            .lpstrFilter = StrConv(.lpstrFilter, vbUnicode)
            .lpstrTitle = StrConv(.lpstrTitle, vbUnicode)
            .lpstrDefExt = StrConv(.lpstrDefExt, vbUnicode)
        End If
        .lpstrFileTitle = .lpstrFile
    End With
    
    If bUnicode Then
        If bSave Then
            rtn = GetSaveFileNameW(ofn)
        Else
            rtn = GetOpenFileNameW(ofn)
        End If
        If rtn > 0& Then
            If bUnicode Then
                rtn = lstrlenW(ByVal ofn.lpstrFile)
                OpenSaveFileDialog = StrConv(Left$(ofn.lpstrFile, rtn * 2), vbFromUnicode)
            End If
        End If
    Else
        If bSave Then
            rtn = GetSaveFileName(ofn)
        Else
            rtn = GetOpenFileName(ofn)
        End If
        If rtn > 0& Then
            rtn = lstrlen(ofn.lpstrFile)
            OpenSaveFileDialog = Left$(ofn.lpstrFile, rtn)
        End If
    End If

ExitRoutine:
End Function

Private Sub CreateNewShadowClass()

    ' the shadow class is static
    ' Whenever the angle or size of your image changes or the shadow attributes change,
    ' the shadow must be recreated.
    ' The other option is to call the RenderDropShadow_JIT routine to draw a shadow on the fly
    '   however that function is far faster on smaller images than larger images.
    '   Recommend using RenderDropShadow_JIT on images < 64x64 and creating a static
    '   shadow class on larger images.
    
    ' Side note: Creating drop shadows on a animated/rotating images is really inefficient. In that
    ' case recommend the following action.
    '   1. Create new DIB of size: source image.Width + dropshadow's blur depth*2 & image.height + blur depth*2
    '           Dim newDIB As c32bppDIB, blurDepth As Long
    '           Set newDIB = New c32bppDIB
    '           blurDepth = 4
    '           newDIB.InitializeDIB srcDIB.Width + blurDepth * 2, srcDIB.Height + blurDepth * 2
    '   2. Call sourceDIB.RenderDropShadow_JIT to the new DIB offsetting the shadow as desired & passing blur depth
    '           dibDC = newDIB.LoadInDC(True)
    '           srcDIB.RenderDropShadow_JIT dibDC, 0, 0, blurDepth, vbBlue, 55
    '           newDIB.LoadDIBinDC False
    '   3. Now sourceDIB.Render the image to the new DIB, passing the new DIB class as one of the optional paramneters
    '           srcDIB.Render 0&, , , , , , , , , , , , newDIB  ' shadow will be right & bottom of image
    '   4. Use this new DIB for rotation
    
    Dim blurDepth As Long
    Dim Color As Long
    Dim i As Integer
    
    For i = mnuShadowDepth.LBound To mnuShadowDepth.UBound
        If mnuShadowDepth(i).Checked = True Then
            Select Case i
            Case mnuShadowDepth.LBound: blurDepth = 2
            Case mnuShadowDepth.UBound: blurDepth = 8
            Case Else: blurDepth = 4
            End Select
            Exit For
        End If
    Next
    mnuShadowDepth(0).Tag = blurDepth
    
    For i = mnuShadowColor.LBound To mnuShadowColor.UBound
        If mnuShadowColor(i).Checked = True Then
            Select Case i
            Case mnuShadowColor.LBound: Color = vbBlack
            Case mnuShadowColor.UBound: Color = vbBlue
            Case mnuShadowColor.LBound + 1: Color = vbRed
            Case mnuShadowColor.UBound - 1: Color = &H8080&
            End Select
            Exit For
        End If
    Next
    
    Set cShadow = cImage.CreateDropShadow(blurDepth, Color)
    cShadow.gdiToken = m_GDItoken   ' assign shared token if one exists
    RefreshImage
    
End Sub

