VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00555555&
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frmRenderSettings 
      BackColor       =   &H00555555&
      Caption         =   "Settings for Chart-Rendering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   12240
      TabIndex        =   10
      Top             =   2700
      Width           =   2895
      Begin VB.ComboBox cmbBSpline 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "fMain.frx":0000
         Left            =   180
         List            =   "fMain.frx":0010
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkShowMinIndicators 
         BackColor       =   &H00555555&
         Caption         =   "Show Min-Indicators"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1140
         Value           =   1  'Aktiviert
         Width           =   2415
      End
      Begin VB.CheckBox chkShowMaxIndicators 
         BackColor       =   &H00555555&
         Caption         =   "Show Max-Indicators"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Aktiviert
         Width           =   2355
      End
   End
   Begin VB.ComboBox cmbGroupBucketSize 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "fMain.frx":0053
      Left            =   12240
      List            =   "fMain.frx":006C
      Style           =   2  'Dropdown-Liste
      TabIndex        =   5
      Top             =   1980
      Width           =   2895
   End
   Begin VB.ComboBox cmbTimeRange 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "fMain.frx":00E7
      Left            =   12240
      List            =   "fMain.frx":0100
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ComboBox cmbRightMostTime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "fMain.frx":0173
      Left            =   12240
      List            =   "fMain.frx":0175
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   420
      Width           =   2895
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  '2D
      BackColor       =   &H00333333&
      ForeColor       =   &H80000008&
      Height          =   8235
      Left            =   180
      ScaleHeight     =   547
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   791
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   180
      Width           =   11895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grp.-BucketSize (for Avg, Min, Max)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   9
      Top             =   1740
      Width           =   3195
   End
   Begin VB.Label lblTimeRange 
      BackStyle       =   0  'Transparent
      Caption         =   "Time-Anchor (Rightmost-Time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   8
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Range (back from Rightmost-Time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   7
      Top             =   960
      Width           =   3195
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MemDB As cMemDB, Srf As cCairoSurface, Chart As New cChart

Private Sub Form_Load()
  Set MemDB = New_c.MemDB
  
  New_c.Timing True
    AddDemoData "2015-04-11 00:00:00" 'add 2*86400 random Demo-Prices for each second of two whole days
  Caption = "Added " & MemDB.GetCount("Ch1") & " Records in" & New_c.Timing & _
            " (use the MouseWheel to scroll through focused ComboBoxes)"
 
  'let's fill the RightMostTime-Combo with Time-Values in 10-minute decrements (from Max-TS downwards)
  Dim FirstTS As Date: FirstTS = MemDB.GetMin("Ch1", "TS") 'get the oldest TimeStamp in our Data-Set
  Dim ComboTS As Date: ComboTS = MemDB.GetMax("Ch1", "TS") 'get the most recent TimeStamp in our Data
  Do Until ComboTS < FirstTS
    cmbRightMostTime.AddItem MemDB.Cnn.GetDateString(ComboTS)
    ComboTS = DateAdd("n", -10, ComboTS) 'decrement by 10 minutes
  Loop
  
  cmbRightMostTime.ListIndex = 0
  cmbTimeRange.ListIndex = 0
  cmbGroupBucketSize.ListIndex = 0
  cmbBSpline.ListIndex = 3
  Redraw False
End Sub

Private Sub AddDemoData(StartFrom As Date)
  MemDB.BeginTrans
    'create a table which takes up Value-Pairs, made out of a TimeStamp- and a Price-Field
    MemDB.Exec "Create Table Ch1(TS Double, Price Double)"
 
    With MemDB.CreateCommand("Insert Into Ch1 Values(?,?)")
      Rnd -1 ' Randomize
      Dim i As Long
      For i = 0 To 86400 * 2 - 1 'loop over all the seconds in two days
        .SetDouble 1, StartFrom + i / 86400 'add these second-fractions to the Start-DateTime
        .SetDouble 2, Round(1 + Rnd * 9, 2) 'add Prices between 5 and 105
        .Execute 'add the new record into the Memory-DB-Table
      Next
    End With
    
    MemDB.Exec "Create Index idx_Ch1_TS On Ch1(TS)" 'an index to speed up the TimeStamp-related Range-queries
  MemDB.CommitTrans
End Sub

Private Sub cmbRightMostTime_Click()
  If cmbRightMostTime.Visible Then Redraw
End Sub
Private Sub cmbTimeRange_Click()
  If cmbTimeRange.Visible Then Redraw
End Sub
Private Sub cmbGroupBucketSize_Click()
  If cmbGroupBucketSize.Visible Then Redraw
End Sub

Private Sub cmbBSpline_Click()
  If cmbBSpline.Visible Then Redraw
End Sub
Private Sub chkShowMaxIndicators_Click()
  If chkShowMaxIndicators.Visible Then Redraw
End Sub
Private Sub chkShowMinIndicators_Click()
   If chkShowMaxIndicators.Visible Then Redraw
End Sub

Private Sub Redraw(Optional ByVal ShowTiming As Boolean = True)
  New_c.Timing True
    'first we prepare a Surface in the same Pixel-size as the Target-PictureBox
    picChart.ScaleMode = vbPixels
    Set Srf = Cairo.CreateSurface(picChart.ScaleWidth, picChart.ScaleHeight)
    
    'Ok, what we now need is the Data (delivered in a cRecordset from the MemDB)
    Dim Rs As cRecordset
    Set Rs = GetData(CDate(cmbRightMostTime), Val(cmbTimeRange) / 60, Val(cmbGroupBucketSize))
    
    '...and a Pattern for the BackGround (which is optional, here we use a plain white color)
    Dim Pat As cCairoPattern 'for plain white BackGround, alternatively: Set Pat = Cairo.CreateSolidPatternLng(vbWhite)
    Set Pat = Cairo.ImageList.AddImage("bgPat", App.Path & "\bgPat.png").CreateSurfacePattern
        Pat.Extend = CAIRO_EXTEND_REPEAT
    
    Chart.ForeColor = &HEEEEEE ' ForeColor (for the Axes Text-Labels and the Axes-LineGrid)
    Chart.AxisAdjustmentFacX = 1.44 'for a more plausible X-Axis-Split for DateTime-Formats (24h*60min=1440, default=1)
    Chart.FormatStrX = "MM\/DD hh\:nn\:ss" 'Format-String for our DateTime-Values at the x-Axis
    Chart.FormatStrY = "0.0" '...and another one for the Text-Renderings at the y-Axis
    Chart.OffsL = 50: Chart.OffsR = 50 '...as well as Left, Right...
    Chart.OffsT = 25: Chart.OffsB = 60 '...and Top, Bottom-Offsets in Pixels
    Chart.BSplineInterpolate = cmbBSpline.ListIndex
    Chart.ShowMaxIndicators = chkShowMaxIndicators
    Chart.ShowMinIndicators = chkShowMinIndicators
   
    Chart.Render Srf.CreateContext, Rs, Pat ' to be able to finally render the Chart
    Set picChart.Picture = Srf.Picture '...and place the result in our Picture-Box
  If ShowTiming Then Caption = "Query and Rendering took: " & New_c.Timing
End Sub
 
Private Function GetData(ByVal MaxDate#, Optional ByVal HoursBack# = 1, Optional ByVal GroupingSeconds& = 60) As cRecordset
  With New_c.StringBuilder
    .AppendNL "Select Avg(TS) AvgT, Avg(Price) AvgP, Min(Price) MinP, Max(Price) MaxP From Ch1"
    .AppendNL "Where TS Between " & Str(MaxDate - HoursBack / 24) & "+1e-6 And " & Str(MaxDate)
    .AppendNL "Group By CLng(0.500001 + TS*24*60/" & Str(GroupingSeconds / 60) & ") Order By TS"

    Set GetData = MemDB.GetRs(.ToString)
  End With
End Function

Private Sub Form_Terminate()
  If Forms.Count = 0 Then New_c.CleanupRichClientDll
End Sub
