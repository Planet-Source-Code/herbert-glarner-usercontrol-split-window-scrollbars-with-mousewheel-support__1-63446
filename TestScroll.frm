VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\AGandaraControls.vbp"
Begin VB.Form TestScroll 
   Caption         =   "Testing Scrollbars and Window Splitting"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.StatusBar staTestScroll 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4815
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   952
            MinWidth        =   952
            Text            =   "Horiz."
            TextSave        =   "Horiz."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Key             =   "Hor"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Vert. Top"
            TextSave        =   "Vert. Top"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Key             =   "VertTop"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Vert. Bot."
            TextSave        =   "Vert. Bot."
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Key             =   "VertBottom"
         EndProperty
      EndProperty
   End
   Begin GandaraControls.gucSplitWindow gswWindow 
      Height          =   3015
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   5318
   End
   Begin VB.Menu mnuScrollbars 
      Caption         =   "&Scrollbars"
      Begin VB.Menu mncSBNone 
         Caption         =   "&None"
      End
      Begin VB.Menu mncSBHorizontal 
         Caption         =   "&Horizontal"
      End
      Begin VB.Menu mncSBVertical 
         Caption         =   "&Vertical"
      End
      Begin VB.Menu mncSBBoth 
         Caption         =   "&Both"
      End
   End
   Begin VB.Menu mnuSplitter 
      Caption         =   "Splitte&r"
      Begin VB.Menu mncSplitterTop 
         Caption         =   "&Top"
      End
      Begin VB.Menu mncSplitterMiddle 
         Caption         =   "&Middle"
      End
      Begin VB.Menu mncSplitterBottom 
         Caption         =   "&Bottom"
      End
   End
End
Attribute VB_Name = "TestScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim picTop As PictureBox, picBottom As PictureBox

Private Sub Form_Load()
    'Setting the two work areas just once. Draw on them like on any other PicBox.
    'Set this *BEFORE* there is a chance for a Resize event for changed client areas
    '(which is the case as soon as you use "ScrollBars").
    Set picTop = gswWindow.WorkArea(egswSCTop)
    Set picBottom = gswWindow.WorkArea(egswSCBottom)
    
    With gswWindow
        'Position the Split Control.
        .Top = 0
        .Left = 0
        
        'Initialize the scrollbars (do that whenever the values change, and also
        'when redisplaying a formerly undisplayed scrollbar)
        .VTopScroll 1, 100, 1, 20, 1    'Order: Min, Max, SmallChange, LargeChange, Value
        .VBottomScroll 1, 100, 1, 20, 1
        .HScroll 1, 100, 1, 20, 1
        
        'Activate the desired scrollbars. Each scrollbar to be displayed *requires*
        'that you use a prior "VScroll" resp. "HScroll".
        .ScrollBars = egswSBBoth
    End With
End Sub

'Usercontrol dimensions changed.
Private Sub Form_Resize()
    Resize
End Sub

'Workarea dimensions changed (after altering the actually displayed scrollbars)
Private Sub gswWindow_Resize()
    Redraw
End Sub

'Event is triggered whenever the window splitter is moved to another position.
Private Sub gswWindow_Split(VPos As Long)
    Redraw
End Sub

Private Sub Resize()
    'Adjust the Split Control to use the whole space.
    With gswWindow
        .Width = ScaleWidth
        If ScaleHeight - staTestScroll.Height > 0& Then
            .Height = ScaleHeight - staTestScroll.Height
        End If
    End With
    
    Redraw
End Sub

Private Sub Redraw()
    'Just demonstrating that the ScaleWidth/Height properties are set properly.
    
    'Top split
    picTop.Cls
    picTop.Line (1, 1)-(picTop.ScaleWidth, picTop.ScaleHeight), vbGreen
    
    picTop.CurrentX = 0: picTop.CurrentY = 0
    picTop.Print "Start of Top picture"
    picTop.CurrentX = 0: picTop.CurrentY = picTop.ScaleHeight - picTop.TextHeight("x")
    picTop.Print "End of Top picture, use Splitter below to split window"
    
    'Bottom split
    picBottom.Cls
    picBottom.Line (1, 1)-(picBottom.ScaleWidth, picBottom.ScaleHeight), vbRed
    
    picBottom.CurrentX = 0: picBottom.CurrentY = 0
    picBottom.Print "Start of Bottom picture, use Splitter above to split window"
    picBottom.Print "Shift+MouseWheel moves the horizontal scrollbar"
    picBottom.CurrentX = 0: picBottom.CurrentY = picBottom.ScaleHeight - picBottom.TextHeight("x")
    picBottom.Print "End of Bottom picture"
End Sub


'Receiving a notification about which scrollbar changed to what value.
Private Sub gswWindow_Change(Area As GandaraControls.egswScrollControl, Scrollbar As GandaraControls.egscSBOrientation, Value As Long)
    Dim sKey As String
    
    If Scrollbar = egscSBOHorizontal Then
        sKey = "Hor"
    Else
        If Area = egswSCTop Then sKey = "VertTop" Else sKey = "VertBottom"
    End If
    staTestScroll.Panels(sKey) = CStr(Value)
End Sub


'Changing the displayed scrollbars
Private Sub mncSBNone_Click()
    With gswWindow
        .ScrollBars = egswSBNone
    End With
End Sub
Private Sub mncSBHorizontal_Click()
    With gswWindow
        .HScroll 1, 100, 1, 20, 1
        .ScrollBars = egswSBHorizontal
    End With
End Sub
Private Sub mncSBVertical_Click()
    With gswWindow
        .VTopScroll 1, 100, 1, 20, 1
        .VBottomScroll 1, 100, 1, 20, 1
        .ScrollBars = egswSBVertical
    End With
End Sub
Private Sub mncSBBoth_Click()
    With gswWindow
        .VTopScroll 1, 100, 1, 20, 1
        .VBottomScroll 1, 100, 1, 20, 1
        .HScroll 1, 100, 1, 20, 1
        .ScrollBars = egswSBBoth
    End With
End Sub


'Some prominent split positions. Just drag & drop to adjust individually.
Private Sub mncSplitterTop_Click()
    gswWindow.Split = 0
End Sub
Private Sub mncSplitterMiddle_Click()
    gswWindow.Split = ScaleHeight \ 2
End Sub
Private Sub mncSplitterBottom_Click()
    gswWindow.Split = ScaleHeight
End Sub

