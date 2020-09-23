VERSION 5.00
Begin VB.UserControl gucSplitWindow 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picSplitting 
      Appearance      =   0  '2D
      BackColor       =   &H80000015&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   930
      MousePointer    =   7  'Größenänderung N S
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   243
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2100
      Width           =   3645
   End
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  '2D
      BackColor       =   &H80000016&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   1260
      MousePointer    =   7  'Größenänderung N S
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3165
   End
   Begin GandaraControls.gucScrollControl gscGridBottom 
      Height          =   1245
      Left            =   960
      TabIndex        =   2
      Top             =   2190
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2196
   End
   Begin GandaraControls.gucScrollControl gscGridTop 
      Height          =   1845
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3254
   End
End
Attribute VB_Name = "gucSplitWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FILE INFO
'---------
'Control:       gucSplitWindow
'IPN prefix:    gsw
'Purpose:       Splitted window, scrollbars, mousewheel support, based on "gucScrollControl".
'               What is drawn into the split windows is up to the user. The client areas are
'               exposed as PictureBox controls: the user can use them like any other PictureBox.

'Author:        Herbert Glarner
'Contact:       herbert.glarner@bluewin.ch
'Copyright:     (c) 2005 by Herbert Glarner
'               Freeware, provided you include credits and mail.



'CONSTITUTING CONTROLS
'---------------------
'gscGridTop         Provides a vertical scrollbar at the right edge only. May be hidden.
'gscGridBottom      Provides scrollbars at the right and bottom edge. Always displayed.
'picSplitter        Separates the two ScrollControls. Used to locate the splitter.
'picSplitting       Moving splitter while in action. Used as a feedback to the user.



'USAGE
'-----
'(1) Add a "gucSplitWindow" control into your form and name it.
'(2) Define 2 global "PictureBox" variables in the form: one for each work area, e.g.:
'    "Dim picTop As PictureBox, picBottom As PictureBox"
'(3) Assign the 2 work areas to these variables (just once, e.g. in the "Load" procedure)
'    "Set picTop = gswGrid.WorkArea(egswSCTop)"
'    "Set picBottom = gswGrid.WorkArea(egswSCBottom)"
'(4) Set the initial scrollbar properties if desired, using "VScroll" and "HScroll", and
'    display them, if desired, using the "Scrollbars" property.
'(-) Use the 2 PictureBox variables ("picTop" and "picBottom" like you would use any other
'    PictureBox control (e.g. "picTop.Line ...", "picTop.Print ...", "x = picTop.ScaleWidth")
'(-) Change the scrollbar values whenever appropriate ("VScroll" and "HScroll"), and display
'    or hide them using the "Scrollbars" property.



'PUBLIC PROPERTIES
'-----------------
'Scrollbars         r/w     Sets/Retrieves the currently set scrollbars.
'Split              r/w     Sets/Retrieves the current window split position.
'Value              r/-     Returns the position of one of the scrollbars
'WorkArea           r/-     Returns one of the 2 picture boxes representing the client
'                           areas into which to draw. Redraw on "Resize" and "Split" events.


'PUBLIC METHODS
'--------------
'HScroll            Sets all properties(1) for the horizontal scrollbar.
'VScroll            Sets all properties(1) for the vertical scrollbars.
'(1): Min, Max, SmallChange, LargeChange, Value

'PRIVATE FUNCTIONS
'-----------------
'vHideScroll        Hides one or both scrollbars.
'vMouseWheel        Processes mouse wheel events received by the "gucScrollControl" objects.
'vResize            Resizes the constituting controls as per UserControl dimensions.
'vSetScroll         Handles settings for one of the 2 scrollbars, in one of the 2 split windows.
'vSetSplitter       Splits the windows as per position of the picSplitter control.
'vShowScroll        Displays one or both scrollbars.



'ENUMS
'-----

'Differentiating between the two scroll controls.
'Public, because public properties ans methods expect this to identify which workarea (the
'upper or lower window) is referred to.
Public Enum egswScrollControl
    egswSCTop = 0&
    egswSCBottom = 1&
End Enum

'Not compatible with "egscSBDefinition" of "gucScrollControl" (this one is an internal Enum
'which lets us combine scrollbars including the possibility of referrng to none, whereas
'"egscSBDefinition" is an Enum used in Windows API calls).
Public Enum egswScrollbars
    egswSBNone = 0&
    egswSBHorizontal = 1&
    egswSBVertical = 2&
    egswSBBoth = 3&
End Enum



'TYPES
'-----

'MS's POINTAPI structure. Used in the "ScreenToClient" API.
Private Type tgswPoint
    X As Long
    Y As Long
End Type



'PRIVATE VARIABLES
'-----------------

'Mouse wheel without pressed key: vertical scrollbar.
'Mouse wheel with pressed Shift key: horizontal scrollbar.
Private glWheeled(egswSCTop To egswSCBottom, egscSBOHorizontal To egscSBOVertical) As Long

'Top position of the splitter. 0 = window is not split.
Private glSplitterTop As Long

'Splitter variables while in action.
Private glSplitterVStart As Long

'Which of the two scroll controls currently has the focus.
Private glFocussedGrid As egswScrollControl

'Stores which scrollbars are active right now
Private glScrollbars As egswScrollbars



'EXPOSED EVENTS
'--------------

'Raised when there is a new splitter position. User needs to redraw the client areas.
Public Event Split(VPos As Long)

'Raised when the client areas were resized. User needs to redraw the client areas.
Public Event Resize()

'Raised whenever a scrollbar is about to change. Use the "SuggestedChange" and "SuggestedPos"
'values, resp., to override the suggested value, if you need a precise control (start of lines,
'grid cells etc.).
'Vertical:
Public Event LineUp(Area As egswScrollControl, SuggestedChange As Long)
Public Event LineDown(Area As egswScrollControl, SuggestedChange As Long)
Public Event PageUp(Area As egswScrollControl, SuggestedChange As Long)
Public Event PageDown(Area As egswScrollControl, SuggestedChange As Long)
Public Event VScroll(Area As egswScrollControl, SuggestedPos As Long)
'Horizontal:
Public Event PosLeft(Area As egswScrollControl, SuggestedChange As Long)
Public Event PosRight(Area As egswScrollControl, SuggestedChange As Long)
Public Event PageLeft(Area As egswScrollControl, SuggestedChange As Long)
Public Event PageRight(Area As egswScrollControl, SuggestedChange As Long)
Public Event HScroll(Area As egswScrollControl, SuggestedPos As Long)

'Raised when any of the scrollbars gets a new value.
Public Event Change(Area As egswScrollControl, Scrollbar As egscSBOrientation, Value As Long)




'EXTERNAL API DECLARATIONS
'-------------------------

'To detect if the mousepointer is within the control
Private Declare Function ScreenToClient Lib "user32.dll" _
    (ByVal hWnd As Long, ByRef lpPoint As tgswPoint) As Long



'CONSTRUCTOR AND DESTRUCTOR
'--------------------------

Private Sub UserControl_Initialize()
    'Initially, no scrollbars are displayed.
    gscGridTop.ActiveScrollbars = egscSBDBoth
    gscGridTop.HideScrollbars
    gscGridBottom.ActiveScrollbars = egscSBDBoth
    gscGridBottom.HideScrollbars
    
    'Initially, the window is not split.
    glSplitterTop = 0&
    picSplitting.Visible = False
End Sub

Private Sub UserControl_Terminate()
    'Nothing to take care for.
End Sub


'USERCONTROL EVENTS
'------------------
Private Sub UserControl_Resize()
    vResize
End Sub



'MOVING THE MOUSE SPLITTER
'-------------------------

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSplitting.Visible = True
    
    'Adjusting exact position while moving
    glSplitterVStart = CLng(Y)
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lPos As Long
    
    'Calculate the vertical position of the splitter. Enforce to stay within the control.
    '"glSplitterVStart" was thevertical position *within* the splitter when the splitter
    'was activated with "MouseDown".
    lPos = CLng(Y) + picSplitter.Top - glSplitterVStart
    If lPos < 0& Then
        lPos = 0&
    ElseIf lPos >= gscGridBottom.Top + gscGridBottom.ScaleHeight Then
        lPos = gscGridBottom.Top + gscGridBottom.ScaleHeight
    End If

    'We don't start at position 0, but 1 to not paint over the control's left border. For
    'the same reason, we draw the width 2 pixels smaller.
    picSplitting.Move 1&, lPos, ScaleWidth - 2&
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Deactivate the splitter
    picSplitting.Visible = False
    
    vSetSplitter CLng(Y)
End Sub



'MOUSE WHEEL NOTIFICATIONS
'-------------------------
'Note: Simply using the wheel within the control differentiates between the upper and the
'      lower grid. Using the wheel while the Shift key is pressed, the (only) horizontal
'      scrollbar is used.


'The value "Y" of the "MouseWheel" events depends on what control has the focus. It is
'registered here into the gloval variable "glFocussedGrid".
Private Sub gscGridTop_GotFocus()
    glFocussedGrid = egswSCTop
End Sub
Private Sub gscGridBottom_GotFocus()
    glFocussedGrid = egswSCBottom
End Sub

'Handling the mouse wheel. It's only one of the two controls which receives the event for
'both controls. We handle them collectively in the "vMouseWheel" procedure.
Private Sub gscGridTop_MouseWheel(hWnd As Long, X As Long, Y As Long, Value As Long, Key As egscMouseKeys)
    vMouseWheel hWnd, X, Y, Value, Key
End Sub
Private Sub gscGridBottom_MouseWheel(hWnd As Long, X As Long, Y As Long, Value As Long, Key As egscMouseKeys)
    vMouseWheel hWnd, X, Y, Value, Key
End Sub



'SCROLLBAR NOTIFICATIONS
'-----------------------

'Top client area
Private Sub gscGridTop_LineDown(SuggestedChange As Long)
    RaiseEvent LineDown(egswSCTop, SuggestedChange)
End Sub
Private Sub gscGridTop_LineUp(SuggestedChange As Long)
    RaiseEvent LineUp(egswSCTop, SuggestedChange)
End Sub
Private Sub gscGridTop_PageDown(SuggestedChange As Long)
    RaiseEvent PageDown(egswSCTop, SuggestedChange)
End Sub
Private Sub gscGridTop_PageUp(SuggestedChange As Long)
    RaiseEvent LineUp(egswSCTop, SuggestedChange)
End Sub
Private Sub gscGridTop_VScroll(SuggestedPos As Long)
    RaiseEvent VScroll(egswSCTop, SuggestedPos)
End Sub

'Bottom client area
Private Sub gscGridBottom_LineDown(SuggestedChange As Long)
    RaiseEvent LineDown(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_LineUp(SuggestedChange As Long)
    RaiseEvent LineUp(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_PageDown(SuggestedChange As Long)
    RaiseEvent PageDown(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_PageLeft(SuggestedChange As Long)
    RaiseEvent PageLeft(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_PageRight(SuggestedChange As Long)
    RaiseEvent PageRight(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_PageUp(SuggestedChange As Long)
    RaiseEvent PageUp(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_PosLeft(SuggestedChange As Long)
    RaiseEvent PosLeft(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_PosRight(SuggestedChange As Long)
    RaiseEvent PosRight(egswSCBottom, SuggestedChange)
End Sub
Private Sub gscGridBottom_HScroll(SuggestedPos As Long)
    RaiseEvent HScroll(egswSCBottom, SuggestedPos)
End Sub
Private Sub gscGridBottom_VScroll(SuggestedPos As Long)
    RaiseEvent VScroll(egswSCBottom, SuggestedPos)
End Sub

'When the scrollbars change (scrollbar, mousewheel, programmatically), the user need to be informed.
Private Sub gscGridTop_Change(Scrollbar As egscSBOrientation, Value As Long)
    'Only raise event if the scrollbar is shown.
    '(Scrollbar+1 is the corresponding bit (1 or 2) in egswScrollbars)
    If glScrollbars And (Scrollbar + 1&) Then RaiseEvent Change(egswSCTop, Scrollbar, Value)
End Sub
Private Sub gscGridBottom_Change(Scrollbar As egscSBOrientation, Value As Long)
    'Only raise event if the scrollbar is shown.
    If glScrollbars And (Scrollbar + 1&) Then RaiseEvent Change(egswSCBottom, Scrollbar, Value)
End Sub



'PUBLIC METHODS
'--------------

'Define the scrollbars.
Public Sub VTopScroll(Min As Long, Max As Long, SmallChange As Long, LargeChange As Long, Value As Long)
    vSetScroll gscGridTop, egscSBOVertical, Min, Max, SmallChange, LargeChange, Value
End Sub
Public Sub VBottomScroll(Min As Long, Max As Long, SmallChange As Long, LargeChange As Long, Value As Long)
    vSetScroll gscGridBottom, egscSBOVertical, Min, Max, SmallChange, LargeChange, Value
End Sub
Public Sub HScroll(Min As Long, Max As Long, SmallChange As Long, LargeChange As Long, Value As Long)
    vSetScroll gscGridBottom, egscSBOHorizontal, Min, Max, SmallChange, LargeChange, Value
End Sub



'PUBLIC PROPERTIES
'-----------------

'Retrieve the client work area onto which to draw (read-only).
'No worries about "Object", this needs not to be called more than once per client session.
'However, it's the only possibility to return a "PictureBox" object.
Public Property Get WorkArea(Area As egswScrollControl) As Object
    If Area = egswSCTop Then
        Set WorkArea = gscGridTop.WorkArea
    Else
        Set WorkArea = gscGridBottom.WorkArea
    End If
End Property

'Return the value of one of the scrollbars.
'Note that we have up to 3 scrollbars (2 vertical ones and 1 horizontal one).
Public Property Get Value(Area As egswScrollControl, Scrollbar As egscSBOrientation) As Long
    If Area = egswSCTop Then
        Value = gscGridTop.Value(Scrollbar)
    Else
        Value = gscGridBottom.Value(Scrollbar)
    End If
End Property

'Set the value of one of the scrollbars.
'Note that we have up to 3 scrollbars (2 vertical ones and 1 horizontal one).
Public Property Let Value(Area As egswScrollControl, Scrollbar As egscSBOrientation, NewValue As Long)
    If Area = egswSCTop Then
        gscGridTop.Value(Scrollbar) = NewValue
    Else
        gscGridBottom.Value(Scrollbar) = NewValue
    End If
End Property

'Retrieve the currently set scrollbars.
Public Property Get Scrollbars() As egswScrollbars
    Scrollbars = glScrollbars
End Property

'Set the desired horizontal and/or vertical scrollbars.
'Note, that after hiding a scrollbar you *must* use "Vscroll" resp. "HScroll" to set new
'values, because for undisplayed scrollbars we need to invalidate their values (or they will
'pop up as soon as there is a chance to, e.g. rotating the mouse wheel).
Public Property Let Scrollbars(ActiveScrollbars As egswScrollbars)
    'If nothing changes, then we don't need to bother the controls.
    If glScrollbars <> ActiveScrollbars Then
        If ActiveScrollbars = egswSBNone Then
            vHideScroll gscGridTop, egscSBDBoth
            vHideScroll gscGridBottom, egscSBDVertical
        ElseIf ActiveScrollbars = egswSBVertical Then
            vHideScroll gscGridBottom, egscSBDHorizontal
            vShowScroll gscGridTop, egscSBDVertical
            vShowScroll gscGridBottom, egscSBDVertical
        ElseIf ActiveScrollbars = egswSBHorizontal Then
            vHideScroll gscGridTop, egscSBDVertical
            vHideScroll gscGridBottom, egscSBDVertical
            vShowScroll gscGridBottom, egscSBDHorizontal
        ElseIf ActiveScrollbars = egswSBBoth Then
            vShowScroll gscGridTop, egscSBDVertical
            vShowScroll gscGridBottom, egscSBDBoth
        End If
        
        'Invalidating the scrollbar values for the now hidden scrollbars.
        'Requires "VScroll" resp. "HScroll" when they need to be redisplayed.
        If (ActiveScrollbars And egswSBVertical) = 0& Then
            vSetScroll gscGridTop, egscSBOVertical, 1, 1, 1, 1, 1
            vSetScroll gscGridBottom, egscSBOVertical, 1, 1, 1, 1, 1
        End If
        If (ActiveScrollbars And egswSBHorizontal) = 0& Then
            vSetScroll gscGridBottom, egscSBOHorizontal, 1, 1, 1, 1, 1
        End If
        
        'Store globally
        glScrollbars = ActiveScrollbars
        
        'Adjust the work areas.
        gscGridTop.WorkArea.Width = gscGridTop.ScaleWidth
        gscGridBottom.WorkArea.Width = gscGridBottom.ScaleWidth
        gscGridBottom.WorkArea.Height = gscGridBottom.ScaleHeight
        
        'Raise Resize event, because changes in the dimensions of the client areas
        'requires that the owner redraws them.
        RaiseEvent Resize
    End If
End Property

'Returns the current split position.
Public Property Get Split() As Long
    Split = picSplitter.Top
End Property

'Split the window programmatically.
Public Property Let Split(VPos As Long)
    picSplitter.Top = VPos      'No need to check here, vSetSplitter will take care for.
    glSplitterVStart = 0&       'No handling adjustments
    vSetSplitter 0&             'No handling adjustments
End Property



'PRIVATE METHODS
'---------------

'Called from: "gscGridTop_MouseWheel" and "gscGridBottom_MouseWheel".
'Collectively handles the mouse wheel for both scroll controls.
Private Sub vMouseWheel(hWnd As Long, ByVal X As Long, ByVal Y As Long, Value As Long, Key As egscMouseKeys)
    Dim lScrollbar As egscSBOrientation, rPoint As tgswPoint
    Dim gscGrid As gucScrollControl, lGrid As egswScrollControl
    
    'Convert the absolute coordinates into such relative to the upper-left corner of the specified
    'client area. Note, that Y can refer to either control, depending on which has the focus.
    With rPoint: .X = X: .Y = Y: ScreenToClient hWnd, rPoint: X = .X: Y = .Y: End With

    'Leaving if it's not for us. Y is tested further below, because we need to consider
    '2 controls in dependence of 2 possible focus owners.
    If X < 0 Or X >= ScaleWidth Then Exit Sub
    
    'The wheel delta usually is a multiple of WHEEL_DELTA which is 120. However, other
    'vendors may provide smaller or arbitrary values. This means that we have to sum up all
    'values, until 120 resp. -120 was exceeded.
    'Because we define, that a mouse wheel activity without any keys relates to the vertical
    'scrollbar and one with a pressed Shift key to the horizontal scrollbar, we keep track of
    'these values individually.
    If Key And egscMKShift Then
        lScrollbar = egscSBOHorizontal
        'In this case, it *always* is the bottom grid which is affected. There never is a
        'horizontal scrollbar for the upper grid. (Why? Just because it looks weird and
        'unfamiliar. However, implement the according logic if you need it.)
        Set gscGrid = gscGridBottom
        lGrid = egswSCBottom
    Else
        lScrollbar = egscSBOVertical
        'For vertical scrollbars we differentiate which grid we mean.
        'The value "Y" depends on what control has the focus. The control with the focus is in
        'the global variable "goFocussedGrid". If it is the lower control, then Y must be
        'transposed so, that it relates to the upper.
        If glFocussedGrid = egswSCBottom Then
            Y = Y + gscGridBottom.Top
        End If
        If Y >= gscGridTop.Top And Y < gscGridTop.Top + gscGridTop.Height Then
            Set gscGrid = gscGridTop
            lGrid = egswSCTop
        ElseIf Y >= gscGridBottom.Top And Y < gscGridBottom.Top + gscGridBottom.Height Then
            Set gscGrid = gscGridBottom
            lGrid = egswSCBottom
        Else
            'Sorry, this is for none of our grids.
            Exit Sub
        End If
    End If
    
    'lDelta is positive or negative.
    glWheeled(lGrid, lScrollbar) = glWheeled(lGrid, lScrollbar) + Value
    
    'As long as we have a WHEEL_DELTA of at least 120, we communicate a 'SmallChange' to the
    'owner. (We can hard-code '120', since this is part of the MS spec; however, other vendors
    'or further upgrades may *communicate* different values. Still, 120 will remain the value
    'to check for.)
    'Negative values (towards user) = down/right; positive values (away from user) = up/left.
    Do While Abs(glWheeled(lGrid, lScrollbar)) >= 120&
        If glWheeled(lGrid, lScrollbar) >= 120& Then
            gscGrid.LineUp lScrollbar       'Suggest a LineUp
            glWheeled(lGrid, lScrollbar) = glWheeled(lGrid, lScrollbar) - 120&
        ElseIf glWheeled(lGrid, lScrollbar) <= -120& Then
            gscGrid.LineDown lScrollbar     'Suggest a LineDown
            glWheeled(lGrid, lScrollbar) = glWheeled(lGrid, lScrollbar) + 120&
        End If
    Loop
End Sub

'Resizes the user control.
Private Sub vResize()
    Dim lValue As Long
    
    'Avoids to redraw when minimizing.
    If ScaleHeight <= 2& Or ScaleWidth <= 2& Then Exit Sub
    
    'Position the splitter.
    With picSplitter
        .Left = 0
        .Width = ScaleWidth
        'We need at least 36 pixels in the bottom control to display the scrollbars.
        If glSplitterTop > ScaleHeight - 36& Then
            glSplitterTop = ScaleHeight - 36&
        End If
        If glSplitterTop < 0& Then glSplitterTop = 0&
        .Top = glSplitterTop
    End With
    
    'Depending from the splitter's position, position the 2 scroll controls.
    With gscGridTop
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = picSplitter.Top
    End With
    With gscGridBottom
        .Left = 0
        .Top = picSplitter.Top + picSplitter.Height
        .Width = ScaleWidth
        lValue = ScaleHeight - .Top
        If lValue > 0& Then .Height = lValue
    End With
End Sub

'Handles settings for one of the 2 scrollbars, in one of the 2 split windows.
'Called from the public "VScroll" and "HScroll".
Private Sub vSetScroll(Ctrl As gucScrollControl, Scrollbar As egscSBOrientation, _
    Min As Long, Max As Long, SmallChange As Long, LargeChange As Long, Value As Long)

    With Ctrl
        .Min(Scrollbar) = Min
        .Max(Scrollbar) = Max
        .SmallChange(Scrollbar) = SmallChange
        .LargeChange(Scrollbar) = LargeChange
        .Value(Scrollbar) = Value
        .SetScrollbar Scrollbar
    End With
End Sub

'Displays one or both scrollbars.
Private Sub vShowScroll(Ctrl As gucScrollControl, Scrollbars As egscSBDefinition)
    With Ctrl
        .ActiveScrollbars = Scrollbars
        .ShowScrollbars
    End With
End Sub

'Hides one or both scrollbars.
Private Sub vHideScroll(Ctrl As gucScrollControl, Scrollbars As egscSBDefinition)
    With Ctrl
        .ActiveScrollbars = Scrollbars
        .HideScrollbars
    End With
End Sub

'Splits the windows as per position of the picSplitter control.
'"Y" is used to adjust minor differences when the splitter was grabbed; to programmatically
'perform a split, just set "picSplitter.Top" and let "Y" be 0.
Private Sub vSetSplitter(Y As Long)
    Dim lPos As Long
    
    'Calculate the vertical position of the splitter. Enforce to stay within the control.
    '"glSplitterVStart" was the vertical position *within* the splitter when the splitter
    'was activated with "MouseDown".
    lPos = Y + picSplitter.Top - glSplitterVStart
    If lPos < 0& Then
        lPos = 0&
    ElseIf lPos >= gscGridBottom.Top + gscGridBottom.ScaleHeight Then
        lPos = gscGridBottom.Top + gscGridBottom.ScaleHeight
    End If
    
    'If this is a new position, store it and raise the "Split" event.
    If glSplitterTop <> lPos Then
        'Store this new position globally as the current splitter position.
        glSplitterTop = lPos
        
        'Resize the controls accordingly
        vResize
        
        'Notify the owner.
        RaiseEvent Split(lPos)
    End If
    
    'We need to give back the focus to one of the 2 controls. Since the lower one is
    'always visible, it is our choice. If we don't give back the focus, we won't be able
    'to immediately perform further scroll events.
    gscGridBottom.SetFocus
End Sub
