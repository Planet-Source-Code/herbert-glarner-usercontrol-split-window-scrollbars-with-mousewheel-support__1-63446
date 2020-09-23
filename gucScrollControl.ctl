VERSION 5.00
Begin VB.UserControl gucScrollControl 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picWorkArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      Height          =   1515
      Left            =   1110
      ScaleHeight     =   1515
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   810
      Width           =   2625
   End
End
Attribute VB_Name = "gucScrollControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FILE INFO
'---------
'Control:       gucScrollControl
'IPN prefix:    gsc
'Purpose:       User control with native windows scrollbars, using gifSubClassing for
'               subclassing. Use this control as a basis to design other controls with full
'               scrollbar functionality.

'Author:        Herbert Glarner
'Contact:       herbert.glarner@bluewin.ch
'Copyright:     (c) 2005 by Herbert Glarner
'               Freeware, provided you include credits and mail.



'CONSTITUTING CONTROLS
'---------------------
'picWorkArea        Represents the inner work area. Note, that the client area is automatically
'                   adjusted, when scrollbars are shown/hidden (i.e. "ScaleWidth" and "ScaleHeight"
'                   always reflect the really visible client area).

'INTERFACES
'----------
'gifSubClassing     IDE-safe subclassing [by Paul Caton]



'USAGE
'-----
'(1) Assign the type of scrollbar(s) to be used (Horiz/Vert/Both) with "ActiveScrollbars".
'
'(2) Use "Min", "Max", "LargeChange" and "Value" to specify the scrollbars' properties.
'    Until now, nothing was displayed: we defined just how the future scrollbar(s) will look like.
'
'(3) Above definitions are communicated to windows now.
'    Use "SetScrollbar" to inform Windows about what we want (once for each scrollbar).
'
'(4) Time to display the scrollbar(s) now (both in one go, if both were defined).
'    Display the scrollbar(s) via "ShowScrollbars" (hide them via "HideScrollbars")
'    The scrollbars are displayed and functional now, i.e., they will trigger events.



'PUBLIC PROPERTIES
'-----------------
'ActiveScrollbars   r/w
'LargeChange        r/w
'Max                r/w
'Min                r/w
'ScaleHeight        r/-
'ScaleWidth         r/-
'SmallChange        r/w
'Value              r/w
'WorkArea           r/-

'PUBLIC METHODS
'--------------
'HideScrollbars
'LineDown           Suggest a line down/pos right, raises events as if the scrollbar was clicked.
'LineUp             Suggest a line up/pos left, raises events as if the scrollbar was clicked.
'SetScrollbar
'ShowScrollbars

'PRIVATE METHODS
'---------------
'GetHiWord
'GetLoWord
'ProcessScrollBar



'CONSTANTS
'---------

'Windows messages that we're going to filter for callback.
Private Const gscWMHScroll      As Long = &H114&
Private Const gscWMVScroll      As Long = &H115&
Private Const gscWMMouseWheel   As Long = &H20A&


'ENUMS
'-----

'Pressed keys while rotating the mouse wheel
Public Enum egscMouseKeys
    egscMKShift = 4&
    egscMKControl = 8&
End Enum

'Type of scrollbar. Used in API calls.
Public Enum egscSBDefinition
    egscSBDHorizontal = 0&
    egscSBDVertical = 1&
    egscSBDBoth = 3&
End Enum

'Our properties allow setting the value for either one of the scrollbars, but not
'for both together: we have an individual record of the "tgswScrollInfo" structure
'for each.
Public Enum egscSBOrientation
    egscSBOHorizontal = 0&
    egscSBOVertical = 1&
End Enum

'The scrollbar notification types are delivered in the low word of the DWord
'"wParam". Use the private function "GetLoWord" to extract that word from wParam.
Public Enum egscSBNotification
    'Set scroll value to value - SmallChange
    egscSBNLineLeft = 0
    egscSBNLineUp = 0
    'Set scroll value to value + SmallChange
    egscSBNLineDown = 1
    egscSBNLineRight = 1
    'Set scroll value to value - LargeChange
    egscSBNPageLeft = 2
    egscSBNPageUp = 2
    'Set scroll value to value + LargeChange
    egscSBNPageRight = 3
    egscSBNPageDown = 3
    'Set scroll value to track position, Track Event if wanted
    egscSBNThumbTrack = 5       'while Tracking
    egscSBNThumbPosition = 4    'End of Tracking
    'Set scroll value to min
    egscSBNLeft = 6
    egscSBNTop = 6
    'Set scroll value to max
    egscSBNRight = 7
    egscSBNBottom = 7
    'Raise a Change Event
    egscSBNEndScroll = 8
End Enum

'Used in the "Mask" field of the structure "tgswScrollInfo".
Public Enum egscScrollInfoMask
    egscSIMRange = &H1
    egscSIMPage = &H2
    egscSIMPos = &H4
    egscSIMDisableNoScroll = &H8
    egscSIMTrackPos = &H10
    egscSIMAll = (egscSIMRange Or egscSIMPage Or egscSIMPos Or egscSIMTrackPos)
End Enum



'TYPES
'-----

'MS's SCROLLINFO structure. Used to set/retrieve scrollbar values.
Private Type tgscScrollInfo
    Size As Long                'Size of (this) structure
    Mask As egscScrollInfoMask  'Values to change
    Min As Long                 'Minimum value of the scrollbar
    Max As Long                 'Maximum value of the scrollbar
    Page As Long                'What VB calls "LargeChange"
    Pos As Long                 'Current value
    TrackPos As Long            '[Is actually in HiWord of wParam]
End Type
Private Const cSizeofScrollInfo As Long = 28&
'Note, that the actual maximal value of the scrollbar is actually equal to the
'structure's "Max" value plus its "Page" value.




'PRIVATE VARIABLES
'-----------------

'Declaring the subclasser
Private gscSubClasser As gclSubClassing  'Declare the subclasser

'Stores the active scrollbar(s). Use "ActiveScrollbars" to set/read this value.
Private glSBDefinition As egscSBDefinition

'We need a "tgswScrollInfo" record per scrollbar, i.e. one each for the
'horizontal (egswSBOHorizontal) and the vertical (egswSBOVertical) scrollbar.
Private grScrollInfo(egscSBOHorizontal To egscSBOVertical) As tgscScrollInfo

'To not destroy above data when it's needed to call the "GetScrollInfo" API (i.e.
'when requesting the 32-bit-thumb value while scrolling), another structure pair
'is defined for that purpose.
Private grScrollInfoTrack(egscSBOHorizontal To egscSBOVertical) As tgscScrollInfo

'We're only raising a "Change" event if there is a new value. This variable holds
'the last value for which such an event was raised.
Private glLastEventValue(egscSBOHorizontal To egscSBOVertical) As Long

'A "small change" is not realized via the structure. Still, we can't assume "1"
'all the time, that depends on the clients implementation. Thus, we store that
'value in a global variable.
'Usually 1, and initialized with that value
Private glSmallChange(egscSBOHorizontal To egscSBOVertical) As Long




'EXPOSED EVENTS
'--------------

'Use these individual events, if you need a precise control (alignments in grids
'and the like). As the second argument implies, this is a *suggested* change value
'only. You can modify this argument and when *your* event procedure was handled the
'changed value will be applied. (You even can 'cancel' the event by setting this
'value to 0).
'Separating the events for the different scrollbars. Vertical scrollbar:
Public Event LineUp(SuggestedChange As Long)
Public Event LineDown(SuggestedChange As Long)
Public Event PageUp(SuggestedChange As Long)
Public Event PageDown(SuggestedChange As Long)
'Horizontal scrollbar
Public Event PosLeft(SuggestedChange As Long)
Public Event PosRight(SuggestedChange As Long)
Public Event PageLeft(SuggestedChange As Long)
Public Event PageRight(SuggestedChange As Long)
'When clicking onto the thumb and when dragging it, a suggested *position* (and
'not a suggested *change* value) is communicated. This position can be manipulated
'by the client's event procedure: if the value is modified, it is that value which
'is applied.
Public Event VScroll(SuggestedPos As Long)
Public Event HScroll(SuggestedPos As Long)

'Raised when there is a new Value for the scrollbar. Communicated *after* above
'events, taking into account a possibly modified suggestion value.
Public Event Change(Scrollbar As egscSBOrientation, Value As Long)


Public Event MouseWheel(hWnd As Long, X As Long, Y As Long, Value As Long, Key As egscMouseKeys)



'API DECLARATIONS
'----------------

'Shows or hides a scrollbar
Private Declare Function ShowScrollBar Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal wBar As egscSBDefinition, _
    ByVal bShow As Boolean) As Long

'Sets the properties of a scrollbar
Private Declare Function SetScrollInfo Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal wBar As egscSBOrientation, _
    ByRef lpScrollInfo As tgscScrollInfo, ByVal bool As Boolean) As Long
    
'Gets the properties of a scrollbar
Private Declare Function GetScrollInfo Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal wBar As egscSBOrientation, _
    ByRef lpScrollInfo As tgscScrollInfo) As Long
    
'Initializing data structures
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
    (Destination As Long, ByVal Length As Long)



'IMTERFACES
'----------

'We're implementing the interfaces declared in iSuperClass. Once the following declaration is in
'place you'll find an entry in the left hand combo-box at the top of the code window for gifSubClassing.
Implements gifSubClassing



'USERCONTROL CONSTRUCTOR AND DESTRUCTOR
'--------------------------------------

Private Sub UserControl_Initialize()
    Dim lSize As Long
    
    'The field "Size" of the two variables of structure type "tgswScrollInfo"
    'needs to be set once only: it won't change.
    lSize = Len(grScrollInfo(egscSBOHorizontal))  'Either one (Hor/Vert) does the job
    grScrollInfo(egscSBOHorizontal).Size = lSize
    grScrollInfo(egscSBOVertical).Size = lSize
    
    grScrollInfo(egscSBOHorizontal).Mask = egscSIMAll
    grScrollInfo(egscSBOVertical).Mask = egscSIMAll
    
    'Initializing the small change value is '1' (can be overwritten with the
    'property "SmallChange").
    glSmallChange(egscSBOHorizontal) = 1&
    glSmallChange(egscSBOVertical) = 1&
    
    'Subclass the scrollbar messages. Create a SubClasser instance.
    Set gscSubClasser = New gclSubClassing
    
    'Position picture box representing the work area.
    picWorkArea.Left = 0&
    picWorkArea.Top = 0&

    'Tell the subclasser which messages to callback on (filtered mode).
    With gscSubClasser
        'Note: There's an optional second parameter to AddMsg which should be set to True if you
        '      wish to receive the message *before* default processing.
        Call .AddMsg(gscWMHScroll, True)
        Call .AddMsg(gscWMVScroll, True)
        Call .AddMsg(gscWMMouseWheel, True)
    
        'Start subclassing.
        Call .Subclass(hWnd, Me)
    End With
End Sub

Private Sub UserControl_Terminate()
    'Destroy the SubClasser.
    Set gscSubClasser = Nothing
End Sub

Private Sub UserControl_Resize()
    'The picture box representing the work area rakes the whole inner client area
    '(without occupying the space needed for the scrollbars).
    picWorkArea.Width = ScaleWidth
    picWorkArea.Height = ScaleHeight
End Sub



'SUBCLASSER INTERFACE MESSAGES
'-----------------------------

Private Sub gifSubClassing_After _
    (lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    
    'Not used, but existence of the Sub is an implementation requirement.
End Sub

'This implemented interface is called BEFORE default processing, i.e. *before* the previous WndProc.
'Set "lReturn" to '0' and "lHandled" to 'True' when the message was handled.
Private Sub gifSubClassing_Before _
    (lHandled As Long, _
     lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
     
    Dim lX As Long, lY As Long
    Dim lDelta As Long, lKeys As egscMouseKeys
    
    Select Case uMsg
        Case gscWMMouseWheel
            'Get the coordinates relative to the *Screen* (not relative to this control or its owner).
            lX = GetLoWord(lParam): lY = GetHiWord(lParam)
            'We need the delta value signed.
            lDelta = (wParam And &HFFFF0000) \ &H10000
            lKeys = GetLoWord(wParam)
            'Owner handles this event.
            RaiseEvent MouseWheel(hWnd, lX, lY, lDelta, lKeys)
            lHandled = True
            lReturn = 0
        Case gscWMHScroll
            'Horizontal scrollbar messages
            ProcessScrollBar egscSBOHorizontal, GetLoWord(wParam), GetHiWord(wParam)
            lHandled = True
            lReturn = 0
        Case gscWMVScroll
            'Vertical scrollbar messages
            ProcessScrollBar egscSBOVertical, GetLoWord(wParam), GetHiWord(wParam)
            lHandled = True
            lReturn = 0
    End Select
End Sub



'PUBLIC PROPERTIES
'-----------------

'Returns the work area (picture box) to the client for direct drawing.
Public Property Get WorkArea() As Object
    Set WorkArea = picWorkArea
End Property

'Assign the type of scrollbar(s) to be displayed, read what type(s) were assigned.
Public Property Let ActiveScrollbars(BarsToDisplay As egscSBDefinition)
    glSBDefinition = BarsToDisplay
End Property
Public Property Get ActiveScrollbars() As egscSBDefinition
    ActiveScrollbars = glSBDefinition
End Property

'Assign/Read the scrollbar property Min/Max/Value/LargeChange/SmallChange for one
'of the two scrollbars (horizontal or vertical).
Public Property Let LargeChange(Scrollbar As egscSBOrientation, NewValue As Long)
    grScrollInfo(Scrollbar).Page = NewValue
End Property
Public Property Get LargeChange(Scrollbar As egscSBOrientation) As Long
    LargeChange = grScrollInfo(Scrollbar).Page
End Property

'(A "small change" is not realized via the structure. Still, we can't assume "1"
'all the time, that depends on the clients implementation. Thus, we store that
'value in a global variable.)
Public Property Let SmallChange(Scrollbar As egscSBOrientation, NewValue As Long)
    glSmallChange(Scrollbar) = NewValue
End Property
Public Property Get SmallChange(Scrollbar As egscSBOrientation) As Long
    SmallChange = glSmallChange(Scrollbar)
End Property

Public Property Let Max(Scrollbar As egscSBOrientation, NewMaximum As Long)
    grScrollInfo(Scrollbar).Max = NewMaximum
End Property
Public Property Get Max(Scrollbar As egscSBOrientation) As Long
    Max = grScrollInfo(Scrollbar).Max
End Property

Public Property Let Min(Scrollbar As egscSBOrientation, NewMinimum As Long)
    grScrollInfo(Scrollbar).Min = NewMinimum
End Property
Public Property Get Min(Scrollbar As egscSBOrientation) As Long
    Min = grScrollInfo(Scrollbar).Min
End Property

Public Property Let Value(Scrollbar As egscSBOrientation, NewValue As Long)
    grScrollInfo(Scrollbar).Pos = NewValue
End Property
Public Property Get Value(Scrollbar As egscSBOrientation) As Long
    Value = grScrollInfo(Scrollbar).Pos
End Property

'Retrieving the work area dimensions (read only)
Public Property Get ScaleHeight() As Long
    ScaleHeight = UserControl.ScaleHeight
End Property
Public Property Get ScaleWidth() As Long
    ScaleWidth = UserControl.ScaleWidth
End Property



'PUBLIC METHODS
'--------------

'Communicating the desired settings (Min, Max, Value, LargeChange) to Windows.
Public Sub SetScrollbar(Scrollbar As egscSBOrientation)
    SetScrollInfo hWnd, Scrollbar, grScrollInfo(Scrollbar), True
End Sub

'Showing the scrollbars as defined in "ActiveScrollbars".
Public Sub ShowScrollbars()
    ShowScrollBar hWnd, glSBDefinition, True
End Sub

'Hide the scrollbars as defined in "ActiveScrollbars".
Public Sub HideScrollbars()
    ShowScrollBar hWnd, glSBDefinition, False
End Sub

'It is possible to tell the control to trigger any event to the owner. Use this instead
'of setting a position with the "Value" property, when your client performs value manipulation
'in order to ensure a dedicated position.
Public Sub LineDown(Scrollbar As egscSBOrientation)
    ProcessScrollBar Scrollbar, egscSBNLineDown, glSmallChange(Scrollbar)
End Sub
Public Sub LineUp(Scrollbar As egscSBOrientation)
    ProcessScrollBar Scrollbar, egscSBNLineUp, glSmallChange(Scrollbar)
End Sub



'PRIVATE METHODS
'---------------
'Processing a scrollbar notification. Called by InterceptedWinMsg for either of
'the two scrollbar orientations (Scrollbar tells for which).
Private Sub ProcessScrollBar(Scrollbar As egscSBOrientation, _
    Notification As egscSBNotification, nPos As Long)
    
    Dim lValue As Long
    Dim lChangeValue As Long            'This is user-modifiable on page/line up/down
    Dim eMask As egscScrollInfoMask
    Dim lEffMax As Long
    
    With grScrollInfo(Scrollbar)
        'The other notifications all change the position (the 'value').
        Select Case Notification
            Case egscSBNThumbTrack, egscSBNThumbPosition
                'Usual 16-bit technique:
                '    'Set scroll value to track position. Here, the scroll position
                '    'is provided in nPos (ex the Hi Word of wParam).
                '    lValue = nPos
                'Circumventing the usual 16 bit value and getting the 32 bit value.
                '   Microsoft states: "The GetScrollInfo function enables applications to use
                '   32-bit scroll positions. Although the messages that indicate scroll-bar position,
                '   WM_HSCROLL and WM_VSCROLL, provide only 16 bits of position data, the functions
                '   SetScrollInfo and GetScrollInfo provide 32 bits of scroll-bar position data.
                '   Thus, an application can call GetScrollInfo while processing either the WM_HSCROLL or
                '   WM_VSCROLL messages to obtain 32-bit scroll-bar position data."
                '   (To not to destroy the data in the usual ScrollInfo structures, we use the separate
                '   structure variable "grScrollInfoTrack()" instead of "grScrollInfo()".)
                ZeroMemory ByVal VarPtr(grScrollInfoTrack(Scrollbar)), cSizeofScrollInfo
                grScrollInfoTrack(Scrollbar).Size = cSizeofScrollInfo
                grScrollInfoTrack(Scrollbar).Mask = egscSIMTrackPos
                GetScrollInfo hWnd, Scrollbar, grScrollInfoTrack(Scrollbar)
                'The function returns the tracking position of the scroll box in the nTrackPos member
                'of the SCROLLINFO structure.
                lValue = grScrollInfoTrack(Scrollbar).TrackPos
                
                'The event *suggests* a final position. This can be changed by
                'the client's event procedure (for example to force a start at
                'the beginning of rows/columns in grids etc.)
                If Scrollbar = egscSBOVertical Then
                    RaiseEvent VScroll(lValue)
                Else
                    RaiseEvent HScroll(lValue)
                End If
'To deactivate if not of use:
'lValue = .Pos
            Case egscSBNLineUp      'also egswSBNLineLeft
                'Set scroll value to value - SmallChange
                lChangeValue = glSmallChange(Scrollbar)
                'Events enabling the client to correct the suggested value
                '(User can change lChangeValue).
                If Scrollbar = egscSBOVertical Then
                    RaiseEvent LineUp(lChangeValue)
                Else
                    RaiseEvent PosLeft(lChangeValue)
                End If
                lValue = .Pos - lChangeValue    'Default is 1
                If lValue < .Min Then lValue = .Min
            Case egscSBNLineDown    'also egswSBNLineRight
                'Set scroll value to value + SmallChange
                lChangeValue = glSmallChange(Scrollbar)
                'Events enabling the client to correct the suggested value
                '(User can change lChangeValue).
                If Scrollbar = egscSBOVertical Then
                    RaiseEvent LineDown(lChangeValue)
                Else
                    RaiseEvent PosRight(lChangeValue)
                End If
                lValue = .Pos + lChangeValue    'Default is 1
                lEffMax = .Max - .Page + 1&
                If lValue > lEffMax Then lValue = lEffMax
            Case egscSBNPageUp      'also egswSBNPageLeft
                'Set scroll value to value - LargeChange
                lChangeValue = .Page
                'Events enabling the client to correct the suggested value
                '(User can change lChangeValue).
                If Scrollbar = egscSBOVertical Then
                    RaiseEvent PageUp(lChangeValue)
                Else
                    RaiseEvent PageLeft(lChangeValue)
                End If
                lValue = .Pos - lChangeValue
                If lValue < .Min Then lValue = .Min
            Case egscSBNPageDown    'also egswSBNPageRight
                'Set scroll value to value + LargeChange
                lChangeValue = .Page
                'Events enabling the client to correct the suggested value
                '(User can change lChangeValue).
                If Scrollbar = egscSBOVertical Then
                    RaiseEvent PageDown(lChangeValue)
                Else
                    RaiseEvent PageRight(lChangeValue)
                End If
                lValue = .Pos + lChangeValue
                lEffMax = .Max - .Page + 1&
                If lValue > lEffMax Then lValue = lEffMax
            Case egscSBNTop         'also egswSBNLeft
                'Set scroll value to min
                lValue = .Min
            Case egscSBNBottom      'also egswSBNRight
                'Set scroll value to max
                lValue = .Max
        End Select
        
        'Provide the new values for Windows (not for egswSBNEndScroll)
        If Notification <> egscSBNEndScroll Then
            .Pos = lValue
            grScrollInfo(Scrollbar).Mask = egscSIMAll
            SetScrollbar Scrollbar
        End If
        
        '"glLastEventValue" holds the last value for which a "Change" event was
        'raised. A new event is raised only when it differs from the last event.
        'If you don't want hot tracking, use "egswSBNEndScroll" to raise a "Change"
        'event and "egswSBNThumbTrack" to raise a "Scroll" event.
        If glLastEventValue(Scrollbar) <> .Pos Then
            RaiseEvent Change(Scrollbar, .Pos)
            glLastEventValue(Scrollbar) = .Pos
        End If
    End With
End Sub

'Extracting the High Word of a DWord.
Private Function GetHiWord(ByVal DWord As Long) As Long
    GetHiWord = (DWord And &HFFFF0000) \ &H10000
    If GetHiWord < 0& Then GetHiWord = GetHiWord + 65536
End Function

'Extracting the Low Word of a DWord.
Private Function GetLoWord(ByVal DWord As Long) As Long
    DWord = DWord And &HFFFF&
    If DWord > 32767 Then GetLoWord = DWord - 65536 Else GetLoWord = DWord
End Function


