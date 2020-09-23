VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gifSubClassing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FILE INFO
'---------
'Control:       gifSubClassing
'IPN prefix:    gsc
'Purpose:       Interface implementation for subclassing Windows messages to report to the owner of
'               gclSubClassing.

'Implementor:   herbert.glarner@bluewin.ch
'Orig. Author:  Paul_Caton@hotmail.com
'Copyright:     None


'USAGE
'-----
'Use "Implements gifSubClassing" in the controls using this interface.


'MULTIPLE SUBCLASSERS
'--------------------
'No issues, but realise that all gclSubClassing instances will callback through the same interfaces.
'If common messages are used among them then use the hWnd parameter to distinguish the source.


'PUBLIC INTERFACE METHODS
'------------------------

'"lReturn": Will contain the return value from the previous WndProc, populate it yourself as the
'           message/situation/Windows dictates.
'"hWnd":    Handle of the window
'"uMsg":    Message number
'"wParam":  Message related data
'"lParam":  Message related data
Public Sub After _
    (ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
End Sub


'"lHandled":If lHandled is set to a non-zero value then the message will not be passed on to previous
'           WndProc after this sub quits.
'"lReturn": Set the value of lReturn as the message/situation/Windows dictates.
'"hWnd":    Handle of the window
'"uMsg":    Message number
'"wParam":  Message related data
'"lParam":  Message related data
Public Sub Before _
    (ByRef lHandled As Long, _
     ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
End Sub
