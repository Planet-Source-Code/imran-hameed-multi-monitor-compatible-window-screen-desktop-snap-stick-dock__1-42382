Attribute VB_Name = "modMain"
' NOTE:
'   I've used VB's built-in subclassing [GHEY] in this
'    demo to save space. Get MsgBlaster or some other
'    decent subclassing control/library.

Option Explicit

Public lpPrevWndProc As Long
Public gHW As Long

Public cx As Long
Public cy As Long

Public Const GWL_WNDPROC = -4
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_MOVING = &H216
Public Const WM_SIZING = &H214

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim result As Integer
    Select Case uMsg
        Case WM_ENTERSIZEMOVE
            PreSnapProc hw, cx, cy
        Case WM_MOVING
            result = SnapProc(hw, uMsg, wParam, lParam, cx, cy)
        Case WM_SIZING
            result = SnapProc(hw, uMsg, wParam, lParam, cx, cy)
    End Select
    If result = 0 Then WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub
