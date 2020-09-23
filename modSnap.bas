Attribute VB_Name = "modSnap"
' modSnap.bas: the wonderful snaping module
' it slices! it dices! it makes your coffee!
' created by Imran Hameed (Xaimus)
'   user: xaimus AT host: [opposite of cold]pop DOT com
'   http://xaimus.vze.com/
' FEATURES:
'   Portable and modularish
'   Multi-Monitor Compatible
'   Non-Flickery
'   Not dependant on nonstandard external libs
' REQUIRES:
'   Win98+
' TODO:
'   allow forms in the same project to snap to each other
' USAGE:
'   subclass your form
'   create two long variables IN THE FORM, named x and y (or something like that)
'   in the form's wndproc, subclass the messages WM_ENTERSIZEMOVE, WM_MOVING, and WM_SIZING.
'   where WM_ENTERSIZEMOVE is, use PreSnapProc(your form's hwnd, the form variable x, the form variable y)
'   where WM_MOVING and WM_SIZING is, use SnapProc(your form's hwnd, the message, the wParam, the lParam, the form variable x, the form variable y)
'   if SnapProc returns '1', have your WndProc(or whatever it's called) return '1' and DO NOT CALL THE ORIGINAL WNDPROC!
'   change the SnapDistance constant to change the... Snap Distance! ... duh :)
' there ya' go. (almost) instant form snapping.
' EXAMPLE:
'   Private Function IMsgTarget_OnMsg(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'       Dim CallOrigWndProc As Integer
'       CallOrigWndProc = 0
'       Select Case msg
'           Case WM_ENTERSIZEMOVE
'               PreSnapProc hWnd, cx, cy
'           Case WM_MOVING
'               CallOrigWndProc = SnapProc(hWnd, msg, wParam, lParam, cx, cy)
'               If CallOrigWndProc = 1 Then IMsgTarget_OnMsg = CallOrigWndProc
'           Case WM_SIZING
'               CallOrigWndProc = SnapProc(hWnd, msg, wParam, lParam, cx, cy)
'               If CallOrigWndProc = 1 Then IMsgTarget_OnMsg = CallOrigWndProc
'       End Select
'       If CallOrigWndProc = 0 Then IMsgTarget_OnMsg = MsgBlaster.CallOrigWndProc(hWnd, msg, wParam, lParam)
'   End Function
' LICENSING:
'   this code is licensed under the GNU LGPL. (you should have recieved a lgpl.txt)
'   feel free to use this in your own programs (open-source and propriatary), but you must
'    distribute the source of this module with your package... and feel free to mess with
'    the source and distribute that.
'   just give me credit where due.
'   Copyright 2003 Imran Hameed.
Option Explicit

Private Const SnapDistance = 10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const MONITOR_DEFAULTTONEAREST = 2
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_MOVING = &H216
Private Const WM_SIZING = &H214
Private Const WMSZ_BOTTOM = 6
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_BOTTOMRIGHT = 8
Private Const WMSZ_LEFT = 1
Private Const WMSZ_RIGHT = 2
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_TOPRIGHT = 5

Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Public Sub PreSnapProc(ByVal hWnd As Long, x As Long, y As Long)
    Dim r As RECT, p As POINTAPI
    GetWindowRect hWnd, r                   'get the current window pos
    GetCursorPos p                          'get the current mouse pos
    
    x = r.Left - p.x                        'get the cursor pos in relation to
    y = r.Top - p.y                         ' the upper-left corner of the window
End Sub

Public Function SnapProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal cx As Long, ByVal cy As Long) As Integer
    Dim r As RECT
    Dim hMonitor As Long
    Dim cm As MONITORINFO
    Dim p As POINTAPI
    
    'why can't VB just include good pointer support? ;)
    CopyMemory r, ByVal lParam, Len(r)

    'get the current monitor
    cm.cbSize = Len(cm)
    hMonitor = MonitorFromRect(r, MONITOR_DEFAULTTONEAREST)
    If hMonitor <> 0 Then                   'make sure we have an active monitor
        GetMonitorInfo hMonitor, cm         'we do have an active monitor
    Else
        SnapProc = 0                        'we don't-let Windows take care of it
        Exit Function
    End If
    
    Select Case msg
        Case WM_MOVING
            'get the width and length (used later)
            Dim w As Long, h As Long
            w = r.Right - r.Left
            h = r.Bottom - r.Top
            
            'the user may have moved the mouse when snapped...
            GetCursorPos p
            r.Left = p.x + cx
            r.Top = p.y + cy
            r.Right = r.Left + w
            r.Bottom = r.Top + h
        
            If r.Left <= (cm.rcWork.Left + SnapDistance) And r.Left >= (cm.rcWork.Left - SnapDistance) Then
                r.Left = cm.rcWork.Left
                r.Right = w + r.Left
            End If
            If r.Top <= (cm.rcWork.Top + SnapDistance) And r.Top >= (cm.rcWork.Top - SnapDistance) Then
                r.Top = cm.rcWork.Top
                r.Bottom = h + r.Top
            End If
            If r.Right <= (cm.rcWork.Right + SnapDistance) And r.Right >= (cm.rcWork.Right - SnapDistance) Then
                r.Right = cm.rcWork.Right
                r.Left = r.Right - w
            End If
            If r.Bottom <= (cm.rcWork.Bottom + SnapDistance) And r.Bottom >= (cm.rcWork.Bottom - SnapDistance) Then
                r.Bottom = cm.rcWork.Bottom
                r.Top = r.Bottom - h
            End If
        Case WM_SIZING
            GetCursorPos p
            Select Case wParam
                Case WMSZ_LEFT
                    r.Left = p.x
                    If r.Left <= (cm.rcWork.Left + SnapDistance) And r.Left >= (cm.rcWork.Left - SnapDistance) Then r.Left = cm.rcWork.Left
                Case WMSZ_TOP
                    r.Top = p.y
                    If r.Top <= (cm.rcWork.Top + SnapDistance) And r.Top >= (cm.rcWork.Top - SnapDistance) Then r.Top = cm.rcWork.Top
                Case WMSZ_RIGHT
                    r.Right = p.x
                    If r.Right <= (cm.rcWork.Right + SnapDistance) And r.Right >= (cm.rcWork.Right - SnapDistance) Then r.Right = cm.rcWork.Right
                Case WMSZ_BOTTOM
                    r.Bottom = p.y
                    If r.Bottom <= (cm.rcWork.Bottom + SnapDistance) And r.Bottom >= (cm.rcWork.Bottom - SnapDistance) Then r.Bottom = cm.rcWork.Bottom
                Case WMSZ_TOPLEFT
                    r.Left = p.x
                    r.Top = p.y
                    If r.Top <= (cm.rcWork.Top + SnapDistance) And r.Top >= (cm.rcWork.Top - SnapDistance) Then r.Top = cm.rcWork.Top
                    If r.Left <= (cm.rcWork.Left + SnapDistance) And r.Left >= (cm.rcWork.Left - SnapDistance) Then r.Left = cm.rcWork.Left
                Case WMSZ_TOPRIGHT
                    r.Right = p.x
                    r.Top = p.y
                    If r.Top <= (cm.rcWork.Top + SnapDistance) And r.Top >= (cm.rcWork.Top - SnapDistance) Then r.Top = cm.rcWork.Top
                    If r.Right <= (cm.rcWork.Right + SnapDistance) And r.Right >= (cm.rcWork.Right - SnapDistance) Then r.Right = cm.rcWork.Right
                Case WMSZ_BOTTOMLEFT
                    r.Left = p.x
                    r.Bottom = p.y
                    If r.Bottom <= (cm.rcWork.Bottom + SnapDistance) And r.Bottom >= (cm.rcWork.Bottom - SnapDistance) Then r.Bottom = cm.rcWork.Bottom
                    If r.Left <= (cm.rcWork.Left + SnapDistance) And r.Left >= (cm.rcWork.Left - SnapDistance) Then r.Left = cm.rcWork.Left
                Case WMSZ_BOTTOMRIGHT
                    r.Right = p.x
                    r.Bottom = p.y
                    If r.Bottom <= (cm.rcWork.Bottom + SnapDistance) And r.Bottom >= (cm.rcWork.Bottom - SnapDistance) Then r.Bottom = cm.rcWork.Bottom
                    If r.Right <= (cm.rcWork.Right + SnapDistance) And r.Right >= (cm.rcWork.Right - SnapDistance) Then r.Right = cm.rcWork.Right
            End Select
    End Select
    
    'if only VB didn't suck with pointers... :D
    CopyMemory ByVal lParam, r, Len(r)
    SnapProc = 1                    'move sucessful-the programmer should appropriately take care of
    Exit Function                   ' the WndProc by returning '1' and ignoring the original WndProc.
End Function
