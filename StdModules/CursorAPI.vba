' Uncomment this and last line to disable CursorAPI (Office for Mac compatibility)
'#If False Then
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' ==============================================================
'
' ##############################################################
' #                                                            #
' #                    PPTGames Cursor API                     #
' #                      CursorAPI Module                      #
' #                                                            #
' ##############################################################
'
' » Version 3.0.0
'
' » https://pptgamespt.wixsite.com/pptg-coding/cursor-api
'
' ===============================================================
' Modified version for codename OrangePath/OS


'Option Explicit

Private Const LOGPIXELSX As Long = 88
Private mPoint As POINTAPI

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As LongPtr
    Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hwndFrom As LongPtr, ByVal hwndTo As LongPtr, lppt As Any, ByVal cPoints As Long) As Long
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As LongPtr
#Else
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
#End If

Private Function GetDpi() As Long
    #If VBA7 Then
        Dim hdcScreen As LongPtr
    #Else
        Dim hdcScreen As Long
    #End If
    Dim iDPI As Long
    iDPI = -1
    hdcScreen = GetDC(FindWindow("screenClass", vbNullString))
    If (hdcScreen) Then
        iDPI = GetDeviceCaps(hdcScreen, LOGPIXELSX)
        ReleaseDC 0, hdcScreen
    End If
    GetDpi = iDPI
End Function

Public Sub SetCursorPosition(x As Single, y As Single, Optional AsPoints As Boolean = True)
    If AsPoints Then
        SetCursorPos x * GetDpi / 72 * ActivePresentation.SlideShowWindow.View.Zoom / 100, y * GetDpi / 72 * ActivePresentation.SlideShowWindow.View.Zoom / 100
    Else
        SetCursorPos x, y
    End If
End Sub

Public Function GetCursorXRaw(Optional Map As Boolean = True) As Long
    Dim p As POINTAPI
    GetCursorPos p
    If Map Then p = MapPoint(p)
    GetCursorXRaw = p.x
End Function

Public Function GetCursorYRaw(Optional Map As Boolean = True) As Long
    Dim p As POINTAPI
    GetCursorPos p
    If Map Then p = MapPoint(p)
    GetCursorYRaw = p.y
End Function

Public Function GetCursorX() As Single
    Dim p As POINTAPI
    #If VBA7 Then
        Dim mWnd As LongPtr
    #Else
        Dim mWnd As Long
    #End If
    Dim sx As Long, sy As Long
    Dim dx As Double, dy As Double
    Dim WR As RECT
    GetCursorPos p
    p = MapPoint(p)
    mWnd = FindWindow("screenClass", vbNullString)
    GetWindowRect mWnd, WR
    sx = WR.lLeft
    sy = WR.lTop
    With ActivePresentation.PageSetup
        dx = (WR.lRight - WR.lLeft) / .SlideWidth
        dy = (WR.lBottom - WR.lTop) / .SlideHeight
        Select Case True
        Case dx > dy
            sx = sx + (dx - dy) * .SlideWidth / 2
            dx = dy
        Case dy > dx
            sy = sy + (dy - dx) * .SlideHeight / 2
            dy = dx
        End Select
    End With
    GetCursorPos p
    GetCursorX = (p.x - sx) / dx
End Function

Public Function GetCursorY() As Single
    Dim p As POINTAPI
    #If VBA7 Then
        Dim mWnd As LongPtr
    #Else
        Dim mWnd As Long
    #End If
    Dim sx As Long, sy As Long
    Dim dx As Double, dy As Double
    Dim WR As RECT
    GetCursorPos p
    p = MapPoint(p)
    mWnd = FindWindow("screenClass", vbNullString)
    GetWindowRect mWnd, WR
    sx = WR.lLeft
    sy = WR.lTop
    With ActivePresentation.PageSetup
        dx = (WR.lRight - WR.lLeft) / .SlideWidth
        dy = (WR.lBottom - WR.lTop) / .SlideHeight
        Select Case True
        Case dx > dy
            sx = sx + (dx - dy) * .SlideWidth / 2
            dx = dy
        Case dy > dx
            sy = sy + (dy - dx) * .SlideHeight / 2
            dy = dx
        End Select
    End With
    GetCursorPos p
    GetCursorY = (p.y - sy) / dy
End Function

Private Function MapPoint(p As POINTAPI) As POINTAPI
    Dim points(0) As POINTAPI
    points(0) = p
    MapWindowPoints GetDesktopWindow, FindWindow("screenClass", vbNullString), points(0), 1
    MapPoint = points(0)
End Function


'#End If