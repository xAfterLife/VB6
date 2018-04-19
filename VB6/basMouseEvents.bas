Attribute VB_Name = "basMouseEvents"
Option Explicit
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MOVE = &H1

Public Enum MEvent
    Up = 0
    Down = 1
    Click = 2
End Enum

Public Enum MoveStyle
    Direkt = 0
    relCurrent = 1
    relForm = 2
End Enum

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Function GetMouseX() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetMouseX = n.x
End Function

Public Function GetMouseY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetMouseY = n.y
End Function

Public Function MLeft(ByVal MouseE As MEvent, ByVal Style As MoveStyle, ByVal mx As Long, ByVal my As Long, ByVal sleeptime As Integer)
    MMove Style, mx, my
    If MouseE = Down Or Click Then
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    End If
    Sleep sleeptime
    If MouseE = Up Or Click Then
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    End If
End Function

Public Function MRight(ByVal MouseE As MEvent, ByVal Style As MoveStyle, ByVal mx As Long, ByVal my As Long, ByVal sleeptime As Integer)
    MMove Style, mx, my
    If MouseE = Down Or Click Then
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    End If
    Sleep sleeptime
    If MouseE = Up Or Click Then
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    End If
End Function

Public Function MMid(ByVal MouseE As MEvent, ByVal Style As MoveStyle, ByVal mx As Long, ByVal my As Long, ByVal sleeptime As Integer)
    MMove Style, mx, my
    If MouseE = Down Or Click Then
        mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
    End If
    Sleep sleeptime
    If MouseE = Up Or Click Then
        mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
    End If
End Function

Public Function MMove(ByVal Style As MoveStyle, ByVal mx As Long, ByVal my As Long, Optional ByVal RelativForm As Form)
    Dim CurrX As Integer
    Dim CurrY As Integer
    
    If Style = relCurrent Then
    CurrX = GetMouseX
    CurrY = GetMouseY
    mx = mx + CurrX
    my = my + CurrY
    End If
    
    If Style = relForm Then
    CurrX = RelativForm.Left
    CurrY = RelativForm.Top
    mx = mx + CurrX
    my = my + CurrY
    End If
    
    SetCursorPos mx, my

End Function

