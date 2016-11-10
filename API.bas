Attribute VB_Name = "API"
Option Explicit

Public Type Point_API

    X As Long
    Y As Long

End Type

Public Declare Function QueryPerformanceCounter Lib "Kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "Kernel32" (lpPerformanceCount As Currency) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (Position As Point_API) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, Position As Point_API) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
