Attribute VB_Name = "modGosHook"
Option Explicit

Public Declare Sub goshStartHook Lib "gosh.dll" (ByVal hParent As Long, ByVal hDest As Long, ByVal hMouseDest As Long, ByVal hThread As Long)
Public Declare Sub goshStopHook Lib "gosh.dll" ()
Public Declare Function goshGetLastWnd Lib "gosh.dll" () As Long

Public Declare Sub goshSetDockable Lib "gosh.dll" (ByVal hWnd As Long, ByVal winID As String)
Public Declare Sub goshSetUndockable Lib "gosh.dll" (ByVal hWnd As Long)
Public Declare Function goshCheckDockable Lib "gosh.dll" (ByVal hWnd As Long) As Boolean
Public Declare Function goshCheckDocked Lib "gosh.dll" (ByVal hWnd As Long) As Boolean

Public Declare Sub goshSetDockingRects Lib "gosh.dll" (ByVal hDockParent As Long, rc As RECT, ByVal Count As Integer)

Public Declare Function goshStartMove Lib "gosh.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub goshMoveToLastRect Lib "gosh.dll" (ByVal hWnd As Long)

Public Declare Function goshSetOwner Lib "gosh.dll" (ByVal hWnd As Long, ByVal hWndOwner As Long) As Long

Public Declare Function goshGetWindowID_ Lib "gosh.dll" Alias "goshGetWindowID" (ByVal hWnd As Long, ByVal sBuff As String, ByVal nLen As Long) As Long

Public Declare Sub goshStopLastMouse Lib "gosh.dll" ()

Public Declare Function goshSetMainWnd Lib "gosh.dll" (ByVal hWnd As Long) As Long
Public Declare Sub goshSetDocked Lib "gosh.dll" (ByVal hWnd As Long)

Public Declare Function goshFindWindow Lib "gosh.dll" (ByVal sWinID As String) As Long

Public Declare Sub goshSetLocked Lib "gosh.dll" (ByVal bValue As Boolean)

Public Function goshGetWindowID(ByVal hWnd As Long) As String
    Dim sBuff As String, lenght As Long
    
    sBuff = Space(128)
    lenght = goshGetWindowID_(hWnd, sBuff, lenght)
    goshGetWindowID = Left$(sBuff, lenght)
End Function
