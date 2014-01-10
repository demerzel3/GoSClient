Attribute VB_Name = "modUser32"
Option Explicit

Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'////////// System Parameters Info /////////////
    Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
        'uAction
        Public Const SPI_GETWHEELSCROLLLINES = 104
        Public Const SPI_SETWHEELSCROLLLINES = 105


'////////// Menus /////////////
    Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
    Public Declare Function CreatePopupMenu Lib "user32" () As Long
    Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
    Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
    Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
    Public Const TPM_LEFTALIGN = &H0&
    Public Const TPM_TOPALIGN = &H0&
    Public Const TPM_LEFTBUTTON = &H0&
    Public Const TPM_RETURNCMD = &H100&
    Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
    Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Public Const MF_BYCOMMAND = &H0&
    Public Const MF_BYPOSITION = &H400&
    Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Public Const MF_STRING = &H0&
    Public Const MF_POPUP = &H10&
    Public Const MF_ENABLED = &H0&
    Public Const MF_SEPARATOR = &H800&
    Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long


'////////// Cursor /////////////
    Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare Function GetCursor Lib "user32" () As Long
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Const IDC_SIZEALL = 32646&
    Public Const IDC_SIZENESW = 32643&
    Public Const IDC_SIZENS = 32645&
    Public Const IDC_SIZENWSE = 32642&
    Public Const IDC_SIZEWE = 32644&


'////////// POINT /////////////
    Public Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


'////////// RECT /////////////
    Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
    Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Public Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Public Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long


'////////// Device Context /////////////
    Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetDCEx Lib "user32" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
        'fdwOptions
        Public Const DCX_WINDOW = &H1&
        Public Const DCX_INTERSECTRGN = &H80&
    
    Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
        'wFormat
        Public Const DT_SINGLELINE = &H20
        Public Const DT_LEFT = &H0
        Public Const DT_RIGHT = &H2
        Public Const DT_VCENTER = &H4
        Public Const DT_CENTER = &H1

    Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long


'////////// Mouse /////////////
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Public Declare Function GetCapture Lib "user32" () As Long
    Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function ReleaseCapture Lib "user32" () As Long
    Public Const WHEEL_DELTA = 120


'////////// Windows /////////////
    Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        'wMsg
        Public Const WM_LBUTTONDOWN = &H201
        Public Const WM_SYSCOMMAND = &H112
        Public Const WM_CLOSE = &H10
        'wParam
        Public Const MK_LBUTTON = &H1
        Public Const SC_CLOSE = &HF060
    
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
        'hWndInsertAfter
        Public Const HWND_TOPMOST = -1
        'wFlags
        Public Const SWP_NOSIZE = &H1
        Public Const SWP_NOMOVE = &H2
        Public Const SWP_FRAMECHANGED = &H20

    Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
        'nIndex
        Public Const GWL_WNDPROC = (-4)
        Public Const GWL_HWNDPARENT = (-8)
        Public Const GWL_EXSTYLE = (-20)
        Public Const GWL_STYLE = (-16)

    Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
        'wCmd
        Public Const GW_OWNER = 4

    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
        'nCmdShow
        Public Const SW_SHOW = 5
        Public Const SW_MAXIMIZE = 3
        Public Const SW_MINIMIZE = 6
        Public Const SW_RESTORE = 9
        Public Const SW_NORMAL = 1
        Public Const SW_HIDE = 0
    
    Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long


'////////// Window styles /////////////
    Public Const WS_POPUP = &H80000000
    Public Const WS_CHILD = &H40000000
    Public Const WS_SYSMENU = &H80000
    Public Const WS_THICKFRAME = &H40000
    Public Const WS_CAPTION = &HC00000
    Public Const WS_OVERLAPPED = &H0&
    Public Const WS_MINIMIZEBOX = &H20000
    Public Const WS_MAXIMIZEBOX = &H10000
    Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    Public Const WS_DLGFRAME = &H400000
    Public Const WS_SIZEBOX = WS_THICKFRAME
    Public Const WS_CLIPCHILDREN = &H2000000


'////////// Window extended styles /////////////
    Public Const WS_EX_APPWINDOW = &H40000


'////////// System /////////////
    Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As GetSysColor_Colors) As Long
        'nIndex
        Public Enum GetSysColor_Colors
             COLOR_3DDKSHADOW = 21
             COLOR_3DFACE = 15
             COLOR_3DHIGHLIGHT = 20
             COLOR_3DHILIGHT = 20
             COLOR_3DLIGHT = 22
             COLOR_3DSHADOW = 16
             COLOR_ACTIVEBORDER = 10
             COLOR_ACTIVECAPTION = 2
             COLOR_APPWORKSPACE = 12
             COLOR_BACKGROUND = 1
             COLOR_BTNFACE = 15
             COLOR_BTNHIGHLIGHT = 20
             COLOR_BTNHILIGHT = 20
             COLOR_BTNSHADOW = 16
             COLOR_BTNTEXT = 18
             COLOR_CAPTIONTEXT = 9
             COLOR_DESKTOP = 1
             COLOR_GRADIENTACTIVECAPTION = 27
             COLOR_GRADIENTINACTIVECAPTION = 28
             COLOR_GRAYTEXT = 17
             COLOR_HIGHLIGHT = 13
             COLOR_HIGHLIGHTTEXT = 14
             COLOR_HOTLIGHT = 26
             COLOR_INACTIVEBORDER = 11
             COLOR_INACTIVECAPTION = 3
             COLOR_INACTIVECAPTIONTEXT = 19
             COLOR_INFOBK = 24
             COLOR_INFOTEXT = 23
             COLOR_MENU = 4
             COLOR_MENUTEXT = 7
             COLOR_SCROLLBAR = 0
             COLOR_WINDOW = 5
             COLOR_WINDOWFRAME = 6
             COLOR_WINDOWTEXT = 8
        End Enum
    
    Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
        'nIndex
        Public Const SM_CYCAPTION = 4
        Public Const SM_CYFRAME = 33
        Public Const SM_CXFRAME = 32
        Public Const SM_CXDLGFRAME = 7
        Public Const SM_CYDLGFRAME = 8
        Public Const SM_CXSIZEFRAME = SM_CXFRAME
        Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
        Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
        Public Const SM_CYSIZEFRAME = SM_CYFRAME

