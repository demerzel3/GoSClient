Attribute VB_Name = "modGosUtils"
Option Explicit

Public Declare Function gosuMakeColorTag Lib "gosu.dll" (ByVal lpBuff As String, ByVal crColor As Long, ByVal bBack As Long) As Long
'gosuMakeColorBackString(LPTSTR buff, COLORREF crColor, COLORREF crBack, LPCTSTR s, int nsLen)
Public Declare Function gosuMakeColorBackString Lib "gosu.dll" (ByVal lpBuff As String, ByVal crColor As Long, ByVal crBack As Long, ByVal sString As String, ByVal nLen As Long) As Long
Public Declare Function gosuMakeColorBackTag Lib "gosu.dll" (ByVal lpBuff As String, ByVal crColor As Long, ByVal crBack As Long) As Long

