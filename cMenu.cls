VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mhPopupMenu As Long 'newly created popupmenu
Private mhMenu As Long      'existing window menu
Private mHWnd As Long       'window owner

Public Function CreatePopupMenu(ByVal hWnd As Long) As Boolean
    Dim Count As Integer, hSubMenu As Long, i As Integer
    Dim Buff As String, Length As Long
    
    If mHWnd = 0 Then
        mHWnd = hWnd
        If mhMenu = 0 Then mhMenu = GetMenu(mHWnd)
    
        mhPopupMenu = modUser32.CreatePopupMenu()
        Count = GetMenuItemCount(mhMenu)
        For i = 0 To Count - 1
            Buff = Space(128)
            Length = GetMenuString(mhMenu, i, Buff, 128, MF_BYPOSITION)
            Buff = Left$(Buff, Length)
            hSubMenu = GetSubMenu(mhMenu, i)
            AppendMenu mhPopupMenu, MF_STRING Or MF_ENABLED Or MF_POPUP, hSubMenu, Buff
        Next i
    Else
        CreatePopupMenu = False
    End If
End Function

Public Function RemoveMenu() As Boolean
    If mHWnd = 0 Then
        RemoveMenu = False
    Else
        SetMenu mHWnd, 0
        RemoveMenu = True
    End If
End Function

Public Function RestoreMenu() As Boolean
    If mhMenu = 0 Then
        RestoreMenu = False
    Else
        SetMenu mHWnd, mhMenu
        RestoreMenu = True
    End If
End Function

Public Function RefreshMenu() As Boolean
    Dim hWnd As Long
    
    If mhMenu = 0 Or mhPopupMenu = 0 Then
        RefreshMenu = False
    Else
        hWnd = mHWnd
        DestroyPopupMenu
        mHWnd = 0
        CreatePopupMenu hWnd
    End If
End Function

Public Function ShowPopupMenu(ByVal X As Long, ByVal Y As Long, Optional ByVal UseCursor As Boolean = False) As Boolean
    Dim pt As POINTAPI, fake As RECT
    
    If mhPopupMenu = 0 Then
        ShowPopupMenu = False
    Else
        If UseCursor Then
            GetCursorPos pt
            X = pt.X
            Y = pt.Y
        End If
        TrackPopupMenu mhPopupMenu, TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_LEFTBUTTON, X, Y, 0, mHWnd, fake
        
        ShowPopupMenu = True
    End If
End Function

Private Sub DestroyPopupMenu()
    Dim i As Integer
    'destroys the created popup menu
    If mhPopupMenu <> 0 Then
        For i = 0 To GetMenuItemCount(mhPopupMenu)
            modUser32.RemoveMenu mhPopupMenu, 0, MF_BYPOSITION
        Next i
        DestroyMenu mhPopupMenu
        mhPopupMenu = 0
    End If
End Sub

Private Sub Class_Initialize()
    mHWnd = 0
    mhMenu = 0
    mhPopupMenu = 0
End Sub

Private Sub Class_Terminate()
    DestroyPopupMenu
End Sub
