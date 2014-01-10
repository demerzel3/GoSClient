VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mhBmp As Long
Private mWidth As Long
Private mHeight As Long

Public Sub CreateFromBitmap(Path As String, Width As Long, Height As Long)
    If mhBmp <> 0 Then Destroy
    mhBmp = LoadImage(ByVal 0&, Path, 0, Width, Height, &H10)
    mWidth = Width
    mHeight = Height
End Sub

Public Sub Create(Width As Long, Height As Long)
    Dim DeskDC As Long

    If mhBmp <> 0 Then Destroy
    DeskDC = GetDC(0)
    'mhDC = CreateCompatibleDC(deskDc)
    mhBmp = CreateCompatibleBitmap(DeskDC, Width, Height)
    Call ReleaseDC(0, DeskDC)
    mWidth = Width
    mHeight = Height
    'SelectObject mhDC, mhBmp
End Sub

Public Sub Destroy()
    Dim rtn As Integer

    If mhBmp <> 0 Then
        rtn = DeleteObject(mhBmp)
        'Open "bitmap.log" For Append As #21
        '    Print #21, "Return deleting bitmap = " & rtn
        'Close #21
        mhBmp = 0
    End If
End Sub

Private Sub Class_Terminate()
    If mhBmp <> 0 Then Destroy
End Sub

Public Property Get Height() As Long
    Height = mHeight
End Property

Public Property Get Width() As Long
    Width = mWidth
End Property

Public Property Get hBmp() As Long
    hBmp = mhBmp
End Property
