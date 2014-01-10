VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mHDC As Long
Private mOriginalBmp As Long
Private mWidth As Long
Private mHeight As Long

Private mhFont As Long
Private mOriginalFont As Long
Private mhBrush As Long
Private mOriginalBrush As Long
Private mhPen As Long
Private mOriginalPen As Long

Public Sub SetSolidPen(Color As Long)
    If mhPen <> 0 Then DeleteObject mhPen
    mhPen = CreatePen(PS_SOLID, 1, Color)
    If mOriginalPen = 0 Then
        mOriginalPen = SelectObject(mHDC, mhPen)
    Else
        Call SelectObject(mHDC, mhPen)
    End If
End Sub

Public Sub SetBrush(Color As Long, Optional Style As Integer = BS_SOLID)
    Dim lb As LOGBRUSH

    If mhBrush <> 0 Then DeleteObject (mhBrush)
    lb.lbColor = Color
    lb.lbHatch = 0
    lb.lbStyle = Style
    mhBrush = CreateBrushIndirect(lb)
    If mOriginalBrush = 0 Then
        mOriginalBrush = SelectObject(mHDC, mhBrush)
    Else
        Call SelectObject(mHDC, mhBrush)
    End If
End Sub

Public Sub Cls(Bitmap As cBitmap, Optional Color As Long = 0)
    Dim hBrush As Long
    Dim Area As RECT

    hBrush = CreateSolidBrush(Color)
    Area.Left = 0
    Area.Top = 0
    Area.Right = Bitmap.Width
    Area.Bottom = Bitmap.Height
    FillRect mHDC, Area, hBrush
    DeleteObject hBrush
End Sub

Public Sub SetBmp(Bitmap As cBitmap)
    If Not mHDC = 0 Then
        If mOriginalBmp = 0 Then
            mOriginalBmp = SelectObject(mHDC, Bitmap.hBmp)
        Else
            Call SelectObject(mHDC, Bitmap.hBmp)
        End If
        mWidth = Bitmap.Width
        mHeight = Bitmap.Height
    End If
End Sub

Public Sub FreeBmp()
    If mHDC <> 0 And mOriginalBmp <> 0 Then SelectObject mHDC, mOriginalBmp
End Sub

Public Sub CreateLight()
    Dim DeskDC As Long

    If mHDC <> 0 Then Destroy
    
    DeskDC = GetDC(0)
    mHDC = CreateCompatibleDC(DeskDC)
    ReleaseDC 0, DeskDC
    
    SetBkMode mHDC, TRANSPARENT
End Sub

Public Sub Create()
    Dim DeskDC As Long
    Dim brush As LOGBRUSH
    Dim rtn As Long

    If mHDC <> 0 Then Destroy
    DeskDC = GetDC(0)
    mHDC = CreateCompatibleDC(DeskDC)
    rtn = ReleaseDC(0, DeskDC)
    
    brush.lbColor = 0
    brush.lbStyle = 1  'nullo, trasparente
    mhBrush = CreateBrushIndirect(brush)
    mOriginalBrush = SelectObject(mHDC, mhBrush)
    
    'rtn = DeleteObject(rtn)
    
    mhFont = CreateFont(13, 0, 0, 0, 100, 0, 0, 0, 0, 0, 0, 0, 64, "Tahoma")
    mOriginalFont = SelectObject(mHDC, mhFont)
    
    'rtn = DeleteObject(rtn)
    
    SetBkMode mHDC, TRANSPARENT

    'Debug.Print "new device context = " & mHDC
    If mHDC = 0 Then
        MsgBox "Failed creating the device context" & vbCrLf & _
               "Restart the client and, if necessary, the computer"
    End If
End Sub

Public Sub Destroy()
    Dim rtn As Integer

    If mHDC <> 0 Then
        If mOriginalBmp <> 0 Then Me.FreeBmp
        If mOriginalBrush <> 0 Then SelectObject mHDC, mOriginalBrush
        If mOriginalFont <> 0 Then SelectObject mHDC, mOriginalFont
        If mOriginalPen <> 0 Then SelectObject mHDC, mOriginalPen
        Call DeleteObject(mhFont)
        Call DeleteObject(mhBrush)
        Call DeleteObject(mhPen)
        
        rtn = DeleteDC(mHDC)
        mHDC = 0
        mOriginalBmp = 0
        mOriginalBrush = 0
        mOriginalFont = 0
    End If
End Sub

Public Property Get hdc() As Long
    hdc = mHDC
End Property

Private Sub Class_Terminate()
    If mHDC <> 0 Then Destroy
End Sub
