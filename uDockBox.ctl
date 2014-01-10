VERSION 5.00
Begin VB.UserControl uDockBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox pctMenu 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   30
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   3825
      Begin VB.Image imgMenu 
         Height          =   120
         Left            =   30
         Picture         =   "uDockBox.ctx":0000
         Top             =   30
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image imgClose 
         Height          =   120
         Left            =   3225
         Picture         =   "uDockBox.ctx":0092
         Top             =   30
         Width           =   120
      End
   End
   Begin VB.PictureBox pctCont 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   3015
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   306
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   4590
   End
End
Attribute VB_Name = "uDockBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mhDocked As Long        'hwnd of the docked window
Private mStatic As Boolean      'a static box hasn't a title bar
Private mLocked As Boolean      'a locked box can't be dragged and hasn't the close button

Private mMenu As cMenu

Public Event Click()
Public Event WindowClosed()
Public Event WindowUndocked()
Public Event StartingMove()

Public Sub SetLocked(ByVal Value As Boolean)
    If mLocked <> Value Then
        mLocked = Value
        imgClose.Visible = Not mLocked
        pctMenu_Resize
    End If
End Sub

Public Function IsHWndValid() As Boolean
    If mhDocked <> 0 Then
        IsHWndValid = IsWindow(mhDocked)
    End If
End Function

Public Function GetCaption() As String
    Dim Buff As String, Length As Long
    
    If Not mhDocked = 0 Then
        Buff = Space(128)
        Length = GetWindowText(mhDocked, Buff, 128)
        Buff = Left$(Buff, Length)
        GetCaption = Buff
    End If
End Function

Public Sub SetStatic(data As Boolean)
    mStatic = data
    imgMenu.Visible = (Not data)
    pctMenu.Visible = False
    pctCont.Top = 0
End Sub

Public Function GetStatic() As Boolean
    GetStatic = mStatic
End Function

Public Property Get DockedHWnd() As Long
    DockedHWnd = mhDocked
End Property

Public Sub ResizeForm()
    Dim rcCont As RECT

'    On Error GoTo ErrorOccured
    
    If Not mhDocked = 0 Then
        pctCont.Visible = False
        Call GetClientRect(pctCont.hWnd, rcCont)
        MoveWindow mhDocked, 0, 0, _
            rcCont.Right, rcCont.Bottom, True
        pctCont.Visible = True
    End If
End Sub

Public Sub Dock(ByVal hWnd As Long)
    Dim Style As Long, Caption As String, lenght As Long
    
    mhDocked = hWnd
    If Not mhDocked = 0 Then
        pctCont.Visible = False
        
        If GetMenu(mhDocked) Then
            Set mMenu = New cMenu
            mMenu.CreatePopupMenu mhDocked
            mMenu.RemoveMenu
            imgMenu.Visible = True
        End If
        
        Style = GetWindowLong(mhDocked, GWL_STYLE)
        Style = Style And (Not WS_CAPTION)
        Style = Style And (Not WS_SIZEBOX)
        Style = Style Or WS_CHILD
        SetWindowLong mhDocked, GWL_STYLE, Style
        SetParent mhDocked, pctCont.hWnd
        ShowWindow mhDocked, SW_MAXIMIZE
        pctCont.Visible = True
        ResizeForm
        
        pctMenu.CurrentX = 3
        pctMenu.CurrentY = 1
        'Caption = mDockForm.Caption
        Caption = Space(128)
        lenght = GetWindowText(mhDocked, Caption, 128)
        Caption = Left$(Caption, lenght)
        
        If InStr(1, Caption, "(") <> 0 Then Caption = Left$(Caption, InStr(1, Caption, "(") - 1)
        If InStr(1, Caption, "[") <> 0 Then Caption = Left$(Caption, InStr(1, Caption, "[") - 1)
        pctMenu.Print UCase$(Caption)
        
        UserControl.Tag = UCase$(Caption)
        
        pctMenu_Resize
    End If
End Sub

Public Sub CloseForm()
    Dim hWnd As Long
    
    If Not mhDocked = 0 Then
        'SendMessage mhDocked, WM_SYSCOMMAND, SC_CLOSE, 0
        'ShowWindow mhDocked, SW_HIDE
        
        hWnd = mhDocked
        UnDock SW_SHOW
        
        modUser32.SetFocus hWnd
        DoEvents
        
        SendMessage hWnd, WM_CLOSE, 0, 0
        
        'If Not IsWindow(hWnd) Then
        '    mhDocked = 0
        '    RaiseEvent WindowClosed
        'Else
        '    MsgBox "Closing failed for some reason"
        'End If
    End If
End Sub

Private Sub RestoreMenu()
    If Not mMenu Is Nothing Then
        mMenu.RestoreMenu
        imgMenu.Visible = False
        Set mMenu = Nothing
    End If
End Sub

Public Sub UnDock(Optional nCmdShow As Long = 0)
    Dim Style As Long

    If Not mhDocked = 0 Then
        Style = GetWindowLong(mhDocked, GWL_STYLE)
        Style = Style Or WS_CAPTION
        Style = Style Or WS_SIZEBOX
        Style = Style And (Not WS_CHILD)
        SetWindowLong mhDocked, GWL_STYLE, Style
        SetParent mhDocked, 0
        SetWindowLong mhDocked, GWL_HWNDPARENT, frmBase.hWnd
        RestoreMenu
        ShowWindow mhDocked, nCmdShow
        
        mhDocked = 0
        RaiseEvent WindowUndocked
    End If
End Sub

Public Property Let BackColor(data As Long)
    UserControl.BackColor = data
End Property

Public Property Get ScaleWidth() As Long
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScaleHeight() As Long
    ScaleHeight = UserControl.ScaleHeight
End Property

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mLocked Then CloseForm
End Sub

Private Sub imgMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mMenu Is Nothing Then mMenu.ShowPopupMenu 0, 0, True
End Sub

Private Sub pctCont_Click()
    RaiseEvent Click
End Sub

Private Sub pctMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mhDocked <> 0 And Not mStatic And Not mLocked Then
        If goshStartMove(mhDocked, X, Y) Then
            RaiseEvent StartingMove
            Debug.Print "starting move"
        End If
    End If
End Sub

Private Sub pctMenu_Resize()
    Dim Sfum As cSfum, Bordi As cGrafica
    Dim dark As Long, light As Long
    Dim pWidth As Long, pLeft As Long
    
    If pctMenu.Visible Or UserControl.Extender.Visible = False Then
        Set Sfum = New cSfum
            dark = VariaColore(pctMenu.BackColor, -20)
            light = VariaColore(dark, 30)
            Sfum.AggiungiColore dark, 0
            Sfum.AggiungiColore light, 100
            Sfum.StampaSfumatura 0, 0, pctMenu.ScaleWidth, pctMenu.ScaleHeight, 0, pctMenu.hdc
        Set Sfum = Nothing
        
        pWidth = pctMenu.ScaleWidth
        If Not mLocked Then pWidth = pWidth - 11
        
        Set Bordi = New cGrafica
            If Not mMenu Is Nothing Then
                pLeft = 10
            End If
        
            Bordi.DisegnaBordi pctMenu.hdc, pLeft, 0, pWidth, pctMenu.ScaleHeight, _
                1, 1, 16777215, 0, 50, , , dark
            
            If Not mLocked Then
                Bordi.DisegnaBordi pctMenu.hdc, pctMenu.ScaleWidth - 11, 0, 11, pctMenu.ScaleHeight, _
                    0, 1, 16777215, 0, 50, , , light
            End If
            
            If Not mMenu Is Nothing Then
                Bordi.DisegnaBordi pctMenu.hdc, 0, 0, 11, pctMenu.ScaleHeight, _
                    0, 1, 16777215, 0, 50, , , light
            End If
        Set Bordi = Nothing
        
        pctMenu.CurrentX = 3
        If Not mMenu Is Nothing Then pctMenu.CurrentX = 14
        pctMenu.CurrentY = 1
        pctMenu.Print UserControl.Tag
        
        'pctMenu.Refresh
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Dim Back As Long

    Back = GetSysColor(COLOR_3DFACE)
    UserControl.BackColor = Back
    pctCont.BackColor = Back
    pctMenu.BackColor = Back
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    pctMenu.Width = Me.ScaleWidth - pctMenu.Left * 2
    imgClose.Left = pctMenu.ScaleWidth - imgClose.Width - 1
    pctCont.Width = Me.ScaleWidth
    pctCont.Height = Me.ScaleHeight - pctCont.Top
    If mhDocked <> 0 And Ambient.UserMode Then
        ResizeForm
    End If
End Sub
