VERSION 5.00
Begin VB.UserControl uToolbar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   Begin VB.Image imgMenu 
      Height          =   360
      Left            =   0
      Picture         =   "uToolbar.ctx":0000
      ToolTipText     =   "Menu"
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSysMenu 
      Height          =   285
      Left            =   4275
      Picture         =   "uToolbar.ctx":022A
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image dis 
      Height          =   360
      Index           =   2
      Left            =   900
      Picture         =   "uToolbar.ctx":0529
      Top             =   2700
      Width           =   360
   End
   Begin VB.Image dis 
      Height          =   360
      Index           =   4
      Left            =   1650
      Picture         =   "uToolbar.ctx":0733
      Top             =   2700
      Width           =   360
   End
   Begin VB.Image dis 
      Height          =   360
      Index           =   1
      Left            =   525
      Picture         =   "uToolbar.ctx":09A5
      Top             =   2700
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   10
      Left            =   3900
      Picture         =   "uToolbar.ctx":0B30
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   9
      Left            =   3525
      Picture         =   "uToolbar.ctx":0D72
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   8
      Left            =   3150
      Picture         =   "uToolbar.ctx":0F52
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   7
      Left            =   2775
      Picture         =   "uToolbar.ctx":115E
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   6
      Left            =   2400
      Picture         =   "uToolbar.ctx":13C7
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   5
      Left            =   2025
      Picture         =   "uToolbar.ctx":15EC
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   4
      Left            =   1650
      Picture         =   "uToolbar.ctx":1818
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   3
      Left            =   1275
      Picture         =   "uToolbar.ctx":1A71
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   2
      Left            =   900
      Picture         =   "uToolbar.ctx":1C9A
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image ena 
      Height          =   360
      Index           =   1
      Left            =   525
      Picture         =   "uToolbar.ctx":1EC2
      Top             =   2325
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   10
      Left            =   4230
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   9
      Left            =   3720
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   8
      Left            =   3270
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   7
      Left            =   2820
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   6
      Left            =   2370
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   5
      Left            =   1920
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   4
      Left            =   1470
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   3
      Left            =   1020
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   2
      Left            =   510
      Top             =   75
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   1
      Left            =   60
      Top             =   75
      Width           =   360
   End
   Begin VB.Image imgRight 
      Height          =   390
      Left            =   2025
      Picture         =   "uToolbar.ctx":20D2
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Image imgLeft 
      Height          =   390
      Left            =   75
      Picture         =   "uToolbar.ctx":278F
      Top             =   825
      Visible         =   0   'False
      Width           =   4410
   End
End
Attribute VB_Name = "uToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Toolbar_Button
    Key As String
    Enabled As String
    Left As Long
    Top As Long
End Type

Private mButt(1 To 10) As Toolbar_Button
Private mFullScreen As Boolean

Public Event ButtonClick(ByVal Key As String)
Public Event DblClick()
Public Event MenuClick(ByVal X As Long, ByVal Y As Long)
Public Event CloseClick()
Public Event RestoreClick()
Public Event MinimizeClick()

Public Sub SetFullScreen(ByVal Value As Boolean)
    Dim i As Integer
    
    If Not mFullScreen = Value Then
        mFullScreen = Value
        If Value Then
            For i = 1 To UBound(mButt, 1)
                mButt(i).Top = 1
                mButt(i).Left = mButt(i).Left + 30
            Next i
        Else
            For i = 1 To UBound(mButt, 1)
                mButt(i).Top = 5
                mButt(i).Left = mButt(i).Left - 30
            Next i
        End If
        
        For i = 1 To UBound(mButt, 1)
            img(i).Top = mButt(i).Top
            img(i).Left = mButt(i).Left
        Next i
        imgMenu.Visible = mFullScreen
        UserControl_Resize
    End If
End Sub

Public Sub SetToolTipText(ByVal Index As Integer, ByVal ttt As String)
    img(Index).ToolTipText = ttt
End Sub

Public Sub SetEnabled(ByVal Index As Integer, ByVal Value As Boolean)
    mButt(Index).Enabled = Value
    'change image
    img(Index).Enabled = Value
    If Value Then
        img(Index).Picture = ena(Index).Picture
    Else
        img(Index).Picture = dis(Index).Picture
    End If
End Sub

Private Sub img_Click(Index As Integer)
    RaiseEvent ButtonClick(mButt(Index).Key)
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    img(Index).Top = mButt(Index).Top + 1
    img(Index).Left = mButt(Index).Left + 1
End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    img(Index).Top = mButt(Index).Top
    img(Index).Left = mButt(Index).Left
End Sub

Private Sub imgMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MenuClick(imgMenu.Left, imgMenu.Top + imgMenu.Height)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    
    For i = 1 To UBound(mButt, 1)
        'mButt(i).Enabled = True
        SetEnabled i, True
        mButt(i).Left = img(i).Left
        mButt(i).Top = img(i).Top
    Next i
    
    mButt(1).Key = "cmdConnect"
    mButt(2).Key = "cmdClose"
    
    mButt(3).Key = "cmdSettings"
    mButt(4).Key = "cmdProfiles"
    mButt(5).Key = "cmdLog"
    mButt(6).Key = "cmdPalette"
    mButt(7).Key = "cmdRubrica"
    mButt(8).Key = "cmdButtons"
    mButt(9).Key = "cmdPlugins"
    
    mButt(10).Key = "cmdExit"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim w As Long
    
    If mFullScreen And Button = 1 Then
        w = UserControl.ScaleWidth
        If X > w - 19 Then
            RaiseEvent CloseClick
        ElseIf X > w - 38 And X < w - 19 Then
            RaiseEvent RestoreClick
        ElseIf X > w - 57 And X < w - 38 Then
            RaiseEvent MinimizeClick
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    Dim Back As Long
    
    Back = GetSysColor(COLOR_3DFACE)
    Dim Sfum As cSfum
    Set Sfum = New cSfum
        If mFullScreen Then
            Sfum.AggiungiColore 0, 0
        Else
            Sfum.AggiungiColore 16777215, 0
            Sfum.AggiungiColore Back, 5
            Sfum.AggiungiColore 0, 15
        End If
        Sfum.AggiungiColore 0, 95
        Sfum.AggiungiColore Back, 96
        Sfum.AggiungiColore 16777215, 100
        Sfum.StampaSfumatura 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1, UserControl.hdc
    Set Sfum = Nothing
    
    If mFullScreen Then
        UserControl.PaintPicture imgSysMenu.Picture, UserControl.ScaleWidth - imgSysMenu.Width, 3
    Else
        If imgLeft.Width + imgRight.Width < UserControl.ScaleWidth Then
            UserControl.PaintPicture imgRight.Picture, UserControl.ScaleWidth - imgRight.Width, 4
        End If
    End If
    
    If mFullScreen Then
        UserControl.PaintPicture imgLeft.Picture, 13 + 30, 0
    Else
        UserControl.PaintPicture imgLeft.Picture, 13, 4
    End If
    UserControl.Refresh
End Sub
