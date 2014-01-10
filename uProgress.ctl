VERSION 5.00
Begin VB.UserControl uProgress 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox pctProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   15
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   291
      TabIndex        =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   4365
   End
End
Attribute VB_Name = "uProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mValue As Integer

Private mTrueValue As Integer
Private mTrueTraguardo As Integer

Private mColor As Long

Public Property Let Color(data As Long)
    mColor = data
    DrawSfum
    Redraw
End Property

Private Sub DrawSfum()
    Dim Sfum As cSfum

    Set Sfum = New cSfum
    Sfum.AggiungiColore mColor, 0
    Sfum.AggiungiColore VariaColore(mColor, -50), 100
    Sfum.StampaSfumatura 0, 0, pctProgress.ScaleWidth, pctProgress.ScaleHeight, , pctProgress.hdc
    Set Sfum = Nothing
End Sub

Public Property Get Value() As Integer
    Value = mValue
End Property

Public Property Let Value(data As Integer)
    mValue = data
    
    GetTrueValue
    
    Redraw
End Property

Private Sub UserControl_Initialize()
    pctProgress.BackColor = GetSysColor(COLOR_3DFACE)
    UserControl.BackColor = GetSysColor(COLOR_3DFACE)
    mColor = 255 'rosso, colore predefinito
End Sub

Private Sub Redraw()
    Dim Sfondi As cGrafica
    Dim Fake As POINTAPI
    'Dim Connect As cConnector

    Set Sfondi = New cGrafica
        Sfondi.FillRectEx UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.BackColor
        'Set Connect = New cConnector
            Sfondi.FillRectEx UserControl.hdc, pctProgress.Left - 1, pctProgress.Top - 1, pctProgress.Width + 2, pctProgress.Height + 2, rgb(240, 240, 240)
            'If Not mValue = mTraguardo Then Sfondi.FillRectEx UserControl.hdc, pctProgress.Left, pctProgress.Top, CLng(mTrueTraguardo), pctProgress.Height, Connect.RetrInfo("Progress_ProgColor", 16777215)
        'Set Connect = Nothing
        Sfondi.DisegnaBordi UserControl.hdc, pctProgress.Left - 1, pctProgress.Top - 1, pctProgress.ScaleWidth + 2, pctProgress.ScaleHeight + 2, 1, 1, 0, 0, 20
    Set Sfondi = Nothing
    
    'If mValue <> mTraguardo Then
    '    Timer.Enabled = True
    'End If
    
    BitBlt UserControl.hdc, pctProgress.Left, pctProgress.Top, mTrueValue, pctProgress.Height, pctProgress.hdc, 0, 0, SRCCOPY

    UserControl.Refresh
End Sub

Private Sub UserControl_Paint()
    Redraw
End Sub

Private Sub UserControl_Resize()
    pctProgress.Width = UserControl.ScaleWidth - pctProgress.Left * 2
    pctProgress.Top = (UserControl.ScaleHeight - pctProgress.Height) / 2

    GetTrueValue
    DrawSfum
    Redraw
End Sub

Private Sub GetTrueValue()
    mTrueValue = (pctProgress.ScaleWidth * mValue) / 100
End Sub
