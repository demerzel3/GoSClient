VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Splash Screen"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ForeColor       =   &H00000000&
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   4200
      Top             =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "http://members.xoom.it/gosclient/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2925
      TabIndex        =   2
      Top             =   4050
      Width           =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   77
      X2              =   390
      Y1              =   175
      Y2              =   175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "gosclient@yahoo.it"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   150
      TabIndex        =   1
      Top             =   4050
      Width           =   1440
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version x.y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4650
      TabIndex        =   0
      Top             =   2625
      Width           =   1140
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Dim Bordi As cGrafica

    lblVersion.Caption = "version " & App.Major & "." & App.Minor

    'Set Bordi = New cGrafica
    '    Bordi.DisegnaBordi Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, _
    '        0, 2, rgb(140, 140, 140), 0, 30, , False, 16777215
    'Set Bordi = Nothing

    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Me.Show
End Sub

Private Sub Timer_Timer()
    Unload Me
    Set frmSplash = Nothing
End Sub
