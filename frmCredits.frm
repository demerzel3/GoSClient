VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About GoSClient"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgEaster 
      Height          =   840
      Left            =   2775
      Top             =   525
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   500
      X2              =   0
      Y1              =   233
      Y2              =   233
   End
   Begin VB.Image imgSeph 
      Height          =   750
      Left            =   0
      MouseIcon       =   "frmCredits.frx":00D2
      MousePointer    =   99  'Custom
      Picture         =   "frmCredits.frx":099C
      Top             =   3525
      Width           =   7500
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   450
      TabIndex        =   1
      Top             =   1575
      Width           =   6540
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version x.y.z"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   1200
      Width           =   2790
   End
   Begin VB.Image imgLogo 
      Height          =   1470
      Left            =   0
      Picture         =   "frmCredits.frx":1FF3
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCredits.Caption = _
    "freeware, entirely free and freely distribuible." & vbCrLf & _
    "completely developed by Seph. © 2002-2003, all rights reserved." & vbCrLf & _
    "this software is given with no warranty of any kind, I'm not responsible of any damage it's use could produce." & vbCrLf & vbCrLf & _
    "special thanks to Esteban (Silmaril) and to people that helped me during the beta-testing" & vbCrLf & vbCrLf & _
    "for any suggestion, info, bug reporting, criticism, ecc..." & vbCrLf & _
    "write at my e-mail address: gosclient@yahoo.it or click the image below"
End Sub

Private Sub imgEaster_DblClick()
    Dim DoEaster As Boolean
    Dim Item As Form
    
    DoEaster = False
    For Each Item In Forms
        If TypeName(Item) = "frmMain" Then
            DoEaster = True
            Exit For
        End If
    Next
    
    If DoEaster Then
        Unload Me
        Set frmCredits = Nothing
        
        frmEaster.Show vbModal
    End If
End Sub

Private Sub imgSeph_Click()
    ShellExecute Me.hWnd, vbNullString, "mailto:gosclient@yahoo.it", vbNullString, vbNullString, vbNormalFocus
End Sub
