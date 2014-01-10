VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug monitor"
   ClientHeight    =   1965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10050
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9315
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
      Begin VB.Menu mnuOptLog 
         Caption         =   "Log on file"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Log(Msg As String)
    Msg = Replace(Msg, Chr$(0), " ")
    lstLog.AddItem Format(Time, "hh:mm:ss") & "| " & Msg
    lstLog.ListIndex = lstLog.ListCount - 1
End Sub

Private Sub Form_Load()
    goshSetOwner Me.hWnd, frmBase.hWnd
    Me.Show
End Sub

Private Sub Form_Resize()
    lstLog.Width = Me.ScaleWidth
    lstLog.Height = Me.ScaleHeight
End Sub

