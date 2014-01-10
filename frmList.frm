VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   2775
      Width           =   1815
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Default         =   -1  'True
      Height          =   315
      Left            =   2625
      TabIndex        =   1
      Top             =   2775
      Width           =   1815
   End
   Begin VB.ListBox lst 
      Height          =   2595
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4365
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mIndex As Integer
Private mKey As String

Private mKeys As Collection

Private Sub cmdAnnulla_Click()
    mIndex = -1
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    mIndex = lst.ListIndex
    mKey = mKeys(lst.ListIndex + 1)
    Me.Hide
End Sub

Private Sub Form_Load()
    Set mKeys = New Collection
End Sub

Public Sub AddItem(ByVal Caption As String, Optional ByVal Key As String = "")
    lst.AddItem Caption
    mKeys.Add Key
End Sub

Public Function ShowForm(Optional ByRef Key As String) As Integer
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        cmdOk.Caption = Connect.Lang("", "Ok")
        cmdAnnulla.Caption = Connect.Lang("", "Cancel")
    Set Connect = Nothing

    Me.Show vbModal
    ShowForm = mIndex
    Key = mKey
    Unload Me
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mKeys = Nothing
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        mIndex = -1
        Me.Hide
    End If
End Sub

Private Sub lst_Click()
    cmdOk.Enabled = True
End Sub

Private Sub lst_DblClick()
    mIndex = lst.ListIndex
    mKey = mKeys(lst.ListIndex + 1)
    Me.Hide
End Sub

