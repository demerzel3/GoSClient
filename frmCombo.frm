VERSION 5.00
Begin VB.Form frmCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combinazione di tasti"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmCombo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   450
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   450
      Width           =   1290
   End
   Begin VB.CommandButton cmdTogli 
      Caption         =   "Togli"
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   1215
   End
   Begin VB.TextBox txtCombo 
      Height          =   315
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   4515
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mCombo As cKeyCombo
Attribute mCombo.VB_VarHelpID = -1
'Private mAlias As cAlias
'Private mAliasID As Integer
Private mKeyCombo As String
Private mSave As Boolean

Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private Sub LoadLang()
    Me.Caption = mLang("combo", "Caption")
    cmdTogli.Caption = mLang("", "Reset")
    cmdOk.Caption = mLang("", "Ok")
    cmdCancel.Caption = mLang("", "Cancel")
End Sub

Public Function GetCombo(ByRef NewCombo As String) As Boolean
    Dim Shift As Integer, Key As Integer
    
    Set mCombo = New cKeyCombo
    mCombo.ScindiCombo NewCombo, Key, Shift
    mCombo.AvviaRiconoscimento txtCombo
    mCombo.IdentCombo Key, Shift
    mKeyCombo = NewCombo

    Me.Show vbModal
    
    If mSave Then
        NewCombo = mKeyCombo
    End If
    GetCombo = mSave
    
    Unload Me
    Set frmCombo = Nothing
End Function

Private Sub cmdChiudi_Click()
    Unload Me
    Set frmCombo = Nothing
End Sub

Private Sub cmdCancel_Click()
    mSave = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    mSave = True
    Me.Hide
End Sub

Private Sub cmdTogli_Click()
    mCombo.IdentCombo 0, 0
    mKeyCombo = "0|0"
    'mAlias.Combo(mAliasID) = "0|0"
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    
    LoadLang
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mCombo = Nothing
    
    Set mLang = Nothing
End Sub

Private Sub mCombo_SetCombo(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode, Shift
    'mAlias.Combo(mAliasID) = KeyCode & "|" & Shift
    mKeyCombo = KeyCode & "|" & Shift
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub
