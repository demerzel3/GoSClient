VERSION 5.00
Begin VB.Form frmDefTrigger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define trigger"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "frmDefTrigger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   5775
      TabIndex        =   5
      Top             =   3900
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   3900
      Width           =   1365
   End
   Begin VB.TextBox txtReaction 
      Height          =   3165
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   675
      Width           =   8490
   End
   Begin VB.TextBox txtExp 
      Height          =   315
      Left            =   1125
      TabIndex        =   1
      Top             =   75
      Width           =   7440
   End
   Begin VB.Label lblReaction 
      Caption         =   "Reaction:"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   450
      Width           =   1740
   End
   Begin VB.Label lblExp 
      Caption         =   "Expression:"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   100
      Width           =   990
   End
End
Attribute VB_Name = "frmDefTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTrig As cTrigger

Private mSave As Boolean

Private mLang As cLang

Private Sub LoadLang()
    lblExp.Caption = mLang("deftrigger", "Expression")
    lblReaction.Caption = mLang("deftrigger", "Reaction")
    cmdOk.Caption = mLang("", "Ok")
    cmdCancel.Caption = mLang("", "Cancel")
End Sub

Public Function NewTrigger(ByRef Col As cTriggers) As Boolean
    Set mTrig = New cTrigger
    
    LoadData
    Me.Show vbModal
    
    If mSave Then
        SaveData
        Col.Add mTrig
    End If
    Set mTrig = Nothing
    Unload Me
    Set frmDefTrigger = Nothing
    
    NewTrigger = mSave
End Function

Public Function ModTrigger(ByRef Trig As cTrigger) As Boolean
    Set mTrig = Trig
    
    LoadData
    Me.Show vbModal
    
    If mSave Then SaveData
    Set mTrig = Nothing
    Unload Me
    Set frmDefTrigger = Nothing
    
    ModTrigger = mSave
End Function

Private Sub LoadData()
    txtExp.Text = mTrig.GetText
    txtReaction.Text = mTrig.Reaction
End Sub

Private Sub SaveData()
    mTrig.SetText txtExp.Text
    mTrig.Reaction = txtReaction.Text
End Sub

Private Sub cmdCancel_Click()
    mSave = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    mSave = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
End Sub
