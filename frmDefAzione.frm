VERSION 5.00
Begin VB.Form frmDefAzione 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definisci azione"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frmDefAzione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Automatico"
      Height          =   240
      Left            =   2925
      TabIndex        =   6
      Top             =   1125
      Width           =   1290
   End
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "Chiudi"
      Default         =   -1  'True
      Height          =   315
      Left            =   5100
      TabIndex        =   5
      Top             =   1800
      Width           =   1365
   End
   Begin VB.TextBox txtPar 
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   1425
      Width           =   6240
   End
   Begin VB.TextBox txtWPar 
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   750
      Width           =   6240
   End
   Begin VB.Label lblWithPar 
      Caption         =   "Con parametri"
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   1125
      Width           =   1740
   End
   Begin VB.Label lblWithoutPar 
      Caption         =   "Senza parametri"
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   450
      Width           =   1740
   End
   Begin VB.Label lblTesto 
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   6390
   End
End
Attribute VB_Name = "frmDefAzione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAlias As cAlias
Private mAliasID As Integer
Private mActionID As Integer

Private mCopy As Boolean

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("defaction", "caption")
    lblWithoutPar.Caption = mLang("defaction", "WithoutPar")
    lblWithPar.Caption = mLang("defaction", "WithPar")
    cmdAuto.Caption = mLang("defaction", "Auto")
    cmdChiudi.Caption = mLang("", "Close")
End Sub

Public Sub Init(ByRef Aliases As cAlias, AliasID As Integer, ActionID As Integer)
    Set mAlias = Aliases
    mAliasID = AliasID
    mActionID = ActionID
    
    txtWPar.Text = mAlias.Action(mAliasID, mActionID)
    txtPar.Text = mAlias.ActionPar(mAliasID, mActionID)
    lblTesto.Caption = mLang("defalias", "Text") & " " & mAlias.Text(mAliasID)
    
    Select Case mAlias.Mode
        Case "aliases"
            If txtWPar.Text & " %a" = txtPar.Text Then
                cmdAuto.Value = True
            End If
        Case "triggers"
            mCopy = True
            cmdAuto.Enabled = False
    End Select
    
    CheckAuto
    
    Me.Show vbModal
End Sub

Private Sub CheckAuto()
    If Not mCopy Then
        txtPar.BackColor = txtWPar.BackColor
        txtPar.Enabled = True
    Else
        txtPar.BackColor = Me.BackColor
        txtPar.Enabled = False
        txtPar.Text = txtWPar.Text & " %a"
        If Me.Visible Then txtWPar.SetFocus
    End If
End Sub

Private Sub cmdAuto_Click()
    mCopy = Not mCopy
    If mCopy Then
        'cmdAuto.Caption = "Personalizza"
        cmdAuto.Caption = mLang("defaction", "Custom")
    Else
        'cmdAuto.Caption = "Automatico"
        cmdAuto.Caption = mLang("defaction", "Auto")
    End If
    CheckAuto
    If Not mCopy And Me.Visible Then txtPar.SetFocus
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
    Set frmDefAzione = Nothing
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mAlias.Action(mAliasID, mActionID) = txtWPar.Text
    mAlias.ActionPar(mAliasID, mActionID) = txtPar.Text
    
    Set mAlias = Nothing
    
    Set mLang = Nothing
End Sub

Private Sub txtWPar_Change()
    If mCopy Then txtPar.Text = txtWPar.Text & " %a"
End Sub

Private Sub txtWPar_GotFocus()
    txtWPar.SelStart = 0
    txtWPar.SelLength = Len(txtWPar.Text)
End Sub

Private Sub txtPar_GotFocus()
    txtPar.SelStart = 0
    txtPar.SelLength = Len(txtPar.Text)
End Sub
