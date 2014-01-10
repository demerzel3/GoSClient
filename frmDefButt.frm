VERSION 5.00
Begin VB.Form frmDefButt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definisci pulsante"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmDefButt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optOption 
      Caption         =   "Doppio"
      Height          =   240
      Left            =   2700
      TabIndex        =   14
      Top             =   375
      Width           =   1890
   End
   Begin VB.OptionButton optNormal 
      Caption         =   "Normale"
      Height          =   240
      Left            =   375
      TabIndex        =   13
      Top             =   375
      Width           =   1890
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   315
      Left            =   4125
      TabIndex        =   8
      Top             =   2325
      Width           =   1290
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2625
      TabIndex        =   7
      Top             =   2325
      Width           =   1440
   End
   Begin VB.CommandButton cmdPresAct 
      Caption         =   "..."
      Height          =   315
      Left            =   4950
      TabIndex        =   6
      Top             =   1950
      Width           =   465
   End
   Begin VB.CommandButton cmdNormAct 
      Caption         =   "..."
      Height          =   315
      Left            =   4950
      TabIndex        =   5
      Top             =   1200
      Width           =   465
   End
   Begin VB.Frame fraButType 
      Caption         =   "Tipo di pulsante"
      Height          =   615
      Left            =   75
      TabIndex        =   4
      Top             =   75
      Width           =   5340
   End
   Begin VB.TextBox txtPresAct 
      Height          =   315
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1950
      Width           =   3315
   End
   Begin VB.TextBox txtPresText 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   1575
      Width           =   3840
   End
   Begin VB.TextBox txtNormAct 
      Height          =   315
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3315
   End
   Begin VB.TextBox txtNormText 
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Top             =   825
      Width           =   3840
   End
   Begin VB.Label lblAliasPres 
      Caption         =   "Alias (premuto):"
      Height          =   240
      Left            =   75
      TabIndex        =   12
      Top             =   1950
      Width           =   1440
   End
   Begin VB.Label lblTextPres 
      Caption         =   "Testo (premuto):"
      Height          =   240
      Left            =   75
      TabIndex        =   11
      Top             =   1575
      Width           =   1440
   End
   Begin VB.Label lblAliasNorm 
      Caption         =   "Alias (normale):"
      Height          =   240
      Left            =   75
      TabIndex        =   10
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label lblTextNorm 
      Caption         =   "Testo (normale):"
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   825
      Width           =   1440
   End
End
Attribute VB_Name = "frmDefButt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAlias As cAlias

Private mButtID As Integer
Private mButt As cButtons

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("defbutt", "caption")
    
    fraButType.Caption = mLang("defbutt", "ButtonType")
    optNormal.Caption = mLang("defbutt", "Normal")
    optOption.Caption = mLang("defbutt", "Double")
    
    lblTextNorm.Caption = mLang("defbutt", "TextNormal")
    lblTextPres.Caption = mLang("defbutt", "TextPressed")
    lblAliasNorm.Caption = mLang("defbutt", "AliasNormal")
    lblAliasPres.Caption = mLang("defbutt", "AliasPressed")
    
    cmdOk.Caption = mLang("", "Ok")
    cmdAnnulla.Caption = mLang("", "Cancel")
End Sub

Public Sub Init(Index As Integer, Buttons As cButtons)
    mButtID = Index
    Set mButt = Buttons

    With mButt
        txtNormText.Text = .NormText(mButtID)
        'mNormAct = .NormAct(mButtID)
        txtNormAct.Text = .NormAct(mButtID)
        txtPresText.Text = .PresText(mButtID)
        'mPresAct = .PresAct(mButtID)
        txtPresAct.Text = .PresAct(mButtID)
        If .OptionButt(mButtID) Then
            optOption.Value = True
        Else
            optNormal.Value = True
        End If
    End With

    Me.Show vbModal
End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
    Set frmDefButt = Nothing
End Sub

Private Sub cmdNormAct_Click()
    Dim rtn As Integer, Name As String

    LoadAliasList
    rtn = frmList.ShowForm(Name)
    If Not rtn = -1 Then
        rtn = rtn + 1
        'mNormAct = rtn
        txtNormAct.Text = Name
    End If
    Unload frmList
    Set frmList = Nothing

    txtNormText.SetFocus
End Sub

Private Sub cmdOk_Click()
    With mButt
        .NormText(mButtID) = txtNormText.Text
        .NormAct(mButtID) = txtNormAct.Text
        .PresText(mButtID) = txtPresText.Text
        .PresAct(mButtID) = txtPresAct.Text
        .OptionButt(mButtID) = optOption.Value
    End With
    Unload Me
    Set frmDefButt = Nothing
End Sub

Private Sub cmdPresAct_Click()
    Dim rtn As Integer, Name As String
    
    LoadAliasList
    rtn = frmList.ShowForm(Name)
    If Not rtn = -1 Then
        rtn = rtn + 1
        'mPresAct = rtn
        txtPresAct.Text = Name
    End If
    Unload frmList
    Set frmList = Nothing

    txtPresText.SetFocus
End Sub

Private Sub LoadAliasList()
    Dim i As Integer

    Load frmList
    frmList.Caption = mLang("defbutt", "SelectAlias")
    For i = 1 To mAlias.Count
        frmList.AddItem mAlias.Text(i), mAlias.Text(i)
    Next i
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    txtNormAct.BackColor = Me.BackColor
    txtPresAct.BackColor = Me.BackColor
    
    Set mAlias = New cAlias
        mAlias.LoadAliases
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmList
    Set frmList = Nothing
    
    Set mButt = Nothing
    Set mAlias = Nothing
    
    Set mLang = Nothing
End Sub

Private Sub optNormal_Click()
    txtPresText.Enabled = False
    txtPresAct.Enabled = False
    txtPresText.BackColor = Me.BackColor
    cmdPresAct.Enabled = False

    If Me.Visible Then txtNormText.SetFocus
End Sub

Private Sub optOption_Click()
    txtPresText.Enabled = True
    txtPresAct.Enabled = True
    txtPresText.BackColor = txtNormText.BackColor
    cmdPresAct.Enabled = True
    
    If Me.Visible Then txtPresText.SetFocus
End Sub

Private Sub txtNormText_GotFocus()
    txtNormText.SelStart = 0
    txtNormText.SelLength = Len(txtNormText.Text)
End Sub

Private Sub txtPresText_GotFocus()
    txtPresText.SelStart = 0
    txtPresText.SelLength = Len(txtPresText.Text)
End Sub

