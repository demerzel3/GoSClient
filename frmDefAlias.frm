VERSION 5.00
Begin VB.Form frmDefAlias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definisci"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDefAlias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDown 
      Height          =   315
      Left            =   4500
      Picture         =   "frmDefAlias.frx":00D2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2100
      Width           =   315
   End
   Begin VB.CommandButton cmdUp 
      Height          =   315
      Left            =   4500
      Picture         =   "frmDefAlias.frx":016C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1500
      Width           =   315
   End
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "Chiudi"
      Default         =   -1  'True
      Height          =   315
      Left            =   5625
      TabIndex        =   8
      Top             =   3975
      Width           =   1140
   End
   Begin VB.CommandButton cmdCombo 
      Caption         =   "Combinazione di tasti"
      Height          =   315
      Left            =   75
      TabIndex        =   7
      Top             =   3975
      Width           =   2340
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   6
      Top             =   1275
      Width           =   1665
   End
   Begin VB.CommandButton cmdModifica 
      Caption         =   "Modifica"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   5
      Top             =   900
      Width           =   1665
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   315
      Left            =   5100
      TabIndex        =   4
      Top             =   525
      Width           =   1665
   End
   Begin VB.ListBox lstAzioni 
      Height          =   3165
      IntegralHeight  =   0   'False
      Left            =   75
      TabIndex        =   3
      Top             =   750
      Width           =   4215
   End
   Begin VB.TextBox txtTesto 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   75
      Width           =   5865
   End
   Begin VB.Label lblMove 
      Alignment       =   2  'Center
      Caption         =   "Move"
      Height          =   240
      Left            =   4350
      TabIndex        =   11
      Top             =   1875
      Width           =   615
   End
   Begin VB.Label lblAzioni 
      Caption         =   "0 Azioni"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   525
      Width           =   4215
   End
   Begin VB.Label lblTesto 
      Caption         =   "Testo:"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   765
   End
End
Attribute VB_Name = "frmDefAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAlias As cAlias
Private mAliasID As Integer

Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private Sub LoadLang()
    If Not mAlias Is Nothing Then Me.Caption = mLang("", "Define") & " " & mAlias.Mode
    lblTesto.Caption = mLang("defalias", "Text")
    cmdCombo.Caption = mLang("combo", "Caption") 'keys combination
    cmdNuovo.Caption = mLang("", "New")
    cmdModifica.Caption = mLang("", "Modify")
    cmdElimina.Caption = mLang("", "Delete")
    cmdChiudi.Caption = mLang("", "Close")
    lblMove.Caption = mLang("defalias", "Move")
End Sub

Public Sub Init(ByRef Aliases As cAlias, AliasID As Integer)
    Set mAlias = Aliases
    mAliasID = AliasID
    
    Me.Caption = mLang("", "Define") & " " & mAlias.Mode
    txtTesto.Text = mAlias.Text(AliasID)
    LoadAction
    
    Me.Show vbModal
End Sub

Private Sub LoadAction()
    Dim i As Integer

    lstAzioni.Clear
    For i = 1 To mAlias.ActionCount(mAliasID)
        lstAzioni.AddItem mAlias.ActionPar(mAliasID, i)
    Next i
    
    lblAzioni.Caption = lstAzioni.ListCount & " " & mLang("defalias", "Actions")
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
    Set frmDefAlias = Nothing
End Sub

Private Sub cmdCombo_Click()
    Dim Combo As String
    
    Load frmCombo
    'frmCombo.Init mAlias, mAliasID
    Combo = mAlias.Combo(mAliasID)
    If frmCombo.GetCombo(Combo) Then
        mAlias.Combo(mAliasID) = Combo
    End If
End Sub

Private Sub cmdDown_Click()
    Dim Index As Integer
    
    If Not lstAzioni.ListIndex = -1 Then
        Index = lstAzioni.ListIndex
        If mAlias.MoveActionDown(mAliasID, Index + 1) Then
            lstAzioni.AddItem lstAzioni.list(Index), Index + 2
            lstAzioni.RemoveItem lstAzioni.ListIndex
            lstAzioni.ListIndex = Index + 1
        End If
    End If
End Sub

Private Sub cmdElimina_Click()
    mAlias.RemoveAction mAliasID, lstAzioni.ListIndex + 1
    lstAzioni.RemoveItem lstAzioni.ListIndex
    lblAzioni.Caption = lstAzioni.ListCount & " " & mLang("defalias", "Actions")

    cmdElimina.Enabled = False
    cmdModifica.Enabled = False
End Sub

Private Sub Modifica()
    Dim Index As Integer
    
    Index = lstAzioni.ListIndex + 1
    If Index <> 0 Then
        mAlias.Text(mAliasID) = txtTesto.Text
        
        Load frmDefAzione
        frmDefAzione.Init mAlias, mAliasID, Index
        Unload frmDefAzione
        Set frmDefAzione = Nothing
        
        lstAzioni.list(lstAzioni.ListIndex) = mAlias.ActionPar(mAliasID, Index)
    End If
End Sub

Private Sub cmdModifica_Click()
    Modifica
End Sub

Private Sub cmdNuovo_Click()
    Dim id As Integer

    id = mAlias.AddAction(mAliasID)
    mAlias.ActionPar(mAliasID, id) = " %a"
    lstAzioni.AddItem ""
    lstAzioni.ListIndex = lstAzioni.ListCount - 1
    lblAzioni.Caption = lstAzioni.ListCount & " " & mLang("defalias", "Actions")
    
    Modifica
End Sub

Private Sub cmdUp_Click()
    Dim Index As Integer
    
    If Not lstAzioni.ListIndex = -1 Then
        Index = lstAzioni.ListIndex
        If mAlias.MoveActionUp(mAliasID, Index + 1) Then
            lstAzioni.AddItem lstAzioni.list(Index), Index - 1
            lstAzioni.RemoveItem lstAzioni.ListIndex
            lstAzioni.ListIndex = Index - 1
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mAlias.Text(mAliasID) = txtTesto.Text
    
    Set mAlias = Nothing

    Set mLang = Nothing
End Sub

Private Sub lstAzioni_Click()
    If lstAzioni.ListIndex <> -1 Then
        If Not cmdModifica.Enabled Then
            cmdModifica.Enabled = True
            cmdElimina.Enabled = True
        End If
    Else
        If cmdModifica.Enabled Then
            cmdModifica.Enabled = False
            cmdElimina.Enabled = False
        End If
    End If
End Sub

Private Sub lstAzioni_DblClick()
    Modifica
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub txtTesto_GotFocus()
    txtTesto.SelStart = 0
    txtTesto.SelLength = Len(txtTesto.Text)
End Sub
