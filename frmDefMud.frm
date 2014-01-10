VERSION 5.00
Begin VB.Form frmDefMud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proprietà"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmDefMud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Annulla"
      Height          =   315
      Left            =   5175
      TabIndex        =   13
      Top             =   4050
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   315
      Left            =   3825
      TabIndex        =   12
      Top             =   4050
      Width           =   1290
   End
   Begin VB.TextBox txtComment 
      Height          =   1140
      Left            =   1275
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2775
      Width           =   5265
   End
   Begin VB.TextBox txtDescr 
      Height          =   1140
      Left            =   1275
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1575
      Width           =   5265
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   1275
      TabIndex        =   8
      Top             =   825
      Width           =   1365
   End
   Begin VB.TextBox txtHost 
      Height          =   315
      Left            =   1275
      TabIndex        =   7
      Top             =   450
      Width           =   5265
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1275
      TabIndex        =   6
      Top             =   75
      Width           =   5265
   End
   Begin VB.Line Line1 
      X1              =   6525
      X2              =   75
      Y1              =   3975
      Y2              =   3975
   End
   Begin VB.Label lblComment 
      Caption         =   "Commento:"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   2775
      Width           =   1665
   End
   Begin VB.Label lblDescr 
      Caption         =   "Descrizione:"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   1595
      Width           =   1665
   End
   Begin VB.Label lblLang 
      Caption         =   "Lingua:"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   1220
      Width           =   1665
   End
   Begin VB.Label lblPort 
      Caption         =   "Porta:"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   845
      Width           =   1665
   End
   Begin VB.Label lblHost 
      Caption         =   "Host:"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   470
      Width           =   1665
   End
   Begin VB.Label lblName 
      Caption         =   "Nome:"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   95
      Width           =   1665
   End
End
Attribute VB_Name = "frmDefMud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSave As Boolean

Private mProp As String
Private mPropOf As String

Private Sub LoadLang(Lang As cLang)
    lblName.Caption = Lang("", "Name") & ":"
    lblHost.Caption = Lang("", "Host") & ":"
    lblPort.Caption = Lang("", "Port") & ":"
    lblLang.Caption = Lang("", "Language") & ":"
    lblDescr.Caption = Lang("", "Description") & ":"
    lblComment.Caption = Lang("", "Comment") & ":"
    
    mProp = Lang("defmud", "Properties")
    mPropOf = Lang("defmud", "PropertiesOf")
    If txtName.Text = "" Then
        Me.Caption = mProp
    Else
        Me.Caption = mPropOf & " " & txtName.Text
    End If
    
    cmdOk.Caption = Lang("", "Ok")
    cmdAbort.Caption = Lang("", "Cancel")
End Sub

Public Function NewMud(Muds As cMuds, Lang As cLang) As Boolean
    Dim lPort As Long
    
    'loading current language settings
    LoadLang Lang
    
    Me.Show vbModal

    If mSave Then
        lPort = Abs(Val(txtPort.Text))
        If lPort > 65535 Then lPort = 65535
        Muds.Add txtName.Text, txtHost.Text, lPort, txtDescr.Text, cboLang.Text, txtComment.Text
        NewMud = True
    End If

    Unload Me
    Set frmDefMud = Nothing
End Function

Public Function EditMud(Mud As cMud, Lang As cLang) As Boolean
    Dim i As Integer
    
    'loading current language settings
    LoadLang Lang
    
    With Mud
        txtName.Text = .Name
        txtName.Enabled = False
        txtHost.Text = .Host
        For i = 0 To cboLang.ListCount - 1
            If LCase$(cboLang.list(i)) = LCase$(.Lang) Then
                cboLang.ListIndex = i
                Exit For
            End If
        Next i
        txtComment.Text = .Comment
        txtDescr.Text = .Descr
        txtPort.Text = Trim$(CStr(.Port))
    End With
    
    Me.Show vbModal
    
    If mSave Then
        SaveMudInfo Mud
        EditMud = True
    End If
    
    Unload Me
    Set frmDefMud = Nothing
End Function

Private Sub SaveMudInfo(Dest As cMud)
    Dim Port As Long
    
    With Dest
        .Name = txtName.Text
        .Host = txtHost.Text
        .Lang = cboLang.Text
        .Comment = txtComment.Text
        .Descr = txtDescr.Text
        Port = Abs(Val(txtPort.Text))
        If Port > 65535 Then Port = 65535
        .Port = Port
    End With
End Sub

Private Sub cmdAbort_Click()
    mSave = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    mSave = True
    Me.Hide
End Sub

Private Sub Form_Load()
    cboLang.AddItem "Italian"
    cboLang.AddItem "English"
End Sub

Private Sub txtName_Change()
    If txtName.Text = "" Then
        Me.Caption = mProp
    Else
        Me.Caption = mPropOf & " " & txtName.Text
    End If
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtHost_GotFocus()
    txtHost.SelStart = 0
    txtHost.SelLength = Len(txtHost.Text)
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End Sub
