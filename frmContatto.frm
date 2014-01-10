VERSION 5.00
Begin VB.Form frmContatto 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proprietà contatto"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmContatto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   315
      Left            =   5700
      TabIndex        =   11
      Top             =   3675
      Width           =   2115
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   3525
      TabIndex        =   10
      Top             =   3675
      Width           =   2115
   End
   Begin VB.TextBox txtNote 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1575
      Width           =   7740
   End
   Begin VB.TextBox txtURL 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      TabIndex        =   3
      Top             =   900
      Width           =   5265
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   525
      Width           =   5040
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4875
      TabIndex        =   1
      Top             =   150
      Width           =   2940
   End
   Begin VB.TextBox txtNick 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   675
      TabIndex        =   0
      Top             =   150
      Width           =   2940
   End
   Begin VB.Label lblNotes 
      Caption         =   "Note:"
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   1275
      Width           =   690
   End
   Begin VB.Label lblUrl 
      Caption         =   "Sito Web:"
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Top             =   900
      Width           =   1140
   End
   Begin VB.Label lblMail 
      Caption         =   "Indirizzo e-mail:"
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   525
      Width           =   1365
   End
   Begin VB.Label lblName 
      Caption         =   "Nome:"
      Height          =   240
      Left            =   4125
      TabIndex        =   6
      Top             =   150
      Width           =   690
   End
   Begin VB.Label lblNick 
      Caption         =   "Nick:"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "frmContatto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRubrica As cRubrica
Private mIndex As Integer
Private mAddNew As Boolean

Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private Sub LoadLang()
    Me.Caption = mLang("contact", "Caption")
    
    'labels
    lblNick.Caption = mLang("contact", "Nick")
    lblName.Caption = mLang("contact", "Name")
    lblMail.Caption = mLang("contact", "EMail")
    lblUrl.Caption = mLang("contact", "Url")
    lblNotes.Caption = mLang("contact", "Notes")
    
    'commands
    cmdOk.Caption = mLang("", "Ok")
    cmdAnnulla.Caption = mLang("", "Cancel")
End Sub

Public Sub ModifyContact(Index As Integer)
    Dim Connect As cConnector

    Set Connect = New cConnector
        Set mRubrica = Connect.Rubrica
    Set Connect = Nothing

    mIndex = Index
    txtNick.Enabled = False
    txtNick.BackColor = Me.BackColor
    txtNick.Text = mRubrica.Nick(Index)
    txtName.Text = mRubrica.Name(Index)
    txtEmail.Text = mRubrica.Email(Index)
    txtURL.Text = mRubrica.URL(Index)
    txtNote.Text = mRubrica.Note(Index)
    Me.Show vbModal
End Sub

Public Sub AddNewContact()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
        Set mRubrica = Connect.Rubrica
    Set Connect = Nothing
    mAddNew = True
    
    txtNick.Enabled = True
    txtNick.BackColor = txtName.BackColor
    txtNick.Text = ""
    txtName.Text = ""
    txtEmail.Text = ""
    txtURL.Text = ""
    txtNote.Text = ""
    Me.Show vbModal
End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
    Set frmContatto = Nothing
End Sub

Private Sub cmdOk_Click()
    If mAddNew Then
        If Len(Trim$(txtNick.Text)) = 0 Then
            'MsgBox "Inserisci un nickname oppure premi annulla per uscire"
            MsgBox mLang("contact", "ErrNoNick")
            txtNick.SetFocus
            Exit Sub
        Else
            mRubrica.Add txtNick.Text, txtName.Text, txtEmail.Text, _
                         txtURL.Text, txtNote.Text
        End If
    Else
        With mRubrica
            .Nick(mIndex) = txtNick.Text
            .Name(mIndex) = txtName.Text
            .Email(mIndex) = txtEmail.Text
            .URL(mIndex) = txtURL.Text
            .Note(mIndex) = txtNote.Text
        End With
    End If
    Unload Me
    Set frmContatto = Nothing
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mRubrica = Nothing
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub txtNick_GotFocus()
    txtNick.SelStart = 0
    txtNick.SelLength = Len(txtNick.Text)
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.SelStart = 0
    txtEmail.SelLength = Len(txtEmail.Text)
End Sub

Private Sub txtURL_GotFocus()
    txtURL.SelStart = 0
    txtURL.SelLength = Len(txtURL.Text)
End Sub
