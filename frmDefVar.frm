VERSION 5.00
Begin VB.Form frmDefVar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define variable"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frmDefVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   975
      TabIndex        =   1
      Top             =   450
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   975
      TabIndex        =   0
      Top             =   75
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4350
      TabIndex        =   3
      Top             =   825
      Width           =   1440
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2850
      TabIndex        =   2
      Top             =   825
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "@"
      Height          =   240
      Left            =   750
      TabIndex        =   6
      Top             =   95
      Width           =   240
   End
   Begin VB.Label lblValue 
      Caption         =   "Value"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   470
      Width           =   1065
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   95
      Width           =   1065
   End
End
Attribute VB_Name = "frmDefVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLang As cLang

Private mVars As cVars
Private mSave As Boolean

Public Function NewVariable(ByRef Name As String, ByRef Value As String, ByRef Vars As cVars) As Boolean
    'mName = (Name)
    'mValue = (Value)
    Set mVars = Vars
    
    Me.Show vbModal
    
    NewVariable = mSave
    If mSave Then
        Name = txtName.Text
        Value = txtValue.Text
    End If

    Unload Me
    Set frmDefVar = Nothing
End Function

Public Function ModVariable(ByVal Name As String, ByRef Value As String) As Boolean
    txtName.Text = Mid$(Name, 2)
    txtName.Enabled = False
    txtValue.Text = Value
    
    Me.Show vbModal
    
    ModVariable = mSave
    If mSave Then Value = txtValue.Text
    
    Unload Me
    Set frmDefVar = Nothing
End Function

Private Sub LoadLang()
    Me.Caption = mLang("cvar", "DefVar")

    cmdOk.Caption = mLang("", "Ok")
    cmdCancel.Caption = mLang("", "Cancel")
    lblName.Caption = mLang("cvar", "Name")
    lblValue.Caption = mLang("cvar", "Value")
End Sub

Private Sub cmdCancel_Click()
    mSave = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    Dim continue As Boolean
    
    If mVars Is Nothing Then
        continue = True
    Else
        continue = DoValidate
    End If
    
    If continue Then
        mSave = True
        Me.Hide
    End If
End Sub

Private Function DoValidate() As Boolean
    Dim Name As String
    
    Name = txtName.Text
    If Not Left$(Name, 1) = "@" Then Name = "@" & Name
    
    If InStr(1, Name, "=") Or InStr(1, Name, " ") Or Len(Trim$(Name)) = 0 Then
        MsgBox "Invalid variable name" & vbCrLf & _
               "equal signs (=) and spaces ar not allowed"
        DoValidate = False
        txtName.SetFocus
    ElseIf Not mVars.FindName(Name) = 0 Then
        MsgBox "A variable with this name already exist, chose another please"
        DoValidate = False
        txtName.SetFocus
    Else
        txtName.Text = Name
        DoValidate = True
    End If
End Function

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mVars = Nothing
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtValue_GotFocus()
    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue.Text)
End Sub
