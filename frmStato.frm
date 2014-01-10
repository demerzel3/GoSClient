VERSION 5.00
Begin VB.Form frmStato 
   AutoRedraw      =   -1  'True
   Caption         =   "Stato"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "frmStato.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCambia 
      Caption         =   "Cambia"
      Default         =   -1  'True
      Height          =   285
      Left            =   4050
      TabIndex        =   5
      Top             =   75
      Width           =   915
   End
   Begin GoS.uProgress pgrPf 
      Height          =   240
      Left            =   450
      TabIndex        =   1
      Top             =   450
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   423
   End
   Begin VB.TextBox txtPrompt 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "Pf:%#/%H Mn:%@/%M Mv:%//%V >"
      Top             =   75
      Width           =   2865
   End
   Begin GoS.uProgress pgrMn 
      Height          =   240
      Left            =   450
      TabIndex        =   2
      Top             =   750
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   423
   End
   Begin GoS.uProgress pgrMv 
      Height          =   240
      Left            =   450
      TabIndex        =   3
      Top             =   1050
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   423
   End
   Begin GoS.uProgress pgrPx 
      Height          =   240
      Left            =   450
      TabIndex        =   4
      Top             =   1350
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   423
   End
   Begin VB.Label lblPe 
      Caption         =   "Pe:"
      Height          =   240
      Left            =   75
      TabIndex        =   10
      Top             =   1350
      Width           =   465
   End
   Begin VB.Label lblMv 
      Caption         =   "Mv:"
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   1050
      Width           =   465
   End
   Begin VB.Label lblMn 
      Caption         =   "Mn:"
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Top             =   750
      Width           =   465
   End
   Begin VB.Label lblPf 
      Caption         =   "Pf:"
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   450
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Prompt:"
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Width           =   915
   End
End
Attribute VB_Name = "frmStato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1

Private mPf As Integer
Private mPfMax As Integer

Private mMana As Integer
Private mManaMax As Integer

Private mMov As Integer
Private mMovMax As Integer

Private mPx As Integer

Private mPrompt As String
Private mChanging As Boolean

Private mProfileSel As Integer

Private Sub cmdCambia_Click()
    mChanging = Not mChanging
    
    If mChanging Then
        cmdCambia.Caption = "Conferma"
        txtPrompt.Enabled = True
        txtPrompt.BackColor = 16777215
        txtPrompt.SelStart = 0
        txtPrompt.SelLength = Len(txtPrompt.Text)
        txtPrompt.SetFocus
        mChanging = True
    Else
        cmdCambia.Caption = "Cambia"
        txtPrompt.Enabled = False
        txtPrompt.BackColor = Me.BackColor
        mChanging = False
        mPrompt = txtPrompt.Text
    End If
End Sub

Private Sub SavePrompt()
    'Dim Config As cIni
    Dim Connect As cConnector
    
    'Set Config = New cIni
    '    Config.CaricaFile "config.ini"
    Set Connect = New cConnector
        Call Connect.SetConfig("Prompt<" & mProfileSel & ">", txtPrompt.Text)
        Connect.SaveConfig
    Set Connect = Nothing
    '    Config.SalvaFile
    'Set Config = Nothing
End Sub

Private Sub Init()
    'Dim Config As cIni, Connect As cConnector
    Dim Connect As cConnector
    
    If Not mProfileSel = -1 Then SavePrompt
    
    'Set Config = New cIni
    Set Connect = New cConnector
        'Config.CaricaFile "config.ini"
        mProfileSel = Connect.ProfileSel
        mPrompt = Connect.GetConfig("Prompt<" & mProfileSel & ">", "Pf:%#/%H Mn:%@/%M Mv:%//%V >")
        txtPrompt.Text = mPrompt
    Set Connect = Nothing
    'Set Config = Nothing
    
    LoadColors

    mPf = 0
    mPfMax = 0
    mMana = 0
    mManaMax = 0
    mMov = 0
    mMovMax = 0
    mPx = 0
    
    pgrPf.Value = 0
    pgrMn.Value = 0
    pgrMv.Value = 0
    pgrPx.Value = 0
End Sub

Private Sub LoadColors()
    Dim Connect As cConnector

    'Set Connect = New cConnector
        'txtPrompt.BackColor = Connect.RetrInfo("Win_Back", rgb(200, 200, 200))
        txtPrompt.BackColor = Me.BackColor
        'pgrPf.Color = Connect.RetrInfo("Pf_Color", rgb(99, 148, 255))
        pgrPf.Color = rgb(99, 148, 255)
        'pgrMn.Color = Connect.RetrInfo("Mn_Color", rgb(198, 253, 102))
        pgrMn.Color = rgb(120, 253, 102)
        'pgrMv.Color = Connect.RetrInfo("Mv_Color", rgb(255, 100, 99))
        pgrMv.Color = rgb(255, 100, 99)
        'pgrPx.Color = Connect.RetrInfo("Px_Color", rgb(255, 251, 95))
        pgrPx.Color = rgb(255, 220, 0)
    'Set Connect = Nothing
End Sub

Private Sub Form_Load()
    goshSetDockable Me.hWnd, "gos.stato"
    
    Set mFine = New cFinestra
    
    txtPrompt.Enabled = False
            
    mProfileSel = -1
    Init
    
    mFine.Init Me, WINREC_OUTPUT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SavePrompt
    
    mFine.UnReg
    
    Set mFine = Nothing
End Sub

Private Sub Form_Resize()
    Dim pheight As Long, pTop As Long

    On Error Resume Next
    
    txtPrompt.Width = Me.ScaleWidth - txtPrompt.Left * 2
    
    cmdCambia.Left = Me.ScaleWidth - cmdCambia.Width - 3
    
    pgrPf.Width = Me.ScaleWidth - pgrPx.Left - 5
    pgrMn.Width = pgrPf.Width
    pgrMv.Width = pgrPf.Width
    pgrPx.Width = pgrPf.Width
    
    pTop = (txtPrompt.Top + txtPrompt.Height + 5)
    pheight = Me.ScaleHeight - pTop - 5 - pgrPf.Height
    pheight = pheight / 3
    
    pgrPf.Top = pTop
    pgrMn.Top = pTop + pheight
    pgrMv.Top = pTop + pheight * 2
    pgrPx.Top = pTop + pheight * 3
    lblPf.Top = pgrPf.Top
    lblMn.Top = pgrMn.Top
    lblMv.Top = pgrMv.Top
    lblPe.Top = pgrPx.Top
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    Select Case uMsg
        Case ENVM_PROFILECHANGED
            Init
        Case ENVM_CLOSE
            pgrPf.Value = 0
            pgrMn.Value = 0
            pgrMv.Value = 0
            pgrPx.Value = 0
    End Select
End Sub

Private Sub mFine_envOutput(data As String, OutType As Integer)
    If OutType = TOUT_LASTLINE Then
        'lstOutput.AddItem data
        'lstOutput.ListIndex = lstOutput.ListCount - 1
        If Not mChanging Then Analizza data
    End If
End Sub

Private Sub Analizza(data As String)
    Dim Inizio As String, Pos As Integer
    Dim i As Integer, Pos2 As Integer
    Dim Limite As String

    'mPrompt = txtPrompt.Text
    
    Pos = InStr(1, mPrompt, "%", vbTextCompare)
    If Pos <= 1 Then Exit Sub
    
    Inizio = Left$(mPrompt, Pos - 1)
    
    Pos = 0
    'hai davanti a te un prompt! complimenti!
    If Inizio = Left$(data, Len(Inizio)) Then
        For i = 1 To Len(mPrompt)
            If Mid$(mPrompt, i, 1) = "%" Then
                If Pos = 0 Then Pos = i
                Pos2 = InStr(i + 2, mPrompt, "%", vbTextCompare)
                If Pos2 = 0 Then Pos2 = Len(mPrompt) + 1
                
                Pos2 = Pos2
                Limite = Mid$(mPrompt, i + 2, Pos2 - (i + 2))
                
                Estrapola Mid$(mPrompt, i, 2), data, Pos, Limite
                    'Pos = Pos + 1
                'End If
                i = i + 2
            Else
                If Not Pos = 0 Then Pos = Pos + 1
            End If
        Next i
    End If
    
    'If Not (mPfMax = 0 Or mManaMax = 0 Or mMovMax = 0) Then
    If Not mPfMax = 0 Then pgrPf.Value = (CLng(100) * mPf) \ mPfMax
    If Not mManaMax = 0 Then pgrMn.Value = (CLng(100) * mMana) \ mManaMax
    If Not mMovMax = 0 Then pgrMv.Value = (CLng(100) * mMov) \ mMovMax
    pgrPx.Value = (CLng(100) * mPx) \ 3000
        'lstOutput.AddItem (CLng(100) * mPf) \ mPfMax & "%  " & (CLng(100) * mMana) \ mManaMax & "%  " & (CLng(100) * mMov) \ mMovMax & "%  "
        'lstOutput.ListIndex = lstOutput.ListCount - 1
    'End If
End Sub

Private Sub Estrapola(Codice As String, Prompt As String, Pos As Integer, Limite As String)
    Dim sInfo As String, iInfo As Integer
    Dim Pos2 As Integer

    'lstOutput.AddItem Codice
    'lstOutput.AddItem Limite
    'lstOutput.ListIndex = lstOutput.ListCount - 1

    Pos2 = InStr(Pos, Prompt, Limite, vbTextCompare) - Pos
    
    If Pos2 <= 0 Then Exit Sub
    
    sInfo = Mid$(Prompt, Pos, Pos2)
    Pos = Pos + Len(sInfo) + 1

    If IsNumeric(sInfo) Then iInfo = Val(sInfo)
    
    Select Case Codice
        Case "%m", "%@"
            mMana = iInfo
        Case "%M"
            mManaMax = iInfo
        Case "%h", "%#"
            mPf = iInfo
        Case "%H"
            mPfMax = iInfo
        Case "%v", "%/"
            mMov = iInfo
        Case "%V"
            mMovMax = iInfo
        Case "%x"
            mPx = iInfo
    End Select
End Sub
