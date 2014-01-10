VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Schermo di gioco"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   3  'Windows Default
   Begin GoS.uOutBox txtMud 
      Height          =   3540
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   6540
      _extentx        =   15240
      _extenty        =   6244
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   3675
      Width           =   8040
   End
   Begin VB.Menu mnuFastCloseP 
      Caption         =   "FastCloseP"
      Begin VB.Menu mnuFastClose 
         Caption         =   "!Close!"
         Shortcut        =   %{BKSP}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1
Private mHistory As cHistory

'Buffers
Private mbStatus As cOutBuff
Private mbMud As cOutBuff

Private mFocusFromTxt As Boolean

'configs
Private mLocalEcho As Boolean
Private mFontSize As Integer
Private mFontName As String

'language reference
Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

'telnet protocol support and variables
Private WithEvents mTelnet As cTelnet
Attribute mTelnet.VB_VarHelpID = -1
Private mSendNAWS As Boolean

Public Function GetMudBuffer() As cOutBuff
    Set GetMudBuffer = mbMud
End Function

Public Sub SetTelnet(ByRef tel As cTelnet)
    Set mTelnet = tel
End Sub

Private Function SplitIntCode(ByVal Code As Integer) As String
    ReDim Buff(1 To 2) As Byte
    Buff(1) = Code And &HFF
    Buff(2) = (Code \ &H100&) And &HFF
    SplitIntCode = Chr$(Buff(2)) & Chr$(Buff(1))
End Function

Private Sub SendWindowSize()
    Dim WSize As String
    
    If mSendNAWS Then
        WSize = SplitIntCode(txtMud.GetWidth) & SplitIntCode(txtMud.GetHeight)
        'Debug.Print "NAWS " & WSize
        mTelnet.SendNAWS WSize
    End If
End Sub

Private Sub LoadLang()
    'Me.Caption = mLang("main", "Caption")
    SetWindowText Me.hWnd, mLang("main", "Caption")
End Sub

Private Sub VerifyConfig()
    Dim Config As cConnector
    Dim ConfIni As cIni

    Set Config = New cConnector
        If Config.GetBoolConfig("DiskBuffer", False) Then
            mbStatus.Mode = SCMODE_DISK
            mbMud.Mode = SCMODE_DISK
        Else
            mbStatus.Mode = SCMODE_MEMORY
            mbMud.Mode = SCMODE_MEMORY
        End If
        mHistory.ErasePrompt = Config.GetBoolConfig("ErasePrompt")
        mLocalEcho = Config.GetBoolConfig("LocalEcho", False)
    Set Config = Nothing
    
    Set ConfIni = New cIni
        ConfIni.CaricaFile App.Path & "\config.ini", True
        mFontSize = ConfIni.RetrInfo("FontSize", 10)
        mFontName = ConfIni.RetrInfo("FontName", "Courier")
        
        txtInput.FontName = mFontName
        txtInput.FontSize = mFontSize
        
        txtMud.SetFontName mFontName
        txtMud.SetFontSize mFontSize
    Set ConfIni = Nothing
    
    SendWindowSize
End Sub

Private Sub Form_Activate()
    SetFocusTxtInput
End Sub

Private Sub Form_GotFocus()
    SetFocusTxtInput
End Sub

Private Sub Form_Load()
    'Dim Config As cIni
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mbStatus = Connect.Envi.GetStatusBuff
        Set mLang = Connect.Lang
        LoadLang
    Set Connect = Nothing
    mbStatus.Name = "Status"
    txtMud.BufferAdd mbStatus
    
    Set mbMud = New cOutBuff
    mbMud.Name = "MUD"
    txtMud.BufferAdd mbMud, False
    'txtMud.Init txtInput
    'txtSplit.Init

    'txtSplit.SetBuffer txtMud.GetBuffer

    Set mHistory = New cHistory
    mHistory.Init txtInput

    VerifyConfig

    Set mFine = New cFinestra
    
    mFine.Init Me, WINREC_OUTPUT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Dim Config As cIni
    Dim Connect As cConnector
    
    If UnloadMode = vbFormControlMenu Then
        Beep
        Cancel = 1
    Else
        txtMud.BufferRemoveAll
        Set mbStatus = Nothing
        Set mbMud = Nothing
        
        Set mHistory = Nothing
        
        mFine.UnReg
        Set mFine = Nothing
    
        Set mTelnet = Nothing
    
        Set mLang = Nothing
    End If
End Sub

Private Sub DoResize(Optional Hide As Boolean = False)
    'Dim Bordi As cGrafica
    On Error Resume Next
    txtMud.Width = Me.ScaleWidth - txtMud.Left * 2
    txtInput.Width = txtMud.Width
    
    'txtMud.Top = 5
    
    txtInput.Top = Me.ScaleHeight - txtInput.Height - 3
    txtMud.Height = Me.ScaleHeight - (Me.ScaleHeight - txtInput.Top) - txtMud.Top - 3

    SendWindowSize
End Sub

Private Sub Form_Resize()
    DoResize
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    Select Case uMsg
        Case ENVM_PROFILECHANGED
            mbMud.Clear
        Case ENVM_CONFIGCHANGED
            VerifyConfig
        Case ENVM_SWITCHTOMUD
            txtMud.BufferChange 2
        Case ENVM_SWITCHTOSTATUS
            txtMud.BufferChange 1
        Case ENVM_MOUSEWHEELUP
            txtMud.DeltaUp
        Case ENVM_MOUSEWHEELDOWN
            txtMud.DeltaDown
    End Select
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub mnuFastClose_Click()
    frmBase.WindowState = vbMinimized
End Sub

Private Sub mTelnet_ReqNAWS()
    mSendNAWS = True
    SendWindowSize
End Sub

Private Sub mTelnet_WILLEcho()
    'the server will do the echo, so hide user input
    txtInput.PasswordChar = "#"
End Sub

Private Sub mTelnet_WONTEcho()
    'the server will no longer do the echo, so show user input
    txtInput.PasswordChar = ""
End Sub

Public Sub SetFocusTxtInput()
    If Not mFocusFromTxt Then
        SendMessage txtMud.hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0
    End If
    'mFocusFromTxt = False
End Sub

Private Sub txtInput_GotFocus()
    'Debug.Print "got focus!! mFocusFromTxt = " & mFocusFromTxt
    mFocusFromTxt = True
    
    SetFocusTxtInput
    mFocusFromTxt = False
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    'If mSplit Then
        Select Case KeyCode
            Case vbKeyPageDown
                txtMud.PageDown
            Case vbKeyPageUp
                txtMud.PageUp
        End Select
    'Else
    '    If KeyCode = vbKeyPageUp Then
    '        cmdSplit.Value = True
    '        txtSplit.PageLast
    '        txtSplit.PageUp
    '    End If
    'End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim Connect As cConnector, ToSend As String

    If KeyAscii = 13 Then
        'KeyAscii = 0
        Set Connect = New cConnector
        Connect.Envi.sendInput txtInput.Text & vbCrLf, TIN_TEXTBOX
        
        If Connect.Envi.ConnState = sckConnected Then
            If txtInput.PasswordChar <> "" Then
                'Connect.Envi.sendOutput String(Len(txtInput.Text), txtInput.PasswordChar) & vbCrLf
                ToSend = String(Len(txtInput.Text), txtInput.PasswordChar)
                
                'txtInput.PasswordChar = ""
                KeyAscii = 0
                txtInput.Text = ""
            Else
                'Connect.Envi.sendOutput txtInput.Text & vbCrLf
                ToSend = txtInput.Text
            End If
            
            If Not ((Left$(txtInput.Text, 1) = "@" And Not CBool(Connect.GetConfig("DisableVarParser"))) _
                Or Left$(txtInput.Text, 1) = "#") Then
                
                If Not mLocalEcho Then ToSend = ""
                Connect.Envi.sendOutput ToSend & vbCrLf
            End If
        End If
        Set Connect = Nothing
        'txtInput.Text = ""
    End If
End Sub

Private Sub mFine_envOutput(data As String, OutType As Integer)
    'txtMud.Text = txtMud.Text & data
    'txtMud.SelStart = Len(txtMud.Text)
    '
    'If InStr(1, data, "ÿû") <> 0 Then
    '    'data = Mid$(data, 1, Len(data) - 3)
    '    data = Replace(data, "ÿû", "")
    '    txtInput.PasswordChar = "#"
    'End If
    '
    'If InStr(1, data, "ÿü") Then
    '    'data = Mid$(data, 6)
    '    data = Replace(data, "ÿü", "")
    'End If
    
    Select Case OutType
        Case TOUT_SOCKET
            'txtMud.AppendText data
            mbMud.AppendANSIText data
    End Select
    'If mSplit Then txtSplit.Aggiorna
End Sub

Private Sub txtInput_LostFocus()
    'Debug.Print "lost focus!! mFocusFromTxt = " & mFocusFromTxt
    mFocusFromTxt = False
End Sub

Private Sub txtInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mFocusFromTxt = True
End Sub

Private Sub txtMud_Click()
    mFocusFromTxt = True
    txtInput.SetFocus
End Sub

Private Sub txtMud_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible Then
        mFocusFromTxt = True
        
        On Error Resume Next
        txtInput.SetFocus
    End If
End Sub
