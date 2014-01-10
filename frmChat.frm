VERSION 5.00
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   Caption         =   "Chat"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraInfo 
      Caption         =   "Personaggi"
      Height          =   5415
      Left            =   7050
      TabIndex        =   2
      Top             =   0
      Width           =   2190
      Begin VB.ListBox lstPeople 
         Height          =   5100
         IntegralHeight  =   0   'False
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   2040
      End
   End
   Begin GoS.uOutBox txtChat 
      Height          =   5340
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   9419
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
      Left            =   60
      TabIndex        =   0
      Top             =   5550
      Width           =   9060
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1
Private mHistory As cHistory

Private mLoadContacts As Boolean
Private mSaveContacts As Boolean

Private mBuffers As Collection

Private mProfileName As String

Private Sub VerifyConfig()
    Dim Config As cConnector

    Set Config = New cConnector
        mHistory.ErasePrompt = Config.GetBoolConfig("ErasePrompt")
        mLoadContacts = Config.GetBoolConfig("ChatLoadContacts")
        mSaveContacts = Config.GetBoolConfig("ChatSaveContacts")
    Set Config = Nothing
End Sub

Private Function NameExist(ByVal Name As String, ByRef Index As Integer) As Boolean
    Dim i As Integer
    
    For i = 0 To lstPeople.ListCount - 1
        If LCase$(lstPeople.list(i)) = LCase$(Name) Then
            Index = i
            NameExist = True
            Exit For
        End If
    Next i
End Function

Private Sub AddMessage(Person As String, Message As String, Optional Title As Integer = 0, Optional Chan As String = "")
    Dim Index As Integer, Draw As Boolean
    Dim Text As String, Invia As Boolean, PrivPers As String
    Dim Connect As cConnector
    Dim IsMe As Boolean, Buff As cOutBuff

    'title 0 = public
    'title 1 = private
    'title 2 = you
    
    If Person = "" Then
        IsMe = True
        Person = mProfileName
    End If
    
    If Chan = "" Then
        PrivPers = "(parla)"
        Chan = ""
    ElseIf Chan = "urli" Or Chan = "urla" Then
        PrivPers = "(urla)"
        Chan = ""
    ElseIf Chan = "chatti" Or Chan = "chatta" Then
        PrivPers = "(chatta)"
        Chan = ""
    End If
    
    Select Case Title
        Case 0
            Text = vbCrLf & "[0;37m<" & Person & "> " & Message
        Case 1
            Text = vbCrLf & "[0m#" & Person & "> " & Message
        Case 2
            Text = vbCrLf & "[0;32m#" & Person & "> " & Message
    End Select
    
    If Not IsMe Then
        If InStr(1, Person, " ", vbTextCompare) = 0 And LCase$(Person) <> "qualcuno" Then
            If Not NameExist(Person, Index) Then
                lstPeople.AddItem Person ', Person
                If mSaveContacts Then
                    Set Connect = New cConnector
                        Connect.Rubrica.Add Person
                    Set Connect = Nothing
                End If
            End If
            
            If Title = 1 Then
                PrivPers = Person
            End If
        End If
    End If
    
    If Chan <> "" Then
        If InStr(1, Chan, " ", vbTextCompare) = 0 And LCase$(Chan) <> "qualcuno" Then
            If Not Chan = "Gruppo" Then
                If Not NameExist(Chan, Index) Then
                    lstPeople.AddItem Chan ', Chan
                    If mSaveContacts Then
                        Set Connect = New cConnector
                            Connect.Rubrica.Add Chan
                        Set Connect = Nothing
                    End If
                End If
            End If
            PrivPers = Chan
        End If
    End If
    
    Set Buff = AddBuffer(PrivPers, True, True)
    Buff.AppendANSIText Text
    Set Buff = Nothing
End Sub

Private Sub Form_GotFocus()
    If Me.Visible Then txtInput.SetFocus
End Sub

Private Function AddBuffer(Name As String, Optional Closeable As Boolean = False, _
                           Optional SwitchOn As Boolean = True)
    Dim newBuff As cOutBuff
    Dim i As Integer
    
    For i = 1 To mBuffers.Count
        If LCase$(mBuffers.Item(i).Name) = LCase$(Name) Then
            Set AddBuffer = mBuffers.Item(i)
            Exit Function
        End If
    Next i
            
    Set newBuff = New cOutBuff
        newBuff.Closeable = Closeable
        newBuff.Name = Name
        mBuffers.Add newBuff, Name
        txtChat.BufferAdd newBuff, SwitchOn
        Set AddBuffer = newBuff
    Set newBuff = Nothing
End Function

Private Sub Form_Load()
    Dim Connect As cConnector
    Dim i As Integer

    goshSetDockable Me.hWnd, "gos.chat"

    'inizializzazione chat
    Set mHistory = New cHistory
    mHistory.Init txtInput
        
    'inizializzazioni buffers pubblici
    Set mBuffers = New Collection
    AddBuffer "(parla)", False
    AddBuffer "(urla)", False, False
    AddBuffer "(chatta)", False, False
    
    FindProfileName
    VerifyConfig
        
    Set Connect = New cConnector
        If mLoadContacts Then
            With Connect.Rubrica
                For i = 1 To .Count
                    lstPeople.AddItem .Nick(i)
                Next i
            End With
        End If
        txtInput.ForeColor = Connect.Palette.rgbDefault
    Set Connect = Nothing

    Set mFine = New cFinestra
    mFine.Init Me, WINREC_OUTPUT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mBuffers = Nothing
    mFine.UnReg
    Set mFine = Nothing
    
    Set mHistory = Nothing
End Sub

Private Sub Form_Resize()
    'riposizionamento dei controlli sulla finestra
    
    On Error Resume Next
    fraInfo.Left = Me.ScaleWidth - fraInfo.Width - 4
    txtChat.Width = fraInfo.Left - txtChat.Left - 4
    txtInput.Top = Me.ScaleHeight - txtInput.Height - 3
    txtChat.Height = txtInput.Top - txtChat.Top - 5
    txtInput.Width = Me.ScaleWidth - txtInput.Left * 2
    fraInfo.Height = txtInput.Top - fraInfo.Top - 5
    lstPeople.Height = (fraInfo.Height - 20) * Screen.TwipsPerPixelY
End Sub

Private Sub lstPeople_Click()
    AddBuffer lstPeople.Text, True
End Sub

Private Sub lstPeople_GotFocus()
    txtInput.SetFocus
End Sub

Private Sub FindProfileName()
    Dim Profiles As cProfili

    Set Profiles = New cProfili
        Profiles.Carica
        If Profiles.ProfileSel = 0 Then
            mProfileName = "io"
        Else
            mProfileName = Profiles.Nick(Profiles.ProfileSel)
        End If
    Set Profiles = Nothing
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    Dim Connect As cConnector

    Select Case uMsg
        Case ENVM_CONFIGCHANGED
            VerifyConfig
        Case ENVM_PROFILECHANGED
            FindProfileName
        Case ENVM_PALCHANGED
            Set Connect = New cConnector
                txtInput.ForeColor = Connect.Palette.rgbDefault
            Set Connect = Nothing
    End Select
End Sub

Private Sub mFine_envOutput(data As String, OutType As Integer)
    If OutType = TOUT_CLEAN Then
        If InStr(1, data, " dici ", vbTextCompare) <> 0 Or _
           InStr(1, data, " urli ", vbTextCompare) <> 0 Or _
           InStr(1, data, " chatti ", vbTextCompare) <> 0 Or _
           InStr(1, data, " dice ", vbTextCompare) <> 0 Or _
           InStr(1, data, " urla ", vbTextCompare) <> 0 Or _
           InStr(1, data, " chatta ", vbTextCompare) <> 0 Then
                ToChat data
        End If
    End If
End Sub

Private Function FindTitle(pLine As String, ByRef Inizio As Long, ByRef Chan As String) As Integer
    Dim Save As Long

    Save = Len(pLine) + 1
    Inizio = InStr(1, pLine, "dice al gruppo", vbTextCompare)
    If Inizio <> 0 Then
        Save = Inizio
        FindTitle = 1
        Chan = "Gruppo"
    End If
    
    Save = Len(pLine) + 1
    Inizio = InStr(1, pLine, "dici", vbTextCompare)
    If Inizio <> 0 Then
        Save = Inizio
        FindTitle = 2
    End If
    
    Inizio = InStr(1, pLine, "urli", vbTextCompare)
    If Inizio <> 0 And Save > Inizio Then
        Save = Inizio
        FindTitle = 2
        Chan = "urli"
    End If
    
    Inizio = InStr(1, pLine, "chatti", vbTextCompare)
    If Inizio <> 0 And Save > Inizio Then
        Save = Inizio
        FindTitle = 2
        Chan = "chatti"
    End If
    
    Inizio = InStr(1, pLine, "ti dice", vbTextCompare)
    If Inizio <> 0 And Save > Inizio Then
        Save = Inizio
        FindTitle = 1
    End If
    
    Inizio = InStr(1, pLine, "dice", vbTextCompare)
    If Inizio <> 0 And Save > Inizio Then
        Save = Inizio
        FindTitle = 0
    End If
    
    Inizio = InStr(1, pLine, "chatta", vbTextCompare)
    If Inizio <> 0 And Save > Inizio Then
        Save = Inizio
        FindTitle = 0
        Chan = "chatta"
    End If
    
    Inizio = InStr(1, pLine, "urla", vbTextCompare)
    If Inizio <> 0 And Save > Inizio Then
        Save = Inizio
        FindTitle = 0
        Chan = "urla"
    End If
    
    Inizio = Save
End Function

Private Sub ToChat(pLine As String)
    Dim Message As String
    Dim Inizio As Long
    Dim Chan As String
    Dim Title As Integer '0 = public, 1 = private, 2 = you
    Dim Person As String
    Dim fine As Long
    Dim Stringa As String

    Inizio = InStr(1, pLine, "'", vbTextCompare)
    If Inizio <> 0 Then
        Message = Mid$(pLine, Inizio + 1, Len(pLine) - 1 - Inizio)
        
        Title = FindTitle(pLine, fine, Chan)
        If Chan = "" Or Chan = "Gruppo" Or Chan = "chatta" Or Chan = "urla" Then
            Select Case Title
                Case 2
                    Stringa = Trim$(Mid$(pLine, fine, Inizio - fine))
                    If InStr(1, Stringa, " a ", vbTextCompare) <> 0 Then
                        Chan = Trim$(Mid$(Stringa, InStr(1, Stringa, " a ", vbTextCompare) + 2))
                    Else
                        Chan = ""
                    End If
                Case Else
                    Person = Trim$(Mid$(pLine, 1, fine - 1))
                    Inizio = InStrRev(Person, ">", Len(Person))
                    If Inizio <> 0 Then
                        Person = Trim$(Mid$(Person, Inizio + 1))
                    End If
                    If Title = 1 Then Chan = Person
            End Select
        End If
        
        Call AddMessage(Person, Message, Title, Chan)
    End If
End Sub

Private Sub txtChat_BuffClosed(Index As Integer)
    mBuffers.Remove Index
End Sub

Private Sub txtChat_Click()
    txtInput.SetFocus
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim ToSend As String, Connect As cConnector

    If KeyAscii = 13 Then
        If Mid$(txtInput.Text, 1, 1) = "/" Then
            ToSend = Mid$(txtInput.Text, 2)
        Else
            Select Case LCase$(mBuffers.Item(txtChat.BufferSel).Name)
                Case "(parla)"
                    ToSend = "parla " & txtInput.Text
                Case "(urla)"
                    ToSend = "urla " & txtInput.Text
                Case "(chatta)"
                    ToSend = "chat " & txtInput.Text
                Case "gruppo"
                    ToSend = "digruppo " & txtInput.Text
                Case Else
                    ToSend = "di " & _
                        mBuffers.Item(txtChat.BufferSel).Name & " " & txtInput.Text
            End Select
        End If
        Set Connect = New cConnector
            Connect.Envi.sendInput ToSend & vbCrLf
        Set Connect = Nothing
    End If
End Sub
