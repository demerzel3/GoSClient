VERSION 5.00
Begin VB.Form frmNota 
   AutoRedraw      =   -1  'True
   Caption         =   "Nuova nota"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "frmNota.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Invia >"
      Height          =   315
      Left            =   8625
      TabIndex        =   19
      Top             =   75
      Width           =   1440
   End
   Begin VB.TextBox txtNota 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   4665
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1650
      Width           =   10065
   End
   Begin VB.PictureBox pctTchiaro 
      BackColor       =   &H0000FFFF&
      Height          =   315
      Left            =   4575
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctT 
      BackColor       =   &H0000C0C0&
      Height          =   315
      Left            =   4275
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctCchiaro 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   3975
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctC 
      BackColor       =   &H00C0C000&
      Height          =   315
      Left            =   3675
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.TextBox txtSubject 
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
      Left            =   1350
      TabIndex        =   1
      Top             =   825
      Width           =   8760
   End
   Begin VB.TextBox txtTo 
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
      Left            =   1350
      TabIndex        =   0
      Top             =   450
      Width           =   8760
   End
   Begin VB.PictureBox pctBchiaro 
      BackColor       =   &H00FF0000&
      Height          =   315
      Left            =   3375
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctB 
      BackColor       =   &H00C00000&
      Height          =   315
      Left            =   3075
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctMchiaro 
      BackColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2775
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctM 
      BackColor       =   &H00C000C0&
      Height          =   315
      Left            =   2475
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctRchiaro 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   2175
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctR 
      BackColor       =   &H000000C0&
      Height          =   315
      Left            =   1875
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctVchiaro 
      BackColor       =   &H0000FF00&
      Height          =   315
      Left            =   1575
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctV 
      BackColor       =   &H0000C000&
      Height          =   315
      Left            =   1275
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctGscuro 
      BackColor       =   &H00808080&
      Height          =   315
      Left            =   975
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctGchiaro 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   675
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctN 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   375
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox pctZ 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   75
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Oggetto:"
      Height          =   240
      Left            =   75
      TabIndex        =   21
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "A:"
      Height          =   240
      Left            =   75
      TabIndex        =   20
      Top             =   450
      Width           =   915
   End
   Begin VB.Image imgRubrica 
      Height          =   195
      Left            =   1050
      Picture         =   "frmNota.frx":00D2
      Top             =   525
      Width           =   195
   End
End
Attribute VB_Name = "frmNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_CHARS As Integer = 80

Private Sub cmdSend_Click()
    Dim i As Integer
    Dim Connect As cConnector
    
    If txtTo.Text = "" Then
        MsgBox "Impossibile inviare un messaggio senza destinatario."
    ElseIf txtSubject.Text = "" Then
        MsgBox "Impossibile inviare un messaggio senza oggetto."
    Else
        Set Connect = New cConnector
            Connect.Envi.sendInput "nota a " & txtTo.Text & vbCrLf, TIN_TOQUEUE
            Connect.Envi.sendInput "nota soggetto " & txtSubject.Text & vbCrLf, TIN_TOQUEUE
            DividiTesto txtNota.Text
            'For i = 1 To txtNota.LineCount
            '    mParent.Send "nota + " & txtNota.GetLine(i) & vbCrLf
            'Next i
            Connect.Envi.sendInput "nota spedisci" & vbCrLf, TIN_TOQUEUE
            Unload Me
        Set Connect = Nothing
    End If
End Sub

Private Sub Form_Load()
    Me.Hide
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'mnbSend.Width = Me.ScaleWidth - mnbSend.Left * 2
    cmdSend.Left = Me.ScaleWidth - cmdSend.Width - 4
    txtTo.Width = Me.ScaleWidth - txtTo.Left - 4
    txtSubject.Width = Me.ScaleWidth - txtSubject.Left - 4
    txtNota.Width = Me.ScaleWidth - txtNota.Left - 4
    txtNota.Height = Me.ScaleHeight - txtNota.Top - 4
End Sub

Private Sub imgRubrica_Click()
    Dim Nick As String, Connect As cConnector

    Set Connect = New cConnector
        Nick = Connect.Rubrica.ChooseContact
    Set Connect = Nothing
    
    If Not Nick = "" Then txtTo.Text = Nick
End Sub

Private Function GetPosSpazio(Linea As String, ByVal Pos As Long) As Long
    Dim i As Integer

    If Not Linea = "" Then
        For i = Pos To 1 Step -1
            If Mid$(Linea, i, 1) = " " Then
                Pos = i + 1
                Exit For
            End If
        Next i
    End If
    GetPosSpazio = Pos
End Function

Private Sub SendLine(ByVal Linea As String)
    Dim Linea1 As String
    Dim Pos As String
    Dim i As Integer
    Dim Clean As String
    Dim Connect As cConnector

    Set Connect = New cConnector
    
    If Len(Linea) < MAX_CHARS Then
        Connect.Envi.sendInput "nota + " & Linea & vbCrLf, TIN_TOQUEUE
        
        Debug.Print Linea
    Else
        Do While Len(Linea) > MAX_CHARS
            Pos = GetPosSpazio(Linea, MAX_CHARS)
            Linea1 = Mid$(Linea, 1, Pos - 1)
            Connect.Envi.sendInput "nota + " & Linea1 & vbCrLf, TIN_TOQUEUE
            
            Debug.Print Linea1
            
            If Len(Linea) > MAX_CHARS Then
                Pos = GetPosSpazio(Linea, MAX_CHARS)
                Linea = Mid$(Linea, Pos)
            End If
        Loop
        Connect.Envi.sendInput "nota + " & Linea & vbCrLf, TIN_TOQUEUE
        
        Debug.Print Linea
    End If

    Set Connect = Nothing
End Sub

Private Sub DividiTesto(Stringa As String)
    Dim Pos As Long
    Dim Start As Long
    Dim Linea As String

    'If mAppVirtual Then
    '    mLines.Remove (mLines.Count)
    '    mAppVirtual = False
    'End If
    
    Start = 1
    Pos = InStr(1, Stringa, vbCrLf, vbTextCompare)
    Do Until Pos = 0
        Linea = Mid$(Stringa, Start, Pos - Start)
        SendLine Linea
        'UserControl.Print Mid$(Stringa, Start, Pos - Start)
        Start = Pos + 2
        Pos = InStr(Start, Stringa, vbCrLf, vbTextCompare)
    Loop
    'UserControl.Print Mid$(Stringa, Start)
    'mRicez = Not mRicez

    SendLine Mid$(Stringa, Start)
End Sub

Private Sub pctB_Click()
    AggiungiTesto "{{b"
End Sub

Private Sub pctBchiaro_Click()
    AggiungiTesto "{{B"
End Sub

Private Sub pctC_Click()
    AggiungiTesto "{{c"
End Sub

Private Sub pctCchiaro_Click()
    AggiungiTesto "{{C"
End Sub

Private Sub pctGchiaro_Click()
    AggiungiTesto "{{G"
End Sub

Private Sub pctGscuro_Click()
    AggiungiTesto "{{g"
End Sub

Private Sub pctM_Click()
    AggiungiTesto "{{m"
End Sub

Private Sub pctMchiaro_Click()
    AggiungiTesto "{{M"
End Sub

Private Sub pctN_Click()
    AggiungiTesto "{{N"
End Sub

Private Sub AggiungiTesto(Testo As String)
    txtNota.SelText = Testo
    txtNota.SetFocus
End Sub

Private Sub pctR_Click()
    AggiungiTesto "{{r"
End Sub

Private Sub pctRchiaro_Click()
    AggiungiTesto "{{R"
End Sub

Private Sub pctT_Click()
    AggiungiTesto "{{t"
End Sub

Private Sub pctTchiaro_Click()
    AggiungiTesto "{{T"
End Sub

Private Sub pctV_Click()
    AggiungiTesto "{{v"
End Sub

Private Sub pctVchiaro_Click()
    AggiungiTesto "{{V"
End Sub

Private Sub pctZ_Click()
    AggiungiTesto "{{Z"
End Sub
