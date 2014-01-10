VERSION 5.00
Begin VB.Form frmPalette 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imposta i colori"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmPalette.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRipristina 
      Caption         =   "Ripristina"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3675
      TabIndex        =   41
      Top             =   4125
      Width           =   1440
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "Conferma"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2175
      TabIndex        =   40
      Top             =   4125
      Width           =   1440
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Salva nuovo profilo"
      Height          =   315
      Left            =   375
      TabIndex        =   39
      Top             =   4125
      Width           =   1740
   End
   Begin VB.ComboBox cboProf 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   75
      Width           =   3315
   End
   Begin VB.Frame fraTextColor 
      Caption         =   "Colore del testo"
      Height          =   4140
      Left            =   75
      TabIndex        =   2
      Top             =   375
      Width           =   5115
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   0
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   19
         Top             =   300
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   1
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   18
         Top             =   675
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   2
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   675
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   3
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   16
         Top             =   1050
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   4
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   1050
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   5
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   14
         Top             =   1425
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   6
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   13
         Top             =   1425
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   7
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   8
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   9
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   10
         Top             =   2175
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   10
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   9
         Top             =   2175
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   11
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   8
         Top             =   2550
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   12
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   7
         Top             =   2550
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   13
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   6
         Top             =   2925
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   14
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   5
         Top             =   2925
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   15
         Left            =   1875
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   4
         Top             =   3300
         Width           =   615
      End
      Begin VB.PictureBox pctColor 
         Height          =   315
         Index           =   16
         Left            =   4350
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   3
         Top             =   3300
         Width           =   615
      End
      Begin VB.Label lblDefault 
         Caption         =   "Colore predefinito:"
         Height          =   240
         Left            =   150
         TabIndex        =   36
         Top             =   300
         Width           =   1740
      End
      Begin VB.Label lblLGray 
         Caption         =   "Grigio chiaro:"
         Height          =   240
         Left            =   150
         TabIndex        =   35
         Top             =   675
         Width           =   1740
      End
      Begin VB.Label lblBlack 
         Caption         =   "Nero:"
         Height          =   240
         Left            =   150
         TabIndex        =   34
         Top             =   1050
         Width           =   1740
      End
      Begin VB.Label lblDRed 
         Caption         =   "Rosso scuro:"
         Height          =   240
         Left            =   150
         TabIndex        =   33
         Top             =   1425
         Width           =   1740
      End
      Begin VB.Label lblDYellow 
         Caption         =   "Giallo scuro:"
         Height          =   240
         Left            =   150
         TabIndex        =   32
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label lblDGreen 
         Caption         =   "Verde scuro:"
         Height          =   240
         Left            =   150
         TabIndex        =   31
         Top             =   2175
         Width           =   1740
      End
      Begin VB.Label lblDCyan 
         Caption         =   "Ciano scuro:"
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   2550
         Width           =   1740
      End
      Begin VB.Label lblDBlue 
         Caption         =   "Blu scuro:"
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   2925
         Width           =   1740
      End
      Begin VB.Label lblDViolet 
         Caption         =   "Viola scuro:"
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   3300
         Width           =   1740
      End
      Begin VB.Label lblWhite 
         Caption         =   "Bianco:"
         Height          =   240
         Left            =   2625
         TabIndex        =   27
         Top             =   675
         Width           =   1740
      End
      Begin VB.Label lblDGray 
         Caption         =   "Grigio scuro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   26
         Top             =   1050
         Width           =   1740
      End
      Begin VB.Label lblLRed 
         Caption         =   "Rosso chiaro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   25
         Top             =   1425
         Width           =   1740
      End
      Begin VB.Label lblLYellow 
         Caption         =   "Giallo chiaro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   24
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label lblLGreen 
         Caption         =   "Verde chiaro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   23
         Top             =   2175
         Width           =   1740
      End
      Begin VB.Label lblLCyan 
         Caption         =   "Ciano chiaro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   22
         Top             =   2550
         Width           =   1740
      End
      Begin VB.Label lblLBlue 
         Caption         =   "Blu chiaro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   21
         Top             =   2925
         Width           =   1740
      End
      Begin VB.Label lblLViolet 
         Caption         =   "Viola chiaro:"
         Height          =   240
         Left            =   2625
         TabIndex        =   20
         Top             =   3300
         Width           =   1740
      End
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Default         =   -1  'True
      Height          =   315
      Left            =   3900
      TabIndex        =   1
      Top             =   4575
      Width           =   1290
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Top             =   4575
      Width           =   1290
   End
   Begin VB.Label lblProfile 
      Caption         =   "Profilo:"
      Height          =   240
      Left            =   150
      TabIndex        =   37
      Top             =   85
      Width           =   1740
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPalette As cPalette

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("palette", "caption")
    
    fraTextColor.Caption = mLang("palette", "TextColor")
    lblProfile.Caption = mLang("palette", "Profile")
    
    'commands
    cmdSaveAs.Caption = mLang("palette", "SaveAs")
    cmdConferma.Caption = mLang("palette", "Confirm")
    cmdRipristina.Caption = mLang("", "Abort")
    cmdOk.Caption = mLang("", "Ok")
    cmdAnnulla.Caption = mLang("", "Cancel")
    
    'color labels
    lblDefault.Caption = mLang("palette", "Default")
    lblWhite.Caption = mLang("palette", "White")
    lblBlack.Caption = mLang("palette", "Black")
    lblLGray.Caption = mLang("palette", "LightGray")
    lblLRed.Caption = mLang("palette", "LightRed")
    lblLYellow.Caption = mLang("palette", "LightYellow")
    lblLGreen.Caption = mLang("palette", "LightGreen")
    lblLCyan.Caption = mLang("palette", "LightCyan")
    lblLBlue.Caption = mLang("palette", "LightBlue")
    lblLViolet.Caption = mLang("palette", "LightViolet")
    lblDGray.Caption = mLang("palette", "DarkGray")
    lblDRed.Caption = mLang("palette", "DarkRed")
    lblDYellow.Caption = mLang("palette", "DarkYellow")
    lblDGreen.Caption = mLang("palette", "DarkGreen")
    lblDCyan.Caption = mLang("palette", "DarkCyan")
    lblDBlue.Caption = mLang("palette", "DarkBlue")
    lblDViolet.Caption = mLang("palette", "DarkViolet")
End Sub

Private Function GetColor(Index As Integer) As Long
    With mPalette
        Select Case Index
            Case 0
                GetColor = .rgbDefault
            Case 1
                GetColor = .rgbLightGrey
            Case 2
                GetColor = .rgbWhite
            Case 3
                GetColor = .rgbBlack
            Case 4
                GetColor = .rgbGrey
            Case 5
                GetColor = .rgbRed
            Case 6
                GetColor = .rgbLightRed
            Case 7
                GetColor = .rgbYellow
            Case 8
                GetColor = .rgbLightYellow
            Case 9
                GetColor = .rgbGreen
            Case 10
                GetColor = .rgbLightGreen
            Case 11
                GetColor = .rgbCyan
            Case 12
                GetColor = .rgbLightCyan
            Case 13
                GetColor = .rgbBlue
            Case 14
                GetColor = .rgbLightBlue
            Case 15
                GetColor = .rgbMagenta
            Case 16
                GetColor = .rgbLightMagenta
        End Select
    End With
End Function

Private Sub SetColor(Index As Integer, NewColor As Long)
    With mPalette
        Select Case Index
            Case 0
                .rgbDefault = NewColor
            Case 1
                .rgbLightGrey = NewColor
            Case 2
                .rgbWhite = NewColor
            Case 3
                .rgbBlack = NewColor
            Case 4
                .rgbGrey = NewColor
            Case 5
                .rgbRed = NewColor
            Case 6
                .rgbLightRed = NewColor
            Case 7
                .rgbYellow = NewColor
            Case 8
                .rgbLightYellow = NewColor
            Case 9
                .rgbGreen = NewColor
            Case 10
                .rgbLightGreen = NewColor
            Case 11
                .rgbCyan = NewColor
            Case 12
                .rgbLightCyan = NewColor
            Case 13
                .rgbBlue = NewColor
            Case 14
                .rgbLightBlue = NewColor
            Case 15
                .rgbMagenta = NewColor
            Case 16
                .rgbLightMagenta = NewColor
        End Select
    End With
End Sub

Private Sub cboProf_Click()
    mPalette.LoadColors cboProf.Text & ".col"
    LoadColors
End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
    Set frmPalette = Nothing
End Sub

Private Sub CheckChanges()
    Dim rtn As VbMsgBoxResult
    
    If cboProf.Enabled = False Then
        'rtn = MsgBox("Confermi le modifiche al profilo " & cboProf.Text & "?", vbYesNo)
        rtn = MsgBox(mLang("palette", "SaveChange") & " " & cboProf.Text & "?", vbYesNo)
        If rtn = vbYes Then
            SaveChanges
        End If
    End If
End Sub

Private Sub SaveChanges(Optional SaveAs As Boolean = False)
    Dim NewFile As String
    Dim i As Integer
    
    If Not SaveAs Then
        mPalette.SaveColors
    Else
        'NewFile = InputBox("Inserisci un nome per il nuovo profilo di colore")
        NewFile = InputBox(mLang("palette", "InsertName"))
        If Not NewFile = "" Then
            If Not LCase$(Right$(NewFile, 4)) = ".col" Then NewFile = NewFile & ".col"
            mPalette.SaveColors NewFile
            NewFile = Left$(NewFile, Len(NewFile) - 4)
            cboProf.AddItem NewFile
            For i = 0 To cboProf.ListCount - 1
                If cboProf.list(i) = NewFile Then
                    cboProf.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
    
    If (Not NewFile = "" Or Not SaveAs) Then
        cboProf.Enabled = True
        'cmdSaveAs.Enabled = False
        cmdRipristina.Enabled = False
        cmdConferma.Enabled = False
    End If
End Sub

Private Sub AbortChanges()
    mPalette.LoadColors cboProf.Text & ".col"
    LoadColors
    
    cboProf.Enabled = True
    'cmdSaveAs.Enabled = False
    cmdRipristina.Enabled = False
    cmdConferma.Enabled = False
End Sub

Private Sub cmdConferma_Click()
    SaveChanges
End Sub

Private Sub cmdOk_Click()
    Dim Connect As cConnector

    CheckChanges
    Set Connect = New cConnector
        Connect.ProfConf.AddInfo "Colours", cboProf.Text & ".col"
        Connect.Envi.SaveProfConfig
        Connect.Palette.LoadColors
    Set Connect = Nothing

    Unload Me
    Set frmPalette = Nothing
End Sub

Private Sub cmdRipristina_Click()
    AbortChanges
End Sub

Private Sub cmdSaveAs_Click()
    SaveChanges True
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    Set mPalette = New cPalette
    LoadList
End Sub

Private Sub LoadList()
    Dim fEnum As String
    Dim Path As String
    Dim Connect As cConnector
    Dim CurProf As String, i As Integer
    
    Set Connect = New cConnector
        CurProf = Connect.ProfConf.RetrInfo("Colours", "(default).col")
    Set Connect = Nothing
    
    cboProf.Clear
    Path = App.Path & "\colours\"
    fEnum = Dir$(Path)
    Do Until fEnum = ""
        If Right$(fEnum, 4) = ".col" Then
            cboProf.AddItem Left$(fEnum, Len(fEnum) - 4)
        End If
        fEnum = Dir$()
    Loop
    
    CurProf = Left$(CurProf, Len(CurProf) - 4)
    For i = 0 To cboProf.ListCount - 1
        If cboProf.list(i) = CurProf Then
            cboProf.ListIndex = i
            Exit For
        End If
    Next i
    
    If cboProf.ListIndex = -1 Then
        mPalette.LoadColors
        LoadColors
    End If
End Sub

Private Sub LoadColors()
    Dim i As Integer
    
    For i = 0 To 16
        pctColor(i).BackColor = GetColor(i)
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mPalette = Nothing

    Set mLang = Nothing
End Sub

Private Sub Changed()
    cboProf.Enabled = False
    cmdSaveAs.Enabled = True
    cmdRipristina.Enabled = True
    cmdConferma.Enabled = True
End Sub

Private Sub pctColor_Click(Index As Integer)
    Dim ComDlg As cCommonDialog

    Set ComDlg = New cCommonDialog
    Set ComDlg.Parent = Me
    ComDlg.Color = pctColor(Index).BackColor
    If ComDlg.ShowColor Then
        Changed
        pctColor(Index).BackColor = ComDlg.Color
        SetColor Index, ComDlg.Color
    End If
    Set ComDlg = Nothing
End Sub
