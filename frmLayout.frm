VERSION 5.00
Begin VB.Form frmLayout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleziona layout"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmLayout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Default         =   -1  'True
      Height          =   315
      Left            =   5775
      TabIndex        =   4
      Top             =   3675
      Width           =   1590
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   4125
      TabIndex        =   3
      Top             =   3675
      Width           =   1590
   End
   Begin VB.PictureBox pctPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   3315
      Left            =   2850
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   2
      Top             =   300
      Width           =   4515
   End
   Begin VB.ListBox lstLyt 
      Height          =   3675
      IntegralHeight  =   0   'False
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   2715
   End
   Begin VB.Label lblCur 
      Caption         =   "Corrente:"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   75
      Width           =   2715
   End
   Begin VB.Label lblPreview 
      Caption         =   "Anteprima"
      Height          =   240
      Left            =   2850
      TabIndex        =   1
      Top             =   75
      Width           =   4515
   End
End
Attribute VB_Name = "frmLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("layout", "caption")

    cmdOk.Caption = mLang("", "Ok")
    cmdAnnulla.Caption = mLang("", "Cancel")
End Sub

Private Sub LoadList()
    Dim fEnum As String
    Dim Path As String
    Dim Connect As cConnector, CurLyt As String
    
    Set Connect = New cConnector
        CurLyt = Connect.ProfConf.RetrInfo("Layout", "(default).lyt")
        ShowPreview CurLyt, LCase$(mLang("layout", "Current")), False
        lblCur.Caption = mLang("layout", "Current") & ": " & Left$(CurLyt, Len(CurLyt) - 4)
    Set Connect = Nothing
    
    Path = App.Path & "\layouts\"
    fEnum = Dir$(Path)
    lstLyt.Clear
    Do Until fEnum = ""
        If Right$(fEnum, 4) = ".lyt" Then
            lstLyt.AddItem Left$(fEnum, Len(fEnum) - 4)
        End If
        fEnum = Dir$()
    Loop
End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
    Set frmLayout = Nothing
End Sub

Private Sub Save()
    Dim Connect As cConnector
    
    If Not lstLyt.Text = "" Then
        Set Connect = New cConnector
            Connect.ProfConf.AddInfo "Layout", lstLyt.Text & ".lyt"
            Connect.Envi.SaveProfConfig
            Connect.Envi.sendNotify ENVM_LYTCHANGED
        Set Connect = Nothing
    End If
End Sub

Private Sub cmdOk_Click()
    Me.Hide
    
    Save

    Unload Me
    Set frmLayout = Nothing
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    LoadList
    lblCur.FontBold = True
End Sub

Private Sub ShowPreview(ByVal LytFile As String, ByVal Name As String, _
    Optional ShowWinNames As Boolean = True)
    
    Dim lyt As cLayout
    Dim i As Integer
    Dim Height As Long, Width As Long
    Dim rc As RECT
    Dim win As String
    
    Set lyt = New cLayout
    lyt.LoadLayout LytFile
    pctPreview.Cls
    Height = pctPreview.ScaleHeight
    Width = pctPreview.ScaleWidth
    For i = 1 To lyt.Count
        With lyt.box(i)
            rc.Left = (.Left * Width) / 100
            rc.Top = (.Top * Height) / 100
            rc.Right = (.Right * Width) / 100
            rc.Bottom = (.Bottom * Height) / 100
            Rectangle pctPreview.hdc, rc.Left, rc.Top, _
                      rc.Right, rc.Bottom
            If ShowWinNames Then
                win = WinToTitle(.Window)
                DrawText pctPreview.hdc, win, Len(win), rc, DT_CENTER Or DT_SINGLELINE
            End If
        End With
    Next i
    Set lyt = Nothing
    
    lblPreview.Caption = mLang("layout", "Preview") & " [" & Name & "]"
End Sub

Private Function WinToTitle(ByVal win As String) As String
    'Select Case win
    '    Case "frmMain"
    '        WinToTitle = "Schermo di gioco"
    '    Case "frmChat"
    '        WinToTitle = "Chat"
    '    Case "frmMapper"
    '        WinToTitle = "Mapper"
    '    Case "frmStato"
    '        WinToTitle = "Stato"
    '    Case "frmMsp"
    '        WinToTitle = "Msp"
    '    Case "frmRubrica"
    '        WinToTitle = "Rubrica"
    '    Case "frmButtons"
    '        WinToTitle = "Pulsanti"
    'End Select
    WinToTitle = mLang("forms", win)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mLang = Nothing
End Sub

Private Sub lstLyt_Click()
    ShowPreview lstLyt.Text & ".lyt", lstLyt.Text
End Sub

Private Sub lstLyt_DblClick()
    Me.Hide
    
    Save
    Unload Me
    Set frmLayout = Nothing
End Sub
