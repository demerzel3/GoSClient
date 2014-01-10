VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "GoS Test Form"
   ClientHeight    =   10635
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14130
   FillStyle       =   0  'Solid
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   942
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstReport 
      Height          =   9810
      Left            =   75
      TabIndex        =   10
      Top             =   525
      Width           =   1665
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Speed report"
      Height          =   390
      Left            =   12450
      TabIndex        =   9
      Top             =   75
      Width           =   1440
   End
   Begin VB.CommandButton cmdSplitText 
      Caption         =   "SplitTextTest"
      Height          =   390
      Left            =   11175
      TabIndex        =   8
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "clean string test"
      Height          =   390
      Left            =   9675
      TabIndex        =   7
      Top             =   75
      Width           =   1440
   End
   Begin VB.TextBox txtSim 
      Height          =   285
      Left            =   7500
      TabIndex        =   6
      Text            =   "1"
      Top             =   150
      Width           =   465
   End
   Begin VB.CommandButton cmdSim 
      Caption         =   "Toggle simulation"
      Height          =   390
      Left            =   8025
      TabIndex        =   5
      Top             =   75
      Width           =   1590
   End
   Begin VB.Timer tmrSim 
      Enabled         =   0   'False
      Left            =   10200
      Top             =   1350
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MakeColorTag test"
      Height          =   390
      Left            =   5775
      TabIndex        =   4
      Top             =   75
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "palette test"
      Height          =   390
      Left            =   4425
      TabIndex        =   3
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ansi test"
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   1365
   End
   Begin GoS.uOutBox out 
      Height          =   9915
      Left            =   1800
      TabIndex        =   0
      Top             =   525
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   17489
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Drawing speed = "
      Height          =   240
      Left            =   1575
      TabIndex        =   2
      Top             =   150
      Width           =   2715
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mConnect As cConnector

Private mBuff As cOutBuff
Private mBuff2 As cOutBuff

'////////simulation variables
    Private mnSimFile As Integer

Private Sub cmdSim_Click()
    If tmrSim.Enabled = False Then
        tmrSim.Interval = Val(txtSim.Text)
        mnSimFile = FreeFile
        Open "logtest.txt" For Input As #mnSimFile
        tmrSim.Enabled = True
    Else
        tmrSim.Enabled = False
        Close #mnSimFile
        mnSimFile = 0
    End If
End Sub

Private Sub cmdSplitText_Click()
    Dim i As Long
    Dim t As Long
    
    AnsiTestSplitted
    out.BufferChange 1
    
    t = GetTickCount
    For i = 1 To 10
        out.SplitText Nothing, 70, mBuff.count
        'out.SplitText Nothing, 68, mBuff.Count - 68
    Next i
    MsgBox GetTickCount - t
End Sub

Private Sub Command1_Click()
    AnsiTestSplitted
End Sub

Private Sub AnsiTestSplitted(Optional buffDim As Long = 1024)
    Dim Buff() As String
    Dim nCount As Long, nPlus As Long
    Dim loa As String
    Dim i As Integer
    
    Open "logtest.txt" For Binary As #25
        nCount = (LOF(25) \ buffDim) + 1
        nPlus = LOF(25) Mod buffDim
        ReDim Buff(1 To nCount) As String
        For i = 1 To nCount
            Buff(i) = Space(buffDim)
            Get #25, , Buff(i)
            Buff(i) = Replace(Buff(i), vbCr & vbCrLf, vbCrLf)
        Next i
        Buff(nCount) = Left$(Buff(nCount), nPlus)
    Close #25
    
    Dim t As Long
    
    t = GetTickCount
    For i = 1 To nCount
        mBuff.AppendANSIText Buff(i)
    Next i
    MsgBox GetTickCount - t
End Sub

Private Sub AnsiTestComplete()
    Dim Buff As String
    Dim loa As String
    
    Open "logtest.txt" For Binary As #25
        Buff = Space(LOF(25))
        Get #25, , Buff
    Close #25
    
    Buff = Replace(Buff, vbCr & vbCrLf, vbCrLf)
    
    Dim t As Long
    
    t = GetTickCount
    mBuff.AppendANSIText Buff
    MsgBox GetTickCount - t
End Sub

Private Sub CleanStringTestComplete()
    Dim Buff As String
    Dim loa As String
    
    Open "logtest_tag.txt" For Binary As #25
        Buff = Space(LOF(25))
        Get #25, , Buff
    Close #25
    
    Dim t As Long
    
    t = GetTickCount
    'mBuff.AppendANSIText Buff
    'Buff = out.CleanString(Buff)
    MsgBox GetTickCount - t
    
    Open "logtest_clean.txt" For Binary As #25
        Put #25, , Buff
    Close #25
End Sub

Private Sub Command2_Click()
    Dim i As Long
    Dim t As Long
    
    t = GetTickCount
    
    For i = 1 To 100000
        mConnect.Palette.AnsiColor "1;37"
    Next i
    
    MsgBox GetTickCount - t
End Sub

Private Sub Command3_Click()
    Dim Buff As String
    Dim i As Long
    Dim t As Long
    Dim lenght As Long
    
    Buff = Space(16)
    t = GetTickCount
    For i = 1 To 100000
        lenght = gosuMakeColorTag(Buff, 16754199, True)
        Buff = Left$(Buff, lenght)
    Next i
    MsgBox GetTickCount - t
    
    Buff = Space(30)
    t = GetTickCount
    For i = 1 To 100000
        lenght = gosuMakeColorBackTag(Buff, 16754199, 16777215)
        Buff = Left$(Buff, lenght)
    Next i
    MsgBox GetTickCount - t
    
    t = GetTickCount
    For i = 1 To 100000
        'lenght = gosuMakeColorBackTag(buff, 16754199, 16777215)
        'buff = Left$(buff, lenght)
        Buff = mBuff.MakeColorTag(16754199) & mBuff.MakeColorTag(16777215, True)
    Next i
    MsgBox GetTickCount - t
    
    'MsgBox mBuff.TestMakeColorTag
End Sub

Private Sub Command4_Click()
    CleanStringTestComplete
End Sub

Private Sub Command5_Click()
    Dim time1 As Integer, time2 As Integer
    Dim loa As String
    Dim sReport As String
    Dim nGood As Integer, nBad As Integer
    
    lstReport.Clear
    Open "speedlog1.log" For Input As #25
    Open "speedlog2.log" For Input As #26
    Do Until EOF(25)
        Line Input #25, loa
        time1 = Val(Trim$(loa))
        Line Input #26, loa
        time2 = Val(Trim$(loa))
        
        If time1 > time2 Then
            sReport = time1 & " -> " & time2 & "  | good"
            nGood = nGood + 1
        ElseIf time2 > time1 Then
            sReport = time1 & " -> " & time2 & "  | BAD"
            nBad = nBad + 1
        Else
            sReport = time1 & " -> " & time2 & "  | equal"
        End If
        lstReport.AddItem sReport
    Loop
    Close #26
    Close #25
    
    MsgBox nGood & " good, " & nBad & " bad"
End Sub

Private Sub Form_Load()
    gMudPath = App.Path & "\TestMud\"
    If Dir$(gMudPath, vbDirectory) = "" Then MkDir gMudPath
    
    Set mConnect = New cConnector
    
    mConnect.Palette.LoadColors
    
    Set mBuff = New cOutBuff
    Set mBuff2 = New cOutBuff
    mBuff.Mode = SCMODE_DISK
    mBuff.Name = "test"
    mBuff.AppendLine TD & "RGB200200200" & TD
    
    out.BufferAdd mBuff
    out.BufferAdd mBuff2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mBuff = Nothing
    Set mBuff2 = Nothing

    Set mConnect = Nothing
End Sub

Private Sub Form_Resize()
    out.Width = Me.ScaleWidth - out.Left
    out.Height = Me.ScaleHeight - out.Top
    lstReport.Height = Me.ScaleHeight - lstReport.Top
End Sub

Private Sub out_DrawingTime(t As Long)
    lblSpeed.Caption = t
    If tmrSim.Enabled Then
        Open "speedlog.log" For Append As #26
            Print #26, t
        Close #26
    End If
End Sub

Private Sub tmrSim_Timer()
    Dim loa As String
    
    Line Input #mnSimFile, loa
    mBuff.AppendANSIText loa & vbCrLf
    
    If EOF(mnSimFile) Then tmrSim.Enabled = False
End Sub
