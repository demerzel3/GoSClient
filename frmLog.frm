VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar Progress 
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   675
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   975
      Width           =   990
   End
   Begin VB.OptionButton optNoBuff 
      Caption         =   "Ignore buffer"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   3315
   End
   Begin VB.OptionButton optUseBuff 
      Caption         =   "Use buffer"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Value           =   -1  'True
      Width           =   3315
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mBuff As cOutBuff
Private mFinal As String

Public Function StartLog() As String
    Set mBuff = frmMain.GetMudBuffer

    If Not mBuff Is Nothing Then
        If mBuff.Count > 0 Then
            frmLog.Show vbModal, frmBase
            
            StartLog = mFinal
            
            Unload Me
            Set frmLog = Nothing
        End If
    End If
End Function

Private Function HtmlColor(rgb As Long) As String
    Dim R As Byte, G As Byte, b As Byte

    R = rgb And &HFF
    G = (rgb \ &H100) And &HFF
    b = (rgb \ &H10000) And &HFF
    
    HtmlColor = Format(Hex(R), "00") & Format(Hex(G), "00") & Format(Hex(b), "00")
End Function

Private Sub cmdOk_Click()
    Dim Connect As cConnector
    
    If optUseBuff.Value Then
        optUseBuff.Enabled = False
        optNoBuff.Enabled = False
        cmdOk.Enabled = False
        Set Connect = New cConnector
            If Connect.GetBoolConfig("LogHtml") Then
                ConvertToHtml
            Else
                ConvertToPlainText
            End If
        Set Connect = Nothing
    Else
        Me.Hide
    End If
End Sub

Private Function TagToHtmlColor(ByVal Tag As String) As String
    Tag = LCase$(Tag)
    If Left$(Tag, 3) = "rgb" Then
        TagToHtmlColor = HtmlColor(rgb( _
            Val(Mid$(Tag, 4, 3)), Val(Mid$(Tag, 7, 3)), Val(Mid$(Tag, 10, 3))))
    End If
End Function

Private Sub ConvertToHtml()
    Dim i As Integer, b As cOutBuff
    Dim Final As String
    Dim nPos As Long, nEndPos As Long
    Dim CurLine As String, Tag As String, Color As String

    Set b = mBuff
    Progress.Max = b.Count
    For i = 1 To b.Count
        CurLine = b.Item((i))
        nPos = InStr(1, CurLine, TD)
        Do Until nPos = 0
            Final = Final & Left$(CurLine, nPos - 1)
            nEndPos = InStr(nPos + 1, CurLine, TD)
            If nEndPos = 0 Then nEndPos = Len(CurLine)
            Tag = Mid$(CurLine, nPos + 1, nEndPos - nPos - 1)
            Color = TagToHtmlColor(Tag)
            If Not Color = "" Then
                Final = Final & "</font><font color=" & Color & ">"
            End If
            CurLine = Mid$(CurLine, nEndPos + 1)
            'nPos = nEndPos + 1
            nPos = InStr(1, CurLine, TD)
        Loop
        Final = Final & CurLine & vbCrLf
        Progress.Value = i
        DoEvents
    Next i
    Set b = Nothing
    
    mFinal = Final
    Me.Hide
End Sub

Private Sub ConvertToPlainText()
    Dim i As Integer, b As cOutBuff
    Dim Final As String
    Dim nPos As Long, nEndPos As Long
    Dim CurLine As String

    Set b = mBuff
    Progress.Max = b.Count
    For i = 1 To b.Count
        CurLine = b.Item((i))
        nPos = InStr(1, CurLine, TD)
        Do Until nPos = 0
            Final = Final & Left$(CurLine, nPos - 1)
            nEndPos = InStr(nPos + 1, CurLine, TD)
            If nEndPos = 0 Then nEndPos = Len(CurLine)
            CurLine = Mid$(CurLine, nEndPos + 1)
            nPos = InStr(1, CurLine, TD)
        Loop
        Final = Final & CurLine & vbCrLf
        Progress.Value = i
        DoEvents
    Next i
    Set b = Nothing
    
    mFinal = Final
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        With Connect
            Me.Caption = .Lang("log", "Caption")
            optUseBuff.Caption = .Lang("log", "SaveBuffer")
            optNoBuff.Caption = .Lang("log", "IgnoreBuffer")
            cmdOk.Caption = .Lang("", "Ok")
        End With
    Set Connect = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mBuff = Nothing
End Sub
