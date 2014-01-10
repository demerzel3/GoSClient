VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/////////////////// interprete della linea di comando \\\\\\\\\\\\\\\\\\\\\\\\\\

'puntatori alle classi per alias, trigger e variabili che la shell dovrà poter gestire

Private mVars As cVars
Private mAlias As cAlias
Private mTrigger As cTriggers

Private mDisableVarParser As Boolean

Public Sub SetInfo(ByRef Alias As cAlias, ByRef Trigger As cTriggers, ByRef Vars As cVars)
    Set mAlias = Alias
    Set mTrigger = Trigger
    Set mVars = Vars
End Sub

Private Sub Class_Initialize()
    mDisableVarParser = False
    VerifyConfig
End Sub

Private Sub Class_Terminate()
    Set mAlias = Nothing
    Set mTrigger = Nothing
    Set mVars = Nothing
End Sub

Public Function ProcessCommand(ByRef cmd As String) As Boolean
    Select Case Left$(cmd, 1)
        Case "#"
            ProcessCommand = False
            If Mid$(cmd, 2, 1) = "@" Then
                If mDisableVarParser Then
                    ProcessCommand = True
                Else
                    mVars.ProcessCommand Mid$(cmd, 2)
                End If
            Else
                DispatchCommand Mid$(cmd, 2)
            End If
        Case "@"
            If mDisableVarParser Then
                ProcessCommand = True
            Else
                ProcessCommand = False
                mVars.ProcessCommand cmd
            End If
        Case Else
            ProcessCommand = True
    End Select
End Function

Private Sub DispatchCommand(ByVal str As String)
    'str arriva qui già senza il # iniziale
    
    Dim nPos As Integer
    Dim sMsg As String, sPar As String
    
    If Right$(str, 2) = vbCrLf Then str = Left$(str, Len(str) - 2)
    
    nPos = InStr(1, str, " ")
    If nPos = 0 Then
        nPos = Len(str)
        sMsg = str
        sPar = ""
    Else
        sMsg = Left$(str, nPos - 1)
        sPar = Mid$(str, nPos + 1)
    End If
    
    sMsg = LCase$(sMsg)
    Select Case sMsg
        Case "", "help"
            ShowCommandList
        Case "notice"
            LogMessage "notice", sPar
        Case "echo"
            LogMessage "echo", sPar
        Case "beep"
            Beep
        Case "alias"
            ProcessAliasCommand sPar
        Case "trigger"
            ProcessTriggerCommand sPar
        Case Else
            Log sMsg & ", invalid or unknown command"
    End Select
End Sub

Private Sub LogMessage(Name As String, data As String)
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Connect.Envi.sendOutput vbCrLf & TD & "RGB098189000" & TD & "[" & Name & "] " & TD & "RGB230230230" & TD & data & vbCrLf
    Set Connect = Nothing
End Sub

Private Sub Log(data As String)
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Connect.Envi.sendOutput vbCrLf & TD & "RGB098189000" & TD & "[msg] " & TD & "RGB230230230" & TD & data & vbCrLf
    Set Connect = Nothing
End Sub

Private Sub LogAlias(data As String)
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Connect.Envi.sendOutput vbCrLf & TD & "RGB098189000" & TD & "[alias] " & TD & "RGB230230230" & TD & data '& vbCrLf
    Set Connect = Nothing
End Sub

Private Sub ShowCommandList()
    LogMessage "# or #help", "commands' list"
    LogMessage "#notice", "shows a message (i.e. '#notice hi!' shows '[notice] hi!' on the screen), useful in aliases and triggers"
    LogMessage "#beep", "sends the standard beep sound, useful in triggers"
    LogMessage "#alias", "manage aliases"
    LogMessage "#trigger", "manage triggers"
End Sub

Private Sub ProcessAliasCommand(sPar As String)
    Dim i As Integer, j As Integer
    
    If sPar = "" Then
        'enumera gli alias
        LogAlias FormatString("Alias", 20) ' & " " & FormatString("Commands", 40)
        LogAlias String(20, "-")
        For i = 1 To mAlias.Count
            LogAlias FormatString(mAlias.Text(i), 20) ' & " " & FormatString(mAlias.ActionPar(i, 1), 40)
            'For j = 2 To mAlias.ActionCount(i)
            '    LogAlias Space(21) & FormatString(mAlias.ActionPar(i, j), 40)
            'Next j
        Next i
        LogAlias String(20, "-") & vbCrLf
    End If
End Sub

Private Sub ProcessTriggerCommand(sPar As String)

End Sub


Public Function FormatString(ByVal str As String, nLen As Integer) As String
    Dim nCurLen As Integer
    
    nCurLen = Len(str)
    If nCurLen < nLen Then
        str = str & Space(nLen - nCurLen)
    ElseIf nCurLen > nLen Then
        str = Left$(str, nLen - 3) & "..."
    End If
    
    FormatString = str
End Function

Public Sub VerifyConfig()
    Dim Conf As cConnector
    
    Set Conf = New cConnector
        mDisableVarParser = Conf.GetConfig("DisableVarParser", 0)
    Set Conf = Nothing
End Sub

