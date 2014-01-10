VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mTrigger As String 'trigger text
Private mTrig() As String 'trigger splitted into singolar units
Private mVars() As String
Private mReaction As String

Private mnCount As Integer
Private mnVarCount As Integer

Public Function FirstLineReact() As String
    Dim nPos As Integer
    
    nPos = InStr(1, mReaction, vbCrLf)
    If nPos = 0 Then
        FirstLineReact = mReaction
    Else
        FirstLineReact = Left$(mReaction, nPos - 1) & " [...]"
    End If
End Function

Public Property Get Reaction() As String
    Reaction = mReaction
End Property

Public Property Let Reaction(data As String)
    mReaction = data
End Property

Public Function GetText() As String
    GetText = mTrigger
End Function

Public Function SetText(ByVal Text As String) As String
    mTrigger = Text
    SplitTrigger
End Function

Private Sub SplitTrigger()
    Dim Text As String
    Dim nCount As Integer
    Dim nPos As Integer, nEndPos As Integer
    Dim nVarCount As Integer
    
    Text = mTrigger
    Erase mTrig()
    mTrig = Split(Text, "%")
    nCount = (UBound(mTrig, 1) + 1) * 2
    ReDim mTrig(1 To nCount) As String
    
    'MsgBox "Trigger = " & Text & ", first_count = " & nCount
    
    nCount = 0
    nVarCount = 0
    'nPos = 1
    nPos = InStr(1, Text, "%")
    Do Until nPos = 0
        If Not nPos = 1 Then
            nCount = nCount + 1
            mTrig(nCount) = Left$(Text, nPos - 1)
        End If
        nCount = nCount + 1
        mTrig(nCount) = Mid$(Text, nPos, 2)
        nVarCount = Val(Mid$(Text, nPos + 1, 1))
        If mnVarCount < nVarCount Then mnVarCount = nVarCount
        Text = Mid$(Text, nPos + 2)
        'Debug.Print Text
    
        nPos = InStr(1, Text, "%")
    Loop
    If Not Text = "" Then
        nCount = nCount + 1
        mTrig(nCount) = Text
    End If
    
    ReDim Preserve mTrig(1 To nCount) As String
    Erase mVars()
    If Not mnVarCount = 0 Then ReDim mVars(1 To mnVarCount) As String
    'MsgBox "Trigger = " & mTrigger & ", real_count = " & nCount
    'Dim i As Integer
    'For i = 1 To nCount
    '    MsgBox """" & mTrig(i) & """"
    'Next i
    'MsgBox "trigger = " & mTrigger & ", varcount = " & mnVarCount
    
    mnCount = nCount
End Sub

Public Function Check(ByVal Text As String) As Boolean
    Dim i As Integer
    Dim nPos As Integer, nIniPos As Integer, nEndPos As Integer
    Dim CurVar As Integer, rtn As Integer
    
    nPos = 1
    nEndPos = 1
    CurVar = 0
    Check = True
    For i = 1 To mnCount
        rtn = IsVar(mTrig(i))
        If rtn > 0 Then
            CurVar = rtn
            nIniPos = nEndPos
        Else
            If i = mnCount Then
                nPos = InStrRev(Text, mTrig(i), -1, vbTextCompare)
                If nPos < nEndPos Then nPos = 0
            Else
                nPos = InStr(nEndPos, Text, mTrig(i), vbTextCompare)
            End If
            
            If nPos = 0 Then
                'one part of the trigger does not match
                Check = False
                CurVar = 0
                Exit For
            Else
                'this part of the trigger does match
                If Not CurVar = 0 Then
                    mVars(CurVar) = Trim$(Mid$(Text, nIniPos, (nPos - nIniPos)))
                End If
                nEndPos = nPos + Len(mTrig(i))
            End If
            If i = mnCount Then CurVar = 0
        End If
    Next i
    
    If Not CurVar = 0 Then
        mVars(CurVar) = Mid$(Text, nIniPos)
    End If
    
    If Not Check Then
        For i = 1 To mnVarCount
            mVars(i) = ""
        Next i
    'Else
    '    MsgBox "Trigger test on """ & Text & """"
    '    For i = 1 To mnVarCount
    '        MsgBox "Var%" & i & " =""" & mVars(i) & """"
    '    Next i
    End If
End Function

Public Function IsVar(ByRef Var As String) As Integer
    If Left$(Var, 1) = "%" And Len(Var) = 2 Then IsVar = Val(Mid$(Var, 2))
End Function

Public Sub Execute()
    Dim Connect As cConnector
    Dim ToExec As String
    Dim i As Integer

    If Trim$(mReaction) = "" Then Exit Sub

    ToExec = Replace(mReaction, "\%", "ÿ")
    For i = 1 To mnVarCount
        ToExec = Replace(ToExec, "%" & i, mVars(i))
    Next i
    ToExec = Replace(ToExec, "ÿ", "%")
    If Not Right$(RTrim$(ToExec), 2) = vbCrLf Then ToExec = ToExec & vbCrLf
    
    Set Connect = New cConnector
        Connect.Envi.sendInput ToExec, TIN_TOQUEUEALIAS
    Set Connect = Nothing
End Sub
