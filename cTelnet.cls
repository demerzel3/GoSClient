VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'this class helps implementing the telnet protocol

Private Type Negotiation_Command 'WILL WON'T DO DON'T
    ncCode As Integer 'command code
    ncPos As Integer 'positive acknowlegdment
    ncNeg As Integer 'negative acknowlegdment
    bNeg As Boolean
End Type

Private Const sIAC As String = "ÿ"
Private Const sWILL As String = "û"
Private Const sWONT As String = "ü"
Private Const sDO As String = "ý"
Private Const sDONT As String = "þ"
Private Const sSB As String = "ú"      'start subnegotiations
Private Const sSE As String = "ð"      'end subnegotiations

Private Const iIAC As Integer = 255
Private Const iWILL As Integer = 251
Private Const iWONT As Integer = 252
Private Const iDO As Integer = 253
Private Const iDONT As Integer = 254
Private Const iSB As Integer = 250      'start subnegotiations
Private Const iSE As Integer = 240      'end subnegotiations

Private Const TEL_SEND As Integer = 1
Private Const TEL_IS As Integer = 0
Private Const TEL_CR As Integer = 13

Private Enum TelnetOptions
    TELOPT_ECHO = 1         'echo
    TELOPT_TTYPE = 24       'terminal type
    TELOPT_EOR = 25         'end of record
    TELOPT_NAWS = 31        'negotiate about window size
    TELOPT_NE = 39          'new environment
End Enum

Private mNeg(251 To 254) As Negotiation_Command

'/////////////////////////general events/////////////////////////
Public Event Send(data As String)

'/////////////////////////command-dependant events///////////////
Public Event ReqNAWS()      'Negotiate About Window Size

'/////////////////////////command-dependant notifications////////
Public Event WILLEcho()
Public Event WONTEcho()


Private Sub Send(data As String)
    'frmLog.Log "    sent data [" & data & "]"

    'If Not Right$(data, 2) = vbCrLf Then data = data & vbCrLf
    RaiseEvent Send(data)
End Sub

Public Sub SendNAWS(ByVal WS As String)
    'IAC SB NAWS <width><height> IAC SE
    Send sIAC & sSB & Chr$(TELOPT_NAWS) & WS & sIAC & sSE
End Sub



Private Sub Class_Initialize()
    'loading negotiation commands informations
    mNeg(iWILL).ncCode = iWILL
    mNeg(iWILL).ncPos = iDO
    mNeg(iWILL).ncNeg = iDONT
    
    mNeg(iWONT).ncCode = iWONT
    mNeg(iWONT).ncPos = iDONT
    mNeg(iWONT).ncNeg = iDO
    mNeg(iWONT).bNeg = True
    
    mNeg(iDO).ncCode = iDO
    mNeg(iDO).ncPos = iWILL
    mNeg(iDO).ncNeg = iWONT
    
    mNeg(iDONT).ncCode = iDONT
    mNeg(iDONT).ncPos = iWONT
    mNeg(iDONT).ncNeg = iWILL
    mNeg(iDONT).bNeg = True
    
    'Load 'frmLog
End Sub

Public Sub ProcessData(ByRef data As String, Optional ByVal StartPos As Long = 1)
    Dim Pos As Long, FinPos As Long, iChar As Integer
    Dim sMsg As String, NegCode As Integer
    Dim SubNegData As String
    
    Pos = InStr(StartPos, data, sIAC)
    If Not Pos = 0 Then
        sMsg = "IAC"
        iChar = Asc(Mid$(data, Pos + 1, 1))
        Select Case iChar
            Case iIAC 'IAC IAC sequence
                data = Left$(data, Pos - 1) & sIAC & Mid$(data, Pos + 2)
                Pos = Pos + 1
            Case iWILL, iWONT, iDO, iDONT
                NegCode = iChar
                sMsg = sMsg & " " & TranslateCommand(iChar)
                iChar = Asc(Mid$(data, Pos + 2, 1))
                sMsg = sMsg & " " & TranslateCommand(iChar) & _
                       " org[" & Mid$(data, Pos, 3) & "] cmd[" & iChar & "]"
                data = Left$(data, Pos - 1) & Mid$(data, Pos + 3)
                'frmLog.Log sMsg
                
                SendResponse NegCode, iChar
            'Case iWONT
            '    sMsg = sMsg & " WON'T"
            '    iChar = Asc(Mid$(data, Pos + 2, 1))
            '    sMsg = sMsg & " org[" & Mid$(data, Pos, 3) & "] cmd[" & iChar & "]"
            '    data = Left$(data, Pos - 1) & Mid$(data, Pos + 3)
            'Case iDO
            '    sMsg = sMsg & " DO"
            '    iChar = Asc(Mid$(data, Pos + 2, 1))
            '    sMsg = sMsg & " org[" & Mid$(data, Pos, 3) & "] cmd[" & iChar & "]"
            '    data = Left$(data, Pos - 1) & Mid$(data, Pos + 3)
            'Case iDONT
            '    sMsg = sMsg & " DON'T"
            '    iChar = Asc(Mid$(data, Pos + 2, 1))
            '    sMsg = sMsg & " org[" & Mid$(data, Pos, 3) & "] cmd[" & iChar & "]"
            '    data = Left$(data, Pos - 1) & Mid$(data, Pos + 3)
            Case iSB
                sMsg = sMsg & " SB"
                iChar = Asc(Mid$(data, Pos + 2, 1))
                sMsg = sMsg & " " & TranslateCommand(iChar)
                FinPos = InStr(Pos, data, sIAC & sSE)
                If Not FinPos = 0 Then
                    FinPos = FinPos + 2
                    sMsg = sMsg & " [" & Mid$(data, Pos, FinPos - Pos) & "]"
                Else
                    FinPos = Pos + 3
                End If
                SubNegData = Mid$(data, Pos + 3, (FinPos - 2) - (Pos + 3))
                data = Left$(data, Pos - 1) & Mid$(data, FinPos)
                'frmLog.Log sMsg
                
                SubNegotiate iChar, SubNegData
            Case Else
                sMsg = sMsg & " cmd[" & iChar & "]"
                data = Left$(data, Pos - 1) & Mid$(data, Pos + 2)
                'frmLog.Log sMsg
        End Select
        ProcessData data, Pos
    End If
End Sub

Private Function SubNegotiate(ByVal CommCode As Integer, data As String)
    Dim i As Integer, Code As String
    
    'frmLog.Log "SubNegotiate " & TranslateCommand(CommCode) & " [" & data & "]"
    
    For i = 1 To Len(data)
        Code = Code & Format(Asc(Mid$(data, i, 1)), "000") & " "
    Next i
    
    'frmLog.Log Code
    
    Select Case CommCode
        Case TELOPT_TTYPE
            If Asc(Left$(data, 1)) = TEL_SEND Then
                Send sIAC & sSB & Chr$(CommCode) & Chr$(TEL_IS) & "gosclient" & sIAC & sSE
            End If
        Case TELOPT_NE
    End Select
End Function

Private Function SendResponse(ByVal NegCode As Integer, CommCode As Integer)
    Dim iRtn As Integer
    
    If IsSupported(CommCode) Then
        iRtn = mNeg(NegCode).ncPos
    Else
        If (mNeg(NegCode).bNeg) Then
            iRtn = mNeg(NegCode).ncPos
        Else
            iRtn = mNeg(NegCode).ncNeg
        End If
    End If
    Send sIAC & Chr$(iRtn) & Chr$(CommCode)
    
    If CommCode = TELOPT_ECHO Then
        If NegCode = iWILL Then RaiseEvent WILLEcho
        If NegCode = iWONT Then RaiseEvent WONTEcho
    End If
    
    If NegCode = iDO Then 'this implies that the client send something
        Select Case CommCode
            Case TELOPT_NAWS
                RaiseEvent ReqNAWS
        End Select
    End If
End Function

Private Function IsSupported(ByRef id As Integer) As Boolean
    If TranslateCommand(id) = "" Then
        IsSupported = False
    Else
        IsSupported = True
    End If
End Function

Private Function TranslateCommand(id As Integer) As String
    Dim rtn As String
    
    Select Case id
        Case 1
            rtn = "ECHO"
        Case 24
            rtn = "TTYPE"
        Case 25
            rtn = "EOR"
        Case 31
            rtn = "NAWS"
'        Case 39
'            rtn = "NEW-ENVIRON"
        Case iWILL
            rtn = "WILL"
        Case iWONT
            rtn = "WONT"
        Case iDO
            rtn = "DO"
        Case iDONT
            rtn = "DONT"
        Case Else
            rtn = ""
    End Select
    
    TranslateCommand = rtn
End Function
