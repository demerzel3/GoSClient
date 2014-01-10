VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cTriggers"
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Enum StopStateEnum
    cannot_stop = -1
    waiting = 0
    leaved = 1
    stopped = 2
End Enum

Private Type gosPluginInfo
    lpTitle As String * 128
    nMajor As Integer
    nMinor As Integer
    nRevision As Integer
End Type

Private Type gosPaths
    lpClientPath As String
    lpMudPath As String
    lpProfilePath As String
End Type

Private m_hLib As Long
Private m_hInfoProc As Long
Private m_hInitProc As Long
Private m_Title As String
Private m_File As String
Private m_SockID As Integer

Private m_Auto 'automatic loading plug-in?

Private mLocalEcho As Boolean 'local echo?

'variabili per il controllo del flusso di output
Private m_Stop As Boolean 'can stop mud's output?
'-1=can't stop 0=waiting for response 1=leaved 2=stopped
Private m_StopState As StopStateEnum
'data replaced to the original output
Private m_StopData As String

Private m_Connect As cConnector

Public Function SendFirstOutput(ByRef wsck As Winsock, ByRef data As String) As Boolean
    'return true if the output wasn't stopped, false otherwise
    
    Dim t As Long
    
    'If m_StopState = waiting Then
    '    Log "dovrebbe esserci accodamento, ma..."
    '    Log "...timeout in accodamento!! ignoro l'output e ritorno true"
    '    SendFirstOutput = True
    '    Exit Function
    '
    '    t = Timer
    '    Do While m_StopState = waiting
    '        DoEvents
    '        If Abs(Timer - t) > 10 Then
    '            'timeout error after 10 seconds
    '            'm_StopState = cannot_stop
    '            Log "timeout in accodamento!! ignoro l'output e ritorno true"
    '            SendFirstOutput = True
    '            Exit Function
    '        End If
    '    Loop
    'End If
    
    If m_Stop And m_hInitProc <> 0 Then
        'if the plug-in can stop the output, then wait for a response
        m_StopState = waiting
        wsck.SendData PIMD & "002 " & data
        'Log "invio output iniziale"
        
        t = Timer
        Do While m_StopState = waiting
            DoEvents
            If Abs(Timer - t) > 10 Then
                'timeout error after 10 seconds
                m_StopState = cannot_stop
                'Log "timeout schifo!!"
            End If
        Loop
        
        If m_StopState = stopped Then
            SendFirstOutput = False
            data = m_StopData 'replace original data with the new one
            'Log "stopped!"
        Else
            SendFirstOutput = True
            'Log "not stopped!"
        End If
        m_StopState = cannot_stop
    Else
        SendFirstOutput = True
    End If
End Function

Public Property Get Auto() As Boolean
    Auto = m_Auto
End Property

Public Property Let Auto(data As Boolean)
    m_Auto = data
End Property

Public Property Get Loaded() As Boolean
    Loaded = Not (m_hInitProc = 0)
End Property

Public Sub Log(ByVal data As String)
    m_Connect.Log "(" & m_File & ") " & data
End Sub

Public Sub VerifyConfig()
    Dim Config As cConnector
    
    Set Config = New cConnector
        mLocalEcho = Config.GetBoolConfig("LocalEcho", False)
    Set Config = Nothing
End Sub

Public Sub DataArrival(ByVal data As String)
    Dim lines() As String
    Dim i As Integer
    
    If Left$(data, 1) <> PIMD Then Exit Sub
    data = Mid$(data, 2)
    lines = Split(data, PIMD)
    For i = 0 To UBound(lines, 1)
        Dispatch lines(i)
    Next i
End Sub

Private Sub Dispatch(ByVal data As String)
    Dim Code As Integer
    
    If Len(data) < 3 Then Exit Sub
    Code = Val(Left$(data, 3))
    If Len(data) > 4 Then
        data = Right$(data, Len(data) - 4)
    Else
        data = ""
    End If
    Select Case Code
        Case 5 'invia input normale
            If Right$(data, 2) <> vbCrLf Then data = data & vbCrLf
            m_Connect.Envi.sendInput (data), TIN_TEXTBOX
            If m_Connect.Envi.ConnState = sckConnected Then
                If mLocalEcho Then
                    m_Connect.Envi.sendOutput (data)
                Else
                    m_Connect.Envi.sendOutput vbCrLf
                End If
            End If
        Case 6 'invia input alla coda di invio
            m_Connect.Envi.sendInput data, TIN_TOQUEUE
        Case 7 'invia output alla console errori
            m_Connect.LogError 0, data, m_File
        Case 8 'invia output alla console staus
            Log data
        Case 11 'non bloccare l'output
            If m_StopState = waiting Then
                m_StopState = leaved
            '    Log "messaggio di non-stop ricevuto"
            'Else
            '    Log "messaggio di non bloccaggio fuori tempo!"
            End If
        Case 12 'blocca l'output
            If m_StopState = waiting Then
                m_StopData = data
                m_StopState = stopped
                'm_Connect.Envi.sendOutput data
            '    Log "messaggio di stop ricevuto"
            'Else
            '    Log "messaggio di bloccaggio fuori-tempo!"
            End If
    End Select
End Sub

Public Function GetPath() As String
    'GetPath = App.Path & "\plugins\" & m_File
    'GetPath = "..\Debug\" & m_File
    GetPath = m_File
End Function

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Get SocketID() As Integer
    SocketID = m_SockID
End Property

Public Property Let SocketID(data As Integer)
    m_SockID = data
End Property

Public Function ReLoadInfo() As Boolean
    ReLoadInfo = LoadInfo(m_File)
End Function

Public Function LoadInfo(ByVal Path As String) As Boolean
    Dim gpi As gosPluginInfo
    
    If m_hLib = 0 Then
        m_File = Path
        'riattivare
        Path = App.Path & "\plugins\" & Path
        'Path = "..\Debug\" & Path
        m_hLib = LoadLibrary(Path)
        If m_hLib = 0 Then
            LoadInfo = False
        Else
            m_hInfoProc = GetProcAddress(m_hLib, "gosInfo")
            If m_hInfoProc = 0 Then
                LoadInfo = False
                FreeLibrary m_hLib
                m_hLib = 0
            Else
                gpi.lpTitle = Space(128)
                CallWindowProc m_hInfoProc, 0, m_Connect.Lang.LangID, 0, gpi
                m_Title = Trim$(gpi.lpTitle) & " (" & m_File & ")"
                LoadInfo = True
            End If
        End If
    End If
End Function

Public Function ConfPlugIn(Optional hWnd As Long = 0) As Boolean
    Dim hConfProc As Long, rtn As Long
    Dim gp As gosPaths, Connect As cConnector
    
    ReLoadInfo
    If Not m_hLib = 0 Then
        hConfProc = GetProcAddress(m_hLib, "gosConf")
        If hConfProc = 0 Then
            ConfPlugIn = False
        Else
            gp.lpClientPath = App.Path & "\"
            gp.lpMudPath = gMudPath
            Set Connect = New cConnector
                gp.lpProfilePath = Connect.ProfileFolder
            Set Connect = Nothing
            rtn = CallWindowProc(hConfProc, hWnd, m_Connect.Lang.LangID, 0, gp)
            If rtn = 0 Then
                ConfPlugIn = False
            Else
                ConfPlugIn = True
            End If
        End If
    End If
End Function

Public Function CredPlugIn(Optional hWnd As Long = 0) As Boolean
    Dim hCredProc As Long, rtn As Long
    
    ReLoadInfo
    If Not m_hLib = 0 Then
        hCredProc = GetProcAddress(m_hLib, "gosCred")
        If hCredProc = 0 Then
            CredPlugIn = False
        Else
            rtn = CallWindowProc(hCredProc, hWnd, m_Connect.Lang.LangID, 0, 0)
            If rtn = 0 Then
                CredPlugIn = False
            Else
                CredPlugIn = True
            End If
        End If
    End If
End Function

Public Function InitPlugIn() As Boolean
    Dim rtn As Long, CanStop As Long
    
    If Not m_hLib = 0 Then
        m_hInitProc = GetProcAddress(m_hLib, "gosInit")
        If m_hInitProc = 0 Then
            InitPlugIn = False
        Else
            rtn = CallWindowProc(m_hInitProc, ByVal frmBase.hWnd, m_Connect.Lang.LangID, gPluginsPort, CanStop)
            If rtn = 0 Then
                m_hInitProc = 0
                InitPlugIn = False
            Else
                m_Stop = Not (CanStop = 0)
                'If m_Stop Then MsgBox "can stop!"
                InitPlugIn = True
            End If
        End If
    End If
End Function

Public Function TermPlugIn() As Boolean
    If Not m_hLib = 0 Then
        FreeLibrary m_hLib
        'If Not m_hInitProc = 0 Then Log "plug-in terminato"
        If Not m_hInitProc = 0 Then Log "Plug-in terminated"
        m_hLib = 0
        m_hInitProc = 0
        m_hInfoProc = 0
    End If
End Function

Private Sub Class_Initialize()
    Set m_Connect = New cConnector

    m_StopState = cannot_stop
End Sub

Private Sub Class_Terminate()
    TermPlugIn
    Set m_Connect = Nothing
End Sub
