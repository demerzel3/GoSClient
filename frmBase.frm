VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBase 
   Caption         =   "GosClient"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8595
   Icon            =   "frmBase.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   StartUpPosition =   3  'Windows Default
   Begin GoS.uToolbar Toolbar 
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   847
   End
   Begin GoS.uDocking docking 
      Height          =   3990
      Left            =   0
      TabIndex        =   0
      Top             =   495
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   7038
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7950
      Top             =   1950
   End
   Begin VB.Timer tmrQueue 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7500
      Top             =   1950
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   7950
      Top             =   825
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "silmaril.novacomp.it"
      RemotePort      =   4000
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7950
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuLink 
      Caption         =   "Collegamento"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connetti"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Disconnetti"
      End
      Begin VB.Menu sel0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitList 
         Caption         =   "Esci all'elenco dei MUD"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu mnuPProfiles 
      Caption         =   "Profili"
      Begin VB.Menu mnuProfile 
         Caption         =   "(nessuno)"
         Index           =   0
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfiles 
         Caption         =   "Gestisci profili"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Strumenti"
      Begin VB.Menu mnuSettings 
         Caption         =   "Impostazioni..."
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Guarda log"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Imposta i colori"
      End
      Begin VB.Menu mnuRubrica 
         Caption         =   "Rubrica"
      End
      Begin VB.Menu mnuButtons 
         Caption         =   "Pulsanti configurabili"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Preferenze..."
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNote 
         Caption         =   "Gestisci le note"
      End
      Begin VB.Menu mnuMapper 
         Caption         =   "Mapper"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
      End
   End
   Begin VB.Menu mnuLayout 
      Caption         =   "Layout"
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockLayout 
         Caption         =   "Lock layout"
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "Plug-Ins"
      Begin VB.Menu mnuPlugIn 
         Caption         =   "plugin0"
         Index           =   0
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfPlugins 
         Caption         =   "Gestisci plug-ins"
      End
   End
   Begin VB.Menu mnuLangs 
      Caption         =   "Langs"
      Begin VB.Menu mnuLang 
         Caption         =   "lang0"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHomePage 
         Caption         =   "Home Page"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "Scrivi all'autore..."
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Informazioni su GosClient..."
      End
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'coda di invio dei dati (per eludere meccanismi anti-spam)
Private mQueue As Collection

'workspace, connector, finestra
Private mConnect As cConnector
Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1

'alias, triggers and variables
Private mAliases As cAlias
'Private mTriggers As cAlias
Private mTriggers As cTriggers
Private mVars As cVars

'combo receiver
Private WithEvents mCombo As cKeyCombo
Attribute mCombo.VB_VarHelpID = -1

'variabili per il controllo del log
Private mLogFile As Integer
Private mLogFileName As String
Private mLogLastLine As String
Private mLogHtml As Boolean

'variabili per il conteggio delle sessione e la gestione del collegamento
Private mSockError As Boolean
Private mNSess As Integer

'separatore fra diversi comandi dati sulla stessa linea
'se e' = " " allora i comandi multipli sono disattivati
Private mInputSep As String * 1

'variabili per la gestione dei plugins
Private mAccept As Boolean
Private mPlugInSel As Integer
Private mPlugins As cPlugIns
Private mPluginSendGo As Boolean

'variabili per la gestione della coda di output (causa plug-ins :P)
Private mOutQueue As Collection

'variabile per la gestione del tempo di connessione
Private mTimeStart As Long
Private mCurTime As Long

'language
Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

'telnet protocol support
Private WithEvents mTelnet As cTelnet
Attribute mTelnet.VB_VarHelpID = -1

'full screen mode
Private mFullScreen As Boolean
Private mMenu As cMenu

'shell
Private mShell As cShell

'var parser
Private mDisableVarParser As Boolean

Private Sub LoadLang()
    'link menu
    mnuLink.Caption = mLang("base", "Link")
    mnuConnect.Caption = mLang("base", "Connect")
    mnuClose.Caption = mLang("base", "CloseConn")
    mnuExit.Caption = mLang("", "Exit")
    mnuExitList.Caption = mLang("base", "ExitMuds")
    
    'profiles menu
    mnuPProfiles.Caption = mLang("base", "Profiles")
    mnuProfiles.Caption = mLang("base", "OrgProfiles")
    mnuProfile(0).Caption = mLang("", "Nobody")
    
    'tools menu
    mnuTools.Caption = mLang("base", "Tools")
    mnuSettings.Caption = mLang("base", "Settings")
    mnuNote.Caption = mLang("base", "Note")
    mnuLog.Caption = mLang("base", "WatchLog")
    'mnuChat.Caption = mLang("base", "Chat")
    mnuMapper.Caption = mLang("base", "Mapper")
    mnuStatus.Caption = mLang("base", "Status")
    mnuColors.Caption = mLang("base", "Colours")
    mnuRubrica.Caption = mLang("base", "Rubrica")
    mnuButtons.Caption = mLang("base", "Buttons")
    mnuOptions.Caption = mLang("base", "Preferences")

    'layout Menu
    mnuFullScreen.Caption = mLang("base", "FullScreen")
    mnuLayout.Caption = mLang("base", "Layout")
    If docking.Locked Then
        mnuLockLayout.Caption = mLang("base", "UnlockLayout")
    Else
        mnuLockLayout.Caption = mLang("base", "LockLayout")
    End If
    
    'plug-ins menu
    mnuPlugins.Caption = mLang("base", "Plugins")
    mnuConfPlugins.Caption = mLang("base", "OrgPlugins")
    
    'langugae menu
    mnuLangs.Caption = mLang("", "Language")
    
    'help menu
    mnuHelp.Caption = mLang("help", "Help")
    mnuHomePage.Caption = mLang("help", "URL")
    mnuMail.Caption = mLang("help", "EMail")
    mnuAbout.Caption = mLang("help", "About")
    
    'toolbar
    With Toolbar
        .SetToolTipText 1, mLang("base", "Connect")
        .SetToolTipText 2, mLang("base", "CloseConn")
        
        .SetToolTipText 3, mLang("base", "Settings")
        .SetToolTipText 4, mLang("base", "OrgProfiles")
        .SetToolTipText 5, mLang("base", "WatchLog")
        .SetToolTipText 6, mLang("base", "Colours")
        .SetToolTipText 7, mLang("base", "Rubrica")
        .SetToolTipText 8, mLang("base", "Buttons")
        .SetToolTipText 9, mLang("base", "OrgPlugins")
    
        .SetToolTipText 10, mLang("", "Exit")
    End With
End Sub

Private Sub LoadLangList()
    Dim fEnum As String, i As Integer
    
    i = 0
    fEnum = Dir$(App.Path & "\lang\")
    Do Until fEnum = ""
        If Right$(fEnum, 4) = ".lng" Then
            If Not i = 0 Then Load mnuLang(i)
            mnuLang(i).Visible = True
            fEnum = UCase$(Left$(fEnum, 1)) & LCase$(Mid$(fEnum, 2, Len(fEnum) - 5))
            If fEnum = mLang.Language Then
                mnuLang(i).Checked = True
            Else
                mnuLang(i).Checked = False
            End If
            mnuLang(i).Caption = fEnum
            i = i + 1
        End If
        fEnum = Dir$()
    Loop
    
    If i = 0 Then mnuLangs.Enabled = False
End Sub

Private Sub PluginLoadList()
    Dim i As Integer
    
    mPlugins.LoadList
    For i = 0 To mPlugins.Count - 1
        If Not i = 0 Then Load mnuPlugIn(i)
        mnuPlugIn(i).Visible = True
        mnuPlugIn(i).Caption = mPlugins.Item(i + 1).Title
        mnuPlugIn(i).Checked = False
        If mPlugins.Item(i + 1).Auto Then PlugInStart i + 1
    Next i

    If mPlugins.Count = 0 Then
        mnuPlugins.Enabled = False
    End If
End Sub

Private Sub PlugInSend(ByVal data As String)
    Dim i As Integer
    
    For i = 1 To sckServer.Count - 1
        If sckServer(i).State = sckConnected Then
            sckServer(i).SendData data
        End If
    Next i
End Sub

Private Sub PlugInSendProfileInfo()
    Dim Final As String
    
    'nome del profilo
    Final = PIMD & "051 " & _
        mConnect.GetConfig("profilo<" & mConnect.ProfileSel & ">", mLang("", "Nobody"))
    'cartella del profilo
    Final = Final & PIMD & "054 " & mConnect.ProfileFolder
    
    PlugInSend Final
End Sub

Private Sub PlugInSendFolders(ByVal Index As Integer)
    Dim Final As String
    
    'nome del mud
    Final = PIMD & "050 " & mConnect.Envi.Mud.Name
    'nome del profilo
    Final = Final & PIMD & "051 " & _
        mConnect.GetConfig("profilo<" & mConnect.ProfileSel & ">", mLang("", "Nobody"))
    'cartella del client
    Final = Final & PIMD & "052 " & App.Path & "\"
    'cartella del mud
    Final = Final & PIMD & "053 " & gMudPath
    'cartella del profilo
    Final = Final & PIMD & "054 " & mConnect.ProfileFolder
    
    sckServer(Index).SendData Final
End Sub

Private Sub PlugInStart(ByVal Index As Integer)
    Dim t As Long
    
    mPlugInSel = Index
    mAccept = True
    With mPlugins.Item(Index)
        .LoadInfo .GetPath
        If .InitPlugIn Then
            mnuPlugIn(Index - 1).Checked = Not mnuPlugIn(Index - 1).Checked
        Else
            mAccept = False
            'MsgBox "Errore nell'inizializzazione del plug-in"
            MsgBox mLang("base_err", "PluginInit")
        End If
    End With
    
    t = Timer
    Do While mAccept
        DoEvents
        If Abs(Timer - t) > 5 Then
            'mConnect.LogError 0, "Impossibile collegarsi con il plug-in, errore di timeout", "frmBase"
            mConnect.LogError 0, mLang("base_err", "PluginTimeout"), "frmBase"
            mnuPlugIn(Index - 1).Checked = False
            mAccept = False
            mPlugins.Item(Index).TermPlugIn
        End If
    Loop
End Sub

Private Sub PlugInSendGo()
    PlugInSend PIMD & "001 go"
    mPluginSendGo = True
End Sub

Private Function PlugInSendOutput(ByRef data As String) As Boolean
    Dim i As Integer
    
    PlugInSendOutput = True
    For i = 1 To mPlugins.Count
        With mPlugins.Item(i)
            If .SendFirstOutput(sckServer(.SocketID), data) = False Then
                PlugInSendOutput = False
            End If
        End With
    Next i
End Function

Private Sub docking_LockChanged(ByVal Value As Boolean)
    If docking.Locked Then
        mnuLockLayout.Caption = mLang("base", "UnlockLayout")
    Else
        mnuLockLayout.Caption = mLang("base", "LockLayout")
    End If
End Sub

Private Sub docking_MouseWheel(ByVal Delta As Integer)
    Dim i As Integer
    
    Delta = Delta / WHEEL_DELTA
    If Abs(Delta) >= 1 Then
        If Delta > 0 Then
            For i = 1 To Abs(Delta)
                mConnect.Envi.sendNotify ENVM_MOUSEWHEELUP
            Next i
        Else
            For i = 1 To Abs(Delta)
                mConnect.Envi.sendNotify ENVM_MOUSEWHEELDOWN
            Next i
        End If
    End If
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub mnuConfPlugins_Click()
    Load frmPlugins
    frmPlugins.Init mPlugins
End Sub

Private Sub ToggleFullScreen(Optional Restore As Boolean = False)
    Dim Style As Long
    
    Style = GetWindowLong(Me.hWnd, GWL_STYLE)
    If mFullScreen Then
        Style = Style Or (WS_CAPTION Or WS_SIZEBOX Or WS_SYSMENU Or _
                WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        SetWindowLong Me.hWnd, GWL_STYLE, Style
        
        Toolbar.Height = 32
        docking.Top = 32
        mMenu.RestoreMenu
        UpdateWindow Me.hWnd
    
        If Restore Then Me.WindowState = vbNormal
    
        mFullScreen = False
    Else
        Style = Style And (Not (WS_CAPTION Or WS_SIZEBOX Or WS_SYSMENU Or _
                WS_MAXIMIZEBOX Or WS_MINIMIZEBOX))
        SetWindowLong Me.hWnd, GWL_STYLE, Style
        
        Toolbar.Height = 28
        docking.Top = 28
        mMenu.RemoveMenu
        
        'ShowWindow Me.hwnd, SW_MAXIMIZE
        MoveWindow Me.hWnd, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, False
        UpdateWindow Me.hWnd
        ShowWindow Me.hWnd, SW_MAXIMIZE

        mFullScreen = True
    End If
    mnuFullScreen.Checked = mFullScreen
    Toolbar.SetFullScreen mFullScreen
End Sub

Private Sub mnuFullScreen_Click()
    ToggleFullScreen
End Sub

Private Sub mnuLang_Click(Index As Integer)
    Dim Conf As cIni, Filename As String
    Dim i As Integer
    
    If Not mnuLang(Index).Checked Then
        Set Conf = New cIni
            Conf.CaricaFile App.Path & "\config.ini", True
            Filename = LCase$(mnuLang(Index).Caption) & ".lng"
            Conf.AddInfo "Lang", Filename
            Conf.SalvaFile
        Set Conf = Nothing
        mLang.LoadLang Filename
    
        'WorkSpace.ReloadFormNames
    
        For i = 0 To mnuLang.Count - 1
            If mnuLang(i).Checked Then mnuLang(i).Checked = False
        Next i
        mnuLang(Index).Checked = True
        
        mMenu.RefreshMenu
    End If
End Sub

Private Sub mnuLockLayout_Click()
    docking.SetLocked (Not docking.Locked)
End Sub

Private Sub mnuPlugin_Click(Index As Integer)
    If mnuPlugIn(Index).Checked Then
        mAccept = False
        mPlugins.Item(Index + 1).TermPlugIn
        mnuPlugIn(Index).Checked = Not mnuPlugIn(Index).Checked
    Else
        PlugInStart Index + 1
    End If
End Sub

Private Sub mTelnet_Send(data As String)
    If Winsock.State = sckConnected Then Winsock.SendData data
End Sub

Private Sub sckServer_Close(Index As Integer)
    Dim PlugInID As Integer
    
    PlugInID = Val(sckServer(Index).Tag)
    mPlugins.Item(PlugInID).TermPlugIn
    mnuPlugIn(PlugInID - 1).Checked = False
    
    sckServer(Index).Close
    'Log "closed " & Index
    
    'close empty boxes (empty because of plug-ins closed)
    docking.CloseSpareBoxes
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Integer, rec As Integer
    
    'Log "Richiesta di connessione " & requestID
    If mAccept Then
        For i = 1 To sckServer.Count - 1
            If sckServer(i).State <> sckConnected Then
                rec = i
                Exit For
            End If
        Next i
        
        If rec = 0 Then
            rec = sckServer.Count
            Load sckServer(rec)
        End If
        
        sckServer(rec).Accept requestID
        PlugInSendFolders rec
        If Winsock.State = sckConnected Then PlugInSend PIMD & "010 " & Trim$(CStr(ENVM_CONNECT))
        'mPluginSendGo will be False until the layout has been initialized
        If mPluginSendGo Then sckServer(rec).SendData PIMD & "001 go"

        sckServer(rec).Tag = mPlugInSel
        mPlugins.Item(mPlugInSel).SocketID = rec
        mAccept = False
        'Log "client " & rec & " at port " & sckServer(rec).RemotePort
        mPlugins.Item(mPlugInSel).Log "Plug-in loaded"
    'Else
        'Log "not accepted"
    End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim data As String
    
    sckServer(Index).GetData data
    'Log "<socket" & Index & "> " & data
    mPlugins.Item(Val(sckServer(Index).Tag)).DataArrival data
End Sub








Public Sub ShowLog()
    If Not mLogFile = 0 Then
        Close #mLogFile
        Open mLogFileName For Append As #mLogFile
        ShellExecute Me.hWnd, vbNullString, mLogFileName, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub AddLineLog(ByVal data As String)
    Dim Palette As cPalette
    Dim iniPos As Integer, FinPos As Integer
    Dim Save As String, lenght As Integer

    Dim i As Integer, iChar As Integer, Char As String

    Dim AnsiCode As String
    Static LastColor As Integer
    Static BackColor As Long
    
    'mLimite = Limite
    If mLogHtml Then
        'logging in html
        If Not Len(data) = 0 Then
            data = Replace(data, "<", "&lt;")
            data = Replace(data, ">", "&gt;")
            
            Set Palette = mConnect.Palette
            
            lenght = Len(data)
            iniPos = 1
            Do
                iniPos = InStr(iniPos, data, ESCCHAR)
                If Not iniPos = 0 Then
                    'finPos = InStr(iniPos, data, "m")
                    
                    For i = iniPos + 1 To lenght
                        Char = Mid$(data, i, 1)
                        iChar = Asc(LCase$(Char))
                        If iChar >= 97 And iChar <= 122 Then
                            FinPos = i
                            Exit For
                        End If
                    Next i
                    
                    If i > lenght Then FinPos = 0
                    
                    If FinPos <> 0 Then
                        AnsiCode = Mid$(data, iniPos + 2, FinPos - iniPos - 2)
                        Select Case Char
                            Case "m"
                                data = Left$(data, iniPos - 1) & _
                                    "</font><font color=#" & HtmlColor(Palette.AnsiColor(AnsiCode, BackColor, LastColor)) & ">" & _
                                    Mid$(data, FinPos + 1)
                                iniPos = FinPos + 1
                            Case "C"
                                data = Left$(data, iniPos - 1) & _
                                    Space(AnsiCode) & _
                                    Mid$(data, FinPos + 1)
                            Case Else
                                data = Left$(data, iniPos - 1) & _
                                    Mid$(data, FinPos + 1)
                                iniPos = FinPos - (Len(AnsiCode) + 3)
                                If iniPos = 0 Then iniPos = 1
                        End Select
                        lenght = Len(data)
                    Else
                        'data = Left$(data, iniPos - 1)
                        iniPos = 0
                    End If
                End If
            Loop Until iniPos = 0
        
            Set Palette = Nothing
        End If
    Else
        'logging in plain text
        data = CleanString(data)
    End If
    
    Print #mLogFile, data
End Sub

Private Function HtmlColor(rgb As Long) As String
    Dim R As Byte, G As Byte, b As Byte

    R = rgb And &HFF
    G = (rgb \ &H100) And &HFF
    b = (rgb \ &H10000) And &HFF
    
    HtmlColor = Format(Hex(R), "00") & Format(Hex(G), "00") & Format(Hex(b), "00")
End Function

Public Sub StartLog()
    Dim fs As cFileSystem
    Dim Path As String
    Dim ToAppend As String
    Dim sExt As String

    'stop log to match configuration
    If mLogFile <> 0 Then
        sExt = LCase$(Right$(mLogFileName, 3))
        If (sExt = "log" And mLogHtml = True) Or _
           (sExt = "htm" And mLogHtml = False) Then
            mLogHtml = Not mLogHtml
            StopLog
            mLogHtml = Not mLogHtml
        End If
    End If

    If mLogFile = 0 Then
        Load frmLog
        ToAppend = frmLog.StartLog
    
        Path = mConnect.ProfileFolder & "logs\"
        Set fs = New cFileSystem
            fs.DirCreate Path
        Set fs = Nothing
        
        mLogFile = FreeFile()
        'path name without extension
        mLogFileName = Path & "gosLog_" & Format(Day(Date), "00") & "-" & Format(Month(Date), "00") & "-" & Format(Year(Date), "00") & "."
        
        If mLogHtml Then
            mLogFileName = mLogFileName & "htm"
        Else
            mLogFileName = mLogFileName & "log"
        End If
    
        Open mLogFileName For Append As #mLogFile
        
        If mLogHtml Then
            'print html headers
            Print #mLogFile, "<body bgcolor=#000000>"
            Print #mLogFile, "<pre><font color=#FFFFFF>"
        End If
        
        If Not ToAppend = "" Then
            Print #mLogFile, ToAppend
        End If
        'Open "LumenEtUmbra(bin).log" For Binary As #50
    End If
End Sub

Public Sub StopLog()
    If Not mLogFile = 0 Then
        If mLogHtml Then Print #mLogFile, "</font></pre>"
        Close mLogFile
        
        mLogFile = 0
    End If
End Sub

Public Function ProseguiInvio(ByRef Stringa As String, Optional ControlSep As Boolean = True) As Boolean
    Dim i As Integer, Alias As String
    Dim ToExec As String, Comando As String
    Dim Pos As Integer

    Stringa = Replace(Stringa, "\" & mInputSep, "ÿ")
    
    If mInputSep = " " Or (ControlSep = False) Then
        'non e' possibile avere piu' di un comando sulla stessa linea
        Pos = 0
    Else
        Pos = InStr(1, Stringa, mInputSep)
    End If
    
    If Pos <> 0 Then
        '//replace the escape character with an unused character
        'mConnect.Envi.sendOutput Stringa
        
        Stringa = Stringa & mInputSep
        ProseguiInvio = False
        ToExec = ""
        Do Until Pos = 0
            Comando = Mid$(Stringa, 1, Pos - 1)
            Stringa = Mid$(Stringa, Pos + 1)
            Comando = Replace(Comando, "ÿ", mInputSep)
            If ProseguiInvio(Comando, False) Then
                ToExec = ToExec & Comando & vbCrLf
            End If
            Pos = InStr(1, Stringa, mInputSep)
        Loop
        'If Not ToExec = "" Then RaiseEvent Esegui(ToExec)
        ToExec = Replace(ToExec, "ÿ", mInputSep)
        If Not ToExec = "" Then mConnect.Envi.sendInput ToExec '//, TIN_TOSEND
    Else
        Stringa = Replace(Stringa, "ÿ", mInputSep)
        ProseguiInvio = True
        'Stringa = Trim$(LCase$(Stringa))
        Stringa = Trim$(Stringa)
        For i = 1 To mAliases.Count
            Alias = LCase$(mAliases.Azione(i))
            If Not Alias = "" Then
                If LCase$(Left$(Stringa, Len(Alias))) = Alias Then
                    If Mid$(Stringa, Len(Alias) + 1, 1) = " " Or Len(Stringa) = Len(Alias) + 2 Then
                        SendAlias i, Right$(Stringa, Len(Stringa) - Len(Alias))
                        ProseguiInvio = False
                        Exit For
                    End If
                End If
            End If
        Next i
        'If ProseguiInvio Then mParser.ToMapper Stringa
    End If
End Function

Private Sub SendAlias(Index As Integer, Optional Param As String = "")
    mAliases.SendAlias Index, Param
End Sub

Private Sub AggiornaProfili()
    Dim i As Integer, Profili As cProfili
    Dim Connect As cConnector

    'Set mnuProfili = New cMenu
    Set Profili = New cProfili
    With Profili
        .Carica
        mConnect.Log "Profiles list loaded"
        'mnuProfili.Add "(nessuno)", "cmdProfilo"
        For i = 1 To mnuProfile.Count - 1
            Unload mnuProfile(i)
        Next i
        
        For i = 0 To .Count + 1
            If Not i = .Count + 1 Then
                If Not i = 0 Then Load mnuProfile(i)
                mnuProfile(i).Caption = .Nick(i)
                mnuProfile(i).Visible = True
                mnuProfile(i).Checked = False
            End If
            If .ProfileSel = i Then
                mnuProfile(i).Checked = True
                mConnect.Log "Current profile: " & .Nick(i) & " (" & i & ")"
                'stbStatus.Panels(1).Text = mConnect.Envi.Mud.Name & " - " & .Nick(i)
            End If
        Next i
        mnuProfile(0).Checked = (.ProfileSel = 0)
        'mnuProfili.Add "Sep1", "sep"
        
        Set Connect = New cConnector
        Connect.SetProfileSel .ProfileSel
        Set Connect = Nothing
    End With
    Set Profili = Nothing
    'mnuProfili.Add "Gestisci profili", "cmdProfili", , "Profiles"
    
    'mnuProfili.Draw Menus, "mnuProfili"
End Sub

Private Sub VerifyConfig()
    mLogHtml = mConnect.GetBoolConfig("LogHtml", True)
    If mConnect.GetBoolConfig("Logging") Then
        StartLog
    Else
        StopLog
    End If

    If mConnect.GetConfig("MultipleInput", 1) <> 0 Then
        mInputSep = mConnect.GetConfig("MultipleInputSep", ";")
    Else
        mInputSep = " "
    End If

    mDisableVarParser = mConnect.GetConfig("DisableVarParser", 0)

    mPlugins.VerifyConfig
    If Not mShell Is Nothing Then mShell.VerifyConfig
End Sub

Private Sub InitMenu()
    VerifyConfig
    AggiornaProfili
End Sub

Private Sub LoadTool(win As Form)
    Dim Control As Form, Show As Boolean

    Show = True
    For Each Control In Forms
        If TypeName(Control) = TypeName(win) Then
            Show = False
            Exit For
        End If
    Next
       
    'If Show Or (Not win.Tag = "d") Then win.Show vbModeless, Me
    If Show Or (goshCheckDocked(win.hWnd) = False) Then win.Show vbModeless, Me
End Sub

Private Sub GetConnInfos()
    'caricamento dati di connessione dalla struttura cMud
    'memorizzata nell'istanza globale di cEnviron
    Winsock.RemoteHost = mConnect.Envi.Mud.Host
    Winsock.RemotePort = mConnect.Envi.Mud.Port
End Sub

Private Sub Form_Activate()
    'If frmMain.Visible And frmMain.Tag = "d" Then
    If frmMain.Visible And goshCheckDocked(frmMain.hWnd) Then
        'Debug.Print "ma funziona?"
        frmMain.SetFocusTxtInput
    End If
End Sub

Private Sub Form_GotFocus()
    'If frmMain.Visible And frmMain.Tag = "d" Then
    If frmMain.Visible And goshCheckDocked(frmMain.hWnd) Then
        Debug.Print "ma funziona?"
        frmMain.SetFocusTxtInput
    End If
End Sub

Private Sub Form_Load()
    Dim winState As Integer
    
    '///////////////////////////////////////////////
    'debug instructions
    Load frmLog
    '///////////////////////////////////////////////
    
    mTimeStart = GetTickCount
    
    If Dir$(gMudPath & "config.ini") = "" Then
        Open gMudPath & "config.ini" For Output As #1: Close #1
    End If
    
    Set mConnect = New cConnector
    
    'load language strings
    Set mLang = mConnect.Lang
    'load translations
    LoadLang
    'load available languages' list
    LoadLangList
    
    Me.Caption = mConnect.Envi.Mud.Name & " - GoSClient"
    
    GetConnInfos
    
    '//////////////////////////////
    Set mTelnet = New cTelnet
    '//////////////////////////////
    
    '//////////// inizializzazione statusbar //////////////////
    'stbStatus.Panels.Add , , mConnect.Envi.Mud.Name
    'stbStatus.Panels(1).Bevel = sbrRaised
    'stbStatus.Panels(1).Alignment = sbrCenter
    'stbStatus.Panels.Add
    'stbStatus.Panels(2).Bevel = sbrNoBevel
    'stbStatus.Panels.Add
    'stbStatus.Panels(3).Bevel = sbrInset
    '//////////// inizializzazione statusbar //////////////////
    
    '/////////////CARICAMENTO PALETTE//////////////////
    mConnect.Palette.LoadColors
    'mConnect.Log mLang("base_log", "ColoursLoaded")
    mConnect.Log "Colours loaded"
    '/////////////CARICAMENTO PALETTE//////////////////
    
    '////////////CARICAMENTO PLUG-INS//////////////////
    Set mPlugins = New cPlugIns
    Set mOutQueue = New Collection
    
    SaveSetting "GosClient", "multiple", "port" & gPluginsPort, -1
    sckServer(0).LocalPort = gPluginsPort
    sckServer(0).Listen
    PluginLoadList
    'mAccept = False
    '////////////CARICAMENTO PLUG-INS//////////////////
    
    '////////////CARICAMENTO DIMENSIONI//////////////////
    winState = mConnect.GetConfig("MainWnd_State", Me.WindowState)
    If Not winState = vbMaximized Then
        Me.Left = mConnect.GetConfig("MainWnd_Left", Me.Left)
        Me.Width = mConnect.GetConfig("MainWnd_Width", Me.Width)
        Me.Top = mConnect.GetConfig("MainWnd_Top", Me.Top)
        Me.Height = mConnect.GetConfig("MainWnd_Height", Me.Height)
    End If
    Me.WindowState = winState
    '////////////CARICAMENTO DIMENSIONI//////////////////
    
    Set mFine = New cFinestra
    mFine.Init Me, WINREC_INPUT Or WINREC_OUTPUT
    
    InitMenu
    
    'create an alternative popup-menu for fullscreen mode
    Set mMenu = New cMenu
    mMenu.CreatePopupMenu Me.hWnd
    
    'caricamento in memoria degli alias
    Set mAliases = New cAlias
    Set mTriggers = New cTriggers
    Set mVars = New cVars
    'inizializza la shell
    Set mShell = New cShell
    AggiornaAlias
        
    'setting docking area properties
    Load frmMain
    frmMain.SetTelnet mTelnet
    goshSetOwner frmMain.hWnd, Me.hWnd
    goshSetDockable frmMain.hWnd, "gos.main"
    docking.SetMainWindow frmMain.hWnd
    docking.LoadLayout
        
    PlugInSendGo
    'WorkSpace.Init mConnect.ProfConf.RetrInfo("Layout", "(default).lyt")
    
    Set mCombo = New cKeyCombo
    mCombo.AvviaRicezione frmMain.txtInput, mAliases
    
    ConnectCmd True

    Set mQueue = New Collection
End Sub

Private Sub AggiornaAlias()
    mAliases.LoadAliases
    mTriggers.Load
    mVars.LoadVars
    
    mShell.SetInfo mAliases, mTriggers, mVars
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frm As Form ', Config As cIni
    'Dim Sett As cBinary

    '//////////////////////////////
    Set mTelnet = Nothing
    '//////////////////////////////

    Set mQueue = Nothing
    Set mOutQueue = Nothing
    
    StopLog
    Set mCombo = Nothing
    
    Set mShell = Nothing
    
    Set mAliases = Nothing
    Set mTriggers = Nothing
    Set mVars = Nothing
    
    docking.SaveLayout
    
    Set mPlugins = Nothing
        
    'mSpace.Destroy
    'Set mSpace = Nothing
    'WorkSpace.SaveWorkspace
    'WorkSpace.Destroy
    
    mFine.UnReg
    Set mFine = Nothing

    'Set Config = New cIni
        'Config.CaricaFile "config.ini"
        Call mConnect.SetConfig("MainWnd_Left", Me.Left)
        Call mConnect.SetConfig("MainWnd_Width", Me.Width)
        Call mConnect.SetConfig("MainWnd_Top", Me.Top)
        Call mConnect.SetConfig("MainWnd_Height", Me.Height)
        Call mConnect.SetConfig("MainWnd_State", Me.WindowState)
        mConnect.SaveConfig
        'Config.SalvaFile
    'Set Config = Nothing

    For Each frm In Forms
        'If frm.Visible = False Then
            Unload frm
            Set frm = Nothing
        'End If
    Next

    Set mConnect = Nothing
    
    SaveSetting "GosClient", "multiple", "port" & gPluginsPort, 0
End Sub

Private Function Sconnetti() As Boolean
    If Winsock.State <> sckClosed Then
        'terminate the current session?
        If MsgBox(mLang("base_quest", "TerminateSess"), vbYesNo) = vbYes Then
            mConnect.Envi.sendOutput vbCrLf
            'mConnect.Envi.sendOutput "Sessione " & mNSess & " terminata", TOUT_SERVICE
            mConnect.LogYellow "Session " & mNSess & " terminated"
            Winsock.Close
            Sconnetti = True
        Else
            Sconnetti = False
        End If
    Else
        Sconnetti = True
    End If
    If Sconnetti Then ConnectCmd True
End Function

Private Function Connetti() As Boolean
    Dim rtn As Boolean

    rtn = Sconnetti
    If rtn Then
        ConnectCmd False
        mNSess = mNSess + 1
        'mConnect.Envi.sendOutput "Sessione " & mNSess, TOUT_SERVICE
        mConnect.LogYellow "Session " & mNSess
        mConnect.Envi.sendOutput "Connecting...", TOUT_STATUS
        'txtMud.AppendText vbCrLf & "#> Sessione " & mNSess & " <#" & vbCrLf & _
        '                    "Connessione in corso..."
        Winsock.Connect
    End If
    Connetti = rtn
End Function

Private Sub Form_Resize()
    '////temporenee
    'txtMud.Width = Me.ScaleWidth - txtMud.Left * 2
    'txtInput.Width = Me.ScaleWidth - txtInput.Left * 2
    '\\\\\\\\\\\\\\
    
    'Menus.Width = Me.ScaleWidth - Menus.Left * 2
    'Toolbar.Width = Me.ScaleWidth - Toolbar.Left * 2
    'TaskBar.Width = Menus.Width
    'TaskBar.Top = Me.ScaleHeight - TaskBar.Height - 3
    
    On Error Resume Next
    'pctBox(0).Width = Me.ScaleWidth - pctBox(0).Left * 2
    'pctBox(0).Height = Me.ScaleHeight - pctBox(0).Top ' - TaskBar.Height - 3
    'WorkSpace.Width = Me.ScaleWidth - WorkSpace.Left * 2
    'WorkSpace.Height = Me.ScaleHeight - WorkSpace.Top ' - stbStatus.Height
    Toolbar.Width = Me.ScaleWidth
    
    docking.Width = Me.ScaleWidth - docking.Left * 2
    docking.Height = Me.ScaleHeight - docking.Top ' - stbStatus.Height
        
    'stbStatus.Width = Me.ScaleWidth
    'stbStatus.Panels(1).Width = 200
    'stbStatus.Panels(2).Width = Me.ScaleWidth - stbStatus.Panels(1).Width - stbStatus.Panels(3).Width
    
    'mSpace.LoadWorkspace
    'WorkSpace.LoadWorkspace
End Sub

Private Sub mCombo_AvviaAlias(Index As Integer)
    SendAlias Index
End Sub

Private Sub WinsockSend(data As String)
    If Not mDisableVarParser Then
        'send the string to the var parser
        mVars.ReplaceVars data
    End If
    
    If Winsock.State = sckConnected Then
        'send the string through the winsock to the mud
        Winsock.SendData data
        
        'reinsert the string into the client environment as sent data
        mConnect.Envi.sendInput data, TIN_SENT
    'Else
    '    mConnect.Envi.sendOutput data
    End If
End Sub

Private Sub mFine_envInput(data As String, InType As Integer)
    Dim i As Integer
    
    'If Left$(Trim$(data), 1) = "@" Then
    '    mVars.ProcessCommand data
    '    Exit Sub
    'End If
    If mShell.ProcessCommand(data) = False Then Exit Sub
    
    If Not InType = TIN_SENT Then
        If Winsock.State = sckConnected Then
            Select Case InType
                Case TIN_TOQUEUE
                    Enqueue data
                Case TIN_TOQUEUEALIAS
                    Enqueue data, True
                Case TIN_TOSEND
                    'Winsock.SendData data
                    'mConnect.Envi.sendOutput data
                    WinsockSend data
                Case TIN_BUTTONS
                    For i = 1 To mAliases.Count
                        If mAliases.Text(i) = data Then
                            SendAlias i
                            Exit For
                        End If
                    Next i
                Case Else
                    If ProseguiInvio(data) Then
                        'mConnect.Envi.sendOutput data
                        'Winsock.SendData data
                        'mConnect.Envi.sendInput data, TIN_SENT
                        WinsockSend data
                    End If
            End Select
        Else
            'mConnect.Envi.sendOutput "nessuna sessione attiva", TOUT_SERVICE
            mConnect.LogYellow "nessuna sessione attiva"
            'If ProseguiInvio(data) Then WinsockSend data
            'mConnect.Envi.sendOutput data & vbCrLf
        End If
    ElseIf InType = TIN_SENT Then
        PlugInSend PIMD & "005 " & data
    End If
End Sub

Private Sub Enqueue(data As String, Optional ControlAlias As Boolean = False)
    Dim lines() As String
    Dim i As Integer
    
    lines = Split(data, vbCrLf)
    For i = LBound(lines, 1) To UBound(lines, 1)
        If Not lines(i) = "" Then
            If ControlAlias Then
                If ProseguiInvio(lines(i) & vbCrLf) Then mQueue.Add lines(i)
            Else
                mQueue.Add lines(i)
            End If
        End If
    Next i
    tmrQueue.Interval = 100
    tmrQueue.Enabled = True
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    Dim i As Integer
    
    If uMsg = ENVM_PROFILECHANGED Then PlugInSendProfileInfo
    If uMsg <= ENVM_SETTCHANGED Then PlugInSend PIMD & "010 " & Trim$(CStr(uMsg))
    
    If uMsg = ENVM_CONFIGCHANGED Then
        VerifyConfig
    'ElseIf uMsg = ENVM_PROFILECHANGED Or uMsg = ENVM_LYTCHANGED Then
    '    WorkSpace.ChangeLayout mConnect.ProfConf.RetrInfo("Layout", "(default).lyt")
    End If
    
    If uMsg = ENVM_PROFILECHANGED Then
        'auto-terminate and start designed plug-ins
        mPlugins.UpdateAutoStart
        For i = 1 To mPlugins.Count
            With mPlugins.Item(i)
                If Not .Auto And .Loaded Then
                    .TermPlugIn
                ElseIf .Auto And Not .Loaded Then
                    PlugInStart i
                End If
            End With
        Next i
    End If
End Sub

Private Sub mFine_envOutput(data As String, OutType As Integer)
    Select Case OutType
        Case TOUT_SOCKET
            PlugInSend PIMD & "003 " & data
        Case TOUT_CLEAN
            PlugInSend PIMD & "004 " & data
    End Select
    
    If OutType = TOUT_SOCKET Then
        DividiTesto data
    End If
End Sub

Private Sub mnuAbout_Click()
    frmCredits.Show vbModal
End Sub

Private Sub mnuButtons_Click()
    LoadTool frmButtons
End Sub

Private Sub mnuClose_Click()
    Sconnetti
End Sub

Private Sub mnuColors_Click()
        frmPalette.Show vbModal, Me
    'mConnect.Envi.sendPalChanged
    mConnect.Envi.sendNotify ENVM_PALCHANGED
End Sub

Private Function BuildLoginString() As String
    Dim User As String, Pass As String
    Dim Profile As Integer
    Dim Count As Integer, i As Integer
    Dim Final As String, Current As String
    
    With mConnect
        Profile = .ProfileSel
        User = .GetConfig("profilo<" & Profile & ">")
        Pass = .GetConfig("pass<" & Profile & ">")
        Count = .GetConfig("login_count", 0)
        Final = ""
        For i = 1 To Count
            Current = .GetConfig("Login<" & i & ">", "")
            Select Case LCase$(Trim$(Current))
                Case "nickname"
                    Final = Final & User
                Case "password"
                    Final = Final & Pass
                Case Else
                    If Current = "" Then Current = " "
                    Final = Final & Current
            End Select
            Final = Final & vbCrLf
        Next i
    End With
    
    BuildLoginString = Final
End Function

Private Sub mnuConnect_Click()
    Dim Connect As cConnector
    'Dim InfoPro As cIni
    Dim Stringa As String
    
    Set Connect = New cConnector
    If Connect.ProfileSel = 0 Or Connect.GetConfig("login_count", 0) = 0 Then
        Connetti
    Else
        'Set InfoPro = New cIni
        'InfoPro.CaricaFile "config.ini"
        'Stringa = Connect.GetConfig("profilo<" & Connect.ProfileSel & ">") & vbCrLf & _
                  Connect.GetConfig("pass<" & Connect.ProfileSel & ">") & vbCrLf
        'Set InfoPro = Nothing
        Stringa = BuildLoginString
        
        If Connetti Then
            mSockError = False
            Do Until Winsock.State = sckConnected
                DoEvents
                If mSockError Then Exit Do
            Loop
            If Not mSockError Then
                'Winsock.SendData Stringa
                Enqueue Stringa
                'If Winsock.State = sckConnected Then Winsock.SendData " " & vbCrLf
                'If Winsock.State = sckConnected Then Winsock.SendData " " & vbCrLf
                'frmMain.txtInput.PasswordChar = ""
            End If
        End If
    End If
    Set Connect = Nothing
End Sub

Private Sub mnuExit_Click()
    'If WorkSpace.Mode = LYTMODE_MODIFY Then WorkSpace.SetMode LYTMODE_DOCK
    Unload Me
    Set frmMain = Nothing
End Sub

Private Sub mnuExitList_Click()
    Unload Me
    Set frmBase = Nothing
    
    frmMuds.Init True
End Sub

Private Sub mnuHomePage_Click()
    ShellExecute Me.hWnd, vbNullString, "http://gosclient.altervista.org/", _
        vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnuLog_Click()
    ShowLog
End Sub

Private Sub mnuMail_Click()
    ShellExecute Me.hWnd, vbNullString, "mailto:gosclient@yahoo.it", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnuMapper_Click()
    LoadTool frmMapper
End Sub

Private Sub mnuNote_Click()
    LoadTool frmNote
End Sub

Private Sub mnuOptions_Click()
    frmConfig.Show vbModal, Me
End Sub

Private Sub mnuProfile_Click(Index As Integer)
    Dim Connect As cConnector

    Set Connect = New cConnector
    mnuProfile(Connect.ProfileSel).Checked = False
    Connect.SetProfileSel Index
    mConnect.Log "Profilo attivo: " & mnuProfile(Index).Caption & " (" & Index & ")"
'    stbStatus.Panels(1).Text = mConnect.Envi.Mud.Name & " - " & mnuProfile(Index).Caption
    mnuProfile(Index).Checked = True
    Set Connect = Nothing
    'mConnect.Envi.sendChangeProfile
    mConnect.Envi.sendNotify ENVM_PROFILECHANGED
    AggiornaAlias

    If Not mLogFile = 0 Then
        StopLog
        StartLog
    End If
End Sub

Private Sub mnuProfiles_Click()
    Dim oldProf As Integer
    
    oldProf = mConnect.ProfileSel
    frmProfili.Show vbModal, Me
    AggiornaProfili
    If mConnect.ProfileSel <> oldProf Then
        AggiornaAlias
        If Not mLogFile = 0 Then
            StopLog
            StartLog
        End If
        mConnect.Envi.sendNotify ENVM_PROFILECHANGED
    End If
End Sub

Private Sub mnuRubrica_Click()
    LoadTool frmRubrica
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show vbModal, Me
    AggiornaAlias
    'per notificare al resto del programma che e' necessatio ricaricare le impostazioni
    mConnect.Envi.sendNotify ENVM_SETTCHANGED
End Sub

Private Sub mnuStatus_Click()
    LoadTool frmStato
End Sub

Private Sub tmrQueue_Timer()
    If Not mQueue.Count = 0 Then
        mConnect.Envi.sendInput mQueue.Item(1) & vbCrLf, TIN_TOSEND
        mQueue.Remove 1
    End If
    If mQueue.Count = 0 Then tmrQueue.Enabled = False
End Sub

Private Sub tmrTime_Timer()
    Dim newTime As Long, Diff As Long
    
    newTime = GetTickCount()
    Diff = ((newTime - mTimeStart) \ 1000)
    If Diff > 0 Then
        mCurTime = mCurTime + Diff
        mTimeStart = newTime
    End If
    'stbStatus.Panels(1).Text = mCurTime
End Sub

Private Sub Toolbar_ButtonClick(ByVal Key As String)
    Select Case Key
        Case "cmdPlugins"
            mnuConfPlugins_Click
        Case "cmdPalette"
            mnuColors_Click
        Case "cmdSettings"
            mnuSettings_Click
        Case "cmdRubrica"
            mnuRubrica_Click
        Case "cmdButtons"
            mnuButtons_Click
        Case "cmdProfiles"
            mnuProfiles_Click
        Case "cmdLog"
            mnuLog_Click
        Case "cmdConnect"
            mnuConnect_Click
        Case "cmdClose"
            mnuClose_Click
        Case "cmdExit"
            mnuExit_Click
    End Select
End Sub

Private Sub Toolbar_CloseClick()
    Unload Me
    Set frmBase = Nothing
End Sub

Private Sub Toolbar_DblClick()
    ToggleFullScreen
End Sub

Private Sub Toolbar_MenuClick(ByVal X As Long, ByVal Y As Long)
    mMenu.ShowPopupMenu X, Y
End Sub

Private Sub Toolbar_MinimizeClick()
    Me.WindowState = vbMinimized
End Sub

Private Sub Toolbar_RestoreClick()
    ToggleFullScreen
End Sub

Private Sub Winsock_Close()
    mConnect.Envi.sendOutput vbCrLf
    'mConnect.Envi.sendOutput "Sessione " & mNSess & " terminata", TOUT_SERVICE
    mConnect.LogYellow "Session " & mNSess & " terminated"
    'mConnect.Envi.sendClose
    mConnect.Envi.sendNotify ENVM_CLOSE
    Winsock.Close
    ConnectCmd True
    mConnect.Envi.sendNotify ENVM_SWITCHTOSTATUS
End Sub

Private Sub ConnectCmd(Enable As Boolean)
    Dim i As Integer

    'Toolbar.SetEnabled Enable, 1 'connetti
    'Toolbar.SetEnabled Enable, 6 'profili
    'Toolbar.SetEnabled (Not Enable), 2 'disconnetti
    'Toolbar2.Buttons.Item(1).Enabled = Enable
    'Toolbar2.Buttons.Item(5).Enabled = Enable
    'Toolbar2.Buttons.Item(2).Enabled = Not Enable
    Toolbar.SetEnabled 1, Enable
    Toolbar.SetEnabled 4, Enable
    Toolbar.SetEnabled 2, Not Enable
    mnuConnect.Enabled = Enable
    mnuClose.Enabled = (Not Enable)
    
    'For i = 0 To mnuProfile.Count - 1
    '    mnuProfile(i).Enabled = Enable
    'Next i
    mnuPProfiles.Enabled = Enable
End Sub

Private Sub Winsock_Connect()
    mConnect.Envi.sendNotify ENVM_SWITCHTOMUD
    mConnect.Envi.sendNotify ENVM_CONNECT
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim data As String, rtn As Boolean

    Static PluginWaiting As Boolean ', t As Long

    'If frmMain.txtInput.PasswordChar <> "" Then frmMain.txtInput.PasswordChar = ""
    
    Winsock.GetData data
    
    'mConnect.Log "got data from socket"
    
    mTelnet.ProcessData data
    
    data = Replace(data, Chr$(27) & "7", "")
    data = Replace(data, vbLf + vbCr, TD)
    data = Replace(data, vbCrLf, TD)
    data = Replace(data, vbCr, "")
    data = Replace(data, vbLf, TD)
    data = Replace(data, Chr(0), "")
    data = Replace(data, TD, vbCrLf)
    data = Replace(data, Chr$(9), "    ")
    
    'Open "dde.txt" For Append As #50
    '    Print #50, data
    'Close #50
    'Open "sorgenteLog.txt" For Append As #50
    '    Print #50, data;
    'Close #50
    
    mOutQueue.Add data
    If PluginWaiting Then
        'mConnect.Log "   waiting for plug-in (" & mOutQueue.Count & " in queue)"
        Exit Sub
        
        't = Timer
        'Do While pluginwaiting
        '    DoEvents
        '    If Abs(Timer - t) > 10 Then
        '        'after 10 seconds stop waiting
        '        Exit Do
        '    End If
        'Loop
    End If
    
    Do Until mOutQueue.Count = 0
        data = mOutQueue.Item(1)
        
        'mConnect.Log "before send to plug-ins, pluginwaiting = true"
        'rtn = true se l'output non e' stato fermato, altrimenti e' false
        PluginWaiting = True
        rtn = PlugInSendOutput(data)
        PluginWaiting = False
        'mConnect.Log "after sending, pluginwaiting = false (rtn = " & rtn & ")"
        
        If Not data = "" Then
            mConnect.Envi.sendOutput data
        'Else
        '    Debug.Print "stopped!"
        End If
        
        mOutQueue.Remove 1
    Loop
End Sub

Private Sub DividiTesto(ByVal Stringa As String)
    Dim Pos As Long
    Dim Start As Long
    Dim Linea As String

    Start = 1
    Pos = InStr(1, Stringa, vbCrLf, vbBinaryCompare)
    Do Until Pos = 0
        Linea = Mid$(Stringa, Start, Pos - Start)
        ControlTrigger Linea
        
        'esegui logging della linea
        If mLogFile <> 0 Then
            If mLogLastLine = "" Then
                AddLineLog Linea
            Else
                AddLineLog mLogLastLine & Linea
                mLogLastLine = ""
            End If
        End If
        
        Start = Pos + 2
        Pos = InStr(Start, Stringa, vbCrLf, vbBinaryCompare)
    Loop

    Stringa = Mid$(Stringa, Start)
            
    mConnect.Envi.sendNotify ENVM_ENDREC
    ControlTrigger Stringa
    If mLogFile <> 0 Then
        mLogLastLine = Stringa
    End If
    
    mConnect.Envi.sendOutput CleanString(Stringa), TOUT_LASTLINE
End Sub

Private Function CleanString(data As String) As String
    Dim iniPos As Integer, FinPos As Integer
    Dim lenght As Integer
    Dim i As Integer, iChar As Integer, Char As String
    Dim AnsiCode As String

    If InStr(1, data, "ÿû") Then data = Replace(data, "ÿû", "")
    If InStr(1, data, "ÿü") Then data = Replace(data, "ÿü", "")
    lenght = Len(data)
    iniPos = 1
    Do
        iniPos = InStr(iniPos, data, ESCCHAR)
        If Not iniPos = 0 Then
            
            For i = iniPos + 1 To lenght
                Char = Mid$(data, i, 1)
                iChar = Asc(LCase$(Char))
                If iChar >= 97 And iChar <= 122 Then
                    FinPos = i
                    Exit For
                End If
            Next i
            
            If i > lenght Then FinPos = 0
            
            If FinPos <> 0 Then
                AnsiCode = Mid$(data, iniPos + 2, FinPos - iniPos - 2)
                data = Left$(data, iniPos - 1) & _
                    Mid$(data, FinPos + 1)
                iniPos = FinPos - (Len(AnsiCode) + 3)
                If iniPos = 0 Then iniPos = 1
                
                lenght = Len(data)
            Else
                iniPos = 0
            End If
        End If
    Loop Until iniPos = 0
    
    CleanString = data
End Function

Private Sub ControlTrigger(ByVal Stringa As String)
    Dim i As Integer ', j As Integer, ToExec As String

    Stringa = CleanString(Stringa)
    mConnect.Envi.sendOutput Stringa, TOUT_CLEAN
    If Not Stringa = "" Then
        For i = 1 To mTriggers.Count
            'If InStr(1, Stringa, mTriggers.Text(i), vbTextCompare) <> 0 Then
            '    SendTrigger i
            'End If
            If mTriggers.Item(i).Check(Stringa) Then mTriggers.Item(i).Execute
        Next i
    End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'MsgBox "errore " & Number & ": " & Description
    'mConnect.Envi.sendOutput "Source = Socket | " & Number & " = " & Description, TOUT_ERROR
    mConnect.LogError Number, Description, "Socket"
    'mConnect.Envi.sendOutput "Sessione " & mNSess & " terminata", TOUT_SERVICE
    mConnect.LogYellow "Session " & mNSess & " terminated"
    mSockError = True
    Winsock.Close
    ConnectCmd True
    mConnect.Envi.sendNotify ENVM_CLOSE
    mConnect.Envi.sendNotify ENVM_SWITCHTOSTATUS
End Sub
