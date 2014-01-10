VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferenze"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Annulla"
      Default         =   -1  'True
      Height          =   315
      Left            =   5475
      TabIndex        =   9
      Top             =   3300
      Width           =   1290
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   4125
      TabIndex        =   8
      Top             =   3300
      Width           =   1290
   End
   Begin VB.Frame fraSects 
      Caption         =   "Sezioni"
      Height          =   3240
      Left            =   75
      TabIndex        =   7
      Top             =   0
      Width           =   1815
      Begin VB.ListBox lstSects 
         Height          =   2925
         IntegralHeight  =   0   'False
         Left            =   75
         TabIndex        =   10
         Top             =   225
         Width           =   1665
      End
   End
   Begin VB.PictureBox pctSect 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   2940
      Index           =   1
      Left            =   2025
      ScaleHeight     =   2940
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CheckBox chkVarParse 
         Caption         =   "Disattiva il Parser delle Variabili"
         Height          =   315
         Left            =   150
         TabIndex        =   34
         Top             =   2100
         Width           =   3540
      End
      Begin VB.CommandButton cmdChangeFont 
         Caption         =   "Change"
         Height          =   240
         Left            =   3675
         TabIndex        =   29
         Top             =   2550
         Width           =   915
      End
      Begin VB.TextBox txtInputSep 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         MaxLength       =   1
         TabIndex        =   26
         Text            =   ";"
         Top             =   1830
         Width           =   465
      End
      Begin VB.CheckBox chkMultInput 
         Caption         =   "Permetti più di un comando sulla stessa riga"
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   1575
         Width           =   3540
      End
      Begin VB.CheckBox chkLocalEcho 
         Caption         =   "Eco locale"
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   1275
         Width           =   2790
      End
      Begin VB.CheckBox chkErasePrompt 
         Caption         =   "Cancella automaticamente il prompt dei comandi"
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   975
         Width           =   3990
      End
      Begin VB.OptionButton optDisk 
         Caption         =   "Su disco (infinito)"
         Height          =   315
         Left            =   450
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optMemory 
         Caption         =   "In memoria (1000 righe)"
         Height          =   315
         Left            =   450
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.Label lblChangeFont 
         Caption         =   "Font:"
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   2550
         Width           =   915
      End
      Begin VB.Label lblSeparator 
         Caption         =   "Separatore comandi:"
         Height          =   240
         Left            =   2325
         TabIndex        =   25
         Top             =   1875
         Width           =   1590
      End
      Begin VB.Label lblBuffer 
         Caption         =   "Tipo di buffer:"
         Height          =   240
         Left            =   150
         TabIndex        =   1
         Top             =   75
         Width           =   1815
      End
      Begin VB.Label lblFont 
         Caption         =   "Courier 10"
         Height          =   240
         Left            =   1125
         TabIndex        =   28
         Top             =   2550
         Width           =   2565
      End
   End
   Begin VB.PictureBox pctSect 
      BorderStyle     =   0  'None
      Height          =   2940
      Index           =   3
      Left            =   2025
      ScaleHeight     =   2940
      ScaleWidth      =   4665
      TabIndex        =   15
      Top             =   225
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CommandButton cmdLogCust 
         Caption         =   "Invia comando"
         Height          =   315
         Left            =   2475
         TabIndex        =   23
         Top             =   2175
         Width           =   2040
      End
      Begin VB.CommandButton cmdLogEnter 
         Caption         =   "Premi Invio"
         Height          =   315
         Left            =   2475
         TabIndex        =   22
         Top             =   1800
         Width           =   2040
      End
      Begin VB.CommandButton cmdLogPass 
         Caption         =   "Invia password"
         Height          =   315
         Left            =   2475
         TabIndex        =   21
         Top             =   1425
         Width           =   2040
      End
      Begin VB.CommandButton cmdLogReset 
         Caption         =   "Ricomincia"
         Height          =   315
         Left            =   2475
         TabIndex        =   20
         Top             =   675
         Width           =   2040
      End
      Begin VB.CommandButton cmdLogUser 
         Caption         =   "Invia nickname"
         Height          =   315
         Left            =   2475
         TabIndex        =   19
         Top             =   1050
         Width           =   2040
      End
      Begin VB.ListBox lstLogin 
         Height          =   1815
         Left            =   375
         TabIndex        =   18
         Top             =   675
         Width           =   2040
      End
      Begin VB.OptionButton optLoginAuto 
         Caption         =   "Esegui l'autologin secondo questa procedura:"
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   375
         Width           =   3690
      End
      Begin VB.OptionButton optLoginNo 
         Caption         =   "Non eseguire l'autologin"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   75
         Width           =   2865
      End
   End
   Begin VB.PictureBox pctSect 
      BorderStyle     =   0  'None
      Height          =   2940
      Index           =   2
      Left            =   2025
      ScaleHeight     =   2940
      ScaleWidth      =   4665
      TabIndex        =   11
      Top             =   225
      Visible         =   0   'False
      Width           =   4665
      Begin VB.ComboBox cboMuds 
         Height          =   315
         Left            =   675
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   675
         Width           =   3915
      End
      Begin VB.OptionButton optStartAuto 
         Caption         =   "Avvia automaticamente con questo mud:"
         Height          =   240
         Left            =   150
         TabIndex        =   13
         Top             =   375
         Width           =   3465
      End
      Begin VB.OptionButton optStartList 
         Caption         =   "Mostra la lista dei MUD"
         Height          =   240
         Left            =   150
         TabIndex        =   12
         Top             =   75
         Width           =   2865
      End
   End
   Begin VB.PictureBox pctSect 
      BorderStyle     =   0  'None
      Height          =   2940
      Index           =   4
      Left            =   2025
      ScaleHeight     =   2940
      ScaleWidth      =   4665
      TabIndex        =   30
      Top             =   225
      Visible         =   0   'False
      Width           =   4665
      Begin VB.OptionButton optLogHtml 
         Caption         =   "Html"
         Enabled         =   0   'False
         Height          =   240
         Left            =   375
         TabIndex        =   33
         Top             =   750
         Width           =   2115
      End
      Begin VB.OptionButton optLogText 
         Caption         =   "Plane text"
         Enabled         =   0   'False
         Height          =   240
         Left            =   375
         TabIndex        =   32
         Top             =   450
         Width           =   2115
      End
      Begin VB.CheckBox chkMudLog 
         Caption         =   "Scrivi Log del MUD"
         Height          =   315
         Left            =   150
         TabIndex        =   31
         Top             =   75
         Width           =   2790
      End
   End
   Begin VB.Frame fraSect 
      Caption         =   "Generale"
      Height          =   3240
      Left            =   1950
      TabIndex        =   6
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private mFontName As String
Private mFontSize As Integer

Private Sub LoadLang()
    Me.Caption = mLang("config", "caption")
    fraSects.Caption = mLang("config", "sections")
    cmdOk.Caption = mLang("", "Ok")
    cmdAbort.Caption = mLang("", "Cancel")
    
    'general sheet
    lblBuffer.Caption = mLang("config", "BufferType")
    optMemory.Caption = mLang("config", "BufferMemory")
    optDisk.Caption = mLang("config", "BufferDisk")
    chkMudLog.Caption = mLang("config", "WriteLog")
    chkErasePrompt.Caption = mLang("config", "ErasePrompt")
    chkLocalEcho.Caption = mLang("config", "LocalEcho")
    chkMultInput.Caption = mLang("config", "MultipleCommands")
    lblSeparator.Caption = mLang("config", "MultipleSeparator")
    lblChangeFont.Caption = mLang("config", "Font")
    cmdChangeFont.Caption = mLang("config", "Change")
    chkVarParse.Caption = mLang("config", "DisableVarParser")
    
    'rubrica sheet
    'chkChatSave.Caption = mLang("config", "RubricaSave")
    'chkChatLoad.Caption = mLang("config", "RubricaLoad")
    
    'startup sheet
    optStartList.Caption = mLang("config", "StartList")
    optStartAuto.Caption = mLang("config", "StartAuto")
    
    'autologin sheet
    optLoginAuto.Caption = mLang("config", "AutologinYes")
    optLoginNo.Caption = mLang("config", "AutologinNo")
    cmdLogReset.Caption = mLang("", "Reset")
    cmdLogUser.Caption = mLang("config", "SendNick")
    cmdLogPass.Caption = mLang("config", "SendPass")
    cmdLogEnter.Caption = mLang("config", "SendEnter")
    cmdLogCust.Caption = mLang("config", "SendCommand")
    
    'log sheet
    optLogText.Caption = mLang("config", "LogPlain")
    optLogHtml.Caption = mLang("config", "LogHtml")
End Sub


Private Sub chkMudLog_Click()
    optLogText.Enabled = chkMudLog.Value
    optLogHtml.Enabled = chkMudLog.Value
End Sub

Private Sub chkMultInput_Click()
    If chkMultInput.Value = vbChecked Then
        txtInputSep.Enabled = True
        If Me.Visible Then txtInputSep.SetFocus
    Else
        txtInputSep.Enabled = False
    End If
End Sub

Private Sub cmdAbort_Click()
    Unload Me
    Set frmConfig = Nothing
End Sub

Private Sub cmdChangeFont_Click()
    Dim font As cCommonDialog
    
    Set font = New cCommonDialog
    font.FontName = mFontName
    font.FontSize = mFontSize
    font.Flags = cdlCFScreenFonts Or cdlCFFixedPitchOnly Or cdlCFInitToLogFontStruct
    Set font.Parent = Me
    If (font.ShowFont) Then
        If Not font.FontName = "" Then
            mFontName = font.FontName
            mFontSize = font.FontSize
            lblFont.Caption = mFontName & " " & mFontSize
        End If
    End If
End Sub

Private Sub cmdLogCust_Click()
    Dim exp As String
    exp = InputBox(mLang("config", "CommandToSend"))
    If Not exp = "" Then LoginAppend exp
End Sub

Private Sub cmdLogEnter_Click()
    LoginAppend ""
End Sub

Private Sub cmdLogPass_Click()
    LoginAppend "Password"
End Sub

Private Sub cmdLogReset_Click()
    lstLogin.Clear
End Sub

Private Sub LoginAppend(ByVal data As String)
    lstLogin.AddItem data & "<" & LCase$(mLang("config", "Enter")) & ">"
    lstLogin.ListIndex = lstLogin.ListCount - 1
End Sub

Private Sub cmdLogUser_Click()
    LoginAppend "Nickname"
End Sub

Private Sub cmdOk_Click()
    SaveSettings
    Unload Me
    Set frmConfig = Nothing
End Sub

Private Sub LoadLogin(Config As cConnector)
    Dim Count As Integer
    Dim i As Integer
    
    Count = Config.GetConfig("Login_Count", 0)
    If Count = 0 Then
        optLoginNo.Value = True
    Else
        optLoginAuto.Value = True
        For i = 1 To Count
            LoginAppend Config.GetConfig("Login<" & i & ">", "")
        Next i
    End If
End Sub

Private Sub SaveLogin(Config As cConnector)
    Dim Count As Integer
    Dim i As Integer
    
    If optLoginNo.Value = True Then
        Count = 0
    Else
        Count = lstLogin.ListCount
        For i = 1 To Count
            Config.SetConfig "Login<" & i & ">", Trim$(Left$(lstLogin.list(i - 1), Len(lstLogin.list(i - 1)) - 7))
        Next i
    End If
    Config.SetConfig "Login_Count", Count
End Sub

Private Sub LoadSettings()
    Dim Config As cConnector
    Dim IniConf As cIni
    Dim Muds As cMuds, i As Integer
    Dim MudSel As String

    Set Config = New cConnector
        'general
        optDisk.Value = Config.GetBoolConfig("DiskBuffer")
        chkMudLog.Value = Abs(Config.GetConfig("Logging", 0))
        chkErasePrompt.Value = Abs(Config.GetConfig("ErasePrompt", -1))
        chkLocalEcho.Value = Abs(Config.GetConfig("LocalEcho", -1))
        'multiple input, true by default
        chkMultInput.Value = Abs(Config.GetConfig("MultipleInput", -1))
        'multiple input separator, is ";" for default
        txtInputSep.Text = Config.GetConfig("MultipleInputSep", ";")
        chkVarParse.Value = Abs(Config.GetConfig("DisableVarParser", 0))
    
        'autologin
        LoadLogin Config
    
        'log
        If Config.GetBoolConfig("LogHtml", True) Then
            optLogHtml.Value = True
        Else
            optLogText.Value = True
        End If
    Set Config = Nothing
    
    'startup
    Set IniConf = New cIni
        IniConf.CaricaFile App.Path & "\config.ini", True
        If IniConf.RetrInfo("startup", 0) = STARTMODE_LIST Then
            optStartList.Value = True
        Else
            optStartAuto.Value = True
        End If
        Set Muds = New cMuds
            Muds.LoadMudList
            For i = 1 To Muds.Count
                cboMuds.AddItem Muds.Name(i)
            Next i
        Set Muds = Nothing
        MudSel = IniConf.RetrInfo("mud", "")
        If Not MudSel = "" Then
            For i = 0 To cboMuds.ListCount - 1
                If cboMuds.list(i) = MudSel Then
                    cboMuds.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
        'general - font size
        mFontSize = IniConf.RetrInfo("FontSize", 10)
        mFontName = IniConf.RetrInfo("FontName", "Courier")
        lblFont.Caption = mFontName & " " & mFontSize
    Set IniConf = Nothing
            
End Sub

Private Sub SaveSettings()
    Dim Config As cConnector
    Dim IniConf As cIni

    Set Config = New cConnector
        'general
        Config.SetBoolConfig "DiskBuffer", optDisk.Value
        Config.SetConfig "Logging", chkMudLog.Value
        Config.SetConfig "ErasePrompt", chkErasePrompt.Value
        Config.SetConfig "LocalEcho", chkLocalEcho.Value
        
        Config.SetConfig "MultipleInput", chkMultInput.Value
        Config.SetConfig "MultipleInputSep", txtInputSep.Text
    
        Config.SetConfig "DisableVarParser", chkVarParse.Value
    
        'rubrica
        'Config.SetConfig "ChatLoadContacts", chkChatLoad.Value
        'Config.SetConfig "ChatSaveContacts", chkChatSave.Value
        

        'startup
        Set IniConf = New cIni
            IniConf.CaricaFile App.Path & "\config.ini", True
            If optStartList.Value Then
                IniConf.AddInfo "startup", STARTMODE_LIST
            Else
                IniConf.AddInfo "startup", STARTMODE_AUTO
            End If
            IniConf.AddInfo "mud", cboMuds.Text
            
            'general - font size
            IniConf.AddInfo "FontSize", mFontSize
            IniConf.AddInfo "FontName", mFontName
    
            IniConf.SalvaFile
        Set IniConf = Nothing

        'autologin
        SaveLogin Config

        'log
        Config.SetBoolConfig "LogHtml", optLogHtml.Value
        Config.SaveConfig
        

        Config.Envi.sendNotify ENVM_CONFIGCHANGED
    Set Config = Nothing
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    lstSects.AddItem mLang("config", "General")     'generale
    lstSects.AddItem mLang("config", "Startup")     'avvio
    lstSects.AddItem mLang("config", "Autologin")   'autologin
    lstSects.AddItem mLang("config", "Log")         'log
    
    'cboFontSize.AddItem 10
    'cboFontSize.AddItem 12
    'cboFontSize.AddItem 15
    
    LoadSettings
    lstSects.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mLang = Nothing
End Sub

Private Sub lstSects_Click()
    Dim Index As Integer
    
    Index = lstSects.ListIndex + 1
    fraSect.Caption = lstSects.Text
    pctSect(Index).Visible = True
    pctSect(Index).ZOrder 0
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub optLoginAuto_Click()
    If optLoginAuto.Value Then
        lstLogin.Enabled = True
        cmdLogReset.Enabled = True
        cmdLogUser.Enabled = True
        cmdLogPass.Enabled = True
        cmdLogEnter.Enabled = True
        cmdLogCust.Enabled = True
    End If
End Sub

Private Sub optLoginNo_Click()
    If optLoginNo.Value Then
        lstLogin.Enabled = False
        cmdLogReset.Enabled = False
        cmdLogUser.Enabled = False
        cmdLogPass.Enabled = False
        cmdLogEnter.Enabled = False
        cmdLogCust.Enabled = False
    End If
End Sub

Private Sub optStartAuto_Click()
    If optStartAuto.Value Then
        cboMuds.Enabled = True
    End If
End Sub

Private Sub optStartList_Click()
    If optStartList.Value Then
        cboMuds.Enabled = False
    End If
End Sub

Private Sub txtInputSep_GotFocus()
    txtInputSep.SelStart = 0
    txtInputSep.SelLength = Len(txtInputSep.Text)
End Sub
