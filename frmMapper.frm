VERSION 5.00
Begin VB.Form frmMapper 
   AutoRedraw      =   -1  'True
   Caption         =   "Mapper"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6480
   Icon            =   "frmMapper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   StartUpPosition =   3  'Windows Default
   Begin GoS.uMapper map 
      Height          =   3315
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5847
   End
   Begin VB.Label lblTitle 
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   3750
      Width           =   6315
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Mapper"
      Visible         =   0   'False
      Begin VB.Menu mnuMap 
         Caption         =   "Mappa"
         Begin VB.Menu mnuNuovo 
            Caption         =   "Inizia nuova"
         End
         Begin VB.Menu mnuCarica 
            Caption         =   "Carica"
         End
         Begin VB.Menu mnuSalva 
            Caption         =   "Salva"
         End
      End
      Begin VB.Menu mnuModee 
         Caption         =   "Modalità"
         Begin VB.Menu mnuOnline 
            Caption         =   "Online (automatico)"
         End
         Begin VB.Menu mnuOffline 
            Caption         =   "Offline (manuale)"
         End
         Begin VB.Menu mnuPausa 
            Caption         =   "Pausa"
         End
         Begin VB.Menu sep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFollow 
            Caption         =   "Segui spostamenti"
         End
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetOrg 
         Caption         =   "Posiziona"
      End
   End
End
Attribute VB_Name = "frmMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents mParser As cParser

Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1

Private mSpost As String
Private mMov As Mapper_Mov

'Private mnuMappa As cMenu
'Private mnuMode As cMenu
Private mFileName As String

Private mPausa As Integer, mRecTitle As Boolean
Private mLocalEcho As Boolean

Private Sub VerifyConfig()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        mLocalEcho = Connect.GetBoolConfig("LocalEcho", False)
    Set Connect = Nothing
End Sub

Private Sub AggiornaTitolo()
    If mFileName = "" Then
        SetWindowText Me.hWnd, "Mapper (nuova mappa)"
    Else
        SetWindowText Me.hWnd, "Mapper [" & Left$(mFileName, Len(mFileName) - 4) & "]"
    End If
End Sub

Private Function LoadMap() As Boolean
    Dim Nomefile As String

    ControllaModifiche
    Nomefile = map.LoadMap
    If Nomefile = "" Then
        LoadMap = False
    Else
        mFileName = Nomefile
        AggiornaTitolo
        LoadMap = True
    End If
End Function

Private Sub NewMap()
    ControllaModifiche
    map.BeginMap
    mFileName = ""
    AggiornaTitolo
End Sub

Private Sub map_ContextMenu()
    PopupMenu mnuPopup
End Sub

Private Sub Form_Load()
    'Dim Config As cIni
    Dim Connect As cConnector

    goshSetDockable Me.hWnd, "gos.mapper"
    
    Set mFine = New cFinestra
    Me.Hide
    
    Set Connect = New cConnector
        'Config.CaricaFile "config.ini"
        'map.Follow = CBool(Val(Config.RetrInfo("Mapper_Follow", 0)))
        map.Follow = Connect.GetBoolConfig("Mapper_Follow")
    Set Connect = Nothing
    
    'mnuMode.Add "Segui spostamenti", "cmdFollow", map.Follow
    'mnuMode.Draw Menus, "mnuMode"
    mnuFollow.Checked = map.Follow
    
    VerifyConfig
    
    map.MapMode = MAPMODE_ONLINE
    
    mFine.Init Me, WINREC_INPUT Or WINREC_OUTPUT
    
    NewMap
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ControllaModifiche
    
    'Set mnuMappa = Nothing
    'Set mnuMode = Nothing
    mFine.UnReg
    Set mFine = Nothing
    'mParser.MapperActive = False
    'Set mParser = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    map.Width = Me.ScaleWidth - map.Left * 2
    map.Height = Me.ScaleHeight - map.Top - lblTitle.Height - 3

    lblTitle.Top = map.Top + map.Height
    lblTitle.Width = map.Width

    'Menus.Width = map.Width
    'mFine.Cls
    'Menus.Draw
    map.DrawBorder
End Sub

Private Sub map_RoomTitle(Title As String)
    If Title = "" Then Title = "PARTENZA"
    
    lblTitle.Caption = Title
End Sub

Private Sub map_Send(Stringa As String)
    Dim Connect As cConnector

    Set Connect = New cConnector
    Connect.Envi.sendInput Stringa
    Set Connect = Nothing
    'mParser.Send Stringa
End Sub

Private Sub ControllaModifiche()
    Dim Mex As String

    If map.Modified Then
        If mFileName = "" Then
            Mex = "Salvare le modifiche apportate?"
        Else
            Mex = "Salvare le mofiche apportate a '" & Left$(mFileName, Len(mFileName) - 4) & "' ?"
        End If
        
        If MsgBox(Mex, vbYesNo, "GoS Mapper") = vbYes Then
            SaveMap
        End If
    End If
End Sub

Private Sub SaveMap()
    mFileName = map.SaveMap(mFileName)
    AggiornaTitolo
End Sub

Private Sub ProcessaTitolo(ByVal Response As String)
    If Not (InStr(1, Response, "non puoi andare da quella parte", vbTextCompare) <> 0 Or _
       InStr(1, Response, "e' chius", vbTextCompare) <> 0 Or _
       InStr(1, Response, "stai gia' combattendo", vbTextCompare) <> 0 Or _
       InStr(1, Response, "hai bisogno di una barca", vbTextCompare) <> 0 Or _
       InStr(1, Response, "adesso sei troppo rilassa", vbTextCompare) <> 0 Or _
       InStr(1, Response, "nei tuoi sogni", vbTextCompare) <> 0 Or _
       InStr(1, Response, "dovresti saper nuotare", vbTextCompare) <> 0) Then
        'lstMov.AddItem mSpost & " --> " & Response
        'lstMov.ListIndex = lstMov.ListCount - 1
        map.AddRoom mMov, Response
    End If
End Sub

Private Sub mFine_envInput(data As String, InType As Integer)
    Dim Continua As Boolean

    'esegui il controllo solo sui comandi realmente inviati!
    If InType = TIN_SENT And map.MapMode = MAPMODE_ONLINE Then
        'un bel select sull'invio
        data = LCase$(Trim$(data))
        If Right$(data, 2) = vbCrLf Then data = Left$(data, Len(data) - 2)
        Continua = True
        Select Case data
            Case "n", "no", "nor", "nord"
                mSpost = "Nord"
                mMov = nord
            Case "s", "su", "sud"
                mSpost = "Sud"
                mMov = sud
            Case "e", "es", "est"
                mSpost = "Est"
                mMov = est
            Case "w", "we", "wes", "west", "o", "ov", "ove", "oves", "ovest"
                mSpost = "Ovest"
                mMov = ovest
            Case "a", "al", "alt", "alto"
                mSpost = "Alto"
                mMov = alto
            Case "b", "ba", "bas", "bass", "basso"
                mSpost = "Basso"
                mMov = basso
            Case Else
                Continua = False
        End Select
    
        If Continua Then
            If mLocalEcho Then
                mPausa = 0
            Else
                mPausa = 1
            End If
            mRecTitle = True
        End If
    End If
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    If uMsg = ENVM_CONFIGCHANGED Then
        VerifyConfig
    End If
End Sub

Private Sub mFine_envOutput(data As String, OutType As Integer)
    'controlla la questione soltanto nell'output pulito dalle sequenze ansi
    If OutType = TOUT_CLEAN And mRecTitle Then
        If Not Trim$(data) = "" Then
            mPausa = mPausa + 1
            If InStr(1, data, "sei affaticato", vbTextCompare) <> 0 Then mPausa = mPausa - 1
            'If InStr(1, data, "!!MUSIC", vbTextCompare) <> 0 Then mPausa = mPausa - 1
            If mPausa = 2 Then
                mRecTitle = False
                ProcessaTitolo data
            End If
        End If
    End If
End Sub

Private Sub mnuCarica_Click()
    LoadMap
End Sub

Private Sub mnuFollow_Click()
    'Dim Config As cIni
    Dim Connect As cConnector

    'If mnuMode.Checked(Index) Then
    '    mnuMode.Checked(Index) = False
    'Else
    '    mnuMode.Checked(Index) = True
    'End If
    mnuFollow.Checked = Not mnuFollow.Checked
    map.Follow = mnuFollow.Checked
    'Set Config = New cIni
    Set Connect = New cConnector
        'Config.CaricaFile "config.ini"
        'map.Follow = CBool(Val(Config.RetrInfo("Mapper_Follow", 0)))
        'Config.AddInfo "Mapper_Follow", CInt(map.Follow)
        Connect.SetBoolConfig "Mapper_Follow", map.Follow
        Connect.SaveConfig
    Set Connect = Nothing
    'Set Config = Nothing
End Sub

Private Sub mnuNuovo_Click()
    NewMap
End Sub

Private Sub mnuOffline_Click()
    If Not mnuOffline.Checked Then
        mnuPausa.Checked = False
        mnuOffline.Checked = True
        mnuOnline.Checked = False
        map.MapMode = MAPMODE_OFFLINE
    End If
End Sub

Private Sub mnuOnline_Click()
    If Not mnuOnline.Checked Then
        mnuPausa.Checked = False
        mnuOffline.Checked = False
        mnuOnline.Checked = True
        map.MapMode = MAPMODE_ONLINE
    End If
End Sub

Private Sub mnuPausa_Click()
    If Not mnuPausa.Checked Then
        mnuPausa.Checked = True
        mnuOffline.Checked = False
        mnuOnline.Checked = False
        map.MapMode = MAPMODE_PAUSE
    End If
End Sub

Private Sub mnuSalva_Click()
    SaveMap
End Sub

Private Sub mnuSetOrg_Click()
    map.SetOrg
End Sub
