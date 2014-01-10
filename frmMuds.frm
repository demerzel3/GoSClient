VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMuds 
   Caption         =   "GosClient - Elenco MUD"
   ClientHeight    =   5040
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9150
   Icon            =   "frmMuds.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwMuds 
      Height          =   4215
      Left            =   2855
      TabIndex        =   2
      Top             =   0
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "GosClient"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   375
      Width           =   2115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   150
      X2              =   2625
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3390
      Left            =   225
      TabIndex        =   1
      Top             =   750
      Width           =   2340
   End
   Begin VB.Image imgLogo 
      Height          =   1080
      Left            =   75
      Picture         =   "frmMuds.frx":00D2
      Top             =   75
      Width           =   1050
   End
   Begin VB.Shape shpLogo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   4290
      Left            =   30
      Top             =   0
      Width           =   2770
   End
   Begin VB.Menu mnuMud 
      Caption         =   "MUD"
      Begin VB.Menu mnuAdd 
         Caption         =   "Aggiungi nuovo MUD"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Proprietà"
      End
      Begin VB.Menu mnuDupl 
         Caption         =   "Duplica"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Apri"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Rimuovi"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu mnuLangs 
      Caption         =   "Language"
      Begin VB.Menu mnuLang 
         Caption         =   "lang0"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuHomePage 
         Caption         =   "Home Page"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "E-Mail"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMuds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMuds As cMuds
Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private mCloseClient As Boolean

Private Sub LoadLang()
    Dim Index As Integer
    
    If Not lvwMuds.SelectedItem Is Nothing Then
        Index = GetKeyIndex(lvwMuds.SelectedItem.Key)
        FillInfoLabel Index
    End If
    
    'column headers
    lvwMuds.ColumnHeaders(1).Text = mLang("general", "Name")
    lvwMuds.ColumnHeaders(2).Text = mLang("general", "Host")
    lvwMuds.ColumnHeaders(3).Text = mLang("general", "Port")

    'mud menu
    mnuAdd.Caption = mLang("muds", "Add")
    mnuProp.Caption = mLang("muds", "Properties")
    mnuDupl.Caption = mLang("muds", "Duplicate")
    mnuOpen.Caption = mLang("general", "Open")
    mnuDelete.Caption = mLang("general", "Remove")
    mnuExit.Caption = mLang("general", "Exit")
    
    'langmenu
    mnuLangs.Caption = mLang("general", "Language")
    
    'help menu
    mnuHomePage.Caption = mLang("help", "URL")
    mnuMail.Caption = mLang("help", "EMail")
    mnuAbout.Caption = mLang("help", "About")
    
    'caption
    Me.Caption = "GoSClient - " & mLang("muds", "Caption")
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

Public Sub Init(Optional ForceList As Boolean = False)
    Dim Conf As cIni
    Dim ShowList As Boolean
    Dim Mud As String, i As Integer
    Dim Found As Boolean
    Dim winState As Integer
    
    Set mMuds = New cMuds
    mMuds.LoadMudList

    Set Conf = New cIni
        Conf.CaricaFile App.Path & "\config.ini", True
        
        winState = Conf.RetrInfo("MudsWnd_State", Me.WindowState)
        If Not winState = vbMaximized Then
            Me.Left = Conf.RetrInfo("MudsWnd_Left", Me.Left)
            Me.Width = Conf.RetrInfo("MudsWnd_Width", Me.Width)
            Me.Top = Conf.RetrInfo("MudsWnd_Top", Me.Top)
            Me.Height = Conf.RetrInfo("MudsWnd_Height", Me.Height)
        End If
        Me.WindowState = winState
        
        LoadList
        If Not ForceList Then
            If Conf.RetrInfo("startup", 0) = STARTMODE_LIST Then
                ShowList = True
            End If
    
            If ShowList Then
                Me.Show
            Else
                Mud = Conf.RetrInfo("mud", "")
                If Mud = "" Then
                    Me.Show
                Else
                    Found = False
                    With lvwMuds.ListItems
                        For i = 1 To .Count
                            If .Item(i).Text = Mud Then
                                .Item(i).Selected = True
                                Found = True
                                Exit For
                            End If
                        Next i
                        If Found Then OpenClient
                    End With
                    If Found Then
                        Unload Me
                        Set frmMuds = Nothing
                    Else
                        Me.Show
                    End If
                End If
            End If
        Else
            Me.Show
        End If
    Set Conf = Nothing
End Sub

Private Sub Form_Load()
    Dim lvwWidth As Long
    Dim Conf As cIni
       
    mCloseClient = True
       
    Me.Hide
    lblInfo.ForeColor = rgb(150, 150, 150)
    lvwWidth = lvwMuds.Width
    lvwMuds.ColumnHeaders.Add , , "Name", (lvwWidth * 50) / 100
    lvwMuds.ColumnHeaders.Add , , "Host", (lvwWidth * 30) / 100
    lvwMuds.ColumnHeaders.Add , , "Port", (lvwWidth * 15) / 100
    lvwMuds.SortKey = 0
    
    Set Conf = New cIni
        Conf.CaricaFile App.Path & "\config.ini", True
        Set mLang = New cLang
        If Not mLang.LoadLang(Conf.RetrInfo("lang", "english.lng")) Then
            'Me.Caption = "GosClient - " & mLang("muds", "Caption")
        'Else
            Me.Caption = "GosClient - <unable to load language file>"
        End If
    Set Conf = Nothing
    
    LoadLangList
End Sub

Private Sub LoadList()
    Dim i As Integer, Item As ListItem
    
    lvwMuds.ListItems.Clear
    For i = 1 To mMuds.Count
        Set Item = lvwMuds.ListItems.Add(, "i" & i, mMuds.Name(i))
            'item.Bold = True
            Item.SubItems(1) = mMuds.Host(i)
            Item.SubItems(2) = mMuds.Port(i)
        Set Item = Nothing
    Next i
    
    If lvwMuds.ListItems.Count > 0 Then lvwMuds.ListItems.Item(1).Selected = True
    If Not lvwMuds.SelectedItem Is Nothing Then FillInfoLabel GetKeyIndex(lvwMuds.SelectedItem.Key)
End Sub

Private Sub FillInfoLabel(id As Integer)
    lblInfo.Caption = _
      "MUD: " & mMuds.Name(id) & vbCrLf & vbCrLf & _
      mLang("", "Host") & ": " & mMuds.Host(id) & vbCrLf & vbCrLf & _
      mLang("", "Port") & ": " & mMuds.Port(id) & vbCrLf & vbCrLf & _
      mLang("", "Description") & ": " & mMuds.Descr(id) & vbCrLf & vbCrLf & _
      mLang("", "Language") & ": " & mMuds.Lang(id) & vbCrLf & vbCrLf & _
      mLang("", "Comment") & ": " & mMuds.Comment(id)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Conf As cIni
    
    If mCloseClient Then SaveSetting "GosClient", "multiple", "port" & gPluginsPort, 0
    
    Set Conf = New cIni
        Conf.CaricaFile App.Path & "/config.ini", True
        Call Conf.AddInfo("MudsWnd_Left", Me.Left)
        Call Conf.AddInfo("MudsWnd_Width", Me.Width)
        Call Conf.AddInfo("MudsWnd_Top", Me.Top)
        Call Conf.AddInfo("MudsWnd_Height", Me.Height)
        Call Conf.AddInfo("MudsWnd_State", Me.WindowState)
        Conf.SalvaFile
    Set Conf = Nothing

    Set mLang = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lvwMuds.Height = Me.ScaleHeight - 50
    lvwMuds.Width = Me.ScaleWidth - lvwMuds.Left - 50
    shpLogo.Height = lvwMuds.Height
    lblInfo.Height = Me.ScaleHeight - lblInfo.Top - 75

    If Err.Number = 380 Then 'errore nel ridimensionamento
        Err.Clear
    End If
End Sub

Private Function GetKeyIndex(Key As String) As Integer
    GetKeyIndex = Val(Mid$(Key, 2))
End Function

Private Sub OpenClient()
    Dim Connect As cConnector
    Dim Index As Integer
    
    mCloseClient = False
    
    Index = GetKeyIndex(lvwMuds.SelectedItem.Key)
    gMudPath = App.Path & "\" & mMuds.Folder(Index) & "\"
    If Dir$(gMudPath, vbDirectory) = "" Then MkDir gMudPath
    
    'creazione di gEnvi, caricamento info iniziali
    Set Connect = New cConnector
        Set Connect.Envi.Mud = mMuds.Mud(Index)
        frmBase.Show
    Set Connect = Nothing
    
    Unload Me
    Set frmMuds = Nothing
End Sub

Private Sub lvwMuds_DblClick()
    OpenClient
End Sub

Private Sub lvwMuds_ItemClick(ByVal Item As ComctlLib.ListItem)
    FillInfoLabel GetKeyIndex(Item.Key)
End Sub

Private Sub lvwMuds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuMud
    End If
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub mnuAbout_Click()
    frmCredits.Show vbModal
End Sub

Private Sub mnuAdd_Click()
    Dim Item As ComctlLib.ListItem
    If frmDefMud.NewMud(mMuds, mLang) Then
        With mMuds.Mud(mMuds.Count)
            Set Item = lvwMuds.ListItems.Add(, "i" & mMuds.Count, .Name)
                Item.SubItems(1) = .Host
                Item.SubItems(2) = .Port
                Item.Selected = True
                FillInfoLabel mMuds.Count
            Set Item = Nothing
        End With
        'LoadList
        mMuds.SaveMudList
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim rtn As VbMsgBoxResult
    Dim Index As Integer
    Dim fs As cFileSystem
    
    'rtn = MsgBox("Rimuovere " & lvwMuds.SelectedItem.Text & " dalla lista?", vbYesNo)
    rtn = MsgBox(mLang("", "Remove") & " " & lvwMuds.SelectedItem.Text & "?", vbYesNo)
    If rtn = vbYes Then
        rtn = MsgBox(mLang("muds", "RemoveData"), vbYesNo)
        Index = GetKeyIndex(lvwMuds.SelectedItem.Key)
        If rtn = vbYes Then
            Set fs = New cFileSystem
                If fs.DirExist(App.Path & "\" & mMuds.Folder(Index)) Then
                    fs.DirDelete App.Path & "\" & mMuds.Folder(Index)
                End If
            Set fs = Nothing
        End If
        lvwMuds.ListItems.Remove lvwMuds.SelectedItem.Index
        mMuds.Remove Index
        LoadList
        
        mMuds.SaveMudList
    End If
End Sub

Private Sub mnuDupl_Click()
    Dim ToCopy As cMud
    Dim Index As Integer
    Dim NewName As String
    Dim fs As cFileSystem
    Dim Item As ComctlLib.ListItem
    
    Index = GetKeyIndex(lvwMuds.SelectedItem.Key)
    Set ToCopy = mMuds.Mud(Index)
    NewName = InputBox(mLang("muds", "NewName"), , ToCopy.Name & " 2")
    If (Not NewName = "") And (NewName <> ToCopy.Name) Then
        With ToCopy
            mMuds.Add NewName, .Host, .Port, .Descr, .Lang, .Comment
            Set fs = New cFileSystem
                If fs.DirExist(App.Path & "\" & .Folder) Then
                    fs.DirCopy App.Path & "\" & .Folder, App.Path & "\" & mMuds.Folder(mMuds.Count)
                End If
            Set fs = Nothing
            
            Set Item = lvwMuds.ListItems.Add(, "i" & mMuds.Count, NewName)
                Item.SubItems(1) = .Host
                Item.SubItems(2) = .Port
                Item.Selected = True
                FillInfoLabel mMuds.Count
            Set Item = Nothing
        End With
        mMuds.SaveMudList
    End If
    Set ToCopy = Nothing
End Sub

Private Sub mnuExit_Click()
    Unload Me
    Set frmMuds = Nothing
End Sub

Private Sub mnuHomePage_Click()
    ShellExecute Me.hWnd, vbNullString, "http://gosclient.altervista.org/", _
        vbNullString, vbNullString, vbNormalFocus
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
    
        For i = 0 To mnuLang.Count - 1
            If mnuLang(i).Checked Then mnuLang(i).Checked = False
        Next i
        mnuLang(Index).Checked = True
    End If
End Sub

Private Sub mnuMail_Click()
    ShellExecute Me.hWnd, vbNullString, "mailto:gosclient@yahoo.it", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnuOpen_Click()
    OpenClient
End Sub

Private Sub mnuProp_Click()
    Dim Item As ComctlLib.ListItem
    Dim Index As Integer
    
    Set Item = lvwMuds.SelectedItem
    Index = GetKeyIndex(Item.Key)
    
    If frmDefMud.EditMud(mMuds.Mud(Index), mLang) Then
        Item.SubItems(1) = mMuds.Host(Index)
        Item.SubItems(2) = mMuds.Port(Index)
        'LoadList
        mMuds.SaveMudList
        FillInfoLabel Index
    End If

    Set Item = Nothing
End Sub
