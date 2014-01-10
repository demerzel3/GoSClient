VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPlugins 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestione plug-ins"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "frmPlugins.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8325
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Chiudi"
      Height          =   315
      Left            =   6225
      TabIndex        =   1
      Top             =   4575
      Width           =   2040
   End
   Begin ComctlLib.ListView lvwList 
      Height          =   4440
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   7832
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuConfig 
         Caption         =   "Configura..."
      End
      Begin VB.Menu mnuStartup 
         Caption         =   "Avvio"
         Begin VB.Menu mnuManual 
            Caption         =   "Manuale"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Automatico"
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "Informazioni su..."
      End
   End
End
Attribute VB_Name = "frmPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPlugins As cPlugIns

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("plugins", "Caption")
    
    mnuConfig.Caption = mLang("plugins", "Configure")
    mnuCredits.Caption = mLang("plugins", "About")
    mnuStartup.Caption = mLang("plugins", "Start")
    mnuAuto.Caption = mLang("plugins", "Automatic")
    mnuManual.Caption = mLang("plugins", "Manual")
    
    cmdClose.Caption = mLang("", "Close")
End Sub

Public Sub Init(list As cPlugIns)
    Dim i As Integer, Item As ComctlLib.ListItem
    
    Set mPlugins = list
    For i = 1 To mPlugins.Count
        With mPlugins.Item(i)
            Set Item = lvwList.ListItems.Add(, , .Title)
                If .Loaded Then
                    Item.SubItems(1) = mLang("plugins", "Started")
                Else
                    Item.SubItems(1) = mLang("plugins", "Stopped")
                End If
                
                If .Auto Then
                    Item.SubItems(2) = mLang("plugins", "Automatic")
                Else
                    Item.SubItems(2) = mLang("plugins", "Manual")
                End If
            Set Item = Nothing
        End With
    Next i
    
    Me.Show vbModal, frmBase
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set frmPlugins = Nothing
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    Me.Hide
    
    lvwList.ColumnHeaders.Add , , mLang("plugins", "Title"), 5400
    lvwList.ColumnHeaders.Add , , mLang("plugins", "State"), 700
    lvwList.ColumnHeaders.Add , , mLang("plugins", "Start"), 1000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mPlugins.SaveAutoLoaded
    Set mPlugins = Nothing
    
    Set mLang = Nothing
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim plugin As cPlugIn
    Dim Item As ComctlLib.ListItem
    Dim mouse As POINTAPI, rc As RECT
    
    GetWindowRect lvwList.hwnd, rc
    GetCursorPos mouse
    mouse.Y = mouse.Y - 2
    Set Item = lvwList.HitTest(((mouse.X - rc.Left) * Screen.TwipsPerPixelX), _
        ((mouse.Y - rc.Top) * Screen.TwipsPerPixelY))
    Debug.Print X, Y
    If Not Item Is Nothing Then
        Set plugin = mPlugins.Item(Item.Index)
            If plugin.Auto Then
                mnuManual.Checked = False
                mnuAuto.Checked = True
            Else
                mnuManual.Checked = True
                mnuAuto.Checked = False
            End If
        Set plugin = Nothing
        PopupMenu mnuPopup
    End If
    
    Set Item = Nothing
End Sub

Private Sub mnuAuto_Click()
    mPlugins.Item(lvwList.SelectedItem.Index).Auto = True
    lvwList.SelectedItem.SubItems(2) = mLang("plugins", "Automatic")
End Sub

Private Sub mnuConfig_Click()
    If Not mPlugins.Item(lvwList.SelectedItem.Index).ConfPlugIn(Me.hwnd) Then
        MsgBox mLang("plugins", "NoConf"), , lvwList.SelectedItem.Text
    End If
End Sub

Private Sub mnuCredits_Click()
    If Not mPlugins.Item(lvwList.SelectedItem.Index).CredPlugIn(Me.hwnd) Then
        MsgBox mLang("plugins", "NoAbout"), , lvwList.SelectedItem.Text
    End If
End Sub

Private Sub mnuManual_Click()
    mPlugins.Item(lvwList.SelectedItem.Index).Auto = False
    lvwList.SelectedItem.SubItems(2) = mLang("plugins", "Manual")
End Sub
