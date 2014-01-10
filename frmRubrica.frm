VERSION 5.00
Begin VB.Form frmRubrica 
   Caption         =   "Rubrica"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "frmRubrica.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstNick 
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   45
      TabIndex        =   1
      Top             =   300
      Width           =   4215
   End
   Begin VB.Image imgMenu 
      Height          =   195
      Left            =   75
      Picture         =   "frmRubrica.frx":00D2
      Top             =   75
      Width           =   195
   End
   Begin VB.Label lblCount 
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
      Height          =   240
      Left            =   300
      TabIndex        =   0
      Top             =   75
      Width           =   1965
   End
   Begin VB.Menu mnuRubrica 
      Caption         =   "Rubrica"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Aggiungi contatto"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Ordina contatti per nick"
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Proprietà"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Elimina contatto"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRubrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mRubrica As cRubrica
Attribute mRubrica.VB_VarHelpID = -1
Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private Sub LoadLang()
    SetWindowText Me.hwnd, mLang("rubrica", "Caption")
    AggiornaCount
    
    'menus
    mnuAdd.Caption = mLang("rubrica", "Add")
    mnuSort.Caption = mLang("rubrica", "Sort")
    mnuProp.Caption = mLang("rubrica", "Properties")
    mnuRemove.Caption = mLang("rubrica", "Remove")
End Sub

Private Sub lstNick_Click()
    If Not lstNick.ListIndex = -1 Then
        mnuProp.Enabled = True
        mnuRemove.Enabled = True
    Else
        mnuProp.Enabled = False
        mnuRemove.Enabled = False
    End If
End Sub

Private Sub lstNick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuRubrica
    End If
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub

Private Sub mnuAdd_Click()
    frmContatto.AddNewContact
End Sub

Private Sub mnuProp_Click()
    frmContatto.ModifyContact lstNick.ListIndex + 1
End Sub

Private Sub AggiornaCount()
    Dim Count As String

    If Not mRubrica Is Nothing Then Count = Trim$(CStr(mRubrica.Count))
    If Not lblCount.Caption = Count Then lblCount.Caption = Count & " " & mLang("rubrica", "contacts")
End Sub

Private Sub mnuRemove_Click()
    If MsgBox(mLang("rubrica", "DelContact") & "'" & lstNick.list(lstNick.ListIndex) & "'?", vbYesNo) = vbYes Then
        mRubrica.Remove lstNick.ListIndex + 1
        AggiornaCount
    End If
End Sub

Private Sub mnuSort_Click()
    If Not lstNick.ListCount = 0 Then
        mRubrica.Sort
        RefreshList
    End If
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    goshSetDockable Me.hwnd, "gos.rubrica"
    
    Set Connect = New cConnector
        Set mRubrica = Connect.Rubrica
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    RefreshList
End Sub

Private Sub RefreshList()
    Dim i As Integer

    lstNick.Clear
    For i = 1 To mRubrica.Count
        lstNick.AddItem mRubrica.Nick(i)
    Next i
    AggiornaCount
    mnuRemove.Enabled = False
    mnuProp.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mRubrica = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'mnuProp.Left = Me.ScaleWidth - mnuProp.Width - 5
    'mnuAdd.Left = mnuProp.Left
    'mnuRemove.Left = mnuProp.Left
    'mnuSort.Left = mnuProp.Left
    'lblCount.Left = mnuProp.Left - 5
    lstNick.Width = Me.ScaleWidth - lstNick.Left - 3
    lstNick.Height = Me.ScaleHeight - lstNick.Top - 3
End Sub

Private Sub imgMenu_Click()
    PopupMenu mnuRubrica
End Sub

Private Sub lstNick_DblClick()
    frmContatto.ModifyContact lstNick.ListIndex + 1
End Sub

Private Sub mRubrica_Added(Index As Integer)
    lstNick.AddItem mRubrica.Nick(Index)
    lstNick.ListIndex = lstNick.ListCount - 1
    AggiornaCount
End Sub

Private Sub mRubrica_Removed(Index As Integer)
    lstNick.RemoveItem Index - 1
    AggiornaCount
End Sub
