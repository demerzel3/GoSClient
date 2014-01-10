VERSION 5.00
Begin VB.Form frmButtons 
   Caption         =   "Pulsanti"
   ClientHeight    =   450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "frmButtons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   3  'Windows Default
   Begin GoS.uMacroBox butt 
      Height          =   315
      Left            =   0
      Top             =   45
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1

Private mButt As cButtons

Private WithEvents mLang As cLang
Attribute mLang.VB_VarHelpID = -1

Private Sub LoadLang()
    'Me.Caption = mLang("buttons", "Caption")
    SetWindowText Me.hwnd, mLang("buttons", "Caption")
End Sub

Private Sub butt_Clicked(Act As String)
    Dim Connect As cConnector

    Set Connect = New cConnector
        Connect.Envi.sendInput Act, TIN_BUTTONS
    Set Connect = Nothing
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    goshSetDockable Me.hwnd, "gos.buttons"
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    Set mFine = New cFinestra
    mFine.Init Me

    Set mButt = New cButtons
    ReloadButt
End Sub

Private Sub ReloadButt()
    Dim i As Integer

    mButt.Clear
    mButt.Load
    butt.Clear
    For i = 1 To mButt.Count
        butt.Add mButt.NormText(i), mButt.NormAct(i), mButt.PresText(i), _
            mButt.PresAct(i), mButt.State(i)
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveStates
    
    Set mButt = Nothing
    
    mFine.UnReg
    Set mFine = Nothing
    
    Set mLang = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    butt.Width = Me.ScaleWidth - butt.Left
    butt.Height = Me.ScaleHeight - butt.Top
End Sub

Private Sub SaveStates()
    Dim i As Integer

    For i = 1 To butt.Count
        mButt.State(i) = butt.GetState(i)
    Next i
    mButt.Save
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    Select Case uMsg
        Case ENVM_BUTTSAVESTATE
            SaveStates
        Case ENVM_BUTTCHANGED, ENVM_PROFILECHANGED, ENVM_SETTCHANGED
            ReloadButt
    End Select
End Sub

Private Sub mLang_RefreshLang()
    LoadLang
End Sub
