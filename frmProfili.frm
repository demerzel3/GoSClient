VERSION 5.00
Begin VB.Form frmProfili 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestione profili"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmProfili.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPass 
      Caption         =   "Cambia password"
      Height          =   315
      Left            =   2925
      TabIndex        =   6
      Top             =   450
      Width           =   1590
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   315
      Left            =   2925
      TabIndex        =   5
      Top             =   4500
      Width           =   1590
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   1275
      TabIndex        =   4
      Top             =   4500
      Width           =   1590
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   315
      Left            =   2925
      TabIndex        =   3
      Top             =   900
      Width           =   1590
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   315
      Left            =   2925
      TabIndex        =   2
      Top             =   75
      Width           =   1590
   End
   Begin VB.ListBox lstElenco 
      Height          =   4125
      IntegralHeight  =   0   'False
      Left            =   75
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   300
      Width           =   2790
   End
   Begin VB.Label lblNProfili 
      Caption         =   "0 profili"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2715
   End
End
Attribute VB_Name = "frmProfili"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mProfili As cProfili

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("profiles", "Caption")

    RefreshCount
    cmdNuovo.Caption = mLang("", "New")
    cmdPass.Caption = mLang("profiles", "ChangePass")
    cmdElimina.Caption = mLang("", "Delete")
    cmdOk.Caption = mLang("", "Ok")
    cmdAnnulla.Caption = mLang("", "Cancel")
End Sub

Private Sub RefreshCount()
    If Not mProfili Is Nothing Then lblNProfili.Caption = mProfili.Count & " " & mLang("profiles", "profiles")
End Sub

Private Sub CaricaProfili()
    Dim i As Integer

    If mProfili Is Nothing Then
        Set mProfili = New cProfili
        mProfili.Carica
    End If
    
    lstElenco.Clear
    For i = 1 To mProfili.Count
        lstElenco.AddItem mProfili.Nick(i)
        If i = mProfili.ProfileSel Then lstElenco.Selected(i - 1) = True
    Next i
    
    RefreshCount
End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
    Set frmProfili = Nothing
End Sub

Private Sub cmdElimina_Click()
    Dim rtn As VbMsgBoxResult
    Dim sel As Integer, Change As Boolean
    
    sel = lstElenco.ListIndex
    If Not sel = -1 Then
        rtn = MsgBox(mLang("profiles", "Remove") & " " & mProfili.Nick(sel + 1) & "?", vbYesNo)
        If rtn = vbYes Then
            mProfili.Remove sel + 1
            If lstElenco.Selected(sel) = True Then Change = True
            lstElenco.RemoveItem sel
            If lstElenco.ListCount > 0 Then
                lstElenco.ListIndex = 0
                If Change Then lstElenco.Selected(0) = True
            End If
            RefreshCount
        End If
    End If
End Sub

Private Sub cmdNuovo_Click()
    Dim Nick As String, Pass As String, Pass2 As String
    Dim Abort As Boolean

    Nick = InputBox(mLang("profiles", "InsNick"), mLang("profiles", "RegNew"))
    If Nick = "" Then
        Abort = True
    Else
        Abort = False
        Do
            Pass = InputBox(mLang("profiles", "InsPass") & " " & Nick, mLang("profiles", "RegNew"))
            If Pass = "" Then Abort = True
            If Not Abort Then Pass2 = InputBox(mLang("profiles", "InsPass2") & " " & Nick, mLang("profiles", "RegNew"))
            If Pass2 = "" Then Abort = True
            
            If Not Abort Then
                If Pass <> Pass2 Then MsgBox mLang("profiles", "DiffPass")
            Else
                Exit Do
            End If
        Loop Until Pass2 = Pass
    End If
    
    If Not Abort Then
        mProfili.Add Nick, Pass
        lstElenco.AddItem Nick
        lstElenco.ListIndex = lstElenco.ListCount - 1
        If lstElenco.ListCount = 1 Then lstElenco.Selected(0) = True
        RefreshCount
    End If
End Sub

Private Sub cmdOk_Click()
    Dim Index As Integer, i As Integer
    Dim Connect As cConnector
    
    Index = 0
    For i = 0 To lstElenco.ListCount - 1
        If lstElenco.Selected(i) = True Then
            Index = i + 1
            Exit For
        End If
    Next i
    mProfili.Save
    
    Set Connect = New cConnector
        Connect.SetProfileSel Index
    Set Connect = Nothing
    
    Unload Me
    Set frmProfili = Nothing
End Sub

Private Sub ChangePass(Index As Integer)
    Dim Confronto As String, Pass As String, Pass2 As String
    Dim Abort As Boolean, Nick As String
    
    Nick = mProfili.Nick(Index)
    Confronto = InputBox(mLang("profiles", "OldPass") & " " & Nick, mLang("profiles", "ChangePass"))
    If Confronto = "" Then
        Abort = True
    ElseIf Confronto = mProfili.Pass(Index) Then
        Abort = False
        Do
            Pass = InputBox(mLang("profiles", "NewPass"), mLang("profiles", "ChangePass"))
            If Pass = "" Then Abort = True
            If Not Abort Then Pass2 = InputBox(mLang("profiles", "NewPass2"), mLang("profiles", "ChangePass"))
            If Pass2 = "" Then Abort = True
            
            If Not Abort Then
                If Pass <> Pass2 Then MsgBox mLang("profiles", "DiffPass")
            Else
                Exit Do
            End If
        Loop Until Pass2 = Pass
    Else
        MsgBox mLang("profiles", "WrongPass")
        Abort = True
    End If
    
    If Not Abort Then mProfili.Pass(Index) = Pass
End Sub

Private Sub cmdPass_Click()
    ChangePass lstElenco.ListIndex + 1
End Sub

Private Sub Form_Load()
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set mLang = Connect.Lang
    Set Connect = Nothing
    LoadLang
    
    CaricaProfili
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mProfili = Nothing

    Set mLang = Nothing
End Sub

Private Sub lstElenco_Click()
    Dim Index As Integer

    Index = lstElenco.ListIndex + 1
    If Index = 0 Then
        cmdElimina.Enabled = False
        cmdPass.Enabled = False
    Else
        cmdElimina.Enabled = True
        cmdPass.Enabled = True
    End If
End Sub

Private Sub lstElenco_DblClick()
    ChangePass lstElenco.ListIndex + 1
End Sub

Private Sub lstElenco_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If lstElenco.Selected(Item) = False Then
        lstElenco.Selected(Item) = True
    Else
        For i = 0 To lstElenco.ListCount - 1
            If i <> Item And lstElenco.Selected(i) = True Then lstElenco.Selected(i) = False
        Next i
    End If
End Sub
