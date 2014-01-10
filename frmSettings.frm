VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impostazioni"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   HasDC           =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   1275
      Width           =   1815
   End
   Begin VB.CommandButton cmdModifica 
      Caption         =   "Modifica"
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   825
      Width           =   1815
   End
   Begin VB.CommandButton cmdDupl 
      Caption         =   "Duplica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Aggiungi"
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Top             =   450
      Width           =   1815
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Ripristina"
      Height          =   315
      Left            =   6975
      TabIndex        =   5
      Top             =   5550
      Width           =   1965
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Applica"
      Height          =   315
      Left            =   4950
      TabIndex        =   6
      Top             =   5550
      Width           =   1965
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   5040
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   8890
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "Chiudi"
      Height          =   315
      Left            =   7500
      TabIndex        =   7
      Top             =   6075
      Width           =   1590
   End
   Begin ComctlLib.TabStrip tabTools 
      Height          =   5940
      Left            =   75
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   75
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10478
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ToolNames
    tn_none = 0
    tn_alias = 1
    tn_trigger = 2
    tn_buttons = 3
    tn_variables = 4
    tn_highlights = 5
End Enum

Private mCont(1 To tn_highlights) As Object
Private mNames(1 To tn_highlights) As String
Private mCurTool As Object
Private mCurToolID As ToolNames
Private mChanged As Boolean
Private mIgnoreChange As Boolean

Private mConnect As cConnector

Private mLang As cLang

Private Sub LoadLang()
    Me.Caption = mLang("settings", "Caption")
    
    cmdChiudi.Caption = mLang("", "Close")
    cmdAdd.Caption = mLang("", "Add")
    cmdModifica.Caption = mLang("", "Modify")
    cmdElimina.Caption = mLang("", "Delete")
    cmdApply.Caption = mLang("", "Apply")
    cmdAbort.Caption = mLang("", "Abort")
End Sub

Private Sub cmdAbort_Click()
    LoadTool mCurToolID
    Set mCurTool = mCont(mCurToolID)
    LoadToolInfo
    mChanged = False
    cmdApply.Enabled = False
    cmdAbort.Enabled = False
End Sub

Private Sub cmdAdd_Click()
    If mCurTool.sett_New Then
        RefreshItem mCurTool.sett_Count, True
    End If
End Sub

Private Sub cmdApply_Click()
    mCurTool.sett_Save
    mChanged = False
    cmdApply.Enabled = False
    cmdAbort.Enabled = False
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
    Set frmSettings = Nothing
End Sub

Private Sub SetChanged()
    mChanged = True
    cmdApply.Enabled = True
    cmdAbort.Enabled = True
End Sub

Private Sub ModifySelected()
    Dim Index As Integer
    
    Index = GetKeyIndex(lvwItems.SelectedItem.Key)
    If mCurTool.sett_Modify(Index) Then
        RefreshItem Index
        SetChanged
    End If
End Sub

Private Sub cmdElimina_Click()
    Dim rtn As VbMsgBoxResult
    Dim Index As Integer
    
    Index = GetKeyIndex(lvwItems.SelectedItem.Key)
    rtn = MsgBox(mLang("", "Delete") & " " & mCurTool.sett_Item(Index, 1) & "?", vbYesNo)
    If rtn = vbYes Then
        mCurTool.sett_Delete (Index)
        LoadToolInfo
        SetChanged
    End If
End Sub

Private Sub cmdModifica_Click()
    ModifySelected
End Sub

Private Sub LoadTool(ByVal tType As ToolNames)
    Dim Alias As cAlias, Trig As cTriggers, Buttons As cButtons, Vars As cVars
    
    Set mCont(tType) = Nothing
    Select Case tType
        Case tn_alias
            Set Alias = New cAlias
                Alias.LoadAliases
                Set mCont(tn_alias) = Alias
            Set Alias = Nothing
        Case tn_trigger
            Set Trig = New cTriggers
                Trig.Load
                Set mCont(tn_trigger) = Trig
            Set Trig = Nothing
        Case tn_buttons
            Set Buttons = New cButtons
                Buttons.Load
                Set mCont(tn_buttons) = Buttons
            Set Buttons = Nothing
        Case tn_variables
            Set Vars = New cVars
                Vars.LoadVars
                Set mCont(tn_variables) = Vars
            Set Vars = Nothing
    End Select
End Sub

Private Sub Form_Load()
    
    Set mConnect = New cConnector
    Set mLang = mConnect.Lang
    LoadLang
    
    mNames(tn_alias) = mLang("settings", "Alias")
    mNames(tn_trigger) = mLang("settings", "Trigger")
    mNames(tn_buttons) = mLang("settings", "Buttons")
    mNames(tn_variables) = mLang("settings", "Variables")
    mNames(tn_highlights) = mLang("settings", "Highlights")
    
    With tabTools
        .Tabs.Clear
        .Tabs.Add , , mNames(tn_alias)
            LoadTool tn_alias
        
        .Tabs.Add , , mNames(tn_trigger)
            LoadTool tn_trigger
        
        .Tabs.Add , , mNames(tn_buttons)
            LoadTool tn_buttons
    
        .Tabs.Add , , mNames(tn_variables)
            LoadTool tn_variables
    
        .Tabs.Item(1).Selected = True
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
        
    If CheckChanges Then
        For i = 1 To UBound(mCont, 1)
            Set mCont(i) = Nothing
        Next i
        
        Set mConnect = Nothing
    Else
        Cancel = 1
    End If
End Sub

Private Sub RefreshItem(Index As Integer, Optional Add As Boolean = False)
    Dim i As Integer, colCount As Integer
    Dim Item As ComctlLib.ListItem
    
    SetChanged
    If Add Then
        Set Item = lvwItems.ListItems.Add(, "i" & Index, mCurTool.sett_Item(Index, 1))
    Else
        For i = 1 To lvwItems.ListItems.Count
            If lvwItems.ListItems.Item(i).Key = "i" & Index Then
                Set Item = lvwItems.ListItems.Item(i)
                Exit For
            End If
        Next i
    End If
    
    Item.Text = mCurTool.sett_Item(Index, 1)
    For i = 2 To mCurTool.sett_ColCount
        Item.SubItems(i - 1) = mCurTool.sett_Item(Index, i)
    Next i
    
    If Add Then
        Item.Selected = True
        lvwItems.SetFocus
    End If
    
    Set Item = Nothing
End Sub

Private Sub LoadToolInfo()
    Dim i As Integer, j As Integer
    Dim col() As String, colCount As Integer, colWidth As Long
    Dim Item As ComctlLib.ListItem
    
    cmdAbort.Enabled = False
    cmdApply.Enabled = False

    With lvwItems
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        'adding column headers
        With lvwItems.ColumnHeaders
            mCurTool.sett_Column col()
            colCount = UBound(col, 1)
            colWidth = 6000 / colCount - 60
            For i = 1 To colCount
                .Add , , col(i), colWidth
            Next i
        End With
        
        If mCurTool.sett_Sorted Then
            .Sorted = True
            .SortKey = 0
        Else
            .Sorted = False
        End If
        
        'adding items
        For i = 1 To mCurTool.sett_Count
            Set Item = lvwItems.ListItems.Add(, "i" & i, mCurTool.sett_Item(i, 1))
            For j = 2 To colCount
                Item.SubItems(j - 1) = mCurTool.sett_Item(i, j)
            Next j
            Set Item = Nothing
        Next i
        
        If .ListItems.Count >= 1 Then
            .ListItems.Item(1).Selected = True
            cmdDupl.Enabled = True
            cmdModifica.Enabled = True
            cmdElimina.Enabled = True
        Else
            cmdDupl.Enabled = False
            cmdModifica.Enabled = False
            cmdElimina.Enabled = False
        End If
        
        If Me.Visible Then .SetFocus
    End With
End Sub

Private Function CheckChanges() As Boolean
    'ritorna true se si puo' proseguire, false se si deve annullare
    'l'operazione.
    
    Dim rtn As VbMsgBoxResult
    
    CheckChanges = True
    If mChanged Then
        rtn = MsgBox(mLang("settings", "SaveChanges") & " " & mNames(mCurToolID) & "?", _
            vbYesNoCancel)
         
        If rtn = vbCancel Then
            CheckChanges = False
        ElseIf rtn = vbYes Then
            mCurTool.sett_Save
        End If
    End If
End Function

Private Sub lvwItems_DblClick()
    ModifySelected
End Sub

Private Sub tabTools_Click()
    If mIgnoreChange Then
        mIgnoreChange = False
        Exit Sub
    End If
    
    If CheckChanges Then
        mCurToolID = tabTools.SelectedItem.Index
        Set mCurTool = mCont(mCurToolID)
        LoadToolInfo
        mChanged = False
    Else
        mIgnoreChange = True
        tabTools.Tabs(mCurToolID).Selected = True
    End If
End Sub

Private Function GetKeyIndex(Key As String) As Integer
    GetKeyIndex = Val(Mid$(Key, 2))
End Function
