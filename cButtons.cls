VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type cbt_button
    OptionButt As Boolean
    
    NormText As String
    NormAct As String
    PresText As String
    PresAct As String

    State As Integer
End Type

Private mButt() As cbt_button
Private mPath As String

Public Property Get sett_ColCount() As Integer
    sett_ColCount = 4
End Property

Public Sub sett_Delete(Index As Integer)
    Me.Remove Index
End Sub

Public Function sett_New() As Boolean
    Me.Add "", "", "", ""
    
    frmDefButt.Init Me.Count, Me
    
    sett_New = True
End Function

Public Function sett_Modify(Index As Integer) As Boolean
    frmDefButt.Init Index, Me
    
    sett_Modify = True
End Function

Public Sub sett_Save()
    Me.Save
End Sub

Public Property Get sett_Sorted() As Boolean
    sett_Sorted = False
End Property

Public Property Get sett_Count() As Integer
    sett_Count = Me.Count
End Property

Public Sub sett_Column(col() As String)
    Dim Connect As cConnector
    
    ReDim col(1 To 4) As String
    Set Connect = New cConnector
        col(1) = Connect.Lang("cbuttons", "TextNormal")
        col(2) = Connect.Lang("cbuttons", "AliasNormal")
        col(3) = Connect.Lang("cbuttons", "TextPressed")
        col(4) = Connect.Lang("cbuttons", "AliasPressed")
    Set Connect = Nothing
End Sub

Public Property Get sett_Item(ByVal Index As Integer, ByVal col As Integer) As String
    Dim Combo As cKeyCombo, Key As Integer, Shift As Integer
    
    Select Case col
        Case 1 'testo normale
            sett_Item = Me.NormText(Index)
        Case 2 'alias normale
            sett_Item = Me.NormAct(Index)
        Case 3 'testo premuto
            sett_Item = Me.PresText(Index)
        Case 4 'alias premuto
            sett_Item = Me.PresAct(Index)
    End Select
End Property

Public Sub Load(Optional Profile As Integer = -1)
    Dim Path As String
    Dim Connect As cConnector
    Dim Free As Integer
    Dim Ini As cIni, i As Integer
    Dim NT As String, NA As String, pt As String, PA As String, ST As Integer
    Dim Count As Integer

    'Path = IIf(Right$(App.Path, 1) <> "\", App.Path & "\", App.Path)
    'Path = Path & "buttons"
    Set Connect = New cConnector
    Path = Connect.ProfileFolder
    If Dir$(Path, vbDirectory) = "" Then
        MkDir Path
        Connect.Log "The directory for the profile was created"
    End If
    
    'If Profile = -1 Then
    '    Profile = Connect.ProfileSel
    'End If
    '
    'Path = "buttons\profile" & Profile & ".ini"
    Path = Path & "buttons.ini"
    
    If Dir$(Path) = "" Then
        Free = FreeFile
        Open Path For Binary As Free: Close Free
        'Connect.Log "Creato il file buttons.ini"
    End If
    
    mPath = Path
    Set Ini = New cIni
        Ini.CaricaFile mPath, True
        Count = Ini.RetrInfo("butt_count")
        For i = 1 To Count
            NT = Ini.RetrInfo("butt<" & i & ">_NormText", "")
            NA = Ini.RetrInfo("butt<" & i & ">_NormAct", "")
            pt = Ini.RetrInfo("butt<" & i & ">_PresText", "")
            PA = Ini.RetrInfo("butt<" & i & ">_PresAct", "")
            ST = Ini.RetrInfo("butt<" & i & ">_State", 0)
            Me.Add NT, NA, pt, PA
            Me.State(i) = ST
        Next i
    Set Ini = Nothing
    'onnect.Log "Loaded buttons from buttons.ini"
    
    Set Connect = Nothing
End Sub

Public Sub Save(Optional Profile As Integer = -1)
    Dim Ini As cIni, i As Integer

    Set Ini = New cIni
        Ini.CaricaFile mPath, True
        Ini.AddInfo "butt_count", Me.Count
        For i = 1 To Me.Count
            With mButt(i)
                Ini.AddInfo "butt<" & i & ">_NormText", .NormText
                Ini.AddInfo "butt<" & i & ">_NormAct", .NormAct
                Ini.AddInfo "butt<" & i & ">_PresText", .PresText
                Ini.AddInfo "butt<" & i & ">_PresAct", .PresAct
                Ini.AddInfo "butt<" & i & ">_State", .State
            End With
        Next i
        Ini.SalvaFile
    Set Ini = Nothing
End Sub

Public Property Get OptionButt(Index As Integer) As Boolean
    OptionButt = mButt(Index).OptionButt
End Property

Public Property Let OptionButt(Index As Integer, data As Boolean)
    mButt(Index).OptionButt = data
End Property

Public Property Get State(Index As Integer) As Integer
    State = mButt(Index).State
End Property

Public Property Let State(Index As Integer, data As Integer)
    mButt(Index).State = data
End Property

Public Property Get NormAct(Index As Integer) As String
    NormAct = mButt(Index).NormAct
End Property

Public Property Let NormAct(Index As Integer, data As String)
    mButt(Index).NormAct = data
End Property

Public Property Get NormText(Index As Integer) As String
    NormText = mButt(Index).NormText
End Property

Public Property Let NormText(Index As Integer, data As String)
    mButt(Index).NormText = data
End Property

Public Property Get PresAct(Index As Integer) As String
    PresAct = mButt(Index).PresAct
End Property

Public Property Let PresAct(Index As Integer, data As String)
    mButt(Index).PresAct = data
End Property

Public Property Get PresText(Index As Integer) As String
    PresText = mButt(Index).PresText
End Property

Public Property Let PresText(Index As Integer, data As String)
    mButt(Index).PresText = data
End Property

Public Property Get Count() As Integer
    Count = UBound(mButt(), 1)
End Property

Public Sub Clear()
    Erase mButt()
    ReDim mButt(0) As cbt_button
End Sub

Public Sub Add(NormText As String, NormAct As String, PresText As String, PresAct As String)
    Dim Count As Integer

    Count = Me.Count + 1
    If Count = 1 Then
        ReDim mButt(1 To 1) As cbt_button
    Else
        ReDim Preserve mButt(1 To Count) As cbt_button
    End If
    
    With mButt(Count)
        .NormText = NormText
        .NormAct = NormAct
        If PresText = "" Then
            .OptionButt = False
        Else
            .OptionButt = True
            .PresText = PresText
            .PresAct = PresAct
        End If
        .State = STATE_NORMAL
    End With
End Sub

Public Sub Remove(Index As Integer)
    Dim Count As Integer
    Dim i As Integer

    Count = Me.Count
    If Not Index = Count Then
        For i = Index To Count - 1
            mButt(i) = mButt(i + 1)
        Next i
    End If
    
    If (Count - 1) = 0 Then
        ReDim mButt(0) As cbt_button
    Else
        ReDim Preserve mButt(1 To (Count - 1)) As cbt_button
    End If
End Sub

Private Sub Class_Initialize()
    ReDim mButt(0 To 0) As cbt_button
End Sub

Private Sub Class_Terminate()
    Erase mButt()
End Sub
