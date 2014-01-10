VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mAlias() As String
Private mMode As Integer

Private mInfo As cIni

Private Const Aliases As Integer = 1
Private Const TRIGGERS As Integer = 2

Public Property Get sett_ColCount() As Integer
    sett_ColCount = 2
End Property

Public Sub sett_Delete(Index As Integer)
    Me.Remove Index
End Sub

Public Function sett_New() As Boolean
    Me.Add ""
    
    Load frmDefAlias
    frmDefAlias.Init Me, Me.Count
    
    sett_New = True
End Function

Public Function sett_Modify(Index As Integer) As Boolean
    Load frmDefAlias
    frmDefAlias.Init Me, Index
    
    sett_Modify = True
End Function

Public Sub sett_Save()
    Me.Save
End Sub

Public Property Get sett_Sorted() As Boolean
    sett_Sorted = True
End Property

Public Property Get sett_Count() As Integer
    sett_Count = Me.Count
End Property

Public Sub sett_Column(col() As String)
    Dim Connect As cConnector
    
    ReDim col(1 To 2) As String
    Set Connect = New cConnector
        col(1) = Connect.Lang("calias", "Text")
        col(2) = Connect.Lang("calias", "KeyCombo")
    Set Connect = Nothing
End Sub

Public Property Get sett_Item(ByVal Index As Integer, ByVal col As Integer) As String
    Dim Combo As cKeyCombo, Key As Integer, Shift As Integer
    
    Select Case col
        Case 1 'testo
            sett_Item = Me.Text(Index)
        Case 2 'combo
            Set Combo = New cKeyCombo
                Combo.ScindiCombo Me.Combo(Index), Key, Shift
                sett_Item = Combo.IdentCombo(Key, Shift)
            Set Combo = Nothing
    End Select
End Property

Public Function MoveActionUp(ByVal AliasID As Integer, ByVal ActionID As Integer) As Boolean
    Dim WPar As String, Par As String
    
    If ActionID = 1 Then
        MoveActionUp = False
    Else
        WPar = Action(AliasID, ActionID - 1)
        Par = ActionPar(AliasID, ActionID - 1)
        Action(AliasID, ActionID - 1) = Action(AliasID, ActionID)
        ActionPar(AliasID, ActionID - 1) = ActionPar(AliasID, ActionID)
        Action(AliasID, ActionID) = WPar
        ActionPar(AliasID, ActionID) = Par
        MoveActionUp = True
    End If
End Function

Public Function MoveActionDown(ByVal AliasID As Integer, ByVal ActionID As Integer) As Boolean
    Dim WPar As String, Par As String
    
    If ActionID = ActionCount(AliasID) Then
        MoveActionDown = False
    Else
        WPar = Action(AliasID, ActionID + 1)
        Par = ActionPar(AliasID, ActionID + 1)
        Action(AliasID, ActionID + 1) = Action(AliasID, ActionID)
        ActionPar(AliasID, ActionID + 1) = ActionPar(AliasID, ActionID)
        Action(AliasID, ActionID) = WPar
        ActionPar(AliasID, ActionID) = Par
        MoveActionDown = True
    End If
End Function

Public Sub SendAlias(Index As Integer, Optional Param As String = "")
    Dim ToExec As String, Parametri() As String, Count As Integer
    Dim ToAdd As String, Action As String, Other As String, CurPar As String
    Dim i As Integer, j As Integer

    Dim Connect As cConnector

    If Param = "" Then
        Count = -1
    Else
        Param = IIf(Right$(Param, 2) = vbCrLf, Left$(Param, Len(Param) - 2), Param)
        Param = Trim$(Param)
        Parametri = Split(Param)
        Count = UBound(Parametri, 1)
    End If
    
    ToExec = ""
    If Count = -1 Then
        For i = 1 To Me.ActionCount(Index)
            If Not i = 1 Then ToExec = ToExec & vbCrLf
            ToExec = ToExec & Me.Action(Index, i)
        Next i
    Else
        For i = 1 To Me.ActionCount(Index)
            If Not Me.ActionPar(Index, i) = "" Then
                If Not i = 1 Then ToExec = ToExec & vbCrLf
                ToAdd = Me.ActionPar(Index, i)
                Action = ToAdd
                If InStr(1, ToAdd, "%x", vbTextCompare) <> 0 Then
                    ToAdd = Replace(Action, "%x", Parametri(0))
                    For j = 1 To Count
                        ToAdd = ToAdd & vbCrLf & Replace(Action, "%x", Parametri(j))
                    Next j
                End If
                
                'ToExec = ToExec & Me.ActionPar(Index, i)
                ToExec = ToExec & ToAdd
            End If
        Next i
        
        ToExec = Replace(ToExec, "%a", Param)
        
        For i = 0 To Count
            CurPar = "%" & (i + 1)
            If InStr(1, ToExec, CurPar) <> 0 Then
                ToExec = Replace(ToExec, CurPar, Parametri(i))
            Else
                Other = Other & Parametri(i) & " "
            End If
        Next i
        If Len(Other) > 0 Then Other = Left$(Other, Len(Other) - 1)
        
        ToExec = Replace(ToExec, "%o", Other)
        
        'If InStr(1, ToExec, "%x", vbTextCompare) <> 0 Then
        '    For i = 0 To Count
        '        mConnect.Envi.sendInput Replace(ToExec, "%x", Parametri(i)) & vbCrLf, TIN_TOSEND
        '        'mConnect.Envi.sendOutput Replace(ToExec, "%x", Parametri(i)) & vbCrLf
        '    Next i
        '    Exit Sub
        'End If
    End If
    
    'ToExec = Me.Comando(Index)
    'If InStr(1, ToExec, "§") <> 0 Then
    '    ToExec = Mid$(ToExec, 1, InStr(1, ToExec, "§") - 1)
    'End If
    'ToExec = ToExec & "|"
    'ToExec = Replace(ToExec, "|", Param & vbCrLf)
    
    Set Connect = New cConnector
        Connect.Envi.sendInput ToExec & vbCrLf, TIN_TOQUEUE
        'Connect.Envi.sendOutput ToExec & vbCrLf
    Set Connect = Nothing
End Sub

Public Property Get Mode() As String
    If mMode = Aliases Then
        Mode = "aliases"
    ElseIf mMode = TRIGGERS Then
        Mode = "triggers"
    End If
End Property

Public Property Get Count2() As Integer
    Count2 = UBound(mAlias, 2)
End Property

Public Property Get Count() As Integer
    Count = Val(mInfo.RetrInfo("Count", 0))
End Property

Private Function GetPath() As String
    If Not mMode = 0 Then
        Select Case mMode
            Case Aliases
                GetPath = "alias.txt"
            Case TRIGGERS
                GetPath = "trigger.txt"
        End Select
    End If
End Function

Private Function EstraiInfo(Stringa As String, ByRef Azione As String) As String
    If InStr(1, Stringa, "=") <> 0 Then
        EstraiInfo = Trim$(Mid$(Stringa, InStr(1, Stringa, "=", vbTextCompare) + 1))
        Azione = Trim$(Mid$(Stringa, 1, InStr(1, Stringa, "=", vbTextCompare) - 2))
    End If
End Function

Private Sub LoadAliasFromFile(Free As Integer, Count As Integer)
    Dim Loa As String
    Dim i As Integer
    Dim Azione As String, Comando As String

    If Not Count = 0 Then
        ReDim mAlias(1 To 3, 1 To Count) As String
        For i = 1 To Count
            Line Input #Free, Loa
            Comando = EstraiInfo(Loa, Azione)
            mAlias(2, i) = Comando
            mAlias(1, i) = Azione
            If mMode = Aliases Then
                If InStr(1, Comando, "§") <> 0 Then
                    mAlias(3, i) = Mid$(Comando, InStr(1, Comando, "§") + 1)
                End If
            End If
        Next i
    End If
End Sub

Private Sub Carica(Optional Profile As Integer = -1)
    Dim Path As String
    Dim Connect As cConnector
    Dim Free As Integer, Tipo As String

    'Path = IIf(Right$(App.Path, 1) <> "\", App.Path & "\", App.Path)
    Set Connect = New cConnector
    Path = Connect.ProfileFolder
    If mMode = Aliases Then
        Tipo = "alias.ini"
    Else
        Tipo = "trigger.ini"
    End If
    If Dir$(Path, vbDirectory) = "" Then
        MkDir Path
        Connect.Log "The directory for the profile was created"
    End If
    
    Path = Path & Tipo
        
    'If Dir$(Path, vbDirectory) = "" Then
    '    MkDir Path
    '    Connect.Log "Creata la directory " & Path
    'End If
    '
    'If Profile = -1 Then
    '    Profile = Connect.ProfileSel
    'End If
    '
    'Path = Tipo & "\profile" & Profile & ".ini"
    '
    If Dir$(Path) = "" Then
        Free = FreeFile
        Open Path For Binary As Free: Close Free
        'Connect.Log "Creato il file " & Tipo
    End If
    
    Set mInfo = New cIni
    mInfo.CaricaFile Path, True
    'Connect.Log "Loaded " & Left$(Tipo, Len(Tipo) - 4) & " from " & Tipo
    
    Set Connect = Nothing
End Sub

Private Sub Carica2(Optional Profile As Integer = -1)
    Dim Path As String
    Dim Free As Integer
    'Dim Azione As String, Comando As String
    Dim Azione As String
    Dim Loa As String
    Dim Count As Integer, i As Integer
    Dim Connect As cConnector

    If Profile = -1 Then
        Set Connect = New cConnector
        Profile = Connect.ProfileSel
        Set Connect = Nothing
    End If
    
    Path = App.Path & "\" & GetPath
    ReDim mAlias(0 To 0, 0 To 0) As String

    If Not Path = "" Then
        Free = FreeFile
        Open Path For Input As #Free
            If Not EOF(Free) Then
                Line Input #Free, Loa
                If LCase$(Left$(Loa, 5)) = "count" Then
                    If Profile = 0 Then
                        Count = CInt(EstraiInfo(Loa, Azione))
                        LoadAliasFromFile Free, Count
                    End If
                ElseIf LCase$(Left$(Loa, 5)) = "profi" Then
                    Do Until Profile = CInt(EstraiInfo(Loa, Azione))
                        Line Input #Free, Loa
                        Count = CInt(EstraiInfo(Loa, Azione))
                        For i = 1 To Count
                            Line Input #Free, Loa
                        Next i
                        If EOF(Free) Then
                            Close #Free
                            Exit Sub
                        End If
                        Line Input #Free, Loa
                    Loop
                    Line Input #Free, Loa
                    Count = CInt(EstraiInfo(Loa, Azione))
                    LoadAliasFromFile Free, Count
                End If
            End If
        Close #Free
    End If
End Sub

Public Sub LoadTriggers(Optional Profile As Integer = -1)
    mMode = TRIGGERS
    Carica Profile
End Sub

Public Sub LoadAliases(Optional Profile As Integer = -1)
    mMode = Aliases
    Carica Profile
End Sub

Public Sub Save()
    mInfo.SalvaFile
End Sub

Private Sub Save2()
    Dim Path As String, Profili As cProfili
    Dim Free As Integer, i As Integer, j As Integer
    Dim Temp As String
    Dim data As cAlias
    
    Path = App.Path & "\" & GetPath
    Temp = Path & ".tmp"
    Free = FreeFile
    Open Temp For Output As #Free
        
        Set Profili = New cProfili
        Profili.Carica
        For j = 0 To Profili.Count
            Print #Free, "profile = " & j
            If Not j = Profili.ProfileSel Then
                Set data = New cAlias
                
                If mMode = Aliases Then
                    data.LoadAliases j
                ElseIf mMode = TRIGGERS Then
                    data.LoadTriggers j
                End If
                
                Print #Free, "count = " & data.Count
                For i = 1 To data.Count
                    Print #Free, data.Azione(i) & " = " & data.Comando(i)
                Next i
                
                Set data = Nothing
            Else
                Print #Free, "count = " & Me.Count
                For i = 1 To Me.Count
                    Print #Free, Me.Azione(i) & " = " & Me.Comando(i)
                Next i
            End If
        Next j
        Set Profili = Nothing
    
    Close #Free

    Kill Path
    FileCopy Temp, Path
    Kill Temp
End Sub

Public Function Add(Testo As String) As Integer
    Dim Count As Integer

    Count = Me.Count
    Count = Count + 1
    mInfo.AddInfo "Count", Count
    mInfo.AddInfo "Text<" & Count & ">", Testo
    mInfo.AddInfo "Combo<" & Count & ">", "0|0"
    mInfo.AddInfo "Action<" & Count & ">_Count", -1
    
    Add = Count
End Function

Public Function AddAction(AliasID As Integer) As Integer
    Dim Count As Integer

    Count = Me.ActionCount(AliasID)
    Count = Count + 1
    ActionCount(AliasID) = Count
    'mInfo.AddInfo "Action<" & AliasID & "_Count", Count
    'Count = Count + 1
    'mInfo.AddInfo "Action<" & AliasID & "," & Count & ">_wpar", ""
    'mInfo.AddInfo "Action<" & AliasID & "," & Count & ">_par", ""
    Me.Action(AliasID, Count) = ""
    Me.ActionPar(AliasID, Count) = ""
    
    AddAction = Count
End Function

Private Sub Add2(Azione As String, Comando As String)
    Dim pCount As Integer

    pCount = UBound(mAlias, 2) + 1
    If pCount = 1 Then
        ReDim mAlias(1 To 3, 1 To 1) As String
    Else
        ReDim Preserve mAlias(1 To 3, 1 To pCount) As String
    End If
    
    mAlias(1, pCount) = Azione
    mAlias(2, pCount) = Comando
    mAlias(3, pCount) = "0|0"
End Sub

Public Sub Remove(AliasID As Integer)
    Dim i As Integer, j As Integer
    Dim Count As Integer

    Count = Me.Count
    If Not Count = 1 Then
        If Not AliasID = Count Then
            For i = AliasID To Count - 1
                ActionCount(i) = ActionCount(i + 1)
                Text(i) = Text(i + 1)
                Combo(i) = Combo(i + 1)
                For j = 1 To ActionCount(i)
                    Action(i, j) = Action(i + 1, j)
                    ActionPar(i, j) = ActionPar(i + 1, j)
                Next j
            Next i
        End If
        
        mInfo.AddInfo "Count", Count - 1
    Else
        mInfo.AddInfo "Count", 0
    End If
End Sub

Public Sub RemoveAction(AliasID As Integer, ActionID As Integer)
    Dim i As Integer
    Dim Count As Integer

    Count = Me.ActionCount(AliasID)
    If Not Count = 1 Then
        If Not ActionID = Count Then
            For i = ActionID To Count - 1
                Action(AliasID, i) = Action(AliasID, i + 1)
                ActionPar(AliasID, i) = ActionPar(AliasID, i + 1)
            Next i
        End If
        
        'mInfo.AddInfo "Count", Count - 1
        ActionCount(AliasID) = Count - 1
    Else
        'mInfo.AddInfo "Count", 0
        ActionCount(AliasID) = 0
    End If
End Sub

Public Sub Remove2(Index As Integer)
    Dim i As Integer
    Dim Count As Integer

    Count = UBound(mAlias, 2)
    
    If Not Count = 1 Then
        If Not Index = Count Then
            For i = Index To Count - 1
                mAlias(1, i) = mAlias(1, i + 1)
                mAlias(2, i) = mAlias(2, i + 1)
                mAlias(3, i) = mAlias(3, i + 1)
            Next i
        End If
        
        ReDim Preserve mAlias(1 To 3, 1 To Count - 1) As String
    Else
        ReDim mAlias(0 To 0, 0 To 0) As String
    End If
End Sub

Public Property Let Azione(Index As Integer, data As String)
    'mAlias(1, Index) = data
End Property

Public Property Let Comando(Index As Integer, data As String)
    'mAlias(2, Index) = data
    'If mMode = ALIASES Then
    '    If InStr(1, data, "§") <> 0 Then
    '        mAlias(3, Index) = Mid$(data, InStr(1, data, "§") + 1)
    '    End If
    'End If
End Property

Public Property Let Text(AliasIndex As Integer, data As String)
    mInfo.AddInfo "Text<" & AliasIndex & ">", data
End Property

Public Property Get Text(AliasIndex As Integer) As String
    Text = mInfo.RetrInfo("Text<" & AliasIndex & ">", "")
End Property

Public Property Get Azione(Index As Integer) As String
    'Azione = mAlias(1, Index)
    Azione = Text(Index)
End Property

Public Property Get ActionCount(AliasID As Integer) As Integer
    ActionCount = Val(mInfo.RetrInfo("Action<" & AliasID & ">_Count", 0)) + 1
End Property

Public Property Get ActionPar(AliasID As Integer, ActionID As Integer) As String
    'azione con parametri
    ActionPar = mInfo.RetrInfo("Action<" & AliasID & "," & ActionID - 1 & ">_par", "")
End Property

Public Property Get Action(AliasID As Integer, ActionID As Integer) As String
    'azione senza parametri
    Action = mInfo.RetrInfo("Action<" & AliasID & "," & ActionID - 1 & ">_wpar", "")
End Property

Private Property Let ActionCount(AliasID As Integer, data As Integer)
    mInfo.AddInfo "Action<" & AliasID & ">_Count", data - 1
End Property

Public Property Let ActionPar(AliasID As Integer, ActionID As Integer, data As String)
    'azione con parametri
    mInfo.AddInfo "Action<" & AliasID & "," & ActionID - 1 & ">_par", data
End Property

Public Property Let Action(AliasID As Integer, ActionID As Integer, data As String)
    'azione senza parametri
    mInfo.AddInfo "Action<" & AliasID & "," & ActionID - 1 & ">_wpar", data
End Property

Public Property Get Comando(Index As Integer) As String
    'Comando = mAlias(2, Index)
    Comando = mInfo.RetrInfo("Action<" & Index & ",0>_wpar", "")
End Property

Public Property Let Combo(AliasID As Integer, sCombo As String)
    mInfo.AddInfo "Combo<" & AliasID & ">", sCombo
End Property

Public Property Get Combo(AliasID As Integer) As String
    'Combo = mAlias(3, Index)
    Combo = mInfo.RetrInfo("Combo<" & AliasID & ">", "")
End Property

Private Sub Class_Terminate()
    Set mInfo = Nothing
End Sub
