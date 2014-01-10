VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mVars() As String
Private mCount As Long

Private Const VAR_NAME As Integer = 1
Private Const VAR_VALUE As Integer = 2
Private Const INCREMENT As Integer = 100

Public Property Get sett_ColCount() As Integer
    sett_ColCount = 2
End Property

Public Sub sett_Delete(Index As Integer)
    Me.Remove Index
End Sub

Public Function sett_New() As Boolean
    Dim Ret As Boolean
    Dim sName As String, sValue As String
    
    Me.Add "", "", False
    
    Load frmDefVar
    sName = Me.Name(mCount)
    sValue = Me.Value(mCount)
    Ret = frmDefVar.NewVariable(sName, sValue, Me)
    If Not Ret Then
        Remove mCount
    Else
        SetName mCount, sName
        Value(mCount) = sValue
    End If
    
    sett_New = Ret
End Function

Public Function sett_Modify(Index As Integer) As Boolean
    Dim sValue As String
    
    Load frmDefVar
    sValue = Me.Value(Index)
    sett_Modify = frmDefVar.ModVariable(Name(Index), sValue)
    If sett_Modify Then Value(Index) = sValue
End Function

Public Sub sett_Save()
    Me.Save False
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
        col(1) = Connect.Lang("cvar", "Name")
        col(2) = Connect.Lang("cval", "Value")
    Set Connect = Nothing
End Sub

Public Property Get sett_Item(ByVal Index As Integer, ByVal col As Integer) As String
    'Dim Combo As cKeyCombo, Key As Integer, Shift As Integer
    
    Select Case col
        Case 1 'name
            sett_Item = Name(Index)
        Case 2 'value
            sett_Item = Value(Index)
    End Select
End Property

Public Sub ReplaceVars(ByRef data As String)
    Dim Pos As Integer, EndPos As Integer, EndPos2 As Integer
    Dim sName As String, sValue As String, Index As Integer

    'replace @... expressions with associated value
    data = Replace(data, "\@", "ÿ")
    Pos = InStr(1, data, "@")
    Do Until Pos = 0
        EndPos = InStr(Pos, data, " ")
        EndPos2 = InStr(Pos, data, vbCrLf)
        If EndPos2 < EndPos And EndPos2 <> 0 Then EndPos = EndPos2
        
        If EndPos = 0 Then EndPos = Len(data) - 1
        sName = Mid$(data, Pos, EndPos - Pos)
        Index = FindName(sName)
        If Index = 0 Then
            sValue = ""
        Else
            sValue = Value(Index)
        End If
        data = Left$(data, Pos - 1) & sValue & Mid$(data, EndPos)
        EndPos = Pos + Len(sValue)
        
        Pos = InStr(EndPos, data, "@")
    Loop
    
    data = Replace(data, "ÿ", "@")
    Debug.Print data
End Sub

Public Sub ProcessCommand(ByVal data As String)
    Dim Pos As Integer
    Dim sName As String, sValue As String

    'process associations, listing and modifications by command line
    data = Trim$(data)
    If Right$(data, 2) = vbCrLf Then data = Left$(data, Len(data) - 2)
    If data = "@" Then
        OutputVarList
    Else
        Pos = InStr(1, data, "=")
        If Not Pos = 0 Then
            sName = Trim$(Left$(data, Pos - 1))
            sValue = Trim$(Mid$(data, Pos + 1))
            If Add(sName, sValue) Then
                Log sName & " was added with value '" & sValue & "'"
            Else
                Log sName & " has now value '" & sValue & "'"
            End If
            Save False
        Else
            Log "Invalid or unknown variable command"
        End If
    End If
End Sub

Private Sub Log(data As String)
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Connect.Envi.sendOutput vbCrLf & TD & "RGB020160255" & TD & "[var] " & TD & "RGB230230230" & TD & data & vbCrLf
    Set Connect = Nothing
End Sub

Private Sub OutputVarList()
    Dim Connect As cConnector
    Dim out As String
    Dim i As Integer
    
    Set Connect = New cConnector
        out = vbCrLf & vbCrLf
        For i = 1 To mCount
            out = out & TD & "RGB020160255" & TD & _
                  Name(i) & TD & "RGB200200200" & TD & " = " & _
                  TD & "RGB230230230" & TD & Value(i) & vbCrLf
        Next i
        out = out & vbCrLf & mCount & " variables" & vbCrLf
        'out = out & "Type @<varname> = <value> to change a value or define a new variable"
        Connect.Envi.sendOutput out
    Set Connect = Nothing
End Sub

Public Sub LoadVars()
    Dim Connect As cConnector
    Dim Path As String, Pos As Integer
    Dim Loa As String, Name As String, Value As String
    Dim Free As Integer
    
    Set Connect = New cConnector
    Path = Connect.ProfileFolder
    If Dir$(Path, vbDirectory) = "" Then
        MkDir Path
        Connect.Log "The directory for the profile was created"
    End If
    
    Path = Path & "vars.ini"
    If Dir$(Path) = "" Then
        Free = FreeFile
        Open Path For Binary As Free: Close Free
    End If
    
    Free = FreeFile
    Open Path For Input As #Free
        Do Until EOF(Free)
            Line Input #Free, Loa
            If Left$(Loa, 1) = "@" Then
                Pos = InStr(1, Loa, "=")
                If Not Pos = 0 Then
                    Name = Left$(Loa, Pos - 1)
                    'Name = Mid$(Name, 2) 'this is to remove the initial @
                    Value = Mid$(Loa, Pos + 1)
                    Add Name, Value, False
                End If
            End If
        Loop
    Close #Free
    
    Call Sort
End Sub

Public Sub Save(Optional bSort As Boolean = True)
    Dim Connect As cConnector
    Dim Path As String
    Dim Free As Integer
    Dim i As Integer
    
    Set Connect = New cConnector
    Path = Connect.ProfileFolder
    If Dir$(Path, vbDirectory) = "" Then
        MkDir Path
    End If
    
    Path = Path & "vars.ini"
    If Dir$(Path) = "" Then
        Free = FreeFile
        Open Path For Binary As Free: Close Free
    End If
    
    If bSort Then Call Sort
    
    Free = FreeFile
    Open Path For Output As #Free
        For i = 1 To mCount
            Print #Free, Name(i) & "=" & Value(i)
        Next i
    Close #Free
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get Name(ByVal Index As Long) As String
    Name = mVars(VAR_NAME, Index)
End Property

Private Sub SetName(ByVal Index As Long, ByVal NewName As String)
    mVars(VAR_NAME, Index) = NewName
End Sub

Public Property Get Value(ByVal Index As Long) As String
    Value = mVars(VAR_VALUE, Index)
End Property

Public Property Let Value(ByVal Index As Long, ByVal data As String)
    mVars(VAR_VALUE, Index) = data
End Property

Public Function FindName(ByVal Name As String) As Integer
    Dim i As Long
    
    Name = LCase$(Trim$(Name))
    For i = 1 To mCount
        If LCase$(mVars(VAR_NAME, i)) = Name Then
            FindName = i
        End If
    Next i
End Function

Public Function Add(ByVal Name As String, ByVal Value As String, Optional ByVal bSort As Boolean = True) As Boolean
    Dim Index As Integer, Bound As Long
    
    'return true if a new variable was added, false if an existing variable was modified
    
    Index = FindName(Name)
    If Index = 0 Then
        'add a new variable to the list
        Index = mCount + 1
        Bound = UBound(mVars, 2)
        If Index > Bound Then
            ReDim Preserve mVars(1 To 2, 1 To (Bound + INCREMENT)) As String
        End If
        mVars(VAR_NAME, Index) = Name
        mCount = Index
        Add = True
    End If
    'update the created variable with the new value
    mVars(VAR_VALUE, Index) = Value
    
    If bSort Then Call Sort
End Function

Public Sub Remove(ByVal Index As Long)
    Dim i As Integer
    
    If Not Index = mCount Then
        For i = Index To mCount - 1
            SetName i, Name(i + 1)
            Value(i) = Value(i + 1)
        Next i
    End If
    mCount = mCount - 1
End Sub

Public Sub Sort()
    Dim exc As String 'exchange variable
    Dim Changed As Boolean 'is changed?
    Dim i As Long
    
    'bubble sort
    Do
        Changed = False
        For i = 1 To mCount - 1
            If LCase$(Name(i)) > LCase$(Name(i + 1)) Then
                'echange names
                exc = Name(i + 1)
                SetName i + 1, Me.Name(i)
                SetName i, exc
                
                'exchange values
                exc = Value(i + 1)
                Me.Value(i + 1) = Me.Value(i)
                Me.Value(i) = exc
                
                'tell the loop that something changed
                Changed = True
            End If
        Next i
    Loop While Changed
End Sub

Private Sub Class_Initialize()
    ReDim mVars(1 To 2, 1 To INCREMENT) As String
    mCount = 0
End Sub
