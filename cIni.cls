VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mInfo As Collection
Private mKeys As Collection
Private mPath As String

Public Sub SalvaFile()
    Dim Free As Integer
    Dim i As Integer

    If Not mInfo Is Nothing Then
        Free = FreeFile
        Open mPath For Output As #Free
            For i = 1 To mInfo.Count
                Print #Free, mKeys.Item(i) & " = " & mInfo.Item(i)
            Next i
        Close #Free
    End If
End Sub

Public Sub AddInfo(Nome As String, Valore As Variant)
    On Error Resume Next
    mInfo.Remove Nome
    mKeys.Remove Nome
    mInfo.Add Valore, Nome
    mKeys.Add Nome, Nome
End Sub

Public Sub RemoveInfo(Nome As String)
    On Error Resume Next
    mInfo.Remove Nome
    mKeys.Remove Nome
End Sub

Public Function CaricaFile(Name As String, Optional AbsPath As Boolean = False) As Boolean
    'mPath = App.Path & "\" & Name
    If AbsPath Then
        mPath = Name
    Else
        mPath = gMudPath & Name
    End If
    CaricaFile = CaricaInfoFile(mPath)
End Function

Private Sub Class_Initialize()
    Set mInfo = New Collection
    Set mKeys = New Collection
End Sub

Private Function EstraiInfo(Stringa As String, ByRef Key As String) As String
    EstraiInfo = Trim$(Mid$(Stringa, InStr(1, Stringa, "=", vbTextCompare) + 1))
    Key = Trim$(Mid$(Stringa, 1, InStr(1, Stringa, "=", vbTextCompare) - 2))
End Function

Private Function CaricaInfoFile(Path As String) As Boolean
    Dim Loa As String
    Dim NFile As Integer

    NFile = FreeFile
    On Error GoTo fine
        
    Open Path For Input As #NFile
        Do Until EOF(NFile)
            Line Input #NFile, Loa
            Loa = Trim$(Loa)
            If InStr(1, Loa, "=", vbTextCompare) <> 0 Then AggiungiInfo Loa
        Loop
    Close #NFile
    CaricaInfoFile = True
    
Exit Function
fine:
    CaricaInfoFile = False
End Function

Public Function RetrInfo(Key As String, Optional Default As Variant = 0) As Variant
    On Error GoTo Errore
    RetrInfo = mInfo(Key)

    Exit Function
Errore:
    RetrInfo = Default
End Function

Private Sub AggiungiInfo(Stringa As String)
    Dim Key As String
    Dim Info As Variant

    Info = EstraiInfo(Stringa, Key)
    On Error Resume Next
    mInfo.Remove Key
    mKeys.Remove Key
    mInfo.Add Info, Key
    mKeys.Add Key, Key
End Sub

Private Sub Class_Terminate()
    Set mInfo = Nothing
    Set mKeys = Nothing
End Sub
