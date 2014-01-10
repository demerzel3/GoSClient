VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mBoxes As Collection
Private mPath As String
Private mLyt As String

Public Property Get LytFile() As String
    LytFile = mLyt
End Property

Public Property Get Count() As Integer
    Count = mBoxes.Count
End Property

Public Property Get box(Index As Integer) As cBox
    Set box = mBoxes.Item(Index)
End Property

Public Sub RemoveBox(Index As Integer)
    mBoxes.Remove Index
End Sub

Public Sub AddBox(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim NewBox As cBox

    Set NewBox = New cBox
    With NewBox
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    
    mBoxes.Add NewBox
    
    Set NewBox = Nothing
End Sub

Public Sub LoadLayout(Path As String)
    Dim lyt As cIni, file As String
    Dim Count As Integer, i As Integer
    Dim NewBox As cBox
    Dim FileExist As Boolean, Free As Integer, ToPrint As String
    Dim Connect As cConnector

    file = Path
    mLyt = file
    Path = App.Path & "\layouts\" ' & Path
    If Dir$(Path, vbDirectory) = "" Then MkDir Path
    Path = Path & file
    mPath = Path
    Set lyt = New cIni
    FileExist = lyt.CaricaFile(Path, True)
    If Not FileExist Then
        Free = FreeFile
        ToPrint = "Box_Count = 1" & vbCrLf & _
                  "Box<1>_Left = 0" & vbCrLf & _
                  "Box<1>_Top = 0" & vbCrLf & _
                  "Box<1>_Right = 100" & vbCrLf & _
                  "Box<1>_Bottom = 100" & vbCrLf & _
                  "Win<1>_Parent = 1" & vbCrLf & _
                  "Win<1> = frmMain" & vbCrLf & _
                  "Win_Count = 1"
        Open Path For Output As #Free
            Print #Free, ToPrint
        Close #Free
        lyt.CaricaFile Path, True
    End If
    
    Set mBoxes = Nothing
    Set mBoxes = New Collection
    
    Count = Val(lyt.RetrInfo("Box_Count"))
    For i = 1 To Count
        Set NewBox = New cBox
        NewBox.Left = Val(lyt.RetrInfo("Box<" & i & ">_Left"))
        NewBox.Top = Val(lyt.RetrInfo("Box<" & i & ">_Top"))
        NewBox.Right = Val(lyt.RetrInfo("Box<" & i & ">_Right"))
        NewBox.Bottom = Val(lyt.RetrInfo("Box<" & i & ">_Bottom"))
        mBoxes.Add NewBox
        Set NewBox = Nothing
    Next i
    
    Count = Val(lyt.RetrInfo("Win_Count"))
    For i = 1 To Count
        Set NewBox = mBoxes.Item(Val(lyt.RetrInfo("Win<" & i & ">_parent")))
        NewBox.Window = lyt.RetrInfo("Win<" & i & ">")
        Set NewBox = Nothing
    Next i
        
    Set lyt = Nothing
    
    Set Connect = New cConnector
        Connect.Log "Caricato Layout da \layouts\" & file
    Set Connect = Nothing
End Sub

Public Sub SaveLayout(Optional OnlyBoxes As Boolean = False, Optional Path As String = "")
    Dim Free As Integer, frmCount As Integer
    Dim i As Integer
    
    If Not Path = "" Then
        mPath = App.Path & "\layouts\" & Path
        mLyt = Path
    End If
    
    If Not mPath = "" Then
        Free = FreeFile
        Open mPath For Output As #Free
            Print #Free, "Box_Count = " & Me.Count
            For i = 1 To Me.Count
                With Me.box(i)
                    Print #Free, "Box<" & i & ">_Left = " & .Left
                    Print #Free, "Box<" & i & ">_Top = " & .Top
                    Print #Free, "Box<" & i & ">_Right = " & .Right
                    Print #Free, "Box<" & i & ">_Bottom = " & .Bottom
                    If Not OnlyBoxes Then
                        If .Window <> "" Then
                            frmCount = frmCount + 1
                            Print #Free, "Win<" & frmCount & ">_Parent = " & i
                            Print #Free, "Win<" & frmCount & "> = " & .Window
                        End If
                    End If
                End With
            Next i
            If Not OnlyBoxes Then Print #Free, "Win_Count = " & frmCount
        Close #Free
    End If
End Sub

Private Sub Class_Initialize()
    Set mBoxes = New Collection
End Sub

Private Sub Class_Terminate()
    Set mBoxes = Nothing
End Sub
