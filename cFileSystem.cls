VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function FileCopyEx(Src As String, Dest As String) As Boolean
    If Dir$(Src) = "" Then
        FileCopyEx = False
    Else
        FileCopy Src, Dest
        FileCopyEx = True
    End If
End Function

Public Sub DirCopy(ByVal Src As String, ByVal Dest As String)
    Src = NormalizeDirPath(Src)
    Dest = NormalizeDirPath(Dest)

    DirManage Src, Dest, True
End Sub

Public Sub DirDelete(Src As String)
    Src = NormalizeDirPath(Src)

    DirManage Src, "", False, True
End Sub

Private Function NormalizeDirPath(ByVal Path As String) As String
    Path = Trim$(Path)
    Path = IIf(Right$(Path, 1) = "\", Path, Path & "\")
    NormalizeDirPath = Path
End Function

Private Sub DirManage(ByVal Src As String, Optional ByVal Dest As String, _
    Optional bCopy As Boolean = False, Optional bDel As Boolean = False)

    Dim Dirs As Collection, fEnum As String
    Dim i As Integer

    If bCopy = False And bDel = False Then Exit Sub
    
    Set Dirs = New Collection
    
    If bCopy Then DirCreate Dest
    
    fEnum = Dir$(Src, vbDirectory)
    Do Until fEnum = ""
        If fEnum <> "." And fEnum <> ".." Then
            If (GetAttr(Src & fEnum) And vbDirectory) = vbDirectory Then
                Dirs.Add fEnum & "\"
            Else
                If bCopy Then
                    FileCopy Src & fEnum, Dest & fEnum
                End If
            
                If bDel Then
                    Kill Src & fEnum
                End If
            End If
        End If
        fEnum = Dir$()
    Loop
    
    If Dirs.Count > 0 Then
        For i = 1 To Dirs.Count
            DirManage Src & Dirs.Item(i), Dest & Dirs.Item(i), bCopy, bDel
        Next i
    End If
    Set Dirs = Nothing

    If bDel Then DirKill Src
End Sub

Private Sub DirKill(Path As String)
    On Error Resume Next
    RmDir Path
    On Error GoTo 0
End Sub

Private Sub FileKill(Path As String)
    On Error Resume Next
    Kill Path
    On Error GoTo 0
End Sub

Public Function DirExist(Path As String) As Boolean
    If Dir$(Path, vbDirectory) = "" Then
        DirExist = False
    Else
        DirExist = True
    End If
End Function

Public Function FileExist(Path As String) As Boolean
    If Dir$(Path) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Sub DirCreate(Path As String)
    If Not DirExist(Path) Then MkDir Path
End Sub
