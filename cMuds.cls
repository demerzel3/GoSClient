VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMuds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mMuds As Collection

Public Property Get Mud(Index As Integer) As cMud
    Set Mud = mMuds.Item(Index)
End Property

Public Property Get Count() As Integer
    Count = mMuds.Count
End Property

Public Sub SaveMudList()
    Dim bin As cBinary
    
    Set bin = New cBinary
        bin.SaveMudList Me
    Set bin = Nothing
End Sub

Public Sub LoadMudList()
    Dim bin As cBinary
    
    Set bin = New cBinary
        bin.LoadMudList Me
    Set bin = Nothing
End Sub

Public Sub Add(Name As String, Host As String, Port As Long, Descr As String, Lang As String, Comment As String)
    Dim NewMud As cMud

    Set NewMud = New cMud
    With NewMud
        .Comment = Comment
        .Name = Name
        .Host = Host
        .Descr = Descr
        .Lang = Lang
        .Port = Port
    End With
    mMuds.Add NewMud
    Set NewMud = Nothing
End Sub

Public Sub Remove(Index As Integer)
    mMuds.Remove Index
End Sub

Public Property Get Name(Index As Integer) As String
    Name = mMuds(Index).Name
End Property

Public Property Get Host(Index As Integer) As String
    Host = mMuds(Index).Host
End Property

Public Property Get Port(Index As Integer) As Long
    Port = mMuds(Index).Port
End Property

Public Property Get Descr(Index As Integer) As String
    Descr = mMuds(Index).Descr
End Property

Public Property Get Lang(Index As Integer) As String
    Lang = mMuds(Index).Lang
End Property

Public Property Get Comment(Index As Integer) As String
    Comment = mMuds(Index).Comment
End Property

Public Property Get Folder(Index As Integer) As String
    Folder = mMuds(Index).Folder
End Property

Private Sub Class_Initialize()
    Set mMuds = New Collection
End Sub

Private Sub Class_Terminate()
    Set mMuds = Nothing
End Sub
