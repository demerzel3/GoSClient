VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mName As String
Public Host As String
Private mPort As Long
Public Descr As String
Public Lang As String
Public Comment As String
Public Folder As String

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(data As String)
    mName = data
    Folder = Replace(LCase$(Trim$(mName)), " ", "_")
    Folder = Replace(Folder, "\", "_")
    Folder = Replace(Folder, "/", "_")
    Folder = Replace(Folder, ":", "_")
    Folder = Replace(Folder, "*", "_")
    Folder = Replace(Folder, "?", "_")
    Folder = Replace(Folder, """", "_")
    Folder = Replace(Folder, "<", "_")
    Folder = Replace(Folder, ">", "_")
    Folder = Replace(Folder, "|", "_")
End Property

Public Property Get Port() As Long
    Port = mPort
End Property

Public Property Let Port(ByVal data As Long)
    data = Abs(data)
    If data > 65535 Then data = 65535
    mPort = data
End Property
