VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private mFile As String
Private mAuthor As String
Private mEmail As String
Private mLang As String
Private mLangID As Integer

Public Event RefreshLang()

Public Property Get LangID() As Integer
    LangID = mLangID
End Property

Public Property Get Language() As String
    Language = mLang
End Property

'Public Property Get Refresh()
'    RaiseEvent RefreshLang
'End Property

Public Property Get Author() As String
    Author = mAuthor
End Property

Public Property Get Email() As String
    Email = mEmail
End Property

Public Function LoadLang(ByRef Filename As String) As Boolean
    Dim Path As String
    
    Path = App.Path & "\lang\"
    If Dir$(Path, vbDirectory) = "" Then MkDir Path
    Path = Path & Filename
    If Dir$(Path) = "" Then
        LoadLang = False
        mFile = ""
        mLang = ""
        mAuthor = ""
        mEmail = ""
        mLangID = 0
    Else
        LoadLang = True
        mFile = Path
        mLang = UCase$(Left$(Filename, 1)) & LCase$(Mid$(Filename, 2, Len(Filename) - 5))
        mAuthor = GetString("translation", "Author", "unknown")
        mEmail = GetString("translation", "authorEmail", "unknown@unknown.com")
        mLangID = GetString("translation", "LangID", 2)
        RaiseEvent RefreshLang
    End If
End Function

Public Function GetString(ByRef sSection As String, ByRef sKey As String, Optional ByRef sDefault = "") As String
Attribute GetString.VB_UserMemId = 0
    Dim sBuff As String, Length As Long
    
    If sDefault = "" Then sDefault = sKey
    If Not mFile = "" Then
        If sSection = "" Then sSection = "general"
        sBuff = Space(128)
        Length = GetPrivateProfileString(sSection, sKey, sDefault, sBuff, 128, mFile)
        GetString = Left$(sBuff, Length)
    Else
        GetString = sDefault
    End If
End Function
