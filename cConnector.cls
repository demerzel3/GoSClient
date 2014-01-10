VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get Lang() As cLang
    Set Lang = gEnvi.GetLang
End Property

Public Property Get ProfConf() As cIni
    Set ProfConf = gEnvi.ProfConfig
End Property

Public Sub LogError(Number As Integer, Descr As String, Src As String)
    Dim Final As String
    
    If Number = 0 Then
        Final = """" & Descr & """ in " & Src
    Else
        Final = Number & ", """ & Descr & """ in " & Src
    End If
    gEnvi.AddToStatusErr Final
End Sub

Public Sub LogYellow(Text As String)
    gEnvi.AddToStatus TD & "RGB255230000" & TD & "  " & Text
End Sub

Public Sub Log(Text As String)
    gEnvi.AddToStatus Text
End Sub

Public Property Get Rubrica() As cRubrica
    Set Rubrica = gEnvi.GetRubrica
End Property

Public Property Get Palette() As cPalette
    Set Palette = gEnvi.Palette
End Property

Public Property Get Envi() As cEnviron
    Set Envi = gEnvi
End Property

Public Sub RegWin(win As cFinestra)
    gEnvi.RegWin win
End Sub

Public Sub UnRegWin(winID As Integer)
    gEnvi.UnRegWin winID
End Sub

Public Property Get ProfileSel() As Integer
    ProfileSel = gEnvi.ProfSel
End Property

Public Property Get ProfileFolder() As String
    ProfileFolder = gEnvi.ProfSelFolder
End Property

Public Sub SaveConfig()
    gEnvi.SaveConfig
End Sub

Public Function GetBoolConfig(Key As String, _
                Optional Default As Boolean = False) As Boolean
    GetBoolConfig = CBool(Val(gEnvi.GetConfig(Key, Default)))
End Function

Public Function GetConfig(Key As String, Optional Default As Variant = 0) As Variant
    GetConfig = gEnvi.GetConfig(Key, Default)
End Function

Public Sub SetBoolConfig(Key As String, Value As Boolean)
    gEnvi.SetConfig Key, CInt(Value)
End Sub

Public Sub SetConfig(Key As String, Value As Variant)
    gEnvi.SetConfig Key, Value
End Sub

Public Sub SetProfileSel(Profile As Integer)
    Dim Profiles As cProfili
    
    gEnvi.ProfSel = Profile
    Me.SetConfig "ProfileSelected", Profile
    Me.SaveConfig
    
    Set Profiles = New cProfili
        Profiles.Carica
        If Not Profile = 0 Then
            gEnvi.ProfSelFolder = gMudPath & "p_" & LCase$(Profiles.Nick(Profile)) & "\"
        Else
            gEnvi.ProfSelFolder = gMudPath & "p_(nessuno)\"
        End If
        If Dir$(gEnvi.ProfSelFolder, vbDirectory) = "" Then MkDir gEnvi.ProfSelFolder
    Set Profiles = Nothing
    gEnvi.SaveProfConfig
    gEnvi.LoadProfConfig
End Sub

Private Sub Class_Initialize()
    If gEnvi Is Nothing Then
        Set gEnvi = New cEnviron
    End If
    gEnvi.NReferenze = gEnvi.NReferenze + 1
End Sub

Private Sub Class_Terminate()
    gEnvi.NReferenze = gEnvi.NReferenze - 1
    If gEnvi.NReferenze = 0 Then
        Set gEnvi = Nothing
        Debug.Print "Killing di gEnvi avvenuto :)"
    End If
End Sub
