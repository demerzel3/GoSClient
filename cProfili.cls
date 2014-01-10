VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProfili"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mProfilo() As String
Private mOldCount As Integer
Private mProfSel As Integer

Public Property Get ProfileSel() As Integer
    ProfileSel = mProfSel
End Property

Public Property Get Count() As Integer
    Count = UBound(mProfilo, 2)
End Property

Public Sub Carica()
    Dim Config As cConnector
    Dim Cont As Integer

    Set Config = New cConnector
    With Config
        '.CaricaFile "config.ini"
        mProfSel = .GetConfig("ProfileSelected", 0)
        Cont = 1
        Do
            If .GetConfig("profilo<" & Cont & ">", "") <> "" Then
                Add .GetConfig("profilo<" & Cont & ">"), .GetConfig("pass<" & Cont & ">")
                Cont = Cont + 1
            Else
                Exit Do
            End If
        Loop
    End With
    Set Config = Nothing
    
    mOldCount = Me.Count
End Sub

Public Sub Save()
    Dim i As Integer
    Dim Config As cConnector
    
    Set Config = New cConnector
    With Config
        '.CaricaFile "config.ini"
        For i = 1 To mOldCount
            .Envi.RemoveConfig "profilo<" & i & ">"
            .Envi.RemoveConfig "pass<" & i & ">"
        Next i
        
        For i = 1 To Me.Count
            .SetConfig "profilo<" & i & ">", Me.Nick(i)
            .SetConfig "pass<" & i & ">", Me.Pass(i)
        Next i
    End With
    Config.SaveConfig
    Set Config = Nothing
    
    mOldCount = Me.Count
End Sub

Public Sub Add(Nick As String, Pass As String)
    Dim pCount As Integer

    pCount = UBound(mProfilo, 2) + 1
    If pCount = 1 Then
        ReDim mProfilo(1 To 2, 1 To 1) As String
    Else
        ReDim Preserve mProfilo(1 To 2, 1 To pCount) As String
    End If
    
    mProfilo(1, pCount) = Nick
    mProfilo(2, pCount) = Pass
End Sub

Public Sub Remove(Index As Integer)
    Dim i As Integer
    Dim Count As Integer

    Count = UBound(mProfilo, 2)
    
    If Not Count = 1 Then
        If Not Index = Count Then
            For i = Index To Count - 1
                mProfilo(1, Index) = mProfilo(1, Index + 1)
                mProfilo(2, Index) = mProfilo(2, Index + 1)
            Next i
        End If
        
        ReDim Preserve mProfilo(1 To 2, 1 To Count - 1) As String
    Else
        ReDim mProfilo(0 To 0, 0 To 0) As String
    End If
End Sub

Public Property Let Nick(Index As Integer, data As String)
    mProfilo(1, Index) = data
End Property

Public Property Let Pass(Index As Integer, data As String)
    mProfilo(2, Index) = data
End Property

Public Property Get Nick(Index As Integer) As String
    Dim Connect As cConnector
    
    If Index = 0 Then
        Set Connect = New cConnector
        Nick = Connect.Lang("", "Nobody")
        Set Connect = Nothing
    Else
        Nick = mProfilo(1, Index)
    End If
End Property

Public Property Get Pass(Index As Integer) As String
    Pass = mProfilo(2, Index)
End Property

Private Sub Class_Initialize()
    ReDim mProfilo(0 To 0, 0 To 0) As String
End Sub
