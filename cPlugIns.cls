VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlugIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPlugIn"
Attribute VB_Ext_KEY = "Member0" ,"cPlugIn"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variabile locale per memorizzare l'insieme.
Private mCol As Collection

Public Sub VerifyConfig()
    Dim i As Integer
    
    For i = 1 To mCol.Count
        mCol.Item(i).VerifyConfig
    Next i
End Sub

Public Sub SaveAutoLoaded()
    Dim Conf As cIni, Connect As cConnector
    Dim Cont As Integer, i As Integer
    
    Set Connect = New cConnector
        Set Conf = Connect.ProfConf
    Set Connect = Nothing
    
    'Set Conf = New cIni
        'Conf.CaricaFile App.Path & "\config.ini", True
        Cont = 0
        For i = 1 To Me.Count
            If Me.Item(i).Auto Then
                Cont = Cont + 1
                Conf.AddInfo "plugin<" & Cont & ">", Me.Item(i).GetPath
            End If
        Next i
        Conf.AddInfo "plugin_count", Cont
        Conf.SalvaFile
    Set Conf = Nothing
End Sub

Public Sub UpdateAutoStart()
    Dim Conf As cIni, Auto As Collection, file As String
    Dim Connect As cConnector, i As Integer
    Dim Cont As Integer
    
    Set Connect = New cConnector
        Set Conf = Connect.ProfConf
    Set Connect = Nothing

    Set Auto = New Collection
        Cont = Val(Conf.RetrInfo("plugin_count", 0))
        For i = 1 To Cont
            file = Conf.RetrInfo("plugin<" & i & ">", "")
            Auto.Add file, file
        Next i
    Set Conf = Nothing
    
    For i = 1 To mCol.Count
        On Error Resume Next
        file = Auto.Item(mCol.Item(i).GetPath)
        If Err.Number <> 0 Then
            Err.Clear
            mCol.Item(i).Auto = False
        Else
            mCol.Item(i).Auto = True
        End If
    Next i
    
    Set Auto = Nothing
End Sub

Public Sub LoadList()
    Dim fEnum As String, Plugins(1 To 1000) As cPlugIn
    Dim newPlugin As cPlugIn
    Dim i As Integer, Cont As Integer, Max As Integer
    Dim Scambio As cPlugIn, Scambiati As Boolean
    
    Dim Conf As cIni, Auto As Collection, file As String
    Dim Connect As cConnector
    
    Set Connect = New cConnector
        Set Conf = Connect.ProfConf
    Set Connect = Nothing

    'Set Conf = New cIni
    Set Auto = New Collection
        'Conf.CaricaFile App.Path & "\config.ini", True
        Cont = Val(Conf.RetrInfo("plugin_count", 0))
        For i = 1 To Cont
            file = Conf.RetrInfo("plugin<" & i & ">", "")
            Auto.Add file, file
        Next i
    Set Conf = Nothing
    
    fEnum = Dir$(App.Path & "\plugins\") 'real
    Cont = 0
    Max = UBound(Plugins, 1)
    Do Until fEnum = ""
        If LCase$(Right$(fEnum, 4)) = ".dll" Then
            Set newPlugin = New cPlugIn
                If newPlugin.LoadInfo(fEnum) Then
                
                    On Error Resume Next
                    file = Auto.Item(fEnum)
                    If Err.Number <> 0 Then
                        Err.Clear
                        newPlugin.Auto = False
                    Else
                        newPlugin.Auto = True
                    End If
                    
                    Cont = Cont + 1
                    If Cont > Max Then Exit Do
                    Set Plugins(Cont) = newPlugin
                End If
            Set newPlugin = Nothing
        End If
        fEnum = Dir$()
    Loop

    Set Auto = Nothing

    Scambiati = True
    Do While Scambiati = True
        Scambiati = False
        For i = 1 To Cont - 1
            If Plugins(i).Title > Plugins(i + 1).Title Then
                Set Scambio = Plugins(i)
                Set Plugins(i) = Plugins(i + 1)
                Set Plugins(i + 1) = Scambio
                Set Scambio = Nothing
                Scambiati = True
            End If
        Next i
    Loop

    For i = 1 To Cont
        Add Plugins(i)
        Set Plugins(i) = Nothing
    Next i
End Sub

Public Function Add(Item As cPlugIn, Optional sKey As String) As cPlugIn
    If Len(sKey) = 0 Then
        mCol.Add Item
    Else
        mCol.Add Item, sKey
    End If
End Function

Public Property Get Item(vntIndexKey As Variant) As cPlugIn
Attribute Item.VB_UserMemId = 0
    'Utilizzato per fare riferimento a un elemento nell'insieme.
    'vntIndexKey contiene la chiave o l'indice dell'insieme,
    'e per questo motivo è dichiarata come Variant.
    'Sintassi: Set foo = x.Item(xyz) oppure Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'Utilizzata per recuperare il numero di elementi dell'insieme.
    'Sintassi: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'Utilizzata per rimuovere un elemento dall'insieme.
    'vntIndexKey contiene l'indice o la chiave, e per questo
    'motivo viene dichiarata come Variant.
    'Sintassi: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Private Sub Class_Initialize()
    'Crea l'insieme quando viene creata questa classe.
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'Rimuove l'insieme quando la classe viene eliminata.
    Set mCol = Nothing
End Sub

