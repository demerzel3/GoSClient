VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTriggers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cTrigger"
Attribute VB_Ext_KEY = "Member0" ,"cTrigger"
Option Explicit

Public Imported As Boolean

'Variabile locale per memorizzare l'insieme.
Private mCol As Collection
Private mPath As String

Public Property Get sett_ColCount() As Integer
    sett_ColCount = 2
End Property

Public Sub sett_Delete(Index As Integer)
    Me.Remove Index
End Sub

Public Function sett_New() As Boolean
    'Load frmDefTrigger
    sett_New = frmDefTrigger.NewTrigger(Me)
End Function

Public Function sett_Modify(Index As Integer) As Boolean
    'Load frmDefTrigger
    sett_Modify = frmDefTrigger.ModTrigger(Me.Item(Index))
End Function

Public Sub sett_Save()
    Me.Save
End Sub

Public Property Get sett_Sorted() As Boolean
    sett_Sorted = True
End Property

Public Property Get sett_Count() As Integer
    sett_Count = Me.Count
End Property

Public Sub sett_Column(Col() As String)
    Dim Connect As cConnector
    
    ReDim Col(1 To 2) As String
    Set Connect = New cConnector
        Col(1) = Connect.Lang("calias", "Expression")
        Col(2) = Connect.Lang("deftrigger", "Reaction")
    Set Connect = Nothing
End Sub

Public Property Get sett_Item(ByVal Index As Integer, ByVal Col As Integer) As String
    Select Case Col
        Case 1 'expression
            sett_Item = Me.Item(Index).GetText
            'sett_Item = Me.Text(Index)
        Case 2 'reaction
            sett_Item = Me.Item(Index).FirstLineReact
            'Set Combo = New cKeyCombo
            '    Combo.ScindiCombo Me.Item(Index).KeyCombo, Key, Shift
            '    sett_Item = Combo.IdentCombo(Key, Shift)
            'Set Combo = Nothing
    End Select
End Property





Public Sub Load(Optional Path As String)
    Dim Trig As cTrigger, fIn As cBinary, Old As cAlias
    Dim Connect As cConnector, i As Integer, j As Integer
    
    Set mCol = Nothing
    Set mCol = New Collection
    
    Set Connect = New cConnector
        If Path = "" Then Path = Connect.ProfileFolder & "trigger.gos"
    Set Connect = Nothing
    
    Set fIn = New cBinary
        fIn.LoadTriggers Me, Path
    Set fIn = Nothing
    mPath = Path
    
    'if old triggers haven't been imported yet...
    If Not Me.Imported Then
        Set Old = New cAlias
            Old.LoadTriggers
            For i = 1 To Old.Count
                Set Trig = New cTrigger
                Trig.SetText Old.Text(i)
                Trig.Reaction = ""
                For j = 1 To Old.ActionCount(i)
                    If Not j = 1 Then Trig.Reaction = Trig.Reaction & vbCrLf
                    Trig.Reaction = Trig.Reaction & Old.Action(i, j)
                Next j
                
                Me.Add Trig
                Set Trig = Nothing
            Next i
        Set Old = Nothing
        
        Me.Imported = True
    End If
End Sub

Public Sub Save()
    Dim fOut As cBinary
    
    'save triggers in mpath
    Set fOut = New cBinary
        fOut.SaveTriggers Me, mPath
    Set fOut = Nothing
End Sub

Public Sub Add(Trig As cTrigger, Optional sKey As String)
    mCol.Add Trig
End Sub

Public Property Get Item(vntIndexKey As Variant) As cTrigger
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
