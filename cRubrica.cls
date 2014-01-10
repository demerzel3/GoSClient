VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRubrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mContatti As Collection

Public Event Added(Index As Integer)
Public Event Removed(Index As Integer)

Public Function ChooseContact() As String
    Dim i As Integer, rtn As Integer
    Dim Connect As cConnector
    
    Load frmList
    For i = 1 To Me.Count
        frmList.AddItem Me.Nick(i)
    Next i
    
    Set Connect = New cConnector
        frmList.Caption = Connect.Lang("rubrica", "Contacts")
    Set Connect = Nothing
    
    rtn = frmList.ShowForm
    rtn = rtn + 1
    If Not rtn = 0 Then ChooseContact = Me.Nick(rtn)
End Function

Public Sub Sort()
    Dim Scambiati As Boolean
    Dim Scambio As cContact
    Dim Contact() As cContact
    Dim i As Integer, Count As Integer

    Count = mContatti.Count
    ReDim Contact(1 To Count) As cContact
    For i = 1 To Count
        Set Contact(i) = mContatti.Item(i)
    Next i
    Set mContatti = Nothing
    
    Scambiati = True
    Do While Scambiati
        Scambiati = False
        For i = 1 To Count - 1
            If Contact(i).acNick > Contact(i + 1).acNick Then
                Set Scambio = Contact(i)
                Set Contact(i) = Contact(i + 1)
                Set Contact(i + 1) = Scambio
                Scambiati = True
            End If
        Next i
    Loop
    
    Set mContatti = New Collection
    For i = 1 To Count
        mContatti.Add Contact(i), Contact(i).acNick
        Set Contact(i) = Nothing
    Next i
    
    Erase Contact()
End Sub

Public Sub LoadRubrica()
    Dim Load As cBinary
    
    Set Load = New cBinary
        Load.LoadRubrica Me
    Set Load = Nothing
End Sub

Public Sub SaveRubrica()
    Dim Save As cBinary

    Set Save = New cBinary
        Save.SaveRubrica Me
    Set Save = Nothing
End Sub

Public Property Get Count() As Integer
    Count = mContatti.Count
End Property

Public Property Get Note(Index As Integer) As String
    Note = mContatti.Item(Index).acNote
End Property

Public Property Let Note(Index As Integer, data As String)
    mContatti.Item(Index).acNote = data
End Property

Public Property Get Url(Index As Integer) As String
    Url = mContatti.Item(Index).acURL
End Property

Public Property Let Url(Index As Integer, data As String)
    mContatti.Item(Index).acURL = data
End Property

Public Property Get Nick(Index As Integer) As String
    Nick = mContatti.Item(Index).acNick
End Property

Public Property Let Nick(Index As Integer, data As String)
    mContatti.Item(Index).acNick = data
End Property

Public Property Get Name(Index As Integer) As String
    Name = mContatti.Item(Index).acName
End Property

Public Property Let Name(Index As Integer, data As String)
    mContatti.Item(Index).acName = data
End Property

Public Property Get Email(Index As Integer) As String
    Email = mContatti.Item(Index).acEmail
End Property

Public Property Let Email(Index As Integer, data As String)
    mContatti.Item(Index).acEmail = data
End Property

Public Sub Add(Nick As String, Optional Name As String, Optional Email As String, _
               Optional Url As String, Optional Note As String)
    Dim newContact As cContact
    
    Set newContact = New cContact
    With newContact
        .acNick = Nick
        .acName = Name
        .acEmail = Email
        .acURL = Url
        .acNote = Note
    End With
    
    On Error GoTo fine
    mContatti.Add newContact, Nick
    Set newContact = Nothing
    
    RaiseEvent Added(mContatti.Count)

Exit Sub
fine:
    Set newContact = Nothing
End Sub

Public Sub Remove(Index As Integer)
    mContatti.Remove Index

    RaiseEvent Removed(Index)
End Sub

Private Sub Class_Initialize()
    Set mContatti = New Collection
End Sub

Private Sub Class_Terminate()
    Set mContatti = Nothing
End Sub
