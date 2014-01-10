VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event envOutput(data As String, OutType As Integer)
Public Event envInput(data As String, InType As Integer)
Public Event envNotify(uMsg As Long)

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
        Debug.Print "Killing di gEnvi avvenuto by cInOut"
    End If
End Sub

