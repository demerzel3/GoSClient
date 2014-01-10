VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub RaiseError(Code As Integer, Descr As String, Source As String)
    Dim Connect As cConnector
    
    MsgBox "Si e' verificato l'errore " & Code & ": " & Descr & _
           " in '" & Source & "'" & vbCrLf & _
           "Per segnalare il bug scrivi a assatarram3@yahoo.it", _
           vbCritical, "Errore"
    
    Set Connect = New cConnector
        Connect.LogError Code, Descr, Source
    Set Connect = Nothing
End Sub
