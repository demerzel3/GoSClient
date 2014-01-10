Attribute VB_Name = "modDebug"
Option Explicit

Public Function DebugString(s As String)
    frmDebug.Log s
End Function
