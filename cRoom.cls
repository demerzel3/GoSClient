VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Caption As String

'coordinate relative/////////////////////
Public PosX As Integer
Public PosY As Integer
Public PosZ As Integer
'////////////////////////////////////////

Public Tag As String
'coordinate assolute/////////////////////
Public RealX As Integer
Public RealY As Integer

Public TagX As Integer
Public TagY As Integer
Public TagH As Integer
Public TagW As Integer
'////////////////////////////////////////

Public Image As String

'infos stabilite a runtime
    'cambiamento di piano in questa room
Public ChangeZ As Boolean
