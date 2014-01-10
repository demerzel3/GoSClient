Attribute VB_Name = "modMapper"
Option Explicit

Public Enum Mapper_Mov
    nord = 0
    sud = 1
    est = 2
    ovest = 3
    alto = 4
    basso = 5
End Enum

Global Const MAPMODE_PAUSE As Integer = 0
Global Const MAPMODE_ONLINE As Integer = 1
Global Const MAPMODE_OFFLINE As Integer = 2

Global Const ROOM_DIM As Integer = 25
Global Const ROOM_MARG As Integer = 4
