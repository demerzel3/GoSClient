Attribute VB_Name = "modKernel32"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub rltMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal lLen As Long)
