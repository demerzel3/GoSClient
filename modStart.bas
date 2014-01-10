Attribute VB_Name = "modStart"
Option Explicit

Global Const STARTMODE_LIST As Integer = 0
Global Const STARTMODE_AUTO As Integer = 1

Global DEBUGGING As Boolean

Sub Main()
    Dim t As Single
    Dim sDebugMode As String
    Dim rtn As VbMsgBoxResult

    Dim SetPortRtn As Boolean
    
    SetPortRtn = SetPluginsPort

    If App.PrevInstance Then
        If SetPortRtn = False Then
            MsgBox "Another instance of the program is already running"
            Exit Sub
        End If
    End If
    
    
    If InStr(1, Command$, "-debug", vbTextCompare) Then
'            sDebugMode = "You're starting GosClient in debug mode." & vbCrLf & _
'                         "In this mode windows can't be docked into layout" & vbCrLf & _
'                         "and the mouse scroller won't work." & vbCrLf & vbCrLf & _
'                         "Use debug mode?"
'            rtn = MsgBox(sDebugMode, vbInformation Or vbYesNoCancel, "GosClient - Run in debug mode?")
'            If rtn = vbYes Then
'                DEBUGGING = True
'                Debug.Print "...........debug mode active"
'            ElseIf rtn = vbNo Then
'                DEBUGGING = False
'                Debug.Print "...........debug mode not active"
'            ElseIf rtn = vbCancel Then
'                Exit Sub
'            End If
        DEBUGGING = True
    End If

    If SetPortRtn = False Then
        gPluginsPort = 10000
    End If
    
    InitCommonControls
    If Dir$(App.Path & "\plugins\", vbDirectory) = "" Then MkDir App.Path & "\plugins\"
    t = Timer
    frmSplash.Show
    Do Until (Timer - t) >= 0.5
        DoEvents
    Loop
    frmMuds.Init
End Sub

Private Function SetPluginsPort() As Boolean
    Dim i As Integer
    
    If Dir$(App.Path & "\multiple.ini") = "" Then
        SetPluginsPort = False
    Else
        'nCount = GetSetting("GosClient", "multiple", "n", -1)
        'nCount = nCount + 1
        'SaveSetting "GosClient", "multiple", "n", nCount
        
        'check for the first empty port in a range from 10000 to 20000
        For i = 10000 To 20000
            If Not CBool(GetSetting("GosClient", "multiple", "port" & i, False)) Then
                Exit For
            End If
        Next i
        gPluginsPort = i
        SaveSetting "GosClient", "multiple", "port" & gPluginsPort, -1
        
        'MsgBox "plug-in port = " & gPluginsPort
        SetPluginsPort = True
    End If
End Function
