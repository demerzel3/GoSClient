VERSION 5.00
Begin VB.UserControl uDocking 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox cmdDrag 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   450
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      Top             =   2475
      Visible         =   0   'False
      Width           =   3840
   End
   Begin GoS.uDockBox box 
      Height          =   1215
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   2143
   End
   Begin VB.PictureBox pctDest 
      Enabled         =   0   'False
      HasDC           =   0   'False
      Height          =   690
      Left            =   525
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox pctMouseDest 
      Enabled         =   0   'False
      Height          =   765
      Left            =   525
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "uDocking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DRAGMARGIN As Integer = 2

Private Enum DockLocation
    dl_none = 0
    dl_left = 1
    dl_top = 2
    dl_right = 3
    dl_bottom = 4
End Enum

Private mWins As Collection 'window rectangles
Private mBoxes() As RECT
Private mIgnoreRemove As Boolean
Private mBoxToRemove As Integer

'mouse management
Private mDragging As Boolean
Private mDragStartX As Long
Private mDragStartY As Long
Private mMouseStart As POINTAPI
Private mDragLoc As DockLocation
Private mDragStartPos As Integer

'locked state
Private mLocked As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Length As Long)

Public Event LockChanged(ByVal Value As Boolean)
Public Event MouseWheel(ByVal Delta As Integer)

Public Property Get Locked() As Boolean
    Locked = mLocked
End Property

Public Sub SetLocked(ByVal Value As Boolean)
    Dim i As Integer
    
    If Not mLocked = Value Then
        'set the lock to all boxes
        For i = 1 To box.Count - 1
            box(i).SetLocked Value
        Next i
        
        mLocked = Value
        'If mLocked Then
        '    goshStopHook
        'Else
        '    goshStartHook frmBase.hWnd, pctDest.hWnd, pctMouseDest.hWnd, App.ThreadID
        'End If
        goshSetLocked mLocked
        
        RaiseEvent LockChanged(Value)
    End If
End Sub

Public Sub CloseSpareBoxes()
    Dim i As Integer
    
    For i = 1 To box.Count - 1
        On Error Resume Next
        If box(i).IsHWndValid = False Then RemoveBox i
        On Error GoTo 0
    Next i
End Sub

Public Sub LoadLayout()
    Dim src As cIni, i As Integer, Count As Integer, w As cWin
    Dim winID As String, frm As Form, continue As Boolean
    Dim hWnd As Long, Connect As cConnector
    
    Set src = New cIni
    'Src.CaricaFile App.Path & "\layout.lyt", True
    src.CaricaFile "layout.lyt"
    Count = src.RetrInfo("count")
    For i = 1 To Count
        If i = 1 Then
            Set w = mWins.Item(1)
        Else
            Set w = New cWin
        End If
        w.Left = src.RetrInfo("win<" & i & ">_left")
        w.Top = src.RetrInfo("win<" & i & ">_top")
        w.Right = src.RetrInfo("win<" & i & ">_right")
        w.Bottom = src.RetrInfo("win<" & i & ">_bottom")
        w.sCaption = src.RetrInfo("win<" & i & ">_caption")
        If Not i = 1 Then mWins.Add w
        Set w = Nothing
    Next i
    
    SetBoxes
    
    For i = 2 To Count
        winID = src.RetrInfo("win<" & i & ">")
        continue = True
        Select Case winID
            Case "gos.buttons"
                Set frm = frmButtons
            'Case "gos.chat"
            '    Set frm = frmChat
            Case "gos.mapper"
                Set frm = frmMapper
            Case "gos.rubrica"
                Set frm = frmRubrica
            Case "gos.stato"
                Set frm = frmStato
            Case Else
                continue = False
        End Select
        
        If continue Then
            'the window belongs to gos
            Load frm
            goshSetOwner frm.hWnd, frmBase.hWnd
            box(i).Dock frm.hWnd
            goshSetDocked frm.hWnd
        Else
            'the window belongs to a plug-in
            hWnd = goshFindWindow(winID)
            'MsgBox hwnd
            If hWnd = 0 Then
                Set Connect = New cConnector
                    Connect.Log "Unable to find window """ & mWins.Item(i).sCaption & """ (" & winID & ") " & _
                                     "If that window belongs to a plug-in, set it as auto-starting"
                Set Connect = Nothing
                RemoveBox i
            Else
                box(i).Dock hWnd
                goshSetDocked hWnd
            End If
        End If
        
        Set frm = Nothing
    Next i
    
    SetLocked CBool(Val(src.RetrInfo("Locked", 0)))
    Set src = Nothing
    
    SetBoxes
End Sub

Public Sub SaveLayout()
    Dim Dest As cIni, i As Integer, w As cWin
    
    Set Dest = New cIni
    'Dest.CaricaFile App.Path & "\layout.lyt", True
    Dest.CaricaFile "layout.lyt"
    Dest.AddInfo "Count", mWins.Count
    For i = 1 To mWins.Count
        Set w = mWins.Item(i)
        Dest.AddInfo "win<" & i & ">_left", w.Left
        Dest.AddInfo "win<" & i & ">_top", w.Top
        Dest.AddInfo "win<" & i & ">_right", w.Right
        Dest.AddInfo "win<" & i & ">_bottom", w.Bottom
        Dest.AddInfo "win<" & i & ">", goshGetWindowID(box(i).DockedHWnd)
        Dest.AddInfo "win<" & i & ">_caption", box(i).GetCaption
        Set w = Nothing
    Next i
    Dest.AddInfo "Locked", CInt(mLocked)
    Dest.SalvaFile
    Set Dest = Nothing
End Sub

Public Sub SetMainWindow(ByVal hWnd As Long)
    goshSetMainWnd hWnd
    If box(1).DockedHWnd <> 0 Then box(1).UnDock SW_SHOW
    box(1).Dock hWnd
    box(1).SetStatic True
End Sub

Private Sub RemoveBox(Index As Integer, Optional ByVal Virtual As Boolean = False)
    Dim i As Integer, win As cWin, rcWin As RECT
    Dim rcWins() As RECT
    Dim rcLeft() As Integer, rcTop() As Integer, rcRight() As Integer, rcBottom() As Integer
    Dim nLeft As Integer, nTop As Integer, nRight As Integer, nBottom As Integer
    
    Dim wWidth As Long, wHeight As Long
    
    If mWins.Count = 1 Then Exit Sub
    
    ReDim rcWins(1 To mWins.Count) As RECT
    ReDim rcLeft(1 To mWins.Count) As Integer
    ReDim rcTop(1 To mWins.Count) As Integer
    ReDim rcRight(1 To mWins.Count) As Integer
    ReDim rcBottom(1 To mWins.Count) As Integer
    
    Set win = mWins.Item(Index)
    Call GetAbsRect(win, rcWin)
    For i = 1 To mWins.Count
        'box(i).BackColor = UserControl.BackColor
        If Not i = Index Then
            GetAbsRect mWins.Item(i), rcWins(i)
            With mWins.Item(i)
                'If .Right = win.Left Then .Right = win.Right
                'If .Bottom = win.Top Then .Bottom = win.Bottom
                
                If rcWins(i).Left >= rcWin.Left And rcWins(i).Right <= rcWin.Right Then
                    If rcWins(i).Top = rcWin.Bottom Then 'sta sotto
                        nBottom = nBottom + 1
                        rcBottom(nBottom) = i
                        'box(i).BackColor = vbBlue
                    ElseIf rcWins(i).Bottom = rcWin.Top Then 'sta sopra
                        nTop = nTop + 1
                        rcTop(nTop) = i
                        'box(i).BackColor = vbBlue
                    End If
                ElseIf rcWins(i).Top >= rcWin.Top And rcWins(i).Bottom <= rcWin.Bottom Then
                    If rcWins(i).Left = rcWin.Right Then 'sta a destra
                        nRight = nRight + 1
                        rcRight(nRight) = i
                        'box(i).BackColor = vbBlue
                    ElseIf rcWins(i).Right = rcWin.Left Then 'sta a sinistra
                        nLeft = nLeft + 1
                        rcLeft(nLeft) = i
                        'box(i).BackColor = vbBlue
                    End If
                End If
            End With
        End If
    Next i
    
    Debug.Print "left = " & nLeft
    Debug.Print "top = " & nTop
    Debug.Print "right = " & nRight
    Debug.Print "bottom = " & nBottom
    
    'MsgBox "continue"
    
    wWidth = rcWin.Right - rcWin.Left
    wHeight = rcWin.Bottom - rcWin.Top
    If CheckHeight(rcLeft, nLeft, rcWins, wHeight) Then
        'ChangeBoxBkColor rcLeft, nLeft
        For i = 1 To nLeft
            mWins.Item(rcLeft(i)).Right = win.Right
        Next i
    ElseIf CheckHeight(rcRight, nRight, rcWins, wHeight) Then
        'ChangeBoxBkColor rcRight, nRight
        For i = 1 To nRight
            mWins.Item(rcRight(i)).Left = win.Left
        Next i
    ElseIf CheckWidth(rcTop, nTop, rcWins, wWidth) Then
        'ChangeBoxBkColor rcTop, nTop
        For i = 1 To nTop
            mWins.Item(rcTop(i)).Bottom = win.Bottom
        Next i
    ElseIf CheckWidth(rcBottom, nBottom, rcWins, wWidth) Then
        'ChangeBoxBkColor rcBottom, nBottom
        For i = 1 To nBottom
            mWins.Item(rcBottom(i)).Top = win.Top
        Next i
    End If
    Set win = Nothing

    Erase rcWins
    
    'MsgBox "continue"
    
    mWins.Remove Index
    If Not Virtual Then
        For i = Index To box.Count - 2
            box(i).Dock box(i + 1).DockedHWnd
        Next i
        Unload box(box.Count - 1)
        SetBoxes
        mBoxToRemove = 0
    Else
        mBoxToRemove = Index
    End If
End Sub

Private Sub ChangeBoxBkColor(Indexes() As Integer, nIndexes As Integer)
    Dim i As Integer
    
    For i = 1 To nIndexes
        box(Indexes(i)).BackColor = vbRed
    Next i
End Sub

Private Function CheckHeight(Indexes() As Integer, nIndexes As Integer, Rcs() As RECT, Height As Long) As Boolean
    Dim i As Integer
    Dim tot As Long
    
    For i = 1 To nIndexes
        tot = tot + (Rcs(Indexes(i)).Bottom - Rcs(Indexes(i)).Top)
    Next i
    CheckHeight = (tot >= Height)
End Function

Private Function CheckWidth(Indexes() As Integer, nIndexes As Integer, Rcs() As RECT, Width As Long) As Boolean
    Dim i As Integer
    Dim tot As Long
    
    For i = 1 To nIndexes
        tot = tot + (Rcs(Indexes(i)).Right - Rcs(Indexes(i)).Left)
    Next i
    CheckWidth = (tot >= Width)
End Function

Private Sub box_Click(Index As Integer)
    RemoveBox Index
End Sub

Private Sub box_StartingMove(Index As Integer)
    RemoveBox Index, True
End Sub

Private Sub box_WindowClosed(Index As Integer)
    If Not mIgnoreRemove Then RemoveBox Index
End Sub

Private Sub box_WindowUndocked(Index As Integer)
    If Not mIgnoreRemove Then RemoveBox Index
End Sub

Private Sub pctDest_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent MouseWheel(KeyCode)
End Sub

Private Sub pctDest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rc() As RECT, i As Integer
    Dim rcMain As RECT
    
    ReDim rc(1 To mWins.Count) As RECT
    Call GetWindowRect(UserControl.hWnd, rcMain)
    
    For i = 1 To mWins.Count
        Call GetAbsRect(mWins.Item(i), rc(i))
        Call OffsetRect(rc(i), rcMain.Left, rcMain.Top)
        'Call GetWindowRect(box(i).hWnd, rc(i))
        'Call OffsetRect(rc(i), -rc(i).Left, -rc(i).Top)
    Next i
    goshSetDockingRects UserControl.hWnd, rc(1), mWins.Count
End Sub

Private Sub pctDest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DockLoc As DockLocation
    Dim rc As RECT, hWnd As Long, wWidth As Long, wHeight As Long
    Dim Index As Integer, win As cWin, winNew As cWin
    Dim i As Integer
    
    'MsgBox "docking rectangle = " & X & ", docking position = " & Y
        
    
    hWnd = goshGetLastWnd
    GetWindowRect hWnd, rc
    OffsetRect rc, -rc.Left, -rc.Top
    wWidth = rc.Right
    wHeight = rc.Bottom
    
    Index = X
    DockLoc = Y
    If Not DockLoc = dl_none Then
        Set win = mWins.Item(Index)
        Set winNew = New cWin
            winNew.Left = win.Left
            winNew.Top = win.Top
            winNew.Right = win.Right
            winNew.Bottom = win.Bottom
            Select Case DockLoc
                Case dl_left
                    winNew.Right = win.Left + wWidth
                    win.Left = win.Left + wWidth
                Case dl_right
                    winNew.Left = win.Right - wWidth
                    win.Right = win.Right - wWidth
                Case dl_top
                    winNew.Bottom = win.Top + wHeight
                    win.Top = win.Top + wHeight
                Case dl_bottom
                    winNew.Top = win.Bottom - wHeight
                    win.Bottom = win.Bottom - wHeight
            End Select
        mWins.Add winNew
        Set winNew = Nothing
        Set win = Nothing
    End If
    
    If Not mBoxToRemove = 0 Then
        mIgnoreRemove = True
            If DockLoc = dl_none Then
                hWnd = box(mBoxToRemove).DockedHWnd
                box(mBoxToRemove).UnDock SW_SHOW
                goshMoveToLastRect hWnd
            Else
                box(mBoxToRemove).UnDock
            End If
        mIgnoreRemove = False
        
        For i = mBoxToRemove To box.Count - 2
            box(i).Dock box(i + 1).DockedHWnd
        Next i
        Unload box(box.Count - 1)
        mBoxToRemove = False
        'SetBoxes
    End If
    
    SetBoxes
    
    If Not DockLoc = dl_none Then box(box.Count - 1).Dock hWnd
    'Dock hWnd, box.Count - 1
End Sub

Private Sub pctMouseDest_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent MouseWheel(KeyCode)
    KeyCode = 0
End Sub

Private Sub pctMouseDest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MsgBox "click!!"
    If HitTest(X, Y, mDragLoc) Then
        goshStopLastMouse
        mDragging = True
        mDragStartX = cmdDrag.Left
        mDragStartY = cmdDrag.Top
        GetCursorPos mMouseStart
        SetCapture pctMouseDest.hWnd
    End If
End Sub

Private Sub pctMouseDest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Index As Integer, Loc As DockLocation
    Dim pt As POINTAPI, Delta As Integer
    
    If mDragging Then
        GetCursorPos pt
        If mDragLoc = dl_top Or mDragLoc = dl_bottom Then
            Delta = (pt.Y - mMouseStart.Y)
            If Abs(Delta) >= 5 Then cmdDrag.Top = mDragStartY + Delta
        Else
            Delta = (pt.X - mMouseStart.X)
            If Abs(Delta) >= 5 Then cmdDrag.Left = mDragStartX + Delta
        End If
        DoDrag
    Else
        'Debug.Print X, Y
        Index = HitTest(X, Y, Loc)
        If Index = 0 Then cmdDrag.Visible = False
    End If
End Sub

Private Function GetDragRect(ByRef Rcs() As RECT, StartRect As Integer, StartPos As Integer, Loc As DockLocation) As RECT
    Dim i As Integer, Count As Integer
    Dim rcFinal As RECT, w As cWin
    
    Count = UBound(Rcs(), 1)
    rcFinal = Rcs(StartRect)
    Select Case Loc
        Case dl_top, dl_bottom
            For i = 1 To Count
                Set w = mWins.Item(i)
                    If w.Top = StartPos Or w.Bottom = StartPos Then
                        If Rcs(i).Left < rcFinal.Left Then rcFinal.Left = Rcs(i).Left
                        If Rcs(i).Right > rcFinal.Right Then rcFinal.Right = Rcs(i).Right
                    End If
                Set w = Nothing
            Next i
        Case dl_left, dl_right
            For i = 1 To Count
                Set w = mWins.Item(i)
                    If w.Left = StartPos Or w.Right = StartPos Then
                        If Rcs(i).Top < rcFinal.Top Then rcFinal.Top = Rcs(i).Top
                        If Rcs(i).Bottom > rcFinal.Bottom Then rcFinal.Bottom = Rcs(i).Bottom
                    End If
                Set w = Nothing
            Next i
    End Select
    GetDragRect = rcFinal
End Function

Private Sub DoDrag()
    Dim i As Integer, Count As Integer
    Dim w As cWin, StartPos As Integer, NewPos As Integer
    Dim Loc As DockLocation, Delta As Integer
    
    Count = mWins.Count
    'rcFinal = Rcs(StartRect)
    Loc = mDragLoc
    StartPos = mDragStartPos
    If Loc = dl_top Or Loc = dl_bottom Then Delta = mDragStartY - cmdDrag.Top
    If Loc = dl_left Or Loc = dl_right Then Delta = mDragStartX - cmdDrag.Left
    
    If Abs(Delta) <= 5 Then Exit Sub
    
    'If StartPos < 0 Then
    '    NewPos = StartPos - Delta
    'Else
    '    NewPos = StartPos + Delta
    'End If
    NewPos = StartPos - Delta
    'Select Case Loc
    '    Case dl_top, dl_bottom
            For i = 1 To Count
                Set w = mWins.Item(i)
                    If Loc = dl_top Or Loc = dl_bottom Then
                        If w.Top = StartPos Then w.Top = NewPos
                        If w.Bottom = StartPos Then w.Bottom = NewPos
                    End If
                    
                    If Loc = dl_left Or Loc = dl_right Then
                        If w.Left = StartPos Then w.Left = NewPos
                        If w.Right = StartPos Then w.Right = NewPos
                    End If
                Set w = Nothing
            Next i
            mDragStartPos = NewPos
            mDragStartX = cmdDrag.Left
            mDragStartY = cmdDrag.Top
            'mMouseStart = pt
            GetCursorPos mMouseStart
            
    '    Case dl_left, dl_right
    '        For i = 1 To Count
    '            Set w = mWins.Item(i)
    '                If w.Left = StartPos Or w.Right = StartPos Then
    '                    If Rcs(i).Top < rcFinal.Top Then rcFinal.Top = Rcs(i).Top
    '                    If Rcs(i).Bottom > rcFinal.Bottom Then rcFinal.Bottom = Rcs(i).Bottom
    '                End If
    '            Set w = Nothing
    '        Next i
    'End Select
    'GetDragRect = rcFinal
    
    SetBoxes
End Sub

Private Function HitTest(ByVal X As Long, ByVal Y As Long, Optional ByRef Loc As DockLocation) As Integer
    Dim i As Integer, rcWin As RECT
    Dim hCursor As Long, rcDrag As RECT
    'Debug.Print X, Y
    
    Call GetWindowRect(UserControl.hWnd, rcWin)
    X = X - rcWin.Left
    Y = Y - rcWin.Top
    For i = 1 To UBound(mBoxes, 1)
        'If cmdDrag.Visible Then cmdDrag.Visible = False
        If PtInRect(mBoxes(i), X, Y) Then
            'HitTest = i
            'Debug.Print x, y, i
            Loc = dl_none
            With mBoxes(i)
                If Y >= .Top And Y <= .Top + DRAGMARGIN And (mWins.Item(i).Top <> 0) Then 'top!
                    Loc = dl_top
                    mDragStartPos = mWins.Item(i).Top
                ElseIf Y >= .Bottom - DRAGMARGIN And Y <= .Bottom And (mWins.Item(i).Bottom <> 0) Then 'bottom!
                    Loc = dl_bottom
                    mDragStartPos = mWins.Item(i).Bottom
                ElseIf X >= .Left And X <= .Left + DRAGMARGIN And (mWins.Item(i).Left <> 0) Then 'left!
                    Loc = dl_left
                    mDragStartPos = mWins.Item(i).Left
                ElseIf X >= .Right - DRAGMARGIN And X <= .Right And (mWins.Item(i).Right <> 0) Then 'right!
                    Loc = dl_right
                    mDragStartPos = mWins.Item(i).Right
                End If
                
                If Loc = dl_none Then Exit For
                
                HitTest = i
                rcDrag = GetDragRect(mBoxes(), i, mDragStartPos, Loc)
                'cmdDrag.Left = rcDrag.Left
                'cmdDrag.Top = rcDrag.Top
                'cmdDrag.Width = rcDrag.Right - rcDrag.Left
                'cmdDrag.Height = rcDrag.Bottom - rcDrag.Top
                'cmdDrag.Visible = True
                'Exit For
                With rcDrag
                    Select Case Loc
                        Case dl_top, dl_bottom
                            cmdDrag.Left = .Left
                            If Loc = dl_top Then
                                cmdDrag.Top = .Top - DRAGMARGIN
                            Else
                                cmdDrag.Top = .Bottom - DRAGMARGIN
                            End If
                            cmdDrag.Width = .Right - .Left
                            cmdDrag.Height = DRAGMARGIN * 2
                            cmdDrag.MousePointer = vbSizeNS
                        Case dl_left, dl_right
                            If Loc = dl_left Then
                                cmdDrag.Left = .Left - DRAGMARGIN
                            Else
                                cmdDrag.Left = .Right - DRAGMARGIN
                            End If
                            cmdDrag.Top = .Top
                            cmdDrag.Width = DRAGMARGIN * 2
                            cmdDrag.Height = .Bottom - .Top
                            cmdDrag.MousePointer = vbSizeWE
                    End Select
                End With
                cmdDrag.Visible = True
            End With
            Exit For
        End If
    Next i

End Function

Private Sub pctMouseDest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mDragging Then
        mDragging = False
        DoDrag
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim win As cWin
    
    Set mWins = New Collection
    
    Set win = New cWin
    win.Left = 0
    win.Top = 0
    win.Right = 0
    win.Bottom = 0
    mWins.Add win
    Set win = Nothing

    'Set win = New cWin
    'win.Left = 0
    'win.Top = 0
    'win.Right = -200
    'win.Bottom = 100
    'mWins.Add win
    'Set win = Nothing

    'Set win = New cWin
    'win.Left = 0
    'win.Top = -100
    'win.Right = -200
    'win.Bottom = 0
    'mWins.Add win
    'Set win = Nothing

    'Set win = New cWin
    'win.Left = -200
    'win.Top = 0
    'win.Right = -100
    'win.Bottom = 0
    'mWins.Add win
    'Set win = Nothing

    'Set win = New cWin
    'win.Left = -100
    'win.Top = 0
    'win.Right = 0
    'win.Bottom = 0
    'mWins.Add win
    'Set win = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        goshStartHook UserControl.Parent.hWnd, pctDest.hWnd, pctMouseDest.hWnd, App.ThreadID
    End If
End Sub

Private Sub UserControl_Resize()
    If Ambient.UserMode Then
        Call SetBoxes
    End If
End Sub

Private Sub UserControl_Terminate()
    Dim i As Integer
    
    Set mWins = Nothing
    goshStopHook

    mIgnoreRemove = True 'this prevents boxes from being removed when each window is closed
    For i = 1 To box.Count - 1
        box(i).CloseForm
    Next i
    mIgnoreRemove = False

    Erase mBoxes
End Sub

Private Sub GetAbsRect(ByRef win As cWin, ByRef rcDest As RECT)
    Dim pWidth As Long, pheight As Long
    
    pWidth = UserControl.ScaleWidth
    pheight = UserControl.ScaleHeight
    With rcDest
        .Left = IIf(win.Left >= 0, win.Left, pWidth + win.Left)
        .Top = IIf(win.Top >= 0, win.Top, pheight + win.Top)
        .Right = IIf(win.Right <= 0, pWidth + win.Right, win.Right)
        .Bottom = IIf(win.Bottom <= 0, pheight + win.Bottom, win.Bottom)
    End With
End Sub

Private Sub SetBoxes()
    Dim win As cWin, i As Integer
    'p... variables are relative to the parent window
    Dim pWidth As Long, pheight As Long
    Dim rc As RECT
    
    pWidth = UserControl.ScaleWidth
    pheight = UserControl.ScaleHeight
    
    Erase mBoxes
    ReDim mBoxes(1 To mWins.Count) As RECT
    For i = 1 To mWins.Count
        Set win = mWins.Item(i)
        On Error Resume Next
        Load box(i)
        
        With box(i)
            .Visible = True
            GetAbsRect win, rc
            '.Left = IIf(win.Left >= 0, win.Left, pwidth + win.Left)
            '.Top = IIf(win.Top >= 0, win.Top, pheight + win.Top)
            '.Width = IIf(win.Right <= 0, pwidth + win.Right, win.Right) - .Left
            '.Height = IIf(win.Bottom <= 0, pheight + win.Bottom, win.Bottom) - .Top
            .Left = rc.Left
            .Top = rc.Top
            .Width = rc.Right - rc.Left
            .Height = rc.Bottom - rc.Top
            
            mBoxes(i) = rc
        End With
        On Error GoTo 0
        Set win = Nothing
    Next i
End Sub
