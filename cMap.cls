VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mRoom() As cRoom
Private mLink() As cLink
Private mActive As Integer
Private mSelected As Integer

Public Property Get Selected() As Integer
    Selected = mSelected
End Property

Public Property Let Selected(data As Integer)
    mSelected = data
End Property

Public Property Get Count() As Integer
    Count = UBound(mRoom, 1)
End Property

Public Property Get LinkCount() As Integer
    LinkCount = UBound(mLink, 1)
End Property

Public Property Get Room(Index As Integer) As cRoom
    Set Room = mRoom(Index)
End Property

Public Property Get Link(Index As Integer) As cLink
    Set Link = mLink(Index)
End Property

Public Property Get ActiveRoom() As Integer
    ActiveRoom = mActive
End Property

Public Property Let ActiveRoom(data As Integer)
    mActive = data
End Property

Public Property Get CurRealX() As Integer
    CurRealX = mRoom(mActive).RealX
End Property

Public Property Get CurRealY() As Integer
    CurRealY = mRoom(mActive).RealY
End Property

Public Property Get CurX() As Integer
    CurX = mRoom(mActive).PosX
End Property

Public Property Get CurY() As Integer
    CurY = mRoom(mActive).PosY
End Property

Public Property Get CurZ() As Integer
    CurZ = mRoom(mActive).PosZ
End Property

Private Sub ShiftY(Inizio As Integer, Shift As Integer)
    Dim i As Integer

    If Shift = -1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosY <= Inizio Then .PosY = .PosY - 1
            End With
        Next i
        
        For i = 1 To UBound(mLink, 1)
            If mLink(i).y1 <= Inizio Then mLink(i).y1 = mLink(i).y1 - 1
            If mLink(i).y2 <= Inizio Then mLink(i).y2 = mLink(i).y2 - 1
        Next i
    ElseIf Shift = 1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosY >= Inizio Then .PosY = .PosY + 1
            End With
        Next i
        
        For i = 1 To UBound(mLink, 1)
            If mLink(i).y1 >= Inizio Then mLink(i).y1 = mLink(i).y1 + 1
            If mLink(i).y2 >= Inizio Then mLink(i).y2 = mLink(i).y2 + 1
        Next i
    End If
End Sub

Private Sub ShiftY2(Inizio As Integer, Shift As Integer)
    Dim i As Integer

    If Shift < 0 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .RealY <= Inizio Then .RealY = .RealY + Shift
            End With
        Next i
    ElseIf Shift > 0 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .RealY >= Inizio Then .RealY = .RealY + Shift
            End With
        Next i
    End If
End Sub

Private Sub ShiftX2(Inizio As Integer, Shift As Integer)
    Dim i As Integer

    If Shift < 0 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .RealX <= Inizio Then .RealX = .RealX + Shift
            End With
        Next i
    ElseIf Shift > 0 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .RealX >= Inizio Then .RealX = .RealX + Shift
            End With
        Next i
    End If
End Sub

Private Sub ShiftX(Inizio As Integer, Shift As Integer)
    Dim i As Integer

    If Shift = -1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosX <= Inizio Then .PosX = .PosX - 1
            End With
        Next i
        
        For i = 1 To UBound(mLink, 1)
            If mLink(i).x1 <= Inizio Then mLink(i).x1 = mLink(i).x1 - 1
            If mLink(i).x2 <= Inizio Then mLink(i).x2 = mLink(i).x2 - 1
        Next i
    ElseIf Shift = 1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosX >= Inizio Then .PosX = .PosX + 1
            End With
        Next i
        
        For i = 1 To UBound(mLink, 1)
            If mLink(i).x1 >= Inizio Then mLink(i).x1 = mLink(i).x1 + 1
            If mLink(i).x2 >= Inizio Then mLink(i).x2 = mLink(i).x2 + 1
        Next i
    End If
End Sub

Private Sub ShiftZ2(Inizio As Integer, Shift As Integer)
    Dim i As Integer

    If Shift = -1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosZ <= Inizio Then .PosZ = .PosZ - 1
            End With
        Next i
    ElseIf Shift = 1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosZ >= Inizio Then .PosZ = .PosZ + 1
            End With
        Next i
    End If
End Sub

Private Sub ShiftZ(Inizio As Integer, Shift As Integer)
    Dim i As Integer

    If Shift = -1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosZ <= Inizio Then .PosZ = .PosZ - 1
            End With
        Next i
        
        For i = 1 To UBound(mLink, 1)
            If mLink(i).z1 <= Inizio Then mLink(i).z1 = mLink(i).z1 - 1
            If mLink(i).z2 <= Inizio Then mLink(i).z2 = mLink(i).z2 - 1
        Next i
    ElseIf Shift = 1 Then
        For i = 1 To UBound(mRoom, 1)
            With mRoom(i)
                If .PosZ >= Inizio Then .PosZ = .PosZ + 1
            End With
        Next i
        
        For i = 1 To UBound(mLink, 1)
            If mLink(i).z1 >= Inizio Then mLink(i).z1 = mLink(i).z1 + 1
            If mLink(i).z2 >= Inizio Then mLink(i).z2 = mLink(i).z2 + 1
        Next i
    End If
End Sub

Public Sub Add2(Caption As String, Optional Mov As Mapper_Mov = -1, Optional RealX As Integer, Optional RealY As Integer, Optional PosZ As Integer, Optional AutoLink As Boolean = True)
    'se mov e' specificato, ovvero diverso da -1, ci pensa la procedura stessa
    'a determinare le coordinate della nuova stanza usando quella corrente e a
    'fare poi tutti i calcoli; se mov = -1 la procedura utilizza realx e realy
    'che devono essere assolutamente specificati
    
    Dim Count As Integer, Index As Integer
    Dim Crea As Boolean
    
    If Mov <> -1 Then
        RealX = mRoom(mActive).RealX
        RealY = mRoom(mActive).RealY
        PosZ = mRoom(mActive).PosZ
        Select Case Mov
            Case nord
                RealY = RealY - ROOM_DIM
            Case sud
                RealY = RealY + ROOM_DIM
            Case est
                RealX = RealX + ROOM_DIM
            Case ovest
                RealX = RealX - ROOM_DIM
            Case alto
                PosZ = PosZ + 1
            Case basso
                PosZ = PosZ - 1
        End Select
    End If
    
    Count = UBound(mRoom, 1) + 1
    Index = Search2(RealX, RealY, PosZ)
    
    Crea = True
    If Index <> 0 Then
        If Caption = mRoom(Index).Caption Then
            Crea = False
        Else
            If RealY <> Me.CurRealY Then ShiftY2 RealY, (RealY - Me.CurRealY) ' / Abs((RealY - Me.CurRealY))
            If RealX <> Me.CurRealX Then ShiftX2 RealX, (RealX - Me.CurRealX) ' / Abs((RealX - Me.CurRealX))
            If PosZ <> Me.CurZ Then ShiftZ2 PosZ, (PosZ - Me.CurZ) ' / Abs((PosZ - Me.CurZ))
        End If
    End If
    
    If Crea Then
        If Count = 1 Then
            ReDim mRoom(1 To 1) As cRoom
        Else
            ReDim Preserve mRoom(1 To Count) As cRoom
        End If
        Set mRoom(Count) = New cRoom
        Index = Count
        'mRoom(Index).PosX = x
        'mRoom(Index).PosY = y
        mRoom(Index).PosZ = PosZ
        'If RealX = -1 Then
        '    mRoom(Index).RealX = x * ROOM_DIM + ROOM_MARG
        '    mRoom(Index).RealY = y * ROOM_DIM + ROOM_MARG
        'Else
            mRoom(Index).RealX = RealX
            mRoom(Index).RealY = RealY
        'End If
        mRoom(Index).Caption = Caption
    End If
    
    If Count <> 1 And AutoLink Then
        AddLink2 mActive, Index, Mov
    End If
    
    mActive = Index
End Sub

Public Sub Add(Caption As String, X As Integer, Y As Integer, z As Integer, Optional AutoLink As Boolean = True, Optional RealX As Integer = -1, Optional RealY As Integer = -1)
    Dim Count As Integer, Index As Integer
    Dim Crea As Boolean
    
    Index = Search(X, Y, z)
    Count = UBound(mRoom, 1) + 1
    If Index = 0 Then
        Crea = True
    Else
        If mRoom(Index).Caption = "" Then
            mRoom(Index).Caption = Caption
        ElseIf mRoom(Index).Caption <> Caption Then
            If Y <> Me.CurY Then ShiftY Y, (Y - Me.CurY) / Abs((Y - Me.CurY))
            If X <> Me.CurX Then ShiftX X, (X - Me.CurX) / Abs((X - Me.CurX))
            If z <> Me.CurZ Then ShiftY z, (z - Me.CurZ) / Abs((z - Me.CurZ))
            Crea = True
        End If
    End If
    
    If Crea Then
        If Count = 1 Then
            ReDim mRoom(1 To 1) As cRoom
        Else
            ReDim Preserve mRoom(1 To Count) As cRoom
        End If
        Set mRoom(Count) = New cRoom
        Index = Count
        mRoom(Index).PosX = X
        mRoom(Index).PosY = Y
        mRoom(Index).PosZ = z
        If RealX = -1 Then
            mRoom(Index).RealX = X * ROOM_DIM + ROOM_MARG
            mRoom(Index).RealY = Y * ROOM_DIM + ROOM_MARG
        Else
            mRoom(Index).RealX = RealX
            mRoom(Index).RealY = RealY
        End If
        mRoom(Index).Caption = Caption
        'Debug.Print Time & " creata"
    Else
        Debug.Print Time & " non-creata"
    End If
    
    If Not Count = 1 And AutoLink Then
        AddLink mRoom(mActive).PosX, mRoom(mActive).PosY, mRoom(mActive).PosZ, _
                mRoom(Index).PosX, mRoom(Index).PosY, mRoom(Index).PosZ
    End If
    mActive = Index
End Sub

Public Sub AddLink2(Room1 As Integer, Room2 As Integer, Optional Mov As Mapper_Mov = -1)
    Dim Count As Integer, Index As Integer

    Count = UBound(mRoom, 1)
    If Room1 > Count Or Room2 > Count Then Exit Sub
    
    Count = UBound(mLink, 1) + 1
    If Count = 1 Then
        ReDim mLink(1 To 1) As cLink
    Else
        ReDim Preserve mLink(1 To Count) As cLink
    End If
    Set mLink(Count) = New cLink
    Index = Count
    mLink(Index).Room1 = Room1
    mLink(Index).Room2 = Room2
    mLink(Index).Mov = Mov
    'If Mov = -1 Then
    '    If mRoom(Room1).PosX < mRoom(Room2).PosX Then
    '        Mov = est
    '        Debug.Print " __mov = est"
    '    ElseIf mRoom(Room1).PosX > mRoom(Room2).PosX Then
    '        Mov = ovest
    '        Debug.Print " __mov = ovest"
    '    ElseIf mRoom(Room1).PosY < mRoom(Room2).PosY Then
    '        Mov = sud
    '        Debug.Print " __mov = sud"
    '    ElseIf mRoom(Room1).PosY > mRoom(Room2).PosY Then
    '        Mov = nord
    '        Debug.Print " __mov = nord"
    '    ElseIf mRoom(Room1).PosZ < mRoom(Room2).PosZ Then
    '        Mov = alto
    '        Debug.Print " __mov = alto"
    '    ElseIf mRoom(Room1).PosZ > mRoom(Room2).PosZ Then
    '        Mov = basso
    '        Debug.Print " __mov = basso"
    '    End If
    'End If
    If mRoom(Room1).PosZ <> mRoom(Room2).PosZ Then
        mRoom(Room1).ChangeZ = True
        mRoom(Room2).ChangeZ = True
    End If
End Sub

Public Sub AddLink(x1 As Integer, y1 As Integer, z1 As Integer, x2 As Integer, y2 As Integer, z2 As Integer)
    Dim Count As Integer, Index As Integer

    Index = SearchLink(x1, y1, z1, x2, y2, z2)
    If Index = 0 Then
        Count = UBound(mLink, 1) + 1
        If Count = 1 Then
            ReDim mLink(1 To 1) As cLink
        Else
            ReDim Preserve mLink(1 To Count) As cLink
        End If
        Set mLink(Count) = New cLink
        Index = Count
        mLink(Index).x1 = x1
        mLink(Index).y1 = y1
        mLink(Index).z1 = z1
        mLink(Index).x2 = x2
        mLink(Index).y2 = y2
        mLink(Index).z2 = z2
        Debug.Print Time & " link_creato"
    Else
        Debug.Print Time & " link_non_creato"
    End If
    'mActive = Index
End Sub

Public Function Jump(rIndex As Integer)
    mActive = rIndex
End Function

Public Function RoomSearch2(Sorg As Integer, Mov As Mapper_Mov) As Integer
    Dim i As Integer, Inverso As Mapper_Mov

    Select Case Mov
        Case nord
            Inverso = sud
        Case sud
            Inverso = nord
        Case est
            Inverso = ovest
        Case ovest
            Inverso = est
        Case alto
            Inverso = basso
        Case basso
            Inverso = alto
    End Select
    
    For i = 1 To UBound(mLink, 1)
        With Link(i)
            If (.Room1 = Sorg And .Mov = Mov) Then
                RoomSearch2 = .Room2
            ElseIf (.Room2 = Sorg And .Mov = Inverso) Then
                RoomSearch2 = .Room1
            End If
        End With
    Next i
End Function

Public Function RoomSearch(sx As Integer, sy As Integer, sz As Integer, cx As Integer, cy As Integer, cz As Integer) As Integer
    'quella s sta per source, la c come di consueto per distanza
    
    Dim i As Integer
    
    If cx = -1 Then
        For i = 1 To UBound(mLink, 1)
            With mLink(i)
                If ((.x1 = sx And .y1 = sy And .z1 = sz) And .x2 <= sx - 1) Then
                    RoomSearch = Search(.x2, .y2, .z2)
                    Exit For
                ElseIf ((.x2 = sx And .y2 = sy And .z2 = sz) And .x1 <= sx - 1) Then
                    RoomSearch = Search(.x1, .y1, .z1)
                    Exit For
                End If
            End With
        Next i
    ElseIf cx = 1 Then
        For i = 1 To UBound(mLink, 1)
            With mLink(i)
                If ((.x1 = sx And .y1 = sy And .z1 = sz) And .x2 >= sx + 1) Then
                    RoomSearch = Search(.x2, .y2, .z2)
                    Exit For
                ElseIf ((.x2 = sx And .y2 = sy And .z2 = sz) And .x1 >= sx + 1) Then
                    RoomSearch = Search(.x1, .y1, .z1)
                    Exit For
                End If
            End With
        Next i
    ElseIf cy = -1 Then
        For i = 1 To UBound(mLink, 1)
            With mLink(i)
                If ((.x1 = sx And .y1 = sy And .z1 = sz) And .y2 <= sy - 1) Then
                    RoomSearch = Search(.x2, .y2, .z2)
                    Exit For
                ElseIf ((.x2 = sx And .y2 = sy And .z2 = sz) And .y1 <= sy - 1) Then
                    RoomSearch = Search(.x1, .y1, .z1)
                    Exit For
                End If
            End With
        Next i
    ElseIf cy = 1 Then
        For i = 1 To UBound(mLink, 1)
            With mLink(i)
                If ((.x1 = sx And .y1 = sy And .z1 = sz) And .y2 >= sy + 1) Then
                    RoomSearch = Search(.x2, .y2, .z2)
                    Exit For
                ElseIf ((.x2 = sx And .y2 = sy And .z2 = sz) And .y1 >= sy + 1) Then
                    RoomSearch = Search(.x1, .y1, .z1)
                    Exit For
                End If
            End With
        Next i
    ElseIf cz = -1 Then
        For i = 1 To UBound(mLink, 1)
            With mLink(i)
                If ((.x1 = sx And .y1 = sy And .z1 = sz) And .z2 <= sz - 1) Then
                    RoomSearch = Search(.x2, .y2, .z2)
                    Exit For
                ElseIf ((.x2 = sx And .y2 = sy And .z2 = sz) And .z1 <= sz - 1) Then
                    RoomSearch = Search(.x1, .y1, .z1)
                    Exit For
                End If
            End With
        Next i
    ElseIf cz = 1 Then
        For i = 1 To UBound(mLink, 1)
            With mLink(i)
                If ((.x1 = sx And .y1 = sy And .z1 = sz) And .z2 >= sz + 1) Then
                    RoomSearch = Search(.x2, .y2, .z2)
                    Exit For
                ElseIf ((.x2 = sx And .y2 = sy And .z2 = sz) And .z1 >= sz + 1) Then
                    RoomSearch = Search(.x1, .y1, .z1)
                    Exit For
                End If
            End With
        Next i
    End If
End Function

Public Function DeleteLink(Index As Integer)
    Dim Count As Integer, i As Integer
    'Dim Link As cLink
    
    Count = UBound(mLink, 1)
    If Count = 1 Then
       ReDim mLink(0) As cLink
    Else
        If Not Index = Count Then
            Set mLink(Index) = Nothing
            For i = Index To Count - 1
                'Set Link = mLink(i + 1)
                Set mLink(i) = mLink(i + 1)
                Set mLink(i + 1) = Nothing
                'Set Link = Nothing
            Next i
        End If
        Set mLink(Count) = Nothing
        ReDim Preserve mLink(1 To Count - 1) As cLink
    End If
End Function

Public Function DeleteRoom(Index As Integer)
    Dim Count As Integer, i As Integer
    'Dim Room As cRoom

    Count = UBound(mRoom, 1)
    If Count = 1 Then
       ReDim mRoom(0) As cRoom
    Else
        If Not Index = Count Then
            Set mRoom(Index) = Nothing
            For i = Index To Count - 1
                'Set Room = mRoom(i + 1)
                'Set mRoom(i) = Nothing
                Set mRoom(i) = mRoom(i + 1)
                Set mRoom(i + 1) = Nothing
                'Set Room = Nothing
            Next i
            
            For i = 1 To UBound(mLink, 1)
                If mLink(i).Room1 > Index Then mLink(i).Room1 = mLink(i).Room1 - 1
                If mLink(i).Room2 > Index Then mLink(i).Room2 = mLink(i).Room2 - 1
            Next i
        End If
        Set mRoom(Count) = Nothing
        ReDim Preserve mRoom(1 To Count - 1) As cRoom
        If mActive >= Index Then mActive = mActive - 1
    End If
End Function

Public Function SearchLink2(Room As Integer) As Integer
    Dim i As Integer

    For i = 1 To UBound(mLink, 1)
        If mLink(i).Room1 = Room Or mLink(i).Room2 = Room Then
            SearchLink2 = i
            Exit For
        End If
    Next i
End Function

Private Function SearchLink(x1 As Integer, y1 As Integer, z1 As Integer, x2 As Integer, y2 As Integer, z2 As Integer) As Integer
    Dim i As Integer

    For i = 1 To UBound(mLink, 1)
        If mLink(i).x1 = x1 And mLink(i).y1 = y1 And mLink(i).z1 = z1 And mLink(i).x2 = x2 And mLink(i).y2 = y2 And mLink(i).z2 = z2 Then
            SearchLink = i
            Exit For
        End If
    Next i
End Function

Public Function Search2(X As Integer, Y As Integer, z As Integer, Optional ByRef Tag As Boolean) As Integer
    Dim i As Integer

    Tag = False
    For i = UBound(mRoom, 1) To 1 Step -1
        If X >= mRoom(i).RealX - ROOM_MARG And X <= mRoom(i).RealX + ROOM_DIM - ROOM_MARG And _
           Y >= mRoom(i).RealY - ROOM_MARG And Y <= mRoom(i).RealY + ROOM_DIM - ROOM_MARG And _
           mRoom(i).PosZ = z Then
            Search2 = i
            Tag = False
            Exit For
        ElseIf (X >= mRoom(i).TagX And X <= mRoom(i).TagX + mRoom(i).TagW And _
           Y >= mRoom(i).TagY And Y <= mRoom(i).TagY + mRoom(i).TagH) And mRoom(i).PosZ = z Then
            Search2 = i
            Tag = True
            Exit For
        End If
    Next i
End Function

Public Function Search(X As Integer, Y As Integer, z As Integer) As Integer
    Dim i As Integer

    For i = 1 To UBound(mRoom, 1)
        If mRoom(i).PosX = X And mRoom(i).PosY = Y And mRoom(i).PosZ = z Then
            Search = i
            Exit For
        End If
    Next i
End Function

Private Sub Class_Initialize()
    ReDim mRoom(0) As cRoom
    ReDim mLink(0) As cLink
End Sub

Public Sub Clear()
    Dim i As Integer

    For i = 1 To UBound(mRoom, 1)
        Set mRoom(i) = Nothing
    Next i
    Erase mRoom

    For i = 1 To UBound(mLink, 1)
        Set mLink(i) = Nothing
    Next i
    Erase mLink
    
    ReDim mRoom(0) As cRoom
    ReDim mLink(0) As cLink
End Sub

Private Sub Class_Terminate()
    Clear
End Sub
