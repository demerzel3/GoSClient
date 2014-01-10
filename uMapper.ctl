VERSION 5.00
Begin VB.UserControl uMapper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   Begin VB.PictureBox pctMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   75
      MouseIcon       =   "uMapper.ctx":0000
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   75
      Width           =   6915
      Begin VB.PictureBox pctRosa 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2250
         Left            =   2775
         ScaleHeight     =   150
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   150
         TabIndex        =   1
         Top             =   525
         Width           =   2250
         Begin VB.Image imgNBasso 
            Height          =   465
            Left            =   1725
            Top             =   1800
            Width           =   465
         End
         Begin VB.Image imgNAlto 
            Height          =   465
            Left            =   1725
            Top             =   0
            Width           =   465
         End
         Begin VB.Image imgNEst 
            Height          =   465
            Left            =   1350
            Top             =   900
            Width           =   915
         End
         Begin VB.Image imgNSud 
            Height          =   915
            Left            =   825
            Top             =   1350
            Width           =   615
         End
         Begin VB.Image imgNOvest 
            Height          =   465
            Left            =   0
            Top             =   900
            Width           =   915
         End
         Begin VB.Image imgNNord 
            Height          =   915
            Left            =   825
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin VB.Image imgPiccola 
      Height          =   750
      Left            =   6450
      Picture         =   "uMapper.ctx":0152
      Top             =   5025
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgPCustom 
      Height          =   255
      Left            =   1800
      Top             =   6075
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgPDown 
      Height          =   255
      Left            =   1500
      Picture         =   "uMapper.ctx":04F2
      Top             =   6075
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgPUp 
      Height          =   255
      Left            =   1125
      Picture         =   "uMapper.ctx":0AD4
      Top             =   6075
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBasso 
      Height          =   2250
      Left            =   4200
      Picture         =   "uMapper.ctx":10B6
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgAlto 
      Height          =   2250
      Left            =   3525
      Picture         =   "uMapper.ctx":1D22
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgEst 
      Height          =   2250
      Left            =   2850
      Picture         =   "uMapper.ctx":2994
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgOvest 
      Height          =   2250
      Left            =   2175
      Picture         =   "uMapper.ctx":3564
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgSud 
      Height          =   2250
      Left            =   1500
      Picture         =   "uMapper.ctx":414F
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgNord 
      Height          =   2250
      Left            =   825
      Picture         =   "uMapper.ctx":4D0C
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgVenti 
      Height          =   2250
      Left            =   75
      Picture         =   "uMapper.ctx":58D6
      Top             =   3525
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgPSel 
      Height          =   255
      Left            =   825
      Picture         =   "uMapper.ctx":6618
      Top             =   6075
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgPMouse 
      Height          =   255
      Left            =   450
      Picture         =   "uMapper.ctx":69CE
      Top             =   6075
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgP 
      Height          =   255
      Left            =   75
      Picture         =   "uMapper.ctx":6D84
      Top             =   6075
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuRoom 
      Caption         =   "Room"
      Begin VB.Menu mnuCaption 
         Caption         =   "Modifica titolo"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Elimina stanza"
      End
      Begin VB.Menu mnuTag 
         Caption         =   "Inserisci commento"
      End
      Begin VB.Menu mnuChangeActive 
         Caption         =   "Posiziona personaggio"
      End
   End
   Begin VB.Menu mnuImages 
      Caption         =   "Images"
      Begin VB.Menu mnuImage 
         Caption         =   "Image"
         Index           =   0
      End
   End
End
Attribute VB_Name = "uMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MARGINE As Integer = 2

Private mOrigineX As Long
Private mOrigineY As Long
Private mMap As cMap
Private mSetOrigin As Boolean

Private mDragStartX As Long
Private mDragStartY As Long
Private mDragged As Boolean
Private mTag As Boolean

Private mZCoord As Integer

'Private WithEvents mnuRoommm As cMenu
'Private WithEvents mnuImagesss As cMenu
'Private WithEvents mnuMap As cMenu
Private mSel As Integer

Private mMapMode As Integer

Private mFollow As Boolean
Private mMod As Boolean

Public Event RoomTitle(Title As String)
Public Event Send(Stringa As String)
Public Event MouseMove()
Public Event ContextMenu()

Public Property Get Follow() As Boolean
    Follow = mFollow
End Property

Public Property Let Follow(data As Boolean)
    mFollow = data
    If Ambient.UserMode Then Draw
End Property

Public Property Get Modified() As Boolean
    Modified = mMod
End Property

Public Property Let Modified(data As Boolean)
    mMod = data
End Property

Public Sub DrawBorder()
    Dim Bordi As cGrafica
    
    Set Bordi = New cGrafica
    Bordi.DisegnaBordi UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1, 2, 0, 0, 30
    UserControl.Refresh
    Set Bordi = Nothing
End Sub

Public Property Get MapMode() As Integer
    MapMode = mMapMode
End Property

Public Property Let MapMode(data As Integer)
    mMapMode = data
    If mMapMode = MAPMODE_PAUSE Then
        pctRosa.Enabled = False
    Else
        pctRosa.Enabled = True
    End If
End Property

Public Function SaveMap(Optional Nomefile As String = "") As String
    'Dim NomeFile As String
    Dim Save As cBinary

    If Nomefile = "" Then Nomefile = InputBox("Inserisci un nome per la mappa")
    If Nomefile <> "" Then
        Nomefile = IIf(LCase$(Right$(Nomefile, 4)) = ".map", Nomefile, Nomefile & ".map")
        Set Save = New cBinary
        Save.SaveMap Nomefile, mMap
        Set Save = Nothing
        'MsgBox "Salvataggio completato"
    End If
    
    SaveMap = Nomefile
    mMod = False
End Function

Public Function LoadMap() As String
    Dim Nomefile As String
    Dim Load As cBinary
    Dim map As cMap
    Dim Elenco As frmList, cElenco As Collection
    Dim Path As String, i As Integer

    'Path = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "Maps\"
    Path = gMudPath & "maps\"
    
    On Error Resume Next
    MkDir Path
    
    Set cElenco = New Collection
        Nomefile = Dir$(Path)
        Do Until Nomefile = ""
            Debug.Print Nomefile
            If LCase$(Right$(Nomefile, 4)) = ".map" Then
                'Elenco.AddItem Left$(Nomefile, Len(Nomefile) - 4), Nomefile
                cElenco.Add Nomefile
            End If
            Nomefile = Dir$
        Loop
    Set Elenco = New frmList
        Elenco.Caption = "Carica mappa"
        For i = 1 To cElenco.Count
           If LCase$(Right$(cElenco(i), 4)) = ".map" Then
                Elenco.AddItem Left$(cElenco(i), Len(cElenco(i)) - 4), cElenco(i)
            End If
        Next i
        Elenco.ShowForm Nomefile
    Set Elenco = Nothing
    Set cElenco = Nothing
    
    'NomeFile = InputBox("Inserisci il nome della mappa")
    If Nomefile <> "" Then
        Nomefile = IIf(Right$(Nomefile, 4) = ".map", Nomefile, Nomefile & ".map")
        Set Load = New cBinary
        Set map = Load.LoadMap(Nomefile)
        Set Load = Nothing
        
        If map Is Nothing Then
            MsgBox "Impossibile caricare la mappa selezionata"
        Else
            Set mMap = Nothing
            Set mMap = map
            'Set map = Nothing
            mMap.ActiveRoom = 1
            'MsgBox "Caricamento completato"
            mOrigineX = pctMap.ScaleWidth / 2
            mOrigineY = pctMap.ScaleHeight / 2
            Draw
            LoadMap = Nomefile
        End If
    End If
    
    mMod = False
End Function

Public Sub SetOrg()
    Dim win As rect

    GetWindowRect pctMap.hwnd, win
    
    mDragStartX = mOrigineX
    mDragStartY = mOrigineY
    pctMap.MousePointer = 99
    SetCursorPos win.Left + pctMap.ScaleWidth / 2, win.Top + pctMap.ScaleHeight / 2
    
    Debug.Print "!!!!!!!!!!getcapture = " & GetCapture
    SetCapture pctMap.hwnd
    Debug.Print "!!!!!!!!!!getcapture = " & GetCapture
    mZCoord = 0
    
    mSetOrigin = True

    mSel = 0
End Sub

Public Sub SelectRoom(i As Integer)
    Dim sel As Integer

    'If Not i = 0 Then
        sel = mMap.Selected
        mMap.Selected = i
        If Not sel = 0 And sel <= mMap.Count Then DrawRoom (sel)
        If Not mMap.Selected = 0 Then DrawRoom (mMap.Selected)
        'If mMap.Room(i).PosZ = mZCoord Then
        '    'If Not mMap.ActiveRoom = i Then
        '        'pctMap.PaintPicture imgP.Picture, mOrigineX + (mMap.Room(i).PosX * ROOM_DIM) + ROOM_MARG, mOrigineY + (mMap.Room(i).PosY * ROOM_DIM) + ROOM_MARG
        '    'Else
        '        pctMap.PaintPicture imgPMouse.Picture, mOrigineX + (mMap.Room(i).PosX * ROOM_DIM) + ROOM_MARG, mOrigineY + (mMap.Room(i).PosY * ROOM_DIM) + ROOM_MARG
        '    'End If
        'End If
    'End If
End Sub

Private Sub DrawRoom(Index As Integer)
    Dim CoordX As Long, CoordY As Long, img As Image
    Dim Area As rect, hBrush As Long

    'CoordX = mOrigineX + (mMap.Room(Index).PosX * ROOM_DIM) + ROOM_MARG
    'CoordY = mOrigineY + (mMap.Room(Index).PosY * ROOM_DIM) + ROOM_MARG
    CoordX = mMap.Room(Index).RealX + mOrigineX
    CoordY = mMap.Room(Index).RealY + mOrigineY
    If mMap.Room(Index).PosZ = mZCoord Then
        If mMap.Selected = Index Then
            Set img = imgPMouse
        ElseIf mMap.ActiveRoom = Index Then
            Set img = imgPSel
        Else
            If mMap.Room(Index).Image = "" Then
                Set img = imgP
            Else
                imgPCustom.Picture = LoadPicture(App.Path & "\mapimages\" & mMap.Room(Index).Image)
                Set img = imgPCustom
            End If
        End If
        
        pctMap.PaintPicture img.Picture, CoordX, CoordY
        Set img = Nothing
        
        If mMap.Room(Index).ChangeZ Then
            'If mMap.RoomSearch(mMap.Room(Index).PosX, mMap.Room(Index).PosY, mZCoord, 0, 0, 1) <> 0 Then
            If mMap.RoomSearch2(Index, alto) <> 0 Then
                pctMap.PaintPicture imgPUp.Picture, CoordX, CoordY
            End If
            
            'If mMap.RoomSearch(mMap.Room(Index).PosX, mMap.Room(Index).PosY, mZCoord, 0, 0, -1) <> 0 Then
            If mMap.RoomSearch2(Index, basso) <> 0 Then
                pctMap.PaintPicture imgPDown.Picture, CoordX, CoordY
            End If
        End If
        
        'If mMap.RoomSearch2(Index, nord) Then pctMap.Circle (CoordX + (ROOM_DIM - ROOM_MARG * 2) \ 2, CoordY), 2
        'If mMap.RoomSearch2(Index, sud) Then pctMap.Circle (CoordX + (ROOM_DIM - ROOM_MARG * 2) \ 2, CoordY + (ROOM_DIM - ROOM_MARG * 2) - 1), 2
        'If mMap.RoomSearch2(Index, ovest) Then pctMap.Circle (CoordX, CoordY + (ROOM_DIM - ROOM_MARG * 2) \ 2), 2
        'If mMap.RoomSearch2(Index, est) Then pctMap.Circle (CoordX + (ROOM_DIM - ROOM_MARG * 2) - 1, CoordY + (ROOM_DIM - ROOM_MARG * 2) \ 2), 2
        
        If mMap.Room(Index).Tag <> "" Then
            With mMap.Room(Index)
                pctMap.DrawWidth = 2
                Area.Top = .TagY + mOrigineY
                Area.Left = .TagX + mOrigineX
                Area.Bottom = Area.Top + .TagH
                Area.Right = Area.Left + .TagW
                If Index = mMap.Selected Then
                    SelectObject pctMap.hdc, GetStockObject(LTGRAY_BRUSH)
                Else
                    SelectObject pctMap.hdc, GetStockObject(WHITE_BRUSH)
                End If
                pctMap.Line (.RealX + mOrigineX - ROOM_MARG + (ROOM_DIM \ 2), .RealY + mOrigineY - ROOM_MARG + (ROOM_DIM \ 2))-(.TagX + mOrigineX, .TagY + mOrigineY + .TagH \ 2)
                pctMap.DrawWidth = 1
                Rectangle pctMap.hdc, Area.Left, Area.Top, Area.Right, Area.Bottom
                DrawText pctMap.hdc, .Tag, Len(.Tag), Area, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
                'TextOut pctMap.hDC, mMap.Room(Index).TagX + mOrigineX, mMap.Room(Index).TagY + mOrigineY, mMap.Room(Index).Tag, Len(mMap.Room(Index).Tag)
            End With
        End If
    End If
End Sub

Private Sub Draw(Optional OnlyLinks As Boolean = False)
    Dim i As Integer
    'Dim Img As Image
    'Dim CoordX As Long, CoordY As Long

    If mMap.ActiveRoom = 0 Then Exit Sub
    
    pctMap.Cls
    mZCoord = mMap.CurZ
    
    If mFollow Then
        mOrigineX = pctMap.ScaleWidth / 2
        mOrigineY = pctMap.ScaleHeight / 2
        mOrigineX = mOrigineX - mMap.Room(mMap.ActiveRoom).RealX
        mOrigineY = mOrigineY - mMap.Room(mMap.ActiveRoom).RealY
    End If

    '[griglia]
    For i = mOrigineX To pctMap.Width Step ROOM_DIM
        pctMap.Line (i, 0)-(i, pctMap.ScaleHeight), rgb(192, 192, 192)
    Next i
        
    For i = mOrigineX To 0 Step -ROOM_DIM
        pctMap.Line (i, 0)-(i, pctMap.ScaleHeight), rgb(192, 192, 192)
    Next i
        
    For i = mOrigineY To pctMap.Height Step ROOM_DIM
        pctMap.Line (0, i)-(pctMap.ScaleWidth, i), rgb(192, 192, 192)
    Next i
    
    For i = mOrigineY To 0 Step -ROOM_DIM
        pctMap.Line (0, i)-(pctMap.ScaleWidth, i), rgb(192, 192, 192)
    Next i
    '[/griglia]
        
    For i = 1 To mMap.LinkCount
        If mMap.Room(mMap.Link(i).Room1).PosZ = mZCoord Or mMap.Room(mMap.Link(i).Room2).PosZ = mZCoord Then
            'pctMap.Line ((mMap.Link(i).x1 * ROOM_DIM + mOrigineX + (ROOM_DIM \ 2)), (mMap.Link(i).y1 * ROOM_DIM + mOrigineY + (ROOM_DIM \ 2)))-((mMap.Link(i).x2 * ROOM_DIM + mOrigineX + (ROOM_DIM \ 2)), (mMap.Link(i).y2 * ROOM_DIM + mOrigineY + (ROOM_DIM \ 2)))
            'Debug.Print "link from " & mMap.Link(i).Room1 & " to " & mMap.Link(i).Room2
            pctMap.Line (mMap.Room(mMap.Link(i).Room1).RealX + (ROOM_DIM \ 2) + mOrigineX - ROOM_MARG, mMap.Room(mMap.Link(i).Room1).RealY + (ROOM_DIM \ 2) + mOrigineY - ROOM_MARG)-(mMap.Room(mMap.Link(i).Room2).RealX + (ROOM_DIM \ 2) + mOrigineX - ROOM_MARG, mMap.Room(mMap.Link(i).Room2).RealY + (ROOM_DIM \ 2) + mOrigineY - ROOM_MARG)
        End If
    Next i

    If mMap.Selected = 0 Then RaiseEvent RoomTitle(mMap.Room(mMap.ActiveRoom).Caption)
    
    If Not OnlyLinks Then
        For i = 1 To mMap.Count
            DrawRoom (i)
        Next i
    End If
            
    pctMap.Refresh
    
    'lblPos.Caption = mMap.CurX & ";" & mMap.CurY & ";" & mMap.CurZ
End Sub

Public Sub BeginMap()
    If Not mMapMode = MAPMODE_PAUSE Then
        mMap.Clear
        mOrigineX = (pctMap.ScaleWidth - ROOM_DIM) / 2
        mOrigineY = (pctMap.ScaleHeight - ROOM_DIM) / 2
        mMap.Add "", 0, 0, 0
        Draw
    End If
End Sub

Public Sub AddRoom(Mov As Mapper_Mov, Caption As String)
    Dim X As Integer, Y As Integer, z As Integer
    Dim cx As Integer, cy As Integer, cz As Integer
    Dim Index As Integer
    
    If Not mMapMode = MAPMODE_PAUSE Then
        X = mMap.CurRealX
        Y = mMap.CurRealY
        z = mMap.CurZ
        
        Select Case Mov
            Case nord 'x+1
                Y = Y - ROOM_DIM
                cy = -1
            Case sud 'x-1
                Y = Y + ROOM_DIM
                cy = 1
            Case est 'y+1
                X = X + ROOM_DIM
                cx = 1
            Case ovest 'y-1
                X = X - ROOM_DIM
                cx = -1
            Case alto 'z+1
                z = z + 1
                cz = 1
            Case basso 'z-1
                z = z - 1
                cz = -1
        End Select
        
        'cerca la presenza di links nella direzione prescelta
        'Index = mMap.RoomSearch(mMap.CurX, mMap.CurY, mMap.CurZ, cx, cy, cz)
        Index = mMap.RoomSearch2(mMap.ActiveRoom, Mov)
        If Index <> 0 Then
            If mMap.Room(Index).Caption = "" Then
                If mMapMode = MAPMODE_ONLINE Then
                    mMap.Room(Index).Caption = Caption
                Else
                    mMap.Room(Index).Caption = InputBox("Inserisci un titolo per la stanza")
                End If
            End If
            If Not mMap.Room(Index).Caption = "" Then mMap.Jump Index
        Else
            If mMapMode = MAPMODE_ONLINE Then
            '    mMap.Add Caption, x, y, z
                mMap.Add2 Caption, Mov
            Else
                'Index = mMap.Search(x, y, z)
                Index = mMap.Search2(X, Y, z)
                If Index <> 0 Then Caption = mMap.Room(Index).Caption
                Caption = InputBox("Inserisci un titolo per la stanza", , Caption)
            '    If Not Caption = "" Then mMap.Add Caption, x, y, z
                If Not Caption = "" Then mMap.Add2 Caption, Mov
            End If
        End If
        
        mMod = True
        Draw
    End If
End Sub

Private Sub imgNAlto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMapMode = MAPMODE_ONLINE Then
        RaiseEvent Send("a" & vbCrLf)
    ElseIf mMapMode = MAPMODE_OFFLINE Then
        AddRoom alto, ""
    End If
End Sub

Private Sub imgNAlto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgAlto.Picture Then SetRosaPicture imgAlto
End Sub

Private Sub imgNBasso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMapMode = MAPMODE_ONLINE Then
        RaiseEvent Send("b" & vbCrLf)
    ElseIf mMapMode = MAPMODE_OFFLINE Then
        AddRoom basso, ""
    End If
End Sub

Private Sub imgNBasso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgBasso.Picture Then SetRosaPicture imgBasso
End Sub

Private Sub imgNEst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMapMode = MAPMODE_ONLINE Then
        RaiseEvent Send("e" & vbCrLf)
    ElseIf mMapMode = MAPMODE_OFFLINE Then
        AddRoom est, ""
    End If
End Sub

Private Sub imgNEst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgEst.Picture Then SetRosaPicture imgEst
End Sub

Private Sub imgNNord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMapMode = MAPMODE_ONLINE Then
        RaiseEvent Send("n" & vbCrLf)
    ElseIf mMapMode = MAPMODE_OFFLINE Then
        AddRoom nord, ""
    End If
End Sub

Private Sub imgNNord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgNord.Picture Then SetRosaPicture imgNord
End Sub

Private Sub imgNOvest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMapMode = MAPMODE_ONLINE Then
        RaiseEvent Send("o" & vbCrLf)
    ElseIf mMapMode = MAPMODE_OFFLINE Then
        AddRoom ovest, ""
    End If
End Sub

Private Sub imgNOvest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgOvest.Picture Then SetRosaPicture imgOvest
End Sub

Private Sub imgNSud_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMapMode = MAPMODE_ONLINE Then
        RaiseEvent Send("s" & vbCrLf)
    ElseIf mMapMode = MAPMODE_OFFLINE Then
        AddRoom sud, ""
    End If
End Sub

Private Sub imgNSud_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgSud.Picture Then SetRosaPicture imgSud
End Sub

Private Sub mnuCaption_Click()
    Dim Capt As String

    Capt = InputBox("Inserisci il nuovo titolo per la stanza", , mMap.Room(mSel).Caption)
    If Not Capt = "" Then mMap.Room(mSel).Caption = Capt
End Sub

Private Sub mnuChangeActive_Click()
    mMap.ActiveRoom = mSel
    Draw
End Sub

Private Sub mnuDelete_Click()
    Dim Index As Integer

    If Not mSel = mMap.ActiveRoom Then
        Index = 1
        Do Until Index = 0
            Index = mMap.SearchLink2(mSel)
            If Index <> 0 Then mMap.DeleteLink (Index)
        Loop
        mMap.DeleteRoom mSel
        Draw
    End If
End Sub

Private Sub mnuImage_Click(Index As Integer)
    If Not mnuImage(Index).Caption = "(normale)" Then
        mMap.Room(mSel).Image = mnuImage(Index).Caption & ".bmp"
    Else
        mMap.Room(mSel).Image = ""
    End If
    mMod = True
End Sub

Private Sub mnuTag_Click()
    Dim Capt As String

    With mMap.Room(mSel)
        Capt = InputBox("Inserisci un commento per " & .Caption, , .Tag)
        If Capt <> "" Then
            .Tag = Capt
            .TagX = .RealX + ROOM_DIM
            .TagW = pctMap.TextWidth(Capt) + 6
            .TagH = pctMap.TextHeight(Capt) + 6
            .TagY = .RealY - .TagH
        End If
    End With
End Sub

Private Sub pctMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSetOrigin Then
        Draw
        ReleaseCapture
        mDragStartX = 0: mDragStartY = 0
        pctMap.MousePointer = 0
    Else
        If mMap.Selected <> 0 Then
            mSel = mMap.Selected
            mDragStartX = X
            mDragStartY = Y
        ElseIf mMap.Selected = 0 Then
            mSel = 0
            If Button = 2 Then RaiseEvent ContextMenu
        End If
    End If
End Sub

Private Sub SetRosaPicture(img As Image)
    Dim Bordi As cGrafica

    pctRosa.Picture = img.Picture
    pctRosa.Width = img.Width
    pctRosa.Height = img.Height
    pctRosa.Top = pctMap.Height - pctRosa.Height
    pctRosa.Left = pctMap.Width - pctRosa.Width
    
    Set Bordi = New cGrafica
    Bordi.DisegnaBordi pctRosa.hdc, 0, 0, pctRosa.ScaleWidth, pctRosa.ScaleHeight, 0, 2, 0, 0, 0
    pctRosa.Refresh
    Set Bordi = Nothing
End Sub

Private Sub pctMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RoomX As Integer, RoomY As Integer
    Dim sel As Integer

    RaiseEvent MouseMove
        
    If mSetOrigin And mDragStartX > 0 Then
        mOrigineX = mDragStartX + X - pctMap.ScaleWidth / 2
        mOrigineY = mDragStartY + Y - pctMap.ScaleHeight / 2
        Draw True
    Else
        If Not pctRosa.Picture = imgPiccola.Picture Then SetRosaPicture imgPiccola
        If Button = 0 Then
            RoomX = (X - mOrigineX) ' \ ROOM_DIM
            RoomY = (Y - mOrigineY) ' \ ROOM_DIM
            'If RoomX < 0 Then RoomX = RoomX - ROOM_DIM
            'If RoomY < 0 Then RoomY = RoomY - ROOM_DIM
            'RoomX = RoomX \ ROOM_DIM
            'RoomY = RoomY \ ROOM_DIM
            sel = mMap.Search2(RoomX, RoomY, mZCoord, mTag)
            'lblPos.Caption = RoomX & ";" & RoomY & ";" & mZCoord & " (r" & sel & ")"
            If sel <> mMap.Selected Then
                If sel = 0 Then
                    RaiseEvent RoomTitle(mMap.Room(mMap.ActiveRoom).Caption)
                Else
                    RaiseEvent RoomTitle(mMap.Room(sel).Caption)
                End If
                'mMap.Selected = Sel
                SelectRoom sel
                'Draw
            End If
        ElseIf Button = 1 And mSel <> 0 Then
            mDragged = True
            If Not mTag Then
                'mMap.Room(mSel).RealX = (X - mDragStartX) - mOrigineX
                'mMap.Room(mSel).RealY = (Y - mDragStartY) - mOrigineY
                mMap.Room(mSel).RealX = mMap.Room(mSel).RealX + (X - mDragStartX)
                mMap.Room(mSel).RealY = mMap.Room(mSel).RealY + (Y - mDragStartY)
                mDragStartX = X
                mDragStartY = Y
            ElseIf mTag Then
                mMap.Room(mSel).TagX = mMap.Room(mSel).TagX + (X - mDragStartX)
                mMap.Room(mSel).TagY = mMap.Room(mSel).TagY + (Y - mDragStartY)
                mDragStartX = X
                mDragStartY = Y
            End If
            'DrawRoom mSel
            Draw
            mMod = True
        End If
    End If
End Sub

Private Sub pctMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSetOrigin Then
        If pctMap.MousePointer = 0 Then
            mSetOrigin = False
        Else
            SetCapture pctMap.hwnd
        End If
    Else
        If mMap.Selected <> 0 Then
            'mSel = mMap.Selected
            If Not mDragged And Not mTag Then
                If Button = 1 Then
                    'mnuRoom.ShowPopup
                    PopupMenu mnuRoom
                ElseIf Button = 2 Then
                    'mnuImages.ShowPopup
                    PopupMenu mnuImages
                End If
            ElseIf mDragged And Not mTag Then
                If mMap.Room(mSel).RealX > 0 Then
                    mMap.Room(mSel).RealX = ((mMap.Room(mSel).RealX - ROOM_MARG + (ROOM_DIM \ 2)) \ ROOM_DIM) * ROOM_DIM + ROOM_MARG
                Else
                    mMap.Room(mSel).RealX = ((mMap.Room(mSel).RealX - ROOM_MARG - (ROOM_DIM \ 2)) \ ROOM_DIM) * ROOM_DIM + ROOM_MARG
                End If
                
                If mMap.Room(mSel).RealY > 0 Then
                    mMap.Room(mSel).RealY = ((mMap.Room(mSel).RealY - ROOM_MARG + (ROOM_DIM \ 2)) \ ROOM_DIM) * ROOM_DIM + ROOM_MARG
                Else
                    mMap.Room(mSel).RealY = ((mMap.Room(mSel).RealY - ROOM_MARG - (ROOM_DIM \ 2)) \ ROOM_DIM) * ROOM_DIM + ROOM_MARG
                End If
                Draw
            End If
        End If
    End If
    mDragged = False
End Sub

Private Sub pctRosa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pctRosa.Picture = imgVenti.Picture Then SetRosaPicture imgVenti
End Sub

Private Sub LoadImages()
    Dim Path As String, Image As String
    Dim sPath As String, Connect As cConnector
    Dim Count As Integer

    'Set Connect = New cConnector
    '    sPath = Connect.SkinPath
    'Set Connect = Nothing
    
    'Set mnuImages = New cMenu
    Path = App.Path & "\mapimages\"
    'With mnuImages
        Image = Dir$(Path)
        Do Until Image = ""
            If Right$(Image, 4) = ".bmp" Then
                'FileCopy Path & Image, sPath & Image
                '.Add Left$(Image, Len(Image) - 4), "mnuImage", False, Path & Image
                If Not Count = 0 Then Load mnuImage(Count)
                mnuImage(Count).Caption = Left$(Image, Len(Image) - 4)
                mnuImage(Count).Visible = True
                Count = Count + 1
            End If
            Image = Dir$()
        Loop
        '.Draw Nothing, "popup"
    'End With
End Sub

Private Sub UserControl_Initialize()
    Set mMap = New cMap

    LoadImages
    
    'Set mnuRoom = New cMenu
    'mnuRoom.Add "Modifica titolo", "mnuCaption"
    'mnuRoom.Add "Elimina stanza", "mnuDelete"
    'mnuRoom.Add "Inserisci commento", "mnuTag"
    'mnuRoom.Add "Posiziona personaggio", "mnuChangeActive"
    'mnuRoom.Draw Nothing, "popup"

    'Set mnuMap = New cMenu
    'mnuMap.Add "Mostra griglia", "mnuGriglia"
    'mnuMap.Draw Nothing, "popup"

    SetRosaPicture imgPiccola
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    pctMap.Left = MARGINE
    pctMap.Top = MARGINE
    pctMap.Width = UserControl.ScaleWidth - MARGINE * 2
    pctMap.Height = UserControl.ScaleHeight - MARGINE * 2
    pctRosa.Left = pctMap.ScaleWidth - pctRosa.Width
    pctRosa.Top = pctMap.ScaleHeight - pctRosa.Height
    SetRosaPicture imgPiccola

    If mMap.Count > 0 Then Draw
End Sub
