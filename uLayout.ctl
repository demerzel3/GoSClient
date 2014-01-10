VERSION 5.00
Begin VB.UserControl uLayout 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   Begin VB.PictureBox pctModify 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   75
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   4515
   End
   Begin GoS.uBox box 
      Height          =   1515
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4065
      _extentx        =   7170
      _extenty        =   2672
   End
   Begin VB.Menu mnuLayout 
      Caption         =   "menu"
      Begin VB.Menu mnuHor 
         Caption         =   "Dividi in orizzontale"
      End
      Begin VB.Menu mnuVert 
         Caption         =   "Dividi in verticale"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Salva con nome"
      End
      Begin VB.Menu mnuFine 
         Caption         =   "Fine"
      End
   End
End
Attribute VB_Name = "uLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CREATE_NONE As Integer = 0
Private Const CREATE_TOP As Integer = 1
Private Const CREATE_BOTTOM As Integer = 2
Private Const CREATE_LEFT As Integer = 3
Private Const CREATE_RIGHT As Integer = 4

Private mLayout As cLayout

Private mMode As Integer

'variabili per modalita' LYTMODE_MODIFY
Private mSizeType As Integer
Private mInitialPos As Long
Private mFinalPos As Long
Private mCreateNew As Integer
Private mDragging As Boolean
Private mSel As Integer

'variabili per modalita' LYTMODE_DOCK
Private mFirstLoading As Boolean

Public Event ModeChanged()

Public Sub ReloadFormNames()
    Dim i As Integer
    
    If mMode = LYTMODE_DOCK Then
        For i = 1 To box.Count
            box(i).PrintWindowCaption
        Next i
    End If
End Sub

Public Property Get Mode() As Integer
    Mode = mMode
End Property

Public Sub SetMode(newMode As Integer)
    If (Not mMode = newMode) Then
        mMode = newMode
        Select Case mMode
            Case LYTMODE_MODIFY
                RetrieveDocked
                UnDockAll SW_MINIMIZE
                pctModify.Visible = True
                UserControl_Resize
            Case LYTMODE_DOCK
                SaveLayout True
                
                pctModify.Visible = False
                mFirstLoading = True
                LoadWorkspace
        End Select
    End If
End Sub

Private Sub UnDockAll(Optional nCmdShow As Long = 0)
    Dim i As Integer
    
    For i = 1 To box.Count
        box(i).UnDock nCmdShow
        box(i).Visible = False
    Next i
End Sub

Public Sub ChangeLayout(newLyt As String)
    If Not newLyt = mLayout.LytFile Then
        Me.SaveWorkspace
        UnDockAll
        Init newLyt
    End If
End Sub

Public Sub Init(Optional lyt As String = "(default).lyt")
    Dim data As cIni, i As Integer
    Dim frm As Form, Connect As cConnector
    Dim Path As String
    
    Path = App.Path & "\layouts\" & lyt
    Set Connect = New cConnector
    'Connect.Log "Inizializzazione layout in corso"
    Set data = New cIni
        data.CaricaFile Path, True
        For i = 1 To data.RetrInfo("Win_Count", 0)
            Select Case data.RetrInfo("Win<" & i & ">")
                Case "frmMain"
                    Load frmMain
                    'Connect.Log "Caricato Schermo di gioco"
                Case "frmChat"
                    Load frmChat
                    'Connect.Log "Caricata Chat"
                Case "frmMapper"
                    Load frmMapper
                    'Connect.Log "Caricato Mapper"
                Case "frmStato"
                    Load frmStato
                    'Connect.Log "Caricata finestra di stato"
                Case "frmRubrica"
                    'Connect.Log "Caricata Rubrica"
                    Load frmRubrica
                Case "frmButtons"
                    'Connect.Log "Caricata finestra Pulsanti"
                    Load frmButtons
            End Select
        Next i
    Set data = Nothing
    
    Set mLayout = New cLayout
    
    mLayout.LoadLayout lyt
    mFirstLoading = True
    LoadWorkspace
    Connect.Log "Layout initialization completed"
    Set Connect = Nothing
End Sub

Public Sub Destroy()
    Dim data As cIni, i As Integer
    
    Set data = New cIni
        data.CaricaFile "layout.lyt"
        For i = 1 To data.RetrInfo("Win_Count", 0)
            Select Case data.RetrInfo("Win<" & i & ">_Name")
                Case "frmMain"
                    Unload frmMain
                    Set frmMain = Nothing
                Case "frmChat"
                    Unload frmChat
                    Set frmChat = Nothing
                Case "frmMapper"
                    Unload frmMapper
                    Set frmMapper = Nothing
                Case "frmStato"
                    Unload frmStato
                    Set frmStato = Nothing
                Case "frmRubrica"
                    Unload frmRubrica
                    Set frmRubrica = Nothing
                Case "frmButtons"
                    Unload frmButtons
                    Set frmButtons = Nothing
            End Select
        Next i
    Set data = Nothing
End Sub

Private Sub RetrieveDocked()
    Dim i As Integer, win As String
    
    For i = 1 To mLayout.Count
        win = box(i).DockedFormName
        If win = "frmMain" Or _
           win = "frmMapper" Or _
           win = "frmChat" Or _
           win = "frmStato" Or _
           win = "frmRubrica" Or _
           win = "frmButtons" Or _
           win = "" Then
            mLayout.box(i).Window = win
        End If
    Next i
End Sub

Public Sub SaveWorkspace()
    RetrieveDocked
    SaveLayout
End Sub

Public Sub LoadWorkspace()
    Dim i As Integer
    Dim ScaleWidth As Long, ScaleHeight As Long
    Dim frm As Form

    'If Changed Then mLayout.LoadLayout "Layout.lyt"
    If mMode = LYTMODE_MODIFY Then Exit Sub
    
    With mLayout
        ScaleWidth = UserControl.ScaleWidth
        ScaleHeight = UserControl.ScaleHeight
        For i = 1 To .Count
            'RaiseEvent LoadBox(i)
            On Error Resume Next
            Load box(i)
            With box(i)
                'SetParent box(i).hWnd, UserControl.hWnd
                .Left = (ScaleWidth * mLayout.box(i).Left) / 100
                .Top = (ScaleHeight * mLayout.box(i).Top) / 100
                .Width = (ScaleWidth * mLayout.box(i).Width) / 100
                .Height = (ScaleHeight * mLayout.box(i).Height) / 100
                'If mFirstLoading Then
                    '.Hide
                    If Not .Visible Then .Visible = True
                'End If
            End With
            
            If mFirstLoading Then
                Select Case .box(i).Window
                    Case "frmMain"
                        Set frm = frmMain
                    Case "frmChat"
                        Set frm = frmChat
                    Case "frmMapper"
                        Set frm = frmMapper
                    Case "frmStato"
                        Set frm = frmStato
                    Case "frmRubrica"
                        Set frm = frmRubrica
                    Case "frmButtons"
                        Set frm = frmButtons
                End Select
            End If
                        
            If mFirstLoading And Not frm Is Nothing Then
                box(i).SetForm frm
                box(i).Dock
            ElseIf Not mFirstLoading Then
                box(i).ResizeForm
            End If
            
            Set frm = Nothing
        Next i
        
        If Not .Count = 0 Then mFirstLoading = False
    End With
End Sub

Public Sub LoadLayout()
    mLayout.LoadLayout "layout.lyt"
    If mMode = LYTMODE_MODIFY Then Draw
End Sub

Public Sub SaveLayout(Optional OnlyBoxes As Boolean = False)
    mLayout.SaveLayout OnlyBoxes
End Sub

Public Sub Draw()
    Dim rcBox As RECT
    Dim i As Integer
    Dim win As String

    If Not mMode = LYTMODE_MODIFY Then Exit Sub
        
    'pctModify.FillColor = 16777215
    pctModify.BackColor = rgb(230, 230, 230)
    pctModify.ForeColor = 0
    pctModify.DrawWidth = 1
    pctModify.Cls
    For i = 1 To mLayout.Count
        rcBox.Left = (pctModify.ScaleWidth * mLayout.box(i).Left) / 100
        rcBox.Top = (pctModify.ScaleHeight * mLayout.box(i).Top) / 100
        rcBox.Right = (pctModify.ScaleWidth * mLayout.box(i).Right) / 100
        rcBox.Bottom = (pctModify.ScaleHeight * mLayout.box(i).Bottom) / 100
        Rectangle pctModify.hdc, rcBox.Left, rcBox.Top, rcBox.Right, rcBox.Bottom
    Next i
End Sub

Private Sub mnuFine_Click()
    SetMode LYTMODE_DOCK
    RaiseEvent ModeChanged
End Sub

Private Sub mnuHor_Click()
    Dim OldWidth As Long

    With mLayout.box(mSel)
        OldWidth = .Width
        .Right = .Left + (.Width / 2)
        mLayout.AddBox .Right, .Top, OldWidth - .Width + .Right, .Bottom
    End With
    Draw
End Sub

Private Sub mnuSaveAs_Click()
    Dim Name As String
    Dim Connect As cConnector
    
    Set Connect = New cConnector
    'Name = InputBox("Inserisci un nuovo nome per questo layout")
    Name = InputBox(Connect.Lang("layout", "NewName"))
    If Not Name = "" Then
        If Not Right$(Name, 4) = ".lyt" Then Name = Name & ".lyt"
        mLayout.SaveLayout True, Name
    
            Connect.ProfConf.AddInfo "Layout", Name
            Connect.Envi.SaveProfConfig
    End If
    Set Connect = Nothing
End Sub

Private Sub mnuVert_Click()
    Dim OldHeight As Long

    With mLayout.box(mSel)
        OldHeight = .Height
        .Bottom = .Top + (.Height / 2)
        mLayout.AddBox .Left, .Bottom, .Right, OldHeight - .Height + .Bottom
    End With
    Draw
End Sub

Private Sub UserControl_Initialize()
    Set mLayout = New cLayout
    mFirstLoading = True
    
    'UserControl.BackColor = GOSRGB_FORM_Back
    UserControl.BackColor = GetSysColor(COLOR_3DFACE)
End Sub

Private Sub pctModify_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rc As RECT, i As Integer
    Dim Connect As cConnector

    If Not mMode = LYTMODE_MODIFY Then Exit Sub
        
    If Button = 1 Then
        If HitTest(X, Y) Then
            mDragging = True
        End If
    ElseIf Button = 2 Then
        For i = 1 To mLayout.Count
            rc = GetAbsoluteRect(i)
            If PtInRect(rc, X, Y) Then
                mSel = i
                Exit For
            End If
        Next i
        
        Set Connect = New cConnector
            With Connect
                mnuHor.Caption = .Lang("layout", "SplitHor")
                mnuVert.Caption = .Lang("layout", "SplitVert")
                mnuSaveAs.Caption = .Lang("layout", "SaveAs")
                mnuFine.Caption = .Lang("layout", "End")
            End With
        Set Connect = Nothing
        PopupMenu mnuLayout
    End If
End Sub

Private Sub pctModify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mMode = LYTMODE_MODIFY Then Exit Sub
    
    If Not mDragging Then
        HitTest X, Y
    Else
        If mSizeType = vbSizeWE Then
            mFinalPos = Int(100 * X / pctModify.ScaleWidth)
        ElseIf mSizeType = vbSizeNS Then
            mFinalPos = Int(100 * Y / pctModify.ScaleHeight)
        End If
        
        If mFinalPos < 0 Then mFinalPos = 0
        If mFinalPos > 100 Then mFinalPos = 100
        
        Draw
        pctModify.ForeColor = 255
        pctModify.DrawWidth = 3
        If mSizeType = vbSizeWE Then
            If X < 0 Then X = 0
            If X > pctModify.ScaleWidth Then X = pctModify.ScaleWidth
            pctModify.Line (X, 0)-(X, pctModify.ScaleHeight)
        ElseIf mSizeType = vbSizeNS Then
            If Y < 0 Then Y = 0
            If Y > pctModify.ScaleHeight Then Y = pctModify.ScaleHeight
            pctModify.Line (0, Y)-(pctModify.ScaleWidth, Y)
        End If
        
        Debug.Print Int(100 * X / pctModify.ScaleWidth) & "%"
    End If
End Sub

Private Function HitTest(X As Single, Y As Single) As Boolean
    Dim i As Integer
    Dim rc As RECT
    Dim Changed As Boolean

    If X < 5 Then
        pctModify.MousePointer = vbSizeWE
        mCreateNew = CREATE_LEFT
        mInitialPos = 0
        Changed = True
    ElseIf X > pctModify.ScaleWidth - 5 Then
        pctModify.MousePointer = vbSizeWE
        mCreateNew = CREATE_RIGHT
        mInitialPos = 100
        Changed = True
    ElseIf Y < 5 Then
        pctModify.MousePointer = vbSizeNS
        mCreateNew = CREATE_TOP
        mInitialPos = 0
        Changed = True
    ElseIf Y > pctModify.ScaleHeight - 5 Then
        pctModify.MousePointer = vbSizeNS
        mCreateNew = CREATE_BOTTOM
        mInitialPos = 100
        Changed = True
    End If
    
    
    If Not Changed Then
        For i = 1 To mLayout.Count
            rc = GetAbsoluteRect(i)
            If X >= rc.Left - 2 And X <= rc.Left + 2 And rc.Left <> 0 Then
                pctModify.MousePointer = vbSizeWE
                Changed = True
                mInitialPos = mLayout.box(i).Left
                mCreateNew = CREATE_NONE
                Exit For
            ElseIf Y >= rc.Top - 2 And Y <= rc.Top + 2 And rc.Top <> 0 Then
                pctModify.MousePointer = vbSizeNS
                Changed = True
                mInitialPos = mLayout.box(i).Top
                mCreateNew = CREATE_NONE
                Exit For
            End If
        Next i
    End If
    
    If Not Changed Then
        pctModify.MousePointer = vbDefault
        HitTest = False
    Else
        mSizeType = pctModify.MousePointer
        HitTest = True
    End If
End Function

Private Function GetAbsoluteRect(Index As Integer) As RECT
    With GetAbsoluteRect
        .Left = (pctModify.ScaleWidth * mLayout.box(Index).Left) / 100
        .Top = (pctModify.ScaleHeight * mLayout.box(Index).Top) / 100
        .Right = (pctModify.ScaleWidth * mLayout.box(Index).Right) / 100
        .Bottom = (pctModify.ScaleHeight * mLayout.box(Index).Bottom) / 100
    End With
End Function

Private Sub pctModify_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    If Not mMode = LYTMODE_MODIFY Then Exit Sub
    
    If mDragging Then
        mDragging = False
        
        For i = 1 To mLayout.Count
            With mLayout.box(i)
                If mSizeType = vbSizeWE Then
                    If .Left = mInitialPos Then .Left = mFinalPos
                    If .Right = mInitialPos Then .Right = mFinalPos
                    If mFinalPos > mInitialPos Then
                        If .Left > mInitialPos And .Left < mFinalPos Then .Left = mFinalPos
                        If .Right > mInitialPos And .Right < mFinalPos Then .Right = mFinalPos
                    Else
                        If .Left < mInitialPos And .Left > mFinalPos Then .Left = mFinalPos
                        If .Right < mInitialPos And .Right > mFinalPos Then .Right = mFinalPos
                    End If
                Else
                    If .Top = mInitialPos Then .Top = mFinalPos
                    If .Bottom = mInitialPos Then .Bottom = mFinalPos
                    If mFinalPos > mInitialPos Then
                        If .Top > mInitialPos And .Top < mFinalPos Then .Top = mFinalPos
                        If .Bottom > mInitialPos And .Bottom < mFinalPos Then .Bottom = mFinalPos
                    Else
                        If .Top < mInitialPos And .Top > mFinalPos Then .Top = mFinalPos
                        If .Bottom < mInitialPos And .Bottom > mFinalPos Then .Bottom = mFinalPos
                    End If
                End If
            End With
        Next i
        
        If Not mCreateNew = CREATE_NONE Then
            Select Case mCreateNew
                Case CREATE_LEFT
                    mLayout.AddBox 0, 0, mFinalPos, 100
                Case CREATE_TOP
                    mLayout.AddBox 0, 0, 100, mFinalPos
                Case CREATE_RIGHT
                    mLayout.AddBox mFinalPos, 0, 100, 100
                Case CREATE_BOTTOM
                    mLayout.AddBox 0, mFinalPos, 100, 100
            End Select
            Debug.Print "box creato", "count = " & mLayout.Count
        End If
        
        For i = 1 To mLayout.Count
            If i > mLayout.Count Then Exit For
            If mLayout.box(i).Width = 0 Or mLayout.box(i).Height = 0 Then
                mLayout.RemoveBox i
                Debug.Print "box " & i & " eliminato", "count = " & mLayout.Count
            End If
        Next i
        
        Draw
    End If
End Sub

Private Sub UserControl_Resize()
    If Ambient.UserMode And mMode = LYTMODE_MODIFY Then
        pctModify.Width = UserControl.ScaleWidth - pctModify.Left * 2
        pctModify.Height = UserControl.ScaleHeight - pctModify.Top * 2
        Draw
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mLayout = Nothing
End Sub
