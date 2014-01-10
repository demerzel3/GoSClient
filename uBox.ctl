VERSION 5.00
Begin VB.UserControl uBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrNothing 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2025
      Top             =   1800
   End
   Begin VB.PictureBox pctMenu 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   30
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   3825
      Begin VB.Image imgMenu 
         Height          =   120
         Left            =   3225
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox pctCont 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   306
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   4590
   End
   Begin VB.Image img 
      Height          =   120
      Index           =   2
      Left            =   1575
      Picture         =   "uBox.ctx":0000
      Top             =   3150
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image img 
      Height          =   120
      Index           =   1
      Left            =   1275
      Picture         =   "uBox.ctx":0092
      Top             =   3150
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "uBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum boxImgState
    idb_undockclose = 2
    idb_dockclose = 0
    idb_docknew = 1
End Enum

Private mImgState As boxImgState

Private mDockForm As Form
Private mDocked As Boolean

Public Sub ResizeForm()
    Dim rcCont As RECT
    
    If Not mDockForm Is Nothing And mDocked Then
        pctCont.Visible = False
        Call GetClientRect(pctCont.hwnd, rcCont)
        MoveWindow mDockForm.hwnd, 0, 0, _
            rcCont.Right, rcCont.Bottom, True
        pctCont.Visible = True
    End If
End Sub

Public Sub Dock()
    Dim Style As Long, Caption As String
    
    If Not mDockForm Is Nothing Then
        UserControl.Cls
        pctCont.Visible = False
        Style = GetWindowLong(mDockForm.hwnd, GWL_STYLE)
        If (Style And WS_CAPTION) = WS_CAPTION Then Style = Style Xor WS_CAPTION
        If (Style And WS_SIZEBOX) = WS_SIZEBOX Then Style = Style Xor WS_SIZEBOX
        If (Style And WS_CHILD) = 0 Then Style = Style Xor WS_CHILD
        SetWindowLong mDockForm.hwnd, GWL_STYLE, Style
        SetParent mDockForm.hwnd, pctCont.hwnd
        ShowWindow mDockForm.hwnd, SW_MAXIMIZE
        pctCont.Visible = True
        mDockForm.Tag = "d"
        mDocked = True
        ChangeMenuState idb_undockclose
        pctMenu.Top = 0
        'tmrNothing.Enabled = True
        
        PrintWindowCaption
    End If
End Sub

Public Sub PrintWindowCaption()
    Dim Caption As String
    
    If Not mDockForm Is Nothing Then
        pctMenu.Cls
        pctMenu.CurrentX = 0
        pctMenu.CurrentY = 0
        Caption = mDockForm.Caption
        If InStr(1, Caption, "(") <> 0 Then Caption = Left$(Caption, InStr(1, Caption, "(") - 1)
        If InStr(1, Caption, "[") <> 0 Then Caption = Left$(Caption, InStr(1, Caption, "[") - 1)
        pctMenu.Print UCase$(Caption)
    End If
End Sub

Public Sub CloseForm()
    If Not mDockForm Is Nothing And mDocked Then
        If Not TypeName(mDockForm) = "frmMain" Then
            tmrNothing.Enabled = False
            mDockForm.Tag = ""
            Unload mDockForm
            Set mDockForm = Nothing
            mDocked = False
            ChangeMenuState idb_docknew
            pctMenu.Top = 2
            pctCont.Visible = False
            UserControl_Paint
            pctMenu.Cls
        End If
    End If
End Sub

Public Sub UnDock(Optional nCmdShow As Long = 0)
    Dim Style As Long

    If Not mDockForm Is Nothing And mDocked Then
        tmrNothing.Enabled = False
        Style = GetWindowLong(mDockForm.hwnd, GWL_STYLE)
        If (Style And WS_CAPTION) = 0 Then Style = Style Or WS_CAPTION
        If (Style And WS_SIZEBOX) = 0 Then Style = Style Or WS_SIZEBOX
        If (Style And WS_CHILD) = WS_CHILD Then Style = Style Xor WS_CHILD
        SetWindowLong mDockForm.hwnd, GWL_STYLE, Style
        SetParent mDockForm.hwnd, 0
        mDockForm.Tag = ""
        mDockForm.Show vbModeless, frmBase
        If nCmdShow <> 0 Then ShowWindow mDockForm.hwnd, nCmdShow
        mDocked = False
        ChangeMenuState idb_docknew
        
        Set mDockForm = Nothing
        pctMenu.Top = 2
        
        pctCont.Visible = False
        UserControl_Paint
        pctMenu.Cls
    End If
End Sub

Public Sub SetForm(ToDock As Form)
    Set mDockForm = ToDock
End Sub

Public Sub Hide()
    pctCont.Visible = False
End Sub

Public Sub Show()
    pctCont.Visible = True
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let BackColor(data As Long)
    UserControl.BackColor = data
End Property

Public Property Get ScaleWidth() As Long
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScaleHeight() As Long
    ScaleHeight = UserControl.ScaleHeight
End Property

Private Sub ChangeMenuState(ByVal NewState As boxImgState)
    mImgState = NewState
    'Set imgMenu.Picture = LoadResPicture(mImgState, vbResBitmap)
    imgMenu.Picture = img(mImgState).Picture
End Sub

Private Function GetForm(Capt As String) As Form
    Dim Connect As cConnector
    
    Set Connect = New cConnector
    With Connect
        Select Case Capt
            Case .Lang("forms", "frmChat")
                Set GetForm = frmChat
            Case .Lang("forms", "frmMapper")
                Set GetForm = frmMapper
            Case .Lang("forms", "frmStato")
                Set GetForm = frmStato
            Case .Lang("forms", "frmRubrica")
                Set GetForm = frmRubrica
            Case .Lang("forms", "frmButtons")
                Set GetForm = frmButtons
        End Select
    End With
    Set Connect = Nothing
End Function

Private Sub imgMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim colForms As Collection
    Dim i As Integer, rtn As Integer, rtnKey As String
    Dim curForm As Form, Connect As cConnector

    X = X / Screen.TwipsPerPixelX
    Select Case mImgState
        Case idb_docknew
            If X < 8 Then
                'visualizza lista finestre dockabili
                Set colForms = New Collection
                'With colForms
                '    .Add "Chat", "frmChat"
                '    .Add "Mapper", "frmMapper"
                '    .Add "Controllo MSP", "frmMsp"
                '    .Add "Status", "frmStato"
                'End With
                
                For Each curForm In Forms
                    'On Error Resume Next
                    'colForms.Remove TypeName(curForm)
                    If curForm.Tag <> "d" Then
                        rtnKey = TypeName(curForm)
                        If rtnKey <> "frmBase" And rtnKey <> "frmNote" And _
                            rtnKey <> "frmNota" Then
                            colForms.Add curForm
                        End If
                    End If
                Next
                
                If colForms.Count > 0 Then
                    Load frmList
                    For i = 1 To colForms.Count
                        frmList.AddItem colForms.Item(i).Caption
                    Next i
                    'frmList.Caption = "Scegli la finestra da ancorare"
                    Set Connect = New cConnector
                        frmList.Caption = Connect.Lang("layout", "SelWindow")
                    Set Connect = Nothing
                    rtn = frmList.ShowForm
                    
                    Unload frmList
                    Set frmList = Nothing
                    
                    If Not rtn = -1 Then
                        Set mDockForm = colForms.Item(rtn + 1)
                        Dock
                    End If
                End If
                
            ElseIf X > 8 Then
                'visualizza lista nuove finestre disponibili
                Set Connect = New cConnector
                
                Set colForms = New Collection
                With colForms
                    .Add Connect.Lang("forms", "frmChat"), "frmChat"
                    .Add Connect.Lang("forms", "frmMapper"), "frmMapper"
                    .Add Connect.Lang("forms", "frmStato"), "frmStato"
                    .Add Connect.Lang("forms", "frmRubrica"), "frmRubrica"
                    .Add Connect.Lang("forms", "frmButtons"), "frmButtons"
                End With
                
                For Each curForm In Forms
                    On Error Resume Next
                    colForms.Remove TypeName(curForm)
                Next
                
                If colForms.Count > 0 Then
                    Load frmList
                    For i = 1 To colForms.Count
                        frmList.AddItem colForms.Item(i), colForms.Item(i)
                    Next i
                    'frmList.Caption = "Scegli la nuova finestra"
                    frmList.Caption = Connect.Lang("layout", "SelWinNew")
                    rtn = frmList.ShowForm(rtnKey)
                    
                    Unload frmList
                    Set frmList = Nothing
                    
                    If Not rtn = -1 Then
                        Set mDockForm = GetForm(rtnKey)
                        Dock
                    End If
                End If
            
                Set Connect = New cConnector
            End If
            Set colForms = Nothing
        
        Case idb_undockclose
            If X < 8 Then
                UnDock
            ElseIf X > 8 Then
                CloseForm
            End If
    End Select
End Sub

Public Function DockedFormName() As String
    If Not mDockForm Is Nothing Then
        DockedFormName = TypeName(mDockForm)
    End If
End Function

Private Sub tmrNothing_Timer()
    If mDockForm.Visible = False Then
        Set mDockForm = Nothing
        
        mDocked = False
        pctMenu.Top = 2
        ChangeMenuState idb_docknew
        
        pctCont.Visible = False
        UserControl_Paint
        tmrNothing.Enabled = False
        pctMenu.Cls
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim Back As Long

    Back = GetSysColor(COLOR_3DFACE)
    'UserControl.BackColor = GOSRGB_FORM_Back
    'pctCont.BackColor = GOSRGB_FORM_Back
    'pctMenu.BackColor = GOSRGB_FORM_Back
    UserControl.BackColor = Back
    pctCont.BackColor = Back
    pctMenu.BackColor = Back
    ChangeMenuState idb_docknew
End Sub

Private Sub UserControl_Paint()
    If mImgState = idb_docknew Then
        UserControl.ForeColor = 0
        UserControl.Cls
        Rectangle UserControl.hdc, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If mImgState = idb_docknew Then
        pctMenu.Top = 2
        UserControl.ForeColor = 0
        UserControl.Cls
        Rectangle UserControl.hdc, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
    End If
    pctMenu.Width = Me.ScaleWidth - pctMenu.Left - 5
    imgMenu.Left = pctMenu.ScaleWidth - imgMenu.Width
    'SetWindowPos Command1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    pctCont.Width = Me.ScaleWidth
    pctCont.Height = Me.ScaleHeight - pctCont.Top
End Sub
