VERSION 5.00
Begin VB.Form frmEaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uh-oh!"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2565
   ControlBox      =   0   'False
   Icon            =   "frmEaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   171
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMatrix 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   55
      Left            =   975
      Top             =   525
   End
   Begin VB.TextBox txtPass 
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
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   2100
      Width           =   1965
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      Caption         =   "close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Left            =   270
      Top             =   2070
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   150
      Picture         =   "frmEaster.frx":000C
      Top             =   150
      Width           =   2250
   End
End
Attribute VB_Name = "frmEaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLUMNLEN As Integer = 30
Private COLUMNCOUNT As Integer

Private Type CharacterColumn
    startX As Long
    startY As Long
    Char(1 To COLUMNLEN) As Integer
    Color(1 To COLUMNLEN) As Long
    Count As Integer
End Type

Private mHDC As Long, mWidth As Long, mHeight As Long
Private mColumn() As CharacterColumn
Private mColors(1 To COLUMNLEN) As Long
Private mTimerCount As Integer
Private mCharHeight As Integer
Private mCharWidth As Integer
Private WithEvents mDest As PictureBox
Attribute mDest.VB_VarHelpID = -1

Private Sub Form_Load()
    txtPass.BackColor = rgb(230, 230, 230)
    txtPass.ForeColor = rgb(150, 150, 150)
    lblClose.BackColor = txtPass.BackColor
    lblClose.ForeColor = txtPass.ForeColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnloadMatrix
End Sub

Private Sub lblClose_Click()
    Unload Me
    Set frmEaster = Nothing
End Sub

Private Sub mDest_Resize()
    mWidth = mDest.ScaleWidth
    mHeight = mDest.ScaleHeight
    mHDC = mDest.hdc
End Sub

Private Sub tmrMatrix_Timer(Index As Integer)
    Randomize
    'If mTimerCount < COLUMNCOUNT And (Rnd * 10) > 8 Then
    If mTimerCount < COLUMNCOUNT And mColumn(Index).Count = 0 Then
        mTimerCount = mTimerCount + 1
        On Error Resume Next
            Load tmrMatrix(mTimerCount)
            'tmrMatrix(mTimerCount).Interval = Rnd * 40 + 55
            tmrMatrix(mTimerCount).Enabled = True
        On Error GoTo 0
    End If
    
    If Not Index = 0 Then
        ExecColumn mColumn(Index), Index
    
        If mColumn(Index).Count = 0 Then
            tmrMatrix(Index).Interval = Rnd * 100 + 55
        End If
    Else
        tmrMatrix(0).Enabled = False
    End If
End Sub

Private Sub ExecColumn(ByRef C As CharacterColumn, ByVal Index As Integer)
    Dim i As Integer, Count As Integer
    
    Randomize
    If C.Count = 0 Then
        C.startX = Rnd * mWidth
        C.startY = (Rnd * (mHeight + 200)) - 200
        For i = 1 To mTimerCount
            If Index <> i Then
                With mColumn(i)
                    If C.startX >= .startX - mCharWidth And C.startX <= .startX + mCharWidth Then
                        If Abs(C.startY - .startY) < COLUMNLEN * mCharHeight Then
                            Exit Sub
                        End If
                    End If
                End With
            End If
        Next i
    End If
    
    If C.Count = COLUMNLEN * 2 Then
        C.Count = 0
    ElseIf C.Count >= COLUMNLEN Then
        C.Count = C.Count + 1
        Count = C.Count - COLUMNLEN
        For i = 1 To Count
            C.Color(i) = 0
            DoEvents
        Next i
        
        For i = Count + 1 To COLUMNLEN
            C.Color(i) = mColors(COLUMNLEN - (i - Count))
            DoEvents
        Next i
    Else
        C.Count = C.Count + 1
        For i = C.Count To 1 Step -1
            C.Color(i) = mColors(C.Count - (i - 1))
            DoEvents
        Next i
        C.Char(C.Count) = CInt(Rnd * 211) + 40
    End If
    
    'If c.count = COLUMNLEN + 3 Then
    '    BkMode = OPAQUE
    'Else
    '    BkMode = TRANSPARENT
    'End If
    
    DrawColumn C
End Sub

Private Sub DrawColumn(ByRef C As CharacterColumn, Optional ByVal BkMode As Long = OPAQUE)
    Dim i As Integer
    Dim Count As Integer
    
    If C.Count < COLUMNLEN Then Count = C.Count Else Count = COLUMNLEN
    SetBkMode mHDC, BkMode
    For i = 1 To Count
        SetTextColor mHDC, C.Color(i)
        TextOut mHDC, C.startX, C.startY + (mCharHeight * (i - 1)), Chr$(C.Char(i)), 1
    Next i
    'mDest.Refresh
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If InStr(1, txtPass.Text, "morpheus", vbTextCompare) Then
            'MsgBox "right!"
            'Me.Hide
            txtPass.Visible = False
            lblClose.Visible = True
            Me.WindowState = vbMinimized
            LoadMatrix
            'Unload Me
            'Set frmEaster = Nothing
        Else
            Unload Me
            Set frmEaster = Nothing
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub LoadMatrix()
    Dim Sfum As cSfum
    Dim i As Integer
    Dim Dest As PictureBox
    
    Set Sfum = New cSfum
    Sfum.AggiungiColore 16777215, 0
    Sfum.AggiungiColore rgb(0, 255, 0), 20
    'Sfum.AggiungiColore rgb(255, 255, 0), 20
    Sfum.AggiungiColore 0, 100
    Sfum.NSfumature = COLUMNLEN
    
    For i = 1 To COLUMNLEN
        mColors(i) = Sfum.Sfumatura(i)
    Next i
    
    'Sfum.StampaSfumatura 0, 0, 100, 100, 0, frmMain.txtMud.hDC
    Set Sfum = Nothing
    mTimerCount = 0
    
    If mDest Is Nothing Then Set mDest = frmMain.txtMud.SetCustomControl(True)
    mWidth = mDest.ScaleWidth
    mHeight = mDest.ScaleHeight
    mDest.AutoRedraw = False
    mHDC = mDest.hdc
    mDest.FontName = "Courier New"
    mDest.FontSize = 8
    mDest.FontBold = True
    mCharHeight = mDest.TextHeight("A")
    mCharWidth = mDest.TextWidth("A")
    COLUMNCOUNT = mWidth / mCharWidth
    ReDim mColumn(0 To COLUMNCOUNT) As CharacterColumn
    'MsgBox COLUMNCOUNT
    
    tmrMatrix(0).Enabled = True
End Sub

Private Sub UnloadMatrix()
    Dim i As Integer
    
    If Not mDest Is Nothing Then
    For i = 0 To tmrMatrix.Count - 1
        tmrMatrix(i).Enabled = False
        If Not i = 0 Then Unload tmrMatrix(i)
    Next i

    mDest.FontBold = False
    Set mDest = Nothing
    frmMain.txtMud.SetCustomControl False
    End If
End Sub
