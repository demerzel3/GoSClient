VERSION 5.00
Begin VB.UserControl uMacroBox 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "uMacroBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type umb_button
    OptionButt As Boolean
    
    NormText As String
    NormAct As String
    PresText As String
    PresAct As String
    State As Integer

    MinWidth As Integer
    Width As Long
    X As Long
    Y As Long
    
    rcButt As rect
End Type

Private Const BUTT_HEIGHT As Integer = 17
Private Const LINE_SPACE As Integer = 2
Private Const BUTT_MARG As Integer = 10

Private mButt() As umb_button
Private mSel As Integer

Public Event Clicked(Act As String)

Public Function GetState(Index As Integer) As Integer
    GetState = mButt(Index).State
End Function

Public Property Get Count() As Integer
    Count = UBound(mButt(), 1)
End Property

Public Sub Clear()
    Erase mButt()
    ReDim mButt(0) As umb_button
    UserControl.Cls
End Sub

Public Sub Add(NormText As String, NormAct As String, _
               Optional PresText As String, Optional PresAct As String, _
               Optional State As Integer = STATE_NORMAL)
    
    Dim Count As Integer

    Count = UBound(mButt(), 1) + 1
    If Count = 1 Then
        ReDim mButt(1 To 1) As umb_button
    Else
        ReDim Preserve mButt(1 To Count) As umb_button
    End If
    
    With mButt(Count)
        .OptionButt = Not (PresText = "")
        .NormText = NormText
        .NormAct = NormAct
        .PresText = PresText
        .PresAct = PresAct
        .MinWidth = ButMinWidth(Count)
        .State = State
    End With
    
    Resize True
End Sub

Private Sub Resize(Optional DoRedraw As Boolean = False)
    Dim i As Integer, j As Integer, Widths() As Integer
    Dim nButt As Integer, Count As Integer, nLines As Integer
    Dim ScaleWidth As Long, CurWidth As Long
    Dim ToAdd As Integer, ButLineCount As Integer

    Count = Me.Count
    ScaleWidth = UserControl.ScaleWidth
    For i = Count To 1 Step -1
        If CalcMinWidth(i, Widths()) <= ScaleWidth Then
            nButt = i
            Exit For
        End If
    Next i
    If nButt = 0 Then nButt = 1
    
    nLines = Count / nButt
    If Count Mod nButt > 0 Then nLines = nLines + 1
    
    For i = 1 To nLines
        CurWidth = 0
        ButLineCount = IIf(i * nButt > Count, Count, i * nButt)
        ToAdd = (ScaleWidth - Widths(i)) \ IIf(i * nButt > Count, Count Mod nButt, nButt)
        For j = ((i - 1) * nButt) + 1 To ButLineCount
            With mButt(j)
                .Width = .MinWidth - 2 + ToAdd
                .X = CurWidth
                .Y = ((i - 1) * BUTT_HEIGHT) + ((i - 1) * LINE_SPACE)
                CurWidth = CurWidth + .Width + 2
            End With
        Next j
    Next i

    If DoRedraw Then Redraw
End Sub

Private Sub Redraw()
    Dim i As Integer

    UserControl.Cls
    For i = 1 To Me.Count
        DrawButt i
    Next i
End Sub

Private Sub DrawButt(Index As Integer)
    Dim Bordi As cGrafica, CurText As String

    With mButt(Index)
        .rcButt.Left = .X
        .rcButt.Top = .Y
        .rcButt.Right = .X + .Width
        .rcButt.Bottom = .Y + BUTT_HEIGHT
        
        Set Bordi = New cGrafica
            With mButt(Index).rcButt
                'Bordi.FillRectEx UserControl.hdc, .Left, .Top, .Right - .Left, _
                    .Bottom - .Top, _
                    IIf(mButt(Index).State = STATE_NORMAL, GetSysColor(COLOR_3DFACE), _
                    GetSysColor(COLOR_3DFACE))
                Bordi.FillRectEx UserControl.hdc, .Left, .Top, .Right - .Left, _
                    .Bottom - .Top, GetSysColor(COLOR_3DFACE)
            End With
            
            Bordi.DisegnaBordi UserControl.hdc, .X, .Y, .Width, BUTT_HEIGHT, _
                Abs(CInt(.State = STATE_PRESSED)), 1, 0, 0, 30
        Set Bordi = Nothing
        
        If .State = STATE_NORMAL Or (Not .OptionButt) Then
            CurText = .NormText
        Else
            CurText = .PresText
        End If
        
        If .State = STATE_PRESSED Then Call OffsetRect(.rcButt, 1, 1)
        
        DrawText UserControl.hdc, CurText, Len(CurText), .rcButt, _
            DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
    
        If .State = STATE_PRESSED Then Call OffsetRect(.rcButt, -1, -1)
    End With
End Sub

Private Function CalcMinWidth(nButt As Integer, ByRef Widths() As Integer) As Integer
    Dim Count As Integer, GroupCount As Integer
    Dim i As Integer, j As Integer
    Dim Max As Integer

    Count = Me.Count
    GroupCount = Count / nButt
    If Count Mod nButt > 0 Then GroupCount = GroupCount + 1
        
    ReDim Widths(1 To GroupCount) As Integer
    
    For i = 1 To GroupCount
        Widths(i) = 0
        For j = ((i - 1) * nButt) + 1 To IIf(i * nButt > Count, Count, i * nButt)
            Widths(i) = Widths(i) + mButt(j).MinWidth
        Next j
    Next i
    
    Max = 1
    For i = 1 To GroupCount
        If Widths(i) > Widths(Max) Then
            Max = i
        End If
    Next i
    
    CalcMinWidth = Widths(Max)
End Function

Private Function ButMinWidth(Index As Integer) As Integer
    Dim NormWidth As Integer, PresWidth As Integer

    With mButt(Index)
        NormWidth = UserControl.TextWidth(.NormText) + BUTT_MARG
        PresWidth = UserControl.TextWidth(.PresText) + BUTT_MARG
    End With
    ButMinWidth = IIf(NormWidth >= PresWidth, NormWidth, PresWidth)
End Function

Private Function HitTest(ByVal X As Long, ByVal Y As Long) As Integer
    Dim i As Integer

    For i = 1 To Me.Count
        If PtInRect(mButt(i).rcButt, X, Y) Then
            HitTest = i
            Exit For
        End If
    Next i
End Function

Private Sub UserControl_Initialize()
    UserControl.BackColor = GetSysColor(COLOR_3DFACE)
    UserControl.ForeColor = 0
    ReDim mButt(0) As umb_button
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSel = HitTest(X, Y)
    If mSel > 0 Then
        With mButt(mSel)
            If Not .OptionButt Then
                .State = STATE_PRESSED
            Else
                If .State = STATE_NORMAL Then
                    .State = STATE_PRESSED
                Else
                    .State = STATE_NORMAL
                End If
            End If
        End With
        DrawButt mSel
        UserControl.Refresh
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Act As String

    If mSel <> 0 Then
        If Not mButt(mSel).OptionButt Then mButt(mSel).State = STATE_NORMAL
        DrawButt mSel
        UserControl.Refresh
        If mSel = HitTest(X, Y) Then
            If mButt(mSel).OptionButt Then
                If mButt(mSel).State = STATE_NORMAL Then
                    Act = mButt(mSel).PresAct
                Else
                    Act = mButt(mSel).NormAct
                End If
            Else
                Act = mButt(mSel).NormAct
            End If
            RaiseEvent Clicked(Act)
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    Resize True
End Sub
