VERSION 5.00
Begin VB.UserControl uOutBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   Begin VB.PictureBox pctHistory 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   2340
      Left            =   0
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   526
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   7890
   End
   Begin VB.PictureBox pctSplit 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   6915
   End
   Begin GoS.uScroller Scroller 
      Height          =   5340
      Left            =   8025
      TabIndex        =   1
      Top             =   150
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   9419
   End
   Begin VB.PictureBox pctText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5340
      Left            =   0
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   356
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   526
      TabIndex        =   0
      Top             =   225
      Width           =   7890
   End
   Begin VB.Image imgClose 
      Height          =   165
      Left            =   4275
      Picture         =   "uOutBox.ctx":0000
      Top             =   15
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "uOutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SelectionStruct
    nIniPos As Integer
    nIniLine As Integer
    nFinPos As Integer
    nFinLine As Integer
    nCurPos As Integer
    nCurLine As Integer
    pctDest As PictureBox
End Type

Private WithEvents mBuff As cOutBuff
Attribute mBuff.VB_VarHelpID = -1
Private mBuffers As Collection
Private mAutoScroll As Boolean
Private mDefaultColor As Long

Private mIgnoreScroll As Boolean

'last printed text
Private mPrinted As String      'for normal output
Private mPrintedBack As String  'for scrollback output

'selection variables
Private mSelect As Boolean
Private mSel As SelectionStruct

'memorizza l'ultimo colore utilizzato
Private mLastColor As Long

'se vero, il box ignora le notifiche dai buffer,
'lasciando cosi' libero il suo hDC per qualsiasi
'altra elaborazione (tipo matrix :D)
Private mCustomControl As Boolean

'variabili per lo split dello schermo
Private mSplit As Boolean
Private mSplitHeight As Integer
Private mDeltaMove As Integer 'got by SystemParametersInfo

'variabili per il movimento della linea di separazione fra schermo e scrollback
Private mDragStartPt As POINTAPI
Private mDragging As Boolean

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event BuffClosed(Index As Integer)
Public Event DrawingTime(t As Long)

Public Property Get SplitHeight() As Integer
    SplitHeight = mSplitHeight
End Property

Public Property Let SplitHeight(data As Integer)
    mSplitHeight = data
    ResizeScrollback
End Property

Private Sub ResizeScrollback()
    Dim NewHeight As Long
    Dim CharHeight As Integer

    CharHeight = pctText.TextHeight("A")
    NewHeight = (CharHeight * mSplitHeight)
    
    If pctText.Height < (NewHeight + CharHeight * 2) Then
        NewHeight = (pctText.ScaleHeight \ CharHeight)
        If NewHeight <= 5 Then
            JoinBuff
            Exit Sub
        Else
            mSplitHeight = NewHeight - 2
        End If
    End If
    pctHistory.Height = pctText.Height - (CharHeight * mSplitHeight) - pctSplit.Height
    
    pctSplit.Top = pctHistory.Top + pctHistory.Height
End Sub

Public Sub SplitBuff()
    If Not mSplit Then
        pctSplit.Visible = True
        pctHistory.Visible = True
        mSplit = True
    End If
End Sub

Public Sub JoinBuff()
    If mSplit Then
        pctHistory.Visible = False
        pctSplit.Visible = False
        mSplit = False
    End If
End Sub

'set the CustomControlFlag
Public Function SetCustomControl(ByVal Value As Boolean) As PictureBox
    Static FontSize As Integer, FontName As String

    If Not mCustomControl = Value Then
        mCustomControl = Value
        If Value Then
            FontSize = pctText.FontSize
            FontName = pctText.FontName
            Set SetCustomControl = pctText
        Else
            pctText.FontSize = FontSize
            pctText.FontName = FontName
            'SplitText
            PrintTextNormal
        End If
    End If
End Function

Public Function GetWidth() As Integer
    GetWidth = pctText.ScaleWidth \ pctText.TextWidth("A")
End Function

Public Function GetHeight() As Integer
    GetHeight = pctText.ScaleWidth \ pctText.TextHeight("A")
End Function

Public Sub SetFontName(NewName As String)
    pctText.FontName = NewName
    pctHistory.FontName = NewName
    'SplitText
    ResizeScrollback
    If mSplit Then PrintTextScrollback
    PrintTextNormal
End Sub

Public Sub SetFontSize(newSize As Integer)
    pctText.FontSize = newSize
    pctHistory.FontSize = newSize
    'SplitText
    ResizeScrollback
    If mSplit Then PrintTextScrollback
    PrintTextNormal
End Sub

Public Property Get hWnd() As Long
    hWnd = pctText.hWnd
End Property

Public Sub PageLast()
    If Scroller.Enabled Then
        Scroller.Value = Scroller.Max
    End If
End Sub

Public Sub DeltaUp()
    If Scroller.Enabled Then
        If Scroller.Value - mDeltaMove < Scroller.Min Then
            Scroller.Value = Scroller.Min
        Else
            Scroller.Value = Scroller.Value - mDeltaMove
        End If
    End If
End Sub

Public Sub DeltaDown()
    If Scroller.Enabled Then
        If Scroller.Value + mDeltaMove > Scroller.Max Then
            Scroller.Value = Scroller.Max
        Else
            Scroller.Value = Scroller.Value + mDeltaMove
        End If
    End If
End Sub

Public Sub PageUp()
    Dim MaxLines As Integer
    
    If Scroller.Enabled Then
        MaxLines = pctText.ScaleHeight / pctText.TextHeight("A") / 2
        If Scroller.Value - MaxLines < Scroller.Min Then
            Scroller.Value = Scroller.Min
        Else
            Scroller.Value = Scroller.Value - MaxLines
        End If
    End If
End Sub

Public Function PageDown() As Integer
    'ritorna 0 se il testo si e' spostato,
    '1 se si era gia' al termine del buffer
    
    Dim MaxLines As Integer
    
    If Scroller.Enabled Then
        If Scroller.Value = Scroller.Max Then PageDown = 1
    
        MaxLines = pctText.ScaleHeight / pctText.TextHeight("A") / 2
        If Scroller.Value + MaxLines > Scroller.Max Then
            Scroller.Value = Scroller.Max
        Else
            Scroller.Value = Scroller.Value + MaxLines
        End If
    End If
End Function

Public Property Let AutoScroll(data As Boolean)
    mAutoScroll = data
End Property

Private Function CleanString(ByVal src As String) As String
    Dim Final As String, Pos As Long, LastPos As Long, FinPos As Long
    
    Pos = 1
    Do
        LastPos = Pos
        Pos = InStr(Pos, src, TD)
        If Not Pos = 0 Then
            FinPos = InStr(Pos + 1, src, TD)
            If FinPos = 0 Then FinPos = Len(src)
            Final = Final & Mid$(src, LastPos, Pos - LastPos)
            Pos = FinPos + 1
            LastPos = Pos
        End If
    Loop Until Pos = 0
    Final = Final & Mid$(src, LastPos)
    CleanString = Final
End Function

Public Property Get BufferSel() As Integer
    If Not mBuff Is Nothing Then BufferSel = mBuff.Index
End Property

Public Sub BufferChange(Index As Integer)
    If Not mBuff Is Nothing Then If mBuff.Index = Index Then Exit Sub
    If Index > 0 And Index <= mBuffers.Count Then
        Set mBuff = mBuffers.Item(Index)
        mBuff.Index = Index
        imgClose.Visible = mBuff.Closeable
        'SplitText
        mLastColor = -1
        PrintTextNormal
        RedrawBuffersBar
    End If
End Sub

Public Sub BufferAdd(Buff As cOutBuff, Optional SwitchOn As Boolean = True)
    mBuffers.Add Buff
    
    If SwitchOn Or mBuffers.Count = 1 Then
        Set mBuff = Buff
        mBuff.Index = mBuffers.Count
        imgClose.Visible = mBuff.Closeable
        PrintTextNormal
        'SplitText
    End If
    
    RedrawBuffersBar
End Sub

Private Sub RedrawBuffersBar()
    Dim i As Integer, rcBox As RECT, CurWidth As Long
    Dim Bordi As cGrafica
    Dim Back As Long
    
    Back = GetSysColor(COLOR_3DFACE)
    CurWidth = 0
    rcBox.Top = 0
    rcBox.Bottom = pctText.Top
    Set Bordi = New cGrafica
    'UserControl.ForeColor = GOSRGB_BUT_BorderLight
    UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
    
    For i = 1 To mBuffers.Count
        rcBox.Left = CurWidth
        rcBox.Right = CurWidth + TextWidth(mBuffers.Item(i).Name) + 8
        Bordi.FillRectEx UserControl.hdc, rcBox.Left, rcBox.Top, _
            rcBox.Right - rcBox.Left, rcBox.Bottom - rcBox.Top, Back
        CurWidth = rcBox.Right
        Bordi.DisegnaBordi UserControl.hdc, rcBox.Left, rcBox.Top, _
            rcBox.Right - rcBox.Left, rcBox.Bottom - rcBox.Top, _
            Abs(CInt(i = mBuff.Index)), 1, 0, 0, 30
        If i = mBuff.Index Then OffsetRect rcBox, 1, 1
        DrawText UserControl.hdc, mBuffers.Item(i).Name, Len(mBuffers.Item(i).Name), _
            rcBox, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
        If i = mBuff.Index Then OffsetRect rcBox, -1, -1
    Next i

    Bordi.FillRectEx UserControl.hdc, CurWidth, 0, _
        UserControl.ScaleWidth - CurWidth, pctText.Top, Back
        'UserControl.ScaleWidth -CurWidth, pctText.Top - 2, GOSRGB_FORM_Back
    Set Bordi = Nothing
    
    UserControl.Refresh
End Sub

Public Sub BufferRemove(Index As Integer)
    Dim Change As Boolean
    
    If Index > 0 And Index <= mBuffers.Count Then
        If Index = mBuff.Index Then
            Set mBuff = Nothing
            Change = True
        End If
        mBuffers.Remove Index
        If Change And mBuffers.Count >= Index Then
            BufferChange Index
        ElseIf Change And mBuffers.Count < Index Then
            BufferChange Index - 1
        ElseIf Change And mBuffers.Count = 0 Then
            RedrawBuffersBar
        End If
    End If
End Sub

Public Sub BufferRemoveAll()
    Set mBuff = Nothing
    Set mBuffers = Nothing
    Set mBuffers = New Collection
    
    RedrawBuffersBar
End Sub

'this function prints out the text on the normal output screen
Private Sub PrintTextNormal()
    Dim t As Long
    t = GetTickCount
    
    If mBuff Is Nothing Then Exit Sub
    SetupScroller
    SplitText pctText, IIf(mSplit, mSplitHeight, 0), mBuff.Count
    
    RaiseEvent DrawingTime(GetTickCount() - t)
End Sub

'this function print out the text in the scrollback screen
Private Sub PrintTextScrollback()
    Dim MaxLines As Integer
    
    Dim t As Long
    t = GetTickCount

    MaxLines = pctHistory.ScaleHeight \ pctHistory.TextHeight("A")
    'Debug.Print "scrollback drawing. value = " & Scroller.Value & " min = " & Scroller.Min
    SplitText pctHistory, MaxLines, Scroller.Value ' - Scroller.Min
    
    RaiseEvent DrawingTime(GetTickCount() - t)
End Sub

'this function sets the min, max and value of the scroller
Private Sub SetupScroller()
    If mBuff Is Nothing Then Exit Sub
    
    If mBuff.Count = 0 Then
        pctText.Cls
        pctHistory.Cls
        Scroller.Enabled = False
        If mSplit Then JoinBuff
        mIgnoreScroll = True
            Scroller.Min = 1
            Scroller.Max = 1
        mIgnoreScroll = False
    Else
        If mBuff.Count > (pctText.ScaleHeight \ pctText.TextHeight("A")) Then
            Scroller.Enabled = True
            mIgnoreScroll = True
                Scroller.Max = mBuff.Count - mSplitHeight + 1
                Scroller.Min = (pctHistory.ScaleHeight \ pctHistory.TextHeight("A"))
                If Not mSplit Then Scroller.Value = Scroller.Max
            mIgnoreScroll = False
        Else
            Scroller.Enabled = False
            If mSplit Then JoinBuff
        End If
    End If
End Sub

Public Sub SplitText(ByRef Dest As PictureBox, ByVal MaxLines As Integer, ByVal FinalIndex As Long)
    Dim CharWidth As Integer, MaxChars As Integer ', MaxLines As Integer
    Dim TextHeight As Long, CurLineIndex As Long
    Dim CurLine As String, FinalText As String, finalLine As String
    Dim iniPos As Integer, FinPos As Integer
    Dim emptyPos As Integer
    Dim CurLineClean As String, i As Integer
    'Dim FinalIndex As Long
    Dim ToAppend As String, ToAppendClean As String


    Dim asFinal() As String, asTemp() As String
    Dim nCount As Long
    
    If mBuff Is Nothing Then
        Exit Sub
    Else
        If mBuff.Count = 0 Then Exit Sub
    End If
        
    If Dest Is Nothing Then
        CharWidth = 10
        MaxChars = 156
        If MaxLines = 0 Then MaxLines = mBuff.Count
    Else
        CharWidth = Dest.TextWidth("A")
        MaxChars = Dest.ScaleWidth \ CharWidth
        If MaxLines = 0 Then MaxLines = Dest.ScaleHeight \ Dest.TextHeight("A")
    End If
    
    '////////////dimensiona buffer finale//////////////////////
    ReDim asFinal(1 To (100 + MaxLines * 2)) As String
    '//////////////////////////////////////////////////////////
    
    TextHeight = 0
    
    CurLineIndex = FinalIndex - MaxLines
    If CurLineIndex < 1 Then CurLineIndex = 1
    
    Do
        CurLine = mBuff.Item(CurLineIndex)
        CurLineClean = CleanString(CurLine)
        
        iniPos = 1 'la ricerca dei tag deve partire dal primo carattere
        finalLine = "" 'la stringa finale all'inizio e' vuota
        
        Do Until Len(CurLineClean) <= MaxChars
            emptyPos = InStrRev(CurLineClean, " ", MaxChars)
            If emptyPos = 0 Then emptyPos = InStr(MaxChars, CurLineClean, " ")
            If emptyPos = 0 Then emptyPos = Len(CurLineClean)
            
            emptyPos = emptyPos + 1
            FinPos = 0
            '//////// cambiato 1 con emptyPos nel for //////////////////////////
            For i = 1 To Len(CurLine)
                If Not Mid$(CurLine, i, 1) = TD Then
                    FinPos = FinPos + 1
                    If FinPos = emptyPos Then
                        FinPos = i
                        Exit For
                    End If
                End If
                
                If Mid$(CurLine, i, 1) = TD Then
                    i = InStr(i + 1, CurLine, TD)
                End If
            Next i
            
            If Mid$(CurLine, FinPos - 1, 1) = " " Then
                ToAppend = Left$(CurLine, FinPos - 2)
                'finalLine = finalLine & vbCrLf & Left$(CurLine, FinPos - 2)
            Else
                ToAppend = Left$(CurLine, FinPos - 1)
                'finalLine = finalLine & vbCrLf & Left$(CurLine, FinPos - 1)
            End If
            'ToAppend = Left$(CurLine, FinPos - 1)
            
            ToAppendClean = CleanString(ToAppend)
            If Len(ToAppendClean) < MaxChars Then ToAppend = ToAppend & Space(MaxChars - Len(ToAppendClean))
            finalLine = finalLine & vbCrLf & ToAppend
            
            CurLine = Mid$(CurLine, FinPos)
            CurLineClean = Mid$(CurLineClean, emptyPos)
        Loop
        
        If Len(CurLineClean) < MaxChars Then CurLine = CurLine & Space(MaxChars - Len(CurLineClean))
        finalLine = finalLine & vbCrLf & CurLine
        
        '///////////////////salvataggio della riga corrente nel buffer finale
        If nCount = 0 Then finalLine = Mid$(finalLine, 3)
        If Len(finalLine) > 1024 Then
            'MsgBox "uh, oh! " & Len(finalLine)
            Do While Len(finalLine) > 1024
                nCount = nCount + 1
                asFinal(nCount) = Left$(finalLine, 1024)
                finalLine = Mid$(finalLine, 1025)
            Loop
            nCount = nCount + 1
            asFinal(nCount) = finalLine
        Else
            nCount = nCount + 1
            asFinal(nCount) = finalLine
        End If
        '/////////////////////////////////////////////////////////////////////
        '///////////////////vecchio metodo////////////////////////////////////
        'FinalText = FinalText & finalLine
        '/////////////////////////////////////////////////////////////////////
                
        If CurLineIndex = mBuff.Count Then
            Exit Do
        Else
            CurLineIndex = CurLineIndex + 1
        End If
    Loop Until CurLineIndex > FinalIndex
       
    'ReDim Preserve asFinal(1 To nCount) As String
    FinalText = Join(asFinal, "")
    
    'FinalText = Mid$(FinalText, 3)
    If Not Dest Is Nothing Then
        TextHeight = Dest.TextHeight(FinalText)
        PrintText FinalText, Dest.ScaleHeight - Dest.TextHeight(FinalText), Dest
    End If
End Sub

Private Sub ExecTag(sTag As String, Optional ByRef Dest As PictureBox)
    Dim Color As Long

    sTag = LCase$(sTag)
    If Left$(sTag, 3) = "rgb" Then
        Color = rgb( _
            Val(Mid$(sTag, 4, 3)), _
            Val(Mid$(sTag, 7, 3)), _
            Val(Mid$(sTag, 10, 3)))
        Dest.ForeColor = Color
        mLastColor = Color
    ElseIf Left$(sTag, 4) = "back" Then
        Color = rgb( _
            Val(Mid$(sTag, 5, 3)), _
            Val(Mid$(sTag, 8, 3)), _
            Val(Mid$(sTag, 11, 3)))
        'dest.FillColor = Color
        'dest.FillStyle = vbSolid
        SetBkColor Dest.hdc, Color
        'Debug.Print "bkcolor = " & GetBkColor(dest.hdc), Mid$(sTag, 5, 3), Mid$(sTag, 8, 3), Mid$(sTag, 11, 3)
        If Color = 0 Then
            SetBkMode Dest.hdc, TRANSPARENT
        Else
            SetBkMode Dest.hdc, OPAQUE
        End If
    ElseIf Left$(sTag, 1) = "s" Then
        Select Case Mid$(sTag, 2, 1)
            Case "n"
                Dest.FontItalic = False
                Dest.FontStrikethru = False
                Dest.FontUnderline = False
            Case "i"
                Dest.FontItalic = True
            Case "u"
                Dest.FontUnderline = True
            Case "s"
                Dest.FontStrikethru = True
        End Select
    ElseIf sTag = "url" Then
        Dest.FontUnderline = True
    ElseIf sTag = "/url" Then
        Dest.FontUnderline = False
    End If
End Sub

Private Sub PrintText(Text As String, Top As Integer, Optional ByRef Dest As PictureBox)
    Dim iniPos As Long, FinPos As Long
    Dim lenght As Long, Tag As String
    Dim LineHeight As Integer, Fill As cGrafica
    Dim Connect As cConnector
    Dim SrcTop As Long

    If Dest.hWnd = pctText.hWnd Then
        mPrinted = Text
    Else
        mPrintedBack = Text
    End If
        
    If Top > 0 Then
        Set Fill = New cGrafica
            Fill.FillRectEx Dest.hdc, 0, 0, Dest.ScaleWidth, (Top), 0
        Set Fill = Nothing
    End If
        
    Dest.CurrentX = 0
    Dest.CurrentY = Top
    Dest.FontItalic = False
    Dest.FontStrikethru = False
    Dest.FontUnderline = False

    If mLastColor = -1 Then
        Set Connect = New cConnector
            mLastColor = Connect.Palette.rgbDefault
        Set Connect = Nothing
    End If
    Dest.ForeColor = mLastColor
    
    iniPos = 1
    FinPos = 0
    lenght = Len(Text)
    Do
        iniPos = InStr(iniPos, Text, TD)
        If iniPos <> 0 Then
            Dest.Print Mid$(Text, FinPos + 1, iniPos - FinPos - 1);
            'TextOutEx dest, Mid$(Text, FinPos + 1, iniPos - FinPos - 1)
            
            FinPos = InStr(iniPos + 1, Text, TD)
            Tag = Mid$(Text, iniPos + 1, FinPos - iniPos - 1)
            'Debug.Print Tag
            ExecTag Tag, Dest
            iniPos = FinPos + 1
        Else
            Dest.Print Mid$(Text, FinPos + 1);
            'TextOutEx dest, Mid$(Text, FinPos + 1)
        End If
        
    Loop Until iniPos = 0
End Sub

Private Sub imgClose_Click()
    Dim Index As Integer
    
    If Not mBuff Is Nothing Then
        Index = mBuff.Index
        Me.BufferRemove Index
        RaiseEvent BuffClosed(Index)
    End If
End Sub

Private Sub mBuff_TextAdded(nLines As Integer)
    If mCustomControl Then Exit Sub
    
    If Not mSelect Then
        If mAutoScroll Then
            PrintTextNormal
        Else
            If mBuff.Count > 1 Then
                Scroller.Max = mBuff.Count
                Scroller.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub pctHistory_Click()
    RaiseEvent Click
End Sub

Private Sub pctHistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mSel.pctDest = pctHistory
    SelMouseDown X, Y
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pctHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelMouseMove Button, X, Y
End Sub

Private Sub pctHistory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelMouseUp
End Sub

Private Sub pctHistory_Paint()
    'SplitText False, pctHistory
    PrintTextScrollback
End Sub

Private Sub pctSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mDragging = True
        GetCursorPos mDragStartPt
    End If
End Sub

Private Sub pctSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Diff As Integer
    Dim pt As POINTAPI
    
    If mDragging And Button = 1 Then
        GetCursorPos pt
        Diff = (mDragStartPt.Y - pt.Y) \ pctText.TextHeight("A")
        'Diff = -Diff
        If Abs(Diff) >= 1 Then
            mSplitHeight = mSplitHeight + Diff
            If mSplitHeight < 2 Then
                mSplitHeight = 2
                ResizeScrollback
                Exit Sub
            Else
                ResizeScrollback
            End If
            mDragStartPt.Y = mDragStartPt.Y - Diff * pctText.TextHeight("A")
        End If
    End If
End Sub

Private Sub pctSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDragging = False
End Sub

Private Sub pctText_Click()
    RaiseEvent Click
End Sub

Private Function LineFromPos(ByVal Y As Long) As Integer
    Dim LineHeight As Integer
    LineHeight = mSel.pctDest.TextHeight("A")
    
    LineFromPos = (mSel.pctDest.ScaleHeight - Y) \ LineHeight + 1
End Function

Private Function CharFromPos(ByVal X As Long) As Integer
    Dim CharWidth As Integer
    CharWidth = mSel.pctDest.TextWidth("A")

    CharFromPos = X \ CharWidth
End Function

Private Sub SelMouseDown(ByVal X As Long, ByVal Y As Long)
    Dim nChar As Integer, nLine As Integer
    
    nChar = CharFromPos(X) '0-based character
    nLine = LineFromPos(Y) '1-based line (from the bottom)
    If nChar >= 0 And nLine > 0 Then
        mSel.nIniLine = nLine
        mSel.nIniPos = nChar
        mSel.nCurLine = nLine
        mSel.nCurPos = nChar
        mSel.nFinLine = nLine
        mSel.nFinPos = nChar
        'frmLog.Log "Line = " & nLine & ", Char = " & nChar
    Else
        mSel.nIniLine = 0
        mSel.nIniPos = 0
    End If
End Sub

Private Sub SelMouseMove(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long)
    Dim nPos As Integer, nLine As Integer
    
    If Button = 1 Then
        If mSel.nIniLine > 0 And Not mSelect Then mSelect = True
   
        If mSelect Then
            nLine = LineFromPos(Y)
            nPos = CharFromPos(X)
            If nPos > mSel.nIniPos Then nPos = nPos + 1
            'frmLog.Log "Line = " & nLine & ", Pos = " & nPos
            DrawSelRect mSel, nLine, nPos
        End If
    End If
End Sub

Private Sub pctText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mSel.pctDest = pctText
    SelMouseDown X, Y
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pctText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelMouseMove Button, X, Y
End Sub

Private Sub DrawSelRect(ByRef s As SelectionStruct, ByVal nLine As Integer, ByVal nPos As Integer)
    Dim ns As SelectionStruct
    Dim dLine As Integer 'diff in term of lines
    Dim dChar As Integer 'difference in term of characters
    
    
    Dim CharWidth As Integer, LineHeight As Integer
    Dim MaxChars As Integer
    Dim pWidth As Long, pt As POINTAPI
    Dim i As Integer
    Dim Dest As PictureBox
    
    Set Dest = mSel.pctDest
    
    dLine = s.nCurLine - nLine
    dChar = s.nCurPos - nPos
    If dLine = 0 And dChar = 0 Then Exit Sub
    
    If dLine = 0 Then
        ns.nIniLine = s.nCurLine
        ns.nFinLine = s.nCurLine
        If dChar > 0 Then
            ns.nIniPos = nPos
            ns.nFinPos = s.nCurPos
        Else
            ns.nIniPos = s.nCurPos
            ns.nFinPos = nPos
        End If
    ElseIf dLine > 0 Then
        ns.nIniLine = s.nCurLine
        ns.nFinLine = nLine
        ns.nIniPos = s.nCurPos
        ns.nFinPos = nPos
    ElseIf dLine < 0 Then
        ns.nIniLine = nLine
        ns.nFinLine = s.nCurLine
        ns.nIniPos = nPos
        ns.nFinPos = s.nCurPos
    End If
    
    dLine = Abs(dLine) + 1
    dChar = Abs(dChar)
    CharWidth = Dest.TextWidth("A")
    LineHeight = Dest.TextHeight("A")
    
    Dest.DrawMode = vbInvert
    Dest.FillStyle = vbSolid
    If dLine > 1 Then
        MaxChars = Dest.ScaleWidth / CharWidth
        'draw the first line
        GetAbsSel pt, ns.nIniLine, ns.nIniPos
        pWidth = (MaxChars - ns.nIniPos) * CharWidth
        Rectangle Dest.hdc, pt.X, pt.Y, pt.X + pWidth, pt.Y + LineHeight
        
        For i = ns.nFinLine + 1 To ns.nIniLine - 1
            GetAbsSel pt, i, 0
            Rectangle Dest.hdc, 0, pt.Y, MaxChars * CharWidth, pt.Y + LineHeight
        Next i
        
        'draw the last line
        GetAbsSel pt, ns.nFinLine, ns.nFinPos
        pWidth = pt.X
        Rectangle Dest.hdc, 0, pt.Y, pWidth, pt.Y + LineHeight
    Else
        GetAbsSel pt, ns.nIniLine, ns.nIniPos
        pWidth = dChar * CharWidth
        Rectangle Dest.hdc, pt.X, pt.Y, pt.X + pWidth, pt.Y + LineHeight
    End If
    Dest.FillStyle = vbTransparent
    Dest.DrawMode = vbCopyPen
    'frmLog.Log "dLine = " & dLine & ", dChar = " & dChar
    
    s.nCurLine = nLine
    s.nCurPos = nPos
    s.nFinLine = nLine
    s.nFinPos = nPos
    
    Set Dest = Nothing
End Sub

Private Sub GetAbsSel(ByRef Dest As POINTAPI, ByVal nLine As Integer, ByVal nPos As Integer)
    Dim CharWidth As Integer, LineHeight As Integer
    
    CharWidth = mSel.pctDest.TextWidth("A")
    LineHeight = mSel.pctDest.TextHeight("A")
    Dest.X = nPos * CharWidth
    Dest.Y = mSel.pctDest.ScaleHeight - (nLine * LineHeight)
End Sub

Private Sub SelMouseUp()
    Dim Printed As String
    Dim ns As SelectionStruct
    Dim dLine As Integer 'diff in term of lines
    Dim dChar As Integer 'difference in term of characters
    
    Dim Final As String, nChars As Integer, nLines As Integer
    Dim lines() As String, CurLine As String, i As Integer
    
    'perform copy routines
    If mSelect Then
        If mSel.pctDest.hWnd = pctText.hWnd Then
            Printed = CleanString(mPrinted)
        Else
            Printed = CleanString(mPrintedBack)
        End If
        
        dLine = mSel.nFinLine - mSel.nIniLine
        dChar = mSel.nFinPos - mSel.nIniPos
        If dLine = 0 And dChar = 0 Then
            mSelect = False
            Exit Sub
        End If
        
        With mSel
            If dLine = 0 Then
                ns.nIniLine = .nFinLine
                ns.nFinLine = .nFinLine
                If dChar > 0 Then
                    ns.nIniPos = .nIniPos
                    ns.nFinPos = .nFinPos
                Else
                    ns.nIniPos = .nFinPos
                    ns.nFinPos = .nIniPos
                End If
            ElseIf dLine > 0 Then
                ns.nIniLine = .nFinLine
                ns.nFinLine = .nIniLine
                ns.nIniPos = .nFinPos
                ns.nFinPos = .nIniPos
            ElseIf dLine < 0 Then
                ns.nIniLine = .nIniLine
                ns.nFinLine = .nFinLine
                ns.nIniPos = .nIniPos
                ns.nFinPos = .nFinPos
            End If
        End With

        ns.nIniPos = ns.nIniPos + 1
        ns.nFinPos = ns.nFinPos + 1

        lines() = Split(Printed, vbCrLf)
        nLines = UBound(lines, 1)

        dLine = Abs(dLine) + 1
        dChar = Abs(dChar)
        If dLine > 1 Then
            'add first line to final
            CurLine = GetAbsLineSel(lines, ns.nIniLine)
            If Not CurLine = "" Then
                Final = RTrim$(Mid$(CurLine, ns.nIniPos))
                'frmLog.Log Final
            End If
        
            For i = ns.nIniLine - 1 To ns.nFinLine + 1 Step -1
                CurLine = GetAbsLineSel(lines, i)
                If Not CurLine = "" Then
                    If Not Final = "" Then Final = Final & vbCrLf
                    Final = Final & RTrim$(CurLine)
                    'frmLog.Log CurLine
                End If
            Next i
        
            'add last line to final
            CurLine = GetAbsLineSel(lines, ns.nFinLine)
            If Not CurLine = "" Then
                If Not Final = "" Then Final = Final & vbCrLf
                CurLine = Left$(CurLine, ns.nFinPos - 1)
                Final = Final & RTrim$(CurLine)
                'frmLog.Log CurLine
            End If
        Else
            nChars = dChar
            Final = GetAbsLineSel(lines, ns.nIniLine)
            If Not Final = "" Then
                Final = Mid$(Final, ns.nIniPos, dChar)
                'frmLog.Log Final
            End If
        End If

        If Not Final = "" Then
            'create a menu to copy the text
            ManageSelection Final
        End If

        mSelect = False

        'SplitText
        If mSel.pctDest.hWnd = pctText.hWnd Then
            PrintTextNormal
        Else
            PrintTextScrollback
        End If
            
        Set mSel.pctDest = Nothing
    End If
End Sub

Private Sub pctText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelMouseUp
End Sub

Private Sub ManageSelection(ByRef sel As String)
    Dim hMenu As Long, pt As POINTAPI, fake As RECT
    Dim rtn As Long
    Dim Email As Collection, Url As Collection
    Dim i As Integer
    
    Dim nPos As Integer
    Dim Temp As String
    Dim Connect As cConnector
    Dim Visit As String, WriteTo As String
    
    Set Url = New Collection
    Set Email = New Collection
    
    nPos = 1
    Temp = FindWord(sel, nPos, "@", ".")
    Do Until Temp = ""
        On Error Resume Next
        If Not Trim$(Temp) = "" Then Email.Add Temp, Temp
        
'        If Err.Number = 0 Then MsgBox Temp
        
        On Error GoTo 0
        Temp = FindWord(sel, nPos, "@", ".")
    Loop
    
    nPos = 1
    Temp = FindWord(sel, nPos, "http://", ".")
    Do Until Temp = ""
        On Error Resume Next
        If Not Trim$(Temp) = "" Then Url.Add Temp, Temp
        
'        If Err.Number = 0 Then MsgBox Temp
        
        On Error GoTo 0
        Temp = FindWord(sel, nPos, "http://", ".")
    Loop
    
    nPos = 1
    Temp = FindWord(sel, nPos, "www", ".")
    Do Until Temp = ""
        On Error Resume Next
        If Not Trim$(Temp) = "" Then Url.Add Temp, Temp
        
'        If Err.Number = 0 Then MsgBox Temp
        
        On Error GoTo 0
        Temp = FindWord(sel, nPos, "www", ".")
    Loop
    
    
    'this loads language information
    Set Connect = New cConnector
    Visit = Connect.Lang("main", "Visit")
    WriteTo = Connect.Lang("main", "WriteTo")
    
    hMenu = CreatePopupMenu()
    Call AppendMenu(hMenu, MF_STRING Or MF_ENABLED, 100, Connect.Lang("main", "Copy"))
    Set Connect = Nothing
    
    If Email.Count > 0 Then Call AppendMenu(hMenu, MF_SEPARATOR, 0, "")
    
    For i = 1 To Email.Count
        AppendMenu hMenu, MF_STRING Or MF_ENABLED, 200 + i, WriteTo & " " & CStr(Email.Item(i))
    Next i
    
    If Url.Count > 0 Then Call AppendMenu(hMenu, MF_SEPARATOR, 0, "")
    
    For i = 1 To Url.Count
        AppendMenu hMenu, MF_STRING Or MF_ENABLED, 100 + i, Visit & " " & CStr(Url.Item(i))
    Next i
    
    GetCursorPos pt
    rtn = TrackPopupMenu(hMenu, TPM_RETURNCMD, pt.X, pt.Y, 0, UserControl.hWnd, fake)
    'MsgBox rtn
    Select Case rtn
        Case 100 'perform copy
            Clipboard.Clear
            Clipboard.SetText sel
        Case Is > 200 'send e-mail
            ShellExecute UserControl.hWnd, vbNullString, "mailto:" & Email.Item(rtn - 200), vbNullString, vbNullString, vbNormalFocus
        Case Is > 100 'visit url
            ShellExecute UserControl.hWnd, vbNullString, Url.Item(rtn - 100), vbNullString, vbNullString, vbNormalFocus
    End Select
    DestroyMenu hMenu
            
    Set Url = Nothing
    Set Email = Nothing
End Sub

Private Function FindWord(ByRef sSrc As String, ByRef nStart As Integer, _
                          ByVal sMatch As String, Optional ByVal sMatch2 As String) As String
    
    Dim nPos As Integer, nPos2 As Integer, nIniPos As Integer
    Dim nEndPos As Integer
    
    nPos = InStr(nStart, sSrc, sMatch, vbTextCompare)
    If nPos = 0 Then
        'first match not found
        FindWord = ""
    Else
        nIniPos = InStrRev(sSrc, " ", nPos) + 1
        nPos2 = InStrRev(sSrc, vbCrLf, nPos) + 2
        If nPos2 > nIniPos And nPos2 > 2 Then nIniPos = nPos2
        'If nIniPos = 0 Then nIniPos = 1
        
        nEndPos = InStr(nPos, sSrc, " ")
        nPos2 = InStr(nPos, sSrc, vbCrLf)
        If nPos2 < nEndPos And nPos2 > 0 Or nEndPos = 0 Then nEndPos = nPos2
        If nEndPos = 0 Then nEndPos = Len(sSrc) + 1
        
        FindWord = Mid$(sSrc, nIniPos, nEndPos - nIniPos)
        If InStr(1, FindWord, sMatch2) = 0 Then
            'matches the first string but not the second one
            FindWord = " "
        End If
        
        nStart = nEndPos
    End If
End Function

Private Function GetAbsLineSel(lines() As String, ByVal nLine As Integer) As String
    Dim nLines As Integer
    
    nLines = UBound(lines, 1) + 1
    If (nLines - nLine) >= 0 And nLine > 0 Then
        GetAbsLineSel = lines(nLines - nLine)
    Else
        GetAbsLineSel = ""
    End If
End Function

Private Sub pctText_Paint()
    'If Not mCustomControl Then SplitText
    If Not mCustomControl Then PrintTextNormal
End Sub

Private Sub DoScroll()
    'If Not mIgnoreScroll Then
    '    SplitText True
    'End If
    If Not mIgnoreScroll Then
        If Scroller.Value < Scroller.Max Then
            SplitBuff
        ElseIf Scroller.Value = Scroller.Max Then
            JoinBuff
        End If
        
        'If mSplit Then SplitText True, pctHistory
        If mSplit Then PrintTextScrollback
    End If
End Sub

Private Sub Scroller_Change()
    DoScroll
End Sub

Private Sub Scroller_GotFocus()
    RaiseEvent MouseDown(1, 0, 0, 0)
End Sub

Private Sub Scroller_Scroll()
    DoScroll
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Set mBuffers = New Collection
    mAutoScroll = True
    
    mIgnoreScroll = True
    Scroller.Min = 1
    Scroller.Max = 1
    Scroller.Value = 1
    mIgnoreScroll = False
    mSplitHeight = 20
        
    mLastColor = -1
End Sub

Private Function BuffersBarHitTest(ByVal X As Long) As Integer
    Dim i As Integer, rcBox As RECT, CurWidth As Long

    CurWidth = 0
    
    For i = 1 To mBuffers.Count
        rcBox.Left = CurWidth
        rcBox.Right = CurWidth + TextWidth(mBuffers.Item(i).Name) + 8
        If X >= rcBox.Left And X <= rcBox.Right Then
            BuffersBarHitTest = i
            Exit For
        End If
        CurWidth = rcBox.Right
    Next i
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < pctText.Top - 2 Then
        'e' stato cliccato nell'area dei buffers
        BufferChange BuffersBarHitTest(X)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim Delta As Integer
    
    If Ambient.UserMode Then
        Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, Delta, 0)
        mDeltaMove = Delta
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    pctText.Width = UserControl.ScaleWidth - pctText.Left - Scroller.Width ' - 2
    pctText.Height = UserControl.ScaleHeight - pctText.Top ' - 2
    pctHistory.Width = pctText.Width
    pctSplit.Width = pctText.Width
    
    ResizeScrollback
    
    Scroller.Top = pctText.Top
    Scroller.Left = pctText.Left + pctText.Width
    Scroller.Height = pctText.Height
    
    imgClose.Left = UserControl.ScaleWidth - imgClose.Width - 1
    
    RedrawBuffersBar
    
    PrintTextNormal
    If mSplit Then PrintTextScrollback
End Sub

Private Sub UserControl_Terminate()
    Set mBuffers = Nothing
End Sub
