VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutBuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'//SCMODE_DISK//////////////////////////////////
Private mStringF As cBinary 'file delle stringhe
Private mIndexF As cBinary 'file indice
Private mFileName As String
Private mEndByte As Long
'\\SCMODE_DISK\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'//SCMODE_MEMORY////////////////////////////////
Private mLines() As String
'\\SCMODE_MEMORY\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private mItemCount As Long
Private mMode As Integer

'/////////////variabili per uOutBox e gestione esterna///////////
Private mName As String
Private mIndex As Integer
Private mCloseable As Boolean
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Event TextAdded(nLines As Integer)
Public Event BuffCleared()

Public Property Let Closeable(data As Boolean)
    mCloseable = data
End Property

Public Property Get Closeable() As Boolean
    Closeable = mCloseable
End Property

Public Property Let Index(data As Integer)
    mIndex = data
End Property

Public Property Get Index() As Integer
    Index = mIndex
End Property

Public Property Let Name(data As String)
    mName = data
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Sub AppendANSILine(data As String)
    AppendANSIText data
End Sub

Public Function MakeColorTag(Color As Long, Optional Back As Boolean = False) As String
    'Dim Code As String
    'Dim r As Byte, g As Byte, b As Byte
    '
    'If Back Then Code = "BACK" Else Code = "RGB"
    '
    'r = Color And &HFF
    'g = (Color \ &H100) And &HFF
    'b = (Color \ &H10000) And &HFF
    '
    'MakeColorTag = TD & Code & Format(r, "000") & _
                               Format(g, "000") & _
                               Format(b, "000") & TD
    
    Dim Ret As String
    Dim nLen As Long
    
    Ret = Space(16)
    nLen = gosuMakeColorTag(Ret, Color, Back)
    Ret = Left$(Ret, nLen)
    MakeColorTag = Ret
End Function

'///////////////////diagnostic routine///////////////////
Public Function TestMakeColorTag(Optional lp As Long = 100000) As Long
    Dim i As Long
    Dim t As Long
    Dim rtn As String
    
    t = GetTickCount
    
    For i = 1 To lp
        rtn = MakeColorTag(16777215)
    Next i
    
    TestMakeColorTag = (GetTickCount - t)
End Function

Private Function AssembleNextString(Color As Long, Back As Long, NextPart As String) As String
    Dim buff1 As String, buff2 As String
     
    buff1 = Space(15)
    gosuMakeColorTag buff1, Color, False
    buff1 = Left$(buff1, 14)
     
    If Back > -1 Then
        buff2 = Space(16)
        gosuMakeColorTag buff2, Back, True
        buff2 = Left$(buff2, 15)
    End If
    'Buff = buff1 & buff2
     
    'AssembleNextString = MakeColorTag(Color) & MakeColorTag(Back) & NextPart
    'AssembleNextString = MakeColorTag(Color) & NextPart
    'gosuMakeColorBackTag buff, Color, Back
    'buff = Left$(buff, 29)
    AssembleNextString = buff1 & buff2 & NextPart
End Function

Public Sub AppendANSIText(data As String)
    Dim Palette As cPalette, Connect As cConnector
    Dim iniPos As Long, FinPos As Long
    Dim Save As String, lenght As Long

    Dim i As Long, iChar As Integer, Char As String

    Dim CurColor As Long, NewBack As Long, NextPart As String

    Dim AnsiCode As String
    Static LastColor As Integer
    Static BackColor As Long

    If Not Me.Count = 0 Then
        Save = Me.Item(Me.Count)
        Me.RemoveLast
        data = Save & data
    End If

    '////////////////////////////////////////
    'la palette va presa invece dal connector
    Set Connect = New cConnector
        Set Palette = Connect.Palette
    Set Connect = Nothing
    'Set Palette = New cPalette
    'Palette.LoadColors
    '////////////////////////////////////////
    
    lenght = Len(data)
    iniPos = 1
    Do
        iniPos = InStr(iniPos, data, ESCCHAR)
        If Not iniPos = 0 Then
            'finPos = InStr(iniPos, data, "m")
            
            For i = iniPos + 1 To lenght
                Char = Mid$(data, i, 1)
                iChar = Asc(LCase$(Char))
                If iChar >= 97 And iChar <= 122 Then
                    FinPos = i
                    Exit For
                End If
            Next i
            
            If i > lenght Then FinPos = 0
            
            If FinPos <> 0 Then
                AnsiCode = Mid$(data, iniPos + 2, FinPos - iniPos - 2)
                'Debug.Print "ANSI = " & AnsiCode & Char
                Select Case Char
                    Case "m"
                        NextPart = Mid$(data, FinPos + 1)
                        CurColor = Palette.AnsiColor(AnsiCode, NewBack, LastColor)
                        If BackColor = NewBack Then NewBack = -1 Else BackColor = NewBack
                        data = Left$(data, iniPos - 1) & _
                            AssembleNextString(CurColor, NewBack, NextPart)
                            'MakeColorTag(Palette.AnsiColor(AnsiCode, BackColor, LastColor)) & _
                            'MakeColorTag(BackColor, True)
                            'Debug.Print Len(NextPart)
                        'data = data & NextPart
                        iniPos = FinPos + 1
                    Case "C"
                        data = Left$(data, iniPos - 1) & _
                            Space(AnsiCode) & _
                            Mid$(data, FinPos + 1)
                    'Case "D"
                    'Case "J"
                    Case Else
                        data = Left$(data, iniPos - 1) & _
                            Mid$(data, FinPos + 1)
                        iniPos = FinPos - (Len(AnsiCode) + 3)
                        If iniPos = 0 Then iniPos = 1
                End Select
                lenght = Len(data)
            Else
                'data = Left$(data, iniPos - 1)
                iniPos = 0
            End If
        End If
    Loop Until iniPos = 0
    
    AppendText data, False
    
    Set Palette = Nothing
End Sub

Public Sub AppendLine(data As String)
    Add data
    RaiseEvent TextAdded(1)
End Sub

Public Sub AppendText(data As String, Optional UseLastLine As Boolean = True)
    Dim lines() As String, i As Integer
    Dim Save As String, nLines As Integer

    If (Not Me.Count = 0) And UseLastLine Then
        Save = Me.Item(Me.Count)
        Me.RemoveLast
        data = Save & data
    End If
    
    lines() = Split(data, vbCrLf)
    For i = LBound(lines, 1) To UBound(lines, 1)
        'If InStr(1, lines(i), "!!MUSIC") = 0 Then
            Add lines(i)
            nLines = nLines + 1
        'End If
    Next i
    RaiseEvent TextAdded(nLines)
End Sub

Public Property Let Mode(data As Integer)
    Dim i As Long, Limit As Integer
    Dim Count As Integer

    If data = mMode Then Exit Property
    
    If data = SCMODE_DISK Then
        mMode = data
        InitDiskMode
        Count = mItemCount
        mItemCount = 0
        For i = 1 To Count
            Me.Add mLines(i)
        Next i
        EndMemoryMode
    Else
        InitMemoryMode
        Limit = UBound(mLines, 1)
        If mItemCount > Limit Then
            For i = mItemCount - Limit + 1 To mItemCount
                mLines(i - (mItemCount - Limit)) = Me.Item(i)
            Next i
            mItemCount = Limit
        Else
            For i = 1 To mItemCount
                mLines(i) = Me.Item(i)
            Next i
        End If
        mMode = data
        EndDiskMode
    End If
End Property

Private Sub InitDiskMode()
    Set mStringF = New cBinary
    Set mIndexF = New cBinary
    
    mStringF.ApriFile mFileName & ".str"
    mIndexF.ApriFile mFileName & ".idx"
    
    'fine del file stringa = 1 (inizio)
    mEndByte = 1
End Sub

Private Sub InitMemoryMode()
    ReDim mLines(1 To 3000) As String
End Sub

Public Sub Init(Mode As Integer)
    mMode = Mode
    
    If mMode = SCMODE_MEMORY Then
        InitMemoryMode
    Else
        InitDiskMode
    End If
End Sub

Public Sub RemoveLast()
    If mMode = SCMODE_DISK Then
        mEndByte = mEndByte - (Len(Me.Item(mItemCount)) + 3)
    End If
    
    mItemCount = mItemCount - 1
End Sub

Public Sub Clear()
    If mMode = SCMODE_DISK Then
        mEndByte = 1
    End If
    
    mItemCount = 0
End Sub

Public Property Get Count() As Long
    Count = mItemCount
End Property

Public Property Get Item(Index As Long) As String
    Dim Pos As Long

    If mMode = SCMODE_MEMORY Then
        Item = mLines(Index)
    Else
        mIndexF.Sposta (Index - 1) * 4 + 1
        Pos = mIndexF.LeggiLong
        
        mStringF.Sposta Pos
        Item = mStringF.LeggiString
    End If
End Property

Public Sub Add(Linea As String)
    Dim i As Integer, Limit As Integer

    If mMode = SCMODE_MEMORY Then
        mItemCount = mItemCount + 1
        Limit = UBound(mLines, 1)
        If mItemCount <= Limit Then
            mLines(mItemCount) = Linea
        Else
            For i = 1 To Limit - 250
                mLines(i) = mLines(i + 250)
            Next i
            mItemCount = mItemCount - 250
            mLines(mItemCount) = Linea
        End If
    Else
        'salvataggio della stringa su file
        mStringF.Sposta mEndByte
        mStringF.ScriviString Linea
        
        'salvataggio della posizione della stringa su index
        mIndexF.Sposta mItemCount * 4 + 1 '4 perche' sono tutti long = 4 bytes
        mIndexF.ScriviLong mEndByte
        
        'aggiornamento del byte finale del files di stringhe
        mEndByte = mEndByte + Len(Linea) + 3
        mItemCount = mItemCount + 1
    End If
End Sub

Private Sub Class_Initialize()
    mMode = SCMODE_MEMORY
    InitMemoryMode
    
    mFileName = IIf(Right$(Environ$("temp"), 1) = "\", Environ$("temp"), Environ$("temp") & "\")
    mFileName = mFileName & "gos_temp_" & Format(CStr(Day(Date)), "00") & Format(CStr(Month(Date)), "00") & CStr(Year(Date))
    Randomize
    mFileName = mFileName & Format(CInt(Rnd * 1000), "000")
    
    'attualmente 0 elementi allocati
    mItemCount = 0
End Sub

Private Sub Class_Terminate()
    If mMode = SCMODE_DISK Then
        EndDiskMode
    Else
        EndMemoryMode
    End If
End Sub

Private Sub EndDiskMode()
    If Not mStringF Is Nothing Then
        mStringF.ChiudiFile
        mIndexF.ChiudiFile
        
        Set mStringF = Nothing
        Set mIndexF = Nothing
        
        Kill mFileName & ".str"
        Kill mFileName & ".idx"
    End If
End Sub

Private Sub EndMemoryMode()
    Erase mLines()
End Sub
