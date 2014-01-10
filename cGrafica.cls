VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGrafica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Private Const DEFAULT_COLOR As Long = 12632256

Private mBackColor As Long
Private mhPen As Long

Public Sub DrawLine(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim fake As POINTAPI
    
    MoveToEx hdc, x1, y1, fake
    LineTo hdc, x2, y2
End Sub

Public Sub BordiTxtBox(ByRef box As Object, hdc As Long)
        Me.DisegnaBordi hdc, box.Left - 2, box.Top - 2, box.Width + 4, _
                        box.Height + 4, 1, 2, 0, 0, 30
End Sub

Public Sub DrawMask(Filename As String, hdc As Long, Width As Long, Height As Long, Optional State = STATE_NORMAL)
    Dim Dc As cDeviceContext
    Dim TempDc As cDeviceContext
    Dim Mask As cBitmap
    Dim puls As cBitmap
    Dim Inizio As Integer, fine As Integer
    Dim i As Integer, Prefix As String, Name As String
    Dim Cont As Integer, sPath As String
    Dim MaskColor As Long
    Dim Connect As cConnector
    Dim hRgn As Long

    Set Connect = New cConnector
        'MaskColor = Connect.RetrInfo("maskcolor", 255)
        MaskColor = 255
        'sPath = Connect.SkinPath
    Set Connect = Nothing
    
    Set Dc = New cDeviceContext
    Set Mask = New cBitmap
    
    Dc.Create
    Select Case State
        Case STATE_NORMAL
            Prefix = ""
        Case STATE_PRESSED
            Prefix = "p_"
        Case STATE_DISABLED
            Prefix = "d_"
        Case STATE_MOUSEON
            Prefix = "m_"
        'Case STATE_GOTFOCUS
        '    Prefix = "s_"
    End Select
    Name = Prefix & Filename
    If Dir$(sPath & Name) = "" Then Name = Filename
    Mask.CreateFromBitmap sPath & Name, MASK_DIM, MASK_DIM
    Dc.SetBmp Mask
    
    Set TempDc = New cDeviceContext
    TempDc.Create
    Set puls = New cBitmap
    puls.Create Width, Height
    TempDc.SetBmp puls
    
    'inutile ma lascialo per il debug
    'TempDC.Cls puls, Picture2.BackColor
    
    'angolo in alto a sinistra
    BitBlt TempDc.hdc, 0, 0, 8, 8, Dc.hdc, 0, 0, SRCCOPY
    
    'angolo in basso a sinistra
    BitBlt TempDc.hdc, 0, Height - 8, 8, 8, Dc.hdc, 0, 9, SRCCOPY
    
    'angolo in alto a destra
    BitBlt TempDc.hdc, Width - 8, 0, 8, 8, Dc.hdc, 9, 0, SRCCOPY
    
    'angolo in basso a destra
    BitBlt TempDc.hdc, Width - 8, Height - 8, 8, 8, Dc.hdc, 9, 9, SRCCOPY
    
    
    'bordo sinistro
    Inizio = 9
    BitBlt TempDc.hdc, 0, 8, 8, 1, Dc.hdc, 0, 8, SRCCOPY
    fine = Height - 9
    Cont = 1
    Do
        If Inizio + Cont > fine Then
            BitBlt TempDc.hdc, 0, Inizio, 8, fine - Inizio + 1, TempDc.hdc, 0, 8, SRCCOPY
            Exit Do
        Else
            BitBlt TempDc.hdc, 0, Inizio, 8, Cont, TempDc.hdc, 0, 8, SRCCOPY
            Inizio = Inizio + Cont
            Cont = Cont * 2
        End If
    Loop
    
    'bordosinistro
    'Inizio = 8
    'Fine = Height - 9
    'For i = Inizio To Fine
    '    BitBlt TempDc.hdc, 0, i, 8, 1, Dc.hdc, 0, 8, SRCCOPY
    'Next i
    
    'bordo destro
    Inizio = 9
    BitBlt TempDc.hdc, Width - 8, 8, 8, 1, Dc.hdc, 9, 8, SRCCOPY
    fine = Height - 9
    Cont = 1
    Do
        If Inizio + Cont > fine Then
            BitBlt TempDc.hdc, Width - 8, Inizio, 8, fine - Inizio + 1, TempDc.hdc, Width - 8, 8, SRCCOPY
            Exit Do
        Else
            BitBlt TempDc.hdc, Width - 8, Inizio, 8, Cont, TempDc.hdc, Width - 8, 8, SRCCOPY
            Inizio = Inizio + Cont
            Cont = Cont * 2
        End If
    Loop

    'bordo destro
    'Inizio = 8
    'Fine = Height - 8
    'For i = Inizio To Fine
    '    BitBlt TempDc.hdc, Width - 8, i, 8, 1, Dc.hdc, 9, 9, SRCCOPY
    'Next i
    
    'bordo alto
    Inizio = 9
    BitBlt TempDc.hdc, 8, 0, 1, 8, Dc.hdc, 8, 0, SRCCOPY
    fine = Width - 9
    Cont = 1
    Do
        If Inizio + Cont > fine Then
            BitBlt TempDc.hdc, Inizio, 0, fine - Inizio + 1, 8, TempDc.hdc, 8, 0, SRCCOPY
            Exit Do
        Else
            BitBlt TempDc.hdc, Inizio, 0, Cont, 8, TempDc.hdc, 8, 0, SRCCOPY
            Inizio = Inizio + Cont
            Cont = Cont * 2
        End If
    Loop

    'bordo alto
    'Inizio = 8
    'Fine = Width - 9
    'For i = Inizio To Fine
    '    BitBlt TempDc.hdc, i, 0, 1, 8, Dc.hdc, 8, 0, SRCCOPY
    'Next i
        
    'bordo basso
    Inizio = 9
    BitBlt TempDc.hdc, 8, Height - 8, 1, 8, Dc.hdc, 8, 9, SRCCOPY
    fine = Width - 9
    Cont = 1
    Do
        If Inizio + Cont > fine Then
            BitBlt TempDc.hdc, Inizio, Height - 8, fine - Inizio + 1, 8, TempDc.hdc, 8, Height - 8, SRCCOPY
            Exit Do
        Else
            BitBlt TempDc.hdc, Inizio, Height - 8, Cont, 8, TempDc.hdc, 8, Height - 8, SRCCOPY
            Inizio = Inizio + Cont
            Cont = Cont * 2
        End If
    Loop
    
    'bordo basso
    'Inizio = 8
    'Fine = Width - 8
    'For i = Inizio To Fine
    '    BitBlt TempDc.hdc, i, Height - 8, 1, 8, Dc.hdc, 8, 9, SRCCOPY
    'Next i
        
    Dc.FreeBmp
    Dc.Destroy
    Set Dc = Nothing
    Mask.Destroy
    Set Mask = Nothing
    
    'hRgn = MakeRegion(TempDc.hdc, Width, Height)
    SelectObject hdc, hRgn
    
    BitBlt hdc, 0, 0, Width, Height, TempDc.hdc, 0, 0, SRCCOPY
    
    Dim hRgn2 As Long
    'hRgn2 = CreateRectRgn(0, 0, Width, Height)
    Call SelectObject(hdc, hRgn2)
    'Debug.Print "DeleteRegion " & DeleteObject(hRgn)
    'Debug.Print "DeleteRegion2 " & DeleteObject(hRgn2)
    DeleteObject (hRgn)
    DeleteObject (hRgn2)
    
    TempDc.FreeBmp
    TempDc.Destroy
    Set TempDc = Nothing
    puls.Destroy
    Set puls = Nothing
End Sub

Private Sub RestoreDCPenColor(ByVal hdc As Long)
    Dim hOld As Long

    hOld = SelectObject(hdc, mhPen)
    DeleteObject hOld
    
    mhPen = 0
End Sub

Public Sub FillRectEx(hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long, Colore As Long)
    Dim Area As RECT, hBrush As Long

    Area.Left = Left
    Area.Top = Top
    Area.Right = Left + Width
    Area.Bottom = Top + Height
    hBrush = CreateSolidBrush(Colore)
    FillRect hdc, Area, hBrush
    Call DeleteObject(hBrush)
End Sub

Private Function SetDCPenColor(ByVal hdc As Long, ByVal Color As Long) As Long
    Dim hPen As Long
    Dim hOld As Long

    hPen = CreatePen(0, 1, Color)
    hOld = SelectObject(hdc, hPen)
    If mhPen = 0 Then
        mhPen = hOld
    Else
        DeleteObject hOld
    End If
    
    SetDCPenColor = hPen
End Function

Public Sub DisegnaBordi(hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long, TipoBordo As Integer, Spessore As Integer, ColoreChiaro As Long, ColoreScuro As Long, Trasparenza As Byte, Optional TraspFissa As Boolean = False, Optional SeguiSkin As Boolean = True, Optional BackColor As Long = -1)
    Dim X As Long
    Dim varTrasp As Single
    'Dim Connect As cConnector
    'Dim Tipo2 As Integer

    'Set Connect = New cConnector
    If SeguiSkin Then
    '    ColoreChiaro = CLng(Connect.RetrInfo("LightColor"))
    '    ColoreScuro = CLng(Connect.RetrInfo("DarkColor"))
        'ColoreChiaro = GOSRGB_BUT_BorderLight
        'ColoreScuro = GOSRGB_BUT_BorderDark
        ColoreChiaro = GetSysColor(COLOR_3DHILIGHT)
        ColoreScuro = GetSysColor(COLOR_3DDKSHADOW)
        'Tipo2 = 0
    End If
    
    'If BackColor = -1 Then BackColor = Val(Connect.RetrInfo("Win_back"))
    'If BackColor = -1 Then BackColor = GOSRGB_FORM_Back
    If BackColor = -1 Then BackColor = GetSysColor(COLOR_3DFACE)
    
    'Tipo2 = CInt(Connect.RetrInfo("BorderStyle"))
    'Set Connect = Nothing
    
    'If Tipo2 = -1 Then
    '    ColoreScuro = ColoreChiaro
    '    TraspFissa = True
    '    TipoBordo = 0
    '    Spessore = 1
    'ElseIf TipoBordo = 1 Then
    '    TipoBordo = CInt(Abs(Not CBool(Tipo2)))
    'ElseIf TipoBordo = 0 Then
    '    If Tipo2 = 1 Then
    '        TipoBordo = 1
    '    Else
    '        TipoBordo = 0
    '    End If
    'End If

    mBackColor = BackColor
    
    If Spessore = 1 Then
        CreaBordo hdc, Left, Top, Width, Height, TipoBordo, ColoreChiaro, ColoreScuro, Trasparenza, Trasparenza
        RestoreDCPenColor hdc
    ElseIf Spessore > 1 Then
        varTrasp = (100 - Trasparenza) / Spessore
        For X = 0 To Spessore - 1
            If TraspFissa = 0 Then
                If TipoBordo = 0 Then
                    CreaBordo hdc, Left + X, Top + X, Width - (2 * X), Height - (2 * X), TipoBordo, ColoreChiaro, ColoreScuro, (100 - varTrasp) - (varTrasp * X), Trasparenza + (varTrasp * X) '+ ((Spessore - 1 - x) * 10)
                Else
                    CreaBordo hdc, Left + X, Top + X, Width - (2 * X), Height - (2 * X), TipoBordo, ColoreChiaro, ColoreScuro, Trasparenza + (varTrasp * X), Trasparenza + (varTrasp * X) '+ ((Spessore - 1 - x) * 10)
                End If
            Else
                CreaBordo hdc, Left + X, Top + X, Width - (2 * X), Height - (2 * X), TipoBordo, ColoreChiaro, ColoreScuro, Trasparenza, Trasparenza
            End If
        Next X
        RestoreDCPenColor hdc
    End If
End Sub

Private Sub HorizontalLine(hdc As Long, PosY As Long, Spessore As Integer, Inizio As Long, fine As Long, ByVal Colore As Long, Trasparenza As Byte)
    Dim punto As POINTAPI

    Colore = CalcolaTrasparenza(Colore, mBackColor, Trasparenza)
    Call SetDCPenColor(hdc, Colore)
    MoveToEx hdc, Inizio, PosY, punto
    LineTo hdc, fine + 1, PosY
    'Dim x As Long, Y As Long
    'For Y = 0 To Spessore - 1
    '    For x = Inizio To fine
    '        SetPixel hDC, x, PosY + Y, CalcolaTrasparenza(Colore, GetPixel(hDC, x, PosY + Y), Trasparenza)
    '    Next x
    'Next Y
End Sub

Private Sub VerticalLine(hdc As Long, PosX As Long, Spessore As Integer, Inizio As Long, fine As Long, ByVal Colore As Long, Optional Trasparenza As Byte)
    Dim punto As POINTAPI
    Dim hPen As Long
    
    Colore = CalcolaTrasparenza(Colore, mBackColor, Trasparenza)
    Call SetDCPenColor(hdc, Colore)
    MoveToEx hdc, PosX, Inizio, punto
    LineTo hdc, PosX, fine + 1
    
    'Dim x As Long, y As Long
    'For x = 0 To Spessore - 1
    '    For y = Inizio To Fine
    '        SetPixel hDC, PosX + x, y, CalcolaTrasparenza(Colore, GetPixel(hDC, PosX + x, y), Trasparenza)
    '    Next y
    'Next x
End Sub

Private Sub CreaBordo(hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long, TipoBordo As Integer, ColoreChiaro As Long, ColoreScuro As Long, TrasparenzaChiaro As Byte, TrasparenzaScuro As Byte)
    Select Case TipoBordo
        Case 0 'bordo in rilievo
            HorizontalLine hdc, Top, 1, Left, Left + Width - 1, ColoreChiaro, TrasparenzaChiaro
            HorizontalLine hdc, Top + Height - 1, 1, Left, Left + Width - 1, ColoreScuro, TrasparenzaScuro
            VerticalLine hdc, Left, 1, Top + 1, Top + Height - 2, ColoreChiaro, TrasparenzaChiaro
            VerticalLine hdc, Left + Width - 1, 1, Top + 1, Top + Height - 2, ColoreScuro, TrasparenzaScuro
        Case 1 'bordo incassato
            HorizontalLine hdc, Top, 1, Left, Left + Width - 1, ColoreScuro, TrasparenzaScuro
            HorizontalLine hdc, Top + Height - 1, 1, Left, Left + Width - 1, ColoreChiaro, TrasparenzaChiaro
            VerticalLine hdc, Left, 1, Top + 1, Top + Height - 2, ColoreScuro, TrasparenzaScuro
            VerticalLine hdc, Left + Width - 1, 1, Top + 1, Top + Height - 2, ColoreChiaro, TrasparenzaChiaro
    End Select
End Sub
