Attribute VB_Name = "modGlobals"
Option Explicit

Public Type RedGreenBlue
    R As Integer
    G As Integer
    b As Integer
End Type

Global Const TD As String * 1 = "Œ" 'TD = TagDelimiter

Global Const ESCCHAR As String = ""

'////////////// ENVIRON MESSAGES \\\\\\\\\\\\\\\\\\'
    Global Const ENVM_CONNECT           As Long = 11
    Global Const ENVM_CLOSE             As Long = 1
    Global Const ENVM_ENDREC            As Long = 2
    Global Const ENVM_PALCHANGED        As Long = 3
    Global Const ENVM_PROFILECHANGED    As Long = 4
    Global Const ENVM_CONFIGCHANGED     As Long = 5
    Global Const ENVM_SWITCHTOMUD       As Long = 6
    Global Const ENVM_SWITCHTOSTATUS    As Long = 7
    Global Const ENVM_BUTTSAVESTATE     As Long = 8
    Global Const ENVM_BUTTCHANGED       As Long = 9
    Global Const ENVM_LYTCHANGED        As Long = 10
    Global Const ENVM_SETTCHANGED       As Long = 12
    Global Const ENVM_MOUSEWHEELUP      As Long = 13
    Global Const ENVM_MOUSEWHEELDOWN    As Long = 14
'///////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\'

Global Const BUTTON_IMG As Long = 17 'dimen. delle immagini sui pulsanti

Global Const MASK_DIM As Long = 17

Global Const STATE_NORMAL As Integer = 0
Global Const STATE_PRESSED As Integer = 1
Global Const STATE_MOUSEON As Integer = 2
Global Const STATE_DISABLED As Integer = 3

Global Const ORIENT_HOR As Integer = 0
Global Const ORIENT_VERT As Integer = -900

Global Const SCMODE_MEMORY As Integer = 0
Global Const SCMODE_DISK As Integer = 1

Global Const LYTMODE_DOCK As Integer = 0
Global Const LYTMODE_MODIFY As Integer = 1

Global Const PIMD As String * 1 = "Œ" 'PIMD = Plug-Ins Message Delimiter

Global gPluginsPort As Integer
Global gMudPath As String
Global gEnvi As cEnviron

Public Sub TextOutEx(Dest As PictureBox, s As String)
    Dim i As Integer
    Dim out() As String, Count As Integer
    Dim LineHeight As Integer
    Dim X As Long, Y As Long
    
    If InStr(1, s, vbCrLf) Then
        LineHeight = Dest.TextHeight("A")
        out = Split(s, vbCrLf)
        Count = UBound(out, 1)
        X = Dest.CurrentX
        Y = Dest.CurrentY
        For i = 0 To Count
            If i <> 0 Then
                X = 0
                Y = Y + LineHeight
            End If
            TextOut Dest.hdc, X, Y, out(i), Len(out(i))
        Next i
        If Not Count = -1 Then X = X + Dest.TextWidth(out(Count))
        Dest.CurrentX = X
        Dest.CurrentY = Y
    Else
        X = Dest.CurrentX
        TextOut Dest.hdc, X, Dest.CurrentY, s, Len(s)
        Dest.CurrentX = X + Dest.TextWidth(s)
    End If
End Sub

Public Function CalcolaTrasparenza(ColoreLong As Long, MischiaLong As Long, Trasparenza As Byte) As Long
    Dim Colore As RedGreenBlue
    Dim Mischia As RedGreenBlue
    Dim sngTrasp As Single
        
    If Trasparenza >= 100 Then
        CalcolaTrasparenza = MischiaLong
    ElseIf Not Trasparenza = 0 Then
        Colore = ScindiColore(ColoreLong)
        sngTrasp = Trasparenza / 10
        'Mischia = ScindiColore(GetPixel(hdc, x, y))
        Mischia = ScindiColore(MischiaLong)
    
        Colore.R = (Colore.R * (10 - sngTrasp) + Mischia.R * sngTrasp) / 10
        Colore.G = (Colore.G * (10 - sngTrasp) + Mischia.G * sngTrasp) / 10
        Colore.b = (Colore.b * (10 - sngTrasp) + Mischia.b * sngTrasp) / 10
        CalcolaTrasparenza = rgb(Colore.R, Colore.G, Colore.b)
    ElseIf Trasparenza = 0 Then
        CalcolaTrasparenza = ColoreLong
    End If
End Function

Public Function ScindiColore(RGBLong As Long) As RedGreenBlue
    ScindiColore.R = RGBLong And &HFF
    ScindiColore.G = (RGBLong \ &H100) And &HFF
    ScindiColore.b = (RGBLong \ &H10000) And &HFF
End Function

Public Function VariaColore(RGBLong As Long, Valore As Integer) As Long
    Dim Colore As RedGreenBlue
    
    Colore = ScindiColore(RGBLong)
    Colore.R = Colore.R + Valore
    Colore.G = Colore.G + Valore
    Colore.b = Colore.b + Valore
    
    If Colore.R < 0 Then Colore.R = 0
    If Colore.G < 0 Then Colore.G = 0
    If Colore.b < 0 Then Colore.b = 0
    VariaColore = rgb(Colore.R, Colore.G, Colore.b)
End Function

Public Function TransparentBlt(ByVal hDC1 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal width1 As Long, ByVal height1 As Long, ByVal hDC2 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal width2 As Long, ByVal height2 As Long, ByVal Colore As Long) As Long
    Dim i As Long, j As Long
    Dim ColorGet As Long
    
    'On Error GoTo ripiegamento
    'Call trueTransparentBlt(hDC1, X1, Y1, width1, height1, hDC2, X2, Y2, width2, height2, Colore)

    'Exit Function
'ripiegamento:
    For i = x2 To x2 + width1
        For j = y2 To y2 + height1
            ColorGet = GetPixel(hDC2, i, j)
            If Not ColorGet = Colore Then Call SetPixel(hDC1, i - x2 + x1, j - y2 + y1, ColorGet)
        Next j
    Next i
End Function
