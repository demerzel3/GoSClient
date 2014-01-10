VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSfum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=============================================
'------cSfum versione 2.0 for Ultra IRC-------
'---------------Codice by Assa----------------
'------------assatarram3@yahoo.it-------------
'=============================================

Option Explicit

Private mSfum() As RedGreenBlue
Private mColori As Collection
Private mPosizioni As Collection
Private mNColori As Integer
Private mSfumDone As Boolean
Private mObject As Object

Private Type RedGreenBlue
    R As Single
    G As Single
    b As Single
End Type

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Declare Function SetDCPenColor Lib "gdi32" (ByVal hDC As Long, ByVal colorref As Long) As Long
'Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Const DC_PEN = 19

Private mhPen As Long

Private Sub RestoreDCPenColor(ByVal hdc As Long)
    Dim hOld As Long

    hOld = SelectObject(hdc, mhPen)
    DeleteObject hOld
    
    mhPen = 0
End Sub

Public Sub SetObject(ByRef Oggetto As Object)
    Set mObject = Oggetto
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

Public Sub StampaSfumatura(Left As Long, Top As Long, Width As Long, Height As Long, Optional Stile As Integer = 0, Optional hdc As Long = 0)
    Dim X As Integer
    Dim i As Integer
    Dim Cont As Integer
    Dim Inizio As Integer, fine As Integer

    If hdc = 0 Then hdc = mObject.hdc
        
    If Stile = 0 Then ' orizzontale
        Me.NSfumature = Width
    
        For i = Left To Left + Width - 1
            SetPixel hdc, i, Top, Me.Sfumatura(i - Left + 1)
        Next i
    
        Inizio = Top + 1
        fine = Top + Height - 1
        Cont = 1
        Do
            If Inizio + Cont > fine Then
                BitBlt hdc, Left, Inizio, Width, fine - Inizio + 1, hdc, Left, Top, SRCCOPY
                Exit Do
            Else
                BitBlt hdc, Left, Inizio, Width, Cont, hdc, Left, Top, SRCCOPY
                Inizio = Inizio + Cont
                Cont = Cont * 2
            End If
        Loop
    ElseIf Stile = 1 Then ' verticale
        Me.NSfumature = Height
    
        For i = Top To Top + Height - 1
            SetPixel hdc, Left, i, Me.Sfumatura(i - Top + 1)
        Next i
    
        Inizio = Left + 1
        fine = Left + Width - 1
        Cont = 1
        Do
            If Inizio + Cont > fine Then
                BitBlt hdc, Inizio, Top, fine - Inizio + 1, Height, hdc, Left, Top, SRCCOPY
                Exit Do
            Else
                BitBlt hdc, Inizio, Top, Cont, Height, hdc, Left, Top, SRCCOPY
                Inizio = Inizio + Cont
                Cont = Cont * 2
            End If
        Loop
    End If
End Sub

Public Property Let NSfumature(data As Integer)
    If data < mColori.Count Then
        MsgBox "Il numero di sfumature deve essere maggiore o uguale al numero dei colori", vbCritical, TypeName(Me)
    Else
        mNColori = data
        ReDim mSfum(1 To mNColori) As RedGreenBlue
        mSfumDone = False
    End If
End Property

Public Property Get NSfumature() As Integer
    NSfumature = mNColori
End Property

Public Property Get NColori() As Integer
    NColori = mColori.Count
End Property

Public Sub AggiungiColore(Colore As Long, Posizione As Single, Optional Index As Integer)
    If mColori.Count = mNColori Then
        MsgBox "Impossibile aggiungere un altro colore, il numero delle sfumature è uguale a quello dei Colori.", vbCritical
    Else
        If Index = 0 Or Index > mColori.Count Then
            mColori.Add Colore
            mPosizioni.Add Posizione
        Else
            mColori.Add Colore, , Index
            mPosizioni.Add Posizione, , Index
        End If
        mSfumDone = False
    End If
End Sub

Public Sub TogliColore(Index As Integer)
    If Not Index > mColori.Count Then
        mColori.Remove (Index)
        mPosizioni.Remove (Index)
        mSfumDone = False
    End If
End Sub

Public Sub Reset()
    Dim X As Integer
    
    For X = 1 To mColori.Count
        mColori.Remove (1)
        mPosizioni.Remove (1)
    Next X
    
    mSfumDone = False
End Sub

Private Sub CreaSfumatura()
    Dim Intervallo As Single, X As Single
    
    If Not mColori.Count = 1 Then
        'Intervallo = (mNColori / (mColori.Count - 1))
        'For x = 0 To mNColori - (Intervallo - 0.1) Step Intervallo
        '    If x = 0 Then
        '        GeneraSfumatura x + 1, x + Intervallo + 0.01, ScindiColore(mColori(CInt(x / Intervallo) + 1)), ScindiColore(mColori(CInt(x / Intervallo) + 2))
        '    Else
        '        GeneraSfumatura CInt(x), x + Intervallo + 0.01, ScindiColore(mColori(CInt(x / Intervallo) + 1)), ScindiColore(mColori(CInt(x / Intervallo) + 2))
        '    End If
        'Next x
        For X = 1 To mColori.Count - 1
            Call GeneraSfumatura(AbsPosizione(mPosizioni(X)), AbsPosizione(mPosizioni(X + 1)), ScindiColore(mColori(X)), ScindiColore(mColori(X + 1)))
        Next X
        mSfumDone = True
    End If
End Sub

Public Property Get Colore(Index As Integer) As Long
    Colore = mColori.Item(Index)
End Property

Public Property Let Colore(Index As Integer, data As Long)
    Dim Temp As Single
    
    Temp = mPosizioni(Index)
    Me.TogliColore Index
    Me.AggiungiColore data, Temp, Index
    'mColori.Item(Index) = data
    mSfumDone = False
End Property

Public Property Get Posizione(Index As Integer) As Single
    Posizione = mPosizioni.Item(Index)
End Property

Public Property Get PosizioneAss(Index As Integer) As Integer
    PosizioneAss = AbsPosizione(mPosizioni.Item(Index))
End Property

Public Property Let Posizione(Index As Integer, data As Single)
    Dim Temp As Long
    
    Temp = mColori(Index)
    Me.TogliColore Index
    Me.AggiungiColore Temp, data, Index
    mSfumDone = False
End Property

Public Property Get Sfumatura(Index As Integer) As Long
    If mSfumDone = False Then CreaSfumatura
    Sfumatura = rgb(mSfum(Index).R, mSfum(Index).G, mSfum(Index).b)
End Property

Private Sub GeneraSfumatura(IDInizio As Integer, IDFine As Integer, Colore1 As RedGreenBlue, Colore2 As RedGreenBlue)
    Dim DiffR As Single, DiffG As Single, DiffB As Single
    Dim PassR As Single, PassG As Single, PassB As Single
    Dim NumeroPassaggi As Single
    Dim X As Integer
    Dim cost As Integer
    Dim IDInizio2 As Single
    
    IDInizio2 = IDInizio - 1

    If Colore1.R > Colore2.R Then
        DiffR = Colore1.R - Colore2.R
    Else
        DiffR = Colore2.R - Colore1.R
    End If

    If Colore1.G > Colore2.G Then
        DiffG = Colore1.G - Colore2.G
    Else
        DiffG = Colore2.G - Colore1.G
    End If

    If Colore1.b > Colore2.b Then
        DiffB = Colore1.b - Colore2.b
    Else
        DiffB = Colore2.b - Colore1.b
    End If

    NumeroPassaggi = IDFine - IDInizio2

    If Not NumeroPassaggi - 1 = 0 Then
        PassR = DiffR / (NumeroPassaggi - 1)
        PassG = DiffG / (NumeroPassaggi - 1)
        PassB = DiffB / (NumeroPassaggi - 1)
    End If
    
    mSfum(IDInizio) = Colore1
    For X = 1 To NumeroPassaggi
        cost = X - 1
        If Colore1.R > Colore2.R Then
            mSfum(X + IDInizio2).R = Colore1.R - (PassR * cost)
        Else
            mSfum(X + IDInizio2).R = Colore1.R + (PassR * cost)
        End If

        If Colore1.G > Colore2.G Then
            mSfum(X + IDInizio2).G = Colore1.G - (PassG * cost)
        Else
            mSfum(X + IDInizio2).G = Colore1.G + (PassG * cost)
        End If

        If Colore1.b > Colore2.b Then
            mSfum(X + IDInizio2).b = Colore1.b - (PassB * cost)
        Else
            mSfum(X + IDInizio2).b = Colore1.b + (PassB * cost)
        End If
        
        If mSfum(X + IDInizio2).R < 0 Then mSfum(X + IDInizio2).R = 0
        If mSfum(X + IDInizio2).G < 0 Then mSfum(X + IDInizio2).G = 0
        If mSfum(X + IDInizio2).b < 0 Then mSfum(X + IDInizio2).b = 0
    Next X
    mSfum(IDFine) = Colore2
End Sub

Private Function ScindiColore(RGBLong As Long) As RedGreenBlue
    ScindiColore.R = RGBLong And &HFF
    ScindiColore.G = (RGBLong \ &H100) And &HFF
    ScindiColore.b = (RGBLong \ &H10000) And &HFF
End Function

Public Function ValoreR(RGBLong As Long) As Byte
    ValoreR = RGBLong And &HFF
End Function

Public Function ValoreG(RGBLong As Long) As Byte
    ValoreG = (RGBLong \ &H100) And &HFF
End Function

Public Function ValoreB(RGBLong As Long) As Byte
    ValoreB = (RGBLong \ &H10000) And &HFF
End Function

Private Sub Class_Initialize()
    Set mColori = New Collection
    Set mPosizioni = New Collection
    mNColori = 255
    ReDim mSfum(1 To 255) As RedGreenBlue
    mSfumDone = False
End Sub

Private Sub Class_Terminate()
    Set mColori = Nothing
    Set mPosizioni = Nothing
    Erase mSfum
End Sub

Private Function AbsPosizione(Valore As Integer) As Integer
    AbsPosizione = CInt((CLng(Valore) * mNColori) / 100)
    If AbsPosizione = 0 Then AbsPosizione = 1
End Function
