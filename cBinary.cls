VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mNFile As Integer
Private mPath As String
'identificatore vecchio, da sostituire quando cambio il mapper con
'una cosa del genere "GoSMap"
Private Const IDENT_MAP As String = "SilmarilMap"
Private Const IDENT_RUBR As String = "GoSRubrica"
Private Const IDENT_MUDS As String = "GoSMuds"
Private Const IDENT_TRIG As String = "GoSTrig"

Public Sub SaveTriggers(src As cTriggers, Path As String)
    Dim i As Integer
    
    ApriFile Path
    ScriviString IDENT_TRIG
    ScriviByte 1
    ScriviByte 0
    ScriviByte 0
    
    ScriviByte Abs(src.Imported) 'have triggers from trigger.ini already been imported?
    
    ScriviLong src.Count
    For i = 1 To src.Count
        With src.Item(i)
            ScriviString .GetText
            ScriviString .Reaction
        End With
    Next i
End Sub

Public Sub LoadTriggers(Dest As cTriggers, Path As String)
    Dim i As Integer, Count As Long
    Dim v1 As Byte, v2 As Byte, v3 As Byte
    Dim Trig As cTrigger
    
    ApriFile Path
    If LeggiString = IDENT_TRIG Then
        v1 = LeggiByte
        v2 = LeggiByte
        v3 = LeggiByte
        If v1 = 1 And v2 = 0 And v3 = 0 Then 'if file version is 1.0.0
            Dest.Imported = CBool(LeggiByte) 'have triggers from trigger.ini already been imported?
            
            Count = LeggiLong
            For i = 1 To Count
                Set Trig = New cTrigger
                    Trig.SetText LeggiString
                    Trig.Reaction = LeggiString
                    Dest.Add Trig
                Set Trig = Nothing
            Next i
        End If
    End If
End Sub

Public Sub SaveMudList(src As cMuds)
    Dim Filename As String
    Dim i As Integer
    
    '/////////////////attuale versione: 1.0.1//////////////
    Filename = App.Path & "\muds.gos"
    ApriFile Filename
    ScriviString IDENT_MUDS
    ScriviByte 1
    ScriviByte 0
    ScriviByte 1
    
    ScriviLong src.Count
    For i = 1 To src.Count
        With src.Mud(i)
            ScriviString .Name
            ScriviString .Host
            ScriviLong .Port
            ScriviString .Descr
            ScriviString .Lang
            ScriviString .Comment
        End With
    Next i
    ChiudiFile
End Sub

Public Sub LoadMudList(Dest As cMuds)
    Dim nMud As cMud, i As Integer
    Dim Filename As String, Count As Long
    Dim v1 As Byte, v2 As Byte, v3 As Byte
    Dim lPort As Long
    
    Filename = App.Path & "\muds.gos"
    ApriFile Filename
    If LeggiString = IDENT_MUDS Then
        v1 = LeggiByte
        v2 = LeggiByte
        v3 = LeggiByte
        '/////////////////////////////////////////////versione 1.0.1
        If v1 = 1 And v2 = 0 And v3 = 1 Then
            Count = LeggiLong
            Set nMud = New cMud
            With nMud
                For i = 1 To Count
                    .Name = LeggiString
                    .Host = LeggiString
                    lPort = LeggiLong
                    If lPort > 65535 Then lPort = 65535
                    .Port = lPort
                    .Descr = LeggiString
                    .Lang = LeggiString
                    .Comment = LeggiString
                    Dest.Add .Name, .Host, .Port, .Descr, .Lang, .Comment
                Next i
            End With
            Set nMud = Nothing
        '/////////////////////////////////////////////versione 1.0.0
        ElseIf v1 = 1 And v2 = 0 And v3 = 0 Then
            Count = LeggiLong
            Set nMud = New cMud
            With nMud
                For i = 1 To Count
                    .Name = LeggiString
                    .Host = LeggiString
                    .Port = LeggiInt
                    .Descr = LeggiString
                    .Lang = LeggiString
                    .Comment = LeggiString
                    Dest.Add .Name, .Host, .Port, .Descr, .Lang, .Comment
                Next i
            End With
            Set nMud = Nothing
        End If
    End If
    ChiudiFile
End Sub

Public Sub SaveRubrica(src As cRubrica)
    Dim Filename As String
    Dim i As Integer
    
    'FileName = App.Path & "\rubrica.gos"
    Filename = gMudPath & "rubrica.gos"
    ApriFile Filename
    ScriviString IDENT_RUBR
    'bytes di identificazione di versione (1.0.0)
    ScriviByte 1
    ScriviByte 0
    ScriviByte 0
    ScriviLong CLng(src.Count)
    For i = 1 To src.Count
        ScriviString src.Nick(i)
        ScriviString src.Name(i)
        ScriviString src.Email(i)
        ScriviString src.Url(i)
        ScriviString src.Note(i)
    Next i
    ChiudiFile
End Sub

Public Sub LoadRubrica(Dest As cRubrica)
    Dim Filename As String
    Dim v1 As Byte, v2 As Byte, v3 As Byte
    Dim Count As Integer, i As Integer
    Dim Nick As String
    Dim Name As String
    Dim Email As String
    Dim Url As String
    Dim Note As String
    
    'FileName = App.Path & "\rubrica.gos"
    Filename = gMudPath & "rubrica.gos"
    'Set Dest = New cRubrica
    If Dir$(Filename) = "" Then SaveRubrica Dest
    ApriFile Filename
    If LeggiString = IDENT_RUBR Then
        v1 = LeggiByte
        v2 = LeggiByte
        v3 = LeggiByte
        'caricamento file di versione 1.0.0
        If v1 = 1 And v2 = 0 And v3 = 0 Then
            Count = CInt(LeggiLong)
            For i = 1 To Count
                Nick = LeggiString
                Name = LeggiString
                Email = LeggiString
                Url = LeggiString
                Note = LeggiString
                Dest.Add Nick, Name, Email, Url, Note
            Next i
        End If
    End If
    ChiudiFile
End Sub

Public Function LoadMapOld(ByVal Nomefile As String) As cMap
    Dim map As cMap
    Dim i As Integer
    Dim Count As Integer
    Dim x1 As Integer, y1 As Integer, z1 As Integer, Capt As String
    Dim x2 As Integer, y2 As Integer, z2 As Integer
    Dim RealX As Integer, RealY As Integer
    Dim Mov As Mapper_Mov

    'Nomefile = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "Maps\" & Nomefile
    Nomefile = gMudPath & "Maps\" & Nomefile
    If Dir$(Nomefile) <> "" Then
        If mNFile <> 0 Then ChiudiFile
        ApriFile Nomefile
        
        'byte iniziale = 255, solo per la versione nuova
        'LeggiByte
        
        If LeggiString2 = IDENT_MAP Then
            Set map = New cMap
            
            Count = CInt(LeggiLong)
            For i = 1 To Count
                'x1 = LeggiInt
                'y1 = LeggiInt
                z1 = LeggiInt
                RealX = LeggiInt
                RealY = LeggiInt
                Capt = LeggiString2
                'map.Add2 Capt, x1, y1, z1, False, RealX, RealY
                map.Add2 Capt, -1, RealX, RealY, z1, False
                Capt = LeggiString2
                map.Room(map.Count).Tag = Capt
                'map.Room(map.Count).RealX = RealX
                'map.Room(map.Count).RealY = RealY
                map.Room(map.Count).TagX = LeggiInt
                map.Room(map.Count).TagY = LeggiInt
                map.Room(map.Count).TagW = LeggiInt
                map.Room(map.Count).TagH = LeggiInt
                Capt = LeggiString2
                map.Room(map.Count).Image = Capt
            Next i
        
            Count = CInt(LeggiLong)
            For i = 1 To Count
                x1 = LeggiInt
                'y1 = LeggiInt
                'z1 = LeggiInt
                x2 = LeggiInt
                Mov = LeggiInt
                'y2 = LeggiInt
                'z2 = LeggiInt
                'map.AddLink x1, y1, z1, x2, y2, z2
                map.AddLink2 x1, x2, Mov
                'Debug.Print "link from " & x1 & " to " & x2
            Next i
            
            ChiudiFile
            Set LoadMapOld = map
            Set map = Nothing
        End If
    Else
        MsgBox "Unable"
    End If
    
End Function

Public Function LoadMap(ByVal Nomefile As String) As cMap
    Dim map As cMap, Config As cIni, MapKey As Long
    Dim i As Integer
    Dim Count As Integer
    Dim x1 As Integer, y1 As Integer, z1 As Integer, Capt As String
    Dim x2 As Integer, y2 As Integer, z2 As Integer
    Dim RealX As Integer, RealY As Integer
    Dim Mov As Mapper_Mov

    'Nomefile = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "Maps\" & Nomefile
    Nomefile = gMudPath & "maps\" & Nomefile
    If Dir$(Nomefile) <> "" Then
        If mNFile <> 0 Then ChiudiFile
        ApriFile Nomefile
        
        'byte iniziale = 255, solo per la versione nuova
        LeggiByte
        
        If LeggiString = IDENT_MAP Then
            
            'controllo del numero di identificazione del pc
            MapKey = LeggiLong
            'Set Config = New cIni
            'Config.CaricaFile "config.ini"
            'If MapKey <> Config.RetrInfo("ChatColor") Then
            '    ChiudiFile
            '    Kill Nomefile
            '    Config.RemoveInfo "DockedWindow"
            '    Config.AddInfo "DockedWindow", True
            '    Config.SalvaFile
            '    Set Config = Nothing
            '    Exit Function
            'End If
            'Set Config = Nothing
            
            Set map = New cMap
            
            Count = CInt(LeggiLong)
            For i = 1 To Count
                'x1 = LeggiInt
                'y1 = LeggiInt
                z1 = LeggiInt
                RealX = LeggiInt
                RealY = LeggiInt
                Capt = LeggiString
                'map.Add2 Capt, x1, y1, z1, False, RealX, RealY
                map.Add2 Capt, -1, RealX, RealY, z1, False
                Capt = LeggiString
                map.Room(map.Count).Tag = Capt
                'map.Room(map.Count).RealX = RealX
                'map.Room(map.Count).RealY = RealY
                map.Room(map.Count).TagX = LeggiInt
                map.Room(map.Count).TagY = LeggiInt
                map.Room(map.Count).TagW = LeggiInt
                map.Room(map.Count).TagH = LeggiInt
                Capt = LeggiString
                map.Room(map.Count).Image = Capt
            Next i
        
            Count = CInt(LeggiLong)
            For i = 1 To Count
                x1 = LeggiInt
                'y1 = LeggiInt
                'z1 = LeggiInt
                x2 = LeggiInt
                Mov = LeggiInt
                'y2 = LeggiInt
                'z2 = LeggiInt
                'map.AddLink x1, y1, z1, x2, y2, z2
                map.AddLink2 x1, x2, Mov
                'Debug.Print "link from " & x1 & " to " & x2
            Next i
            
            ChiudiFile
            Set LoadMap = map
            Set map = Nothing
        End If
    Else
        MsgBox "Unable to find this file"
    End If
    
End Function

Public Sub SaveMap(ByVal Nomefile As String, map As cMap)
    Dim i As Integer
    Dim Path As String
    Dim Config As cIni
    
    'Path = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "Maps\"
    Path = gMudPath & "maps\"
    Nomefile = Path & Nomefile
    
    On Error Resume Next
    MkDir Path

    If mNFile <> 0 Then ChiudiFile
    ApriFile Nomefile
    
    'byte iniziale = 255, per riconoscere la nuova versione da quella vecchia
    ScriviByte 255
    
    ScriviString IDENT_MAP
    
    'scrittura del numero per impediare la diffusione delle mappe
    'Set Config = New cIni
    'Config.CaricaFile "config.ini"
    'ScriviLong Config.RetrInfo("ChatColor")
    'Set Config = Nothing
    
    ScriviLong 0
    
    ScriviLong map.Count
    For i = 1 To map.Count
        With map.Room(i)
            'ScriviInt .PosX
            'ScriviInt .PosY
            
            ScriviInt .PosZ
            ScriviInt .RealX
            ScriviInt .RealY
            ScriviString .Caption
            ScriviString .Tag
            ScriviInt .TagX
            ScriviInt .TagY
            ScriviInt .TagW
            ScriviInt .TagH
            ScriviString .Image
        End With
    Next i
    
    ScriviLong map.LinkCount
    For i = 1 To map.LinkCount
        With map.Link(i)
            'ScriviInt .x1
            'ScriviInt .y1
            'ScriviInt .z1
            'ScriviInt .x2
            'ScriviInt .y2
            'ScriviInt .z2
            'ScriviInt map.Search(.x1, .y1, .z1)
            'ScriviInt map.Search(.x2, .y2, .z2)
            ScriviInt .Room1
            ScriviInt .Room2
            ScriviInt .Mov
        End With
    Next i
    
    'Set map = Nothing
    ChiudiFile
End Sub

Private Function LunghezzaFile() As Long
    'LunghezzaFile = FileLen(mPath)
    LunghezzaFile = RealFileLen
End Function

Private Function RealFileLen() As Long
    Dim Pos As Long
    Dim ProvaLen As Long

    ProvaLen = FileLen(mPath)
    Pos = Seek(mNFile)
    Close mNFile
    RealFileLen = FileLen(mPath)
    Open mPath For Binary As mNFile
    Seek mNFile, Pos
    
    Debug.Print "Len = " & ProvaLen & " | RealLen = " & RealFileLen
End Function

Private Sub Shift(Inizio As Long, Spost As Long)
    Dim Buff() As Byte
    Dim Lunghezza As Long

    Lunghezza = RealFileLen - Inizio
    If Lunghezza > 0 Then
        Seek mNFile, Inizio
        Debug.Print "Shift di " & Spost & " di " & Lunghezza & " bytes"
        ReDim Buff(0 To Lunghezza)
        Get mNFile, , Buff
        Seek mNFile, Inizio + Spost
        Put mNFile, , Buff
        'If Spost < 0 Then
        '    ReDim BuffNeg(1 To Abs(Spost))
        '    Seek mNFile, Inizio + Spost + UBound(Buff, 1)
        '    Put mNFile, , BuffNeg
        'End If
        Erase Buff()
    End If
End Sub

Private Sub TornaInizio()
    Seek mNFile, 1
End Sub

Private Function PosCorrente() As Long
    PosCorrente = Seek(mNFile)
End Function

Public Sub ApriFile(Path As String)
    mPath = Path
    mNFile = FreeFile
    Open Path For Binary As mNFile
End Sub

Public Sub ChiudiFile()
    Close #mNFile
    mNFile = 0
End Sub

Private Sub SpostaAvantiByte()
    Seek mNFile, Seek(mNFile) + 1
End Sub

Private Sub SpostaAvantiString2()
    Dim LenString As Byte
    
    Get mNFile, , LenString
    'LenString = LeggiInt
    Seek mNFile, Seek(mNFile) + LenString
End Sub

Private Sub SpostaAvantiString()
    Dim LenString As Integer
    
    'Get mNFile, , LenString
    LenString = LeggiInt
    Seek mNFile, Seek(mNFile) + LenString
End Sub

Private Sub SpostaAvantiLong()
    Seek mNFile, Seek(mNFile) + 4
End Sub

Private Sub SpostaAvantiBuff()
    Dim LenBuff As Long

    LenBuff = LeggiLong
    Seek mNFile, Seek(mNFile) + LenBuff
End Sub

Public Sub Sposta(NByte As Long)
    Seek mNFile, NByte
End Sub

Private Sub ScriviBytes(Buff() As Byte)
    Put mNFile, , Buff
End Sub

Private Sub LeggiBytes(NBytes As Long, Buff() As Byte)
    ReDim Buff(1 To NBytes) As Byte
    Get mNFile, , Buff
End Sub

Public Function LeggiByte() As Byte
    Get mNFile, , LeggiByte
End Function

Private Sub ScriviByte(Dato As Byte)
    Put mNFile, , Dato
End Sub

Private Sub ScriviBuff(Buff() As Byte)
    ScriviLong UBound(Buff, 1)
    Put mNFile, , Buff()
End Sub

Private Sub LeggiBuff(Buff() As Byte)
    Dim NBytes As Long

    NBytes = LeggiLong
    If NBytes = 0 Then
        'NBytes = 1
        ReDim Buff(0) As Byte
    Else
        ReDim Buff(1 To NBytes) As Byte
        Get mNFile, , Buff()
    End If
    'ReDim Buff(0 To NBytes) As Byte
End Sub

Public Function LeggiString2() As String
    Dim LenStringa As Byte, bStringa() As Byte

    Get mNFile, , LenStringa
    If Not LenStringa = 0 Then
        ReDim bStringa(1 To LenStringa) As Byte
        Get mNFile, , bStringa
        
        LeggiString2 = UnisciStr(bStringa)
    Else
        LeggiString2 = ""
    End If
End Function

Public Function LeggiString() As String
    Dim LenStringa As Integer, bStringa() As Byte

    'Get mNFile, , LenStringa
    LenStringa = LeggiInt
    If Not LenStringa = 0 Then
        ReDim bStringa(1 To LenStringa) As Byte
        Get mNFile, , bStringa
        
        LeggiString = UnisciStr(bStringa)
    Else
        LeggiString = ""
    End If
End Function

Public Sub ScriviString2(Stringa As String)
    Dim LenStringa As Byte, bStringa() As Byte

    'Stringa = Trim$(Stringa)
    LenStringa = Len(Stringa)
    If Not LenStringa = 0 Then DividiStr Stringa, bStringa
    
    Put mNFile, , LenStringa
    'ScriviInt LenStringa
    If Not LenStringa = 0 Then Put mNFile, , bStringa
End Sub

Public Sub ScriviString(Stringa As String)
    Dim LenStringa As Integer, bStringa() As Byte

    'Stringa = Trim$(Stringa)
    LenStringa = Len(Stringa)
    If Not LenStringa = 0 Then DividiStr Stringa, bStringa
    
    'Put mNFile, , LenStringa
    ScriviInt LenStringa
    If Not LenStringa = 0 Then Put mNFile, , bStringa
End Sub

Private Sub DividiStr(Stringa As String, Buff() As Byte)
    Dim X As Integer

    ReDim Buff(1 To Len(Stringa)) As Byte
    For X = 1 To UBound(Buff, 1)
        Buff(X) = CByte(Asc(Mid$(Stringa, X, 1)))
    Next X
End Sub

Private Function UnisciStr(Buff() As Byte) As String
    Dim Stringa As String
    Dim X As Integer

    For X = 1 To UBound(Buff, 1)
        'Debug.Print Buff(x)
        Stringa = Stringa & Chr(Buff(X))
    Next X
    
    UnisciStr = Stringa
End Function

Public Function LeggiLong() As Long
    Dim bNumero(1 To 4) As Byte

    Get mNFile, , bNumero
    LeggiLong = Unisci(bNumero)
End Function

Private Function LeggiInt() As Integer
    Dim bNumero(1 To 2) As Byte
    Dim Segno As Byte
    Dim Numero As Integer
    
    Get mNFile, , Segno
    Get mNFile, , bNumero
    Numero = UnisciInt(bNumero())
    If Segno = 1 Then LeggiInt = -Numero Else LeggiInt = Numero
End Function

Public Sub ScriviLong(Numero As Long)
    Dim bNumero() As Byte

    Dividi Numero, bNumero
    Put mNFile, , bNumero
End Sub

Private Sub ScriviInt(Numero As Integer)
    Dim Segno As Byte
    Dim Buff() As Byte
    
    If Numero < 0 Then Segno = 1 Else Segno = 0
    Put mNFile, , Segno
    Numero = Abs(Numero)
    DividiInt Numero, Buff()
    Put mNFile, , Numero
End Sub

Private Sub DividiInt(Numero As Integer, Buff() As Byte)
    ReDim Buff(1 To 2) As Byte
    Buff(1) = Numero And &HFF
    Buff(2) = (Numero \ &H100&) And &HFF
    'Buff(3) = (Numero \ &H10000) And &HFF
    'Buff(4) = (Numero \ &H1000000) And &HFF
End Sub

Private Function UnisciInt(Buff() As Byte) As Integer
    Dim X As Integer

    UnisciInt = UnisciInt + Buff(1) ' And &HFF)
    UnisciInt = UnisciInt + (Buff(2) * &H100&) ' And &HFF)
    'Unisci = Unisci + (Buff(3) * &H10000) ' And &HFF)
    'Unisci = Unisci + (Buff(4) * &H1000000) ' And &HFF)
End Function

Private Sub Dividi(Numero As Long, Buff() As Byte)
    ReDim Buff(1 To 4) As Byte
    Buff(1) = Numero And &HFF
    Buff(2) = (Numero \ &H100&) And &HFF
    Buff(3) = (Numero \ &H10000) And &HFF
    Buff(4) = (Numero \ &H1000000) And &HFF
End Sub

Private Function Unisci(Buff() As Byte) As Long
    Dim X As Integer

    Unisci = Unisci + Buff(1) ' And &HFF)
    Unisci = Unisci + (Buff(2) * &H100&) ' And &HFF)
    Unisci = Unisci + (Buff(3) * &H10000) ' And &HFF)
    Unisci = Unisci + (Buff(4) * &H1000000) ' And &HFF)
End Function

Private Sub Class_Terminate()
    ChiudiFile
End Sub

