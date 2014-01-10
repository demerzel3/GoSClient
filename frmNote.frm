VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNote 
   Caption         =   "Gestione delle note"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "frmNote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lvwElenco 
      Height          =   2490
      Left            =   75
      TabIndex        =   5
      Top             =   450
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cboVisual 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   75
      Width           =   2415
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nuova"
      Height          =   315
      Left            =   2175
      TabIndex        =   2
      Top             =   75
      Width           =   1140
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Scarica intestazioni"
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   2040
   End
   Begin VB.TextBox txtNota 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3015
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3000
      Width           =   9315
   End
   Begin VB.Label Label1 
      Caption         =   "Visualizza"
      Height          =   240
      Left            =   3975
      TabIndex        =   3
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents mParser As cParser
Private WithEvents mFine As cFinestra
Attribute mFine.VB_VarHelpID = -1
'Private mState As Integer
Private mSel As Integer
Private mDownloading As Boolean

'robaccia ripigliata dal parser

'var per vedere se la ricezione e' avvenuta tutta insieme o meno
Private mRicez As Boolean
Private mNota As String
Private mInitState As Boolean

'intestazioni
Private mIntest As Collection
Private mEndIntes As Boolean
Private mInibisci As Boolean
Private mCount As Integer

Private mPagina As Integer
Private mPausa As Integer

Private Enum ActiveMode
    pause = -1 'nessun parsering
    
    none = 0 'ricezione e parsering normale
    'equip = 1 'ricezione dell'equipaggiamento
    'inve = 2 'ricezione dell'inventario
    intest = 4 'ricezione delle intestazioni delle note
    attiva_intest = 3 'attivazione della ricezione delle intestazioni
    
    Nota = 6 'ricezione della nota (continua fino a che si spezza la ricezione)
    attiva_nota = 5 'attivazione della ricezione della nota
    
    pagina = 7 'ricezione lunghezza pagina
    
    'mapper = 8 'ricezione della risposta al cambio di direzione
End Enum

Private Const CONTINUE_STRING As String = "[Batti (C)ontinua, (R)ipeti, (I)ndietro, (E)sci, o INVIO]:"

Private mMode As ActiveMode

Private Sub ProcessaPausa()
    'If mPausa = -1 Then
    '    If Me.RicezState <> mInitState Then
    '        mMode = none
    '    End If
    'Else
        mPausa = mPausa - 1
        If mPausa = 0 Then mMode = none
    'End If
End Sub

Private Sub ProcessaPagina(pLine As String)
    Dim Cercato As String
    
    Cercato = "Lunghezza pagina impostata a "
    If InStr(1, pLine, Cercato, vbTextCompare) Then
        mPagina = Val(Mid$(pLine, InStr(1, pLine, Cercato) + Len(Cercato), 2))
        mMode = none
        'Pausa 0, True
    End If
End Sub

Private Sub Pausa(Lunghezza As Integer, Optional FineRicez As Boolean = False)
    'If Not FineRicez Then
        mPausa = Lunghezza
    'Else
    '    mPausa = -1
    '    mInitState = Me.RicezState
    'End If
    mMode = pause
    
    Do Until mMode = none
        DoEvents
    Loop
End Sub

Private Function CambiaPagina(Lunghezza As Integer, Optional RetrVecchia As Boolean = True) As Integer
    Dim Connect As cConnector

    Set Connect = New cConnector
    If RetrVecchia Then
        mMode = pagina
        'RaiseEvent Send("Info" & vbCrLf)
        Connect.Envi.sendInput "info " & vbCrLf
        Do Until mMode = none
            DoEvents
        Loop
        CambiaPagina = mPagina
    End If
    
    'RaiseEvent Send("pausa " & Lunghezza & vbCrLf)
    Connect.Envi.sendInput "pausa " & Lunghezza & vbCrLf
    Set Connect = Nothing
    Pausa 1
End Function

Private Sub ScaricaIntestazioni(ByRef list As Collection)
    Dim Connect As cConnector
    
    mEndIntes = False
    Set mIntest = list
    Set Connect = New cConnector
    Connect.Envi.sendInput vbCrLf & "nota lista" & vbCrLf
    Set Connect = Nothing
    mMode = attiva_intest
    
    Do Until mEndIntes
        DoEvents
        If mMode <> intest And mMode <> attiva_intest Then Exit Do
    Loop
    
    Set mIntest = Nothing
End Sub

Private Sub cboVisual_Click()
    If Not mDownloading Then
        Select Case cboVisual.ListIndex
            Case 0 'tutte
                Carica
            Case 1 'nuove
                Carica False, True
            Case 2 'private
                Carica True
            Case 3 'inviate da..
                Carica False, False, True
        End Select
    End If
End Sub

Private Sub cmdDownload_Click()
    Dim Connect As cConnector
    Dim Lista As Collection
    Dim i As Integer, Subj As String, From As String
    Dim Nuovo As Boolean, Privato As Boolean, id As String
    
    Set Connect = New cConnector
    If Connect.Envi.ConnState = sckConnected Then
        lvwElenco.ListItems.Clear
        
        mDownloading = True
            cboVisual.ListIndex = 0
        mDownloading = False
        
        Set Lista = New Collection
        Me.Enabled = False
        'ScaricaIntestazioni Lista
        ScaricaIntestazioni Lista
        For i = 1 To Lista.Count
            id = Trim$(Mid$(Lista(i), 1, InStr(1, Lista(i), "]")))
            Subj = Trim$(Mid$(Lista(i), InStr(1, Lista(i), ":") + 1))
            If Not InStr(1, Lista(i), "]") + 1 > Len(Lista(i)) Then From = Trim$(Mid$(Lista(i), InStr(1, Lista(i), "]") + 1, InStr(1, Lista(i), ":") - InStr(1, Lista(i), "]") - 1))
            If Mid$(Subj, Len(Subj), 1) = Chr(10) Then Subj = Mid$(Subj, 1, Len(Subj) - 1)
            If InStr(1, id, "N") <> 0 Then Nuovo = True Else Nuovo = False
            If InStr(1, id, "P") <> 0 Then Privato = True Else Privato = False
            AggiungiNota i, From, Subj, Nuovo, Privato
        Next i
        Me.Enabled = True
        Set Lista = Nothing
        
        Salva
    Else
        MsgBox "Per scaricare le intestazioni e' necessario connettersi"
    End If
    Set Connect = Nothing
End Sub

Private Sub cmdNew_Click()
    Dim Nota As frmNota
    Dim Connect As cConnector
    
    Set Connect = New cConnector
    'If Connect.Envi.ConnState = sckConnected Then
        Set Nota = New frmNota
        Load Nota
        'Nota.Init Me
        Nota.Show
        Set Nota = Nothing
    'Else
    '    MsgBox "Per scrivere una nota e' necessario connettersi"
    'End If
    Set Connect = Nothing
End Sub

Private Sub mFine_envNotify(uMsg As Long)
    If uMsg = ENVM_ENDREC Then
        mRicez = Not mRicez
    End If
End Sub

Private Sub mFine_envOutput(data As String, OutType As Integer)
    Dim Line As String

    If OutType <> TOUT_CLEAN Then Exit Sub
    
    data = Trim$(data)
    Line = LCase$(data)
    Select Case mMode
        'Case mapper
        '    ProcessaMapper data
        Case pagina
            ProcessaPagina data
        Case pause
            ProcessaPausa
        'Case equip
        '    ProcessaEquip data
        Case attiva_nota
            If InStr(1, data, "Non c'e' quel messaggio", vbTextCompare) <> 0 Then
                mMode = none
                mNota = ""
                Exit Sub
            End If
    
            If InStr(1, data, "[", vbTextCompare) <> 0 Then
                mMode = Nota
                ProcessaNota Mid$(data, InStr(1, data, "["))
            End If
        Case Nota
            ProcessaNota data
        Case attiva_intest
            If InStr(1, data, "[", vbTextCompare) <> 0 Then
                mMode = intest
                ProcessaIntest Mid$(data, InStr(1, data, "["))
            End If
        Case intest
            ProcessaIntest data
    End Select
End Sub

Private Sub ProcessaIntest(pLine As String)
    Dim ContString As String, Connect As cConnector

    ContString = Mid$(pLine, 1, 58)
    If Mid$(pLine, 1, 1) = "[" And Mid$(pLine, 7, 1) = "]" Then
        mInibisci = False
        mCount = 0
        mIntest.Add Trim$(Mid$(pLine, InStr(1, pLine, "[")))
    Else
        If Not ContString = CONTINUE_STRING Then
            mCount = mCount + 1
            If mCount = 2 Then mMode = none
        ElseIf ContString = CONTINUE_STRING Then
            'If Not mInibisci Then RaiseEvent Send("C" & vbCrLf)
            If Not mInibisci Then
                Set Connect = New cConnector
                    Connect.Envi.sendInput "C" & vbCrLf
                Set Connect = Nothing
            End If
            mInibisci = True
            mCount = 0
        End If
    End If
End Sub

Private Sub ProcessaNota(pLine As String)
    If mIntest.Count = 0 Then
        mInitState = mRicez
        'mNota = pLine
        mIntest.Add pLine
    Else
        If mInitState = mRicez Then
            'mNota = mNota & vbCrLf & pLine
            mIntest.Add pLine
        Else
            mMode = none
        End If
    End If
End Sub

Private Sub Salva()
    Dim Config As cConnector
    Dim i As Integer, Nuovo As Boolean, Privato As Boolean

    Set Config = New cConnector
    With Config
        '.CaricaFile "config.ini"
        .SetConfig "nota_{count}", lvwElenco.ListItems.Count
        For i = 1 To lvwElenco.ListItems.Count
            .SetConfig "nota_{id<" & i & ">}", lvwElenco.ListItems(i).SubItems(2)
            .SetConfig "nota_{from<" & i & ">}", lvwElenco.ListItems(i).SubItems(1)
            .SetConfig "nota_{subject<" & i & ">}", lvwElenco.ListItems(i).Text
            If lvwElenco.ListItems(i).SubItems(3) = "*" Then Nuovo = True Else Nuovo = False
            .SetBoolConfig "nota_{new<" & i & ">}", Nuovo
            If lvwElenco.ListItems(i).SubItems(4) = "P" Then Privato = True Else Privato = False
            .SetBoolConfig "nota_{private<" & i & ">}", Privato
        Next i
        .SaveConfig
    End With
    Set Config = Nothing
End Sub

Private Sub Carica(Optional Privato As Boolean = False, Optional Nuove As Boolean = False, Optional Inviate As Boolean = False)
    Dim Config As cConnector
    Dim i As Integer
    Dim Count As Integer, Continua As Boolean
    Dim Nick As String

    lvwElenco.ListItems.Clear
    
    If Inviate = True Then
        Nick = InputBox("Inserisci un nickname")
        If Nick = "" Then
            mDownloading = True
            'mnbVisual.Selected = 1
            cboVisual.ListIndex = 0
            mDownloading = False
            Inviate = False
        Else
            Nick = LCase$(Nick)
        End If
    End If
        
    Set Config = New cConnector
    With Config
        Count = CInt(Val(.GetConfig("nota_{count}", 0)))
        For i = 1 To Count
            Continua = False
            If Privato = False And Inviate = False And Nuove = False Then
                Continua = True
            ElseIf Privato = True Then
                If .GetBoolConfig("nota_{private<" & i & ">}") Then Continua = True
            ElseIf Nuove = True Then
                If .GetBoolConfig("nota_{new<" & i & ">}") Then Continua = True
            ElseIf Inviate = True Then
                If LCase$(.GetConfig("nota_{from<" & i & ">}", "")) = Nick Then Continua = True
            End If
            
            If Continua Then
                AggiungiNota .GetConfig("nota_{id<" & i & ">}"), _
                             .GetConfig("nota_{from<" & i & ">}", ""), _
                             .GetConfig("nota_{subject<" & i & ">}", ""), _
                             .GetBoolConfig("nota_{new<" & i & ">}"), _
                             .GetBoolConfig("nota_{private<" & i & ">}", 0)
            End If
        Next i
    End With
    Set Config = Nothing
End Sub

Private Sub AggiungiNota(id As Integer, Mittente As String, Soggetto As String, Optional Nuovo As Boolean = False, Optional Privato As Boolean = False)
    Dim Item As ComctlLib.ListItem

    Set Item = lvwElenco.ListItems.Add(, , Soggetto)
    'lvwElenco.AddInfo Trim$(CStr(id))
    'lvwElenco.ModInfo lvwElenco.Count, 2, Mittente
    'lvwElenco.ModInfo lvwElenco.Count, 3, Soggetto
        Item.SubItems(1) = Mittente
        Item.SubItems(2) = Trim$(CStr(id))
    'If Nuovo Then lvwElenco.ModInfo lvwElenco.Count, 4, "*"
    'If Privato Then lvwElenco.ModInfo lvwElenco.Count, 5, "P"
        If Nuovo Then Item.SubItems(3) = "*"
        If Privato Then Item.SubItems(4) = "P"
    Set Item = Nothing
End Sub

Private Sub Form_Load()
    Dim i As Integer, Width As Long
    
    'mnbDown.Add "cmdNew", "Nuova"
    'mnbDown.Add "cmdDownload", "Scarica intestazioni"
    
    'mnbVisual.Add "cmdAll", "Tutte"
    'mnbVisual.Add "cmdNew", "Nuove"
    'mnbVisual.Add "cmdPrivate", "Private"
    'mnbVisual.Add "cmdSend", "Inviate da.."
    cboVisual.AddItem "Tutte"
    cboVisual.AddItem "Nuove"
    cboVisual.AddItem "Private"
    cboVisual.AddItem "Inviate da..."
    
    Width = lvwElenco.Width
    With lvwElenco.ColumnHeaders
        .Add , "Subject", "Oggetto", Width * 70 / 100
        .Add , "From", "Mittente", Width * 19 / 100
        .Add , "ID", "N°", Width * 5 / 100
        .Add , "New", "*", Width * 3 / 100
        .Add , "Private", "P", Width * 3 / 100
    End With
    Carica
    
    Set mFine = New cFinestra
    mFine.Init Me, WINREC_OUTPUT

    cboVisual.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Set mParser = Nothing
    mFine.UnReg
    Set mFine = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtNota.Width = Me.ScaleWidth - txtNota.Left - 5
    txtNota.Height = Me.ScaleHeight - txtNota.Top - 5
    lvwElenco.Width = Me.ScaleWidth - lvwElenco.Left - 5
End Sub

Private Sub lvwElenco_DblClick()
    Dim Testo As String, Index As Integer

    'MsgBox Index
    Index = lvwElenco.SelectedItem.Index
    Testo = ScaricaNota(Index)
    Select Case Testo
        Case "retr_subj_error"
            MsgBox "Il corpo della nota ed il suo oggetto non corrispondono" & vbCrLf & _
                   "La nota potrebbe essere stata spostata o rimossa." & vbCrLf & _
                   "Scaricare nuovamente le intestazioni per correggere."
        Case "not_found"
            MsgBox "Impossibile trovare la nota, potrebbe essere spostata o rimossa." & vbCrLf & _
                   "Scaricare nuovamente le intestazioni per correggere."
        Case Is <> ""
            'MsgBox Testo
            txtNota.Text = Trim$(Testo)
    End Select
End Sub

Private Sub DownloadNota(Index As Integer, pNota As Collection)
    Dim Connect As cConnector

    CambiaPagina 60, True
    Set Connect = New cConnector
    Connect.Envi.sendInput ("nota leggi " & Index & vbCrLf)
    Set Connect = Nothing
    mMode = attiva_nota
    mNota = ""
    Set mIntest = pNota
    
    Do Until mMode = none
        DoEvents
        If mMode <> Nota And mMode <> attiva_nota Then Exit Do
    Loop
    
    CambiaPagina mPagina, False
    
    Set mIntest = Nothing
End Sub

Private Function ScaricaNota(Index As Integer) As String
    Dim Config As cConnector
    Dim Subj As String, Corpo As String, RetrSubj As String
    Dim Nota As Collection ', Connect As cConnector

    'MsgBox Asc(Mid$(lvwElenco.GetInfo(Index, 3), Len(lvwElenco.GetInfo(Index, 3)), 1))
    Subj = lvwElenco.ListItems(Index).SubItems(1) & ": " & lvwElenco.ListItems(Index).Text
    Subj = Replace(Subj, " ", "+")
    'MsgBox Subj
    Set Config = New cConnector
    With Config
        '.CaricaFile "config.ini"
        Corpo = .GetConfig("nota_" & Subj, "")
        If Corpo = "" Then
            If Config.Envi.ConnState = sckConnected Then
                Set Nota = New Collection
                'Corpo = mParser.ScaricaNota(Index)
                
                DownloadNota Val(lvwElenco.ListItems(Index).SubItems(2)), Nota
                
                Corpo = RetrCorpo(Nota)
                Set Nota = Nothing
                'MsgBox "retrsubj result = " & RetrSubject(Corpo)
                If Corpo = "" Then
                    Corpo = "not_found"
                Else
                    RetrSubj = Replace(RetrSubject(Corpo), " ", "+")
                    If RetrSubj <> Subj Then
                        .SetConfig "nota_" & RetrSubj, Replace(Corpo, vbCrLf, "§")
                        Corpo = "retr_subj_error"
                    Else
                        .SetConfig "nota_" & Subj, Replace(Corpo, vbCrLf, "§")
                    End If
                    .SaveConfig
                End If
            Else
                MsgBox "Impossibile recuperare il testo della nota"
            End If
        Else
            Corpo = Replace(Corpo, "§", vbCrLf)
        End If
    End With
    Set Config = Nothing
    
    ScaricaNota = Corpo
End Function

Private Function RetrCorpo(Nota As Collection) As String
    Dim i As Integer

    For i = 1 To Nota.Count
        If i = 1 Then
            RetrCorpo = Nota.Item(i)
        Else
            RetrCorpo = RetrCorpo & vbCrLf & Nota.Item(i)
        End If
    Next i
End Function

Private Function RetrSubject(Corpo As String) As String
    Dim Riga As String

    Riga = Mid$(Corpo, 1, InStr(1, Corpo, vbCrLf) - 1)
    Riga = Mid$(Riga, InStr(1, Riga, "]") + 1)
    RetrSubject = Trim$(Riga)
End Function
