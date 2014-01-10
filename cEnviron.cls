VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnviron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ProfSel As Integer
Public ProfSelFolder As String
Public Mud As cMud

Private mLang As cLang
Private mPalette As cPalette
Private mRubrica As cRubrica
Private mConfig As cIni
Private mProfConfig As cIni
Private mStatusBuff As cOutBuff
Private mNRef As Integer

Private mWin() As cFinestra

Public Function GetLang() As cLang
    Set GetLang = mLang
End Function

Public Sub LoadProfConfig()
    If Not mProfConfig Is Nothing Then
        Set mProfConfig = Nothing
    End If
    
    Set mProfConfig = New cIni
    mProfConfig.CaricaFile ProfSelFolder & "config.ini", True
End Sub

Public Sub SaveProfConfig()
    If Not mProfConfig Is Nothing Then mProfConfig.SalvaFile
End Sub

Public Property Get ProfConfig() As cIni
    Set ProfConfig = mProfConfig
End Property

Public Function GetStatusBuff() As cOutBuff
    Set GetStatusBuff = mStatusBuff
End Function

Public Sub AddToStatusErr(Text As String)
    AddToStatus TD & "RGB220000000" & TD & "error " & Text
End Sub

Public Sub AddToStatus(Text As String, Optional IgnoreConv As Boolean = False)
    If IgnoreConv Then
        mStatusBuff.AppendText Text
    Else
        mStatusBuff.AppendText TD & "RGB220220220" & TD & "#" & _
            Format(CStr(Time), "hh.mm.ss") & "# " & TD & _
            "RGB200200200" & TD & "| " & Text & vbCrLf
    End If
End Sub

Public Function GetRubrica() As cRubrica
    Set GetRubrica = mRubrica
End Function

Public Property Get ConnState() As StateConstants
    ConnState = frmBase.Winsock.State
End Property

Public Sub sendNotify(uMsg As Long)
    Dim i As Integer
    
    For i = 1 To UBound(mWin, 1)
        mWin(i).sendNotify uMsg
        If i >= UBound(mWin, 1) Then Exit For
    Next i
End Sub

Public Sub sendOutput(out As String, Optional OutType As Integer = TOUT_SOCKET)
    Dim i As Integer
    
    If Not (OutType = TOUT_STATUS Or OutType = TOUT_ERROR) Then
        For i = 1 To UBound(mWin, 1)
            If (mWin(i).WinFlags And WINREC_OUTPUT) = WINREC_OUTPUT Then
                mWin(i).sendOutput out, OutType
            End If
        Next i
    ElseIf OutType = TOUT_STATUS Then
        AddToStatus out
    ElseIf OutType = TOUT_ERROR Then
        AddToStatusErr out
    End If
End Sub

Public Sub sendInput(inp As String, Optional InType As Integer = 0)
    Dim i As Integer
    
    For i = 1 To UBound(mWin, 1)
        If (mWin(i).WinFlags And WINREC_INPUT) = WINREC_INPUT Then
            mWin(i).sendInput inp, InType
        End If
    Next i
End Sub

Public Sub RegWin(win As cFinestra)
    Dim Count As Integer
    
    Count = UBound(mWin, 1) + 1
    If Count = 1 Then
        ReDim mWin(1 To 1) As cFinestra
    Else
        ReDim Preserve mWin(1 To Count) As cFinestra
    End If
    
    Set mWin(Count) = win
    win.winID = Count
    
    Dim j As Integer
    For j = 1 To Count
        Debug.Print j, mWin(j).Caption
    Next j
End Sub

Public Sub UnRegWin(winID As Integer)
    Dim i As Integer
    Dim Count As Integer
    
    Debug.Print "UnRegwin eseguito"
    On Error Resume Next
    Count = UBound(mWin, 1)
    If winID < Count Then
        Set mWin(winID) = Nothing
        For i = winID To Count - 1
            Set mWin(i) = mWin(i + 1)
            Set mWin(i + 1) = Nothing
            mWin(i).winID = mWin(i).winID - 1
        Next i
    End If
    
    If Not Count = 1 Then
        ReDim Preserve mWin(1 To Count - 1) As cFinestra
    Else
        ReDim mWin(0) As cFinestra
    End If
    
    Count = Count - 1
    
    Dim j As Integer
    For j = 1 To Count
        Debug.Print j, mWin(j).Caption
    Next j
End Sub

Public Property Let NReferenze(data As Integer)
    mNRef = data
End Property

Public Property Get NReferenze() As Integer
    NReferenze = mNRef
End Property

Private Sub Class_Initialize()
    Dim Conf As cIni
   
    Set Conf = New cIni
        Conf.CaricaFile App.Path & "\config.ini", True
        Set mLang = New cLang
            mLang.LoadLang Conf.RetrInfo("lang", "english.lng")
    Set Conf = Nothing
    
    Set mStatusBuff = New cOutBuff

    'AddToStatus TD & "SU" & TD & _
                TD & "RGB254041013" & TD & "Ga" & _
                TD & "RGB255133044" & TD & "te" & _
                TD & "RGB254186034" & TD & "s " & _
                TD & "RGB254226014" & TD & "of" & _
                TD & "RGB229248006" & TD & " S" & _
                TD & "RGB149228026" & TD & "il" & _
                TD & "RGB069208046" & TD & "ma" & _
                TD & "RGB067187118" & TD & "ri" & _
                TD & "RGB091163212" & TD & "l " & _
                TD & "RGB077116254" & TD & "v" & App.Major & _
                TD & "RGB127000254" & TD & "." & App.Minor & vbCrLf & _
                TD & "SN" & TD & vbCrLf, True
    
    AddToStatus TD & "RGB000124212" & TD & "                       _JMMMM_." & vbCrLf & _
                TD & "RGB019134215" & TD & "       ,_MMMMMM     ,JMMMM4MMMMML       ,JMMMMMMMMMML" & vbCrLf & _
                TD & "RGB039144218" & TD & "     JMMMMMMMMF    JMMF`      4MMM.    (MMMM""""""""4MMM`" & vbCrLf & _
                TD & "RGB058154221" & TD & "   JMMM""          JMM`          4MM.   MMM" & vbCrLf & _
                TD & "RGB078164225" & TD & "  MMM`           ,MM`            4MM   `MMM_" & vbCrLf & _
                TD & "RGB098174228" & TD & " MMM ,M_,_____.  JMF             `MM)    ""MMMM_." & vbCrLf & _
                TD & "RGB117184231" & TD & "JMM  (MMMMMMMMM) MM)              MM)       ""MMMMM_" & vbCrLf & _
                TD & "RGB137194235" & TD & "MM`          4M) MM)             ,MM)          `4MMMM." & vbCrLf & _
                TD & "RGB156204238" & TD & "MM           (M) (MM             JMM ,ML          `4MMM" & vbCrLf & _
                TD & "RGB137194235" & TD & "MM           (M)  MMM.          JMMF `MM)            MM)" & vbCrLf & _
                TD & "RGB117184231" & TD & "MM.          (M)  `MMML.      _MMMF   4MML          ,MM)," & vbCrLf & _
                TD & "RGB098174228" & TD & "4MML.        JM)    4MMMMMMMMMMMM`     `MMMM_______MMMM" & vbCrLf & _
                TD & "RGB078164225" & TD & " `MMMMMMMMMMMMM)     `4MMMMMMMM""         `""MMMMMMMMMM`" & vbCrLf & _
                TD & "RGB058154221" & TD & "     `""""""""""""""" & vbCrLf, True
    
    'AddToStatus TD & "RGB255173110" & TD & "                       _JMMMM_.                            " & vbCrLf & _
                TD & "RGB254169106" & TD & "       ,_MMMMMM     ,JMMMM4MMMMML       ,JMMMMMMMMMML       " & vbCrLf & _
                TD & "RGB252166102" & TD & "     JMMMMMMMMF    JMMF`      4MMM.    (MMMM""""""""4MMM`   " & vbCrLf & _
                TD & "RGB251162098" & TD & "   JMMM""          JMM`          4MM.   MMM                 " & vbCrLf & _
                TD & "RGB249158094" & TD & "  MMM`           ,MM`            4MM   `MMM_                " & vbCrLf & _
                TD & "RGB248154090" & TD & " MMM ,M_,_____.  JMF             `MM)    ""MMMM_.           " & vbCrLf & _
                TD & "RGB246151086" & TD & "JMM  (MMMMMMMMM) MM)              MM)       ""MMMMM_        " & vbCrLf & _
                TD & "RGB245147082" & TD & "MM`          4M) MM)             ,MM)          `4MMMM.      " & vbCrLf & _
                TD & "RGB214129072" & TD & "MM           (M) (MM             JMM ,ML          `4MMM     " & vbCrLf & _
                TD & "RGB184110062" & TD & "MM           (M)  MMM.          JMMF `MM)            MM)    " & vbCrLf & _
                TD & "RGB153092051" & TD & "MM.          (M)  `MMML.      _MMMF   4MML          ,MM),   " & vbCrLf & _
                TD & "RGB122074041" & TD & "4MML.        JM)    4MMMMMMMMMMMM`     `MMMM_______MMMM     " & vbCrLf & _
                TD & "RGB092055031" & TD & " `MMMMMMMMMMMMM)     `4MMMMMMMM""         `""MMMMMMMMMM`    " & vbCrLf & _
                TD & "RGB061037020" & TD & "     `""""""""""""""                                        " & vbCrLf, True
    
    'AddToStatus vbCrLf & _
                TD & "RGB220220220" & TD & "                         GosClient                           " & vbCrLf, True
    'AddToStatus TD & "RGB000000000" & TD & _
                TD & "BACK000124212" & TD & "                         _JMMMM_.                            " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK028139217" & TD & "         ,_MMMMMM     ,JMMMM4MMMMML       ,JMMMMMMMMMML      " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK057153222" & TD & "       JMMMMMMMMF    JMMF`      4MMM.    (MMMM""""""""4MMM`      " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK085168226" & TD & "     JMMM""          JMM`          4MM.   MMM                 " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK113182231" & TD & "    MMM`           ,MM`            4MM   `MMM_               " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK142197236" & TD & "   MMM ,M_,_____.  JMF             `MM)    ""MMMM_.           " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK170211241" & TD & "  JMM  (MMMMMMMMM) MM)              MM)       ""MMMMM_        " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK198226245" & TD & "  MM`          4M) MM)             ,MM)          `4MMMM.     " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK227240250" & TD & "  MM           (M) (MM             JMM ,ML          `4MMM    " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK255255255" & TD & "  MM           (M)  MMM.          JMMF `MM)            MM)   " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK255220197" & TD & "  MM.          (M)  `MMML.      _MMMF   4MML          ,MM),  " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK255186139" & TD & "  4MML.        JM)    4MMMMMMMMMMMM`     `MMMM_______MMMM    " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK255151081" & TD & "   `MMMMMMMMMMMMM)     `4MMMMMMMM""         `""MMMMMMMMMM`     " & TD & "BACK000000000" & TD & vbCrLf & TD & "RGB000000000" & TD & _
                TD & "BACK255116023" & TD & "       `""""""""""""""                                              " & TD & "BACK000000000" & TD & vbCrLf & _
                TD & "BACK000000000" & TD, True
    
    AddToStatus TD & "RGB220220220" & TD & "#" & _
            "  Time  # " & TD & "RGB200200200" & TD & "| Message" & vbCrLf, True

    AddToStatus "Going to " & gMudPath

    Set mRubrica = New cRubrica
    mRubrica.LoadRubrica
    'AddToStatus "Caricata rubrica da rubrica.gos"
    AddToStatus "Loaded rubrica from rubrica.gos"
    ReDim mWin(0) As cFinestra
    
    Set mConfig = New cIni
    mConfig.CaricaFile "config.ini"
    'AddToStatus "Caricata configurazione da config.ini"
    AddToStatus "Loaded configurations from config.ini"
    
    Set mPalette = New cPalette

    ProfSel = Me.GetConfig("ProfileSelected")
    If ProfSel = 0 Then
        ProfSelFolder = gMudPath & "p_(nessuno)\"
    Else
        ProfSelFolder = gMudPath & "p_" & LCase$(Me.GetConfig("profilo<" & ProfSel & ">")) & "\"
    End If
    If Dir$(ProfSelFolder, vbDirectory) = "" Then MkDir ProfSelFolder
    
    'carica in mProfConfig (cIni) il file config.ini per il profilo
    LoadProfConfig
End Sub

Public Sub RemoveConfig(Key As String)
    mConfig.RemoveInfo Key
End Sub

Public Sub SaveConfig()
    mConfig.SalvaFile
End Sub

Public Function GetConfig(Key As String, Optional Default As Variant = 0) As Variant
    GetConfig = mConfig.RetrInfo(Key, Default)
End Function

Public Sub SetConfig(Key As String, Value As Variant)
    mConfig.AddInfo Key, Value
End Sub

Private Sub Class_Terminate()
    mConfig.SalvaFile
    Set mConfig = Nothing
    
    mProfConfig.SalvaFile
    Set mProfConfig = Nothing
    
    Erase mWin()
    mRubrica.SaveRubrica
    Set mRubrica = Nothing
    Set mPalette = Nothing

    Set mLang = Nothing
End Sub

Public Property Get Palette() As cPalette
    Set Palette = mPalette
End Property
