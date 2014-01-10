VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKeyCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mShift As Integer
Private mValid As Boolean
Private WithEvents mtxtCombo As TextBox
Attribute mtxtCombo.VB_VarHelpID = -1
Private mAliases As cAlias
Private mMode As Integer
Private mAliasSent As Boolean

Private Const MODE_RICEZIONE As Integer = 1
Private Const MODE_RICONOSCIMENTO As Integer = 2

Private Const KEY_SHIFT = 16
Private Const KEY_CTRL = 17
Private Const KEY_ALT = 18

Public Event SetCombo(KeyCode As Integer, Shift As Integer)
Public Event AvviaAlias(Index As Integer)

Public Sub ScindiCombo(Combo As String, ByRef KeyCode As Integer, ByRef Shift As Integer)
    Dim Mezzo As Integer

    If Combo = "" Or Combo = "0|0" Then
        Shift = 0
        KeyCode = 0
    Else
        Mezzo = InStr(1, Combo, "|", vbTextCompare)
        KeyCode = CInt(Mid$(Combo, 1, Mezzo - 1))
        Shift = CInt(Mid$(Combo, Mezzo + 1))
    End If
End Sub

Public Function IdentCombo(KeyCode As Integer, Shift As Integer) As String
    Dim Combo As String
    Dim Valid As Boolean

    Valid = True
        If Not Shift = 0 Then
            If (Shift And vbCtrlMask) = vbCtrlMask Then Combo = Combo & "CTRL + "
            If (Shift And vbShiftMask) = vbShiftMask Then Combo = Combo & "SHIFT + "
            If (Shift And vbAltMask) = vbAltMask Then Combo = Combo & "ALT + "
        End If
        
        If (KeyCode >= vbKeyA And KeyCode <= vbKeyZ) Or _
           (KeyCode >= vbKey0 And KeyCode <= vbKey9) Then
            Combo = Combo & Chr$(KeyCode)
        ElseIf KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
            Combo = Combo & "F" & (KeyCode - vbKeyF1 + 1)
        Else
            Valid = False
        End If
        
        If Not mtxtCombo Is Nothing Then
            If Valid Then
                mtxtCombo.Text = Combo
            Else
                mtxtCombo.Text = "(none)"
            End If
        End If
        
    IdentCombo = Combo
End Function

Public Sub AvviaRicezione(box As TextBox, Alias As cAlias)
    Set mtxtCombo = box
    Set mAliases = Alias
    mMode = MODE_RICEZIONE
End Sub

Public Sub AvviaRiconoscimento(box As TextBox)
    Set mtxtCombo = box
    mMode = MODE_RICONOSCIMENTO
End Sub

Private Sub Class_Terminate()
    Set mAliases = Nothing
    Set mtxtCombo = Nothing
End Sub

Private Sub mtxtCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Combo As String
    
    If mMode = MODE_RICONOSCIMENTO Then
        mShift = Shift
        If KeyCode <> KEY_SHIFT And KeyCode <> KEY_CTRL And KeyCode <> KEY_ALT Then
            mValid = True
        End If
    Else
        Combo = KeyCode & "|" & Shift
        mAliasSent = False
        If Combo <> "0|0" Then
            For i = 1 To mAliases.Count
                If Combo = mAliases.Combo(i) Then
                    mAliasSent = True
                    RaiseEvent AvviaAlias(i)
                    KeyCode = 0
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub mtxtCombo_KeyPress(KeyAscii As Integer)
    If mMode = MODE_RICONOSCIMENTO Then
        If mShift <> 0 Then
            mValid = True
        Else
            mValid = False
        End If
    Else
        If mAliasSent = True Then KeyAscii = 0
    End If
End Sub

Private Sub mtxtCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Combo As String

    If mMode = MODE_RICONOSCIMENTO Then
        If KeyCode = KEY_SHIFT Or KeyCode = KEY_CTRL Or KeyCode = KEY_ALT Or KeyCode = 13 Then
            mValid = False
        End If
        
        If mValid Then
            IdentCombo KeyCode, mShift
            If mtxtCombo.Text <> "(nessuna)" Then
                'mtxtCombo.Text = Combo
                RaiseEvent SetCombo(KeyCode, mShift)
            End If
        End If
        
        mShift = 0
    End If
End Sub
