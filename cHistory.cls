VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents mtxtBox As TextBox
Attribute mtxtBox.VB_VarHelpID = -1
Private mMemory() As String
Private mCurPos As Long
Private mErasePrompt As Boolean

Public Property Let ErasePrompt(data As Boolean)
    mErasePrompt = data
End Property

Public Sub Init(txt As TextBox)
    Set mtxtBox = txt
End Sub

Private Sub Class_Initialize()
    ReDim mMemory(0 To 0) As String
    mCurPos = 1
End Sub

Private Sub Class_Terminate()
    Set mtxtBox = Nothing
End Sub

Private Sub mtxtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim NMem As Long

    If KeyCode = vbKeyEscape Then
        'clear textbox and put the pointer to the end of the history list
        mtxtBox.Text = ""
        mCurPos = UBound(mMemory, 1) + 1
    End If

    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        NMem = UBound(mMemory, 1)
        Select Case KeyCode
            Case vbKeyUp 'alto
                If Not mCurPos = 1 Then mCurPos = mCurPos - 1 Else Beep
            Case vbKeyDown 'basso vbkeydown
                If Not mCurPos = NMem + 1 Then
                    mCurPos = mCurPos + 1
                Else
                    If mtxtBox.Text = "" Then
                        Beep
                    Else
                        NMem = AddString(mtxtBox.Text)
                    End If
                End If
        End Select
        
        If mCurPos = NMem + 1 Then
            mtxtBox.Text = ""
        Else
            mtxtBox.Text = mMemory(mCurPos)
        End If
        
        mtxtBox.SelStart = Len(mtxtBox.Text)
        KeyCode = 0
    End If
End Sub

Private Sub mtxtBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then KeyAscii = 0
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        AddString mtxtBox.Text
        
        If mErasePrompt Then
            mtxtBox.Text = ""
        Else
            mtxtBox.SelStart = 0
            mtxtBox.SelLength = Len(mtxtBox.Text)
        End If
    End If
End Sub

Private Function AddString(ByRef Command As String) As Long
    Dim NMem As Long
    
    NMem = UBound(mMemory, 1)
    If Not (mtxtBox.Text = mMemory(NMem) Or mtxtBox.Text = "") Then
        NMem = NMem + 1
        If NMem = 1 Then
            ReDim mMemory(1 To 1) As String
        Else
            ReDim Preserve mMemory(1 To NMem) As String
        End If
        mMemory(NMem) = Command
    End If
    mCurPos = NMem + 1
    
    AddString = NMem
End Function
