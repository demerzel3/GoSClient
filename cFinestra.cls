VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFinestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mRif As Form
'Private mInit As Boolean
'Private mScartoX As Integer, mScartoY As Integer
Private mWinFlags As Integer

'da sostituire con property get e let
Public winID As Integer
    
Public Event envOutput(data As String, OutType As Integer)
Public Event envInput(data As String, InType As Integer)
'Public Event envChangeProfile()
'Public Event envClose()
'Public Event envEndRec()
'Public Event envPalChanged()
Public Event envNotify(uMsg As Long)

Public Sub sendNotify(uMsg As Long)
    RaiseEvent envNotify(uMsg)
End Sub

Public Sub sendOutput(out As String, OutType As Integer)
    RaiseEvent envOutput(out, OutType)
End Sub

Public Sub sendInput(inp As String, InType As Integer)
    RaiseEvent envInput(inp, InType)
End Sub

Public Property Get WinFlags() As Integer
    WinFlags = mWinFlags
End Property

Public Property Get Caption() As String
    Caption = mRif.Caption
End Property

Public Sub Init(Form As Form, Optional WinFlags As Integer = WINREC_NONE)
    'Dim Style As Long
    'Dim hRgn As Long
    Dim Connect As cConnector

    'Dim data As cIni, i As Integer, Subclassa As Boolean
        
    mWinFlags = WinFlags
        
    Set mRif = Form
    mRif.AutoRedraw = True
    mRif.ScaleMode = vbPixels
    'mRif.BackColor = GOSRGB_FORM_Back
    mRif.BackColor = GetSysColor(COLOR_3DFACE)

    'CalcScarto
    
    'If DoCls Then Cls
        
    'commentare la riga sotto per prevenire il subclassing
    'Subclassa = True
    'Set data = New cIni
    '    data.CaricaFile "Workspace.ini"
    '    For i = 1 To data.RetrInfo("frmsCount", 0)
    '        If TypeName(Form) = data.RetrInfo("frm<" & i & ">_Name") Then
    '            Subclassa = False
    '            Exit For
    '        End If
    '    Next i
    'Set data = Nothing
    '
    'If Subclassa Then
    '    glProcOld = SetWindowLong(mRif.hWnd, GWL_WNDPROC, AddressOf WindowProc)
    '    Style = GetWindowLong(mRif.hWnd, GWL_STYLE)
    '    'SetWindowLong mRif.hwnd, GWL_STYLE, Style Xor WS_SYSMENU
    '    If (Style And WS_SYSMENU) = WS_SYSMENU Then
    '        SetWindowLong mRif.hWnd, GWL_STYLE, Style Xor (WS_SYSMENU)
    '    End If
    '    mInit = True
    'End If
    
    'Style = GetWindowLong(mRif.hWnd, GWL_STYLE)
    'If Not mRif.hWnd = frmBase.hWnd Then
        'SetWindowLong mRif.hWnd, GWL_HWNDPARENT, frmBase.hWnd
        'SetWindowLong mRif.hWnd, GWL_STYLE, Style Or WS_CHILD
    'End If
    
    'hRgn = CreateRectRgn(0, 0, Screen.Width / Screen.TwipsPerPixelX + BorderWidth * 2, Screen.Height / Screen.TwipsPerPixelY + BorderWidth * 2)
    'SetWindowRgn mRif.hwnd, hRgn, True
    'SetWindowPos mRif.hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
    
    Set Connect = New cConnector
        Connect.RegWin Me
    Set Connect = Nothing
End Sub

Public Sub UnReg()
    Dim Connect As cConnector

    Set Connect = New cConnector
    Connect.UnRegWin winID
    Set Connect = Nothing
End Sub

Private Sub Class_Terminate()
    'If mInit Then Call SetWindowLong(mRif.hwnd, GWL_WNDPROC, glProcOld)
    
    'Debug.Print "class terminate"

    Set mRif = Nothing
End Sub
