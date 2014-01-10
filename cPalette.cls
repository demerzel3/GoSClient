VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Enum IntensityStates
    isNormal = 0
    isBold = 1
    isFaint = 2
End Enum

Public rgbDefault As Long
Public rgbBlack As Long
Public rgbGrey As Long
Public rgbLightGrey As Long
Public rgbWhite As Long
Public rgbGreen As Long
Public rgbLightGreen As Long
Public rgbRed As Long
Public rgbLightRed As Long
Public rgbMagenta As Long
Public rgbLightMagenta As Long
Public rgbBlue As Long
Public rgbLightBlue As Long
Public rgbCyan As Long
Public rgbLightCyan As Long
Public rgbYellow As Long
Public rgbLightYellow As Long

Private mPath As String

Private mBold As Boolean
Private mIntensity As IntensityStates
Private mLast As Long

Public Function AnsiColor(ByVal AnsiCode As String, _
    Optional ByRef BackColor As Long = 0, _
    Optional ByRef LastColor As Integer = 39) As Long
    
    Dim Color As Long
    
    Dim sFlags() As String, i As Integer
    Dim iFlags() As Integer ', bBold As Boolean
    
    'bBold = mBold
    'Color = mLast
    sFlags = Split(AnsiCode, ";")
    If UBound(sFlags, 1) = -1 Then ReDim sFlags(0) As String
    ReDim iFlags(0 To UBound(sFlags, 1)) As Integer
    For i = 0 To UBound(sFlags, 1)
        iFlags(i) = Val(sFlags(i))
    Next i
    Erase sFlags()
    'If Left$(AnsiCode, 1) = "0" Then AnsiCode = Mid$(AnsiCode, 2)
    'If Left$(AnsiCode, 1) = ";" Then AnsiCode = Mid$(AnsiCode, 2)
    
    'If Len(AnsiCode) >= 3 Then
    '    If Val(Left$(AnsiCode, 2)) >= 40 Then
    '        AnsiCode = Mid(AnsiCode, 4)
    '    End If
    'End If
    For i = 0 To UBound(iFlags, 1)
        Select Case iFlags(i)
            '////////////////options//////////////////////////////////
            Case 0 'normal, erase all modifications
                'bBold = False
                mIntensity = isNormal
                LastColor = 39
                BackColor = 0
                'Color = rgbDefault
            Case 1 'bold, increase intensity
                mIntensity = isBold
                'bBold = True
            Case 2 'faint, decrase intensity
                mIntensity = isFaint
'            Case 3 'italic
'            Case 4 'underlined
'            Case 5 'slow blinking
'            Case 6 'rapidly blinking
'            Case 7 'negative image
'            Case 8 'invisible characters
'            Case 9 'crossed out
'            Case 10 To 20 'change font
'            Case 21 'doubly underlined
            Case 22 'default foreground color (disable bold and faint)
                mIntensity = isNormal
                'bBold = False
'            Case 23 'not italic
'            Case 24 'not underlined
'            Case 25 'steady (not blinking)
'            Case 26 '[reserved]
'            Case 27 'positive image
'            Case 28 'visible characters
'            Case 29 'not crossed out
            Case 49 'default background color
                BackColor = 0

            '////////////////foreground//////////////////////////////////
            Case 30 To 37, 39
                LastColor = iFlags(i)

'            Case 30 'nero
'                Color = IIf(Not bBold, rgbBlack, rgbGrey)
'            Case 31 'rosso
'                Color = IIf(Not bBold, rgbRed, rgbLightRed)
'            Case 32 'verde
'                Color = IIf(Not bBold, rgbGreen, rgbLightGreen)
'            Case 33 'marrone
'                Color = IIf(Not bBold, rgbYellow, rgbLightYellow)
'            Case 34 'blu
'                Color = IIf(Not bBold, rgbBlue, rgbLightBlue)
'            Case 35 'magenta
'                Color = IIf(Not bBold, rgbMagenta, rgbLightMagenta)
'            Case 36 'ciano
'                Color = IIf(Not bBold, rgbCyan, rgbLightCyan)
'            Case 37 'grigio chiaro
'                Color = IIf(Not bBold, rgbLightGrey, rgbWhite)
            
            '////////////////backgound//////////////////////////////////
            Case 40 'nero
                BackColor = rgbBlack
            Case 41 'rosso
                BackColor = rgbRed
            Case 42 'verde
                BackColor = rgbGreen
            Case 43 'marrone
                BackColor = rgbYellow
            Case 44 'blu
                BackColor = rgbBlue
            Case 45 'magenta
                BackColor = rgbMagenta
            Case 46 'ciano
                BackColor = rgbCyan
            Case 47 'grigio chiaro
                BackColor = rgbLightGrey
            
            '////////////////unrecognized////////////////////////////////
            Case Else
                Color = rgbDefault
        End Select
    Next
    'frmLog.Log "ANSI: " & Replace(AnsiCode, ";", "  ")
    
    Color = GetForeColor(LastColor, mIntensity)
    
    'mLast = Color
    AnsiColor = Color
    'mBold = bBold
End Function

Private Function GetForeColor(ByVal Code As Integer, ByVal Intensity As IntensityStates) As Long
    Dim Color As Long
    
    Select Case Code
        Case 30 'nero
            Color = IIf(Not Intensity = isBold, rgbBlack, rgbGrey)
        Case 31 'rosso
            Color = IIf(Not Intensity = isBold, rgbRed, rgbLightRed)
        Case 32 'verde
            Color = IIf(Not Intensity = isBold, rgbGreen, rgbLightGreen)
        Case 33 'marrone
            Color = IIf(Not Intensity = isBold, rgbYellow, rgbLightYellow)
        Case 34 'blu
            Color = IIf(Not Intensity = isBold, rgbBlue, rgbLightBlue)
        Case 35 'magenta
            Color = IIf(Not Intensity = isBold, rgbMagenta, rgbLightMagenta)
        Case 36 'ciano
            Color = IIf(Not Intensity = isBold, rgbCyan, rgbLightCyan)
        Case 37 'grigio chiaro
            Color = IIf(Not Intensity = isBold, rgbLightGrey, rgbWhite)
    End Select
    
    If Intensity = isFaint Then Color = VariaColore(Color, -50)
    If Code = 39 Then Color = rgbDefault
    GetForeColor = Color
End Function

Public Sub LoadColors(Optional file As String = "")
    Dim Pal As cIni
    Dim Connect As cConnector

    If file = "" Then
        Set Connect = New cConnector
            file = Connect.ProfConf.RetrInfo("Colours", "(default).col")
        Set Connect = Nothing
    End If
    
    mPath = App.Path & "\colours\" ' & file
    If Dir$(mPath, vbDirectory) = "" Then MkDir mPath
    mPath = mPath & file
    
    Set Pal = New cIni
    Pal.CaricaFile mPath, True
    With Pal
        rgbDefault = .RetrInfo("pal_Default", 16777215)
        rgbBlack = .RetrInfo("pal_Black", 0)
        rgbGrey = .RetrInfo("pal_Grey", rgb(100, 100, 100))
        rgbLightGrey = .RetrInfo("pal_LightGrey", rgb(200, 200, 200))
        rgbWhite = .RetrInfo("pal_White", rgb(255, 255, 255))
        rgbGreen = .RetrInfo("pal_Green", rgb(0, 200, 0))
        rgbLightGreen = .RetrInfo("pal_LightGreen", rgb(100, 255, 100))
        rgbRed = .RetrInfo("pal_Red", rgb(200, 0, 0))
        rgbLightRed = .RetrInfo("pal_LightRed", rgb(255, 100, 100))
        rgbMagenta = .RetrInfo("pal_Magenta", rgb(200, 0, 200))
        rgbLightMagenta = .RetrInfo("pal_LightMagenta", rgb(255, 100, 255))
        rgbBlue = .RetrInfo("pal_Blue", rgb(0, 0, 200))
        rgbLightBlue = .RetrInfo("pal_LightBlue", rgb(100, 100, 255))
        rgbCyan = .RetrInfo("pal_Cyan", rgb(0, 200, 200))
        rgbLightCyan = .RetrInfo("pal_LightCyan", rgb(100, 255, 255))
        rgbYellow = .RetrInfo("pal_Yellow", rgb(200, 200, 0))
        rgbLightYellow = .RetrInfo("pal_LightYellow", rgb(255, 255, 100))
    End With
    Set Pal = Nothing
    
    mLast = rgbDefault
End Sub

Public Sub SaveColors(Optional file As String = "")
    Dim Pal As cIni

    If file <> "" Then
        mPath = App.Path & "\colours\" ' & file
        If Dir$(mPath, vbDirectory) = "" Then MkDir mPath
        mPath = mPath & file
    End If
    
    Set Pal = New cIni
    Pal.CaricaFile mPath, True
    With Pal
        Call .AddInfo("pal_Default", rgbDefault)
        Call .AddInfo("pal_Black", rgbBlack)
        Call .AddInfo("pal_Grey", rgbGrey)
        Call .AddInfo("pal_LightGrey", rgbLightGrey)
        Call .AddInfo("pal_White", rgbWhite)
        Call .AddInfo("pal_Green", rgbGreen)
        Call .AddInfo("pal_LightGreen", rgbLightGreen)
        Call .AddInfo("pal_Red", rgbRed)
        Call .AddInfo("pal_LightRed", rgbLightRed)
        Call .AddInfo("pal_Magenta", rgbMagenta)
        Call .AddInfo("pal_LightMagenta", rgbLightMagenta)
        Call .AddInfo("pal_Blue", rgbBlue)
        Call .AddInfo("pal_LightBlue", rgbLightBlue)
        Call .AddInfo("pal_Cyan", rgbCyan)
        Call .AddInfo("pal_LightCyan", rgbLightCyan)
        Call .AddInfo("pal_Yellow", rgbYellow)
        Call .AddInfo("pal_LightYellow", rgbLightYellow)
    End With
    Pal.SalvaFile
    Set Pal = Nothing
End Sub
