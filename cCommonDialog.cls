VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'///////////////////////////////////////////////////////////////////////////////
'//                                                                           //
'// PROJECT     :                                                             //
'// FILE        : ComDlg32.cls                                                //
'// VERSION     : 1.2.0                                                       //
'// CREATED     : ??/11/1997  BY : Roal Zanazzi                               //
'// MODIFIED    : 06/05/1999  BY : Roal Zanazzi                               //
'// LANGUAGE    : Visual Basic 5, 5+SP3, 6, 6+SP3                             //
'// ENVIRONMENT : Windows 95, 98                                              //
'// DESCRIPTION : The CommonDialog class is designed to substitute MS Common  //
'//   Dialog OCX with pure Visual Basic code.                                 //
'//   This class works like "MS Common Dialogs" OCX.                          //
'//   For more informations look at "MS Common Dialogs" OCX online help;      //
'//   all Flags constants are supported.                                      //
'//                                                                           //
'// Copyright © 1997-1999 Roal Zanazzi <zrs@iol.it>                           //
'// This work is partially based on the clsPrintDialog class written by       //
'// Luca Minudel <luca.minudel@nline.it>                                      //
'///////////////////////////////////////////////////////////////////////////////
'//                                                                           //
'// NOTE: VB4 doesn't support public costant declaration in class module.     //
'//       To use this code with VB4 remove Enum declarations and put          //
'//       costant declarations in a .BAS module.                              //
'//                                                                           //
'///////////////////////////////////////////////////////////////////////////////
'//                                                                           //
'// HISTORY:                                                                  //
'//                                                                           //
'//   1.0.0  ??/11/1997  Roal Zanazzi                                         //
'//     -                                                                     //
'//                                                                           //
'//   1.0.1  20/03/1998  Roal Zanazzi                                         //
'//     - Corrected bug in ShowOpen and ShowSave methods, using the function  //
'//       sLeftOfNull() to return filename and filetitle.                     //
'//                                                                           //
'//   1.1.0  19/01/1999  Roal Zanazzi                                         //
'//     - Added ShowColor, ShowFont and ShowHelp methods.                     //
'//     - Modified ShowPrinter method to correctly return Copies for printer  //
'//       drivers that always return 1 with the MS ComDlg OCX.                //
'//                                                                           //
'//   1.2.0  06/05/1999  Roal Zanazzi                                         //
'//     - Corrected some bugs in ShowOpen and ShowSave methods.               //
'//                                                                           //
'///////////////////////////////////////////////////////////////////////////////

Option Explicit


' Comment any of these constants if you are not using the related feature.
#Const USE_COLOR = True
#Const USE_FILEOPEN = False
#Const USE_FILESAVE = False
#Const USE_FONT = True
#Const USE_HELP = False
#Const USE_PRINTER = False


'///////////////////////////////////////////////////////////////////////////////
' API constants
'
' ---- ShowColor
#If USE_COLOR Then
' Flags
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_RGBINIT = &H1&
Private Const CC_SHOWHELP = &H8&
#End If
' ---- ShowFont
#If USE_FONT Then
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
' Flags
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_TTONLY = &H40000
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_USESTYLE = &H80&
Private Const CF_WYSIWYG = &H8000
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
#End If
' ---- ShowPrinter
#If USE_PRINTER Then
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
' API memory functions.
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
#End If


'///////////////////////////////////////////////////////////////////////////////
' API types
'
' ---- ShowColor
#If USE_COLOR Then
Private Type CHOOSECOLOR_TYPE
    lStructSize As Long
    hWndOwner As Long ' HWND
    hInstance As Long ' HINSTANCE
    rgbResult As Long ' COLORREF
    lpCustColors As Long ' COLORREF[16]
    Flags As Long
    lCustData As Long ' LPARAM
    lpfnHook As Long ' LPCCHOOKPROC
    lpTemplateName As String
End Type
#End If
' ---- ShowFont
#If USE_FONT Then
'constants for LOGFONT
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type CHOOSEFONT_TYPE
    lStructSize As Long
    hWndOwner As Long 'HWND
    hdc As Long 'HDC
    lpLogFont As Long 'LPLOGFONT
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long 'LPCFHOOKPROC
    lpTemplateName As String
    hInstance As Long 'HINSTANCE
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
#End If
' ---- ShowOpen and ShowSave
#If USE_FILEOPEN Or USE_FILESAVE Then
Private Type OPENFILENAME
   lStructSize As Long
   hWndOwner As Long 'HWND
   hInstance As Long 'HINSTANCE
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   Flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type
#End If
' ---- ShowPrinter
#If USE_PRINTER Then
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long 'HWND
    hDevMode As Long
    hDevNames As Long
    hdc As Long 'HDC
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long 'HINSTANCE
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
#End If


'///////////////////////////////////////////////////////////////////////////////
' API functions declarations
'
' API function to retrieve extended error information
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
' ---- ShowColor
#If USE_COLOR Then
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR_TYPE) As Long
#End If
' ---- ShowFont
#If USE_FONT Then
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT_TYPE) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90
#End If
' ---- ShowHelp
#If USE_HELP Then
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
#End If
' ---- ShowOpen
#If USE_FILEOPEN Then
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If
' ---- ShowPrinter
#If USE_PRINTER Then
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
#End If
' ---- ShowSave
#If USE_FILESAVE Then
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If


'///////////////////////////////////////////////////////////////////////////////
' Class public enums.
'
Public Enum COMDLGConstants
    dummy = 0 ' Just in case no useful USE_*** defined.
' ---- ShowColor
#If USE_COLOR Then
    cdlCCFullOpen = CC_FULLOPEN
    cdlCCHelpButton = CC_SHOWHELP
    cdlCCPreventFullOpen = CC_PREVENTFULLOPEN
    cdlCCRGBInit = CC_RGBINIT
#End If
' ---- ShowFont
#If USE_FONT Then
    cdlCFANSIOnly = CF_ANSIONLY
    cdlCFApply = CF_APPLY
    cdlCFBoth = CF_BOTH
    cdlCFEffects = CF_EFFECTS
    cdlCFFixedPitchOnly = CF_FIXEDPITCHONLY
    cdlCFForceFontExist = CF_FORCEFONTEXIST
    cdlCFHelpButton = CF_SHOWHELP
    cdlCFInitToLogFontStruct = CF_INITTOLOGFONTSTRUCT
    cdlCFLimitSize = CF_LIMITSIZE
    cdlCFNoFaceSel = CF_NOFACESEL
    cdlCFNoSimulations = CF_NOSIMULATIONS
    cdlCFNoSizeSel = CF_NOSIZESEL
    cdlCFNoStyleSel = CF_NOSTYLESEL
    cdlCFNoVectorFonts = CF_NOVECTORFONTS
'    cdlNoVerticalFonts = CF_NOVERTFONTS ' New!
    cdlCFPrinterFonts = CF_PRINTERFONTS
    cdlCFScalableOnly = CF_SCALABLEONLY
    cdlCFScreenFonts = CF_SCREENFONTS
    cdlCFTTOnly = CF_TTONLY
    cdlCFWYSIWYG = CF_WYSIWYG
#End If
' ---- ShowOpen and ShowSave
#If USE_FILEOPEN Or USE_FILESAVE Then
    cdlOFNAllowMultiselect = &H200
    cdlOFNCreatePrompt = &H2000
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400
    cdlOFNFileMustExist = &H1000
    cdlOFNHelpButton = &H10
    cdlOFNHideReadOnly = &H4
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000
    cdlOFNNoValidate = &H100
    cdlOFNOverwritePrompt = &H2
    cdlOFNPathMustExist = &H800
    cdlOFNReadOnly = &H1
    cdlOFNShareAware = &H4000
#End If
' ---- ShowPrinter
#If USE_PRINTER Then
    cdlPDAllPages = &H0
    cdlPDCollate = &H10
    cdlPDDisablePrintToFile = &H80000
    cdlPDHelpButton = &H800
    cdlPDHidePrintToFile = &H100000
    cdlPDNoPageNums = &H8
    cdlPDNoSelection = &H4
    cdlPDNoWarning = &H80
    cdlPDPageNums = &H2
    cdlPDPrintSetup = &H40
    cdlPDPrintToFile = &H20
    cdlPDReturnDC = &H100
    cdlPDReturnDefault = &H400
    cdlPDReturnIC = &H200
    cdlPDSelection = &H1
    cdlPDUseDevModeCopies = &H40000
#End If
End Enum
' Common
Public Enum ErrorConstants
    cdlCancel = 32755
End Enum


'///////////////////////////////////////////////////////////////////////////////
' Class private variables (internal storage for class properties).
'
Public p_DialogTitle As String
Private p_ExtendedError As Long
Private p_Parent As Object
' ---- ShowColor
#If USE_COLOR Or USE_FONT Then
Private p_Color As Long
#End If
' ---- ShowFont
#If USE_FONT Then
Private p_FontName As String
Private p_FontSize As Integer
Private p_FontBold As Boolean
Private p_FontItalic As Boolean
Private p_FontStrikethru As Boolean
Private p_FontUnderline As Boolean
#End If
' ---- ShowHelp
Private p_HelpFile As String
Private p_HelpCommand As Long
Private p_HelpContext As String
Private p_HelpKey As String
' ---- ShowOpen and ShowSave
#If USE_FILEOPEN Or USE_FILESAVE Then
Private p_FileTitle As String
#End If
' ---- ShowPrinter
Private p_hDC As Long
#If USE_PRINTER Then
Private p_Copies As Integer
' Note: 0 >= Min >= FromPage >= ToPage >= Max
'       If Max=0 then no limits.
Private p_Min As Integer
Private p_Max As Integer
Private p_FromPage As Integer
Private p_ToPage As Integer
#End If


'///////////////////////////////////////////////////////////////////////////////
' Class public variables (class properties).
'
Public CancelError As Boolean
Public Flags As COMDLGConstants
#If USE_FILEOPEN Or USE_FILESAVE Then
Public DefaultExt As String
Public Filename As String
Public Filter As String
Public FilterIndex As Long
Public InitDir As String
Public MaxFileSize As Long
#End If


' Class initialization.
Private Sub Class_Initialize()
    CancelError = False
    ' ShowOpen and ShowSave
#If USE_FILEOPEN Or USE_FILESAVE Then
    MaxFileSize = 256
#End If
    ' ShowPrinter
#If USE_PRINTER Then
    p_Min = 0
    p_Max = 0
    p_FromPage = 0
    p_ToPage = 0
#End If
End Sub


'///////////////////////////////////////////////////////////////////////////////
' Class properties.
'

Public Property Get DialogTitle() As String
    DialogTitle = p_DialogTitle
End Property
Public Property Let DialogTitle(Value As String)
    p_DialogTitle = Value
End Property

Public Property Get ExtendedError() As Long
    ExtendedError = p_ExtendedError
End Property

Property Get Parent() As Object
    Parent = p_Parent
End Property
Property Set Parent(Value As Object)
    Set p_Parent = Value
    p_hDC = p_Parent.hdc
End Property

' ShowOpen and ShowSave
#If USE_FILEOPEN Or USE_FILESAVE Then
Property Get FileTitle() As String
    FileTitle = p_FileTitle
End Property
#End If

' ShowPrinter
Property Get hdc() As Integer
    hdc = p_hDC
End Property

#If USE_PRINTER Then
Property Get Copies() As Integer
    If ((Flags And cdlPDUseDevModeCopies) = cdlPDUseDevModeCopies) Then
        Copies = 1
    Else
        Copies = p_Copies
    End If
End Property
Property Let Copies(ByVal Value As Integer)
    If (Value > 0) Then p_Copies = Value Else p_Copies = 0
End Property

Property Get Max() As Integer
    Max = p_Max
End Property
Property Let Max(ByVal Value As Integer)
    If (Value > 0) Then p_Max = Value Else p_Max = 0
    If (Value < p_ToPage) Then p_ToPage = Value
    If (Value < p_FromPage) Then p_FromPage = Value
    If (Value < p_Min) Then p_Min = Value
End Property

Property Get Min() As Integer
    Min = p_Min
End Property
Property Let Min(ByVal Value As Integer)
    If (Value > 0) Then p_Min = Value Else p_Min = 0
    If (Value > p_Max) Then p_Max = Value
    If (Value > p_FromPage) Then p_FromPage = Value
    If (Value > p_ToPage) Then p_ToPage = Value
End Property

Property Get FromPage() As Integer
    FromPage = p_FromPage
End Property
Property Let FromPage(ByVal Value As Integer)
    If (Value > 0) Then p_FromPage = Value Else p_FromPage = 0
    If (Value > p_Max) Then p_Max = Value
    If (Value < p_Min) Then p_Min = Value
    If (Value > p_ToPage) Then p_ToPage = Value
End Property

Property Get ToPage() As Integer
    ToPage = p_ToPage
End Property
Property Let ToPage(ByVal Value As Integer)
    If (Value > 0) Then p_ToPage = Value Else p_ToPage = 0
    If (Value > p_Max) Then p_Max = Value
    If (Value < p_Min) Then p_Min = Value
    If (Value < p_FromPage) Then p_FromPage = Value
End Property
#End If

' ShowColor
#If USE_COLOR Or USE_FONT Then
Property Get Color() As Long
    Color = p_Color
End Property
Property Let Color(Value As Long)
    p_Color = Value
End Property
#End If

' ShowFont
#If USE_FONT Then
Public Property Get FontBold() As Boolean
    FontBold = p_FontBold
End Property
Public Property Let FontBold(Value As Boolean)
    p_FontBold = Value
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = p_FontItalic
End Property
Public Property Let FontItalic(Value As Boolean)
    p_FontItalic = Value
End Property

Public Property Get FontName() As String
    FontName = p_FontName
End Property
Public Property Let FontName(Value As String)
    p_FontName = Value
End Property

Public Property Get FontSize() As Long
    FontSize = p_FontSize
End Property
Public Property Let FontSize(Value As Long)
    p_FontSize = Value
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = p_FontStrikethru
End Property
Public Property Let FontStrikethru(Value As Boolean)
    p_FontStrikethru = Value
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = p_FontUnderline
End Property
Public Property Let FontUnderline(Value As Boolean)
    p_FontUnderline = Value
End Property
#End If

' ShowHelp
Property Let HelpCommand(Value As Long)
    p_HelpCommand = Value
End Property

Public Property Get HelpContext() As String
    HelpContext = p_HelpContext
End Property
Public Property Let HelpContext(Value As String)
    p_HelpContext = Value
End Property

Public Property Get HelpFile() As String
    HelpFile = p_HelpFile
End Property
Property Let HelpFile(Value As String)
    p_HelpFile = Value
End Property

Public Property Get HelpKey() As String
    HelpKey = p_HelpKey
End Property
Public Property Let HelpKey(Value As String)
    p_HelpKey = Value
End Property


'///////////////////////////////////////////////////////////////////////////////
' Private functions.
'

' Return the part of string preceding Chr$(0).
Private Function sLeftOfNull(ByVal s As String)
    Dim lNullPos As Long
    ' Get position of first Chr$(0) in string.
    lNullPos = InStr(s, Chr$(0))
    ' Return part of string to left of first Chr$(0) if found.
    If (lNullPos > 0) Then
        sLeftOfNull = Mid$(s, 1, lNullPos - 1)
    Else
        sLeftOfNull = s
    End If
End Function

' Return a string from a byte array.
Private Function sByteArrayToString(abBytes() As Byte) As String
    Dim lBytePoint As Long
    Dim lByteVal As Long
    Dim sOut As String
    
    'init array pointer
    lBytePoint = LBound(abBytes)
    ' Fill sOut with characters in array.
    While lBytePoint <= UBound(abBytes)
        lByteVal = abBytes(lBytePoint)
        ' Return sOut and stop if Chr$(0) is encountered.
        If lByteVal = 0 Then
            sByteArrayToString = sOut
            Exit Function
        Else
            sOut = sOut & Chr$(lByteVal)
        End If
        lBytePoint = lBytePoint + 1
    Wend
    ' Return string if Chr$(0) wasn't encountered.
    sByteArrayToString = sOut
End Function

' Copy a string into a byte array
Private Function StringToByteArray(sSrc As String, abBytes() As Byte) As String
    Dim i As Integer
    
    For i = 0 To Len(sSrc) - 1
        abBytes(i) = Asc(Mid$(sSrc, i + 1, 1))
    Next i
    abBytes(i) = 0
End Function


'///////////////////////////////////////////////////////////////////////////////
' Class methods.
'

' ShowColor method.
#If USE_COLOR Then
Public Function ShowColor() As Boolean
    Dim cc As CHOOSECOLOR_TYPE
    Dim alCustColor(15) As Long ' COLORREF
    Dim bCancel As Boolean
    
    ' Initialize general properties.
    p_ExtendedError = 0
    bCancel = False
    ' Fill custom colors array with grey scale.
    Dim i As Byte
    For i = LBound(alCustColor) To UBound(alCustColor)
        alCustColor(i) = &HFFFFFF - (&H101010 * i)
    Next
    ' Prepare CHOOSECOLOR structure.
    With cc
        .lStructSize = Len(cc)
        If (p_Parent Is Nothing) Then .hWndOwner = 0 Else .hWndOwner = p_Parent.hWnd
        '.hInstance = App.hInstance
        .rgbResult = p_Color
        .lpCustColors = VarPtr(alCustColor(0))
        .Flags = Flags
        
        '.lCustData
        '.lpfnHook
        '.lpTemplateName
    End With
    ' Call the ChooseColor API function and properly handle return code.
    If ChooseColor(cc) Then
        p_Color = cc.rgbResult
    Else
        bCancel = True
        p_ExtendedError = CommDlgExtendedError
        If CancelError Then Err.Raise cdlCancel, "ColorDialog", "Cancel."
    End If
    ShowColor = Not bCancel
End Function
#End If


' ShowFont method.
#If USE_FONT Then
Public Function ShowFont() As Boolean
    Dim cf As CHOOSEFONT_TYPE
    Dim tLogFont As LOGFONT
    Dim bCancel As Boolean

    ' Initialize general properties.
    p_ExtendedError = 0
    bCancel = False
    
    If p_FontBold Then tLogFont.lfWeight = FW_BOLD
    If p_FontItalic Then tLogFont.lfItalic = 1
    If p_FontUnderline Then tLogFont.lfUnderline = 1
    If p_FontStrikethru Then tLogFont.lfStrikeOut = 1
    StringToByteArray p_FontName, tLogFont.lfFaceName
    tLogFont.lfHeight = -MulDiv(p_FontSize, GetDeviceCaps(p_hDC, LOGPIXELSY), 72)

    With cf
        .lStructSize = Len(cf)
        If (p_Parent Is Nothing) Then .hWndOwner = 0 Else .hWndOwner = p_Parent.hWnd
        .hdc = p_hDC
        .lpLogFont = VarPtr(tLogFont)
        .iPointSize = p_FontSize
        .Flags = Flags
        .rgbColors = p_Color
        '.lCustData
        '.lpfnHook
        '.lpTemplateName
        '.hInstance = App.hInstance
        .lpszStyle = vbNull
        .nFontType = 0
        '.MISSING_ALIGNMENT
        .nSizeMin = 0
        .nSizeMax = 0
    End With
    ' Call the CHOOSEFONT API function.
    If (ChooseFont(cf)) Then
        p_FontName = sByteArrayToString(tLogFont.lfFaceName())
        p_FontSize = cf.iPointSize \ 10
        p_FontBold = (tLogFont.lfWeight > FW_NORMAL)
        p_FontItalic = tLogFont.lfItalic
        p_FontUnderline = tLogFont.lfUnderline
        p_FontStrikethru = tLogFont.lfStrikeOut
        p_Color = cf.rgbColors
    Else
        bCancel = True
        p_ExtendedError = CommDlgExtendedError
        If CancelError Then Err.Raise cdlCancel, "FontDialog", "Cancel."
    End If
    ShowFont = Not bCancel
End Function
#End If


' ShowHelp method.
' Run winhelp.exe with the specified help file.
#If USE_HELP Then
Public Function ShowHelp() As Boolean
    Dim lData As Long
    Dim hWndOwner As Long 'HWND
    Dim bCancel As Boolean
    
    ' Initialize general properties.
    p_ExtendedError = 0
    bCancel = False
    
    lData = 0
    If (p_Parent Is Nothing) Then hWndOwner = 0 Else hWndOwner = p_Parent.hWnd
    ' Call the API function.
    If (WinHelp(hWndOwner, p_HelpFile, p_HelpCommand, lData)) Then
        '
    Else
        bCancel = True
        p_ExtendedError = CommDlgExtendedError
        If CancelError Then Err.Raise cdlCancel, "HelpDialog", "Cancel."
    End If
    ShowHelp = Not bCancel
End Function
#End If


' ShowOpen method.
#If USE_FILEOPEN Then
Public Function ShowOpen() As Boolean
    Dim OFN As OPENFILENAME
    Dim lMaxSize As Long
    Dim sFileNameBuff As String
    Dim sFileTitleBuff As String
    Dim intFilter As String
    Dim bCancel As Boolean
    Dim i As Integer
    
    ' Initialize general properties.
    p_ExtendedError = 0
    bCancel = False
    ' Modify Filter string to change any | in Chr(0), as expected by the API.
    For i = 1 To Len(Filter)
        If (Mid$(Filter, i, 1) = "|") Then
            intFilter = intFilter & Chr$(0)
        Else
            intFilter = intFilter & Mid$(Filter, i, 1)
        End If
    Next
    intFilter = intFilter & Chr$(0) & Chr$(0)  ' double null terminated.
    
'    If (FileName = "") Then FileName = String$(MaxFileSize, 0)
    'determine size of buffer from MaxFileSize property
    If (MaxFileSize > 0) Then
        lMaxSize = MaxFileSize
    Else
        lMaxSize = 255
    End If
    sFileNameBuff = Filename
    'pad with spaces
    While Len(sFileNameBuff) < lMaxSize - 1
        sFileNameBuff = sFileNameBuff & " "
    Wend
    'trim to length of lMaxFileSize - 1
    sFileNameBuff = Mid$(sFileNameBuff, 1, MaxFileSize - 1)
    'null terminate
    sFileNameBuff = sFileNameBuff & Chr$(0)
    
    If (p_FileTitle = "") Then p_FileTitle = String$(256, 0)
    With OFN
        .lStructSize = Len(OFN)
        If (p_Parent Is Nothing) Then .hWndOwner = 0 Else .hWndOwner = p_Parent.hWnd
        '.hInstance = App.hInstance
        .lpstrFilter = intFilter
        '.lpstrCustomFilter
        '.nMaxCustFilter
        .nFilterIndex = FilterIndex
        .lpstrFile = sFileNameBuff
        .nMaxFile = MaxFileSize
        .lpstrFileTitle = p_FileTitle
        .nMaxFileTitle = 256
        .lpstrInitialDir = InitDir
        .lpstrTitle = DialogTitle
        .Flags = Flags
        '.nFileOffset
        '.nFileExtension
        .lpstrDefExt = DefaultExt
        '.lCustData
        '.lpfnHook
        '.lpTemplateName
    End With
    If (GetOpenFileName(OFN)) Then
        With OFN
        FilterIndex = .nFilterIndex
        Filename = sLeftOfNull(.lpstrFile)
        p_FileTitle = sLeftOfNull(.lpstrFileTitle)
        Flags = .Flags
        End With
    Else
        bCancel = True
        Filename = ""
        p_ExtendedError = CommDlgExtendedError
        If CancelError Then Err.Raise cdlCancel, "FileOpenDialog", "Cancel."
    End If
    ShowOpen = Not bCancel
End Function
#End If


' ShowSave method.
#If USE_FILESAVE Then
Public Function ShowSave() As Boolean
    Dim OFN As OPENFILENAME
    Dim intFilter As String
    Dim bCancel As Boolean
    Dim i As Integer
    
    bCancel = False
    ' Modify Filter string to change any | in Chr(0), as expected by the API.
    For i = 1 To Len(Filter)
        If (Mid$(Filter, i, 1) = "|") Then
            intFilter = intFilter & Chr$(0)
        Else
            intFilter = intFilter & Mid$(Filter, i, 1)
        End If
    Next
    intFilter = intFilter & Chr$(0) & Chr$(0)  ' double null terminated
    If (Filename = "") Then Filename = String$(MaxFileSize, 0)
    If (p_FileTitle = "") Then p_FileTitle = String$(256, 0)
    With OFN
        .lStructSize = Len(OFN)
        If (p_Parent Is Nothing) Then .hWndOwner = 0 Else .hWndOwner = p_Parent.hWnd
        '.hInstance
        .lpstrFilter = intFilter
        '.lpstrCustomFilter
        '.nMaxCustFilter
        .nFilterIndex = FilterIndex
        .lpstrFile = Filename
        .nMaxFile = MaxFileSize
        .lpstrFileTitle = p_FileTitle
        .nMaxFileTitle = 256
        .lpstrInitialDir = InitDir
        .lpstrTitle = DialogTitle
        .Flags = Flags
        '.nFileOffset
        '.nFileExtension
        .lpstrDefExt = DefaultExt
        '.lCustData
        '.lpfnHook
        '.lpTemplateName
    End With
    If (GetSaveFileName(OFN)) Then
        With OFN
        FilterIndex = .nFilterIndex
        Filename = sLeftOfNull(.lpstrFile)
        p_FileTitle = sLeftOfNull(.lpstrFileTitle)
        Flags = .Flags
        End With
    Else
        bCancel = True
        If CancelError Then Err.Raise cdlCancel, "FileSaveDialog", "Cancel."
    End If
    ShowSave = Not bCancel
End Function
#End If


' ShowPrinter method.
#If USE_PRINTER Then
Public Function ShowPrinter() As Boolean
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE
    Dim lpDevMode As Long, lpDevName As Long
    Dim objPrinter As Printer
    Dim strNewPrinterName As String
    Dim bCancel As Boolean
    
    ' Initialize general properties.
    p_ExtendedError = 0
    bCancel = False
    ' Use PrintDialog to get the handle to a memory block with a DevMode and DevName structures.
    With PrintDlg
        .lStructSize = Len(PrintDlg)
        If (p_Parent Is Nothing) Then .hWndOwner = 0 Else .hWndOwner = p_Parent.hWnd
        '.hDevMode
        '.hDevNames
        '.hDC
        .Flags = Flags
        .nFromPage = p_FromPage
        .nToPage = p_ToPage
        .nMinPage = p_Min
        .nMaxPage = p_Max
        .nCopies = p_Copies
        '.hInstance = App.hInstance
        '.lCustData
        '.lpfnPrintHook
        '.lpfnSetupHook
        '.lpPrintTemplateName
        '.lpSetupTemplateName
        '.hPrintTemplate
        '.hSetupTemplate
    End With
    ' Set the current orientation and duplex setting.
    With DevMode
        .dmDeviceName = Printer.DeviceName
        .dmSize = Len(DevMode)
        .dmFields = DM_ORIENTATION Or DM_DUPLEX
        .dmOrientation = Printer.Orientation
        On Error Resume Next
        .dmDuplex = Printer.Duplex
        On Error GoTo 0
    End With
    ' Allocate memory for the initialization hDevMode structure
    ' and copy the settings gathered above into this memory.
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If (lpDevMode > 0) Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        Call GlobalUnlock(lpDevMode)
    End If
    ' Set the current driver, device, and port name strings.
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    ' Allocate memory for the initial hDevName structure
    ' and copy the settings gathered above into this memory.
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        Call GlobalUnlock(lpDevName)
    End If
    ' Call the print dialog and let the user make changes.
    If PrintDialog(PrintDlg) Then
        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        Call GlobalUnlock(lpDevName)
        With PrintDlg
            p_hDC = .hdc
            Flags = .Flags
            p_Max = .nMaxPage
            p_Min = .nMinPage
            p_FromPage = .nFromPage
            p_ToPage = .nToPage
            p_Copies = .nCopies
        End With
        GlobalFree PrintDlg.hDevNames
        ' Get the DevMode structure and set the printer properties appropriately.
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        Call GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        strNewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> strNewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = strNewPrinterName Then Set Printer = objPrinter
            Next
        End If
        On Error Resume Next
        ' Set printer object properties according to selections made by user.
        With Printer
            .Copies = DevMode.dmCopies
            .Duplex = DevMode.dmDuplex
            .Orientation = DevMode.dmOrientation
        End With
        On Error GoTo 0
    Else
        GlobalFree PrintDlg.hDevMode
        GlobalFree PrintDlg.hDevNames
        bCancel = True
        p_ExtendedError = CommDlgExtendedError
        If CancelError Then Err.Raise cdlCancel, "PrintDialog", "Cancel."
    End If
    ShowPrinter = Not bCancel
End Function
#End If
