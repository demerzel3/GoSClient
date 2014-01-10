VERSION 5.00
Begin VB.UserControl uScroller 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.VScrollBar vscroll 
      Height          =   2565
      LargeChange     =   10
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "uScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Scroll()
Event Change()

Private Const MAXVALUE As Integer = 32767

Private mMax As Long
Private mValue As Long

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
    Enabled = vscroll.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    vscroll.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un oggetto."
    vscroll.Refresh
End Sub

Private Sub UserControl_Initialize()
    vscroll.Max = MAXVALUE
    mMax = MAXVALUE
End Sub

Private Sub UserControl_Resize()
    vscroll.Width = UserControl.ScaleWidth
    vscroll.Height = UserControl.ScaleHeight
End Sub

Private Sub vscroll_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub vscroll_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub vscroll_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get Value() As Long
Attribute Value.VB_Description = "Restituisce o imposta il valore di un oggetto."
    Value = mValue
End Property

Public Property Let Value(ByVal New_Value As Long)
    mValue = New_Value
    SetValue
    'vscroll.Value() = New_Value
    'vscroll.Value() = (New_Value * MAXVALUE) / mMax
    'PropertyChanged "Value"
End Property

Private Sub vscroll_Scroll()
    If mMax > MAXVALUE Then
        mValue = (vscroll.Value * CDec(mMax)) / MAXVALUE
    Else
        mValue = vscroll.Value
    End If
    RaiseEvent Scroll
End Sub

Private Sub vscroll_Change()
    If mMax > MAXVALUE Then
        mValue = (vscroll.Value * CDec(mMax)) / MAXVALUE
    Else
        mValue = vscroll.Value
    End If
    RaiseEvent Change
End Sub

Public Property Get Max() As Long
Attribute Max.VB_Description = "Restituisce o imposta il valore massimo per l'impostazione della proprietà Value relativa alla posizione di una barra di scorrimento."
    Max = mMax
End Property

Private Sub SetValue()
    Dim NewValue As Long
    
    If mMax <= MAXVALUE Then
        NewValue = mValue
    Else
        NewValue = (CDec(mValue) * MAXVALUE) / mMax
    End If
    
    If NewValue > vscroll.Max Then
        vscroll.Value = vscroll.Max
    ElseIf NewValue < vscroll.Min Then
        vscroll.Value = vscroll.Min
    Else
        vscroll.Value = NewValue
    End If
End Sub

Public Property Let Max(ByVal New_Max As Long)
    If New_Max <= MAXVALUE Then
        vscroll.Max = New_Max
    Else
        vscroll.Max = MAXVALUE
    End If
    
    mMax = New_Max
    SetValue
    'vscroll.Value() = (mValue * MAXVALUE) / mMax
    'vscroll.Max() = New_Max
    'PropertyChanged "Max"
End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Restituisce o imposta il valore massimo per l'impostazione della proprietà Value relativa alla posizione di una barra di scorrimento."
    Min = vscroll.Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    vscroll.Min() = New_Min
    'PropertyChanged "Min"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    vscroll.Enabled = PropBag.ReadProperty("Enabled", True)
    mValue = PropBag.ReadProperty("Value", 0)
    mMax = PropBag.ReadProperty("Max", 32767)
    vscroll.Min = PropBag.ReadProperty("Min", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", vscroll.Enabled, True)
    Call PropBag.WriteProperty("Value", mValue, 0)
    Call PropBag.WriteProperty("Max", mMax, MAXVALUE)
    Call PropBag.WriteProperty("Min", vscroll.Min, 0)
End Sub
