VERSION 5.00
Begin VB.UserControl HSlider 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ScaleHeight     =   1950
   ScaleWidth      =   3180
   ToolboxBitmap   =   "HSlider.ctx":0000
   Begin VB.PictureBox SliderBack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      Picture         =   "HSlider.ctx":0312
      ScaleHeight     =   255
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Image ImgCur 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      Picture         =   "HSlider.ctx":1B82
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Image ImgCur_S 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   240
      Picture         =   "HSlider.ctx":1F4A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   225
   End
   Begin VB.Image ImgCur_A 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   720
      Picture         =   "HSlider.ctx":2312
      Stretch         =   -1  'True
      Top             =   480
      Width           =   225
   End
   Begin VB.Image ImgCur_S 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   240
      Picture         =   "HSlider.ctx":26DA
      Stretch         =   -1  'True
      Top             =   840
      Width           =   225
   End
   Begin VB.Image ImgCur_A 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   720
      Picture         =   "HSlider.ctx":29EE
      Stretch         =   -1  'True
      Top             =   840
      Width           =   225
   End
   Begin VB.Image ImgCur_S 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   240
      Picture         =   "HSlider.ctx":2D02
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   225
   End
   Begin VB.Image ImgCur_A 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   720
      Picture         =   "HSlider.ctx":3016
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   225
   End
   Begin VB.Image ImgCur_S 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   240
      Picture         =   "HSlider.ctx":332A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   225
   End
   Begin VB.Image ImgCur_A 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   720
      Picture         =   "HSlider.ctx":363E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   225
   End
End
Attribute VB_Name = "HSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Horizontal Slider
' Nome del File ..: HSLIDER
' Data............: 27/11/2004
' Versione........: 1.31
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
Option Explicit
'
Private M_Value As Long
Private M_MinValue As Long
Private M_MaxValue As Long
'
Private CursRaporto As Double
'Private CursRange As Long
Private CursRange As Double
Private CursBlk As Boolean
Private Cur_Stato As Boolean
Private Flag As Boolean
'
Enum HscrollStyle
    Glass
    green
    red
    Yellow
End Enum
Private m_Style As HscrollStyle

' Dichiarazione Eventi
Public Event Change(Value As Long)
Public Event Scroll(Value As Long)
'-----------------------------------------------------------------
'                      AmbientChanged
'-----------------------------------------------------------------
Private Sub UserControl_AmbientChanged(PropertyName As String)
     Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Sub
'-----------------------------------------------------------------
'      Inizializza le Variabili ( Solo Progetazione )
'-----------------------------------------------------------------
Private Sub UserControl_InitProperties()
     M_Value = 0                   ' Valore Iniziale
     M_MinValue = 0                ' Valore Iniziale
     M_MaxValue = 10               ' Valore Iniziale
     UserControl.Height = 255      ' Altezza
     UserControl.width = 1830      ' Larghezza
     Flag = Not Cur_Stato
End Sub
'-----------------------------------------------------------------
'                      Inizializza
'-----------------------------------------------------------------
Private Sub UserControl_Initialize()
     Cur_Stato = False
End Sub
'-----------------------------------------------------------------
'                      Show
'-----------------------------------------------------------------
Private Sub UserControl_Show()
    Init_Style
End Sub
'-----------------------------------------------------------
' Resize
'-----------------------------------------------------------
Private Sub UserControl_Resize()
Dim X1, Y1, X2, Y2, L1, H1, L2, H2 As Long
    '  On Error Resume Next
    ImgCur.Left = 0
    ImgCur.Height = ScaleHeight
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
    '
    X1 = 0 'TitleLeft.width
    Y1 = 0 '300 ' AAAAA
    L1 = UserControl.width '* (UserControl.ScaleWidth)   'Risultato Destinazione
    H1 = UserControl.Height                             'Risultato Destinazione
    X2 = 0
    Y2 = 0
    L2 = SliderBack.width
    H2 = SliderBack.Height
    '
    UserControl.PaintPicture SliderBack.Image, X1, Y1, L1, H1, X2, Y2, L2, H2, vbSrcCopy
End Sub
'-----------------------------------------------------------------
'
'                                Property
'
'-----------------------------------------------------------------
Public Property Let Style(xVal As HscrollStyle)
    If xVal <> m_Style Then
        m_Style = xVal
        PropertyChanged "Style"
        Init_Style
    End If
End Property
Public Property Get Style() As HscrollStyle
    Style = m_Style
End Property
'
Public Property Get Value() As Long
   Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
   
'   If NewValue = M_Value Then Exit Property
   
   If NewValue > M_MaxValue Then NewValue = M_MaxValue
   If NewValue < M_MinValue Then NewValue = M_MinValue
   
   M_Value = NewValue
   PropertyChanged "Value"
   ChangeEvent Value
   ScrollEvent Value
   Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Property
'
Public Property Get MinValue() As Long
   MinValue = M_MinValue
End Property
Public Property Let MinValue(ByVal NewValue As Long)
    M_MinValue = NewValue
    PropertyChanged "MinValue"
    If M_Value < M_MinValue Then
        M_Value = M_MinValue
        PropertyChanged "Value"
    End If
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
'
Public Property Get MaxValue() As Long
   MaxValue = M_MaxValue
End Property
Public Property Let MaxValue(ByVal NewValue As Long)
   M_MaxValue = NewValue
   PropertyChanged "MaxValue"
   CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
'
Public Property Get Picture() As Picture
   Set Picture = SliderBack.Picture
End Property

Public Property Set Picture(ByVal NewPic As Picture)
   Set SliderBack.Picture = NewPic
   PropertyChanged "Picture"
End Property
'
'
Public Property Get PictureCursor() As Picture
   Set PictureCursor = ImgCur_S(0).Picture
End Property

Public Property Set PictureCursor(ByVal NewPic As Picture)
   Set ImgCur_S(0).Picture = NewPic
   PropertyChanged "PictureCursor"
   Set ImgCur.Picture = ImgCur_S(0).Picture
End Property
'
Public Property Get PicCursor_Selected() As Picture
   Set PicCursor_Selected = ImgCur_A(0).Picture
End Property

Public Property Set PicCursor_Selected(ByVal NewPic As Picture)
   Set ImgCur_A(0).Picture = NewPic
   PropertyChanged "PicCursor_Selected"
End Property
'-----------------------------------------------------------------
'                 Read/Write Properties
'-----------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        M_Value = .ReadProperty("Value", 0)
        M_MinValue = .ReadProperty("MinValue", 0)
        M_MaxValue = .ReadProperty("MaxValue", 10)
        m_Style = .ReadProperty("Style", 0)
    End With
  '
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
    Call Sposta((M_Value - M_MinValue) * CursRaporto)
    
    Set SliderBack.Picture = PropBag.ReadProperty("Picture", Nothing)
    Set ImgCur_A(0).Picture = PropBag.ReadProperty("PicCursor_Selected", Nothing)
    Set ImgCur_S(0).Picture = PropBag.ReadProperty("PictureCursor", Nothing)
    Set ImgCur.Picture = ImgCur_S(0).Picture
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Value", M_Value, 0
        .WriteProperty "MinValue", M_MinValue, 0
        .WriteProperty "MaxValue", M_MaxValue, 10
        .WriteProperty "Picture", SliderBack.Picture, Nothing
        .WriteProperty "PicCursor_Selected", ImgCur_A(0).Picture, Nothing
        .WriteProperty "PictureCursor", ImgCur_S(0).Picture, Nothing
        .WriteProperty "Style", m_Style, 0
    End With
End Sub
'-----------------------------------------------------------------
'                        Eventi
'-----------------------------------------------------------------
Private Sub ChangeEvent(Valore As Long)
    RaiseEvent Change(Valore)
End Sub
Private Sub ScrollEvent(Valore As Long)
    RaiseEvent Scroll(Valore)
End Sub

'-----------------------------------------------------------------
'
'                        Inizio
'
'-----------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Inizializa Style
'---------------------------------------------------------------------------------------
Private Sub Init_Style()
    ImgCur = ImgCur_S(m_Style)
End Sub
'
Public Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CursBlk = True
    Cur_Stato = True
    Cursore (Cur_Stato)
End Sub
'
Public Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Usercontrol_MouseMove(Button, Shift, X, Y)
    CursBlk = False
    Cur_Stato = False
    Cursore (Cur_Stato)
    Call ChangeEvent(Value)
End Sub
'
Public Sub Usercontrol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MaxDX As Long
Dim MinSX As Long
 '
    If CursRaporto = 0 Then Exit Sub    ' AAAA Bloca in Caso Sia Zero
 '
    If CursBlk = False Then Exit Sub
    MaxDX = ScaleWidth - (ImgCur.width / 2)
    MinSX = (ImgCur.width / 2)
    Select Case X
        Case Is < MinSX              ' Minimo
            ImgCur.Left = 0
            M_Value = M_MinValue
            GoTo SetValue
        Case Is > MaxDX              ' Massimo
            ImgCur.Left = ScaleWidth - ImgCur.width
            M_Value = M_MaxValue
            GoTo SetValue
    End Select

    Call Sposta(X - MinSX)
    M_Value = (ImgCur.Left / CursRaporto) + M_MinValue

SetValue:
    Call ScrollEvent(Value)
End Sub
'
Private Sub Sposta(Posizione As Long)
    If CursRaporto = 0 Then Exit Sub    ' AAAA Bloca in Caso Sia Zero
    ImgCur.Left = Posizione
End Sub
'
Private Function Raporto(Min As Long, Max As Long) As Double
    CursRange = Max - Min
    If CursRange <> 0 Then
        Raporto = (ScaleWidth - ImgCur.width) / CursRange
    Else
        Raporto = 0
    End If
End Function
'
Private Sub Cursore(Stato As Boolean) ' Corrected
    If Cur_Stato <> Flag Then
        Flag = Cur_Stato
        If Stato Then
            Set ImgCur.Picture = ImgCur_A(m_Style).Picture
        Else
            Set ImgCur.Picture = ImgCur_S(m_Style).Picture
        End If
    End If

End Sub
