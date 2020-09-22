VERSION 5.00
Begin VB.UserControl TextScroll 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   2760
   ScaleWidth      =   1935
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   2160
   End
   Begin VB.PictureBox PictBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.Label LblCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -60
         TabIndex        =   2
         Top             =   360
         Width           =   1365
         WordWrap        =   -1  'True
      End
      Begin VB.Label LblBorder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "TextScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Glass TextScroll
' Nome del File ..: TextScroll
' Data............: 27/10/2007
' Versione........: 0.90 Beta
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte

' AlphaBlend
Private Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal WidthSrc As Long, _
  ByVal HeightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean
  
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type
Private tProperties As typeBlendProperties
Private lngBlend As Long
Private m_left As Long
Private m_Top As Long
Private m_Width As Long
Private m_Height As Long
Private hDCSrc As Long
Private hDCDst As Long

' SetGlass
Private Type vRGB
  r As Byte
  G As Byte
  B As Byte
End Type
Private iRGB As vRGB
'
Enum GlsTxt_Style
    GlsTxtStyle_x1 = 0
    GlsTxtStyle_x2 = 1
    GlsTxtStyle_x3 = 2
    GlsTxtStyle_x4 = 3
End Enum
Private m_Style As GlsTxt_Style

Enum GlsTxt_Status
    GlsTxtStop = 0
    GlsTxtStart = 1
    GlsTxtPause = 2
End Enum
Private m_Status As GlsTxt_Status

Private i As Integer
Private m_Caption As String
Private m_ColorBack As OLE_COLOR
Private m_Transparency As Byte
Private m_GlassEffect As Long
Private m_IsDown As Long
Private m_Redraw As Boolean
' Dichiarazione Eventi
Public Event Click()
'
Private WithEvents M_Frm As Form
Attribute M_Frm.VB_VarHelpID = -1
'----------------------------------------------------------
' m_Frm Events ( For Capture Parent Events  )
'----------------------------------------------------------
Private Sub M_Frm_Resize()
 '   GoAlpha2 m_Transparency
End Sub

Private Sub M_Frm_Load()
    Timer1.Enabled = True
End Sub
'-----------------------------------------------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
     LblCaption.Caption = Empty
     m_Style = GlsTxtStyle_x1
     m_Status = GlsTxtStop
     m_ColorBack = &H404040
     m_Transparency = 160
     m_GlassEffect = True
     Set Font = Ambient.Font
     '
     UserControl.Height = 2655
     UserControl.width = 1935
End Sub
'-----------------------------------------------------------------------------------------------
' Resizing
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Resize()
Dim Wt, Ht As Long
    With PictBack
        .Left = 0
        .Top = 0
        .width = ScaleWidth
        .Height = ScaleHeight
    End With
    '
    With LblBorder
        .Height = PictBack.ScaleHeight + 2 ' - 10
        .width = PictBack.ScaleWidth - 10
        .Left = (PictBack.ScaleWidth - .width) / 2
        .Top = (PictBack.ScaleHeight - .Height) / 2
    End With
    '
    With LblCaption
        .BorderStyle = 0
        .Height = LblBorder.Height - 6
        .width = LblBorder.width - 6
        .WordWrap = True
        .AutoSize = True
        .Left = (PictBack.ScaleWidth - .width) / 2
        .Top = LblBorder.Top + LblBorder.Height
    End With
        GoAlpha2 m_Transparency
End Sub
'-----------------------------------------------------------------------------------------------
' Show
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Show()
    On Error Resume Next
    Set M_Frm = UserControl.Parent
    PictBack.BackColor = m_ColorBack
    GoAlpha2 m_Transparency
End Sub
'-----------------------------------------------------------------------------------------------
' Inizializa
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
 '   GoScroll
End Sub
'-----------------------------------------------------------------------------------------------
' Eventi
'-----------------------------------------------------------------------------------------------
Private Sub ClickEvent()
    RaiseEvent Click
End Sub
'-----------------------------------------------------------------------------------------------
' Property Let Set
'-----------------------------------------------------------------------------------------------
Public Property Get Font() As StdFont
    Set Font = LblCaption.Font
End Property
Public Property Set Font(ByVal new_Font As StdFont)
    Set LblCaption.Font = new_Font
    PropertyChanged "Font"
End Property
' ColorText
Public Property Get ColorText() As OLE_COLOR
   ColorText = LblCaption.ForeColor
End Property
Public Property Let ColorText(ByVal NewValue As OLE_COLOR)
   PropertyChanged "ColorText"
   LblCaption.ForeColor = NewValue
End Property
'
Public Property Get Status() As GlsTxt_Status
   Status = m_Status
End Property
Public Property Let Status(NewValue As GlsTxt_Status)
    m_Status = NewValue
    PropertyChanged "Status"
    Select Case m_Status
        Case GlsTxtStart
            Timer1.Enabled = True
        Case GlsTxtStop
            Timer1.Enabled = False
        Case GlsTxtPause
            Timer1.Enabled = False
    End Select
End Property
'
Public Property Get Caption() As String
   Caption = m_Caption
End Property
Public Property Let Caption(NewValue As String)
    m_Caption = NewValue
    PropertyChanged "Caption"
    LblCaption.Caption = m_Caption
End Property
' Style
Public Property Get Style() As GlsTxt_Style
   Style = m_Style
End Property
Public Property Let Style(ByVal NewValue As GlsTxt_Style)
   m_Style = NewValue
   PropertyChanged "Style"
   GoAlpha2 m_Transparency
End Property
' ColorBack
Public Property Get ColorBack() As OLE_COLOR
   ColorBack = m_ColorBack
End Property
Public Property Let ColorBack(ByVal NewValue As OLE_COLOR)
   m_ColorBack = NewValue
   PropertyChanged "ColorBack"
   PictBack.BackColor = m_ColorBack
   GoAlpha2 m_Transparency
End Property
' Transparency
Public Property Let Transparency(bTransparency As Byte)
    m_Transparency = bTransparency
    PropertyChanged "Transparency"
'
    GoAlpha2 m_Transparency
End Property
Public Property Get Transparency() As Byte
    Transparency = m_Transparency
End Property
' GlassEffect
Public Property Let GlassEffect(bGlassEffect As Boolean)
    m_GlassEffect = bGlassEffect
    PropertyChanged "GlassEffect"
'
    GoAlpha2 m_Transparency
End Property
Public Property Get GlassEffect() As Boolean
    GlassEffect = m_GlassEffect
End Property
'-----------------------------------------------------------------------------------------------
' Property Read Write
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Transparency", m_Transparency, 160
        .WriteProperty "GlassEffect", m_GlassEffect, True
        .WriteProperty "Caption", m_Caption, Empty
        .WriteProperty "Style", m_Style, 0
        .WriteProperty "Status", m_Status, 0
        .WriteProperty "ColorBack", m_ColorBack, &H404040
        .WriteProperty "ColorText", ColorText, &H0
        Call .WriteProperty("Font", Font, Ambient.Font)
    End With
End Sub
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Transparency = .ReadProperty("Transparency", 160)
        m_GlassEffect = .ReadProperty("GlassEffect", True)
        m_Caption = .ReadProperty("Caption", Empty)
        m_ColorBack = .ReadProperty("ColorBack", &H404040)
        m_Style = .ReadProperty("Style", 0)
        m_Status = .ReadProperty("Status", 0)
        LblCaption.ForeColor = .ReadProperty("ColorText", &H0)
        Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'
'
' Inizio Routine GlassButton
'
'
'-----------------------------------------------------------------------------------------------
Private Sub Timer1_Timer()
    With LblCaption
        If .Top + .Height < LblBorder.Top Then .Top = LblBorder.Top + LblBorder.Height
        .Top = .Top - 2
    End With
'     DoEvents
End Sub
'------------------------------------------------------------------------
' MouseEvents
'------------------------------------------------------------------------
Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_IsDown = True
End Sub
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_IsDown = False
    ClickEvent
End Sub
'------------------------------------------------------------------------
' AlphaBlend 2 nota bene Autoredraw deve essere disattivato
'------------------------------------------------------------------------
Private Sub GoAlpha2(Alpha As Byte)
    On Error Resume Next
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    
    PictBack.FillColor = m_ColorBack
'    PictBack.Cls
    tProperties.tBlendAmount = Alpha  'Set translucency level
    
    m_Width = UserControl.width / Screen.TwipsPerPixelX
    m_Height = UserControl.Height / Screen.TwipsPerPixelY
    m_left = UserControl.Extender.Left / Screen.TwipsPerPixelX
    m_Top = UserControl.Extender.Top / Screen.TwipsPerPixelY
    '
    hDCSrc = UserControl.Parent.hdc
    hDCDst = PictBack.hdc
    '
    CopyMemory lngBlend, tProperties, 4 'Blend colors
    AlphaBlend hDCDst, 0, 0, m_Width, m_Height, hDCSrc, m_left, m_Top, m_Width, m_Height, lngBlend 'Blend together
    '
    UserControl.Parent.AutoRedraw = m_Redraw ' Ripristina
    '
    If m_GlassEffect = True Then SetGlass PictBack ' Glass Effect
End Sub
'------------------------------------------------------------------------
' Glass Tipo 1
'------------------------------------------------------------------------
Private Sub SetGlass(My_Obj As Object)
Dim Sta As Integer
Dim m_X As Long
Dim m_Y As Long
Dim m_V As Long
Dim m_T As Long
Dim m_H As Long
Dim m_X1 As Long
Dim m_Y1 As Long
Dim m_X2 As Long
Dim m_Y2 As Long
    
    m_X1 = 0
    m_Y1 = 0
  
    m_X2 = My_Obj.width / Screen.TwipsPerPixelX
    m_Y2 = My_Obj.Height / Screen.TwipsPerPixelY
    '
    m_H = (m_Y2 / 100) * 20
    If m_H > 15 Then m_H = 15

' Parte Alta
    
    For m_Y = m_H To 0 Step -1
        
        Sta = (m_H - m_Y) * 6
        
        For m_X = m_X1 To m_X1 + m_X2
            m_V = GetPixel(My_Obj.hdc, m_X, m_Y1 + m_Y)

            CopyMemory iRGB, m_V, LenB(iRGB)

            m_V = iRGB.r + Sta
            If m_V > 255 Then m_V = 255
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G + Sta
            If m_V > 255 Then m_V = 255
            iRGB.G = Int(m_V)
                     
            m_V = iRGB.B + Sta
            If m_V > 255 Then m_V = 255
            iRGB.B = m_V
'
            SetPixel My_Obj.hdc, m_X, m_Y1 + m_Y, RGB(iRGB.r, iRGB.G, iRGB.B)
        Next
    Next
    DoEvents
' Parte Bassa
    m_T = m_H
    For m_Y = m_Y2 - m_H To m_Y2
        Sta = ((m_H - m_T) * 6)
        For m_X = m_X1 To m_X1 + m_X2
            m_V = GetPixel(My_Obj.hdc, m_X, m_Y1 + m_Y)
            
            CopyMemory iRGB, m_V, LenB(iRGB)
        
            m_V = iRGB.r - Sta
            If m_V < 0 Then m_V = 0
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G - Sta
            If m_V < 0 Then m_V = 0
            iRGB.G = Int(m_V)
                        
            m_V = iRGB.B - Sta
            If m_V < 0 Then m_V = 0
            iRGB.B = m_V
        
'            SetPixel My_Obj.hdc, m_X, m_Y1 + m_Y, RGB(iRGB.r, iRGB.G, iRGB.B)
             SetPixelV My_Obj.hdc, m_X, m_Y1 + m_Y, RGB(iRGB.r, iRGB.G, iRGB.B)
        Next
        m_T = m_T - 1
    Next
    DoEvents
    My_Obj.Refresh
End Sub
