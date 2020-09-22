VERSION 5.00
Begin VB.UserControl ProgressBar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   ScaleHeight     =   1800
   ScaleWidth      =   3660
   ToolboxBitmap   =   "ProgressBar.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   1320
   End
   Begin VB.PictureBox PictCur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFEAD1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.PictureBox PictBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFEAD1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: ProgressBar
' Nome del File ..: ProgressBar
' Data............: 18/10/2007
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
Private M_Value As Long
Private M_MinValue As Long
Private M_MaxValue As Long
Private m_ViewValue As Boolean
Private m_Transparency As Byte
Private m_TrspCur As Byte
Private m_GlassEffect As Long
Private m_IsHot As Boolean
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_Redraw As Boolean
'
Private CursRaporto As Double
Private CursRange As Long
'
Private WithEvents M_Frm As Form
Attribute M_Frm.VB_VarHelpID = -1
'                                Dichiarazione Eventi
Public Event Change(Value As Long)
'-------------------------------------------------------------
' Eventi
'-------------------------------------------------------------
Private Sub ChangeEvent(Valore As Long)
    RaiseEvent Change(Valore)
End Sub
'----------------------------------------------------------
' m_Frm Events ( For Capture Parent Events  )
'----------------------------------------------------------
Private Sub M_Frm_Resize()
    GoAlpha2 m_Transparency
End Sub
Private Sub M_Frm_activate()
    Timer1.Enabled = True
End Sub
Private Sub M_Frm_show()
'    GoAlpha2 m_Transparency
End Sub
'-------------------------------------------------------------
' Inizializza
'-------------------------------------------------------------
Private Sub UserControl_Initialize()
'     Call Sposta((m_Value - M_MinValue) * CursRaporto)
End Sub
'-------------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'-------------------------------------------------------------
Private Sub UserControl_InitProperties()
    M_Value = 0                   ' Valore Iniziale
    M_MinValue = 0                ' Valore Iniziale
    M_MaxValue = 100              ' Valore Iniziale
    m_ViewValue = True            ' Visualizza Numero
    m_Transparency = 200
    m_TrspCur = 200
    m_GlassEffect = True
    Timer1.Enabled = False
    UserControl.Height = 255      ' Altezza
    UserControl.width = 1830      ' Larghezza
End Sub
'-------------------------------------------------------------
' Resizing
'-------------------------------------------------------------
Private Sub UserControl_Resize()
    With PictBack
        .Left = 0
        .Top = 0
        .width = ScaleWidth
        .Height = ScaleHeight
    End With
    
    With PictCur
        .Left = 0
        .Top = 0
        .width = ScaleWidth
        .Height = ScaleHeight
    End With

    PictCur.Height = PictBack.Height - (PictBack.Height / 100) * 40
    PictCur.Top = (PictBack.Height - PictCur.Height) / 2
        
   If Not Ambient.UserMode Then
        GoAlpha2 m_Transparency
    End If
End Sub
'-----------------------------------------------------------------------------------------------
' Show
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Show()
    On Error Resume Next
    Set M_Frm = UserControl.Parent
    PictBack.BackColor = m_BackColor
    PictCur.BackColor = m_BackColor
   If Not Ambient.UserMode Then
        GoAlpha2 m_Transparency
    End If
    Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Sub
'-------------------------------------------------------------
' Property
'-------------------------------------------------------------
' Value
Public Property Get Value() As Long
    Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
   
    If NewValue > M_MaxValue Then NewValue = M_MaxValue
    If NewValue < M_MinValue Then NewValue = M_MinValue
   
    M_Value = NewValue
    PropertyChanged "Value"
    Call Sposta((M_Value - M_MinValue) * CursRaporto)
    If Ambient.UserMode = True Then
        Call Timer1_Timer
    End If
End Property
' MinValue
Public Property Get MinValue() As Long
    MinValue = M_MinValue
End Property
Public Property Let MinValue(ByVal NewValue As Long)
    M_MinValue = NewValue
    PropertyChanged "MinValue"
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
' MaxValue
Public Property Get MaxValue() As Long
    MaxValue = M_MaxValue
End Property
Public Property Let MaxValue(ByVal NewValue As Long)
    M_MaxValue = NewValue
    PropertyChanged "MaxValue"
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
End Property
' PictureBackG
Public Property Get PictureBackG() As Picture
    Set PictureBackG = PictBack.Picture
End Property
Public Property Set PictureBackG(ByVal NewPic As Picture)
    Set PictBack.Picture = NewPic
    Set PictCur.Picture = NewPic
    PropertyChanged "PictureBackG"
End Property
' Transparency
Public Property Let Transparency(bTransparency As Byte)
    m_Transparency = bTransparency
    If m_Transparency > 230 Then
        m_TrspCur = 230
    Else
        m_TrspCur = m_Transparency
    End If
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
' BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    PropertyChanged "BackColor"
    PictBack.BackColor = m_BackColor
    PictCur.BackColor = m_BackColor
    GoAlpha2 m_Transparency
End Property
' ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    PropertyChanged "ForeColor"
    GoAlpha2 m_Transparency
End Property
'-------------------------------------------------------------
' Read/Write Properties
'-------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Transparency = .ReadProperty("Transparency", 200)
        m_GlassEffect = .ReadProperty("GlassEffect", True)
        m_BackColor = .ReadProperty("BackColor", &HFFEAD1)
        m_ForeColor = .ReadProperty("ForeColor", &HFF)
        M_Value = .ReadProperty("Value", 0)
        M_MinValue = .ReadProperty("MinValue", 0)
        M_MaxValue = .ReadProperty("MaxValue", 100)
        m_ViewValue = .ReadProperty("ViewValue", True)
        '
        Set PictBack.Picture = .ReadProperty("PictureBackG", Nothing)
    End With
    
    If m_Transparency > 230 Then
        m_TrspCur = 230
    Else
        m_TrspCur = m_Transparency
    End If
    
    CursRaporto = Raporto(M_MinValue, M_MaxValue)
    Call Sposta((M_Value - M_MinValue) * CursRaporto)
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Transparency", m_Transparency, 200
        .WriteProperty "GlassEffect", m_GlassEffect, True
        .WriteProperty "BackColor", m_BackColor, &HFFEAD1
        .WriteProperty "Forecolor", m_ForeColor, &HFF
        .WriteProperty "Value", M_Value, 0
        .WriteProperty "MinValue", M_MinValue, 0
        .WriteProperty "MaxValue", M_MaxValue, 100
        .WriteProperty "PictureForG", PictCur.Picture, Nothing
    End With
End Sub
'-------------------------------------------------------------
' Sposta
'-------------------------------------------------------------
Private Sub Sposta(Posizione As Long)
    PictCur.width = Posizione
End Sub
'-------------------------------------------------------------
' Rapporto
'-------------------------------------------------------------
Private Function Raporto(Min As Long, Max As Long) As Single
    CursRange = Max - Min
    Raporto = UserControl.ScaleWidth / CursRange
End Function
'------------------------------------------------------------------------
' AlphaBlend 1 nota bene Autoredraw deve essere Attivato
'------------------------------------------------------------------------
Public Sub GoAlpha1(M_Obj As Object, Alpha As Byte, m_left As Long, m_Top As Long, m_Width As Long, m_Height As Long)
    Timer1.Enabled = False
    On Error Resume Next
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    '
    tProperties.tBlendAmount = Alpha  'Set translucency level
    '
    hDCSrc = UserControl.Parent.hdc
    hDCDst = M_Obj.hdc
    '
    CopyMemory lngBlend, tProperties, 4 'Blend colors
'    AlphaBlend hDCDst, m_left, m_Top, m_Width, m_Height, hDCSrc,( m_left + UserControl.Extender.Left )/ Screen.TwipsPerPixelX,( m_Top + UserControl.Extender.Top )/ Screen.TwipsPerPixelY, m_Width, m_Height, lngBlend 'Blend together
    AlphaBlend hDCDst, m_left, m_Top, m_Width, m_Height, hDCSrc, (m_left + UserControl.Extender.Left) / Screen.TwipsPerPixelX, (PictCur.Top + UserControl.Extender.Top) / Screen.TwipsPerPixelY, m_Width, m_Height, lngBlend  'Blend together
    '
    UserControl.Parent.AutoRedraw = m_Redraw ' Ripristina
    '
    Timer1.Enabled = True
End Sub
'------------------------------------------------------------------------
' AlphaBlend 2 nota bene Autoredraw deve essere Attivato
'------------------------------------------------------------------------
Private Sub GoAlpha2(Alpha As Byte)
    Timer1.Enabled = False
    On Error Resume Next
    PictBack.Cls
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    
    PictBack.FillColor = m_BackColor

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
    
    If Ambient.UserMode = True Then
        Timer1.Enabled = True
    End If
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
'------------------------------------------------------------------------
' Timer1
'------------------------------------------------------------------------
Private Sub Timer1_Timer()
Dim Stp As Double
Dim Y As Double
    
    PictCur.Cls
    PictCur.DrawWidth = 25
    Stp = 0.05
        For Y = 0 To PictCur.ScaleHeight Step Stp
        PictCur.PSet (PictCur.ScaleWidth * Rnd * -Tan(Cos(Sin(Y))) + (PictCur.ScaleWidth), Y / 3), m_ForeColor
        PictCur.PSet (PictCur.ScaleWidth * Rnd * Tan(Cos(Sin(Y))) - (PictCur.ScaleWidth), Y), m_BackColor    ' RGB(R, G, B)
    Next Y
    
    Call GoAlpha1(PictCur, m_TrspCur, 0, 0, PictCur.ScaleWidth, PictCur.ScaleHeight)
    
End Sub
