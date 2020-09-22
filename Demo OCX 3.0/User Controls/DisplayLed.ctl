VERSION 5.00
Begin VB.UserControl DisplayLed 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   ScaleHeight     =   525
   ScaleWidth      =   525
   ToolboxBitmap   =   "DisplayLed.ctx":0000
   Begin VB.PictureBox PictBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "DisplayLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Display Led
' Nome del File ..: DisplayLed
' Data............: 27/11/2004
' Versione........: 1.41 beta
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
Enum DspLd_Zoom
    Zoom_x1 = 1
    Zoom_x2 = 2
    Zoom_x3 = 3
    Zoom_x4 = 4
    Zoom_x5 = 5
    Zoom_x6 = 6
    Zoom_x7 = 7
    Zoom_x8 = 8
    Zoom_x9 = 9
End Enum
Private m_Zoom As DspLd_Zoom

Enum DspLd_ScrlStyle
    none = 1
    RightToLeft = 2
    LeftToRight = 3
    DownToUp = 4
    UpToDown = 5
End Enum
Private m_ScrollStyle As DspLd_ScrlStyle

Enum DspLd_Style
    Style_x1 = 0
    Style_x2 = 1
    Style_x3 = 2
    Style_x4 = 3
    Style_x5 = 4
    Style_x6 = 5
    Style_x7 = 6
    Style_x8 = 7
    Style_x9 = 8
    Style_x10 = 9
End Enum
Private m_Style As DspLd_Style

Private i As Integer
Private I1 As Integer
Private Matrice(255, 4) As Byte
Private Matr_V(6) As Byte
Private m_Redraw As Boolean

Private M_Value As Integer
Private m_LedColor As OLE_COLOR
Private m_LedColorLow As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_ColorLed_H(9) As Long
Private m_ColorLed_L(9) As Long
Private m_Transparency As Byte
Private m_GlassEffect As Long
'                                Dichiarazione Eventi
Public Event Change(Value As Integer)
'
Private WithEvents M_Frm As Form
Attribute M_Frm.VB_VarHelpID = -1
'----------------------------------------------------------
' m_Frm Events ( For Capture Parent Events  )
'----------------------------------------------------------
Private Sub M_Frm_Resize()
    CaricaFondo m_Zoom, m_Style
    GoAlpha2 m_Transparency
    Scrive M_Value
End Sub
'-----------------------------------------------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
     M_Value = 0
     m_Zoom = Zoom_x1
     m_ScrollStyle = none
     m_Style = Style_x1
     m_LedColor = m_ColorLed_H(m_Style)
     
     UserControl.Height = 330
     UserControl.width = 240
End Sub
'-----------------------------------------------------------------------------------------------
' Resizing
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Resize() ' Corrected
    With UserControl
        .Height = 330 * m_Zoom
        .width = 240 * m_Zoom
    End With
    PictBack.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
'-----------------------------------------------------------------------------------------------
' Show
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Show()
    On Error Resume Next
    Set M_Frm = UserControl.Parent
    If Not Ambient.UserMode Then
        GoAlpha2 m_Transparency
    End If
End Sub
'-----------------------------------------------------------------------------------------------
' Inizializa
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
  UserControl.Height = 330 * m_Zoom
  UserControl.width = 240 * m_Zoom
  Call LeggeMatrici
  m_ColorLed_H(0) = &HFFFFFF
  m_ColorLed_L(0) = &H404040
  m_ColorLed_H(1) = RGB(168, 255, 0)
  m_ColorLed_L(1) = &HFF00FF
  m_ColorLed_H(2) = &H404040
  m_ColorLed_L(2) = &HFFFFFF
  m_ColorLed_H(3) = &HFF00&
  m_ColorLed_L(3) = &HC000&
  m_ColorLed_H(4) = &HFFFFFF
  m_ColorLed_L(4) = &HFF0000
  m_ColorLed_H(5) = &HFFFFFF
  m_ColorLed_L(5) = &H9B8B40
  m_ColorLed_H(6) = &H404040
  m_ColorLed_L(6) = &H70B7C6 ' &H408B9B

End Sub
'-----------------------------------------------------------------------------------------------
' Eventi
'-----------------------------------------------------------------------------------------------
Private Sub ChangeEvent(Valore As Integer)
    RaiseEvent Change(Valore)
End Sub
'-----------------------------------------------------------------------------------------------
' Property Let Set
'-----------------------------------------------------------------------------------------------
Public Property Get Value() As Long
   Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
   PropertyChanged "Value"
   ChangeEvent Value
   If M_Value = NewValue Then Exit Property
   M_Value = NewValue
   '
   CaricaFondo m_Zoom, m_Style
   GoAlpha2 m_Transparency
   Scrive M_Value
End Property
' Zoom
Public Property Get Zoom() As DspLd_Zoom
   Zoom = m_Zoom
End Property
Public Property Let Zoom(ByVal NewValue As DspLd_Zoom)
    m_Zoom = NewValue
    PropertyChanged "Zoom"
    '
    UserControl.Height = 330 * m_Zoom
    UserControl.width = 240 * m_Zoom
    '
    CaricaFondo m_Zoom, m_Style
    GoAlpha2 m_Transparency
    Scrive M_Value
End Property
' ScrollStyle
Public Property Get ScrollStyle() As DspLd_ScrlStyle
   ScrollStyle = m_ScrollStyle
End Property
Public Property Let ScrollStyle(ByVal NewValue As DspLd_ScrlStyle)
    m_ScrollStyle = NewValue
    PropertyChanged "ScrollStyle"
    '
    CaricaFondo m_Zoom, m_Style
    GoAlpha2 m_Transparency
    Scrive M_Value
End Property
' Style
Public Property Get Style() As DspLd_Style
   Style = m_Style
End Property
Public Property Let Style(ByVal NewValue As DspLd_Style)
   m_Style = NewValue
   PropertyChanged "Style"
   m_LedColor = m_ColorLed_H(m_Style)
   m_BackColor = m_ColorLed_L(m_Style)
    '
    CaricaFondo m_Zoom, m_Style
    GoAlpha2 m_Transparency
    Scrive M_Value
End Property
' LedColor
Public Property Get LedColor() As OLE_COLOR
   LedColor = m_LedColor
End Property
Public Property Let LedColor(ByVal NewValue As OLE_COLOR)
   m_LedColor = NewValue
   PropertyChanged "LedColor"
   Call Scrive(M_Value)
End Property
' BackColor
Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
   m_BackColor = NewValue
   PropertyChanged "BackColor"
   Call Scrive(M_Value)
End Property
' Transparency
Public Property Let Transparency(bTransparency As Byte)
    m_Transparency = bTransparency
    PropertyChanged "Transparency"
'
    CaricaFondo m_Zoom, m_Style
    GoAlpha2 m_Transparency
    Scrive M_Value
End Property
Public Property Get Transparency() As Byte
    Transparency = m_Transparency
End Property
' GlassEffect
Public Property Let GlassEffect(bGlassEffect As Boolean)
    m_GlassEffect = bGlassEffect
    PropertyChanged "GlassEffect"
'
    CaricaFondo m_Zoom, m_Style
    GoAlpha2 m_Transparency
    Scrive M_Value
End Property
Public Property Get GlassEffect() As Boolean
    GlassEffect = m_GlassEffect
End Property
'-----------------------------------------------------------------------------------------------
' Property Read Write
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Transparency", m_Transparency, 0
        .WriteProperty "GlassEffect", m_GlassEffect, True
        .WriteProperty "Value", M_Value, 0
        .WriteProperty "Zoom", m_Zoom, 1
        .WriteProperty "ScrollStyle", m_ScrollStyle, 1
        .WriteProperty "LedColor", m_LedColor, RGB(168, 255, 0)
        .WriteProperty "BackColor", m_BackColor, &H404040
        .WriteProperty "Style", m_Style, 0
    End With
End Sub
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Transparency = .ReadProperty("Transparency", 0)
        m_GlassEffect = .ReadProperty("GlassEffect", True)
        M_Value = .ReadProperty("Value", 0)
        m_Zoom = .ReadProperty("Zoom", 1)
        m_ScrollStyle = .ReadProperty("ScrollStyle", 1)
        m_LedColor = .ReadProperty("LedColor", RGB(168, 255, 0))
        m_BackColor = .ReadProperty("BackColor", &H404040)
        m_Style = .ReadProperty("Style", 0)
    End With
    ' Dim M_SelectedColor As OLE_COLOR
  
    UserControl.Height = 330 * m_Zoom
    UserControl.width = 240 * m_Zoom
    CaricaFondo m_Zoom, m_Style
End Sub
'-----------------------------------------------------------------------------------------------
'
'
'         Inizio Routine DisplayLed
'
'
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
' CaricaFondo ( Load BackGround )
'-----------------------------------------------------------------------------------------------
Private Sub CaricaFondo(Zm As Long, Style As Long)
Dim Col As Long
Dim Rig As Long
Dim Brgh As Byte
Dim tmp As Long
    PictBack.BackColor = m_BackColor
 '   PictBack.Cls
    Select Case Style
        Case 0
        
        Case 1 To 9
            '
            iRGB.r = m_BackColor And &HFF&
            iRGB.G = (m_BackColor And &HFF00&) / 2 ^ 8
            iRGB.B = (m_BackColor And &HFF0000) / 2 ^ 16
            Brgh = 30
            tmp = iRGB.r: If (tmp - Brgh) > -1 Then iRGB.r = iRGB.r - Brgh
            tmp = iRGB.G: If (tmp - Brgh) > -1 Then iRGB.G = iRGB.G - Brgh
            tmp = iRGB.B: If (tmp - Brgh) > -1 Then iRGB.B = iRGB.B - Brgh
            m_LedColorLow = RGB(iRGB.r, iRGB.G, iRGB.B)
                       
            
'            For Col = 0 To 4
'                For Rig = 0 To 6
'                    Call PlotMask(Rig, Col, m_LedColorLow)
'                Next Rig
'            Next Col
    End Select
End Sub
'-----------------------------------------------------------------------------------------------
' ReverseRGB
'-----------------------------------------------------------------------------------------------
Function ReverseRGB(red, green, blue)
    ReverseRGB = CLng(blue + (green * 256) + (red * 65536))
End Function
'-----------------------------------------------------------------------------------------------
' Scrive ( Write )
'-----------------------------------------------------------------------------------------------
Private Sub Scrive(Valore As Integer)
 Dim Vrt As Long
 Dim Rig As Long
 Dim Col As Long
 Dim Nib As Integer
 '
 Dim tmp(6) As Byte     ' 4
 Dim Scr As Integer
 Dim St As Integer
 Dim ClrTmp As OLE_COLOR
 '
 If m_Style <> Style_x1 Then
        ClrTmp = m_LedColorLow
    Else
        ClrTmp = m_BackColor
 End If
 '
 Select Case ScrollStyle
 
 Case 1 ' ===============  Style 1 Standard
  
  For Col = 0 To 4
    Nib = 1
    For Rig = 0 To 6
        If (Matrice(Valore, Col) Or Nib) = Matrice(Valore, Col) Then
            Call Plot(Rig, Col, m_LedColor)
        Else
            PlotMask Rig, Col, ClrTmp
        End If
        Nib = Nib * 2
    Next Rig
  Next Col
 
 Case 2 ' ===============  Style 2 Scroll Da Destra a Sinistra

  For Scr = 0 To 4 Step 1           ' Loop Pricipale Da Destra a Sinistra
    Call Pause(0.03)                ' Ritardo
    For St = 0 To 3                 ' Loop Shift Carattere
      tmp(St) = tmp(St + 1)
    Next St                         ' Fine Loop Shift Carattere
    tmp(4) = Matrice(Valore, Scr)   '
    '
    For Col = 0 To 4                ' Loop Scrive Carattere
        Nib = 1
        For Rig = 0 To 6
            If (tmp(Col) Or Nib) = tmp(Col) Then
                Call Plot(Rig, Col, m_LedColor)
            Else
                PlotMask Rig, Col, ClrTmp
            End If
            Nib = Nib * 2
        Next Rig
    Next Col                        ' Fine Loop Scrive Carattere
  Next Scr                          ' Fine Loop Pricipale
  
Case 3 ' ===============  Style 3 Scroll Da Sinistra a Destra

 For Scr = 4 To 0 Step -1           ' Loop Pricipale Da Sinistra a Destra
    Call Pause(0.03)                ' Ritardo
    For St = 4 To 1 Step -1         ' Loop Shift Carattere
        tmp(St) = tmp(St - 1)
    Next St                        ' Fine Loop Shift Carattere
    tmp(0) = Matrice(Valore, Scr)  '
    '
    For Col = 0 To 4               ' Loop Scrive Carattere
        Nib = 1
        For Rig = 0 To 6
            If (tmp(Col) Or Nib) = tmp(Col) Then
                Call Plot(Rig, Col, m_LedColor)
            Else
                PlotMask Rig, Col, ClrTmp
            End If
            Nib = Nib * 2
        Next Rig
    Next Col                        ' Fine Loop Scrive Carattere
 Next Scr                           ' Fine Loop Pricipale
 
 Case 4 ' ===============  Style 4 Scroll Da Giu a Su
  
    Call Convert_Car(Valore)        ' Result in Matr_V(n)
    For Scr = 0 To 6 Step 1         ' Loop Pricipale Da Giu a Su
        Call Pause(0.03)            ' Ritardo
        For St = 0 To 5             ' Loop Shift Carattere
            tmp(St) = tmp(St + 1)
        Next St                     ' Fine Loop Shift Carattere
        tmp(6) = Matr_V(Scr)        '
        '
        For Rig = 0 To 6
            Nib = 1
            For Col = 0 To 4
                If (tmp(Rig) Or Nib) = tmp(Rig) Then
                    Call Plot(Rig, Col, m_LedColor)
                Else
                    PlotMask Rig, Col, ClrTmp
                End If
            Nib = Nib * 2
            Next Col
        Next Rig
    Next Scr
      
 Case 5 ' ===============  Style 5 Scroll Da Su a Giu
  
    Call Convert_Car(Valore)        ' Result in Matr_V(n)
    For Scr = 6 To 0 Step -1        ' Loop Pricipale Da Su a Giu
        Call Pause(0.03)            ' Ritardo ( Wait )
        For St = 6 To 1 Step -1     ' ---------------------
            tmp(St) = tmp(St - 1)   ' Loop Shift Carattere
        Next St                     ' ---------------------
        tmp(0) = Matr_V(Scr)        ' Put New Value
        '
        For Rig = 0 To 6
            Nib = 1
            For Col = 0 To 4
                If (tmp(Rig) Or Nib) = tmp(Rig) Then
                    Call Plot(Rig, Col, m_LedColor)
                Else
                    PlotMask Rig, Col, ClrTmp
                End If
                Nib = Nib * 2
            Next Col
       Next Rig
    Next Scr                        ' Fine Loop Principale
 End Select
End Sub
'-----------------------------------------------------------------------------------------------
' Pausa ( Delay )
'-----------------------------------------------------------------------------------------------
Private Sub Pause(Tempo As Double) ' Corrected

    Dim start As Single
    start = Timer + Tempo
    Do While Timer < start
        DoEvents
    Loop
End Sub
'-----------------------------------------------------------------------------------------------
' Ploting
'-----------------------------------------------------------------------------------------------
Private Sub Plot(Rig As Long, Col As Long, ValCol As OLE_COLOR)
 Dim Vrt As Long
 Dim Hr As Long
 Dim Lr As Long
 
    Lr = 3 * m_Zoom
    Hr = Lr * Col
 
    For i = 1 To m_Zoom * 2
        Vrt = (m_Zoom + i) - 1        ' Base 0
        Vrt = (Vrt + (Lr * Rig))
        PictBack.Line (Hr + m_Zoom, Vrt)-(Hr + Lr, Vrt), ValCol
    Next i
    '
End Sub
'-----------------------------------------------------------------------------------------------
' Plot Mask
'-----------------------------------------------------------------------------------------------
Private Sub PlotMask(Rig As Long, Col As Long, ValCol As OLE_COLOR)
Dim Vrt As Long
Dim Stp As Long
Dim Hr As Long
Dim Lr As Long
    
    Lr = 3 * m_Zoom
    Hr = Lr * Col
    Vrt = (m_Zoom + (Lr * Rig))
    Stp = (m_Zoom * 2) - 1

    PictBack.Line (Hr + m_Zoom, Vrt)-Step(Stp, Stp), ValCol, BF
    
'    Call GoAlpha1(m_Transparency, Hr + m_Zoom, Vrt, Stp, Stp) ' Effeto Figo 2 AAA
'    Call GoAlpha1(m_Transparency, Hr + m_Zoom + 1, Vrt + 1, Stp, Stp)' Effeto Figo 2 AAA
     Call GoAlpha1(m_Transparency, Hr + m_Zoom, Vrt, Stp + 1, Stp + 1)

End Sub
'-----------------------------------------------------------------------------------------------
' Make new array for Vertical shift
'-----------------------------------------------------------------------------------------------
Public Sub Convert_Car(Asc As Integer)
 Dim Nib As Byte
 Dim Nib1 As Byte
 '
 For i = 0 To 6       ' ----------------------------------
   Matr_V(i) = 0      '             Set to 0
 Next i               ' ----------------------------------
 '
    Nib = 1
 For i = 0 To 6
   Nib1 = 1
    For I1 = 0 To 4
      If (Matrice(Asc, I1) Or Nib) = Matrice(Asc, I1) Then
        Matr_V(i) = Matr_V(i) + Nib1
      End If
      Nib1 = Nib1 * 2
    Next I1
   Nib = Nib * 2
 Next i
End Sub
'-----------------------------------------------------------------------------------------------
' Matrice
'-----------------------------------------------------------------------------------------------
Private Sub LeggeMatrici()
'
'                        0 1 2 3 4
'                     01 * * * * *
'                     02 * * * * *
'                     04 * * * * *
'                     08 * * * * *
'                     16 * * * * *
'                     32 * * * * *
'                     64 * * * * *
'
'
'
'                           Vari
'
'                                   "
Matrice(34, 0) = 0
Matrice(34, 1) = 7
Matrice(34, 2) = 0
Matrice(34, 3) = 7
Matrice(34, 4) = 0
'                                   &
Matrice(38, 0) = 50
Matrice(38, 1) = 77
Matrice(38, 2) = 89
Matrice(38, 3) = 38
Matrice(38, 4) = 80
'                                   '
Matrice(39, 0) = 0
Matrice(39, 1) = 4
Matrice(39, 2) = 2
Matrice(39, 3) = 1
Matrice(39, 4) = 0
'                                   (
Matrice(40, 0) = 0
Matrice(40, 1) = 62
Matrice(40, 2) = 65
Matrice(40, 3) = 0
Matrice(40, 4) = 0
'                                   )
Matrice(41, 0) = 0
Matrice(41, 1) = 0
Matrice(41, 2) = 65
Matrice(41, 3) = 62
Matrice(41, 4) = 0
'                                   +
Matrice(43, 0) = 8
Matrice(43, 1) = 8
Matrice(43, 2) = 62
Matrice(43, 3) = 8
Matrice(43, 4) = 8
'                                   -
Matrice(45, 0) = 8
Matrice(45, 1) = 8
Matrice(45, 2) = 8
Matrice(45, 3) = 8
Matrice(45, 4) = 8
'                                   .
Matrice(46, 0) = 0
Matrice(46, 1) = 0
Matrice(46, 2) = 32
Matrice(46, 3) = 0
Matrice(46, 4) = 0
'                                   /
Matrice(47, 0) = 32
Matrice(47, 1) = 16
Matrice(47, 2) = 8
Matrice(47, 3) = 4
Matrice(47, 4) = 2
'
'                                   :
Matrice(58, 0) = 0
Matrice(58, 1) = 0
Matrice(58, 2) = 20
Matrice(58, 3) = 0
Matrice(58, 4) = 0
'                                   =
Matrice(61, 0) = 20
Matrice(61, 1) = 20
Matrice(61, 2) = 20
Matrice(61, 3) = 20
Matrice(61, 4) = 20
'                                   \
Matrice(92, 0) = 2
Matrice(92, 1) = 4
Matrice(92, 2) = 8
Matrice(92, 3) = 16
Matrice(92, 4) = 32
'
'                          Numerici
'
'                                   0
Matrice(48, 0) = 62
Matrice(48, 1) = 65
Matrice(48, 2) = 65
Matrice(48, 3) = 65
Matrice(48, 4) = 62
'                                   1
Matrice(49, 0) = 0
Matrice(49, 1) = 4
Matrice(49, 2) = 2
Matrice(49, 3) = 127
Matrice(49, 4) = 0
'                                   2
Matrice(50, 0) = 121
Matrice(50, 1) = 73
Matrice(50, 2) = 73
Matrice(50, 3) = 73
Matrice(50, 4) = 79
'                                   3
Matrice(51, 0) = 73
Matrice(51, 1) = 73
Matrice(51, 2) = 73
Matrice(51, 3) = 73
Matrice(51, 4) = 127
'                                   4
Matrice(52, 0) = 15
Matrice(52, 1) = 8
Matrice(52, 2) = 8
Matrice(52, 3) = 8
Matrice(52, 4) = 127
'                                   5
Matrice(53, 0) = 79
Matrice(53, 1) = 73
Matrice(53, 2) = 73
Matrice(53, 3) = 73
Matrice(53, 4) = 121
'                                   6
Matrice(54, 0) = 127
Matrice(54, 1) = 73
Matrice(54, 2) = 73
Matrice(54, 3) = 73
Matrice(54, 4) = 121
'                                   7
Matrice(55, 0) = 65
Matrice(55, 1) = 33
Matrice(55, 2) = 17
Matrice(55, 3) = 9
Matrice(55, 4) = 7
'                                   8
Matrice(56, 0) = 127
Matrice(56, 1) = 73
Matrice(56, 2) = 73
Matrice(56, 3) = 73
Matrice(56, 4) = 127
'                                   9
Matrice(57, 0) = 79
Matrice(57, 1) = 73
Matrice(57, 2) = 73
Matrice(57, 3) = 73
Matrice(57, 4) = 127
'
'                             AlfaNumerici
'
'                                   A
Matrice(65, 0) = 126
Matrice(65, 1) = 9
Matrice(65, 2) = 9
Matrice(65, 3) = 9
Matrice(65, 4) = 126
'                                   B
Matrice(66, 0) = 127
Matrice(66, 1) = 73
Matrice(66, 2) = 73
Matrice(66, 3) = 73
Matrice(66, 4) = 54
'                                   C
Matrice(67, 0) = 62
Matrice(67, 1) = 65
Matrice(67, 2) = 65
Matrice(67, 3) = 65
Matrice(67, 4) = 65
'                                   D
Matrice(68, 0) = 127
Matrice(68, 1) = 65
Matrice(68, 2) = 65
Matrice(68, 3) = 65
Matrice(68, 4) = 62
'                                   E
Matrice(69, 0) = 127
Matrice(69, 1) = 73
Matrice(69, 2) = 73
Matrice(69, 3) = 73
Matrice(69, 4) = 65
'                                   F
Matrice(70, 0) = 127
Matrice(70, 1) = 9
Matrice(70, 2) = 9
Matrice(70, 3) = 9
Matrice(70, 4) = 1
'                                   G
Matrice(71, 0) = 62
Matrice(71, 1) = 65
Matrice(71, 2) = 65
Matrice(71, 3) = 73
Matrice(71, 4) = 121
'                                   H
Matrice(72, 0) = 127
Matrice(72, 1) = 8
Matrice(72, 2) = 8
Matrice(72, 3) = 8
Matrice(72, 4) = 127
'                                   I
Matrice(73, 0) = 0
Matrice(73, 1) = 65
Matrice(73, 2) = 127
Matrice(73, 3) = 65
Matrice(73, 4) = 0
'                                   J
Matrice(74, 0) = 48
Matrice(74, 1) = 64
Matrice(74, 2) = 65
Matrice(74, 3) = 65
Matrice(74, 4) = 63
'                                   K
Matrice(75, 0) = 127
Matrice(75, 1) = 8
Matrice(75, 2) = 20
Matrice(75, 3) = 34
Matrice(75, 4) = 65
'                                   L
Matrice(76, 0) = 127
Matrice(76, 1) = 64
Matrice(76, 2) = 64
Matrice(76, 3) = 64
Matrice(76, 4) = 64
'                                   M
Matrice(77, 0) = 127
Matrice(77, 1) = 4
Matrice(77, 2) = 8
Matrice(77, 3) = 4
Matrice(77, 4) = 127
'                                   N
Matrice(78, 0) = 127
Matrice(78, 1) = 4
Matrice(78, 2) = 8
Matrice(78, 3) = 16
Matrice(78, 4) = 127
'                                   O
Matrice(79, 0) = 127
Matrice(79, 1) = 65
Matrice(79, 2) = 65
Matrice(79, 3) = 65
Matrice(79, 4) = 127
'                                   P
Matrice(80, 0) = 127
Matrice(80, 1) = 9
Matrice(80, 2) = 9
Matrice(80, 3) = 9
Matrice(80, 4) = 6
'                                   Q
Matrice(81, 0) = 127
Matrice(81, 1) = 65
Matrice(81, 2) = 81
Matrice(81, 3) = 33
Matrice(81, 4) = 95
'                                   R
Matrice(82, 0) = 127
Matrice(82, 1) = 9
Matrice(82, 2) = 25
Matrice(82, 3) = 41
Matrice(82, 4) = 70
'                                   S
Matrice(83, 0) = 79
Matrice(83, 1) = 73
Matrice(83, 2) = 73
Matrice(83, 3) = 73
Matrice(83, 4) = 121
'                                   T
Matrice(84, 0) = 1
Matrice(84, 1) = 1
Matrice(84, 2) = 127
Matrice(84, 3) = 1
Matrice(84, 4) = 1
'                                   U
Matrice(85, 0) = 127
Matrice(85, 1) = 64
Matrice(85, 2) = 64
Matrice(85, 3) = 64
Matrice(85, 4) = 127
'                                   V
Matrice(86, 0) = 31
Matrice(86, 1) = 32
Matrice(86, 2) = 64
Matrice(86, 3) = 32
Matrice(86, 4) = 31
'                                   W
Matrice(87, 0) = 127
Matrice(87, 1) = 32
Matrice(87, 2) = 16
Matrice(87, 3) = 32
Matrice(87, 4) = 127
'                                   X
Matrice(88, 0) = 99
Matrice(88, 1) = 20
Matrice(88, 2) = 8
Matrice(88, 3) = 20
Matrice(88, 4) = 99
'                                   Y
Matrice(89, 0) = 7
Matrice(89, 1) = 8
Matrice(89, 2) = 120
Matrice(89, 3) = 8
Matrice(89, 4) = 7
'                                   Z
Matrice(90, 0) = 97
Matrice(90, 1) = 81
Matrice(90, 2) = 73
Matrice(90, 3) = 69
Matrice(90, 4) = 67
End Sub
'------------------------------------------------------------------------
' AlphaBlend 1 nota bene Autoredraw deve essere Attivato
'------------------------------------------------------------------------
Public Sub GoAlpha1(Alpha As Byte, m_left As Long, m_Top As Long, m_Width As Long, m_Height As Long)
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    '
    tProperties.tBlendAmount = Alpha  'Set translucency level
    '
    hDCSrc = UserControl.Parent.hdc
    hDCDst = PictBack.hdc
    '
    CopyMemory lngBlend, tProperties, 4 'Blend colors
    AlphaBlend hDCDst, m_left, m_Top, m_Width, m_Height, hDCSrc, m_left + UserControl.Extender.Left / Screen.TwipsPerPixelX, m_Top + UserControl.Extender.Top / Screen.TwipsPerPixelY, m_Width, m_Height, lngBlend 'Blend together
    '
    UserControl.Parent.AutoRedraw = m_Redraw ' Ripristina
    '
End Sub
'------------------------------------------------------------------------
' AlphaBlend 2 nota bene Autoredraw deve essere disattivato
'------------------------------------------------------------------------
Private Sub GoAlpha2(Alpha As Byte)
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True

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
' Glass Tipo 0
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
    m_H = (m_Y2 / 100) * 5
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
        
            m_V = iRGB.r + Sta
            If m_V > 255 Then m_V = 255
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G + Sta
            If m_V > 255 Then m_V = 255
            iRGB.G = Int(m_V)
                        
            m_V = iRGB.B + Sta
            If m_V > 255 Then m_V = 255
            iRGB.B = m_V
        
'            SetPixel My_Obj.hdc, m_X, m_Y1 + m_Y, RGB(iRGB.r, iRGB.G, iRGB.B)
             SetPixelV My_Obj.hdc, m_X, m_Y1 + m_Y, RGB(iRGB.r, iRGB.G, iRGB.B)
        Next
        m_T = m_T - 1
    Next
    DoEvents
    My_Obj.Refresh
End Sub

