VERSION 5.00
Begin VB.UserControl GlassButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1485
   ScaleWidth      =   1935
   Begin VB.PictureBox PictBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.Image ImgIcon 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
End
Attribute VB_Name = "GlassButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Glass Button
' Nome del File ..: GlassButton
' Data............: 27/11/2004
' Versione........: 0.90 Beta
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
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

Private m_X As Long
Private m_Y As Long
Private m_V As Long
Private m_T As Long
Private m_H As Long
Private m_X1 As Long
Private m_Y1 As Long
Private m_X2 As Long
Private m_Y2 As Long
'
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private pt As POINTAPI
' DrawText
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private m_Rect As RECT

' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
'
Private Const WM_GETFONT As Long = &H31

Enum GlsBt_Align
    nLeft = DT_LEFT
    nCenter = DT_CENTER
    nRight = DT_RIGHT
End Enum
Private m_Align As GlsBt_Align
'
'
Enum GlsBt_Style
    GlsBtStyle_x1 = 0
    GlsBtStyle_x2 = 1
    GlsBtStyle_x3 = 2
    GlsBtStyle_x4 = 3
End Enum
Private m_Style As GlsBt_Style

Private i As Integer
Private m_Caption As String
Private m_ColorBack As OLE_COLOR
Private m_Transparency As Byte
Private m_GlassEffect As Long
Private m_Border As Boolean
'
Private m_IsHot As Boolean
Private m_IsHotC As Boolean
Private m_IsDown As Boolean
Private m_TextShadow As Boolean
Private m_ColorHotBorder As OLE_COLOR
Private m_Enabled As Boolean
Private m_Redraw As Boolean

' Dichiarazione Eventi
Public Event Click()
Public Event MouseDown()
'
Private WithEvents M_Frm As Form
Attribute M_Frm.VB_VarHelpID = -1
'----------------------------------------------------------
' m_Frm Events ( For Capture Parent Events  )
'----------------------------------------------------------
Private Sub M_Frm_Resize()
'
End Sub
Private Sub M_Frm_activate()
'
End Sub
'-----------------------------------------------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
     m_Style = GlsBtStyle_x1
     m_Align = nCenter
     m_ColorBack = &H404040
     m_Transparency = 160
     m_GlassEffect = True
     m_TextShadow = False
     m_Border = True
     m_ColorHotBorder = &H47AFF
     m_Enabled = True
     '
     Set Font = Ambient.Font
     '
     m_IsHot = False
     m_IsHotC = False
     '
     UserControl.Height = 495
     UserControl.width = 1935
End Sub
'-----------------------------------------------------------------------------------------------
' Resizing
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Resize()
    Dim Ht As Long
    PictBack.Move 0, 0, ScaleWidth, ScaleHeight
    Ht = PictBack.ScaleHeight - (PictBack.ScaleHeight / 100) * 30
    ImgIcon.Move ImgIcon.Left, (PictBack.ScaleHeight - Ht) / 2, Ht, Ht
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
    PictBack.BackColor = m_ColorBack
    GoAlpha2 m_Transparency
    If m_Enabled = False Then
        BtnGray PictBack
    End If
End Sub
'-----------------------------------------------------------------------------------------------
' Inizializa
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    m_IsHot = False
End Sub
'-----------------------------------------------------------------------------------------------
' Eventi
'-----------------------------------------------------------------------------------------------
Private Sub ClickEvent()
    RaiseEvent Click
End Sub
Private Sub MouseDownEvent()
    RaiseEvent MouseDown
End Sub
'-----------------------------------------------------------------------------------------------
' Property Let Set
'-----------------------------------------------------------------------------------------------
' Font
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(new_Font As StdFont)
    Set UserControl.Font = new_Font
    Set PictBack.Font = new_Font
    PropertyChanged "Font"
    GoAlpha2 m_Transparency
End Property
' Icon
Public Property Get Icon() As Picture
   Set Icon = ImgIcon.Picture
End Property
Public Property Set Icon(ByVal NewPic As Picture)
   Set ImgIcon.Picture = NewPic
   PropertyChanged "Icon"
   GoAlpha2 m_Transparency
End Property
' Caption
Public Property Get Caption() As String
   Caption = m_Caption
End Property
Public Property Let Caption(NewValue As String)
    m_Caption = NewValue
    PropertyChanged "Caption"
    GoAlpha2 m_Transparency
End Property
' Style
Public Property Get Style() As GlsBt_Style
   Style = m_Style
End Property
Public Property Let Style(ByVal NewValue As GlsBt_Style)
   m_Style = NewValue
   PropertyChanged "Style"
   GoAlpha2 m_Transparency
End Property
' Align
Public Property Get Align() As GlsBt_Align
   Align = m_Align
End Property
Public Property Let Align(ByVal NewValue As GlsBt_Align)
   m_Align = NewValue
   PropertyChanged "Align"
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
' ColorText
Public Property Get ColorText() As OLE_COLOR
   ColorText = UserControl.ForeColor
End Property
Public Property Let ColorText(ByVal NewValue As OLE_COLOR)
   PropertyChanged "ColorText"
   UserControl.ForeColor = NewValue
    GoAlpha2 m_Transparency
End Property
' ColorHotBorder
Public Property Get ColorHotBorder() As OLE_COLOR
   ColorHotBorder = m_ColorHotBorder
End Property
Public Property Let ColorHotBorder(ByVal NewValue As OLE_COLOR)
   PropertyChanged "ColorHotBorder"
    m_ColorHotBorder = NewValue
End Property
' Transparency
Public Property Let Transparency(bTransparency As Byte)
    m_Transparency = bTransparency
    PropertyChanged "Transparency"
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
' TextShadow
Public Property Get TextShadow() As Boolean
    TextShadow = m_TextShadow
End Property
Public Property Let TextShadow(bTextShadow As Boolean)
    m_TextShadow = bTextShadow
    PropertyChanged "TextShadow"
    GoAlpha2 m_Transparency
End Property
' Border
Public Property Get Border() As Boolean
    Border = m_Border
End Property
Public Property Let Border(bBorder As Boolean)
    m_Border = bBorder
    PropertyChanged "Border"
    PictBack.BorderStyle = Abs(m_Border)
    GoAlpha2 m_Transparency
End Property
' Enabled
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(bEnabled As Boolean)
    m_Enabled = bEnabled
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
    If m_Enabled = True Then
        GoAlpha2 m_Transparency
    Else
        BtnGray PictBack
    End If
End Property
'-----------------------------------------------------------------------------------------------
' Property Read Write
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "Transparency", m_Transparency, 160
        .WriteProperty "GlassEffect", m_GlassEffect, True
        .WriteProperty "TextShadow", m_TextShadow, False
        .WriteProperty "Border", m_Border, True
        .WriteProperty "Caption", m_Caption, Empty
        .WriteProperty "Style", m_Style, GlsBt_Style.GlsBtStyle_x1
        .WriteProperty "Align", m_Align, GlsBt_Align.nCenter
        .WriteProperty "ColorBack", m_ColorBack, &H404040
        .WriteProperty "ColorText", ColorText, &H0
        .WriteProperty "ColorHotBorder", m_ColorHotBorder, &H47AFF
        .WriteProperty "Icon", ImgIcon.Picture, Nothing
        .WriteProperty "Font", Font, Ambient.Font
    End With
End Sub
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Enabled = .ReadProperty("Enabled", True)
        m_Transparency = .ReadProperty("Transparency", 160)
        m_GlassEffect = .ReadProperty("GlassEffect", True)
        m_TextShadow = .ReadProperty("TextShadow", False)
        m_Border = .ReadProperty("Border", True)
        m_Caption = .ReadProperty("Caption", Empty)
        m_ColorBack = .ReadProperty("ColorBack", &H404040)
        m_ColorHotBorder = .ReadProperty("ColorHotBorder", &H47AFF)
        m_Style = .ReadProperty("Style", GlsBt_Style.GlsBtStyle_x1)
        m_Align = .ReadProperty("Align", GlsBt_Align.nCenter)
        UserControl.ForeColor = .ReadProperty("ColorText", &H0)
        Set ImgIcon.Picture = .ReadProperty("Icon", Nothing)
        Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    End With
    UserControl.Enabled = m_Enabled
    PictBack.BorderStyle = Abs(m_Border)
End Sub
'-----------------------------------------------------------------------------------------------
'
'
' Inizio Routine GlassButton
'
'
'-----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------
' UserControl Processing
'----------------------------------------------------------------------------------------
Private Sub Usercontrol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With UserControl
        If X < 0 Or Y < 0 Or X > .ScaleWidth Or Y > .ScaleHeight Then
Lp:         m_IsHot = False
            If Button <> 1 Then
                ReleaseCapture
            End If
        Else
            GetCursorPos pt
            If WindowFromPoint(pt.X, pt.Y) <> .hwnd Then
                GoTo Lp
            Else
                SetCapture hwnd
            End If
            m_IsHot = True
        End If
    End With

    If m_IsHot = m_IsHotC Then Exit Sub
    m_IsHotC = m_IsHot
    If m_IsHot = True Then
        HotBorder
    Else
        GoAlpha2 (m_Transparency)
    End If
End Sub
'
Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_IsDown = True
    BtnGray PictBack
    MouseDownEvent
End Sub
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_IsDown = False
    GoAlpha2 m_Transparency
    If m_IsHot = True Then HotBorder
    ClickEvent
End Sub
'-----------------------------------------------------------------------------------------------
' HotBorder
'-----------------------------------------------------------------------------------------------
Private Sub HotBorder()
    m_V = (m_ColorHotBorder And &HFF&) * 1.3
    If m_V > 255 Then m_V = 255
    iRGB.r = m_V
            
    m_V = ((m_ColorHotBorder And &HFF00&) / 2 ^ 8) * 1.3
    If m_V > 255 Then m_V = 255
    iRGB.G = m_V
            
    m_V = ((m_ColorHotBorder And &HFF0000) / 2 ^ 16) * 1.3
    If m_V > 255 Then m_V = 255
    iRGB.B = m_V
        
    With PictBack
        PictBack.Line (0, 0)-(.ScaleWidth - 1, 0), m_ColorHotBorder
        PictBack.Line (0, .ScaleHeight - 1)-(.ScaleWidth - 1, .ScaleHeight - 1), m_ColorHotBorder
        PictBack.Line (0, 0)-(0, .ScaleHeight - 1), m_ColorHotBorder
        PictBack.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), m_ColorHotBorder
            
        PictBack.Line (1, 1)-(.ScaleWidth - 2, 1), RGB(iRGB.r, iRGB.G, iRGB.B)
        PictBack.Line (1, .ScaleHeight - 2)-(.ScaleWidth - 2, .ScaleHeight - 2), RGB(iRGB.r, iRGB.G, iRGB.B)
        PictBack.Line (1, 1)-(1, .ScaleHeight - 2), RGB(iRGB.r, iRGB.G, iRGB.B)
        PictBack.Line (.ScaleWidth - 2, 1)-(.ScaleWidth - 2, .ScaleHeight - 2), RGB(iRGB.r, iRGB.G, iRGB.B)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' SizeCaption
'-----------------------------------------------------------------------------------------------
Private Sub SizeCaption(Colore As OLE_COLOR, Optional Shift_X As Long, Optional Shift_Y As Long)
Dim Res, m_Bt, m_W, m_H, m_Mrg As Long
Dim m_St As Long

    If ImgIcon.Picture <> Empty Then
        ImgIcon.Left = 0
    Else
        ImgIcon.Left = -ImgIcon.width
    End If
        
    PictBack.ForeColor = Colore


' Define the Rectangle for the Text - if Multiline must put the width do you want in m_Rect.Right
    m_Mrg = 5
    m_Rect.Top = 0
    m_Rect.Left = ImgIcon.Left + ImgIcon.width + m_Mrg
    m_Rect.Right = PictBack.ScaleWidth - m_Mrg * 2

    m_St = DT_CALCRECT Or m_Align Or DT_WORDBREAK
        
'    Res = DrawText(PictBack.hdc, m_Caption, -1, m_Rect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    Res = DrawText(PictBack.hdc, m_Caption, -1, m_Rect, m_St)

    m_Rect.Right = PictBack.ScaleWidth - m_Mrg * 2
    m_Bt = m_Rect.Bottom

    m_Rect.Left = m_Rect.Left + Shift_X
    m_Rect.Top = ((PictBack.ScaleHeight - m_Bt) / 2) + Shift_Y
    m_Rect.Bottom = m_Bt + m_Rect.Top
    
    m_St = m_Align Or DT_WORDBREAK
    
    DrawText PictBack.hdc, m_Caption, Len(m_Caption), m_Rect, m_St ' Print text
'    DrawText PictBack.hdc, m_Caption, Len(m_Caption), m_Rect, DT_CENTER Or DT_WORDBREAK  ' Print text

End Sub
'------------------------------------------------------------------------
' Refresh
'------------------------------------------------------------------------
Public Sub Refresh()
    If m_Enabled = True Then
        GoAlpha2 m_Transparency
    Else
        BtnGray PictBack
    End If
End Sub
'------------------------------------------------------------------------
' BtnGray
'------------------------------------------------------------------------
Private Sub BtnGray(My_Obj As Object)
Dim Rgb_R, Rgb_G, Rgb_B As Integer
    
    For m_Y = 0 To My_Obj.ScaleHeight
        For m_X = 0 To My_Obj.ScaleWidth
            m_V = GetPixel(My_Obj.hdc, m_X, m_Y)
            
            Rgb_R = m_V And &HFF&
            Rgb_G = (m_V And &HFF00&) \ &H100&
            Rgb_B = (m_V And &HFF0000) \ &H10000
            
            m_V = (Rgb_R + Rgb_G + Rgb_B) / 3
            SetPixelV My_Obj.hdc, m_X, m_Y1 + m_Y, RGB(m_V, m_V, m_V)
        Next
    Next
    
    SizeCaption &H404040
    My_Obj.Refresh
End Sub
'------------------------------------------------------------------------
' AlphaBlend 2 nota bene Autoredraw deve essere Attivato
'------------------------------------------------------------------------
Private Sub GoAlpha2(Alpha As Byte)
Dim SclMd As Long
    On Error Resume Next
        
    PictBack.Cls
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    
    If m_TextShadow = True Then SizeCaption &HFFFFFF, -3, 2
    
    tProperties.tBlendAmount = Alpha  'Set translucency level
    '
    SclMd = UserControl.Parent.Scalemode
    UserControl.Parent.Scalemode = 1
    m_Width = UserControl.width / Screen.TwipsPerPixelX
    m_Height = UserControl.Height / Screen.TwipsPerPixelY
    m_left = UserControl.Extender.Left / Screen.TwipsPerPixelX
    m_Top = UserControl.Extender.Top / Screen.TwipsPerPixelY
    UserControl.Parent.Scalemode = SclMd
    '
    hDCSrc = UserControl.Parent.hdc
    hDCDst = PictBack.hdc
    '
    CopyMemory lngBlend, tProperties, 4 'Blend colors
    AlphaBlend hDCDst, 0, 0, m_Width, m_Height, hDCSrc, m_left, m_Top, m_Width, m_Height, lngBlend 'Blend together
    '
    UserControl.Parent.AutoRedraw = m_Redraw ' Ripristina
    '
    SizeCaption ColorText

    If m_GlassEffect = True Then
        SetGlass PictBack ' Glass Effect
    Else
        PictBack.Refresh
    End If
End Sub
'------------------------------------------------------------------------
' Glass Tipo 1
'------------------------------------------------------------------------
Private Sub SetGlass(My_Obj As Object)
Dim Sta As Long 'Integer
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





