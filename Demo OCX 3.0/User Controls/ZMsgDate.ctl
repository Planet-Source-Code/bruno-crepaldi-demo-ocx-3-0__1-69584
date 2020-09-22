VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ZMsgDate 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Picture         =   "ZMsgDate.ctx":0000
   ScaleHeight     =   2190
   ScaleWidth      =   4575
   ToolboxBitmap   =   "ZMsgDate.ctx":89B2
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   45416449
      UpDown          =   -1  'True
      CurrentDate     =   38971
   End
   Begin VB.Timer ZTimer 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4080
      Top             =   3960
   End
   Begin VB.PictureBox PicButDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "ZMsgDate.ctx":8CC4
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "ZMsgDate.ctx":AC7C
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButEnabled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "ZMsgDate.ctx":CC34
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      Picture         =   "ZMsgDate.ctx":EBEC
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
      Begin VB.Image LblButton 
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         Picture         =   "ZMsgDate.ctx":10BA4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PictLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "ZMsgDate.ctx":121EE
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox PictBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      Picture         =   "ZMsgDate.ctx":12BF6
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   6
      Top             =   0
      Width           =   4575
      Begin VB.Label LblPrompt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LblTitle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   120
         Picture         =   "ZMsgDate.ctx":1B5A8
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1560
      End
   End
End
Attribute VB_Name = "ZMsgDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Date Message Box
' Nome del File ..: ZMsgDate
' Data............: 27/08/2007
' Versione........: 1.00
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
'   Ritorna la posizione Assoluta del mouse in PIXEL  X e Y
Private Declare Function M_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
'
Private Type POINTAPI
    X       As Long
    Y       As Long
End Type
Private pt   As POINTAPI
'
Private Type Myform
    Scalemode As Long
    width As Long
    Height As Long
    Top As Long
    Left As Long
End Type
'
Private MouseIsDown As Boolean
Private i As Long
Private Bt_Stat As Integer
Private m_Active As Boolean
Private m_Transparency As Byte
Private m_GlassEffect As Boolean
Private m_Redraw As Boolean
Private ZMsgBoxResult As Long

Private Myform As Form
'----------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'----------------------------------------------------------
Private Sub UserControl_InitProperties()
    If Not Ambient.UserMode Then
        Extender.Height = PictLogo.Height    ' Altezza
        Extender.width = PictLogo.width      ' Larghezza
    End If
        Extender.Visible = False
        '
End Sub
'----------------------------------------------------------
' UserControl Terminate
'----------------------------------------------------------
Private Sub UserControl_Terminate()
'
End Sub
'-----------------------------------------------------------
' Inizializa
'-----------------------------------------------------------
Private Sub UserControl_Initialize()
        m_Active = False
        ZMsgBoxResult = -1
End Sub
'----------------------------------------------------------
' Show
'----------------------------------------------------------
Private Sub UserControl_Show()
    DTPicker1.Value = Date
    Extender.ZOrder 0
    GoAlpha2 (m_Transparency)
End Sub
'----------------------------------------------------------
' Property Let / Get
'----------------------------------------------------------
Public Property Let Transparency(bTransparency As Byte)
    m_Transparency = bTransparency
    PropertyChanged "Transparency"
    GoAlpha2 (m_Transparency)
End Property
Public Property Get Transparency() As Byte
    Transparency = m_Transparency
End Property
'
Public Property Let GlassEffect(bGlassEffect As Boolean)
    m_GlassEffect = bGlassEffect
    PropertyChanged "GlassEffect"
    GoAlpha2 (m_Transparency)
End Property
Public Property Get GlassEffect() As Boolean
    GlassEffect = m_GlassEffect
End Property
'---------------------------------------------------------------------------------------
' PropertyBag Read / Write
'---------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Transparency", m_Transparency, 0
        .WriteProperty "GlassEffect", m_GlassEffect, True
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Transparency = .ReadProperty("Transparency", 0)
        m_GlassEffect = .ReadProperty("GlassEffect", True)
    End With
End Sub
'---------------------------------------------------------------------------------------
' Timer Gestione Mouse
'---------------------------------------------------------------------------------------
Private Sub ZTimer_Timer()
    On Error Resume Next
    If MouseIsDown = True Then Exit Sub
    
    Call M_GetCursorPos(pt)
    Select Case Myform.Scalemode
        Case 1
            pt.X = ScaleX(pt.X, vbPixels, vbTwips) - (Extender.Left + Myform.Left)  ' convert Pixels to Twips
            pt.Y = ScaleY(pt.Y, vbPixels, vbTwips) - (Extender.Top + Myform.Top)
        Case 3
            pt.X = ScaleX(pt.X - Extender.Left, vbPixels, vbTwips) - Myform.Left    ' convert Pixels to Twips
            pt.Y = ScaleY(pt.Y - Extender.Top, vbPixels, vbTwips) - Myform.Top
    End Select
    
        With PicButton
            If pt.X >= .Left And pt.X <= (.Left + .width) And pt.Y >= .Top And pt.Y <= (.Top + .Height) Then
                If Bt_Stat <> 1 Then
                    Bt_Stat = 1
                    UserControl.Parent.Enabled = True
                    .Picture = PicButHot.Picture 'Acceso
                End If
            Else
                If Bt_Stat <> 2 Then
                    Bt_Stat = 2
                    UserControl.Parent.Enabled = False
                    .Picture = PicButEnabled.Picture 'Spento
                End If
            End If
        End With
       
        With DTPicker1
            If pt.X >= .Left And pt.X <= (.Left + .width) And pt.Y >= .Top And pt.Y <= (.Top + .Height) Then
                UserControl.Parent.Enabled = True
            Else
                If Bt_Stat = 0 Then
                    UserControl.Parent.Enabled = False
                End If
            End If
        End With
    
End Sub
'---------------------------------------------------------------------------------------
' Mouse Down
'---------------------------------------------------------------------------------------
Private Sub PicButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = True
    Bt_Stat = 0
    PicButton.Picture = PicButDown.Picture
End Sub
'---------------------------------------------------------------------------------------
' Mouse Up
'---------------------------------------------------------------------------------------
Private Sub PicButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
    ZMsgBoxResult = 1
End Sub
'-----------------------------------------------------------
' Resize
'-----------------------------------------------------------
Private Sub Resize()
    On Error Resume Next
    Set Myform = UserControl.Parent         ' corrected
    '
    With Myform
        .width = Myform.width
        .Height = Myform.Height
        .Top = Myform.Top
        .Left = Myform.Left
    End With
    '
    With UserControl
        Select Case Myform.Scalemode
            Case 1
                .Extender.Left = (Myform.width - .width) / 2
                .Extender.Top = (Myform.Height - .Height) / 2
            Case 3
                .Extender.Left = ScaleX(Myform.width - .width, vbTwips, vbPixels) / 2
                .Extender.Top = ScaleY(Myform.Height - .Height, vbTwips, vbPixels) / 2
        End Select
    End With
End Sub
'---------------------------------------------------------------------------------------
' ZMsDate
'---------------------------------------------------------------------------------------
Public Function ZMsDate(Prompt As String) As Date
    
    If m_Active = True Then GoTo Go
        UserControl.Parent.Enabled = False
        m_Active = True
        ZMsgBoxResult = -1
        '
        Extender.Visible = True
        PictLogo.Visible = False
        '
        With UserControl
            .Height = 2175 '4575
            .width = 4575
        End With
        Resize
        LblPrompt.Caption = Prompt
        ZTimer = True
        '
Go:
    While ZMsgBoxResult = -1
        DoEvents
    Wend
    UserControl.Parent.Enabled = True
    Extender.Visible = False
    m_Active = False
    ZTimer = False
    ZMsDate = DTPicker1.Value

End Function
'------------------------------------------------------------------------
' AlphaBlend 2
'------------------------------------------------------------------------
Private Sub GoAlpha2(Alpha As Byte)
    m_Redraw = UserControl.Parent.AutoRedraw
    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True

    PictBack.Picture = UserControl.Picture
    tProperties.tBlendAmount = Alpha 'Set translucency level
    
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
    m_H = (m_Y2 / 100) * 7
    If m_H > 10 Then m_H = 10

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
