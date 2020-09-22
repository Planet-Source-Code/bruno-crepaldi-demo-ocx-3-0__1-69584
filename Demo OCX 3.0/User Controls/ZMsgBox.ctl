VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ZMsgBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   Picture         =   "ZMsgBox.ctx":0000
   ScaleHeight     =   2715
   ScaleWidth      =   4845
   ToolboxBitmap   =   "ZMsgBox.ctx":2ABB4
   Begin VB.PictureBox PictLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "ZMsgBox.ctx":2AEC6
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   0
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ZMsgBox.ctx":2B8CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ZMsgBox.ctx":2CF28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ZMsgBox.ctx":2E582
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ZMsgBox.ctx":2FBDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ZMsgBox.ctx":31236
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicButDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3360
      Picture         =   "ZMsgBox.ctx":32890
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1800
      Picture         =   "ZMsgBox.ctx":34848
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "ZMsgBox.ctx":36800
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButDisabled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      Picture         =   "ZMsgBox.ctx":387B8
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer ZTimer 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   4680
   End
   Begin VB.PictureBox PicButHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3360
      Picture         =   "ZMsgBox.ctx":3A770
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1800
      Picture         =   "ZMsgBox.ctx":3C728
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "ZMsgBox.ctx":3E6E0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButEnabled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3360
      Picture         =   "ZMsgBox.ctx":40698
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButEnabled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1800
      Picture         =   "ZMsgBox.ctx":42650
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButEnabled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "ZMsgBox.ctx":44608
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3360
      Picture         =   "ZMsgBox.ctx":465C0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
      Begin VB.Image ImgButton 
         Enabled         =   0   'False
         Height          =   465
         Index           =   2
         Left            =   360
         Picture         =   "ZMsgBox.ctx":48578
         Stretch         =   -1  'True
         Top             =   30
         Width           =   465
      End
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1800
      Picture         =   "ZMsgBox.ctx":49BC2
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
      Begin VB.Image ImgButton 
         Enabled         =   0   'False
         Height          =   465
         Index           =   1
         Left            =   360
         Picture         =   "ZMsgBox.ctx":4BB7A
         Stretch         =   -1  'True
         Top             =   30
         Width           =   465
      End
   End
   Begin VB.PictureBox PicButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   240
      Picture         =   "ZMsgBox.ctx":4D1C4
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
      Begin VB.Image ImgButton 
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   360
         Picture         =   "ZMsgBox.ctx":4F17C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox PictBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   0
      Picture         =   "ZMsgBox.ctx":507C6
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   323
      TabIndex        =   17
      Top             =   0
      Width           =   4845
      Begin VB.Label LblTitle 
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
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   4575
      End
      Begin VB.Image ImgIcon 
         Height          =   960
         Left            =   3720
         Picture         =   "ZMsgBox.ctx":7B37A
         Top             =   720
         Width           =   960
      End
      Begin VB.Label LblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prompt"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3495
      End
   End
   Begin VB.Label LblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   16
      Top             =   1720
      Width           =   1215
   End
   Begin VB.Label LblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   1720
      Width           =   1215
   End
   Begin VB.Label LblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1720
      Width           =   1215
   End
End
Attribute VB_Name = "ZMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Message Box
' Nome del File ..: ZMsgBox
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
'   API Constant Declarations
Private Const BM_SETSTATE = &HF3
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const LWA_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
' Mouse
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
Private Bt_Enabled(2) As Boolean
Private Bt_Stat(2) As Integer
Private m_Active As Boolean
Private m_Transparency As Byte
Private m_GlassEffect As Boolean
Private m_Redraw As Boolean
Private Myform As Form

Private ZMsgBoxResult As Long
'----------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'----------------------------------------------------------
Private Sub UserControl_InitProperties()
    If Not Ambient.UserMode Then
    End If
    '
    Extender.Height = PictLogo.Height    ' Altezza
    Extender.width = PictLogo.width      ' Larghezza
    Extender.Visible = False
    '
End Sub
'-----------------------------------------------------------
' Inizializa
'-----------------------------------------------------------
Private Sub UserControl_Initialize()
    Bt_Enabled(0) = True
    Bt_Enabled(1) = True
    Bt_Enabled(2) = True
    m_Active = False
    m_Transparency = 0 ' Opaque
    m_GlassEffect = False
    ZMsgBoxResult = -1
End Sub
'----------------------------------------------------------
' Resize
'----------------------------------------------------------
Private Sub UserControl_Resize()
  With PictBack
    .Left = 0
    .Top = 0
    .width = UserControl.width
    .Height = UserControl.Height
  End With
End Sub
'----------------------------------------------------------
' Show
'----------------------------------------------------------
Private Sub UserControl_Show()
   Extender.ZOrder 0
   GoAlpha2 (m_Transparency)
End Sub
'----------------------------------------------------------
' Terminate
'----------------------------------------------------------
Private Sub UserControl_Terminate()
'
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
            pt.X = ScaleX(pt.X, vbPixels, vbTwips) - (Extender.Left + Myform.Left)    ' convert Pixels to Twips
            pt.Y = ScaleY(pt.Y, vbPixels, vbTwips) - (Extender.Top + Myform.Top)
        Case 3
            pt.X = ScaleX(pt.X - Extender.Left, vbPixels, vbTwips) - Myform.Left   ' convert Pixels to Twips
            pt.Y = ScaleY(pt.Y - Extender.Top, vbPixels, vbTwips) - Myform.Top
    End Select
    
    For i = 0 To 2
    If Bt_Enabled(i) = True Then
        With PicButton(i)
            pt.Y = pt.Y
            If pt.X >= .Left And pt.X <= (.Left + .width) And pt.Y >= .Top And pt.Y <= (.Top + .Height) Then
                If Bt_Stat(i) <> 1 Then
                    Bt_Stat(i) = 1
                    UserControl.Parent.Enabled = True       ' Sblocca il form di chiamata
                    .Picture = PicButHot(i).Picture 'Acceso
                End If
            Else
                If Bt_Stat(i) <> 2 Then
                    Bt_Stat(i) = 2
                    UserControl.Parent.Enabled = False      ' blocca il form di chiamata
                    .Picture = PicButEnabled(i).Picture 'Spento
                End If
            End If
        End With
    End If
    Next i
End Sub
'---------------------------------------------------------------------------------------
' Mouse Down
'---------------------------------------------------------------------------------------
Private Sub PicButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = True
    Bt_Stat(Index) = 0
    PicButton(Index).Picture = PicButDown(Index).Picture
End Sub
'---------------------------------------------------------------------------------------
' Mouse Up
'---------------------------------------------------------------------------------------
Private Sub PicButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
    ZMsgBoxResult = Index
  '  ZMsbox (Empty)
End Sub
'-----------------------------------------------------------
' Resize
'-----------------------------------------------------------
Private Sub Resize()
    On Error Resume Next
    Set Myform = UserControl.Parent
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
' ZMsbox
'---------------------------------------------------------------------------------------
Public Function ZMsbox(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String) As Long
    
    If m_Active = True Then GoTo Go
        UserControl.Parent.Enabled = False      ' blocca il form di chiamata
        m_Active = True
        ZMsgBoxResult = -1

        Extender.Visible = True
        PictLogo.Visible = False
        
        With UserControl
            .Height = 2715 ' 2460
            .width = 4845
        End With
        Resize
                
        LblPrompt.Caption = Prompt
        LblPrompt.BorderStyle = 0
        LblTitle.Caption = Title
        
        Set_Control (Buttons)
    
        ZTimer.Enabled = True
    
Go:
    While ZMsgBoxResult = -1
        DoEvents
    Wend
    UserControl.Parent.Enabled = True      ' blocca il form di chiamata
    Extender.Visible = False
    m_Active = False
    ZTimer.Enabled = False
    ZMsbox = ZMsgBoxResult

End Function
'---------------------------------------------------------------------------------------
' Set_Control
'---------------------------------------------------------------------------------------
Private Sub Set_Control(Buttons As VbMsgBoxStyle)
        
        If (Buttons And vbOKOnly) = vbOKOnly Then
            Bt_Enabled(0) = False
            PicButton(0).Visible = False
            LblButton(0) = Empty
            Bt_Enabled(1) = True
            PicButton(1).Visible = True
            LblButton(1) = "OK"
            Bt_Enabled(2) = False
            PicButton(2).Visible = False
            LblButton(2) = Empty
            Set_Icon (Buttons)
        End If
        
        If (Buttons And vbYesNo) = vbYesNo Then
            Bt_Enabled(0) = True
            PicButton(0).Visible = True
            LblButton(0) = "SI"
            Bt_Enabled(1) = True
            PicButton(1).Visible = True
            LblButton(1) = "NO"
            Bt_Enabled(2) = False
            PicButton(2).Visible = False
            LblButton(2) = Empty
            Set_Icon (Buttons)
        End If
        
        If (Buttons And vbOKCancel) = vbOKCancel Then
            Bt_Enabled(0) = True
            PicButton(0).Visible = True
            LblButton(0) = "SI"
            Bt_Enabled(1) = False
            PicButton(1).Visible = False
            LblButton(1) = Empty
            Bt_Enabled(2) = True
            PicButton(2).Visible = True
            LblButton(2) = "Cancel"
            Set_Icon (Buttons)
        End If

        If (Buttons And vbYesNoCancel) = vbYesNoCancel Then
            Bt_Enabled(0) = True
            PicButton(0).Visible = True
            LblButton(0) = "SI"
            Bt_Enabled(1) = True
            PicButton(1).Visible = True
            LblButton(1) = "NO"
            Bt_Enabled(2) = True
            PicButton(2).Visible = True
            LblButton(2) = "Cancel"
            Set_Icon (Buttons)
        End If

'
'            Bt_Enabled(0) = False
'            PicButton(0).Visible = False
'            LblButton(0) = Empty
'            Bt_Enabled(1) = False
'            PicButton(1).Visible = False
'            LblButton(1) = Empty
'            Bt_Enabled(2) = False
'            PicButton(2).Visible = False
'            LblButton(2) = Empty
'            Set_Icon (Buttons)
'
End Sub
'---------------------------------------------------------------------------------------
' Set_Icon
'---------------------------------------------------------------------------------------
Public Sub Set_Icon(Buttons As VbMsgBoxStyle)
   ' vbCritical vbExclamation vbInformation
    If (Buttons And vbExclamation) = vbExclamation Then
        ImgIcon.Picture = ImageList1.ListImages.Item(1).Picture
        Exit Sub
    End If
    '
    If (Buttons And vbInformation) = vbInformation Then
        ImgIcon.Picture = ImageList1.ListImages.Item(2).Picture
        Exit Sub
    End If
    '
    If (Buttons And vbCritical) = vbCritical Then
        ImgIcon.Picture = ImageList1.ListImages.Item(3).Picture
        Exit Sub
    End If
End Sub
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
