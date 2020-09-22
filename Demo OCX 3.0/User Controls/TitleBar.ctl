VERSION 5.00
Begin VB.UserControl TitleBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   2550
   ScaleWidth      =   8595
   ToolboxBitmap   =   "TitleBar.ctx":0000
   Begin VB.PictureBox TitleLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      Picture         =   "TitleBar.ctx":0312
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox PictIconBase 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1200
      Picture         =   "TitleBar.ctx":1166
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   600
      Begin VB.Image ImgIconBase 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   90
         Picture         =   "TitleBar.ctx":246A
         Stretch         =   -1  'True
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   960
      Picture         =   "TitleBar.ctx":2B70
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   6
      Left            =   1320
      Picture         =   "TitleBar.ctx":2C04
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   5
      Left            =   1200
      Picture         =   "TitleBar.ctx":2C98
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   4
      Left            =   1080
      Picture         =   "TitleBar.ctx":2D2C
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   3
      Left            =   960
      Picture         =   "TitleBar.ctx":2DC0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   840
      Picture         =   "TitleBar.ctx":2E54
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   720
      Picture         =   "TitleBar.ctx":2EE8
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PicMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   960
      Picture         =   "TitleBar.ctx":2F7C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   960
      Picture         =   "TitleBar.ctx":3470
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   960
      Picture         =   "TitleBar.ctx":3964
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   9
      ToolTipText     =   "Maximize"
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   480
      Picture         =   "TitleBar.ctx":3E56
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   8
      ToolTipText     =   "Minimize"
      Top             =   1560
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   480
      Picture         =   "TitleBar.ctx":434A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      ToolTipText     =   "Minimize"
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   120
      Top             =   2040
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   0
      Picture         =   "TitleBar.ctx":483E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   1560
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   0
      Picture         =   "TitleBar.ctx":4D32
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox titleButton2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1440
      Picture         =   "TitleBar.ctx":5226
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox titleRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1680
      Picture         =   "TitleBar.ctx":5718
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox titleButton 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1440
      Picture         =   "TitleBar.ctx":584A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   480
      Picture         =   "TitleBar.ctx":5D3C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      ToolTipText     =   "Minimize"
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   0
      Picture         =   "TitleBar.ctx":622E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   840
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1920
      Picture         =   "TitleBar.ctx":6722
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   45
   End
End
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Form Title Bar
' Nome del File ..: TitleBar
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
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Private Declare Function SetBkColor& Lib "gdi32" (ByVal hdc&, ByVal crColor&)
Private Declare Function GetPixel& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&)
Private Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hdc&, ByVal nWidth&, ByVal nHeight&)
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hdc&)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc&, ByVal hObject&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Private Declare Function CreateBitmap& Lib "gdi32" (ByVal nWidth&, ByVal nHeight&, ByVal nPlanes&, ByVal nBitCount&, ByVal lpBits As Any)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hdc&)

Private Const SRCCOPY = &HCC0020
Private Const SRCINVERT = &H660046
Private Const SRCAND = &H8800C6

'   API Constant Declarations
Private Const BM_SETSTATE = &HF3
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const LWA_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
'
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' Form Trasparente
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'
Private Declare Function ReleaseCapture Lib "user32" () As Long
' Serve per la routine di stampa diretta sul controllo utente
Const LF_FACESIZE = 32
Const LF_FULLFACESIZE = 64
Private Declare Function TextColor Lib "gdi32" Alias "SetTextColor" (ByVal hdc As Long, ByVal Colore As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

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
'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'   Ritorna la posizione Assoluta del mouse in PIXEL  X e Y
Private Declare Function M_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
'
Private Type POINTAPI
    X       As Long
    Y       As Long
End Type
Private pt   As POINTAPI
'
Private mX, mY As Long
'
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private MouseIsDown As Boolean
Private i As Long
'
Enum TitleBarStyle
    blue
    Gray
    red
    green
    Violet
    Yellow
End Enum
Private m_Style As TitleBarStyle
'
Private m_bEnabled As Boolean
Private m_kClose As Boolean
Private m_kMaximize As Boolean
Private m_kMinimize As Boolean
Private m_IconEnable As Boolean
Private m_Flocked As Boolean
Private m_Transparency As Integer
'
Private P_Hwnd As Long ' Obbligatorio Passare usercontrol.hwnd per Trasparenza ( se no nn Funziona !)
'
Private myform As Form
Attribute myform.VB_VarHelpID = -1
Private WithEvents PicFormRight As PictureBox
Attribute PicFormRight.VB_VarHelpID = -1
Private WithEvents PicFormLeft As PictureBox
Attribute PicFormLeft.VB_VarHelpID = -1

Private WithEvents PictIcon As PictureBox
Attribute PictIcon.VB_VarHelpID = -1
Private WithEvents ImgIcon As Image
Attribute ImgIcon.VB_VarHelpID = -1

Private Msg As Long
'----------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'----------------------------------------------------------
Private Sub UserControl_InitProperties()
    If Not Ambient.UserMode Then
    End If
        UserControl.Extender.Align = 1
        m_bEnabled = True
        m_kClose = True
        m_kMaximize = True
        m_kMinimize = True
        m_IconEnable = True
        m_Flocked = False
        m_Transparency = 255
End Sub
'-----------------------------------------------------------
' Inizializa
'-----------------------------------------------------------
Private Sub UserControl_Initialize()
    '
End Sub
'----------------------------------------------------------
' Show
'----------------------------------------------------------
Private Sub UserControl_Show()
'Dim P_Hwnd As Long
        On Error Resume Next
        
     '   P_Hwnd = UserControl.ContainerHwnd
        P_Hwnd = UserControl.Extender.Container.hwnd
        
        Msg = GetWindowLong(P_Hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong P_Hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes P_Hwnd, 0, m_Transparency, LWA_ALPHA
        '
        UserControl.width = 1000
        Init_Style
        Timer1_Timer                                     '<<<<<< Calling it this way is dodgy but allows the control to draw properly
        Timer1.Enabled = UserControl.Ambient.UserMode    '<<<<<< this stops the timer running in the IDE
End Sub
'----------------------------------------------------------
' Terminate
'----------------------------------------------------------
Private Sub UserControl_Terminate()
  '
End Sub
'-----------------------------------------------------------
' Resize
'-----------------------------------------------------------
Private Sub UserControl_Resize()
Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, L1 As Long, H1 As Long, L2 As Long, H2 As Long ' Corrected
  '  On Error Resume Next
    UserControl.Height = 300
    
    Set myform = UserControl.Parent
    CreaFormIcon
    '
    TitleLeft.Left = 0
    TitleLeft.Top = 0
    '
    X1 = 0 'TitleLeft.width
    Y1 = 0 '300 ' AAAAA
    L1 = PictBack(0).width * (UserControl.ScaleWidth)   'Risultato Destinazione
    H1 = PictBack(0).Height                             'Risultato Destinazione
    X2 = 0
    Y2 = 0
    L2 = PictBack(0).width
    H2 = PictBack(0).Height
    '
    UserControl.PaintPicture PictBack(0).Image, X1, Y1, L1, H1, X2, Y2, L2, H2, vbSrcCopy
    '
    CopiaTrsp TitleLeft
    '
    With picClose(0)
        .Top = 0
        .Left = UserControl.width - 30 - picClose(0).width
    End With
    
    With PicMax(0)
        .Top = 0
        .Left = UserControl.width - 60 - (picClose(0).width * 2)
    End With
    
    With picMin(0)
        .Top = 0
        .Left = UserControl.width - 90 - (picClose(0).width * 3)
    End With
    '
    Call Scrive_H(UserControl.hdc, 2, 71, 14, vbWhite, myform.Caption)
    Call Scrive_H(UserControl.hdc, 1, 70, 14, vbBlack, myform.Caption)
End Sub
'---------------------------------------------------------------------------------------
' Property Let / Get
'---------------------------------------------------------------------------------------
Public Property Let Style(xVal As TitleBarStyle)
    If xVal <> m_Style Then
        m_Style = xVal
        PropertyChanged "Style"
        Init_Style
    End If
End Property
Public Property Get Style() As TitleBarStyle
    Style = m_Style
End Property
'
Public Property Let Enabled(bEnabled As Boolean)
On Error GoTo Handler
    m_bEnabled = bEnabled
    PropertyChanged "Enabled"
Handler:
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property
'
Public Property Let KeyClose(bClose As Boolean)
    m_kClose = bClose
    PropertyChanged "KeyClose"
    If m_kClose = True Then
            picClose(0) = picClose(1)
        Else
            picClose(0) = titleButton
    End If
    picClose(0).Enabled = m_kClose
End Property
Public Property Get KeyClose() As Boolean
    KeyClose = m_kClose
End Property
'
Public Property Let KeyMaximize(bMaximize As Boolean)
    m_kMaximize = bMaximize
    PropertyChanged "KeyMaximize"
    
        If m_kMaximize = True Then
            PicMax(0) = PicMax(1)
        Else
            PicMax(0) = titleButton
    End If
    PicMax(0).Enabled = m_kMaximize
End Property
Public Property Get KeyMaximize() As Boolean
    KeyMaximize = m_kMaximize
End Property
'
Public Property Let KeyMinimize(bMinimize As Boolean)
    m_kMinimize = bMinimize
    PropertyChanged "KeyMinimize"
    
    If m_kMinimize = True Then
            picMin(0) = picMin(1)
        Else
            picMin(0) = titleButton
    End If
    picMin(0).Enabled = m_kMinimize
End Property
Public Property Get KeyMinimize() As Boolean
    KeyMinimize = m_kMinimize
End Property
'
Public Property Let IconEnable(bIconEnable As Boolean)
    m_IconEnable = bIconEnable
    PropertyChanged "IconEnable"
    CreaFormIcon
    UserControl.Refresh
End Property

Public Property Get IconEnable() As Boolean
    IconEnable = m_IconEnable
End Property
'
Public Property Let Transparency(bTransparency As Integer)
    m_Transparency = bTransparency
    PropertyChanged "Transparency"
'    MakeTransparent myform.hwnd, 0, m_Transparency, LWA_ALPHA
    MakeTransparent UserControl.Extender.Container.hwnd, 0, m_Transparency, LWA_ALPHA
End Property

Public Property Get Transparency() As Integer
    Transparency = m_Transparency
End Property
'---------------------------------------------------------------------------------------
' PropertyBag Read / Write
'---------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "KeyClose", m_kClose, True
        .WriteProperty "KeyMaximize", m_kMaximize, True
        .WriteProperty "KeyMinimize", m_kMinimize, True
        .WriteProperty "IconEnable", m_IconEnable, True
        .WriteProperty "Transparency", m_Transparency, 255
        .WriteProperty "Style", m_Style, 0
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_bEnabled = .ReadProperty("Enabled", True)
        m_kClose = .ReadProperty("KeyClose", True)
        m_kMaximize = .ReadProperty("KeyMaximize", True)
        m_kMinimize = .ReadProperty("KeyMinimize", True)
        m_IconEnable = .ReadProperty("IconEnable", True)
        m_Transparency = .ReadProperty("Transparency", 255)
        m_Style = .ReadProperty("Style", 0)
    End With
    
    picClose(0).Enabled = m_kClose
    picMin(0).Enabled = m_kMinimize
    PicMax(0).Enabled = m_kMaximize
End Sub
'--------------------------------------------------------
'  CopiaTrsp
'--------------------------------------------------------
Private Sub CopiaTrsp(m_Pict As PictureBox)
Dim m_Col As Long
    With m_Pict
        .Visible = True
        .Left = 0
        .Top = 0
        m_Col = GetPixel(.hdc, .ScaleWidth - 1, .ScaleHeight - 1) '&HFF00FF  '
        TransBltNow UserControl.hdc, .Left, .Top, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, m_Col
        .Visible = False
    End With
End Sub
Public Sub TransBltNow(hDestDC As Long, lDestX As Long, lDestY As Long, lWidth As Long, lHeight As Long, hSourceDC As Long, lSourceX As Long, lSourceY As Long, lTransColor As Long)
Dim lOldColor As Long
Dim hMaskDC As Long
Dim hMaskBmp As Long
Dim hOldMaskBmp As Long
Dim hTempBmp As Long
Dim hTempDC As Long
Dim hOldTempBmp As Long
Dim hDummy As Long
Dim lRet As Long

    lOldColor = SetBkColor&(hSourceDC, lTransColor)
    lOldColor = SetBkColor&(hDestDC, lTransColor)
    hMaskDC = CreateCompatibleDC(hDestDC)
    hMaskBmp = CreateCompatibleBitmap(hDestDC, lWidth, lHeight)
    hOldMaskBmp = SelectObject(hMaskDC, hMaskBmp)
    hTempBmp = CreateBitmap(lWidth, lHeight, 1, 1, 0&)
    hTempDC = CreateCompatibleDC(hDestDC)
    hOldTempBmp = SelectObject(hTempDC, hTempBmp)
    If BitBlt(hTempDC, 0, 0, lWidth, lHeight, hSourceDC, lSourceX, lSourceY, SRCCOPY) Then
        hDummy = BitBlt(hMaskDC, 0, 0, lWidth, lHeight, hTempDC, 0, 0, SRCCOPY)
    End If
    hTempBmp = SelectObject(hTempDC, hOldTempBmp)
    hDummy = DeleteObject(hTempBmp)
    hDummy = DeleteDC(hTempDC)
    lRet = BitBlt(hDestDC, lDestX, lDestY, lWidth, lHeight, hSourceDC, lSourceX, lSourceY, SRCINVERT)
    lRet = BitBlt(hDestDC, lDestX, lDestY, lWidth, lHeight, hMaskDC, 0, 0, SRCAND)
    lRet = BitBlt(hDestDC, lDestX, lDestY, lWidth, lHeight, hSourceDC, lSourceX, lSourceY, SRCINVERT)
    hMaskBmp = SelectObject(hMaskDC, hOldMaskBmp)
    hDummy = DeleteObject(hMaskBmp)
    hDummy = DeleteDC(hMaskDC)
End Sub
'-----------------------------------------------------------------------
' Form Trasparente
' Es: 1) MakeTransparent(Me.hWnd, &H00FF00FF&,0, LWA_COLORKEY) As Long
' Es: 2) MakeTransparent(Me.hWnd, 0, 0-255 , LWA_ALPHA) As Long
'-----------------------------------------------------------------------
Private Function MakeTransparent(ByVal hwnd As Long, clr As Long, iValue As Integer, dFlags As Long) As Long
 '   Dim Msg As Long
    If iValue < 0 Or iValue > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, clr, iValue, dFlags
        MakeTransparent = 0
    End If
    If Err Then
        MakeTransparent = 2
    End If
End Function
'---------------------------------------------------------------------------------------
' Inizializa Style
'---------------------------------------------------------------------------------------
Private Sub Init_Style()
    PictBack(0) = PictBack(m_Style + 1)
    UserControl_Resize
End Sub
'---------------------------------------------------------------------------------------
' Timer Gestione Mouse
'---------------------------------------------------------------------------------------
Private Sub Timer1_Timer()
    
    If MouseIsDown = True Then Exit Sub
    
    Call M_GetCursorPos(pt)
    pt.X = ScaleX(pt.X, vbPixels, vbTwips) - myform.Left    ' convert Pixels to Twips
    pt.Y = ScaleY(pt.Y, vbPixels, vbTwips) - myform.Top
    
    If m_kClose = True Then
        If pt.X >= picClose(0).Left And pt.X <= (picClose(0).Left + picClose(0).width) And pt.Y >= picClose(0).Top And pt.Y <= (picClose(0).Top + picClose(0).Height) Then
            picClose(0) = picClose(2) 'Acceso
        Else
            picClose(0) = picClose(1) 'Spento
        End If
    End If
'
    If m_kMinimize = True Then
        If pt.X >= picMin(0).Left And pt.X <= (picMin(0).Left + picMin(0).width) And pt.Y >= picMin(0).Top And pt.Y <= (picMin(0).Top + picMin(0).Height) Then
            picMin(0) = picMin(2) 'Acceso
        Else
            picMin(0) = picMin(1) 'Spento
        End If
    End If
'
    If m_kMaximize = True Then
        If pt.X >= PicMax(0).Left And pt.X <= (PicMax(0).Left + PicMax(0).width) And pt.Y >= PicMax(0).Top And pt.Y <= (PicMax(0).Top + PicMax(0).Height) Then
            PicMax(0) = PicMax(2) 'Acceso
        Else
            PicMax(0) = PicMax(1) 'Spento
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' MoveForm
'---------------------------------------------------------------------------------------
Private Function MoveForm(TheForm As Form)
    Dim ret
    If m_Flocked = True Then Exit Function
    ReleaseCapture
    SendMessage TheForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function
'---------------------------------------------------------------------------------------
' usercontrol_MouseDown
'---------------------------------------------------------------------------------------
Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm myform
End Sub
'---------------------------------------------------------------------------------------
' picClose
'---------------------------------------------------------------------------------------
Private Sub picClose_Click(Index As Integer)
    Unload myform
End Sub

Private Sub picClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = True
    picClose(0) = titleButton2
End Sub

Private Sub picClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
    picClose(0) = picClose(2)
End Sub
'---------------------------------------------------------------------------------------
' PicMax
'---------------------------------------------------------------------------------------
Private Sub PicMax_Click(Index As Integer)
    If Parent.WindowState = 0 Then
        myform.WindowState = 2              ' Full Screen
        m_Flocked = True
    Else
        m_Flocked = False
        myform.WindowState = 0             ' Normal
    End If
End Sub

Private Sub PicMax_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = True
    PicMax(0) = titleButton2
End Sub

Private Sub PicMax_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
    PicMax(0) = PicMax(2)
End Sub
'---------------------------------------------------------------------------------------
' picMin
'---------------------------------------------------------------------------------------
Private Sub picMin_Click(Index As Integer)
    myform.WindowState = 1             ' Minimize
End Sub

Private Sub picMin_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = True
    picMin(0) = titleButton2
End Sub

Private Sub picMin_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
    picMin(0) = picMin(2)
End Sub
'----------------------------------------------------------------------
' Scrive_H      ( Scrive un Testo in Horizontale in Un Form
' Call Scrive_H(V_Riga, 280, 16, Riga)
'----------------------------------------------------------------------
Private Sub Scrive_H(ByVal hdc As Long, Py As Long, Px As Long, FontSize As Long, TColor As Long, Xlabel As String)
Dim HnFont As Long, HoFont As Long
Dim lf As LOGFONT
Dim r As Long

    'FontSize = 12
    lf.lfHeight = FontSize
    lf.lfItalic = False
    lf.lfUnderline = False
    r = TextColor(hdc, TColor)
    
    HnFont = CreateFontIndirect(lf)
    HoFont = SelectObject(hdc, HnFont)
    r = TextOut(hdc, Px, Py, Xlabel, Len(Xlabel))
End Sub
'---------------------------------------------------------------------------------------
' CreaFormIcon
'---------------------------------------------------------------------------------------
Private Sub CreaFormIcon()
Dim ScMod As Long
    ScMod = myform.Scalemode
    On Error Resume Next
    Set PictIcon = myform.Controls.Add("VB.PictureBox", "PictIcon")
    PictIcon.Picture = PictIconBase.Picture
    myform.Scalemode = vbPixels
    With PictIcon
        .Top = 0
        .Left = 6
        .width = 40
        .Height = 40
        .Align = 0 ' Right
        .Appearance = 0
        .BorderStyle = 0
        .AutoRedraw = True
        .Visible = m_IconEnable
        .BackColor = vbBlack
        .ZOrder 0
    End With
    
    Set ImgIcon = myform.Controls.Add("VB.Image", "ImgIcon", PictIcon)
    ImgIcon.Picture = myform.Icon ' ImgIconBase.Picture
    With ImgIcon
        .Top = 90
        .Left = 90
        .width = 420
        .Height = 420
        .Stretch = True
        .Appearance = 0
        .BorderStyle = 0
        .Visible = True
        .ZOrder 0
    End With
    myform.Scalemode = ScMod
End Sub
