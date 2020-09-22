VERSION 5.00
Begin VB.UserControl BottomBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   ScaleHeight     =   975
   ScaleWidth      =   7020
   ToolboxBitmap   =   "BottomBar.ctx":0000
   Begin VB.PictureBox TitleRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5640
      Picture         =   "BottomBar.ctx":0312
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox TitleLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      Picture         =   "BottomBar.ctx":1166
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox PicResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6600
      MousePointer    =   8  'Size NW SE
      Picture         =   "BottomBar.ctx":1FBA
      ScaleHeight     =   300
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   0
      Width           =   270
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   6
      Left            =   720
      Picture         =   "BottomBar.ctx":245E
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   5
      Left            =   600
      Picture         =   "BottomBar.ctx":24F2
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   4
      Left            =   480
      Picture         =   "BottomBar.ctx":2586
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   3
      Left            =   360
      Picture         =   "BottomBar.ctx":261A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   2
      Left            =   240
      Picture         =   "BottomBar.ctx":26AE
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   120
      Picture         =   "BottomBar.ctx":2742
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox PictBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   5520
      Picture         =   "BottomBar.ctx":27D6
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "BottomBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Form Bottom Bar
' Nome del File ..: BottomBar
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
Private M_Stato As Long
Private i As Long
'
Enum BottomBarStyle
    blue
    Gray
    red
    green
    Violet
    Yellow
End Enum
Private m_Style As BottomBarStyle

Private m_Enabled As Boolean
Private m_Resize As Boolean
Private m_Mdown As Boolean
Private m_Caption As String
Private m_MinSize As Long
'Dichiarazione Eventi
Public Event KeyPressed(Value As Long)
'----------------------------------------------------------
'      Inizializza le Variabili ( Solo Progetazione )
'----------------------------------------------------------
Private Sub UserControl_InitProperties()
    If Not Ambient.UserMode Then
        UserControl.Extender.Align = 2
        m_Enabled = True
        m_Resize = True
        m_Caption = Empty
        m_MinSize = 2000
    End If
End Sub
'-----------------------------------------------------------
'                       Inizializa
'-----------------------------------------------------------
Private Sub UserControl_Initialize()
    m_Style = Style
End Sub
'----------------------------------------------------------
' Show
'----------------------------------------------------------
Private Sub UserControl_Show()
    UserControl.Height = 300
    UserControl.width = 1000
    m_MinSize = 2000
    Init_Style
End Sub
'----------------------------------------------------------
' Terminate
'----------------------------------------------------------
Private Sub UserControl_Terminate()
  '
End Sub
'-----------------------------------------------------------
' Resizing
'-----------------------------------------------------------
Private Sub UserControl_Resize()
Dim X1, Y1, X2, Y2, L1, H1, L2, H2 As Long
    
 '   On Error Resume Next
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
    titleRight.Left = 0
    titleRight.Top = 0
    CopiaTrsp TitleLeft
    titleRight.Top = 0
    titleRight.Left = UserControl.width - titleRight.width
    CopiaTrsp titleRight
    '
    With PicResize
        .Top = 0
        .Left = UserControl.width - .width
    End With
    '
    WriteCaption
End Sub
'-----------------------------------------------------------------
'                        Eventi
'-----------------------------------------------------------------
Private Sub KeyPressedEvent(Valore As Long)
    RaiseEvent KeyPressed(Valore)
End Sub
'---------------------------------------------------------------------------------------
' Property Let / Get
'---------------------------------------------------------------------------------------
Public Property Let Style(xVal As BottomBarStyle)
    If xVal <> m_Style Then
        m_Style = xVal
        PropertyChanged "Style"
        Init_Style
    End If
End Property

Public Property Get Style() As BottomBarStyle
    Style = m_Style
End Property
'
Public Property Let Caption(bCaption As String)
    m_Caption = bCaption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property
'
Public Property Let Enabled(bEnabled As Boolean)
On Error GoTo Handler
    m_Enabled = bEnabled
    PropertyChanged "Enabled"
Handler:
End Property

Public Property Get Enabled() As Boolean
On Error GoTo Handler
    Enabled = m_Enabled
    Refresh
    Exit Property
Handler:
End Property
'
Public Property Let Resize(bResize As Boolean)
On Error GoTo Handler
    m_Resize = bResize
    PropertyChanged "Resize"
    If m_Resize = True Then
        PicResize.MousePointer = 8
    Else
        PicResize.MousePointer = 0
    End If
Handler:
End Property

Public Property Get Resize() As Boolean
    Resize = m_Resize
End Property
'---------------------------------------------------------------------------------------
' PropertyBag Read / Write
'---------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "Resize", m_Resize, True
        .WriteProperty "Style", m_Style, 0
        .WriteProperty "Caption", m_Caption, Empty
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Enabled = .ReadProperty("Enabled", True)
        m_Resize = .ReadProperty("Resize", True)
        m_Style = .ReadProperty("Style", 0)
        m_Caption = .ReadProperty("Caption", Empty)
    End With
  
    If m_Resize = True Then
        PicResize.MousePointer = 8
    Else
        PicResize.MousePointer = 0
    End If
    UserControl_Resize
End Sub
'--------------------------------------------------------
'  WriteCaption
'--------------------------------------------------------
Private Sub WriteCaption()
    Call Scrive_H(4, 71, 14, vbWhite, m_Caption)
    Call Scrive_H(3, 70, 14, vbBlack, m_Caption)
End Sub
'--------------------------------------------------------
'  CopiaTrsp
'--------------------------------------------------------
Private Sub CopiaTrsp(m_Pict As PictureBox)
Dim m_Col As Long
    With m_Pict
        .Visible = True
        m_Col = &HFF00FF ' GetPixel(.hdc, .ScaleWidth - 1, .ScaleHeight - 1)
        TransBltNow UserControl.hdc, ScaleX(.Left, vbTwips, vbPixels), ScaleX(.Top, vbTwips, vbPixels), .ScaleWidth, .ScaleHeight, .hdc, 0, 0, m_Col
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
'---------------------------------------------------------------------------------------
' Inizializa Style
'---------------------------------------------------------------------------------------
Private Sub Init_Style()
    PictBack(0) = PictBack(m_Style + 1)
    UserControl_Resize
End Sub
'----------------------------------------------------------------------
' Scrive_H      ( Scrive un Testo in Horizontale in Un Form
' Call Scrive_H(V_Riga, 280, 16, Riga)
'----------------------------------------------------------------------
Private Sub Scrive_H(Py As Long, Px As Long, FontSize As Long, TColor As Long, Xlabel As String)
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
' picResize
'---------------------------------------------------------------------------------------
Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Mdown = True
    UserControl.Parent.BottomBar1.ZOrder 0
End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Resize = False Then Exit Sub
    On Error Resume Next
    If m_Mdown = True Then
        Call M_GetCursorPos(pt)
        pt.X = ScaleX(pt.X, vbPixels, vbTwips) ' - UserControl.Parent.Left  ' convert Pixels to Twips
        pt.Y = ScaleY(pt.Y, vbPixels, vbTwips) ' - UserControl.Parent.Top
        If pt.X - UserControl.Parent.Left > m_MinSize Then UserControl.Parent.width = pt.X - UserControl.Parent.Left
        If pt.Y - UserControl.Parent.Top > m_MinSize Then UserControl.Parent.Height = pt.Y - UserControl.Parent.Top
    End If
End Sub

Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Mdown = False
    UserControl.Parent.Refresh
End Sub

