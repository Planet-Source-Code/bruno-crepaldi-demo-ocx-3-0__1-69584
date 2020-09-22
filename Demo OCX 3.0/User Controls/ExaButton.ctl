VERSION 5.00
Begin VB.UserControl ExaButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "ExaButton.ctx":0000
   Picture         =   "ExaButton.ctx":77FE
   ScaleHeight     =   1650
   ScaleWidth      =   11325
   ToolboxBitmap   =   "ExaButton.ctx":EFFC
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7080
      Top             =   0
   End
   Begin VB.Image ImageDown 
      Height          =   1455
      Index           =   1
      Left            =   5400
      Picture         =   "ExaButton.ctx":F30E
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image ImageFocused 
      Height          =   1455
      Index           =   1
      Left            =   9600
      Picture         =   "ExaButton.ctx":1203E
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image ImageDown 
      Height          =   1455
      Index           =   0
      Left            =   3600
      Picture         =   "ExaButton.ctx":1983E
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image ImgIcon 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   900
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   900
   End
   Begin VB.Image ImageUp 
      Height          =   1455
      Left            =   1680
      Picture         =   "ExaButton.ctx":1C56E
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image ImageFocused 
      Height          =   1455
      Index           =   0
      Left            =   7920
      Picture         =   "ExaButton.ctx":23D6C
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "ExaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Exa Button
' Nome del File ..: ExaButton
' Data............: 27/08/2007
' Versione........: 0.1 Beta
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
'   Ritorna la posizione Assoluta del mouse in PIXEL  X e Y
Private Declare Function M_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
'
Private Type POINTAPI
    X       As Long
    Y       As Long
End Type
Private pt   As POINTAPI
'
Private m_Caption As String
Private m_Hit As Integer
Private MouseIsDown As Boolean
Enum E_Style
    Distort
    Pushed
End Enum
Private m_Style As E_Style
' Dichiarazione Eventi
Public Event Click()
'----------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'----------------------------------------------------------
Private Sub UserControl_InitProperties()
    If Not Ambient.UserMode Then
        UserControl.HitBehavior = 1
        UserControl.ClipBehavior = 1
    End If
    With UserControl
        .width = 1575
        .Height = 1455
    End With
End Sub
'-----------------------------------------------------------
' Inizializa
'-----------------------------------------------------------
Private Sub UserControl_Initialize()
    MouseIsDown = False
    
End Sub
'-----------------------------------------------------------
' Resize
'-----------------------------------------------------------
Private Sub UserControl_Resize()
    With UserControl
        .width = 1575
        .Height = 1455
    End With
    SizeIcon
End Sub
'---------------------------------------------------------------------------------------
' Property Let / Get
'---------------------------------------------------------------------------------------
Public Property Get Picture() As Picture
   Set Picture = ImgIcon.Picture
End Property

Public Property Set Picture(ByVal NewPic As Picture)
   Set ImgIcon.Picture = NewPic
   PropertyChanged "Picture"
End Property
'
Public Property Let Caption(bCaption As String)
    m_Caption = bCaption
    PropertyChanged "Caption"
    LblCaption.Caption = m_Caption
  '  UserControl_Resize
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property
'
Public Property Let Style(xVal As E_Style)
    If xVal <> m_Style Then
        m_Style = xVal
        PropertyChanged "Style"
    End If
End Property
Public Property Get Style() As E_Style
    Style = m_Style
End Property
'-----------------------------------------------------------------
'                 Read/Write Properties
'-----------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Caption = .ReadProperty("Caption", Empty)
        Set ImgIcon.Picture = .ReadProperty("Picture", Nothing)
        m_Style = .ReadProperty("Style", 0)
    End With
    LblCaption.Caption = m_Caption
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", m_Caption, Empty
        .WriteProperty "Picture", ImgIcon.Picture, Nothing
        .WriteProperty "Style", m_Style, 0
    End With
End Sub
'-----------------------------------------------------------------
' Eventi
'-----------------------------------------------------------------
Private Sub ClickEvent()
    RaiseEvent Click
End Sub
'-----------------------------------------------------------------
' SizeIcon
'-----------------------------------------------------------------
Private Sub SizeIcon()
    With ImgIcon
        If MouseIsDown = True Then
            .width = 900
            .Height = 900
            .Top = 190
            .Left = ((UserControl.ScaleWidth - .width) / 2)
        Else
            .width = 495
            .Height = 495
            .Top = 120 '(UserControl.Scaleheight - .height) / 2
            .Left = ((UserControl.ScaleWidth - .width) / 2)
        End If
    End With
End Sub
'-----------------------------------------------------------
' HitTest
'-----------------------------------------------------------
Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
If UserControl.Ambient.UserMode Then  '<<<< IDE GUARD
    m_Hit = HitResult
    Select Case HitResult
        Case 0
            UserControl.Picture = ImageUp.Picture
            Timer1.Enabled = False
        Case 3
            Extender.ZOrder 0
            If MouseIsDown = True Then Exit Sub
            UserControl.Picture = ImageFocused(m_Style).Picture
            Timer1.Enabled = True
    End Select
    End If
End Sub
'-----------------------------------------------------------
' Timer
'-----------------------------------------------------------
Private Sub Timer1_Timer()
    If MouseIsDown = True Then Exit Sub
    
    Call M_GetCursorPos(pt)
    With UserControl
        pt.X = ScaleX(pt.X, vbPixels, vbTwips) - .Parent.Left    ' convert Pixels to Twips
        pt.Y = ScaleY(pt.Y, vbPixels, vbTwips) - .Parent.Top
        
        If Not (pt.X >= .Extender.Left And pt.X <= (.Extender.Left + .width) And pt.Y >= .Extender.Top And pt.Y <= (.Extender.Top + .Height)) Then
            .Picture = ImageUp.Picture
            Timer1.Enabled = False
        End If
    End With
End Sub
'-----------------------------------------------------------
' Mouse
'-----------------------------------------------------------
Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDown
End Sub
Private Sub LblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDown
End Sub
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MUp
End Sub
Private Sub LblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MUp
End Sub
'
Private Sub MDown()
    MouseIsDown = True
    UserControl.Picture = ImageDown(m_Style).Picture
    SizeIcon
End Sub
Private Sub MUp()
    MouseIsDown = False
    SizeIcon
    Call ClickEvent
End Sub

