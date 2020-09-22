VERSION 5.00
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Main Form"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":324A
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DemoOCX30.ZMsgBox ZMsgBox1 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Transparency    =   160
   End
   Begin DemoOCX30.ZMsgDate ZMsgDate1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Transparency    =   160
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   0
      Left            =   6120
      TabIndex        =   34
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      TextShadow      =   -1  'True
      Caption         =   "Alpha Text"
      ColorText       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   4800
      TabIndex        =   30
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1085
      Forecolor       =   12583104
   End
   Begin DemoOCX30.TextScroll TextScroll1 
      Height          =   6855
      Left            =   9000
      TabIndex        =   29
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   12091
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PictDateBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00CEB7AF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2265
      ScaleWidth      =   2505
      TabIndex        =   20
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CheckBox CheckGlass1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   160
         TabIndex        =   21
         Top             =   1110
         Value           =   1  'Checked
         Width           =   200
      End
      Begin DemoOCX30.HSlider HSlider2 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Value           =   160
         MaxValue        =   255
         Picture         =   "FrmMain.frx":1E0A8
         PicCursor_Selected=   "FrmMain.frx":1F614
         PictureCursor   =   "FrmMain.frx":1F9EC
      End
      Begin DemoOCX30.GlassButton GlassButton1 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date Insert Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transparency"
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
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   460
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GlassEffect"
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   2175
      End
   End
   Begin DemoOCX30.TitleBar TitleBar1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   529
   End
   Begin DemoOCX30.BottomBar BottomBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   15
      Top             =   8700
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   529
      Resize          =   0   'False
   End
   Begin VB.PictureBox PictResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4080
      ScaleHeight     =   855
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Result"
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
         Height          =   855
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox PictMsgBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00CEB7AF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   840
      ScaleHeight     =   2865
      ScaleWidth      =   2505
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CheckBox CheckGlass 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   160
         TabIndex        =   19
         Top             =   1760
         Value           =   1  'Checked
         Width           =   200
      End
      Begin VB.ComboBox CmbMsg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "FrmMain.frx":1FDC4
         Left            =   120
         List            =   "FrmMain.frx":1FDC6
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox CmbMsg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "FrmMain.frx":1FDC8
         Left            =   120
         List            =   "FrmMain.frx":1FDD5
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin DemoOCX30.HSlider HSlider1 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Value           =   160
         MaxValue        =   255
         Picture         =   "FrmMain.frx":1FE03
         PicCursor_Selected=   "FrmMain.frx":2136F
         PictureCursor   =   "FrmMain.frx":21747
      End
      Begin DemoOCX30.GlassButton GlassButton1 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Show"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GlassEffect"
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
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transparency"
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1180
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Message Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2535
      End
   End
   Begin DemoOCX30.HSlider HSlider3 
      Height          =   255
      Left            =   4800
      TabIndex        =   31
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      Value           =   160
      MaxValue        =   255
      Picture         =   "FrmMain.frx":21B1F
      PicCursor_Selected=   "FrmMain.frx":2308B
      PictureCursor   =   "FrmMain.frx":23463
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   1
      Left            =   6120
      TabIndex        =   35
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      TextShadow      =   -1  'True
      Caption         =   "Nothing"
      ColorText       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.ExaButton ExaButton1 
      Height          =   1455
      Index           =   5
      Left            =   2640
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "DateInsertBox"
      Picture         =   "FrmMain.frx":2383B
   End
   Begin DemoOCX30.ExaButton ExaButton1 
      Height          =   1455
      Index           =   4
      Left            =   2640
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "MessageBox"
      Picture         =   "FrmMain.frx":26A95
   End
   Begin DemoOCX30.ExaButton ExaButton1 
      Height          =   1455
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "GlassButton"
      Picture         =   "FrmMain.frx":29CEF
   End
   Begin DemoOCX30.ExaButton ExaButton1 
      Height          =   1455
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "AlphaBlend demo"
      Picture         =   "FrmMain.frx":2CF49
   End
   Begin DemoOCX30.ExaButton ExaButton1 
      Height          =   1455
      Index           =   3
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "DisplayLed"
      Picture         =   "FrmMain.frx":2D85C
   End
   Begin DemoOCX30.ExaButton ExaButton1 
      Height          =   1455
      Index           =   0
      Left            =   1440
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Title/Bottom Bar"
      Picture         =   "FrmMain.frx":2EEB6
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EXA Buttons"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Progressbar /  Hslider"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   32
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================
' Demo Oxc 3.0
' Thanks to Roger Gilchrist for his Help"
'=====================================================


Option Explicit

Private i As Long
Private LeMsg As Long
Private TextMsg As String
Private CntDs1 As Integer
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'----------------------------------------------------------------------
' Form Initialize
'----------------------------------------------------------------------
Private Sub Form_Initialize()
'
End Sub
'----------------------------------------------------------------------
' Form Activate
'----------------------------------------------------------------------
Private Sub Form_Activate()
    TextScroll1.ColorBack = &H408B9B
    TextScroll1.Status = 1
    LoadTxt
End Sub
'---------------------------------------------------------------------------------------
' Form Load
'---------------------------------------------------------------------------------------
Private Sub Form_Load()

    With Me
        .Height = Screen.Height
        .width = Screen.width
    End With
    StretchPicture Me
    '
    GlassButton2(0).Refresh
    GlassButton2(1).Refresh
    '
    CmbMsg(0).Clear
    CmbMsg(1).Clear
    
    CmbMsg(0).AddItem "vbOKOnly": CmbMsg(0).ItemData(0) = vbOKOnly
    CmbMsg(0).AddItem "vbYesNo": CmbMsg(0).ItemData(1) = vbYesNo
    CmbMsg(0).AddItem "vbOKCancel": CmbMsg(0).ItemData(2) = vbOKCancel
    CmbMsg(0).AddItem "vbYesNoCancel": CmbMsg(0).ItemData(3) = vbYesNoCancel
    CmbMsg(0).Text = CmbMsg(0).List(0)

    CmbMsg(1).AddItem "vbExclamation": CmbMsg(1).ItemData(0) = vbExclamation
    CmbMsg(1).AddItem "vbInformation": CmbMsg(1).ItemData(1) = vbInformation
    CmbMsg(1).AddItem "vbCritical": CmbMsg(1).ItemData(2) = vbCritical
    CmbMsg(1).Text = CmbMsg(1).List(0)
    ' Form Clock
    With FrmClock
 '       .Left = Me.width - FrmClock.width - 120
 '       .Top = 500 'Me.Height - .Height - 500
        .Show
    End With

    CheckGlass.Value = Abs(ZMsgBox1.GlassEffect)
    HSlider1.Value = ZMsgBox1.Transparency
    
    With ProgressBar1
        .MinValue = HSlider3.MinValue
        .MaxValue = HSlider3.MaxValue
        .Value = HSlider3.Value
    End With
End Sub
'--------------------------------------------------------
' Form_Resize
'--------------------------------------------------------
Private Sub Form_Resize()
 Call FormOnTop(FrmClock.hwnd, True)
    PictDateBox.Top = PictMsgBox.Top
    PictDateBox.Left = PictMsgBox.Left
    
    With TextScroll1
        .Left = Me.width - .width - 120
        .Top = (Me.Height - .Height) / 2
    End With
    
End Sub
'--------------------------------------------------------
' Form_Unload
'--------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
   Dim frm As Form
   For Each frm In Forms
     Unload frm
   Next
   End
End Sub
'---------------------------------------------------------------------------------------
' LoadTxt
'---------------------------------------------------------------------------------------
Private Sub LoadTxt()
Dim L_Txt As String
    L_Txt = "[ DEMO OCX 3.0 ]" & vbCrLf
    L_Txt = L_Txt & "By Bruno Crepaldi 2007" & vbCrLf & vbCrLf
    L_Txt = L_Txt & "[ Elenco OCX ]" & vbCrLf
    L_Txt = L_Txt & "Title Bar" & vbCrLf
    L_Txt = L_Txt & "Bottom Bar" & vbCrLf
    L_Txt = L_Txt & "Display Led Translucent" & vbCrLf
    L_Txt = L_Txt & "Message Box Translucent" & vbCrLf
    L_Txt = L_Txt & "Date Insert Box Translucent" & vbCrLf
    L_Txt = L_Txt & "TextScroll Translucent" & vbCrLf
    L_Txt = L_Txt & "button Translucent" & vbCrLf
    L_Txt = L_Txt & "Progress Bar Translucent" & vbCrLf
    L_Txt = L_Txt & "Exagonal button" & vbCrLf
    L_Txt = L_Txt & "Horizontal Slider" & vbCrLf
    L_Txt = L_Txt & "Horizontal Updown" & vbCrLf
    L_Txt = L_Txt & "Alpha Text" & vbCrLf
    L_Txt = L_Txt & vbCrLf & "[ Commenti ]" & vbCrLf
    L_Txt = L_Txt & "Just First Beta Version" & vbCrLf
    L_Txt = L_Txt & "Many Bugs To be corrected" & vbCrLf
    L_Txt = L_Txt & "Compile me before use !!" & vbCrLf
    L_Txt = L_Txt & vbCrLf & "[ Thanks ]" & vbCrLf
    L_Txt = L_Txt & "Thanks to Roger Gilchrist" & vbCrLf
    L_Txt = L_Txt & "for his Help" & vbCrLf
        
    
    TextScroll1.Caption = L_Txt
End Sub
'--------------------------------------------------------
'  Command Menu
'--------------------------------------------------------
Private Sub ExaButton1_Click(Index As Integer)
    PictMsgBox.Visible = False
    PictResult.Visible = False
    PictDateBox.Visible = False
    Select Case Index
        Case 0
            FrmEx01.Show
        Case 1
            SetParent FrmAlpha.hwnd, Me.hwnd
            FrmAlpha.Show
        Case 2
'            SetParent FrmEx03.hwnd, Me.hwnd
            FrmEx03.Show
        Case 3
            FrmEx02.Show
        Case 4
            PictMsgBox.Visible = True
            PictResult.Visible = True
        Case 5
            PictDateBox.Visible = True
        Case 6
'
    End Select
End Sub
'
Private Sub GlassButton1_Click(Index As Integer)
    Select Case Index
        Case 0
            messagebox
        Case 1
            messageDate
    End Select
End Sub
'
Private Sub GlassButton2_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmEx04.Show
        Case 1
    
    End Select
End Sub
'---------------------------------------------------------------------------------------
' Messagebox
'---------------------------------------------------------------------------------------
Private Sub messagebox()
Dim Rtn As Long
Dim P As Long
    ' vbExclamation vbInformation vbCritical
    ' vbOKOnly  vbYesNo vbOKCancel vbYesNoCancel
    P = CmbMsg(0).ItemData(CmbMsg(0).ListIndex) + CmbMsg(1).ItemData(CmbMsg(1).ListIndex)
    Rtn = ZMsgBox1.ZMsbox("Area del Messaggio", P, "Titolo")
    Text2 = "Key" + Str(Rtn) + " Pressed"
End Sub
'---------------------------------------------------------------------------------------
' HSlider1
'---------------------------------------------------------------------------------------
Private Sub HSlider1_Change(Value As Long)
    ZMsgBox1.Transparency = Value
End Sub
'---------------------------------------------------------------------------------------
' CheckGlass
'---------------------------------------------------------------------------------------
Private Sub CheckGlass_Click()
    ZMsgBox1.GlassEffect = CheckGlass.Value
End Sub
'---------------------------------------------------------------------------------------
' DateInsertbox
'---------------------------------------------------------------------------------------
Private Sub messageDate()
    PictResult.Visible = True
    Text2 = ZMsgDate1.ZMsDate("Inserire Data")
End Sub
'---------------------------------------------------------------------------------------
' HSlider2
'---------------------------------------------------------------------------------------
Private Sub HSlider2_Change(Value As Long)
    ZMsgDate1.Transparency = Value
End Sub
'---------------------------------------------------------------------------------------
' CheckGlass
'---------------------------------------------------------------------------------------
Private Sub CheckGlass1_Click()
    ZMsgDate1.GlassEffect = CheckGlass1.Value
End Sub
'---------------------------------------------------------------------------------------
' HSlider3
'---------------------------------------------------------------------------------------
Private Sub HSlider3_scroll(Value As Long)
    ProgressBar1.Value = Value
End Sub
'--------------------------------------------------------------------
' TextScroll1
'--------------------------------------------------------------------
Private Sub TextScroll1_Click()
    With TextScroll1
        Select Case .Status
            Case GlsTxtStop
                .Status = GlsTxtStart
            Case GlsTxtPause
                .Status = GlsTxtStart
            Case GlsTxtStart
                .Status = GlsTxtPause
    End Select
    End With
End Sub
