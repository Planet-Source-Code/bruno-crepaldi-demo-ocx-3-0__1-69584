VERSION 5.00
Begin VB.Form FrmEx02 
   Appearance      =   0  'Flat
   BackColor       =   &H00CEB7AF&
   BorderStyle     =   0  'None
   Caption         =   "Form Esempio DisplayLed"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   Icon            =   "FrmEx02.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmEx02.frx":324A
   ScaleHeight     =   7005
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2970
      Left            =   120
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   17
      Top             =   600
      Width           =   2775
      Begin VB.TextBox TxtDspTest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin DemoOCX30.H_UpDown H_UpDown 
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ForeColor       =   -2147483640
         Value           =   3
         MaxValue        =   9
      End
      Begin DemoOCX30.H_UpDown H_UpDown 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   20
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ForeColor       =   -2147483640
         Value           =   1
         MinValue        =   1
      End
      Begin DemoOCX30.H_UpDown H_UpDown 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ForeColor       =   -2147483640
         Value           =   9
         MinValue        =   1
         MaxValue        =   9
      End
      Begin DemoOCX30.H_UpDown H_UpDown 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   22
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ForeColor       =   -2147483640
         MaxValue        =   6
      End
      Begin DemoOCX30.H_UpDown H_UpDown 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   23
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         ForeColor       =   -2147483640
         MaxValue        =   9
         LoopValue       =   -1  'True
      End
      Begin DemoOCX30.HSlider HSlider3 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Value           =   160
         MaxValue        =   255
         Picture         =   "FrmEx02.frx":44BA6
         PicCursor_Selected=   "FrmEx02.frx":46112
         PictureCursor   =   "FrmEx02.frx":464EA
      End
   End
   Begin DemoOCX30.BottomBar BottomBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6705
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   529
      Resize          =   0   'False
   End
   Begin DemoOCX30.TitleBar TitleBar1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   529
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8760
      Top             =   6000
   End
   Begin DemoOCX30.DisplayLed DspTest 
      Height          =   2970
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   5239
      Transparency    =   160
      Zoom            =   9
      LedColor        =   65280
      BackColor       =   49152
      Style           =   3
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   495
      Index           =   0
      Left            =   5640
      TabIndex        =   4
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Transparency    =   220
      Caption         =   "Tasto Num 1"
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Transparency    =   220
      Caption         =   "Tasto Num 2"
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Transparency    =   220
      Caption         =   "Tasto Num 3"
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   495
      Index           =   3
      Left            =   5640
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Transparency    =   220
      Caption         =   "Tasto Num 4"
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   495
      Index           =   4
      Left            =   5640
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Transparency    =   220
      Caption         =   "Tasto Num 5"
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
   Begin DemoOCX30.TextScroll TextScroll1 
      Height          =   6255
      Left            =   8280
      TabIndex        =   9
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   11033
      ColorText       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1085
      Transparency    =   180
      BackColor       =   16777152
      Forecolor       =   32768
      Value           =   60
      MaxValue        =   255
   End
   Begin VB.Label Label8 
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
      Left            =   240
      TabIndex        =   16
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Style"
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
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ScrollStyle"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Zoom"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LedColor"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Value"
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
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BackColor"
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
      Height          =   270
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DisplayLed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   4935
   End
End
Attribute VB_Name = "FrmEx02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ii As Byte
Dim Style(1) As Long
Dim Trsp As Integer
'----------------------------------------------------------------------
' Form Initialize
'----------------------------------------------------------------------
Private Sub Form_Initialize()
'
End Sub
'---------------------------------------------------------------------------------------
' Form Load
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
    ii = 40
    Call H_UpDown_Change(0, 0)
    LoadTxt
End Sub
'----------------------------------------------------------------------
' Form Activate
'----------------------------------------------------------------------
Private Sub Form_Activate()
    GoAlpha2 Picture2, 160 ' 0=Opaque
    SetGlass Picture2
    TextScroll1.Status = 1
    ProgressBar1.Value = HSlider3.Value
End Sub
'----------------------------------------------------------------------
' Form Resize
'----------------------------------------------------------------------
Private Sub Form_Resize()
'
End Sub
'--------------------------------------------------------------------
' H_UpDown_Change
'--------------------------------------------------------------------
'--------------------------------------------------------------------
' H_UpDown_Change
'--------------------------------------------------------------------
Private Sub H_UpDown_Change(Index As Integer, Value As Long)
    Select Case Index
        Case 0
            DspTest.Value = Asc(Trim(Str(Value))) '    Valore Ascii
        Case 1
            Select Case Value
                Case 0
                    DspTest.LedColor = RGB(168, 255, 0)
                Case 1
                    DspTest.LedColor = &HFFFFFF
                Case 2
                    DspTest.LedColor = RGB(255, 90, 0)
                Case 3
                    DspTest.LedColor = RGB(252, 255, 0)
                Case 4
                    DspTest.LedColor = RGB(168, 250, 255)
                Case 5
                    DspTest.LedColor = RGB(255, 150, 200)
                Case 6
                    DspTest.LedColor = RGB(124, 142, 252)
            End Select
        Case 2
            DspTest.Zoom = Value
        Case 3
            DspTest.ScrollStyle = Value
        Case 4
            DspTest.Style = Value
        Case 5
            DspTest.BackColor = Value
    End Select
End Sub
'---------------------------------------------------------------------------------------
' HSlider3
'---------------------------------------------------------------------------------------
Private Sub HSlider3_change(Value As Long)
    DspTest.Transparency = Value
End Sub
Private Sub HSlider3_scroll(Value As Long)
    ProgressBar1.Value = Value
End Sub
'---------------------------------------------------------------------------------------
' Timer1
'---------------------------------------------------------------------------------------
Private Sub Timer1_Timer()
' If ii < 255 Then ii = ii + 5: TitleBar1.Transparency = ii
 If ii < 230 Then
    ii = ii + 5: TitleBar1.Transparency = ii
 Else
    Timer1.Enabled = False
 End If
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
Private Sub LoadTxt()
Dim L_Txt As String
    L_Txt = "[ TextScroll ]" & vbCrLf
    L_Txt = L_Txt & "By Bruno Crepaldi 2007" & vbCrLf & vbCrLf
    L_Txt = L_Txt & vbCrLf & "[ Commenti ]" & vbCrLf
    L_Txt = L_Txt & "Click To Stop Scroll" & vbCrLf
    L_Txt = L_Txt & "Click To Start Scroll" & vbCrLf
    
    TextScroll1.Caption = L_Txt
End Sub
