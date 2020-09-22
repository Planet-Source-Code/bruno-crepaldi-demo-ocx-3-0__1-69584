VERSION 5.00
Begin VB.Form FrmEx01 
   Appearance      =   0  'Flat
   BackColor       =   &H00CEB7AF&
   BorderStyle     =   0  'None
   Caption         =   "Form Esempio Title/Bottom Bar"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   Icon            =   "FrmEx01.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtCaption 
      Height          =   285
      Left            =   6720
      TabIndex        =   14
      Text            =   "Test Caption"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   5160
   End
   Begin DemoOCX30.HSlider HSlider1 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   5400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      Value           =   40
      MinValue        =   40
      MaxValue        =   255
      Picture         =   "FrmEx01.frx":324A
      PicCursor_Selected=   "FrmEx01.frx":47B6
      PictureCursor   =   "FrmEx01.frx":4B8E
   End
   Begin DemoOCX30.BottomBar BottomBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   529
   End
   Begin DemoOCX30.TitleBar TitleBar1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   529
      Transparency    =   40
   End
   Begin DemoOCX30.ExaButton CmdButton01 
      Height          =   1455
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Enable/disable Icon"
   End
   Begin DemoOCX30.ExaButton CmdButton01 
      Height          =   1455
      Index           =   4
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Change Style"
   End
   Begin DemoOCX30.ExaButton CmdButton02 
      Height          =   1455
      Index           =   0
      Left            =   7920
      TabIndex        =   11
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Caption"
   End
   Begin DemoOCX30.ExaButton CmdButton02 
      Height          =   1455
      Index           =   1
      Left            =   9120
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Resize On/Off"
   End
   Begin DemoOCX30.ExaButton CmdButton02 
      Height          =   1455
      Index           =   2
      Left            =   9120
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Change Style"
   End
   Begin DemoOCX30.ExaButton CmdButton01 
      Height          =   1455
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "KeyMinimize on/off"
   End
   Begin DemoOCX30.ExaButton CmdButton01 
      Height          =   1455
      Index           =   5
      Left            =   2760
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "Form Transparency"
   End
   Begin DemoOCX30.ExaButton CmdButton01 
      Height          =   1455
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "KeyClose on/off"
   End
   Begin DemoOCX30.ExaButton CmdButton01 
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      Caption         =   "KeyMaximize on/off"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Form Transparency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   720
      TabIndex        =   15
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottom Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Title Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "FrmEx01"
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
  '  Forms.Add Me
End Sub
'---------------------------------------------------------------------------------------
' Form Load
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
    ii = 40
    HSlider1.Value = 235
End Sub
'---------------------------------------------------------------------------------------
' Commandi
'---------------------------------------------------------------------------------------

Private Sub CmdButton01_Click(Index As Integer) '------------------- TitleBar

    Select Case Index
        Case 0
            TitleBar1.KeyClose = Not TitleBar1.KeyClose
        Case 1
            TitleBar1.KeyMaximize = Not TitleBar1.KeyMaximize
        Case 2
            TitleBar1.KeyMinimize = Not TitleBar1.KeyMinimize
        Case 3
            TitleBar1.IconEnable = Not TitleBar1.IconEnable
        Case 4
            If TitleBar1.Style < 5 Then
                TitleBar1.Style = TitleBar1.Style + 1
            Else
                TitleBar1.Style = 0
            End If
        Case 5
            TitleBar1.Transparency = Trsp
    End Select
End Sub
Private Sub CmdButton02_Click(Index As Integer)
    Select Case Index
        Case 0
            BottomBar1.Caption = TxtCaption.Text
        Case 1
            BottomBar1.Resize = Not BottomBar1.Resize
        Case 2
            If BottomBar1.Style < 5 Then
                BottomBar1.Style = BottomBar1.Style + 1
            Else
                BottomBar1.Style = 0
            End If
    End Select
End Sub
'---------------------------------------------------------------------------------------
' HSlider
'---------------------------------------------------------------------------------------
Private Sub HSlider1_scroll(Value As Long)
    Trsp = Value
    CmdButton01(5).Caption = "Form Transparency =" + Str(Trsp)
    TitleBar1.Transparency = Trsp
End Sub
'---------------------------------------------------------------------------------------
' Timer1
'---------------------------------------------------------------------------------------
Private Sub Timer1_Timer()
' If ii < 255 Then ii = ii + 5: TitleBar1.Transparency = ii
 If ii < 230 Then ii = ii + 5: TitleBar1.Transparency = ii
End Sub
