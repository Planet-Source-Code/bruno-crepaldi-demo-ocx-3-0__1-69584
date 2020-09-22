VERSION 5.00
Begin VB.Form FrmEx04 
   Appearance      =   0  'Flat
   BackColor       =   &H00CEB7AF&
   BorderStyle     =   0  'None
   Caption         =   "Form Esempio AlphaText"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEx04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmEx04.frx":164A
   ScaleHeight     =   7005
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enable/Disable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   5520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer 
      Interval        =   10
      Left            =   120
      Top             =   360
   End
   Begin DemoOCX30.GlassButton GlassButton1 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Change Text"
      ColorText       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin DemoOCX30.BottomBar BottomBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6705
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   529
      Resize          =   0   'False
   End
   Begin DemoOCX30.AlphaText AlphaText2 
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   2990
      Transparency    =   255
      TextShadow      =   -1  'True
      Caption         =   "Prova STAMPA"
      ColorText       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Highlight LET"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.HSlider HSlider1 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      MaxValue        =   255
      Picture         =   "FrmEx04.frx":F7ADE
      PicCursor_Selected=   "FrmEx04.frx":F904A
      PictureCursor   =   "FrmEx04.frx":F9422
   End
   Begin DemoOCX30.GlassButton GlassButton1 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Change Color"
      ColorText       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Now you can drag text with the Mouse"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shadow"
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
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transparency"
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
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label LblResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmEx04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Col_Sw As Long
'----------------------------------------------------------------------
' Form_Load
'----------------------------------------------------------------------
Private Sub Form_Load()
    Col_Sw = 0
End Sub
'----------------------------------------------------------------------
' GlassButton1
'----------------------------------------------------------------------
Private Sub GlassButton1_Click(Index As Integer)
    Select Case Index
        Case 0  ' Caption
            If AlphaText2.Caption = "Prova STAMPA" Then
                AlphaText2.Caption = "Hello World"
            Else
                AlphaText2.Caption = "Prova STAMPA"
            End If
        Case 1  ' Color
            Col_Sw = Col_Sw + 1: If Col_Sw > 2 Then Col_Sw = 0
            Select Case Col_Sw
                Case 0
                    AlphaText2.ColorText = &HFF0000
                Case 1
                    AlphaText2.ColorText = &HFF00&
                Case 2
                    AlphaText2.ColorText = &HFF
            End Select
     End Select
End Sub
'----------------------------------------------------------------------
' Check1
'----------------------------------------------------------------------
Private Sub Check1_Click()
    AlphaText2.TextShadow = -Check1.Value
End Sub
'----------------------------------------------------------------------
' HSlider1
'----------------------------------------------------------------------
Private Sub HSlider1_scroll(Value As Long)
    AlphaText2.Transparency = Value
End Sub
'----------------------------------------------------------------------
' Timer
'----------------------------------------------------------------------
Private Sub Timer_Timer()
    If AlphaText2.Transparency < 15 Then
        AlphaText2.Transparency = 0
        GlassButton1(0).Visible = True
        GlassButton1(1).Visible = True
        Label1.Visible = True
        Check1.Visible = True
        Label9.Visible = True
        HSlider1.Visible = True
        Label2.Visible = True
        Timer.Enabled = False: Exit Sub
    End If
    AlphaText2.Transparency = AlphaText2.Transparency - 15
End Sub
