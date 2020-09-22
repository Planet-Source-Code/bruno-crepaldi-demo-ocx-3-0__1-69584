VERSION 5.00
Begin VB.Form FrmEx03 
   Appearance      =   0  'Flat
   BackColor       =   &H00CEB7AF&
   BorderStyle     =   0  'None
   Caption         =   "Form Esempio GlassButton"
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
   Icon            =   "FrmEx03.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmEx03.frx":164A
   ScaleHeight     =   7005
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      TextShadow      =   -1  'True
      Border          =   0   'False
      Caption         =   "NoBorder/NoIcon/Shadow"
      ColorBack       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      Caption         =   "NoIcon/TextRed/NoShadow"
      ColorBack       =   16711680
      ColorText       =   255
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
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      Caption         =   "Icon Left"
      ColorBack       =   65535
      Icon            =   "FrmEx03.frx":4F6F2
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
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      TextShadow      =   -1  'True
      Caption         =   "No Icon/Text Yellow"
      ColorText       =   65535
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
      Height          =   1095
      Index           =   10
      Left            =   2760
      TabIndex        =   6
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Transparency    =   220
      Caption         =   "Icon /Green Text  No Shadow Big Button"
      ColorText       =   65280
      Icon            =   "FrmEx03.frx":50D4C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   1215
      Index           =   4
      Left            =   3720
      TabIndex        =   7
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2143
      Transparency    =   220
      Border          =   0   'False
      ColorHotBorder  =   65280
      Icon            =   "FrmEx03.frx":523A6
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
      Height          =   1215
      Index           =   5
      Left            =   4560
      TabIndex        =   8
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2143
      Transparency    =   220
      Border          =   0   'False
      ColorHotBorder  =   65280
      Icon            =   "FrmEx03.frx":53A00
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
      Height          =   1215
      Index           =   11
      Left            =   5400
      TabIndex        =   9
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2143
      Transparency    =   220
      Border          =   0   'False
      ColorHotBorder  =   65280
      Icon            =   "FrmEx03.frx":5505A
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
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      Transparency    =   220
      TextShadow      =   -1  'True
      Caption         =   "No Icon / Shadow"
      ColorBack       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Blackletter686 BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   1095
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Transparency    =   220
      Caption         =   "Icon / Blu Text  No Shadow Big Button"
      ColorText       =   16711680
      Icon            =   "FrmEx03.frx":566B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   1095
      Index           =   12
      Left            =   6240
      TabIndex        =   12
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Transparency    =   220
      TextShadow      =   -1  'True
      Border          =   0   'False
      Caption         =   "Icon / RedText / Shadow / No Border"
      ColorText       =   255
      Icon            =   "FrmEx03.frx":57D0E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   1215
      Index           =   2
      Left            =   9720
      TabIndex        =   14
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Transparency    =   220
      ColorBack       =   16711935
      ColorHotBorder  =   16744576
      Icon            =   "FrmEx03.frx":59368
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
      Height          =   1215
      Index           =   3
      Left            =   9720
      TabIndex        =   15
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Transparency    =   220
      ColorBack       =   16711935
      ColorHotBorder  =   16744576
      Icon            =   "FrmEx03.frx":5A9C2
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
      Height          =   1095
      Index           =   13
      Left            =   2760
      TabIndex        =   16
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Enabled         =   0   'False
      Transparency    =   220
      TextShadow      =   -1  'True
      Border          =   0   'False
      Caption         =   "Disabled"
      ColorText       =   12582912
      Icon            =   "FrmEx03.frx":5C01C
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   1215
      Index           =   14
      Left            =   6240
      TabIndex        =   17
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2143
      Transparency    =   220
      Border          =   0   'False
      ColorHotBorder  =   65280
      Icon            =   "FrmEx03.frx":5D676
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
      Height          =   615
      Index           =   15
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      TextShadow      =   -1  'True
      Caption         =   "No Icon/Text Green"
      ColorText       =   65280
      ColorHotBorder  =   16711680
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   16
      Left            =   2760
      TabIndex        =   19
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      TextShadow      =   -1  'True
      Caption         =   "Align Left"
      Align           =   0
      ColorText       =   65280
      ColorHotBorder  =   16711680
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   17
      Left            =   5160
      TabIndex        =   20
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      TextShadow      =   -1  'True
      Caption         =   "align Center"
      ColorText       =   65280
      ColorHotBorder  =   16711680
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
   Begin DemoOCX30.GlassButton GlassButton2 
      Height          =   615
      Index           =   18
      Left            =   7560
      TabIndex        =   21
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Transparency    =   220
      TextShadow      =   -1  'True
      Caption         =   "Align Right"
      Align           =   2
      ColorText       =   65280
      ColorHotBorder  =   16711680
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
      TabIndex        =   13
      Top             =   6120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmEx03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GlassButton2_Click(Index As Integer)
 LblResult.Caption = Empty
End Sub
'
Private Sub GlassButton2_MouseDown(Index As Integer)
    LblResult.Caption = "Button" & Str(Index) & " Pressed"
End Sub
