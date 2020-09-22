VERSION 5.00
Begin VB.Form FrmAlpha 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "FrmAlpha"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   Icon            =   "FrmAlpha.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmAlpha.frx":164A
   ScaleHeight     =   467
   ScaleMode       =   0  'User
   ScaleWidth      =   667
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   5400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   600
      Top             =   5400
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3435
      Left            =   5280
      Picture         =   "FrmAlpha.frx":3C168E
      ScaleHeight     =   3375
      ScaleWidth      =   4500
      TabIndex        =   3
      Top             =   1800
      Width           =   4560
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   3435
      Left            =   120
      Picture         =   "FrmAlpha.frx":3E2982
      ScaleHeight     =   3375
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   1800
      Width           =   4560
   End
   Begin DemoOCX30.BottomBar BottomBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   529
   End
   Begin DemoOCX30.TitleBar TitleBar1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   529
      Transparency    =   200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AlphaBlend Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   1215
      Left            =   1560
      TabIndex        =   4
      Top             =   5400
      Width           =   6855
   End
End
Attribute VB_Name = "FrmAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Not Mine !!!
'
Const AC_SRC_OVER = &H0

Private Type BLENDFUNCTION
 BlendOp As Byte
 BlendFlags As Byte
 SourceConstantAlpha As Byte
 AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" _
  (ByVal hdc As Long, _
  ByVal lInt As Long, _
  ByVal lInt As Long, _
  ByVal lInt As Long, _
  ByVal lInt As Long, _
  ByVal hdc As Long, _
  ByVal lInt As Long, _
  ByVal lInt As Long, _
  ByVal lInt As Long, _
  ByVal lInt As Long, _
  ByVal BLENDFUNCT As Long) As Long
  
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
  (Destination As Any, _
  Source As Any, _
  ByVal Length As Long)

Dim BlendVal As Integer
Dim BF As BLENDFUNCTION, lBF As Long

'//Variable to hold images so we can swap
Dim tempPic1 As New StdPicture, tempPic2 As New StdPicture
'----------------------------------------------------------------------
' Form Initialize
'----------------------------------------------------------------------
Private Sub Form_Initialize()
'    Forms.Add Me
    BottomBar1.Caption = Me.Name
End Sub
'----------------------------------------------------------------------
' Form Load
'----------------------------------------------------------------------
Private Sub Form_Load()
    Label1.Caption = "Basilica Santuario Santo Stefano" & vbCrLf & "Bologna" & vbCrLf & "Italy"
  
    With Picture1
        .AutoRedraw = True
        .Scalemode = vbPixels
    End With
    
    With Picture2
        .AutoRedraw = True
        .Scalemode = vbPixels
    End With
  
    'set the parameters
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    BlendVal = 1
    Set tempPic1 = Picture1.Picture
    Set tempPic2 = Picture2.Picture
  
    With Timer1
        .Interval = 60
        .Enabled = True
    End With
    
    With Timer2
        .Interval = 60
        .Enabled = False
    End With
End Sub
'----------------------------------------------------------------------
' Form_Unload
'----------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer2.Enabled = False
End Sub
'----------------------------------------------------------------------
' DoAlphablend
'----------------------------------------------------------------------
Public Sub DoAlphablend(SrcPicBox As PictureBox, DestPicBox As PictureBox, AlphaVal As Integer)
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = AlphaVal
        .AlphaFormat = 0
    End With
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend DestPicBox.hdc, 0, 0, DestPicBox.ScaleWidth, DestPicBox.ScaleHeight, SrcPicBox.hdc, 0, 0, SrcPicBox.ScaleWidth, SrcPicBox.ScaleHeight, lBF
End Sub
'----------------------------------------------------------------------
' Timer
'----------------------------------------------------------------------
Private Sub Timer1_Timer()
    Picture1.Refresh
    Picture2.Refresh

    BlendVal = BlendVal + 5
    If BlendVal >= 155 Then
        Flag = True
        Timer1.Enabled = False
        Picture2.Picture = tempPic1
        Timer2.Enabled = True
        BlendVal = 1
    End If
  
    DoAlphablend Picture2, Picture1, BlendVal

    Me.Caption = CStr(BlendVal)
End Sub
'----------------------------------------------------------------------
' Timer
'----------------------------------------------------------------------
Private Sub Timer2_Timer()
    Picture1.Refresh
    Picture2.Refresh

    BlendVal = BlendVal + 5
    If BlendVal >= 155 Then
        BlendVal = 1
        Timer1.Enabled = True
        Timer2.Enabled = False
        Picture2.Picture = tempPic2
    End If
    DoAlphablend Picture2, Picture1, BlendVal

    Me.Caption = CStr(BlendVal)
End Sub

