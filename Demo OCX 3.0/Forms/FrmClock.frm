VERSION 5.00
Begin VB.Form FrmClock 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0DECA&
   BorderStyle     =   0  'None
   Caption         =   "FrmClock"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmClock.frx":0000
   ScaleHeight     =   1695
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer TimerTransparency 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   600
      Top             =   1200
   End
   Begin VB.Timer TimerDisplay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   1200
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   7
      Left            =   6960
      TabIndex        =   0
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   6
      Left            =   6000
      TabIndex        =   1
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   5
      Left            =   5040
      TabIndex        =   2
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   4
      Left            =   4080
      TabIndex        =   3
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   3
      Left            =   3120
      TabIndex        =   4
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
   Begin DemoOCX30.DisplayLed Led_Time 
      Height          =   1320
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Transparency    =   160
      GlassEffect     =   0
      Zoom            =   4
   End
End
Attribute VB_Name = "FrmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ii As Integer
Private Flg00 As Long
Private Flg01 As Long
Private M_Step As Long
'---------------------------------------------------------------------------------------
' Form Load
'---------------------------------------------------------------------------------------
Private Sub Form_Load()
    ii = 0
    Flg00 = 0
    M_Step = 5
    '
    MakeTransparent FrmClock.hwnd, &HFF00FF, 255, LWA_ALPHA + LWA_COLORKEY
    If ScaleX(Screen.width, vbTwips, vbPixels) <= 1024 Then
        Flg01 = 1
    Else
        Flg01 = 0
    End If
    '
    For i = 0 To 7
        Led_Time(i).Style = Style_x3
        Led_Time(i).GlassEffect = False
        Led_Time(i).Transparency = 220
        Led_Time(i).LedColor = vbBlack
        Led_Time(i).Zoom = Zoom_x4
    Next
    '
    TimerTransparency.Enabled = True
    TimerDisplay.Enabled = True
    TimerTransparency_Timer
    TimerDisplay_Timer
    '
    Call FormOnTop(Me.hwnd, True)
    '
End Sub
'---------------------------------------------------------------------------------------
' Form Unload
'---------------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Call FormOnTop(Me.hwnd, False)

    TimerTransparency.Enabled = False
    TimerDisplay.Enabled = False
End Sub
'---------------------------------------------------------------------------------------
' TimerDisplay
'---------------------------------------------------------------------------------------
Private Sub TimerDisplay_Timer()
Dim I1 As Long
Dim Ch As String * 1
Dim Str As String

    Str = Trim(Time)
    If Len(Str) = 7 Then Str = "0" + Str
    For I1 = 0 To 7
        Ch = Mid(Str, I1 + 1, 1)
        If Ch = "." Then Ch = ":"
        Led_Time(I1).Value = Asc(Ch)
    Next I1
End Sub
'---------------------------------------------------------------------------------------
' TimerTransparency
'---------------------------------------------------------------------------------------
Private Sub TimerTransparency_Timer()
    
    If Flg00 = 0 Then
        Select Case Flg01
            Case 0 ' Alto e Destra
                Me.Left = Screen.width - Me.width - 120
                Me.Top = 500
            Case 1 ' Basso e Destra
                Me.Left = Screen.width - Me.width - 120
                Me.Top = Screen.Height - Me.Height - 500
            Case 2 ' Basso e Sinistra
                Me.Left = 120
                Me.Top = Screen.Height - Me.Height - 500
            Case 3 ' Alto e Sinistra
                Me.Left = 120
                Me.Top = 500
        End Select

        If Flg01 < 2 Then
            Flg01 = Flg01 + 1
        Else
            If ScaleX(Screen.width, vbTwips, vbPixels) <= 1024 Then
                Flg01 = 1
            Else
                Flg01 = 0
            End If
        End If
    
    End If
    '
    If Flg00 = 500 Then Flg00 = -500
        
    Flg00 = Flg00 + M_Step

    If (ii < 180) And (Flg00 > 0 And Flg00 < 500) Then
        ii = ii + M_Step
        MakeTransparent FrmClock.hwnd, &HFF00FF, ii, LWA_ALPHA + LWA_COLORKEY
    End If
        
    If (ii > 0) And (Flg00 < 0) Then
        ii = ii - M_Step
        If ii = 0 Then Flg00 = 0
        MakeTransparent FrmClock.hwnd, &HFF00FF, ii, LWA_ALPHA + LWA_COLORKEY
    End If
    DoEvents
End Sub
