VERSION 5.00
Begin VB.UserControl H_UpDown 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2400
   ScaleWidth      =   3645
   ToolboxBitmap   =   "H_UpDown.ctx":0000
   Begin VB.CommandButton Cmd_Left 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      DownPicture     =   "H_UpDown.ctx":0314
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MaskColor       =   &H8000000F&
      Picture         =   "H_UpDown.ctx":0664
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Cmd_Right 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      CausesValidation=   0   'False
      DownPicture     =   "H_UpDown.ctx":093C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      MaskColor       =   &H8000000F&
      Picture         =   "H_UpDown.ctx":0C8C
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   340
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "H_UpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Horizontal UpDown
' Nome del File ..: H_UpDown
' Data............: 27/11/2004
' Versione........: 1.31
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
Option Explicit

Private M_Value As Long
Private M_MinValue As Long
Private M_MaxValue As Long
Private M_LoopValue As Boolean
Private M_Step As Long
'                                Dichiarazione Eventi
Public Event Change(Value As Long)

'
'      Inizializza le Variabili ( Solo Progetazione )
'
Private Sub UserControl_InitProperties()
     M_Value = 0
     M_MinValue = 0
     M_MaxValue = 10
     M_Step = 1
     M_LoopValue = False
     UserControl.Height = 255
     UserControl.width = 1080
End Sub
'
'                        Resizing
'
Private Sub UserControl_Resize()
  
    Cmd_Left.Left = 0
    Cmd_Left.Height = ScaleHeight
 
    Cmd_Right.Left = ScaleWidth - 360
    Cmd_Right.Top = 0
    Cmd_Right.Height = ScaleHeight
 
    LblValue.Left = 360
    LblValue.Top = 0
    If ScaleWidth - (360 * 2) < 0 Then ' Corrected
        LblValue.width = 0
    Else
        LblValue.width = ScaleWidth - (360 * 2)
    End If
    
    LblValue.Height = ScaleHeight
    
    LblValue.FontSize = ScaleHeight / 22
 
End Sub
'
'                       inizializa
'
Private Sub UserControl_Initialize()
  LblValue.Caption = M_Value
End Sub
'
'                         Eventi
'
Private Sub ChangeEvent(Valore As Long)
    RaiseEvent Change(Valore)
End Sub
'
'                                Property
'
'
Public Property Get Value() As Long
   Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
   M_Value = NewValue
   PropertyChanged "Value"
   LblValue.Caption = Value
   ChangeEvent Value
End Property
'
Public Property Get MinValue() As Long
   MinValue = M_MinValue
End Property
Public Property Let MinValue(ByVal NewValue As Long)
   M_MinValue = NewValue
   PropertyChanged "MinValue"
    If M_Value < M_MinValue Then
        M_Value = M_MinValue
        PropertyChanged "Value"
    End If
End Property
'
Public Property Get MaxValue() As Long
   MaxValue = M_MaxValue
End Property
Public Property Let MaxValue(ByVal NewValue As Long)
   M_MaxValue = NewValue
   PropertyChanged "MaxValue"
End Property
'
Public Property Get Step() As Long
   Step = M_Step
End Property
Public Property Let Step(ByVal NewValue As Long)
   M_Step = NewValue
   PropertyChanged "Step"
End Property
'
Public Property Get LoopValue() As Boolean
  LoopValue = M_LoopValue
End Property
Public Property Let LoopValue(ByVal NewValue As Boolean)
   M_LoopValue = NewValue
   PropertyChanged "LoopValue"
End Property
'
Public Property Get BackColor() As OLE_COLOR
BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
UserControl.BackColor() = NewValue
PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = LblValue.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
LblValue.ForeColor = NewValue
PropertyChanged "ForeColor"
End Property
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
  LblValue.ForeColor = PropBag.ReadProperty("ForeColor", &H0)
  M_Value = PropBag.ReadProperty("Value", 0)
  M_MinValue = PropBag.ReadProperty("MinValue", 0)
  M_MaxValue = PropBag.ReadProperty("MaxValue", 5)
  M_Step = PropBag.ReadProperty("Step", 1)
  M_LoopValue = PropBag.ReadProperty("LoopValue", False)
  LblValue.Caption = M_Value
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
  Call PropBag.WriteProperty("ForeColor", LblValue.ForeColor, &H0)
  Call PropBag.WriteProperty("Value", M_Value, 0)
  Call PropBag.WriteProperty("MinValue", M_MinValue, 0)
  Call PropBag.WriteProperty("MaxValue", M_MaxValue, 5)
  Call PropBag.WriteProperty("Step", M_Step, 1)
  Call PropBag.WriteProperty("LoopValue", M_LoopValue, False)
End Sub
'
'              Cursori Updown
'

Private Sub Cmd_Left_Click()
'
    If M_Value <= M_MinValue Then
     If M_LoopValue = True Then
       M_Value = M_MaxValue + M_Step
      Else
       Exit Sub
     End If
    End If
'
    M_Value = M_Value - M_Step
    LblValue.Caption = M_Value
    ChangeEvent Value
End Sub

Private Sub Cmd_Right_Click()
'
    If M_Value >= M_MaxValue Then    ' -----------------------
     If M_LoopValue = True Then      '
       M_Value = M_MinValue - M_Step '  Controllo Fine Corsa
      Else                           '  LoopValue = true
       Exit Sub                      '  if False Exit sub
     End If                          '
    End If                           ' -----------------------
'
   
   M_Value = M_Value + M_Step
   LblValue.Caption = M_Value
   ChangeEvent Value
End Sub
