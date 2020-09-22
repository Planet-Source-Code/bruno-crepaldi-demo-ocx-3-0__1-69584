Attribute VB_Name = "Mdl_Forms"
' Retrieves the DC for the entire window, including title bar,
' menus, and scroll bars. A window DC permits painting anywhere
' in a window, because the origin of the DC is the upper-left
' corner of the window instead of the client area.
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
' Returns a handle to the Desktop window.  The desktop
' window covers the entire screen and is the area on top
' of which all icons and other windows are painted.
Private Declare Function GetDesktopWindow Lib "user32" () As Long

' AlphaBlend
Private Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal WidthSrc As Long, _
  ByVal HeightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean
  

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, Source As Any, ByVal Length As Long)

Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Private tProperties As typeBlendProperties
Private lngBlend As Long
Private m_left As Long
Private m_Top As Long
Private m_Width As Long
Private m_Height As Long

Private hDCSrc As Long
Private hDCDst As Long

' Form Trasparente
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wflags As Long) As Long
'API Constant Declarations
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const LWA_OPAQUE = &H4
Private Const BM_SETSTATE = &HF3
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
'
Private Pictmp As PictureBox
' SetGlass
Type vRGB
  r As Byte
  G As Byte
  B As Byte
End Type
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Dim iRGB  As vRGB
'
Dim i As Long
'------------------------------------------------------------------------
' AlphaBlend 2
'------------------------------------------------------------------------
Public Sub GoAlpha2(myPict As Object, Alpha As Byte)

'Dim hDCSrc As Long
'Dim hDCDst As Long
    
    myPict.Cls
    '
    tProperties.tBlendAmount = Alpha 'Set translucency level
    
    m_Width = myPict.width / Screen.TwipsPerPixelX
    m_Height = myPict.Height / Screen.TwipsPerPixelY
    m_left = myPict.Left / Screen.TwipsPerPixelX
    m_Top = myPict.Top / Screen.TwipsPerPixelY
    '
    
    hDCSrc = myPict.Parent.hdc
    hDCDst = myPict.hdc
    
    myPict.Visible = False

    DoEvents
    
    CopyMemory lngBlend, tProperties, 4 'Blend colors
    AlphaBlend hDCDst, 0, 0, m_Width, m_Height, hDCSrc, m_left, m_Top, m_Width, m_Height, lngBlend 'Blend together
    '
    myPict.Visible = True
End Sub
'------------------------------------------------------------------------
' Glass Tipo 0
'------------------------------------------------------------------------
Public Sub SetGlass(myPict As Object)
Dim Sta As Integer
Dim m_X As Long
Dim m_Y As Long
Dim m_V As Long
Dim m_T As Long
Dim m_H As Long
Dim m_X1 As Long
Dim m_Y1 As Long
Dim m_X2 As Long
Dim m_Y2 As Long

    m_X1 = 0
    m_Y1 = 0
  
    m_X2 = myPict.width / Screen.TwipsPerPixelX
    m_Y2 = myPict.Height / Screen.TwipsPerPixelY
    '
    m_H = (m_Y2 / 100) * 5
    If m_H > 15 Then m_H = 15
' Parte Alta
    For m_Y = m_H To 0 Step -1
           
        Sta = (m_H - m_Y) * 6
        
        For m_X = m_X1 To m_X1 + m_X2
            m_V = myPict.Point(m_X, m_Y1 + m_Y)

            CopyMemory iRGB, m_V, LenB(iRGB)
            
            m_V = iRGB.r + Sta
            If m_V > 255 Then m_V = 255
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G + Sta
            If m_V > 255 Then m_V = 255
            iRGB.G = Int(m_V)
                        
            m_V = iRGB.B + Sta
            If m_V > 255 Then m_V = 255
            iRGB.B = m_V

            myPict.ForeColor = RGB(iRGB.r, iRGB.G, iRGB.B)
            myPict.PSet (m_X, m_Y1 + m_Y)
        Next
    Next
    DoEvents
' Parte Bassa
    m_T = m_H
    For m_Y = m_Y2 - m_H To m_Y2
        Sta = ((m_H - m_T) * 6)
        For m_X = m_X1 To m_X1 + m_X2
            m_V = myPict.Point(m_X, m_Y1 + m_Y)
            CopyMemory iRGB, m_V, LenB(iRGB)

            m_V = iRGB.r + Sta
            If m_V > 255 Then m_V = 255
            If m_V < 0 Then m_V = 0
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G + Sta
            If m_V > 255 Then m_V = 255
            If m_V < 0 Then m_V = 0
            iRGB.G = Int(m_V)
                        
            m_V = iRGB.B + Sta
            If m_V > 255 Then m_V = 255
            If m_V < 0 Then m_V = 0
            iRGB.B = m_V
        
            myPict.ForeColor = RGB(iRGB.r, iRGB.G, iRGB.B)
            myPict.PSet (m_X, m_Y1 + m_Y)
        Next
        m_T = m_T - 1
    Next
    DoEvents
End Sub
'------------------------------------------------------------------------
' Glass Tipo 1
'------------------------------------------------------------------------
Public Sub SetGlass1(myPict As Object)
Dim Sta As Integer
Dim m_X As Long
Dim m_Y As Long
Dim m_V As Long
Dim m_T As Long
Dim m_H As Long
Dim m_X1 As Long
Dim m_Y1 As Long
Dim m_X2 As Long
Dim m_Y2 As Long

    m_X1 = 0
    m_Y1 = 0
  
    m_X2 = myPict.width / Screen.TwipsPerPixelX
    m_Y2 = myPict.Height / Screen.TwipsPerPixelY
    '
    m_H = (m_Y2 / 100) * 20
    If m_H > 15 Then m_H = 15
' Parte Alta
    For m_Y = m_H To 0 Step -1
        
        Sta = (m_H - m_Y) * 6
        
        For m_X = m_X1 To m_X1 + m_X2
            m_V = myPict.Point(m_X, m_Y1 + m_Y)

            CopyMemory iRGB, m_V, LenB(iRGB)
            
            m_V = iRGB.r + Sta
            If m_V > 255 Then m_V = 255
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G + Sta
            If m_V > 255 Then m_V = 255
            iRGB.G = Int(m_V)
                        
            m_V = iRGB.B + Sta
            If m_V > 255 Then m_V = 255
            iRGB.B = m_V

            myPict.ForeColor = RGB(iRGB.r, iRGB.G, iRGB.B)
            myPict.PSet (m_X, m_Y1 + m_Y)
        Next
    Next
    DoEvents
' Parte Bassa
    m_T = m_H
    For m_Y = m_Y2 - m_H To m_Y2
        Sta = ((m_H - m_T) * 6)
        For m_X = m_X1 To m_X1 + m_X2
            m_V = myPict.Point(m_X, m_Y1 + m_Y)
            CopyMemory iRGB, m_V, LenB(iRGB)

            m_V = iRGB.r - Sta
            If m_V < 0 Then m_V = 0
            iRGB.r = Int(m_V)
            
            m_V = iRGB.G - Sta
            If m_V < 0 Then m_V = 0
            iRGB.G = Int(m_V)
                        
            m_V = iRGB.B - Sta
            If m_V < 0 Then m_V = 0
            iRGB.B = m_V
        
            myPict.ForeColor = RGB(iRGB.r, iRGB.G, iRGB.B)
            myPict.PSet (m_X, m_Y1 + m_Y)
        Next
        m_T = m_T - 1
    Next
    DoEvents
End Sub
'------------------------------------------------------------------------
' StretchPicture
'------------------------------------------------------------------------
Public Sub StretchPicture(M_Obj As Object)
Dim X1, Y1, L1, H1, X2, Y2, L2, H2 As Long
Dim Redraw As Boolean

    Set Pictmp = M_Obj.Controls.Add("VB.Picturebox", "PicTmp")
    Pictmp.Picture = M_Obj.Picture
    With Pictmp
        .Visible = False
        .AutoRedraw = True
        .AutoSize = True
    End With
    Redraw = M_Obj.AutoRedraw
    M_Obj.AutoRedraw = True
                
    Select Case M_Obj.Scalemode
        Case 1
            X1 = 0                                          ' Destinazione
            Y1 = 0                                          '
            L1 = M_Obj.width                                ' Area Destinazione
            H1 = M_Obj.Height                               ' Area Destinazione
            X2 = 0                                          ' Sorgente
            Y2 = 0                                          '
            L2 = Pictmp.width                               ' Area Sorgente
            H2 = Pictmp.Height                              ' Area Sorgente
        Case 3
            X1 = 0                                          ' Destinazione
            Y1 = 0                                          '
            L1 = M_Obj.width / Screen.TwipsPerPixelX       ' Area Destinazione
            H1 = M_Obj.Height / Screen.TwipsPerPixelY       ' Area Destinazione
            X2 = 0                                          ' Sorgente
            Y2 = 0                                          '
            L2 = Pictmp.ScaleWidth / Screen.TwipsPerPixelX  ' Area Sorgente
            H2 = Pictmp.ScaleHeight / Screen.TwipsPerPixelY ' Area Sorgente
    End Select
    '
    
    M_Obj.PaintPicture Pictmp.Picture, X1, Y1, L1, H1, X2, Y2, L2, H2, vbSrcCopy
    
    M_Obj.Refresh
    M_Obj.Controls.Remove "PicTmp"
    M_Obj.AutoRedraw = Redraw

End Sub
'-----------------------------------------------------------------------
' Form Trasparente
' Es: 1) MakeTransparent(Me.hWnd, &H00FF00FF&,0, LWA_COLORKEY) As Long
' Es: 2) MakeTransparent(Me.hWnd, 0, 0-255 , LWA_ALPHA) As Long
'-----------------------------------------------------------------------
Public Function MakeTransparent(ByVal hwnd As Long, clr As Long, tValue As Integer, dFlags As Long) As Long
    Dim Msg As Long
    If tValue < 0 Or tValue > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, clr, tValue, dFlags
        MakeTransparent = 0
    End If
    If Err Then
        MakeTransparent = 2
    End If
End Function
'-----------------------------------------------------------------------
' Form on Top
'-----------------------------------------------------------------------
Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
Dim wflags, Placement
' Example: Call FormOnTop(me.hWnd, True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    wflags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wflags
End Sub
