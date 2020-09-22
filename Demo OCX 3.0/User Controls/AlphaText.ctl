VERSION 5.00
Begin VB.UserControl AlphaText 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   94
End
Attribute VB_Name = "AlphaText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: AlphaText
' Nome del File ..: AlphaText
' Data............: 01/10/2007
' Versione........: 0.1 beta
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
'Option Explicit
'
' DrawText
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private m_Rect As RECT

' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
'
Private m_Caption As String
Private m_BackColor As String
Private m_TextShadow As Boolean ' Corrected
Private m_Trsp As Long
Private m_TrspSh As Long
Private m_Redraw As Boolean
Private m_SclMd As Long
Private i As Long
'Private m_ForeColor As OLE_COLOR

Private m_X, m_Y As Long
Private Pict_Src As PictureBox
Private Pict_Msk As PictureBox
Private MouseIsDown As Boolean

'
Enum TxtAlpha_Align
    nLeft = DT_LEFT
    nCenter = DT_CENTER
    nRight = DT_RIGHT
End Enum
Private m_Align As TxtAlpha_Align
'
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Const BI_RGB = 0&
Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private m_Width As Long
Private m_Height As Long

Private msk As Long, MSKO1 As Long, MSKI As BITMAPINFO, MSKBITS() As Byte
Private nSRC As Long, nSRCO1 As Long, nSRCI As BITMAPINFO, SRCBITS() As Byte
Private DST As Long, DSTO1 As Long, DSTI As BITMAPINFO, DSTBITS() As Byte
Private BB As Long, BBO As Long
Private X_Base, Y_Base As Long
Private LX As Long, LY As Long
'
Private WithEvents M_Frm As Form
Attribute M_Frm.VB_VarHelpID = -1
'----------------------------------------------------------
' m_Frm Events ( For Capture Parent Events  )
'----------------------------------------------------------
Private Sub M_Frm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = True
    X_Base = X - Extender.Left
    Y_Base = Y - Extender.Top
End Sub
'
Private Sub M_Frm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown = False
End Sub
'
Private Sub M_Frm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MouseIsDown = False Then Exit Sub
        
    With UserControl
        If X >= .Extender.Left And X <= (.Extender.Left + .width) And Y >= .Extender.Top And Y <= (.Extender.Top + .Height) Then
            .Extender.Left = (X - X_Base)
            .Extender.Top = (Y - Y_Base)
            Select Case UserControl.Parent.Scalemode
                Case vbTwips
                    m_X = ScaleX((X - X_Base), vbTwips, vbPixels)   ' convert Twips/Pixels
                    m_Y = ScaleY((Y - Y_Base), vbTwips, vbPixels)   ' convert Twips/Pixels
                Case vbPixels
                    m_X = (X - X_Base)
                    m_Y = (Y - Y_Base)
            End Select
            BLTIT
        End If
    End With

End Sub
'-----------------------------------------------------------------------------------------------
' Inizializza le Variabili ( Solo Progetazione )
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    On Error Resume Next
    Set M_Frm = UserControl.Parent
     
    m_Align = nCenter
    m_BackColor = &H404040
    m_TextShadow = False
    m_Trsp = 0
    m_TrspSh = 143 + (m_Trsp / 2)
    MouseIsDown = False
    '
    Set Font = Ambient.Font
End Sub
'----------------------------------------------------------------------------------------
' usercontrol_Initialize
'----------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    Set Pict_Src = UserControl.Controls.Add("VB.Picturebox", "Pict_Src")
    Set Pict_Msk = UserControl.Controls.Add("VB.Picturebox", "Pict_Msk")
    With Pict_Src
        .AutoRedraw = True
        .Appearance = 0
        .BackColor = vbWhite
        .BorderStyle = 0
        .Enabled = True
        .Scalemode = vbPixels
        .Visible = False
    End With
    With Pict_Msk
        .AutoRedraw = True
        .Appearance = 0
        .BackColor = vbWhite
        .BorderStyle = 0
        .Enabled = True
        .Scalemode = vbPixels
        .Visible = False
    End With
    
    m_TrspSh = 143 + (m_Trsp / 2)
    If m_TrspSh > 255 Then m_TrspSh = 255
End Sub
'----------------------------------------------------------------------------------------
' usercontrol_Show
'----------------------------------------------------------------------------------------
Private Sub UserControl_Show()
    'setup for alphablitting
    On Error Resume Next
    Set M_Frm = UserControl.Parent

    If m_TextShadow = True Then
        ' Calcul Shadow Value
        m_TrspSh = 143 + (m_Trsp / 2)
        If m_TrspSh > 255 Then m_TrspSh = 255
        ' Print Shadow
        SizeCaption Pict_Src, &H0, 15, 5
        SizeCaption Pict_Msk, RGB(m_TrspSh, m_TrspSh, m_TrspSh), 15, 5
    End If
    ' Setup & execute
    Setup0 Pict_Src.hdc, Pict_Msk.hdc
    BLTIT
End Sub
'----------------------------------------------------------------------------------------
' usercontrol_Resize
'----------------------------------------------------------------------------------------
Private Sub UserControl_Resize()
    With Pict_Src
        .Height = UserControl.ScaleHeight - 15
        .width = UserControl.ScaleWidth - 15
    End With
    With Pict_Msk
        .Height = UserControl.ScaleHeight - 15
        .width = UserControl.ScaleWidth - 15
    End With
End Sub
'----------------------------------------------------------------------------------------
' usercontrol_Terminate
'----------------------------------------------------------------------------------------
Private Sub UserControl_Terminate()
    Cleanup 'delete all memory consumers
End Sub
'-----------------------------------------------------------------------------------------------
' Property Let Set
'-----------------------------------------------------------------------------------------------
' Font
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Set Font(new_Font As StdFont)
    Set UserControl.Font = new_Font
    Set Pict_Src.Font = new_Font
    Set Pict_Msk.Font = new_Font
    PropertyChanged "Font"
'    Setup Pict_Src.hdc, Pict_Msk.hdc 'Pict_Src, Pict_Msk
'    BLTIT
End Property
' Caption
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(NewValue As String)
    m_Caption = NewValue
    PropertyChanged "Caption"
    ' Clear Pictures
    Pict_Src.Cls
    Pict_Msk.Cls
    ' Redraw
    Redraw
End Property
' Align
Public Property Get Align() As TxtAlpha_Align
   Align = m_Align
End Property
Public Property Let Align(ByVal NewValue As TxtAlpha_Align)
    m_Align = NewValue
    PropertyChanged "Align"
    SizeCaption Pict_Src, UserControl.ForeColor
    ' Clear Pictures
    Pict_Src.Cls
    Pict_Msk.Cls
    ' Redraw
    Redraw
End Property
' ColorText
Public Property Get ColorText() As OLE_COLOR
   ColorText = UserControl.ForeColor
End Property
Public Property Let ColorText(ByVal NewValue As OLE_COLOR)
    PropertyChanged "ColorText"
    UserControl.ForeColor = NewValue
    SizeCaption Pict_Src, UserControl.ForeColor
    ' Clear Pictures
    Pict_Src.Cls
    Pict_Msk.Cls
    ' Redraw
    Redraw
End Property
' TextShadow
Public Property Get TextShadow() As Boolean
    TextShadow = m_TextShadow
End Property
Public Property Let TextShadow(bTextShadow As Boolean)
    m_TextShadow = bTextShadow
    PropertyChanged "TextShadow"
    SizeCaption Pict_Src, UserControl.ForeColor
    Pict_Src.Cls
    Pict_Msk.Cls
    ' Redraw
    Redraw
End Property
' Transparency
Public Property Let Transparency(bTransparency As Byte)
    m_Trsp = bTransparency
    PropertyChanged "Transparency"
    Redraw
End Property
Public Property Get Transparency() As Byte
    Transparency = m_Trsp
End Property
'-----------------------------------------------------------------------------------------------
' Property Read Write
'-----------------------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Transparency", m_Trsp, 0
        .WriteProperty "TextShadow", m_TextShadow, False
        .WriteProperty "Caption", m_Caption, Empty
        .WriteProperty "Align", m_Align, TxtAlpha_Align.nCenter
        .WriteProperty "ColorText", ColorText, &H0
        .WriteProperty "Font", Font, Ambient.Font
    End With
End Sub
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Trsp = .ReadProperty("Transparency", 0)
        m_TextShadow = .ReadProperty("TextShadow", False)
        m_Caption = .ReadProperty("Caption", Empty)
        m_Align = .ReadProperty("Align", TxtAlpha_Align.nCenter)
        UserControl.ForeColor = .ReadProperty("ColorText", &H0)
        Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Redraw
'-----------------------------------------------------------------------------------------------
Private Sub Redraw()
    If m_TextShadow = True Then
        ' Calcul Shadow Value
        m_TrspSh = 143 + (m_Trsp / 2)
        If m_TrspSh > 255 Then m_TrspSh = 255
        ' Print Shadow
        SizeCaption Pict_Src, &H0, 15, 5
        SizeCaption Pict_Msk, RGB(m_TrspSh, m_TrspSh, m_TrspSh), 15, 5
    End If
    ' Print Text
    SizeCaption Pict_Src, UserControl.ForeColor
    SizeCaption Pict_Msk, RGB(m_Trsp, m_Trsp, m_Trsp)
    ' Setup1 & execute
    Setup1 Pict_Src.hdc, Pict_Msk.hdc
    BLTIT
End Sub
'-----------------------------------------------------------------------------------------------
' SizeCaption
'-----------------------------------------------------------------------------------------------
Private Sub SizeCaption(My_Obj As Object, Colore As OLE_COLOR, Optional Shift_X As Long, Optional Shift_Y As Long)
Dim Res, m_Bt, m_W, m_H, m_Mrg As Long
Dim m_St As Long

        
    My_Obj.ForeColor = Colore


    ' Define the Rectangle for the Text - if Multiline must put the width do you want in m_Rect.Right
    m_Mrg = 5
    m_Rect.Top = Shift_Y
    m_Rect.Left = Shift_X
    m_Rect.Right = My_Obj.ScaleWidth

    m_St = DT_CALCRECT Or m_Align Or DT_WORDBREAK
        
    Res = DrawText(My_Obj.hdc, m_Caption, -1, m_Rect, m_St)

    m_Rect.Right = My_Obj.ScaleWidth
    m_Bt = m_Rect.Bottom

    m_Rect.Top = Shift_Y
    m_Rect.Bottom = m_Bt + m_Rect.Top
    
    m_St = m_Align Or DT_WORDBREAK
    
    DrawText My_Obj.hdc, m_Caption, Len(m_Caption), m_Rect, m_St ' Print text

End Sub
'----------------------------------------------------------------------------------------
' Setup
'----------------------------------------------------------------------------------------
Private Sub Setup0(Hdc1 As Long, Hdc2 As Long)
    m_Redraw = UserControl.Parent.AutoRedraw
    m_SclMd = UserControl.Parent.Scalemode
    
    'set the width and height
    UserControl.Parent.Scalemode = vbPixels ' Cambia Scalemode form
     
    m_Width = Pict_Src.ScaleWidth
    m_Height = Pict_Src.ScaleHeight
    m_X = UserControl.Extender.Left
    m_Y = UserControl.Extender.Top
    LX = m_X
    LY = m_Y

    'set bitmap info for the Pict_Src, Pict_Msk, and destination bitmaps
    With MSKI.bmiHeader
        .biBitCount = 24        '24 bits per pixel (R,G,B per pixel)
        .biSize = Len(MSKI)     'size of this information
        .biHeight = m_Height    'height
        .biWidth = m_Width      'width
        .biPlanes = 1           'bitmap planes (2D, so 1)
        .biCompression = BI_RGB 'Type of color compression
    End With
    'the following is the same for all bitmaps
    With DSTI.bmiHeader
        .biBitCount = 24
        .biSize = Len(DSTI)
        .biHeight = m_Height
        .biWidth = m_Width
        .biPlanes = 1
        .biCompression = BI_RGB
    End With
    With nSRCI.bmiHeader
        .biBitCount = 24
        .biSize = Len(nSRCI)
        .biHeight = m_Height
        .biPlanes = 1
        .biWidth = m_Width
        .biCompression = BI_RGB
    End With

    'create the device contexts
    msk = CreateCompatibleDC(GetDC(0))
    nSRC = CreateCompatibleDC(GetDC(0))
    DST = CreateCompatibleDC(GetDC(0))
    BB = CreateCompatibleDC(GetDC(0))

    'variable that defines how many color bits there are in one bit array
    '[Width * Height] (all pixels) [* 3] (R,G,B - 3 values) per pixel
    Dim nl As Long
    nl = ((m_Width + 1) * (m_Height + 1)) * 3

    'redimension the bit color information arrays to fit all the color information
    ReDim MSKBITS(1 To nl)
    ReDim SRCBITS(1 To nl)
    ReDim DSTBITS(1 To nl)

    'create a DIB section based on the bitmapinfo we provided above.
    'this is like creating a compatible bitmap, but used for modifying bitmap bits
    MSKO1 = CreateDIBSection(GetDC(0), MSKI, DIB_RGB_COLORS, 0, 0, 0)
    nSRCO1 = CreateDIBSection(GetDC(0), nSRCI, DIB_RGB_COLORS, 0, 0, 0)
    DSTO1 = CreateDIBSection(GetDC(0), DSTI, DIB_RGB_COLORS, 0, 0, 0)

    'create a permanent image of the form, so we can restore drawn-over parts
    BBO = CreateCompatibleBitmap(GetDC(0), UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight)

    'link the device contexts to thier bitmap objects
    SelectObject msk, MSKO1
    SelectObject DST, DSTO1
    SelectObject nSRC, nSRCO1
    SelectObject BB, BBO
    '
     UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    ' copia tutto il form in BB
    BitBlt BB, 0, 0, UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight, UserControl.Parent.hdc, 0, 0, vbSrcCopy
    '
    UserControl.Parent.AutoRedraw = m_Redraw ' Very Important Autoredraw must be True
    
    UserControl.Parent.Scalemode = m_SclMd  ' Ripristina scalmode del form
   
    Setup1 Hdc1, Hdc2
End Sub

Private Sub Setup1(Hdc1 As Long, Hdc2 As Long)

    'blt the Pict_Msk and Pict_Src images into the bitmap object so we can copy the color information
    BitBlt msk, 0, 0, m_Width, m_Height, Hdc2, 0, 0, vbSrcCopy
    BitBlt nSRC, 0, 0, m_Width, m_Height, Hdc1, 0, 0, vbSrcCopy
    'load the color information into arrays we only do this once because the Pict_Src and Pict_Msk
    'images never change, but the destination image will change frequently, depending on where the mouse
    'is on the form, so we have to update the DST bit array every time we alphablt.

    GetDIBits msk, MSKO1, 0, m_Height, MSKBITS(1), MSKI, DIB_RGB_COLORS
    GetDIBits nSRC, nSRCO1, 0, m_Height, SRCBITS(1), nSRCI, DIB_RGB_COLORS

End Sub
'----------------------------------------------------------------------------------------
' BLTIT
'----------------------------------------------------------------------------------------
Private Sub BLTIT()
    m_Redraw = UserControl.Parent.AutoRedraw

    'copy image from the permanant image of the form to the destination bitmap, so we have a 'background'
    BitBlt DST, 0, 0, m_Width, m_Height, BB, m_X, m_Y, vbSrcCopy
    
    'copy the destination image data into its bit array so we can process it
    GetDIBits DST, DSTO1, 0, m_Height, DSTBITS(1), DSTI, DIB_RGB_COLORS

    'some processing variables
    Dim SrcC(2) As Integer
    Dim DstC(2) As Integer
    Dim Alpha(2) As Integer
    Dim tmp(2) As Integer

    'bit array Temporaneo
    Dim tmpBits() As Byte

    'make the temporary bit array large enough to hold all the color information from the resulting alpha blitted bitmap
    ReDim tmpBits(UBound(SRCBITS))

    'a for loop to loop through the pixels of the bitmaps we do step3 because for every pixel, there are RED,GREEN, and BLUE color values in the bit array
    For i = 1 To UBound(SRCBITS) Step 3
        'pixel: (i) to (i+2)
        SrcC(0) = SRCBITS(i)     'blue value
        SrcC(1) = SRCBITS(i + 1) 'green value
        SrcC(2) = SRCBITS(i + 2) 'red value
    
        Alpha(0) = MSKBITS(i)
        Alpha(1) = MSKBITS(i + 1)
        Alpha(2) = MSKBITS(i + 2)
    
        DstC(0) = DSTBITS(i)
        DstC(1) = DSTBITS(i + 1)
        DstC(2) = DSTBITS(i + 2)
        '
        tmp(0) = SrcC(0) + (((DstC(0) - SrcC(0)) / 255) * Alpha(0))
        tmp(1) = SrcC(1) + (((DstC(1) - SrcC(1)) / 255) * Alpha(1))
        tmp(2) = SrcC(2) + (((DstC(2) - SrcC(2)) / 255) * Alpha(2))

        'set the alpha values into the temporary bit array
        tmpBits(i) = tmp(0)     'Alpha Blue
        tmpBits(i + 1) = tmp(1) 'Alpha Green
        tmpBits(i + 2) = tmp(2) 'Alpha Red
    Next

    UserControl.Parent.AutoRedraw = True ' Very Important Autoredraw must be True
    'Copia l'Imagine precedente per cancellare le modifiche effetuate con il Testo
    BitBlt UserControl.Parent.hdc, LX, LY, m_Width, m_Height, BB, LX, LY, vbSrcCopy
    'blt the alpha values to the screen
    SetDIBitsToDevice UserControl.Parent.hdc, m_X, m_Y, m_Width, m_Height, 0, 0, 0, m_Height, tmpBits(1), nSRCI, DIB_RGB_COLORS
    'set the Last X and Last Y values so we know where tp clear the screen next time.
    LX = m_X
    LY = m_Y
    UserControl.Parent.Refresh
    UserControl.Parent.AutoRedraw = m_Redraw ' Ripristina il Valore Autoredraw del form
End Sub
'----------------------------------------------------------------------------------------
' Cleanup
'----------------------------------------------------------------------------------------
Private Sub Cleanup()
'cleanup all the memory space we have used
DeleteDC msk
DeleteDC nSRC
DeleteDC DST
DeleteDC BB

DeleteObject BBO
DeleteObject MSKO1
DeleteObject nSRCO1
DeleteObject DSTO1

'erase any array data left over
Erase MSKBITS
Erase SRCBITS
Erase DSTBITS
End Sub

