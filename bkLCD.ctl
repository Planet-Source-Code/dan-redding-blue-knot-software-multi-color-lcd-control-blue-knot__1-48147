VERSION 5.00
Begin VB.UserControl bkLCD 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   225
   ScaleWidth      =   540
   ToolboxBitmap   =   "bkLCD.ctx":0000
   Begin VB.PictureBox picOut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   9  'Not Mask Pen
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "bkLCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'bkLCD control
'2003 Dan Redding / Blue Knot Software

'Print LCD-like characters to any form/picturebox/etc.
'Original characters ('Regular Font') by Peter Wilson 2002
'New Characters & Bold font by Dan Redding 2003
'(Added +*'×÷!@#$_"";<> and ? ;  Revised - and D)

'Free to use, though credits/mentions and
'links to http://www.blueknot.com are appreciated!

Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal X As Long, ByVal X As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum bkLcdColors
    bkLcd_Red = 0
    bkLcd_Orange = 1
    bkLcd_Yellow = 2
    bkLcd_Lime = 3
    bkLcd_Green = 4
    bkLcd_Teal = 5
    bkLcd_Cyan = 6
    bkLcd_Blue = 7
    bkLcd_Purple = 8
    bkLcd_Violet = 9
    bkLcd_Magenta = 10
    bkLcd_Pink = 11
    bkLcd_White = 12
    bkLcd_Custom = 13
End Enum

Public Enum bkLcdBackgroundColors
    bkLcd_RedBack = 0
    bkLcd_AmberBack = 1
    bkLcd_GreenBack = 2
    bkLcd_TealBack = 3
    bkLcd_BlueBack = 4
    bkLcd_PurpleBack = 5
    bkLcd_BlackBack = 6
End Enum

Private Const m_sKey = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ\_×÷"

'Case protection for Enum in VB IDE
'Private bkLcd_Red, bkLcd_Orange, bkLcd_Yellow, bkLcd_Lime, bkLcd_Green, bkLcd_Teal, bkLcd_Cyan, bkLcd_Blue, bkLcd_Purple, bkLcd_Violet, bkLcd_Magenta, bkLcd_Pink, bkLcd_White, bkLcd_Custom
'Private bkLcd_RedBack, bkLcd_AmberBack, bkLcd_GreenBack, bkLcd_TealBack, bkLcd_BlueBack, bkLcd_PurpleBack, bkLcd_BlackBack, bkLcd_CustomBack

Private m_blnBold As Boolean, m_Hue As Byte, m_blnBright As Boolean, _
    m_Background As bkLcdBackgroundColors, m_Color As bkLcdColors, _
    lColor As Long, lBackColor As Long, lLastX As Long, lLastY As Long
'Default Property Values:

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Use Light/Bold font"
    Bold = m_blnBold
End Property

Public Property Let Bold(ByVal NewBold As Boolean)
    m_blnBold = NewBold
    LoadFont
    PropertyChanged "Bold"
End Property

Public Sub PrintOut(oWindow As Object, X As Long, Y As Long, Message As String, _
    Optional blnEndLine As Boolean = True, Optional lDelay As Long = 0&, _
    Optional blnPad As Boolean = False)
    'oWindow is usually a form or Picturebox, but it can be anything with an
    '.hDC, .Width and a .Refresh method
    
    'Sub could also be modified to use/set .CurrentX,/.CurrentY properties directly...
Dim iLoop As Integer, iPos As Integer, lX As Long, lY As Long, lSrcX As Long, _
    sMess As String, iPad As Integer, lhDC As Long, lWidth As Long
        
    If oWindow Is Nothing Then
        lhDC = UserControl.hDC
        lWidth = UserControl.Width
    Else
        lhDC = oWindow.hDC
        lWidth = oWindow.Width
    End If
    lX = X \ Screen.TwipsPerPixelX
    lY = Y \ Screen.TwipsPerPixelY
    If Message <> "{SHOWALL}" Then
        sMess = UCase$(Message)
        sMess = Replace$(sMess, "[", "(")
        sMess = Replace$(sMess, "]", ")")
        sMess = Replace$(sMess, "{", "(")
        sMess = Replace$(sMess, "}", ")")
        sMess = Replace$(sMess, "~", "-")
        sMess = Replace$(sMess, "`", "'")
        sMess = Replace$(sMess, "|", "I")
    Else
        sMess = m_sKey
    End If
    If blnPad Then
        iPad = Fix(lWidth \ CharWidth())
        If iPad > Len(sMess) Then
            sMess = sMess & Space$(iPad - Len(sMess))
        End If
    End If
    For iLoop = 1 To Len(sMess)
        iPos = InStr(1, m_sKey, Mid$(sMess, iLoop, 1))
        If iPos > 0 Then iPos = iPos - 1
        lSrcX = 12 * iPos
        BitBlt lhDC, lX, lY, 12, 15, _
            picOut.hDC, lSrcX, 0, vbSrcCopy
        lX = lX + 12
        If lDelay > 0 Then
            Sleep lDelay
            If Not oWindow Is Nothing Then
                oWindow.Refresh
            End If
            DoEvents
        End If
    Next iLoop
    If Not oWindow Is Nothing Then
        oWindow.Refresh
    End If
    'Return 'next' values if program is using .CurrentX/.CurrentY
    '(See WriteOnForm sub in demo)
    If blnEndLine Then
        lLastX = 0
        lLastY = Y + 15 * Screen.TwipsPerPixelY
    Else
        lLastX = lX * Screen.TwipsPerPixelX
        lLastY = Y
    End If
End Sub

Public Property Get Hue() As Byte
Attribute Hue.VB_Description = "Actual or Custom Hue set by Color"
    Hue = m_Hue
End Property

Public Property Let Hue(ByVal NewHue As Byte)
    'Hue can be set to any value 0 to 240 (the Color property sets this in
    'increments of 20 which provides a pretty full rainbow)
    If NewHue >= 0 And NewHue <= 240 Then
        m_Hue = NewHue
        If NewHue Mod 20 = 0 Then
            'Standard Color -- Convert to 0-11
            m_Color = (NewHue Mod 240) \ 20
        Else
            m_Color = bkLcd_Custom
        End If
        'Get New Color & Tint the Font
        RecalcColor
        LoadFont
        PropertyChanged "Hue"
    End If
End Property

Public Property Get Bright() As Boolean
Attribute Bright.VB_Description = "High or Low Intensity"
    Bright = m_blnBright
End Property

Public Property Let Bright(ByVal NewBright As Boolean)
    'This is boolean for ease of use.
    'It could also have been made a byte and accpted a range of 0-240
    'which would have made the 'white' color unnecessary as any color
    'with a luminosity of 240 comes out white.
    'See RecalcColor for details of how the boolean affects the Luminosity
    m_blnBright = NewBright
    'Get new color & tint the font
    RecalcColor
    LoadFont
    PropertyChanged "Bright"
End Property

Public Property Get Background() As bkLcdBackgroundColors
Attribute Background.VB_Description = "Background color of LCDs (Use with Caution!)"
    Background = m_Background
End Property

Public Property Let Background(ByVal NewBackground As bkLcdBackgroundColors)
    m_Background = NewBackground
    'Get new BG Color & Tint the font
    RecalcBackground
    LoadFont
    PropertyChanged "Background"
End Property

Public Property Get Color() As bkLcdColors
Attribute Color.VB_Description = "Forecolor Selection"
    Color = m_Color
End Property

Public Property Let Color(ByVal NewColor As bkLcdColors)
    m_Color = NewColor
    If m_Color <> bkLcd_Custom Then
        'Get New Color & Tint the Font
        RecalcColor
        LoadFont
    End If
    PropertyChanged "Color"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    'Defaults
    m_blnBold = False
    m_blnBright = False
    m_Background = bkLcd_BlackBack
    'Defaults to Cyan in memory of original ;-)
    m_Color = bkLcd_Cyan
    RecalcBackground
    RecalcColor
    LoadFont
End Sub

Private Sub UserControl_Paint()
    'Display current colorset on 'control' at design time
    PrintOut Nothing, 0, 0, "LCD"
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_blnBold = PropBag.ReadProperty("Bold", False)
    m_Hue = PropBag.ReadProperty("Hue", 0)
    m_blnBright = PropBag.ReadProperty("Bright", False)
    m_Background = PropBag.ReadProperty("Background", bkLcd_BlackBack)
    m_Color = PropBag.ReadProperty("Color", bkLcd_Cyan)
    'Set Colors & Font
    RecalcColor
    RecalcBackground
    LoadFont
End Sub

Private Sub UserControl_Resize()
    'Control invisible at runtime, why do you need to resize it?
    UserControl.Width = 40 * Screen.TwipsPerPixelX
    UserControl.Height = 19 * Screen.TwipsPerPixelY
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Bold", m_blnBold
    PropBag.WriteProperty "Hue", m_Hue
    PropBag.WriteProperty "Bright", m_blnBright
    PropBag.WriteProperty "Background", m_Background
    PropBag.WriteProperty "Color", m_Color
End Sub

Private Sub RecalcColor()
Dim bLum As Byte, bSat As Byte, C As cColor
    'Hue really doesn't matter for White
    If m_Color < bkLcd_White Then
        m_Hue = m_Color * 20
    End If
    'No color (saturation) for White, full saturation for others
    'Luminosity (intensity) determined by 'Bright' property
    If m_Color = bkLcd_White Then
        bLum = IIf(m_blnBright, 240, 160)
        bSat = 0
    Else
        bLum = IIf(m_blnBright, 180, 100)
        bSat = 240
    End If
    'Use Color class to convert HSL values to Long Color
    Set C = New cColor
    With C
        .Hue = m_Hue
        .Luminance = bLum
        .Saturation = bSat
        .RecalcFromHSL
        lColor = .Color
    End With
    Set C = Nothing
End Sub

Private Sub RecalcBackground()
Dim bLum As Byte, bSat As Byte, C As cColor, bHue As Byte
    'Hue doesn't matter for Black
    If m_Background < bkLcd_BlackBack Then
        bHue = m_Background * 40
    End If
    'Set Appropriate Luminosity & Saturation levels for Non-black backgrounds (tweak to taste!)
    If m_Background = bkLcd_BlackBack Then
        bLum = 0
        bSat = 0
    Else
        bLum = 30
        bSat = 240
    End If
    'Use Color class to convert HSL values to Long color value
    Set C = New cColor
    With C
        .Hue = bHue
        .Luminance = bLum
        .Saturation = bSat
        .RecalcFromHSL
        lBackColor = .Color
    End With
    Set C = Nothing
End Sub

Private Sub LoadFont()
    'Get either the regular (LO) or bold (HI) font
    'Font is a JPG image stored in Custom resource.
    Set picOut.Picture = LoadResPic(IIf(m_blnBold, "HI", "LO"), 101)
    '
    Tint picOut, lColor, lBackColor
    If Ambient.UserMode = False Then UserControl.Refresh
End Sub

Private Sub Tint(pic As PictureBox, lCol As Long, Optional lCol2 As Long = 0&)
    With pic
        '.FillStyle = vbSolid
        .DrawMode = 9
        pic.Line (0, 0)-(.Width, .Height), lCol, BF
    
        If lCol2 > 0 Then
            .FillStyle = vbSolid
            .DrawMode = 15 'vbmerge tints background/black!
            pic.Line (0, 0)-(.Width, .Height), lCol2, BF
        End If
    End With
End Sub
Public Property Get BackgroundColor() As OLE_COLOR
Attribute BackgroundColor.VB_MemberFlags = "400"
    BackgroundColor = lBackColor
End Property

Public Property Get LastX() As Long
Attribute LastX.VB_MemberFlags = "400"
    LastX = lLastX
End Property

Public Property Get LastY() As Variant
Attribute LastY.VB_MemberFlags = "400"
    LastY = lLastY
End Property

Public Property Get CharWidth() As Long
Attribute CharWidth.VB_MemberFlags = "400"
    'constant for now, might be calculated if different fonts
    CharWidth = 12 * Screen.TwipsPerPixelX
End Property

Public Property Get CharHeight() As Long
Attribute CharHeight.VB_MemberFlags = "400"
    'constant for now, might be calculated if different fonts
    CharHeight = 15 * Screen.TwipsPerPixelY
End Property


