VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const HSLMAX As Double = 240# '***
    'H, S and L values can be 0 - HSLMAX
    '240 matches what is used by MS Win;
    'any number less than 1 byte is OK;
    'works best if it is evenly divisible by 6
Const RGBMAX As Integer = 255 '***
    'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
    'Hue is undefined if Saturation = 0 (greyscale)

Private mdHue As Double, mdSat As Double, mdLum As Double, _
    miRed As Integer, miGreen As Integer, miBlue As Integer, _
    mlCol As Long, mstrName As String
    
Private Function iMax(a As Integer, B As Integer) _
    As Integer
'Return the Larger of two values
    If a > B Then
        iMax = a
    Else
        iMax = B
    End If
End Function

Private Function iMin(a As Integer, B As Integer) _
    As Integer
'Return the smaller of two values
    If a < B Then
        iMin = a
    Else
        iMin = B
    End If
End Function

'Adapted from C Code in MS KB
Public Sub RGBtoHSL()
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value
Dim H As Long, S As Long, L As Long
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, BDelta As Double
Dim iMinus As Integer, iPlus As Integer
    
    cMax = iMax(iMax(miRed, miGreen), miBlue) 'Highest and lowest
    cMin = iMin(iMin(miRed, miGreen), miBlue) 'color values
    
    iMinus = cMax - cMin 'Used to simplify the
    iPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((iPlus * HSLMAX) + RGBMAX) \ (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        S = 0# 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            S = ((iMinus * HSLMAX) + 0.5) / iPlus
        Else
            S = ((iMinus * HSLMAX) + 0.5) / (2 * RGBMAX - iPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - miRed) * (HSLMAX / 6)) + 0.5) / iMinus
        GDelta = (((cMax - miGreen) * (HSLMAX / 6)) + 0.5) / iMinus
        BDelta = (((cMax - miBlue) * (HSLMAX / 6)) + 0.5) / iMinus
    
        Select Case cMax
            Case miRed
                H = BDelta - GDelta
            Case miGreen
                H = (HSLMAX / 3) + RDelta - BDelta
            Case miBlue
                H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If H < 0 Then H = H + HSLMAX
    End If
    Hue() = CInt(H)
    Saturation() = CInt(S)
    Luminance() = CInt(L)
End Sub

Public Sub HSLtoRGB()
Dim Magic1 As Double, Magic2 As Double

'Part of the HSL to RGB function when it was 'Stand-alone',
'This is now taken care of in Property Let
'
'    If mdHue > HSLMAX Then
'        mdHue = HSLMAX
'    ElseIf mdHue < 0 Then
'        mdHue = 0
'    End If
'
'    If mdLum > HSLMAX Then
'        mdLum = HSLMAX
'    ElseIf mdLum < 0 Then
'        mdLum = 0
'    End If
'
'    If mdSat > HSLMAX Then
'        mdSat = HSLMAX
'    ElseIf mdSat < 0 Then
'        mdSat = 0
'    End If
'
    If CInt(mdSat) = 0 Then 'Greyscale
        miRed = (mdLum * RGBMAX) / HSLMAX 'luminescence,
                'converted to the proper range
        miGreen = miRed 'All RGB values same in greyscale
        miBlue = miRed
        'If CInt(mdHue) <> UNDEFINED Then
            'This is technically an error.
            'The RGBtoHSL routine will always return
            'Hue = UNDEFINED (160) when HSLMAX is 240
            'when Sat = 0.
            'if you are writing a color mixer and
            'letting the user input color values,
            'you may want to set Hue = UNDEFINED
            'in this case.
        'End If
    Else
        'Get the "Magic Numbers"
        If mdLum <= HSLMAX / 2 Then
            Magic2 = (mdLum * (HSLMAX + mdSat) + 0.5) / HSLMAX
        Else
            Magic2 = mdLum + mdSat - ((mdLum * mdSat) + 0.5) / HSLMAX
        End If
        
        Magic1 = 2 * mdLum - Magic2
        
        'get miRed, miGreen, miBlue; change units from HSLMAX range
        'to RGBMAX range
        Red() = (HuetoRGB(Magic1, Magic2, mdHue + (HSLMAX / 3)) _
            * RGBMAX + 0.5) / HSLMAX
        Green() = (HuetoRGB(Magic1, Magic2, mdHue) * RGBMAX + 0.5) / HSLMAX
        Blue() = (HuetoRGB(Magic1, Magic2, mdHue - (HSLMAX / 3)) _
            * RGBMAX + 0.5) / HSLMAX
    End If
    mlCol = RGB(miRed, miGreen, miBlue)
End Sub

'Long -> RGB -> HSL
Public Sub RecalcFromColor()
    RecalcRGB
    RGBtoHSL
End Sub

'RGB -> Long, RGB -> HSL
Public Sub RecalcFromRGB()
    mlCol = RGB(miRed, miGreen, miBlue)
    RGBtoHSL
End Sub

'HSL -> RGB -> Long
Public Sub RecalcFromHSL()
    HSLtoRGB
    mlCol = RGB(miRed, miGreen, miBlue)
End Sub

'Long -> RGB
Public Sub RecalcRGB()
    Red() = mlCol And &HFF
    Green() = ((mlCol And &H100FF00) / &H100)
    Blue() = (mlCol And &HFF0000) / &H10000
End Sub

Private Function HuetoRGB(mag1 As Double, mag2 As Double, _
    ByVal Hue As Double) As Double '***
'Utility function for HSLtoRGB

'Range check
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    
    'Return r, g, or b value from parameters
    Select Case Hue 'Values get progressively larger.
                'Only the first true condition will execute
        Case Is < (HSLMAX / 6)
            HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
            HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
            HuetoRGB = (mag1 + (((mag2 - mag1) * _
                ((HSLMAX * 2 / 3) - Hue) + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
            HuetoRGB = mag1
    End Select
End Function

Private Function ContrastingColor() As Long
'Returns Black or White, whichever will show up better
'on the specified color.
'Useful for setting label forecolors with transparent
'backgrounds (send it the form backcolor - RGB value, not
'system value!)
'(also produces a monochrome negative when applied to
'all pixels in an image)

'Where did the ridiculous formula for the B/W cutoff come from?
'I eyeballed a series of H/S values to see where the ideal
'cutoff would be, made a chart of it in Excel, then made a chart
'based on a mathematical formula and played with it till it
'strongly resembled the sample charts.  If anyone knows a better
'way, please let me know.  In the mean time, this does work
'surprisingly well...

    If mdLum > (127 - Sin((mdHue / 42) - 0.3) * _
        ((mdSat / 4) * ((320 - mdHue) / 140)) * _
        0.7) Then
        ContrastingColor = vbBlack
    Else
        ContrastingColor = vbWhite
    End If
End Function

Public Property Let Color(lColor As Long)
    mlCol = lColor
End Property

Public Property Get Color() As Long
    Color = mlCol
End Property

Public Property Get Red() As Integer
    Red = miRed
End Property

Public Property Let Red(iRed As Integer)
    If iRed > 255 Then
        miRed = 255
    ElseIf iRed < 0 Then
        miRed = 0
    Else
        miRed = iRed
    End If
End Property

Public Property Get Green() As Integer
    Green = miGreen
End Property

Public Property Let Green(iGreen As Integer)
    If iGreen > 255 Then
        miGreen = 255
    ElseIf iGreen < 0 Then
        miGreen = 0
    Else
        miGreen = iGreen
    End If
End Property

Public Property Get Blue() As Integer
    Blue = miBlue
End Property

Public Property Let Blue(iBlue As Integer)
    If iBlue > 255 Then
        miBlue = 255
    ElseIf iBlue < 0 Then
        miBlue = 0
    Else
        miBlue = iBlue
    End If
End Property

Public Property Get Hue() As Integer
    Hue = CInt(mdHue)
End Property

Public Property Let Hue(iHue As Integer)
    If iHue > HSLMAX Then
        mdHue = HSLMAX
    ElseIf iHue < 0 Then
        mdHue = 0#
    Else
        mdHue = CDbl(iHue)
    End If
End Property

Public Property Get Saturation() As Integer
    Saturation = CInt(mdSat)
End Property

Public Property Let Saturation(iSat As Integer)
    If iSat > HSLMAX Then
        mdSat = HSLMAX
    ElseIf iSat < 0 Then
        mdSat = 0#
    Else
        mdSat = CDbl(iSat)
    End If
End Property

Public Property Get Luminance() As Integer
    Luminance = CInt(mdLum)
End Property

Public Property Let Luminance(iLum As Integer)
    If iLum > HSLMAX Then
        mdLum = HSLMAX
    ElseIf iLum < 0 Then
        mdLum = 0#
    Else
        mdLum = CDbl(iLum)
    End If
End Property

'Returns a two digit hex code from a number (w/ leading 0)
Private Function TwoHex(iHex As Integer) As String
    TwoHex = Right$("0" & Hex$(iHex), 2)
End Function

'HTML Color Code output
Public Property Get HTMLString() As String
    HTMLString = TwoHex(miRed) & TwoHex(miGreen) & TwoHex(miBlue)
End Property

'Find a good fore color if main color is back, and vice versa
Public Property Get Contrasting() As Long
    Contrasting = ContrastingColor
End Property

'really nothing to do with color...
Public Property Let Name(strName As String)
    mstrName = strName
End Property

Public Property Get Name() As String
    Name = mstrName
End Property

'VB color hex constant in string form
Public Property Get VBString() As String
    VBString = "&H" & Hex$(mlCol) & IIf(mlCol > vbYellow, "", "&")
End Property

