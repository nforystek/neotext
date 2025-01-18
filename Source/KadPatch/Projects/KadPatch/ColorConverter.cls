VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------
'(c) Branco Medeiros, 1999
'    branco@apis.com.br
'-----------------------------------------------------------------------

'======================================================================
'  Abstract:
'======================================================================
'  ColorConverter presents methods to easy conversion between the
'  RGB color model and the HSB color model.
'----------------------------------------------------------------------
'  * The RGB color model represents colors as they are used in light
'  emiting objects, such as the beams in a monitor. In this model,
'  each light beam represents the Red, Green or Blue ligths.
'  RGB colors are called "Additive", that is, the frequencies
'  of each light beam are combined to emit the final color.
'  This way, as any beam value is incremented, the resulting color
'  becomes lighter. The white color, for instance, is given by the
'  maximum value of each beam (Red=255, Green=255, Blue=255), and
'  black is given by the absence of light in each beam
'  (Red=0, Green=0, Blue=0).
'  Compare this with the "Subtractive" model (such us the one used
'  in printing and painting and other real life objects that don't
'  emit light): the light that touches a paint spot
'  fires energy in specific frequencies, thus perceived by our
'  eyes as a given color. The practical effect of this is that the
'  color spot "absorbs" all colors but the one that we see.
'  Take, for example, a blue object. We see it blue because it
'  absorbs all light frequencies but the blue frequency (It's
'  funny to think that actually a blue object is any color but blue).
'  In the subtractive model, more quantities of "paint" means darker
'  colors, while less quantities give us a lighter color.
'----------------------------------------------------------------------
'  * The HSB color model is just a mathematical representation of
'  color, in a way more similar to our color perception.
'  For instance, the HSB model breaks the color into three
'  components: The Hue (which would be the "pure" color), the
'  percentage of Saturation ("how much" color) and the Brightness
'  of the color (also a percentage). The Hue comes in the
'  range of 0 to 360. The HSB model is also called HSL (L is for
'  "Lightness") or HSV (V is for "Value" [!!]).
'  This model comes in handy where the RGB model can't help, for
'  instance, classifying similar colors, classifying colors by
'  levels of darkness, sorting colors, etc.
'
'======================================================================
'  Warning:
'======================================================================
'  Because of the differences between the two color models,
'  converting between RGB and HSB values is not a transitive
'  operation, that is, given an RGB value converted to HSB, when
'  converting back to RGB the result may be different from the
'  original. And Vice-versa.
'  This is "by design"... ;-))
'
'-----------------------------------------------------------------------
#Const clMaxRed = MAXRGB
#Const clMaxGreen = MAXRGB
#Const clMaxBlue = MAXRGB
#Const clMaxHue = MAXHUE
#Const clMaxSaturation = MAXHSB
#Const clMaxBrightness = MAXHSB
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
Private Const MAXRGB = 255
Private Const MAXHUE = 359
Private Const MAXHSB = 100
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
Public Enum EColorConstants
  clMaxRed = MAXRGB
  clMaxGreen = MAXRGB
  clMaxBlue = MAXRGB
  clMaxHue = MAXHUE
  clMaxSaturation = MAXHSB
  clMaxBrightness = MAXHSB
End Enum

'======================================================================
Public Function RGBAsString(ByVal RGBValue As Long) As String
'----------------------------------------------------------------------
' Returns the RGB value as string... :-P
'----------------------------------------------------------------------
Dim R As Long, G As Long, B As Long
  Call SplitRGB(RGBValue, R, G, B)
  RGBAsString = CStr(R) & ", " & CStr(G) & ", " & CStr(B)
End Function

'======================================================================
Public Function HSBAsString(ByVal HSBValue As Long) As String
'----------------------------------------------------------------------
' Returns the HSB value as string... :-*
'----------------------------------------------------------------------
Dim H As Long, S As Long, L As Long
  Call SplitHSB(HSBValue, H, S, L)
  HSBAsString = CStr(H) & ", " & CStr(S) & ", " & CStr(L)
End Function

'======================================================================
Public Function RGBToLong(ByVal Red As Long, _
ByVal Green As Long, _
ByVal Blue As Long)
'----------------------------------------------------------------------
'This function is just a wrap around the built in RGB function
'----------------------------------------------------------------------
  Call xLimit(Red, 0, MAXRGB)
  Call xLimit(Green, 0, MAXRGB)
  Call xLimit(Blue, 0, MAXRGB)
  RGBToLong = RGB(Red, Green, Blue)
End Function

'======================================================================
Public Function HSBToLong(ByVal hue As Long, _
ByVal Saturation As Long, _
ByVal brightness As Long) As Long
'----------------------------------------------------------------------
' This function packs the HSB components into a Long, thus
' easing it's handling
'----------------------------------------------------------------------
  Call xLimit(hue, 0, MAXHUE)
  Call xLimit(Saturation, 0, MAXHSB)
  Call xLimit(brightness, 0, MAXHSB)
  HSBToLong = (hue * &H10000) Or (Saturation * 256) Or (brightness)
End Function

'======================================================================
Function Red(ByVal RGBValue As Long) As Long
'----------------------------------------------------------------------
' Returns the Red component of an RGB value
'(In Windows, the Red, Green and Blue values are stored backwards)
'----------------------------------------------------------------------
  Red = RGBValue And 255&
End Function

'======================================================================
Function Green(ByVal RGBValue As Long) As Long
'----------------------------------------------------------------------
' Returns the Green component of an RGB value
'----------------------------------------------------------------------
  Green = (RGBValue And &HFF00&) \ 256
End Function

'======================================================================
Function Blue(ByVal RGBValue As Long) As Long
'----------------------------------------------------------------------
' Returns the Blue component of an RGB value
'----------------------------------------------------------------------
  Blue = (RGBValue And &HFF0000) \ &H10000
End Function

'======================================================================
Function hue(ByVal HSBValue As Long) As Long
'----------------------------------------------------------------------
' Returns the Hue component of an HSB value
'----------------------------------------------------------------------
  hue = xLimit((HSBValue \ &H10000), 0, MAXHUE)
End Function

'======================================================================
Function Saturation(ByVal HSBValue As Long) As Long
'----------------------------------------------------------------------
' Returns the Saturation component of an HSB value
'----------------------------------------------------------------------
  Saturation = xLimit(((HSBValue And &HFF00&) \ 256), 0, MAXHSB)
End Function

'======================================================================
Function brightness(ByVal HSBValue As Long) As Long
'----------------------------------------------------------------------
' Returns the Brightness component of an HSB value
'----------------------------------------------------------------------
  brightness = xLimit((HSBValue And &HFF&), 0, MAXHSB)
End Function

'======================================================================
Sub SplitRGB(ByVal RGBValue As Long, _
ByRef Red As Long, _
ByRef Green As Long, _
ByRef Blue As Long)
'----------------------------------------------------------------------
' Breaks an RGB value into its components
'----------------------------------------------------------------------
  Red = (RGBValue And &HFF&)
  Green = (RGBValue And &HFF00&) \ 256
  Blue = (RGBValue And &HFF0000) \ &H10000
End Sub

'======================================================================
Sub SplitHSB(ByVal HSBValue As Long, _
ByRef hue As Long, _
ByRef Saturation As Long, _
ByRef brightness As Long)
'----------------------------------------------------------------------
' Breaks an HSB value into its components
'----------------------------------------------------------------------
  hue = xLimit((HSBValue \ &H10000), 0, MAXHUE)
  Saturation = xLimit(((HSBValue And &HFF00&) \ 256), 0, MAXHSB)
  brightness = xLimit((HSBValue And &HFF&), 0, MAXHSB)
End Sub

'======================================================================
Function HSBToRGB(ByVal HSBValue As Long) As Long
'----------------------------------------------------------------------
' Converts an HSB value to the RGB color model
'----------------------------------------------------------------------
' Adapted from Java.awt.Color.java
'----------------------------------------------------------------------
Dim R As Long, G As Long, B As Long
Dim H As Long, S As Long, L As Long
Dim nH As Single, nS As Single, nL As Single
Dim nF As Single, nP As Single, nQ As Single, nT As Single
Dim lH As Long

  Call SplitHSB(HSBValue, H, S, L)
  
  If S > 0 Then
  
    nH = H / 60: nL = L / 100: nS = S / 100
    
    lH = Int(nH)
    nF = nH - lH
    nP = nL * (1 - nS)
    nQ = nL * (1 - nS * nF)
    nT = nL * (1 - nS * (1 - nF))
    Select Case lH
    Case 0
      R = nL * 255
      G = nT * 255
      B = nP * 255
    Case 1
      R = nQ * 255
      G = nL * 255
      B = nP * 255
    Case 2
      R = nP * 255
      G = nL * 255
      B = nT * 255
    Case 3
      R = nP * 255
      G = nQ * 255
      B = nL * 255
    Case 4
      R = nT * 255
      G = nP * 255
      B = nL * 255
    Case 5
      R = nL * 255
      G = nP * 255
      B = nQ * 255
    End Select
  Else
    R = (L * 255) / 100
    G = R: B = R
  End If
  
  HSBToRGB = RGBToLong(R, G, B)
  
End Function

'======================================================================
Function RGBToHSB(ByVal RGBValue As Long) As Long
'----------------------------------------------------------------------
' Converts an RGB value to the HSB color model
'----------------------------------------------------------------------
' Adapted from Java.awt.Color.java
'----------------------------------------------------------------------
Dim nTemp As Single
Dim lMin As Long, lMax As Long, lDelta As Long
Dim R As Long, G As Long, B As Long
Dim H As Long, S As Long, L As Long

  Call SplitRGB(RGBValue, R, G, B)
  
  lMax = IIf(R > G, IIf(R > B, R, B), IIf(G > B, G, B))
  lMin = IIf(R < G, IIf(R < B, R, B), IIf(G < B, G, B))
  
  lDelta = lMax - lMin
  
  L = (lMax * 100) / 255
  
  If lMax > 0 Then
    S = (lDelta / lMax) * 100
    If lDelta > 0 Then
      If lMax = R Then
        nTemp = (G - B) / lDelta
      ElseIf lMax = G Then
        nTemp = 2 + (B - R) / lDelta
      Else
        nTemp = 4 + (R - G) / lDelta
      End If
      H = nTemp * 60
      If H < 0 Then H = H + 360
    End If
  End If
  
  RGBToHSB = HSBToLong(H, S, L)
  
End Function

'======================================================================
Private Function xLimit(Value As Long, _
Lower As Long, Higher As Long) As Long
'----------------------------------------------------------------------
' Auxiliary procedure to limit (!) a given value to a specified
' range
'----------------------------------------------------------------------
  If Value < Lower Then Value = Lower
  If Value > Higher Then Value = Higher
  xLimit = Value
End Function



