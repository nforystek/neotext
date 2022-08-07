VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Choose Application"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
'@                                                                                       @'
'@ Ok everyone, after looking at all the transparency examples on PSCode, I decided to   @'
'@ write my own because all the others contain absurd amounts of code and could all be   @'
'@ made much smaller. This example uses a Function created by Robert Gainor that was     @'
'@ cleaned up a bit. Robert Gainor made it look like you needed 4 forms and 4 modules    @'
'@ just to create a transparent form when all you need is one form and one line of code. @'
'@                                                                                       @'
'@      Created by Chad Roe, Jeepman6@aol.com, http://www.offroadextremes.com/chad       @'
'@                                                                                       @'
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'

'FormMove and FormOnTop Declarations and Constants
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

'Transparency Declarations and Constants
'I copied these from Robert Gainor's Example
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5

'Transparency Function
'I copied this from Robert Gainor's Example
Private Function MakeTransparent(ByRef Frm As Form, ByVal TransparentColor As Long) As Long
Dim rgnMain As Long, rgnPixel As Long, bmpMain As Long, dcMain As Long
Dim Width As Long, Height As Long, x As Long, y As Long
Dim ScaleSize As Long, RGBColor As Long
ScaleSize& = Frm.ScaleMode
Frm.ScaleMode = 3
Frm.BorderStyle = 0
Width& = Frm.ScaleX(Frm.Picture.Width, vbHimetric, vbPixels)
Height& = Frm.ScaleY(Frm.Picture.Height, vbHimetric, vbPixels)
Frm.Width = Width& * Screen.TwipsPerPixelX
Frm.Height = Height& * Screen.TwipsPerPixelY
rgnMain& = CreateRectRgn(0&, 0&, Width&, Height&)
dcMain& = CreateCompatibleDC(Frm.hDc)
bmpMain& = SelectObject(dcMain&, Frm.Picture.Handle)
For y& = 0& To Height&
  For x& = 0& To Width&
    RGBColor& = GetPixel(dcMain&, x&, y&)
    If RGBColor& = TransparentColor& Then
      rgnPixel& = CreateRectRgn(x&, y&, x& + 1&, y& + 1&)
      CombineRgn rgnMain&, rgnMain&, rgnPixel&, RGN_XOR
      DeleteObject rgnPixel&
    Else
'      rgnPixel& = CreateRectRgn(x&, y&, x& + 1&, y& + 1&)
'      CombineRgn rgnPixel, rgnMain&, rgnPixel&, RGN_DIFF
'      CombineRgn rgnMain&, rgnMain&, rgnPixel&, RGN_AND
'      DeleteObject rgnPixel&
    End If
  Next x&
Next y&
SelectObject dcMain&, bmpMain&
DeleteDC dcMain&
DeleteObject bmpMain&
If rgnMain& <> 0& Then
  SetWindowRgn Frm.hWnd, rgnMain&, True
  MakeTransparent = rgnMain&
End If
Frm.ScaleMode = ScaleSize&
End Function


'FormMove and FormOnTop Subs
Private Sub FormOnTop(Frm As Form)
Call SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Private Sub FormMoveXP(Frm As Form)
Call ReleaseCapture
Call SendMessage(Frm.hWnd, &HA1, 2, 0&)
End Sub

Private Sub CenterForm(Frm As Form)
Frm.Left = Screen.Width / 2 - Frm.Width / 2
Frm.Top = Screen.Height / 2 - Frm.Height / 2
End Sub

Private Sub Form_Load()

 Call FormOnTop(Me)
Call CenterForm(Me)
Call MakeTransparent(Me, RGB(255, 255, 255))
End Sub
