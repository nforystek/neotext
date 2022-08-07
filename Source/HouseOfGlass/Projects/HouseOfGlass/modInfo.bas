Attribute VB_Name = "modInfo"
#Const modInfo = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Private CenterText As String
Private FadeTime As Long
Private FadeText As String

Public Sub CreateInfo()

End Sub

Public Sub CleanupInfo()

End Sub

Public Sub CenterMessage(ByVal txt As String)
    CenterText = txt
End Sub

Public Sub FadeMessage(ByVal txt As String)
    FadeTime = Timer
    FadeText = txt
    AddMessage txt
End Sub

Private Function Row(ByVal num As Long) As Long
    Row = ((TextHeight \ Screen.TwipsPerPixelY) * num) + (2 * num)
End Function

Public Sub RenderInfo()
    Dim txt As String
            
    DDevice.SetRenderState D3DRS_ZENABLE, False
    DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    DDevice.SetRenderState D3DRS_LIGHTING, False
    DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    DDevice.SetRenderState D3DRS_ALPHATESTENABLE, False

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        
    If Not ConsoleVisible Then
    
        DrawText "E=Forward, D=Back, W/R=Side, Q/A=Speed", 2, 2, &HFF000000

        txt = "ESC=Exit"
        DrawText txt, (frmMain.ScaleWidth / Screen.TwipsPerPixelX) - (frmMain.TextWidth(txt) / Screen.TwipsPerPixelX), 2, &HFF000000
        
    End If
    
    If Not (CenterText = "") Then
    
        DrawText CenterText, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(CenterText) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - (((TextHeight * CountWord(CenterText, vbCrLf)) / Screen.TwipsPerPixelY) / 2)
    
    End If
    
    If (Not (Level.Loaded = "")) And (Not (Level.Elapsed = 0)) Then
    
        txt = "Elapsed Time: " & GetElapsed
        DrawText txt, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(txt) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight - (TextHeight * 2)) / Screen.TwipsPerPixelY)
    
    End If
    
    If Not (FadeText = "") Then
        If (Timer - FadeTime) >= 6 Then
            FadeText = ""
        Else
            DrawText FadeText, ((frmMain.ScaleWidth / Screen.TwipsPerPixelX) / 2) - ((frmMain.TextWidth(FadeText) / Screen.TwipsPerPixelX) / 2), ((frmMain.ScaleHeight / Screen.TwipsPerPixelY) / 2) - (TextHeight / 2)
        End If
    End If

    DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    
    DDevice.SetRenderState D3DRS_ZENABLE, 1
    DDevice.SetRenderState D3DRS_LIGHTING, 1
End Sub

Public Function CountWord(ByVal text As String, ByVal Word As String) As Long
    Dim cnt As Long
    Dim pos As Long
    cnt = 0
    pos = InStr(text, Word)
    Do Until pos = 0
        cnt = cnt + 1
        pos = InStr(pos + 1, text, Word)
    Loop
    CountWord = cnt
End Function

