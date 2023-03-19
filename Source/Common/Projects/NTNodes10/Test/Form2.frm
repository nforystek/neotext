VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BlockSize As Long = 4
Private Const SpaceSize As Long = 3


Public Sub DrawNodes(ByRef n As iNode)
    Dim nTop As Long
    nTop = (TextHeight("X") * 8)
    If Not n.Count = 0 Then
        
        Dim ColPerRow As Long

        Dim nLeft As Long
        Dim nHeight As Long
        Dim nWidth As Long

        
        nLeft = SpaceSize
        
        nWidth = (frmMain.ScaleWidth - (SpaceSize * 2)) \ BlockSize
    
        Dim ptr As Long
        Dim eol As Long
        Dim bol As Long
        Dim cnt As Long
        Dim chk As Long
'        Dim ff As Long
'        Dim pp As Long
        chk = n.check
        eol = n.Final
        bol = n.First
        ptr = n.Point
        cnt = n.Count
'        ff = n.Ahead
'        pp = n.Backs
        
        Do Until n.Point = bol
            n.forward
        Loop
        
        Do
            If nLeft > (frmMain.ScaleWidth - (SpaceSize * 2)) Then
                nLeft = SpaceSize
                nTop = nTop + BlockSize + SpaceSize
            End If
    
            If ptr = n.Point Then
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), vbGreen, BF
            ElseIf chk = n.Point Then
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), vbYellow, BF
'            ElseIf ff = n.Point Then
'                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), ColorConstants.vbMagenta, BF
'            ElseIf pp = n.Point Then
'                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), ColorConstants.vbCyan, BF
            Else
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), n.Value, BF
            End If
            If bol = n.Point Then
                frmMain.Line (nLeft - 1, nTop - 1)-(nLeft + BlockSize + 1, nTop + BlockSize + 1), vbBlue, B
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), vbBlue, B
            ElseIf eol = n.Point Then
                frmMain.Line (nLeft - 1, nTop - 1)-(nLeft + BlockSize + 1, nTop + BlockSize + 1), vbRed, B
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), vbRed, B
            End If
            nLeft = nLeft + BlockSize + SpaceSize
    
            n.forward
            cnt = cnt - 1
            If cnt = 0 Then nTop = nTop + BlockSize + SpaceSize
        Loop Until n.Point = bol
        
        frmMain.Height = nTop + (frmMain.Height - frmMain.ScaleHeight) + BlockSize + SpaceSize
        
        Do Until n.Point = ptr
            n.forward
        Loop
    Else
        frmMain.Height = nTop + (frmMain.Height - frmMain.ScaleHeight)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 1 Then Me.Visible = False
End Sub


