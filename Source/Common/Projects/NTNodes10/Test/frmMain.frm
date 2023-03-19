VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyper Node Example"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListFirst 
         Caption         =   "First"
      End
      Begin VB.Menu mnuListPoint 
         Caption         =   "Point"
      End
      Begin VB.Menu mnuListFinal 
         Caption         =   "Final"
      End
      Begin VB.Menu mnuForth 
         Caption         =   "Forth"
      End
      Begin VB.Menu mnuTotal 
         Caption         =   "Total"
      End
      Begin VB.Menu mnuListReway 
         Caption         =   "Reway"
      End
   End
   Begin VB.Menu mnuNode 
      Caption         =   "Node"
      Visible         =   0   'False
      Begin VB.Menu mnuNodeAddr 
         Caption         =   "Addr"
         Index           =   0
         Begin VB.Menu mnuNodeValue 
            Caption         =   "value"
         End
         Begin VB.Menu mnuNodeForth 
            Caption         =   "Forth"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Const BlockSize As Long = 4
Private Const SpaceSize As Long = 3


Public Sub DrawNodes(ByRef n As inode)
    Dim nTop As Long
    nTop = (TextHeight("X") * 8)
    
    If Not n.Count = 0 Then
        
        Dim ColPerRow As Long

        Dim nLeft As Long
        Dim nHeight As Long
        Dim nWidth As Long

        nLeft = SpaceSize
        
        nWidth = ((frmMain.ScaleWidth - (SpaceSize * 2)) \ BlockSize)
    
        Dim att As Long
        Dim tot As Long
        Dim ptr As Long
        Dim eol As Long
        Dim bol As Long
        Dim cnt As Long
        Dim chk As Long
        chk = n.Check
        eol = n.Final
        bol = n.First
        ptr = n.Point
        cnt = n.Count

        Do Until (n.Point = bol) Or (bol = 0)
            n.Forward
        Loop
        
        Do
            If nLeft > (frmMain.ScaleWidth - (SpaceSize * 2)) Then
                nLeft = SpaceSize
                nTop = nTop + BlockSize + SpaceSize
            End If

            If ptr = n.Point Then
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), vbGreen, BF
            ElseIf chk = n.Point Then
                frmMain.Line (nLeft, nTop)-(nLeft + BlockSize, nTop + BlockSize), RGB(255, 128, 0), BF
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

            
            n.Forward
            cnt = cnt - 1

        Loop Until n.Point = bol Or (bol = 0)
        
        nTop = nTop + BlockSize + SpaceSize
        frmMain.Height = nTop + (frmMain.Height - frmMain.ScaleHeight)
        
        Do Until n.Point = ptr Or (bol = 0)
            n.Forward
        Loop
    Else
        frmMain.Height = nTop + (frmMain.Height - frmMain.ScaleHeight)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 1 Then Me.Visible = False
End Sub

