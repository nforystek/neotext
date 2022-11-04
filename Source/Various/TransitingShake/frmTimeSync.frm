VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTimeSync 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   768
      Top             =   768
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTimeSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This Time Synchronizer is modified from of the original posting found at this url:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=53903&lngWId=1

Private sngDelay As Single
Private strTime As String

Private Sub Form_Load()
    Winsock.RemoteHost = "time.windows.com" '"time.nist.gov"
    Winsock.RemotePort = 37
End Sub

Private Sub WinSock_Close()

    On Error Resume Next
        Do Until Winsock.State = sckClosed
            Winsock.Close
            DoEvents
        Loop
        sngDelay = ((Timer - sngDelay) / 2)
        Winsock.Tag = "FINISH"
    On Error GoTo 0

End Sub

Private Sub WinSock_Connect()

    sngDelay = Timer

End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)

  Dim strData As String

    Winsock.GetData strData, vbString
    strTime = strTime & strData

End Sub

Public Function SynchronizeClock() As Variant

    strTime = Empty
    Winsock.Tag = "AQUIRE"
    Winsock.Connect

    Do While Winsock.Tag = "AQUIRE"
        DoEvents
    Loop
    
    Dim datTime As Date
    Dim dblTime As Double
    Dim lngTime As Long

    strTime = Trim$(strTime)
    If Len(strTime) <> 4 Then Err.Raise 8, App.Title, "Error synchronizing!"
    Debug.Print LenB(strTime)

    dblTime = Asc(Left$(strTime, 1)) * 256 ^ 3 + Asc(Mid$(strTime, 2, 1)) * 256 ^ 2 + Asc(Mid$(strTime, 3, 1)) * 256 ^ 1 + Asc(Right$(strTime, 1))
    lngTime = dblTime - 2840140800#
    datTime = DateAdd("s", CDbl(lngTime + CLng(sngDelay)), #1/1/1990#)
   
    SynchronizeClock = Month(datTime) & "/" & Day(datTime) & "/" & Year(datTime) & " " & Hour(datTime) & ":" & Minute(datTime) & ":" & Second(datTime)

End Function
