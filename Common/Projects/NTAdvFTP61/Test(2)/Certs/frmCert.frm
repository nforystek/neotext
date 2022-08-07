VERSION 5.00
Begin VB.Form frmCert 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cert As New NTAdvFTP61.Certificate


Private Sub Form_Load()

'    Dim txt As String
'    txt = ReadFile("C:\Temp\Temp.txt")
'    Dim nxt As String
'    Dim out As String
'
'    Do Until txt = ""
'        nxt = RemoveNextArg(txt, " ")
'        If nxt <> "" Then
'            out = out & Chr("&H" & nxt)
'        End If
'    Loop
'    WriteFile "C:\Temp\Temp.cer", out
'    End
    
   ' cert.LoadCertificatefile "C:\Temp\Server.cer"
   ' cert.LoadCertificatefile "C:\Temp\Client.cer"
   ' cert.LoadCertificatefile "C:\Temp\Signing.cer"
    cert.LoadCertificatefile "C:\Temp\Temp.cer"
    
    Dim fields As NTNodes10.Collection
    Set fields = cert.fields
    Dim num As Integer
    For num = 1 To fields.count
       ' Debug.Print cert.Names(fields.Key(num)) & " " & cert.HexStream(cert.fields(fields.Key(num)))
        Debug.Print fields.Key(num) & " " & cert.HexStream(cert.fields(fields.Key(num)))
        Debug.Print cert.Namely(fields.Key(num)) & " " & cert.Terms(fields.Key(num))
        Debug.Print
    Next
    
    

'    Debug.Print "Version " & cert.HexStream(cert.fields(Version))
'    Debug.Print
'    Debug.Print "SerialNumber " & cert.HexStream(cert.fields(SerialNumber))
'    Debug.Print
'    Debug.Print "Algorithm " & cert.HexStream(cert.fields(Algorithm))
'    Debug.Print
'    Debug.Print "Issuer " & cert.HexStream(cert.fields(Issuer))
'    Debug.Print
'    Debug.Print "Validity " & cert.HexStream(cert.fields(Validity))
'    Debug.Print
'    Debug.Print "Subject " & cert.HexStream(cert.fields(Subject))
'    Debug.Print
'    Debug.Print "PublicKey " & cert.HexStream(cert.fields(PublicKey))
'    Debug.Print
'    Debug.Print "Extensions " & cert.HexStream(cert.fields(Extensions))
'    Debug.Print
'    Debug.Print "SignatureAlgorithm " & cert.HexStream(cert.fields(SignatureAlgorithm))
'    Debug.Print
'    Debug.Print "Signature " & cert.HexStream(cert.fields(Signature))
'    Debug.Print
                                            
End Sub
