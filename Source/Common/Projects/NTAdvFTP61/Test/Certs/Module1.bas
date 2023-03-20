Attribute VB_Name = "Module1"
Option Explicit


Private cert As NTAdvFTP61.Certificate


Public Sub Main()
    Set cert = New NTAdvFTP61.Certificate
    
    Dim fil As String
   ' fil = "C:\Temp\Server.cer"
    fil = "C:\Temp\Client.cer"
  '  fil = "C:\Temp\Signing.cer"
    'fil = "C:\Temp\Temp.cer"
   ' fil = "C:\Development\Neotext\Certificates\RootCert.cer"
    

    cert.LoadCertificatefile fil

    If ToString(cert.fields(CertificateSequence)) = ReadFile(fil) Then
        Debug.Print "EQUAL"
    Else
        Debug.Print "NON EQUAL"
    End If

    
    Dim fields As Collection
    Set fields = cert.fields
    Dim num As Integer
    For num = 1 To fields.Count
       ' Debug.Print cert.Names(fields.Key(num)) & " " & cert.HexStream(cert.fields(fields.Key(num)))
        Debug.Print cert.Keys(num) & " " & cert.HexStream(cert.fields(cert.Keys(num)))
        Debug.Print cert.Namely(cert.Keys(num)) & " " & cert.Terms(cert.Keys(num))
        Debug.Print
    Next

    
    cert.ViewCertificate
    
    

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
