Attribute VB_Name = "modCert"
Option Explicit

Global KeyContainer As String

Global CertObjs As NTNodes10.Collection

Public Sub ViewCertificate(ByRef cert As Certificate)

    If CertObjs Is Nothing Then
        Set CertObjs = New NTNodes10.Collection
    End If
    
    Dim CertForm As New frmCert
    CertForm.ViewCert cert
    Unload CertForm
    Set CertForm = Nothing
End Sub

Public Function CheckCertificate(ByRef cert As Certificate) As Boolean
    If CertObjs Is Nothing Then
        Set CertObjs = New NTNodes10.Collection
    End If
    Dim sn As String
    
    sn = "SN_" & Replace(cert.Terms(SerialNumber), " ", "")
    
    If CertObjs.Exists(sn) Then
        Set cert = CertObjs(sn)
        CertObjs.Remove sn
    End If
    
    If (Not cert.NoPrompt) Then 'And (LCase(AppEXE(True, True)) <> "maxservice") Then
        Dim CertForm As New frmCert
        CertForm.CheckCert cert
        Unload CertForm
        Set CertForm = Nothing
    Else 'If (LCase(AppEXE(True, True)) = "maxservice") Then
        cert.Accepted = True
    End If

    CheckCertificate = cert.Accepted
    
    If Not CertObjs Is Nothing Then
        CertObjs.Add cert, sn
    End If
End Function

Public Sub TermCerts()
    If Not CertObjs Is Nothing Then
        CertObjs.Clear
        Set CertObjs = Nothing
    End If
End Sub
