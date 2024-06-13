VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Certificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CERT_SEQUENCE As Byte = &H30
Private Const CERT_SETTINGS As Byte = &H31

Private Const CERT_LEN_BYTE As Byte = &H0
Private Const CERT_LEN_LONG As Byte = &H82
Private Const CERT_LEN_WORD As Byte = &H2

Private Const CERT_MAX_WORD = &H100
Private Const CERT_HEX_UNIT = &H2

Private Const CERT_TYPE_NULL As Byte = &H5
Private Const CERT_TYPE_OIDS As Byte = &H6

Private Const CERT_TYPE_BOOLEAN As Byte = &H1
Private Const CERT_TYPE_BITSET As Byte = &H3
Private Const CERT_TYPE_OCTETS As Byte = &H4
Private Const CERT_TYPE_STRING As Byte = &H13
Private Const CERT_TYPE_UTCTIME As Byte = &H17

Private Const CERT_CONTEXT_CERTFORMAT As Byte = &HA0
Private Const CERT_CONTEXT_EXTENSIONS As Byte = &HA3
Private Const CERT_CONTEXT_EMBEDEDKEY As Byte = &H80

Public Enum CertificateFields
    CertificateCollection = 0
    
    CertificateSequence = 2
    CertificateInfoSequence = 3
    Version = 4
    SerialNumber = 5
          
    Issuer = 20: Issuer1 = 21: Issuer2 = 22: Issuer3 = 23: Issuer4 = 24
    Issuer5 = 25: Issuer6 = 26: Issuer7 = 27: Issuer8 = 28: Issuer9 = 29

    IssuerCountry = 21
    IssuerOrganizationalUnit = 22

    Subject = 30: Subject1 = 31: Subject2 = 32: Subject3 = 33: Subject4 = 34
    Subject5 = 35: Subject6 = 36: Subject7 = 37: Subject8 = 38: Subject9 = 39

    SubjectCountry = 31
    SubjectCommonName = 32

    Validity = 6
    ValidityBeginDate = 7
    ValidityExpireDate = 8
    
    PublicKeyBlock = 9
    
    PublicKey = 1

    IntendedUses = 10: IntendedUses1 = 11: IntendedUses2 = 12: IntendedUses3 = 13: IntendedUses4 = 14
    IntendedUses5 = 15: IntendedUses6 = 16: IntendedUses7 = 17: IntendedUses8 = 18: IntendedUses9 = 19

    Extensions = 50: Extensions1 = 51: Extensions2 = 52: Extensions3 = 53: Extensions4 = 54
    Extensions5 = 55: Extensions6 = 56: Extensions7 = 57: Extensions8 = 58: Extensions9 = 59

    ExtensionKeyUsage = 51
    ExtensionExtendedKeyUsage = 52
    ExtensionAuthorityKeyIdentifier = 53

    Algorithm = 70: Algorithm1 = 71: Algorithm2 = 72: Algorithm3 = 73: Algorithm4 = 74
    Algorithm5 = 75: Algorithm6 = 76: Algorithm7 = 77: Algorithm8 = 78: Algorithm9 = 79

    SignatureAlgorithm = 80: SignatureAlgorithm1 = 81: SignatureAlgorithm2 = 82: SignatureAlgorithm3 = 83: SignatureAlgorithm4 = 84
    SignatureAlgorithm5 = 85: SignatureAlgorithm6 = 86: SignatureAlgorithm7 = 87: SignatureAlgorithm8 = 88: SignatureAlgorithm9 = 89

    Signature = 90
    
    ObjectIdentifyer = 100
    
End Enum

Private pNoPrompt As Boolean
Private pAccepted As Boolean

Private pContextBit As Integer
Private pKeySize As Integer

Private pKeys As Collection
Private pFields As Collection

Public Property Get RSAKeySize() As Integer ' _
Returns the PublicKey length of the RSA cipher.
    RSAKeySize = pKeySize
End Property
    
Public Sub ViewCertificate() ' _
Opens a window for the current loaded certificate for viewing.
Attribute ViewCertificate.VB_Description = "Opens a window for the current loaded certificate for viewing."

    modCrypto.ViewCertificate Me
End Sub

Public Function CheckCertificate() As Boolean ' _
Opens a window for the current loaded certificate prompting for acceptance or rejection response.
Attribute CheckCertificate.VB_Description = "Opens a window for the current loaded certificate prompting for acceptance or rejection response."

    CheckCertificate = modCrypto.CheckCertificate(Me)
End Function

Friend Property Get NoPrompt() As Boolean
    NoPrompt = pNoPrompt
End Property

Friend Property Get Accepted() As Boolean
    Accepted = pAccepted
End Property

Friend Property Let NoPrompt(ByVal RHS As Boolean)
    pNoPrompt = RHS
End Property

Friend Property Let Accepted(ByVal RHS As Boolean)
    pAccepted = RHS
End Property


Public Sub LoadCertificateFile(ByVal FileName As String) ' _
Loads a certificate in binary form by the FileName supplied, if it is a .CER file.
    If PathExists(FileName, True) Then
        LoadCertificate Stream(ReadFile(FileName))
    Else
        Err.Raise 53
    End If
End Sub

Public Function HexStream(Stream) As String ' _
Inputs a stream of byte array or Stream object and outputs the length followed by hexidecimal form per byte seperated by spaces for debugging.
Attribute HexStream.VB_Description = "Inputs a stream of byte array or Stream object and outputs the length followed by hexidecimal form per byte seperated by spaces for debugging."

    HexStream = modStream.HexStream(Stream)
End Function
Private Sub AddLength(ByRef out As Stream)
    Dim inc As Integer
    inc = out.Length
    If inc > CERT_MAX_WORD - (CERT_LEN_WORD * CERT_HEX_UNIT) Then
        out.Prepend ToBytes(Chr(CERT_LEN_LONG) & _
                            Chr(Val("&H" & Left(Padding(CERT_HEX_UNIT * CERT_LEN_WORD, Hex(inc), "0"), 2))) & _
                            Chr(Val("&H" & Right(Padding(CERT_HEX_UNIT * CERT_LEN_WORD, Hex(inc), "0"), 2))))
    Else
        out.Prepend ToBytes(Chr(out.Length))
    End If
    
End Sub

Private Function StripNumbers(ByVal Text As String) As String
    Text = Replace(Text, "0", "")
    Text = Replace(Text, "1", "")
    Text = Replace(Text, "2", "")
    Text = Replace(Text, "3", "")
    Text = Replace(Text, "4", "")
    Text = Replace(Text, "5", "")
    Text = Replace(Text, "6", "")
    Text = Replace(Text, "7", "")
    Text = Replace(Text, "8", "")
    Text = Replace(Text, "9", "")
End Function
Public Property Get Namely(ByVal Refer As Variant) As Variant ' _
Converts the identifier number or key to it's respective text literal, or text literal to identifier key, depending on which is passed as Refer.
Attribute Namely.VB_Description = "Converts the identifier number or key to it's respective text literal, or text literal to identifier key, depending on which is passed as Refer."

    If Not IsNumeric(Refer) Then
        If IsNumeric(Replace(Refer, "ID_", "")) Then
            Refer = Replace(Refer, "ID_", "")
        End If
    End If
    
    Dim ext As Integer
    If Not IsNumeric(Refer) Then
        If IsNumeric(Right(Refer, 1)) Then ext = CInt(Right(Refer, 1))
        Select Case LCase(StripNumbers(Refer))
            Case LCase("CertificateCollection")
                Namely = CInt(CertificateCollection)
            Case LCase("PublicKey")
                Namely = CInt(PublicKey)
            Case LCase("CertificateSequence")
                Namely = CInt(CertificateSequence)
            Case LCase("CertificateInfoSequence")
                Namely = CInt(CertificateInfoSequence)
            Case LCase("Version")
                Namely = CInt(Version)
            Case LCase("SerialNumber")
                Namely = CInt(SerialNumber)
            Case LCase("Validity")
                Namely = CInt(Validity)
            Case LCase("ValidityBeginDate")
                Namely = CInt(ValidityBeginDate)
            Case LCase("ValidityExpireDate")
                Namely = CInt(ValidityExpireDate)
            Case LCase("PublicKeyBlock")
                Namely = CInt(PublicKeyBlock)
            Case LCase("IntendedUses")
                Namely = CInt(IntendedUses)
            Case LCase("Issuer")
                Namely = CInt(Issuer)
            Case LCase("Subject")
                Namely = CInt(Subject)
            Case LCase("Extensions")
                Namely = CInt(Extensions)
            Case LCase("Algorithm")
                Namely = CInt(Algorithm)
            Case LCase("SignatureAlgorithm")
                Namely = CInt(SignatureAlgorithm)
            Case LCase("Signature")
                Namely = CInt(Signature)
            Case LCase("ObjectIdentifyer")
                Namely = CInt(ObjectIdentifyer)
            Case Else
                Err.Raise 8, , "Field identifier unrecognized."
        End Select
        Namely = CInt(Namely + ext)
    Else
        If CInt(Refer) > 10 Then
            ext = CInt(Refer) Mod 10
            Refer = ((CInt(Refer) \ 10) * 10)
        End If
        Select Case CInt(Refer)
            Case CInt(CertificateCollection)
                Namely = "CertificateCollection"
            Case CInt(PublicKey)
                Namely = "PublicKey"
            Case CInt(CertificateSequence)
                Namely = "CertificateSequence"
            Case CInt(CertificateInfoSequence)
                Namely = "CertificateInfoSequence"
            Case CInt(Version)
                Namely = "Version"
            Case CInt(SerialNumber)
                Namely = "SerialNumber"
            Case CInt(Validity)
                Namely = "Validity"
            Case CInt(ValidityBeginDate)
                Namely = "ValidityBeginDate"
            Case CInt(ValidityExpireDate)
                Namely = "ValidityExpireDate"
            Case CInt(PublicKeyBlock)
                Namely = "PublicKeyBlock"
            Case CInt(IntendedUses)
                Namely = "IntendedUses"
            Case CInt(Issuer)
                Namely = "Issuer"
            Case CInt(Subject)
                Namely = "Subject"
            Case CInt(Extensions)
                Namely = "Extensions"
            Case CInt(Algorithm)
                Namely = "Algorithm"
            Case CInt(SignatureAlgorithm)
                Namely = "SignatureAlgorithm"
            Case CInt(Signature)
                Namely = "Signature"
            Case CInt(ObjectIdentifyer)
                Namely = "ObjectIdentifyer"
            Case Else
                Err.Raise 8, , "Field identifier unrecognized."
        End Select
        If ext <> 0 Then Namely = Namely & ext
    End If

End Property

Public Property Get Terms(ByVal Refer As Variant) As String ' _
Returns the specified Refer property of the certificate fields as humanly understandable as possible, i.e. English terminology.
Attribute Terms.VB_Description = "Returns the specified Refer property of the certificate fields as humanly understandable as possible, i.e. English."

    Dim tmp As Stream

    Dim name As String
    
    If Not IsNumeric(Refer) Then
        If IsNumeric(Replace(Refer, "ID_", "")) Then
            name = Replace(Refer, "ID_", "")
        Else
            name = Namely(Refer)
            Refer = Namely(Refer)
        End If
    Else
        name = Refer
        Refer = "ID_" & Refer
    End If
    
    Dim inc As Integer
    
    Select Case CInt(name)
        Case CertificateCollection

        Case Version

            Terms = pFields(Refer)

        Case CertificateSequence, CertificateInfoSequence
        
            Set tmp = pFields(Refer)
            Terms = tmp.Length
            Set tmp = Nothing
        Case Is >= ObjectIdentifyer
            inc = -2
        Case Signature, Algorithm, IntendedUses, PublicKeyBlock, SerialNumber, Extensions, SignatureAlgorithm, Subject, Issuer
            inc = -1
        Case Is > SignatureAlgorithm
            inc = -2
        Case Is > Algorithm
            inc = -2
        Case Is > Extensions
            inc = -2
        Case Is > Subject
            inc = -2
        Case Is > Issuer
            inc = -2
        Case Is > IntendedUses
            inc = -2

        Case ValidityBeginDate, ValidityExpireDate, Validity

            Set tmp = pFields(Refer)

            Terms = (CInt(Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(0, 1)(1)), "0"))) & Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(1, 1)(1)), "0")))) + 2000) & "-" & _
                            Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(2, 1)(1)), "0"))) & Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(3, 1)(1)), "0"))) & "-" & _
                            Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(4, 1)(1)), "0"))) & Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(5, 1)(1)), "0"))) & " " & _
                            Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(6, 1)(1)), "0"))) & Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(7, 1)(1)), "0"))) & ":" & _
                            Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(8, 1)(1)), "0"))) & Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(9, 1)(1)), "0"))) & ":" & _
                            Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(10, 1)(1)), "0"))) & Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(11, 1)(1)), "0"))) & _
                            Chr(Val("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(12, 1)(1)), "0"))) & " GMT"
            Set tmp = Nothing

        Case PublicKey

    End Select
    
    If inc = -1 Then
        Terms = RemoveArg(HexStream(pFields(Refer)), ":")
        
    ElseIf inc = -2 Then
        If CInt(name) >= ObjectIdentifyer Then
            inc = -1
            Refer = "ID_" & (CInt(name) - CInt(ObjectIdentifyer))
        Else
            inc = 0
        End If

        Set tmp = Stream(pFields(Refer))
        
        If tmp.Partial(0, 1)(1) = CERT_SEQUENCE Or tmp.Partial(0, 1)(1) = CERT_SETTINGS Then
            inc = inc - 2
        Else
            tmp.Push 1
        End If

        Set tmp = GetSequence(tmp)
        If inc < -1 Then
            GetInteger tmp
        End If
        If inc = -1 Or inc = -3 Then
            Terms = GetInteger(tmp)
        Else
            GetInteger tmp
            Set tmp = GetSequence(tmp)

            Do Until tmp.Length = 0
                Select Case LCase(Chr("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(0, 1)(1)), "0")))
                    Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        Terms = Terms & Chr("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(0, 1)(1)), "0"))
                    Case "~", "`", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "=", "+", "[", "{", "]", "}", "\", "|", ";", ":", "'", """", ",", "<", ".", ">", "/", "?", " "
                        Terms = Terms & Chr("&H" & Padding(CERT_HEX_UNIT, Hex(tmp.Partial(0, 1)(1)), "0"))
                End Select
                
                tmp.Push 1
            Loop

        End If

        Set tmp = Nothing
    End If
    
End Property


Public Property Get Fields(Optional ByVal Refer As Variant = CertificateCollection) ' _
Returns the specified Refer property of the certificate fields as original, and where isolated applicable, binary sectioned form.  No Refer returns the collection object of all fields.
Attribute Fields.VB_Description = "Returns the specified Refer property of the certificate fields as original, and where isolated applicable, binary sectioned form.  No Refer returns the collection object of all fields."

    If Not IsNumeric(Refer) Then
        If Not IsNumeric(Replace(Refer, "ID_", "")) Then
            Refer = Namely(Refer)
        Else
            Refer = Replace(Refer, "ID_", "")
        End If
    End If
                
    If Refer = CertificateCollection Then
        Set Fields = pFields
    Else
        Dim itm As Stream
        Dim out As Stream
        
        Dim inc As Integer
        Select Case Refer
            Case Validity, ValidityBeginDate, ValidityExpireDate
                Set out = New Stream
                If Refer = Validity Or Refer = ValidityBeginDate Then
                    If Exists("ID_" & ValidityBeginDate) Then
                        Set itm = pFields("ID_" & ValidityBeginDate)
                        If Refer = Validity Then
                            AddLength itm
                            itm.Prepend ToBytes(Chr(CERT_TYPE_UTCTIME))
                            out.Concat itm.Partial
                        Else
                            Fields = itm.Partial
                        End If
                    End If
                End If
                If Refer = Validity Or Refer = ValidityExpireDate Then
                    If Exists("ID_" & ValidityExpireDate) Then
                        Set itm = pFields("ID_" & ValidityExpireDate)
                        If Refer = Validity Then
                            AddLength itm
                            itm.Prepend ToBytes(Chr(CERT_TYPE_UTCTIME))
                            out.Concat itm.Partial
                        Else
                            Fields = itm.Partial
                        End If
                    End If
                End If

                AddLength out
                out.Prepend ToBytes(Chr(CERT_SEQUENCE))
                Fields = out.Partial
                
            Case PublicKey
                Set out = Stream(Fields(IntendedUses))
                out.Concat ToBytes(Chr(CERT_TYPE_BITSET))
                Set itm = pFields("ID_" & PublicKeyBlock)
                AddLength itm
                out.Concat itm.Partial
                
                AddLength out
                out.Prepend ToBytes(Chr(CERT_SEQUENCE))
                Fields = out.Partial
            
            Case Extensions, Subject, Issuer, SerialNumber, Signature, Algorithm, SignatureAlgorithm
            
                Set out = New Stream
                out.Concat ToBytes(pFields("ID_" & Refer))
                AddLength out
                out.Prepend ToBytes(IIf(Refer = Extensions, Chr(CERT_CONTEXT_EXTENSIONS), Chr(CERT_SEQUENCE)))
                Fields = out.Partial
                
            Case Is > SignatureAlgorithm
                inc = CERT_TYPE_OIDS
            Case Is > Algorithm
                inc = CERT_TYPE_OIDS
            Case Is > Extensions
                inc = CERT_SETTINGS
            Case Is > Subject
                inc = CERT_SETTINGS
            Case Is > Issuer
                inc = CERT_SETTINGS
            Case Is > IntendedUses
                inc = CERT_TYPE_OIDS
                
            Case CertificateSequence, CertificateInfoSequence
                Set out = New Stream
                out.Clone pFields("ID_" & Refer)
                AddLength out
                out.Prepend ToBytes(Chr(CERT_SEQUENCE))
                Fields = out.Partial
                
            Case Version
                Dim b As Byte
                b = CByte(pFields("ID_" & Refer))
                Set out = New Stream
                out.Concat ToBytes(Chr(CERT_CONTEXT_CERTFORMAT) & Chr(LenB(b) + CERT_HEX_UNIT) & _
                            Chr(CERT_LEN_WORD) & Chr(LenB(b)) & Chr(b))
      
                Fields = out.Partial

            Case Else
                If Exists("ID_" & Refer) Then
                    Select Case TypeName(pFields("ID_" & Refer))
                        Case "Long", "String"
                            Fields = pFields("ID_" & Refer)
                        Case "Stream"
                            Dim tmp As Stream
                            Set tmp = pFields("ID_" & Refer)
                            If tmp.Length > 0 Then
                                Fields = tmp.Partial
                            End If
                            Set tmp = Nothing
                    End Select
                End If
        End Select
        If inc = CERT_TYPE_OIDS Or inc = CERT_SETTINGS Then
            Set out = New Stream
            Set itm = New Stream

            Set itm = pFields("ID_" & Refer)
            If inc = CERT_SETTINGS Then
                AddLength itm
                out.Concat ToBytes(Chr(inc))
            ElseIf inc = CERT_TYPE_OIDS Then
                out.Concat ToBytes(Chr(inc) & Chr(itm.Length))
            End If
            out.Concat itm.Partial

            If inc = CERT_TYPE_OIDS Then
                out.Concat ToBytes(Chr(CERT_TYPE_NULL) & Chr(0))
            End If
            Fields = out.Partial

        End If

        Set out = Nothing
        Set itm = Nothing
    End If

End Property

Private Function PeekByte(ByRef FromData As Stream) As Byte
    PeekByte = FromData.Partial(0, 1)(1)
End Function

Private Function NextByte(ByRef FromData As Stream) As Byte
    NextByte = FromData.Partial(0, 1)(1)
    FromData.Push 1
End Function

Private Function GetLongNum(ByRef FromData As Stream, Optional ByRef RunLen As Variant) As Long
    If FromData.Length > 0 Then

        If PeekByte(FromData) = CERT_SEQUENCE Or PeekByte(FromData) = CERT_TYPE_BITSET Then
            If FromData.Partial(1, 1)(1) = CERT_LEN_LONG Or FromData.Partial(1, 1)(1) >= (CERT_MAX_WORD \ CERT_LEN_WORD) Then

                If Not IsMissing(RunLen) Then
                    If RunLen <= 0 Then Exit Function
                End If

                Dim ilng As Long
                Dim ostr As String
                ilng = FromData.Partial(1, 1)(1) - (CERT_MAX_WORD \ CERT_LEN_WORD)
                If Not IsMissing(RunLen) Then RunLen = (RunLen - (2 + ilng))
                FromData.Push CERT_HEX_UNIT
                Do While ilng > 0
                    ostr = ostr & Padding(CERT_HEX_UNIT, Hex(NextByte(FromData)), "0")
                    ilng = ilng - 1
                Loop
                GetLongNum = Val("&H" & ostr)
            End If
        End If
    End If
End Function

Private Function GetInteger(ByRef FromData As Stream, Optional ByRef RunLen As Variant) As Variant
    If FromData.Length > 0 Then
        Dim ilng As Long
        ilng = CLng(PeekByte(FromData))
        If ilng = CERT_SEQUENCE Or ilng = CERT_SETTINGS Or ilng = CERT_TYPE_OIDS Or _
            ilng = CERT_TYPE_UTCTIME Or ilng = CERT_LEN_WORD Or _
            ilng - pContextBit = CERT_TYPE_BITSET Or _
            ilng - pContextBit = CERT_LEN_WORD Or ilng = CERT_TYPE_BITSET Then
            
            If Not IsMissing(RunLen) Then
                If RunLen <= 0 Then Exit Function
            End If
            Dim notInt As Boolean
            notInt = (ilng = CERT_SEQUENCE Or ilng = CERT_SETTINGS Or _
                ilng = CERT_TYPE_OIDS Or ilng = CERT_TYPE_UTCTIME Or _
                ilng - pContextBit = CERT_TYPE_BITSET Or ilng = CERT_TYPE_BITSET)
                
            FromData.Push 1

            If ilng = CERT_SEQUENCE Or ilng = CERT_SETTINGS Then
                ilng = CLng(NextByte(FromData))
                GetInteger = ilng
                ilng = 0
            Else
                ilng = CLng(NextByte(FromData))
            End If

            If ilng > (CERT_MAX_WORD \ CERT_LEN_WORD) And ilng < ((CERT_MAX_WORD \ CERT_LEN_WORD) + (CERT_LEN_WORD * CERT_HEX_UNIT)) Then
                Dim ostr As String
                ilng = ilng - (CERT_MAX_WORD \ CERT_LEN_WORD)

                Do While ilng > 0
                    ostr = ostr & Padding(CERT_HEX_UNIT, Hex(NextByte(FromData)), "0")
                    ilng = ilng - 1
                Loop
                ilng = Val("&H" & ostr)

            Else
                If Not IsMissing(RunLen) Then RunLen = RunLen - (ilng + 2)
            End If

            If ilng > 0 Then
                Dim bAry() As Byte
                ReDim bAry(1 To ilng) As Byte

                RtlMoveMemory ByVal VarPtr(bAry(1)), ByVal FromData.Address, ByVal ilng
                FromData.Push ilng
                
                Dim strInt As String
                
                For ilng = LBound(bAry) To UBound(bAry)
                    If notInt Then
                        strInt = strInt & Padding(CERT_HEX_UNIT, Hex(bAry(ilng)), "0") & " "
                    Else
                        strInt = strInt & Padding(CERT_HEX_UNIT, bAry(ilng), "0")
                    End If
                Next

                If notInt Then
                    GetInteger = strInt
                Else
                    GetInteger = Val("&H" & strInt)
                End If
                
                Erase bAry

            End If
        End If
    End If
End Function

Private Function GetSequence(ByRef FromData As Stream, Optional ByRef RunLen As Variant) As Stream
    If FromData.Length > 0 Then
        Dim ilng As Long
doagain:
        
        ilng = CLng(PeekByte(FromData))
        If ilng = CERT_SEQUENCE Or ilng = CERT_SETTINGS Or ilng = CERT_TYPE_OIDS Or _
            ilng = CERT_TYPE_UTCTIME Or ilng - pContextBit = CERT_TYPE_BITSET Or _
            ilng - pContextBit = CERT_LEN_WORD Or ilng = CERT_TYPE_BITSET Or _
            ilng = CERT_LEN_WORD Or ilng = CERT_TYPE_STRING Or ilng = CERT_TYPE_OCTETS Or _
            ilng = CERT_TYPE_BOOLEAN Then

            If Not IsMissing(RunLen) Then
                If RunLen <= 0 Then Exit Function
            End If
            If ilng = CERT_TYPE_OIDS Then
                FromData.Push 1
                ilng = CLng(NextByte(FromData))
            Else
                FromData.Push 1
                ilng = CLng(NextByte(FromData))
            End If

            If ilng > (CERT_MAX_WORD \ CERT_LEN_WORD) And ilng < ((CERT_MAX_WORD \ CERT_LEN_WORD) + (CERT_LEN_WORD * CERT_HEX_UNIT)) Then
                Dim ostr As String
                ilng = ilng - (CERT_MAX_WORD \ CERT_LEN_WORD)
                Do While ilng > 0
                    ostr = ostr & Padding(CERT_HEX_UNIT, Hex(NextByte(FromData)), "0")
                    ilng = ilng - 1
                Loop
                ilng = Val("&H" & ostr)
            Else
                If Not IsMissing(RunLen) Then RunLen = RunLen - (ilng + 2)
            End If

            If ilng > 0 Then
                Dim bAry() As Byte
                ReDim bAry(1 To ilng) As Byte

                RtlMoveMemory ByVal VarPtr(bAry(1)), ByVal FromData.Address, ByVal ilng
                FromData.Push ilng

                Set GetSequence = Stream(bAry)
                Erase bAry

            End If
        Else
            Set GetSequence = New Stream
            GetSequence.Clone FromData
            FromData.Reset

        End If
    End If
End Function
Public Property Get IssuerName() As String ' _
Gets the terminology of the last issuer related property entry of the loaded certificate.
    Dim inc As Integer
    inc = 9
    Do Until Exists("ID_" & (Issuer + inc)) Or inc = 0
        inc = inc - 1
    Loop
    If inc > 0 Then IssuerName = Terms("ID_" & (Issuer + inc))
End Property
Public Property Get SubjectName() As String ' _
Gets the terminology of the last subject related property entry of the loaded certificate.
    Dim inc As Integer
    inc = 9
    Do Until Exists("ID_" & (Subject + inc)) Or inc = 0
        inc = inc - 1
    Loop
    If inc > 0 Then SubjectName = Terms("ID_" & (Subject + inc))
End Property
Public Property Get Keys(ByVal Number As Integer) As String ' _
Returns the key of a specific ordinal Number in the collection of properties for the loaded certificate.

    If Number > 0 And Number <= pKeys.count Then
        Keys = pKeys(Number)
    Else
        Err.Raise 8, , "Ordinal does not exist in collection."
    End If
    
End Property
Public Function Exists(Optional ByVal Refer As Variant = CertificateCollection) As Boolean ' _
Checks for the existance of a property in the loaded certificate by key or name supplied as Refer parameter.
    If Refer = CertificateCollection Then
        Exists = pFields.count > 0
    ElseIf pFields.count > 0 Then

        If Not IsNumeric(Refer) Then
            If Not IsNumeric(Replace(Refer, "ID_", "")) Then
                Refer = Namely(Refer)
            End If
        Else
            Refer = "ID_" & Refer
        End If
        
        Dim cnt As Integer
        For cnt = 1 To pFields.count
            If pKeys(cnt) = Refer Then
                Exists = True
                Exit Function
            End If
        Next
    End If
End Function
Private Sub RemoveField(ByVal Key As Variant)

    pFields.Remove Key
    pKeys.Remove Key
    
End Sub
Private Sub AddField(ByRef Anything, ByVal Key As Variant)

    pFields.Add Anything, Key
    pKeys.Add Key, Key
    
End Sub
Public Function LoadCertificate(ByRef CertData As Stream) As Boolean ' _
Load a certificate in binary form contained in the CertData stream object.
    On Error GoTo badload
    
    Dim FromData As New Stream
    Dim ilng As Long
    Dim tmp As New Stream
    Dim inc As Integer

    FromData.Clone CertData
        
    pNoPrompt = False
    pAccepted = False
    
    Do Until pFields.count = 0
        pFields.Remove 1
    Loop
    Do Until pKeys.count = 0
        pKeys.Remove 1
    Loop
    
    AddField GetLongNum(FromData), "ID_" & CertificateSequence
    
    ilng = CLng(pFields("ID_" & CertificateSequence))
    RemoveField "ID_" & CertificateSequence
    AddField Stream(FromData.Partial(0, ilng)), "ID_" & CertificateSequence
    
    AddField GetLongNum(FromData), "ID_" & CertificateInfoSequence

    ilng = CLng(pFields("ID_" & CertificateInfoSequence))
    RemoveField "ID_" & CertificateInfoSequence
    AddField Stream(FromData.Partial(0, ilng)), "ID_" & CertificateInfoSequence
            
    pContextBit = NextByte(FromData)
        
    AddField GetInteger(FromData, CLng(NextByte(FromData))), "ID_" & Version
        
    AddField GetSequence(FromData), "ID_" & SerialNumber
    
    AddField GetSequence(FromData), "ID_" & Algorithm
    tmp.Clone pFields("ID_" & Algorithm)
    If tmp.Length > 0 Then
        inc = 0
        Do While tmp.Length > 0
            If tmp.Partial(0, 1)(1) = CERT_TYPE_NULL Then
                tmp.Push 2
            Else
                inc = inc + 1
                AddField GetSequence(tmp), "ID_" & (Algorithm + inc)
            End If
        Loop
    End If
    RemoveField "ID_" & Algorithm
            
    AddField GetSequence(FromData), "ID_" & Issuer
    tmp.Clone pFields("ID_" & Issuer)
    inc = 0
    Do While tmp.Length > 0
        inc = inc + 1
        AddField GetSequence(tmp), "ID_" & (Issuer + inc)
    Loop

    AddField GetSequence(FromData), "ID_" & Validity
    tmp.Clone pFields("ID_" & Validity)
    inc = 0
    Do While tmp.Length > 0
        inc = inc + 1
        AddField GetSequence(tmp), "ID_" & (Validity + inc)
    Loop
    RemoveField "ID_" & Validity
        
    AddField GetSequence(FromData), "ID_" & Subject
    tmp.Clone pFields("ID_" & Subject)
    inc = 0
    Do While tmp.Length > 0
        inc = inc + 1
        AddField GetSequence(tmp), "ID_" & (Subject + inc)
    Loop

    AddField GetSequence(FromData), "ID_" & PublicKeyBlock
    tmp.Clone pFields("ID_" & PublicKeyBlock)
    ilng = CLng(GetInteger(tmp))
    tmp.Length = ilng
    AddField Stream(tmp.Partial), "ID_" & IntendedUses
    If tmp.Length > 0 Then
        inc = 0
        Do While tmp.Length > 0
            If tmp.Partial(0, 1)(1) = CERT_TYPE_NULL Then
                tmp.Push 2
            Else
                inc = inc + 1
                AddField GetSequence(tmp), "ID_" & (IntendedUses + inc)
            End If
        Loop
    End If
    tmp.Clone pFields("ID_" & PublicKeyBlock)
    ilng = CLng(GetInteger(tmp))
    tmp.Push ilng
    RemoveField "ID_" & PublicKeyBlock
    RemoveField "ID_" & IntendedUses
    AddField GetSequence(tmp), "ID_" & PublicKeyBlock

    pContextBit = PeekByte(FromData)

    FromData.Placeat ToBytes(CERT_SEQUENCE), 0, 1

    If FromData.Partial(1, 1)(1) = CERT_LEN_LONG Then
        ilng = GetLongNum(FromData)
    ElseIf FromData.Partial(1, 1)(1) = CERT_LEN_WORD Then
        ilng = GetInteger(FromData)
    Else
        ilng = CByte(FromData.Partial(1, 1)(1))
        FromData.Push 2
    End If
       
    tmp.Length = ilng
    tmp.Placeat FromData.Partial, 0, ilng
    FromData.Push ilng

    AddField tmp, "ID_" & Extensions

    Set tmp = New Stream

    tmp.Clone pFields("ID_" & Extensions)
    
    GetSequence tmp
    
    inc = 0
    Do While tmp.Length > 0
        inc = inc + 1
        AddField GetSequence(tmp), "ID_" & (Extensions + inc)
    Loop
    
    AddField GetSequence(FromData), "ID_" & SignatureAlgorithm
    tmp.Clone pFields("ID_" & SignatureAlgorithm)
    If tmp.Length > 0 Then
        inc = 0
        Do While tmp.Length > 0
            If tmp.Partial(0, 1)(1) = CERT_TYPE_NULL Then
                tmp.Push 2
            Else
                inc = inc + 1
                AddField GetSequence(tmp), "ID_" & (SignatureAlgorithm + inc)
            End If
        Loop
    End If
    RemoveField "ID_" & SignatureAlgorithm
        
    If FromData.Length > 0 Then
        AddField GetSequence(FromData), "ID_" & Signature
    End If

    Set tmp = pFields("ID_" & PublicKeyBlock)
    If tmp.Length < 512 And tmp.Length >= 256 Then
        pKeySize = 256
    ElseIf tmp.Length < 256 And tmp.Length >= 128 Then
        pKeySize = 128
    Else
        pKeySize = 256
    End If
badload:
    
    Set tmp = Nothing
    Set FromData = Nothing
        
    If Err Then
        On Error GoTo 0
        Err.Raise 8, , "Unable to load certificate"
    End If

End Function

Private Sub Class_Initialize()
    Set pFields = New Collection
    Set pKeys = New Collection
End Sub

Private Sub Class_Terminate()
    Do Until pFields.count = 0
        pFields.Remove 1
    Loop
    Set pFields = Nothing
    Do Until pKeys.count = 0
        pKeys.Remove 1
    Loop
    Set pKeys = Nothing
End Sub
