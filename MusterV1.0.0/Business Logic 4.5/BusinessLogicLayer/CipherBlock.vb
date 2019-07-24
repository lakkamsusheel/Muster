Imports System.Security.Cryptography
Imports System.Text
Public Class CipherBlock


    Public Enum CryptoProviders
        Rijndael
        RC2
        TripleDES
        DES
    End Enum

    'The plaintext message in a byte array
    Private PlainTextBArray As Byte()
    'The cyphertext message in a byte array
    Private CypherTextBArray As Byte()
    'The key used to encrypt/decrypt
    Private TheKey As String = String.Empty
    'The initialization vector used for encryption
    Private TheIV As String = String.Empty
    'The key used to encrypt/decrypt as a byte array
    Private barrTheKey As Byte()
    'The initialization vector used for encryption as a byte array
    Private barrTheIV As Byte()
    'The provider selected for encryption/decryption
    Private TheProvider As SymmetricAlgorithm = Nothing
    'The enumeration selected for the cryptoprovider
    Private TheProviderEnum As Integer = 0


    '
    ' Get/Set the encryption key
    Public Property Password() As String
        Get
            Return TheKey
        End Get
        Set(ByVal Value As String)
            TheKey = Value
        End Set
    End Property
    '
    ' Get/Set the initialization vector
    Public Property IV() As String
        Get
            Return TheIV
        End Get
        Set(ByVal Value As String)
            TheIV = Value
        End Set
    End Property
    '
    ' Get/Set the provider
    Public Property Provider() As CryptoProviders
        Get
            Return TheProviderEnum
        End Get
        Set(ByVal Value As CryptoProviders)
            TheProviderEnum = Value
            Select Case TheProviderEnum
                Case CryptoProviders.Rijndael
                    TheProvider = New RijndaelManaged
                Case CryptoProviders.DES
                    TheProvider = New DESCryptoServiceProvider
                Case CryptoProviders.RC2
                    TheProvider = New RC2CryptoServiceProvider
                Case CryptoProviders.TripleDES
                    TheProvider = New TripleDESCryptoServiceProvider
            End Select
        End Set
    End Property
    '
    'Pad the private key provided (if necessary) with spaces to meet 
    ' legal key size requirements
    Private Function GetLegalKey() As Byte()
        Dim LessSize As Integer = 0
        Dim MoreSize As Integer = TheProvider.LegalKeySizes(0).MinSize
        If TheProvider.LegalKeySizes.Length > 0 Then
            While TheKey.Length * 8 < MoreSize And TheProvider.LegalKeySizes(0).SkipSize > 0 And MoreSize < TheProvider.LegalKeySizes(0).MaxSize
                LessSize = MoreSize
                MoreSize += TheProvider.LegalKeySizes(0).SkipSize
            End While

            If TheKey.Length > MoreSize / 8 Then
                Return ASCIIEncoding.ASCII.GetBytes(TheKey.Substring(0, (MoreSize / 8)))
            Else
                Return ASCIIEncoding.ASCII.GetBytes(TheKey.PadRight(MoreSize / 8, " "))
            End If
        Else
            Return ASCIIEncoding.ASCII.GetBytes(TheKey)
        End If
    End Function
    '
    'Pad the initialization vector to the proper length
    Private Function GetLegalIV() As Byte()
        If TheIV.Length > TheProvider.IV.Length Then
            Return ASCIIEncoding.ASCII.GetBytes(TheIV.Substring(0, TheProvider.IV.Length))
        Else
            Return ASCIIEncoding.ASCII.GetBytes(TheIV.PadRight(TheProvider.IV.Length, " "))
        End If
    End Function
    '
    'Encrypt the string
    Public Function Encrypt(ByVal TheSource As String) As String
        Dim ms As New System.IO.MemoryStream
        Dim strSource As Byte() = ASCIIEncoding.ASCII.GetBytes(TheSource)
        Dim EncryptStream As CryptoStream
        Dim strReturn As String

        Try
            TheProvider.Key = GetLegalKey()
            TheProvider.IV = GetLegalIV()
            EncryptStream = New CryptoStream(ms, TheProvider.CreateEncryptor, CryptoStreamMode.Write)

            EncryptStream.Write(strSource, 0, strSource.Length)
            EncryptStream.FlushFinalBlock()

            strReturn = System.Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
            ms.Close()
        Catch ex As Exception
            Throw ex
            If Not ms Is Nothing Then
                ms.Close()
                Exit Function
            End If
        End Try

        Return strReturn

    End Function
    '
    'Decrypt a string
    Public Function Decrypt(ByVal TheSource As String) As String
        Dim ms As System.IO.MemoryStream
        Dim strSource As Byte() = System.Convert.FromBase64String(TheSource)
        Dim DecryptStream As CryptoStream
        Dim sr As System.IO.StreamReader
        Dim strReturn As String

        ms = New System.IO.MemoryStream(strSource, 0, strSource.Length)
        TheProvider.Key = GetLegalKey()
        TheProvider.IV = GetLegalIV()
        DecryptStream = New CryptoStream(ms, TheProvider.CreateDecryptor, CryptoStreamMode.Read)
        sr = New System.IO.StreamReader(DecryptStream)

        Try
            strReturn = sr.ReadToEnd
            ms.Close()
            Return strReturn
        Catch ex As Exception
            If ex.Message.ToUpper.StartsWith("BAD DATA") Then
                Return "Bad Password"
            Else
                Throw ex
            End If
        End Try

    End Function

End Class
