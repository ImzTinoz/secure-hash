#Region "Imports"
Imports System.Security
Imports System.Security.Cryptography
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Text
#End Region

Module Main

#Region "Subs"
    Sub Main()

        Console.Clear()

        Console.WindowWidth = 78
        Console.BufferWidth = 78

        Console.BufferHeight = Console.WindowHeight

        Console.WriteLine(My.Resources.splash)
        Console.WriteLine("")
        Console.WriteLine("")
        Console.WriteLine("")
        Console.Write("                              > ")
        Dim cki As ConsoleKeyInfo = Console.ReadKey()
        If cki.Key = ConsoleKey.E Then
            Encrypt()
        ElseIf cki.Key = ConsoleKey.D Then
            Decrypt()
        Else
            Main()
        End If

    End Sub
    Sub Encrypt()

        Console.Clear()

        Console.WindowWidth = 78
        Console.BufferWidth = 78

        Console.BufferHeight = Console.WindowHeight

        Console.WriteLine(My.Resources.encrypt)
        Console.WriteLine("")
        Console.Write("                              > ")
        Dim cki As ConsoleKeyInfo = Console.ReadKey()

        If cki.Key = ConsoleKey.D1 Then
            StartEncrypt(1)
        ElseIf cki.Key = ConsoleKey.D2 Then
            StartEncrypt(2)
        ElseIf cki.Key = ConsoleKey.D3 Then
            StartEncrypt(4)
        ElseIf cki.Key = ConsoleKey.D4 Then
            StartEncrypt(8)
        Else
            Main()
        End If

    End Sub
    Sub Decrypt()

        Console.Clear()

        Console.WindowWidth = 78
        Console.BufferWidth = 78

        Console.BufferHeight = Console.WindowHeight

        Console.WriteLine(My.Resources.decrypt)
        Console.WriteLine("")
        Console.Write("                              > ")
        Dim cki As ConsoleKeyInfo = Console.ReadKey()

        If cki.Key = ConsoleKey.D1 Then
            StartDecrypt(1)
        ElseIf cki.Key = ConsoleKey.D2 Then
            StartDecrypt(2)
        ElseIf cki.Key = ConsoleKey.D3 Then
            StartDecrypt(4)
        ElseIf cki.Key = ConsoleKey.D4 Then
            StartDecrypt(8)
        Else
            Main()
        End If

    End Sub
    Sub StartEncrypt(ByVal rounds As Integer)

        Dim passw As Integer
        Dim partinput As String

        Dim passwList As New List(Of String)

        Console.Clear()

        Console.WindowWidth = 78
        Console.BufferWidth = 78

        Console.BufferHeight = Console.WindowHeight

        Console.WriteLine(My.Resources.encrypt_2)
        Console.WriteLine("")
        Console.Write("                              > ")
        Dim input As String = Console.ReadLine()
        partinput = input
        'passw(10) = 0
        Console.WriteLine("")
        Console.WriteLine("")
        Console.WriteLine("   Log:")
        Console.WriteLine("")

        For a = 10 To (rounds + 9) Step 1
            Dim r As New Random
            passw = r.Next(100000, 999999)
            passwList.Add(passw)
            Console.Write("   >  PW(" & a & ") ")
            Console.ForegroundColor = ConsoleColor.DarkGreen
            Console.WriteLine(passw)
            Console.ResetColor()
            Console.ForegroundColor = ConsoleColor.DarkGray
            Console.WriteLine("   >  ENC =  " & Truncate(partinput, 20))
            Console.ResetColor()
            If a = 10 Then
                partinput = Encrypt(input, passw)
            Else
                partinput = Encrypt(partinput, passw)
            End If
            Console.ForegroundColor = ConsoleColor.DarkGray
            Console.WriteLine("   >  PI(" & a & ") " & Truncate(partinput, 20))
            Console.ResetColor()

            Threading.Thread.Sleep(CInt(Int((500 * Rnd()) + 100)))
        Next

        Console.WriteLine("")
        Console.WriteLine("   Encryption done. Below you will find a list of passwords in order. " & vbCrLf & "   Write these down - you will need them for decryption.")

        Console.WriteLine("")

        Dim message = String.Join(Environment.NewLine & "   > ", passwList.ToArray())
        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("   > " & message)
        Console.ResetColor()

        Console.WriteLine("")
        Console.WriteLine("   The fully encrypted input string has been copied to your clipboard. " & vbCrLf & vbCrLf & "   If you want to decrypt it, make sure you have the crypted input" & vbCrLf & "   and the passwords in order." & vbCrLf & vbCrLf & "   Press any key to exit.")
        Console.WriteLine("")
        My.Computer.Clipboard.SetText(partinput)
        Console.ReadKey()
        Main()

    End Sub
    Sub StartDecrypt(ByVal rounds As Integer)

        Dim passwList As New List(Of String)

        Console.Clear()

        Console.WindowWidth = 78
        Console.BufferWidth = 78

        Console.BufferHeight = Console.WindowHeight

        Console.WriteLine(My.Resources.encrypt_2)
        Console.WriteLine("")
        Console.WriteLine("         Copy the encrypted string to clipboard and press any key.")
        Dim confirmationKey As ConsoleKeyInfo = Console.ReadKey()
        Dim input_dc As String = My.Computer.Clipboard.GetText
        Console.WriteLine("")
        Dim decStr As String

        For a = rounds To 1 Step -1
            Console.Write("   >  Enter password number " & a & ": ")
            Console.ForegroundColor = ConsoleColor.DarkGreen
            Dim pwinput As String = Console.ReadLine
            Console.ResetColor()
            Console.ForegroundColor = ConsoleColor.DarkGray
            Threading.Thread.Sleep(CInt(Int((500 * Rnd()) + 100)))
            If a = rounds Then
                decStr = Decrypt(input_dc, pwinput)
            Else
                decStr = Decrypt(decStr, pwinput)
            End If

            Console.WriteLine("   >  DEC =  " & decStr) 'Truncate(decStr, 20))
            Console.ResetColor()
        Next

        Console.WriteLine("")
        Console.WriteLine("   Decryption done.")

        Console.WriteLine("")

        Console.WriteLine("   > " & decStr)

        Console.WriteLine("")
        Console.WriteLine("   Press any key to exit.")
        Console.WriteLine("")
        My.Computer.Clipboard.SetText(decStr)
        Console.ReadKey()
        Main()

    End Sub
#End Region

#Region "Cryptography"
    Public Function Encrypt(ByVal input As String, ByVal pass As String) As String

        Dim passPhrase As String = pass
        Dim saltValue As String = "9h1d7i0r1s6v8n7m6s"
        Dim hashAlgorithm As String = "SHA1"

        Dim passwordIterations As Integer = 2
        Dim initVector As String = "@1B2c3D4e5F6g7H8"
        Dim keySize As Integer = 256

        Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(initVector)
        Dim saltValueBytes As Byte() = Encoding.ASCII.GetBytes(saltValue)

        Dim plainTextBytes As Byte() = Encoding.UTF8.GetBytes(input)


        Dim password As New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)

        Dim keyBytes As Byte() = password.GetBytes(keySize \ 8)
        Dim symmetricKey As New RijndaelManaged()

        symmetricKey.Mode = CipherMode.CBC

        Dim encryptor As ICryptoTransform = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)

        Dim memoryStream As New MemoryStream()
        Dim cryptoStream As New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)

        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)
        cryptoStream.FlushFinalBlock()
        Dim cipherTextBytes As Byte() = memoryStream.ToArray()
        memoryStream.Close()
        cryptoStream.Close()
        Dim cipherText As String = Convert.ToBase64String(cipherTextBytes)
        Return cipherText
    End Function
    Public Function Decrypt(ByVal input As String, ByVal pass As String) As String
        Dim passPhrase As String = pass
        Dim saltValue As String = "9h1d7i0r1s6v8n7m6s"
        Dim hashAlgorithm As String = "SHA1"

        Dim passwordIterations As Integer = 2
        Dim initVector As String = "@1B2c3D4e5F6g7H8"
        Dim keySize As Integer = 256

        Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(initVector)
        Dim saltValueBytes As Byte() = Encoding.ASCII.GetBytes(saltValue)

        Dim cipherTextBytes As Byte() = Convert.FromBase64String(input)

        Dim password As New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)

        Dim keyBytes As Byte() = password.GetBytes(keySize \ 8)

        Dim symmetricKey As New RijndaelManaged()

        symmetricKey.Mode = CipherMode.CBC

        Dim decryptor As ICryptoTransform = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)

        Dim memoryStream As New MemoryStream(cipherTextBytes)

        Dim cryptoStream As New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)

        Dim plainTextBytes As Byte() = New Byte(cipherTextBytes.Length - 1) {}

        Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)

        memoryStream.Close()
        cryptoStream.Close()

        Dim plainText As String = Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)

        Return plainText
    End Function
#End Region

#Region "Unused"
    Public Function AES_Encrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim encrypted As String = ""
        Dim error_msg As String = "Something went wrong while encrypting."
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = Security.Cryptography.CipherMode.ECB
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
            encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return encrypted
        Catch ex As Exception
            MsgBox("Unexpected error occured.")
            End
        End Try
    End Function
    Public Function AES_Decrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim decrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = Security.Cryptography.CipherMode.ECB
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = Convert.FromBase64String(input)
            decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return decrypted
        Catch ex As Exception
            MsgBox(ex.Message)
            End
        End Try
    End Function
#End Region

#Region "Other Functions"
    Public Function Truncate(value As String, length As Integer) As String
        If length > value.Length Then
            Return value
        Else
            Return value.Substring(0, length) & "..."
        End If
    End Function
#End Region

End Module
