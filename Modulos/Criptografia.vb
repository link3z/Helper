Imports System.IO
Imports System.Windows.Forms
Imports System.Security.Cryptography
Imports System.Text

Namespace Criptografia
    Public Module modRecompilaCriptografia
#Region " METODOS AUXILIARES "
        ''' <summary>
        ''' Crea una clave válida a través de la clave especificada
        ''' </summary>
        Private Function CreateKey(ByVal strPassword As String) As Byte()
            'Convert strPassword to an array and store in chrData.
            Dim chrData() As Char = strPassword.ToCharArray
            'Use intLength to get strPassword size.
            Dim intLength As Integer = chrData.GetUpperBound(0)
            'Declare bytDataToHash and make it the same size as chrData.
            Dim bytDataToHash(intLength) As Byte

            'Use For Next to convert and store chrData into bytDataToHash.
            For i As Integer = 0 To chrData.GetUpperBound(0)
                bytDataToHash(i) = CByte(Asc(chrData(i)))
            Next

            'Declare what hash to use.
            Dim SHA512 As New System.Security.Cryptography.SHA512Managed
            'Declare bytResult, Hash bytDataToHash and store it in bytResult.
            Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
            'Declare bytKey(31).  It will hold 256 bits.
            Dim bytKey(31) As Byte

            'Use For Next to put a specific size (256 bits) of 
            'bytResult into bytKey. The 0 To 31 will put the first 256 bits
            'of 512 bits into bytKey.
            For i As Integer = 0 To 31
                bytKey(i) = bytResult(i)
            Next

            Return bytKey 'Return the key.
        End Function

        ''' <summary>
        ''' Crea una clave IV válida a través de la clave especificada
        ''' </summary>
        Private Function CreateIV(ByVal strPassword As String) As Byte()
            'Convert strPassword to an array and store in chrData.
            Dim chrData() As Char = strPassword.ToCharArray
            'Use intLength to get strPassword size.
            Dim intLength As Integer = chrData.GetUpperBound(0)
            'Declare bytDataToHash and make it the same size as chrData.
            Dim bytDataToHash(intLength) As Byte

            'Use For Next to convert and store chrData into bytDataToHash.
            For i As Integer = 0 To chrData.GetUpperBound(0)
                bytDataToHash(i) = CByte(Asc(chrData(i)))
            Next

            'Declare what hash to use.
            Dim SHA512 As New System.Security.Cryptography.SHA512Managed
            'Declare bytResult, Hash bytDataToHash and store it in bytResult.
            Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
            'Declare bytIV(15).  It will hold 128 bits.
            Dim bytIV(15) As Byte

            'Use For Next to put a specific size (128 bits) of 
            'bytResult into bytIV. The 0 To 30 for bytKey used the first 256 bits.
            'of the hashed password. The 32 To 47 will put the next 128 bits into bytIV.
            For i As Integer = 32 To 47
                bytIV(i - 32) = bytResult(i)
            Next

            Return bytIV 'return the IV
        End Function
#End Region

#Region " MD5 Y SHA256 "
        ''' <summary>
        ''' Obtiene el MD5 correspondiente a la cadena especificada
        ''' </summary>
        ''' <param name="eCadena">Cadena de la que se quiere obtener el MD5</param>
        ''' <returns>MD5 calculado para la cadena de entrada</returns>
        Public Function encriptarEnMD5(ByVal eCadena As String) As String
            ' Si no se le pasa una cadena no se puede convertir
            If String.IsNullOrEmpty(eCadena) Then Return ""

            Dim Md5csp As New Security.Cryptography.MD5CryptoServiceProvider
            Dim bites() As Byte = System.Text.Encoding.ASCII.GetBytes(eCadena)
            Dim Resultado As New System.Text.StringBuilder()

            ' Encriptración en md5
            bites = Md5csp.ComputeHash(bites)

            For Each b As Byte In bites
                Resultado.Append(b.ToString("x2"))
            Next

            Return Resultado.ToString()
        End Function

        ''' <summary>
        ''' Obtiene el SHA256 correspondiente a la cadena especificada
        ''' </summary>
        ''' <param name="eCadena">Cadena de la que se quiere obtener el SHA256</param>
        ''' <returns>SHA256 calculado para la cadena de entrada</returns>
        Public Function encriptarEnSHA256(ByVal eCadena As String) As String
            ' Si no se le pasa una cadena no se puede convertir
            If String.IsNullOrEmpty(eCadena) Then Return ""

            Dim uEncode As New UnicodeEncoding()
            Dim bytClearString() As Byte = uEncode.GetBytes(eCadena)
            Dim sha As New System.Security.Cryptography.SHA256Managed()
            Dim hash() As Byte = sha.ComputeHash(bytClearString)
            Return Convert.ToBase64String(hash)
        End Function

        ''' <summary>
        ''' Calcula el MD5 para el fichero que se pasa como argumento
        ''' </summary>
        ''' <param name="eRutaFichero">Fichero para el cálculo del MD5</param>
        ''' <returns>MD5 del fichero que se le pasa como parámetro</returns>
        Public Function calcularMD5Fichero(ByVal eRutaFichero As String) As String
            If System.IO.File.Exists(eRutaFichero) Then
                Dim cadenaMD5 As String = ""
                Dim cadenaFichero As FileStream
                Dim bytesFichero As [Byte]()
                Dim MD5Crypto As New MD5CryptoServiceProvider
                Try
                    cadenaFichero = File.Open(eRutaFichero, FileMode.Open, FileAccess.Read)
                    bytesFichero = MD5Crypto.ComputeHash(cadenaFichero)
                    cadenaFichero.Close()
                    cadenaMD5 = BitConverter.ToString(bytesFichero)
                    cadenaMD5 = cadenaMD5.Replace("-", "")
                Catch ex As System.Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al calcular el MD5 del fichero.", ex, New StackTrace(0, True))
                    Return ""
                End Try
                Return cadenaMD5
            Else
                If Log._LOG_ACTIVO Then Log.escribirLog("No existe el fichero de entrada en '" & eRutaFichero & "'.", , New StackTrace(0, True))
                Throw New IO.FileNotFoundException
            End If
        End Function
#End Region

#Region " ENCRIPTACIÓN "
        ''' <summary>
        ''' Encripta el fichero especificado como parámetro con la clave que se 
        ''' pasa como parámetro
        ''' </summary>
        ''' <param name="eRutaFichero">Ruta del archivo que queremos comprimir</param>
        ''' <param name="eRutaFicheroEncriptado">Ruta del archivo resultante comprimido</param>
        ''' <param name="eClave">Clave utilizada para la encripcación</param>
        ''' <param name="eSegundaClave">Segunda clave de encripcación</param>
        ''' <returns>True o False dependiendo del resultado de la encriptación</returns>
        Public Function encriptarFichero(ByVal eRutaFichero As String, _
                                         ByVal eRutaFicheroEncriptado As String, _
                                         ByVal eClave As String, _
                                         Optional ByVal eSegundaClave As String = "") As Boolean
            Dim fsOriginal As System.IO.FileStream = Nothing
            Dim fsEncriptado As System.IO.FileStream = Nothing
            Dim cryptostream As CryptoStream = Nothing

            If String.IsNullOrEmpty(eSegundaClave) Then eSegundaClave = eClave

            Try

                'Setup file streams to handle input and output.
                fsOriginal = New System.IO.FileStream(eRutaFichero, FileMode.Open, FileAccess.Read)
                fsEncriptado = New System.IO.FileStream(eRutaFicheroEncriptado, FileMode.OpenOrCreate, FileAccess.Write)
                fsEncriptado.SetLength(0) 'make sure fsOutput is empty

                'Declare variables for encrypt/decrypt process.
                Dim bytBuffer(4096) As Byte 'holds a block of bytes for processing
                Dim lngBytesProcessed As Long = 0 'running count of bytes processed
                Dim lngFileLength As Long = fsOriginal.Length 'the input file's length
                Dim intBytesInCurrentBlock As Integer 'current bytes being processed
                'Declare your CryptoServiceProvider.
                Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged

                cryptostream = New CryptoStream(fsEncriptado, cspRijndael.CreateEncryptor(CreateKey(eClave), CreateIV(eSegundaClave)), CryptoStreamMode.Write)


                'Use While to loop until all of the file is processed.
                While lngBytesProcessed < lngFileLength
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsOriginal.Read(bytBuffer, 0, 4096)

                    'Write output file with the cryptostream.
                    cryptostream.Write(bytBuffer, 0, intBytesInCurrentBlock)

                    'Update lngBytesProcessed
                    lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                End While

                Return True

            Catch ex As System.Exception
                Throw ex
                Return False
            Finally
                If cryptostream IsNot Nothing Then cryptostream.Close()
                If fsOriginal IsNot Nothing Then fsOriginal.Close()
                If fsEncriptado IsNot Nothing Then fsEncriptado.Close()
            End Try

        End Function

        ''' <summary>
        ''' Desencripta el fichero especificado como parámetro con la clave dada
        ''' </summary>
        ''' <param name="eRutaFichero">Ruta del archivo que queremos desencriptar</param>
        ''' <param name="eRutaFicheroDesencriptado">Ruta del archivo resultante desencriptado</param>
        ''' <param name="eClave">Clave utilizada para desencriptación</param>
        ''' <param name="eSegundaClave">Segunda clave de encripcación</param>
        ''' <returns>True o False dependiendo del resultado de la desencreiptación</returns>
        Public Function desencriptarFichero(ByVal eRutaFichero As String, _
                                            ByVal eRutaFicheroDesencriptado As String, _
                                            ByVal eClave As String, _
                                            Optional ByVal eSegundaClave As String = "") As Boolean
            Dim fsEncriptado As System.IO.FileStream = Nothing
            Dim fsDesencriptado As System.IO.FileStream = Nothing
            Dim cryptostream As CryptoStream = Nothing

            If String.IsNullOrEmpty(eSegundaClave) Then eSegundaClave = eClave

            Try

                'Setup file streams to handle input and output.
                fsEncriptado = New System.IO.FileStream(eRutaFichero, FileMode.Open, FileAccess.Read)
                fsDesencriptado = New System.IO.FileStream(eRutaFicheroDesencriptado, FileMode.OpenOrCreate, FileAccess.Write)
                fsDesencriptado.SetLength(0) 'make sure fsOutput is empty

                'Declare variables for encrypt/decrypt process.
                Dim bytBuffer(4096) As Byte 'holds a block of bytes for processing
                Dim lngBytesProcessed As Long = 0 'running count of bytes processed
                Dim lngFileLength As Long = fsEncriptado.Length 'the input file's length
                Dim intBytesInCurrentBlock As Integer 'current bytes being processed
                'Declare your CryptoServiceProvider.
                Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged

                cryptostream = New CryptoStream(fsDesencriptado, cspRijndael.CreateDecryptor(CreateKey(eClave), CreateIV(eSegundaClave)), CryptoStreamMode.Write)


                'Use While to loop until all of the file is processed.
                While lngBytesProcessed < lngFileLength
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsEncriptado.Read(bytBuffer, 0, 4096)

                    'Write output file with the cryptostream.
                    cryptostream.Write(bytBuffer, 0, intBytesInCurrentBlock)

                    'Update lngBytesProcessed
                    lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                End While

                Return True

            Catch ex As System.Exception
                Throw ex
                Return False
            Finally
                If cryptostream IsNot Nothing Then cryptostream.Close()
                If fsEncriptado IsNot Nothing Then fsEncriptado.Close()
                If fsDesencriptado IsNot Nothing Then fsDesencriptado.Close()
            End Try

        End Function
#End Region

#Region " BASE 64 "
        ''' <summary>
        ''' Convierte el string a base 64 string
        ''' </summary>
        ''' <param name="eEntrada">Cadena de entrada</param>
        ''' <returns>Cadena convertida a base 64</returns>
        Public Function string2Base64(ByVal eEntrada As String) As String
            If Not String.IsNullOrEmpty(eEntrada) Then
                Try
                    Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(eEntrada)
                    Return (Convert.ToBase64String(byt))
                Catch ex As Exception
                    Return String.Empty
                End Try
            Else
                Return eEntrada
            End If
        End Function

        ''' <summary>
        ''' Convierte una cadena en base 64 a una cadena nomrmal
        ''' </summary>
        ''' <param name="eEntrada">Cadena en base 64</param>
        ''' <returns>Cadena nomrmal</returns>
        Public Function base642String(ByVal eEntrada As String) As String
            If Not String.IsNullOrEmpty(eEntrada) Then
                Try
                    Dim b As Byte() = Convert.FromBase64String(eEntrada)
                    Return (System.Text.Encoding.UTF8.GetString(b))
                Catch ex As Exception
                    Return String.Empty
                End Try
            Else
                Return eEntrada
            End If
        End Function
#End Region
    End Module
End Namespace