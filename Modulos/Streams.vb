Imports System.IO

Namespace Streams
    Public Module modRecompilaStreams
#Region " METODOS "
        ''' <summary>
        ''' Convierte un Stream de cualquier tipo en un Array de Bytes
        ''' </summary>
        ''' <param name="eStream">Stream que se va a convertir en un Array de Bytes</param>
        ''' <returns>Array de bytes generado a partir del Stream</returns>
        Public Function Stream2ByteArray(ByVal eStream As System.IO.Stream) As Byte()
            Try
                Dim laLongitud As Integer = Convert.ToInt32(eStream.Length)
                Dim paraDevolver As Byte() = New Byte(laLongitud) {}

                eStream.Read(paraDevolver, 0, laLongitud)
                eStream.Close()

                Return paraDevolver
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error de conversión...", ex, New StackTrace(0, True))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Convierte un Array de Bytes en un Stream
        ''' </summary>
        ''' <param name="eByte">Array de bytes a convertir</param>
        ''' <returns>Stream generado a partir del array de bytes</returns>
        Public Function ByteArray2Stream(ByVal eByte As Byte()) As System.IO.Stream
            Try
                Dim paraDevolver As New MemoryStream
                paraDevolver.Write(eByte, 0, eByte.Length)
                paraDevolver.Position = 0
                Return paraDevolver
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error de conversión...", ex, New StackTrace(0, True))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Permite guardar un Stream en un fichero
        ''' </summary>
        ''' <param name="eStream">Stream que se quiere guardar</param>
        ''' <param name="eRutaFichero">Ruta del fichero donde se va a guardar</param>
        ''' <param name="eCrearFichero">Si se tiene que crear el fichero en caso de que este no exista</param>
        ''' <returns>True/False dependiendo de si se pudo efectuar o no la operación</returns>
        Public Function Stream2File(ByVal eStream As Stream, _
                                    ByVal eRutaFichero As String, _
                                    Optional ByVal eCrearFichero As Boolean = True) As Boolean

            ' Si no existe el fichero y no se permite crear no se puede efectuar la operación
            If Not File.Exists(eRutaFichero) And Not eCrearFichero Then Return False

            Try
                Dim elFichero As New FileStream(eRutaFichero, FileMode.Create, System.IO.FileAccess.Write)
                Dim bytes As Byte() = New Byte(eStream.Length - 1) {}

                eStream.Read(bytes, 0, CInt(eStream.Length))
                elFichero.Write(bytes, 0, bytes.Length)
                eStream.Position = 0

                elFichero.Close()
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error de conversión...", ex, New StackTrace(0, True))
                Return Nothing
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Lee un ficheor y lo devuelve como un Stream
        ''' </summary>
        ''' <param name="eRutaFichero">Ruta del ficheor que se va a leer</param>
        ''' <returns>Stream obtenido del fichero</returns>
        Public Function File2Stream(ByVal eRutaFichero As String) As Stream
            Try
                Dim paraDevolver As New MemoryStream()
                Dim elFichero As New FileStream(eRutaFichero, FileMode.Open, FileAccess.Read)
                Dim bytes As Byte() = New Byte(elFichero.Length - 1) {}
                elFichero.Read(bytes, 0, CInt(elFichero.Length))
                paraDevolver.Write(bytes, 0, CInt(elFichero.Length))
                paraDevolver.Position = 0
                elFichero.Close()

                Return paraDevolver
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error de conversión...", ex, New StackTrace(0, True))
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Convierte un fichero en un array de bytes
        ''' </summary>
        ''' <param name="eRutaArchivo">Ruta del archivo a convertir</param>
        ''' <returns>Byte array generado a partir del archivo</returns>
        Public Function File2Byte(ByVal eRutaArchivo As String) As Byte()
            Return Ficheros.Archivo2Byte(eRutaArchivo)
        End Function

        ''' <summary>
        ''' Guarda un array de bytes en un fichero
        ''' </summary>
        ''' <param name="eByte">Array de bytes a guardar</param>
        ''' <param name="eRutaSalida">Ruta donde se debe guardar el array de bytes</param>
        ''' <returns>True o False dependiendo de si se ejecutó correctamente la operación</returns>
        Public Function Byte2File(ByVal eByte() As Byte, _
                                  ByVal eRutaSalida As String) As Boolean
            Return Ficheros.Byte2Archivo(eByte, eRutaSalida)
        End Function
#End Region
    End Module
End Namespace
