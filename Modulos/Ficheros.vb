Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Drawing
Imports System.Reflection

Namespace Ficheros
    ''' <summary>
    ''' Metodos para facilitar el trabajo con ficheros
    ''' </summary>
    Public Module modRecompilaFicheros
#Region " DECLARACIONES "
        ''' <summary>
        ''' Listado con caracteres no permitidos para los nombres de los archivos y carpetas.
        ''' Algunos de estos caracteres son permitidos por el sistema operativo pero se excluyen
        ''' para facilitar el uso de otras funciones y métodos
        ''' </summary>
        Public _CARACTERES_NO_PERMITIDOS As New List(Of String) From {"'", "?", "[", "]", "/", "\", "*", ":", """", "|", "<", ">", "´"}
#End Region

#Region " METODOS COMUNES "
        ''' <summary>
        ''' Elimina los caracteres no permitidos en el nombre de un archivo o carpeta
        ''' </summary>
        ''' <param name="eNombre">Nombre original</param>
        ''' <returns>Nuevo nombre eliminando los caracteres no permitidos</returns>
        Public Function LimpiarNombre(ByVal eNombre As String) As String
            Dim paraDevolver As String = eNombre

            If _CARACTERES_NO_PERMITIDOS IsNot Nothing AndAlso _CARACTERES_NO_PERMITIDOS.Count > 0 Then
                For Each unCaracter As String In _CARACTERES_NO_PERMITIDOS
                    paraDevolver = Replace(paraDevolver, unCaracter, "")
                Next
            End If

            Return paraDevolver
        End Function
#End Region

#Region " METODOS FICHEROS "
        ''' <summary>
        ''' Obtiene un nombre para un documento evitando que este sea igual que otro 
        ''' en que exista en el mismo directorio, para ello, concatena _XXXX donde XXXX
        ''' es un número consecutivo al anterior creado
        ''' </summary>
        ''' <param name="eCarpetaDestino">Carpeta donde se va a comprobar el nombre</param>
        ''' <param name="eNombreDestino">Nombre del documento que se quiere guardar</param>
        ''' <param name="eExtensionDestinoConPunto">Extensión del documento (Con punto)</param>
        ''' <param name="eDevolverConExtension">Determina si el nuevo nombre se devuelve con la extensión</param>
        ''' <returns>Devuelve el nombre del documento para que no se repita</returns>
        Public Function obtenerNombreConsecutivo(ByVal eCarpetaDestino As String, _
                                                 ByVal eNombreDestino As String, _
                                                 ByVal eExtensionDestinoConPunto As String, _
                                                 Optional ByVal eDevolverConExtension As Boolean = False) As String
            Dim Aux As Long = 0
            Dim nuevoNombre As String = eNombreDestino

            ' Si existe el documento en destino se realiza un bucle incrementando el contador
            ' hasta que se encuentra uno que todavía no exista
            If IO.File.Exists(eCarpetaDestino & nuevoNombre & eExtensionDestinoConPunto) Then
                Do
                    Aux += 1
                    nuevoNombre = eNombreDestino & "_" & Format(Aux, "0000")
                Loop While IO.File.Exists(eCarpetaDestino & nuevoNombre & eExtensionDestinoConPunto)
            End If

            If eDevolverConExtension Then
                Return nuevoNombre & eExtensionDestinoConPunto
            Else
                Return nuevoNombre
            End If
        End Function

        ''' <summary>
        ''' Extrae el nombre y la extensión del fichero de la ruta que se le pasa a la función.
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa del fichero de donde extraer el nombre</param>
        ''' <returns>Nombre y extensión del fichero</returns>
        Public Function extraerNombreFichero(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetFileName(eRutaCompleta)
        End Function

        ''' <summary>
        ''' Extrae el nombre sin la extensión del fichero de la ruta que se le pasa a la función.
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa del fichero de donde extraer el nombre</param>
        ''' <returns>Nombre del fichero</returns>
        Public Function extraerNombreFicheroSinExtension(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetFileNameWithoutExtension(eRutaCompleta)
        End Function

        ''' <summary>
        ''' Extrae la extensión del fichero que se le pasa como parametro.
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa de la ubicación del fichero</param>
        ''' <returns>Extensión del fichero</returns>
        Public Function extraerExtensionFicheroConPunto(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetExtension(eRutaCompleta)
        End Function

        ''' <summary>
        ''' Extrae la extensión del fichero que se le pasa como parametro sin el punto.
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa de la ubicación del fichero</param>
        ''' <returns>Extensión del fichero sin punto</returns>
        Public Function extraerExtensionFicheroSinPunto(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetExtension(eRutaCompleta).Replace(".", "")
        End Function

        ''' <summary>
        ''' Obtiene la ruta del directorio que contiene al fichero que se le pasa.
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa de donde se encuentra el fichero</param>
        ''' <returns>Directorio donde se encuentra el fichero</returns>
        Public Function extraerRutaFichero(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetDirectoryName(eRutaCompleta)
        End Function

        ''' <summary>
        ''' Permite realizar una concatenación de dos ficheros, añadiendo el fichero
        ''' de origen al final del fichero de destino
        ''' </summary>
        ''' <param name="eFicheroOrigen">Ruta del archivo de origen (se concatena al final del destino)</param>
        ''' <param name="eFicheroDestino">Ruta del archivo de destino</param>
        ''' <param name="eCodificacion">Codificación utilizada para la lectura y escritura</param>
        ''' <param name="eEliminarOrigen">Determina si se tiene que eliminar el archivo de origen una vez concatenado</param>
        ''' <param name="eEliminarLineasBlanco">A la hora de realizar la concanetación elimina las lineas en blanco</param>
        ''' <remarks></remarks>
        Public Function concatenarFicheros(ByVal eFicheroOrigen As String, _
                                           ByVal eFicheroDestino As String, _
                                           ByVal eCodificacion As System.Text.Encoding, _
                                           Optional ByVal eEliminarOrigen As Boolean = False, _
                                           Optional ByVal eEliminarLineasBlanco As Boolean = True) As Boolean
            ' Se lee todo el contenido del fichero de entrada con la codificación especificada,
            ' Si se produciera un error no se puede continuar y se devuelve el error
            Dim elContenido As String = ""
            Try
                Dim elLector As New System.IO.StreamReader(eFicheroOrigen, eCodificacion)
                elContenido = elLector.ReadToEnd
                elLector.Close()
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error al concatenar fichero " & eFicheroOrigen & " en " & eFicheroDestino, ex, New StackTrace(0, True))
                Return False
            End Try

            ' Se separan todas las lineas que componen el fichero de entrada para analizar
            ' linea a linea ya que se pueden eliminar las lineas en blanco
            Dim lasCadenas As String() = elContenido.Trim.Split(vbCrLf, Chr(10), Chr(13))
            Try
                If lasCadenas.Length > 0 Then
                    Dim elEscritor As New System.IO.StreamWriter(eFicheroDestino, True, System.Text.Encoding.UTF8)
                    For Each S As String In lasCadenas
                        If (eEliminarLineasBlanco And Not String.IsNullOrEmpty(S)) Or (Not eEliminarLineasBlanco) Then
                            elEscritor.WriteLine(S.Trim(vbCrLf, Chr(10), Chr(13)))
                        End If
                    Next
                    elEscritor.Close()
                End If
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error al concatenar...", ex, New StackTrace(0, True))
                Return False
            End Try

            ' Si el código llega hasta este punto la concatenación se realizó correctamente,
            ' por lo que se elimina el archivo de origien si se configuró así
            If eEliminarOrigen Then
                Try
                    System.IO.File.Delete(eFicheroOrigen)
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Error al borrar el fichero de origen de concatenacion : " & eFicheroOrigen, ex, New StackTrace(0, True))
                End Try
            End If

            Return True
        End Function

        ''' <summary>
        ''' Determina si un fichero está en uso
        ''' </summary>
        ''' <param name="eRuta">Ruta del archivo del que se quiere comprobar si está en uso o no</param>
        ''' <returns>True o False indicando si está en uso</returns>
        Public Function estaEnUso(ByVal eRuta As String) As Boolean
            Dim paraDevolver As Boolean = False

            If IO.File.Exists(eRuta) Then
                Try
                    ' Se trata de abrir el archivo, si se produce un error el archivo está en uso
                    Dim elFileStream As New IO.FileStream(eRuta, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None)
                    elFileStream.Close()
                    elFileStream.Dispose()
                    elFileStream = Nothing
                Catch ex As Exception
                    paraDevolver = True
                End Try
            End If

            Return paraDevolver
        End Function
#End Region

#Region " METODOS CARPETAS "
        ''' <summary>
        ''' Obtiene un array de strings con cada uno de los directorios que componen la ruta completa
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa</param>
        ''' <returns>Cada uno de los directorios que componen la ruta completa</returns>
        Public Function splitDirectorios(ByVal eRutaCompleta As String) As String()
            Dim lasPartes() As String = eRutaCompleta.Split("\")
            Dim paraDevolver As New List(Of String)

            If eRutaCompleta.StartsWith("\\") Then
                For i As Integer = 1 To lasPartes.Length - 1
                    paraDevolver.Add(lasPartes(i))
                Next
            Else
                For i As Integer = 0 To lasPartes.Length - 1
                    paraDevolver.Add(lasPartes(i))
                Next
            End If

            Return paraDevolver.ToArray
        End Function

        ''' <summary>
        ''' Obtiene la carpeta desde donde fué lanzada la aplicación.
        ''' Similar a My.Application.StartUpPath pero utilizando Reflectión, lo que
        ''' permite obtener las rutas desde DLL
        ''' </summary>
        ''' <returns>Carpeta desde donde fué lanzada la aplicación.</returns>
        Public Function StartUpPath() As String
            Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        End Function
#End Region

#Region " TEMPORALES "
        ''' <summary>
        ''' Obtiene la ruta donde Windows almacena los archivos temporales
        ''' </summary>
        ''' <returns>Ruta donde windows genera los archivos temporales</returns>
        Public Function RutaDirectorioTemporalWindows() As String
            Dim ArchivoTemporal As String = Microsoft.VisualBasic.FileIO.FileSystem.GetTempFileName()
            Return extraerRutaFichero(ArchivoTemporal)
        End Function

        ''' <summary>
        ''' Obtiene un fichero temporal para poder trabajar con el.
        ''' Se le añade un prefijo al archivo temporal y el formato yyyyMMddHHmmss para poder realizar
        ''' operaciones de borrado mediante otros métodos
        ''' </summary>
        ''' <param name="ePrefijo">Prefijo a añadir al archivo temporal</param>
        ''' <returns>Ruta al archivo temporal</returns>
        Public Function obtenerFicheroTemporal(Optional ByVal ePrefijo As String = "_") As String
            Dim directorioTemp As String = My.Computer.FileSystem.SpecialDirectories.Temp
            Dim nombreFichero As String = String.Empty

            ' Se busca un archivo temporal que no exista en la carpeta
            Do
                nombreFichero = directorioTemp & "\" & ePrefijo & "rec_" & Format(Now, "yyyyMMddHHmmss") & "_" & Aleatorios.cadenaAleatoria(6, True) & ".tmp"
            Loop While IO.File.Exists(nombreFichero)

            Return nombreFichero
        End Function

        ''' <summary>
        ''' Obtiene un directorio temporal en carpeta de temporales
        ''' </summary>
        ''' <param name="eCrear">Crea el directorio temporal</param>
        ''' <returns>Ruta al nuevo directorio temporal</returns>
        Public Function ObtenerDirectorioTemporal(Optional ByVal eCrear As Boolean = True) As String
            Dim paraDevolver As String = ""
            Dim hayError As Boolean = True

            While hayError
                Do
                    paraDevolver = RutaDirectorioTemporalWindows() & "\" & Aleatorios.cadenaAleatoria(8, True)
                Loop While IO.Directory.Exists(paraDevolver)

                ' Una vez generada la ruta no hay error a no ser que se 
                ' tenga que crear el direcotrio y este no se pueda crear
                hayError = False

                If eCrear Then
                    Try
                        IO.Directory.CreateDirectory(paraDevolver)
                    Catch ex As Exception
                        If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de crear und irectorio temporal en " & paraDevolver, ex, New StackTrace(0, True))
                        hayError = True
                    End Try
                End If
            End While

            Return paraDevolver
        End Function
#End Region

#Region " TAMAÑOS "
        ''' <summary>
        ''' Obtiene el tamaño de un archivo en Bytes
        ''' </summary>
        ''' <param name="eRuta">Ruta del archivo</param>
        ''' <returns>Tamaño en Bytes</returns>
        Public Function TamanhoFicheroBytes(ByVal eRuta As String) As Double
            If IO.File.Exists(eRuta) Then
                Try
                    Return (New IO.FileInfo(eRuta).Length)
                Catch ex As Exception
                    Return -1
                End Try
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' Obtiene el tamaño de un archivo en Kb
        ''' </summary>
        ''' <param name="eRuta">Ruta del archivo</param>
        ''' <returns>Tamaño en Kb</returns>
        Public Function TamanhoFicheroKB(ByVal eRuta As String) As Double
            If IO.File.Exists(eRuta) Then
                Try
                    Return (TamanhoFicheroBytes(eRuta) / 1024)
                Catch ex As Exception
                    Return -1
                End Try
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' Obtiene el tamaño de un fichero en Mb
        ''' </summary>
        ''' <param name="eRuta">Ruta al fichero</param>
        ''' <returns>Tamaño del fichero en Mb</returns>
        Public Function TamanhoFicheroMB(ByVal eRuta As String) As Double
            If IO.File.Exists(eRuta) Then
                Try
                    Return (TamanhoFicheroKB(eRuta) / 1024)
                Catch ex As Exception
                    Return -1
                End Try
            Else
                Return -1
            End If
        End Function
#End Region

#Region " CONVERSIONES "
        ''' <summary>
        ''' Convierte un archivo a un array de bytes
        ''' </summary>
        ''' <param name="eRutaFichero">Ruta del fichero que se quiere convertir</param>
        ''' <returns>Fichero como array de bytes o Nothing si no existe</returns>
        Public Function Archivo2Byte(ByVal eRutaFichero As String) As Byte()
            If IO.File.Exists(eRutaFichero) Then
                Try
                    Dim bteRead() As Byte
                    Dim bteArray(256) As Byte
                    Dim myMemortStream As MemoryStream
                    Dim myFileStream As FileStream = New FileStream(eRutaFichero, FileMode.Open, FileAccess.Read)
                    ReDim bteRead(myFileStream.Length)
                    myFileStream.Read(bteRead, 0, myFileStream.Length)
                    myMemortStream = New MemoryStream(myFileStream.Length)
                    myMemortStream.Write(bteRead, 0, myFileStream.Length)
                    Return (myMemortStream.ToArray)
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Error al convertir el archivo '" & eRutaFichero & "' en un aray de bytes...", ex, New StackTrace(0, True))
                    Return Nothing
                End Try
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Convierte un array de btyres en un fichero
        ''' </summary>
        ''' <param name="eByte">Array de bytes a convertir en fichero</param>
        ''' <param name="eFicheroDestino">Ruta del fichero de salida para el guardado</param>
        ''' <returns>True o false dependiendo de la correcta ejecución</returns>
        Public Function Byte2Archivo(ByVal eByte As Object, _
                                     ByVal eFicheroDestino As String) As Boolean
            Dim ParaDevolver As Boolean = False

            Try
                Dim myFileStream As FileStream = New FileStream(eFicheroDestino, FileMode.OpenOrCreate, FileAccess.Write)
                myFileStream.Write(eByte, 0, eByte.Length)
                myFileStream.Close()

                ParaDevolver = IO.File.Exists(eFicheroDestino)
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error al guardar un array de bytes en el fichero '" & eFicheroDestino & "'...", ex, New StackTrace(0, True))
                ParaDevolver = False
            End Try

            Return ParaDevolver
        End Function
#End Region
    End Module
End Namespace
