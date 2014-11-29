Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Drawing
Imports System.Reflection
Imports ComponentFactory.Krypton.Toolkit

Namespace Ficheros
    Public Module modFicheros

        ''' <summary>
        ''' Limpia el nombre del archivo de caracteres no permitidos para una exportación
        ''' </summary>
        Public Function LimpiarNombreArchivoExportacion(ByVal eNombreNormal) As String
            Return eNombreNormal.Replace("\", "-").Replace(":", "").Replace("/", "-").Replace("*", "").Replace("?", "").Replace("<", "").Replace(">", "").Replace("|", "")
        End Function

        Public Function LimpiarNombreArchivo(ByVal eNombreNomrmal As String) As String
            Dim NombreNuevo As String = eNombreNomrmal

            NombreNuevo = Replace(NombreNuevo, "'", "")
            NombreNuevo = Replace(NombreNuevo, "?", "")
            NombreNuevo = Replace(NombreNuevo, "[", "")
            NombreNuevo = Replace(NombreNuevo, "]", "")
            NombreNuevo = Replace(NombreNuevo, "/", "")
            NombreNuevo = Replace(NombreNuevo, "\", "")
            NombreNuevo = Replace(NombreNuevo, ":", "")
            NombreNuevo = Replace(NombreNuevo, "*", "")
            NombreNuevo = Replace(NombreNuevo, """", "")
            NombreNuevo = Replace(NombreNuevo, "|", "")
            NombreNuevo = Replace(NombreNuevo, "<", "")
            NombreNuevo = Replace(NombreNuevo, ">", "")
            NombreNuevo = Replace(NombreNuevo, "´", "")

            Return NombreNuevo
        End Function

        ''' <summary>
        ''' Obtiene un nombre para un documento evitando que este sea igual que otro 
        ''' en que exista en el mismo directorio
        ''' </summary>
        ''' <param name="eCarpetaDestino">Carpeta donde se va a comprobar el nombre</param>
        ''' <param name="eNombreDestino">Nombre del documento</param>
        ''' <param name="eExtensionDestinoConPunto">Extensión del documento (Con punto)</param>
        ''' <returns>Devuelve el nombre del documento para que no se repita (OJO, SIN EXTENSIÓN)</returns>
        ''' <remarks>Se devuelve solamente el nombre del documento, no el documento con la extensión</remarks>
        Public Function obtenerNombreConsecutivo(ByVal eCarpetaDestino As String, _
                                                 ByVal eNombreDestino As String, _
                                                 ByVal eExtensionDestinoConPunto As String) As String
            ' Se genera un nombre consecutivo al documento por si ya existe el nombre
            ' en el directorio de destino
            Dim Aux As Long = 0
            Dim Sufijo As String = String.Empty
            Dim nuevoNombre As String = eNombreDestino
            If IO.File.Exists(eCarpetaDestino & nuevoNombre & eExtensionDestinoConPunto) Then
                Do
                    Aux += 1
                    Sufijo = Format(Aux, "0000")
                    nuevoNombre = eNombreDestino & "_" & Sufijo
                Loop While IO.File.Exists(eCarpetaDestino & nuevoNombre & eExtensionDestinoConPunto)
            End If

            Return nuevoNombre
        End Function

        ''' <summary>
        ''' Extrae el nombre y la extensión del fichero de la ruta que se le pasa a la función.
        ''' </summary>
        ''' <param name="rutaCompleta">Ruta completa del fichero de donde extraer el nombre</param>
        ''' <returns>Nombre y extensión del fichero</returns>
        ''' <remarks></remarks>
        Public Function extraerNombreFichero(ByVal rutaCompleta As String) As String
            Return System.IO.Path.GetFileName(rutaCompleta)
        End Function

        Public Function extraerNombreFicheroSinExtension(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetFileNameWithoutExtension(eRutaCompleta)
        End Function

        ''' <summary>
        ''' Extrae la extensión del fichero que se le pasa como parametro.
        ''' </summary>
        ''' <param name="eRutaCompleta">Ruta completa de la ubicación del fichero</param>
        ''' <returns>Extensión del fichero</returns>
        ''' <remarks></remarks>
        Public Function extraerExtensionFicheroConPunto(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetExtension(eRutaCompleta)
        End Function

        Public Function extraerExtensionFicheroSinPunto(ByVal eRutaCompleta As String) As String
            Return System.IO.Path.GetExtension(eRutaCompleta).Replace(".", "")
        End Function


        ''' <summary>
        ''' Obtiene la ruta del directorio que contiene al fichero que se le pasa.
        ''' </summary>
        ''' <param name="rutaCompleta">Ruta completa de donde se encuentra el fichero</param>
        ''' <returns>Directorio donde se encuentra el fichero</returns>
        ''' <remarks></remarks>
        Public Function extraerRutaFichero(ByVal rutaCompleta As String) As String
            Return System.IO.Path.GetDirectoryName(rutaCompleta)
        End Function

        ''' <summary>
        ''' Obtiene un array de strings con cada uno de los directorios que componen la ruta completa
        ''' </summary>
        ''' <param name="rutaCompleta">Ruta completa</param>
        ''' <returns>Cada uno de los directorios que componen la ruta completa</returns>
        Public Function obtenerDirectorios(ByVal rutaCompleta As String) As String()
            Dim lasPartes() As String = rutaCompleta.Split("\")
            Dim paraDevolver() As String = {}
            If rutaCompleta.StartsWith("\\") Then
                For i As Integer = 1 To lasPartes.Length - 1
                    paraDevolver(i - 1) = lasPartes(i)
                Next
            Else
                For i As Integer = 0 To lasPartes.Length - 1
                    paraDevolver(i) = lasPartes(i)
                Next
            End If

            Return paraDevolver
        End Function

        Public Sub concatenarFicheros(ByVal ficheroOrigen As String, _
                                       ByVal ficheroDestino As String)

            Dim elContenido As String = ""
            Try
                Dim elLector As New System.IO.StreamReader(ficheroOrigen, System.Text.Encoding.UTF8)
                elContenido = elLector.ReadToEnd
                elLector.Close()
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error al concatenar fichero " & ficheroOrigen & " en " & ficheroDestino, ex, New StackTrace(0, True))
            End Try

            Try
                System.IO.File.Delete(ficheroOrigen)
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error al borrar el fichero de origen de concatenacion : " & ficheroOrigen, ex, New StackTrace(0, True))
            End Try

            Dim lasCadenas As String() = elContenido.Trim.Split(vbCrLf, Chr(10), Chr(13))

            Try
                If lasCadenas.Length > 0 Then
                    Dim elEscritor As New System.IO.StreamWriter(ficheroDestino, True, System.Text.Encoding.UTF8)
                    For Each S As String In lasCadenas
                        If S.Trim.Length > 1 Then
                            elEscritor.WriteLine(S.Trim(vbCrLf, Chr(10), Chr(13)))
                        End If
                    Next
                    elEscritor.Close()
                End If
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Error al concatenar...", ex, New StackTrace(0, True))
            End Try
        End Sub

        ''' <summary>
        ''' Obtiene un fichero temporal para poder trabajar con el.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function obtenerFicheroTemporal(Optional ByVal ePrefijo As String = "_") As String
            Dim directorioTemp As String = My.Computer.FileSystem.SpecialDirectories.Temp
            Dim nombreFichero As String = String.Empty
            Do
                nombreFichero = directorioTemp & "\" & ePrefijo & "dfs_" & Format(Now, "yyyyMMddHHmmss") & "_" & Aleatorios.CadenaAleatoria(6, True) & ".tmp"
            Loop While IO.File.Exists(nombreFichero)
            Return nombreFichero

            'Dim rutaWindows As String = Microsoft.VisualBasic.FileIO.FileSystem.GetTempFileName()
            'Dim nombreWindows As String = extraerNombreFichero(rutaWindows)
            'Dim directorioWindows As String = extraerRutaFichero(rutaWindows)
            'Dim temporal As String = directorioWindows & "\" & ePrefijo & "_" & Format(Now, "yyyyMMddHHmmss") & "_" & nombreWindows

            'Return temporal
        End Function

        Public Function ObtenerDirectorioTemporal() As String
            Dim RutaObtenida As String = ""
            Do
                RutaObtenida = RutaDirectorioTemporalWindows() & "\" & Aleatorios.CadenaAleatoria(8, True)
            Loop While IO.Directory.Exists(RutaObtenida)

            IO.Directory.CreateDirectory(RutaObtenida)
            Return RutaObtenida
        End Function

        Public Function RutaDirectorioTemporalWindows() As String
            Dim ArchivoTemporal As String = Microsoft.VisualBasic.FileIO.FileSystem.GetTempFileName()
            Return extraerRutaFichero(ArchivoTemporal)
        End Function

        Public Function buscarImagen(Optional ByVal eIdioma As Traductor.enmLenguajesCodigos = Traductor.enmLenguajesCodigos.es_ES, _
                                     Optional ByVal eRutaInicial As String = "") As String
            ' MSG 1 : Archivos de imagen
            Dim lasExtensiones As String = theMessages(eIdioma)(1) & ": (*.jpg, *.jpeg, *.gif, *.png, *.bmp, *.tif, *.tiff|*.jpg;*.jpeg;*.gif;*.png;*.bmp;*.tif;*.tiff|*.jpg|*.jpg|*.gif|*.gif|*.jpeg|*.jpeg|*.png|*.png|*.bmp|*.bmp|*.tif|*.tif|*.tiff|*.tiff"
            Return buscarArchivo(eRutaInicial, lasExtensiones, "")
        End Function

        Public Function buscarArchivo(ByVal _rutaInicial As String, _
                                      ByVal _extensionesArchivo As String, _
                                      ByVal _nombreArchivo As String) As String
            Dim paraDevolver As String = ""
            Try
                Dim dlgAbrirArchivo As New System.Windows.Forms.OpenFileDialog

                If _rutaInicial <> "" AndAlso System.IO.Directory.Exists(_rutaInicial) Then dlgAbrirArchivo.InitialDirectory = _rutaInicial
                If _extensionesArchivo <> "" Then dlgAbrirArchivo.Filter = _extensionesArchivo
                If _nombreArchivo <> "" Then dlgAbrirArchivo.FileName = _nombreArchivo


                dlgAbrirArchivo.ShowDialog()

                If System.IO.File.Exists(dlgAbrirArchivo.FileName) Then paraDevolver = dlgAbrirArchivo.FileName
            Catch
                paraDevolver = ""
            End Try

            Return paraDevolver
        End Function

        Public Function BuscarArchivoGuardar(ByVal eRutaInicial As String, _
                                             ByVal eExtensionArchivo As String, _
                                             ByVal eNombreArchvio As String) As String
            Dim ParaDevolver As String = ""
            Try
                Dim dlgGuardarArchivo As New System.Windows.Forms.SaveFileDialog
                If eRutaInicial <> "" AndAlso System.IO.Directory.Exists(eRutaInicial) Then dlgGuardarArchivo.InitialDirectory = eRutaInicial
                If eExtensionArchivo <> "" Then dlgGuardarArchivo.Filter = eExtensionArchivo
                If eNombreArchvio <> "" Then dlgGuardarArchivo.FileName = eNombreArchvio

                dlgGuardarArchivo.ShowDialog()

                ParaDevolver = dlgGuardarArchivo.FileName
            Catch ex As Exception
                ParaDevolver = ""
            End Try

            Return ParaDevolver

        End Function

        Public Function buscarDirectorio(ByVal _rutaIncial As String) As String
            Dim paraDevolver As String = ""
            Try
                Dim dlAbrirDirectorio As New FolderBrowserDialog

                dlAbrirDirectorio.SelectedPath = _rutaIncial

                dlAbrirDirectorio.ShowDialog()

                If System.IO.Directory.Exists(dlAbrirDirectorio.SelectedPath) Then paraDevolver = dlAbrirDirectorio.SelectedPath
            Catch ex As System.Exception
                paraDevolver = ""
            End Try

            Return paraDevolver
        End Function

        Public Function buscarDirectorioRed(ByVal oFolderBrowserDialog As FolderBrowserDialog, _
                                            ByVal eRutaInicial As String) As String
            Dim type As Type = oFolderBrowserDialog.[GetType]
            Dim fieldInfo As Reflection.FieldInfo = type.GetField("rootFolder", BindingFlags.NonPublic Or BindingFlags.Instance)
            fieldInfo.SetValue(oFolderBrowserDialog, DirectCast(18, Environment.SpecialFolder))
            If IO.Directory.Exists(eRutaInicial) Then oFolderBrowserDialog.SelectedPath = eRutaInicial
            If oFolderBrowserDialog.ShowDialog() = DialogResult.OK Then
                Return oFolderBrowserDialog.SelectedPath
            Else
                Return ""
            End If
        End Function

        Public Function buscarDirectorioLocal(ByVal oFolderBrowserDialog As FolderBrowserDialog, _
                                              ByVal eRutaInicial As String) As String
            Dim type As Type = oFolderBrowserDialog.[GetType]
            Dim fieldInfo As Reflection.FieldInfo = type.GetField("rootFolder", BindingFlags.NonPublic Or BindingFlags.Instance)
            fieldInfo.SetValue(oFolderBrowserDialog, DirectCast(17, Environment.SpecialFolder))
            If IO.Directory.Exists(eRutaInicial) Then oFolderBrowserDialog.SelectedPath = eRutaInicial
            If oFolderBrowserDialog.ShowDialog() = DialogResult.OK Then
                Return oFolderBrowserDialog.SelectedPath
            Else
                Return ""
            End If
        End Function

        Public Function StartUpPath() As String
            Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        End Function

        ''' <summary>
        ''' Copia un archivo de origen en destino. 
        ''' Se puede especificar que se cree la estructura de directorios de destino
        ''' si esta no existe
        ''' </summary>
        ''' <param name="eOrigen">Archivo d eorigen que quiere ser copiado</param>
        ''' <param name="eDestino">Archiov de destino que quiere ser copiado</param>
        ''' <param name="eSobreescribir">Si se va a sobreescribir el destino o no</param>
        ''' <param name="eCrearRuta">Si se tiene que crear la ruta d edestino si no existe</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function copiarArchivo(ByVal eOrigen As String, _
                                      ByVal eDestino As String, _
                                      ByVal eSobreescribir As Boolean, _
                                      ByVal eCrearRuta As Boolean)

            eDestino.Replace("/", "\")
            If eDestino.StartsWith("\\") Then Throw New Exception("No se pueden hacer copias en rutas UNC.")

            If eCrearRuta Then
                Dim lasPartes() As String = eDestino.Split("\")
                Dim laUnidad As String = lasPartes(0) & "\"
                Dim rutaParcial As String = laUnidad

                For i As Integer = 1 To lasPartes.Length - 2
                    rutaParcial &= lasPartes(i) & "\"
                    If Not IO.Directory.Exists(rutaParcial) Then IO.Directory.CreateDirectory(rutaParcial)
                Next
            End If

            IO.File.Copy(eOrigen, eDestino, eSobreescribir)

            Return True
        End Function

        ''' <summary>
        ''' Copia todos los archivos y directorios del Origen al Destino de manera recursiva
        ''' </summary>
        ''' <param name="diSource"></param>
        ''' <param name="diDestination"></param>
        ''' <param name="blOverwrite"></param>
        ''' <remarks></remarks>
        Public Sub copiarArchivos(ByVal diSource As DirectoryInfo, _
                                  ByVal diDestination As DirectoryInfo, _
                                  ByVal blOverwrite As Boolean, _
                                  ByVal eLabelInfo As Object, _
                                  ByVal eBarra As Object)

            Dim diSourceSubDirectories() As DirectoryInfo
            Dim fiSourceFiles() As FileInfo

            'obtengo todos los archivos del directorio origen
            fiSourceFiles = diSource.GetFiles()
            'obtengo los subdirectorios (si existen)
            diSourceSubDirectories = diSource.GetDirectories()

            'si no existe el directorio destino crearlo
            If Not diDestination.Exists Then diDestination.Create()

            'Usar la recursividad para navegar por los subdirectorios
            'e ir obteniendo los archivos hasta llegar al final
            For Each diSourceSubDirectory As DirectoryInfo In diSourceSubDirectories
                copiarArchivos(diSourceSubDirectory, New DirectoryInfo(diDestination.FullName & "\" & diSourceSubDirectory.Name), blOverwrite, eLabelInfo, eBarra)
            Next

            Try
                If eBarra IsNot Nothing Then WinForms.ProgressBar.fijarMaximoBarra(eBarra, fiSourceFiles.Length)
            Catch ex As Exception
            End Try

            For Each fiSourceFile As FileInfo In fiSourceFiles
                Dim archivoDestino As String = diDestination.FullName + "\" + fiSourceFile.Name

                Try
                    If eLabelInfo IsNot Nothing Then eLabelInfo.Text = archivoDestino
                    If eBarra IsNot Nothing Then WinForms.ProgressBar.AumentarBarra(eBarra)
                Catch ex As Exception
                End Try

                If Log._LOG_ACTIVO Then Log.escribirLog(fiSourceFile.FullName & " -> " & archivoDestino, , New StackTrace(0, True))

                If System.IO.File.Exists(archivoDestino) Then
                    Try
                        System.IO.File.Delete(archivoDestino)
                    Catch ex As Exception
                        If Log._LOG_ACTIVO Then Log.escribirLog("No se pudo borrar el archivo de destino '" & archivoDestino & "' (Forzando Kill Proceso)", , New StackTrace(0, True))
                        Try
                            Dim nombreProceso As String = Ficheros.extraerNombreFichero(fiSourceFile.Name)
                            Try
                                If eLabelInfo IsNot Nothing Then eLabelInfo.Text = "Deteniendo el proceso " & nombreProceso
                            Catch exProceso As Exception

                            End Try
                            Procesos.mataProcesoWhile(nombreProceso, 3)
                            System.IO.File.Delete(archivoDestino)
                        Catch
                            If Log._LOG_ACTIVO Then Log.escribirLog("ERROR eliminando archivo (incluso con KILL) '" & archivoDestino & "'", ex, New StackTrace(0, True))
                        End Try
                    End Try
                End If

                Try
                    fiSourceFile.CopyTo(diDestination.FullName + "\" + fiSourceFile.Name, blOverwrite)
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("ERRROR al copiar " & fiSourceFile.FullName & " -> " & archivoDestino, ex, New StackTrace(0, True))
                End Try
            Next
        End Sub

        Public Function TamanhoFicheroKB(ByVal eRuta As String) As Double
            If IO.File.Exists(eRuta) Then
                Try
                    Return ((New IO.FileInfo(eRuta).Length) / 1024)
                Catch ex As Exception
                    Return -1
                End Try
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' Convierte un archivo a un array de bytes
        ''' </summary>
        Public Function Archivo2Byte(ByVal eRutaFichero) As Byte()
            If IO.File.Exists(eRutaFichero) Then
                Dim bteRead() As Byte
                Dim bteArray(256) As Byte
                Dim myMemortStream As MemoryStream
                Dim myFileStream As FileStream = New FileStream(eRutaFichero, FileMode.Open, FileAccess.Read)
                ReDim bteRead(myFileStream.Length)
                myFileStream.Read(bteRead, 0, myFileStream.Length)
                myMemortStream = New MemoryStream(myFileStream.Length)
                myMemortStream.Write(bteRead, 0, myFileStream.Length)
                Return (myMemortStream.ToArray)
            Else
                Return Nothing
            End If
        End Function

        Public Function Byte2Archivo(ByVal eByte As Object, ByVal eRutaFichero As String) As Boolean
            Dim ParaDevolver As Boolean = False
            Try
                Dim myFileStream As FileStream = New FileStream(eRutaFichero, FileMode.OpenOrCreate, FileAccess.Write)
                myFileStream.Write(eByte, 0, eByte.Length)
                myFileStream.Close()
                ParaDevolver = True
            Catch ex As Exception
                ParaDevolver = False
            End Try
            Return ParaDevolver
        End Function

        <DllImport("urlmon.dll", CharSet:=CharSet.Auto)> _
        Private Function FindMimeFromData(ByVal pBC As System.UInt32, <MarshalAs(UnmanagedType.LPStr)> ByVal pwzUrl As System.String, <MarshalAs(UnmanagedType.LPArray)> ByVal pBuffer As Byte(), ByVal cbSize As System.UInt32, <MarshalAs(UnmanagedType.LPStr)> ByVal pwzMimeProposed As System.String, ByVal dwMimeFlags As System.UInt32, _
                                              ByRef ppwzMimeOut As System.UInt32, ByVal dwReserverd As System.UInt32) As System.UInt32
        End Function

        ''' <summary>
        ''' Obtiene el tipo de archivo correcto en función del contenido del mismo
        ''' </summary>
        ''' <param name="filename">Ruta del archivo que deseamos analizar</param>
        Public Function ObtenerTipoArchivo(ByVal filename As String) As String
            If Not File.Exists(filename) Then
                Throw New FileNotFoundException("Archivo " & filename & " no encontrado")
                Return String.Empty
            Else
                Dim buffer As Byte() = New Byte(255) {}
                Using fs As New FileStream(filename, FileMode.Open)
                    If fs.Length >= 256 Then
                        fs.Read(buffer, 0, 256)
                    Else
                        fs.Read(buffer, 0, CInt(fs.Length))
                    End If
                End Using

                Return ObtenerTipoArchivo(buffer)
            End If
        End Function

        ''' <summary>
        ''' Obtiene el tipo de archivo correcto en función del contenido del mismo
        ''' </summary>
        ''' <param name="Bytes">Contenido del archivo</param>
        Public Function ObtenerTipoArchivo(ByVal Bytes As Byte()) As String
            If Bytes Is Nothing OrElse Bytes.Length = 0 Then Return String.Empty

            ' Tomo los 256 primeros caracteres del archivo
            Dim buffer As Byte() = New Byte(255) {}
            Dim Longitud As Integer = 256

            If Bytes.Length < 256 Then
                Longitud = Bytes.Length
            End If

            For i As Integer = 0 To Longitud - 1
                buffer(i) = Bytes(i)
            Next

            Try
                Dim mimetype As System.UInt32
                FindMimeFromData(0, Nothing, buffer, 256, Nothing, 0, mimetype, 0)
                Dim mimeTypePtr As System.IntPtr = New IntPtr(mimetype)
                Dim mime As String = Marshal.PtrToStringUni(mimeTypePtr)
                Marshal.FreeCoTaskMem(mimeTypePtr)
                Return mime
            Catch e As Exception
                Return "unknown/unknown"
            End Try
        End Function

        ''' <summary>
        ''' Obtiene la extensión por defecto del tipo mime especificado
        ''' </summary>
        ''' <param name="mimeType"></param>
        Public Function ExtensionPorDefecto(ByVal mimeType As String) As String
            Dim result As String
            Dim key As RegistryKey
            Dim value As Object

            Try
                key = Registry.ClassesRoot.OpenSubKey("MIME\Database\Content Type\" & mimeType, False)
                value = If(key IsNot Nothing, key.GetValue("Extension", Nothing), Nothing)
                result = If(value IsNot Nothing, value.ToString(), String.Empty)
            Catch ex As Exception
                result = String.Empty
            End Try

            Return result
        End Function

        Public Function estaEnUso(ByVal eRuta As String) As Boolean
            If IO.File.Exists(eRuta) Then
                Try
                    Dim myfilestream As New IO.FileStream(eRuta, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None)
                    myfilestream.Close()
                    myfilestream.Dispose()
                    myfilestream = Nothing
                    Return False
                Catch h As Exception
                    Return True
                End Try
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Obtiene la imagen (icono) asociado a un archivo
        ''' </summary>
        ''' <param name="eRuta">Ruda del archivo del que se quiere obetner la imagen</param>
        ''' <returns>Imagen asociada al archivo</returns>
        Public Function obtenerIconoArchivo(ByVal eRuta As String) As Icon
            If Not IO.File.Exists(eRuta) Then Return Nothing
            Dim paraDevolver As Icon = Nothing

            Try
                Dim laExtension As String = extraerExtensionFicheroSinPunto(eRuta)
                If laExtension <> "" Then
                    Dim rutaEjecutable As String = obtenerProgramaAsociado(laExtension)
                    paraDevolver = Drawing.Icon.ExtractAssociatedIcon(rutaEjecutable)
                End If
            Catch ex As Exception

            End Try

            Return paraDevolver
        End Function

        Public Function obtenerProgramaAsociado(ByVal eExtension As String) As String


            ' Returns the application associated with the specified
            ' FileExtension
            ' ie, path\denenv.exe for "VB" files
            Dim objExtReg As Microsoft.Win32.RegistryKey = _
                 Microsoft.Win32.Registry.ClassesRoot
            Dim objAppReg As Microsoft.Win32.RegistryKey = _
                Microsoft.Win32.Registry.ClassesRoot
            Dim strExtValue As String
            Try
                ' Add trailing period if doesn't exist
                If eExtension.Substring(0, 1) <> "." Then _
                    eExtension = "." & eExtension
                ' Open registry areas containing launching app details
                objExtReg = objExtReg.OpenSubKey(eExtension.Trim)
                strExtValue = objExtReg.GetValue("").ToString
                objAppReg = objAppReg.OpenSubKey(strExtValue & _
                                "\shell\open\command")
                ' Parse out, tidy up and return result
                Dim SplitArray() As String
                SplitArray = Split(objAppReg.GetValue(Nothing).ToString, """")
                If SplitArray(0).Trim.Length > 0 Then
                    Return SplitArray(0).Replace("%1", "")
                Else
                    Return SplitArray(1).Replace("%1", "")
                End If
            Catch
                Return ""
            End Try
        End Function
    End Module
End Namespace
