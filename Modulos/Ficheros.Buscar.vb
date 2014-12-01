Imports System.Windows.Forms
Imports System.Reflection

Namespace Ficheros
    Namespace Buscar
        Public Module modRecompilaFicherosBuscar
#Region " PROPIEDADES "
            ''' <summary>
            ''' Listado de extensiones de archivos del tipo imagen
            ''' </summary>
            Public Property _EXTENSIONES_IMAGEN As New List(Of String) From {"jpg", "jpeg", "gif", "png", "bmp", "tif", "tiff"}

            ''' <summary>
            ''' Singleton para la apertura de archivos
            ''' </summary>
            Public Property _DIALOGO_APERTURA As System.Windows.Forms.OpenFileDialog
                Get
                    If i_DIALOGO_APERTURA Is Nothing Then
                        i_DIALOGO_APERTURA = New System.Windows.Forms.OpenFileDialog
                    End If
                    Return i_DIALOGO_APERTURA
                End Get
                Set(value As System.Windows.Forms.OpenFileDialog)
                    i_DIALOGO_APERTURA = value
                End Set
            End Property
            Private i_DIALOGO_APERTURA As System.Windows.Forms.OpenFileDialog = Nothing

            ''' <summary>
            ''' Singleton para el guardado de archivos
            ''' </summary>
            Public Property _DIALOGO_GUARDADO As System.Windows.Forms.SaveFileDialog
                Get
                    If i_DIALOGO_GUARDADO Is Nothing Then
                        i_DIALOGO_GUARDADO = New System.Windows.Forms.SaveFileDialog
                    End If
                    Return i_DIALOGO_GUARDADO
                End Get
                Set(value As System.Windows.Forms.SaveFileDialog)
                    i_DIALOGO_GUARDADO = value
                End Set
            End Property
            Private i_DIALOGO_GUARDADO As System.Windows.Forms.SaveFileDialog = Nothing

            ''' <summary>
            ''' Singleton para la búsqueda de carpetas
            ''' </summary>
            Public Property _DIALOGO_APERTURA_FOLDER As FolderBrowserDialog
                Get
                    If i_DIALOGO_APERTURA_FOLDER Is Nothing Then
                        i_DIALOGO_APERTURA_FOLDER = New FolderBrowserDialog
                    End If
                    Return i_DIALOGO_APERTURA_FOLDER
                End Get
                Set(value As FolderBrowserDialog)
                    i_DIALOGO_APERTURA_FOLDER = value
                End Set
            End Property
            Private i_DIALOGO_APERTURA_FOLDER As FolderBrowserDialog = Nothing
#End Region

#Region " METODOS ARCHIVOS "
            ''' <summary>
            ''' Permite la búsqueda y apertura de un archivo
            ''' </summary>
            ''' <param name="eTitulo">Título a mostrar en el FileOpenDialog</param>
            ''' <param name="eRutaInicial">Ruta inicial donde se tiene que realizar la búsqueda</param>
            ''' <param name="eExtensionesArchivo">Extensiones permitidas en formato *.ex1|*.ex1|*.ex2|*.ex2</param>
            ''' <param name="eNombreArchivo">Nombre inicial del archivo</param>
            ''' <returns>Ruta al archivo que se va a abrir o una cadena vacía</returns>

            Public Function buscarArchivo(ByVal eTitulo As String, _
                                          Optional ByVal eRutaInicial As String = "", _
                                          Optional ByVal eExtensionesArchivo As String = "*.*", _
                                          Optional ByVal eNombreArchivo As String = "") As String

                Return buscarArchivoInterno(eTitulo, eRutaInicial, eExtensionesArchivo, eNombreArchivo)
            End Function


            ''' <summary>
            ''' Permite la búsqueda y apertura de un archivo
            ''' </summary>
            ''' <param name="eTitulo">Título a mostrar en el FileOpenDialog</param>
            ''' <param name="eRutaInicial">Ruta inicial donde se tiene que realizar la búsqueda</param>
            ''' <param name="eExtensionesArchivo">Listado con las extensiones que se pueden abrir</param>
            ''' <param name="eNombreArchivo">Nombre inicial del archivo</param>
            ''' <returns>Ruta al archivo que se va a abrir o una cadena vacía</returns>
            Public Function buscarArchivo(ByVal eTitulo As String, _
                                          Optional ByVal eRutaInicial As String = "", _
                                          Optional ByVal eExtensionesArchivo As List(Of String) = Nothing, _
                                          Optional ByVal eNombreArchivo As String = "") As String
                ' Se crea el filtro para la apertura a partir de las extensiones
                Dim elFiltro As String = "*.*|*.*"
                If eExtensionesArchivo IsNot Nothing AndAlso eExtensionesArchivo.Count > 0 Then
                    Dim filtroA As String = ""
                    Dim filtroB As String = ""

                    ' Se añade un primer filtro con todas las extensiones
                    For Each unaExtension As String In eExtensionesArchivo
                        If Not String.IsNullOrEmpty(unaExtension.Trim) Then
                            filtroA &= "*." & unaExtension.Trim & ", "
                            filtroB &= "*." & unaExtension.Trim & ";"
                        End If
                    Next

                    If filtroA.EndsWith(", ") Then filtroA = filtroA.Substring(0, filtroA.Length - 2)
                    If filtroB.EndsWith(";") Then filtroB = filtroB.Substring(0, filtroB.Length - 1)
                    elFiltro = "(" & filtroA & ")|" & filtroB

                    ' Se añaden cada una de las extensiones por separador
                    For Each unaExtension As String In eExtensionesArchivo
                        If Not String.IsNullOrEmpty(unaExtension.Trim) Then
                            elFiltro &= "|*." & unaExtension & "|*." & unaExtension
                        End If
                    Next
                End If

                ' Se raliza la búsuqeda interna con los parámetros configurados
                Return buscarArchivoInterno(eTitulo, eRutaInicial, elFiltro, eNombreArchivo)
            End Function

            ''' <summary>
            ''' Permite la búsqueda y apertura de un archivo
            ''' </summary>
            ''' <param name="eTitulo">Título a mostrar en el FileOpenDialog</param>
            ''' <param name="eRutaInicial">Ruta inicial donde se tiene que realizar la búsqueda</param>
            ''' <param name="eExtensionesArchivo">Listado con las extensiones que se pueden abrir</param>
            ''' <param name="eNombreArchivo">Nombre inicial del archivo</param>
            ''' <returns>Ruta al archivo que se va a abrir o una cadena vacía</returns>
            Private Function buscarArchivoInterno(ByVal eTitulo As String, _
                                                  ByVal eRutaInicial As String, _
                                                  ByVal eExtensionesArchivo As String, _
                                                  ByVal eNombreArchivo As String) As String
                Dim paraDevolver As String = ""

                ' Se realiza la busqueda del documento configurado el FileOpenDialog
                Try
                    _DIALOGO_APERTURA.Title = eTitulo
                    If Not String.IsNullOrEmpty(eRutaInicial) AndAlso System.IO.Directory.Exists(eRutaInicial) Then
                        _DIALOGO_APERTURA.InitialDirectory = eRutaInicial
                    Else
                        _DIALOGO_APERTURA.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                    End If
                    _DIALOGO_APERTURA.Filter = eExtensionesArchivo
                    If Not String.IsNullOrEmpty(eNombreArchivo) Then
                        _DIALOGO_APERTURA.FileName = eNombreArchivo
                    Else
                        _DIALOGO_APERTURA.FileName = ""
                    End If

                    If _DIALOGO_APERTURA.ShowDialog() = DialogResult.OK Then
                        If System.IO.File.Exists(_DIALOGO_APERTURA.FileName) Then
                            paraDevolver = _DIALOGO_APERTURA.FileName
                        End If
                    End If
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de buscar un archivo.", ex, New StackTrace(0, True))
                    paraDevolver = ""
                End Try

                Return paraDevolver
            End Function


            ''' <summary>
            ''' Crea un FileOpenDialog y obtiene la ruta a la imagen seleccionada por el usuario
            ''' </summary>
            ''' <param name="eTitulo">Título a mostrar en el FileOpenDialog</param>
            ''' <param name="eRutaInicial">Ruta inicial para la apertura de la imagen</param>
            ''' <param name="eExtensiones">Extensiones de imagen que puede seleccionar el usuario</param>
            ''' <param name="eNombreArchivo">Nombre del archivo predeterminado a abrir</param>
            ''' <returns>Ruta a la imagen seleccionada</returns>
            Public Function buscarImagen(Optional ByVal eTitulo As String = "", _
                                         Optional ByVal eRutaInicial As String = "", _
                                         Optional ByVal eExtensiones As List(Of String) = Nothing, _
                                         Optional ByVal eNombreArchivo As String = "") As String
                ' Si no están establecidos los parámetros opcionales se utilizan los predeterminados
                If String.IsNullOrEmpty(eTitulo) Then eTitulo = "Seleccione una imagen"
                If String.IsNullOrEmpty(eRutaInicial) OrElse (Not String.IsNullOrEmpty(eRutaInicial) AndAlso Not IO.Directory.Exists(eRutaInicial)) Then
                    eRutaInicial = My.Computer.FileSystem.SpecialDirectories.MyPictures
                End If
                If eExtensiones Is Nothing OrElse (eExtensiones IsNot Nothing AndAlso eExtensiones.Count = 0) Then
                    eExtensiones = _EXTENSIONES_IMAGEN
                End If

                ' Se llama a buscar archivo con la configuración de búsqueda de imágenes
                Return buscarArchivo(eTitulo, eRutaInicial, eExtensiones, eNombreArchivo)
            End Function

            ''' <summary>
            ''' Abre un FileSaveDialog para localizar una ruta de guardado de un archivo
            ''' </summary>
            ''' <param name="eTitulo">Titulo para el FileSaveDialog</param>
            ''' <param name="eRutaInicial">Ruta inicial donde se tiene que posicionar inicialmente</param>
            ''' <param name="eExtensionArchivo">Extensión predeterminada del archivo</param>
            ''' <param name="eNombreArchvio">Nombre del archivo si existe uno previo</param>
            ''' <returns>Ruta completa al archivo a guardar</returns>
            Public Function BuscarArchivoGuardar(ByVal eTitulo As String, _
                                                 Optional ByVal eRutaInicial As String = "", _
                                                 Optional ByVal eExtensionArchivo As String = "", _
                                                 Optional ByVal eNombreArchvio As String = "") As String
                Dim ParaDevolver As String = ""
                Try
                    _DIALOGO_GUARDADO.Title = eTitulo
                    If Not String.IsNullOrEmpty(eRutaInicial) AndAlso System.IO.Directory.Exists(eRutaInicial) Then
                        _DIALOGO_GUARDADO.InitialDirectory = eRutaInicial
                    Else
                        _DIALOGO_GUARDADO.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                    End If
                    If Not String.IsNullOrEmpty(eExtensionArchivo) Then
                        _DIALOGO_GUARDADO.Filter = eExtensionArchivo
                    Else
                        _DIALOGO_GUARDADO.Filter = "*.*"
                    End If
                    If Not String.IsNullOrEmpty(eNombreArchvio) Then
                        _DIALOGO_GUARDADO.FileName = eNombreArchvio
                    Else
                        _DIALOGO_GUARDADO.FileName = ""
                    End If

                    If _DIALOGO_GUARDADO.ShowDialog() = DialogResult.OK Then
                        ParaDevolver = _DIALOGO_GUARDADO.FileName
                    End If
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de buscar un archivo para guardar.", ex, New StackTrace(0, True))
                    ParaDevolver = ""
                End Try

                Return ParaDevolver

            End Function
#End Region

#Region " METODOS CARPETAS "
            ''' <summary>
            ''' Abre un FolderBrowserDialog para escoger un directorio
            ''' </summary>
            ''' <param name="eRutaIncial">Ruta inicial donde se posicionará FolderBrowserDialog</param>
            ''' <returns>Ruta completa al direcotorio seleccionado</returns>
            Public Function buscarDirectorio(Optional ByVal eRutaIncial As String = "") As String
                Dim paraDevolver As String = ""

                Try
                    If Not String.IsNullOrEmpty(eRutaIncial) AndAlso IO.Directory.Exists(eRutaIncial) Then
                        _DIALOGO_APERTURA_FOLDER.SelectedPath = eRutaIncial
                    Else
                        _DIALOGO_APERTURA_FOLDER.SelectedPath = My.Computer.FileSystem.SpecialDirectories.Desktop
                    End If

                    If _DIALOGO_APERTURA_FOLDER.ShowDialog() = DialogResult.OK Then
                        If System.IO.Directory.Exists(_DIALOGO_APERTURA_FOLDER.SelectedPath) Then
                            paraDevolver = _DIALOGO_APERTURA_FOLDER.SelectedPath
                        End If
                    End If
                Catch ex As System.Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de buscar una carpeta.", ex, New StackTrace(0, True))
                    paraDevolver = ""
                End Try

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Permite seleccionar un directorio asegurándose que este sea una carpeta
            ''' compartida de red 
            ''' </summary>
            ''' <param name="eRutaInicial">Ruta inicial para la apertura</param>
            ''' <returns>Ruta completa al directorio seleccionado o una cadena vacía</returns>
            Public Function buscarDirectorioRed(Optional ByVal eRutaInicial As String = "") As String
                Dim paraDevolver As String = ""

                Try
                    Dim type As Type = _DIALOGO_APERTURA_FOLDER.[GetType]
                    Dim fieldInfo As Reflection.FieldInfo = type.GetField("rootFolder", BindingFlags.NonPublic Or BindingFlags.Instance)
                    fieldInfo.SetValue(_DIALOGO_APERTURA_FOLDER, DirectCast(18, Environment.SpecialFolder))
                    If IO.Directory.Exists(eRutaInicial) Then
                        _DIALOGO_APERTURA_FOLDER.SelectedPath = eRutaInicial
                    Else
                        _DIALOGO_APERTURA_FOLDER.SelectedPath = ""
                    End If

                    If _DIALOGO_APERTURA_FOLDER.ShowDialog() = DialogResult.OK Then
                        If IO.Directory.Exists(_DIALOGO_APERTURA_FOLDER.SelectedPath) Then
                            paraDevolver = _DIALOGO_APERTURA_FOLDER.SelectedPath
                        End If
                    End If
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de buscar una carpeta en red.", ex, New StackTrace(0, True))
                    paraDevolver = ""
                End Try

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Permite seleccionar una carpeta asegurándose que este es local
            ''' </summary>
            ''' <param name="eRutaInicial">Ruta inicial para la apertura</param>
            ''' <returns>Ruta completa al directorio seleccionado o una cadena vacía</returns>
            Public Function buscarDirectorioLocal(ByVal eRutaInicial As String) As String
                Dim paraDevolver As String = ""

                Try
                    Dim type As Type = _DIALOGO_APERTURA_FOLDER.[GetType]
                    Dim fieldInfo As Reflection.FieldInfo = type.GetField("rootFolder", BindingFlags.NonPublic Or BindingFlags.Instance)
                    fieldInfo.SetValue(_DIALOGO_APERTURA_FOLDER, DirectCast(17, Environment.SpecialFolder))
                    If IO.Directory.Exists(eRutaInicial) Then
                        _DIALOGO_APERTURA_FOLDER.SelectedPath = eRutaInicial
                    Else
                        _DIALOGO_APERTURA_FOLDER.SelectedPath = ""
                    End If

                    If _DIALOGO_APERTURA_FOLDER.ShowDialog() = DialogResult.OK Then
                        If IO.Directory.Exists(_DIALOGO_APERTURA_FOLDER.SelectedPath) Then
                            paraDevolver = _DIALOGO_APERTURA_FOLDER.SelectedPath
                        End If
                    End If
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de buscar una carpeta local.", ex, New StackTrace(0, True))
                    paraDevolver = ""
                End Try

                Return paraDevolver
            End Function
#End Region

        End Module
    End Namespace
End Namespace

