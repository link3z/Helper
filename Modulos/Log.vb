Imports System.IO
Imports System.Windows.Forms

Namespace Log
    ''' <summary>
    ''' Funciones necesarias para la implementación del sistema de LOG con rotación
    ''' </summary>
    Public Module modRecompilaHelperLog
#Region " PROPIEDADES "
        ''' <summary>
        ''' Flag que controla si está activo el sistema de LOG de debug, este sistema
        ''' escribirá todo lo que se pase a la función escribirLOG para poder depurar 
        ''' la aplicación
        ''' </summary>            
        Public Property _LOG_ACTIVO As Boolean = False

        ''' <summary>
        ''' Timer utilizado para la rotación de los archivos de LOG, evitando que estos
        ''' crezcan de forma desproporcionada y mantiendo los logs más modernos. Se utiliza
        ''' conjuntamente con _ROTACION_LOG para determinar cuandos archivos de log
        ''' se van a mantener
        ''' </summary>
        Private Property _TIMER_ROTACION As Timers.Timer

        ''' <summary>
        ''' Determina cada cuanto tiempo se tiene que ejecutar el _TIMER_ROTACION para 
        ''' eliminar los archivos de log antiguos y rotar los nuevos
        ''' </summary>
        ''' <remarks>El intervalo de rotación se establece en minutos</remarks>
        Private Property _INTERVALO_ROTACION As Integer = 5

        ''' <summary>
        ''' Número de ficheros de log que se mantendrán antes de ser eliminados de forma
        ''' automática por un ciclo de rotación de log. Se mantendrán siempre los 
        ''' _ROTACION_LOG archivos de log más actuales
        ''' </summary>
        Private Property _ROTACION_LOG As Short = 10

        ''' <summary>
        ''' Determina el tamaño del código identificativo de la linea de código desde donde
        ''' se está llamando al log si se le pasa información de la pila. Esto permite identificar
        ''' el fichero, el método y la linea desde donde se llama a la salida del log.
        ''' </summary>
        Private Property _TAMANHO_CODIGO_PILA As Integer = 35

        ''' <summary>
        ''' Flag para evitar escribir en el log mientras se está realizando una
        ''' rotación de este y evitar que se produzca un error de escritura si el archivo
        ''' se encuentra en uso
        ''' </summary>
        Private Property _ROTANDO_LOG As Boolean = False

        ''' <summary>
        ''' Ruta donde se guardará el LOG del fichero de debug. Estas ruta se inicializa
        ''' cuando se activa el sistema de Log. Se iran generando documentos de log
        ''' con el nombre del fichero _XXX donde XXX será incremental cuando más antiguo
        ''' sea el fichero de log
        ''' </summary>
        Private Property _RUTA_FICHERO_LOG As String = ""
#End Region

#Region " METODOS "
        ''' <summary>
        ''' Activa las funcionalidades del Debug. Si no se le pasa ningún parámetro, el Debug se iniciará
        ''' si existe un fichero llamado .conDebug en la unidad donde se encuentre instalado Windows 
        ''' o en el directorio desde donde se está ejecutando la aplicación
        ''' </summary>
        ''' <param name="eRutaFicheroLog">Ruta donde se creará el fichero de LOG</param>
        ''' <param name="eActivar">Forzar la activación del sistema de log</param>
        ''' <param name="eIntervaloRotacion">Establece el intervalo de rotación del log en minutos</param>
        ''' <param name="eTamanhoCodigoPila">Tamaño del código para la identifiación de la pila de llamada</param>
        Public Sub iniciarSistemaLog(ByVal eRutaFicheroLog As String,
                                     Optional ByVal eActivar As Boolean = False, _
                                     Optional ByVal eIntervaloRotacion As Integer = 5, _
                                     Optional ByVal eTamanhoCodigoPila As Integer = 35)
            Try
                ' Se calculan las rutas de activación del log de forma automática
                Dim rutaFicheroConDebugGeneral As String = Environment.GetFolderPath(Environment.SpecialFolder.System).Substring(0, 1) & ":\.conDebug"
                Dim rutaFicheroConDebugParticular As String = Ficheros.StartUpPath & "\.conDebug"

                ' Se activa el log validando cualquiera de las tres opciones de inicio
                _LOG_ACTIVO = (File.Exists(rutaFicheroConDebugGeneral) Or File.Exists(rutaFicheroConDebugParticular) Or eActivar)

                If _LOG_ACTIVO = True Then
                    ' Inicialización de las variables
                    _RUTA_FICHERO_LOG = eRutaFicheroLog
                    _INTERVALO_ROTACION = eIntervaloRotacion
                    _TIMER_ROTACION = New Timers.Timer(_INTERVALO_ROTACION * 60 * 1000)
                    _TAMANHO_CODIGO_PILA = eTamanhoCodigoPila

                    ' Creación de los manejadores para la rotación del log
                    AddHandler _TIMER_ROTACION.Elapsed, AddressOf realizarRotacion
                    _TIMER_ROTACION.Start()

                    ' Se escribe la cabecera en el fichero de log
                    escribirLog("LOG INICIADO A LAS " & Now)
                    escribirLog(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & " | V " & My.Application.Info.Version.ToString)
                    escribirLog("-----------------------------------------------------")
                End If
            Catch ex As Exception
                ' Si se produce algún error en la inicialización del log este no se puede activar
                ' La mayor parte de estos errores se producen por problemas de permisos en la carpeta de
                ' salida del fichero de log
                _LOG_ACTIVO = False
            End Try
        End Sub

        ''' <summary>
        ''' Evento que se ejecuta a intervalos regulares para realizar la rotación de los ficheros de log
        ''' </summary>
        Private Sub realizarRotacion(sender As Object, e As Timers.ElapsedEventArgs)
            ' Se activa el flag de rotación para evitar cambios en el fichero de log mientras se rota
            _ROTANDO_LOG = True

            Try
                Dim rutaFichero As String = Ficheros.extraerRutaFichero(_RUTA_FICHERO_LOG)
                Dim nombreFichero As String = Ficheros.extraerNombreFicheroSinExtension(_RUTA_FICHERO_LOG)

                ' Si existe el último fichero de log, este se elimina para poder realizar
                ' la rotación y eliminar el log más antiguo
                Dim nombreUltimoLog As String = rutaFichero & "\" & nombreFichero & "_" & Format(_ROTACION_LOG, "000") & ".log"
                If IO.File.Exists(nombreUltimoLog) Then
                    Try
                        IO.File.Delete(nombreUltimoLog)
                    Catch ex As Exception
                    End Try
                End If

                ' Se recorren todos los ficheros de log empezando por el ultimo y moviendolos
                ' al siguiente, dejando libre el fichero de log inicial 
                For i As Integer = _ROTACION_LOG To 1 Step -1
                    Dim nombreOrigen As String = rutaFichero & "\" & nombreFichero & "_" & Format(i - 1, "000") & ".log"
                    Dim nombreDestino As String = rutaFichero & "\" & nombreFichero & "_" & Format(i, "000") & ".log"

                    Try
                        IO.File.Move(nombreOrigen, nombreDestino)
                    Catch ex As Exception
                    End Try
                Next
            Catch ex As Exception
                ' Si se produce un error en la rotación este se ignora, en la siguiente
                ' rotación se volverá a rotar en el peor de los casos.
            End Try

            _ROTANDO_LOG = False
        End Sub

        ''' <summary>
        ''' Escribe en el log el mensaje y los datos de la excepción que se le pasan como parámetro
        ''' </summary>
        ''' <param name="eMensaje">Texto que se va a escribir en el fichero de debug</param>
        ''' <param name="eExcepcion">Excepción que se produce si se trata de un error</param>
        ''' <param name="eStack">Información de la pila de llamadas para mostrar donde se genera el log</param>
        Public Sub escribirLog(ByVal eMensaje As String, _
                               Optional ByVal eExcepcion As System.Exception = Nothing, _
                               Optional ByVal eStack As System.Diagnostics.StackTrace = Nothing)
            Try
                ' Si se está realizando una rotación de log se espera a que esta se complete
                While _ROTANDO_LOG
                    Application.DoEvents()
                End While

                ' Se obtienen la ruta del fichero de log a utilizar
                Dim rutaFichero As String = Ficheros.extraerRutaFichero(_RUTA_FICHERO_LOG)
                Dim nombreFichero As String = Ficheros.extraerNombreFicheroSinExtension(_RUTA_FICHERO_LOG)
                Dim nombreFinal As String = rutaFichero & "\" & nombreFichero & "_000.log"

                ' Se genera el código identificativo de la pila de llamadas
                Dim Codigo As String = ""
                If eStack IsNot Nothing Then
                    Try
                        Dim eFrame As StackFrame = eStack.GetFrame(0)
                        Codigo = Ficheros.extraerNombreFicheroSinExtension(eFrame.GetFileName) & "." & eFrame.GetFileLineNumber.ToString.PadLeft(4, "0")
                        If Codigo.Length > _TAMANHO_CODIGO_PILA Then
                            Codigo = Codigo.Substring(Codigo.Length - _TAMANHO_CODIGO_PILA)
                        ElseIf Codigo.Length < _TAMANHO_CODIGO_PILA Then
                            Codigo = Codigo.PadLeft(_TAMANHO_CODIGO_PILA, " ")
                        End If
                    Catch ex As Exception
                    End Try
                End If

                ' Se crea el escritor encargado de insertar las nuevas lineas de LOG
                Dim elEscritor As New StreamWriter(nombreFinal, True)
                elEscritor.Write(Format(Now, "HH:mm:ss") & " [" & Codigo & "]")
                elEscritor.WriteLine(" " & eMensaje)

                ' Si la llamada incluye una excepción esta se añade al fichero de log para poder
                ' ser analizada en el fichero
                If eExcepcion IsNot Nothing Then
                    With elEscritor
                        .WriteLine("--- EXCEPCIÓN !!-------------------------------------")
                        .WriteLine(eExcepcion.Message)
                        .WriteLine()
                        .WriteLine(eExcepcion.StackTrace)
                        .WriteLine()
                        .WriteLine(eExcepcion.HelpLink)
                        .WriteLine()
                        If eExcepcion.InnerException IsNot Nothing Then
                            .WriteLine(eExcepcion.InnerException.Message)
                            .WriteLine()
                            .WriteLine(eExcepcion.InnerException.StackTrace)
                            .WriteLine()
                            .WriteLine(eExcepcion.InnerException.HelpLink)
                            .WriteLine()
                        End If
                        .WriteLine("-----------------------------------------------------")
                    End With
                End If
                elEscritor.Close()
            Catch ex As System.Exception
            End Try
        End Sub
#End Region
    End Module
End Namespace


