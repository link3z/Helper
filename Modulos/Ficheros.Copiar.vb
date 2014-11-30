Imports System.IO

Namespace Ficheros
    Namespace Copiar
        Public Module modRecompilaFicherosCopiar
#Region " METODOS "
            ''' <summary>
            ''' Copia un archivo de origen en destino. 
            ''' Se puede especificar que se cree la estructura de directorios de destino
            ''' si esta no existe creando la parte de la ruta que no existe
            ''' </summary>
            ''' <param name="eOrigen">Archivo d eorigen que quiere ser copiado</param>
            ''' <param name="eDestino">Archiov de destino que quiere ser copiado</param>
            ''' <param name="eSobreescribir">Si se va a sobreescribir el destino o no</param>
            ''' <param name="eCrearRuta">Si se tiene que crear la ruta d edestino si no existe</param>
            ''' <returns>True o False dependiendo de si se pudo copiar o no el archivo</returns>
            Public Function copiarArchivo(ByVal eOrigen As String, _
                                          ByVal eDestino As String, _
                                          ByVal eSobreescribir As Boolean, _
                                          ByVal eCrearRuta As Boolean) As Boolean
                Dim lasPartes() As String = Ficheros.splitDirectorios(eDestino.Replace("/", "\"))

                If eCrearRuta Then
                    Dim rutaBase As String = lasPartes(0) & "\"
                    Dim rutaParcial As String = rutaBase

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
            ''' creando las rutas si estas no existen
            ''' </summary>
            ''' <param name="eOrigen">DirectoryInfo con el directorio de origen</param>
            ''' <param name="eDestino">DirectoryInfo con el directorio de destino</param>
            ''' <param name="eSobreescribir">Determina si se sobreescriben los archivos si estos ya existen en destino</param>
            ''' <param name="eLabelInfo">Label para mostrar el estado de la copia</param>
            ''' <param name="eBarraProgreso">Barra de progreso para mostrar el avance de la copia</param>
            Public Function copiarArchivos(ByVal eOrigen As DirectoryInfo, _
                                           ByVal eDestino As DirectoryInfo, _
                                           ByVal eSobreescribir As Boolean, _
                                           Optional ByVal eLabelInfo As Object = Nothing, _
                                           Optional ByVal eBarraProgreso As Object = Nothing) As Boolean
                Dim paraDevolver As Boolean = True

                ' Determina si el progreso se va a ejecutar por ficheros o por carpetas
                Dim usarProgresoFile As Boolean = True

                ' Se obtiene todos los archivos del directorio de origen
                Dim fiSourceFiles() As FileInfo
                fiSourceFiles = eOrigen.GetFiles()

                ' Se obtienen todas las carpetas del directorio de origen
                Dim diSourceSubDirectories() As DirectoryInfo
                diSourceSubDirectories = eOrigen.GetDirectories()

                ' Si no existe el directorio destino crearlo
                If Not eDestino.Exists Then eDestino.Create()

                ' Si existen subdirectorios se utilizarán estos para indicar
                ' el progreso, en caso contrario, el progreso se indica mediante
                ' los archivos copiados, recorriendo cada uno de los subdirectorios
                ' y aplicando de forma recursiva la misma función sin pasar los
                ' parámetros de progreso
                If diSourceSubDirectories IsNot Nothing AndAlso diSourceSubDirectories.Count > 0 Then
                    Try
                        If eBarraProgreso IsNot Nothing Then
                            WinForms.ProgressBar.fijarMaximoBarra(eBarraProgreso, diSourceSubDirectories.Length + 1)
                        End If
                    Catch ex As Exception
                    End Try

                    usarProgresoFile = False

                    For Each unaCarpeta As DirectoryInfo In diSourceSubDirectories
                        copiarArchivos(unaCarpeta, New DirectoryInfo(eDestino.FullName & "\" & unaCarpeta.Name), eSobreescribir, eLabelInfo, Nothing)

                        Try
                            If eBarraProgreso IsNot Nothing Then
                                WinForms.ProgressBar.AumentarBarra(eBarraProgreso)
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                End If

                ' Se realiza la copia de cada archivo 
                For Each unArchivo As FileInfo In fiSourceFiles
                    Dim archivoDestino As String = eDestino.FullName + "\" + unArchivo.Name
                    If Log._LOG_ACTIVO Then Log.escribirLog(unArchivo.FullName & " -> " & archivoDestino, , New StackTrace(0, True))

                    ' Se actualiza la información del progreso
                    Try
                        If eLabelInfo IsNot Nothing Then eLabelInfo.Text = archivoDestino
                        If eBarraProgreso IsNot Nothing AndAlso usarProgresoFile Then
                            WinForms.ProgressBar.AumentarBarra(eBarraProgreso)
                        End If
                    Catch ex As Exception
                    End Try

                    ' Si hay que sobreescribir el archivo este se elimina antes utilizando
                    ' el KILL de procesos por si se trata de un ejecutable y e este está
                    ' en uso
                    If System.IO.File.Exists(archivoDestino) AndAlso eSobreescribir Then
                        Try
                            System.IO.File.Delete(archivoDestino)
                        Catch ex As Exception
                            If Log._LOG_ACTIVO Then Log.escribirLog("No se pudo borrar el archivo de destino '" & archivoDestino & "' (Forzando Kill Proceso)", , New StackTrace(0, True))
                            Try
                                Dim nombreProceso As String = Ficheros.extraerNombreFichero(unArchivo.Name)
                                Try
                                    If eLabelInfo IsNot Nothing Then eLabelInfo.Text = "Deteniendo el proceso " & nombreProceso & " para sobreescribir archivo."
                                Catch exProceso As Exception
                                End Try
                                Procesos.mataProcesoWhile(nombreProceso, 3)
                                System.IO.File.Delete(archivoDestino)
                            Catch
                                If Log._LOG_ACTIVO Then Log.escribirLog("ERROR eliminando archivo (incluso con KILL) '" & archivoDestino & "'...", ex, New StackTrace(0, True))
                            End Try
                        End Try
                    End If

                    ' Se realiza la copia del archivo
                    Try
                        unArchivo.CopyTo(eDestino.FullName + "\" + unArchivo.Name, eSobreescribir)
                    Catch ex As Exception
                        If Log._LOG_ACTIVO Then Log.escribirLog("ERRROR al copiar " & unArchivo.FullName & " -> " & archivoDestino, ex, New StackTrace(0, True))
                    End Try
                Next

                Return True
            End Function
#End Region
        End Module
    End Namespace
End Namespace

