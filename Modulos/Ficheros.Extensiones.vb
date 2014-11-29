Imports System.Drawing

Namespace Ficheros
    Namespace Extensiones
        Public Module modRecompilaFicherosExtensiones
            ''' <summary>
            ''' Obtiene la imagen (icono) asociado a un archivo
            ''' </summary>
            ''' <param name="eRuta">Ruda del archivo del que se quiere obetner la imagen</param>
            ''' <returns>Icono asociada al archivo</returns>
            Public Function obtenerIconoArchivo(ByVal eRuta As String) As Icon
                If Not IO.File.Exists(eRuta) Then Return Nothing
                Dim paraDevolver As Icon = Nothing

                Try
                    Dim laExtension As String = extraerExtensionFicheroSinPunto(eRuta).Trim
                    If Not String.IsNullOrEmpty(laExtension) Then
                        paraDevolver = obtenerIconoExtension(laExtension)
                    End If
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de obtener el icono asociado a '" & eRuta & "'", ex, New StackTrace(0, True))
                    paraDevolver = Nothing
                End Try

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Obtiene el icono asociado a una extensión de un archivo
            ''' </summary>
            ''' <param name="eExtension">Extensión de la que se quiere obtener el icono asociado</param>
            ''' <returns>Icono asociado a la extensión</returns>
            Public Function obtenerIconoExtension(ByVal eExtension As String) As Icon
                Dim paraDevolver As Icon = Nothing

                If Not String.IsNullOrEmpty(eExtension) Then
                    Dim rutaEjecutable As String = obtenerProgramaAsociado(eExtension)
                    paraDevolver = Drawing.Icon.ExtractAssociatedIcon(rutaEjecutable)
                End If

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Obtiene el programa asociado a una determinada extensión de archivo
            ''' </summary>
            ''' <param name="eExtension">Extensión del fichero del que se quiere obtener el ejecutable</param>
            ''' <returns>Ruta al ejecutable asociado a la extensión</returns>
            Public Function obtenerProgramaAsociado(ByVal eExtension As String) As String
                Dim paraDevolver As String = ""

                ' Si no se indica extensión no se puede obtener la ruta
                If String.IsNullOrEmpty(eExtension) Then Return paraDevolver

                Try
                    ' Se le añade el punto a la extensión si no e le pasó
                    If eExtension.Substring(0, 1) <> "." Then eExtension = "." & eExtension

                    ' Se obtienen las claves del registro donde se almacena la informaicón de las apicaciones
                    Dim objExtReg As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot
                    Dim objAppReg As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot

                    objExtReg = objExtReg.OpenSubKey(eExtension.Trim)
                    paraDevolver = objExtReg.GetValue("").ToString
                    objAppReg = objAppReg.OpenSubKey(paraDevolver & "\shell\open\command")

                    ' Se eliminan los comodines de entrada de los ficheros para la aplicación
                    Dim SplitArray() As String
                    SplitArray = Split(objAppReg.GetValue(Nothing).ToString, """")
                    If SplitArray(0).Trim.Length > 0 Then
                        Return SplitArray(0).Replace("%1", "")
                    Else
                        Return SplitArray(1).Replace("%1", "")
                    End If
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al obtener el programa asociado a la extensión '" & eExtension & "'", ex, New StackTrace(0, True))
                    paraDevolver = ""
                End Try

                Return paraDevolver
            End Function
        End Module
    End Namespace
End Namespace

