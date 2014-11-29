Imports System.Windows.Forms

Namespace Windows
    Namespace UAC
        Public Module modWindowsUAC
            ''' <summary>
            ''' Permite ejecutar un comando o aplicación utilizando UAC para subir el nivel
            ''' de ejecución si fuese necesario
            ''' </summary>
            ''' <param name="eComando">Ruta al ejecutable que se va a ejecutar</param>
            ''' <param name="eWorkingDirectory">Directorio donde se tiene que realizar la ejecución</param>
            ''' <param name="eParametros">Parámetros para la ejecución del programa o comando</param>
            ''' <param name="eUseShell">Determina si se va utilizar Shell para la ejecución o no</param>
            ''' <param name="eWindowStyle">Como se debe mostrar la ventana si es una aplicación WinForms</param>
            ''' <param name="eConExcepcion">Determina si se tiene que lanzar una excepción en caso de error</param>
            ''' <returns>True o False dependiendo de si se pudo ejecutar el comando</returns>
            ''' <remarks></remarks>
            Public Function ejecutarUAC(ByVal eComando As String, _
                                        ByVal eWorkingDirectory As String, _
                                        ByVal eParametros As String, _
                                        ByVal eUseShell As Boolean, _
                                        ByVal eWindowStyle As ProcessWindowStyle, _
                                        Optional ByVal eConExcepcion As Boolean = False) As Boolean
                Try
                    Dim laInfoInicio As New ProcessStartInfo
                    With laInfoInicio
                        .FileName = eComando
                        .WorkingDirectory = eWorkingDirectory
                        If Not String.IsNullOrEmpty(eParametros) Then .Arguments = eParametros
                        .UseShellExecute = eUseShell
                        .WindowStyle = eWindowStyle
                        .Verb = "runas"
                    End With

                    Dim elProceso As Process = Nothing
                    elProceso = Process.Start(laInfoInicio)
                Catch ex As Exception
                    If eConExcepcion Then
                        Throw New Exception("No se ha podido ejecutar '" & eComando & "'.", ex)
                    Else
                        Return False
                    End If
                End Try

                Return True
            End Function
        End Module
    End Namespace
End Namespace
