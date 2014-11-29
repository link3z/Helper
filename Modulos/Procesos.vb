Imports System.Windows.Forms
Imports System.Drawing

Namespace Procesos
    Public Module modProcesos
        ''' <summary>
        ''' Determina si la aplicación ya está en ejecución, para ello, se busca el número de procesos
        ''' que se llamen igual en la lista de procesos en ejecución, si hay más de uno se detiene la ejecución
        ''' </summary>            
        ''' <returns>True o False dependiendo de si ya está en ejecución o no</returns>
        Public Function yaEnEjecucion() As Boolean
            ' Usamos la clase Process para sacar información de procesos en ejecución.                
            Dim losProcesos() As Process

            ' Asignamos a la matríz todos los procesos en ejecución que tengan por nombre el de nuestra aplicación.
            losProcesos = Process.GetProcessesByName(Application.ProductName.ToString)

            ' MisProcesos.Length nunca es cero, porque este mismo proceso cuenta;
            ' por eso miramos si solo hay una coincidencia. 
            Return (losProcesos.Length > 1)
        End Function

        Public Function BuscarProcesoPorRutaParcial(ByVal rutaParcial As String) As Long
            Dim paraDevolver As Long = -1
            If rutaParcial <> "" Then
                For Each proceso As Process In Process.GetProcesses()
                    Try
                        If proceso.MainModule.FileName.Contains(rutaParcial) Then
                            paraDevolver = proceso.Id
                            Exit For
                        End If
                    Catch ex As Exception
                        paraDevolver = -1
                    End Try
                Next
            End If

            Return paraDevolver
        End Function

        ''' <summary>
        ''' Busca un proceso a partir de la ruta donde se está ejecutando
        ''' </summary>
        ''' <param name="RutaEjecutable"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function BuscarProcesoPorRuta(ByVal RutaEjecutable As String) As Integer
            Dim paraDevolver As Long = -1
            If RutaEjecutable <> "" Then
                For Each proceso As Process In Process.GetProcesses()
                    Try
                        If proceso.MainModule.FileName = RutaEjecutable Then
                            paraDevolver = proceso.Id
                            Exit For
                        End If
                    Catch ex As Exception
                        paraDevolver = -1
                    End Try
                Next
            End If

            Return paraDevolver
        End Function

        Public Function buscarProcesoPorNombre(ByVal eNombreProceo As String) As Integer
            Dim paraDevolver As Long = -1
            If Not String.IsNullOrEmpty(eNombreProceo) Then
                Try
                    For Each proceso As Process In Process.GetProcesses()
                        Try
                            If Ficheros.extraerNombreFichero(proceso.MainModule.FileName).ToUpper = eNombreProceo.ToUpper Then
                                paraDevolver = proceso.Id
                                Exit For
                            End If
                        Catch ex As Exception
                            ' No se muestra el error ya que puede intentar acceder a procesos
                            ' del sistema y mostraría error siempre... XD
                        End Try
                    Next
                Catch ex As Exception
                End Try
            End If

            Return paraDevolver
        End Function

        Public Function buscarProcesoporNombreEmpieza(ByVal eNombreProceo As String) As Integer
            Dim paraDevolver As Long = -1
            If Not String.IsNullOrEmpty(eNombreProceo) Then
                Try
                    For Each proceso As Process In Process.GetProcesses()
                        Try
                            If Ficheros.extraerNombreFichero(proceso.MainModule.FileName).ToUpper.StartsWith(eNombreProceo.ToUpper) Then
                                paraDevolver = proceso.Id
                                Exit For
                            End If
                        Catch ex As Exception
                            ' No se muestra el error ya que puede intentar acceder a procesos
                            ' del sistema y mostraría error siempre... XD
                        End Try
                    Next
                Catch ex As Exception
                End Try
            End If

            Return paraDevolver

        End Function

        Public Function obtenerIconoEjecutable(ByVal rutaCompleta As String) As Bitmap
            Try
                Dim elIcono As Icon = Icon.ExtractAssociatedIcon(rutaCompleta)
                Return elIcono.ToBitmap
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Sub matarProceso(ByVal _idProceso As Long)
            Try
                Process.GetProcessById(_idProceso).Kill()
            Catch ex As Exception
            End Try
        End Sub

        Public Sub mataProcesoWhile(ByVal nombreProceso As String, ByVal numeroIntentos As Integer)
            Dim idProceso As Long = 1
            Dim Intentos = 0

            idProceso = 1
            Intentos = 0
            While (idProceso > 0) And (Intentos < numeroIntentos)
                idProceso = buscarProcesoPorNombre(nombreProceso)
                Try
                    If idProceso > 0 Then
                        Process.GetProcessById(idProceso).Kill()
                    Else
                        Intentos = numeroIntentos + 1
                    End If
                Catch ex As Exception
                    idProceso = 1
                    Intentos += 1
                End Try
            End While
        End Sub
    End Module
End Namespace

