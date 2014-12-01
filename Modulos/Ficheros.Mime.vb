Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.Win32
Imports System.Drawing.Imaging

Namespace Ficheros
    Namespace Mime
        Public Module modRecompilaFicherosMime
#Region " API WINDOWS "
            <DllImport("urlmon.dll", CharSet:=CharSet.Auto)> _
            Private Function FindMimeFromData(ByVal pBC As System.UInt32, <MarshalAs(UnmanagedType.LPStr)> ByVal pwzUrl As System.String, <MarshalAs(UnmanagedType.LPArray)> ByVal pBuffer As Byte(), ByVal cbSize As System.UInt32, <MarshalAs(UnmanagedType.LPStr)> ByVal pwzMimeProposed As System.String, ByVal dwMimeFlags As System.UInt32, _
                                              ByRef ppwzMimeOut As System.UInt32, ByVal dwReserverd As System.UInt32) As System.UInt32
            End Function
#End Region

#Region " METODOS "
            ''' <summary>
            ''' Obtiene el tipo de archivo correcto en función del contenido del mismo
            ''' </summary>
            ''' <param name="eRutaArchivo">Ruta del archivo a analizar</param>
            ''' <returns>Tipo mime que determina el tipo de archivo</returns>
            Public Function obtenerTipoArchivo(ByVal eRutaArchivo As String) As String
                Dim paraDevolver As String = "unknown/unknown"

                If IO.File.Exists(eRutaArchivo) Then
                    Try
                        ' Se lee la cabecera del archivo (255 bytes) donde se encuentra
                        ' el tipo mime
                        Dim elBuffer As Byte() = New Byte(255) {}
                        Using elFileStream As New FileStream(eRutaArchivo, FileMode.Open)
                            If elFileStream.Length >= 256 Then
                                elFileStream.Read(elBuffer, 0, 256)
                            Else
                                elFileStream.Read(elBuffer, 0, CInt(elFileStream.Length))
                            End If
                        End Using

                        paraDevolver = obtenerTipoArchivo(eRutaArchivo)
                    Catch ex As Exception
                        If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de obtener el tipo MIME del archivo '" & eRutaArchivo & "'...", ex, New StackTrace(0, True))
                        paraDevolver = "unknown/unknown"
                    End Try
                Else
                    If Log._LOG_ACTIVO Then Log.escribirLog("No existe el archivo '" & eRutaArchivo & "' por lo que no se puede obtener su tipo MIME...", , New StackTrace(0, True))
                End If

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Obtiene el tipo de archivo correcto en función del contenido del mismom
            ''' a partir del array de bytes del fichero
            ''' </summary>
            ''' <param name="eBytes">Contenido del archivo o cabecera de este</param>
            ''' <returns>Mime del archivo</returns>
            Public Function obtenerTipoArchivo(ByVal eBytes As Byte()) As String
                Dim paraDevolver As String = "unknown/unknown"

                ' Si no existe la cabecdera del fichero se devuelve el mime desconocido
                If eBytes Is Nothing OrElse eBytes.Length = 0 Then Return paraDevolver

                ' El tipo mime está en los 255 bytes
                Dim buffer As Byte() = New Byte(255) {}
                Dim Longitud As Integer = 256

                If eBytes.Length < 256 Then
                    Longitud = eBytes.Length
                End If

                For i As Integer = 0 To Longitud - 1
                    buffer(i) = eBytes(i)
                Next

                Try
                    Dim mimetype As System.UInt32
                    FindMimeFromData(0, Nothing, buffer, 256, Nothing, 0, mimetype, 0)
                    Dim mimeTypePtr As System.IntPtr = New IntPtr(mimetype)
                    Dim mime As String = Marshal.PtrToStringUni(mimeTypePtr)
                    Marshal.FreeCoTaskMem(mimeTypePtr)
                    paraDevolver = mime
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Error obteniendo la extensión Mime...", ex, New StackTrace(0, True))
                    paraDevolver = "unknown/unknown"
                End Try

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Obtiene la extensión por defecto del tipo mime especificado
            ''' </summary>
            ''' <param name="eMime">Tipo mime del que se quiere obtener la extensión por defecto</param>
            ''' <returns>Extensión por defecto asociada a un tipo mime</returns>
            Public Function extensionPorDefecto(ByVal eMime As String) As String
                Dim paraDevolver As String = ""

                Try
                    Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey("MIME\Database\Content Type\" & eMime, False)
                    Dim value As Object = If(key IsNot Nothing, key.GetValue("Extension", Nothing), Nothing)
                    extensionPorDefecto = If(value IsNot Nothing, value.ToString(), String.Empty)
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error obteniendo pa extensión asociada al tipo mime '" & eMime & "'...", ex, New StackTrace(0, True))
                    paraDevolver = String.Empty
                End Try

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Obtiene información de un codec de imagen a partir de su tipo mime
            ''' </summary>
            ''' <param name="eMime">Tipo mime del que se quiere obtener la información</param>
            ''' <returns>Información del codec</returns>
            Public Function obtenerCodec(ByVal eMime As String) As ImageCodecInfo
                Return Imagenes.obtenerCodec(eMime)
            End Function
#End Region
        End Module
    End Namespace
End Namespace
