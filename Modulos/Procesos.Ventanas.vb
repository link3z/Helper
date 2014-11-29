Namespace Procesos
    Namespace Ventanas
        Module modProcesosVentanas
#Region " DECLARACIONES API WINDOWS "
            Public Structure POINTAPI
                Public x As Integer
                Public y As Integer
            End Structure

            Public Structure RECT
                Public Left As Integer
                Public Top As Integer
                Public Right As Integer
                Public Bottom As Integer
            End Structure

            Public Structure WINDOWPLACEMENT
                Public Length As Integer
                Public flags As Integer
                Public showCmd As Integer
                Public ptMinPosition As POINTAPI
                Public ptMaxPosition As POINTAPI
                Public rcNormalPosition As RECT
            End Structure

            Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer
            Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As IntPtr, ByRef ProcessID As Integer) As Integer
            Public Delegate Function CallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean
            Public Declare Function EnumWindows Lib "user32" (ByVal Adress As CallBack, ByVal y As Integer) As Integer
            Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
            Public Declare Function GetForegroundWindow Lib "user32.dll" () As IntPtr
            Public Declare Function GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As IntPtr
            Public ActiveWindows As New System.Collections.ObjectModel.Collection(Of IntPtr)
            Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpWindowText As String, ByVal cch As Integer) As Long
            Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
            Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As UInteger, ByVal wParam As Integer, ByVal lParam As Integer) As Boolean
            Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As IntPtr) As Int32
#End Region

            ''' <summary>
            ''' Obtiene un listado con los punteros a las ventanas que se encuentran activas en windows
            ''' </summary>
            ''' <param name="eConExcepcion">Determina si se debe lanzar una excepción si se produce un error</param>
            ''' <returns>Listado a los punteros que manejan las ventanas activas</returns>
            Public Function obtenerVentanasActivas(Optional ByVal eConExcepcion As Boolean = False) As ObjectModel.Collection(Of IntPtr)
                Try
                    Return ActiveWindows
                Catch ex As Exception
                    If eConExcepcion Then
                        Throw New Exception("Error al obtener el listado de ventnaas activas...", ex)
                    Else
                        Return Nothing
                    End If
                End Try
            End Function

            ''' <summary>
            ''' Obtiene un listado con la información de las ventanas activas en Windows
            ''' </summary>
            ''' <param name="eConExcepcion">Determina si se tiene que lanzar una excepción en caso de error</param>
            ''' <returns>Listado con la información relevante de las ventnas</returns>
            Public Function obtenerInfoVentanasActivas(Optional ByVal eConExcepcion As Boolean = False) As List(Of strInfoVentana)
                Try
                    iInfoVentanas = New List(Of strInfoVentana)
                    EnumWindows(AddressOf enumerarVentanas, 0)
                    Return iInfoVentanas
                Catch ex As Exception
                    If eConExcepcion Then
                        Throw New Exception("Se ha producido un error al tratar de obtener información de las ventanas...", ex)
                    Else
                        Return Nothing
                    End If
                End Try
            End Function
            Private iInfoVentanas As List(Of strInfoVentana)

            ''' <summary>
            ''' Obtiene información sobre la ventana asociada al puntero que se le pasa como parámetro
            ''' </summary>
            ''' <param name="eHWND">Puntero a la ventana de la que se quiere obtener información</param>
            ''' <param name="eConExcepcion">Lanza una excepción en caso de error</param>
            ''' <returns>Información de la ventana o Nothing en caso de no localizarla</returns>
            Public Function obtenerInfoVentana(ByVal eHWND As IntPtr, _
                                               Optional ByVal eConExcepcion As Boolean = False) As strInfoVentana
                Dim paraDevolver As strInfoVentana = Nothing

                Try
                    If IsWindowVisible(eHWND) Then
                        Dim wpTemp As WINDOWPLACEMENT
                        Dim intRet As Integer

                        Dim text As String = Space(GetWindowTextLength(eHWND) + 1)
                        GetWindowText(eHWND, text, Len(text))
                        If text.Length > 0 Then text = text.Substring(0, text.Length - 1)
                        text = text.Trim

                        wpTemp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wpTemp)
                        intRet = GetWindowPlacement(eHWND.ToInt32, wpTemp)

                        Dim elPID As Long
                        GetWindowThreadProcessId(eHWND, elPID)

                        If Not String.IsNullOrEmpty(text) Then
                            paraDevolver = New strInfoVentana
                            With paraDevolver
                                .eNombre = text
                                .eHWND = eHWND

                                .eLeft = wpTemp.rcNormalPosition.Left
                                .eTop = wpTemp.rcNormalPosition.Top
                                .eBotton = wpTemp.rcNormalPosition.Bottom
                                .eRight = wpTemp.rcNormalPosition.Right

                                .eCMD = wpTemp.showCmd
                                .ePID = elPID
                            End With
                        End If
                    End If
                Catch ex As Exception
                    If eConExcepcion Then
                        Throw New Exception("Se ha producido un error al obtener información de la ventana asociada al HWND " & eHWND.ToString & ".", ex)
                    Else
                        paraDevolver = Nothing
                    End If
                End Try

                Return paraDevolver
            End Function

            ''' <summary>
            ''' Metodo interno utilizado para obtener información de todas las ventanas
            ''' </summary>
            ''' <param name="eHWND">Puntero que apunta a la ventana de la que se va a obtener la información</param>
            ''' <param name="eParam">Necesario para un correcta firma del método</param>
            ''' <returns>True o False dependiendo de si se pudo ejecutar correctamente</returns>
            ''' <remarks>Este método se utiliza internamente</remarks>
            Private Function enumerarVentanas(ByVal eHWND As IntPtr, _
                                              ByVal eParam As Integer) As Boolean
                Try
                    Dim intRet As Integer
                    Dim wpTemp As WINDOWPLACEMENT

                    If IsWindowVisible(eHWND) Then
                        Dim text As String = Space(GetWindowTextLength(eHWND) + 1)
                        GetWindowText(eHWND, text, Len(text))
                        If text.Length > 0 Then text = text.Substring(0, text.Length - 1)
                        text = text.Trim

                        wpTemp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wpTemp)
                        intRet = GetWindowPlacement(eHWND.ToInt32, wpTemp)

                        Dim elPID As Long
                        GetWindowThreadProcessId(eHWND, elPID)

                        If Not String.IsNullOrEmpty(text) Then

                            Dim nuevaVentana As New strInfoVentana
                            With nuevaVentana
                                .eNombre = text
                                .eHWND = eHWND

                                .eLeft = wpTemp.rcNormalPosition.Left
                                .eTop = wpTemp.rcNormalPosition.Top
                                .eBotton = wpTemp.rcNormalPosition.Bottom
                                .eRight = wpTemp.rcNormalPosition.Right

                                .eCMD = wpTemp.showCmd
                                .ePID = elPID
                            End With
                            If Not iInfoVentanas.Contains(nuevaVentana) Then iInfoVentanas.Add(nuevaVentana)
                        End If
                    End If
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            End Function
        End Module
    End Namespace
End Namespace
